using System;
using System.Collections.Concurrent;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace MailMind;

// Preview handlers and their child window live on a dedicated STA with a message pump.
public sealed class SystemPreviewHost : IDisposable
{
    private const uint WorkMessage = 0x8001;
    private readonly ConcurrentQueue<Action> work = new();
    private readonly ManualResetEventSlim ready = new();
    private readonly Thread thread;
    private uint threadId;
    private IntPtr child;
    private IntPtr owner;
    private IPreviewHandler? handler;
    private object? handlerObject;
    private PreviewSite? site;
    private bool disposed;
    private static readonly string[] Blocked = { ".exe", ".bat", ".cmd", ".ps1", ".vbs", ".js", ".scr", ".cpl", ".msi", ".jar", ".com", ".hta" };

    public SystemPreviewHost()
    {
        thread = new Thread(Run) { IsBackground = true, Name = "Attachment preview" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        ready.Wait();
    }
    private void Run()
    {
        threadId = GetCurrentThreadId();
        PeekMessage(out _, IntPtr.Zero, 0, 0, 0);
        ready.Set();
        while (GetMessage(out var msg, IntPtr.Zero, 0, 0) > 0)
        {
            if (msg.message == WorkMessage) { while (work.TryDequeue(out var action)) action(); }
            else { TranslateMessage(ref msg); DispatchMessage(ref msg); }
        }
        CloseCore();
    }
    private void Post(Action action) { if (disposed) return; work.Enqueue(action); PostThreadMessage(threadId, WorkMessage, UIntPtr.Zero, IntPtr.Zero); }
    public Task ShowAsync(IntPtr parent, MailAttachment attachment, int x, int y, int width, int height)
    {
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        Post(() =>
        {
            try
            {
                CloseCore();
                var extension = Path.GetExtension(attachment.Name).ToLowerInvariant();
                if (extension.Length < 2 || extension.Length > 16 || extension.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || Array.IndexOf(Blocked, extension) >= 0)
                    throw new InvalidOperationException(Strings.Get("More_9c718678dedd"));
                var clsid = FindHandler(extension) ?? throw new InvalidOperationException(Strings.Get("More_f76ea878382e"));
                var directory = Path.Combine(PreviewCacheRoot(), "MailMind", "cache", "attachments");
                Directory.CreateDirectory(directory);
                var file = Path.Combine(directory, Convert.ToHexString(SHA256.HashData(attachment.Data)) + extension);
                File.WriteAllBytes(file, attachment.Data);
                var type = Type.GetTypeFromCLSID(clsid, true)!;
                handlerObject = Activator.CreateInstance(type) ?? throw new InvalidOperationException(Strings.Get("More_ce2072231a75"));
                if (handlerObject is IObjectWithSite objectWithSite)
                {
                    site = new PreviewSite(parent);
                    objectWithSite.SetSite(site);
                }
                if (handlerObject is IInitializeWithFile fileInitializer) fileInitializer.Initialize(file, 0);
                else if (handlerObject is IInitializeWithStream streamInitializer)
                {
                    Marshal.ThrowExceptionForHR(SHCreateStreamOnFileEx(file, 0, 0, false, null, out var stream));
                    try { streamInitializer.Initialize(stream, 0); }
                    finally { Marshal.ReleaseComObject(stream); }
                }
                else throw new InvalidOperationException(Strings.Get("More_540b07df5820"));
                handler = (IPreviewHandler)handlerObject;
                // An owned popup avoids WinUI's compositor covering native child windows.
                owner = parent;
                child = CreateWindowEx(0x00000080, "STATIC", "", 0x80000000 | 0x10000000 | 0x04000000 | 0x02000000 | 0x00000100, -32000, -32000, Math.Max(width, 1), Math.Max(height, 1), parent, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                if (child == IntPtr.Zero) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                var rect = new RECT { right = Math.Max(width, 1), bottom = Math.Max(height, 1) };
                handler.SetWindow(child, ref rect); handler.DoPreview();
                var position = ScreenPosition(x, y);
                SetWindowPos(child, IntPtr.Zero, position.x, position.y, 0, 0, 0x0001 | 0x0040 | 0x0010);
                completion.TrySetResult();
            }
            catch (Exception ex) { CloseCore(); completion.TrySetException(ex); }
        });
        return completion.Task;
    }
    private static string PreviewCacheRoot()
    {
        // Packaged apps can have LocalAppData writes redirected to LocalCache.
        // Pass the physical path so out-of-process preview handlers can read it.
        try { return Windows.Storage.ApplicationData.Current.LocalCacheFolder.Path; }
        catch (InvalidOperationException) { return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData); }
        catch (COMException) { return Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData); }
    }
    private static Guid? FindHandler(string extension)
    {
        const string key = @"\shellex\{8895b1c6-b41f-4c1c-a562-0d564250836f}";
        using var ext = Registry.ClassesRoot.OpenSubKey(extension);
        var progId = ext?.GetValue(null) as string;
        foreach (var path in new[] { extension + key, progId == null ? null : progId + key, @"SystemFileAssociations\" + extension + key })
        {
            if (path == null) continue;
            using var registration = Registry.ClassesRoot.OpenSubKey(path);
            if (Guid.TryParse(registration?.GetValue(null) as string, out var id)) return id;
        }
        return null;
    }
    public void Resize(int x, int y, int width, int height) => Post(() =>
    {
        if (child == IntPtr.Zero || handler == null) return;
        try
        {
            var position = ScreenPosition(x, y);
            MoveWindow(child, position.x, position.y, Math.Max(width, 1), Math.Max(height, 1), true);
            var rect = new RECT { right = Math.Max(width, 1), bottom = Math.Max(height, 1) };
            handler.SetRect(ref rect);
        }
        catch (Exception ex) when (ex is COMException or System.ComponentModel.Win32Exception) { }
    });
    private POINT ScreenPosition(int x, int y)
    {
        var point = new POINT { x = x, y = y };
        if (!ClientToScreen(owner, ref point)) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        return point;
    }
    public void Close() => Post(CloseCore);
    private void CloseCore()
    {
        if (handler != null) { try { handler.Unload(); } catch (COMException) { } handler = null; }
        if (handlerObject is IObjectWithSite objectWithSite) { try { objectWithSite.SetSite(null); } catch (COMException) { } }
        if (handlerObject != null) { Marshal.ReleaseComObject(handlerObject); handlerObject = null; }
        site = null;
        if (child != IntPtr.Zero) { DestroyWindow(child); child = IntPtr.Zero; }
    }
    public void Dispose() { if (disposed) return; Post(CloseCore); disposed = true; PostThreadMessage(threadId, 0x0012, UIntPtr.Zero, IntPtr.Zero); }

    [ComImport, Guid("8895b1c6-b41f-4c1c-a562-0d564250836f"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPreviewHandler { void SetWindow(IntPtr hwnd, ref RECT rect); void SetRect(ref RECT rect); void DoPreview(); void Unload(); void SetFocus(); void QueryFocus(out IntPtr hwnd); void TranslateAccelerator(ref MSG msg); }
    [ComImport, Guid("b7d14566-0509-4cce-a71f-0a554233bd9b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IInitializeWithFile { void Initialize([MarshalAs(UnmanagedType.LPWStr)] string file, uint mode); }
    [ComImport, Guid("b824b49d-22ac-4161-ac8a-9916e8fa3f7f"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IInitializeWithStream { void Initialize(IStream stream, uint mode); }
    [ComImport, Guid("fc4801a3-2ba9-11cf-a229-00aa003d7352"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IObjectWithSite { void SetSite([MarshalAs(UnmanagedType.IUnknown)] object? site); void GetSite(ref Guid iid, out IntPtr site); }
    [ComVisible(true), Guid("00000114-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleWindow { void GetWindow(out IntPtr window); void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool enterMode); }
    [ComVisible(true), Guid("fec87aaf-35f9-447a-adb7-20234491401a"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPreviewHandlerFrame { void GetWindowContext(out PreviewHandlerFrameInfo info); [PreserveSig] int TranslateAccelerator(ref MSG message); }
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public sealed class PreviewSite : IOleWindow, IPreviewHandlerFrame
    {
        private readonly IntPtr window;
        public PreviewSite(IntPtr window) => this.window = window;
        public void GetWindow(out IntPtr result) => result = window;
        public void ContextSensitiveHelp(bool enterMode) { }
        public void GetWindowContext(out PreviewHandlerFrameInfo info) => info = default;
        public int TranslateAccelerator(ref MSG message) => 1; // S_FALSE: not handled.
    }
    [StructLayout(LayoutKind.Sequential)] public struct PreviewHandlerFrameInfo { public IntPtr accelerator; public uint acceleratorCount; }
    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int left, top, right, bottom; }
    [StructLayout(LayoutKind.Sequential)] private struct POINT { public int x, y; }
    [StructLayout(LayoutKind.Sequential)] public struct MSG { public IntPtr hwnd; public uint message; public UIntPtr wParam; public IntPtr lParam; public uint time; public int x, y; public uint lPrivate; }
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] private static extern int GetMessage(out MSG msg, IntPtr hwnd, uint min, uint max);
    [DllImport("user32.dll")] private static extern bool PeekMessage(out MSG msg, IntPtr hwnd, uint min, uint max, uint remove);
    [DllImport("user32.dll")] private static extern bool TranslateMessage(ref MSG msg);
    [DllImport("user32.dll")] private static extern IntPtr DispatchMessage(ref MSG msg);
    [DllImport("user32.dll")] private static extern bool PostThreadMessage(uint thread, uint message, UIntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr CreateWindowEx(uint exStyle, string className, string windowName, uint style, int x, int y, int width, int height, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr param);
    [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr hwnd, IntPtr insertAfter, int x, int y, int width, int height, uint flags);
    [DllImport("user32.dll")] private static extern bool MoveWindow(IntPtr hwnd, int x, int y, int width, int height, bool repaint);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool ClientToScreen(IntPtr hwnd, ref POINT point);
    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, PreserveSig = true)] private static extern int SHCreateStreamOnFileEx(string file, uint mode, uint attributes, bool create, IStream? template, out IStream stream);
}
