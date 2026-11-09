using System;
using System.Runtime.InteropServices;

namespace MailMind;

internal sealed class TrayIcon : IDisposable
{
    public const uint CallbackMessage = 0x8001;
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NotifyData
    {
        public uint Size; public nint Window; public uint Id, Flags, Callback; public nint Icon;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)] public string Tip;
        public uint State, StateMask;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string Info;
        public uint Timeout;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)] public string InfoTitle;
        public uint InfoFlags; public Guid Guid; public nint BalloonIcon;
    }
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)] private static extern bool Shell_NotifyIcon(uint message, ref NotifyData data);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern nint LoadImage(nint instance, string name, uint type, int width, int height, uint flags);
    [DllImport("user32.dll")] private static extern bool DestroyIcon(nint icon);
    [DllImport("user32.dll")] private static extern nint CreatePopupMenu();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern bool AppendMenu(nint menu, uint flags, nuint id, string text);
    [DllImport("user32.dll")] private static extern bool DestroyMenu(nint menu);
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(nint window);
    [DllImport("user32.dll")] private static extern uint TrackPopupMenu(nint menu, uint flags, int x, int y, int reserved, nint window, nint rect);
    [DllImport("user32.dll")] private static extern bool GetCursorPos(out Point point);
    [StructLayout(LayoutKind.Sequential)] private struct Point { public int X, Y; }
    private NotifyData data;
    public TrayIcon(nint window)
    {
        data = new NotifyData { Size = (uint)Marshal.SizeOf<NotifyData>(), Window = window, Id = 1, Flags = 7, Callback = CallbackMessage, Icon = LoadImage(0, System.IO.Path.Combine(AppContext.BaseDirectory, "Assets", "TrayNormal.ico"), 1, 0, 0, 0x10), Tip = "MailMind", Info = "", InfoTitle = "" };
        if (data.Icon == 0 || !Shell_NotifyIcon(0, ref data)) { if (data.Icon != 0) DestroyIcon(data.Icon); throw new InvalidOperationException("Could not create the system tray icon."); }
    }
    public void SetUnread(bool unread)
    {
        var icon = LoadImage(0, System.IO.Path.Combine(AppContext.BaseDirectory, "Assets", unread ? "TrayImportant.ico" : "TrayNormal.ico"), 1, 0, 0, 0x10);
        if (icon == 0) return;
        var old = data.Icon; data.Icon = icon; data.Flags = 2;
        if (Shell_NotifyIcon(1, ref data)) DestroyIcon(old);
        else { data.Icon = old; DestroyIcon(icon); }
    }
    public void Notify(string title, string summary)
    {
        data.Flags = 0x10; data.InfoTitle = title.Length > 63 ? title[..63] : title;
        data.Info = summary.Length > 255 ? summary[..255] : summary; data.InfoFlags = 1;
        Shell_NotifyIcon(1, ref data);
    }
    public int ShowMenu()
    {
        var menu = CreatePopupMenu();
        if (menu == 0) return 0;
        try
        {
            AppendMenu(menu, 0, 1, Strings.Get("TrayOpen")); AppendMenu(menu, 0, 2, Strings.Get("TrayExit"));
            GetCursorPos(out var point); SetForegroundWindow(data.Window);
            return (int)TrackPopupMenu(menu, 0x0100 | 0x0002, point.X, point.Y, 0, data.Window, 0);
        }
        finally { DestroyMenu(menu); }
    }
    public void Dispose() { Shell_NotifyIcon(2, ref data); if (data.Icon != 0) { DestroyIcon(data.Icon); data.Icon = 0; } }
}
