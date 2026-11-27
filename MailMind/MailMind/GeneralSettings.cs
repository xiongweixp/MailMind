using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Win32;
using Windows.ApplicationModel;

namespace MailMind;

internal sealed record GeneralSettings(string Language = "system", bool CloseToTray = false)
{
    private static readonly string PathName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailMind", "general.json");
    public static GeneralSettings Load()
    {
        try { return File.Exists(PathName) ? JsonSerializer.Deserialize<GeneralSettings>(File.ReadAllText(PathName)) ?? new() : new(); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException) { return new(); }
    }
    public void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(PathName)!);
        File.WriteAllText(PathName, JsonSerializer.Serialize(this));
    }
    private const string RunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string StartupTaskId = "MailMindStartup";
    private static bool HasPackageIdentity()
    {
        try { _ = Package.Current; return true; }
        catch (InvalidOperationException) { return false; }
    }
    public static async Task<bool> StartupEnabledAsync()
    {
        if (HasPackageIdentity())
            return (await StartupTask.GetAsync(StartupTaskId)).State == StartupTaskState.Enabled;
        using var key = Registry.CurrentUser.OpenSubKey(RunKey);
        return key?.GetValue("MailMind") is string value && !string.IsNullOrEmpty(value);
    }
    public static async Task SetStartupAsync(bool enabled)
    {
        if (HasPackageIdentity())
        {
            var task = await StartupTask.GetAsync(StartupTaskId);
            if (!enabled) { task.Disable(); return; }
            if (task.State != StartupTaskState.Enabled && await task.RequestEnableAsync() != StartupTaskState.Enabled)
                throw new InvalidOperationException("无法启用开机启动。请检查 Windows 设置中的启动应用权限。");
            return;
        }
        using var key = Registry.CurrentUser.CreateSubKey(RunKey);
        if (enabled) key.SetValue("MailMind", "\"" + Path.Combine(AppContext.BaseDirectory, "MailMind.exe") + "\" --startup-hidden");
        else key.DeleteValue("MailMind", false);
    }
}
