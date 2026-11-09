using System;
using System.IO;
using System.Text.Json;
using Microsoft.Win32;

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
    public static bool StartupEnabled()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RunKey);
        return key?.GetValue("MailMind") is string value && !string.IsNullOrEmpty(value);
    }
    public static void SetStartup(bool enabled)
    {
        using var key = Registry.CurrentUser.CreateSubKey(RunKey);
        if (enabled) key.SetValue("MailMind", "\"" + Path.Combine(AppContext.BaseDirectory, "MailMind.exe") + "\"");
        else key.DeleteValue("MailMind", false);
    }
}
