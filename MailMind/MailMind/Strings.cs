using System;
using System.Globalization;
using System.Resources;
using System.Runtime.InteropServices;

namespace MailMind;

internal static class Strings
{
    private static readonly ResourceManager Resources = new("MailMind.Resources.Strings", typeof(Strings).Assembly);

    [DllImport("kernel32.dll")]
    private static extern ushort GetUserDefaultUILanguage();

    private static readonly CultureInfo Language = SelectConfiguredLanguage();

    private static CultureInfo SelectConfiguredLanguage()
    {
        var configured = GeneralSettings.Load().Language;
        return configured is "zh-CN" or "zh-TW" or "en-US" ? CultureInfo.GetCultureInfo(configured) : SelectLanguage(CultureInfo.GetCultureInfo(GetUserDefaultUILanguage()));
    }

    public static string LanguageTag => Language.Name;

    private static CultureInfo SelectLanguage(CultureInfo culture)
    {
        if (culture.TwoLetterISOLanguageName != "zh") return CultureInfo.GetCultureInfo("en-US");

        // Windows uses Hant for traditional Chinese, including Hong Kong and Macau.
        for (var current = culture; !current.Equals(CultureInfo.InvariantCulture); current = current.Parent)
            if (current.Name.Equals("zh-Hant", StringComparison.OrdinalIgnoreCase) ||
                current.Name.Equals("zh-TW", StringComparison.OrdinalIgnoreCase) ||
                current.Name.Equals("zh-HK", StringComparison.OrdinalIgnoreCase) ||
                current.Name.Equals("zh-MO", StringComparison.OrdinalIgnoreCase))
                return CultureInfo.GetCultureInfo("zh-TW");

        return CultureInfo.GetCultureInfo("zh-CN");
    }

    public static string Get(string key) => Resources.GetString(key, Language) ?? key;

    public static string FolderName(string name) => name switch
    {
        "收件箱" => Get("FolderInbox"),
        "已发送" => Get("FolderSent"),
        "草稿箱" => Get("FolderDrafts"),
        "已删除" => Get("FolderDeleted"),
        "垃圾邮件" => Get("FolderJunk"),
        "归档" => Get("FolderArchive"),
        _ => name
    };

    public static string Importance(string value) => value switch
    {
        "重要" => Get("ImportanceHigh"),
        "一般" => Get("ImportanceNormal"),
        "可忽略" => Get("ImportanceLow"),
        "失败" => Get("ImportanceFailed"),
        _ => value
    };
}
