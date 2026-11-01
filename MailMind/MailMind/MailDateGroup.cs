using System;
using System.Collections.ObjectModel;
using System.Globalization;

namespace MailMind;

public sealed class MailDateGroup : ObservableCollection<MailItem>
{
    public int Order { get; }
    public string Title { get; }
    public MailDateGroup(int order, string title) { Order = order; Title = title; }

    public static (int Order, string Title) Classify(string value, DateTime today)
    {
        if (!DateTime.TryParseExact(value, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)) return (4, "日期未知");
        var days = (today.Date - date.Date).Days;
        if (days <= 0) return (0, "今天");
        if (days < 7) return (1, "一天前");
        if (days < 14) return (2, "一周前");
        return (3, "更早前");
    }
}
