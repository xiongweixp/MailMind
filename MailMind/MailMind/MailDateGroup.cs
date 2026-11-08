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
        if (!DateTime.TryParseExact(value, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)) return (4, Strings.Get("DateUnknown"));
        var days = (today.Date - date.Date).Days;
        if (days <= 0) return (0, Strings.Get("Today"));
        if (days < 7) return (1, Strings.Get("DaysAgo"));
        if (days < 14) return (2, Strings.Get("WeekAgo"));
        return (3, Strings.Get("Older"));
    }
}
