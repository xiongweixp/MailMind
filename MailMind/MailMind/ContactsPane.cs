using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Text;
using Microsoft.Data.Sqlite;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;

namespace MailMind;

public sealed class ContactsPane : Grid
{
    public sealed record Contact(long Id, string Name, string Email, string Phone, string Company, string Department, string Title, string Notes);
    private readonly string connection;
    private readonly ListView list = new() { HorizontalContentAlignment = HorizontalAlignment.Stretch, SelectionMode = ListViewSelectionMode.Single };
    private readonly TextBox search = new() { PlaceholderText = Strings.Get("Contact_42b013f6a361") };
    private readonly StackPanel detail = new() { Spacing = 12, Padding = new Thickness(12, 0, 18, 18) };
    private readonly TextBlock count = new() { Margin = new Thickness(12), Opacity = .65 };
    private List<Contact> contacts = new();
    private bool editing;
    private readonly ColumnDefinition contactListColumn = new() { Width = new GridLength(340), MinWidth = 220 };
    private bool resizingContacts;
    private uint resizePointerId;
    private double resizeStartX, resizeStartWidth;
    public event Action? ContactListWidthChanged;
    public double ContactListWidth { get => contactListColumn.Width.Value; set => contactListColumn.Width = new GridLength(Math.Clamp(value, 220, 800)); }
    private static readonly int[] InitialStarts = { -20319, -20284, -19776, -19219, -18711, -18527, -18240, -17923, -17418, -16475, -16213, -15641, -15166, -14923, -14915, -14631, -14150, -14091, -13319, -12839, -12557, -11848, -11056 };
    private static readonly Encoding PinyinEncoding = CreatePinyinEncoding();
    public event Action<string>? ComposeRequested;
    public event Action<RecentMail>? MailRequested;
    public sealed record RecentMail(string Account, string Folder, uint Validity, uint Uid, string Subject, string Date);

    public ContactsPane()
    {
        var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailMind");
        Directory.CreateDirectory(dir);
        connection = new SqliteConnectionStringBuilder { DataSource = Path.Combine(dir, "mailmind.db") }.ToString();
        using (var db = Open()) { using var cmd = db.CreateCommand(); cmd.CommandText = "CREATE TABLE IF NOT EXISTS contacts (id INTEGER PRIMARY KEY, name TEXT NOT NULL, email TEXT NOT NULL COLLATE NOCASE UNIQUE, phone TEXT NOT NULL, company TEXT NOT NULL, department TEXT NOT NULL, title TEXT NOT NULL, notes TEXT NOT NULL)"; cmd.ExecuteNonQuery(); }

        ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(340) });
        ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        var left = new Grid { Padding = new Thickness(18, 0, 12, 12), RowSpacing = 12 };
        var shell = new Grid { RowSpacing = 0 };
        shell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        shell.RowDefinitions.Add(new RowDefinition());
        SetColumnSpan(shell, 2); Children.Add(shell);
        var top = new Grid { Padding = new Thickness(18, 18, 18, 18), ColumnSpacing = 16 };
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        top.Children.Add(new TextBlock { Text = Strings.Get("Contact_9070cb0eb695"), FontSize = 28, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, VerticalAlignment = VerticalAlignment.Center });
        var add = new Button { Content = Strings.Get("Contact_163297d713b7"), Background = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 40, 99, 207)), Foreground = new SolidColorBrush(Microsoft.UI.Colors.White), Padding = new Thickness(16, 8, 16, 8) };
        add.Click += (_, _) => { if (!editing) Edit(null); };
        var import = new Button { Content = Strings.Get("Contact_77912c43b1e9"), Padding = new Thickness(12, 8, 12, 8) };
        import.Click += async (_, _) =>
        {
            if (editing) return;
            try
            {
                int imported = ImportCachedSenders();
                Refresh();
                await new ContentDialog { XamlRoot = XamlRoot, Title = Strings.Get("Contact_75c33f5c2ecc"), Content = string.Format(Strings.Get("Contact_42635ed1f6f1"), imported), CloseButtonText = Strings.Get("Contact_cb63c62e50f8") }.ShowAsync();
            }
            catch (Exception ex)
            {
                await new ContentDialog { XamlRoot = XamlRoot, Title = Strings.Get("Contact_a01a8d53938f"), Content = ex.Message, CloseButtonText = Strings.Get("Contact_cb63c62e50f8") }.ShowAsync();
            }
        };
        var topActions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 }; topActions.Children.Add(add); topActions.Children.Add(import);
        SetColumn(topActions, 1); top.Children.Add(topActions); shell.Children.Add(top);
        var columns = new Grid(); columns.ColumnDefinitions.Add(contactListColumn); columns.ColumnDefinitions.Add(new ColumnDefinition());
        SetRow(columns, 1); shell.Children.Add(columns);
        left.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); left.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); left.RowDefinitions.Add(new RowDefinition()); left.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        search.TextChanged += (_, _) => Filter(); left.Children.Add(search);
        var filterLabel = new TextBlock { Text = Strings.Get("Contact_93362321e1bb"), FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, Margin = new Thickness(4, 0, 0, 0) }; SetRow(filterLabel, 1); left.Children.Add(filterLabel);
        list.SelectionChanged += (_, _) => ShowDetail(); SetRow(list, 2); left.Children.Add(list); SetRow(count, 3); left.Children.Add(count); columns.Children.Add(left);
        var scroll = new ScrollViewer { Content = detail, VerticalScrollBarVisibility = ScrollBarVisibility.Auto }; SetColumn(scroll, 1); columns.Children.Add(scroll);
        var splitter = new ColumnSplitter { Width = 8, HorizontalAlignment = HorizontalAlignment.Right };
        ToolTipService.SetToolTip(splitter, Strings.Get("ResizeColumns"));
        splitter.PointerPressed += (_, e) =>
        {
            if (resizingContacts || !e.GetCurrentPoint(splitter).Properties.IsLeftButtonPressed || !splitter.CapturePointer(e.Pointer)) return;
            resizingContacts = true; resizePointerId = e.Pointer.PointerId;
            resizeStartX = e.GetCurrentPoint(columns).Position.X; resizeStartWidth = contactListColumn.ActualWidth;
            e.Handled = true;
        };
        splitter.PointerMoved += (_, e) =>
        {
            if (!resizingContacts || e.Pointer.PointerId != resizePointerId) return;
            double maxWidth = Math.Max(220, Math.Min(800, columns.ActualWidth - 300));
            contactListColumn.Width = new GridLength(Math.Clamp(resizeStartWidth + e.GetCurrentPoint(columns).Position.X - resizeStartX, 220, maxWidth));
            e.Handled = true;
        };
        splitter.PointerReleased += (_, e) =>
        {
            if (!resizingContacts || e.Pointer.PointerId != resizePointerId) return;
            resizingContacts = false; splitter.ReleasePointerCapture(e.Pointer); ContactListWidthChanged?.Invoke(); e.Handled = true;
        };
        splitter.PointerCaptureLost += (_, e) =>
        {
            if (!resizingContacts || e.Pointer.PointerId != resizePointerId) return;
            resizingContacts = false; ContactListWidthChanged?.Invoke();
        };
        columns.Children.Add(splitter);
        columns.SizeChanged += (_, _) => contactListColumn.MaxWidth = Math.Max(220, Math.Min(800, columns.ActualWidth - 300));
        Refresh();
    }
    private SqliteConnection Open() { var db = new SqliteConnection(connection); db.Open(); return db; }
    public (List<Contact> Matches, bool Truncated) SearchContacts(string query, int limit = 20)
    {
        query = query.Trim();
        if (query.Length == 0) throw new ArgumentException(Strings.Get("Contact_d33caa10efe3"), nameof(query));
        var matches = new List<Contact>();
        using var db = Open(); using var cmd = db.CreateCommand();
        cmd.CommandText = "SELECT id,name,email,phone,company,department,title,notes FROM contacts ORDER BY name COLLATE NOCASE";
        using var rows = cmd.ExecuteReader();
        while (rows.Read())
        {
            var name = rows.GetString(1); var email = rows.GetString(2);
            if (!name.Contains(query, StringComparison.OrdinalIgnoreCase) && !email.Contains(query, StringComparison.OrdinalIgnoreCase)) continue;
            if (matches.Count == limit) return (matches, true);
            matches.Add(new Contact(rows.GetInt64(0), name, email, rows.GetString(3), rows.GetString(4), rows.GetString(5), rows.GetString(6), rows.GetString(7)));
        }
        return (matches, false);
    }
    public List<Contact> SearchRecipients(string query, int limit = 12)
    {
        query = query.Trim();
        if (query.Length == 0) return new();
        using var db = Open(); using var cmd = db.CreateCommand();
        cmd.CommandText = "SELECT id,name,email,phone,company,department,title,notes FROM contacts ORDER BY name COLLATE NOCASE";
        using var rows = cmd.ExecuteReader(); var matches = new List<Contact>();
        while (rows.Read())
        {
            string name = rows.GetString(1), email = rows.GetString(2);
            string pinyin = PinyinSearch.Transliterate(name);
            string initials = string.Concat(name.Select(c => PinyinSearch.Transliterate(c.ToString()).FirstOrDefault()).Where(c => c != '\0'));
            if (!name.Contains(query, StringComparison.OrdinalIgnoreCase) && !email.Contains(query, StringComparison.OrdinalIgnoreCase) && !pinyin.Contains(query, StringComparison.OrdinalIgnoreCase) && !initials.Contains(query, StringComparison.OrdinalIgnoreCase)) continue;
            matches.Add(new Contact(rows.GetInt64(0), name, email, rows.GetString(3), rows.GetString(4), rows.GetString(5), rows.GetString(6), rows.GetString(7)));
            if (matches.Count == limit) break;
        }
        return matches;
    }
    private int ImportCachedSenders()
    {
        using var db = Open();
        using var check = db.CreateCommand();
        check.CommandText = "SELECT 1 FROM sqlite_master WHERE type='table' AND name='messages'";
        if (check.ExecuteScalar() == null) return 0;
        var senders = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var accounts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using (var read = db.CreateCommand())
        {
            read.CommandText = "SELECT account,sender FROM messages";
            using var rows = read.ExecuteReader();
            while (rows.Read())
            {
                accounts.Add(rows.GetString(0).Trim());
                if (!MimeKit.InternetAddressList.TryParse(rows.GetString(1), out var addresses)) continue;
                foreach (var address in addresses.Mailboxes)
                {
                    var email = address.Address.Trim();
                    if (string.IsNullOrWhiteSpace(email)) continue;
                    var name = string.IsNullOrWhiteSpace(address.Name) ? email : address.Name.Trim();
                    if (!senders.TryGetValue(email, out var existing) || (existing == email && name != email)) senders[email] = name;
                }
            }
        }
        using var transaction = db.BeginTransaction();
        using var insert = db.CreateCommand(); insert.Transaction = transaction;
        insert.CommandText = "INSERT OR IGNORE INTO contacts(name,email,phone,company,department,title,notes) VALUES($name,$email,'','','','','')";
        var nameParameter = insert.Parameters.Add("$name", SqliteType.Text);
        var emailParameter = insert.Parameters.Add("$email", SqliteType.Text);
        int count = 0;
        foreach (var (email, name) in senders)
        {
            if (accounts.Contains(email)) continue;
            nameParameter.Value = name; emailParameter.Value = email;
            count += insert.ExecuteNonQuery();
        }
        transaction.Commit();
        return count;
    }
    public void Refresh()
    {
        if (editing) return;
        long? selected = SelectedContact()?.Id;
        using var db = Open(); using var cmd = db.CreateCommand(); cmd.CommandText = "SELECT id,name,email,phone,company,department,title,notes FROM contacts ORDER BY name COLLATE NOCASE";
        using var r = cmd.ExecuteReader(); contacts = new(); while (r.Read()) contacts.Add(new(r.GetInt64(0), r.GetString(1), r.GetString(2), r.GetString(3), r.GetString(4), r.GetString(5), r.GetString(6), r.GetString(7)));
        Filter(); if (selected.HasValue) list.SelectedItem = list.Items.OfType<ListViewItem>().FirstOrDefault(item => (item.Tag as Contact)?.Id == selected);
    }
    private void Filter()
    {
        var query = search.Text.Trim();
        var results = contacts.Where(c => c.Name.Contains(query, StringComparison.OrdinalIgnoreCase) || c.Email.Contains(query, StringComparison.OrdinalIgnoreCase))
            .OrderBy(c => Initial(c.Name) == '#' ? '[' : Initial(c.Name))
            .ThenBy(c => c.Name, StringComparer.Create(CultureInfo.GetCultureInfo("zh-CN"), true)).ToList();
        list.Items.Clear();
        char? previous = null;
        foreach (var c in results)
        {
            char initial = Initial(c.Name);
            if (initial != previous)
            {
                list.Items.Add(new ListViewItem { IsEnabled = false, IsHitTestVisible = false, Content = new TextBlock { Text = initial.ToString(), FontSize = 16, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, Margin = new Thickness(2, 8, 0, 0) } });
                previous = initial;
            }
            var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12, Padding = new Thickness(0, 6, 0, 6) };
            row.Children.Add(Avatar(c, 48));
            var text = new StackPanel { VerticalAlignment = VerticalAlignment.Center, Spacing = 3 };
            text.Children.Add(new TextBlock { Text = c.Name, FontSize = 16, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
            text.Children.Add(new TextBlock { Text = Subtitle(c), Opacity = .65, MaxWidth = 210, TextTrimming = TextTrimming.CharacterEllipsis }); row.Children.Add(text);
            list.Items.Add(new ListViewItem { Content = row, Tag = c, HorizontalContentAlignment = HorizontalAlignment.Stretch });
        }
        count.Text = string.Format(Strings.Get("Contact_0464af7ccbfe"), results.Count);
        if (results.Count > 0) list.SelectedItem = list.Items.OfType<ListViewItem>().First(item => item.Tag is Contact); else ShowDetail();
    }
    private static Encoding CreatePinyinEncoding()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding("GB2312", EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
    }
    private static char Initial(string name)
    {
        var trimmed = name.Trim();
        if (trimmed.Length == 0) return '#';
        char first = char.ToUpperInvariant(trimmed[0]);
        if (first >= 'A' && first <= 'Z') return first;
        try
        {
            var encoded = PinyinEncoding.GetBytes(trimmed.Substring(0, 1));
            if (encoded.Length != 2) return '#';
            int code = encoded[0] * 256 + encoded[1] - 65536;
            const string letters = "ABCDEFGHJKLMNOPQRSTWXYZ";
            for (int i = InitialStarts.Length - 1; i >= 0; i--) if (code >= InitialStarts[i] && code <= -10247) return letters[i];
        }
        catch (EncoderFallbackException) { }
        return '#';
    }
    private void ShowDetail()
    {
        detail.Children.Clear();
        if (SelectedContact() is not Contact c) { detail.Children.Add(new TextBlock { Text = Strings.Get("Contact_78fdf87306a8"), Opacity = .6, FontSize = 18, Margin = new Thickness(20) }); return; }
        var actions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        var edit = new Button { Content = Strings.Get("Contact_a7f814c0a40d") }; edit.Click += (_, _) => Edit(c); actions.Children.Add(edit);
        var compose = new Button { Content = Strings.Get("Contact_29558466833b") }; compose.Click += (_, _) => ComposeRequested?.Invoke(c.Email); actions.Children.Add(compose);
        var delete = new Button { Content = Strings.Get("Contact_3755f56f2f83") }; delete.Click += async (_, _) => { var dialog = new ContentDialog { XamlRoot = XamlRoot, Title = Strings.Get("Contact_74cff4c9e39b"), Content = string.Format(Strings.Get("Contact_11e0c9499eda"), c.Name), PrimaryButtonText = Strings.Get("Contact_3755f56f2f83"), CloseButtonText = Strings.Get("Contact_4d0b4688c787") }; if (await dialog.ShowAsync() == ContentDialogResult.Primary) { using var db = Open(); using var cmd = db.CreateCommand(); cmd.CommandText = "DELETE FROM contacts WHERE id=$id"; cmd.Parameters.AddWithValue("$id", c.Id); cmd.ExecuteNonQuery(); Refresh(); } }; actions.Children.Add(delete); detail.Children.Add(actions);
        var identity = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 18 }; identity.Children.Add(Avatar(c, 72));
        var identityText = new StackPanel { Spacing = 5, VerticalAlignment = VerticalAlignment.Center };
        identityText.Children.Add(new TextBlock { Text = c.Name, FontSize = 26, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold }); identityText.Children.Add(new TextBlock { Text = Subtitle(c), Opacity = .65 }); identity.Children.Add(identityText); detail.Children.Add(Card(identity));
        var info = new StackPanel { Spacing = 16 };
        info.Children.Add(new TextBlock { Text = Strings.Get("Contact_b0df5c459f31"), FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        foreach (var (label, value, glyph) in new[] { (Strings.Get("Contact_9ed627bcf63d"), c.Email, "\uE715"), (Strings.Get("Contact_f39c6b629f80"), c.Phone, "\uE717"), (Strings.Get("Contact_5d80806d7f55"), c.Company, "\uE80F"), (Strings.Get("Contact_e0361480e3a5"), c.Notes, "\uE8A5") })
        {
            var row = InfoRow(label, glyph);
            var text = new TextBlock { Text = string.IsNullOrWhiteSpace(value) ? "—" : value, TextWrapping = TextWrapping.Wrap, VerticalAlignment = VerticalAlignment.Center };
            SetColumn(text, 2); row.Children.Add(text); info.Children.Add(row);
        }
        detail.Children.Add(Card(info));
        var recent = RecentMessages(c.Email);
        var history = new StackPanel { Spacing = 10 };
        var historyHeader = new Grid(); historyHeader.ColumnDefinitions.Add(new ColumnDefinition()); historyHeader.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        historyHeader.Children.Add(new TextBlock { Text = Strings.Get("Contact_02b5d33f3897"), FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, VerticalAlignment = VerticalAlignment.Center });
        var showAll = new Button { Content = Strings.Get("Contact_ed2172fd7894"), Visibility = recent.Count > 3 ? Visibility.Visible : Visibility.Collapsed };
        SetColumn(showAll, 1); historyHeader.Children.Add(showAll); history.Children.Add(historyHeader);
        var entries = new StackPanel { Spacing = 4 }; history.Children.Add(entries);
        void RenderMessages(int limit)
        {
            entries.Children.Clear();
            if (recent.Count == 0) entries.Children.Add(new TextBlock { Text = Strings.Get("Contact_a00575679ceb"), Opacity = .65, TextWrapping = TextWrapping.Wrap });
            foreach (var message in recent.Take(limit))
            {
                var row = new Grid { ColumnSpacing = 12 }; row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); row.ColumnDefinitions.Add(new ColumnDefinition()); row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                row.Children.Add(InfoIcon("\uE715"));
                var subject = new TextBlock { Text = message.Subject, TextTrimming = TextTrimming.CharacterEllipsis, VerticalAlignment = VerticalAlignment.Center }; SetColumn(subject, 1); row.Children.Add(subject);
                var date = new TextBlock { Text = message.Date, Opacity = .6, VerticalAlignment = VerticalAlignment.Center }; SetColumn(date, 2); row.Children.Add(date);
                var button = new Button { Content = row, HorizontalAlignment = HorizontalAlignment.Stretch, HorizontalContentAlignment = HorizontalAlignment.Stretch, Background = new SolidColorBrush(Microsoft.UI.Colors.Transparent), BorderThickness = new Thickness(0), Padding = new Thickness(0, 8, 0, 8) };
                button.Click += (_, _) => MailRequested?.Invoke(message); entries.Children.Add(button);
            }
        }
        showAll.Click += (_, _) => { RenderMessages(recent.Count); showAll.Visibility = Visibility.Collapsed; };
        RenderMessages(3); detail.Children.Add(Card(history));
    }
    private List<RecentMail> RecentMessages(string email)
    {
        var result = new List<RecentMail>();
        using var db = Open(); using var check = db.CreateCommand();
        check.CommandText = "SELECT 1 FROM sqlite_master WHERE type='table' AND name='messages'";
        if (check.ExecuteScalar() == null) return result;
        using var cmd = db.CreateCommand();
        cmd.CommandText = "SELECT m.account,m.folder,m.validity,m.uid,m.sender,m.subject,m.date FROM messages m JOIN folders f ON f.account=m.account AND f.path=m.folder AND f.validity=m.validity ORDER BY m.date DESC, m.uid DESC";
        using var rows = cmd.ExecuteReader();
        while (rows.Read())
        {
            if (!MimeKit.InternetAddressList.TryParse(rows.GetString(4), out var addresses) || !addresses.Mailboxes.Any(a => string.Equals(a.Address, email, StringComparison.OrdinalIgnoreCase))) continue;
            result.Add(new RecentMail(rows.GetString(0), rows.GetString(1), (uint)rows.GetInt64(2), (uint)rows.GetInt64(3), rows.GetString(5), rows.GetString(6)));
        }
        return result;
    }
    private Contact? SelectedContact() => (list.SelectedItem as ListViewItem)?.Tag as Contact;
    private static FontIcon InfoIcon(string glyph) => new() { Glyph = glyph, FontSize = 18, VerticalAlignment = VerticalAlignment.Center };
    private static Grid InfoRow(string label, string glyph)
    {
        var row = new Grid { ColumnSpacing = 12 };
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(44) });
        row.ColumnDefinitions.Add(new ColumnDefinition());
        row.Children.Add(InfoIcon(glyph));
        var text = new TextBlock { Text = label, VerticalAlignment = VerticalAlignment.Center }; SetColumn(text, 1); row.Children.Add(text);
        return row;
    }
    private static string Subtitle(Contact c) { var parts = new[] { c.Department, c.Title }.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray(); return parts.Length == 0 ? c.Email : string.Join(" · ", parts); }
    private static Border Avatar(Contact c, double size)
    {
        byte[][] palette = { new byte[] { 143, 161, 230 }, new byte[] { 143, 199, 188 }, new byte[] { 202, 137, 146 }, new byte[] { 232, 164, 83 }, new byte[] { 165, 143, 219 } };
        int hash = 0; foreach (char ch in c.Email) hash = (hash * 31 + ch) & int.MaxValue;
        var color = palette[hash % palette.Length];
        char initial = Initial(c.Name);
        return new Border { Width = size, Height = size, CornerRadius = new CornerRadius(8), Background = new SolidColorBrush(Windows.UI.Color.FromArgb(255, color[0], color[1], color[2])), Child = new TextBlock { Text = initial == '#' ? (c.Name.Length > 0 ? c.Name.Substring(0, 1) : "?") : initial.ToString(), FontSize = size * .4, Foreground = new SolidColorBrush(Microsoft.UI.Colors.White), HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center } };
    }
    private static Border Card(UIElement child) => new() { Child = child, Padding = new Thickness(20), CornerRadius = new CornerRadius(8), BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(Windows.UI.Color.FromArgb(100, 150, 160, 175)) };
    private void Edit(Contact? c)
    {
        editing = true;
        list.IsEnabled = false;
        search.IsEnabled = false;
        detail.Children.Clear();
        var fields = new[] { new TextBox { PlaceholderText = Strings.Get("Contact_be4c2616b177"), Text = c?.Name ?? "", FontSize = 20 }, new TextBox { PlaceholderText = Strings.Get("Contact_9ed627bcf63d"), Text = c?.Email ?? "" }, new TextBox { PlaceholderText = Strings.Get("Contact_f39c6b629f80"), Text = c?.Phone ?? "" }, new TextBox { PlaceholderText = Strings.Get("Contact_5d80806d7f55"), Text = c?.Company ?? "" }, new TextBox { PlaceholderText = Strings.Get("Contact_91061a56c00f"), Text = c?.Department ?? "" }, new TextBox { PlaceholderText = Strings.Get("Contact_a2a92e50a114"), Text = c?.Title ?? "" }, new TextBox { PlaceholderText = Strings.Get("Contact_e0361480e3a5"), Text = c?.Notes ?? "", AcceptsReturn = true } };
        var actions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        var save = new Button { Content = Strings.Get("Contact_fadf24dbc5a9"), Background = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 40, 99, 207)), Foreground = new SolidColorBrush(Microsoft.UI.Colors.White) };
        var cancel = new Button { Content = Strings.Get("Contact_4d0b4688c787") };
        actions.Children.Add(save); actions.Children.Add(cancel); detail.Children.Add(actions);
        var identity = new Grid { ColumnSpacing = 18 };
        identity.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        identity.ColumnDefinitions.Add(new ColumnDefinition());
        identity.Children.Add(Avatar(c ?? new Contact(0, "", "", "", "", "", "", ""), 72));
        var identityFields = new StackPanel { Spacing = 6 };
        identityFields.Children.Add(fields[0]);
        var roleFields = new Grid { ColumnSpacing = 6 };
        roleFields.ColumnDefinitions.Add(new ColumnDefinition()); roleFields.ColumnDefinitions.Add(new ColumnDefinition());
        roleFields.Children.Add(fields[4]); SetColumn(fields[5], 1); roleFields.Children.Add(fields[5]);
        identityFields.Children.Add(roleFields); SetColumn(identityFields, 1); identity.Children.Add(identityFields);
        detail.Children.Add(Card(identity));
        var info = new StackPanel { Spacing = 16 };
        info.Children.Add(new TextBlock { Text = Strings.Get("Contact_b0df5c459f31"), FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        foreach (var (label, index, glyph) in new[] { (Strings.Get("Contact_9ed627bcf63d"), 1, "\uE715"), (Strings.Get("Contact_f39c6b629f80"), 2, "\uE717"), (Strings.Get("Contact_5d80806d7f55"), 3, "\uE80F"), (Strings.Get("Contact_e0361480e3a5"), 6, "\uE8A5") })
        {
            var row = InfoRow(label, glyph);
            SetColumn(fields[index], 2); row.Children.Add(fields[index]); info.Children.Add(row);
        }
        var error = new TextBlock { TextWrapping = TextWrapping.Wrap, Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 190, 45, 45)) };
        info.Children.Add(error); detail.Children.Add(Card(info));
        void EndEdit() { editing = false; list.IsEnabled = true; search.IsEnabled = true; }
        cancel.Click += (_, _) => { EndEdit(); ShowDetail(); };
        save.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(fields[0].Text) || !MimeKit.MailboxAddress.TryParse(fields[1].Text.Trim(), out var parsed) || parsed.Address != fields[1].Text.Trim()) { error.Text = Strings.Get("Contact_3b303cde8823"); return; }
            if (contacts.Any(x => x.Id != c?.Id && string.Equals(x.Email, fields[1].Text.Trim(), StringComparison.OrdinalIgnoreCase))) { error.Text = Strings.Get("Contact_d2b8c0e173d6"); return; }
            try
            {
                using var db = Open(); using var cmd = db.CreateCommand(); cmd.CommandText = c == null ? "INSERT INTO contacts(name,email,phone,company,department,title,notes) VALUES($n,$e,$p,$c,$d,$t,$o)" : "UPDATE contacts SET name=$n,email=$e,phone=$p,company=$c,department=$d,title=$t,notes=$o WHERE id=$id";
                string[] keys = { "$n", "$e", "$p", "$c", "$d", "$t", "$o" }; for (int i = 0; i < keys.Length; i++) cmd.Parameters.AddWithValue(keys[i], fields[i].Text.Trim()); if (c != null) cmd.Parameters.AddWithValue("$id", c.Id); cmd.ExecuteNonQuery();
                var email = fields[1].Text.Trim(); EndEdit(); Refresh(); list.SelectedItem = list.Items.OfType<ListViewItem>().FirstOrDefault(item => string.Equals((item.Tag as Contact)?.Email, email, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex) { error.Text = Strings.Get("Contact_b12163000cd3") + ex.Message; }
        };
        fields[0].Focus(FocusState.Programmatic);
    }
}
