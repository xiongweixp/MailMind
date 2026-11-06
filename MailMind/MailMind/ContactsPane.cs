using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Data.Sqlite;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;

namespace MailMind;

public sealed class ContactsPane : Grid
{
    public sealed record Contact(long Id, string Name, string Email, string Phone, string Company, string Department, string Title, string Notes);
    private readonly string connection;
    private readonly ListView list = new() { HorizontalContentAlignment = HorizontalAlignment.Stretch, SelectionMode = ListViewSelectionMode.Single };
    private readonly TextBox search = new() { PlaceholderText = "搜索姓名或邮箱" };
    private readonly StackPanel detail = new() { Spacing = 12, Padding = new Thickness(12, 0, 18, 18) };
    private readonly TextBlock count = new() { Margin = new Thickness(12), Opacity = .65 };
    private List<Contact> contacts = new();
    private bool editing;
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
        top.Children.Add(new TextBlock { Text = "通讯录", FontSize = 28, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, VerticalAlignment = VerticalAlignment.Center });
        var add = new Button { Content = "+  新增联系人", Background = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 40, 99, 207)), Foreground = new SolidColorBrush(Microsoft.UI.Colors.White), Padding = new Thickness(16, 8, 16, 8) };
        add.Click += (_, _) => { if (!editing) Edit(null); };
        var import = new Button { Content = "从邮件导入", Padding = new Thickness(12, 8, 12, 8) };
        import.Click += async (_, _) =>
        {
            if (editing) return;
            try
            {
                int imported = ImportCachedSenders();
                Refresh();
                await new ContentDialog { XamlRoot = XamlRoot, Title = "导入完成", Content = $"已从本地邮件缓存新增 {imported} 位联系人。已有联系人资料保持不变。", CloseButtonText = "知道了" }.ShowAsync();
            }
            catch (Exception ex)
            {
                await new ContentDialog { XamlRoot = XamlRoot, Title = "导入失败", Content = ex.Message, CloseButtonText = "知道了" }.ShowAsync();
            }
        };
        var topActions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 }; topActions.Children.Add(add); topActions.Children.Add(import);
        SetColumn(topActions, 1); top.Children.Add(topActions); shell.Children.Add(top);
        var columns = new Grid(); columns.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(340) }); columns.ColumnDefinitions.Add(new ColumnDefinition());
        SetRow(columns, 1); shell.Children.Add(columns);
        left.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); left.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); left.RowDefinitions.Add(new RowDefinition()); left.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        search.TextChanged += (_, _) => Filter(); left.Children.Add(search);
        var filterLabel = new TextBlock { Text = "全部联系人", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, Margin = new Thickness(4, 0, 0, 0) }; SetRow(filterLabel, 1); left.Children.Add(filterLabel);
        list.SelectionChanged += (_, _) => ShowDetail(); SetRow(list, 2); left.Children.Add(list); SetRow(count, 3); left.Children.Add(count); columns.Children.Add(left);
        var scroll = new ScrollViewer { Content = detail, VerticalScrollBarVisibility = ScrollBarVisibility.Auto }; SetColumn(scroll, 1); columns.Children.Add(scroll);
        Refresh();
    }
    private SqliteConnection Open() { var db = new SqliteConnection(connection); db.Open(); return db; }
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
        var query = search.Text.Trim(); var results = contacts.Where(c => c.Name.Contains(query, StringComparison.OrdinalIgnoreCase) || c.Email.Contains(query, StringComparison.OrdinalIgnoreCase)).ToList();
        list.Items.Clear();
        foreach (var c in results)
        {
            var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12, Padding = new Thickness(0, 6, 0, 6) };
            row.Children.Add(Avatar(c, 48));
            var text = new StackPanel { VerticalAlignment = VerticalAlignment.Center, Spacing = 3 };
            text.Children.Add(new TextBlock { Text = c.Name, FontSize = 16, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
            text.Children.Add(new TextBlock { Text = Subtitle(c), Opacity = .65, MaxWidth = 210, TextTrimming = TextTrimming.CharacterEllipsis }); row.Children.Add(text);
            list.Items.Add(new ListViewItem { Content = row, Tag = c, HorizontalContentAlignment = HorizontalAlignment.Stretch });
        }
        count.Text = $"共 {results.Count} 位联系人";
        if (results.Count > 0) list.SelectedIndex = 0; else ShowDetail();
    }
    private void ShowDetail()
    {
        detail.Children.Clear();
        if (SelectedContact() is not Contact c) { detail.Children.Add(new TextBlock { Text = "选择联系人查看详情", Opacity = .6, FontSize = 18, Margin = new Thickness(20) }); return; }
        var actions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        var edit = new Button { Content = "编辑" }; edit.Click += (_, _) => Edit(c); actions.Children.Add(edit);
        var compose = new Button { Content = "发邮件" }; compose.Click += (_, _) => ComposeRequested?.Invoke(c.Email); actions.Children.Add(compose);
        var delete = new Button { Content = "删除" }; delete.Click += async (_, _) => { var dialog = new ContentDialog { XamlRoot = XamlRoot, Title = "删除联系人？", Content = $"确定删除 {c.Name}？", PrimaryButtonText = "删除", CloseButtonText = "取消" }; if (await dialog.ShowAsync() == ContentDialogResult.Primary) { using var db = Open(); using var cmd = db.CreateCommand(); cmd.CommandText = "DELETE FROM contacts WHERE id=$id"; cmd.Parameters.AddWithValue("$id", c.Id); cmd.ExecuteNonQuery(); Refresh(); } }; actions.Children.Add(delete); detail.Children.Add(actions);
        var identity = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 18 }; identity.Children.Add(Avatar(c, 72));
        var identityText = new StackPanel { Spacing = 5, VerticalAlignment = VerticalAlignment.Center };
        identityText.Children.Add(new TextBlock { Text = c.Name, FontSize = 26, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold }); identityText.Children.Add(new TextBlock { Text = Subtitle(c), Opacity = .65 }); identity.Children.Add(identityText); detail.Children.Add(Card(identity));
        var info = new StackPanel { Spacing = 16 };
        info.Children.Add(new TextBlock { Text = "联系信息", FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        foreach (var (label, value) in new[] { ("邮箱", c.Email), ("电话", c.Phone), ("公司", c.Company), ("备注", c.Notes) }) info.Children.Add(new TextBlock { Text = $"{label}       {(string.IsNullOrWhiteSpace(value) ? "—" : value)}", TextWrapping = TextWrapping.Wrap });
        detail.Children.Add(Card(info));
        var recent = RecentMessages(c.Email);
        var history = new StackPanel { Spacing = 10 };
        var historyHeader = new Grid(); historyHeader.ColumnDefinitions.Add(new ColumnDefinition()); historyHeader.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        historyHeader.Children.Add(new TextBlock { Text = "最近往来", FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, VerticalAlignment = VerticalAlignment.Center });
        var showAll = new Button { Content = "查看全部", Visibility = recent.Count > 3 ? Visibility.Visible : Visibility.Collapsed };
        SetColumn(showAll, 1); historyHeader.Children.Add(showAll); history.Children.Add(historyHeader);
        var entries = new StackPanel { Spacing = 4 }; history.Children.Add(entries);
        void RenderMessages(int limit)
        {
            entries.Children.Clear();
            if (recent.Count == 0) entries.Children.Add(new TextBlock { Text = "本地缓存中暂无该联系人的来信", Opacity = .65, TextWrapping = TextWrapping.Wrap });
            foreach (var message in recent.Take(limit))
            {
                var row = new Grid { ColumnSpacing = 12 }; row.ColumnDefinitions.Add(new ColumnDefinition()); row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                row.Children.Add(new TextBlock { Text = "✉  " + message.Subject, TextTrimming = TextTrimming.CharacterEllipsis, VerticalAlignment = VerticalAlignment.Center });
                var date = new TextBlock { Text = message.Date, Opacity = .6, VerticalAlignment = VerticalAlignment.Center }; SetColumn(date, 1); row.Children.Add(date);
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
    private static string Subtitle(Contact c) { var parts = new[] { c.Department, c.Title }.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray(); return parts.Length == 0 ? c.Email : string.Join(" · ", parts); }
    private static Border Avatar(Contact c, double size)
    {
        byte[][] palette = { new byte[] { 143, 161, 230 }, new byte[] { 143, 199, 188 }, new byte[] { 202, 137, 146 }, new byte[] { 232, 164, 83 }, new byte[] { 165, 143, 219 } };
        int hash = 0; foreach (char ch in c.Email) hash = (hash * 31 + ch) & int.MaxValue;
        var color = palette[hash % palette.Length];
        return new Border { Width = size, Height = size, CornerRadius = new CornerRadius(8), Background = new SolidColorBrush(Windows.UI.Color.FromArgb(255, color[0], color[1], color[2])), Child = new TextBlock { Text = c.Name.Length > 0 ? c.Name.Substring(0, 1).ToUpperInvariant() : "?", FontSize = size * .4, Foreground = new SolidColorBrush(Microsoft.UI.Colors.White), HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center } };
    }
    private static Border Card(UIElement child) => new() { Child = child, Padding = new Thickness(20), CornerRadius = new CornerRadius(8), BorderThickness = new Thickness(1), BorderBrush = new SolidColorBrush(Windows.UI.Color.FromArgb(100, 150, 160, 175)) };
    private void Edit(Contact? c)
    {
        editing = true;
        list.IsEnabled = false;
        search.IsEnabled = false;
        detail.Children.Clear();
        var fields = new[] { new TextBox { PlaceholderText = "姓名", Text = c?.Name ?? "", FontSize = 20 }, new TextBox { PlaceholderText = "邮箱", Text = c?.Email ?? "" }, new TextBox { PlaceholderText = "电话", Text = c?.Phone ?? "" }, new TextBox { PlaceholderText = "公司", Text = c?.Company ?? "" }, new TextBox { PlaceholderText = "部门", Text = c?.Department ?? "" }, new TextBox { PlaceholderText = "职位", Text = c?.Title ?? "" }, new TextBox { PlaceholderText = "备注", Text = c?.Notes ?? "", AcceptsReturn = true } };
        var actions = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8 };
        var save = new Button { Content = "保存", Background = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 40, 99, 207)), Foreground = new SolidColorBrush(Microsoft.UI.Colors.White) };
        var cancel = new Button { Content = "取消" };
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
        info.Children.Add(new TextBlock { Text = "联系信息", FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold });
        foreach (var (label, index) in new[] { ("邮箱", 1), ("电话", 2), ("公司", 3), ("备注", 6) })
        {
            var row = new Grid { ColumnSpacing = 12 };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(44) }); row.ColumnDefinitions.Add(new ColumnDefinition());
            row.Children.Add(new TextBlock { Text = label, VerticalAlignment = VerticalAlignment.Center });
            SetColumn(fields[index], 1); row.Children.Add(fields[index]); info.Children.Add(row);
        }
        var error = new TextBlock { TextWrapping = TextWrapping.Wrap, Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 190, 45, 45)) };
        info.Children.Add(error); detail.Children.Add(Card(info));
        void EndEdit() { editing = false; list.IsEnabled = true; search.IsEnabled = true; }
        cancel.Click += (_, _) => { EndEdit(); ShowDetail(); };
        save.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(fields[0].Text) || !MimeKit.MailboxAddress.TryParse(fields[1].Text.Trim(), out var parsed) || parsed.Address != fields[1].Text.Trim()) { error.Text = "请填写姓名和有效的邮箱地址。"; return; }
            if (contacts.Any(x => x.Id != c?.Id && string.Equals(x.Email, fields[1].Text.Trim(), StringComparison.OrdinalIgnoreCase))) { error.Text = "该邮箱已存在。"; return; }
            try
            {
                using var db = Open(); using var cmd = db.CreateCommand(); cmd.CommandText = c == null ? "INSERT INTO contacts(name,email,phone,company,department,title,notes) VALUES($n,$e,$p,$c,$d,$t,$o)" : "UPDATE contacts SET name=$n,email=$e,phone=$p,company=$c,department=$d,title=$t,notes=$o WHERE id=$id";
                string[] keys = { "$n", "$e", "$p", "$c", "$d", "$t", "$o" }; for (int i = 0; i < keys.Length; i++) cmd.Parameters.AddWithValue(keys[i], fields[i].Text.Trim()); if (c != null) cmd.Parameters.AddWithValue("$id", c.Id); cmd.ExecuteNonQuery();
                var email = fields[1].Text.Trim(); EndEdit(); Refresh(); list.SelectedItem = list.Items.OfType<ListViewItem>().FirstOrDefault(item => string.Equals((item.Tag as Contact)?.Email, email, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex) { error.Text = "保存失败：" + ex.Message; }
        };
        fields[0].Focus(FocusState.Programmatic);
    }
}
