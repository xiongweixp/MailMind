using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Media;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;

namespace MailMind;

public sealed partial class MainWindow : Window
{
    private readonly MailService mail = new();
    private readonly ObservableCollection<MailItem> messages = new();
    private readonly ObservableCollection<MailDateGroup> dateGroups = new();
    private readonly CollectionViewSource groupedMessages = new() { IsSourceGrouped = true };
    private bool aiVisible = true, loading;
    private int total, offset;
    private string folder = "收件箱";
    private ScrollViewer? listScroll;
    private bool pageLoadFailed, pageCheckQueued;
    private DateTime? searchDate;
    private bool htmlResourceHandlerReady;
    private readonly System.Collections.Generic.Dictionary<string, string> htmlDocuments = new();
    private readonly System.Collections.Generic.Dictionary<string, Button> folderButtons = new();
    public MainWindow()
    {
        InitializeComponent();
        ExtendsContentIntoTitleBar = true;
        SetTitleBar(DragRegion);
        AppWindow.TitleBar.ButtonBackgroundColor = Microsoft.UI.Colors.Transparent;
        AppWindow.TitleBar.ButtonInactiveBackgroundColor = Microsoft.UI.Colors.Transparent;
        AppWindow.TitleBar.ButtonForegroundColor = Microsoft.UI.Colors.Black;
        AppWindow.TitleBar.ButtonInactiveForegroundColor = Microsoft.UI.Colors.Gray;
        AppWindow.SetIcon(System.IO.Path.Combine(AppContext.BaseDirectory, "Assets", "MailMind.ico"));
        AppWindow.Resize(new Windows.Graphics.SizeInt32(1440, 900));
        groupedMessages.Source = dateGroups;
        MailList.ItemsSource = groupedMessages.View;
        messages.CollectionChanged += (_, e) =>
        {
            if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Add && e.NewItems != null)
            {
                foreach (MailItem item in e.NewItems) if (MatchesSearch(item)) AddToDateGroup(item);
            }
            else ApplySearch();
        };
        try { mail.Load(); if (mail.Account != null) { ShowAccount(mail.Account.Address); _ = LoadFoldersAsync(); } }
        catch (Exception ex) { ListStatus.Text = "读取账户失败：" + ex.Message; }
    }
    private void ShowAccount(string address)
    {
        AccountLabel.Text = address;
        AccountExpander.Visibility = Visibility.Visible;
        NoAccountHint.Visibility = Visibility.Collapsed;
    }
    private void UpdateFolderSelection(string name)
    {
        foreach (var (path, button) in folderButtons)
        {
            button.Background = new SolidColorBrush(path == name ? Windows.UI.Color.FromArgb(255, 225, 235, 252) : Microsoft.UI.Colors.Transparent);
            button.Foreground = new SolidColorBrush(path == name ? Windows.UI.Color.FromArgb(255, 27, 82, 181) : Windows.UI.Color.FromArgb(255, 23, 35, 58));
        }
    }
    private void ShowFolders(System.Collections.Generic.IEnumerable<MailFolderInfo> folders)
    {
        FolderList.Children.Clear(); folderButtons.Clear();
        foreach (var entry in folders)
        {
            var button = new Button { Content = entry.Name, IsEnabled = entry.Selectable, Margin = new Thickness(Math.Min(entry.Depth, 5) * 12, 0, 0, 0), HorizontalAlignment = HorizontalAlignment.Stretch, Style = (Style)((FrameworkElement)Content).Resources["SidebarButtonStyle"] };
            button.Click += async (_, _) => await OpenFolderAsync(entry.Path);
            FolderList.Children.Add(button); folderButtons[entry.Path] = button;
        }
        UpdateFolderSelection(folder);
    }
    private async Task LoadFoldersAsync()
    {
        var cached = mail.CachedFolders();
        if (cached.Count > 0)
        {
            ShowFolders(cached);
            var cachedInbox = cached.Find(f => f.Name == "收件箱" && f.Selectable);
            if (cachedInbox != null) await OpenFolderAsync(cachedInbox.Path);
        }
        try
        {
            if (messages.Count == 0) ListStatus.Text = "正在同步目录…";
            var folders = await mail.GetFoldersAsync();
            ShowFolders(folders);
            var inbox = System.Linq.Enumerable.FirstOrDefault(folders, f => f.Name == "收件箱" && f.Selectable);
            if (inbox != null && (cached.Count == 0 || folder != inbox.Path)) await OpenFolderAsync(inbox.Path);
            else if (inbox == null) ListStatus.Text = "服务器没有返回收件箱";
        }
        catch (Exception ex)
        {
            ListStatus.Text = cached.Count > 0 ? "目录同步失败，显示本地目录：" + ex.Message : "目录同步失败：" + ex.Message;
        }
    }
    private async Task OpenFolderAsync(string name)
    {
        if (loading) return;
        loading = true; pageLoadFailed = false; folder = name; UpdateFolderSelection(name); FolderTitle.Text = folderButtons.TryGetValue(name, out var selectedFolder) ? selectedFolder.Content.ToString() : name; messages.Clear(); offset = total = 0;
        ListEmptyText.Visibility = Visibility.Collapsed;
        if (mail.Account == null) { ListStatus.Text = "请先配置邮箱"; loading = false; return; }
        try
        {
            total = mail.BeginCachedFolder(name);
            int count = Math.Max(15, (int)Math.Ceiling(MailList.ActualHeight / 70) + 2);
            await foreach (var item in mail.GetPageAsync(0, count))
            {
                offset++;
                if (item != null) messages.Add(item);
            }
            ListStatus.Text = messages.Count > 0 ? $"本地缓存：{messages.Count} 封 · 正在同步…" : "正在连接邮箱…";
            int serverTotal = await mail.BeginFolderAsync(name);
            if (!mail.Offline)
            {
                messages.Clear(); offset = 0; total = serverTotal;
            }
            else
            {
                total = serverTotal;
                ListStatus.Text = $"离线缓存：已加载 {messages.Count} / {total} 封。连接失败：{mail.OfflineReason}";
            }
        }
        catch (Exception ex)
        {
            ListStatus.Text = "同步失败，显示本地缓存：" + ex.Message;
            loading = false;
            return;
        }
        loading = false;
        if (offset < total || total == 0) await LoadPageAsync();
        else QueuePageCheck();
    }
    private async Task LoadPageAsync()
    {
        if (loading || offset >= total) { if (total == 0 && !loading) { ListEmptyText.Visibility = Visibility.Visible; ListStatus.Text = "没有邮件"; } return; }
        loading = true;
        pageLoadFailed = false;
        bool pageLoaded = false;
        try
        {
            ListStatus.Text = "正在加载邮件…";
            int count = Math.Max(15, (int)Math.Ceiling(MailList.ActualHeight / 70) + 2);
            await foreach (var item in mail.GetPageAsync(offset, count))
            {
                // A missing UID was deleted on the server; still advance past it.
                offset++;
                if (item != null) messages.Add(item);
                ListStatus.Text = mail.Offline
                    ? $"离线缓存：已加载 {messages.Count} / {total} 封。连接失败：{mail.OfflineReason}"
                    : $"已加载 {messages.Count} / {total} 封";
            }
            pageLoaded = true;
        }
        catch (Exception ex) { pageLoadFailed = true; ListStatus.Text = "加载失败：" + ex.Message + "，滚动可重试"; }
        finally { loading = false; }
        if (pageLoaded) QueuePageCheck();
    }
    private void MailScroll_Loaded(object sender, RoutedEventArgs e)
    {
        listScroll = (ScrollViewer)sender;
        QueuePageCheck();
    }
    private void QueuePageCheck(bool retry = false)
    {
        if (pageCheckQueued || loading || offset >= total || (pageLoadFailed && !retry)) return;
        pageCheckQueued = true;
        if (!DispatcherQueue.TryEnqueue(async () =>
        {
            pageCheckQueued = false;
            if (!loading && offset < total && listScroll != null &&
                listScroll.VerticalOffset >= listScroll.ScrollableHeight - 150)
                await LoadPageAsync();
        })) pageCheckQueued = false;
    }
    private void MailList_Loaded(object sender, RoutedEventArgs e) => QueuePageCheck();
    private void MailList_LayoutUpdated(object? sender, object e) => QueuePageCheck();
    private void ListScroll_ViewChanged(object? sender, ScrollViewerViewChangedEventArgs e) => QueuePageCheck(retry: true);
    private void MailList_SizeChanged(object sender, SizeChangedEventArgs e) => QueuePageCheck();
    private void MailList_ContainerContentChanging(ListViewBase sender, ContainerContentChangingEventArgs args)
    {
        if (!args.InRecycleQueue && args.ItemIndex >= MailList.Items.Count - 3) QueuePageCheck();
    }
    private async void MailList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (MailList.SelectedItem is not MailItem item) return;
        if (ComposeWorkspace.Visibility == Visibility.Visible) { MailList.SelectedItem = null; await ShowNotice("正在写邮件", "请先关闭写信区，再打开邮件。"); return; }
        EmptyWorkspace.Visibility = Visibility.Collapsed; ReadingWorkspace.Visibility = Visibility.Visible;
        ReadSubject.Text = item.Subject; ReadSender.Text = item.Sender + "  ·  " + item.Date; ReadBody.Text = "正在加载正文…"; HtmlView.Visibility = Visibility.Collapsed; PlainScroll.Visibility = Visibility.Visible;
        if (item.IsSeen != true) _ = MarkSeenAsync(item);
        try { var body = await mail.GetBodyAsync(item.Uid); if (ReferenceEquals(MailList.SelectedItem, item)) { if (body.IsHtml) await ShowHtmlAsync(body.Content, item); else ReadBody.Text = body.Content; } }
        catch (Exception ex) { if (ReferenceEquals(MailList.SelectedItem, item)) ReadBody.Text = "加载正文失败：" + ex.Message; }
    }
    private async Task MarkSeenAsync(MailItem item)
    {
        var selectedFolder = folder;
        var selectedAccount = mail.Account;
        try
        {
            await mail.MarkSeenAsync(item.Uid);
            item.MarkSeen();
            if (folder == selectedFolder && ReferenceEquals(mail.Account, selectedAccount))
                foreach (var current in messages)
                    if (current.Uid == item.Uid && !ReferenceEquals(current, item)) current.MarkSeen();
        }
        catch (Exception ex) { ListStatus.Text = "已读状态同步失败，请重新选择邮件重试：" + ex.Message; }
    }
    private async Task ShowHtmlAsync(string html, MailItem item)
    {
        await HtmlView.EnsureCoreWebView2Async();
        if (!ReferenceEquals(MailList.SelectedItem, item)) return;
        HtmlView.CoreWebView2.Settings.IsScriptEnabled = false;
        HtmlView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
        if (!htmlResourceHandlerReady)
        {
            HtmlView.CoreWebView2.AddWebResourceRequestedFilter("http://mailmind.invalid/*", Microsoft.Web.WebView2.Core.CoreWebView2WebResourceContext.Document);
            HtmlView.CoreWebView2.WebResourceRequested += (_, args) =>
            {
                if (htmlDocuments.Remove(args.Request.Uri, out var content))
                {
                    var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(content));
                    args.Response = HtmlView.CoreWebView2.Environment.CreateWebResourceResponse(stream.AsRandomAccessStream(), 200, "OK", "Content-Type: text/html; charset=utf-8\r\nCache-Control: no-store");
                }
            };
            htmlResourceHandlerReady = true;
        }
        // The sandbox prevents email content from running scripts or navigating the app.
        var document = "<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; img-src data: https: http:; style-src 'unsafe-inline'\"><style>body{margin:0}iframe{border:0;width:100%;height:100vh}</style><iframe sandbox srcdoc=\"" + System.Net.WebUtility.HtmlEncode("<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; img-src data: https: http:; style-src 'unsafe-inline'\">" + html) + "\"></iframe>";
        var documentUri = "http://mailmind.invalid/" + Guid.NewGuid().ToString("N");
        htmlDocuments.Clear();
        htmlDocuments[documentUri] = document;
        HtmlView.Source = new Uri(documentUri);
        PlainScroll.Visibility = Visibility.Collapsed; HtmlView.Visibility = Visibility.Visible;
    }
    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e) => ApplySearch();
    private void ApplySearch()
    {
        if (SearchBox == null) return;
        dateGroups.Clear();
        foreach (var item in messages)
            if (MatchesSearch(item)) AddToDateGroup(item);
    }
    private void AddToDateGroup(MailItem item)
    {
        var (order, title) = MailDateGroup.Classify(item.Date, DateTime.Today);
        MailDateGroup? group = null;
        int index = 0;
        while (index < dateGroups.Count && dateGroups[index].Order < order) index++;
        if (index < dateGroups.Count && dateGroups[index].Order == order) group = dateGroups[index];
        if (group == null) { group = new MailDateGroup(order, title); dateGroups.Insert(index, group); }
        group.Add(item);
    }
    private bool MatchesSearch(MailItem item)
    {
        var query = SearchBox.Text.Trim();
        var matchesText = query.Length == 0 || item.Subject.Contains(query, StringComparison.CurrentCultureIgnoreCase) || item.Sender.Contains(query, StringComparison.CurrentCultureIgnoreCase) || item.Preview.Contains(query, StringComparison.CurrentCultureIgnoreCase);
        return matchesText && (!searchDate.HasValue || (DateTime.TryParse(item.Date, out var date) && date.Date == searchDate.Value));
    }
    private void SearchCalendar_SelectedDatesChanged(CalendarView sender, CalendarViewSelectedDatesChangedEventArgs args)
    {
        searchDate = sender.SelectedDates.Count > 0 ? sender.SelectedDates[0].Date : null;
        ToolTipService.SetToolTip(DateFilterButton, searchDate.HasValue ? $"日期：{searchDate:yyyy-MM-dd}（仅已加载邮件）" : "按日期筛选已加载邮件");
        DateFilterButton.Foreground = new SolidColorBrush(searchDate.HasValue ? Windows.UI.Color.FromArgb(255, 40, 99, 207) : Windows.UI.Color.FromArgb(255, 23, 35, 58));
        ApplySearch();
    }
    private void ClearDateFilter_Click(object sender, RoutedEventArgs e)
    {
        SearchCalendar.SelectedDates.Clear();
        searchDate = null;
        ApplySearch();
    }
    private async void Refresh_Click(object sender, RoutedEventArgs e) { if (mail.Account != null) await LoadFoldersAsync(); }
    private void Compose_Click(object sender, RoutedEventArgs e) { EmptyWorkspace.Visibility = Visibility.Collapsed; ReadingWorkspace.Visibility = Visibility.Collapsed; ComposeWorkspace.Visibility = Visibility.Visible; }
    private async void CloseCompose_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(ToBox.Text + SubjectBox.Text + BodyBox.Text))
        {
            var dialog = new ContentDialog { XamlRoot = ComposeWorkspace.XamlRoot, Title = "关闭写信？", Content = "邮件内容尚未保存。关闭后内容会丢失。", PrimaryButtonText = "继续编辑", CloseButtonText = "丢弃" };
            if (await dialog.ShowAsync() == ContentDialogResult.Primary) return;
        }
        ToBox.Text = SubjectBox.Text = BodyBox.Text = ""; ComposeWorkspace.Visibility = Visibility.Collapsed; EmptyWorkspace.Visibility = Visibility.Visible;
    }
    private async void Send_Click(object sender, RoutedEventArgs e) => await ShowNotice("尚不能发送", "SMTP 发送功能尚未实现，当前内容不会发送。");
    private async void Settings_Click(object sender, RoutedEventArgs e)
    {
        var a = mail.Account;
        var address = new TextBox { Header = "邮箱地址", Text = a?.Address ?? "", PlaceholderText = "name@example.com" };
        var user = new TextBox { Header = "登录用户名", Text = a?.UserName ?? "", PlaceholderText = "通常与邮箱地址相同" };
        var host = new TextBox { Header = "IMAP 服务器", Text = a?.ImapHost ?? "", PlaceholderText = "imap.example.com" };
        var port = new TextBox { Header = "端口", Text = (a?.ImapPort ?? 993).ToString() };
        var ssl = new CheckBox { Content = "使用 SSL/TLS", IsChecked = !(a?.StartTls ?? false) };
        ssl.Checked += (_, _) => port.Text = "993";
        ssl.Unchecked += (_, _) => port.Text = "143";
        var smtpHost = new TextBox { Header = "SMTP 服务器", Text = a?.SmtpHost ?? "", PlaceholderText = "smtp.example.com（可稍后填写）" };
        var smtpPort = new TextBox { Header = "SMTP 端口", Text = (a?.SmtpPort > 0 ? a.SmtpPort : 465).ToString() };
        var smtpSsl = new CheckBox { Content = "使用 SSL/TLS", IsChecked = !(a?.SmtpStartTls ?? false) };
        smtpSsl.Checked += (_, _) => smtpPort.Text = "465";
        smtpSsl.Unchecked += (_, _) => smtpPort.Text = "587";
        var secret = new PasswordBox { Header = "密码或邮箱授权码", PlaceholderText = "请输入授权码" };
        var error = new TextBlock { TextWrapping = TextWrapping.Wrap, Foreground = new Microsoft.UI.Xaml.Media.SolidColorBrush(Microsoft.UI.Colors.Firebrick) };
        var panel = new StackPanel { Spacing = 10, Children = {
            new TextBlock { Text = "收发服务器", FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold },
            new TextBlock { Text = "配置邮箱账户及服务器。保存时测试 IMAP 连接；SMTP 发送功能尚未接入。", TextWrapping = TextWrapping.Wrap },
            address, user, secret,
            new TextBlock { Text = "接收邮件 · IMAP", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, Margin = new Thickness(0, 8, 0, 0) },
            host, port, ssl,
            new TextBlock { Text = "发送邮件 · SMTP", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, Margin = new Thickness(0, 8, 0, 0) },
            smtpHost, smtpPort, smtpSsl } };
        var navigation = new ListView { Width = 160, SelectionMode = ListViewSelectionMode.Single, Items = { "收发服务器" }, SelectedIndex = 0 };
        var layout = new Grid { Width = 720, MaxWidth = 720, ColumnSpacing = 24 };
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        layout.Children.Add(navigation);
        var scroll = new ScrollViewer { Content = panel, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
        var save = new Button { Content = "测试并保存" };
        var cancel = new Button { Content = "取消" };
        var actions = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Spacing = 8, Margin = new Thickness(0, 12, 0, 0), Children = { cancel, save } };
        var page = new Grid { Height = 500, RowSpacing = 8 };
        page.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        page.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        page.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        page.Children.Add(scroll);
        Grid.SetRow(error, 1); page.Children.Add(error);
        Grid.SetRow(actions, 2); page.Children.Add(actions);
        Grid.SetColumn(page, 1); layout.Children.Add(page);
        var dialog = new ContentDialog { XamlRoot = ComposeWorkspace.XamlRoot, Title = "设置", Content = layout };
        dialog.Resources["ContentDialogMaxWidth"] = 800.0;
        bool saved = false;
        cancel.Click += (_, _) => dialog.Hide();
        save.Click += async (_, _) =>
        {
            save.IsEnabled = false; cancel.IsEnabled = false;
            try
            {
                if (string.IsNullOrWhiteSpace(address.Text) || string.IsNullOrWhiteSpace(host.Text) || string.IsNullOrWhiteSpace(secret.Password) || !int.TryParse(port.Text, out int number) || number < 1 || number > 65535) { error.Text = "请填写邮箱、服务器、有效端口和密码或授权码。"; return; }
                if (!int.TryParse(smtpPort.Text, out int smtpNumber) || smtpNumber < 1 || smtpNumber > 65535) { error.Text = "请输入有效的 SMTP 端口。"; return; }
                var candidate = new MailAccount(address.Text.Trim(), string.IsNullOrWhiteSpace(user.Text) ? address.Text.Trim() : user.Text.Trim(), host.Text.Trim(), number, ssl.IsChecked != true, smtpHost.Text.Trim(), smtpNumber, smtpSsl.IsChecked != true);
                error.Text = "正在测试连接…"; await mail.TestAsync(candidate, secret.Password); mail.Save(candidate, secret.Password); ShowAccount(candidate.Address); saved = true; dialog.Hide();
            }
            catch (Exception ex) { error.Text = "连接或保存失败：" + ex.Message; }
            finally { save.IsEnabled = true; cancel.IsEnabled = true; }
        };
        await dialog.ShowAsync();
        if (saved) await LoadFoldersAsync();
    }
    private async Task ShowNotice(string title, string content) => await new ContentDialog { XamlRoot = ComposeWorkspace.XamlRoot, Title = title, Content = content, CloseButtonText = "知道了" }.ShowAsync();
    private void ToggleAi_Click(object sender, RoutedEventArgs e)
    {
        aiVisible = !aiVisible;
        AiColumn.Width = new GridLength(aiVisible ? 320 : 0);
        AiPane.Visibility = aiVisible ? Visibility.Visible : Visibility.Collapsed;
        ReadingToolbar.Padding = new Thickness(18, aiVisible ? 14 : 40, 18, 14);
        AiToggleButton.Visibility = aiVisible ? Visibility.Visible : Visibility.Collapsed;
        AiFloatingButton.Visibility = aiVisible ? Visibility.Collapsed : Visibility.Visible;
    }
}

