using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Input;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace MailMind;

public sealed partial class MainWindow : Window
{
    private readonly MailService mail = new();
    private readonly DecisionService decision = new();
    private readonly System.Threading.SemaphoreSlim decisionQueue = new(1, 1);
    private readonly System.Threading.CancellationTokenSource windowLifetime = new();
    private System.Threading.CancellationTokenSource? searchCancellation;
    private readonly System.Collections.Generic.List<MailItem> aiResults = new();
    private bool aiSearchShowing;
    private bool closing;
    private static readonly string LayoutPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailMind", "layout.json");
    private double expandedAiWidth = 320, dragStartX, dragFirstWidth, dragSecondWidth;
    private int activeSplitter = -1;
    private uint dragPointerId;
    private record ColumnWidths(double Sidebar, double Mail, double Ai);
    private delegate nint WindowSubclassProc(nint hwnd, uint message, nuint wParam, nint lParam, nuint subclassId, nuint referenceData);
    private WindowSubclassProc? windowSubclass;
    [StructLayout(LayoutKind.Sequential)]
    private struct NativePoint { public int X, Y; }
    [StructLayout(LayoutKind.Sequential)]
    private struct MinMaxInfo { public NativePoint Reserved, MaxSize, MaxPosition, MinTrackSize, MaxTrackSize; }
    [DllImport("comctl32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetWindowSubclass(nint hwnd, WindowSubclassProc callback, nuint id, nuint data);
    [DllImport("comctl32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool RemoveWindowSubclass(nint hwnd, WindowSubclassProc callback, nuint id);
    [DllImport("comctl32.dll")]
    private static extern nint DefSubclassProc(nint hwnd, uint message, nuint wParam, nint lParam);
    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(nint hwnd);
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
    private MailAttachment? previewAttachment;
    private SystemPreviewHost? systemPreview;
    private readonly System.Collections.Generic.Dictionary<string, string> htmlDocuments = new();
    private readonly System.Collections.Generic.Dictionary<string, Button> folderButtons = new();
    public MainWindow()
    {
        InitializeComponent();
        LoadColumnWidths();
        windowSubclass = WindowMessage;
        if (!SetWindowSubclass(WinRT.Interop.WindowNative.GetWindowHandle(this), windowSubclass, 1, 0))
            throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        AppWindow.Closing += (_, _) => { closing = true; searchCancellation?.Cancel(); windowLifetime.Cancel(); systemPreview?.Dispose(); };
        AppWindow.Changed += (_, _) => UpdatePreviewBounds();
        ExtendsContentIntoTitleBar = true;
        SetTitleBar(DragRegion);
        AppWindow.TitleBar.ButtonBackgroundColor = Microsoft.UI.Colors.Transparent;
        AppWindow.TitleBar.ButtonInactiveBackgroundColor = Microsoft.UI.Colors.Transparent;
        ((FrameworkElement)Content).ActualThemeChanged += (_, _) => UpdateThemeColors();
        AppWindow.SetIcon(System.IO.Path.Combine(AppContext.BaseDirectory, "Assets", "MailMind.ico"));
        AppWindow.Resize(new Windows.Graphics.SizeInt32(1440, 900));
        groupedMessages.Source = dateGroups;
        MailList.ItemsSource = groupedMessages.View;
        messages.CollectionChanged += (_, e) =>
        {
            if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Add && e.NewItems != null)
            {
                foreach (MailItem item in e.NewItems) { if (!aiSearchShowing && MatchesSearch(item)) AddToDateGroup(item); _ = ClassifyAsync(item); }
            }
            else ApplySearch();
        };
        UpdateThemeColors();
        try { decision.Load(); mail.Load(); if (mail.Account != null) { ShowAccount(mail.Account.Address); _ = LoadFoldersAsync(); } }
        catch (Exception ex) { ListStatus.Text = "读取账户失败：" + ex.Message; }
    }
    private nint WindowMessage(nint hwnd, uint message, nuint wParam, nint lParam, nuint subclassId, nuint referenceData)
    {
        const uint getMinMaxInfo = 0x0024, ncDestroy = 0x0082;
        if (message == getMinMaxInfo)
        {
            var info = Marshal.PtrToStructure<MinMaxInfo>(lParam);
            double dpi = GetDpiForWindow(hwnd) / 96.0;
            // The reading pane takes the remaining width, so preserve it alongside the saved fixed columns.
            double width = SidebarColumn.Width.Value + MailColumn.Width.Value + 220 + expandedAiWidth + 16;
            info.MinTrackSize.X = Math.Max(info.MinTrackSize.X, (int)Math.Ceiling(width * dpi));
            Marshal.StructureToPtr(info, lParam, false);
            return 0;
        }
        if (message == ncDestroy && windowSubclass != null) RemoveWindowSubclass(hwnd, windowSubclass, subclassId);
        return DefSubclassProc(hwnd, message, wParam, lParam);
    }
    private async Task ClassifyAsync(MailItem item, bool force = false)
    {
        var settings = decision.Settings;
        var fingerprint = DecisionService.Fingerprint(settings, item);
        string? cached;
        try { cached = force ? null : mail.CachedClassification(item, fingerprint); }
        catch (Exception ex) { item.Importance = "失败"; if (!closing) ListStatus.Text = "读取 AI 分类缓存失败：" + ex.Message; return; }
        if (cached != null) { item.Importance = cached; return; }
        if (!settings.Enabled) { item.Importance = ""; return; }
        item.Importance = "排队中";
        try { await decisionQueue.WaitAsync(windowLifetime.Token); }
        catch (OperationCanceledException) { return; }
        try
        {
            if (!System.Linq.Enumerable.Any(messages, current => ReferenceEquals(current, item)) || decision.Settings != settings) return;
            cached = force ? null : mail.CachedClassification(item, fingerprint);
            if (cached != null) { item.Importance = cached; return; }
            item.Importance = "判断中";
            var result = await decision.ClassifyAsync(settings, item, windowLifetime.Token);
            mail.CacheClassification(item, fingerprint, result);
            if (!closing && System.Linq.Enumerable.Any(messages, current => ReferenceEquals(current, item)) && decision.Settings == settings) item.Importance = result;
        }
        catch (OperationCanceledException) when (closing) { }
        catch (Exception ex) { if (!closing && System.Linq.Enumerable.Any(messages, current => ReferenceEquals(current, item))) { item.Importance = "失败"; ListStatus.Text = "AI 分类失败：" + ex.Message; } }
        finally { decisionQueue.Release(); }
    }
    private void ShowAccount(string address)
    {
        AccountLabel.Text = address;
        AccountExpander.Visibility = Visibility.Visible;
        NoAccountHint.Visibility = Visibility.Collapsed;
    }
    private void UpdateFolderSelection(string name)
    {
        bool dark = ((FrameworkElement)Content).ActualTheme == ElementTheme.Dark;
        foreach (var (path, button) in folderButtons)
        {
            bool selected = path == name;
            button.Background = new SolidColorBrush(selected ? (dark ? Windows.UI.Color.FromArgb(255, 43, 68, 98) : Windows.UI.Color.FromArgb(255, 225, 235, 252)) : Microsoft.UI.Colors.Transparent);
            button.Foreground = new SolidColorBrush(selected ? (dark ? Windows.UI.Color.FromArgb(255, 173, 207, 255) : Windows.UI.Color.FromArgb(255, 27, 82, 181)) : (dark ? Windows.UI.Color.FromArgb(255, 242, 243, 245) : Windows.UI.Color.FromArgb(255, 23, 35, 58)));
        }
    }
    private void UpdateThemeColors()
    {
        bool dark = ((FrameworkElement)Content).ActualTheme == ElementTheme.Dark;
        AppWindow.TitleBar.ButtonForegroundColor = dark ? Microsoft.UI.Colors.White : Microsoft.UI.Colors.Black;
        AppWindow.TitleBar.ButtonInactiveForegroundColor = dark ? Windows.UI.Color.FromArgb(255, 165, 165, 165) : Microsoft.UI.Colors.Gray;
        UpdateFolderSelection(folder);
        UpdateDateFilterColor();
    }
    private void UpdateDateFilterColor()
    {
        if (DateFilterButton == null) return;
        if (!searchDate.HasValue) { DateFilterButton.ClearValue(Control.ForegroundProperty); return; }
        bool dark = ((FrameworkElement)Content).ActualTheme == ElementTheme.Dark;
        DateFilterButton.Foreground = new SolidColorBrush(dark ? Windows.UI.Color.FromArgb(255, 145, 186, 255) : Windows.UI.Color.FromArgb(255, 40, 99, 207));
    }
    private void ShowFolders(System.Collections.Generic.IEnumerable<MailFolderInfo> folders)
    {
        FolderList.Children.Clear(); folderButtons.Clear();
        foreach (var entry in System.Linq.Enumerable.OrderBy(folders, f => f.Name == "收件箱" ? 0 : 1))
        {
            object content = entry.Name;
            if (entry.Depth == 0)
            {
                var row = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 12 };
                row.Children.Add(new FontIcon { Glyph = FolderGlyph(entry.Name), FontSize = 16 });
                row.Children.Add(new TextBlock { Text = entry.Name, VerticalAlignment = VerticalAlignment.Center });
                content = row;
            }
            var button = new Button { Content = content, Tag = entry.Name, IsEnabled = entry.Selectable, Margin = new Thickness(Math.Min(entry.Depth, 5) * 12, 0, 0, 0), HorizontalAlignment = HorizontalAlignment.Stretch, Style = (Style)((FrameworkElement)Content).Resources["SidebarButtonStyle"] };
            button.Click += async (_, _) => await OpenFolderAsync(entry.Path);
            FolderList.Children.Add(button); folderButtons[entry.Path] = button;
        }
        UpdateFolderSelection(folder);
    }
    private static string FolderGlyph(string name) => name.ToLowerInvariant() switch
    {
        "收件箱" or "inbox" => "\uE715",
        "已发送" or "sent" or "sent items" => "\uE724",
        "草稿箱" or "草稿" or "drafts" => "\uE8A5",
        "已删除" or "trash" or "deleted items" => "\uE74D",
        "归档" or "archive" => "\uE7B8",
        "垃圾邮件" or "junk" or "spam" => "\uE7BA",
        _ => "\uE8B7"
    };
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
            if (closing) return;
            ShowFolders(folders);
            var inbox = System.Linq.Enumerable.FirstOrDefault(folders, f => f.Name == "收件箱" && f.Selectable);
            if (inbox != null && (cached.Count == 0 || folder != inbox.Path)) await OpenFolderAsync(inbox.Path);
            else if (inbox == null) ListStatus.Text = "服务器没有返回收件箱";
        }
        catch (Exception ex)
        {
            if (!closing) ListStatus.Text = cached.Count > 0 ? "目录同步失败，显示本地目录：" + ex.Message : "目录同步失败：" + ex.Message;
        }
    }
    private async Task OpenFolderAsync(string name)
    {
        if (loading) return;
        searchCancellation?.Cancel(); searchCancellation = null; aiSearchShowing = false; aiResults.Clear();
        AiSearchStatusPanel.Visibility = Visibility.Collapsed;
        loading = true; pageLoadFailed = false; folder = name; UpdateFolderSelection(name); FolderTitle.Text = folderButtons.TryGetValue(name, out var selectedFolder) ? (string)selectedFolder.Tag : name; messages.Clear(); offset = total = 0;
        ListEmptyText.Visibility = Visibility.Collapsed;
        if (mail.Account == null) { ListStatus.Text = "请先配置邮箱"; loading = false; return; }
        try
        {
            total = mail.BeginCachedFolder(name);
            int count = Math.Max(15, (int)Math.Ceiling(MailList.ActualHeight / 70) + 2);
            await foreach (var item in mail.GetPageAsync(0, count))
            {
                if (closing) return;
                offset++;
                if (item != null) messages.Add(item);
            }
            ListStatus.Text = messages.Count > 0 ? $"本地缓存：{messages.Count} 封 · 正在同步…" : "正在连接邮箱…";
            int serverTotal = await mail.BeginFolderAsync(name);
            if (closing) return;
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
            if (!closing) ListStatus.Text = "同步失败，显示本地缓存：" + ex.Message;
            loading = false;
            return;
        }
        loading = false;
        if (offset < total || total == 0) await LoadPageAsync();
        else QueuePageCheck();
    }
    private async Task LoadPageAsync()
    {
        if (closing) return;
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
                if (closing) return;
                // A missing UID was deleted on the server; still advance past it.
                offset++;
                if (item != null) messages.Add(item);
                ListStatus.Text = mail.Offline
                    ? $"离线缓存：已加载 {messages.Count} / {total} 封。连接失败：{mail.OfflineReason}"
                    : $"已加载 {messages.Count} / {total} 封";
            }
            pageLoaded = true;
        }
        catch (Exception ex) { pageLoadFailed = true; if (!closing) ListStatus.Text = "加载失败：" + ex.Message + "，滚动可重试"; }
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
        if (closing || aiSearchShowing || pageCheckQueued || loading || offset >= total || (pageLoadFailed && !retry)) return;
        pageCheckQueued = true;
        if (!DispatcherQueue.TryEnqueue(async () =>
        {
            pageCheckQueued = false;
            if (!closing && !aiSearchShowing && !loading && offset < total && listScroll != null &&
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
        if (closing || MailList.SelectedItem is not MailItem item) return;
        if (ComposeWorkspace.Visibility == Visibility.Visible) { MailList.SelectedItem = null; await ShowNotice("正在写邮件", "请先关闭写信区，再打开邮件。"); return; }
        EmptyWorkspace.Visibility = Visibility.Collapsed; PreviewWorkspace.Visibility = Visibility.Collapsed; ReadingWorkspace.Visibility = Visibility.Visible; ReadingToolbar.Visibility = Visibility.Visible;
        systemPreview?.Close(); previewAttachment = null; AttachmentPanel.Visibility = Visibility.Collapsed; AttachmentList.ItemsSource = null;
        ReadSubject.Text = item.Subject; ReadSender.Text = item.Sender + "  ·  " + item.Date; ReadBody.Text = "正在加载正文…"; HtmlView.Visibility = Visibility.Collapsed; PlainScroll.Visibility = Visibility.Visible;
        if (item.IsSeen != true) _ = MarkSeenAsync(item);
        try { var body = await mail.GetBodyAsync(item.Uid); if (!closing && ReferenceEquals(MailList.SelectedItem, item)) { if (body.IsHtml) await ShowHtmlAsync(body.Content, item); else ReadBody.Text = body.Content; } }
        catch (Exception ex) { if (!closing && ReferenceEquals(MailList.SelectedItem, item)) ReadBody.Text = "加载正文失败：" + ex.Message; }
        try
        {
            var attachments = await mail.GetAttachmentsAsync(item);
            if (!closing && ReferenceEquals(MailList.SelectedItem, item) && attachments.Count > 0)
            {
                AttachmentHeading.Text = $"附件（{attachments.Count}）"; AttachmentList.ItemsSource = attachments; AttachmentPanel.Visibility = Visibility.Visible;
            }
        }
        catch (Exception ex) { if (!closing && ReferenceEquals(MailList.SelectedItem, item)) { AttachmentHeading.Text = "附件加载失败：" + ex.Message; AttachmentPanel.Visibility = Visibility.Visible; } }
    }
    private async void Attachment_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: MailAttachment attachment }) return;
        previewAttachment = attachment; ReadingWorkspace.Visibility = Visibility.Collapsed; ReadingToolbar.Visibility = Visibility.Collapsed; PreviewWorkspace.Visibility = Visibility.Visible;
        PreviewName.Text = attachment.Name; PreviewSize.Text = attachment.SizeLabel; PreviewMessage.Text = "正在加载预览…";
        PreviewProgress.Visibility = Visibility.Visible; PreviewProgress.IsActive = true;
        try
        {
            await Task.Yield();
            if (!ReferenceEquals(previewAttachment, attachment)) return;
            PreviewWorkspace.UpdateLayout();
            systemPreview ??= new SystemPreviewHost();
            var bounds = PreviewBounds();
            if (bounds.width <= 1 || bounds.height <= 1) throw new InvalidOperationException("预览区域尚未完成布局，请返回邮件后重试。");
            await systemPreview.ShowAsync(WinRT.Interop.WindowNative.GetWindowHandle(this), attachment, bounds.x, bounds.y, bounds.width, bounds.height);
            if (ReferenceEquals(previewAttachment, attachment))
            {
                // Layout can change while the out-of-process Office previewer starts.
                bounds = PreviewBounds(); systemPreview.Resize(bounds.x, bounds.y, bounds.width, bounds.height);
                PreviewMessage.Text = "";
            }
        }
        catch (Exception ex) { if (ReferenceEquals(previewAttachment, attachment)) PreviewMessage.Text = "无法预览此附件：" + ex.Message; }
        finally { if (ReferenceEquals(previewAttachment, attachment)) { PreviewProgress.IsActive = false; PreviewProgress.Visibility = Visibility.Collapsed; } }
    }
    private (int x, int y, int width, int height) PreviewBounds()
    {
        var point = PreviewSurface.TransformToVisual((UIElement)Content).TransformPoint(new Windows.Foundation.Point(0, 0));
        var scale = PreviewSurface.XamlRoot.RasterizationScale;
        return ((int)Math.Round(point.X * scale), (int)Math.Round(point.Y * scale), (int)Math.Round(PreviewSurface.ActualWidth * scale), (int)Math.Round(PreviewSurface.ActualHeight * scale));
    }
    private void PreviewSurface_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        UpdatePreviewBounds();
    }
    private void UpdatePreviewBounds()
    {
        if (previewAttachment == null || systemPreview == null || PreviewWorkspace.Visibility != Visibility.Visible) return;
        var bounds = PreviewBounds(); systemPreview.Resize(bounds.x, bounds.y, bounds.width, bounds.height);
    }
    private void BackFromPreview_Click(object sender, RoutedEventArgs e)
    {
        previewAttachment = null; systemPreview?.Close(); PreviewWorkspace.Visibility = Visibility.Collapsed; ReadingWorkspace.Visibility = Visibility.Visible; ReadingToolbar.Visibility = Visibility.Visible;
    }
    private async void SaveAttachment_Click(object sender, RoutedEventArgs e)
    {
        var attachment = previewAttachment; if (attachment == null) return;
        try
        {
            var picker = new Windows.Storage.Pickers.FileSavePicker();
            WinRT.Interop.InitializeWithWindow.Initialize(picker, WinRT.Interop.WindowNative.GetWindowHandle(this));
            var safeName = Path.GetFileName(attachment.Name.Replace('\\', '/'));
            var extension = Path.GetExtension(safeName);
            if (string.IsNullOrWhiteSpace(extension) || extension.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) extension = ".bin";
            picker.SuggestedFileName = Path.GetFileNameWithoutExtension(safeName); picker.FileTypeChoices.Add("附件", new System.Collections.Generic.List<string> { extension });
            var file = await picker.PickSaveFileAsync();
            if (file != null) await Windows.Storage.FileIO.WriteBytesAsync(file, attachment.Data);
        }
        catch (Exception ex) { await ShowNotice("保存附件失败", ex.Message); }
    }
    private async Task MarkSeenAsync(MailItem item)
    {
        var selectedFolder = folder;
        var selectedAccount = mail.Account;
        try
        {
            await mail.MarkSeenAsync(item.Uid);
            if (closing) return;
            item.MarkSeen();
            if (folder == selectedFolder && ReferenceEquals(mail.Account, selectedAccount))
                foreach (var current in messages)
                    if (current.Uid == item.Uid && !ReferenceEquals(current, item)) current.MarkSeen();
        }
        catch (Exception ex) { if (!closing) ListStatus.Text = "已读状态同步失败，请重新选择邮件重试：" + ex.Message; }
    }
    private async Task ShowHtmlAsync(string html, MailItem item)
    {
        await HtmlView.EnsureCoreWebView2Async();
        if (closing || !ReferenceEquals(MailList.SelectedItem, item)) return;
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
    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (aiSearchShowing) { searchCancellation?.Cancel(); searchCancellation = null; aiSearchShowing = false; aiResults.Clear(); if (AiSearchStatusPanel != null) AiSearchStatusPanel.Visibility = Visibility.Collapsed; }
        if (AiSearchButton?.IsChecked == true && AiSearchStatusPanel != null && AiSearchStopButton != null && AiSearchStatus != null)
        {
            AiSearchStatusPanel.Visibility = string.IsNullOrWhiteSpace(SearchBox.Text) ? Visibility.Collapsed : Visibility.Visible;
            AiSearchStopButton.Visibility = Visibility.Collapsed;
            AiSearchStatus.Text = "按回车开始 AI 搜索（本地缓存）";
        }
        ApplySearch();
    }
    private async void SearchBox_KeyDown(object sender, KeyRoutedEventArgs e)
    {
        if (e.Key == Windows.System.VirtualKey.Enter && AiSearchButton.IsChecked == true) { e.Handled = true; await StartAiSearchAsync(); }
    }
    private async void AiSearchModeChanged(object sender, RoutedEventArgs e)
    {
        if (AiSearchButton == null || SearchBox == null) return;
        bool enabled = AiSearchButton.IsChecked == true;
        if (AiSwitchTrack != null) AiSwitchTrack.Background = new SolidColorBrush(enabled ? Windows.UI.Color.FromArgb(255, 40, 99, 207) : Windows.UI.Color.FromArgb(255, 200, 209, 222));
        if (AiSwitchThumb != null) AiSwitchThumb.Margin = new Thickness(enabled ? 14 : 2, 0, 0, 0);
        SearchBox.PlaceholderText = enabled ? "输入问题，按回车搜索" : "搜索邮件";
        if (enabled && !string.IsNullOrWhiteSpace(SearchBox.Text)) await StartAiSearchAsync();
        else if (!enabled) { searchCancellation?.Cancel(); searchCancellation = null; aiSearchShowing = false; aiResults.Clear(); AiSearchStatusPanel.Visibility = Visibility.Collapsed; ApplySearch(); QueuePageCheck(); }
    }
    private void AiSearchStop_Click(object sender, RoutedEventArgs e) => searchCancellation?.Cancel();
    private void ClearSearch_Click(object sender, RoutedEventArgs e)
    {
        searchCancellation?.Cancel(); searchCancellation = null; aiSearchShowing = false; aiResults.Clear();
        SearchBox.Text = "";
        SearchCalendar.SelectedDates.Clear(); searchDate = null;
        AiSearchStatusPanel.Visibility = Visibility.Collapsed;
        ApplySearch(); QueuePageCheck();
    }
    private async Task StartAiSearchAsync()
    {
        string query = SearchBox.Text.Trim();
        if (query.Length == 0) { AiSearchStatusPanel.Visibility = Visibility.Visible; AiSearchStopButton.Visibility = Visibility.Collapsed; AiSearchStatus.Text = "输入搜索问题后按回车。"; return; }
        searchCancellation?.Cancel();
        var cancellation = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(windowLifetime.Token);
        searchCancellation = cancellation;
        aiSearchShowing = true; aiResults.Clear(); dateGroups.Clear();
        UpdateClearSearchVisibility();
        AiSearchStatusPanel.Visibility = Visibility.Visible; AiSearchStopButton.Visibility = Visibility.Visible;
        ListEmptyText.Visibility = Visibility.Collapsed;
        var settings = decision.Settings;
        try
        {
            var items = mail.CachedSearchItems(folder);
            int completed = 0, failed = 0;
            AiSearchStatus.Text = $"AI 搜索中 · 本地缓存 0/{items.Count} · 匹配 0";
            foreach (var item in items)
            {
                cancellation.Token.ThrowIfCancellationRequested();
                try
                {
                    var body = await mail.GetSearchBodyAsync(item, cancellation.Token);
                    bool matches = false;
                    const int chunkSize = 2400, overlap = 240;
                    for (int start = 0; ;)
                    {
                        int end = Math.Min(start + chunkSize, body.Length);
                        if (end < body.Length && char.IsHighSurrogate(body[end - 1]) && char.IsLowSurrogate(body[end])) end--;
                        if (await decision.MatchesAsync(settings, item, body[start..end], query, cancellation.Token)) { matches = true; break; }
                        if (end == body.Length) break;
                        start = end - overlap;
                        if (start > 0 && char.IsLowSurrogate(body[start]) && char.IsHighSurrogate(body[start - 1])) start--;
                    }
                    cancellation.Token.ThrowIfCancellationRequested();
                    if (matches)
                    {
                        item.Importance = mail.CachedClassification(item, DecisionService.Fingerprint(settings, item)) ?? "";
                        aiResults.Add(item);
                        if (!searchDate.HasValue || (DateTime.TryParse(item.Date, out var date) && date.Date == searchDate.Value)) AddToDateGroup(item);
                    }
                }
                catch (OperationCanceledException) when (cancellation.IsCancellationRequested) { throw; }
                catch (System.Net.Http.HttpRequestException) { throw; }
                catch (Exception) { failed++; }
                completed++;
                if (ReferenceEquals(searchCancellation, cancellation)) AiSearchStatus.Text = $"AI 搜索中 · 本地缓存 {completed}/{items.Count} · 匹配 {aiResults.Count}" + (failed > 0 ? $" · 失败 {failed}" : "");
            }
            if (ReferenceEquals(searchCancellation, cancellation)) AiSearchStatus.Text = $"AI 搜索完成 · 本地邮件索引 {completed}/{items.Count} · 匹配 {aiResults.Count}" + (failed > 0 ? $" · 正文获取或判断失败 {failed}" : "") + "（按邮件内容判断）";
        }
        catch (OperationCanceledException) { if (!closing && ReferenceEquals(searchCancellation, cancellation)) AiSearchStatus.Text = $"AI 搜索已中止 · 匹配 {aiResults.Count}（本地缓存）"; }
        catch (Exception ex) { if (!closing && ReferenceEquals(searchCancellation, cancellation)) AiSearchStatus.Text = "AI 搜索失败：" + ex.Message; }
        finally { if (!closing && ReferenceEquals(searchCancellation, cancellation)) { searchCancellation = null; AiSearchStopButton.Visibility = Visibility.Collapsed; } cancellation.Dispose(); }
    }
    private void ApplySearch()
    {
        if (SearchBox == null) return;
        UpdateClearSearchVisibility();
        dateGroups.Clear();
        foreach (var item in aiSearchShowing ? (System.Collections.Generic.IEnumerable<MailItem>)aiResults : messages)
            if (aiSearchShowing ? !searchDate.HasValue || (DateTime.TryParse(item.Date, out var date) && date.Date == searchDate.Value) : MatchesSearch(item)) AddToDateGroup(item);
    }
    private void UpdateClearSearchVisibility()
    {
        if (ClearSearchButton == null || SearchBox == null) return;
        bool filtered = aiSearchShowing || searchDate.HasValue || (AiSearchButton?.IsChecked != true && !string.IsNullOrWhiteSpace(SearchBox.Text));
        ClearSearchButton.Visibility = filtered ? Visibility.Visible : Visibility.Collapsed;
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
        var query = AiSearchButton?.IsChecked == true ? "" : SearchBox.Text.Trim();
        var matchesText = query.Length == 0 || item.Subject.Contains(query, StringComparison.CurrentCultureIgnoreCase) || item.Sender.Contains(query, StringComparison.CurrentCultureIgnoreCase) || item.Preview.Contains(query, StringComparison.CurrentCultureIgnoreCase);
        return matchesText && (!searchDate.HasValue || (DateTime.TryParse(item.Date, out var date) && date.Date == searchDate.Value));
    }
    private void SearchCalendar_SelectedDatesChanged(CalendarView sender, CalendarViewSelectedDatesChangedEventArgs args)
    {
        searchDate = sender.SelectedDates.Count > 0 ? sender.SelectedDates[0].Date : null;
        ToolTipService.SetToolTip(DateFilterButton, searchDate.HasValue ? $"日期：{searchDate:yyyy-MM-dd}（仅已加载邮件）" : "按日期筛选已加载邮件");
        UpdateDateFilterColor();
        ApplySearch();
    }
    private void ClearDateFilter_Click(object sender, RoutedEventArgs e)
    {
        SearchCalendar.SelectedDates.Clear();
        searchDate = null;
        ApplySearch();
    }
    private async void Refresh_Click(object sender, RoutedEventArgs e) { if (mail.Account != null) await LoadFoldersAsync(); }
    private bool composeReady, sending;
    private int sendStatusVersion;
    private async void Compose_Click(object sender, RoutedEventArgs e)
    {
        if (sending) return;
        systemPreview?.Close(); previewAttachment = null; PreviewWorkspace.Visibility = Visibility.Collapsed; EmptyWorkspace.Visibility = Visibility.Collapsed; ReadingWorkspace.Visibility = Visibility.Collapsed; ReadingToolbar.Visibility = Visibility.Collapsed;
        FromBox.Items.Clear(); if (mail.Account != null) { FromBox.Items.Add(mail.Account.Address); FromBox.SelectedIndex = 0; }
        ComposeWorkspace.Visibility = Visibility.Visible;
        if (!composeReady)
        {
            try
            {
                await BodyBox.EnsureCoreWebView2Async();
                BodyBox.NavigateToString("<!doctype html><html><head><meta charset='utf-8'><style>body{font:16px 'Microsoft YaHei',sans-serif;line-height:1.6;margin:14px;outline:none}table{border-collapse:collapse}td,th{border:1px solid #888;min-width:80px;padding:6px}</style></head><body contenteditable='true'></body></html>");
                composeReady = true;
            }
            catch (Exception ex) { await ShowNotice("无法打开正文编辑器", ex.Message); }
        }
    }
    private async Task<string> EditorValueAsync(string expression)
    {
        if (!composeReady) throw new InvalidOperationException("正文编辑器尚未准备好。");
        return System.Text.Json.JsonSerializer.Deserialize<string>(await BodyBox.CoreWebView2.ExecuteScriptAsync(expression)) ?? "";
    }
    private async Task EditorCommandAsync(string command, string? value = null)
    {
        if (!composeReady) return;
        await BodyBox.CoreWebView2.ExecuteScriptAsync($"document.execCommand({System.Text.Json.JsonSerializer.Serialize(command)}, false, {System.Text.Json.JsonSerializer.Serialize(value)}); document.body.focus();");
    }
    private async void CloseCompose_Click(object sender, RoutedEventArgs e)
    {
        if (sending) return;
        string body = composeReady ? await EditorValueAsync("document.body.innerText") : "";
        string html = composeReady ? await EditorValueAsync("document.body.innerHTML") : "";
        if (!string.IsNullOrWhiteSpace(ToBox.Text + CcBox.Text + SubjectBox.Text + body) || html.Contains("<table", StringComparison.OrdinalIgnoreCase))
        {
            var dialog = new ContentDialog { XamlRoot = ComposeWorkspace.XamlRoot, Title = "关闭写信？", Content = "邮件内容尚未保存。关闭后内容会丢失。", PrimaryButtonText = "继续编辑", CloseButtonText = "丢弃" };
            if (await dialog.ShowAsync() == ContentDialogResult.Primary) return;
        }
        await ClearComposeAsync();
    }
    private async Task ClearComposeAsync(bool closeWorkspace = true)
    {
        ToBox.Text = CcBox.Text = SubjectBox.Text = "";
        if (composeReady) await BodyBox.CoreWebView2.ExecuteScriptAsync("document.body.innerHTML = ''; ");
        if (closeWorkspace) { ComposeWorkspace.Visibility = Visibility.Collapsed; EmptyWorkspace.Visibility = Visibility.Visible; }
    }
    private async void FormatBold_Click(object sender, RoutedEventArgs e) => await EditorCommandAsync("bold");
    private async void FontBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (sender is ComboBox box && box.SelectedItem is ComboBoxItem item && item.Tag is string font) await EditorCommandAsync("fontName", font);
    }
    private async void FontSizeBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (sender is ComboBox box && box.SelectedItem is ComboBoxItem item && item.Tag is string size) await EditorCommandAsync("fontSize", size);
    }
    private async void FormatItalic_Click(object sender, RoutedEventArgs e) => await EditorCommandAsync("italic");
    private async void FormatUnderline_Click(object sender, RoutedEventArgs e) => await EditorCommandAsync("underline");
    private async void FormatBullet_Click(object sender, RoutedEventArgs e) => await EditorCommandAsync("insertUnorderedList");
    private async void FormatNumber_Click(object sender, RoutedEventArgs e) => await EditorCommandAsync("insertOrderedList");
    private async void FormatLeft_Click(object sender, RoutedEventArgs e) => await EditorCommandAsync("justifyLeft");
    private async void FormatCenter_Click(object sender, RoutedEventArgs e) => await EditorCommandAsync("justifyCenter");
    private async void FormatRight_Click(object sender, RoutedEventArgs e) => await EditorCommandAsync("justifyRight");
    private void InsertTable_Click(object sender, RoutedEventArgs e)
    {
        var flyout = new Flyout();
        var label = new TextBlock { Text = "选择表格大小（最多 10 × 10）", Margin = new Thickness(0, 0, 0, 8) };
        var grid = new Grid();
        for (int i = 0; i < 10; i++) { grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); }
        var cells = new Button[10, 10];
        void Highlight(int rows, int columns)
        {
            label.Text = $"{columns} 列 × {rows} 行";
            for (int r = 0; r < 10; r++) for (int c = 0; c < 10; c++)
                cells[r, c].Background = new SolidColorBrush(r < rows && c < columns ? Windows.UI.Color.FromArgb(255, 135, 180, 245) : Windows.UI.Color.FromArgb(255, 235, 239, 245));
        }
        for (int r = 0; r < 10; r++) for (int c = 0; c < 10; c++)
        {
            int rows = r + 1, columns = c + 1;
            var cell = new Button { Width = 22, Height = 22, MinWidth = 0, MinHeight = 0, Padding = new Thickness(0), Margin = new Thickness(1), Background = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 235, 239, 245)) };
            Microsoft.UI.Xaml.Automation.AutomationProperties.SetName(cell, $"{columns} 列 {rows} 行");
            cell.PointerEntered += (_, _) => Highlight(rows, columns);
            cell.GotFocus += (_, _) => Highlight(rows, columns);
            cell.Click += async (_, _) =>
            {
                flyout.Hide();
                var html = new System.Text.StringBuilder("<table style='border-collapse:collapse'><tbody>");
                for (int row = 0; row < rows; row++) { html.Append("<tr>"); for (int col = 0; col < columns; col++) html.Append("<td style='border:1px solid #888;padding:6px;min-width:80px'>&nbsp;</td>"); html.Append("</tr>"); }
                html.Append("</tbody></table><p><br></p>");
                await EditorCommandAsync("insertHTML", html.ToString());
            };
            cells[r, c] = cell; Grid.SetRow(cell, r); Grid.SetColumn(cell, c); grid.Children.Add(cell);
        }
        flyout.Content = new StackPanel { Children = { label, grid } };
        flyout.ShowAt(TableButton);
    }
    private async void ComposeAttachment_Click(object sender, RoutedEventArgs e) => await ShowNotice("添加附件", "附件功能尚未接入。");
    private async void ComposeAi_Click(object sender, RoutedEventArgs e) => await ShowNotice("AI 帮我写", "AI 写作功能尚未接入。");
    private async void Send_Click(object sender, RoutedEventArgs e)
    {
        if (sending) return;
        sending = true;
        int statusVersion = ++sendStatusVersion;
        ComposeWorkspace.Visibility = Visibility.Collapsed;
        EmptyWorkspace.Visibility = Visibility.Visible;
        SendStatusText.Text = "邮件发送中";
        SendResultIcon.Visibility = Visibility.Collapsed;
        SendProgress.IsActive = true; SendProgress.Visibility = Visibility.Visible;
        SendStatusBanner.Visibility = Visibility.Visible;
        bool accepted = false;
        try
        {
            string html = await EditorValueAsync("document.body.innerHTML");
            string text = await EditorValueAsync("document.body.innerText");
            await mail.SendAsync(FromBox.SelectedItem as string ?? "", ToBox.Text, CcBox.Text, SubjectBox.Text, "<!doctype html><html><body>" + html + "</body></html>", text, windowLifetime.Token);
            accepted = true;
            await ClearComposeAsync(false);
            SendStatusText.Text = "邮件发送成功";
        }
        catch (Exception ex) { if (!closing) SendStatusText.Text = accepted ? "邮件已发送，但清理编辑区失败，请勿重复发送。" : "发送失败：" + ex.Message + "。点击写邮件可继续编辑。"; }
        finally
        {
            SendProgress.IsActive = false; SendProgress.Visibility = Visibility.Collapsed;
            SendResultIcon.Glyph = accepted ? "\uE73E" : "\uE711";
            SendResultIcon.Foreground = new SolidColorBrush(accepted ? Microsoft.UI.Colors.ForestGreen : Microsoft.UI.Colors.Firebrick);
            SendResultIcon.Visibility = Visibility.Visible;
            sending = false;
            await Task.Delay(TimeSpan.FromSeconds(5));
            if (!closing && statusVersion == sendStatusVersion) SendStatusBanner.Visibility = Visibility.Collapsed;
        }
    }    private async void Settings_Click(object sender, RoutedEventArgs e)
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
            new TextBlock { Text = "配置邮箱账户及服务器。保存时测试 IMAP 连接；发送邮件时使用 SMTP 设置。", TextWrapping = TextWrapping.Wrap },
            address, user, secret,
            new TextBlock { Text = "接收邮件 · IMAP", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, Margin = new Thickness(0, 8, 0, 0) },
            host, port, ssl,
            new TextBlock { Text = "发送邮件 · SMTP", FontWeight = Microsoft.UI.Text.FontWeights.SemiBold, Margin = new Thickness(0, 8, 0, 0) },
            smtpHost, smtpPort, smtpSsl } };
        var ai = decision.Settings;
        var engine = new ComboBox { Header = "AI 服务", HorizontalAlignment = HorizontalAlignment.Stretch, Items = { "LM Studio", "Lemonade Server" }, SelectedIndex = ai.Engine == "Lemonade Server" ? 1 : 0 };
        var aiPort = new TextBox { Header = "端口", Text = ai.Port.ToString() };
        var model = new ComboBox { Header = "判断模型（Decision）", HorizontalAlignment = HorizontalAlignment.Stretch };
        var chatModel = new ComboBox { Header = "生成模型（用于后续总结和写作）", HorizontalAlignment = HorizontalAlignment.Stretch };
        var profile = new TextBox { Header = "自我介绍", Text = ai.Profile ?? "", PlaceholderText = "例如：我是产品经理，关注项目进度、客户反馈和安全问题。", AcceptsReturn = true, TextWrapping = TextWrapping.Wrap, Height = 100 };
        var modelStatus = new TextBlock { TextWrapping = TextWrapping.Wrap };
        var refreshModels = new Button { Content = "刷新模型列表" };
        async Task RefreshModelsAsync()
        {
            if (!int.TryParse(aiPort.Text, out int number) || number < 1 || number > 65535) { modelStatus.Text = "请输入有效端口。"; return; }
            string selectedDecision = model.SelectedItem as string ?? ai.Model;
            string selectedChat = chatModel.SelectedItem as string ?? ai.ChatModel;
            refreshModels.IsEnabled = false;
            try
            {
                modelStatus.Text = "正在获取模型列表…";
                var models = await decision.ListModelsAsync(engine.SelectedItem?.ToString() ?? "LM Studio", number, windowLifetime.Token);
                model.Items.Clear(); chatModel.Items.Clear();
                foreach (var entry in models)
                {
                    if (entry.Type == "decision") model.Items.Add(entry.Id);
                    else if (entry.Type is "llm" or "vlm") chatModel.Items.Add(entry.Id);
                }
                if (model.Items.Contains(selectedDecision)) model.SelectedItem = selectedDecision;
                else if (model.Items.Count > 0) model.SelectedIndex = 0;
                if (chatModel.Items.Contains(selectedChat)) chatModel.SelectedItem = selectedChat;
                else if (chatModel.Items.Count > 0) chatModel.SelectedIndex = 0;
                modelStatus.Text = $"发现 {model.Items.Count} 个判断模型、{chatModel.Items.Count} 个生成模型。";
            }
            catch (Exception ex) { modelStatus.Text = "获取模型失败：" + ex.Message; }
            finally { refreshModels.IsEnabled = true; }
        }
        refreshModels.Click += async (_, _) => await RefreshModelsAsync();
        engine.SelectionChanged += async (_, _) => { aiPort.Text = engine.SelectedIndex == 0 ? "1234" : "13305"; await RefreshModelsAsync(); };
        var enabled = new CheckBox { Content = "为加载的邮件自动判断重要性", IsChecked = ai.Enabled };
        var relabel = new Button { Content = "重新标记当前文件夹已加载的邮件" };
        var aiPanel = new StackPanel { Spacing = 10, Children = { new TextBlock { Text = "AI 设置", FontSize = 20, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold }, new TextBlock { Text = "判断模型用于重要性标记和 AI 搜索。个人介绍帮助 AI 理解你的角色与关注点。", TextWrapping = TextWrapping.Wrap }, engine, aiPort, refreshModels, modelStatus, model, chatModel, profile, enabled, relabel } };
        var navigation = new ListView { Width = 160, SelectionMode = ListViewSelectionMode.Single, Items = { "收发服务器", "AI 设置" }, SelectedIndex = 0 };
        var layout = new Grid { Width = 720, MaxWidth = 720, ColumnSpacing = 24 };
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
        layout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        layout.Children.Add(navigation);
        var scroll = new ScrollViewer { Content = panel, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
        var save = new Button { Content = "测试并保存" };
        navigation.SelectionChanged += async (_, _) => { scroll.Content = navigation.SelectedIndex == 0 ? panel : aiPanel; save.Content = navigation.SelectedIndex == 0 ? "测试并保存" : "保存"; error.Text = ""; if (navigation.SelectedIndex == 1) await RefreshModelsAsync(); };
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
        relabel.Click += async (_, _) =>
        {
            if (!decision.Settings.Enabled) { error.Text = "请先保存并启用 AI 重要性判断。"; return; }
            relabel.IsEnabled = false;
            try
            {
                var items = System.Linq.Enumerable.ToArray(messages);
                for (int i = 0; i < items.Length; i++) { error.Text = $"正在重新标记 {i + 1}/{items.Length}…"; await ClassifyAsync(items[i], true); }
                error.Text = $"重新标记完成，共 {items.Length} 封；失败的邮件会显示“失败”。";
            }
            finally { relabel.IsEnabled = true; }
        };
        cancel.Click += (_, _) => dialog.Hide();
        save.Click += async (_, _) =>
        {
            save.IsEnabled = false; cancel.IsEnabled = false;
            try
            {
                if (navigation.SelectedIndex == 1)
                {
                    if (!int.TryParse(aiPort.Text, out int aiNumber) || aiNumber < 1 || aiNumber > 65535 || (enabled.IsChecked == true && model.SelectedItem is not string)) { error.Text = "请填写有效端口并选择判断模型。"; return; }
                    var settings = new AiSettings(engine.SelectedItem?.ToString() ?? "LM Studio", aiNumber, model.SelectedItem as string ?? "", enabled.IsChecked == true, profile.Text.Trim(), chatModel.SelectedItem as string ?? "");
                    if (settings.Enabled)
                    {
                        error.Text = "正在测试模型…";
                        await decision.ClassifyAsync(settings, new MailItem(0, "test@example.com", "Meeting tomorrow", "", "Please confirm your availability.", true));
                    }
                    decision.Save(settings);
                    foreach (var item in messages) _ = ClassifyAsync(item);
                    dialog.Hide(); return;
                }
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
    private void LoadColumnWidths()
    {
        try
        {
            if (!File.Exists(LayoutPath)) return;
            var widths = System.Text.Json.JsonSerializer.Deserialize<ColumnWidths>(File.ReadAllText(LayoutPath));
            if (widths == null || !double.IsFinite(widths.Sidebar) || !double.IsFinite(widths.Mail) || !double.IsFinite(widths.Ai)) return;
            SidebarColumn.Width = new GridLength(Math.Clamp(widths.Sidebar, 160, 600));
            MailColumn.Width = new GridLength(Math.Clamp(widths.Mail, 250, 800));
            expandedAiWidth = Math.Clamp(widths.Ai, 220, 800);
            AiColumn.Width = new GridLength(expandedAiWidth);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Text.Json.JsonException) { }
    }
    private void SaveColumnWidths()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LayoutPath)!);
            File.WriteAllText(LayoutPath, System.Text.Json.JsonSerializer.Serialize(new ColumnWidths(SidebarColumn.Width.Value, MailColumn.Width.Value, expandedAiWidth)));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }
    private void Splitter_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        if (sender is not FrameworkElement splitter || !e.GetCurrentPoint(splitter).Properties.IsLeftButtonPressed || activeSplitter >= 0) return;
        if (!int.TryParse(splitter.Tag?.ToString(), out int index) || !splitter.CapturePointer(e.Pointer)) return;
        activeSplitter = index; dragPointerId = e.Pointer.PointerId;
        dragStartX = e.GetCurrentPoint(MainColumns).Position.X;
        dragFirstWidth = index == 0 ? SidebarColumn.ActualWidth : index == 1 ? MailColumn.ActualWidth : AiColumn.ActualWidth;
        dragSecondWidth = index == 0 ? MailColumn.ActualWidth : ReadingColumn.ActualWidth;
        e.Handled = true;
    }
    private void Splitter_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (activeSplitter < 0 || e.Pointer.PointerId != dragPointerId) return;
        double delta = e.GetCurrentPoint(MainColumns).Position.X - dragStartX;
        if (activeSplitter == 0)
        {
            double width = Math.Clamp(dragFirstWidth + delta, Math.Max(160, dragFirstWidth + dragSecondWidth - 800), Math.Min(600, Math.Max(160, dragFirstWidth + dragSecondWidth - 250)));
            SidebarColumn.Width = new GridLength(width);
            MailColumn.Width = new GridLength(dragFirstWidth + dragSecondWidth - width);
        }
        else if (activeSplitter == 1)
            MailColumn.Width = new GridLength(Math.Clamp(dragFirstWidth + delta, 250, Math.Min(800, Math.Max(250, dragFirstWidth + dragSecondWidth - 220))));
        else
        {
            expandedAiWidth = Math.Clamp(dragFirstWidth - delta, 220, Math.Min(800, Math.Max(220, dragFirstWidth + dragSecondWidth - 220)));
            AiColumn.Width = new GridLength(expandedAiWidth);
        }
        e.Handled = true;
    }
    private void Splitter_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        if (activeSplitter < 0 || e.Pointer.PointerId != dragPointerId) return;
        activeSplitter = -1;
        (sender as UIElement)?.ReleasePointerCapture(e.Pointer);
        SaveColumnWidths(); e.Handled = true;
    }
    private void Splitter_PointerCaptureLost(object sender, PointerRoutedEventArgs e)
    {
        if (activeSplitter < 0 || e.Pointer.PointerId != dragPointerId) return;
        activeSplitter = -1; SaveColumnWidths();
    }
    private void ToggleAi_Click(object sender, RoutedEventArgs e)
    {
        aiVisible = !aiVisible;
        AiColumn.Width = new GridLength(aiVisible ? expandedAiWidth : 0);
        AiSplitter.Visibility = aiVisible ? Visibility.Visible : Visibility.Collapsed;
        AiPane.Visibility = aiVisible ? Visibility.Visible : Visibility.Collapsed;
        ReadingToolbar.Padding = new Thickness(18, aiVisible ? 14 : 40, 18, 14);
        AiToggleButton.Visibility = aiVisible ? Visibility.Visible : Visibility.Collapsed;
        AiFloatingButton.Visibility = aiVisible ? Visibility.Collapsed : Visibility.Visible;
    }
}

