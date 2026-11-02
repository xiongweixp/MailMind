using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using Windows.Security.Credentials;

namespace MailMind;

public record MailAccount(string Address, string UserName, string ImapHost, int ImapPort, bool StartTls, string SmtpHost = "", int SmtpPort = 465, bool SmtpStartTls = false);
public record MailFolderInfo(string Path, string Name, int Depth, bool Selectable);
public record MailItem(uint Uid, string Sender, string Subject, string Date, string Preview, bool? IsSeen) : System.ComponentModel.INotifyPropertyChanged
{
    public string DisplaySender => InternetAddressList.TryParse(Sender, out var addresses) && addresses.Count > 0
        ? string.Join(", ", addresses.Select(address => !string.IsNullOrWhiteSpace(address.Name) ? address.Name : address is MailboxAddress mailbox ? mailbox.Address : address.ToString()))
        : Sender;
    public string Account { get; set; } = "";
    public string Folder { get; set; } = "";
    public uint Validity { get; set; }
    private bool? hasAttachments;
    public bool? HasAttachments
    {
        get => hasAttachments;
        set { hasAttachments = value; PropertyChanged?.Invoke(this, new(nameof(HasAttachments))); PropertyChanged?.Invoke(this, new(nameof(AttachmentVisibility))); }
    }
    public Microsoft.UI.Xaml.Visibility AttachmentVisibility => HasAttachments == true ? Microsoft.UI.Xaml.Visibility.Visible : Microsoft.UI.Xaml.Visibility.Collapsed;
    private string importance = "";
    public string Importance
    {
        get => importance;
        set
        {
            importance = value;
            PropertyChanged?.Invoke(this, new(nameof(Importance)));
            PropertyChanged?.Invoke(this, new(nameof(ImportanceVisibility)));
            PropertyChanged?.Invoke(this, new(nameof(ImportanceBrush)));
        }
    }
    public Microsoft.UI.Xaml.Visibility ImportanceVisibility => string.IsNullOrEmpty(importance) ? Microsoft.UI.Xaml.Visibility.Collapsed : Microsoft.UI.Xaml.Visibility.Visible;
    public Microsoft.UI.Xaml.Media.Brush ImportanceBrush => new Microsoft.UI.Xaml.Media.SolidColorBrush(importance switch
    {
        "重要" => Windows.UI.Color.FromArgb(255, 190, 49, 49),
        "一般" => Windows.UI.Color.FromArgb(255, 37, 130, 75),
        "可忽略" => Windows.UI.Color.FromArgb(255, 164, 108, 0),
        _ => Windows.UI.Color.FromArgb(255, 105, 116, 132)
    });
    public bool? IsSeen { get; private set; } = IsSeen;
    public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
    public void MarkSeen()
    {
        IsSeen = true;
        PropertyChanged?.Invoke(this, new(nameof(IsSeen)));
        PropertyChanged?.Invoke(this, new(nameof(UnreadVisibility)));
        PropertyChanged?.Invoke(this, new(nameof(MailWeight)));
    }
    public Microsoft.UI.Xaml.Visibility UnreadVisibility => IsSeen == false ? Microsoft.UI.Xaml.Visibility.Visible : Microsoft.UI.Xaml.Visibility.Collapsed;
    public Windows.UI.Text.FontWeight MailWeight => IsSeen == false ? Microsoft.UI.Text.FontWeights.SemiBold : Microsoft.UI.Text.FontWeights.Normal;
}
public record MailBody(string Content, bool IsHtml);
public record MailAttachment(string Name, string ContentType, byte[] Data)
{
    public string SizeLabel => Data.Length < 1024 ? $"{Data.Length} B" : Data.Length < 1024 * 1024 ? $"{Data.Length / 1024.0:F1} KB" : $"{Data.Length / (1024.0 * 1024):F1} MB";
}

public sealed class MailService
{
    private const string VaultResource = "MailMind.IMAP";
    private static readonly string ConfigPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailMind", "account.json");
    private readonly PasswordVault vault = new();
    private readonly MailCache cache = new();
    private uint validity;
    public bool Offline { get; private set; }
    public string? OfflineReason { get; private set; }
    private MailAccount? account;
    private string? password;
    private List<UniqueId> ids = new();
    private string? folderName;
    private bool cachedOnly;
    public MailAccount? Account => account;
    public string? CachedClassification(MailItem item, string fingerprint) => cache.Classification(item, fingerprint);
    public void CacheClassification(MailItem item, string fingerprint, string label) => cache.PutClassification(item, fingerprint, label);
    public List<MailItem> CachedSearchItems(string path)
    {
        var selectedAccount = account ?? throw new InvalidOperationException("请先配置邮箱。");
        var currentValidity = cache.Validity(selectedAccount.Address, path);
        var result = new List<MailItem>();
        foreach (var uid in cache.Uids(selectedAccount.Address, path, currentValidity))
        {
            var item = cache.Get(selectedAccount.Address, path, currentValidity, uid);
            if (item == null) continue;
            SetIdentity(item, selectedAccount.Address, path, currentValidity);
            result.Add(item);
        }
        return result;
    }

    public void Load()
    {
        if (!File.Exists(ConfigPath)) return;
        account = JsonSerializer.Deserialize<MailAccount>(File.ReadAllText(ConfigPath));
        if (account == null) return;
        try { var credential = vault.Retrieve(VaultResource, account.Address); credential.RetrievePassword(); password = credential.Password; }
        catch (Exception) { password = null; }
    }
    public async Task TestAsync(MailAccount candidate, string secret)
    {
        using var client = new ImapClient();
        await ConnectAsync(client, candidate, secret);
        await client.DisconnectAsync(true);
    }
    public void Save(MailAccount candidate, string secret)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(ConfigPath)!);
        vault.Add(new PasswordCredential(VaultResource, candidate.Address, secret));
        File.WriteAllText(ConfigPath, JsonSerializer.Serialize(candidate));
        account = candidate; password = secret; ids.Clear(); folderName = null;
    }
    private static async Task ConnectAsync(ImapClient client, MailAccount a, string secret, System.Threading.CancellationToken cancellationToken = default)
    {
        client.Timeout = 30000;
        try { await client.ConnectAsync(a.ImapHost, a.ImapPort, a.StartTls ? SecureSocketOptions.StartTls : SecureSocketOptions.SslOnConnect, cancellationToken); }
        catch (Exception ex) when (IsTimeout(ex)) { throw new TimeoutException($"连接 IMAP 服务器 {a.ImapHost}:{a.ImapPort} 超时，请检查网络、服务器地址和端口。", ex); }
        try { await client.AuthenticateAsync(a.UserName, secret, cancellationToken); }
        catch (Exception ex) when (IsTimeout(ex)) { throw new TimeoutException("IMAP 登录超时，请检查网络和邮箱服务器状态。", ex); }
    }
    private static bool IsTimeout(Exception ex) => ex is TimeoutException || ex is System.Net.Sockets.SocketException socket && socket.SocketErrorCode == System.Net.Sockets.SocketError.TimedOut || ex.InnerException != null && IsTimeout(ex.InnerException);
    private async Task<ImapClient> ConnectAsync()
    {
        if (account == null || string.IsNullOrEmpty(password)) throw new InvalidOperationException("请先配置邮箱和授权码。");
        var client = new ImapClient();
        try { await ConnectAsync(client, account, password); return client; }
        catch { client.Dispose(); throw; }
    }
    public List<MailFolderInfo> CachedFolders() => account == null ? new() : cache.Folders(account.Address);
    public async Task<IReadOnlyList<MailFolderInfo>> GetFoldersAsync()
    {
        using var client = await ConnectAsync();
        var result = new Dictionary<string, MailFolderInfo>(StringComparer.OrdinalIgnoreCase);
        foreach (var ns in client.PersonalNamespaces)
        {
            IList<IMailFolder> serverFolders;
            try { serverFolders = await client.GetFoldersAsync(ns, StatusItems.None, false); }
            catch (Exception ex) when (IsTimeout(ex)) { throw new TimeoutException("邮箱服务器读取目录列表超时，请稍后点击左侧刷新重试。", ex); }
            foreach (var f in serverFolders)
            {
                var label = (f.Attributes & FolderAttributes.Inbox) != 0 ? "收件箱" :
                    (f.Attributes & FolderAttributes.Junk) != 0 ? "垃圾邮件" :
                    (f.Attributes & FolderAttributes.Sent) != 0 ? "已发送" :
                    (f.Attributes & FolderAttributes.Drafts) != 0 ? "草稿箱" :
                    (f.Attributes & FolderAttributes.Trash) != 0 ? "已删除" : f.Name;
                result[f.FullName] = new(f.FullName, label, f.DirectorySeparator == '\0' ? 0 : f.FullName.Count(c => c == f.DirectorySeparator), (f.Attributes & FolderAttributes.NoSelect) == 0);
            }
        }
        var inbox = client.Inbox;
        if (result.TryGetValue(inbox.FullName, out var inboxInfo)) result[inbox.FullName] = inboxInfo with { Name = "收件箱" };
        else result[inbox.FullName] = new(inbox.FullName, "收件箱", 0, true);
        var folders = result.Values.OrderBy(f => f.Path, StringComparer.OrdinalIgnoreCase).ToList();
        cache.SaveFolders(account!.Address, folders);
        return folders;
    }
    private static async Task<IMailFolder> OpenFolderAsync(ImapClient client, string path)
    {
        var folder = await client.GetFolderAsync(path);
        await folder.OpenAsync(FolderAccess.ReadOnly); return folder;
    }
    public int BeginCachedFolder(string path)
    {
        cachedOnly = true;
        if (account == null) return 0;
        folderName = path;
        validity = cache.Validity(account.Address, path);
        ids = cache.Uids(account.Address, path, validity).Select(uid => new UniqueId(uid)).ToList();
        Offline = false;
        OfflineReason = null;
        return ids.Count;
    }
    public async Task<int> BeginFolderAsync(string path)
    {
        cachedOnly = false;
        folderName = path; Offline = false; OfflineReason = null;
        try
        {
            using var client = await ConnectAsync();
            var folder = await OpenFolderAsync(client, path);
            var found = await folder.SearchAsync(SearchQuery.All);
            // Some servers return a restricted SEARCH result despite reporting more
            // messages in SELECT. Fetch UIDs by sequence number in that case.
            if (found.Count < folder.Count)
            {
                var summaries = await folder.FetchAsync(0, -1, MessageSummaryItems.UniqueId);
                var all = summaries.Select(summary => summary.UniqueId).Where(uid => uid.IsValid).ToList();
                if (all.Count < folder.Count)
                    throw new InvalidDataException($"服务器报告 {folder.Count} 封邮件，但只返回 {all.Count} 个邮件标识，请检查邮箱的 IMAP 同步范围设置。");
                found = all;
            }
            validity = folder.UidValidity;
            cache.SetValidity(account!.Address, path, validity);
            ids = found.Reverse().ToList();
        }
        catch (Exception ex) when (ex is not InvalidDataException)
        {
            validity = cache.Validity(account!.Address, path);
            var cached = cache.Uids(account.Address, path, validity);
            if (validity == 0 || cached.Count == 0) throw;
            ids = cached.Select(uid => new UniqueId(uid)).ToList(); Offline = true; OfflineReason = ex.Message;
        }
        return ids.Count;
    }
    public async IAsyncEnumerable<MailItem?> GetPageAsync(int offset, int count)
    {
        if (folderName == null) yield break;
        var page = ids.Skip(offset).Take(count).ToList();
        if (page.Count == 0) yield break;
        ImapClient? client = null;
        IMailFolder? folder = null;
        try
        {
            foreach (var uid in page)
            {
                var cached = cache.Get(account!.Address, folderName, validity, uid.Id);
                if (cached != null && (cachedOnly || Offline)) { SetIdentity(cached, account!.Address, folderName, validity); yield return cached; continue; }
                client ??= await ConnectAsync();
                folder ??= await OpenFolderAsync(client, folderName);
                if (folder.UidValidity != validity) throw new InvalidOperationException("服务器目录已重建，请重新打开目录。");
                IList<IMessageSummary> summaries;
                try { summaries = await folder.FetchAsync(new[] { uid }, MessageSummaryItems.Envelope | MessageSummaryItems.UniqueId | MessageSummaryItems.PreviewText | MessageSummaryItems.Flags | MessageSummaryItems.BodyStructure); }
                catch (ImapCommandException) { summaries = await folder.FetchAsync(new[] { uid }, MessageSummaryItems.Envelope | MessageSummaryItems.UniqueId | MessageSummaryItems.Flags | MessageSummaryItems.BodyStructure); }
                var summary = summaries.FirstOrDefault();
                var item = summary == null ? null : new MailItem(uid.Id, summary.Envelope?.From.ToString() ?? "未知发件人", summary.Envelope?.Subject ?? "(无主题)", summary.Envelope?.Date?.ToLocalTime().ToString("yyyy-MM-dd HH:mm") ?? "", summary.PreviewText ?? cached?.Preview ?? "", summary.Flags.HasValue ? (summary.Flags.Value & MessageFlags.Seen) != 0 : null);
                if (item != null) { item.HasAttachments = summary!.Body != null ? summary.Attachments.Any() : cached?.HasAttachments; SetIdentity(item, account.Address, folderName, validity); cache.Put(account.Address, folderName, validity, item); cache.PutAttachmentFlag(item); }
                yield return item;
            }
        }
        finally { client?.Dispose(); }
    }
    private static void SetIdentity(MailItem item, string account, string folder, uint validity) { item.Account = account; item.Folder = folder; item.Validity = validity; }
    public async Task<string> GetSearchBodyAsync(MailItem item, System.Threading.CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var cached = cache.Body(item.Account, item.Folder, item.Validity, item.Uid);
        if (cached != null) return cached;
        var html = cache.HtmlBody(item.Account, item.Folder, item.Validity, item.Uid);
        if (html != null) return HtmlTextExtractor.Extract(html);
        var selectedAccount = account;
        if (selectedAccount == null || selectedAccount.Address != item.Account || string.IsNullOrEmpty(password)) throw new InvalidOperationException("无法连接邮件账户获取正文。");
        using var client = new ImapClient();
        await ConnectAsync(client, selectedAccount, password, cancellationToken);
        var selectedFolder = await client.GetFolderAsync(item.Folder, cancellationToken);
        await selectedFolder.OpenAsync(FolderAccess.ReadOnly, cancellationToken);
        if (selectedFolder.UidValidity != item.Validity) throw new InvalidOperationException("服务器目录已重建，请重新打开目录。");
        var message = await selectedFolder.GetMessageAsync(new UniqueId(item.Uid), cancellationToken);
        var text = message.TextBody ?? (message.HtmlBody != null ? HtmlTextExtractor.Extract(message.HtmlBody) : "");
        cache.PutBody(item.Account, item.Folder, item.Validity, item.Uid, text);
        return text;
    }
    public async Task<MailBody> GetBodyAsync(uint uid)
    {
        var selectedAccount = account ?? throw new InvalidOperationException("请先配置邮箱。");
        var path = folderName ?? throw new InvalidOperationException("请先选择目录。");
        var expectedValidity = validity;
        var html = cache.HtmlBody(selectedAccount.Address, path, expectedValidity, uid);
        // Rebuild older HTML caches once so Content-Location images can be resolved too.
        if (html != null && (Offline || cache.HtmlVersion(selectedAccount.Address, path, expectedValidity, uid) >= 2)) return new(html, true);
        var cached = cache.Body(selectedAccount.Address, path, expectedValidity, uid);
        if (Offline && cached != null) return new(cached, false);
        try
        {
        using var client = await ConnectAsync();
        var folder = await OpenFolderAsync(client, path);
        if (folder.UidValidity != expectedValidity) throw new InvalidOperationException("服务器目录已重建，请重新打开目录。");
        var message = await folder.GetMessageAsync(new UniqueId(uid));
        if (message.HtmlBody != null) { var rendered = InlineImageResolver.Resolve(message); cache.PutHtmlBody(selectedAccount.Address, path, expectedValidity, uid, rendered); return new(rendered, true); }
        var body = message.TextBody ?? "(无正文)";
        cache.PutBody(selectedAccount.Address, path, expectedValidity, uid, body);
        return new(body, false);
        }
        catch when (html != null) { return new(html, true); }
        catch when (cached != null) { return new(cached, false); }
    }
    public async Task<IReadOnlyList<MailAttachment>> GetAttachmentsAsync(MailItem item)
    {
        var cached = cache.Attachments(item);
        if (cached != null) { item.HasAttachments = cached.Count > 0; return cached; }
        var selectedAccount = account;
        if (selectedAccount == null || selectedAccount.Address != item.Account || string.IsNullOrEmpty(password)) throw new InvalidOperationException("无法连接邮件账户获取附件。");
        using var client = new ImapClient();
        await ConnectAsync(client, selectedAccount, password);
        var selectedFolder = await OpenFolderAsync(client, item.Folder);
        if (selectedFolder.UidValidity != item.Validity) throw new InvalidOperationException("服务器目录已重建，请重新打开目录。");
        var message = await selectedFolder.GetMessageAsync(new UniqueId(item.Uid));
        var result = new List<MailAttachment>();
        foreach (var entity in message.Attachments)
        {
            using var stream = new MemoryStream();
            string name;
            if (entity is MimePart part) { name = part.FileName ?? "附件"; if (part.Content != null) await part.Content.DecodeToAsync(stream); }
            else if (entity is MessagePart attachedMessage) { name = attachedMessage.ContentDisposition?.FileName ?? "邮件附件.eml"; if (attachedMessage.Message != null) await attachedMessage.Message.WriteToAsync(stream); }
            else continue;
            result.Add(new(name, entity.ContentType.MimeType, stream.ToArray()));
        }
        cache.PutAttachments(item, result);
        item.HasAttachments = result.Count > 0;
        cache.PutAttachmentFlag(item);
        return result;
    }
    public async Task MarkSeenAsync(uint uid)
    {
        var selectedAccount = account ?? throw new InvalidOperationException("请先配置邮箱。");
        var path = folderName ?? throw new InvalidOperationException("请先选择目录。");
        var expectedValidity = validity;
        using var client = await ConnectAsync();
        var selectedFolder = await client.GetFolderAsync(path);
        await selectedFolder.OpenAsync(FolderAccess.ReadWrite);
        if (selectedFolder.Access != FolderAccess.ReadWrite) throw new InvalidOperationException("服务器未允许修改此目录的邮件状态。");
        if (selectedFolder.UidValidity != expectedValidity) throw new InvalidOperationException("服务器目录已重建，请重新打开目录。");
        var id = new UniqueId(uid);
        await selectedFolder.AddFlagsAsync(id, MessageFlags.Seen, true);
        var result = await selectedFolder.FetchAsync(new[] { id }, MessageSummaryItems.UniqueId | MessageSummaryItems.Flags);
        var flags = result.Count == 0 ? null : result[0].Flags;
        if (!flags.HasValue || (flags.GetValueOrDefault() & MessageFlags.Seen) == 0)
            throw new InvalidOperationException("服务器未确认邮件已读，邮件可能已被删除或移动。");
        cache.MarkSeen(selectedAccount.Address, path, expectedValidity, uid);
    }
}
