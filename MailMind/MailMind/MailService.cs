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

public record MailAccount(string Address, string UserName, string ImapHost, int ImapPort, bool StartTls, string SmtpHost = "", int SmtpPort = 465, bool SmtpStartTls = false, string DisplayName = "");
public record MailFolderInfo(string Path, string Name, int Depth, bool Selectable);
public record MailItem(uint Uid, string Sender, string Subject, string Date, string Preview, bool? IsSeen) : System.ComponentModel.INotifyPropertyChanged
{
    private bool bulkMode, bulkHover, bulkSelected;
    public bool IsBulkSelected { get => bulkSelected; set { bulkSelected = value; PropertyChanged?.Invoke(this, new(nameof(IsBulkSelected))); PropertyChanged?.Invoke(this, new(nameof(BulkSelectedVisibility))); } }
    public Microsoft.UI.Xaml.Visibility BulkSelectedVisibility => bulkSelected ? Microsoft.UI.Xaml.Visibility.Visible : Microsoft.UI.Xaml.Visibility.Collapsed;
    public void SetBulkMode(bool value) { bulkMode = value; UpdateBulkVisibility(); }
    public void SetBulkHover(bool value) { bulkHover = value; UpdateBulkVisibility(); }
    private void UpdateBulkVisibility() { PropertyChanged?.Invoke(this, new(nameof(BulkCheckVisibility))); PropertyChanged?.Invoke(this, new(nameof(MailIconVisibility))); }
    public Microsoft.UI.Xaml.Visibility BulkCheckVisibility => bulkMode || bulkHover ? Microsoft.UI.Xaml.Visibility.Visible : Microsoft.UI.Xaml.Visibility.Collapsed;
    public Microsoft.UI.Xaml.Visibility MailIconVisibility => bulkMode || bulkHover ? Microsoft.UI.Xaml.Visibility.Collapsed : Microsoft.UI.Xaml.Visibility.Visible;
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
            PropertyChanged?.Invoke(this, new(nameof(DisplayImportance)));
            PropertyChanged?.Invoke(this, new(nameof(ImportanceVisibility)));
            PropertyChanged?.Invoke(this, new(nameof(ImportanceBrush)));
        }
    }
    public string DisplayImportance => Strings.Importance(importance);
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
        SetSeen(true);
    }
    public void SetSeen(bool? value)
    {
        if (IsSeen == value) return;
        IsSeen = value;
        PropertyChanged?.Invoke(this, new(nameof(IsSeen)));
        PropertyChanged?.Invoke(this, new(nameof(UnreadVisibility)));
        PropertyChanged?.Invoke(this, new(nameof(MailWeight)));
    }
    public Microsoft.UI.Xaml.Visibility UnreadVisibility => IsSeen == false ? Microsoft.UI.Xaml.Visibility.Visible : Microsoft.UI.Xaml.Visibility.Collapsed;
    public Windows.UI.Text.FontWeight MailWeight => IsSeen == false ? Microsoft.UI.Text.FontWeights.SemiBold : Microsoft.UI.Text.FontWeights.Normal;
}
public record MailBody(string Content, bool IsHtml);
public record ReplyDraft(string To, string Cc, string Subject, string Text, string Sender, string Date, string? MessageId, string[] References);
public record ForwardDraft(string Subject, string Text, string Sender, string Date, string To, string Cc);
public record MailAttachment(string Name, string ContentType, byte[] Data)
{
    public string SizeLabel => Data.Length < 1024 ? $"{Data.Length} B" : Data.Length < 1024 * 1024 ? $"{Data.Length / 1024.0:F1} KB" : $"{Data.Length / (1024.0 * 1024):F1} MB";
    public bool IsImage => Path.GetExtension(Name).ToLowerInvariant() is ".jpg" or ".jpeg" or ".png" or ".gif" or ".bmp";
    public bool IsText => Path.GetExtension(Name).ToLowerInvariant() is ".txt" or ".md" or ".markdown";
    private string Kind => Path.GetExtension(Name).ToLowerInvariant() switch
    {
        ".jpg" or ".jpeg" or ".png" or ".gif" or ".bmp" or ".webp" or ".heic" or ".svg" or ".tif" or ".tiff" => "image",
        ".mp4" or ".mov" or ".avi" or ".mkv" or ".wmv" or ".webm" or ".m4v" => "video",
        ".doc" or ".docx" or ".docm" or ".rtf" => "word",
        ".xls" or ".xlsx" or ".xlsm" or ".csv" => "excel",
        ".ppt" or ".pptx" or ".pptm" => "powerpoint",
        ".pdf" => "pdf",
        ".txt" or ".md" or ".markdown" => "text",
        ".cs" or ".c" or ".cpp" or ".h" or ".hpp" or ".py" or ".js" or ".jsx" or ".ts" or ".tsx" or ".java" or ".go" or ".rs" or ".html" or ".css" or ".json" or ".xml" or ".xaml" or ".sql" or ".sh" or ".ps1" => "code",
        ".zip" or ".rar" or ".7z" or ".tar" or ".gz" or ".bz2" or ".xz" => "archive",
        _ => "file"
    };
    public string TypeGlyph => Kind == "video" ? "\uE714" : "\uEB9F";
    public Microsoft.UI.Xaml.Visibility ImageIconVisibility => Kind is "image" or "video" ? Microsoft.UI.Xaml.Visibility.Visible : Microsoft.UI.Xaml.Visibility.Collapsed;
    public Microsoft.UI.Xaml.Visibility TypeLabelVisibility => Kind is "image" or "video" ? Microsoft.UI.Xaml.Visibility.Collapsed : Microsoft.UI.Xaml.Visibility.Visible;
    public string TypeLabel => Kind switch { "word" => "W", "excel" => "X", "powerpoint" => "P", "pdf" => "PDF", "text" => "TXT", "code" => "</>", "archive" => "ZIP", _ => Strings.Get("More_49deaf7da20d") };
    public Microsoft.UI.Xaml.Media.Brush TypeBrush => new Microsoft.UI.Xaml.Media.SolidColorBrush(Kind switch
    {
        "image" => Windows.UI.Color.FromArgb(255, 70, 135, 160), "video" => Windows.UI.Color.FromArgb(255, 148, 90, 140),
        "word" => Windows.UI.Color.FromArgb(255, 43, 87, 154), "excel" => Windows.UI.Color.FromArgb(255, 24, 120, 70),
        "powerpoint" => Windows.UI.Color.FromArgb(255, 196, 72, 44), "pdf" => Windows.UI.Color.FromArgb(255, 190, 48, 48),
        "text" => Windows.UI.Color.FromArgb(255, 90, 123, 151), "code" => Windows.UI.Color.FromArgb(255, 64, 110, 100),
        "archive" => Windows.UI.Color.FromArgb(255, 167, 119, 47), _ => Windows.UI.Color.FromArgb(255, 97, 112, 133)
    });
}

public sealed class MailService
{
    private const string VaultResource = "MailMind.IMAP";
    private static readonly string ConfigPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailMind", "account.json");
    private static readonly string AccountsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailMind", "accounts.json");
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
    private readonly List<MailAccount> accounts = new();
    public IReadOnlyList<MailAccount> Accounts => accounts;
    public string? CachedClassification(MailItem item, string fingerprint) => cache.Classification(item, fingerprint);
    public void CacheClassification(MailItem item, string fingerprint, string label) => cache.PutClassification(item, fingerprint, label);
    public List<MailItem> CachedSearchItems(string path)
    {
        var selectedAccount = account ?? throw new InvalidOperationException(Strings.Get("More_5ae31f43f796"));
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
    public List<MailItem> SearchAgentMessages(string? folder, string query, string sender, string from, string to, int offset, int limit)
    {
        var selected = account ?? throw new InvalidOperationException(Strings.Get("More_5ae31f43f796"));
        return cache.SearchAgentMessages(selected.Address, folder, query, sender, from, to, offset, limit);
    }
    public MailItem? GetCachedAgentMessage(MailItem item)
    {
        if (!string.Equals(account?.Address, item.Account, StringComparison.OrdinalIgnoreCase) || cache.Validity(item.Account, item.Folder) != item.Validity) return null;
        var current = cache.Get(item.Account, item.Folder, item.Validity, item.Uid);
        if (current != null) SetIdentity(current, item.Account, item.Folder, item.Validity);
        return current;
    }

    public void Load()
    {
        if (File.Exists(AccountsPath)) accounts.AddRange(JsonSerializer.Deserialize<List<MailAccount>>(File.ReadAllText(AccountsPath)) ?? new());
        if (File.Exists(ConfigPath)) account = JsonSerializer.Deserialize<MailAccount>(File.ReadAllText(ConfigPath));
        if (account == null) return;
        if (!accounts.Any(a => string.Equals(a.Address, account.Address, StringComparison.OrdinalIgnoreCase))) accounts.Add(account);
        try { var credential = vault.Retrieve(VaultResource, account.Address); credential.RetrievePassword(); password = credential.Password; }
        catch (Exception) { password = null; }
    }
    public async Task TestAsync(MailAccount candidate, string secret)
    {
        using var client = new ImapClient();
        await ConnectAsync(client, candidate, secret);
        await client.DisconnectAsync(true);
    }
    public async Task TestSmtpAsync(MailAccount candidate, string secret)
    {
        using var client = new MailKit.Net.Smtp.SmtpClient { Timeout = 30000 };
        await client.ConnectAsync(candidate.SmtpHost, candidate.SmtpPort, candidate.SmtpStartTls ? SecureSocketOptions.StartTls : SecureSocketOptions.SslOnConnect);
        await client.AuthenticateAsync(candidate.UserName, secret);
        await client.DisconnectAsync(true);
    }
    public string StoredPassword(string address)
    {
        try { var credential = vault.Retrieve(VaultResource, address); credential.RetrievePassword(); return credential.Password; }
        catch (Exception) { return ""; }
    }
    public void Save(MailAccount candidate, string secret)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(ConfigPath)!);
        vault.Add(new PasswordCredential(VaultResource, candidate.Address, secret));
        accounts.RemoveAll(a => string.Equals(a.Address, candidate.Address, StringComparison.OrdinalIgnoreCase));
        accounts.Add(candidate);
        File.WriteAllText(AccountsPath, JsonSerializer.Serialize(accounts));
        File.WriteAllText(ConfigPath, JsonSerializer.Serialize(candidate));
        account = candidate; password = secret; ids.Clear(); folderName = null;
    }
    public async Task SendAsync(string from, string to, string cc, string subject, string html, string text, System.Threading.CancellationToken token, ReplyDraft? reply = null, System.Collections.Generic.IReadOnlyList<MailAttachment>? attachments = null)
    {
        var selected = account ?? throw new InvalidOperationException(Strings.Get("More_5ae31f43f796"));
        if (from != selected.Address || string.IsNullOrEmpty(password)) throw new InvalidOperationException(Strings.Get("More_c4bfac5f682f"));
        if (string.IsNullOrWhiteSpace(selected.SmtpHost) || selected.SmtpPort is < 1 or > 65535) throw new InvalidOperationException(Strings.Get("More_2c207ab6b221"));
        var message = new MimeMessage();
        message.From.Add(new MailboxAddress(selected.DisplayName, selected.Address));
        if (!InternetAddressList.TryParse(to, out var recipients) || recipients.Mailboxes.Count() == 0) throw new InvalidOperationException(Strings.Get("More_6a66c747edcb"));
        message.To.AddRange(recipients);
        if (!string.IsNullOrWhiteSpace(cc))
        {
            if (!InternetAddressList.TryParse(cc, out var copies) || copies.Mailboxes.Count() == 0) throw new InvalidOperationException(Strings.Get("More_87b629422272"));
            message.Cc.AddRange(copies);
        }
        message.Subject = subject;
        if (reply?.MessageId is string messageId && !string.IsNullOrWhiteSpace(messageId))
        {
            message.InReplyTo = messageId;
            foreach (var reference in reply.References) message.References.Add(reference);
            if (!message.References.Contains(messageId)) message.References.Add(messageId);
        }
        var builder = new BodyBuilder { HtmlBody = html, TextBody = text };
        if (attachments != null) foreach (var attachment in attachments) builder.Attachments.Add(attachment.Name, attachment.Data);
        message.Body = builder.ToMessageBody();
        using var client = new MailKit.Net.Smtp.SmtpClient { Timeout = 30000 };
        await client.ConnectAsync(selected.SmtpHost, selected.SmtpPort, selected.SmtpStartTls ? SecureSocketOptions.StartTls : SecureSocketOptions.SslOnConnect, token);
        await client.AuthenticateAsync(selected.UserName, password, token);
        await client.SendAsync(message, token);
        try { await client.DisconnectAsync(true, token); } catch (Exception) { /* The server already accepted the message. */ }
    }
    private static async Task ConnectAsync(ImapClient client, MailAccount a, string secret, System.Threading.CancellationToken cancellationToken = default)
    {
        client.Timeout = 30000;
        try { await client.ConnectAsync(a.ImapHost, a.ImapPort, a.StartTls ? SecureSocketOptions.StartTls : SecureSocketOptions.SslOnConnect, cancellationToken); }
        catch (Exception ex) when (IsTimeout(ex)) { throw new TimeoutException(string.Format(Strings.Get("Fmt_9e6ea7a754b7"), a.ImapHost, a.ImapPort), ex); }
        try { await client.AuthenticateAsync(a.UserName, secret, cancellationToken); }
        catch (Exception ex) when (IsTimeout(ex)) { throw new TimeoutException(Strings.Get("More_8b42227ca4a1"), ex); }
    }
    private static bool IsTimeout(Exception ex) => ex is TimeoutException || ex is System.Net.Sockets.SocketException socket && socket.SocketErrorCode == System.Net.Sockets.SocketError.TimedOut || ex.InnerException != null && IsTimeout(ex.InnerException);
    private async Task<ImapClient> ConnectAsync()
    {
        if (account == null || string.IsNullOrEmpty(password)) throw new InvalidOperationException(Strings.Get("More_9cf8aaefb0c2"));
        var client = new ImapClient();
        try { await ConnectAsync(client, account, password); return client; }
        catch { client.Dispose(); throw; }
    }
    public List<MailFolderInfo> CachedFolders() => account == null ? new() : cache.Folders(account.Address);
    public async Task<bool> SyncFolderAsync(string path, System.Threading.CancellationToken token, Action<int, int>? progress = null)
    {
        var selected = account ?? throw new InvalidOperationException("No mail account configured.");
        var secret = password ?? throw new InvalidOperationException("No mail password configured.");
        using var client = new ImapClient();
        await ConnectAsync(client, selected, secret, token);
        var remote = await client.GetFolderAsync(path, token);
        await remote.OpenAsync(FolderAccess.ReadOnly, token);
        var found = await remote.SearchAsync(SearchQuery.All, token);
        if (found.Count < remote.Count)
        {
            var all = await remote.FetchAsync(0, -1, MessageSummaryItems.UniqueId, token);
            found = all.Select(s => s.UniqueId).Where(id => id.IsValid).ToList();
            if (found.Count < remote.Count) throw new InvalidDataException("The server did not return all message IDs.");
        }
        uint currentValidity = remote.UidValidity;
        if (cache.Validity(selected.Address, path) != currentValidity)
            cache.SetValidity(selected.Address, path, currentValidity);
        var complete = cache.CompleteUids(selected.Address, path, currentValidity);
        var existing = found.Where(id => complete.Contains(id.Id)).OrderBy(id => id.Id).ToList();
        foreach (var batch in existing.Chunk(500))
        {
            var flags = await remote.FetchAsync(batch, MessageSummaryItems.UniqueId | MessageSummaryItems.Flags, token);
            if (flags.Count != batch.Length) throw new InvalidDataException("The server message list changed during synchronization.");
            cache.UpdateSeen(selected.Address, path, currentValidity, flags.Where(s => s.Flags.HasValue).Select(s => (s.UniqueId.Id, (s.Flags!.Value & MessageFlags.Seen) != 0)));
        }
        var missing = found.Where(id => !complete.Contains(id.Id)).OrderByDescending(id => id.Id).ToList();
        int completed = found.Count - missing.Count;
        foreach (var batch in missing.Take(10).Chunk(10))
        {
            token.ThrowIfCancellationRequested();
            progress?.Invoke(completed + 1, found.Count);
            Array.Sort(batch, (left, right) => left.Id.CompareTo(right.Id));
            IList<IMessageSummary> summaries;
            try { summaries = await remote.FetchAsync(batch, MessageSummaryItems.Envelope | MessageSummaryItems.UniqueId | MessageSummaryItems.PreviewText | MessageSummaryItems.Flags | MessageSummaryItems.BodyStructure, token); }
            catch (ImapCommandException) { summaries = await remote.FetchAsync(batch, MessageSummaryItems.Envelope | MessageSummaryItems.UniqueId | MessageSummaryItems.Flags | MessageSummaryItems.BodyStructure, token); }
            if (summaries.Count != batch.Length) throw new InvalidDataException("The server message list changed during synchronization.");
            foreach (var summary in summaries)
            {
                progress?.Invoke(++completed, found.Count);
                var uid = summary.UniqueId.Id;
                var previous = cache.Get(selected.Address, path, currentValidity, uid);
                var item = new MailItem(uid, summary.Envelope?.From.ToString() ?? "", summary.Envelope?.Subject ?? "", summary.Envelope?.Date?.ToLocalTime().ToString("yyyy-MM-dd HH:mm") ?? "", summary.PreviewText ?? previous?.Preview ?? "", summary.Flags.HasValue ? (summary.Flags.Value & MessageFlags.Seen) != 0 : null);
                SetIdentity(item, selected.Address, path, currentValidity);
                item.HasAttachments = summary.Body != null ? summary.Attachments.Any() : previous?.HasAttachments;
                cache.Put(selected.Address, path, currentValidity, item);
                if (summary.Envelope != null) cache.PutRecipients(item, summary.Envelope.To?.ToString() ?? "", summary.Envelope.Cc?.ToString() ?? "");
                cache.PutAttachmentFlag(item);
                if ((cache.HtmlVersion(selected.Address, path, currentValidity, uid) >= 2 || cache.Body(selected.Address, path, currentValidity, uid) != null) && cache.HasAttachmentSet(item)) continue;
                var message = await remote.GetMessageAsync(summary.UniqueId, token);
                if (message.HtmlBody != null) cache.PutHtmlBody(selected.Address, path, currentValidity, uid, InlineImageResolver.Resolve(message));
                else cache.PutBody(selected.Address, path, currentValidity, uid, message.TextBody ?? "");
                var attachments = new List<MailAttachment>();
                foreach (var entity in message.Attachments)
                {
                    using var stream = new MemoryStream();
                    string name;
                    if (entity is MimePart part) { name = part.FileName ?? "attachment"; if (part.Content != null) await part.Content.DecodeToAsync(stream, token); }
                    else if (entity is MessagePart nested) { name = nested.ContentDisposition?.FileName ?? "message.eml"; if (nested.Message != null) await nested.Message.WriteToAsync(stream, token); }
                    else continue;
                    attachments.Add(new(name, entity.ContentType.MimeType, stream.ToArray()));
                }
                cache.PutAttachments(item, attachments);
                item.HasAttachments = attachments.Count > 0; cache.PutAttachmentFlag(item);
            }
        }
        token.ThrowIfCancellationRequested();
        if (!ReferenceEquals(account, selected)) return false;
        var verified = await remote.SearchAsync(SearchQuery.All, token);
        if (verified.Count < remote.Count)
        {
            var all = await remote.FetchAsync(0, -1, MessageSummaryItems.UniqueId, token);
            verified = all.Select(s => s.UniqueId).Where(id => id.IsValid).ToList();
        }
        if (verified.Count != remote.Count || !verified.OrderBy(id => id.Id).SequenceEqual(found.OrderBy(id => id.Id)))
            return true;
        cache.RemoveMissing(selected.Address, path, currentValidity, found.Select(id => id.Id).ToArray());
        return missing.Count > 10;
    }
    public async Task<IReadOnlyList<MailFolderInfo>> GetFoldersAsync()
    {
        using var client = await ConnectAsync();
        var result = new Dictionary<string, MailFolderInfo>(StringComparer.OrdinalIgnoreCase);
        foreach (var ns in client.PersonalNamespaces)
        {
            IList<IMailFolder> serverFolders;
            try { serverFolders = await client.GetFoldersAsync(ns, StatusItems.None, false); }
            catch (Exception ex) when (IsTimeout(ex)) { throw new TimeoutException(Strings.Get("More_7591f45427c7"), ex); }
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
                    throw new InvalidDataException(string.Format(Strings.Get("Fmt_ce6690dbc149"), folder.Count, all.Count));
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
    public async Task<bool> FolderChangedAsync(string path)
    {
        using var client = await ConnectAsync();
        var folder = await OpenFolderAsync(client, path);
        if (folder.UidValidity != validity || folder.Count != ids.Count) return true;
        var found = await folder.SearchAsync(SearchQuery.All);
        if (found.Count != folder.Count) return true;
        return !found.Reverse().SequenceEqual(ids);
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
                if (folder.UidValidity != validity) throw new InvalidOperationException(Strings.Get("More_1ab279dc6512"));
                IList<IMessageSummary> summaries;
                try { summaries = await folder.FetchAsync(new[] { uid }, MessageSummaryItems.Envelope | MessageSummaryItems.UniqueId | MessageSummaryItems.PreviewText | MessageSummaryItems.Flags | MessageSummaryItems.BodyStructure); }
                catch (ImapCommandException) { summaries = await folder.FetchAsync(new[] { uid }, MessageSummaryItems.Envelope | MessageSummaryItems.UniqueId | MessageSummaryItems.Flags | MessageSummaryItems.BodyStructure); }
                var summary = summaries.FirstOrDefault();
                var item = summary == null ? null : new MailItem(uid.Id, summary.Envelope?.From.ToString() ?? Strings.Get("More_1cc645d7a525"), summary.Envelope?.Subject ?? Strings.Get("More_dd6026e30d48"), summary.Envelope?.Date?.ToLocalTime().ToString("yyyy-MM-dd HH:mm") ?? "", summary.PreviewText ?? cached?.Preview ?? "", summary.Flags.HasValue ? (summary.Flags.Value & MessageFlags.Seen) != 0 : null);
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
        if (selectedAccount == null || selectedAccount.Address != item.Account || string.IsNullOrEmpty(password)) throw new InvalidOperationException(Strings.Get("More_10bee1228539"));
        using var client = new ImapClient();
        await ConnectAsync(client, selectedAccount, password, cancellationToken);
        var selectedFolder = await client.GetFolderAsync(item.Folder, cancellationToken);
        await selectedFolder.OpenAsync(FolderAccess.ReadOnly, cancellationToken);
        if (selectedFolder.UidValidity != item.Validity) throw new InvalidOperationException(Strings.Get("More_1ab279dc6512"));
        var message = await selectedFolder.GetMessageAsync(new UniqueId(item.Uid), cancellationToken);
        var text = message.TextBody ?? (message.HtmlBody != null ? HtmlTextExtractor.Extract(message.HtmlBody) : "");
        cache.PutBody(item.Account, item.Folder, item.Validity, item.Uid, text);
        return text;
    }
    public async Task<ReplyDraft> GetReplyDraftAsync(MailItem item, bool replyAll, System.Threading.CancellationToken cancellationToken)
    {
        var selected = account;
        if (selected == null || selected.Address != item.Account || string.IsNullOrEmpty(password)) throw new InvalidOperationException(Strings.Get("More_76be3e92a26f"));
        using var client = new ImapClient();
        await ConnectAsync(client, selected, password, cancellationToken);
        var selectedFolder = await client.GetFolderAsync(item.Folder, cancellationToken);
        await selectedFolder.OpenAsync(FolderAccess.ReadOnly, cancellationToken);
        if (selectedFolder.UidValidity != item.Validity) throw new InvalidOperationException(Strings.Get("More_1ab279dc6512"));
        var original = await selectedFolder.GetMessageAsync(new UniqueId(item.Uid), cancellationToken);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { selected.Address };
        var to = new InternetAddressList();
        var cc = new InternetAddressList();
        void Add(InternetAddressList target, IEnumerable<MailboxAddress> addresses)
        {
            foreach (var address in addresses) if (seen.Add(address.Address)) target.Add(address);
        }
        Add(to, (original.ReplyTo.Count > 0 ? original.ReplyTo : original.From).Mailboxes);
        if (replyAll) { Add(to, original.To.Mailboxes); Add(cc, original.Cc.Mailboxes); }
        if (to.Count == 0) Add(to, original.To.Mailboxes);
        if (to.Count == 0) throw new InvalidOperationException(Strings.Get("More_b8cb72e2c4e2"));
        var subject = original.Subject ?? "";
        if (!subject.StartsWith("Re:", StringComparison.OrdinalIgnoreCase)) subject = "Re: " + subject;
        var text = original.TextBody ?? (original.HtmlBody != null ? HtmlTextExtractor.Extract(original.HtmlBody) : "");
        return new(to.ToString(), cc.ToString(), subject, text, original.From.ToString(), original.Date.ToLocalTime().ToString("g"), original.MessageId, original.References.ToArray());
    }
    public async Task<ForwardDraft> GetForwardDraftAsync(MailItem item, System.Threading.CancellationToken cancellationToken)
    {
        var selected = account;
        if (selected == null || selected.Address != item.Account || string.IsNullOrEmpty(password)) throw new InvalidOperationException(Strings.Get("ForwardConnectError"));
        using var client = new ImapClient();
        await ConnectAsync(client, selected, password, cancellationToken);
        var source = await client.GetFolderAsync(item.Folder, cancellationToken);
        await source.OpenAsync(FolderAccess.ReadOnly, cancellationToken);
        if (source.UidValidity != item.Validity) throw new InvalidOperationException(Strings.Get("MessageFolderRebuilt"));
        var original = await source.GetMessageAsync(new UniqueId(item.Uid), cancellationToken);
        var subject = original.Subject ?? "";
        if (!subject.StartsWith("Fwd:", StringComparison.OrdinalIgnoreCase)) subject = "Fwd: " + subject;
        return new(subject, original.TextBody ?? (original.HtmlBody != null ? HtmlTextExtractor.Extract(original.HtmlBody) : ""), original.From.ToString(), original.Date.ToLocalTime().ToString("g"), original.To.ToString(), original.Cc.ToString());
    }
    public async Task DeleteAsync(MailItem item, System.Threading.CancellationToken cancellationToken)
    {
        var selected = account;
        if (selected == null || selected.Address != item.Account || string.IsNullOrEmpty(password)) throw new InvalidOperationException(Strings.Get("DeleteConnectError"));
        using var client = new ImapClient();
        await ConnectAsync(client, selected, password, cancellationToken);
        var source = await client.GetFolderAsync(item.Folder, cancellationToken);
        await source.OpenAsync(FolderAccess.ReadWrite, cancellationToken);
        if (source.Access != FolderAccess.ReadWrite) throw new InvalidOperationException(Strings.Get("MessageFolderReadOnly"));
        if (source.UidValidity != item.Validity) throw new InvalidOperationException(Strings.Get("MessageFolderRebuilt"));
        var uid = new UniqueId(item.Uid);
        var found = await source.FetchAsync(new[] { uid }, MessageSummaryItems.UniqueId, cancellationToken);
        if (found.Count == 0) throw new InvalidOperationException(Strings.Get("MessageMovedOrDeleted"));
        var trash = client.GetFolder(SpecialFolder.Trash);
        if (trash != null && trash.FullName != source.FullName)
        {
            if ((client.Capabilities & (ImapCapabilities.Move | ImapCapabilities.UidPlus)) == 0) throw new InvalidOperationException(Strings.Get("MessageMoveUnsupported"));
            await source.MoveToAsync(uid, trash, cancellationToken);
        }
        else
        {
            if ((client.Capabilities & ImapCapabilities.UidPlus) == 0) throw new InvalidOperationException(Strings.Get("MessageDeleteUnsupported"));
            await source.AddFlagsAsync(uid, MessageFlags.Deleted, true, cancellationToken);
            await source.ExpungeAsync(new[] { uid }, cancellationToken);
        }
        cache.Remove(item);
        if (folderName == item.Folder && validity == item.Validity) ids.Remove(uid);
    }
    public async Task SetSeenAsync(MailItem item, bool seen, System.Threading.CancellationToken token)
    {
        if (account == null || account.Address != item.Account || string.IsNullOrEmpty(password)) throw new InvalidOperationException(Strings.Get("DeleteConnectError"));
        using var client = new ImapClient();
        await ConnectAsync(client, account, password, token);
        var source = await client.GetFolderAsync(item.Folder, token);
        await source.OpenAsync(FolderAccess.ReadWrite, token);
        if (source.Access != FolderAccess.ReadWrite) throw new InvalidOperationException(Strings.Get("MessageFolderReadOnly"));
        if (source.UidValidity != item.Validity) throw new InvalidOperationException(Strings.Get("MessageFolderRebuilt"));
        var uid = new UniqueId(item.Uid);
        if (seen) await source.AddFlagsAsync(uid, MessageFlags.Seen, true, token);
        else await source.RemoveFlagsAsync(uid, MessageFlags.Seen, true, token);
        var result = await source.FetchAsync(new[] { uid }, MessageSummaryItems.UniqueId | MessageSummaryItems.Flags, token);
        var flags = result.Count == 0 ? null : result[0].Flags;
        if (!flags.HasValue || flags.GetValueOrDefault().HasFlag(MessageFlags.Seen) != seen) throw new InvalidOperationException(Strings.Get("MessageMovedOrDeleted"));
        cache.SetSeen(item, seen);
    }
    public async Task MoveAsync(MailItem item, string destination, System.Threading.CancellationToken token)
    {
        if (account == null || account.Address != item.Account || string.IsNullOrEmpty(password)) throw new InvalidOperationException(Strings.Get("DeleteConnectError"));
        using var client = new ImapClient();
        await ConnectAsync(client, account, password, token);
        var source = await client.GetFolderAsync(item.Folder, token);
        await source.OpenAsync(FolderAccess.ReadWrite, token);
        if (source.Access != FolderAccess.ReadWrite) throw new InvalidOperationException(Strings.Get("MessageFolderReadOnly"));
        if (source.UidValidity != item.Validity) throw new InvalidOperationException(Strings.Get("MessageFolderRebuilt"));
        var uid = new UniqueId(item.Uid);
        if ((await source.FetchAsync(new[] { uid }, MessageSummaryItems.UniqueId, token)).Count == 0) throw new InvalidOperationException(Strings.Get("MessageMovedOrDeleted"));
        if ((client.Capabilities & (ImapCapabilities.Move | ImapCapabilities.UidPlus)) == 0) throw new InvalidOperationException(Strings.Get("MessageMoveUnsupported"));
        await source.MoveToAsync(uid, await client.GetFolderAsync(destination, token), token);
        cache.Remove(item);
        if (folderName == item.Folder && validity == item.Validity) ids.Remove(uid);
    }
    public async Task<(string To, string Cc)> GetRecipientsAsync(MailItem item)
    {
        var cached = cache.Recipients(item);
        if (cached.HasValue) return cached.Value;
        if (account == null || account.Address != item.Account || string.IsNullOrEmpty(password)) throw new InvalidOperationException(Strings.Get("More_96055ce0ed7d"));
        using var client = new ImapClient();
        await ConnectAsync(client, account, password);
        var selectedFolder = await OpenFolderAsync(client, item.Folder);
        if (selectedFolder.UidValidity != item.Validity) throw new InvalidOperationException(Strings.Get("More_1ab279dc6512"));
        var summaries = await selectedFolder.FetchAsync(new[] { new UniqueId(item.Uid) }, MessageSummaryItems.Envelope);
        var envelope = summaries.FirstOrDefault()?.Envelope;
        if (envelope == null) throw new InvalidDataException("The server did not return the message envelope.");
        var result = (envelope.To?.ToString() ?? "", envelope.Cc?.ToString() ?? "");
        cache.PutRecipients(item, result.Item1, result.Item2);
        return result;
    }
    public async Task<MailBody> GetBodyAsync(uint uid)
    {
        var selectedAccount = account ?? throw new InvalidOperationException(Strings.Get("More_5ae31f43f796"));
        var path = folderName ?? throw new InvalidOperationException(Strings.Get("More_e0702ff2cb5e"));
        var expectedValidity = validity;
        var html = cache.HtmlBody(selectedAccount.Address, path, expectedValidity, uid);
        // Rebuild older HTML caches once so Content-Location images can be resolved too.
        if (html != null && (Offline || cache.HtmlVersion(selectedAccount.Address, path, expectedValidity, uid) >= 2)) return new(html, true);
        var cached = cache.Body(selectedAccount.Address, path, expectedValidity, uid);
        if (cached != null) return new(cached, false);
        try
        {
        using var client = await ConnectAsync();
        var folder = await OpenFolderAsync(client, path);
        if (folder.UidValidity != expectedValidity) throw new InvalidOperationException(Strings.Get("More_1ab279dc6512"));
        var message = await folder.GetMessageAsync(new UniqueId(uid));
        if (message.HtmlBody != null) { var rendered = InlineImageResolver.Resolve(message); cache.PutHtmlBody(selectedAccount.Address, path, expectedValidity, uid, rendered); return new(rendered, true); }
        var body = message.TextBody ?? Strings.Get("More_38ca211767ce");
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
        if (selectedAccount == null || selectedAccount.Address != item.Account || string.IsNullOrEmpty(password)) throw new InvalidOperationException(Strings.Get("More_2c067f642cdc"));
        using var client = new ImapClient();
        await ConnectAsync(client, selectedAccount, password);
        var selectedFolder = await OpenFolderAsync(client, item.Folder);
        if (selectedFolder.UidValidity != item.Validity) throw new InvalidOperationException(Strings.Get("More_1ab279dc6512"));
        var message = await selectedFolder.GetMessageAsync(new UniqueId(item.Uid));
        var result = new List<MailAttachment>();
        foreach (var entity in message.Attachments)
        {
            using var stream = new MemoryStream();
            string name;
            if (entity is MimePart part) { name = part.FileName ?? Strings.Get("More_99f6fe6c41ad"); if (part.Content != null) await part.Content.DecodeToAsync(stream); }
            else if (entity is MessagePart attachedMessage) { name = attachedMessage.ContentDisposition?.FileName ?? Strings.Get("More_81592e481edb"); if (attachedMessage.Message != null) await attachedMessage.Message.WriteToAsync(stream); }
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
        var selectedAccount = account ?? throw new InvalidOperationException(Strings.Get("More_5ae31f43f796"));
        var path = folderName ?? throw new InvalidOperationException(Strings.Get("More_e0702ff2cb5e"));
        var expectedValidity = validity;
        using var client = await ConnectAsync();
        var selectedFolder = await client.GetFolderAsync(path);
        await selectedFolder.OpenAsync(FolderAccess.ReadWrite);
        if (selectedFolder.Access != FolderAccess.ReadWrite) throw new InvalidOperationException(Strings.Get("More_e8dc9306cda8"));
        if (selectedFolder.UidValidity != expectedValidity) throw new InvalidOperationException(Strings.Get("More_1ab279dc6512"));
        var id = new UniqueId(uid);
        await selectedFolder.AddFlagsAsync(id, MessageFlags.Seen, true);
        var result = await selectedFolder.FetchAsync(new[] { id }, MessageSummaryItems.UniqueId | MessageSummaryItems.Flags);
        var flags = result.Count == 0 ? null : result[0].Flags;
        if (!flags.HasValue || (flags.GetValueOrDefault() & MessageFlags.Seen) == 0)
            throw new InvalidOperationException(Strings.Get("More_791c29dbb42e"));
        cache.MarkSeen(selectedAccount.Address, path, expectedValidity, uid);
    }
}
