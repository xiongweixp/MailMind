using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Data.Sqlite;

namespace MailMind;

// Cached UIDs are scoped by account, folder and UIDVALIDITY.
public sealed class MailCache
{
    private readonly string connectionString;
    public MailCache()
    {
        var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MailMind");
        Directory.CreateDirectory(dir);
        connectionString = new SqliteConnectionStringBuilder { DataSource = Path.Combine(dir, "mailmind.db") }.ToString();
        using var db = Open();
        using var cmd = db.CreateCommand();
        cmd.CommandText = @"PRAGMA journal_mode=WAL;
CREATE TABLE IF NOT EXISTS folders (account TEXT NOT NULL, path TEXT NOT NULL, name TEXT NOT NULL, depth INTEGER NOT NULL, selectable INTEGER NOT NULL, validity INTEGER NOT NULL DEFAULT 0, PRIMARY KEY(account,path));
CREATE TABLE IF NOT EXISTS messages (account TEXT NOT NULL, folder TEXT NOT NULL, validity INTEGER NOT NULL, uid INTEGER NOT NULL, sender TEXT NOT NULL, subject TEXT NOT NULL, date TEXT NOT NULL, preview TEXT NOT NULL, body TEXT, PRIMARY KEY(account,folder,validity,uid));
CREATE TABLE IF NOT EXISTS classifications (account TEXT NOT NULL, folder TEXT NOT NULL, validity INTEGER NOT NULL, uid INTEGER NOT NULL, fingerprint TEXT NOT NULL, label TEXT NOT NULL, PRIMARY KEY(account,folder,validity,uid,fingerprint));";
        cmd.ExecuteNonQuery();
        cmd.CommandText = @"CREATE TABLE IF NOT EXISTS attachment_sets (account TEXT NOT NULL, folder TEXT NOT NULL, validity INTEGER NOT NULL, uid INTEGER NOT NULL, PRIMARY KEY(account,folder,validity,uid));
CREATE TABLE IF NOT EXISTS attachments (account TEXT NOT NULL, folder TEXT NOT NULL, validity INTEGER NOT NULL, uid INTEGER NOT NULL, position INTEGER NOT NULL, name TEXT NOT NULL, content_type TEXT NOT NULL, data BLOB NOT NULL, PRIMARY KEY(account,folder,validity,uid,position));";
        cmd.ExecuteNonQuery();
        using var migration = db.CreateCommand();
        migration.CommandText = "PRAGMA table_info(messages)";
        using var columns = migration.ExecuteReader();
        bool hasHtml = false, hasSeen = false, hasHtmlVersion = false;
        while (columns.Read()) { if (columns.GetString(1) == "html_body") hasHtml = true; if (columns.GetString(1) == "is_seen") hasSeen = true; if (columns.GetString(1) == "html_version") hasHtmlVersion = true; }
        columns.Close();
        if (!hasHtml) { migration.CommandText = "ALTER TABLE messages ADD COLUMN html_body TEXT"; migration.ExecuteNonQuery(); }
        if (!hasSeen) { migration.CommandText = "ALTER TABLE messages ADD COLUMN is_seen INTEGER"; migration.ExecuteNonQuery(); }
        if (!hasHtmlVersion) { migration.CommandText = "ALTER TABLE messages ADD COLUMN html_version INTEGER NOT NULL DEFAULT 0"; migration.ExecuteNonQuery(); }
        migration.CommandText = "CREATE TABLE IF NOT EXISTS attachment_flags (account TEXT NOT NULL, folder TEXT NOT NULL, validity INTEGER NOT NULL, uid INTEGER NOT NULL, has_attachments INTEGER NOT NULL, PRIMARY KEY(account,folder,validity,uid))"; migration.ExecuteNonQuery();
        migration.CommandText = "INSERT OR IGNORE INTO attachment_flags SELECT s.account,s.folder,s.validity,s.uid,EXISTS(SELECT 1 FROM attachments a WHERE a.account=s.account AND a.folder=s.folder AND a.validity=s.validity AND a.uid=s.uid) FROM attachment_sets s"; migration.ExecuteNonQuery();
    }
    private SqliteConnection Open() { var db = new SqliteConnection(connectionString); db.Open(); return db; }
    public string? Classification(MailItem item, string fingerprint) { using var db=Open(); using var c=db.CreateCommand(); c.CommandText="SELECT label FROM classifications WHERE account=$a AND folder=$f AND validity=$v AND uid=$u AND fingerprint=$fingerprint"; Key(c,item.Account,item.Folder,item.Validity,item.Uid); c.Parameters.AddWithValue("$fingerprint",fingerprint); return c.ExecuteScalar() as string; }
    public void PutClassification(MailItem item, string fingerprint, string label) { using var db=Open(); using var c=db.CreateCommand(); c.CommandText="INSERT OR REPLACE INTO classifications(account,folder,validity,uid,fingerprint,label) VALUES($a,$f,$v,$u,$fingerprint,$label)"; Key(c,item.Account,item.Folder,item.Validity,item.Uid); c.Parameters.AddWithValue("$fingerprint",fingerprint); c.Parameters.AddWithValue("$label",label); c.ExecuteNonQuery(); }
    public void SaveFolders(string account, IReadOnlyList<MailFolderInfo> folders)
    {
        using var db = Open(); using var tx = db.BeginTransaction();
        using var del = db.CreateCommand(); del.Transaction = tx; del.CommandText = "DELETE FROM folders WHERE account=$a"; del.Parameters.AddWithValue("$a",account); del.ExecuteNonQuery();
        foreach (var folder in folders) { using var c = db.CreateCommand(); c.Transaction = tx; c.CommandText = "INSERT INTO folders(account,path,name,depth,selectable,validity) VALUES($a,$p,$n,$d,$s,COALESCE((SELECT MAX(validity) FROM messages WHERE account=$a AND folder=$p),0))"; c.Parameters.AddWithValue("$a",account); c.Parameters.AddWithValue("$p",folder.Path); c.Parameters.AddWithValue("$n",folder.Name); c.Parameters.AddWithValue("$d",folder.Depth); c.Parameters.AddWithValue("$s",folder.Selectable ? 1 : 0); c.ExecuteNonQuery(); }
        tx.Commit();
    }
    public List<MailFolderInfo> Folders(string account)
    {
        using var db=Open(); using var c=db.CreateCommand(); c.CommandText="SELECT path,name,depth,selectable FROM folders WHERE account=$a ORDER BY path COLLATE NOCASE"; c.Parameters.AddWithValue("$a",account); using var r=c.ExecuteReader(); var result=new List<MailFolderInfo>(); while(r.Read()) result.Add(new(r.GetString(0),r.GetString(1),r.GetInt32(2),r.GetInt32(3)!=0)); return result;
    }
    public List<MailItem> SearchAgentMessages(string account, string? folder, string query, string sender, string from, string to, int offset, int limit)
    {
        using var db = Open(); using var cmd = db.CreateCommand();
        cmd.CommandText = @"SELECT m.folder,m.validity,m.uid,m.sender,m.subject,m.date,m.preview,m.is_seen
FROM messages m JOIN folders f ON f.account=m.account AND f.path=m.folder AND f.validity=m.validity
WHERE m.account=$account AND ($folder='' OR m.folder=$folder)
AND ($query='' OR instr(lower(m.sender),lower($query))>0 OR instr(lower(m.subject),lower($query))>0 OR instr(lower(m.preview),lower($query))>0)
AND ($sender='' OR instr(lower(m.sender),lower($sender))>0)
AND ($from='' OR substr(m.date,1,10)>=$from) AND ($to='' OR substr(m.date,1,10)<=$to)
ORDER BY m.date DESC,m.folder,m.uid DESC LIMIT $limit OFFSET $offset";
        cmd.Parameters.AddWithValue("$account", account); cmd.Parameters.AddWithValue("$folder", folder ?? "");
        cmd.Parameters.AddWithValue("$query", query); cmd.Parameters.AddWithValue("$sender", sender);
        cmd.Parameters.AddWithValue("$from", from); cmd.Parameters.AddWithValue("$to", to);
        cmd.Parameters.AddWithValue("$offset", offset); cmd.Parameters.AddWithValue("$limit", limit);
        using var rows = cmd.ExecuteReader(); var result = new List<MailItem>();
        while (rows.Read()) result.Add(new MailItem((uint)rows.GetInt64(2), rows.GetString(3), rows.GetString(4), rows.GetString(5), rows.GetString(6), rows.IsDBNull(7) ? null : rows.GetInt32(7) != 0) { Account = account, Folder = rows.GetString(0), Validity = (uint)rows.GetInt64(1) });
        return result;
    }
    public void SetValidity(string account,string folder,uint validity)
    {
        using var db=Open(); using var tx=db.BeginTransaction(); using var c=db.CreateCommand(); c.Transaction=tx; c.CommandText="DELETE FROM messages WHERE account=$a AND folder=$f AND validity<>$v; DELETE FROM attachments WHERE account=$a AND folder=$f AND validity<>$v; DELETE FROM attachment_sets WHERE account=$a AND folder=$f AND validity<>$v; DELETE FROM attachment_flags WHERE account=$a AND folder=$f AND validity<>$v; UPDATE folders SET validity=$v WHERE account=$a AND path=$f"; c.Parameters.AddWithValue("$a",account); c.Parameters.AddWithValue("$f",folder); c.Parameters.AddWithValue("$v",validity); c.ExecuteNonQuery(); tx.Commit();
    }
    public uint Validity(string account,string folder) { using var db=Open(); using var c=db.CreateCommand(); c.CommandText="SELECT validity FROM folders WHERE account=$a AND path=$f"; c.Parameters.AddWithValue("$a",account); c.Parameters.AddWithValue("$f",folder); return Convert.ToUInt32(c.ExecuteScalar() ?? 0); }
    public List<uint> Uids(string account,string folder,uint validity) { using var db=Open(); using var c=db.CreateCommand(); c.CommandText="SELECT uid FROM messages WHERE account=$a AND folder=$f AND validity=$v ORDER BY uid DESC"; c.Parameters.AddWithValue("$a",account); c.Parameters.AddWithValue("$f",folder); c.Parameters.AddWithValue("$v",validity); using var r=c.ExecuteReader(); var list=new List<uint>(); while(r.Read()) list.Add((uint)r.GetInt64(0)); return list; }
    public MailItem? Get(string account,string folder,uint validity,uint uid) { using var db=Open(); using var c=db.CreateCommand(); c.CommandText="SELECT sender,subject,date,preview,is_seen,(SELECT has_attachments FROM attachment_flags f WHERE f.account=messages.account AND f.folder=messages.folder AND f.validity=messages.validity AND f.uid=messages.uid) FROM messages WHERE account=$a AND folder=$f AND validity=$v AND uid=$u"; Key(c,account,folder,validity,uid); using var r=c.ExecuteReader(); return r.Read()?new MailItem(uid,r.GetString(0),r.GetString(1),r.GetString(2),r.GetString(3),r.IsDBNull(4) ? null : r.GetInt32(4) != 0) { HasAttachments = r.IsDBNull(5) ? null : r.GetInt32(5) != 0 }:null; }
    public void PutAttachmentFlag(MailItem item) { if (!item.HasAttachments.HasValue) return; using var db=Open(); using var c=db.CreateCommand(); Key(c,item.Account,item.Folder,item.Validity,item.Uid); c.CommandText="INSERT OR REPLACE INTO attachment_flags VALUES($a,$f,$v,$u,$h)"; c.Parameters.AddWithValue("$h",item.HasAttachments.Value ? 1 : 0); c.ExecuteNonQuery(); }
    public void Put(string account,string folder,uint validity,MailItem item) { using var db=Open(); using var c=db.CreateCommand(); c.CommandText="INSERT INTO messages(account,folder,validity,uid,sender,subject,date,preview,is_seen) VALUES($a,$f,$v,$u,$s,$t,$d,$p,$seen) ON CONFLICT(account,folder,validity,uid) DO UPDATE SET sender=$s,subject=$t,date=$d,preview=$p,is_seen=$seen"; Key(c,account,folder,validity,item.Uid); c.Parameters.AddWithValue("$s",item.Sender); c.Parameters.AddWithValue("$t",item.Subject); c.Parameters.AddWithValue("$d",item.Date); c.Parameters.AddWithValue("$p",item.Preview); c.Parameters.AddWithValue("$seen",item.IsSeen.HasValue ? (object)(item.IsSeen.Value ? 1 : 0) : DBNull.Value); c.ExecuteNonQuery(); }
    public string? Body(string account,string folder,uint validity,uint uid) { using var db=Open(); using var c=db.CreateCommand(); c.CommandText="SELECT body FROM messages WHERE account=$a AND folder=$f AND validity=$v AND uid=$u"; Key(c,account,folder,validity,uid); var value=c.ExecuteScalar(); return value is string body ? body : null; }
    public void PutBody(string account,string folder,uint validity,uint uid,string body) { using var db=Open(); using var c=db.CreateCommand(); c.CommandText="UPDATE messages SET body=$b WHERE account=$a AND folder=$f AND validity=$v AND uid=$u"; Key(c,account,folder,validity,uid); c.Parameters.AddWithValue("$b",body); c.ExecuteNonQuery(); }
    public string? HtmlBody(string account,string folder,uint validity,uint uid) { using var db=Open(); using var c=db.CreateCommand(); c.CommandText="SELECT html_body FROM messages WHERE account=$a AND folder=$f AND validity=$v AND uid=$u"; Key(c,account,folder,validity,uid); return c.ExecuteScalar() as string; }
    public int HtmlVersion(string account,string folder,uint validity,uint uid) { using var db=Open(); using var c=db.CreateCommand(); c.CommandText="SELECT html_version FROM messages WHERE account=$a AND folder=$f AND validity=$v AND uid=$u"; Key(c,account,folder,validity,uid); return Convert.ToInt32(c.ExecuteScalar() ?? 0); }
    public void PutHtmlBody(string account,string folder,uint validity,uint uid,string html) { using var db=Open(); using var c=db.CreateCommand(); c.CommandText="UPDATE messages SET html_body=$b,html_version=2 WHERE account=$a AND folder=$f AND validity=$v AND uid=$u"; Key(c,account,folder,validity,uid); c.Parameters.AddWithValue("$b",html); c.ExecuteNonQuery(); }
    public void MarkSeen(string account,string folder,uint validity,uint uid) { using var db=Open(); using var c=db.CreateCommand(); c.CommandText="UPDATE messages SET is_seen=1 WHERE account=$a AND folder=$f AND validity=$v AND uid=$u"; Key(c,account,folder,validity,uid); c.ExecuteNonQuery(); }
    public List<MailAttachment>? Attachments(MailItem item)
    {
        using var db = Open(); using var c = db.CreateCommand();
        c.CommandText = "SELECT 1 FROM attachment_sets WHERE account=$a AND folder=$f AND validity=$v AND uid=$u"; Key(c,item.Account,item.Folder,item.Validity,item.Uid);
        if (c.ExecuteScalar() == null) return null;
        c.CommandText = "SELECT name,content_type,data FROM attachments WHERE account=$a AND folder=$f AND validity=$v AND uid=$u ORDER BY position";
        using var r = c.ExecuteReader(); var result = new List<MailAttachment>();
        while (r.Read()) result.Add(new(r.GetString(0), r.GetString(1), (byte[])r[2]));
        return result;
    }
    public void PutAttachments(MailItem item, IReadOnlyList<MailAttachment> attachments)
    {
        using var db = Open(); using var tx = db.BeginTransaction(); using var c = db.CreateCommand(); c.Transaction = tx;
        c.CommandText = "DELETE FROM attachments WHERE account=$a AND folder=$f AND validity=$v AND uid=$u"; Key(c,item.Account,item.Folder,item.Validity,item.Uid); c.ExecuteNonQuery();
        for (int i = 0; i < attachments.Count; i++)
        {
            using var row = db.CreateCommand(); row.Transaction = tx; Key(row,item.Account,item.Folder,item.Validity,item.Uid);
            row.CommandText = "INSERT INTO attachments VALUES($a,$f,$v,$u,$i,$n,$t,$d)";
            row.Parameters.AddWithValue("$i",i); row.Parameters.AddWithValue("$n",attachments[i].Name); row.Parameters.AddWithValue("$t",attachments[i].ContentType); row.Parameters.AddWithValue("$d",attachments[i].Data); row.ExecuteNonQuery();
        }
        c.CommandText = "INSERT OR REPLACE INTO attachment_sets VALUES($a,$f,$v,$u)"; c.ExecuteNonQuery(); tx.Commit();
    }
    private static void Key(SqliteCommand c,string a,string f,uint v,uint u) { c.Parameters.AddWithValue("$a",a); c.Parameters.AddWithValue("$f",f); c.Parameters.AddWithValue("$v",v); c.Parameters.AddWithValue("$u",u); }
}
