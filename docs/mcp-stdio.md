# MailMind MCP stdio

Start MailMind and sign in to the account before connecting an MCP client. Configure the client to run the installed MailMind executable with `--mcp-stdio`, for example:

```json
{
  "mcpServers": {
    "mailmind": {
      "command": "C:\\path\\to\\MailMind.exe",
      "args": ["--mcp-stdio"]
    }
  }
}
```

Use the actual path to the installed executable. Each client starts a small stdio process; it connects to the running MailMind window through a local pipe restricted to the current Windows user. Standard output is reserved for MCP messages. The connection ends when MailMind exits. Restart the client after reopening MailMind.

The server exposes the same tools as the built-in assistant. Mail search uses the currently cached mailbox; current-mail and draft tools operate on the open window. Mail deletion still requires confirmation in that window. Tool access includes mail and contacts and can modify contacts and drafts, so only configure trusted local agents. All processes running as the same Windows user can access the local pipe.
