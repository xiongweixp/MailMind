using System;
using System.Net;
using System.Text;

namespace MailMind;

public static class HtmlTextExtractor
{
    public static string Extract(string html)
    {
        var output = new StringBuilder();
        int position = 0;
        while (position < html.Length)
        {
            int tagStart = html.IndexOf('<', position);
            if (tagStart < 0) { output.Append(html, position, html.Length - position); break; }
            output.Append(html, position, tagStart - position);
            if (html.AsSpan(tagStart).StartsWith("<!--", StringComparison.Ordinal))
            {
                int commentEnd = html.IndexOf("-->", tagStart + 4, StringComparison.Ordinal);
                if (commentEnd < 0) break;
                position = commentEnd + 3;
                continue;
            }
            int tagEnd = FindTagEnd(html, tagStart + 1);
            if (tagEnd < 0) { output.Append(html, tagStart, html.Length - tagStart); break; }
            var tag = html.AsSpan(tagStart + 1, tagEnd - tagStart - 1).TrimStart();
            int nameLength = 0;
            while (nameLength < tag.Length && char.IsLetterOrDigit(tag[nameLength])) nameLength++;
            var name = tag[..nameLength];
            if (name.Equals("script", StringComparison.OrdinalIgnoreCase) || name.Equals("style", StringComparison.OrdinalIgnoreCase) || name.Equals("head", StringComparison.OrdinalIgnoreCase))
            {
                string closing = "</" + name.ToString();
                int closeStart = html.IndexOf(closing, tagEnd + 1, StringComparison.OrdinalIgnoreCase);
                if (closeStart < 0) break;
                int closeEnd = FindTagEnd(html, closeStart + closing.Length);
                if (closeEnd < 0) break;
                position = closeEnd + 1;
            }
            else position = tagEnd + 1;
            output.Append(' ');
        }
        return WebUtility.HtmlDecode(output.ToString());
    }

    private static int FindTagEnd(string html, int position)
    {
        char quote = '\0';
        for (int i = position; i < html.Length; i++)
        {
            char ch = html[i];
            if (quote != '\0') { if (ch == quote) quote = '\0'; }
            else if (ch is '\'' or '"') quote = ch;
            else if (ch == '>') return i;
        }
        return -1;
    }
}
