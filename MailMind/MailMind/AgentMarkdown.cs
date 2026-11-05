using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Documents;
using System;
using System.Text.RegularExpressions;

namespace MailMind;

// Render common Markdown directly in the selectable text block, including partial streamed text.
internal static class AgentMarkdown
{
    private static readonly Regex Inline = new(@"(\*\*[^*\r\n]+\*\*|`[^`\r\n]+`|(?<!\*)\*[^*\r\n]+\*)", RegexOptions.Compiled);
    private static readonly Regex ListItem = new(@"^(\s*)([-*+]\s+|\d+[.)]\s+)", RegexOptions.Compiled);

    public static void Render(TextBlock target, string text)
    {
        target.Inlines.Clear();
        var lines = text.Replace("\r\n", "\n").Split('\n');
        bool code = false;
        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i];
            if (line.StartsWith("```", StringComparison.Ordinal)) { code = !code; continue; }
            if (i > 0 && target.Inlines.Count > 0) target.Inlines.Add(new LineBreak());
            bool heading = !code && Regex.IsMatch(line, @"^#{1,6}\s+");
            if (heading) line = Regex.Replace(line, @"^#{1,6}\s+", "");
            if (!code) line = ListItem.Replace(line, m => m.Groups[1].Value + (char.IsDigit(m.Groups[2].Value[0]) ? m.Groups[2].Value : "• "));
            int start = 0;
            foreach (Match match in Inline.Matches(line))
            {
                Add(target, line[start..match.Index], heading, code);
                string part = match.Value;
                bool strong = part.StartsWith("**", StringComparison.Ordinal);
                Add(target, part[(strong ? 2 : 1)..^(strong ? 2 : 1)], heading || strong, part[0] == '`', !strong && part[0] == '*');
                start = match.Index + match.Length;
            }
            Add(target, line[start..], heading, code);
        }
    }
    private static void Add(TextBlock target, string text, bool bold, bool code, bool italic = false)
    {
        if (text.Length == 0) return;
        var run = new Run { Text = text };
        if (bold) run.FontWeight = Microsoft.UI.Text.FontWeights.SemiBold;
        if (italic) run.FontStyle = Windows.UI.Text.FontStyle.Italic;
        if (code) run.FontFamily = new Microsoft.UI.Xaml.Media.FontFamily("Consolas");
        target.Inlines.Add(run);
    }
}
