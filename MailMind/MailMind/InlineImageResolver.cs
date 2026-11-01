using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using MimeKit;

namespace MailMind;

internal static class InlineImageResolver
{
    private static readonly Regex References = new(
        "(?<prefix>\\b(?:src|background|poster)\\s*=\\s*|url\\(\\s*)(?<quote>[\"']?)(?<url>[^\\s\"'<>\\)]+)(?<end>[\"']?)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public static string Resolve(MimeMessage message)
    {
        var ids = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var locations = new Dictionary<string, string>(StringComparer.Ordinal);
        var htmlPart = message.BodyParts.OfType<TextPart>().FirstOrDefault(p => p.IsHtml);
        var htmlBase = AbsoluteUri(htmlPart?.ContentBase) ?? AbsoluteUri(htmlPart?.ContentLocation);
        foreach (var part in message.BodyParts.OfType<MimePart>())
        {
            var type = ImageType(part);
            if (type == null || part.Content == null || (part.ContentId == null && part.ContentLocation == null)) continue;
            using var bytes = new MemoryStream();
            part.Content.DecodeTo(bytes);
            var data = $"data:{type};base64,{Convert.ToBase64String(bytes.ToArray())}";
            if (part.ContentId != null) ids[part.ContentId.Trim('<', '>')] = data;
            if (part.ContentLocation != null)
            {
                var location = part.ContentLocation.ToString();
                locations[location] = data;
                var absolute = ResolveUri(AbsoluteUri(part.ContentBase) ?? htmlBase, location);
                if (absolute != null) locations[absolute] = data;
            }
        }
        return References.Replace(message.HtmlBody ?? "", match =>
        {
            var url = WebUtility.HtmlDecode(match.Groups["url"].Value);
            string? data = null;
            if (url.StartsWith("cid:", StringComparison.OrdinalIgnoreCase))
            {
                var id = Uri.UnescapeDataString(url[4..]).Trim('<', '>');
                ids.TryGetValue(id, out data);
            }
            else if (!locations.TryGetValue(url, out data))
            {
                var absolute = ResolveUri(htmlBase, url);
                if (absolute != null) locations.TryGetValue(absolute, out data);
            }
            return data == null ? match.Value : match.Groups["prefix"].Value + match.Groups["quote"].Value + data + match.Groups["end"].Value;
        });
    }

    private static Uri? AbsoluteUri(Uri? value) => value?.IsAbsoluteUri == true ? value : null;
    private static string? ResolveUri(Uri? baseUri, string value) =>
        Uri.TryCreate(value, UriKind.Absolute, out var absolute) ? absolute.ToString() :
        baseUri != null && Uri.TryCreate(baseUri, value, out var resolved) ? resolved.ToString() : null;

    private static string? ImageType(MimePart part)
    {
        var type = part.ContentType.MimeType.ToLowerInvariant();
        if (type is "image/png" or "image/jpeg" or "image/gif" or "image/webp" or "image/bmp" or "image/avif") return type;
        if (type != "application/octet-stream") return null;
        var extension = Path.GetExtension(part.FileName ?? part.ContentLocation?.ToString() ?? "").ToLowerInvariant();
        return extension switch { ".png" => "image/png", ".jpg" or ".jpeg" => "image/jpeg", ".gif" => "image/gif", ".webp" => "image/webp", ".bmp" => "image/bmp", ".avif" => "image/avif", _ => null };
    }
}
