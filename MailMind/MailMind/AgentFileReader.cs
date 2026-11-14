using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace MailMind;

internal static class AgentFileReader
{
    internal static string Read(string name, byte[] data)
    {
        string extension = Path.GetExtension(name).ToLowerInvariant();
        if (extension == ".pdf")
        {
            using var document = UglyToad.PdfPig.PdfDocument.Open(data);
            var result = new StringBuilder();
            bool hasText = false;
            foreach (var page in document.GetPages())
            {
                string pageText = page.Text;
                hasText |= !string.IsNullOrWhiteSpace(pageText);
                result.AppendLine($"--- Page {page.Number} ---");
                result.AppendLine(pageText);
                if (result.Length > 2_000_000) throw new InvalidDataException("PDF 文本过大，无法读取。");
            }
            if (!hasText)
                throw new InvalidDataException("PDF 没有可提取的文本；扫描件需要 OCR，目前尚不支持。");
            return result.ToString();
        }
        if (extension == ".docx")
        {
            using var archive = new ZipArchive(new MemoryStream(data), ZipArchiveMode.Read);
            var entry = archive.GetEntry("word/document.xml") ?? throw new InvalidDataException("无效的 DOCX 文件。");
            if (entry.Length > 20 * 1024 * 1024) throw new InvalidDataException("文档内容过大。");
            using var stream = entry.Open();
            var document = XDocument.Load(stream);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            return string.Join("\n", document.Descendants(w + "p").Select(p => string.Concat(p.Descendants(w + "t").Select(t => t.Value))));
        }
        if (extension is not (".txt" or ".md" or ".csv" or ".tsv" or ".json" or ".xml" or ".html" or ".htm" or ".log"))
            throw new InvalidOperationException("不支持读取此文件类型。支持 PDF、文本、CSV、JSON、HTML 和 DOCX。");
        using var reader = new StreamReader(new MemoryStream(data), Encoding.UTF8, true);
        string text = reader.ReadToEnd();
        return extension is ".html" or ".htm" ? HtmlTextExtractor.Extract(text) : text;
    }
}
