using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlMcp.Server.Services;

public class OpenXmlDocumentService
{
    public string CreateDocumentBase64(string title, string body)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        ArgumentException.ThrowIfNullOrWhiteSpace(body);

        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(
                new Body(
                    new Paragraph(new Run(new Text(title))),
                    new Paragraph(new Run(new Text(body)))));
        }

        return Convert.ToBase64String(stream.ToArray());
    }

    public string ExtractPlainText(string base64Docx)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(base64Docx);

        var bytes = Convert.FromBase64String(base64Docx);
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var body = document.MainDocumentPart?.Document?.Body;

        if (body is null)
        {
            return string.Empty;
        }

        var lines = body.Descendants<Paragraph>()
            .Select(paragraph => paragraph.InnerText)
            .Where(text => !string.IsNullOrWhiteSpace(text));

        return string.Join(Environment.NewLine, lines);
    }

    public int CountParagraphs(string base64Docx)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(base64Docx);

        var bytes = Convert.FromBase64String(base64Docx);
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart?.Document?.Body?.Descendants<Paragraph>().Count() ?? 0;
    }

    public string ExtractPlainTextFromPath(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("DOCX file not found.", filePath);
        }

        using var document = WordprocessingDocument.Open(filePath, false);
        var body = document.MainDocumentPart?.Document?.Body;

        if (body is null)
        {
            return string.Empty;
        }

        var lines = body.Descendants<Paragraph>()
            .Select(paragraph => paragraph.InnerText)
            .Where(text => !string.IsNullOrWhiteSpace(text));

        return string.Join(Environment.NewLine, lines);
    }

    public int CountParagraphsFromPath(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("DOCX file not found.", filePath);
        }

        using var document = WordprocessingDocument.Open(filePath, false);
        return document.MainDocumentPart?.Document?.Body?.Descendants<Paragraph>().Count() ?? 0;
    }
}
