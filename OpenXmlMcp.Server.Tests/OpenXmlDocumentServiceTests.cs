using OpenXmlMcp.Server.Services;

namespace OpenXmlMcp.Server.Tests;

public class OpenXmlDocumentServiceTests
{
    private readonly OpenXmlDocumentService _service = new();

    [Fact]
    public void CreateDocumentBase64_ThenExtractPlainText_ReturnsExpectedText()
    {
        var base64Docx = _service.CreateDocumentBase64("Invoice", "Line item A");

        var extractedText = _service.ExtractPlainText(base64Docx);

        Assert.Equal("Invoice" + Environment.NewLine + "Line item A", extractedText);
    }

    [Fact]
    public void CreateDocumentBase64_CountParagraphs_ReturnsTwo()
    {
        var base64Docx = _service.CreateDocumentBase64("Title", "Body");

        var paragraphCount = _service.CountParagraphs(base64Docx);

        Assert.Equal(2, paragraphCount);
    }

    [Fact]
    public void ExtractPlainText_WithInvalidBase64_ThrowsFormatException()
    {
        Assert.Throws<FormatException>(() => _service.ExtractPlainText("not-base64"));
    }

    [Fact]
    public void ExtractPlainTextFromPath_WithValidFile_ReturnsExpectedText()
    {
        var filePath = CreateTempDocxFile("Invoice", "Line item A");

        try
        {
            var extractedText = _service.ExtractPlainTextFromPath(filePath);

            Assert.Equal("Invoice" + Environment.NewLine + "Line item A", extractedText);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void CountParagraphsFromPath_WithValidFile_ReturnsTwo()
    {
        var filePath = CreateTempDocxFile("Title", "Body");

        try
        {
            var paragraphCount = _service.CountParagraphsFromPath(filePath);

            Assert.Equal(2, paragraphCount);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ExtractPlainTextFromPath_WithMissingFile_ThrowsFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() => _service.ExtractPlainTextFromPath("missing-file.docx"));
    }

    private string CreateTempDocxFile(string title, string body)
    {
        var base64Docx = _service.CreateDocumentBase64(title, body);
        var bytes = Convert.FromBase64String(base64Docx);
        var filePath = Path.Combine(Path.GetTempPath(), $"openxmlmcp-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(filePath, bytes);
        return filePath;
    }
}
