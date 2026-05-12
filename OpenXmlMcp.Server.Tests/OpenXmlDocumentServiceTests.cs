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
}
