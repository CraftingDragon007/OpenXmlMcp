using System.ComponentModel;
using OpenXmlMcp.Server.Services;
using ModelContextProtocol.Server;

namespace OpenXmlMcp.Server.Tools;

internal class OpenXmlTools(OpenXmlDocumentService documentService)
{
    [McpServerTool]
    [Description("Creates a DOCX document and returns it as a base64 string.")]
    public string CreateDocxBase64(
        [Description("Document title for the first paragraph.")] string title,
        [Description("Main content for the second paragraph.")] string body)
    {
        return documentService.CreateDocumentBase64(title, body);
    }

    [McpServerTool]
    [Description("Extracts plain text from a base64 encoded DOCX document.")]
    public string ExtractDocxText(
        [Description("Base64 encoded DOCX document.")] string base64Docx)
    {
        return documentService.ExtractPlainText(base64Docx);
    }

    [McpServerTool]
    [Description("Counts paragraphs in a base64 encoded DOCX document.")]
    public int CountDocxParagraphs(
        [Description("Base64 encoded DOCX document.")] string base64Docx)
    {
        return documentService.CountParagraphs(base64Docx);
    }
}
