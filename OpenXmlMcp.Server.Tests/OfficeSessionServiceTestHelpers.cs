using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlMcp.Server.Services;

namespace OpenXmlMcp.Server.Tests;

public static class OfficeSessionServiceTestHelpers
{
    /// <summary>
    /// Creates a fully wired <see cref="OfficeSessionService"/> with all its dependencies.
    /// </summary>
    public static OfficeSessionService CreateService()
    {
        var sessionManager = new SessionManager();
        var wordService = new WordDocumentService(sessionManager);
        var excelService = new ExcelDocumentService(sessionManager);
        var powerPointService = new PowerPointDocumentService(sessionManager);
        return new OfficeSessionService(sessionManager, wordService, excelService, powerPointService);
    }

    public static string GetTempPath(string extension)
    {
        return Path.Combine(Path.GetTempPath(), $"openxmlmcp-{Guid.NewGuid():N}.{extension}");
    }

    public static void DeleteIfExists(string filePath)
    {
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }
}
