using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlMcp.Server.Services;

namespace OpenXmlMcp.Server.Tests;

public static class OfficeSessionServiceTestHelpers
{
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
