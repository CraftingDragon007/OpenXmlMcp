namespace OpenXmlMcp.Server.Models;

public sealed record WordListItem(string Text, int Level, string Kind, string? BulletStyle = null, string? NumberStyle = null);
