namespace Docx2Md.Core.Models;

/// <summary>
/// Source metadata for a segment as defined in PRD Section 6.2
/// </summary>
public class SourceMetadata
{
    public string? StyleName { get; set; }
    public int? OutlineLevel { get; set; }
    public string? NumberingId { get; set; }
    public int? NumberingLevel { get; set; }
    public bool IsNumberedList { get; set; }
    public int? NumberingStartValue { get; set; }
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public bool IsUnderline { get; set; }
    public Dictionary<string, object> AdditionalProperties { get; set; } = new();
}
