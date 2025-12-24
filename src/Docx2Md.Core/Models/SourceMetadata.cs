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
    /// <summary>
    /// True if the paragraph explicitly has numId=0 (disabling any style-based numbering)
    /// </summary>
    public bool NumberingExplicitlyDisabled { get; set; }
    public bool IsNumberedList { get; set; }
    public int? NumberingStartValue { get; set; }
    /// <summary>
    /// The actual sequence number for this list item (computed during parsing)
    /// </summary>
    public int? ListItemNumber { get; set; }
    /// <summary>
    /// The LevelText format string from Word numbering definition (e.g., "%1.", "Article %1", "%1.%2.")
    /// </summary>
    public string? LevelTextFormat { get; set; }
    /// <summary>
    /// The resolved numbering prefix after substituting placeholders with actual numbers
    /// (e.g., "1.", "Article 225", "11.1.")
    /// </summary>
    public string? ResolvedNumberingPrefix { get; set; }
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public bool IsUnderline { get; set; }
    public Dictionary<string, object> AdditionalProperties { get; set; } = new();
}
