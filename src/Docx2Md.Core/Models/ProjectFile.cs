namespace Docx2Md.Core.Models;

/// <summary>
/// Represents a project file that stores document overrides and settings
/// </summary>
public class ProjectFile
{
    /// <summary>
    /// Project file format version
    /// </summary>
    public string Version { get; set; } = "1.0";

    /// <summary>
    /// Path to the source DOCX file
    /// </summary>
    public string SourceDocxPath { get; set; } = string.Empty;

    /// <summary>
    /// Last export path (if any)
    /// </summary>
    public string? LastExportPath { get; set; }

    /// <summary>
    /// When the project was last modified
    /// </summary>
    public DateTime LastModified { get; set; } = DateTime.Now;

    /// <summary>
    /// Conversion settings used for this project
    /// </summary>
    public ConversionSettings Settings { get; set; } = new();

    /// <summary>
    /// Segment overrides applied to this document
    /// </summary>
    public List<SegmentOverride> SegmentOverrides { get; set; } = new();

    /// <summary>
    /// Front matter template for this project (overrides settings)
    /// </summary>
    public FrontMatterTemplate? FrontMatter { get; set; }
}

/// <summary>
/// Stores overrides for a specific segment
/// </summary>
public class SegmentOverride
{
    /// <summary>
    /// Segment ID to match
    /// </summary>
    public string SegmentId { get; set; } = string.Empty;

    /// <summary>
    /// Order index as a fallback identifier
    /// </summary>
    public int OrderIndex { get; set; }

    /// <summary>
    /// Hash of segment content for detecting changes
    /// </summary>
    public string? ContentHash { get; set; }

    /// <summary>
    /// Whether segment is excluded from output
    /// </summary>
    public bool? ExcludeFromOutput { get; set; }

    /// <summary>
    /// Override heading level
    /// </summary>
    public int? OverrideHeadingLevel { get; set; }

    /// <summary>
    /// Override segment type
    /// </summary>
    public SegmentType? OverrideType { get; set; }

    /// <summary>
    /// Manual markdown override
    /// </summary>
    public string? ManualMarkdownOverride { get; set; }
}
