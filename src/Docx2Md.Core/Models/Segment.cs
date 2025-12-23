namespace Docx2Md.Core.Models;

/// <summary>
/// Represents a document segment - the atomic unit of conversion (PRD Section 5.1 and 6.2)
/// </summary>
public class Segment
{
    /// <summary>
    /// Unique identifier for this segment
    /// </summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>
    /// Document order index (0-based)
    /// </summary>
    public int OrderIndex { get; set; }

    /// <summary>
    /// Type of segment
    /// </summary>
    public SegmentType Type { get; set; }

    /// <summary>
    /// Source metadata from DOCX
    /// </summary>
    public SourceMetadata Metadata { get; set; } = new();

    /// <summary>
    /// Extracted content from DOCX
    /// </summary>
    public string Content { get; set; } = string.Empty;

    /// <summary>
    /// Converted Markdown output
    /// </summary>
    public string MarkdownOutput { get; set; } = string.Empty;

    /// <summary>
    /// Conversion diagnostics for this segment
    /// </summary>
    public List<Diagnostic> Diagnostics { get; set; } = new();

    /// <summary>
    /// Whether this segment should be excluded from output
    /// </summary>
    public bool ExcludeFromOutput { get; set; }

    /// <summary>
    /// User override for heading level (if applicable)
    /// </summary>
    public int? OverrideHeadingLevel { get; set; }

    /// <summary>
    /// User override for segment type
    /// </summary>
    public SegmentType? OverrideType { get; set; }

    /// <summary>
    /// Manual markdown override
    /// </summary>
    public string? ManualMarkdownOverride { get; set; }

    /// <summary>
    /// Get effective segment type (considering overrides)
    /// </summary>
    public SegmentType EffectiveType => OverrideType ?? Type;

    /// <summary>
    /// Get effective markdown output (considering overrides)
    /// </summary>
    public string EffectiveMarkdown
    {
        get
        {
            // Manual override takes highest priority
            if (!string.IsNullOrEmpty(ManualMarkdownOverride))
                return ManualMarkdownOverride;

            // If heading level is overridden and this is a heading, regenerate the markdown
            if (OverrideHeadingLevel.HasValue && EffectiveType == SegmentType.Heading)
            {
                var level = Math.Clamp(OverrideHeadingLevel.Value, 1, 6);
                var prefix = new string('#', level);
                return $"{prefix} {Content}";
            }

            return MarkdownOutput;
        }
    }

    /// <summary>
    /// Add a diagnostic to this segment
    /// </summary>
    public void AddDiagnostic(DiagnosticLevel level, string code, string message, string? details = null)
    {
        Diagnostics.Add(new Diagnostic(level, code, message, details));
    }

    /// <summary>
    /// Get a snippet of the content for display (first 100 chars)
    /// </summary>
    public string GetContentSnippet(int maxLength = 100)
    {
        if (string.IsNullOrEmpty(Content))
            return string.Empty;

        return Content.Length <= maxLength 
            ? Content 
            : Content.Substring(0, maxLength) + "...";
    }

    /// <summary>
    /// Get a snippet of the markdown output for display
    /// </summary>
    public string GetMarkdownSnippet(int maxLength = 100)
    {
        var markdown = EffectiveMarkdown;
        if (string.IsNullOrEmpty(markdown))
            return string.Empty;

        return markdown.Length <= maxLength 
            ? markdown 
            : markdown.Substring(0, maxLength) + "...";
    }
}
