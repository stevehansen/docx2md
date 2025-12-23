namespace Docx2Md.Core.Models;

/// <summary>
/// User-configurable conversion settings (PRD Section 8)
/// </summary>
public class ConversionSettings
{
    /// <summary>
    /// Enable style-based heading detection
    /// </summary>
    public bool EnableStyleBasedHeadingDetection { get; set; } = true;

    /// <summary>
    /// Infer headings from formatting (bold, larger font)
    /// </summary>
    public bool InferHeadingsFromFormatting { get; set; } = true;

    /// <summary>
    /// Convert underline to emphasis
    /// </summary>
    public bool ConvertUnderlineToEmphasis { get; set; } = false;

    /// <summary>
    /// Emit HTML for unsupported constructs
    /// </summary>
    public bool EmitHtmlForUnsupported { get; set; } = false;

    /// <summary>
    /// Append "Lost & Found" appendix section for unprocessed content
    /// </summary>
    public bool AppendLostAndFoundSection { get; set; } = true;

    /// <summary>
    /// Maximum heading level to output (1-6)
    /// </summary>
    public int MaxHeadingLevel { get; set; } = 6;

    /// <summary>
    /// Image output folder relative to markdown file
    /// </summary>
    public string ImageOutputFolder { get; set; } = "images";

    /// <summary>
    /// Generate diagnostic report on export
    /// </summary>
    public bool GenerateDiagnosticReport { get; set; } = true;

    /// <summary>
    /// Diagnostic report format
    /// </summary>
    public DiagnosticReportFormat DiagnosticReportFormat { get; set; } = DiagnosticReportFormat.Markdown;

    /// <summary>
    /// Preserve line breaks in paragraphs
    /// </summary>
    public bool PreserveLineBreaks { get; set; } = true;

    /// <summary>
    /// Create default settings
    /// </summary>
    public static ConversionSettings Default => new ConversionSettings();
}

public enum DiagnosticReportFormat
{
    Markdown,
    Json
}
