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
    /// Enable hyperlink conversion to Markdown links
    /// </summary>
    public bool EnableHyperlinkConversion { get; set; } = true;

    /// <summary>
    /// Enable code block detection from monospace fonts
    /// </summary>
    public bool EnableCodeBlockDetection { get; set; } = true;

    /// <summary>
    /// Enable footnote/endnote conversion
    /// </summary>
    public bool EnableFootnoteConversion { get; set; } = true;

    /// <summary>
    /// Font families to detect as code (monospace fonts)
    /// </summary>
    public List<string> MonospaceFonts { get; set; } = new()
    {
        "Courier New",
        "Consolas",
        "Lucida Console",
        "Monaco",
        "Source Code Pro",
        "Fira Code",
        "Courier",
        "Menlo"
    };

    /// <summary>
    /// Default language hint for detected code blocks (empty for no language)
    /// </summary>
    public string CodeBlockLanguage { get; set; } = "";

    /// <summary>
    /// Custom style mappings (Word style name to Markdown conversion action)
    /// </summary>
    public List<StyleMapping> StyleMappings { get; set; } = new();

    /// <summary>
    /// Front matter template to use (null for no front matter)
    /// </summary>
    public FrontMatterTemplate? FrontMatterTemplate { get; set; }

    /// <summary>
    /// Include resolved numbering prefix in heading output
    /// (e.g., "# 11 INCIDENT MANAGEMENT" instead of "# INCIDENT MANAGEMENT")
    /// </summary>
    public bool IncludeHeadingNumbers { get; set; } = true;

    /// <summary>
    /// Convert list items with custom LevelText prefixes (like "Article %1") to paragraphs
    /// with the prefix prepended, rather than Markdown list items
    /// </summary>
    public bool ConvertPrefixedListsToParagraphs { get; set; } = true;

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
