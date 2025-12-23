namespace Docx2Md.Core.Diagnostics;

/// <summary>
/// Centralized catalog of diagnostic codes (PRD Section 7.2)
/// </summary>
public static class DiagnosticCodes
{
    // Unsupported features
    public const string HEADER_FOOTER_IGNORED = "HEADER_FOOTER_IGNORED";
    public const string TEXT_BOX_IGNORED = "TEXT_BOX_IGNORED";
    public const string SHAPE_IGNORED = "SHAPE_IGNORED";
    public const string SMARTART_IGNORED = "SMARTART_IGNORED";
    public const string TRACK_CHANGES_IGNORED = "TRACK_CHANGES_IGNORED";
    public const string COMMENT_IGNORED = "COMMENT_IGNORED";
    public const string SECTION_BREAK_IGNORED = "SECTION_BREAK_IGNORED";
    public const string MULTI_COLUMN_LAYOUT_IGNORED = "MULTI_COLUMN_LAYOUT_IGNORED";
    public const string FLOATING_IMAGE_IGNORED = "FLOATING_IMAGE_IGNORED";
    public const string FIELD_IGNORED = "FIELD_IGNORED";

    // Table issues
    public const string TABLE_MERGED_CELLS_LOSSY = "TABLE_MERGED_CELLS_LOSSY";
    public const string TABLE_NESTED_TABLE = "TABLE_NESTED_TABLE";
    public const string TABLE_COMPLEX_FORMATTING = "TABLE_COMPLEX_FORMATTING";

    // Heading detection
    public const string HEADING_INFERRED = "HEADING_INFERRED";
    public const string HEADING_STYLE_AMBIGUOUS = "HEADING_STYLE_AMBIGUOUS";

    // List issues
    public const string LIST_CONTINUITY_WARNING = "LIST_CONTINUITY_WARNING";
    public const string LIST_NESTING_COMPLEX = "LIST_NESTING_COMPLEX";

    // Formatting
    public const string INLINE_FORMATTING_LOSS = "INLINE_FORMATTING_LOSS";
    public const string FONT_NOT_PRESERVED = "FONT_NOT_PRESERVED";
    public const string COLOR_NOT_PRESERVED = "COLOR_NOT_PRESERVED";

    // Images
    public const string IMAGE_MISSING_REFERENCE = "IMAGE_MISSING_REFERENCE";
    public const string IMAGE_EXTRACTION_FAILED = "IMAGE_EXTRACTION_FAILED";
    public const string IMAGE_ALT_TEXT_MISSING = "IMAGE_ALT_TEXT_MISSING";

    // Quality checks
    public const string HEADING_HIERARCHY_INVALID = "HEADING_HIERARCHY_INVALID";
    public const string EXCESSIVE_STYLE_AMBIGUITY = "EXCESSIVE_STYLE_AMBIGUITY";
}
