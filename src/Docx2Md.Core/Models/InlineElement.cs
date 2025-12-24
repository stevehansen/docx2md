namespace Docx2Md.Core.Models;

/// <summary>
/// Base class for inline elements within paragraph text.
/// Captures hyperlinks, formatted runs, footnote references, etc.
/// </summary>
public abstract class InlineElement
{
    /// <summary>
    /// Start offset within the parent segment's Content string
    /// </summary>
    public int StartOffset { get; set; }

    /// <summary>
    /// End offset within the parent segment's Content string
    /// </summary>
    public int EndOffset { get; set; }

    /// <summary>
    /// The text content of this inline element
    /// </summary>
    public string Text { get; set; } = string.Empty;
}

/// <summary>
/// Represents a text run with formatting information
/// </summary>
public class TextRun : InlineElement
{
    /// <summary>
    /// Whether the text is bold
    /// </summary>
    public bool IsBold { get; set; }

    /// <summary>
    /// Whether the text is italic
    /// </summary>
    public bool IsItalic { get; set; }

    /// <summary>
    /// Whether the text is underlined
    /// </summary>
    public bool IsUnderline { get; set; }

    /// <summary>
    /// Whether the text has strikethrough
    /// </summary>
    public bool IsStrikethrough { get; set; }

    /// <summary>
    /// Whether the text should be rendered as code (monospace font detected)
    /// </summary>
    public bool IsCode { get; set; }

    /// <summary>
    /// The font family name (for code detection)
    /// </summary>
    public string? FontFamily { get; set; }
}

/// <summary>
/// Represents a hyperlink element
/// </summary>
public class HyperlinkElement : InlineElement
{
    /// <summary>
    /// The URL target of the hyperlink
    /// </summary>
    public string Url { get; set; } = string.Empty;

    /// <summary>
    /// Optional tooltip text
    /// </summary>
    public string? Tooltip { get; set; }

    /// <summary>
    /// Whether this is an internal bookmark link (starts with #)
    /// </summary>
    public bool IsBookmark => Url.StartsWith("#");
}

/// <summary>
/// Represents a footnote or endnote reference
/// </summary>
public class FootnoteReferenceElement : InlineElement
{
    /// <summary>
    /// The footnote/endnote ID
    /// </summary>
    public int NoteId { get; set; }

    /// <summary>
    /// Whether this is an endnote (false = footnote)
    /// </summary>
    public bool IsEndnote { get; set; }
}

/// <summary>
/// Represents a footnote or endnote definition
/// </summary>
public class FootnoteDefinition
{
    /// <summary>
    /// The footnote/endnote ID
    /// </summary>
    public int Id { get; set; }

    /// <summary>
    /// Whether this is an endnote (false = footnote)
    /// </summary>
    public bool IsEndnote { get; set; }

    /// <summary>
    /// The text content of the footnote/endnote
    /// </summary>
    public string Content { get; set; } = string.Empty;
}
