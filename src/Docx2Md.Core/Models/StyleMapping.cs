namespace Docx2Md.Core.Models;

/// <summary>
/// Defines a custom mapping from a Word style to Markdown output behavior
/// </summary>
public class StyleMapping
{
    /// <summary>
    /// The Word style name to match (case-insensitive)
    /// </summary>
    public string WordStyleName { get; set; } = string.Empty;

    /// <summary>
    /// The action to take when this style is encountered
    /// </summary>
    public StyleMappingAction Action { get; set; } = StyleMappingAction.Paragraph;

    /// <summary>
    /// Heading level (1-6) when Action is Heading
    /// </summary>
    public int? HeadingLevel { get; set; }

    /// <summary>
    /// Custom prefix to add before the content
    /// </summary>
    public string? CustomPrefix { get; set; }

    /// <summary>
    /// Custom suffix to add after the content
    /// </summary>
    public string? CustomSuffix { get; set; }

    /// <summary>
    /// Language hint for code blocks (when Action is CodeBlock)
    /// </summary>
    public string? CodeLanguage { get; set; }
}

/// <summary>
/// Actions that can be applied when a style mapping matches
/// </summary>
public enum StyleMappingAction
{
    /// <summary>
    /// Treat as a normal paragraph
    /// </summary>
    Paragraph,

    /// <summary>
    /// Convert to a heading (use HeadingLevel to specify level)
    /// </summary>
    Heading,

    /// <summary>
    /// Wrap in a fenced code block
    /// </summary>
    CodeBlock,

    /// <summary>
    /// Wrap in a blockquote
    /// </summary>
    Blockquote,

    /// <summary>
    /// Remove style formatting, treat as plain text
    /// </summary>
    IgnoreStyle,

    /// <summary>
    /// Exclude from output entirely
    /// </summary>
    Exclude
}
