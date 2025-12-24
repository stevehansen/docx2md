using Docx2Md.Core.Models;
using System.Text;

namespace Docx2Md.Core.Conversion;

/// <summary>
/// Converts segments to Markdown (PRD Section 6.3)
/// Rule-based and deterministic converter
/// </summary>
public class MarkdownConverter
{
    private readonly ConversionSettings _settings;

    public MarkdownConverter(ConversionSettings? settings = null)
    {
        _settings = settings ?? ConversionSettings.Default;
    }

    /// <summary>
    /// Convert all segments in a document to Markdown
    /// </summary>
    public void ConvertDocument(DocumentModel document)
    {
        for (int i = 0; i < document.Segments.Count; i++)
        {
            var segment = document.Segments[i];
            segment.MarkdownOutput = ConvertSegment(segment, document);
        }

        // Validate heading hierarchy
        ValidateHeadingHierarchy(document);
    }

    /// <summary>
    /// Convert a single segment to Markdown
    /// </summary>
    public string ConvertSegment(Segment segment, DocumentModel document)
    {
        if (string.IsNullOrWhiteSpace(segment.Content) && segment.EffectiveType != SegmentType.Table && segment.EffectiveType != SegmentType.Image)
            return string.Empty;

        // Check for custom style mapping first
        var styleMapping = GetStyleMapping(segment.Metadata.StyleName);
        if (styleMapping != null)
        {
            return ApplyStyleMapping(segment, styleMapping, document);
        }

        return segment.EffectiveType switch
        {
            SegmentType.Heading => ConvertHeading(segment),
            SegmentType.Paragraph => ConvertParagraph(segment),
            SegmentType.ListItem => ConvertListItem(segment),
            SegmentType.Table => ConvertTable(segment),
            SegmentType.Image => ConvertImage(segment, document),
            SegmentType.PageBreak => ConvertPageBreak(segment),
            SegmentType.SectionBreak => ConvertSectionBreak(segment),
            _ => ConvertParagraph(segment)
        };
    }

    /// <summary>
    /// Find a style mapping for the given style name
    /// </summary>
    private StyleMapping? GetStyleMapping(string? styleName)
    {
        if (string.IsNullOrEmpty(styleName))
            return null;

        return _settings.StyleMappings.FirstOrDefault(m =>
            m.WordStyleName.Equals(styleName, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Apply a style mapping to produce Markdown output
    /// </summary>
    private string ApplyStyleMapping(Segment segment, StyleMapping mapping, DocumentModel document)
    {
        var content = ProcessInlineFormatting(segment);

        switch (mapping.Action)
        {
            case StyleMappingAction.Heading:
                var level = Math.Clamp(mapping.HeadingLevel ?? 1, 1, _settings.MaxHeadingLevel);
                var prefix = new string('#', level);
                return $"{prefix} {content}";

            case StyleMappingAction.CodeBlock:
                var lang = mapping.CodeLanguage ?? _settings.CodeBlockLanguage;
                return $"```{lang}\n{segment.Content.Trim()}\n```";

            case StyleMappingAction.Blockquote:
                // Handle multi-line blockquotes
                var lines = content.Split('\n');
                return string.Join("\n", lines.Select(l => $"> {l}"));

            case StyleMappingAction.Exclude:
                segment.ExcludeFromOutput = true;
                return string.Empty;

            case StyleMappingAction.IgnoreStyle:
            case StyleMappingAction.Paragraph:
            default:
                var result = content;
                if (!string.IsNullOrEmpty(mapping.CustomPrefix))
                    result = mapping.CustomPrefix + result;
                if (!string.IsNullOrEmpty(mapping.CustomSuffix))
                    result = result + mapping.CustomSuffix;
                return result;
        }
    }

    private string ConvertHeading(Segment segment)
    {
        // Determine heading level
        int level = DetermineHeadingLevel(segment);

        if (level < 1) level = 1;
        if (level > _settings.MaxHeadingLevel) level = _settings.MaxHeadingLevel;

        var prefix = new string('#', level);
        var content = ProcessInlineFormatting(segment);

        if (_settings.InferHeadingsFromFormatting && level > 1)
        {
            segment.AddDiagnostic(
                DiagnosticLevel.Info,
                Diagnostics.DiagnosticCodes.HEADING_INFERRED,
                $"Heading level {level} inferred from style");
        }

        return $"{prefix} {content}";
    }

    private int DetermineHeadingLevel(Segment segment)
    {
        // Check for user override
        if (segment.OverrideHeadingLevel.HasValue)
            return segment.OverrideHeadingLevel.Value;

        // Check outline level
        if (segment.Metadata.OutlineLevel.HasValue)
            return segment.Metadata.OutlineLevel.Value + 1;

        // Parse from style name (e.g., "Heading1" -> 1)
        if (!string.IsNullOrEmpty(segment.Metadata.StyleName))
        {
            var styleName = segment.Metadata.StyleName;
            if (styleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
            {
                var levelStr = styleName.Substring(7);
                if (int.TryParse(levelStr, out int level))
                    return level;
            }

            if (styleName.Equals("Title", StringComparison.OrdinalIgnoreCase))
                return 1;
        }

        // Default to level 1
        return 1;
    }

    private string ConvertParagraph(Segment segment)
    {
        // Check if entire paragraph is monospace (should be code block)
        if (_settings.EnableCodeBlockDetection && IsEntirelyMonospace(segment))
        {
            return FormatAsCodeBlock(segment.Content);
        }

        var content = ProcessInlineFormatting(segment);

        if (_settings.PreserveLineBreaks)
        {
            // Preserve line breaks as double spaces + newline in Markdown
            content = content.Replace("\n", "  \n");
        }

        return content;
    }

    /// <summary>
    /// Check if segment is entirely monospace (code)
    /// </summary>
    private bool IsEntirelyMonospace(Segment segment)
    {
        if (segment.InlineElements.Count == 0)
            return false;

        // Check if all text runs are code
        var textRuns = segment.InlineElements.OfType<TextRun>().ToList();
        if (textRuns.Count == 0)
            return false;

        return textRuns.All(tr => tr.IsCode);
    }

    /// <summary>
    /// Format content as a fenced code block
    /// </summary>
    private string FormatAsCodeBlock(string content)
    {
        var lang = _settings.CodeBlockLanguage;
        return $"```{lang}\n{content.Trim()}\n```";
    }

    private string ConvertListItem(Segment segment)
    {
        var level = segment.Metadata.NumberingLevel ?? 0;
        var indent = new string(' ', level * 2);
        var content = ProcessInlineFormatting(segment);

        if (segment.Metadata.IsNumberedList)
        {
            // Numbered list - use actual item number from parsing
            var number = segment.Metadata.ListItemNumber ?? 1;
            return $"{indent}{number}. {content}";
        }
        else
        {
            // Bulleted list
            return $"{indent}- {content}";
        }
    }

    private string ConvertTable(Segment segment)
    {
        // Get table data from metadata
        if (!segment.Metadata.AdditionalProperties.TryGetValue("TableData", out var tableDataObj) ||
            tableDataObj is not List<List<string>> tableData ||
            tableData.Count == 0)
        {
            // Fallback for tables without extracted data
            segment.AddDiagnostic(
                DiagnosticLevel.Warning,
                Diagnostics.DiagnosticCodes.TABLE_COMPLEX_FORMATTING,
                "Table data could not be extracted.");
            return $"*[Table: {segment.Content}]*";
        }

        var sb = new StringBuilder();

        // First row is the header
        var headerRow = tableData[0];
        sb.AppendLine("| " + string.Join(" | ", headerRow.Select(EscapeTableCell)) + " |");

        // Separator row
        sb.AppendLine("| " + string.Join(" | ", headerRow.Select(_ => "---")) + " |");

        // Data rows
        for (int i = 1; i < tableData.Count; i++)
        {
            sb.AppendLine("| " + string.Join(" | ", tableData[i].Select(EscapeTableCell)) + " |");
        }

        return sb.ToString().TrimEnd();
    }

    private static string EscapeTableCell(string cell)
    {
        if (string.IsNullOrEmpty(cell))
            return " ";

        // Escape pipe characters and trim
        return cell.Replace("|", "\\|").Replace("\n", " ").Replace("\r", "");
    }

    private string ConvertImage(Segment segment, DocumentModel document)
    {
        // Get alt text from segment content or use default
        var altText = !string.IsNullOrEmpty(segment.Content) ? segment.Content : "Image";

        // Try to find image by relationship ID
        string imageName;
        if (segment.Metadata.AdditionalProperties.TryGetValue("ImageRelationshipId", out var rIdObj) &&
            rIdObj is string relationshipId)
        {
            var imageInfo = document.Images.FirstOrDefault(img => img.RelationshipId == relationshipId);
            if (imageInfo != null)
            {
                imageName = imageInfo.GetFileName();
            }
            else
            {
                // Image not found in document - use fallback name
                imageName = $"image_{segment.OrderIndex}.png";
                segment.AddDiagnostic(
                    DiagnosticLevel.Warning,
                    Diagnostics.DiagnosticCodes.IMAGE_EXTRACTION_FAILED,
                    $"Image with relationship ID '{relationshipId}' not found");
            }
        }
        else
        {
            // No relationship ID - use fallback name
            imageName = $"image_{segment.OrderIndex}.png";
        }

        var imagePath = $"{_settings.ImageOutputFolder}/{imageName}";

        if (string.IsNullOrEmpty(segment.Content))
        {
            segment.AddDiagnostic(
                DiagnosticLevel.Info,
                Diagnostics.DiagnosticCodes.IMAGE_ALT_TEXT_MISSING,
                "Image has no alt text");
        }

        return $"![{altText}]({imagePath})";
    }

    private string ConvertPageBreak(Segment segment)
    {
        return "\n---\n";
    }

    private string ConvertSectionBreak(Segment segment)
    {
        segment.AddDiagnostic(
            DiagnosticLevel.Info,
            Diagnostics.DiagnosticCodes.SECTION_BREAK_IGNORED,
            "Section break converted to horizontal rule");
        return "\n---\n";
    }

    /// <summary>
    /// Process inline formatting using InlineElements from the segment
    /// </summary>
    private string ProcessInlineFormatting(Segment segment)
    {
        // If no inline elements, just return trimmed content
        if (segment.InlineElements.Count == 0)
            return segment.Content.Trim();

        var sb = new StringBuilder();

        // Process elements in order
        foreach (var element in segment.InlineElements.OrderBy(e => e.StartOffset))
        {
            switch (element)
            {
                case HyperlinkElement link:
                    sb.Append(FormatHyperlink(link));
                    break;

                case TextRun run:
                    sb.Append(FormatTextRun(run));
                    break;

                case FootnoteReferenceElement fnRef:
                    sb.Append(FormatFootnoteReference(fnRef));
                    break;

                default:
                    sb.Append(element.Text);
                    break;
            }
        }

        return sb.ToString().Trim();
    }

    /// <summary>
    /// Format a hyperlink as Markdown
    /// </summary>
    private string FormatHyperlink(HyperlinkElement link)
    {
        if (string.IsNullOrEmpty(link.Url))
            return link.Text;

        var escapedText = EscapeMarkdownText(link.Text);
        return $"[{escapedText}]({link.Url})";
    }

    /// <summary>
    /// Format a text run with inline formatting (bold, italic, code, etc.)
    /// </summary>
    private string FormatTextRun(TextRun run)
    {
        var text = run.Text;

        if (string.IsNullOrEmpty(text))
            return text;

        // Handle inline code first (takes precedence)
        if (run.IsCode && _settings.EnableCodeBlockDetection)
        {
            return $"`{text}`";
        }

        // Apply formatting in order: bold+italic > bold > italic > strikethrough
        if (run.IsBold && run.IsItalic)
        {
            text = $"***{text}***";
        }
        else if (run.IsBold)
        {
            text = $"**{text}**";
        }
        else if (run.IsItalic)
        {
            text = $"*{text}*";
        }

        if (run.IsStrikethrough)
        {
            text = $"~~{text}~~";
        }

        // Convert underline to emphasis if setting is enabled
        if (run.IsUnderline && _settings.ConvertUnderlineToEmphasis && !run.IsItalic)
        {
            text = $"*{text}*";
        }

        return text;
    }

    /// <summary>
    /// Format a footnote reference
    /// </summary>
    private string FormatFootnoteReference(FootnoteReferenceElement fnRef)
    {
        return $"[^{fnRef.NoteId}]";
    }

    /// <summary>
    /// Escape special Markdown characters in text
    /// </summary>
    private static string EscapeMarkdownText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        // Escape brackets which are special in links
        return text.Replace("[", "\\[").Replace("]", "\\]");
    }

    private void ValidateHeadingHierarchy(DocumentModel document)
    {
        var headings = document.Segments
            .Where(s => s.EffectiveType == SegmentType.Heading)
            .ToList();

        int? previousLevel = null;
        foreach (var heading in headings)
        {
            var level = DetermineHeadingLevel(heading);
            
            if (previousLevel.HasValue && level > previousLevel.Value + 1)
            {
                heading.AddDiagnostic(
                    DiagnosticLevel.Warning,
                    Diagnostics.DiagnosticCodes.HEADING_HIERARCHY_INVALID,
                    $"Heading level jumps from {previousLevel.Value} to {level}");
            }

            previousLevel = level;
        }
    }
}
