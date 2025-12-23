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
        if (string.IsNullOrWhiteSpace(segment.Content))
            return string.Empty;

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

    private string ConvertHeading(Segment segment)
    {
        // Determine heading level
        int level = DetermineHeadingLevel(segment);
        
        if (level < 1) level = 1;
        if (level > _settings.MaxHeadingLevel) level = _settings.MaxHeadingLevel;

        var prefix = new string('#', level);
        var content = ProcessInlineFormatting(segment.Content);

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
        var content = ProcessInlineFormatting(segment.Content);
        
        if (_settings.PreserveLineBreaks)
        {
            // Preserve line breaks as double spaces + newline in Markdown
            content = content.Replace("\n", "  \n");
        }

        return content;
    }

    private string ConvertListItem(Segment segment)
    {
        var level = segment.Metadata.NumberingLevel ?? 0;
        var indent = new string(' ', level * 2);
        var content = ProcessInlineFormatting(segment.Content);

        if (segment.Metadata.IsNumberedList)
        {
            // Numbered list - use "1." format (Markdown auto-numbers)
            return $"{indent}1. {content}";
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
        // Find image by relationship ID or create placeholder
        var altText = segment.Content.Length > 0 ? segment.Content : "Image";
        
        // Generate relative image path
        var imageName = $"image_{segment.OrderIndex}.png";
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

    private string ProcessInlineFormatting(string content)
    {
        // Basic inline formatting processing
        // This is simplified - full implementation would track formatting runs
        return content.Trim();
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
