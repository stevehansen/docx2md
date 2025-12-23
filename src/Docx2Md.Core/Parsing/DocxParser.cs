using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docx2Md.Core.Models;

namespace Docx2Md.Core.Parsing;

/// <summary>
/// Parses DOCX files and extracts segments (PRD Section 6.1)
/// </summary>
public class DocxParser
{
    private readonly ConversionSettings _settings;

    public DocxParser(ConversionSettings? settings = null)
    {
        _settings = settings ?? ConversionSettings.Default;
    }

    /// <summary>
    /// Parse a DOCX file and extract segments
    /// </summary>
    public DocumentModel Parse(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"DOCX file not found: {filePath}");

        var document = new DocumentModel
        {
            SourcePath = filePath
        };

        using (var wordDoc = WordprocessingDocument.Open(filePath, false))
        {
            if (wordDoc.MainDocumentPart == null)
                throw new InvalidOperationException("Document has no main document part");

            var body = wordDoc.MainDocumentPart.Document.Body;
            if (body == null)
            {
                document.AddDiagnostic(DiagnosticLevel.Warning, "EMPTY_DOCUMENT", "Document body is empty");
                return document;
            }

            // Extract document title from core properties
            ExtractDocumentMetadata(wordDoc, document);

            // Process body elements
            int orderIndex = 0;
            foreach (var element in body.Elements())
            {
                var segment = ProcessElement(element, orderIndex, wordDoc);
                if (segment != null)
                {
                    document.Segments.Add(segment);
                    orderIndex++;
                }
            }

            // Extract images
            ExtractImages(wordDoc, document);

            // Detect unsupported features
            DetectUnsupportedFeatures(wordDoc, document);
        }

        return document;
    }

    private void ExtractDocumentMetadata(WordprocessingDocument wordDoc, DocumentModel document)
    {
        try
        {
            var coreProps = wordDoc.PackageProperties;
            document.Title = coreProps.Title;
            document.Metadata["Creator"] = coreProps.Creator ?? string.Empty;
            document.Metadata["Created"] = coreProps.Created ?? DateTime.MinValue;
            document.Metadata["Modified"] = coreProps.Modified ?? DateTime.MinValue;
        }
        catch
        {
            // Metadata extraction is best-effort
        }
    }

    private Segment? ProcessElement(DocumentFormat.OpenXml.OpenXmlElement element, int orderIndex, WordprocessingDocument wordDoc)
    {
        return element switch
        {
            Paragraph paragraph => ProcessParagraph(paragraph, orderIndex, wordDoc),
            Table table => ProcessTable(table, orderIndex),
            _ => null
        };
    }

    private Segment ProcessParagraph(Paragraph paragraph, int orderIndex, WordprocessingDocument wordDoc)
    {
        var segment = new Segment
        {
            OrderIndex = orderIndex
        };

        // Extract paragraph style
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        segment.Metadata.StyleName = styleId;

        // Extract text content
        segment.Content = GetParagraphText(paragraph);

        // Detect segment type based on style
        segment.Type = DetectSegmentType(paragraph, styleId);

        // Extract formatting
        ExtractParagraphFormatting(paragraph, segment);

        // Check for images in paragraph
        ProcessInlineImages(paragraph, segment, wordDoc);

        return segment;
    }

    private string GetParagraphText(Paragraph paragraph)
    {
        var texts = paragraph.Descendants<Text>();
        return string.Join("", texts.Select(t => t.Text));
    }

    private SegmentType DetectSegmentType(Paragraph paragraph, string? styleId)
    {
        // Check if it's a heading based on style
        if (!string.IsNullOrEmpty(styleId))
        {
            if (styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
                return SegmentType.Heading;
            
            if (styleId.StartsWith("Title", StringComparison.OrdinalIgnoreCase))
                return SegmentType.Heading;
        }

        // Check outline level
        var outlineLevel = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (outlineLevel.HasValue)
        {
            return SegmentType.Heading;
        }

        // Check for list item
        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        if (numPr != null)
        {
            return SegmentType.ListItem;
        }

        return SegmentType.Paragraph;
    }

    private void ExtractParagraphFormatting(Paragraph paragraph, Segment segment)
    {
        var runProps = paragraph.Descendants<RunProperties>().FirstOrDefault();
        if (runProps != null)
        {
            segment.Metadata.IsBold = runProps.Bold != null;
            segment.Metadata.IsItalic = runProps.Italic != null;
            segment.Metadata.IsUnderline = runProps.Underline != null;
        }

        var outlineLevel = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (outlineLevel.HasValue)
        {
            segment.Metadata.OutlineLevel = outlineLevel.Value;
        }

        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        if (numPr != null)
        {
            segment.Metadata.NumberingId = numPr.NumberingId?.Val?.Value.ToString();
            segment.Metadata.NumberingLevel = numPr.NumberingLevelReference?.Val?.Value;
        }
    }

    private void ProcessInlineImages(Paragraph paragraph, Segment segment, WordprocessingDocument wordDoc)
    {
        var drawings = paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Blip>();
        if (drawings.Any())
        {
            segment.Type = SegmentType.Image;
            // Image relationship IDs would be stored in metadata
        }
    }

    private Segment ProcessTable(Table table, int orderIndex)
    {
        var segment = new Segment
        {
            OrderIndex = orderIndex,
            Type = SegmentType.Table
        };

        // Extract table content
        var rows = table.Elements<TableRow>();
        var rowCount = rows.Count();
        
        segment.Content = $"Table with {rowCount} rows";
        segment.Metadata.AdditionalProperties["RowCount"] = rowCount;

        // Check for merged cells
        var mergedCells = table.Descendants<TableCell>()
            .Where(c => c.TableCellProperties?.GridSpan != null || 
                       c.TableCellProperties?.VerticalMerge != null);
        
        if (mergedCells.Any())
        {
            segment.AddDiagnostic(
                DiagnosticLevel.Warning,
                Diagnostics.DiagnosticCodes.TABLE_MERGED_CELLS_LOSSY,
                "Table contains merged cells which may not convert perfectly to Markdown");
        }

        return segment;
    }

    private void ExtractImages(WordprocessingDocument wordDoc, DocumentModel document)
    {
        if (wordDoc.MainDocumentPart?.ImageParts == null)
            return;

        foreach (var imagePart in wordDoc.MainDocumentPart.ImageParts)
        {
            try
            {
                using (var stream = imagePart.GetStream())
                using (var ms = new MemoryStream())
                {
                    stream.CopyTo(ms);
                    var imageInfo = new ImageInfo
                    {
                        Data = ms.ToArray(),
                        ContentType = imagePart.ContentType,
                        RelationshipId = wordDoc.MainDocumentPart.GetIdOfPart(imagePart)
                    };
                    document.Images.Add(imageInfo);
                }
            }
            catch (Exception ex)
            {
                document.AddDiagnostic(
                    DiagnosticLevel.Error,
                    Diagnostics.DiagnosticCodes.IMAGE_EXTRACTION_FAILED,
                    "Failed to extract image",
                    ex.Message);
            }
        }
    }

    private void DetectUnsupportedFeatures(WordprocessingDocument wordDoc, DocumentModel document)
    {
        // Check for headers/footers
        if (wordDoc.MainDocumentPart?.HeaderParts?.Any() == true ||
            wordDoc.MainDocumentPart?.FooterParts?.Any() == true)
        {
            document.AddDiagnostic(
                DiagnosticLevel.Info,
                Diagnostics.DiagnosticCodes.HEADER_FOOTER_IGNORED,
                "Document contains headers or footers which are not included in the conversion");
        }

        // Check for comments
        if (wordDoc.MainDocumentPart?.WordprocessingCommentsPart != null)
        {
            document.AddDiagnostic(
                DiagnosticLevel.Info,
                Diagnostics.DiagnosticCodes.COMMENT_IGNORED,
                "Document contains comments which are not included in the conversion");
        }
    }
}
