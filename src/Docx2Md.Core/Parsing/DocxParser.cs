using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docx2Md.Core.Models;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DrawingWp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Vml = DocumentFormat.OpenXml.Vml;

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

            // Activate style-based numbering for headings after the first numbered one
            // This handles Word's "from here on" behavior for outline numbering
            ActivateStyleNumbering(document, wordDoc);

            // Calculate list item numbers and resolve numbering prefixes after all segments are parsed
            CalculateListItemNumbers(document, wordDoc);
        }

        return document;
    }

    /// <summary>
    /// Activate style-based numbering for headings after the first numbered heading.
    /// Word has a "from here on" behavior where front-matter headings don't have numbers,
    /// but once outline numbering starts, subsequent headings continue the sequence.
    /// </summary>
    private void ActivateStyleNumbering(DocumentModel document, WordprocessingDocument wordDoc)
    {
        // First, find the index where numbering "starts" (first heading with explicit numbering)
        int numberingStartIndex = -1;
        for (int i = 0; i < document.Segments.Count; i++)
        {
            var seg = document.Segments[i];
            if (seg.Type == SegmentType.Heading && !string.IsNullOrEmpty(seg.Metadata.NumberingId))
            {
                numberingStartIndex = i;
                break;
            }
        }

        // If no heading has explicit numbering, nothing to activate
        if (numberingStartIndex < 0)
            return;

        // Cache style numbering info for styles that have numbering defined
        var styleNumberingInfo = new Dictionary<string, (int numId, int level)>();

        // Second pass: for headings after the start index, apply style numbering if not already set
        for (int i = numberingStartIndex; i < document.Segments.Count; i++)
        {
            var segment = document.Segments[i];
            if (segment.Type != SegmentType.Heading)
                continue;

            var styleName = segment.Metadata.StyleName;
            if (string.IsNullOrEmpty(styleName))
                continue;

            // Skip headings that have explicitly disabled numbering (numId=0)
            if (segment.Metadata.NumberingExplicitlyDisabled)
                continue;

            // If this heading already has numbering, cache it for this style
            if (!string.IsNullOrEmpty(segment.Metadata.NumberingId) &&
                int.TryParse(segment.Metadata.NumberingId, out var numId))
            {
                var level = segment.Metadata.NumberingLevel ?? 0;
                styleNumberingInfo[styleName] = (numId, level);
            }
            // If this heading doesn't have numbering, try to get it from style definition or cache
            else
            {
                // First check our cache (from previously seen headings of this style)
                if (styleNumberingInfo.TryGetValue(styleName, out var cachedInfo))
                {
                    segment.Metadata.NumberingId = cachedInfo.numId.ToString();
                    segment.Metadata.NumberingLevel = cachedInfo.level;
                    DetectNumberingFormatFromValues(wordDoc, segment, cachedInfo.numId, cachedInfo.level);
                }
                // Otherwise, check the style definition
                else
                {
                    var styleNumPr = GetStyleNumberingProperties(styleName, wordDoc);
                    if (styleNumPr != null)
                    {
                        segment.Metadata.NumberingId = styleNumPr.Value.numId.ToString();
                        segment.Metadata.NumberingLevel = styleNumPr.Value.level;
                        DetectNumberingFormatFromValues(wordDoc, segment, styleNumPr.Value.numId, styleNumPr.Value.level);

                        // Cache for future headings of this style
                        styleNumberingInfo[styleName] = styleNumPr.Value;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Calculate sequential list item numbers and resolve numbering prefixes for lists and headings
    /// </summary>
    private void CalculateListItemNumbers(DocumentModel document, WordprocessingDocument wordDoc)
    {
        var resolver = new NumberingResolver(wordDoc);

        foreach (var segment in document.Segments)
        {
            // Process both list items and headings that have numbering
            if (string.IsNullOrEmpty(segment.Metadata.NumberingId))
                continue;

            if (!int.TryParse(segment.Metadata.NumberingId, out var numId))
                continue;

            var level = segment.Metadata.NumberingLevel ?? 0;

            // Get level info and resolve the numbering prefix
            var levelInfo = resolver.GetLevelInfo(numId, level);
            if (levelInfo != null)
            {
                segment.Metadata.ResolvedNumberingPrefix = resolver.ResolveNumberingPrefix(numId, level, levelInfo, segment.Type);
                segment.Metadata.ListItemNumber = resolver.GetCurrentCount(numId, level);

                // Also store the LevelText if not already set
                if (string.IsNullOrEmpty(segment.Metadata.LevelTextFormat))
                {
                    segment.Metadata.LevelTextFormat = levelInfo.LevelText;
                }
            }
        }
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

        // Extract text content and inline elements
        var (text, inlineElements) = ExtractParagraphContent(paragraph, wordDoc);
        segment.Content = text;
        segment.InlineElements = inlineElements;

        // Detect segment type based on style
        segment.Type = DetectSegmentType(paragraph, styleId, wordDoc);

        // Extract formatting
        ExtractParagraphFormatting(paragraph, segment, wordDoc);

        // Check for images in paragraph
        ProcessInlineImages(paragraph, segment, wordDoc);

        return segment;
    }

    /// <summary>
    /// Extracts paragraph content including inline elements (hyperlinks, formatting, footnotes)
    /// </summary>
    private (string text, List<InlineElement> inlineElements) ExtractParagraphContent(
        Paragraph paragraph,
        WordprocessingDocument wordDoc)
    {
        var textBuilder = new System.Text.StringBuilder();
        var elements = new List<InlineElement>();
        int currentOffset = 0;

        foreach (var child in paragraph.ChildElements)
        {
            if (child is Run run)
            {
                var (runText, runElements) = ProcessRun(run, currentOffset, wordDoc);
                textBuilder.Append(runText);
                elements.AddRange(runElements);
                currentOffset += runText.Length;
            }
            else if (child is Hyperlink hyperlink)
            {
                var (linkText, linkElement) = ProcessHyperlink(hyperlink, wordDoc, currentOffset);
                textBuilder.Append(linkText);
                if (linkElement != null)
                    elements.Add(linkElement);
                currentOffset += linkText.Length;
            }
        }

        return (textBuilder.ToString(), elements);
    }

    /// <summary>
    /// Process a Run element to extract text and formatting
    /// </summary>
    private (string text, List<InlineElement> elements) ProcessRun(Run run, int startOffset, WordprocessingDocument wordDoc)
    {
        var elements = new List<InlineElement>();

        // Check for footnote/endnote references first
        var footnoteRef = run.GetFirstChild<FootnoteReference>();
        if (footnoteRef != null && _settings.EnableFootnoteConversion)
        {
            elements.Add(new FootnoteReferenceElement
            {
                StartOffset = startOffset,
                EndOffset = startOffset,
                Text = "",
                NoteId = (int)(footnoteRef.Id?.Value ?? 0),
                IsEndnote = false
            });
            // Footnote references don't contribute to text
            return ("", elements);
        }

        var endnoteRef = run.GetFirstChild<EndnoteReference>();
        if (endnoteRef != null && _settings.EnableFootnoteConversion)
        {
            elements.Add(new FootnoteReferenceElement
            {
                StartOffset = startOffset,
                EndOffset = startOffset,
                Text = "",
                NoteId = (int)(endnoteRef.Id?.Value ?? 0),
                IsEndnote = true
            });
            return ("", elements);
        }

        // Extract text from run
        var text = string.Join("", run.Descendants<Text>().Select(t => t.Text));
        if (string.IsNullOrEmpty(text))
            return (text, elements);

        // Extract formatting
        var runProps = run.RunProperties;
        var fontFamily = runProps?.RunFonts?.Ascii?.Value ??
                         runProps?.RunFonts?.HighAnsi?.Value;

        var textRun = new TextRun
        {
            StartOffset = startOffset,
            EndOffset = startOffset + text.Length,
            Text = text,
            IsBold = runProps?.Bold != null || runProps?.BoldComplexScript != null,
            IsItalic = runProps?.Italic != null || runProps?.ItalicComplexScript != null,
            IsUnderline = runProps?.Underline != null && runProps.Underline.Val?.Value != UnderlineValues.None,
            IsStrikethrough = runProps?.Strike != null || runProps?.DoubleStrike != null,
            FontFamily = fontFamily,
            IsCode = _settings.EnableCodeBlockDetection && IsMonospaceFont(fontFamily)
        };

        elements.Add(textRun);
        return (text, elements);
    }

    /// <summary>
    /// Process a Hyperlink element
    /// </summary>
    private (string text, HyperlinkElement? element) ProcessHyperlink(
        Hyperlink hyperlink,
        WordprocessingDocument wordDoc,
        int startOffset)
    {
        // Extract text from hyperlink runs
        var text = string.Join("", hyperlink.Descendants<Text>().Select(t => t.Text));

        if (!_settings.EnableHyperlinkConversion)
            return (text, null);

        string url = "";

        // Resolve URL from relationship ID
        var rId = hyperlink.Id?.Value;
        if (!string.IsNullOrEmpty(rId))
        {
            try
            {
                var relationship = wordDoc.MainDocumentPart?.HyperlinkRelationships
                    .FirstOrDefault(r => r.Id == rId);
                if (relationship != null)
                {
                    url = relationship.Uri.ToString();
                }
            }
            catch
            {
                // Best effort - hyperlink relationship may be invalid
            }
        }
        // Check for anchor (internal bookmark links)
        else if (!string.IsNullOrEmpty(hyperlink.Anchor?.Value))
        {
            url = "#" + hyperlink.Anchor.Value;
        }

        if (string.IsNullOrEmpty(text) && string.IsNullOrEmpty(url))
            return (text, null);

        var element = new HyperlinkElement
        {
            StartOffset = startOffset,
            EndOffset = startOffset + text.Length,
            Text = text,
            Url = url,
            Tooltip = hyperlink.Tooltip?.Value
        };

        return (text, element);
    }

    /// <summary>
    /// Check if a font family is a monospace font
    /// </summary>
    private bool IsMonospaceFont(string? fontFamily)
    {
        if (string.IsNullOrEmpty(fontFamily))
            return false;

        return _settings.MonospaceFonts.Any(f =>
            fontFamily.Contains(f, StringComparison.OrdinalIgnoreCase));
    }

    private SegmentType DetectSegmentType(Paragraph paragraph, string? styleId, WordprocessingDocument wordDoc)
    {
        // Check if it's a heading based on style name (English built-in styles)
        if (!string.IsNullOrEmpty(styleId))
        {
            if (styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
                return SegmentType.Heading;

            if (styleId.StartsWith("Title", StringComparison.OrdinalIgnoreCase))
                return SegmentType.Heading;
        }

        // Check outline level directly on paragraph
        var outlineLevel = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (outlineLevel.HasValue)
        {
            return SegmentType.Heading;
        }

        // Check outline level from style definition (handles localized styles like Dutch "Kop1")
        // This must be checked BEFORE NumberingProperties because heading styles can have numbering
        if (!string.IsNullOrEmpty(styleId))
        {
            var styleOutlineLevel = GetStyleOutlineLevel(styleId, wordDoc);
            if (styleOutlineLevel.HasValue)
            {
                return SegmentType.Heading;
            }
        }

        // Check for list item
        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        if (numPr != null)
        {
            return SegmentType.ListItem;
        }

        return SegmentType.Paragraph;
    }

    /// <summary>
    /// Get the outline level defined in a style definition.
    /// Returns null if the style has no outline level (not a heading style).
    /// </summary>
    private int? GetStyleOutlineLevel(string styleId, WordprocessingDocument wordDoc)
    {
        var stylesPart = wordDoc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null)
            return null;

        var style = stylesPart.Styles.Elements<Style>()
            .FirstOrDefault(s => s.StyleId?.Value == styleId);

        if (style == null)
            return null;

        // Check if this style has an outline level in its paragraph properties
        var outlineLevel = style.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
        if (outlineLevel.HasValue)
            return outlineLevel.Value;

        // Check if this style is based on another style that might have outline level
        var basedOnStyleId = style.BasedOn?.Val?.Value;
        if (!string.IsNullOrEmpty(basedOnStyleId))
        {
            return GetStyleOutlineLevel(basedOnStyleId, wordDoc);
        }

        return null;
    }

    private void ExtractParagraphFormatting(Paragraph paragraph, Segment segment, WordprocessingDocument wordDoc)
    {
        var runProps = paragraph.Descendants<RunProperties>().FirstOrDefault();
        if (runProps != null)
        {
            segment.Metadata.IsBold = runProps.Bold != null;
            segment.Metadata.IsItalic = runProps.Italic != null;
            segment.Metadata.IsUnderline = runProps.Underline != null;
        }

        // Check outline level directly on paragraph first
        var outlineLevel = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (outlineLevel.HasValue)
        {
            segment.Metadata.OutlineLevel = outlineLevel.Value;
        }
        else
        {
            // Fall back to style definition for outline level (handles localized styles)
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrEmpty(styleId))
            {
                var styleOutlineLevel = GetStyleOutlineLevel(styleId, wordDoc);
                if (styleOutlineLevel.HasValue)
                {
                    segment.Metadata.OutlineLevel = styleOutlineLevel.Value;
                }
            }
        }

        // Check for numbering on the paragraph
        var numPr = paragraph.ParagraphProperties?.NumberingProperties;
        if (numPr != null)
        {
            var numId = numPr.NumberingId?.Val?.Value;

            // numId of 0 means "explicitly no numbering" (overrides style numbering)
            if (numId.HasValue && numId.Value != 0)
            {
                segment.Metadata.NumberingId = numId.Value.ToString();
                segment.Metadata.NumberingLevel = numPr.NumberingLevelReference?.Val?.Value;

                // Detect numbered vs bulleted list from numbering definition
                DetectNumberingFormat(wordDoc, segment, numPr);
            }
            else if (numId.HasValue && numId.Value == 0)
            {
                // Mark that numbering is explicitly disabled (don't apply style numbering later)
                segment.Metadata.NumberingExplicitlyDisabled = true;
            }
        }
        // Note: Style-based numbering continuation is handled in a second pass
        // (see ActivateStyleNumbering) to properly handle the "from here on" behavior
        // where front-matter headings don't have numbers but content headings do.
    }

    /// <summary>
    /// Get numbering properties from a style definition (used for heading styles with numbering)
    /// </summary>
    private (int numId, int level)? GetStyleNumberingProperties(string styleId, WordprocessingDocument wordDoc)
    {
        var stylesPart = wordDoc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null)
            return null;

        var style = stylesPart.Styles.Elements<Style>()
            .FirstOrDefault(s => s.StyleId?.Value == styleId);

        if (style == null)
            return null;

        // Check if this style has numbering properties
        var numPr = style.StyleParagraphProperties?.NumberingProperties;
        if (numPr != null)
        {
            var numId = numPr.NumberingId?.Val?.Value;
            var level = numPr.NumberingLevelReference?.Val?.Value ?? 0;

            if (numId.HasValue)
                return (numId.Value, level);
        }

        // Check if based on another style
        var basedOnStyleId = style.BasedOn?.Val?.Value;
        if (!string.IsNullOrEmpty(basedOnStyleId))
        {
            return GetStyleNumberingProperties(basedOnStyleId, wordDoc);
        }

        return null;
    }

    /// <summary>
    /// Detect numbering format using numId and level values directly
    /// </summary>
    private void DetectNumberingFormatFromValues(WordprocessingDocument wordDoc, Segment segment, int numId, int levelIndex)
    {
        var numberingPart = wordDoc.MainDocumentPart?.NumberingDefinitionsPart;
        if (numberingPart?.Numbering == null)
            return;

        // Find the numbering instance
        var numberingInstance = numberingPart.Numbering
            .Elements<NumberingInstance>()
            .FirstOrDefault(ni => ni.NumberID?.Value == numId);

        if (numberingInstance?.AbstractNumId?.Val == null)
            return;

        // Find the abstract numbering definition
        var abstractNumId = numberingInstance.AbstractNumId.Val.Value;
        var abstractNum = numberingPart.Numbering
            .Elements<AbstractNum>()
            .FirstOrDefault(an => an.AbstractNumberId?.Value == abstractNumId);

        if (abstractNum == null)
            return;

        // Find the level definition
        var level = abstractNum.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == levelIndex);

        if (level?.NumberingFormat?.Val == null)
            return;

        // Determine if numbered based on numbering format
        var format = level.NumberingFormat.Val.Value;
        segment.Metadata.IsNumberedList = IsNumberedFormat(format);

        // Get start value if available
        if (level.StartNumberingValue?.Val != null)
        {
            segment.Metadata.NumberingStartValue = level.StartNumberingValue.Val.Value;
        }

        // Extract LevelText format string
        segment.Metadata.LevelTextFormat = level.LevelText?.Val?.Value;
    }

    private void DetectNumberingFormat(WordprocessingDocument wordDoc, Segment segment, NumberingProperties numPr)
    {
        var numberingPart = wordDoc.MainDocumentPart?.NumberingDefinitionsPart;
        if (numberingPart?.Numbering == null)
            return;

        var numId = numPr.NumberingId?.Val?.Value;
        var levelIndex = numPr.NumberingLevelReference?.Val?.Value ?? 0;

        if (numId == null)
            return;

        // Find the numbering instance
        var numberingInstance = numberingPart.Numbering
            .Elements<NumberingInstance>()
            .FirstOrDefault(ni => ni.NumberID?.Value == numId);

        if (numberingInstance?.AbstractNumId?.Val == null)
            return;

        // Find the abstract numbering definition
        var abstractNumId = numberingInstance.AbstractNumId.Val.Value;
        var abstractNum = numberingPart.Numbering
            .Elements<AbstractNum>()
            .FirstOrDefault(an => an.AbstractNumberId?.Value == abstractNumId);

        if (abstractNum == null)
            return;

        // Find the level definition
        var level = abstractNum.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == levelIndex);

        if (level?.NumberingFormat?.Val == null)
            return;

        // Determine if numbered based on numbering format
        var format = level.NumberingFormat.Val.Value;
        segment.Metadata.IsNumberedList = IsNumberedFormat(format);

        // Get start value if available
        if (level.StartNumberingValue?.Val != null)
        {
            segment.Metadata.NumberingStartValue = level.StartNumberingValue.Val.Value;
        }

        // Extract LevelText format string for numbering prefix resolution
        segment.Metadata.LevelTextFormat = level.LevelText?.Val?.Value;
    }

    private static bool IsNumberedFormat(NumberFormatValues format)
    {
        // NumberFormatValues uses EnumValue comparison
        if (format == NumberFormatValues.Decimal ||
            format == NumberFormatValues.DecimalZero ||
            format == NumberFormatValues.LowerLetter ||
            format == NumberFormatValues.UpperLetter ||
            format == NumberFormatValues.LowerRoman ||
            format == NumberFormatValues.UpperRoman ||
            format == NumberFormatValues.Ordinal ||
            format == NumberFormatValues.CardinalText ||
            format == NumberFormatValues.OrdinalText)
        {
            return true;
        }

        // Bullet and other formats are not numbered
        return false;
    }

    private void ProcessInlineImages(Paragraph paragraph, Segment segment, WordprocessingDocument wordDoc)
    {
        // Look for Drawing elements which contain images
        var drawings = paragraph.Descendants<Drawing.Blip>().ToList();
        if (!drawings.Any())
            return;

        // Get the first image's relationship ID
        var firstBlip = drawings.First();
        var embedId = firstBlip.Embed?.Value;

        if (!string.IsNullOrEmpty(embedId))
        {
            segment.Type = SegmentType.Image;
            segment.Metadata.AdditionalProperties["ImageRelationshipId"] = embedId;

            // Try to extract alt text from DocPr (Document Properties) element
            var docPr = paragraph.Descendants<DrawingWp.DocProperties>().FirstOrDefault();
            if (docPr != null)
            {
                var altText = docPr.Description?.Value ?? docPr.Name?.Value;
                if (!string.IsNullOrEmpty(altText))
                {
                    segment.Content = altText;
                }
            }

            // If multiple images in paragraph, add diagnostic
            if (drawings.Count > 1)
            {
                segment.AddDiagnostic(
                    DiagnosticLevel.Info,
                    Diagnostics.DiagnosticCodes.IMAGE_ALT_TEXT_MISSING,
                    $"Paragraph contains {drawings.Count} images. Only the first is converted.");
            }
        }
    }

    private Segment ProcessTable(Table table, int orderIndex)
    {
        var segment = new Segment
        {
            OrderIndex = orderIndex,
            Type = SegmentType.Table
        };

        // Extract table content - all rows and cells
        var rows = table.Elements<TableRow>().ToList();
        var rowCount = rows.Count;
        var tableData = new List<List<string>>();
        var maxColumns = 0;

        foreach (var row in rows)
        {
            var rowData = new List<string>();
            foreach (var cell in row.Elements<TableCell>())
            {
                // Extract text from all paragraphs in the cell
                var cellText = string.Join(" ",
                    cell.Elements<Paragraph>()
                        .Select(p => string.Join("", p.Descendants<Text>().Select(t => t.Text))));

                // Handle merged cells (GridSpan) by adding empty columns
                var gridSpan = cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
                rowData.Add(cellText.Trim());

                // Add empty strings for spanned columns (GFM doesn't support colspan)
                for (int i = 1; i < gridSpan; i++)
                {
                    rowData.Add("");
                }
            }
            tableData.Add(rowData);
            if (rowData.Count > maxColumns)
                maxColumns = rowData.Count;
        }

        // Normalize rows to have same column count
        foreach (var row in tableData)
        {
            while (row.Count < maxColumns)
                row.Add("");
        }

        // Store table data for converter
        segment.Metadata.AdditionalProperties["TableData"] = tableData;
        segment.Metadata.AdditionalProperties["RowCount"] = rowCount;
        segment.Metadata.AdditionalProperties["ColumnCount"] = maxColumns;

        // Create content summary for display
        segment.Content = $"Table with {rowCount} rows, {maxColumns} columns";

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
        var body = wordDoc.MainDocumentPart?.Document.Body;

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

        // Check for track changes (revisions)
        var hasRevisions = body?.Descendants<InsertedRun>().Any() == true ||
                           body?.Descendants<DeletedRun>().Any() == true ||
                           body?.Descendants<Inserted>().Any() == true ||
                           body?.Descendants<Deleted>().Any() == true;
        if (hasRevisions)
        {
            document.AddDiagnostic(
                DiagnosticLevel.Warning,
                Diagnostics.DiagnosticCodes.TRACK_CHANGES_IGNORED,
                "Document contains tracked changes. Only the current version is converted.");
        }

        // Extract or detect footnotes
        if (wordDoc.MainDocumentPart?.FootnotesPart != null)
        {
            var footnotes = wordDoc.MainDocumentPart.FootnotesPart.Footnotes?
                .Elements<Footnote>()
                .Where(f => f.Type?.Value != FootnoteEndnoteValues.Separator &&
                            f.Type?.Value != FootnoteEndnoteValues.ContinuationSeparator);
            if (footnotes?.Any() == true)
            {
                if (_settings.EnableFootnoteConversion)
                {
                    ExtractFootnotes(footnotes, document);
                }
                else
                {
                    document.AddDiagnostic(
                        DiagnosticLevel.Info,
                        Diagnostics.DiagnosticCodes.FOOTNOTE_IGNORED,
                        $"Document contains {footnotes.Count()} footnote(s) which are not included in the conversion");
                }
            }
        }

        // Extract or detect endnotes
        if (wordDoc.MainDocumentPart?.EndnotesPart != null)
        {
            var endnotes = wordDoc.MainDocumentPart.EndnotesPart.Endnotes?
                .Elements<Endnote>()
                .Where(e => e.Type?.Value != FootnoteEndnoteValues.Separator &&
                            e.Type?.Value != FootnoteEndnoteValues.ContinuationSeparator);
            if (endnotes?.Any() == true)
            {
                if (_settings.EnableFootnoteConversion)
                {
                    ExtractEndnotes(endnotes, document);
                }
                else
                {
                    document.AddDiagnostic(
                        DiagnosticLevel.Info,
                        Diagnostics.DiagnosticCodes.ENDNOTE_IGNORED,
                        $"Document contains {endnotes.Count()} endnote(s) which are not included in the conversion");
                }
            }
        }

        // Check for text boxes (using VML or drawing canvas)
        var hasTextBoxes = body?.Descendants<Vml.TextBox>().Any() == true ||
                           body?.Descendants<DrawingWp.WrapNone>()
                               .Where(w => w.Parent?.Descendants<Drawing.Blip>().Any() != true)
                               .Any() == true;
        if (hasTextBoxes)
        {
            document.AddDiagnostic(
                DiagnosticLevel.Warning,
                Diagnostics.DiagnosticCodes.TEXT_BOX_IGNORED,
                "Document contains text boxes which are not included in the conversion");
        }

        // Check for shapes (VML shapes)
        var hasShapes = body?.Descendants<Vml.Shape>().Any() == true ||
                        body?.Descendants<Vml.Shapetype>().Any() == true ||
                        body?.Descendants<Vml.Oval>().Any() == true ||
                        body?.Descendants<Vml.Rectangle>().Any() == true ||
                        body?.Descendants<Vml.Line>().Any() == true;
        if (hasShapes)
        {
            document.AddDiagnostic(
                DiagnosticLevel.Warning,
                Diagnostics.DiagnosticCodes.SHAPE_IGNORED,
                "Document contains shapes which are not included in the conversion");
        }

        // Check for SmartArt (diagrams) - look for dgm:relIds elements in the document
        var hasDiagrams = body?.Descendants<DocumentFormat.OpenXml.Drawing.Diagrams.RelationshipIds>().Any() == true ||
                          wordDoc.MainDocumentPart?.Parts.Any(p =>
                              p.OpenXmlPart.ContentType.Contains("diagram")) == true;
        if (hasDiagrams)
        {
            document.AddDiagnostic(
                DiagnosticLevel.Warning,
                Diagnostics.DiagnosticCodes.SMARTART_IGNORED,
                "Document contains SmartArt graphics which are not included in the conversion");
        }

        // Check for section breaks and multi-column layouts
        var sectionProps = body?.Descendants<SectionProperties>().ToList() ?? new List<SectionProperties>();
        if (sectionProps.Count > 1)
        {
            document.AddDiagnostic(
                DiagnosticLevel.Info,
                Diagnostics.DiagnosticCodes.SECTION_BREAK_IGNORED,
                $"Document contains {sectionProps.Count - 1} section break(s). Section formatting is not preserved.");
        }

        // Check for multi-column layouts
        var multiColumnSections = sectionProps.Where(sp =>
        {
            var columns = sp.GetFirstChild<Columns>();
            return columns?.ColumnCount?.Value > 1 ||
                   (columns?.HasChildren == true && columns.ChildElements.Count > 1);
        });
        if (multiColumnSections.Any())
        {
            document.AddDiagnostic(
                DiagnosticLevel.Warning,
                Diagnostics.DiagnosticCodes.MULTI_COLUMN_LAYOUT_IGNORED,
                "Document contains multi-column layouts which are converted as single-column");
        }

        // Check for floating images (non-inline positioned images)
        var floatingImages = body?.Descendants<DrawingWp.Anchor>().Any() == true;
        if (floatingImages)
        {
            document.AddDiagnostic(
                DiagnosticLevel.Warning,
                Diagnostics.DiagnosticCodes.FLOATING_IMAGE_IGNORED,
                "Document contains floating/positioned images. Position information is not preserved.");
        }

        // Check for Word fields (TOC, page numbers, etc.)
        var hasFields = body?.Descendants<FieldCode>().Any() == true ||
                        body?.Descendants<SimpleField>().Any() == true;
        if (hasFields)
        {
            var fieldTypes = new HashSet<string>();
            foreach (var field in body?.Descendants<FieldCode>() ?? Enumerable.Empty<FieldCode>())
            {
                var fieldText = field.Text?.Trim().Split(' ').FirstOrDefault()?.ToUpperInvariant();
                if (!string.IsNullOrEmpty(fieldText))
                    fieldTypes.Add(fieldText);
            }
            foreach (var field in body?.Descendants<SimpleField>() ?? Enumerable.Empty<SimpleField>())
            {
                var instruction = field.Instruction?.Value?.Trim().Split(' ').FirstOrDefault()?.ToUpperInvariant();
                if (!string.IsNullOrEmpty(instruction))
                    fieldTypes.Add(instruction);
            }

            var fieldList = fieldTypes.Count > 0 ? $" ({string.Join(", ", fieldTypes)})" : "";
            document.AddDiagnostic(
                DiagnosticLevel.Info,
                Diagnostics.DiagnosticCodes.FIELD_IGNORED,
                $"Document contains Word fields{fieldList} which show static values only");
        }
    }

    /// <summary>
    /// Extract footnote definitions from the document
    /// </summary>
    private void ExtractFootnotes(IEnumerable<Footnote> footnotes, DocumentModel document)
    {
        foreach (var footnote in footnotes)
        {
            var id = (int)(footnote.Id?.Value ?? 0);
            // Skip special footnotes (separator, continuation)
            if (id < 1)
                continue;

            var content = ExtractNoteContent(footnote);
            document.Footnotes.Add(new FootnoteDefinition
            {
                Id = id,
                IsEndnote = false,
                Content = content
            });
        }
    }

    /// <summary>
    /// Extract endnote definitions from the document
    /// </summary>
    private void ExtractEndnotes(IEnumerable<Endnote> endnotes, DocumentModel document)
    {
        foreach (var endnote in endnotes)
        {
            var id = (int)(endnote.Id?.Value ?? 0);
            // Skip special endnotes (separator, continuation)
            if (id < 1)
                continue;

            var content = ExtractNoteContent(endnote);
            document.Endnotes.Add(new FootnoteDefinition
            {
                Id = id,
                IsEndnote = true,
                Content = content
            });
        }
    }

    /// <summary>
    /// Extract text content from a footnote or endnote
    /// </summary>
    private string ExtractNoteContent(OpenXmlCompositeElement note)
    {
        // Get all paragraphs in the note and join their text
        var paragraphs = note.Elements<Paragraph>();
        var textParts = new List<string>();

        foreach (var para in paragraphs)
        {
            var paraText = string.Join("", para.Descendants<Text>().Select(t => t.Text));
            if (!string.IsNullOrWhiteSpace(paraText))
                textParts.Add(paraText.Trim());
        }

        return string.Join(" ", textParts);
    }
}
