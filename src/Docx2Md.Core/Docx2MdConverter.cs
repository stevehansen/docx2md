using Docx2Md.Core.Models;
using Docx2Md.Core.Parsing;
using Docx2Md.Core.Conversion;
using Docx2Md.Core.Export;

namespace Docx2Md.Core;

/// <summary>
/// Main orchestrator for DOCX to Markdown conversion
/// Coordinates parsing, conversion, and export
/// </summary>
public class Docx2MdConverter
{
    private readonly ConversionSettings _settings;
    private readonly DocxParser _parser;
    private readonly MarkdownConverter _converter;
    private readonly MarkdownExporter _exporter;

    public Docx2MdConverter(ConversionSettings? settings = null)
    {
        _settings = settings ?? ConversionSettings.Default;
        _parser = new DocxParser(_settings);
        _converter = new MarkdownConverter(_settings);
        _exporter = new MarkdownExporter(_settings);
    }

    /// <summary>
    /// Convert a DOCX file to Markdown
    /// </summary>
    /// <param name="inputPath">Path to input DOCX file</param>
    /// <param name="outputPath">Path to output Markdown file</param>
    /// <returns>Document model with all segments and diagnostics</returns>
    public DocumentModel ConvertFile(string inputPath, string outputPath)
    {
        // Step 1: Parse DOCX and extract segments
        var document = _parser.Parse(inputPath);

        // Step 2: Convert segments to Markdown
        _converter.ConvertDocument(document);

        // Step 3: Export to file
        _exporter.ExportMarkdown(document, outputPath);

        return document;
    }

    /// <summary>
    /// Parse a DOCX file without converting or exporting
    /// Useful for UI scenarios where conversion happens separately
    /// </summary>
    public DocumentModel ParseFile(string inputPath)
    {
        return _parser.Parse(inputPath);
    }

    /// <summary>
    /// Convert an already-parsed document
    /// </summary>
    public void ConvertDocument(DocumentModel document)
    {
        _converter.ConvertDocument(document);
    }

    /// <summary>
    /// Export a converted document
    /// </summary>
    public void ExportDocument(DocumentModel document, string outputPath)
    {
        _exporter.ExportMarkdown(document, outputPath);
    }

    /// <summary>
    /// Get Markdown preview without writing to file
    /// </summary>
    public string GetMarkdownPreview(DocumentModel document)
    {
        return _exporter.GenerateMarkdown(document);
    }
}
