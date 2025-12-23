using Docx2Md.Core.Models;
using System.Text;
using System.Text.Json;

namespace Docx2Md.Core.Export;

/// <summary>
/// Exports Markdown and diagnostics (PRD Section 9)
/// </summary>
public class MarkdownExporter
{
    private readonly ConversionSettings _settings;

    public MarkdownExporter(ConversionSettings? settings = null)
    {
        _settings = settings ?? ConversionSettings.Default;
    }

    /// <summary>
    /// Export document to Markdown file
    /// </summary>
    public void ExportMarkdown(DocumentModel document, string outputPath)
    {
        var markdown = GenerateMarkdown(document);
        
        // Ensure output directory exists
        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(outputPath, markdown);

        // Export images if any
        ExportImages(document, outputPath);

        // Export diagnostics if enabled
        if (_settings.GenerateDiagnosticReport)
        {
            ExportDiagnostics(document, outputPath);
        }
    }

    /// <summary>
    /// Generate Markdown content from document
    /// </summary>
    public string GenerateMarkdown(DocumentModel document)
    {
        var sb = new StringBuilder();

        // Add title if available
        if (!string.IsNullOrEmpty(document.Title))
        {
            sb.AppendLine($"# {document.Title}");
            sb.AppendLine();
        }

        // Add segments
        foreach (var segment in document.Segments)
        {
            if (segment.ExcludeFromOutput)
                continue;

            var markdown = segment.EffectiveMarkdown;
            if (!string.IsNullOrWhiteSpace(markdown))
            {
                sb.AppendLine(markdown);
                sb.AppendLine();
            }
        }

        // Add Lost & Found section if enabled
        if (_settings.AppendLostAndFoundSection)
        {
            AppendLostAndFound(document, sb);
        }

        return sb.ToString().TrimEnd();
    }

    private void AppendLostAndFound(DocumentModel document, StringBuilder sb)
    {
        var excludedSegments = document.Segments
            .Where(s => s.ExcludeFromOutput && !string.IsNullOrWhiteSpace(s.Content))
            .ToList();

        if (excludedSegments.Count == 0)
            return;

        sb.AppendLine();
        sb.AppendLine("---");
        sb.AppendLine();
        sb.AppendLine("## Lost & Found");
        sb.AppendLine();
        sb.AppendLine("*The following content was excluded from the main conversion:*");
        sb.AppendLine();

        foreach (var segment in excludedSegments)
        {
            sb.AppendLine($"### Segment {segment.OrderIndex} ({segment.Type})");
            sb.AppendLine();
            sb.AppendLine(segment.Content);
            sb.AppendLine();
        }
    }

    private void ExportImages(DocumentModel document, string markdownPath)
    {
        if (document.Images.Count == 0)
            return;

        var markdownDir = Path.GetDirectoryName(markdownPath) ?? ".";
        var imageDir = Path.Combine(markdownDir, _settings.ImageOutputFolder);

        if (!Directory.Exists(imageDir))
        {
            Directory.CreateDirectory(imageDir);
        }

        foreach (var image in document.Images)
        {
            try
            {
                var fileName = image.GetFileName();
                var imagePath = Path.Combine(imageDir, fileName);
                File.WriteAllBytes(imagePath, image.Data);
            }
            catch (Exception ex)
            {
                document.AddDiagnostic(
                    DiagnosticLevel.Error,
                    Diagnostics.DiagnosticCodes.IMAGE_EXTRACTION_FAILED,
                    $"Failed to export image {image.Id}",
                    ex.Message);
            }
        }
    }

    private void ExportDiagnostics(DocumentModel document, string markdownPath)
    {
        var diagnostics = document.GetAllDiagnostics().ToList();
        if (diagnostics.Count == 0)
            return;

        var basePath = Path.GetFileNameWithoutExtension(markdownPath);
        var directory = Path.GetDirectoryName(markdownPath) ?? ".";

        if (_settings.DiagnosticReportFormat == DiagnosticReportFormat.Json)
        {
            ExportDiagnosticsJson(diagnostics, directory, basePath);
        }
        else
        {
            ExportDiagnosticsMarkdown(diagnostics, document, directory, basePath);
        }
    }

    private void ExportDiagnosticsJson(List<Diagnostic> diagnostics, string directory, string basePath)
    {
        var jsonPath = Path.Combine(directory, $"{basePath}_diagnostics.json");
        var json = JsonSerializer.Serialize(diagnostics, new JsonSerializerOptions
        {
            WriteIndented = true
        });
        File.WriteAllText(jsonPath, json);
    }

    private void ExportDiagnosticsMarkdown(List<Diagnostic> diagnostics, DocumentModel document, string directory, string basePath)
    {
        var mdPath = Path.Combine(directory, $"{basePath}_diagnostics.md");
        var sb = new StringBuilder();

        sb.AppendLine("# Conversion Diagnostics");
        sb.AppendLine();
        sb.AppendLine($"**Source:** {Path.GetFileName(document.SourcePath)}");
        sb.AppendLine($"**Generated:** {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine();

        // Group by level
        var byLevel = diagnostics.GroupBy(d => d.Level).OrderBy(g => g.Key);

        foreach (var group in byLevel)
        {
            sb.AppendLine($"## {group.Key} ({group.Count()})");
            sb.AppendLine();

            foreach (var diag in group)
            {
                sb.AppendLine($"- **{diag.Code}**: {diag.Message}");
                if (!string.IsNullOrEmpty(diag.Details))
                {
                    sb.AppendLine($"  - *Details:* {diag.Details}");
                }
            }
            sb.AppendLine();
        }

        File.WriteAllText(mdPath, sb.ToString());
    }
}
