namespace Docx2Md.Core.Models;

/// <summary>
/// Represents a DOCX document with its segments and metadata
/// </summary>
public class DocumentModel
{
    /// <summary>
    /// Source file path
    /// </summary>
    public string SourcePath { get; set; } = string.Empty;

    /// <summary>
    /// Document title (if available)
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// All segments in document order
    /// </summary>
    public List<Segment> Segments { get; set; } = new();

    /// <summary>
    /// Embedded images extracted from the document
    /// </summary>
    public List<ImageInfo> Images { get; set; } = new();

    /// <summary>
    /// Document-level diagnostics
    /// </summary>
    public List<Diagnostic> DocumentDiagnostics { get; set; } = new();

    /// <summary>
    /// Metadata about the document
    /// </summary>
    public Dictionary<string, object> Metadata { get; set; } = new();

    /// <summary>
    /// Get all diagnostics (document-level and segment-level)
    /// </summary>
    public IEnumerable<Diagnostic> GetAllDiagnostics()
    {
        foreach (var diag in DocumentDiagnostics)
            yield return diag;

        foreach (var segment in Segments)
        {
            foreach (var diag in segment.Diagnostics)
                yield return diag;
        }
    }

    /// <summary>
    /// Get diagnostics by level
    /// </summary>
    public IEnumerable<Diagnostic> GetDiagnosticsByLevel(DiagnosticLevel level)
    {
        return GetAllDiagnostics().Where(d => d.Level == level);
    }

    /// <summary>
    /// Add a document-level diagnostic
    /// </summary>
    public void AddDiagnostic(DiagnosticLevel level, string code, string message, string? details = null)
    {
        DocumentDiagnostics.Add(new Diagnostic(level, code, message, details));
    }
}
