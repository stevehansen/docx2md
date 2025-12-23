namespace Docx2Md.Core.Models;

/// <summary>
/// Represents a diagnostic message for a segment conversion
/// </summary>
public class Diagnostic
{
    public DiagnosticLevel Level { get; set; }
    public string Code { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public string? Details { get; set; }

    public Diagnostic() { }

    public Diagnostic(DiagnosticLevel level, string code, string message, string? details = null)
    {
        Level = level;
        Code = code;
        Message = message;
        Details = details;
    }

    public override string ToString()
    {
        return $"[{Level}] {Code}: {Message}";
    }
}
