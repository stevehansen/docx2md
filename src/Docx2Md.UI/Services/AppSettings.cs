using System.Collections.Generic;

namespace Docx2Md.UI.Services;

/// <summary>
/// Application settings model for persistence between sessions.
/// Stores UI state, conversion settings, and recent files.
/// </summary>
public class AppSettings
{
    // UI State
    public bool ShowDocxPreview { get; set; } = true;
    public bool ShowSegmentInspector { get; set; } = true;
    public bool ShowMarkdownPreview { get; set; } = true;
    public bool ShowRawMarkdown { get; set; } = false;

    // Conversion Settings
    public bool EnableStyleBasedHeadingDetection { get; set; } = true;
    public bool InferHeadingsFromFormatting { get; set; } = true;
    public bool ConvertUnderlineToEmphasis { get; set; } = false;
    public bool GenerateDiagnosticReport { get; set; } = true;

    // Recent Files
    public List<string> RecentFiles { get; set; } = new();
}
