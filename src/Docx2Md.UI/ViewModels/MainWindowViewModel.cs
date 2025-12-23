using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Platform.Storage;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Docx2Md.Core;
using Docx2Md.Core.Models;

namespace Docx2Md.UI.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    private readonly Docx2MdConverter _converter;
    private DocumentModel? _document;
    private string _currentFilePath = string.Empty;

    // Delegate for file dialogs (injected by View)
    public Func<Task<string?>>? ShowOpenFileDialog { get; set; }
    public Func<Task<string?>>? ShowSaveFileDialog { get; set; }

    /// <summary>
    /// Available heading levels for override ComboBox (null = use original)
    /// </summary>
    public static List<int?> AvailableHeadingLevels { get; } = new() { null, 1, 2, 3, 4, 5, 6 };

    /// <summary>
    /// Available segment types for override ComboBox (null = use original)
    /// </summary>
    public static List<SegmentType?> AvailableSegmentTypes { get; } = new()
    {
        null,
        SegmentType.Heading,
        SegmentType.Paragraph,
        SegmentType.ListItem,
        SegmentType.Table,
        SegmentType.Image,
        SegmentType.PageBreak,
        SegmentType.SectionBreak,
        SegmentType.Unknown
    };

    [ObservableProperty]
    private string _statusText = "Ready. Open a DOCX file to begin.";

    [ObservableProperty]
    private string _documentTitle = "No document loaded";

    [ObservableProperty]
    private string _docxPreviewText = "Open a DOCX file to preview its content here.";

    [ObservableProperty]
    private string _markdownOutput = string.Empty;

    [ObservableProperty]
    private ObservableCollection<SegmentViewModel> _segments = new();

    [ObservableProperty]
    private SegmentViewModel? _selectedSegment;

    [ObservableProperty]
    private bool _showDocxPreview = true;

    [ObservableProperty]
    private bool _showSegmentInspector = true;

    [ObservableProperty]
    private bool _showMarkdownPreview = true;

    [ObservableProperty]
    private bool _showRawMarkdown = false;

    // Settings properties with change notification
    [ObservableProperty]
    private bool _enableStyleBasedHeadingDetection = true;

    [ObservableProperty]
    private bool _inferHeadingsFromFormatting = true;

    [ObservableProperty]
    private bool _convertUnderlineToEmphasis = false;

    [ObservableProperty]
    private bool _generateDiagnosticReport = true;

    [ObservableProperty]
    private bool _hasDocument = false;

    [ObservableProperty]
    private ConversionSettings _settings = new();

    public MainWindowViewModel()
    {
        _converter = new Docx2MdConverter(Settings);
    }

    // View toggle commands
    [RelayCommand]
    private void ToggleDocxPreview() => ShowDocxPreview = !ShowDocxPreview;

    [RelayCommand]
    private void ToggleSegmentInspector() => ShowSegmentInspector = !ShowSegmentInspector;

    [RelayCommand]
    private void ToggleMarkdownPreview() => ShowMarkdownPreview = !ShowMarkdownPreview;

    // Settings toggle commands
    [RelayCommand]
    private void ToggleStyleBasedHeadingDetection() => EnableStyleBasedHeadingDetection = !EnableStyleBasedHeadingDetection;

    [RelayCommand]
    private void ToggleInferHeadingsFromFormatting() => InferHeadingsFromFormatting = !InferHeadingsFromFormatting;

    [RelayCommand]
    private void ToggleConvertUnderlineToEmphasis() => ConvertUnderlineToEmphasis = !ConvertUnderlineToEmphasis;

    [RelayCommand]
    private void ToggleGenerateDiagnosticReport() => GenerateDiagnosticReport = !GenerateDiagnosticReport;

    // Reconvert when settings properties change
    partial void OnEnableStyleBasedHeadingDetectionChanged(bool value) => ApplySettingsAndReconvert();
    partial void OnInferHeadingsFromFormattingChanged(bool value) => ApplySettingsAndReconvert();
    partial void OnConvertUnderlineToEmphasisChanged(bool value) => ApplySettingsAndReconvert();
    partial void OnGenerateDiagnosticReportChanged(bool value) => ApplySettingsAndReconvert();

    private void ApplySettingsAndReconvert()
    {
        // Update the Settings object
        Settings.EnableStyleBasedHeadingDetection = EnableStyleBasedHeadingDetection;
        Settings.InferHeadingsFromFormatting = InferHeadingsFromFormatting;
        Settings.ConvertUnderlineToEmphasis = ConvertUnderlineToEmphasis;
        Settings.GenerateDiagnosticReport = GenerateDiagnosticReport;

        // Reconvert if document is loaded
        if (_document != null)
        {
            var newConverter = new Docx2MdConverter(Settings);
            newConverter.ConvertDocument(_document);
            UpdateMarkdownOutput();
            StatusText = "Document reconverted with new settings.";
        }
    }

    partial void OnSettingsChanged(ConversionSettings value)
    {
        // Sync individual properties when entire Settings object changes
        EnableStyleBasedHeadingDetection = value.EnableStyleBasedHeadingDetection;
        InferHeadingsFromFormatting = value.InferHeadingsFromFormatting;
        ConvertUnderlineToEmphasis = value.ConvertUnderlineToEmphasis;
        GenerateDiagnosticReport = value.GenerateDiagnosticReport;
    }

    partial void OnSelectedSegmentChanged(SegmentViewModel? value)
    {
        // Selection changed - in a full implementation, this would
        // highlight the selected segment in both DOCX and Markdown previews
    }

    private void OnSegmentOverrideChanged(object? sender, EventArgs e)
    {
        // Regenerate markdown when any segment override changes
        UpdateMarkdownOutput();
    }

    [RelayCommand]
    private async Task OpenFileAsync()
    {
        try
        {
            if (ShowOpenFileDialog == null)
            {
                StatusText = "File dialog not available. Load a sample document instead.";
                LoadSampleDocument();
                return;
            }

            var filePath = await ShowOpenFileDialog();
            if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
            {
                LoadDocument(filePath);
            }
        }
        catch (Exception ex)
        {
            StatusText = $"Error: {ex.Message}";
        }
    }

    [RelayCommand]
    private async Task ExportMarkdownAsync()
    {
        if (!HasDocument || _document == null)
        {
            StatusText = "No document loaded to export.";
            return;
        }

        try
        {
            if (ShowSaveFileDialog == null)
            {
                StatusText = "Save dialog not available.";
                return;
            }

            var filePath = await ShowSaveFileDialog();
            if (!string.IsNullOrEmpty(filePath))
            {
                StatusText = "Exporting Markdown...";
                
                // Export the document
                _converter.ExportDocument(_document, filePath);
                
                StatusText = $"Successfully exported to {Path.GetFileName(filePath)}";
            }
        }
        catch (Exception ex)
        {
            StatusText = $"Error exporting: {ex.Message}";
        }
    }

    [RelayCommand]
    private void Exit()
    {
        // Use Avalonia's proper shutdown mechanism
        if (Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime lifetime)
        {
            lifetime.Shutdown();
        }
    }

    private void LoadDocument(string filePath)
    {
        try
        {
            StatusText = $"Loading {Path.GetFileName(filePath)}...";

            // Parse and convert the document
            _document = _converter.ParseFile(filePath);
            _converter.ConvertDocument(_document);
            _currentFilePath = filePath;

            // Update UI
            DocumentTitle = Path.GetFileName(filePath);
            HasDocument = true;

            // Unsubscribe from old segments
            foreach (var vm in Segments)
            {
                vm.OverrideChanged -= OnSegmentOverrideChanged;
            }

            // Load segments wrapped in ViewModels
            Segments.Clear();
            foreach (var segment in _document.Segments)
            {
                var vm = new SegmentViewModel(segment);
                vm.OverrideChanged += OnSegmentOverrideChanged;
                Segments.Add(vm);
            }

            // Generate markdown output
            UpdateMarkdownOutput();

            // Update DOCX preview
            UpdateDocxPreview();

            var diagnosticCount = _document.GetAllDiagnostics().Count();
            StatusText = $"Loaded {Segments.Count} segments. {diagnosticCount} diagnostics.";
        }
        catch (Exception ex)
        {
            StatusText = $"Error loading document: {ex.Message}";
            HasDocument = false;
        }
    }

    private void UpdateMarkdownOutput()
    {
        if (_document == null)
        {
            MarkdownOutput = string.Empty;
            return;
        }

        // Concatenate all segment markdown outputs (using SegmentViewModel properties)
        var markdown = string.Join("\n\n",
            Segments
                .Where(s => !s.ExcludeFromOutput)
                .Select(s => s.EffectiveMarkdown));

        MarkdownOutput = markdown;
    }

    private void UpdateDocxPreview()
    {
        if (_document == null)
        {
            DocxPreviewText = "Open a DOCX file to preview its content here.";
            return;
        }

        // Simple text representation of the document
        var preview = string.Join("\n\n",
            _document.Segments.Select(s => 
                $"[{s.Type}] {s.Content}"));

        DocxPreviewText = preview;
    }

    // Method to load a sample document for testing
    public void LoadSampleDocument()
    {
        // Unsubscribe from old segments
        foreach (var vm in Segments)
        {
            vm.OverrideChanged -= OnSegmentOverrideChanged;
        }

        // Create a sample document for demonstration
        _document = new DocumentModel();

        var segment1 = new Segment
        {
            OrderIndex = 0,
            Type = SegmentType.Heading,
            Content = "Sample Document",
            MarkdownOutput = "# Sample Document",
            Metadata = new SourceMetadata { StyleName = "Heading 1", OutlineLevel = 1 }
        };

        var segment2 = new Segment
        {
            OrderIndex = 1,
            Type = SegmentType.Paragraph,
            Content = "This is a sample paragraph demonstrating the DOCX to Markdown conversion workbench.",
            MarkdownOutput = "This is a sample paragraph demonstrating the DOCX to Markdown conversion workbench.",
            Metadata = new SourceMetadata { StyleName = "Normal" }
        };

        var segment3 = new Segment
        {
            OrderIndex = 2,
            Type = SegmentType.Heading,
            Content = "Features",
            MarkdownOutput = "## Features",
            Metadata = new SourceMetadata { StyleName = "Heading 2", OutlineLevel = 2 }
        };

        var segment4 = new Segment
        {
            OrderIndex = 3,
            Type = SegmentType.ListItem,
            Content = "Three-pane workbench layout",
            MarkdownOutput = "- Three-pane workbench layout",
            Metadata = new SourceMetadata { StyleName = "List Paragraph" }
        };

        segment4.AddDiagnostic(DiagnosticLevel.Info, "LIST_DETECTED", "List item detected and converted to Markdown");

        _document.Segments.Add(segment1);
        _document.Segments.Add(segment2);
        _document.Segments.Add(segment3);
        _document.Segments.Add(segment4);

        DocumentTitle = "Sample Document";
        HasDocument = true;

        Segments.Clear();
        foreach (var segment in _document.Segments)
        {
            var vm = new SegmentViewModel(segment);
            vm.OverrideChanged += OnSegmentOverrideChanged;
            Segments.Add(vm);
        }

        UpdateMarkdownOutput();
        UpdateDocxPreview();

        StatusText = "Sample document loaded for demonstration.";
    }
}
