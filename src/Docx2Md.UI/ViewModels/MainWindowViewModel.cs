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
using Docx2Md.Core.Services;
using Docx2Md.UI.Models;
using Docx2Md.UI.Services;

namespace Docx2Md.UI.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    private readonly Docx2MdConverter _converter;
    private readonly SettingsService _settingsService;
    private readonly ProjectFileService _projectFileService;
    private readonly AppSettings _appSettings;
    private DocumentModel? _document;
    private string _currentFilePath = string.Empty;
    private string _currentProjectPath = string.Empty;

    // Delegate for file dialogs (injected by View)
    public Func<Task<string?>>? ShowOpenFileDialog { get; set; }
    public Func<string?, Task<string?>>? ShowSaveFileDialog { get; set; }
    public Func<Task<string?>>? ShowOpenProjectDialog { get; set; }
    public Func<string?, Task<string?>>? ShowSaveProjectDialog { get; set; }
    public Func<string, Task>? CopyToClipboard { get; set; }

    // Services for configuration
    private readonly StyleMappingService _styleMappingService = new();

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
    private ObservableCollection<SegmentViewModel> _filteredSegments = new();

    [ObservableProperty]
    private SegmentViewModel? _selectedSegment;

    // Collection for multi-select in DataGrid
    private ObservableCollection<SegmentViewModel> _selectedSegments = new();
    public ObservableCollection<SegmentViewModel> SelectedSegments => _selectedSegments;

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

    [ObservableProperty]
    private ObservableCollection<RecentFileItem> _recentFiles = new();

    [ObservableProperty]
    private bool _isDirty = false;

    [ObservableProperty]
    private string _windowTitle = "DOCX → Markdown Translation Workbench";

    // Filter properties
    [ObservableProperty]
    private string _searchText = string.Empty;

    [ObservableProperty]
    private SegmentType? _filterType = null;

    [ObservableProperty]
    private bool _filterHasDiagnostics = false;

    [ObservableProperty]
    private bool _filterHasOverrides = false;

    // Statistics panel
    [ObservableProperty]
    private bool _showStatistics = false;

    // Diff view
    [ObservableProperty]
    private bool _showDiffView = false;

    // Document statistics (computed)
    public int TotalSegments => Segments.Count;
    public int WordCount => Segments.Sum(s => CountWords(s.Content));
    public int CharacterCount => Segments.Sum(s => s.Content?.Length ?? 0);
    public int HeadingCount => Segments.Count(s => s.EffectiveType == SegmentType.Heading);
    public int ParagraphCount => Segments.Count(s => s.EffectiveType == SegmentType.Paragraph);
    public int ListItemCount => Segments.Count(s => s.EffectiveType == SegmentType.ListItem);
    public int TableCount => Segments.Count(s => s.EffectiveType == SegmentType.Table);
    public int ImageCount => Segments.Count(s => s.EffectiveType == SegmentType.Image);
    public int TotalDiagnosticCount => Segments.Sum(s => s.DiagnosticCount);
    public int OverrideCount => Segments.Count(s => s.HasOverrides);
    public int ExcludedCount => Segments.Count(s => s.ExcludeFromOutput);

    private static int CountWords(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return 0;
        return text.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Length;
    }

    private const int MaxRecentFiles = 10;

    // Undo/redo support
    private record OverrideAction(string SegmentId, string PropertyName, object? OldValue, object? NewValue);
    private readonly Stack<OverrideAction> _undoStack = new();
    private readonly Stack<OverrideAction> _redoStack = new();
    private bool _isUndoingOrRedoing = false;

    public MainWindowViewModel()
    {
        // Load settings from disk
        _settingsService = new SettingsService();
        _projectFileService = new ProjectFileService();
        _appSettings = _settingsService.Load();

        // Apply loaded settings to observable properties
        _showDocxPreview = _appSettings.ShowDocxPreview;
        _showSegmentInspector = _appSettings.ShowSegmentInspector;
        _showMarkdownPreview = _appSettings.ShowMarkdownPreview;
        _showRawMarkdown = _appSettings.ShowRawMarkdown;
        _enableStyleBasedHeadingDetection = _appSettings.EnableStyleBasedHeadingDetection;
        _inferHeadingsFromFormatting = _appSettings.InferHeadingsFromFormatting;
        _convertUnderlineToEmphasis = _appSettings.ConvertUnderlineToEmphasis;
        _generateDiagnosticReport = _appSettings.GenerateDiagnosticReport;

        // Sync settings to ConversionSettings object
        Settings.EnableStyleBasedHeadingDetection = _enableStyleBasedHeadingDetection;
        Settings.InferHeadingsFromFormatting = _inferHeadingsFromFormatting;
        Settings.ConvertUnderlineToEmphasis = _convertUnderlineToEmphasis;
        Settings.GenerateDiagnosticReport = _generateDiagnosticReport;

        _converter = new Docx2MdConverter(Settings);

        // Load recent files
        LoadRecentFiles();

        // Set initial window title with version
        UpdateWindowTitle();
    }

    private void LoadRecentFiles()
    {
        RecentFiles.Clear();
        foreach (var path in _appSettings.RecentFiles)
        {
            RecentFiles.Add(new RecentFileItem { FilePath = path });
        }
    }

    private void AddToRecentFiles(string filePath)
    {
        // Remove if already exists (will be re-added at top)
        _appSettings.RecentFiles.Remove(filePath);

        // Add to front
        _appSettings.RecentFiles.Insert(0, filePath);

        // Trim to max
        while (_appSettings.RecentFiles.Count > MaxRecentFiles)
        {
            _appSettings.RecentFiles.RemoveAt(_appSettings.RecentFiles.Count - 1);
        }

        _settingsService.Save(_appSettings);
        LoadRecentFiles();
    }

    [RelayCommand]
    private async Task OpenRecentFileAsync(RecentFileItem? item)
    {
        if (item == null) return;

        if (!File.Exists(item.FilePath))
        {
            // Remove from list and show error
            _appSettings.RecentFiles.Remove(item.FilePath);
            _settingsService.Save(_appSettings);
            LoadRecentFiles();
            StatusText = $"File not found: {item.FileName}";
            return;
        }

        // Check if it's a project file or a DOCX file
        if (item.FilePath.EndsWith(".docx2md", StringComparison.OrdinalIgnoreCase))
        {
            await OpenProjectFromPathAsync(item.FilePath);
        }
        else
        {
            LoadDocument(item.FilePath);
        }
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

    // Save settings when view toggle properties change
    partial void OnShowDocxPreviewChanged(bool value) => SaveViewSettings();
    partial void OnShowSegmentInspectorChanged(bool value) => SaveViewSettings();
    partial void OnShowMarkdownPreviewChanged(bool value) => SaveViewSettings();
    partial void OnShowRawMarkdownChanged(bool value) => SaveViewSettings();

    private void SaveViewSettings()
    {
        _appSettings.ShowDocxPreview = ShowDocxPreview;
        _appSettings.ShowSegmentInspector = ShowSegmentInspector;
        _appSettings.ShowMarkdownPreview = ShowMarkdownPreview;
        _appSettings.ShowRawMarkdown = ShowRawMarkdown;
        _settingsService.Save(_appSettings);
    }

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

        // Save to disk
        _appSettings.EnableStyleBasedHeadingDetection = EnableStyleBasedHeadingDetection;
        _appSettings.InferHeadingsFromFormatting = InferHeadingsFromFormatting;
        _appSettings.ConvertUnderlineToEmphasis = ConvertUnderlineToEmphasis;
        _appSettings.GenerateDiagnosticReport = GenerateDiagnosticReport;
        _settingsService.Save(_appSettings);

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

    partial void OnSelectedSegmentChanged(SegmentViewModel? oldValue, SegmentViewModel? newValue)
    {
        // Update IsSelected state for visual highlighting in preview panes
        if (oldValue != null)
            oldValue.IsSelected = false;
        if (newValue != null)
            newValue.IsSelected = true;
    }

    partial void OnIsDirtyChanged(bool value) => UpdateWindowTitle();

    partial void OnSearchTextChanged(string value) => UpdateFilteredSegments();
    partial void OnFilterTypeChanged(SegmentType? value) => UpdateFilteredSegments();
    partial void OnFilterHasDiagnosticsChanged(bool value) => UpdateFilteredSegments();
    partial void OnFilterHasOverridesChanged(bool value) => UpdateFilteredSegments();

    private void UpdateFilteredSegments()
    {
        var filtered = Segments.AsEnumerable();

        // Apply search text filter
        if (!string.IsNullOrWhiteSpace(SearchText))
        {
            var searchLower = SearchText.ToLowerInvariant();
            filtered = filtered.Where(s =>
                (s.Content?.ToLowerInvariant().Contains(searchLower) ?? false) ||
                (s.EffectiveMarkdown?.ToLowerInvariant().Contains(searchLower) ?? false) ||
                (s.Metadata?.StyleName?.ToLowerInvariant().Contains(searchLower) ?? false));
        }

        // Apply type filter
        if (FilterType.HasValue)
        {
            filtered = filtered.Where(s => s.EffectiveType == FilterType.Value);
        }

        // Apply diagnostics filter
        if (FilterHasDiagnostics)
        {
            filtered = filtered.Where(s => s.DiagnosticCount > 0);
        }

        // Apply overrides filter
        if (FilterHasOverrides)
        {
            filtered = filtered.Where(s => s.HasOverrides);
        }

        FilteredSegments.Clear();
        foreach (var segment in filtered)
        {
            FilteredSegments.Add(segment);
        }
    }

    [RelayCommand]
    private void ClearFilters()
    {
        SearchText = string.Empty;
        FilterType = null;
        FilterHasDiagnostics = false;
        FilterHasOverrides = false;
    }

    private void UpdateWindowTitle()
    {
        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        var versionStr = version != null ? $" v{version.Major}.{version.Minor}.{version.Build}" : "";
        var title = $"DOCX → Markdown Workbench{versionStr}";
        if (HasDocument)
        {
            var docName = Path.GetFileName(_currentFilePath);
            title = $"{docName} - {title}";
        }
        if (IsDirty)
        {
            title = "* " + title;
        }
        WindowTitle = title;
    }

    [RelayCommand]
    private void SelectSegment(SegmentViewModel? segment)
    {
        if (segment != null)
            SelectedSegment = segment;
    }

    [RelayCommand]
    private void SelectPreviousSegment()
    {
        if (Segments.Count == 0) return;

        if (SelectedSegment == null)
        {
            SelectedSegment = Segments[^1]; // Select last
        }
        else
        {
            var index = Segments.IndexOf(SelectedSegment);
            if (index > 0)
                SelectedSegment = Segments[index - 1];
        }
    }

    [RelayCommand]
    private void SelectNextSegment()
    {
        if (Segments.Count == 0) return;

        if (SelectedSegment == null)
        {
            SelectedSegment = Segments[0]; // Select first
        }
        else
        {
            var index = Segments.IndexOf(SelectedSegment);
            if (index < Segments.Count - 1)
                SelectedSegment = Segments[index + 1];
        }
    }

    [RelayCommand]
    private void SelectFirstSegment()
    {
        if (Segments.Count > 0)
            SelectedSegment = Segments[0];
    }

    [RelayCommand]
    private void SelectLastSegment()
    {
        if (Segments.Count > 0)
            SelectedSegment = Segments[^1];
    }

    [RelayCommand]
    private void ToggleExcludeSelected()
    {
        if (SelectedSegment != null)
            SelectedSegment.ExcludeFromOutput = !SelectedSegment.ExcludeFromOutput;
    }

    [RelayCommand]
    private void SetHeadingLevel(int? level)
    {
        if (SelectedSegment != null)
            SelectedSegment.OverrideHeadingLevel = level;
    }

    [RelayCommand]
    private void ClearSelectedOverrides()
    {
        if (SelectedSegment != null)
        {
            SelectedSegment.ExcludeFromOutput = false;
            SelectedSegment.OverrideHeadingLevel = null;
            SelectedSegment.OverrideType = null;
            SelectedSegment.ManualMarkdownOverride = null;
        }
    }

    // Bulk operations for multi-select
    private IEnumerable<SegmentViewModel> GetTargetSegments()
    {
        if (SelectedSegments.Count > 0)
            return SelectedSegments;
        if (SelectedSegment != null)
            return new[] { SelectedSegment };
        return Array.Empty<SegmentViewModel>();
    }

    [RelayCommand]
    private void BulkToggleExclude()
    {
        foreach (var segment in GetTargetSegments())
        {
            segment.ExcludeFromOutput = !segment.ExcludeFromOutput;
        }
    }

    [RelayCommand]
    private void BulkSetHeadingLevel(int? level)
    {
        foreach (var segment in GetTargetSegments())
        {
            segment.OverrideHeadingLevel = level;
        }
    }

    [RelayCommand]
    private void BulkSetType(SegmentType? type)
    {
        foreach (var segment in GetTargetSegments())
        {
            segment.OverrideType = type;
        }
    }

    [RelayCommand]
    private void BulkClearOverrides()
    {
        foreach (var segment in GetTargetSegments())
        {
            segment.ExcludeFromOutput = false;
            segment.OverrideHeadingLevel = null;
            segment.OverrideType = null;
            segment.ManualMarkdownOverride = null;
        }
    }

    [RelayCommand]
    private void BulkExclude()
    {
        foreach (var segment in GetTargetSegments())
        {
            segment.ExcludeFromOutput = true;
        }
    }

    [RelayCommand]
    private void BulkInclude()
    {
        foreach (var segment in GetTargetSegments())
        {
            segment.ExcludeFromOutput = false;
        }
    }

    private void OnSegmentOverrideChanged(object? sender, EventArgs e)
    {
        // Regenerate markdown when any segment override changes
        UpdateMarkdownOutput();
        // Mark as dirty when overrides change
        IsDirty = true;
        // Refresh statistics
        RefreshStatistics();
    }

    [RelayCommand]
    private void ToggleStatistics() => ShowStatistics = !ShowStatistics;

    [RelayCommand]
    private void ToggleDiffView() => ShowDiffView = !ShowDiffView;

    [RelayCommand]
    private void OpenStyleMappingsConfig()
    {
        var filePath = _styleMappingService.GetFilePath();
        try
        {
            // Ensure file exists
            if (!File.Exists(filePath))
            {
                _styleMappingService.Load(); // Creates default file
            }

            // Open in default editor
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            };
            System.Diagnostics.Process.Start(psi);
            StatusText = $"Opened style mappings config: {filePath}";
        }
        catch (Exception ex)
        {
            StatusText = $"Error opening config: {ex.Message}";
        }
    }

    private void RefreshStatistics()
    {
        OnPropertyChanged(nameof(TotalSegments));
        OnPropertyChanged(nameof(WordCount));
        OnPropertyChanged(nameof(CharacterCount));
        OnPropertyChanged(nameof(HeadingCount));
        OnPropertyChanged(nameof(ParagraphCount));
        OnPropertyChanged(nameof(ListItemCount));
        OnPropertyChanged(nameof(TableCount));
        OnPropertyChanged(nameof(ImageCount));
        OnPropertyChanged(nameof(TotalDiagnosticCount));
        OnPropertyChanged(nameof(OverrideCount));
        OnPropertyChanged(nameof(ExcludedCount));
    }

    // Undo/Redo functionality
    public bool CanUndo => _undoStack.Count > 0;
    public bool CanRedo => _redoStack.Count > 0;

    public void RecordOverrideChange(string segmentId, string propertyName, object? oldValue, object? newValue)
    {
        if (_isUndoingOrRedoing) return;

        _undoStack.Push(new OverrideAction(segmentId, propertyName, oldValue, newValue));
        _redoStack.Clear();
        OnPropertyChanged(nameof(CanUndo));
        OnPropertyChanged(nameof(CanRedo));
    }

    [RelayCommand]
    private void Undo()
    {
        if (_undoStack.Count == 0) return;

        _isUndoingOrRedoing = true;
        try
        {
            var action = _undoStack.Pop();
            _redoStack.Push(action);

            // Find segment and apply old value
            var segment = Segments.FirstOrDefault(s => s.Id == action.SegmentId);
            if (segment != null)
            {
                ApplyOverrideValue(segment, action.PropertyName, action.OldValue);
            }

            OnPropertyChanged(nameof(CanUndo));
            OnPropertyChanged(nameof(CanRedo));
        }
        finally
        {
            _isUndoingOrRedoing = false;
        }
    }

    [RelayCommand]
    private void Redo()
    {
        if (_redoStack.Count == 0) return;

        _isUndoingOrRedoing = true;
        try
        {
            var action = _redoStack.Pop();
            _undoStack.Push(action);

            // Find segment and apply new value
            var segment = Segments.FirstOrDefault(s => s.Id == action.SegmentId);
            if (segment != null)
            {
                ApplyOverrideValue(segment, action.PropertyName, action.NewValue);
            }

            OnPropertyChanged(nameof(CanUndo));
            OnPropertyChanged(nameof(CanRedo));
        }
        finally
        {
            _isUndoingOrRedoing = false;
        }
    }

    private void ApplyOverrideValue(SegmentViewModel segment, string propertyName, object? value)
    {
        switch (propertyName)
        {
            case nameof(SegmentViewModel.ExcludeFromOutput):
                segment.ExcludeFromOutput = value is bool b && b;
                break;
            case nameof(SegmentViewModel.OverrideHeadingLevel):
                segment.OverrideHeadingLevel = value as int?;
                break;
            case nameof(SegmentViewModel.OverrideType):
                segment.OverrideType = value as SegmentType?;
                break;
            case nameof(SegmentViewModel.ManualMarkdownOverride):
                segment.ManualMarkdownOverride = value as string;
                break;
        }
    }

    private void ClearUndoHistory()
    {
        _undoStack.Clear();
        _redoStack.Clear();
        OnPropertyChanged(nameof(CanUndo));
        OnPropertyChanged(nameof(CanRedo));
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

            // Suggest filename based on source document
            var suggestedName = !string.IsNullOrEmpty(_currentFilePath)
                ? Path.GetFileNameWithoutExtension(_currentFilePath) + ".md"
                : null;
            var filePath = await ShowSaveFileDialog(suggestedName);
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
    private async Task CopyMarkdownToClipboardAsync()
    {
        if (!HasDocument || Segments.Count == 0)
        {
            StatusText = "No document loaded to copy.";
            return;
        }

        try
        {
            // Build markdown from non-excluded segments
            var markdown = string.Join("\n\n",
                Segments
                    .Where(s => !s.ExcludeFromOutput)
                    .Select(s => s.EffectiveMarkdown)
                    .Where(m => !string.IsNullOrWhiteSpace(m)));

            if (CopyToClipboard != null)
            {
                await CopyToClipboard(markdown);
                StatusText = "Markdown copied to clipboard.";
            }
            else
            {
                StatusText = "Clipboard not available.";
            }
        }
        catch (Exception ex)
        {
            StatusText = $"Error copying to clipboard: {ex.Message}";
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

    [RelayCommand]
    private async Task SaveProjectAsync()
    {
        if (!HasDocument || _document == null)
        {
            StatusText = "No document loaded to save project.";
            return;
        }

        try
        {
            if (ShowSaveProjectDialog == null)
            {
                StatusText = "Save project dialog not available.";
                return;
            }

            // Generate suggested filename from source DOCX
            string? suggestedFileName = null;
            if (!string.IsNullOrEmpty(_currentFilePath))
            {
                var docxName = Path.GetFileNameWithoutExtension(_currentFilePath);
                suggestedFileName = $"{docxName}.docx2md";
            }

            var filePath = await ShowSaveProjectDialog(suggestedFileName);
            if (!string.IsNullOrEmpty(filePath))
            {
                StatusText = "Saving project...";

                // Create and save project file
                var project = _projectFileService.CreateFromDocument(_document, Settings);
                _projectFileService.Save(project, filePath);

                _currentProjectPath = filePath;
                IsDirty = false;

                // Add project to recent files
                AddToRecentFiles(filePath);

                StatusText = $"Project saved to {Path.GetFileName(filePath)}";
            }
        }
        catch (Exception ex)
        {
            StatusText = $"Error saving project: {ex.Message}";
        }
    }

    [RelayCommand]
    private async Task OpenProjectAsync()
    {
        try
        {
            if (ShowOpenProjectDialog == null)
            {
                StatusText = "Open project dialog not available.";
                return;
            }

            var filePath = await ShowOpenProjectDialog();
            if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
            {
                await OpenProjectFromPathAsync(filePath);
            }
        }
        catch (Exception ex)
        {
            StatusText = $"Error loading project: {ex.Message}";
        }
    }

    public async Task OpenProjectFromPathAsync(string filePath)
    {
        try
        {
            StatusText = $"Loading project {Path.GetFileName(filePath)}...";

            var project = _projectFileService.Load(filePath);
            if (project == null)
            {
                StatusText = "Failed to load project file.";
                return;
            }

            // Load the source DOCX if available
            if (!string.IsNullOrEmpty(project.SourceDocxPath) && File.Exists(project.SourceDocxPath))
            {
                LoadDocument(project.SourceDocxPath);

                // Apply project overrides
                if (_document != null)
                {
                    _projectFileService.ApplyToDocument(project, _document);

                    // Refresh UI segments to reflect applied overrides
                    foreach (var vm in Segments)
                    {
                        vm.RefreshFromModel();
                    }

                    // Regenerate markdown output and refresh statistics
                    UpdateMarkdownOutput();
                    RefreshStatistics();
                }

                _currentProjectPath = filePath;
                IsDirty = false;
                UpdateWindowTitle();

                // Add project to recent files
                AddToRecentFiles(filePath);

                var overrideCount = project.SegmentOverrides.Count;
                StatusText = $"Project loaded. {overrideCount} overrides applied.";
            }
            else
            {
                StatusText = $"Source document not found: {project.SourceDocxPath}";
            }
        }
        catch (Exception ex)
        {
            StatusText = $"Error loading project: {ex.Message}";
        }
    }

    public void LoadDocument(string filePath)
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
                vm.OverrideChangedWithValues += OnSegmentOverrideChangedWithValues;
                Segments.Add(vm);
            }

            // Clear undo history when loading new document
            ClearUndoHistory();

            // Update filtered view and statistics
            UpdateFilteredSegments();
            RefreshStatistics();

            // Generate markdown output
            UpdateMarkdownOutput();

            // Update DOCX preview
            UpdateDocxPreview();

            // Add to recent files
            AddToRecentFiles(filePath);

            // Reset project state
            _currentProjectPath = string.Empty;
            IsDirty = false;
            UpdateWindowTitle();

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
        // Filter out excluded segments and empty paragraphs (common in Word for spacing)
        var markdown = string.Join("\n\n",
            Segments
                .Where(s => !s.ExcludeFromOutput)
                .Select(s => s.EffectiveMarkdown)
                .Where(md => !string.IsNullOrWhiteSpace(md)));

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
            vm.OverrideChangedWithValues += OnSegmentOverrideChangedWithValues;
            Segments.Add(vm);
        }

        ClearUndoHistory();
        UpdateFilteredSegments();
        RefreshStatistics();
        UpdateMarkdownOutput();
        UpdateDocxPreview();

        StatusText = "Sample document loaded for demonstration.";
    }

    private void OnSegmentOverrideChangedWithValues(object? sender, OverrideChangedEventArgs e)
    {
        if (sender is SegmentViewModel segment)
        {
            RecordOverrideChange(segment.Id, e.PropertyName, e.OldValue, e.NewValue);
        }
    }
}
