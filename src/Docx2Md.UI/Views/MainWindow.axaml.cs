using System;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Platform.Storage;
using Docx2Md.UI.ViewModels;
using SukiUI.Controls;

namespace Docx2Md.UI.Views;

public partial class MainWindow : SukiWindow
{
    public MainWindow()
    {
        InitializeComponent();
        
        // Wire up file dialog support
        if (DataContext is MainWindowViewModel viewModel)
        {
            viewModel.ShowOpenFileDialog = ShowOpenFileDialogAsync;
            viewModel.ShowSaveFileDialog = ShowSaveFileDialogAsync;
            viewModel.ShowOpenProjectDialog = ShowOpenProjectDialogAsync;
            viewModel.ShowSaveProjectDialog = ShowSaveProjectDialogAsync;
        }
    }

    protected override void OnDataContextChanged(EventArgs e)
    {
        base.OnDataContextChanged(e);

        if (DataContext is MainWindowViewModel viewModel)
        {
            viewModel.ShowOpenFileDialog = ShowOpenFileDialogAsync;
            viewModel.ShowSaveFileDialog = ShowSaveFileDialogAsync;
            viewModel.ShowOpenProjectDialog = ShowOpenProjectDialogAsync;
            viewModel.ShowSaveProjectDialog = ShowSaveProjectDialogAsync;
        }
    }

    private async Task<string?> ShowOpenFileDialogAsync()
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Open DOCX File",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Word Documents")
                {
                    Patterns = new[] { "*.docx" }
                }
            }
        });

        return files.Count > 0 ? files[0].Path.LocalPath : null;
    }

    private async Task<string?> ShowSaveFileDialogAsync()
    {
        var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Export Markdown",
            DefaultExtension = "md",
            SuggestedFileName = "output.md",
            FileTypeChoices = new[]
            {
                new FilePickerFileType("Markdown Files")
                {
                    Patterns = new[] { "*.md" }
                }
            }
        });

        return file?.Path.LocalPath;
    }

    private async Task<string?> ShowOpenProjectDialogAsync()
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Open Project File",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("DOCX2MD Project Files")
                {
                    Patterns = new[] { "*.docx2md" }
                }
            }
        });

        return files.Count > 0 ? files[0].Path.LocalPath : null;
    }

    private async Task<string?> ShowSaveProjectDialogAsync()
    {
        var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Save Project",
            DefaultExtension = "docx2md",
            SuggestedFileName = "project.docx2md",
            FileTypeChoices = new[]
            {
                new FilePickerFileType("DOCX2MD Project Files")
                {
                    Patterns = new[] { "*.docx2md" }
                }
            }
        });

        return file?.Path.LocalPath;
    }

    private void OnDocxSegmentClick(object? sender, PointerPressedEventArgs e)
    {
        SelectSegmentFromBorder(sender);
    }

    private void OnMarkdownSegmentClick(object? sender, PointerPressedEventArgs e)
    {
        SelectSegmentFromBorder(sender);
    }

    private void SelectSegmentFromBorder(object? sender)
    {
        if (sender is Border border && border.DataContext is SegmentViewModel segment)
        {
            if (DataContext is MainWindowViewModel vm)
            {
                vm.SelectedSegment = segment;
            }
        }
    }
}