using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Platform.Storage;
using Avalonia.VisualTree;
using Docx2Md.UI.ViewModels;
using SukiUI.Controls;

namespace Docx2Md.UI.Views;

public partial class MainWindow : SukiWindow
{
    private MainWindowViewModel? _viewModel;

    public MainWindow()
    {
        InitializeComponent();

        // Wire up file dialog support
        if (DataContext is MainWindowViewModel viewModel)
        {
            SetupViewModel(viewModel);
        }
    }

    protected override void OnDataContextChanged(EventArgs e)
    {
        base.OnDataContextChanged(e);

        // Unsubscribe from old view model
        if (_viewModel != null)
        {
            _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
        }

        if (DataContext is MainWindowViewModel viewModel)
        {
            SetupViewModel(viewModel);
        }
    }

    private void SetupViewModel(MainWindowViewModel viewModel)
    {
        _viewModel = viewModel;
        viewModel.ShowOpenFileDialog = ShowOpenFileDialogAsync;
        viewModel.ShowSaveFileDialog = ShowSaveFileDialogAsync;
        viewModel.ShowOpenProjectDialog = ShowOpenProjectDialogAsync;
        viewModel.ShowSaveProjectDialog = (suggestedFileName) => ShowSaveProjectDialogAsync(suggestedFileName);
        viewModel.PropertyChanged += OnViewModelPropertyChanged;
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(MainWindowViewModel.SelectedSegment))
        {
            ScrollToSelectedSegment();
        }
    }

    private void ScrollToSelectedSegment()
    {
        if (_viewModel?.SelectedSegment == null)
            return;

        var selectedIndex = _viewModel.Segments.IndexOf(_viewModel.SelectedSegment);
        if (selectedIndex < 0)
            return;

        // Scroll DOCX preview
        ScrollItemsControlToIndex(DocxPreviewItemsControl, selectedIndex);

        // Scroll Markdown preview (use the visible one)
        if (_viewModel.ShowRawMarkdown)
        {
            ScrollItemsControlToIndex(RawMarkdownItemsControl, selectedIndex);
        }
        else
        {
            ScrollItemsControlToIndex(RenderedMarkdownItemsControl, selectedIndex);
        }
    }

    private void ScrollItemsControlToIndex(ItemsControl? itemsControl, int index)
    {
        if (itemsControl == null)
            return;

        // Get the container at the index
        var container = itemsControl.ContainerFromIndex(index);
        if (container is Control control)
        {
            control.BringIntoView();
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

    private async Task<string?> ShowSaveProjectDialogAsync(string? suggestedFileName = null)
    {
        var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Save Project",
            DefaultExtension = "docx2md",
            SuggestedFileName = suggestedFileName ?? "project.docx2md",
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