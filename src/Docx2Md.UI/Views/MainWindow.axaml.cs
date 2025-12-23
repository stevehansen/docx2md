using System;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Docx2Md.UI.ViewModels;

namespace Docx2Md.UI.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        
        // Wire up file dialog support
        if (DataContext is MainWindowViewModel viewModel)
        {
            viewModel.ShowOpenFileDialog = ShowOpenFileDialogAsync;
            viewModel.ShowSaveFileDialog = ShowSaveFileDialogAsync;
        }
    }

    protected override void OnDataContextChanged(EventArgs e)
    {
        base.OnDataContextChanged(e);
        
        if (DataContext is MainWindowViewModel viewModel)
        {
            viewModel.ShowOpenFileDialog = ShowOpenFileDialogAsync;
            viewModel.ShowSaveFileDialog = ShowSaveFileDialogAsync;
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
}