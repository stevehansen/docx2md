using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Data.Core;
using Avalonia.Data.Core.Plugins;
using System.IO;
using System.Linq;
using Avalonia.Markup.Xaml;
using Docx2Md.UI.ViewModels;
using Docx2Md.UI.Views;

namespace Docx2Md.UI;

public partial class App : Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            // Avoid duplicate validations from both Avalonia and the CommunityToolkit.
            // More info: https://docs.avaloniaui.net/docs/guides/development-guides/data-validation#manage-validationplugins
            DisableAvaloniaDataAnnotationValidation();
            var viewModel = new MainWindowViewModel();

            // Check for command-line file argument
            var fileArg = desktop.Args?.FirstOrDefault(arg =>
                !arg.StartsWith("-") &&
                (arg.EndsWith(".docx", System.StringComparison.OrdinalIgnoreCase) ||
                 arg.EndsWith(".docx2md", System.StringComparison.OrdinalIgnoreCase)));

            if (!string.IsNullOrEmpty(fileArg) && File.Exists(fileArg))
            {
                // Use dispatcher to load after window is ready
                Avalonia.Threading.Dispatcher.UIThread.Post(async () =>
                {
                    if (fileArg.EndsWith(".docx2md", System.StringComparison.OrdinalIgnoreCase))
                    {
                        await viewModel.OpenProjectFromPathAsync(fileArg);
                    }
                    else
                    {
                        viewModel.LoadDocument(fileArg);
                    }
                });
            }

            desktop.MainWindow = new MainWindow
            {
                DataContext = viewModel,
            };
        }

        base.OnFrameworkInitializationCompleted();
    }

    private void DisableAvaloniaDataAnnotationValidation()
    {
        // Get an array of plugins to remove
        var dataValidationPluginsToRemove =
            BindingPlugins.DataValidators.OfType<DataAnnotationsValidationPlugin>().ToArray();

        // remove each entry found
        foreach (var plugin in dataValidationPluginsToRemove)
        {
            BindingPlugins.DataValidators.Remove(plugin);
        }
    }
}