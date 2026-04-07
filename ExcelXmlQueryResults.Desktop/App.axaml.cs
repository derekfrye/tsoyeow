using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using ExcelXmlQueryResults.App.Services;
using ExcelXmlQueryResults.Desktop.ViewModels;
using ExcelXmlQueryResults.Desktop.Views;

namespace ExcelXmlQueryResults.Desktop;

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
            desktop.MainWindow = new MainWindow
            {
                DataContext = new MainWindowViewModel(new AppSettingsStore(), new ExportWorkflowService()),
            };
        }

        base.OnFrameworkInitializationCompleted();
    }
}
