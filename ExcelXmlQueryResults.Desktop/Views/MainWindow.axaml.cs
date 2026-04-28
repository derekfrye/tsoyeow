using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Platform.Storage;
using ExcelXmlQueryResults.Desktop.ViewModels;

namespace ExcelXmlQueryResults.Desktop.Views;

public partial class MainWindow : Window
{
    bool initialized;

    public MainWindow()
    {
        InitializeComponent();
        Opened += OnOpened;
        Closing += OnClosing;
    }

    MainWindowViewModel ViewModel => (MainWindowViewModel)DataContext!;

    void InitializeComponent()
    {
        AvaloniaXamlLoader.Load(this);
    }

    async void OnOpened(object? sender, EventArgs e)
    {
        if (initialized)
        {
            return;
        }

        initialized = true;
        await ViewModel.InitializeAsync();
    }

    void OnClosing(object? sender, WindowClosingEventArgs e)
    {
        ViewModel.CancelExport();
        ViewModel.Save();
        ViewModel.Dispose();
    }

    async void BrowseOutputFile_Click(object? sender, RoutedEventArgs e)
    {
        var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Choose output workbook",
            SuggestedFileName = string.IsNullOrWhiteSpace(ViewModel.OutputPath)
                ? "export.xlsx"
                : Path.GetFileName(ViewModel.OutputPath),
            DefaultExtension = "xlsx",
            FileTypeChoices =
            [
                new FilePickerFileType("Excel Workbook")
                {
                    Patterns = ["*.xlsx"],
                },
            ],
        });

        if (file is not null)
        {
            ViewModel.SetOutputPath(file.Path.LocalPath);
        }
    }

    async void BrowseQueryFiles_Click(object? sender, RoutedEventArgs e)
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Select SQL files",
            AllowMultiple = true,
            FileTypeFilter =
            [
                new FilePickerFileType("SQL files")
                {
                    Patterns = ["*.sql"],
                },
            ],
        });

        ViewModel.SetQueryFiles(files.Select(file => file.Path.LocalPath));
    }

    void ClearQueryFiles_Click(object? sender, RoutedEventArgs e)
    {
        ViewModel.ClearQueryFiles();
    }

    async void RunExport_Click(object? sender, RoutedEventArgs e)
    {
        await ViewModel.RunExportAsync();
    }

    void CancelExport_Click(object? sender, RoutedEventArgs e)
    {
        ViewModel.CancelExport();
    }
}
