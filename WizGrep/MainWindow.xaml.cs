using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Windows.Storage;
using Windows.Storage.Pickers;
using WinRT.Interop;
using WizGrep.Helpers;
using WizGrep.Services;
using WizGrep.ViewModels;
using WizGrep.Views;

namespace WizGrep;

/// <summary>
/// Represents the main window of the application, providing the primary user interface for configuring grep settings,
/// managing file operations, and displaying search results.
/// </summary>
/// <remarks>The MainWindow class initializes core services and the main view model, and manages user interactions
/// such as adjusting result column widths and opening settings dialogs. It serves as the entry point for user
/// interaction with grep-related features and file export functionality.</remarks>
public sealed partial class MainWindow : Window
{
    public MainWindow()
    {
        // Initialize services required for the application's core functionality.
        var settingsService = new SettingsService();
        var fileReaderService = new FileReaderService();
        var indexService = new IndexService();
        var grepService = new GrepService(fileReaderService, indexService);

        // Initialize the main view model
        ViewModel = new MainViewModel(grepService, settingsService);
        ViewModel.ExportToFileAsync = ExportToFileAsync;

        InitializeComponent();

        _columnWidthsHelper = (ResultColumnWidthsHelper)RootGrid.Resources["ResultColumnWidthsHelper"];
    }

    private readonly ResultColumnWidthsHelper _columnWidthsHelper;
    private ColumnDefinition? _draggingColumn;
    private double _dragStartX;
    private double _dragStartWidth;

    public MainViewModel ViewModel { get; }

    private async void OnGrepSettingsClick(object sender, RoutedEventArgs e)
    {
        await ShowGrepSettingsDialogAsync();
    }

    private async void OnWizGrepSettingsClick(object sender, RoutedEventArgs e)
    {
        await ShowWizGrepSettingsDialogAsync();
    }

    private async Task ShowGrepSettingsDialogAsync()
    {
        try
        {
            var dialog = new GrepSettingsDialog(this, ViewModel.GrepSettings)
            {
                XamlRoot = Content.XamlRoot
            };

            var result = await dialog.ShowAsync();

            if (result == ContentDialogResult.Primary && dialog.StartSearch)
            {
                ViewModel.SaveGrepSettings(dialog.Settings);
                await ViewModel.StartSearchAsync();
            }
        }
        catch (Exception ex)
        {
            LoggerHelper.Instance.LogError($"Error showing Grep Settings dialog: {ex}");
            await ShowErrorDialogAsync($"{ResourceLoaderHelper.GetString("ErrorMessage_GrepSetting")}", ex.Message);
        }
    }

    private async Task ShowWizGrepSettingsDialogAsync()
    {
        try
        {
            var dialog = new WizGrepSettingsDialog(this, ViewModel.WizGrepSettings)
            {
                XamlRoot = Content.XamlRoot
            };

            var result = await dialog.ShowAsync();

            if (result == ContentDialogResult.Primary)
            {
                ViewModel.SaveWizGrepSettings(dialog.Settings);
            }
        }
        catch (Exception ex)
        {
            LoggerHelper.Instance.LogError($"Error showing WizGrep Settings dialog: {ex}");
            await ShowErrorDialogAsync($"{ResourceLoaderHelper.GetString("ErrorMessage_WizGrepSetting")}", ex.Message);
        }
    }

    private async Task ExportToFileAsync(string content, string suggestedFileName)
    {
        try
        {
            var picker = new FileSavePicker
            {
                SuggestedStartLocation = PickerLocationId.DocumentsLibrary,
                SuggestedFileName = suggestedFileName
            };
            picker.FileTypeChoices.Add($"{ResourceLoaderHelper.GetString("TextFileLabel")}", new List<string> { ".txt" });

            var hwnd = WindowNative.GetWindowHandle(this);
            InitializeWithWindow.Initialize(picker, hwnd);

            var file = await picker.PickSaveFileAsync();
            if (file != null)
            {
                await FileIO.WriteTextAsync(file, content);
            }
        }
        catch (Exception ex)
        {
            LoggerHelper.Instance.LogError($"Error exporting to file: {ex}");
            await ShowErrorDialogAsync($"{ResourceLoaderHelper.GetString("ErrorMessage_ExportFile")}", ex.Message);
        }
    }

    private async Task ShowErrorDialogAsync(string title, string message)
    {
        var dialog = new ContentDialog
        {
            Title = $"{ResourceLoaderHelper.GetString("ErrorLabel")}",
            Content = $"{title}\n{message}",
            CloseButtonText = "OK",
            XamlRoot = Content.XamlRoot
        };
        await dialog.ShowAsync();
    }

    private void OnSplitterPointerPressed(object sender, PointerRoutedEventArgs e)
    {
        if (sender is not ColumnSplitterHelper splitter)
            return;

        var tag = splitter.Tag?.ToString();
        var grid = (Grid)splitter.Parent;

        _draggingColumn = tag switch
        {
            "0" => grid.ColumnDefinitions[0],
            "1" => grid.ColumnDefinitions[2],
            _ => null
        };

        if (_draggingColumn is null)
            return;

        _dragStartX = e.GetCurrentPoint(grid).Position.X;
        _dragStartWidth = _draggingColumn.ActualWidth;
        splitter.CapturePointer(e.Pointer);
        e.Handled = true;
    }

    private void OnSplitterPointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (_draggingColumn is null || sender is not ColumnSplitterHelper splitter)
            return;

        var grid = (Grid)splitter.Parent;
        var currentX = e.GetCurrentPoint(grid).Position.X;
        var delta = currentX - _dragStartX;
        var newWidth = Math.Max(_draggingColumn.MinWidth, _dragStartWidth + delta);
        _draggingColumn.Width = new GridLength(newWidth);

        var tag = splitter.Tag?.ToString();
        if (tag == "0")
            _columnWidthsHelper.FileNameWidth = newWidth;
        else if (tag == "1")
            _columnWidthsHelper.LocationWidth = newWidth;

        e.Handled = true;
    }

    private void OnSplitterPointerReleased(object sender, PointerRoutedEventArgs e)
    {
        if (_draggingColumn is null || sender is not ColumnSplitterHelper splitter)
            return;

        splitter.ReleasePointerCapture(e.Pointer);
        _draggingColumn = null;
        e.Handled = true;
    }
}