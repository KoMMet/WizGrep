using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.UI.Input;
using Windows.Storage;
using Microsoft.UI.Xaml.Automation;
using Microsoft.Windows.Storage.Pickers;
using NPOI.SS.Formula.Functions;
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
    private bool _isPaneSplitterDragging;
    private double _paneSplitterDragStartX;
    private double _paneSplitterDragStartWidth;

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
            var picker = new FileSavePicker(AppWindow.Id)
            {
                SuggestedStartLocation = PickerLocationId.Desktop
            };
            picker.FileTypeChoices.Add($"{ResourceLoaderHelper.GetString("TextFileLabel")}", new List<string> { ".txt" });
            
            var pickResult = await picker.PickSaveFileAsync();
                if (pickResult != null && !string.IsNullOrEmpty(pickResult.Path))
                {
                    var file = await StorageFile.GetFileFromPathAsync(pickResult.Path);
                    await FileIO.WriteTextAsync(file, content);
                }
        }
        catch (Exception ex)
        {
            LoggerHelper.Instance.LogError($"Error exporting to file: {ex.StackTrace}");
            await ShowErrorDialogAsync($"{ResourceLoaderHelper.GetString("ErrorMessage_ExportFile")}", ex.StackTrace);
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

    // --- Left/Right pane splitter ---

    private void PaneSplitter_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        if (sender is not UIElement el)
            return;

        _isPaneSplitterDragging = true;
        _paneSplitterDragStartX = e.GetCurrentPoint(RootGrid).Position.X;
        _paneSplitterDragStartWidth = LeftPaneColumn.ActualWidth;
        el.CapturePointer(e.Pointer);
        e.Handled = true;
    }

    private void PaneSplitter_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (!_isPaneSplitterDragging)
            return;

        var currentX = e.GetCurrentPoint(RootGrid).Position.X;
        var delta = currentX - _paneSplitterDragStartX;
        var newWidth = Math.Max(LeftPaneColumn.MinWidth, _paneSplitterDragStartWidth + delta);
        LeftPaneColumn.Width = new GridLength(newWidth);
        e.Handled = true;
    }

    private void PaneSplitter_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        if (!_isPaneSplitterDragging)
            return;

        _isPaneSplitterDragging = false;
        if (sender is UIElement el)
            el.ReleasePointerCapture(e.Pointer);
        e.Handled = true;
    }

    private void UIElement_OnDoubleTapped_leftPainPath(object sender, DoubleTappedRoutedEventArgs e)
    {
        if(sender is TextBlock { Text: { } path })
        {
            ViewModel.OpenFileCommand(path);
        }
    }

    private void UIElement_OnDoubleTapped_writePainPath(object sender, DoubleTappedRoutedEventArgs e)
    {
        if (sender is TextBlock textBlock)
        {
            var toolTip = ToolTipService.GetToolTip(textBlock);
            if (toolTip != null)
            {
                var path = toolTip.ToString();
                ViewModel.OpenFileCommand(path);
            }
        }

    }
}