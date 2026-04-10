using System;
using System.Threading.Tasks;
using Microsoft.Windows.Storage.Pickers;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using WinRT.Interop;
using WizGrep.Helpers;
using WizGrep.Models;
using WizGrep.ViewModels;

namespace WizGrep.Views;

/// <summary>
/// Represents a dialog for configuring grep settings in the application.
/// </summary>
/// <remarks>
/// This dialog allows users to modify grep settings and optionally start a search operation.
/// It is initialized with an optional parent window and existing settings.
/// </remarks>
public sealed partial class GrepSettingsDialog : ContentDialog
{
    private readonly Window _parentWindow;
    private readonly GrepSettings _settings;

    /// <summary>
    /// Initializes a new instance of the GrepSettingsDialog class, allowing the user to configure grep settings in a
    /// modal dialog.
    /// </summary>
    /// <remarks>
    /// This constructor sets up the dialog's view model and loads settings from the provided
    /// GrepSettings instance. If no existing settings are supplied, a new instance with default values is created. The
    /// dialog is intended to be used as a modal child of the specified parent window.
    /// </remarks>
    /// <param name="parentWindow">The parent window that will own the dialog. This is used to display the dialog modally and to position it
    /// relative to the parent.</param>
    /// <param name="existingSettings">An optional GrepSettings instance containing existing configuration values to load into the dialog. If null,
    /// default settings are used.</param>
    public GrepSettingsDialog(Window parentWindow, GrepSettings? existingSettings = null)
    {
        _parentWindow = parentWindow;
        _settings = existingSettings ?? new GrepSettings();
        ViewModel = new GrepSettingsDialogViewModel();
        ViewModel.LoadFromSettings(_settings);
        ViewModel.BrowseFolderAsync = PickFolderAsync;

        InitializeComponent();
    }

    public GrepSettingsDialogViewModel ViewModel { get; }
    public GrepSettings Settings => _settings;
    public bool StartSearch { get; private set; }

    /// <summary>
    /// Prompts the user to select a folder and returns the path of the selected folder.
    /// </summary>
    /// <remarks>
    /// The folder picker is initialized with the desktop as the suggested start location and allows
    /// selection of any folder type. If an exception occurs during the folder selection process, null is returned.
    /// </remarks>
    /// <returns>
    /// The path of the selected folder as a string. Returns null if the user cancels the operation or if an error occurs.
    ///</returns>
    private async Task<string?> PickFolderAsync()
    {
        try
        {
            var picker = new FolderPicker(_parentWindow.AppWindow.Id);
            picker.SuggestedStartLocation = PickerLocationId.Desktop;

            var folder = await picker.PickSingleFolderAsync();
            return folder?.Path;
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error picking folder: {e.StackTrace}");
            return null;
        }
    }

    /// <summary>
    /// Handles the click event for the primary button in the content dialog, validating user input and saving settings
    /// if validation succeeds.
    /// </summary>
    /// <remarks>If input validation fails, the event is canceled and settings are not saved. If validation
    /// succeeds, the current settings are saved and a search operation is initiated.</remarks>
    /// <param name="sender">The content dialog that raised the primary button click event.</param>
    /// <param name="args">The event data for the button click, which can be used to cancel the event if validation fails.</param>
    private void OnPrimaryButtonClick(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        if (!ViewModel.Validate())
        {
            args.Cancel = true;
            return;
        }

        ViewModel.SaveToSettings(_settings);
        StartSearch = true;
    }
}