using System;
using System.Threading.Tasks;
using Microsoft.UI;
using Microsoft.Windows.Storage.Pickers;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using WinRT.Interop;
using WizGrep.Helpers;
using WizGrep.Models;
using WizGrep.ViewModels;

namespace WizGrep.Views;

/// <summary>
/// Represents a settings dialog for the WizGrep application.
/// </summary>
/// <remarks>
/// This dialog is used to configure and manage settings for the WizGrep application. 
/// It initializes with an optional existing settings object and provides a ViewModel 
/// for data binding and interaction.
/// </remarks>
public sealed partial class WizGrepSettingsDialog : ContentDialog
{
    private readonly Window _parentWindow;
    private readonly WizGrepSettings _settings;

    /// <summary>
    /// Initializes a new instance of the WizGrepSettingsDialog class, allowing the user to configure settings for the
    /// WizGrep feature.
    /// </summary>
    /// <remarks>This constructor sets up the dialog's ViewModel and loads settings from the provided
    /// configuration or defaults. It also initializes the dialog's components. Use this constructor to present the
    /// settings dialog to users and optionally pre-populate it with existing values.</remarks>
    /// <param name="parentWindow">The parent window that will host the settings dialog. This window is used for modal display and positioning of
    /// the dialog.</param>
    /// <param name="existingSettings">An optional WizGrepSettings object containing existing configuration values to load into the dialog. If null,
    /// default settings are used.</param>
    public WizGrepSettingsDialog(Window parentWindow, WizGrepSettings? existingSettings = null)
    {
        _parentWindow = parentWindow;
        _settings = existingSettings ?? new WizGrepSettings();
        ViewModel = new WizGrepSettingsDialogViewModel();
        ViewModel.LoadFromSettings(_settings);
        ViewModel.BrowseFolderAsync = PickFolderAsync;

        InitializeComponent();
    }

    public WizGrepSettingsDialogViewModel ViewModel { get; }
    public WizGrepSettings Settings => _settings;

    /// <summary>
    /// Prompts the user to select a folder and returns the path of the selected folder.
    /// </summary>
    /// <remarks>This method initializes a folder picker dialog starting at the desktop location. It allows
    /// the user to select a single folder. If an exception occurs during the folder selection process, null is
    /// returned.</remarks>
    /// <returns>The path of the selected folder as a string. Returns null if the user cancels the operation or if an error
    /// occurs.</returns>
    private async Task<string?> PickFolderAsync()
    {
        try
        {
            var picker = new FolderPicker(_parentWindow.AppWindow.Id)
            {
                SuggestedStartLocation = PickerLocationId.DocumentsLibrary,
            };

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
    /// Handles the event that occurs when the primary button of the content dialog is clicked, and saves the current
    /// settings.
    /// </summary>
    /// <remarks>Call this method to persist user settings when the primary button is activated. This ensures
    /// that any changes made in the dialog are saved.</remarks>
    /// <param name="sender">The content dialog that raised the click event.</param>
    /// <param name="args">The event data associated with the primary button click.</param>
    private void OnPrimaryButtonClick(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        ViewModel.SaveToSettings(_settings);
    }

    private void OnOssLibClick(object sender, RoutedEventArgs e)
    {
        SettingsPanel.Visibility = Visibility.Collapsed;
        OssPanel.Visibility = Visibility.Visible;
    }

    private void OnBackFromOssClick(object sender, RoutedEventArgs e)
    {
        OssPanel.Visibility = Visibility.Collapsed;
        SettingsPanel.Visibility = Visibility.Visible;
    }
}