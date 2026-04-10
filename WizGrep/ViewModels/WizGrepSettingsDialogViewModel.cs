using System;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using WizGrep.Models;

namespace WizGrep.ViewModels;

/// <summary>
/// Represents the view model for the WizGrep settings dialog, managing user settings related to indexing and folder
/// selection.
/// </summary>
/// <remarks>
/// This view model provides functionality to browse for a folder and load or save settings related to
/// the indexing process. It utilizes observable properties to notify the UI of changes.
/// </remarks>
public partial class WizGrepSettingsDialogViewModel : ObservableObject
{
    [ObservableProperty]
    private string _indexBasePath = String.Empty;

    [ObservableProperty]
    private bool _rebuildIndex;

    [ObservableProperty]
    private bool _showSearchConditionsInExport;

    /// <summary>
    /// Gets or sets a function that asynchronously retrieves the path of a selected folder.
    /// </summary>
    /// <remarks>
    /// Assign a function to this property to provide custom folder browsing functionality. The function
    /// should return a string representing the selected folder path, or null if no folder is chosen. It is recommended that
    /// the function handle any exceptions that may occur during folder selection to ensure robust behavior.
    /// </remarks>
    public Func<Task<string?>>? BrowseFolderAsync { get; set; }

    /// <summary>
    /// Opens a dialog that allows the user to select a folder and updates the base path if a folder is selected.
    /// </summary>
    /// <remarks>If the user cancels the folder selection dialog, the base path is not updated. This method
    /// relies on the BrowseFolderAsync delegate to perform the folder browsing operation.</remarks>
    /// <returns>A task that represents the asynchronous operation.</returns>
    [RelayCommand]
    private async Task BrowseFolder()
    {
        if (BrowseFolderAsync is { } browse)
        {
            var path = await browse();
            if (path != null) IndexBasePath = path;
        }
    }

    /// <summary>
    /// Loads configuration values from the specified settings object, updating the index base path and rebuild index
    /// flag accordingly.
    /// </summary>
    /// <remarks>This method updates the internal state of the object based on the provided settings. Ensure
    /// that the settings object is properly initialized before calling this method.</remarks>
    /// <param name="settings">The settings object containing the configuration values to apply. Cannot be null.</param>
    public void LoadFromSettings(WizGrepSettings settings)
    {
        IndexBasePath = settings.IndexBasePath;
        RebuildIndex = settings.RebuildIndex;
        ShowSearchConditionsInExport = settings.ShowSearchConditionsInExport;
    }

    /// <summary>
    /// Saves the current index base path and rebuild index settings to the specified WizGrepSettings instance.
    /// </summary>
    /// <remarks>This method updates the IndexBasePath and RebuildIndex properties of the provided settings
    /// object with the current values.</remarks>
    /// <param name="settings">The WizGrepSettings object to which the current index base path and rebuild index values will be applied. This
    /// parameter cannot be null.</param>
    public void SaveToSettings(WizGrepSettings settings)
    {
        settings.IndexBasePath = IndexBasePath;
        settings.RebuildIndex = RebuildIndex;
        settings.ShowSearchConditionsInExport = ShowSearchConditionsInExport;
    }
}
