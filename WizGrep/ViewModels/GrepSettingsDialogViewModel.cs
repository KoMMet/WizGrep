using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using WizGrep.Models;

namespace WizGrep.ViewModels;

/// <summary>
/// Provides a view model for the Grep settings dialog, enabling configuration and validation of search parameters, file
/// type filters, and keyword options for a Grep operation.
/// </summary>
/// <remarks>This class encapsulates the state and logic required to configure and manage search settings in a
/// user interface. It supports asynchronous folder browsing, maintains multiple keyword and file type selections, and
/// offers methods to load from and save to a settings object. Use this view model to bind search configuration options
/// in a dialog and to ensure that all required parameters are validated before executing a search.</remarks>
public partial class GrepSettingsDialogViewModel : ObservableObject
{

    /// <summary>
    /// Gets or sets a function that asynchronously retrieves the path of a selected folder.
    /// </summary>
    /// <remarks>
    /// The assigned function should prompt the user to select a folder and return its path as a
    /// string, or null if no folder is selected. Implementations should handle user interaction and any exceptions that
    /// may occur during the folder selection process.
    /// </remarks>
    public Func<Task<string?>>? BrowseFolderAsync { get; set; }

    [RelayCommand]
    private async Task BrowseFolder()
    {
        if (BrowseFolderAsync is { } browse)
        {
            var path = await browse();
            if (path != null) TargetFolderPath = path;
        }
    }

    [ObservableProperty]
    private string _targetFolderPath = string.Empty;

    [ObservableProperty]
    private string _keyword0 = string.Empty;

    [ObservableProperty]
    private string _keyword1 = string.Empty;

    [ObservableProperty]
    private string _keyword2 = string.Empty;

    [ObservableProperty]
    private string _keyword3 = string.Empty;

    [ObservableProperty]
    private string _keyword4 = string.Empty;

    [ObservableProperty]
    private bool _keywordEnabled0 = true;

    [ObservableProperty]
    private bool _keywordEnabled1 = true;

    [ObservableProperty]
    private bool _keywordEnabled2 = true;

    [ObservableProperty]
    private bool _keywordEnabled3 = true;

    [ObservableProperty]
    private bool _keywordEnabled4 = true;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsOrSearch))]
    private bool _isAndSearch = true;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasAnyTargetFileSelected))]
    private bool _includeExcel = true;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasAnyTargetFileSelected))]
    private bool _includeWord = true;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasAnyTargetFileSelected))]
    private bool _includePowerPoint = true;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasAnyTargetFileSelected))]
    private bool _includePdf = true;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasAnyTargetFileSelected))]
    private bool _includeText = true;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasAnyTargetFileSelected))]
    [NotifyPropertyChangedFor(nameof(IsIndividualFileTypesEnabled))]
    [NotifyPropertyChangedFor(nameof(IsCustomExtensionsTextEnabled))]
    private bool _includeAll;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasAnyTargetFileSelected))]
    [NotifyPropertyChangedFor(nameof(IsCustomExtensionsTextEnabled))]
    private bool _useCustomExtensions;

    [ObservableProperty]
    private string _customExtensions = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsExcludeExtensionsTextEnabled))]
    private bool _useExcludeExtensions;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsExcelDisplayValueEnabled))]
    private bool _isExcelDisplayValue = true;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsExcelFormulaValueEnabled))]
    private bool _isExcelFormulaValue;

    [ObservableProperty]
    private string _excludeExtensions = string.Empty;

    [ObservableProperty]
    private string _excludeFolders = string.Empty;

    [ObservableProperty]
    private bool _caseSensitive;

    [ObservableProperty]
    private bool _useRegex;

    [ObservableProperty]
    private bool _realTimeDisplay = true;

    [ObservableProperty]
    private bool _removeExcelLineBreaks;

    /// <summary>
    /// Gets or sets a value indicating whether the search operation uses 'OR' logic. When set to <see
    /// langword="true"/>, the search returns results that match any of the specified criteria; otherwise, it returns
    /// results that match all criteria.
    /// </summary>
    /// <remarks>
    /// Setting this property to <see langword="true"/> automatically sets the 'AND' search mode to
    /// <see langword="false"/>, and vice versa. Use this property to toggle between 'AND' and 'OR' search modes when
    /// configuring search behavior.
    /// </remarks>
    public bool IsOrSearch
    {
        get => !IsAndSearch;
        set => IsAndSearch = !value;
    }

    /// <summary>
    /// Gets a value indicating whether individual file types are enabled for processing.
    /// </summary>
    /// <remarks>Use this property to determine if the application is configured to handle file types
    /// individually rather than including all file types together. This can affect how file selection and filtering are
    /// applied.</remarks>
    public bool IsIndividualFileTypesEnabled => !IncludeAll;

    /// <summary>
    /// Gets a value indicating whether the custom extensions text option is enabled based on the current configuration
    /// settings.
    /// </summary>
    /// <remarks>This property returns <see langword="true"/> if custom extensions are enabled and all
    /// extensions are not included. It is useful for determining whether the user can specify custom extensions in the
    /// current context.</remarks>
    public bool IsCustomExtensionsTextEnabled => UseCustomExtensions && !IncludeAll;

    /// <summary>
    /// Gets a value indicating whether exclusion of extensions text is enabled.
    /// </summary>
    /// <remarks>This property reflects the current state of the setting that determines if extensions text
    /// should be excluded from processing. Use this property to check whether the exclusion feature is
    /// active.</remarks>
    public bool IsExcludeExtensionsTextEnabled => UseExcludeExtensions;

    /// <summary>
    /// Gets a value indicating whether any target file type is selected for processing.
    /// </summary>
    /// <remarks>This property evaluates to <see langword="true"/> if at least one of the specified file types
    /// (Excel, Word, PowerPoint, PDF, Text, All, or custom extensions) is selected. It is useful for determining if any
    /// file processing should occur based on user selections.</remarks>
    public bool HasAnyTargetFileSelected =>
        IncludeExcel || IncludeWord || IncludePowerPoint ||
        IncludePdf || IncludeText || IncludeAll || UseCustomExtensions;

    public bool IsExcelDisplayValueEnabled => IsExcelDisplayValue;
    public bool IsExcelFormulaValueEnabled => IsExcelFormulaValue;

    /// <summary>
    /// Validates the current settings to ensure that all required conditions for performing a search are met.
    /// </summary>
    /// <remarks>
    /// This method should be called before initiating a search operation to confirm that the
    /// necessary parameters are set. It checks for a non-empty target folder path, the selection of at least one target
    /// file, and the presence of at least one enabled keyword that is not null or whitespace.
    /// </remarks>
    /// <returns>
    /// true if the target folder path is specified, at least one target file is selected, and at least one enabled
    /// keyword is provided; otherwise, false.
    /// </returns>
    public bool Validate()
    {
        if (string.IsNullOrWhiteSpace(TargetFolderPath))
            return false;

        if (!HasAnyTargetFileSelected)
            return false;

        return (!string.IsNullOrWhiteSpace(Keyword0) && KeywordEnabled0) ||
               (!string.IsNullOrWhiteSpace(Keyword1) && KeywordEnabled1) ||
               (!string.IsNullOrWhiteSpace(Keyword2) && KeywordEnabled2) ||
               (!string.IsNullOrWhiteSpace(Keyword3) && KeywordEnabled3) ||
               (!string.IsNullOrWhiteSpace(Keyword4) && KeywordEnabled4);
    }

    /// <summary>
    /// Loads configuration values from the specified GrepSettings object and applies them to the current instance.
    /// </summary>
    /// <remarks>
    /// This method updates the instance's properties based on the values provided in the settings
    /// object, including target folder path, keywords, search options, and file type filters. Ensure that the settings
    /// object contains valid data to avoid unexpected behavior.
    /// </remarks>
    /// <param name="settings">
    /// The GrepSettings object containing the configuration values to be loaded. Must be properly initialized before
    /// calling this method.
    /// </param>
    public void LoadFromSettings(GrepSettings settings)
    {
        TargetFolderPath = settings.TargetFolderPath;

        var keywords = settings.Keywords;
        if (keywords.Count > 0) { Keyword0 = keywords[0].Keyword; KeywordEnabled0 = keywords[0].IsEnabled; }
        if (keywords.Count > 1) { Keyword1 = keywords[1].Keyword; KeywordEnabled1 = keywords[1].IsEnabled; }
        if (keywords.Count > 2) { Keyword2 = keywords[2].Keyword; KeywordEnabled2 = keywords[2].IsEnabled; }
        if (keywords.Count > 3) { Keyword3 = keywords[3].Keyword; KeywordEnabled3 = keywords[3].IsEnabled; }
        if (keywords.Count > 4) { Keyword4 = keywords[4].Keyword; KeywordEnabled4 = keywords[4].IsEnabled; }

        IsAndSearch = settings.IsAndSearch;

        IncludeExcel = settings.IncludeExcel;
        IncludeWord = settings.IncludeWord;
        IncludePowerPoint = settings.IncludePowerPoint;
        IncludePdf = settings.IncludePdf;
        IncludeText = settings.IncludeText;
        IncludeAll = settings.IncludeAll;
        UseCustomExtensions = settings.UseCustomExtensions;
        CustomExtensions = settings.CustomExtensions;
        UseExcludeExtensions = settings.UseExcludeExtensions;
        ExcludeExtensions = settings.ExcludeExtensions;
        ExcludeFolders = settings.ExcludeFolders;
        IsExcelDisplayValue= settings.IsExcelDisplayValue;
        IsExcelFormulaValue = settings.IsExcelFormulaValue;

        CaseSensitive = settings.CaseSensitive;
        UseRegex = settings.UseRegex;
        RealTimeDisplay = settings.RealTimeDisplay;
        RemoveExcelLineBreaks = settings.RemoveExcelLineBreaks;
    }

    /// <summary>
    /// Copies the current search configuration values to the specified <see cref="GrepSettings"/> instance.
    /// </summary>
    /// <remarks>
    /// This method updates the provided settings object with the current values for search
    /// parameters, including folder path, keywords, file type inclusions, and search options. Ensure that the settings
    /// object is properly initialized before calling this method.
    /// </remarks>
    /// <param name="settings">The <see cref="GrepSettings"/>
    /// object to which the current settings will be applied. This parameter must not be null.
    /// </param>
    public void SaveToSettings(GrepSettings settings)
    {
        settings.TargetFolderPath = TargetFolderPath;

        settings.Keywords =
        [
            new() { Keyword = Keyword0, IsEnabled = KeywordEnabled0 },
            new() { Keyword = Keyword1, IsEnabled = KeywordEnabled1 },
            new() { Keyword = Keyword2, IsEnabled = KeywordEnabled2 },
            new() { Keyword = Keyword3, IsEnabled = KeywordEnabled3 },
            new() { Keyword = Keyword4, IsEnabled = KeywordEnabled4 }
        ];

        settings.IsAndSearch = IsAndSearch;

        settings.IncludeExcel = IncludeExcel;
        settings.IncludeWord = IncludeWord;
        settings.IncludePowerPoint = IncludePowerPoint;
        settings.IncludePdf = IncludePdf;
        settings.IncludeText = IncludeText;
        settings.IncludeAll = IncludeAll;
        settings.UseCustomExtensions = UseCustomExtensions;
        settings.CustomExtensions = CustomExtensions;
        settings.UseExcludeExtensions = UseExcludeExtensions;
        settings.ExcludeExtensions = ExcludeExtensions;
        settings.ExcludeFolders = ExcludeFolders;
        settings.IsExcelDisplayValue = IsExcelDisplayValue;
        settings.IsExcelFormulaValue = IsExcelFormulaValue;

        settings.CaseSensitive = CaseSensitive;
        settings.UseRegex = UseRegex;
        settings.RealTimeDisplay = RealTimeDisplay;
        settings.RemoveExcelLineBreaks = RemoveExcelLineBreaks;
    }
}
