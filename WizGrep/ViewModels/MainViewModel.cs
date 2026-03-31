using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.UI.Xaml;
using WizGrep.Helpers;
using WizGrep.Models;
using WizGrep.Services;

namespace WizGrep.ViewModels;


/// <summary>
/// Represents the main view model for the application, providing search functionality and managing application settings
/// and results.
/// </summary>
/// <remarks>
/// MainViewModel coordinates search operations by interacting with GrepService and SettingsService. It
/// exposes properties and commands for starting and canceling searches, tracking progress, and exporting results. The
/// view model maintains collections of search results and matched files, and updates UI-related properties to reflect
/// the current search status. Use this class as the central binding point for UI elements that display search progress,
/// results, and settings.
/// </remarks>
public partial class MainViewModel : ObservableObject
{
    private readonly GrepService _grepService;
    private readonly SettingsService _settingsService;
    private readonly HashSet<string> _matchedFilesSet = new(StringComparer.OrdinalIgnoreCase);
    private CancellationTokenSource? _cancellationTokenSource;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ProgressVisibility))]
    [NotifyPropertyChangedFor(nameof(CancelVisibility))]
    [NotifyPropertyChangedFor(nameof(IsNotSearching))]
    private bool _isSearching;

    [ObservableProperty]
    private string _currentProcessingFile = string.Empty;

    [ObservableProperty]
    private ObservableCollection<GrepResult> _grepResults = null!;

    [ObservableProperty]
    private ObservableCollection<string> _matchedFiles = null!;

    [ObservableProperty]
    private int _processedFiles;

    [ObservableProperty]
    private string _statusMessage = $"{ResourceLoaderHelper.GetString("PreparationCompleteLabel")}";

    [ObservableProperty]
    private int _totalFiles;
    
    /// <summary>
    /// Initializes a new instance of the MainViewModel class, configuring search and settings services and loading
    /// initial application settings.
    /// </summary>
    /// <remarks>
    /// The constructor subscribes to changes in the search results collection to update the result
    /// count when the collection changes.
    /// </remarks>
    /// <param name="grepService">The service used to perform search operations within files and directories.</param>
    /// <param name="settingsService">The service responsible for loading and managing application settings, including search configuration.</param>
    public MainViewModel(GrepService grepService, SettingsService settingsService)
    {
        _grepResults = [];
        _matchedFiles = [];
        
        _grepService = grepService;
        _settingsService = settingsService;
        GrepSettings = _settingsService.LoadGrepSettings();
        WizGrepSettings = _settingsService.LoadWizGrepSettings();

        // Updates the result count text whenever the grep results collection changes.
        _grepResults.CollectionChanged += (_, _) => OnPropertyChanged(nameof(ResultCountText));
    }

    public GrepSettings GrepSettings { get; private set; }
    public WizGrepSettings WizGrepSettings { get; private set; }

    // Determines the visibility of the progress indicator based on the current search state.
    public Visibility ProgressVisibility => IsSearching ? Visibility.Visible : Visibility.Collapsed;
    public Visibility CancelVisibility => IsSearching ? Visibility.Visible : Visibility.Collapsed;
    public bool IsNotSearching => !IsSearching;
    public string ResultCountText => $"{GrepResults.Count}{ResourceLoaderHelper.GetString("ResultCountSuffix")}";
    public string ProgressText => $"{ResourceLoaderHelper.GetString("ProcessingLabel")}: {CurrentProcessingFile} ({ProcessedFiles}/{TotalFiles})";

    partial void OnCurrentProcessingFileChanged(string value) => OnPropertyChanged(nameof(ProgressText));
    partial void OnProcessedFilesChanged(int value) => OnPropertyChanged(nameof(ProgressText));
    partial void OnTotalFilesChanged(int value) => OnPropertyChanged(nameof(ProgressText));

   /// <summary>
   /// Reloads the application settings by loading the latest Grep and WizGrep settings from the settings service.
   /// </summary>
   /// <remarks>Call this method to refresh the application's configuration at runtime without restarting. This
   /// is useful when settings may have changed externally and need to be re-applied.</remarks>
    public void ReloadSettings()
    {
        // Load the latest settings for Grep and WizGrep from the settings service.
        GrepSettings = _settingsService.LoadGrepSettings();
        WizGrepSettings = _settingsService.LoadWizGrepSettings();
    }

   /// <summary>
   /// Saves the specified <see cref="GrepSettings"/> to the application's settings storage.
   /// </summary>
   /// <param name="settings">The <see cref="GrepSettings"/> instance containing the configuration to be saved.</param>
   /// <remarks>
   /// This method uses the <see cref="SettingsService"/> to persist the provided settings. 
   /// After saving, it reloads the settings to ensure the application reflects the latest changes.
   /// </remarks>
    public void SaveGrepSettings(GrepSettings settings)
    {
        // Persist the provided GrepSettings and refresh the application settings.
        _settingsService.SaveGrepSettings(settings);
        ReloadSettings();
    }


   /// <summary>
   /// Saves the specified WizGrepSettings to persistent storage and reloads the current settings.
   /// </summary>
   /// <remarks>
   /// This method ensures that the latest settings are persisted and subsequently reloaded for use.
   /// It is important to ensure that the settings provided are valid before calling this method.
   /// </remarks>
   /// <param name="settings">The WizGrepSettings object containing the settings to be saved. This parameter cannot be null.</param>
    public void SaveWizGrepSettings(WizGrepSettings settings)
    {
        // Persist the provided settings and refresh the current configuration.
        _settingsService.SaveWizGrepSettings(settings);
        ReloadSettings();
    }

    /// <summary>
    /// Starts an asynchronous search operation for files matching the specified criteria.
    /// </summary>
    /// <remarks>
    /// The search operation can be canceled using a cancellation token. The method updates the status
    /// message to indicate the progress of the search and handles exceptions that may occur during the operation. If
    /// real-time display is enabled, results are added as they are found; otherwise, all results are added after the search
    /// completes.
    /// </remarks>
    /// <returns>This method does not return a value.</returns>
    [RelayCommand]
    public async Task StartSearchAsync()
    {
        if (IsSearching) return;

        IsSearching = true;
        MatchedFiles.Clear();
        GrepResults.Clear();
        _matchedFilesSet.Clear();
        StatusMessage = $"{ResourceLoaderHelper.GetString("SearchingLabel")}...";

        _cancellationTokenSource = new CancellationTokenSource();

        var grepSettings = GrepSettings;
        var wizGrepSettings = WizGrepSettings;
        var realTimeDisplay = grepSettings.RealTimeDisplay;
        var token = _cancellationTokenSource.Token;

        try
        {
            var progress = new Progress<GrepProgress>(p =>
            {
                CurrentProcessingFile = p.CurrentFile;
                ProcessedFiles = p.ProcessedFiles;
                TotalFiles = p.TotalFiles;

                if (p.FoundResult != null && realTimeDisplay) AddResult(p.FoundResult);
            });

            List<GrepResult> results = await Task.Run(() => _grepService.ExecuteGrepAsync(
                grepSettings,
                wizGrepSettings,
                progress,
                token));

            // Auto-reset RebuildIndex after successful rebuild so the index
            // is not deleted on every subsequent search.
            if (wizGrepSettings.RebuildIndex)
            {
                wizGrepSettings.RebuildIndex = false;
                SaveWizGrepSettings(wizGrepSettings);
            }

            if (!realTimeDisplay)
                foreach (var result in results)
                    AddResult(result);

            StatusMessage = $"{ResourceLoaderHelper.GetString("SearchCompleteLabel")}: {GrepResults.Count}{ResourceLoaderHelper.GetString("FoundLabel")}";
        }
        catch (OperationCanceledException e)
        {
            LoggerHelper.Instance.LogInfo($"Search operation was canceled: {e.Message}");
            StatusMessage = $"{ResourceLoaderHelper.GetString("SearchCanceledLabel")}";
        }
        catch (Exception ex)
        {
            LoggerHelper.Instance.LogError($"An error occurred during the search operation: {ex.StackTrace}");
            StatusMessage = $"Error: {ex.Message} ";
        }
        finally
        {
            IsSearching = false;
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;
        }
    }

    /// <summary>
    /// Cancels the ongoing search operation, if any.
    /// </summary>
    /// <remarks>
    /// This method signals cancellation to any active search process. The search operation should be
    /// designed to respond to cancellation requests promptly. Calling this method has no effect if no search is
    /// currently in progress.
    /// </remarks>
    [RelayCommand]
    public void CancelSearch()
    {
        _cancellationTokenSource?.Cancel();
    }

    /// <summary>
    /// Adds a grep result to the collection and records the associated file path if it has not already been matched.
    /// </summary>
    /// <remarks>
    /// This method ensures that each file path is tracked only once in the matched files collection,
    /// preventing duplicate entries.</remarks>
    /// <param name="result">The grep result to add, containing information about the match and its file path.
    /// </param>
    private void AddResult(GrepResult result)
    {
        GrepResults.Add(result);

        if (_matchedFilesSet.Add(result.FilePath))
        {
            MatchedFiles.Add(result.FilePath);
        }
    }

    /// <summary>
    /// Gets or sets a function that asynchronously exports data from a source file to a destination file.
    /// </summary>
    /// <remarks>
    /// The function should accept the source file path and the destination file path as parameters
    /// and return a task representing the asynchronous operation. Ensure that both file paths are valid and accessible
    /// to avoid exceptions during execution.
    /// </remarks>
    public Func<string, string, Task>? ExportToFileAsync { get; set; }

    /// <summary>
    /// Exports the list of matched files to a text file asynchronously.
    /// </summary>
    /// <remarks>
    /// This method uses the delegate specified by <c>ExportToFileAsync</c> to write the file list
    /// generated by <c>BuildFileListText()</c> to a file named 'MatchedFiles.txt'. Ensure that <c>ExportToFileAsync</c>
    /// is properly initialized before calling this method. The operation is performed asynchronously and may take time
    /// depending on the number of files.
    /// </remarks>
    /// <returns></returns>
    [RelayCommand]
    private async Task ExportFiles()
    {
        if (ExportToFileAsync is { } export)
            await export(BuildFileListText(), "MatchedFiles.txt");
    }

    /// <summary>
    /// Asynchronously exports the search results to a file named 'GrepResults.txt'.
    /// </summary>
    /// <remarks>
    /// This method relies on the ExportToFileAsync delegate being initialized. The results are
    /// generated by the BuildResultsText method and written to the specified file. If ExportToFileAsync is not set, the
    /// method does not perform any export operation.
    /// </remarks>
    /// <returns></returns>
    [RelayCommand]
    private async Task ExportResults()
    {
        if (ExportToFileAsync is { } export)
            await export(BuildResultsText(), "GrepResults.txt");
    }

    /// <summary>
    /// Builds a text representation of the matched files.
    /// </summary>
    /// <remarks>
    /// This method concatenates the file paths of all matched files into a single string, 
    /// with each file path separated by a new line. It is typically used for exporting 
    /// or displaying the list of matched files.
    /// </remarks>
    /// <returns>A string containing the file paths of all matched files, separated by new lines.</returns>
    public string BuildFileListText()
    {
        var sb = new StringBuilder();
        if (WizGrepSettings.ShowSearchConditionsInExport)
        {
            sb.Append(BuildSearchConditionsHeader());
        }
        sb.Append(string.Join(Environment.NewLine, MatchedFiles));
        return sb.ToString();
    }

    /// <summary>
    /// Generates a tab-separated string that represents the grep results, including file paths, file names, locations,
    /// and content.
    /// </summary>
    /// <remarks>
    /// The returned string begins with a header row and includes all entries from the GrepResults
    /// collection. This format is suitable for easy parsing or exporting to spreadsheet applications.
    /// </remarks>
    /// <returns>A string containing the formatted grep results. Each entry appears on a new line, with fields separated by tabs.</returns>
    public string BuildResultsText()
    {
        var lines = new List<string>();
        if (WizGrepSettings.ShowSearchConditionsInExport)
        {
            lines.Add(BuildSearchConditionsHeader().TrimEnd());
        }
        // header line
        lines.Add($"{ResourceLoaderHelper.GetString("FilePathLabel")}\t{ResourceLoaderHelper.GetString("FileNameLabel")}\t{ResourceLoaderHelper.GetString("LocationLabel")}\t{ResourceLoaderHelper.GetString("ContentLabel")}");
        foreach (var result in GrepResults)
        {
            lines.Add($"{result.FilePath}\t{result.FileName}\t{EscapeForTsv(result.Location)}\t{EscapeForTsv(result.Content)}");
        }
        return string.Join(Environment.NewLine, lines);
    }

    /// <summary>
    /// Builds a search conditions header string based on the current <see cref="GrepSettings"/>.
    /// Only active/selected settings are included in the output.
    /// </summary>
    private string BuildSearchConditionsHeader()
    {
        const string separator = "--------------";
        var sb = new StringBuilder();
        sb.AppendLine(separator);

        var settings = GrepSettings;

        // Target folder
        if (!string.IsNullOrWhiteSpace(settings.TargetFolderPath))
            sb.AppendLine($"{ResourceLoaderHelper.GetString("ExportHeader_TargetFolder")}:{settings.TargetFolderPath}");

        // Keywords (only enabled and non-empty)
        int keywordIndex = 0;
        foreach (var kw in settings.Keywords)
        {
            keywordIndex++;
            if (kw.IsEnabled && !string.IsNullOrWhiteSpace(kw.Keyword))
                sb.AppendLine($"{ResourceLoaderHelper.GetString("ExportHeader_Keyword")}{keywordIndex}:{kw.Keyword}");
        }

        // Search condition (AND/OR)
        sb.AppendLine($"{ResourceLoaderHelper.GetString("ExportHeader_SearchCondition")}:{(settings.IsAndSearch ? "AND" : "OR")}");

        // Target files
        var fileTypes = new List<string>();
        if (settings.IncludeAll)
        {
            fileTypes.Add("ALL");
        }
        else
        {
            if (settings.IncludeExcel) fileTypes.Add("Excel");
            if (settings.IncludeWord) fileTypes.Add("Word");
            if (settings.IncludePowerPoint) fileTypes.Add("PowerPoint");
            if (settings.IncludePdf) fileTypes.Add("PDF");
            if (settings.IncludeText) fileTypes.Add("Text");
        }
        if (fileTypes.Count > 0)
            sb.AppendLine($"{ResourceLoaderHelper.GetString("ExportHeader_TargetFiles")}: {string.Join(", ", fileTypes)}");

        // Custom extensions
        if (settings.UseCustomExtensions && !string.IsNullOrWhiteSpace(settings.CustomExtensions))
            sb.AppendLine($"{ResourceLoaderHelper.GetString("ExportHeader_CustomExtensions")}: {settings.CustomExtensions}");

        // Exclude extensions
        if (settings.UseExcludeExtensions && !string.IsNullOrWhiteSpace(settings.ExcludeExtensions))
            sb.AppendLine($"{ResourceLoaderHelper.GetString("ExportHeader_ExcludeExtensions")}: {settings.ExcludeExtensions}");

        // Exclude folders
        if (!string.IsNullOrWhiteSpace(settings.ExcludeFolders))
            sb.AppendLine($"{ResourceLoaderHelper.GetString("ExportHeader_ExcludeFolders")}: {settings.ExcludeFolders}");

        // Search options (only checked ones)
        var options = new List<string>();
        if (settings.CaseSensitive) options.Add(ResourceLoaderHelper.GetString("ExportHeader_CaseSensitive")!);
        if (settings.UseRegex) options.Add(ResourceLoaderHelper.GetString("ExportHeader_UseRegex")!);
        if (settings.RealTimeDisplay) options.Add(ResourceLoaderHelper.GetString("ExportHeader_RealTimeDisplay")!);
        if (settings.RemoveExcelLineBreaks) options.Add(ResourceLoaderHelper.GetString("ExportHeader_RemoveExcelLineBreaks")!);
        if (options.Count > 0)
        {
            sb.AppendLine(ResourceLoaderHelper.GetString("ExportHeader_SearchOptions"));
            foreach (var option in options)
                sb.AppendLine($"          {option}");
        }

        // Excel search target (only if Excel is included)
        if (settings.IncludeExcel || settings.IncludeAll)
        {
            var excelTarget = settings.IsExcelDisplayValue
                ? ResourceLoaderHelper.GetString("ExportHeader_DisplayValue")
                : ResourceLoaderHelper.GetString("ExportHeader_FormulaValue");
            sb.AppendLine($"{ResourceLoaderHelper.GetString("ExportHeader_ExcelSearchTarget")}:{excelTarget}");
        }

        sb.AppendLine(separator);
        return sb.ToString();
    }

    /// <summary>
    /// Escapes special characters in a string for use in a TSV (Tab-Separated Values) format by replacing newline and
    /// tab characters with spaces.
    /// </summary>
    /// <param name="value">
    /// The input string to be escaped for TSV formatting. This string may contain newline or tab
    /// characters that will be replaced.
    /// </param>
    /// <returns>
    /// A string with newline and tab characters replaced by spaces, making it suitable for TSV output.
    /// </returns>
    private static string EscapeForTsv(string value)
    {
        return value
            .Replace("\r\n", " ")
            .Replace("\r", " ")
            .Replace("\n", " ")
            .Replace("\t", " ");
    }
}