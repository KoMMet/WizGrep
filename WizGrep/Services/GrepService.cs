using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using WizGrep.Helpers;
using WizGrep.Models;

namespace WizGrep.Services;

/// <summary>
/// Represents the progress of a grep operation, including information about the current file being processed, the
/// number of files processed, the total number of files, and any result found so far.
/// </summary>
/// <remarks>Use this class to monitor the state of an ongoing grep operation. The properties provide real-time
/// updates on which file is currently being searched, how many files have been processed, and the overall scope of the
/// operation. The FoundResult property may be set if a matching result is found during processing.</remarks>
public class GrepProgress
{
    /// <summary>Full path of the file currently being scanned.</summary>
    public string CurrentFile { get; set; } = string.Empty;

    /// <summary>Number of files processed so far.</summary>
    public int ProcessedFiles { get; set; }

    /// <summary>Total number of files to process in this grep operation.</summary>
    public int TotalFiles { get; set; }

    /// <summary>A match found during the current file scan, or null if none yet.</summary>
    public GrepResult? FoundResult { get; set; }
}

/// <summary>
/// Provides functionality to perform grep searches on files within a specified directory, utilizing optional indexing
/// for improved performance.
/// </summary>
/// <remarks>This class allows for both real-time display of results and the option to rebuild the index. It
/// supports keyword searches with both plain text and regular expressions, and handles various file types based on
/// specified extensions.</remarks>
/// <param name="fileReaderService">The service responsible for reading the contents of files during the grep operation.</param>
/// <param name="indexService">The service used to manage and access the index for faster grep searches, if enabled.</param>
public class GrepService(FileReaderService fileReaderService, IndexService indexService)
{
    /// <summary>
    /// The maximum duration allowed for a regular expression match operation before timing out.
    /// </summary>
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(5);
    /// <summary>
    /// Asynchronously executes a grep operation using the specified settings, searching for keywords within target
    /// files and returning the matching results.
    /// </summary>
    /// <remarks>If index rebuilding is enabled and an index base path is specified, the method will delete
    /// and recreate the search index before performing the operation. The method supports both direct file scanning and
    /// indexed searches, depending on the availability and validity of the index. Progress updates and cancellation are
    /// supported for long-running operations.</remarks>
    /// <param name="grepSettings">The settings that define the grep operation, including the target folder path, keywords to search for, file
    /// extension filters, and search options such as case sensitivity and regular expression usage.</param>
    /// <param name="wizGrepSettings">The settings related to WizGrep functionality, including options for index rebuilding and the base path for
    /// storing or retrieving search indexes.</param>
    /// <param name="progress">An optional progress reporter that receives updates about the progress of the grep operation, such as the
    /// current file being processed and the number of files completed.</param>
    /// <param name="cancellationToken">A cancellation token that can be used to cancel the grep operation before it completes.</param>
    /// <returns>A list of grep results, where each result contains information about the file path, line number, and content
    /// that matched the specified keywords. Returns an empty list if no matches are found or if the search criteria are
    /// not met.</returns>
    /// <exception cref="InvalidOperationException">Thrown if a provided regular expression pattern in the search keywords is invalid during the validation process.</exception>
    public async Task<List<GrepResult>> ExecuteGrepAsync(
        GrepSettings grepSettings,
        WizGrepSettings wizGrepSettings,
        IProgress<GrepProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var results = new List<GrepResult>();
        var keywords = grepSettings.GetActiveKeywords();

        if (keywords.Count == 0 || string.IsNullOrEmpty(grepSettings.TargetFolderPath)) return results;

        // Resolve relative paths against the application's base directory
        // because simply using Path.GetFullPath defaults to System32 in packaged apps.
        var targetFolderPath = Path.IsPathRooted(grepSettings.TargetFolderPath)
            ? Path.GetFullPath(grepSettings.TargetFolderPath)
            : Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, grepSettings.TargetFolderPath));

        // Check if regex is enabled and configure options accordingly
        if (grepSettings.UseRegex)
        {
            var regexOptions = grepSettings.CaseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
            foreach (var keyword in keywords)
                try
                {
                    _ = Regex.Match("", keyword, regexOptions, RegexTimeout);
                }
                catch (ArgumentException ex)
                {
                    LoggerHelper.Instance.LogError($"Invalid regular expression pattern: \"{keyword}\" - {ex.StackTrace}");

                    throw new InvalidOperationException(
                        $"{ResourceLoaderHelper.GetString("RegularExceptionMessage")}: \"{keyword}\" - {ex.Message}", ex);
                }
        }

        // If the rebuild index option is enabled and a valid index base path is provided,
        // delete the existing index for the specified target folder path.
        if (wizGrepSettings.RebuildIndex && !string.IsNullOrEmpty(wizGrepSettings.IndexBasePath))
            indexService.DeleteIndex(wizGrepSettings.IndexBasePath, targetFolderPath);

        // Enumerate target files
        var targetExtensions = grepSettings.GetTargetExtensions();
        var excludeExtensions = grepSettings.GetExcludeExtensions();

        // If no target extensions are specified and "Include All" is not selected, there are no files to search
        if (targetExtensions.Count == 0 && !grepSettings.IncludeAll) return results;

        var excludeFolders = NormalizeExcludeFolders(grepSettings.ExcludeFolders, targetFolderPath);
        var targetFiles = EnumerateTargetFiles(
            targetFolderPath,
            targetExtensions,
            excludeExtensions,
            excludeFolders);
        var totalFiles = targetFiles.Count;

        // Reflect the setting for whether to search formulas in Excel files
        var excelFormula = grepSettings.IsExcelFormulaValue;

        // Check whether per-file indexing is available
        var useIndex = !string.IsNullOrEmpty(wizGrepSettings.IndexBasePath);

        var processedFiles = 0;

        foreach (var filePath in targetFiles)
        {
            cancellationToken.ThrowIfCancellationRequested();

            progress?.Report(new GrepProgress
            {
                CurrentFile = filePath,
                ProcessedFiles = processedFiles,
                TotalFiles = totalFiles
            });

            var fileInfo = new FileInfo(filePath);
            var lastModified = fileInfo.LastWriteTime;

            List<GrepResult> fileContents;
            var loadedFromIndex = false;

            // Try to load from the per-file index when the timestamp and excelFormula flag match
            if (useIndex &&
                indexService.FileIndexExists(wizGrepSettings.IndexBasePath, targetFolderPath, filePath))
            {
                var (storedTime, storedFormula) = indexService.LoadFileTimestamp(
                    wizGrepSettings.IndexBasePath, targetFolderPath, filePath);

                if (storedTime.HasValue &&
                    Math.Abs((storedTime.Value - lastModified).TotalSeconds) < 1 &&
                    storedFormula == excelFormula)
                {
                    fileContents = indexService.LoadFileIndex(
                        wizGrepSettings.IndexBasePath, targetFolderPath, filePath);
                    loadedFromIndex = true;
                }
                else
                {
                    fileContents = fileReaderService.ReadFile(filePath, excelFormula).ToList();
                }
            }
            else
            {
                fileContents = fileReaderService.ReadFile(filePath, excelFormula).ToList();
            }

            // Save the per-file index only when the file was actually read from disk
            if (useIndex && !loadedFromIndex)
            {
                try
                {
                    await indexService.SaveFileIndexAsync(
                        wizGrepSettings.IndexBasePath,
                        targetFolderPath,
                        filePath,
                        fileContents,
                        lastModified,
                        excelFormula);
                }
                catch (Exception e)
                {
                    LoggerHelper.Instance.LogError(
                        $"Error saving index for '{filePath}': {e.StackTrace}");
                }
            }

            // Execute search
            var isExcel = IsExcelFile(filePath);
            var useRowAnd = grepSettings.IsAndSearch && isExcel;

            // Pre-compute the set of entries whose AND evaluation should be performed per-row
            // (Excel row cells only; shapes/comments are evaluated per entry).
            HashSet<GrepResult>? rowMatchedEntries = null;
            if (useRowAnd)
            {
                rowMatchedEntries = new HashSet<GrepResult>();
                var rowGroups = fileContents
                    .Where(c => c.LineNumber > 0
                                && !string.IsNullOrEmpty(c.CellAddress)
                                && string.IsNullOrEmpty(c.ObjectName))
                    .GroupBy(c => (c.SheetName, c.LineNumber));

                foreach (var rowGroup in rowGroups)
                {
                    var rowText = string.Join("\n", rowGroup.Select(c => c.Content));
                    if (!MatchesKeywords(rowText, keywords, grepSettings)) continue;

                    // The row matches AND across all keywords: emit only cells that contain at least one keyword.
                    foreach (var cell in rowGroup)
                        if (keywords.Any(k => MatchesKeyword(cell.Content, k, grepSettings)))
                            rowMatchedEntries.Add(cell);
                }
            }

            foreach (var content in fileContents)
            {
                bool matched;
                if (useRowAnd
                    && content.LineNumber > 0
                    && !string.IsNullOrEmpty(content.CellAddress)
                    && string.IsNullOrEmpty(content.ObjectName))
                {
                    // Excel row cell: decided by the per-row AND evaluation above.
                    matched = rowMatchedEntries!.Contains(content);
                }
                else
                {
                    // Non-Excel, OR search, or Excel shapes/comments: evaluate per entry.
                    matched = MatchesKeywords(content.Content, keywords, grepSettings);
                }

                if (!matched) continue;

                // If the file is an Excel file, remove line breaks within cells according to the settings
                var displayContent = content.Content;
                if (grepSettings.RemoveExcelLineBreaks && isExcel)
                    displayContent = displayContent.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");

                var result = new GrepResult
                {
                    FilePath = content.FilePath,
                    LineNumber = content.LineNumber,
                    SheetName = content.SheetName,
                    CellAddress = content.CellAddress,
                    ObjectName = content.ObjectName,
                    Content = displayContent
                };
                results.Add(result);

                if (grepSettings.RealTimeDisplay)
                    progress?.Report(new GrepProgress
                    {
                        CurrentFile = filePath,
                        ProcessedFiles = processedFiles,
                        TotalFiles = totalFiles,
                        FoundResult = result
                    });
            }

            processedFiles++;
        }

        return results;
    }

    /// <summary>
    /// Enumerates all files in the specified directory and its subdirectories, including only files with the specified
    /// extensions and excluding files with the specified extensions.
    /// </summary>
    /// <remarks>The search is performed recursively through all subdirectories. Inaccessible files and
    /// directories are ignored.</remarks>
    /// <param name="folderPath">The full path of the directory to search. The path must exist and be accessible.</param>
    /// <param name="extensions">A list of file extensions to include in the results. If the list is empty, all file extensions are included.</param>
    /// <param name="excludeExtensions">A list of file extensions to exclude from the results. Files with these extensions are omitted.</param>
    /// <returns>A list of file paths that match the specified inclusion and exclusion criteria.</returns>
    /// <exception cref="InvalidOperationException">Thrown when an error occurs during file enumeration, such as if the directory is inaccessible or the path is
    /// invalid.</exception>
    private List<string> EnumerateTargetFiles(
        string folderPath,
        List<string> extensions,
        List<string> excludeExtensions,
        List<string> excludeFolders)
    {
        var files = new List<string>();

        try
        {
            var options = new EnumerationOptions
            {
                RecurseSubdirectories = true,
                IgnoreInaccessible = true,
                AttributesToSkip = 0
            };

            foreach (var file in Directory.EnumerateFiles(folderPath, "*", options))
            {
                var normalizedFile = Path.GetFullPath(file);
                if (excludeFolders.Count > 0 && excludeFolders.Any(folder => normalizedFile.StartsWith(folder, StringComparison.OrdinalIgnoreCase)))
                    continue;

                var ext = Path.GetExtension(file).ToLowerInvariant();

                // Skip files with excluded extensions
                if (excludeExtensions.Count > 0 && excludeExtensions.Contains(ext))
                    continue;

                // If no target extensions are specified (ALL selected), include all files
                if (extensions.Count == 0)
                {
                    files.Add(file);
                }
                else
                {
                    if (extensions.Contains(ext)) files.Add(file);
                }
            }
        }
        catch (Exception ex)
        {
            LoggerHelper.Instance.LogError($"Error enumerating files in '{folderPath}': {ex.StackTrace}");

            throw new InvalidOperationException($"{ResourceLoaderHelper.GetString("EnumeratingFilesExceptionMessage")}: {ex.Message}", ex);
        }

        return files;
    }

    /// <summary>
    /// Normalizes a semicolon-separated list of folder paths by resolving relative paths to absolute paths
    /// and ensuring consistent formatting for comparison.
    /// </summary>
    /// <param name="excludeFolders">
    /// A semicolon-separated string containing folder paths to exclude. 
    /// Relative paths will be resolved based on the provided <paramref name="basePath"/>.
    /// </param>
    /// <param name="basePath">
    /// The base directory used to resolve relative folder paths in <paramref name="excludeFolders"/>.
    /// </param>
    /// <returns>
    /// A list of normalized and distinct folder paths, formatted with a trailing directory separator.
    /// </returns>
    private static List<string> NormalizeExcludeFolders(string? excludeFolders, string basePath)
    {
        if (string.IsNullOrWhiteSpace(excludeFolders))
            return new List<string>();

        return excludeFolders
            .Split([';'], StringSplitOptions.RemoveEmptyEntries)
            .Select(folder => folder.Trim())
            .Where(folder => !string.IsNullOrWhiteSpace(folder))
            .Select(folder => Path.IsPathRooted(folder) ? folder : Path.Combine(basePath, folder))
            .Select(folder => Path.GetFullPath(folder))
            .Select(folder => Path.TrimEndingDirectorySeparator(folder) + Path.DirectorySeparatorChar)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>
    /// Determines whether the specified file path has an extension that corresponds to a supported Excel file format.
    /// </summary>
    /// <remarks>The comparison is case-insensitive. Only the file extension is evaluated; the method does not
    /// verify the file's contents.</remarks>
    /// <param name="filePath">The path of the file to check. This value must not be null.</param>
    /// <returns>true if the file has an Excel extension (.xlsx, .xlsm, or .xls); otherwise, false.</returns>
    private static bool IsExcelFile(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext is ".xlsx" or ".xlsm" or ".xls";
    }

    /// <summary>
    /// Determines whether the specified content matches all or any of the provided keywords according to the specified
    /// search settings.
    /// </summary>
    /// <remarks>Use this method to flexibly match content against multiple keywords, supporting both strict
    /// (AND) and broad (OR) search modes as defined by the provided settings.</remarks>
    /// <param name="content">The text to search for keyword matches.</param>
    /// <param name="keywords">A list of keywords to evaluate against the content. The search will require all or any keywords to match,
    /// depending on the settings.</param>
    /// <param name="settings">The search settings that specify whether to perform an AND (all keywords) or OR (any keyword) match.</param>
    /// <returns>true if the content matches all keywords when an AND search is specified, or any keyword when an OR search is
    /// specified; otherwise, false.</returns>
    private static bool MatchesKeywords(string content, List<string> keywords, GrepSettings settings)
    {
        if (settings.IsAndSearch)
            // AND search: all keywords must match
            return keywords.All(k => MatchesKeyword(content, k, settings));

        // OR search: any keyword match is sufficient
        return keywords.Any(k => MatchesKeyword(content, k, settings));
    }

    /// <summary>
    /// Determines whether the specified content matches the given keyword, using either regular expression or plain
    /// text matching based on the provided settings.
    /// </summary>
    /// <remarks>If the UseRegex property of settings is true, the method performs a regular expression match;
    /// otherwise, it uses plain text matching. Case sensitivity is determined by the CaseSensitive property of
    /// settings.</remarks>
    /// <param name="content">The string content to be searched for the specified keyword.</param>
    /// <param name="keyword">The keyword to search for within the content. This can be a plain text string or a regular expression depending
    /// on the settings.</param>
    /// <param name="settings">An instance of GrepSettings that specifies whether to use regular expressions and case sensitivity for the
    /// matching operation.</param>
    /// <returns>true if the content matches the keyword according to the specified settings; otherwise, false.</returns>
    private static bool MatchesKeyword(string content, string keyword, GrepSettings settings)
    {
        if (settings.UseRegex)
        {
            var options = settings.CaseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
            try
            {
                return Regex.IsMatch(content, keyword, options, RegexTimeout);
            }
            catch (RegexMatchTimeoutException e)
            {
                LoggerHelper.Instance.LogError($"Regex match timeout for pattern: \"{keyword}\" - {e.StackTrace}");
                return false;
            }
        }

        return MatchesPlainText(content, keyword, settings.CaseSensitive);
    }

    /// <summary>
    /// Determines whether the specified keyword exists within the given content, using the specified case sensitivity.
    /// </summary>
    /// <param name="content">The string to search for the keyword.</param>
    /// <param name="keyword">The keyword to locate within the content.</param>
    /// <param name="caseSensitive">A value indicating whether the search should be case-sensitive. Specify <see langword="true"/> to perform a
    /// case-sensitive search; otherwise, <see langword="false"/>.</param>
    /// <returns>true if the keyword is found in the content; otherwise, false.</returns>
    private static bool MatchesPlainText(string content, string keyword, bool caseSensitive)
    {
        var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        return content.Contains(keyword, comparison);
    }
}