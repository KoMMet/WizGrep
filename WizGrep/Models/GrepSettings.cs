using System;
using System.Collections.Generic;
using System.Linq;

namespace WizGrep.Models;

/// <summary>
/// Configuration model for a grep (search) operation.
/// Persisted via <see cref="Services.SettingsService"/> and consumed by <see cref="Services.GrepService"/>.
/// Holds search keywords, target file-type flags, extension filters, and search options.
/// </summary>
public class GrepSettings
{
    /// <summary>
    /// Up to five keyword slots that define the search terms.
    /// Only enabled, non-empty keywords participate in the search
    /// (see <see cref="GetActiveKeywords"/>).
    /// </summary>
    public List<SearchKeyword> Keywords { get; set; } =
    [
        new(),
        new(),
        new(),
        new(),
        new()
    ];

    /// <summary>
    /// <c>true</c> for AND search (all keywords must match); <c>false</c> for OR search
    /// (any keyword may match).
    /// </summary>
    public bool IsAndSearch { get; set; } = true;

    /// <summary>Absolute path of the root folder to search recursively.</summary>
    public string TargetFolderPath { get; set; } = string.Empty;

    /// <summary>Include Excel files (.xlsx, .xlsm, .xls) in the search.</summary>
    public bool IncludeExcel { get; set; } = true;

    /// <summary>Include Word files (.docx, .docm, .doc) in the search.</summary>
    public bool IncludeWord { get; set; } = true;

    /// <summary>Include PowerPoint files (.pptx, .pptm, .ppt) in the search.</summary>
    public bool IncludePowerPoint { get; set; } = true;

    /// <summary>Include PDF files (.pdf) in the search.</summary>
    public bool IncludePdf { get; set; } = true;

    /// <summary>Include plain-text files (.txt) in the search.</summary>
    public bool IncludeText { get; set; } = true;

    /// <summary>
    /// When <c>true</c>, all file types are included regardless of individual flags.
    /// Unsupported extensions are read as plain text via the default reader.
    /// </summary>
    public bool IncludeAll { get; set; } = false;

    /// <summary>Perform case-sensitive keyword matching when <c>true</c>.</summary>
    public bool CaseSensitive { get; set; } = false;

    /// <summary>Interpret keywords as .NET regular expressions when <c>true</c>.</summary>
    public bool UseRegex { get; set; } = false;

    /// <summary>
    /// When <c>true</c>, matching results are pushed to the UI via
    /// <see cref="IProgress{GrepProgress}"/> as they are found, rather than
    /// waiting until the entire search completes.
    /// </summary>
    public bool RealTimeDisplay { get; set; } = true;

    /// <summary>
    /// When <c>true</c>, line-break characters inside Excel cell values are replaced
    /// with spaces before displaying the result content.
    /// </summary>
    public bool RemoveExcelLineBreaks { get; set; } = false;

    /// <summary>Enable the custom-extensions list (<see cref="CustomExtensions"/>).</summary>
    public bool UseCustomExtensions { get; set; } = false;

    /// <summary>
    /// Comma-, semicolon-, or space-separated additional file extensions to include
    /// (e.g., ".csv, .xml, .json"). Each extension is normalized to lowercase
    /// and prefixed with a dot if missing.
    /// </summary>
    public string CustomExtensions { get; set; } = string.Empty;

    /// <summary>Enable the exclude-extensions list (<see cref="ExcludeExtensions"/>).</summary>
    public bool UseExcludeExtensions { get; set; } = false;

    /// <summary>Search Excel cell display values (the computed/visible text).</summary>
    public bool IsExcelDisplayValue { get; set; } = true;

    /// <summary>Search Excel cell formula text instead of display values.</summary>
    public bool IsExcelFormulaValue { get; set; } = false;

    /// <summary>
    /// Comma-, semicolon-, or space-separated file extensions to exclude
    /// (e.g., ".tmp, .bak, .log").
    /// </summary>
    public string ExcludeExtensions { get; set; } = string.Empty;

    /// <summary>
    /// Semicolon-separated folder paths to exclude from the search
    /// (e.g., "C:\\Temp; D:\\Logs").
    /// </summary>
    public string ExcludeFolders { get; set; } = string.Empty;

    /// <summary>
    /// Builds the effective list of file extensions to search based on the current flag settings.
    /// </summary>
    /// <returns>
    /// A list of lowercase, dot-prefixed extensions. Returns an empty list when
    /// <see cref="IncludeAll"/> is <c>true</c> (meaning all extensions are accepted).
    /// </returns>
    /// <remarks>
    /// OOXML formats (.xlsx, .docx, .pptx, etc.) are read via the OpenXml SDK;
    /// legacy binary formats (.xls, .doc, .ppt) are read via NPOI / OpenMcdf.
    /// </remarks>
    public List<string> GetTargetExtensions()
    {
        // When "ALL" is selected, return an empty list so every file is included
        if (IncludeAll) return new List<string>();

        var extensions = new List<string>();

        // OOXML formats are read by OpenXml SDK; legacy formats (.xls, .doc, .ppt) by NPOI / OpenMcdf
        if (IncludeExcel) extensions.AddRange([".xlsx", ".xlsm", ".xls"]);
        if (IncludeWord) extensions.AddRange([".docx", ".docm", ".doc"]);
        if (IncludePowerPoint) extensions.AddRange([".pptx", ".pptm", ".ppt"]);
        if (IncludePdf) extensions.Add(".pdf");
        if (IncludeText) extensions.Add(".txt");

        // Append user-specified custom extensions, avoiding duplicates
        if (UseCustomExtensions && !string.IsNullOrWhiteSpace(CustomExtensions))
        {
            var customExts = CustomExtensions
                .Split([',', ';', ' '], StringSplitOptions.RemoveEmptyEntries)
                .Select(e => e.Trim().ToLowerInvariant())
                .Select(e => e.StartsWith(".") ? e : "." + e)
                .Where(e => !extensions.Contains(e));
            extensions.AddRange(customExts);
        }

        return extensions;
    }

    /// <summary>
    /// Builds the effective list of file extensions to exclude from the search.
    /// Extensions are normalized to lowercase with a leading dot and de-duplicated.
    /// </summary>
    /// <returns>An empty list when <see cref="UseExcludeExtensions"/> is <c>false</c> or the text is blank.</returns>
    public List<string> GetExcludeExtensions()
    {
        if (!UseExcludeExtensions || string.IsNullOrWhiteSpace(ExcludeExtensions))
            return new List<string>();

        return ExcludeExtensions
            .Split([',', ';', ' '], StringSplitOptions.RemoveEmptyEntries)
            .Select(e => e.Trim().ToLowerInvariant())
            .Select(e => e.StartsWith(".") ? e : "." + e)
            .Distinct()
            .ToList();
    }

    /// <summary>
    /// Returns only the keywords that are both enabled and non-empty.
    /// These are the terms that <see cref="Services.GrepService"/> will use for matching.
    /// </summary>
    public List<string> GetActiveKeywords()
    {
        return Keywords
            .Where(k => k.IsEnabled && !string.IsNullOrWhiteSpace(k.Keyword))
            .Select(k => k.Keyword)
            .ToList();
    }
}