using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using WizGrep.Models;

namespace WizGrep.Services;

/// <summary>
/// Provides per-file index and timestamp management for grep results.
/// Each source file gets its own <c>{filename}_index.txt</c> / <c>{filename}_timestamp.txt</c>
/// pair, stored in a directory structure that mirrors the source folder hierarchy.
/// </summary>
public class IndexService
{
    private const string IndexSuffix = "_index.txt";
    private const string TimestampSuffix = "_timestamp.txt";

    /// <summary>
    /// Constructs the root index folder path for a given target search folder.
    /// The drive letter colon is removed so the path can be nested under <paramref name="basePath"/>.
    /// </summary>
    private static string GetIndexFolderPath(string basePath, string targetFolder)
    {
        var targetUri = new Uri(targetFolder);
        var relativePath = targetUri.LocalPath.TrimStart(Path.DirectorySeparatorChar);

        if (relativePath.Length >= 2 && relativePath[1] == ':')
            relativePath = relativePath[0] + relativePath.Substring(2);

        return Path.Combine(basePath, relativePath);
    }

    /// <summary>
    /// Gets the per-file index path.
    /// The source file's relative path from the target folder is preserved so that
    /// identically named files in different sub-folders never collide.
    /// </summary>
    private static string GetFileIndexPath(string basePath, string targetFolder, string sourceFilePath)
    {
        var indexFolder = GetIndexFolderPath(basePath, targetFolder);
        var relativeFilePath = Path.GetRelativePath(targetFolder, sourceFilePath);
        return Path.Combine(indexFolder, relativeFilePath + IndexSuffix);
    }

    /// <summary>
    /// Gets the per-file timestamp path (same naming convention as the index path).
    /// </summary>
    private static string GetFileTimestampPath(string basePath, string targetFolder, string sourceFilePath)
    {
        var indexFolder = GetIndexFolderPath(basePath, targetFolder);
        var relativeFilePath = Path.GetRelativePath(targetFolder, sourceFilePath);
        return Path.Combine(indexFolder, relativeFilePath + TimestampSuffix);
    }

    /// <summary>
    /// Determines whether the index and timestamp files both exist for a specific source file.
    /// </summary>
    public bool FileIndexExists(string basePath, string targetFolder, string sourceFilePath)
    {
        return File.Exists(GetFileIndexPath(basePath, targetFolder, sourceFilePath)) &&
               File.Exists(GetFileTimestampPath(basePath, targetFolder, sourceFilePath));
    }

    /// <summary>
    /// Loads the stored last-modified timestamp and <c>excelFormula</c> flag for a specific source file.
    /// Returns <c>(null, false)</c> if the timestamp file does not exist or cannot be parsed.
    /// </summary>
    public (DateTime? Timestamp, bool ExcelFormula) LoadFileTimestamp(string basePath, string targetFolder, string sourceFilePath)
    {
        var timestampPath = GetFileTimestampPath(basePath, targetFolder, sourceFilePath);
        if (!File.Exists(timestampPath)) return (null, false);

        var line = File.ReadAllText(timestampPath).Trim();
        var parts = line.Split('|');

        if (!DateTime.TryParse(parts[0], CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var ts))
            return (null, false);

        var excelFormula = parts.Length > 1 && bool.TryParse(parts[1], out var f) && f;
        return (ts, excelFormula);
    }

    /// <summary>
    /// Loads the indexed grep results for a specific source file.
    /// </summary>
    public List<GrepResult> LoadFileIndex(string basePath, string targetFolder, string sourceFilePath)
    {
        var results = new List<GrepResult>();
        var indexPath = GetFileIndexPath(basePath, targetFolder, sourceFilePath);

        if (!File.Exists(indexPath)) return results;

        foreach (var line in File.ReadLines(indexPath))
        {
            var result = GrepResult.FromIndexLine(line);
            if (result != null) results.Add(result);
        }

        return results;
    }

    /// <summary>
    /// Asynchronously saves the index and timestamp for a single source file.
    /// Directories are created automatically if they do not exist.
    /// The <paramref name="excelFormula"/> flag is stored alongside the timestamp
    /// so that a setting change correctly invalidates the cache.
    /// </summary>
    public async Task SaveFileIndexAsync(
        string basePath,
        string targetFolder,
        string sourceFilePath,
        IEnumerable<GrepResult> results,
        DateTime timestamp,
        bool excelFormula)
    {
        var indexPath = GetFileIndexPath(basePath, targetFolder, sourceFilePath);
        var timestampPath = GetFileTimestampPath(basePath, targetFolder, sourceFilePath);

        var dir = Path.GetDirectoryName(indexPath);
        if (dir != null) Directory.CreateDirectory(dir);

        await using (var writer = new StreamWriter(indexPath, false, Encoding.UTF8))
        {
            foreach (var result in results)
                await writer.WriteLineAsync(result.ToIndexLine());
        }

        await File.WriteAllTextAsync(timestampPath, $"{timestamp:O}|{excelFormula}");
    }

    /// <summary>
    /// Deletes all per-file index and timestamp files for the specified target folder
    /// by removing the entire index directory tree.
    /// </summary>
    public void DeleteIndex(string basePath, string targetFolder)
    {
        var indexFolder = GetIndexFolderPath(basePath, targetFolder);

        if (Directory.Exists(indexFolder))
            Directory.Delete(indexFolder, true);
    }
}