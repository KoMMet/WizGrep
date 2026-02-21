using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using WizGrep.Models;

namespace WizGrep.Services;

/// <summary>
/// Provides methods for managing and accessing index and timestamp files associated with a target folder.
/// </summary>
/// <remarks>The IndexService class enables the retrieval, loading, saving, and deletion of index and timestamp
/// data for specified folders. It ensures that index and timestamp files are stored in a structured manner within a
/// given base path, supporting efficient file management and access. This service is intended for scenarios where
/// maintaining up-to-date file indices and modification timestamps is required, such as search or synchronization
/// operations.</remarks>
public class IndexService
{
    private const string IndexFileName = "GrepIndex.txt";
    private const string TimestampFileName = "GrepTimestamp.txt";

    /// <summary>
    /// Constructs the index folder path by combining a base path with the relative path of the specified target folder.
    /// </summary>
    /// <remarks>The method extracts the drive letter from the target folder path and constructs a new path by
    /// removing the drive letter's colon. It is important to ensure that the target folder is a valid URI to avoid
    /// unexpected behavior.</remarks>
    /// <param name="basePath">The base path to which the relative path of the target folder will be appended. This path must be a valid
    /// directory path.</param>
    /// <param name="targetFolder">The full path of the target folder from which the relative path will be derived. This must be a valid URI
    /// format.</param>
    /// <returns>A string representing the combined index folder path, constructed from the base path and the modified relative
    /// path of the target folder.</returns>
    private string GetIndexFolderPath(string basePath, string targetFolder)
    {
        // Remove the colon from the drive letter in the relative path
        var targetUri = new Uri(targetFolder);
        var relativePath = targetUri.LocalPath.TrimStart(Path.DirectorySeparatorChar);

        // Adjust the relative path by removing the colon after the drive letter if present.
        if (relativePath.Length >= 2 && relativePath[1] == ':')
            relativePath = relativePath[0] + relativePath.Substring(2);

        return Path.Combine(basePath, relativePath);
    }

    /// <summary>
    /// Generates the full file path to the index file located within the specified target folder under the given base
    /// path.
    /// </summary>
    /// <remarks>The index file name is determined by the constant value 'IndexFileName'. The method combines
    /// the base path, target folder, and index file name to construct the full path.</remarks>
    /// <param name="basePath">The root directory in which the target folder containing the index file is located. This value must not be null
    /// or empty.</param>
    /// <param name="targetFolder">The name of the folder within the base path where the index file resides. This value must not be null or empty.</param>
    /// <returns>A string containing the absolute path to the index file within the specified target folder.</returns>
    public string GetIndexFilePath(string basePath, string targetFolder)
    {
        return Path.Combine(GetIndexFolderPath(basePath, targetFolder), IndexFileName);
    }

    /// <summary>
    /// Gets the full file path to the timestamp file within the specified target folder.
    /// </summary>
    /// <remarks>The method does not verify the existence of the target folder. If the folder does not exist,
    /// the returned path may not correspond to an existing file.</remarks>
    /// <param name="basePath">The base directory that contains the target folder.</param>
    /// <param name="targetFolder">The name of the folder within the base path where the timestamp file is located.</param>
    /// <returns>A string containing the full path to the timestamp file in the specified target folder.</returns>
    public string GetTimestampFilePath(string basePath, string targetFolder)
    {
        return Path.Combine(GetIndexFolderPath(basePath, targetFolder), TimestampFileName);
    }

    /// <summary>
    /// Determines whether both the index file and its associated timestamp file exist for the specified folder.
    /// </summary>
    /// <remarks>Both the index file and its corresponding timestamp file must be present for this method to
    /// return true. This check is useful for verifying the integrity or readiness of an index before performing further
    /// operations.</remarks>
    /// <param name="basePath">The base directory path where the index and timestamp files are expected to be located. This path must be valid
    /// and accessible.</param>
    /// <param name="targetFolder">The name of the target folder for which the existence of the index and timestamp files is being checked.</param>
    /// <returns>true if both the index file and the timestamp file exist for the specified folder; otherwise, false.</returns>
    public bool IndexExists(string basePath, string targetFolder)
    {
        return File.Exists(GetIndexFilePath(basePath, targetFolder)) &&
               File.Exists(GetTimestampFilePath(basePath, targetFolder));
    }

    /// <summary>
    /// Loads the last modified timestamps of files from a timestamp file located in the specified directory.
    /// </summary>
    /// <remarks>The method reads a timestamp file line by line, extracting file paths and their last modified
    /// dates. It is important to ensure that the timestamp file is formatted correctly for successful
    /// parsing.</remarks>
    /// <param name="basePath">The base directory path where the timestamp file is located. Cannot be null or empty.</param>
    /// <param name="targetFolder">The name of the target folder used to construct the full path to the timestamp file. Cannot be null or empty.</param>
    /// <returns>A dictionary containing file paths as keys and their corresponding last modified timestamps as values. The
    /// dictionary is empty if the timestamp file does not exist or contains no valid entries.</returns>
    public Dictionary<string, DateTime> LoadTimestamps(string basePath, string targetFolder)
    {
        var timestamps = new Dictionary<string, DateTime>();
        var timestampFilePath = GetTimestampFilePath(basePath, targetFolder);

        if (!File.Exists(timestampFilePath)) return timestamps;

        foreach (var line in File.ReadLines(timestampFilePath))
        {
            var fileTimestamp = FileTimestamp.FromTimestampLine(line);
            if (fileTimestamp != null) timestamps[fileTimestamp.FilePath] = fileTimestamp.LastModified;
        }

        return timestamps;
    }

    /// <summary>
    /// Loads and parses the index file from the specified target folder within the given base path, returning a list of
    /// grep results.
    /// </summary>
    /// <remarks>If the index file is not found at the constructed path, the method returns an empty list.
    /// Each line in the index file is parsed into a GrepResult instance if possible.</remarks>
    /// <param name="basePath">The base directory that contains the target folder. This path is used to locate the index file.</param>
    /// <param name="targetFolder">The name of the folder within the base path from which to load the index file.</param>
    /// <returns>A list of GrepResult objects representing the entries parsed from the index file. Returns an empty list if the
    /// index file does not exist or contains no valid entries.</returns>
    public List<GrepResult> LoadIndex(string basePath, string targetFolder)
    {
        var results = new List<GrepResult>();
        var indexFilePath = GetIndexFilePath(basePath, targetFolder);

        if (!File.Exists(indexFilePath)) return results;

        foreach (var line in File.ReadLines(indexFilePath))
        {
            var result = GrepResult.FromIndexLine(line);
            if (result != null) results.Add(result);
        }

        return results;
    }

    /// <summary>
    /// Asynchronously saves the index and timestamp files to the specified target folder.
    /// </summary>
    /// <remarks>This method creates the necessary directories and files if they do not already exist. It
    /// overwrites any existing index and timestamp files.</remarks>
    /// <param name="basePath">The base path where the index and timestamp files will be saved.</param>
    /// <param name="targetFolder">The target folder within the base path where the index and timestamp files will be created.</param>
    /// <param name="results">An enumerable collection of GrepResult objects representing the results to be written to the index file.</param>
    /// <param name="timestamps">A dictionary mapping file paths to their last modified timestamps, used to create the timestamp file.</param>
    /// <returns>A Task representing the asynchronous operation of saving the index and timestamp files.</returns>
    public async Task SaveIndexAsync(
        string basePath,
        string targetFolder,
        IEnumerable<GrepResult> results,
        Dictionary<string, DateTime> timestamps)
    {
        var indexFolder = GetIndexFolderPath(basePath, targetFolder);
        Directory.CreateDirectory(indexFolder);

        var indexFilePath = GetIndexFilePath(basePath, targetFolder);
        var timestampFilePath = GetTimestampFilePath(basePath, targetFolder);

        // Save the index file
        await using (var writer = new StreamWriter(indexFilePath, false, Encoding.UTF8))
        {
            foreach (var result in results) await writer.WriteLineAsync(result.ToIndexLine());
        }

        // Save the timestamp file asynchronously
        await using (var writer = new StreamWriter(timestampFilePath, false, Encoding.UTF8))
        {
            foreach (var kvp in timestamps)
            {
                var fileTimestamp = new FileTimestamp
                {
                    FilePath = kvp.Key,
                    LastModified = kvp.Value
                };
                await writer.WriteLineAsync(fileTimestamp.ToTimestampLine());
            }
        }
    }

    /// <summary>
    /// Deletes the index and timestamp files associated with the specified base path and target folder, if they exist.
    /// </summary>
    /// <remarks>If the specified files do not exist, no action is taken. This method does not throw
    /// exceptions for missing files.</remarks>
    /// <param name="basePath">The base directory that contains the index and timestamp files to delete. This path must be valid and
    /// accessible.</param>
    /// <param name="targetFolder">The name of the target folder within the base path that identifies the location of the files to be deleted.</param>
    public void DeleteIndex(string basePath, string targetFolder)
    {
        var indexFilePath = GetIndexFilePath(basePath, targetFolder);
        var timestampFilePath = GetTimestampFilePath(basePath, targetFolder);

        if (File.Exists(indexFilePath)) File.Delete(indexFilePath);

        if (File.Exists(timestampFilePath)) File.Delete(timestampFilePath);
    }
}