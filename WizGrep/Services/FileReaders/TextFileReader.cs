using System;
using System.Collections.Generic;
using System.IO;
using WizGrep.Helpers;
using WizGrep.Models;

namespace WizGrep.Services.FileReaders;

/// <summary>
/// Reads plain-text files (.txt) using auto-detected encoding (via <see cref="EncodingDetectorHelper"/>)
/// and returns one <see cref="GrepResult"/> per line.
/// Also serves as the default/fallback reader for any unrecognized file extension.
/// </summary>
public class TextFileReader : IFileReader
{
    public IEnumerable<string> SupportedExtensions => [".txt"];

    /// <summary>
    /// Reads all lines from the file, detecting encoding automatically.
    /// Returns an empty collection on any read error (file not found, access denied, etc.).
    /// </summary>
    public IEnumerable<GrepResult> ReadFile(string filePath, bool excelFormula)
    {
        var results = new List<GrepResult>();

        try
        {
            var encoding = EncodingDetectorHelper.DetectEncoding(filePath);
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(stream, encoding);

            var lineNumber = 1;
            while (reader.ReadLine() is { } line)
            {
                results.Add(new GrepResult
                {
                    FilePath = filePath,
                    LineNumber = lineNumber,
                    Content = line
                });
                lineNumber++;
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading text file '{filePath}' with excelFormula={excelFormula}: {e.StackTrace}");
        }

        return results;
    }
}