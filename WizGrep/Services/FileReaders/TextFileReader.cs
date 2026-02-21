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
            var lines = File.ReadAllLines(filePath, encoding);

            for (var i = 0; i < lines.Length; i++)
                results.Add(new GrepResult
                {
                    FilePath = filePath,
                    LineNumber = i + 1,
                    Content = lines[i]
                });
        }
        catch (Exception)
        {
            // Silently ignore file read errors (e.g., access denied, encoding issues)
        }

        return results;
    }
}