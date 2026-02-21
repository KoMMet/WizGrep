using System;
using System.Globalization;

namespace WizGrep.Models;

/// <summary>
/// Associates a file's absolute path with its last-modified timestamp.
/// Used by <see cref="Services.IndexService"/> to track whether indexed files
/// have changed since the index was last built, enabling incremental re-indexing.
/// </summary>
public class FileTimestamp
{
    /// <summary>The absolute file path.</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>The file's last-modified date/time (from <see cref="System.IO.FileInfo.LastWriteTime"/>).</summary>
    public DateTime LastModified { get; set; }

    /// <summary>
    /// Serializes this instance to a pipe-delimited line: <c>FilePath|LastModified</c>.
    /// The timestamp uses the round-trip ("O") format specifier for lossless storage.
    /// </summary>
    public string ToTimestampLine()
    {
        return $"{FilePath}|{LastModified:O}";
    }

    /// <summary>
    /// Parses a pipe-delimited line (<c>FilePath|Timestamp</c>) and creates a
    /// <see cref="FileTimestamp"/> instance. Returns <c>null</c> if the format is invalid.
    /// </summary>
    /// <param name="line">A line previously produced by <see cref="ToTimestampLine"/>.</param>
    public static FileTimestamp? FromTimestampLine(string line)
    {
        var parts = line.Split('|');
        if (parts.Length < 2) return null;

        if (!DateTime.TryParse(parts[1], CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var lastModified))
            return null;

        return new FileTimestamp
        {
            FilePath = parts[0],
            LastModified = lastModified
        };
    }
}