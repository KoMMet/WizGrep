using Microsoft.Windows.ApplicationModel.Resources;
using System.IO;
using System.Text;
using WizGrep.Helpers;

namespace WizGrep.Models;

/// <summary>
/// Represents a single match result from a grep operation.
/// Contains the source file path, location metadata (line number, sheet name,
/// cell address, shape/object name), and the matched text content.
/// Provides serialization/deserialization to a tab-delimited index format
/// used by <see cref="Services.IndexService"/> for persistent caching.
/// </summary>
public class GrepResult
{
    /// <summary>Absolute path of the file that contains the match.</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>File name with extension, derived from <see cref="FilePath"/>.</summary>
    public string FileName => Path.GetFileName(FilePath);

    /// <summary>
    /// 1-based line number within the file or sheet. Set to 0 for non-line-based
    /// matches (e.g., shapes, comments, headers/footers).
    /// </summary>
    public int LineNumber { get; set; }

    /// <summary>
    /// Excel sheet name, PDF page label, or PowerPoint slide label.
    /// <c>null</c> for file types that have no sheet/page concept (e.g., plain text, Word body).
    /// </summary>
    public string? SheetName { get; set; }

    /// <summary>
    /// Excel cell reference (e.g., "A1"). <c>null</c> for non-Excel results.
    /// </summary>
    public string? CellAddress { get; set; }

    /// <summary>
    /// Descriptive name of the container object, such as "Comment", "Shape1",
    /// "Header1", "Note", etc. <c>null</c> for regular body/cell text.
    /// </summary>
    public string? ObjectName { get; set; }

    /// <summary>The matched text content (one line or one cell value).</summary>
    public string Content { get; set; } = string.Empty;

    /// <summary>
    /// Human-readable location string shown in the results grid.
    /// Composed of <see cref="SheetName"/>, <see cref="CellAddress"/>,
    /// <see cref="ObjectName"/>, and <see cref="LineNumber"/> depending on which
    /// values are available.
    /// </summary>
    public string Location
    {
        get
        {
            if (!string.IsNullOrEmpty(SheetName) && !string.IsNullOrEmpty(CellAddress) && !string.IsNullOrEmpty(ObjectName))
                return $"[{SheetName}] {CellAddress} {ObjectName}";
            if (!string.IsNullOrEmpty(SheetName) && !string.IsNullOrEmpty(CellAddress))
                return $"[{SheetName}] {CellAddress}";
            if (!string.IsNullOrEmpty(SheetName) && !string.IsNullOrEmpty(ObjectName))
                return $"[{SheetName}] {ObjectName}";
            if (!string.IsNullOrEmpty(SheetName))
                return $"[{SheetName}] {LineNumber} {ResourceLoaderHelper.GetString("RowLabel")}";
            if (!string.IsNullOrEmpty(ObjectName)) return $"[{ResourceLoaderHelper.GetString("ObjectLabel")}] {ObjectName}";
            return $"{LineNumber} {ResourceLoaderHelper.GetString("RowLabel")}";
        }
    }

    /// <summary>
    /// Serializes this result to a single tab-delimited line for index file storage.
    /// Format: <c>FilePath\tLineNumber\tSheetName\tCellAddress\tObjectName\tEscapedContent</c>.
    /// </summary>
    public string ToIndexLine()
    {
        return $"{FilePath}\t{LineNumber}\t{SheetName ?? ""}\t{CellAddress ?? ""}\t{ObjectName ?? ""}\t{EscapeForIndex(Content)}";
    }

    /// <summary>
    /// Deserializes a tab-delimited index line (produced by <see cref="ToIndexLine"/>)
    /// back into a <see cref="GrepResult"/>.
    /// </summary>
    /// <param name="line">A tab-separated string with at least 6 fields.</param>
    /// <returns>A populated <see cref="GrepResult"/>, or <c>null</c> if the line is malformed.</returns>
    public static GrepResult? FromIndexLine(string line)
    {
        var parts = line.Split('\t');
        if (parts.Length < 6) return null;

        return new GrepResult
        {
            FilePath = parts[0],
            LineNumber = int.TryParse(parts[1], out var ln) ? ln : 0,
            SheetName = string.IsNullOrEmpty(parts[2]) ? null : parts[2],
            CellAddress = string.IsNullOrEmpty(parts[3]) ? null : parts[3],
            ObjectName = string.IsNullOrEmpty(parts[4]) ? null : parts[4],
            Content = UnescapeFromIndex(string.Join("\t", parts, 5, parts.Length - 5))
        };
    }

    /// <summary>
    /// Escapes backslash, CR, LF, and tab characters so that the content field
    /// can be safely stored in a tab-delimited index line.
    /// </summary>
    private static string EscapeForIndex(string value)
    {
        return value
            .Replace("\\", "\\\\")
            .Replace("\r", "\\r")
            .Replace("\n", "\\n")
            .Replace("\t", "\\t");
    }

    /// <summary>
    /// Reverses the escaping applied by <see cref="EscapeForIndex"/>, restoring
    /// the original content string from an index line.
    /// </summary>
    private static string UnescapeFromIndex(string value)
    {
        var sb = new StringBuilder(value.Length);
        for (var i = 0; i < value.Length; i++)
        {
            if (value[i] == '\\' && i + 1 < value.Length)
            {
                switch (value[i + 1])
                {
                    case '\\':
                        sb.Append('\\');
                        i++;
                        break;
                    case 'r':
                        sb.Append('\r');
                        i++;
                        break;
                    case 'n':
                        sb.Append('\n');
                        i++;
                        break;
                    case 't':
                        sb.Append('\t');
                        i++;
                        break;
                    default:
                        sb.Append(value[i]);
                        break;
                }
            }
            else
            {
                sb.Append(value[i]);
            }
        }

        return sb.ToString();
    }
}