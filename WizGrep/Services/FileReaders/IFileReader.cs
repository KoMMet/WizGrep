using System.Collections.Generic;
using WizGrep.Models;

namespace WizGrep.Services.FileReaders;

/// <summary>
/// Contract for file readers that extract searchable text from a specific document format.
/// Each implementation declares the file extensions it handles and returns
/// <see cref="GrepResult"/> items representing individual lines, cells, shapes, etc.
/// </summary>
/// <remarks>
/// Implementations are registered by <see cref="Services.FileReaderService"/> at startup.
/// During a grep operation, the service dispatches each file to the reader whose
/// <see cref="SupportedExtensions"/> match the file's extension.
/// </remarks>
public interface IFileReader
{
    /// <summary>
    /// Lowercase, dot-prefixed file extensions this reader can handle
    /// (e.g., <c>".xlsx"</c>, <c>".docx"</c>).
    /// </summary>
    IEnumerable<string> SupportedExtensions { get; }

    /// <summary>
    /// Reads the file at <paramref name="filePath"/> and extracts all searchable text segments.
    /// </summary>
    /// <param name="filePath">Absolute path to the file to read.</param>
    /// <param name="excelFormula">
    /// When <c>true</c>, Excel readers return the formula text (<c>="=SUM(...)"</c>)
    /// instead of the computed display value. Ignored by non-Excel readers.
    /// </param>
    /// <returns>
    /// Zero or more <see cref="GrepResult"/> items, each representing a line, cell,
    /// shape text, comment, header, footer, or slide note found in the file.
    /// Returns an empty collection if the file cannot be read.
    /// </returns>
    IEnumerable<GrepResult> ReadFile(string filePath, bool excelFormula);
}