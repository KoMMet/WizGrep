using System.Collections.Generic;
using System.IO;
using System.Linq;
using WizGrep.Models;
using WizGrep.Services.FileReaders;

namespace WizGrep.Services;

/// <summary>
/// Central dispatcher that routes file reads to the appropriate <see cref="IFileReader"/>
/// implementation based on the file's extension.
/// Falls back to <see cref="TextFileReader"/> for unrecognized extensions.
/// </summary>
public class FileReaderService
{
    /// <summary>Extension Å® reader mapping (keys are lowercase, dot-prefixed).</summary>
    private readonly Dictionary<string, IFileReader> _readers = new();

    /// <summary>Fallback reader used when no registered reader matches the file extension.</summary>
    private readonly IFileReader _defaultReader = new TextFileReader();

    /// <summary>
    /// Registers all built-in file readers:
    /// plain text, Word (.docx/.docm), Excel (.xlsx/.xlsm), PowerPoint (.pptx/.pptm),
    /// PDF, and legacy binary formats (.xls via NPOI, .doc/.ppt via OpenMcdf).
    /// </summary>
    public FileReaderService()
    {
        RegisterReader(_defaultReader);
        RegisterReader(new WordFileReader());
        RegisterReader(new ExcelFileReader());
        RegisterReader(new PowerPointFileReader());
        RegisterReader(new PdfFileReader());

        // Legacy binary formats
        RegisterReader(new NpoiExcelFileReader()); // .xls (NPOI)
        RegisterReader(new DocFileReader());        // .doc (OpenMcdf)
        RegisterReader(new PptFileReader());        // .ppt (OpenMcdf)
    }

    /// <summary>
    /// Maps each of the reader's supported extensions to the reader instance.
    /// Later registrations overwrite earlier ones for the same extension.
    /// </summary>
    private void RegisterReader(IFileReader reader)
    {
        foreach (var ext in reader.SupportedExtensions) 
            _readers[ext.ToLowerInvariant()] = reader;
    }

    /// <summary>
    /// Reads the file and returns searchable text segments.
    /// Dispatches to the reader registered for the file's extension,
    /// or to the default text reader if none matches.
    /// </summary>
    public IEnumerable<GrepResult> ReadFile(string filePath, bool excelFormula)
    {
        var extension = Path.GetExtension(filePath).ToLowerInvariant();

        if (_readers.TryGetValue(extension, out var reader)) 
            return reader.ReadFile(filePath, excelFormula);

        // Unrecognized extension ? treat as plain text
        return _defaultReader.ReadFile(filePath, excelFormula);
    }
}