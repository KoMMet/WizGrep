using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OpenMcdf;
using WizGrep.Helpers;
using WizGrep.Models;

namespace WizGrep.Services.FileReaders;

/// <summary>
/// Reads legacy PowerPoint binary files (.ppt) by parsing the OLE2 compound
/// document structure via OpenMcdf and extracting text atom records.
/// </summary>
/// <remarks>
/// Supports both UTF-16 LE text atoms (<see cref="TextCharsAtomType"/>) and
/// single-byte Windows-1252 text atoms (<see cref="TextBytesAtomType"/>).
/// The SlideListWithTextContainer is intentionally skipped to avoid duplicating
/// text that already appears inside individual SlideContainer/NotesContainer records.
/// </remarks>
public class PptFileReader : IFileReader
{
    /// <summary>Record type for TextCharsAtom – contains UTF-16 LE encoded text.</summary>
    private const ushort TextCharsAtomType = 0x0FA0;

    /// <summary>Record type for TextBytesAtom – contains single-byte Windows-1252 text.</summary>
    private const ushort TextBytesAtomType = 0x0FA8;

    /// <summary>Record type for SlideContainer – marks the boundary of a slide.</summary>
    private const ushort SlideContainerType = 0x03EE;

    /// <summary>Record type for NotesContainer – marks the boundary of a notes page.</summary>
    private const ushort NotesContainerType = 0x03F0;

    /// <summary>
    /// Record type for SlideListWithTextContainer – a document-level container that
    /// duplicates text already found inside individual slide/notes containers; skipped
    /// during extraction to avoid double-counting.
    /// </summary>
    private const ushort SlideListWithTextContainerType = 0x0FF0;

    /// <summary>Windows-1252 encoding for single-byte text atoms.</summary>
    private static readonly Encoding Windows1252;
    
    /// <summary>
    /// Registers the code-pages encoding provider so that Windows-1252 is available
    /// for decoding single-byte text atoms.
    /// </summary>
    static PptFileReader()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Windows1252 = Encoding.GetEncoding(1252);
    }

    public IEnumerable<string> SupportedExtensions => [".ppt"];

    /// <summary>
    /// Reads the .ppt file and extracts text from all slides and notes.
    /// Returns an empty collection on any error.
    /// </summary>
    public IEnumerable<GrepResult> ReadFile(string filePath, bool excelFormula)
    {
        var results = new List<GrepResult>();

        try
        {
            using var root = RootStorage.OpenRead(filePath);
            using var pptStream = root.OpenStream("PowerPoint Document");

            using var ms = new MemoryStream();
            pptStream.CopyTo(ms);
            var data = ms.ToArray();

            var slideCounter = 0;
            ExtractTextRecords(data, 0, data.Length, ref slideCounter, false, filePath, results);
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading PowerPoint file '{filePath}' with excelFormula={excelFormula}: {e.StackTrace}");
        }

        return results;
    }

    /// <summary>
    /// Recursively walks the binary record tree within <paramref name="data"/> and
    /// extracts text from TextCharsAtom and TextBytesAtom records.
    /// </summary>
    /// <param name="data">The raw bytes of the "PowerPoint Document" stream.</param>
    /// <param name="offset">Start offset within <paramref name="data"/>.</param>
    /// <param name="endOffset">Exclusive end offset.</param>
    /// <param name="slideCounter">Running slide counter, incremented on each SlideContainer.</param>
    /// <param name="inNotes"><c>true</c> if currently inside a NotesContainer.</param>
    /// <param name="filePath">Source file path for populating results.</param>
    /// <param name="results">Accumulator list for extracted results.</param>
    private void ExtractTextRecords(
        byte[] data, int offset, int endOffset,
        ref int slideCounter, bool inNotes,
        string filePath, List<GrepResult> results)
    {
        while (offset + 8 <= endOffset)
        {
            var recVerInstance = BitConverter.ToUInt16(data, offset);
            var recType = BitConverter.ToUInt16(data, offset + 2);
            var recLen = BitConverter.ToUInt32(data, offset + 4);

            var recVer = recVerInstance & 0x0F;
            var dataStart = offset + 8;

            // Abort if the record extends beyond the buffer boundary
            if (dataStart + (long)recLen > endOffset)
                break;

            if (recVer == 0x0F) // Container record – recurse into children
            {
                // Skip SlideListWithTextContainer to avoid duplicate text
                // (individual SlideContainer/NotesContainer already contain the same text)
                if (recType == SlideListWithTextContainerType)
                {
                    offset = dataStart + (int)recLen;
                    continue;
                }

                var childInNotes = inNotes;

                if (recType == SlideContainerType)
                    slideCounter++;
                else if (recType == NotesContainerType)
                    childInNotes = true;

                ExtractTextRecords(data, dataStart, dataStart + (int)recLen,
                    ref slideCounter, childInNotes, filePath, results);
            }
            else // Atom record – extract text if it is a text atom type
            {
                if (recType == TextCharsAtomType && recLen >= 2)
                {
                    var text = Encoding.Unicode.GetString(data, dataStart, (int)recLen);
                    AddTextResults(text, slideCounter, inNotes, filePath, results);
                }
                else if (recType == TextBytesAtomType && recLen >= 1)
                {
                    var text = Windows1252.GetString(data, dataStart, (int)recLen);
                    AddTextResults(text, slideCounter, inNotes, filePath, results);
                }
            }

            offset = dataStart + (int)recLen;
        }
    }

    /// <summary>
    /// Splits the extracted text into lines (PowerPoint uses <c>\r</c> as the line break)
    /// and appends each non-empty line to <paramref name="results"/> with appropriate
    /// slide/note metadata.
    /// </summary>
    private static void AddTextResults(
        string text, int slideCounter, bool inNotes,
        string filePath, List<GrepResult> results)
    {
        if (string.IsNullOrWhiteSpace(text)) return;

        var slideName = slideCounter > 0 ? $"{ResourceLoaderHelper.GetString("SlideLabel")}{slideCounter}" : $"{ResourceLoaderHelper.GetString("DocumentInfoLabel")}";
        var objectName = inNotes ? $"{ResourceLoaderHelper.GetString("NoteLabel")}" : null;

        // PowerPoint binary format uses \r as the line break character
        var lines = text.Split('\r');
        var lineNumber = 1;
        foreach (var line in lines)
        {
            var trimmedLine = line.Trim('\n');
            if (!string.IsNullOrWhiteSpace(trimmedLine))
            {
                results.Add(new GrepResult
                {
                    FilePath = filePath,
                    LineNumber = lineNumber,
                    SheetName = slideName,
                    ObjectName = objectName,
                    Content = trimmedLine
                });
            }

            lineNumber++;
        }
    }
}
