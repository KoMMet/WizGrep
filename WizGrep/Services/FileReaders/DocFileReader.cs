using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OpenMcdf;
using WizGrep.Helpers;
using WizGrep.Models;

namespace WizGrep.Services.FileReaders;

/// <summary>
/// Reads legacy Word binary files (.doc) by parsing the OLE2 compound document
/// structure via OpenMcdf and extracting text from the FIB/CLX piece table.
/// </summary>
/// <remarks>
/// The reader supports both compressed (Windows-1252) and uncompressed (UTF-16 LE)
/// text pieces. Field codes (mail merge, IF, SET, etc.) are stripped using a
/// stack-based approach in <see cref="CleanText"/>. Returns an empty collection
/// on any error.
/// </remarks>
public class DocFileReader : IFileReader
{
    /// <summary>Windows-1252 encoding used for compressed text pieces in .doc files.</summary>
    private static readonly Encoding Windows1252;

    static DocFileReader()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Windows1252 = Encoding.GetEncoding(1252);
    }

    public IEnumerable<string> SupportedExtensions => [".doc"];

    public IEnumerable<GrepResult> ReadFile(string filePath, bool excelFormula)
    {
        var results = new List<GrepResult>();

        try
        {
            using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var memStream = new MemoryStream();
            fs.CopyTo(memStream);
            memStream.Position = 0;
            using var root = RootStorage.Open(memStream);

            using var wordDocStream = root.OpenStream("WordDocument");
            var wordDocBytes = ReadAllBytes(wordDocStream);

            if (wordDocBytes.Length < 0x01AA)
                return results;

            // Validate the Word magic number (wIdent = 0xA5EC)
            var wIdent = BitConverter.ToUInt16(wordDocBytes, 0);
            if (wIdent != 0xA5EC)
                return results;

            // Read FIB (File Information Block) fields needed for text extraction
            var flags = BitConverter.ToUInt16(wordDocBytes, 0x000A);
            var fWhichTblStm = (flags & 0x0200) != 0;
            var ccpText = BitConverter.ToUInt32(wordDocBytes, 0x004C);
            var fcClx = BitConverter.ToUInt32(wordDocBytes, 0x01A2);
            var lcbClx = BitConverter.ToUInt32(wordDocBytes, 0x01A6);

            if (lcbClx == 0 || ccpText == 0)
                return results;

            // Open the correct table stream (0Table or 1Table) based on the FIB flag
            var tableName = fWhichTblStm ? "1Table" : "0Table";
            using var tableStream = root.OpenStream(tableName);
            var tableBytes = ReadAllBytes(tableStream);

            // Parse the CLX structure to extract the piece table and reconstruct the text
            var text = ParseClx(wordDocBytes, tableBytes, fcClx, lcbClx, ccpText);

            // Split text into lines and return as results
            var lines = text.Split(['\r', '\n'], StringSplitOptions.None);
            var lineNumber = 1;
            foreach (var line in lines)
            {
                // Remove field codes and control characters from each line
                var cleaned = CleanText(line);
                if (!string.IsNullOrWhiteSpace(cleaned))
                {
                    results.Add(new GrepResult
                    {
                        FilePath = filePath,
                        LineNumber = lineNumber,
                        Content = cleaned
                    });
                }

                lineNumber++;
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading .doc file '{filePath}' with excelFormula={excelFormula}: {e.StackTrace}");
        }

        return results;
    }

    /// <summary>
    /// Parses the CLX (Complex) structure in the table stream, locates the Pcdt entry,
    /// and delegates to <see cref="ParsePlcPcd"/> to reconstruct the document text.
    /// </summary>
    private string ParseClx(byte[] wordDocBytes, byte[] tableBytes, uint fcClx, uint lcbClx, uint ccpText)
    {
        if (fcClx + lcbClx > tableBytes.Length)
            return string.Empty;

        var offset = fcClx;
        var endOffset = fcClx + lcbClx;

        // Skip Pcr entries (0x01) until we find the Pcdt marker (0x02)
        while (offset < endOffset)
        {
            var clxt = tableBytes[offset];
            if (clxt == 0x01) // Pcr
            {
                var cbGrpprl = BitConverter.ToUInt16(tableBytes, (int)(offset + 1));
                offset += 3u + cbGrpprl;
            }
            else if (clxt == 0x02) // Pcdt
            {
                offset++;
                var lcb = BitConverter.ToUInt32(tableBytes, (int)offset);
                offset += 4;
                return ParsePlcPcd(wordDocBytes, tableBytes, offset, lcb, ccpText);
            }
            else
            {
                break;
            }
        }

        return string.Empty;
    }

    /// <summary>
    /// Reads the PlcPcd (Piece Descriptor Table) and reconstructs the document text
    /// by concatenating each text piece, handling both compressed (Windows-1252)
    /// and uncompressed (UTF-16 LE) pieces.
    /// </summary>
    private string ParsePlcPcd(byte[] wordDocBytes, byte[] tableBytes, uint plcPcdOffset, uint lcb, uint ccpText)
    {
        // ピース数: (n+1)*4 + n*8 = lcb → n = (lcb - 4) / 12
        var n = (int)(lcb - 4) / 12;
        if (n <= 0) return string.Empty;

        var sb = new StringBuilder();

        for (var i = 0; i < n; i++)
        {
            var cpStart = BitConverter.ToUInt32(tableBytes, (int)(plcPcdOffset + i * 4));
            var cpEnd = BitConverter.ToUInt32(tableBytes, (int)(plcPcdOffset + (i + 1) * 4));

            // Only extract characters within the main document text range
            if (cpStart >= ccpText) break;
            if (cpEnd > ccpText) cpEnd = ccpText;

            var charCount = cpEnd - cpStart;
            if (charCount == 0) continue;

            // Read PCD (Piece Descriptor) to get the file offset and compression flag
            var pcdOffset = plcPcdOffset + (uint)(n + 1) * 4 + (uint)i * 8;
            var fc = BitConverter.ToUInt32(tableBytes, (int)(pcdOffset + 2));

            var isCompressed = (fc & 0x40000000) != 0;

            if (isCompressed)
            {
                // Compressed text piece – single-byte Windows-1252 encoding
                var realOffset = (fc & 0x3FFFFFFFU) / 2;
                if (realOffset + charCount <= wordDocBytes.Length)
                    sb.Append(Windows1252.GetString(wordDocBytes, (int)realOffset, (int)charCount));
            }
            else
            {
                // Uncompressed text piece – UTF-16 LE encoding
                var realOffset = fc & 0x3FFFFFFFU;
                var byteCount = charCount * 2;
                if (realOffset + byteCount <= wordDocBytes.Length)
                    sb.Append(Encoding.Unicode.GetString(wordDocBytes, (int)realOffset, (int)byteCount));
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// Removes control characters and Word field codes from extracted text.
    /// Handles nested fields (IF, SET, mail merge, etc.) using a stack to track
    /// the visible/hidden state across field begin/separator/end markers.
    /// </summary>
    private static string CleanText(string text)
    {
        var sb = new StringBuilder(text.Length);

        // Stack tracks the visibility state before each nested field was entered
        var isVisible = true;
        var fieldStack = new Stack<bool>();

        foreach (var c in text)
        {
            switch (c)
            {
                case '\x13': // Field begin – entering the field code (hidden) region
                    fieldStack.Push(isVisible);
                    isVisible = false;
                    break;
                case '\x14': // Field separator – beginning of the display-text region
                    // Restore visibility only if the parent context was visible
                    isVisible = fieldStack.Count > 0 && fieldStack.Peek();
                    break;
                case '\x15': // Field end – restore the parent's visibility state
                    isVisible = fieldStack.Count <= 0 || fieldStack.Pop();
                    break;
                case '\x07': // Table cell / row separator – render as a tab
                    if (isVisible)
                        sb.Append('\t');
                    break;
                default:
                    if (isVisible && !char.IsControl(c))
                        sb.Append(c);
                    break;
            }
        }

        return sb.ToString();
    }

    private static byte[] ReadAllBytes(CfbStream stream)
    {
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }
}
