using System;
using System.IO;
using System.Text;

namespace WizGrep.Helpers;

/// <summary>
/// Utility class that detects the text encoding of a file by analyzing its byte content.
/// </summary>
/// <remarks>
/// Supports BOM (Byte Order Mark) detection for UTF-8, UTF-16 (LE/BE), and UTF-32 (LE/BE).
/// When no BOM is present, it performs strict decode validation for both UTF-8 and Shift_JIS.
/// If both encodings decode successfully, the encoding with the higher Japanese character count
/// (Hiragana, Katakana, CJK Ideographs, full-width forms) is chosen; otherwise UTF-8 is the default.
/// This class is used by <see cref="Services.FileReaders.TextFileReader"/> to determine the
/// correct encoding before reading plain-text files.
/// </remarks>
public class EncodingDetectorHelper
{
    /// <summary>
    /// Cached Shift_JIS encoding instance used for non-strict decoding and as the return value
    /// when content is identified as Shift_JIS.
    /// </summary>
    private static readonly Encoding ShiftJisEncoding = Encoding.GetEncoding("Shift_JIS");

    /// <summary>
    /// Strict UTF-8 encoding instance that throws <see cref="DecoderFallbackException"/> on
    /// invalid byte sequences. Created in <see cref="DetectEncodingFromBytes"/> to validate
    /// whether the data is valid UTF-8.
    /// </summary>
    private static UTF8Encoding? StrictUtf8;

    /// <summary>
    /// Strict Shift_JIS encoding instance that throws <see cref="DecoderFallbackException"/> on
    /// invalid byte sequences. Created in <see cref="DetectEncodingFromBytes"/> to validate
    /// whether the data is valid Shift_JIS.
    /// </summary>
    private static Encoding? StrictShiftJis;

    /// <summary>
    /// Maximum byte length of a single multi-byte character (4 for UTF-8).
    /// When the read buffer is truncated, this many trailing bytes are excluded from
    /// strict decode validation to avoid false negatives caused by a partially-read character.
    /// Shift_JIS uses at most 2 bytes per character, so 4 bytes is sufficient for both encodings.
    /// </summary>
    private const int MaxMultiByteLength = 4;
    
    /// <summary>
    /// Detects the text encoding of a file by reading the first 4 KB of its content.
    /// </summary>
    /// <param name="filePath">The absolute or relative path to the file to analyze.</param>
    /// <returns>
    /// The detected <see cref="Encoding"/>. Returns UTF-8 as the default when the encoding
    /// cannot be determined with certainty.
    /// </returns>
    /// <remarks>
    /// A 4096-byte buffer is read from the beginning of the file. If the file is larger than
    /// the buffer, the data is marked as truncated so that <see cref="DetectEncodingFromBytes"/>
    /// can trim trailing bytes to avoid multi-byte boundary issues.
    /// </remarks>
    /// <exception cref="IOException">Thrown when an I/O error occurs while reading the file.</exception>
    /// <exception cref="UnauthorizedAccessException">Thrown when access to the file is denied.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="filePath"/> is null or invalid.</exception>
    public static Encoding DetectEncoding(string filePath)
    {
        // Read the first 4 KB ? enough for BOM detection and statistical analysis
        const int bufferSize = 4096;
        var buffer = new byte[bufferSize];

        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var bytesRead = stream.Read(buffer, 0, buffer.Length);

        var isTruncated = bytesRead == bufferSize && stream.Position < stream.Length;
        return DetectEncodingFromBytes(buffer.AsSpan(0, bytesRead), isTruncated);
    }

    /// <summary>
    /// Detects the text encoding from a raw byte sequence.
    /// </summary>
    /// <param name="data">The byte sequence to analyze.</param>
    /// <param name="isTruncated">
    /// <c>true</c> when the byte sequence was truncated (i.e., the file is larger than the buffer).
    /// Trailing bytes equal to <see cref="MaxMultiByteLength"/> are excluded to prevent
    /// false decode failures at multibyte character boundaries.
    /// </param>
    /// <returns>
    /// The detected <see cref="Encoding"/>. Defaults to UTF-8 when encoding cannot be determined.
    /// </returns>
    /// <remarks>
    /// Detection order:
    /// <list type="number">
    ///   <item>BOM check ? UTF-8, UTF-32 LE/BE (checked before UTF-16 to avoid prefix collisions), UTF-16 LE/BE.</item>
    ///   <item>Strict decode ? attempt UTF-8 and Shift_JIS; if only one succeeds, that encoding wins.</item>
    ///   <item>Heuristic ? when both succeed (or both fail), the encoding whose decoded text contains more
    ///         Japanese characters (Hiragana, Katakana, CJK, full-width) is chosen.</item>
    ///   <item>Fallback ? UTF-8 is returned (e.g., pure ASCII content).</item>
    /// </list>
    /// </remarks>
    private static Encoding DetectEncodingFromBytes(ReadOnlySpan<byte> data, bool isTruncated)
    {
        if (data.Length < 2)
            return Encoding.UTF8;

        // BOM detection ? UTF-32 must be checked before UTF-16 because:
        //   UTF-32 LE BOM (FF FE 00 00) starts with the same bytes as UTF-16 LE BOM (FF FE)
        //   UTF-32 BE BOM (00 00 FE FF) starts with the same bytes as UTF-16 BE BOM (FE FF)
        // UTF-8 BOM (EF BB BF) has no such ambiguity.
        // When a BOM is present it unambiguously identifies the encoding, so return immediately.
        if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF)
            return Encoding.UTF8;

        if (data.Length >= 4 && data[0] == 0xFF && data[1] == 0xFE && data[2] == 0x00 && data[3] == 0x00)
            return new UTF32Encoding(bigEndian: false, byteOrderMark: true); // UTF-32 LE

        if (data.Length >= 4 && data[0] == 0x00 && data[1] == 0x00 && data[2] == 0xFE && data[3] == 0xFF)
            return new UTF32Encoding(bigEndian: true, byteOrderMark: true); // UTF-32 BE

        if (data[0] == 0xFF && data[1] == 0xFE)
            return Encoding.Unicode; // UTF-16 LE

        if (data[0] == 0xFE && data[1] == 0xFF)
            return Encoding.BigEndianUnicode; // UTF-16 BE

        // When the buffer was truncated, trim trailing bytes to avoid decode failures
        // caused by a partially-read multibyte character at the end of the buffer
        var decodeData = (isTruncated && data.Length > MaxMultiByteLength)
            ? data[..^MaxMultiByteLength]
            : data;

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        StrictUtf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
        StrictShiftJis = Encoding.GetEncoding(
            "Shift_JIS", EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);

        // No BOM found ? use strict decoding to determine encoding validity
        var isValidUtf8 = TryDecodeStrict(StrictUtf8, decodeData);
        var isValidShiftJis = TryDecodeStrict(StrictShiftJis, decodeData);

        if (isValidUtf8 && !isValidShiftJis)
            return Encoding.UTF8;

        if (isValidShiftJis && !isValidUtf8)
            return ShiftJisEncoding;

        // Both valid or both invalid ? use heuristic: compare Japanese character counts
        // from non-strict decoding to decide which encoding is more likely
        var utf8Text = Encoding.UTF8.GetString(decodeData);
        var sjisText = ShiftJisEncoding.GetString(decodeData);

        var utf8JapaneseCount = CountJapaneseCharacters(utf8Text);
        var sjisJapaneseCount = CountJapaneseCharacters(sjisText);

        if (sjisJapaneseCount > utf8JapaneseCount)
            return ShiftJisEncoding;

        // Default to UTF-8 (e.g., pure ASCII content where both encodings are valid)
        return Encoding.UTF8;
    }

    /// <summary>
    /// Counts characters that fall within common Japanese Unicode ranges.
    /// </summary>
    /// <param name="text">The decoded string to analyze. Must not be null.</param>
    /// <returns>The number of characters in the Hiragana, Katakana, CJK Unified Ideographs,
    /// or full-width alphanumeric/symbol ranges.</returns>
    /// <remarks>
    /// Used as a heuristic in <see cref="DetectEncodingFromBytes"/> to disambiguate between
    /// UTF-8 and Shift_JIS when both produce valid decode results. A higher count indicates
    /// the encoding is more likely correct for Japanese text.
    /// </remarks>
    private static int CountJapaneseCharacters(string text)
    {
        var count = 0;
        foreach (var c in text)
        {
            if (c is (>= '\u3040' and <= '\u309F')   // Hiragana
                or (>= '\u30A0' and <= '\u30FF')    // Katakana
                or (>= '\u4E00' and <= '\u9FFF')    // CJK Unified Ideographs
                or (>= '\uFF00' and <= '\uFFEF'))   // Full-width alphanumeric & symbols
            {
                count++;
            }
        }
        return count;
    }
    
    /// <summary>
    /// Attempts to decode the byte span using the given encoding configured with
    /// <see cref="DecoderFallback.ExceptionFallback"/>.
    /// </summary>
    /// <param name="encoding">
    /// A strict encoding instance that throws <see cref="DecoderFallbackException"/>
    /// on invalid byte sequences.
    /// </param>
    /// <param name="data">The raw bytes to decode.</param>
    /// <returns>
    /// <c>true</c> if decoding succeeds without errors; <c>false</c> if a
    /// <see cref="DecoderFallbackException"/> is thrown.
    /// </returns>
    /// <remarks>
    /// This is a validation-only call; the decoded string result is discarded.
    /// Used by <see cref="DetectEncodingFromBytes"/> to determine whether the byte
    /// sequence is valid under a given encoding.
    /// </remarks>
    private static bool TryDecodeStrict(Encoding encoding, ReadOnlySpan<byte> data)
    {
        try
        {
            encoding.GetString(data);
            return true;
        }
        catch (DecoderFallbackException e)
        {
            LoggerHelper.Instance.LogError($"Error decoding bytes with {encoding.WebName}: {e.Message}");
            return false;
        }
    }
}