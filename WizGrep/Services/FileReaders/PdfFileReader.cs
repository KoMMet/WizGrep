using System;
using System.Collections.Generic;
using UglyToad.PdfPig;
using WizGrep.Helpers;
using WizGrep.Models;

namespace WizGrep.Services.FileReaders;

/// <summary>
/// Reads PDF files using the PdfPig library and extracts text page by page,
/// splitting each page's content into individual lines.
/// </summary>
public class PdfFileReader : IFileReader
{
    public IEnumerable<string> SupportedExtensions => [".pdf"];

    public IEnumerable<GrepResult> ReadFile(string filePath, bool excelFormula)
    {
        var results = new List<GrepResult>();

        try
        {
            using var document = PdfDocument.Open(filePath);

            foreach (var page in document.GetPages())
            {
                var pageNumber = page.Number;
                var text = page.Text;

                if (string.IsNullOrWhiteSpace(text)) continue;

                // Split page text into individual lines
                var lines = text.Split(["\r\n", "\r", "\n"], StringSplitOptions.None);

                for (var i = 0; i < lines.Length; i++)
                {
                    var line = lines[i];
                    if (!string.IsNullOrWhiteSpace(line))
                        results.Add(new GrepResult
                        {
                            FilePath = filePath,
                            LineNumber = i + 1,
                            SheetName = $"{ResourceLoaderHelper.GetString("PageLabel")}{pageNumber}",
                            Content = line.Trim()
                        });
                }
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading PDF file '{filePath}' with excelFormula={excelFormula}: {e.Message}");
        }

        return results;
    }
}