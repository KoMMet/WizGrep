using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WizGrep.Helpers;
using WizGrep.Models;

namespace WizGrep.Services.FileReaders;

/// <summary>
/// Reads modern Excel files (.xlsx, .xlsm) using the Open XML SDK.
/// Extracts cell values, shape/textbox text, and cell comments from every worksheet.
/// </summary>
public class ExcelFileReader : IFileReader
{
    public IEnumerable<string> SupportedExtensions => new[] { ".xlsx", ".xlsm" };

    public IEnumerable<GrepResult> ReadFile(string filePath, bool excelFormula)
    {
        var results = new List<GrepResult>();

        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var document = SpreadsheetDocument.Open(stream, false);
            OpenXmlSearchHelper.AddPackageProperties(document, filePath, results);
            OpenXmlSearchHelper.AddHyperlinks(document, filePath, results);
            OpenXmlSearchHelper.AddAlternativeText(document, filePath, results);
            OpenXmlSearchHelper.AddChartText(document, filePath, results);
            OpenXmlSearchHelper.AddSmartArtText(document, filePath, results);
            OpenXmlSearchHelper.AddThreadedComments(document, filePath, results);

            var workbookPart = document.WorkbookPart;

            if (workbookPart == null) return results;

            var sheets = workbookPart.Workbook?.Descendants<Sheet>();
            var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
            var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;

            if(sheets == null) return results;
            
            foreach (var sheet in sheets)
            {
                var sheetName = sheet.Name?.Value ?? "Sheet";
                var worksheetPart = (WorksheetPart?)workbookPart.GetPartById(sheet.Id!);

                if (worksheetPart == null) continue;

                OpenXmlSearchHelper.AddHeaderFooterText(worksheetPart, filePath, sheetName, results);

                // Extract cell values from every row/cell in the sheet
                var sheetData = worksheetPart.Worksheet?.Descendants<SheetData>().FirstOrDefault();
                if (sheetData != null)
                    foreach (var row in sheetData.Descendants<Row>())
                    foreach (var cell in row.Descendants<Cell>())
                    {
                        var cellValue = GetCellValue(cell, sharedStringTable, stylesheet, excelFormula);
                        if (!string.IsNullOrWhiteSpace(cellValue))
                        {
                            var cellAddress = cell.CellReference?.Value ?? "";
                            results.Add(new GrepResult
                            {
                                FilePath = filePath,
                                LineNumber = (int)(row.RowIndex?.Value ?? 0),
                                SheetName = sheetName,
                                CellAddress = cellAddress,
                                Content = cellValue
                            });
                        }
                    }

                // Extract text from drawing shapes (textboxes, callouts, etc.)
                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart != null)
                {
                    var textBodies = drawingsPart.WorksheetDrawing?.Descendants<TextBody>();

                    if (textBodies != null)
                    {
                        var shapeIndex = 1;
                        foreach (var textBody in textBodies)
                        {
                            var hasParagraph = false;
                            foreach (var paragraph in textBody.Elements<Paragraph>())
                            {
                                var text = paragraph.InnerText;
                                if (!string.IsNullOrWhiteSpace(text))
                                {
                                    results.Add(new GrepResult
                                    {
                                        FilePath = filePath,
                                        LineNumber = 0,
                                        SheetName = sheetName,
                                        ObjectName = $"{ResourceLoaderHelper.GetString("ShapeLabel")}{shapeIndex}",
                                        Content = text
                                    });
                                    hasParagraph = true;
                                }
                            }

                            if (hasParagraph)
                                shapeIndex++;
                        }
                    }
                }

                // Extract cell comments / notes
                var commentsPart = worksheetPart.WorksheetCommentsPart;
                if (commentsPart != null)
                {
                    var comments = commentsPart.Comments?.CommentList?.Descendants<Comment>();
                    if (comments == null) continue;
                    foreach (var comment in comments)
                    {
                        var cellRef = comment.Reference?.Value ?? "";
                        var commentText = comment.InnerText;
                        if (!string.IsNullOrWhiteSpace(commentText))
                            results.Add(new GrepResult
                            {
                                FilePath = filePath,
                                LineNumber = 0,
                                SheetName = sheetName,
                                CellAddress = cellRef,
                                ObjectName = $"{ResourceLoaderHelper.GetString("CommentLabel")}",
                                Content = commentText
                            });
                    }
                }
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading Excel file '{filePath}' with excelFormula={excelFormula}: {e.StackTrace}");
        }

        return results;
    }

    private static string GetCellValue(Cell cell, SharedStringTable? sharedStringTable, Stylesheet? stylesheet, bool excelFormula)
    {
        if (excelFormula)
        {
            if (cell.CellFormula != null && !string.IsNullOrEmpty(cell.CellFormula.Text))
                return "=" + cell.CellFormula.Text;
            return string.Empty;
        }

        if (cell.DataType?.Value == CellValues.InlineString)
            return cell.InlineString?.InnerText ?? string.Empty;

        if (cell.CellValue == null) return string.Empty;

        var value = cell.CellValue.InnerText;

        if (cell.DataType?.Value == CellValues.SharedString && sharedStringTable != null)
        {
            if (int.TryParse(value, out var index))
            {
                var sharedStringItem = sharedStringTable.ElementAt(index);
                return sharedStringItem.InnerText;
            }
        }

        if (cell.DataType?.Value == CellValues.Boolean)
            return value == "1" ? "TRUE" : "FALSE";

        // Convert date-formatted numeric cells to DateTime strings (consistent with .xls behavior)
        if (cell.DataType == null || cell.DataType.Value == CellValues.Number)
        {
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var numericValue)
                && IsDateCell(cell, stylesheet))
            {
                try
                {
                    return DateTime.FromOADate(numericValue).ToString();
                }
                catch
                {
                    // OADate conversion can fail for out-of-range values; fall through to raw value
                }
            }
        }

        return value;
    }

    /// <summary>
    /// Determines whether a cell is formatted as a date/time based on its style and number format.
    /// </summary>
    private static bool IsDateCell(Cell cell, Stylesheet? stylesheet)
    {
        if (stylesheet?.CellFormats == null || cell.StyleIndex == null)
            return false;

        var styleIndex = (int)cell.StyleIndex.Value;
        var cellFormats = stylesheet.CellFormats.Elements<CellFormat>().ToList();
        if (styleIndex < 0 || styleIndex >= cellFormats.Count)
            return false;

        var numberFormatId = cellFormats[styleIndex].NumberFormatId?.Value ?? 0;

        // Built-in date/time format IDs
        if (numberFormatId >= 14 && numberFormatId <= 22) return true;
        if (numberFormatId >= 27 && numberFormatId <= 36) return true;
        if (numberFormatId >= 45 && numberFormatId <= 47) return true;
        if (numberFormatId >= 50 && numberFormatId <= 58) return true;

        // Custom number formats (IDs >= 164): check format code for date/time characters
        if (numberFormatId >= 164 && stylesheet.NumberingFormats != null)
        {
            var numFormat = stylesheet.NumberingFormats
                .Elements<NumberingFormat>()
                .FirstOrDefault(nf => nf.NumberFormatId?.Value == numberFormatId);

            if (numFormat?.FormatCode?.Value != null)
            {
                // Strip quoted literals and escaped characters before checking
                var code = Regex.Replace(numFormat.FormatCode.Value, "\"[^\"]*\"", "");
                code = Regex.Replace(code, "\\\\.", "");
                // Strip underscore-skip-width patterns (e.g., _m, _)) and bracketed sections (e.g., [Red], [$-409])
                code = Regex.Replace(code, "_.", "");
                code = Regex.Replace(code, @"\[.*?\]", "");
                code = code.ToLowerInvariant();

                if (code.Contains('y') || code.Contains('d') ||
                    code.Contains('h') || code.Contains('s') ||
                    code.Contains("am") || code.Contains("pm"))
                    return true;

                // 'm' alone also indicates month/minute in Excel format codes
                if (code.Contains('m'))
                    return true;
            }
        }

        return false;
    }
}