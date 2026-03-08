using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WizGrep.Helpers;
using WizGrep.Models;
using Shape = DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape;

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
            using var document = SpreadsheetDocument.Open(filePath, false);
            var workbookPart = document.WorkbookPart;

            if (workbookPart == null) return results;

            var sheets = workbookPart.Workbook?.Descendants<Sheet>();
            var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;

            if(sheets == null) return results;
            
            foreach (var sheet in sheets)
            {
                var sheetName = sheet.Name?.Value ?? "Sheet";
                var worksheetPart = (WorksheetPart?)workbookPart.GetPartById(sheet.Id!);

                if (worksheetPart == null) continue;

                // Extract cell values from every row/cell in the sheet
                var sheetData = worksheetPart.Worksheet?.Descendants<SheetData>().FirstOrDefault();
                if (sheetData != null)
                    foreach (var row in sheetData.Descendants<Row>())
                    foreach (var cell in row.Descendants<Cell>())
                    {
                        var cellValue = GetCellValue(cell, sharedStringTable, excelFormula);
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

                if(excelFormula)
                    continue; // If we're only interested in formulas, skip shapes and comments
                
                // Extract text from drawing shapes (textboxes, callouts, etc.)
                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart != null)
                {
                    var shapes = drawingsPart.WorksheetDrawing?
                        .Descendants<Shape>();

                    if (shapes != null)
                    {
                        var shapeIndex = 1;
                        foreach (var shape in shapes)
                        {
                            var textBody = shape.Descendants<Paragraph>();
                            foreach (var para in textBody)
                            {
                                var text = para.InnerText;
                                if (!string.IsNullOrWhiteSpace(text))
                                    results.Add(new GrepResult
                                    {
                                        FilePath = filePath,
                                        LineNumber = 0,
                                        SheetName = sheetName,
                                        ObjectName = $"{ResourceLoaderHelper.GetString("ShapeLabel")}{shapeIndex}",
                                        Content = text
                                    });
                            }

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

    private static string GetCellValue(Cell cell, SharedStringTable? sharedStringTable, bool excelFormula)
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

        return value;
    }
}