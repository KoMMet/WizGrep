using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using WizGrep.Helpers;
using WizGrep.Models;

namespace WizGrep.Services.FileReaders;

/// <summary>
/// Reads legacy Excel binary files (.xls) using the NPOI library.
/// Extracts cell values (including date formatting), cell comments, and text
/// from drawing shapes (including grouped shapes via recursive traversal).
/// </summary>
public class NpoiExcelFileReader : IFileReader
{
    public IEnumerable<string> SupportedExtensions => [".xls"];

    public IEnumerable<GrepResult> ReadFile(string filePath, bool excelFormula)
    {
        var results = new List<GrepResult>();

        try
        {
            using var stream = File.OpenRead(filePath);
            IWorkbook workbook = new HSSFWorkbook(stream);

            for (var i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                var sheetName = sheet.SheetName ?? $"Sheet{i + 1}";

                // Extract cell values from every row
                for (var rowIdx = 0; rowIdx <= sheet.LastRowNum; rowIdx++)
                {
                    var row = sheet.GetRow(rowIdx);
                    if (row == null) continue;

                    foreach (var cell in row.Cells)
                    {
                        var cellValue = GetCellStringValue(cell, excelFormula);
                        if (!string.IsNullOrWhiteSpace(cellValue))
                        {
                            results.Add(new GrepResult
                            {
                                FilePath = filePath,
                                LineNumber = rowIdx + 1,
                                SheetName = sheetName,
                                CellAddress = cell.Address.ToString(),
                                Content = cellValue
                            });
                        }

                        // Extract the cell comment text, if present
                        var comment = cell.CellComment;
                        if (comment != null)
                        {
                            var commentText = comment.String?.String;
                            if (!string.IsNullOrWhiteSpace(commentText))
                            {
                                results.Add(new GrepResult
                                {
                                    FilePath = filePath,
                                    LineNumber = 0,
                                    SheetName = sheetName,
                                    CellAddress = cell.Address.ToString(),
                                    ObjectName = $"{ResourceLoaderHelper.GetString("CommentLabel")}",
                                    Content = commentText
                                });
                            }
                        }
                    }
                }

                // Extract text from drawing shapes (recursively includes grouped shapes)
                if (sheet.DrawingPatriarch is HSSFPatriarch patriarch)
                {
                    var shapeIndex = 1;
                    foreach (var shape in GetAllShapes(patriarch))
                    {
                        var text = GetShapeText(shape);
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
                        }

                        shapeIndex++;
                    }
                }
            }
        }
        catch (Exception)
        {
            // Silently ignore file read errors
        }

        return results;
    }

    /// <summary>
    /// Recursively enumerates all shapes in the container, flattening nested
    /// <see cref="HSSFShapeGroup"/> hierarchies.
    /// </summary>
    private static IEnumerable<HSSFShape> GetAllShapes(HSSFShapeContainer container)
    {
        foreach (var shape in container.Children)
        {
            if (shape is HSSFShapeGroup group)
            {
                foreach (var child in GetAllShapes(group))
                    yield return child;
            }
            else
            {
                yield return shape;
            }
        }
    }

    /// <summary>
    /// Returns the text content of a simple shape, or <c>null</c> for other shape types.
    /// </summary>
    private static string? GetShapeText(HSSFShape shape)
    {
        if (shape is HSSFSimpleShape simpleShape)
            return simpleShape.String?.String;
        return null;
    }

    /// <summary>
    /// Converts a cell's value to its string representation.
    /// Handles string, numeric (including date-formatted), boolean, and formula cells.
    /// When <paramref name="excelFormula"/> is <c>true</c> and the cell is a formula,
    /// the raw formula text is returned prefixed with '='.
    /// </summary>
    private static string GetCellStringValue(ICell cell, bool excelFormula)
    {
        try
        {
            if (excelFormula && cell.CellType == CellType.Formula)
                return "=" + cell.CellFormula;

            return cell.CellType switch
            {
                CellType.String => cell.StringCellValue,
                CellType.Numeric => DateUtil.IsCellDateFormatted(cell)
                    ? cell.DateCellValue?.ToString() ?? ""
                    : cell.NumericCellValue.ToString(),
                CellType.Boolean => cell.BooleanCellValue ? "TRUE" : "FALSE",
                CellType.Formula => GetFormulaCellValue(cell),
                _ => cell.ToString() ?? ""
            };
        }
        catch
        {
            return cell.ToString() ?? "";
        }
    }

    /// <summary>
    /// Retrieves the cached result value of a formula cell as a string.
    /// Falls back to <c>cell.ToString()</c> for unrecognized result types.
    /// </summary>
    private static string GetFormulaCellValue(ICell cell)
    {
        try
        {
            return cell.CachedFormulaResultType switch
            {
                CellType.String => cell.StringCellValue,
                CellType.Numeric => cell.NumericCellValue.ToString(),
                CellType.Boolean => cell.BooleanCellValue ? "TRUE" : "FALSE",
                _ => cell.ToString() ?? ""
            };
        }
        catch
        {
            return cell.ToString() ?? "";
        }
    }
}
