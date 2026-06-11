using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using UglyToad.PdfPig.AcroForms;
using UglyToad.PdfPig.AcroForms.Fields;
using UglyToad.PdfPig.Annotations;
using UglyToad.PdfPig.Actions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Outline;
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
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var document = PdfDocument.Open(stream);
            AddDocumentInformation(document, filePath, results);
            AddBookmarks(document, filePath, results);
            AddFormFields(document, filePath, results);

            foreach (var page in document.GetPages())
            {
                var pageNumber = page.Number;
                var text = page.Text;

                if (!string.IsNullOrWhiteSpace(text))
                {
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
                                Content = line
                            });
                    }
                }

                AddHyperlinks(page, filePath, results);

                // Extract text from PDF annotations (comments, sticky notes, etc.)
                try
                {
                    var shapeIndex = 1;
                    var commentIndex = 1;
                    var annotations = page.GetAnnotations();
                    foreach (var annotation in annotations)
                    {
                        if (IsShapeTextAnnotation(annotation.Type))
                        {
                            var hasParagraph = false;
                            foreach (var paragraph in GetAnnotationParagraphs(annotation))
                            {
                                results.Add(new GrepResult
                                {
                                    FilePath = filePath,
                                    LineNumber = 0,
                                    SheetName = $"{ResourceLoaderHelper.GetString("PageLabel")}{pageNumber}",
                                    ObjectName = $"{ResourceLoaderHelper.GetString("ShapeLabel")}{shapeIndex}",
                                    Content = paragraph
                                });
                                hasParagraph = true;
                            }

                            if (hasParagraph)
                                shapeIndex++;

                            continue;
                        }

                        var annotationText = GetAnnotationSearchText(annotation);
                        if (string.IsNullOrWhiteSpace(annotationText))
                        {
                            continue;
                        }

                        results.Add(new GrepResult
                        {
                            FilePath = filePath,
                            LineNumber = 0,
                            SheetName = $"{ResourceLoaderHelper.GetString("PageLabel")}{pageNumber}",
                            ObjectName = $"{ResourceLoaderHelper.GetString("CommentLabel")}{commentIndex}",
                            Content = annotationText
                        });
                        commentIndex++;
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.Instance.LogError($"Error reading annotations from page {pageNumber} of '{filePath}': {ex.Message}");
                }
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading PDF file '{filePath}' with excelFormula={excelFormula}: {e.StackTrace}");
        }

        return results;
    }

    private static void AddDocumentInformation(PdfDocument document, string filePath, IList<GrepResult> results)
    {
        try
        {
            var info = document.Information;
            var values = new[]
            {
                info.Title,
                info.Author,
                info.Subject,
                info.Keywords,
                info.Creator,
                info.Producer,
                info.CreationDate,
                info.ModifiedDate
            };

            AddDistinctValues(values, filePath, null, ResourceLoaderHelper.GetString("DocumentPropertiesLabel"), results);
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading PDF document information for '{filePath}': {e.StackTrace}");
        }
    }

    private static void AddBookmarks(PdfDocument document, string filePath, IList<GrepResult> results)
    {
        try
        {
            if (!document.TryGetBookmarks(out var bookmarks)) return;

            var values = new List<string>();
            foreach (var bookmark in bookmarks.GetNodes())
            {
                if (!string.IsNullOrWhiteSpace(bookmark.Title))
                    values.Add(bookmark.Title);

                if (bookmark is UriBookmarkNode uriBookmark && !string.IsNullOrWhiteSpace(uriBookmark.Uri))
                    values.Add(uriBookmark.Uri);
                else if (bookmark is ExternalBookmarkNode externalBookmark && !string.IsNullOrWhiteSpace(externalBookmark.FileName))
                    values.Add(externalBookmark.FileName);
            }

            AddDistinctValues(values, filePath, null, ResourceLoaderHelper.GetString("BookmarkLabel"), results);
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading PDF bookmarks for '{filePath}': {e.StackTrace}");
        }
    }

    private static void AddFormFields(PdfDocument document, string filePath, IList<GrepResult> results)
    {
        try
        {
            if (!document.TryGetForm(out var form)) return;

            var index = 1;
            foreach (var field in form.GetFields())
            {
                var values = new List<string?>
                {
                    field.Information.PartialName,
                    field.Information.AlternateName,
                    field.Information.MappingName,
                    GetFieldValue(field)
                };

                var text = string.Join(" ", values.Where(value => !string.IsNullOrWhiteSpace(value)).Select(value => value!.Trim()).Distinct());
                if (string.IsNullOrWhiteSpace(text)) continue;

                results.Add(new GrepResult
                {
                    FilePath = filePath,
                    LineNumber = 0,
                    SheetName = field.PageNumber.HasValue ? $"{ResourceLoaderHelper.GetString("PageLabel")}{field.PageNumber.Value}" : null,
                    ObjectName = $"{ResourceLoaderHelper.GetString("FormFieldLabel")}{index++}",
                    Content = text
                });
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading PDF form fields for '{filePath}': {e.StackTrace}");
        }
    }

    private static void AddHyperlinks(Page page, string filePath, IList<GrepResult> results)
    {
        try
        {
            var index = 1;
            foreach (var hyperlink in page.GetHyperlinks())
            {
                var values = new[] { hyperlink.Text, hyperlink.Uri };
                var text = string.Join(" ", values.Where(value => !string.IsNullOrWhiteSpace(value)).Select(value => value!.Trim()).Distinct());
                if (string.IsNullOrWhiteSpace(text)) continue;

                results.Add(new GrepResult
                {
                    FilePath = filePath,
                    LineNumber = 0,
                    SheetName = $"{ResourceLoaderHelper.GetString("PageLabel")}{page.Number}",
                    ObjectName = $"{ResourceLoaderHelper.GetString("HyperlinkLabel")}{index++}",
                    Content = text
                });
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading PDF hyperlinks from page {page.Number} of '{filePath}': {e.StackTrace}");
        }
    }

    private static string? GetFieldValue(AcroFieldBase field)
    {
        return field switch
        {
            AcroTextField textField => textField.Value,
            AcroCheckboxField checkboxField => checkboxField.IsChecked ? checkboxField.CurrentValue?.Data : null,
            AcroRadioButtonsField radioButtonsField => string.Join(" ", radioButtonsField.Children.Select(GetFieldValue).Where(value => !string.IsNullOrWhiteSpace(value))),
            _ => AcroFormExtensions.GetFieldValue(field).Value?.ToString()
        };
    }

    private static string? GetAnnotationSearchText(Annotation annotation)
    {
        var values = new List<string?>
        {
            annotation.Content,
            annotation.Name,
            annotation.Action is UriAction uriAction ? uriAction.Uri : null
        };

        return string.Join(" ", values.Where(value => !string.IsNullOrWhiteSpace(value)).Select(value => value!.Trim()).Distinct());
    }

    private static IEnumerable<string> GetAnnotationParagraphs(Annotation annotation)
    {
        if (string.IsNullOrWhiteSpace(annotation.Content))
            yield break;

        var paragraphs = annotation.Content.Split(["\r\n", "\r", "\n"], StringSplitOptions.None);
        foreach (var paragraph in paragraphs)
        {
            if (!string.IsNullOrWhiteSpace(paragraph))
                yield return paragraph;
        }
    }

    private static bool IsShapeTextAnnotation(AnnotationType annotationType)
    {
        return annotationType == AnnotationType.FreeText || annotationType == AnnotationType.Widget;
    }

    private static bool IsCommentAnnotation(AnnotationType annotationType)
    {
        return annotationType != AnnotationType.Link && annotationType != AnnotationType.Widget;
    }

    private static void AddDistinctValues(IEnumerable<string?> values, string filePath, string? sheetName, string objectName, IList<GrepResult> results)
    {
        var index = 1;
        foreach (var value in values.Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v!.Trim()).Distinct())
        {
            results.Add(new GrepResult
            {
                FilePath = filePath,
                LineNumber = 0,
                SheetName = sheetName,
                ObjectName = $"{objectName}{index++}",
                Content = value
            });
        }
    }
}