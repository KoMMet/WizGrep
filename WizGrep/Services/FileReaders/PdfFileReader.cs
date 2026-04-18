using System;
using System.Collections.Generic;
using System.IO;
using UglyToad.PdfPig.Annotations;
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
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var document = PdfDocument.Open(stream);

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

                        if (!IsCommentAnnotation(annotation.Type) || string.IsNullOrWhiteSpace(annotation.Content))
                        {
                            continue;
                        }

                        results.Add(new GrepResult
                        {
                            FilePath = filePath,
                            LineNumber = 0,
                            SheetName = $"{ResourceLoaderHelper.GetString("PageLabel")}{pageNumber}",
                            ObjectName = $"{ResourceLoaderHelper.GetString("CommentLabel")}{commentIndex}",
                            Content = annotation.Content
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
        return annotationType == AnnotationType.Text;
    }
}