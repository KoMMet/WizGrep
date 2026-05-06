using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using WizGrep.Helpers;
using WizGrep.Models;
using A = DocumentFormat.OpenXml.Drawing;

namespace WizGrep.Services.FileReaders;

/// <summary>
/// Reads modern PowerPoint files (.pptx, .pptm) using the Open XML SDK.
/// Extracts text from slide shapes (including named placeholders), graphic frames
/// (tables, charts), and speaker notes.
/// </summary>
public class PowerPointFileReader : IFileReader
{
    public IEnumerable<string> SupportedExtensions => [".pptx", ".pptm"];

    /// <summary>
    /// Reads all slides in the PowerPoint file and extracts text from shapes,
    /// graphic frames, and speaker notes. Returns an empty collection on any error.
    /// </summary>
    public IEnumerable<GrepResult> ReadFile(string filePath, bool excelFormula)
    {
        var results = new List<GrepResult>();

        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var document = PresentationDocument.Open(stream, false);
            OpenXmlSearchHelper.AddPackageProperties(document, filePath, results);
            OpenXmlSearchHelper.AddHyperlinks(document, filePath, results);
            OpenXmlSearchHelper.AddAlternativeText(document, filePath, results);
            OpenXmlSearchHelper.AddChartText(document, filePath, results);
            OpenXmlSearchHelper.AddSmartArtText(document, filePath, results);

            var presentationPart = document.PresentationPart;

            if (presentationPart?.Presentation?.SlideIdList == null) return results;

            ExtractPresentationSharedText(presentationPart, filePath, results);

            var slideIndex = 1;
            foreach (var slideId in presentationPart.Presentation.SlideIdList.Descendants<SlideId>())
            {
                var slidePart = (SlidePart?)presentationPart.GetPartById(slideId.RelationshipId!);
                if (slidePart == null) continue;

                var slideName = $"{ResourceLoaderHelper.GetString("SlideLabel")}{slideIndex}";
                var slide = slidePart.Slide;

                if (slide != null)
                {
                    // Extract text from shapes on this slide
                    var shapeIndex = 1;

                    foreach (var shape in slide.Descendants<Shape>())
                    {
                        var textBody = shape.TextBody;
                        if (textBody != null)
                        {
                            var lineNumber = 1;
                            foreach (var paragraph in textBody.Descendants<A.Paragraph>())
                            {
                                var text = paragraph.InnerText;
                                if (!string.IsNullOrWhiteSpace(text))
                                {
                                    // Use the shape's drawing name if available (e.g., placeholder title)
                                    var nvSpPr = shape.NonVisualShapeProperties;
                                    var shapeName = nvSpPr?.NonVisualDrawingProperties?.Name?.Value;

                                    results.Add(new GrepResult
                                    {
                                        FilePath = filePath,
                                        LineNumber = lineNumber,
                                        SheetName = slideName,
                                        ObjectName = shapeName ?? $"{ResourceLoaderHelper.GetString("ShapeLabel")}{shapeIndex}",
                                        Content = text
                                    });
                                    lineNumber++;
                                }
                            }
                        }

                        shapeIndex++;
                    }

                    // Extract text from graphic frames (tables, charts, SmartArt, etc.)
                    foreach (var frame in slide.Descendants<GraphicFrame>())
                    {
                        var paragraphs = frame.Descendants<A.Paragraph>();
                        var textIndex = 1;
                        foreach (var paragraph in paragraphs)
                        {
                            var text = paragraph.InnerText;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                results.Add(new GrepResult
                                {
                                    FilePath = filePath,
                                    LineNumber = textIndex,
                                    SheetName = slideName,
                                    ObjectName = $"{ResourceLoaderHelper.GetString("TableChartLabel")}",
                                    Content = text
                                });
                                textIndex++;
                            }
                        }
                    }
                }

                // Extract speaker notes for this slide
                var notesSlidePart = slidePart.NotesSlidePart;
                if (notesSlidePart != null)
                {
                    var noteShapes = notesSlidePart.NotesSlide?.Descendants<Shape>();
                    if (noteShapes != null)
                    {
                        var noteLineNumber = 1;
                        foreach (var shape in noteShapes)
                        {
                            var textBody = shape.TextBody;
                            if (textBody == null) continue;

                            foreach (var paragraph in textBody.Descendants<A.Paragraph>())
                            {
                                var text = paragraph.InnerText;
                                if (!string.IsNullOrWhiteSpace(text))
                                {
                                    results.Add(new GrepResult
                                    {
                                        FilePath = filePath,
                                        LineNumber = noteLineNumber,
                                        SheetName = slideName,
                                        ObjectName = $"{ResourceLoaderHelper.GetString("NoteLabel")}",
                                        Content = text
                                    });
                                    noteLineNumber++;
                                }
                            }
                        }
                    }
                }

                // Extract slide comments (review comments)
                var commentsPart = slidePart.SlideCommentsPart;
                if (commentsPart?.CommentList != null)
                {
                    var commentIndex = 1;
                    foreach (var comment in commentsPart.CommentList.Descendants<Comment>())
                    {
                        var commentText = comment.InnerText;
                        if (!string.IsNullOrWhiteSpace(commentText))
                        {
                            results.Add(new GrepResult
                            {
                                FilePath = filePath,
                                LineNumber = 0,
                                SheetName = slideName,
                                ObjectName = $"{ResourceLoaderHelper.GetString("CommentLabel")}{commentIndex}",
                                Content = commentText
                            });
                            commentIndex++;
                        }
                    }
                }

                slideIndex++;
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading PowerPoint file '{filePath}' with excelFormula={excelFormula}: {e.StackTrace}");
        }

        return results;
    }

    private static void ExtractPresentationSharedText(PresentationPart presentationPart, string filePath, IList<GrepResult> results)
    {
        var layoutIndex = 1;
        foreach (var slideMasterPart in presentationPart.SlideMasterParts)
        {
            var masterName = $"{ResourceLoaderHelper.GetString("SlideMasterLabel")}{layoutIndex++}";
            OpenXmlSearchHelper.AddDrawingParagraphText(slideMasterPart, filePath, masterName,
                ResourceLoaderHelper.GetString("ShapeLabel"), results);

            var slideLayoutIndex = 1;
            foreach (var slideLayoutPart in slideMasterPart.SlideLayoutParts)
            {
                var layoutName = $"{ResourceLoaderHelper.GetString("SlideLayoutLabel")}{slideLayoutIndex++}";
                OpenXmlSearchHelper.AddDrawingParagraphText(slideLayoutPart, filePath, layoutName,
                    ResourceLoaderHelper.GetString("ShapeLabel"), results);
            }
        }
    }
}