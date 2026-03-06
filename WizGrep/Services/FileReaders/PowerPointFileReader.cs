using System;
using System.Collections.Generic;
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
            using var document = PresentationDocument.Open(filePath, false);
            var presentationPart = document.PresentationPart;

            if (presentationPart?.Presentation?.SlideIdList == null) return results;

            var slideIndex = 1;
            foreach (var slideId in presentationPart.Presentation.SlideIdList.Descendants<SlideId>())
            {
                var slidePart = (SlidePart?)presentationPart.GetPartById(slideId.RelationshipId!);
                if (slidePart == null) continue;

                var slideName = $"{ResourceLoaderHelper.GetString("SlideLabel")}{slideIndex}";

                // Extract text from shapes on this slide
                var shapes = slidePart.Slide?.Descendants<Shape>();
                var shapeIndex = 1;
                
                if(shapes == null) continue;

                foreach (var shape in shapes)
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
                var graphicFrames = slidePart.Slide?.Descendants<GraphicFrame>();
                
                if(graphicFrames == null) continue;
                
                foreach (var frame in graphicFrames)
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

                slideIndex++;
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading PowerPoint file '{filePath}' with excelFormula={excelFormula}: {e.StackTrace}");
        }

        return results;
    }
}