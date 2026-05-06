using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using WizGrep.Helpers;
using WizGrep.Models;

namespace WizGrep.Services.FileReaders;

/// <summary>
/// Reads modern Word documents (.docx, .docm) using the Open XML SDK.
/// Extracts text from body paragraphs, VML shapes, headers, and footers.
/// Paragraphs inside VML shapes are excluded from the body pass to avoid duplication.
/// </summary>
public class WordFileReader : IFileReader
{
    public IEnumerable<string> SupportedExtensions => [".docx", ".docm"];

    /// <summary>
    /// Reads the Word document and returns one <see cref="GrepResult"/> per paragraph,
    /// shape, header, and footer found in the file.
    /// Returns an empty collection on any read error.
    /// </summary>
    public IEnumerable<GrepResult> ReadFile(string filePath, bool excelFormula)
    {
        var results = new List<GrepResult>();

        try
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var document = WordprocessingDocument.Open(stream, false);
            OpenXmlSearchHelper.AddPackageProperties(document, filePath, results);
            OpenXmlSearchHelper.AddHyperlinks(document, filePath, results);
            OpenXmlSearchHelper.AddAlternativeText(document, filePath, results);

            var body = document.MainDocumentPart?.Document?.Body;

            if (body != null)
            {
                var lineNumber = 1;
                var paragraphs = body.Descendants<Paragraph>();

                foreach (var paragraph in paragraphs)
                {
                    // Skip paragraphs inside VML shapes (they are extracted separately below)
                    if (paragraph.Ancestors<Shape>().Any())
                        continue;

                    // Exclude Runs containing VML shapes to prevent duplicate shape text
                    var text = string.Concat(
                        paragraph.Descendants<Run>()
                            .Where(r => !r.Descendants<Shape>().Any())
                            .Select(r => r.InnerText));
                    if (!string.IsNullOrWhiteSpace(text))
                        results.Add(new GrepResult
                        {
                            FilePath = filePath,
                            LineNumber = lineNumber,
                            Content = text
                        });
                    lineNumber++;
                }

                // Extract text from VML shapes
                var shapes = body.Descendants<Shape>();
                var shapeIndex = 1;
                foreach (var shape in shapes)
                {
                    var textContent = shape.InnerText;
                    if (!string.IsNullOrWhiteSpace(textContent))
                        results.Add(new GrepResult
                        {
                            FilePath = filePath,
                            LineNumber = 0,
                            ObjectName = $"{ResourceLoaderHelper.GetString("ShapeLabel")}{shapeIndex}",
                            Content = textContent
                        });
                    shapeIndex++;
                }
            }

            if (document.MainDocumentPart != null)
            {
                var shapeName = ResourceLoaderHelper.GetString("ShapeLabel");
                OpenXmlSearchHelper.AddWordDrawingTextBoxes(document.MainDocumentPart, filePath, shapeName, results);

                foreach (var headerPart in document.MainDocumentPart.HeaderParts)
                    OpenXmlSearchHelper.AddWordDrawingTextBoxes(headerPart, filePath, shapeName, results);

                foreach (var footerPart in document.MainDocumentPart.FooterParts)
                    OpenXmlSearchHelper.AddWordDrawingTextBoxes(footerPart, filePath, shapeName, results);
            }

            // Extract text from headers and footers
            if (document.MainDocumentPart?.HeaderParts != null)
            {
                var headerIndex = 1;
                foreach (var headerPart in document.MainDocumentPart.HeaderParts)
                {
                    var headerText = headerPart.Header?.InnerText;
                    if (!string.IsNullOrWhiteSpace(headerText))
                        results.Add(new GrepResult
                        {
                            FilePath = filePath,
                            LineNumber = 0,
                            ObjectName = $"{ResourceLoaderHelper.GetString("HeaderLabel")}{headerIndex}",
                            Content = headerText
                        });
                    headerIndex++;
                }
            }

            if (document.MainDocumentPart?.FooterParts != null)
            {
                var footerIndex = 1;
                foreach (var footerPart in document.MainDocumentPart.FooterParts)
                {
                    var footerText = footerPart.Footer?.InnerText;
                    if (!string.IsNullOrWhiteSpace(footerText))
                        results.Add(new GrepResult
                        {
                            FilePath = filePath,
                            LineNumber = 0,
                            ObjectName = $"{ResourceLoaderHelper.GetString("FooterLabel")}{footerIndex}",
                            Content = footerText
                        });
                    footerIndex++;
                }
            }

            // Extract text from comments
            var commentsPart = document.MainDocumentPart?.WordprocessingCommentsPart;
            if (commentsPart?.Comments != null)
            {
                var commentIndex = 1;
                foreach (var comment in commentsPart.Comments.Descendants<Comment>())
                {
                    var commentText = comment.InnerText;
                    if (!string.IsNullOrWhiteSpace(commentText))
                        results.Add(new GrepResult
                        {
                            FilePath = filePath,
                            LineNumber = 0,
                            ObjectName = $"{ResourceLoaderHelper.GetString("CommentLabel")}{commentIndex}",
                            Content = commentText
                        });
                    commentIndex++;
                }
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading Word file '{filePath}' with excelFormula={excelFormula}: {e.StackTrace}");
        }

        return results;
    }
}