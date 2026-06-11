using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using WizGrep.Helpers;
using WizGrep.Models;

namespace WizGrep.Services.FileReaders;

internal static class OpenXmlSearchHelper
{
    public static void AddPackageProperties(OpenXmlPackage package, string filePath, IList<GrepResult> results)
    {
        try
        {
            var properties = package.PackageProperties;
            var values = new[]
            {
                properties.Title,
                properties.Subject,
                properties.Creator,
                properties.Keywords,
                properties.Description,
                properties.Category,
                properties.ContentStatus,
                properties.ContentType,
                properties.Identifier,
                properties.Language,
                properties.LastModifiedBy,
                properties.Revision,
                properties.Version
            };

            AddDistinctValues(values, filePath, null, ResourceLoaderHelper.GetString("DocumentPropertiesLabel"), results);
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading OpenXML package properties for '{filePath}': {e.StackTrace}");
        }
    }

    public static void AddHyperlinks(OpenXmlPackage package, string filePath, IList<GrepResult> results)
    {
        try
        {
            var values = new List<string>();
            foreach (var part in GetAllParts(package))
            {
                foreach (var relationship in part.HyperlinkRelationships)
                {
                    if (relationship.Uri != null)
                        values.Add(relationship.Uri.ToString());
                }
            }

            AddDistinctValues(values, filePath, null, ResourceLoaderHelper.GetString("HyperlinkLabel"), results);
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading OpenXML hyperlinks for '{filePath}': {e.StackTrace}");
        }
    }

    public static void AddAlternativeText(OpenXmlPackage package, string filePath, IList<GrepResult> results)
    {
        try
        {
            var values = new List<string>();
            foreach (var part in GetAllParts(package))
            {
                var document = TryLoadXml(part);
                if (document == null) continue;

                foreach (var element in document.Descendants().Where(IsNonVisualDrawingProperties))
                {
                    AddAttributeValue(element, "descr", values);
                    AddAttributeValue(element, "title", values);
                }
            }

            AddDistinctValues(values, filePath, null, ResourceLoaderHelper.GetString("AltTextLabel"), results);
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading OpenXML alternative text for '{filePath}': {e.StackTrace}");
        }
    }

    public static void AddChartText(OpenXmlPackage package, string filePath, IList<GrepResult> results, string? sheetName = null)
    {
        AddPartParagraphText(package, filePath, results, ResourceLoaderHelper.GetString("TableChartLabel"),
            part => part.ContentType.Contains("chart", StringComparison.OrdinalIgnoreCase), sheetName);
    }

    public static void AddSmartArtText(OpenXmlPackage package, string filePath, IList<GrepResult> results, string? sheetName = null)
    {
        AddPartParagraphText(package, filePath, results, ResourceLoaderHelper.GetString("SmartArtLabel"),
            part => part.ContentType.Contains("diagram", StringComparison.OrdinalIgnoreCase), sheetName);
    }

    public static void AddThreadedComments(OpenXmlPackage package, string filePath, IList<GrepResult> results)
    {
        try
        {
            var index = 1;
            foreach (var part in GetAllParts(package).Where(part => part.ContentType.Contains("threadedcomments", StringComparison.OrdinalIgnoreCase)))
            {
                var document = TryLoadXml(part);
                if (document == null) continue;

                foreach (var text in document.Descendants()
                             .Where(e => e.Name.LocalName is "text" or "t")
                             .Select(e => e.Value)
                             .Where(text => !string.IsNullOrWhiteSpace(text)))
                {
                    results.Add(new GrepResult
                    {
                        FilePath = filePath,
                        LineNumber = 0,
                        ObjectName = $"{ResourceLoaderHelper.GetString("CommentLabel")}{index++}",
                        Content = text
                    });
                }
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading OpenXML threaded comments for '{filePath}': {e.StackTrace}");
        }
    }

    public static void AddHeaderFooterText(OpenXmlPart part, string filePath, string? sheetName, IList<GrepResult> results)
    {
        try
        {
            var document = TryLoadXml(part);
            if (document == null) return;

            var headerFooterNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"
            };

            var values = document.Descendants()
                .Where(e => headerFooterNames.Contains(e.Name.LocalName))
                .Select(e => CleanExcelHeaderFooter(e.Value))
                .Where(text => !string.IsNullOrWhiteSpace(text));

            AddDistinctValues(values, filePath, sheetName, ResourceLoaderHelper.GetString("HeaderFooterLabel"), results);
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading OpenXML header/footer text for '{filePath}': {e.StackTrace}");
        }
    }

    public static void AddDrawingParagraphText(OpenXmlPart part, string filePath, string? sheetName, string objectName, IList<GrepResult> results)
    {
        try
        {
            var document = TryLoadXml(part);
            if (document == null) return;

            var index = 1;
            foreach (var text in GetDrawingParagraphTexts(document))
            {
                results.Add(new GrepResult
                {
                    FilePath = filePath,
                    LineNumber = index,
                    SheetName = sheetName,
                    ObjectName = objectName,
                    Content = text
                });
                index++;
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading OpenXML drawing text for '{filePath}': {e.StackTrace}");
        }
    }

    public static void AddWordDrawingTextBoxes(OpenXmlPart part, string filePath, string objectName, IList<GrepResult> results)
    {
        try
        {
            var document = TryLoadXml(part);
            if (document == null) return;

            var index = 1;
            foreach (var textBox in document.Descendants().Where(e => e.Name.LocalName == "txbxContent"))
            {
                foreach (var text in GetWordParagraphTexts(textBox))
                {
                    results.Add(new GrepResult
                    {
                        FilePath = filePath,
                        LineNumber = 0,
                        ObjectName = $"{objectName}{index}",
                        Content = text
                    });
                }

                index++;
            }

            foreach (var text in GetDrawingParagraphTexts(document))
            {
                results.Add(new GrepResult
                {
                    FilePath = filePath,
                    LineNumber = 0,
                    ObjectName = $"{objectName}{index++}",
                    Content = text
                });
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading Word DrawingML text for '{filePath}': {e.StackTrace}");
        }
    }

    private static void AddPartParagraphText(OpenXmlPackage package, string filePath, IList<GrepResult> results, string objectName, Func<OpenXmlPart, bool> predicate, string? sheetName)
    {
        try
        {
            foreach (var part in GetAllParts(package).Where(predicate))
            {
                AddDrawingParagraphText(part, filePath, sheetName, objectName, results);
            }
        }
        catch (Exception e)
        {
            LoggerHelper.Instance.LogError($"Error reading OpenXML part text for '{filePath}': {e.StackTrace}");
        }
    }

    private static IEnumerable<OpenXmlPart> GetAllParts(OpenXmlPackage package)
    {
        var visited = new HashSet<Uri>();
        foreach (var part in package.Parts.Select(p => p.OpenXmlPart))
        {
            foreach (var child in GetAllParts(part, visited))
                yield return child;
        }
    }

    private static IEnumerable<OpenXmlPart> GetAllParts(OpenXmlPart part, HashSet<Uri> visited)
    {
        if (!visited.Add(part.Uri)) yield break;

        yield return part;

        foreach (var child in part.Parts.Select(p => p.OpenXmlPart))
        {
            foreach (var descendant in GetAllParts(child, visited))
                yield return descendant;
        }
    }

    private static XDocument? TryLoadXml(OpenXmlPart part)
    {
        if (!part.ContentType.Contains("xml", StringComparison.OrdinalIgnoreCase)) return null;

        try
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            if (stream.Length == 0) return null;
            return XDocument.Load(stream);
        }
        catch
        {
            return null;
        }
    }

    private static IEnumerable<string> GetDrawingParagraphTexts(XContainer container)
    {
        foreach (var paragraph in container.Descendants().Where(e => e.Name.LocalName == "p" && IsDrawingNamespace(e.Name.NamespaceName)))
        {
            var text = string.Concat(paragraph.Descendants().Where(e => e.Name.LocalName == "t").Select(e => e.Value));
            if (string.IsNullOrWhiteSpace(text))
                text = paragraph.Value;

            if (!string.IsNullOrWhiteSpace(text))
                yield return text;
        }
    }

    private static IEnumerable<string> GetWordParagraphTexts(XContainer container)
    {
        foreach (var paragraph in container.Descendants().Where(e => e.Name.LocalName == "p" && e.Name.NamespaceName.Contains("wordprocessingml", StringComparison.OrdinalIgnoreCase)))
        {
            var text = string.Concat(paragraph.Descendants().Where(e => e.Name.LocalName == "t").Select(e => e.Value));
            if (!string.IsNullOrWhiteSpace(text))
                yield return text;
        }
    }

    private static bool IsDrawingNamespace(string namespaceName)
    {
        return namespaceName.Contains("drawingml", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsNonVisualDrawingProperties(XElement element)
    {
        return element.Name.LocalName is "docPr" or "cNvPr";
    }

    private static void AddAttributeValue(XElement element, string name, IList<string> values)
    {
        var value = element.Attributes().FirstOrDefault(a => a.Name.LocalName == name)?.Value;
        if (!string.IsNullOrWhiteSpace(value))
            values.Add(value);
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

    private static string CleanExcelHeaderFooter(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;

        return text
            .Replace("&L", " ", StringComparison.OrdinalIgnoreCase)
            .Replace("&C", " ", StringComparison.OrdinalIgnoreCase)
            .Replace("&R", " ", StringComparison.OrdinalIgnoreCase)
            .Replace("&P", " ", StringComparison.OrdinalIgnoreCase)
            .Replace("&N", " ", StringComparison.OrdinalIgnoreCase)
            .Replace("&D", " ", StringComparison.OrdinalIgnoreCase)
            .Replace("&T", " ", StringComparison.OrdinalIgnoreCase)
            .Replace("&F", " ", StringComparison.OrdinalIgnoreCase)
            .Replace("&A", " ", StringComparison.OrdinalIgnoreCase)
            .Trim();
    }
}
