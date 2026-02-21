using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using WizGrep.Models;
using WizGrep.Services;

namespace WizGrep.Tests;

[TestClass]
public class GrepServiceTests
{
    [TestMethod]
    public async Task ExecuteGrepAsync_ExcludesSpecifiedFolder()
    {
        var tempRoot = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        var excludedFolder = Path.Combine(tempRoot, "excluded");
        Directory.CreateDirectory(excludedFolder);

        var rootFile = Path.Combine(tempRoot, "root.txt");
        var excludedFile = Path.Combine(excludedFolder, "excluded.txt");
        await File.WriteAllTextAsync(rootFile, "hello");
        await File.WriteAllTextAsync(excludedFile, "hello");

        try
        {
            var grepSettings = CreateSettings(tempRoot, "hello");
            grepSettings.ExcludeFolders = excludedFolder;

            var results = await CreateService().ExecuteGrepAsync(grepSettings, new WizGrepSettings());

            Assert.AreEqual(1, results.Count);
            Assert.AreEqual(rootFile, results.Single().FilePath);
        }
        finally
        {
            Directory.Delete(tempRoot, true);
        }
    }

    [TestMethod]
    public async Task ExecuteGrepAsync_UsesRegexWhenEnabled()
    {
        var tempRoot = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        var rootFile = Path.Combine(tempRoot, "root.txt");
        Directory.CreateDirectory(tempRoot);
        await File.WriteAllTextAsync(rootFile, "hello");

        try
        {
            var grepSettings = CreateSettings(tempRoot, "h.llo");
            grepSettings.UseRegex = true;

            var results = await CreateService().ExecuteGrepAsync(grepSettings, new WizGrepSettings());

            Assert.AreEqual(1, results.Count);
            Assert.AreEqual(rootFile, results.Single().FilePath);
        }
        finally
        {
            Directory.Delete(tempRoot, true);
        }
    }

    [TestMethod]
    public async Task ExecuteGrepAsync_ThrowsForInvalidRegex()
    {
        var tempRoot = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        var rootFile = Path.Combine(tempRoot, "root.txt");
        Directory.CreateDirectory(tempRoot);
        await File.WriteAllTextAsync(rootFile, "hello");

        try
        {
            var grepSettings = CreateSettings(tempRoot, "(");
            grepSettings.UseRegex = true;

            await Assert.ThrowsExceptionAsync<InvalidOperationException>(
                () => CreateService().ExecuteGrepAsync(grepSettings, new WizGrepSettings()));
        }
        finally
        {
            Directory.Delete(tempRoot, true);
        }
    }

    [TestMethod]
    public async Task ExecuteGrepAsync_ExcludesFilesByExtension()
    {
        var tempRoot = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempRoot);

        var txtFile = Path.Combine(tempRoot, "root.txt");
        var logFile = Path.Combine(tempRoot, "root.log");
        await File.WriteAllTextAsync(txtFile, "hello");
        await File.WriteAllTextAsync(logFile, "hello");

        try
        {
            var grepSettings = CreateSettings(tempRoot, "hello");
            grepSettings.UseExcludeExtensions = true;
            grepSettings.ExcludeExtensions = ".log";

            var results = await CreateService().ExecuteGrepAsync(grepSettings, new WizGrepSettings());

            Assert.AreEqual(1, results.Count);
            Assert.AreEqual(txtFile, results.Single().FilePath);
        }
        finally
        {
            Directory.Delete(tempRoot, true);
        }
    }

    [TestMethod]
    public async Task ExecuteGrepAsync_IgnoresRelativeExcludeFolders()
    {
        var tempRoot = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        var excludedFolder = Path.Combine(tempRoot, "excluded");
        Directory.CreateDirectory(excludedFolder);

        var rootFile = Path.Combine(tempRoot, "root.txt");
        var excludedFile = Path.Combine(excludedFolder, "excluded.txt");
        await File.WriteAllTextAsync(rootFile, "hello");
        await File.WriteAllTextAsync(excludedFile, "hello");

        try
        {
            var grepSettings = CreateSettings(tempRoot, "hello");
            grepSettings.ExcludeFolders = "excluded";

            var results = await CreateService().ExecuteGrepAsync(grepSettings, new WizGrepSettings());

            Assert.AreEqual(2, results.Count);
        }
        finally
        {
            Directory.Delete(tempRoot, true);
        }
    }

    private static GrepService CreateService()
    {
        return new GrepService(new FileReaderService(), new IndexService());
    }

    private static GrepSettings CreateSettings(string targetFolderPath, string keyword)
    {
        var settings = new GrepSettings
        {
            TargetFolderPath = targetFolderPath,
            IncludeExcel = false,
            IncludeWord = false,
            IncludePowerPoint = false,
            IncludePdf = false,
            IncludeText = true,
            IncludeAll = false,
            UseRegex = false,
            CaseSensitive = false
        };

        settings.Keywords[0].Keyword = keyword;
        settings.Keywords[0].IsEnabled = true;
        settings.Keywords[1].IsEnabled = false;
        settings.Keywords[2].IsEnabled = false;
        settings.Keywords[3].IsEnabled = false;
        settings.Keywords[4].IsEnabled = false;

        return settings;
    }
}
