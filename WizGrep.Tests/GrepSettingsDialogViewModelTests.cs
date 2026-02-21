using Microsoft.VisualStudio.TestTools.UnitTesting;
using WizGrep.ViewModels;

namespace WizGrep.Tests;

[TestClass]
public class GrepSettingsDialogViewModelTests
{
    [TestMethod]
    public void Validate_ReturnsFalse_WhenTargetFolderIsEmpty()
    {
        var viewModel = new GrepSettingsDialogViewModel
        {
            Keyword0 = "keyword",
            KeywordEnabled0 = true
        };

        Assert.IsFalse(viewModel.Validate());
    }

    [TestMethod]
    public void Validate_ReturnsTrue_WhenTargetFolderAndKeywordAreSet()
    {
        var viewModel = new GrepSettingsDialogViewModel
        {
            TargetFolderPath = "C:\\Temp",
            Keyword0 = "keyword",
            KeywordEnabled0 = true
        };

        Assert.IsTrue(viewModel.Validate());
    }
}
