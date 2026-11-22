using Xunit;

namespace MarkdownToDocx.Tests;

public sealed class MainViewModelTests
{
    [Fact]
    public void SwitchingSourcesPreservesTextAndIndependentDestinations()
    {
        var model = new MainViewModel(Path.GetTempPath()) { InputPath = "Datei.md" };
        model.SourceIndex = 1;
        Assert.False(model.CanConvert);
        Assert.False(model.IsOutputReadOnly);
        Assert.True(model.TryPasteMarkdown("# Meine Notizen\n\nHallo"));
        Assert.Equal(Path.Combine(Path.GetTempPath(), "Meine Notizen.docx"), model.OutputPath);
        model.OutputPath = "Eigener Name.docx";
        model.SourceIndex = 0;
        Assert.Equal("Datei.docx", model.OutputPath);
        Assert.True(model.UseSourceFolder);
        model.InputPath = "Andere.md";
        model.SourceIndex = 1;
        Assert.Equal("Eigener Name.docx", model.OutputPath);
        Assert.Equal("# Meine Notizen\n\nHallo", model.MarkdownText);
        model.TryPasteMarkdown("# Neuer Text");
        Assert.Equal("Eigener Name.docx", model.OutputPath);
        Assert.True(model.CanConvert);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" \r\n\t")]
    public void EmptyClipboardPreservesPreviewAndCompletedDocument(string? text)
    {
        var model = new MainViewModel { SourceIndex = 1, MarkdownText = "# Behalten", CompletedPath = "Fertig.docx" };
        Assert.False(model.TryPasteMarkdown(text));
        Assert.Equal("# Behalten", model.MarkdownText);
        Assert.True(model.CanOpen);
        Assert.Contains("Kein Text", model.StatusTitle);
    }

    [Fact]
    public void OversizedPastePreservesPreviewAndBusyStateDisablesConversion()
    {
        var model = new MainViewModel { SourceIndex = 1, MarkdownText = "# Behalten" };
        Assert.False(model.TryPasteMarkdown(new string('ä', 10 * 1024 * 1024 + 1)));
        Assert.Equal("# Behalten", model.MarkdownText);
        Assert.Contains("20 MB", model.StatusMessage);
        model.IsReadingClipboard = true;
        Assert.False(model.CanConvert);
        Assert.False(model.IsIdle);
        model.IsReadingClipboard = false;
        model.IsBusy = true;
        Assert.False(model.CanConvert);
        model.IsBusy = false;
        Assert.True(model.CanConvert);
    }

    [Fact]
    public void EditingPreviewOrChangingSourceInvalidatesPreviousDocument()
    {
        var model = new MainViewModel { SourceIndex = 1, MarkdownText = "# Eins", CompletedPath = "Eins.docx" };
        model.MarkdownText = "# Zwei";
        Assert.False(model.CanOpen);
        Assert.EndsWith("Zwei.docx", model.OutputPath);
        model.CompletedPath = "Zwei.docx";
        model.SourceIndex = 0;
        Assert.False(model.CanOpen);
    }

    [Theory]
    [InlineData("# Reise: Schweiz / Italien?", "Reise Schweiz  Italien.docx")]
    [InlineData("# CON", "Dokument-CON.docx")]
    [InlineData("# LPT1", "Dokument-LPT1.docx")]
    [InlineData("# ../", "Dokument.docx")]
    [InlineData("## **Notizen** ##", "Notizen.docx")]
    [InlineData("Normaler Text", "Dokument.docx")]
    [InlineData("#Titel", "Dokument.docx")]
    public void SuggestedFileNameIsUsableForWindows(string markdown, string expected)
    {
        var model = new MainViewModel(Path.GetTempPath()) { SourceIndex = 1, MarkdownText = markdown };
        Assert.Equal(Path.Combine(Path.GetTempPath(), expected), model.OutputPath);
    }

    [Fact]
    public void DefaultDestinationFollowsSource()
    {
        var model = new MainViewModel { InputPath = Path.Combine(Path.GetTempPath(), "Eins.md") };
        Assert.Equal(Path.ChangeExtension(model.InputPath, ".docx"), model.OutputPath);
        model.InputPath = Path.Combine(Path.GetTempPath(), "Zwei.markdown");
        Assert.Equal(Path.ChangeExtension(model.InputPath, ".docx"), model.OutputPath);
        Assert.True(model.CanConvert);
    }

    [Fact]
    public void CustomDestinationPersistsUntilDefaultIsSelectedAgain()
    {
        var model = new MainViewModel { UseSourceFolder = false, OutputPath = "Bericht.docx", InputPath = "Notizen.md" };
        Assert.Equal("Bericht.docx", model.OutputPath);
        model.UseSourceFolder = true;
        Assert.Equal("Notizen.docx", model.OutputPath);
    }

    [Fact]
    public void EditingPathsInvalidatesPreviousSuccess()
    {
        var model = new MainViewModel { InputPath = "Eins.md", CompletedPath = "Eins.docx" };
        Assert.True(model.CanOpen);
        model.InputPath = "Zwei.md";
        Assert.False(model.CanOpen);
        Assert.Null(model.CompletedPath);
    }

    [Fact]
    public void BusyStateDisablesActionsAndEmptySourceCannotConvert()
    {
        var model = new MainViewModel();
        Assert.False(model.CanConvert);
        model.InputPath = "Notizen.md";
        model.CompletedPath = "Notizen.docx";
        model.IsBusy = true;
        Assert.False(model.CanConvert);
        Assert.False(model.CanOpen);
        Assert.False(model.IsIdle);
        model.IsBusy = false;
        Assert.True(model.CanConvert);
        Assert.True(model.CanOpen);
    }
}
