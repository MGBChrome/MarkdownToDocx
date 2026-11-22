using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using MarkdownToDocx.Core;
using Xunit;

namespace MarkdownToDocx.Tests;

public sealed class DocumentDesignTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "MarkdownToDocx.DesignTests", Guid.NewGuid().ToString("N"));

    public DocumentDesignTests() => Directory.CreateDirectory(_folder);

    private async Task<WordprocessingDocument> Convert(string markdown)
    {
        var input = Path.Combine(_folder, "input.md");
        var output = Path.Combine(_folder, "output.docx");
        await File.WriteAllTextAsync(input, markdown);
        await new MarkdownConverter().ConvertAsync(input, output);
        return WordprocessingDocument.Open(output, false);
    }

    [Fact]
    public async Task NumberedHeadingsRetainLiteralTextWithoutInventedOutlineNumbers()
    {
        using var document = await Convert("""
            # Mein Dokument

            ## Die sechs Punkte

            ### **1. Erste Frage?**

            Erklärung.

            ### **2. Zweite Frage?**

            Erklärung.

            #### 2.1. Ein Detail

            ##### Ebene fünf

            ###### Ebene sechs
            """);
        var main = document.MainDocumentPart!;
        var headings = main.Document!.Body!.Elements<Paragraph>()
            .Where(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.StartsWith("Heading") == true).ToArray();

        Assert.Equal(new[] { "Mein Dokument", "Die sechs Punkte", "1. Erste Frage?", "2. Zweite Frage?", "2.1. Ein Detail", "Ebene fünf", "Ebene sechs" },
            headings.Select(p => p.InnerText));
        Assert.All(headings, p => Assert.Null(p.ParagraphProperties!.NumberingProperties));
        for (var level = 1; level <= 6; level++)
        {
            var style = main.StyleDefinitionsPart!.Styles!.Elements<Style>().Single(s => s.StyleId == $"Heading{level}");
            Assert.Equal("000000", style.StyleRunProperties!.Color!.Val!.Value);
            Assert.Null(style.StyleRunProperties.Color.ThemeColor);
            Assert.Equal("Arial", style.StyleRunProperties.RunFonts!.Ascii!.Value);
            Assert.Equal(level - 1, style.StyleParagraphProperties!.OutlineLevel!.Val!.Value);
            Assert.NotNull(style.StyleParagraphProperties.KeepNext);
            Assert.Null(style.StyleParagraphProperties.NumberingProperties);
        }
        AssertValid(document);
    }

    [Fact]
    public async Task SeparateOrderedListsRestartAndExplicitStartIsPreserved()
    {
        using var document = await Convert("""
            1. Eins
            2. Zwei

            Ein neuer Abschnitt.

            1. Wieder eins

            Noch ein Abschnitt.

            7. Beginnt bei sieben
            """);
        var main = document.MainDocumentPart!;
        var lists = main.Document!.Body!.Elements<Paragraph>().Where(p => p.ParagraphProperties?.NumberingProperties is not null).ToArray();
        Assert.Equal(4, lists.Length);
        var ids = lists.Select(p => p.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value).ToArray();
        Assert.Equal(ids[0], ids[1]);
        Assert.NotEqual(ids[1], ids[2]);
        Assert.NotEqual(ids[2], ids[3]);
        var numbering = main.NumberingDefinitionsPart!.Numbering!;
        int Start(int id) => numbering.Elements<NumberingInstance>().Single(n => n.NumberID!.Value == id)
            .Elements<LevelOverride>().Single(l => l.LevelIndex!.Value == 0).StartOverrideNumberingValue!.Val!.Value;
        Assert.Equal(1, Start(ids[0]));
        Assert.Equal(1, Start(ids[2]));
        Assert.Equal(7, Start(ids[3]));
        AssertValid(document);
    }

    [Fact]
    public async Task KnownFaviconExportBecomesCompactSourceNoteWithoutChangingRealLinks()
    {
        using var document = await Convert("""
            Inhalt mit **Hervorhebung**.

            [image](https://www.google.com/s2/favicons?domain=https://example.org&sz=32)

            example.org+2

            **Fragen an die Werkstatt:**

            - Erste Frage
            - Zweite Frage

            [image](https://example.org/picture.png)

            Echte Bildverknüpfung.

            [Originalbericht](https://example.org/report)
            """);
        var main = document.MainDocumentPart!;
        var paragraphs = main.Document!.Body!.Elements<Paragraph>().ToArray();
        var source = Assert.Single(paragraphs, p => p.ParagraphProperties?.ParagraphStyleId?.Val == "SourceNote");
        Assert.Equal("example.org+2", source.InnerText);
        Assert.Single(paragraphs, p => p.InnerText == "image");
        Assert.Contains(main.HyperlinkRelationships, l => l.Uri.AbsoluteUri == "https://example.org/report");
        Assert.Contains(main.HyperlinkRelationships, l => l.Uri.AbsoluteUri == "https://example.org/picture.png");
        var leadIn = paragraphs.Single(p => p.InnerText == "Fragen an die Werkstatt:");
        Assert.Equal("LeadIn", leadIn.ParagraphProperties!.ParagraphStyleId!.Val!.Value);
        var lastItem = paragraphs.Single(p => p.InnerText == "Zweite Frage");
        Assert.Equal("160", lastItem.ParagraphProperties!.SpacingBetweenLines!.After!.Value);
        AssertValid(document);
    }

    [Fact]
    public async Task FaviconLinkWithoutClearCaptionIsNotRemoved()
    {
        using var document = await Convert("""
            [image](https://www.google.com/s2/favicons?domain=https://example.org&sz=32)

            Dies ist ein eigenständiger Satz.
            """);
        Assert.Contains("image", document.MainDocumentPart!.Document!.InnerText);
        Assert.Contains("Dies ist ein eigenständiger Satz.", document.MainDocumentPart.Document.InnerText);
    }

    private static void AssertValid(WordprocessingDocument document)
    {
        var errors = new OpenXmlValidator(FileFormatVersions.Office2013).Validate(document)
            .Select(error => error.Description + " at " + error.Path?.XPath).ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
    }

    public void Dispose() => Directory.Delete(_folder, recursive: true);
}
