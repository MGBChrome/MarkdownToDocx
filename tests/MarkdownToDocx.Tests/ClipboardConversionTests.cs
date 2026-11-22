using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using MarkdownToDocx.Core;
using Xunit;

namespace MarkdownToDocx.Tests;

public sealed class ClipboardConversionTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "MarkdownToDocx.Tests", Guid.NewGuid().ToString("N"));
    private readonly MarkdownConverter _converter = new();
    public ClipboardConversionTests() => Directory.CreateDirectory(_folder);
    private string FilePath(string name) => Path.Combine(_folder, name);

    [Fact]
    public async Task TextAndFileSourcesProduceTheSameValidDocumentDesign()
    {
        const string markdown = "# Reiseplanung\n\n### 1. Fragen\n\n**Wichtig:** Grüße 😀\n\n- Segel\n- Motor\n\n| Frage | Antwort |\n| --- | --- |\n| Hafen | Ja |\n\n[Quelle](https://example.com)\n\n```csharp\nvar x = 1;\n```";
        var result = await _converter.ConvertTextAsync(markdown, FilePath("Text.docx"));
        Assert.Single(Directory.GetFiles(_folder)); // No intermediate Markdown file.
        Assert.Null(result.Warning);
        await File.WriteAllTextAsync(FilePath("Quelle.md"), markdown);
        await _converter.ConvertAsync(FilePath("Quelle.md"), FilePath("Datei.docx"));
        using var text = WordprocessingDocument.Open(result.OutputPath, false);
        using var file = WordprocessingDocument.Open(FilePath("Datei.docx"), false);
        Assert.Equal(NormalizeRelationships(file.MainDocumentPart!), NormalizeRelationships(text.MainDocumentPart!));
        Assert.Equal(file.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml, text.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml);
        Assert.Contains("1. Fragen", text.MainDocumentPart!.Document!.InnerText);
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2013).Validate(text));
    }

    [Theory]
    [InlineData("")]
    [InlineData(" \r\n\t")]
    public async Task EmptyTextPreservesExistingDestination(string markdown)
    {
        var output = FilePath("Unverändert.docx");
        await File.WriteAllTextAsync(output, "Original");
        var error = await Assert.ThrowsAsync<ConversionException>(() => _converter.ConvertTextAsync(markdown, output, true));
        Assert.Contains("Markdown-Text", error.Message);
        Assert.Equal("Original", await File.ReadAllTextAsync(output));
        Assert.Single(Directory.GetFiles(_folder));
    }

    [Fact]
    public async Task SizeLimitCountsUtf8BytesInsteadOfCharacters()
    {
        var text = new string('ä', (int)MarkdownConverter.MaxInputBytes / 2 + 1);
        var error = await Assert.ThrowsAsync<ConversionException>(() => _converter.ConvertTextAsync(text, FilePath("Gross.docx")));
        Assert.Contains("20 MB", error.Message);
        Assert.Empty(Directory.GetFiles(_folder));
    }

    [Fact]
    public async Task OverwriteRequiresConsentAndCancellationPreservesOriginal()
    {
        var output = FilePath("Vorhanden.docx");
        await File.WriteAllTextAsync(output, "Original");
        await Assert.ThrowsAsync<DestinationExistsException>(() => _converter.ConvertTextAsync("# Neu", output));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => _converter.ConvertTextAsync("# Neu", output, true, new CancellationToken(true)));
        Assert.Equal("Original", await File.ReadAllTextAsync(output));
        await _converter.ConvertTextAsync("# Neu", output, true);
        using var document = WordprocessingDocument.Open(output, false);
        Assert.Contains("Neu", document.MainDocumentPart!.Document!.InnerText);
        Assert.Empty(Directory.GetFiles(_folder, ".markdown-docx-*.tmp"));
    }

    [Theory]
    [InlineData("", "Speicherort")]
    [InlineData("Ausgabe.pdf", ".docx")]
    [InlineData("Fehlt/Ausgabe.docx", "Zielordner")]
    public async Task TextConversionValidatesDestination(string name, string message)
    {
        var error = await Assert.ThrowsAsync<ConversionException>(() => _converter.ConvertTextAsync("# Text", name == "" ? "" : FilePath(name)));
        Assert.Contains(message, error.Message);
        Assert.Empty(Directory.GetFiles(_folder));
    }

    [Fact]
    public async Task RelativeClipboardImagesKeepAltTextAndExplainMissingSourceFolder()
    {
        var image = FilePath("Bild.png");
        await File.WriteAllBytesAsync(image, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII="));
        var relativePath = Path.GetRelativePath(Environment.CurrentDirectory, image).Replace('\\', '/');
        var result = await _converter.ConvertTextAsync($"# Bild\n\n![Bildbeschreibung]({relativePath})", FilePath("Bilder.docx"));
        Assert.Contains("Quellordner", result.Warning);
        using var document = WordprocessingDocument.Open(result.OutputPath, false);
        Assert.Empty(document.MainDocumentPart!.Document!.Descendants<Drawing>());
        Assert.Contains("Bildbeschreibung", document.MainDocumentPart.Document!.InnerText);
    }

    private static string NormalizeRelationships(MainDocumentPart main)
    {
        var content = (Document)main.Document!.CloneNode(true);
        foreach (var link in content.Descendants<Hyperlink>().Where(link => link.Id is not null))
            link.Id = main.HyperlinkRelationships.Single(relation => relation.Id == link.Id!.Value).Uri.AbsoluteUri;
        return content.OuterXml;
    }

    public void Dispose() => Directory.Delete(_folder, recursive: true);
}
