using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using MarkdownToDocx.Core;
using Xunit;

namespace MarkdownToDocx.Tests;

public sealed class MarkdownConverterTests : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "MarkdownToDocx.Tests", Guid.NewGuid().ToString("N"));
    private readonly MarkdownConverter _converter = new();

    public MarkdownConverterTests() => Directory.CreateDirectory(_folder);
    private string FilePath(string name) => Path.Combine(_folder, name);
    private async Task<string> Source(string content, string name = "Quelle mit Ümlaut.md")
    {
        var path = FilePath(name);
        await File.WriteAllTextAsync(path, content, new UTF8Encoding(false));
        return path;
    }

    [Fact]
    public async Task RichMarkdownProducesValidWordDocument()
    {
        var input = await Source("""
            # Projektübersicht

            Ein **wichtiger** und *formatierter* Text mit ä, ö, ü und 😀.

            ## Aufgaben

            - Erster Punkt
            - Zweiter Punkt

            1. Einleitung
            2. Abschluss

            | Name | Status |
            | --- | --- |
            | Bericht | Fertig |

            > Ein Zitat

            [Avalonia](https://avaloniaui.net)

            ```csharp
            Console.WriteLine("Hallo");
            ```
            """);
        var output = MarkdownConverter.GetDefaultOutputPath(input);

        var result = await _converter.ConvertAsync(input, output);

        Assert.Null(result.Warning);
        Assert.Equal(output, result.OutputPath);
        Assert.True(result.Bytes > 1000);
        using var package = WordprocessingDocument.Open(output, false);
        var main = package.MainDocumentPart!;
        var body = main.Document!;
        Assert.Contains("Projektübersicht", body.InnerText);
        Assert.Contains("ä, ö, ü und 😀", body.InnerText);
        Assert.Contains("Console.WriteLine", body.InnerText);
        Assert.Single(body.Descendants<Table>());
        Assert.NotEmpty(body.Descendants<Bold>());
        Assert.NotEmpty(body.Descendants<Italic>());
        Assert.NotNull(main.NumberingDefinitionsPart);
        Assert.Contains(main.HyperlinkRelationships, link => link.Uri.AbsoluteUri == "https://avaloniaui.net/");
        Assert.Contains(body.Descendants<ParagraphStyleId>(), style => style.Val?.Value?.Contains("Heading", StringComparison.OrdinalIgnoreCase) == true);
        var errors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2013)
            .Validate(package).Select(error => error.Description + " at " + error.Path?.XPath).ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
        Assert.Empty(Directory.GetFiles(_folder, ".markdown-docx-*.tmp"));
    }

    [Fact]
    public async Task CustomDestinationIsUsedAndSourceIsUntouched()
    {
        var input = await Source("# Unverändert");
        var before = await File.ReadAllBytesAsync(input);
        var destination = Directory.CreateDirectory(FilePath("Anderer Ordner"));
        var output = Path.Combine(destination.FullName, "Bericht.docx");
        await _converter.ConvertAsync(input, output);
        Assert.True(File.Exists(output));
        Assert.False(File.Exists(MarkdownConverter.GetDefaultOutputPath(input)));
        Assert.Equal(before, await File.ReadAllBytesAsync(input));
    }

    [Fact]
    public async Task ExistingDestinationRequiresExplicitOverwrite()
    {
        var input = await Source("# Neu");
        var output = FilePath("Bericht.docx");
        await File.WriteAllTextAsync(output, "Vorheriger Inhalt");
        await Assert.ThrowsAsync<DestinationExistsException>(() => _converter.ConvertAsync(input, output));
        Assert.Equal("Vorheriger Inhalt", await File.ReadAllTextAsync(output));
        await _converter.ConvertAsync(input, output, overwrite: true);
        using var document = WordprocessingDocument.Open(output, false);
        Assert.Contains("Neu", document.MainDocumentPart!.Document!.InnerText);
    }

    [Fact]
    public async Task FailedConversionPreservesExistingDestination()
    {
        var input = await Source("   ");
        var output = FilePath("Bericht.docx");
        await File.WriteAllTextAsync(output, "Nicht ersetzen");
        var error = await Assert.ThrowsAsync<ConversionException>(() => _converter.ConvertAsync(input, output, overwrite: true));
        Assert.Contains("leer", error.Message);
        Assert.Equal("Nicht ersetzen", await File.ReadAllTextAsync(output));
    }

    [Fact]
    public async Task CancellationPreservesExistingDestinationAndCleansTemporaryFiles()
    {
        var input = await Source("# Hallo");
        var output = FilePath("Bericht.docx");
        await File.WriteAllTextAsync(output, "Nicht ersetzen");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => _converter.ConvertAsync(input, output, true, cancellation.Token));
        Assert.Equal("Nicht ersetzen", await File.ReadAllTextAsync(output));
        Assert.Empty(Directory.GetFiles(_folder, ".markdown-docx-*.tmp"));
    }

    [Fact]
    public async Task LockedDestinationProducesActionableErrorAndPreservesOriginal()
    {
        if (!OperatingSystem.IsWindows()) return;
        var input = await Source("# Neu");
        var output = FilePath("Bericht.docx");
        await File.WriteAllTextAsync(output, "Nicht ersetzen");
        using (var locked = new FileStream(output, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            var error = await Assert.ThrowsAsync<ConversionException>(() => _converter.ConvertAsync(input, output, true));
            Assert.Contains("Schliesse", error.Message);
        }
        Assert.Equal("Nicht ersetzen", await File.ReadAllTextAsync(output));
        Assert.Empty(Directory.GetFiles(_folder, ".markdown-docx-*.tmp"));
    }

    [Fact]
    public async Task RelativeImageIsEmbeddedFromSourceDirectory()
    {
        await File.WriteAllBytesAsync(FilePath("bild.png"), Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII="));
        var input = await Source("# Bild\n\n![Ein Pixel](bild.png)");
        var result = await _converter.ConvertAsync(input, FilePath("Bild.docx"));
        Assert.Null(result.Warning);
        using var document = WordprocessingDocument.Open(result.OutputPath, false);
        Assert.Single(document.MainDocumentPart!.ImageParts);
        Assert.Single(document.MainDocumentPart.Document!.Descendants<Drawing>());
    }

    [Fact]
    public async Task MissingImageIsReportedInsteadOfSilentSuccess()
    {
        var input = await Source("# Bild\n\n![Fehlendes Bild](nicht-vorhanden.png)");
        var result = await _converter.ConvertAsync(input, FilePath("Bild.docx"));
        Assert.NotNull(result.Warning);
        Assert.Contains("Bild", result.Warning);
        Assert.True(File.Exists(result.OutputPath));
    }

    [Fact]
    public async Task InvalidUtf8ProducesEncodingMessage()
    {
        var input = FilePath("Zeichen.md");
        await File.WriteAllBytesAsync(input, [0xC3, 0x28]);
        var error = await Assert.ThrowsAsync<ConversionException>(() => _converter.ConvertAsync(input, FilePath("Ausgabe.docx")));
        Assert.Contains("UTF-8", error.Message);
    }

    [Theory]
    [InlineData("", "leer")]
    [InlineData(" \r\n \t", "leer")]
    public async Task EmptyInputProducesClearError(string content, string message)
    {
        var input = await Source(content);
        var error = await Assert.ThrowsAsync<ConversionException>(() => _converter.ConvertAsync(input, FilePath("Ausgabe.docx")));
        Assert.Contains(message, error.Message);
        Assert.False(File.Exists(FilePath("Ausgabe.docx")));
    }

    [Fact]
    public async Task OversizedInputIsRejectedBeforeParsing()
    {
        var input = FilePath("Gross.md");
        using (var file = File.Create(input)) file.SetLength(MarkdownConverter.MaxInputBytes + 1);
        var error = await Assert.ThrowsAsync<ConversionException>(() => _converter.ConvertAsync(input, FilePath("Ausgabe.docx")));
        Assert.Contains("20 MB", error.Message);
    }

    [Fact]
    public async Task InvalidPathsProduceUsefulMessages()
    {
        var input = await Source("# Hallo");
        Assert.Contains("nicht gefunden", Assert.Throws<ConversionException>(() => MarkdownConverter.ValidatePaths(FilePath("Fehlt.md"), FilePath("Ausgabe.docx"))).Message);
        Assert.Contains(".docx", Assert.Throws<ConversionException>(() => MarkdownConverter.ValidatePaths(input, FilePath("Ausgabe.pdf"))).Message);
        Assert.Contains("Zielordner", Assert.Throws<ConversionException>(() => MarkdownConverter.ValidatePaths(input, FilePath("Fehlt/Ausgabe.docx"))).Message);
        Assert.Contains("Markdown-Datei", Assert.Throws<ConversionException>(() => MarkdownConverter.ValidatePaths(FilePath("Text.txt"), FilePath("Ausgabe.docx"))).Message);
    }

    public void Dispose() => Directory.Delete(_folder, recursive: true);
}
