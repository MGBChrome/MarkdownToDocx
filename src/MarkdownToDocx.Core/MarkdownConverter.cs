using System.Net.Http;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using HtmlToOpenXml.IO;
using Markdig;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace MarkdownToDocx.Core;

public sealed record ConversionResult(string OutputPath, long Bytes, string? Warning);

public sealed class ConversionException(string message, Exception? innerException = null)
    : Exception(message, innerException);

public sealed class DestinationExistsException(string path)
    : IOException($"Die Datei „{Path.GetFileName(path)}“ existiert bereits.");

/// <summary>Converts locally, then commits the completed document in one file move.</summary>
public sealed class MarkdownConverter
{
    public const long MaxInputBytes = 20 * 1024 * 1024;

    private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder()
        .UsePipeTables()
        .UseEmphasisExtras()
        .UseAutoLinks()
        .DisableHtml()
        .Build();

    private static readonly HttpClient ImageClient = new() { Timeout = TimeSpan.FromSeconds(20) };

    public static string GetDefaultOutputPath(string inputPath) => Path.ChangeExtension(inputPath, ".docx");

    public static (string Input, string Output) ValidatePaths(string inputPath, string outputPath)
    {
        if (string.IsNullOrWhiteSpace(inputPath))
            throw new ConversionException("Bitte wähle zuerst eine Markdown-Datei aus.");
        string input;
        try
        {
            input = Path.GetFullPath(inputPath.Trim());
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            throw new ConversionException("Ein Dateipfad ist ungültig oder zu lang. Bitte wähle die Datei über „Browse …“ aus.", ex);
        }

        if (!new[] { ".md", ".markdown", ".mdown" }.Contains(Path.GetExtension(input), StringComparer.OrdinalIgnoreCase))
            throw new ConversionException("Bitte wähle eine Markdown-Datei mit der Endung .md, .markdown oder .mdown aus.");
        if (!File.Exists(input))
            throw new ConversionException($"Die Markdown-Datei wurde nicht gefunden:\n{input}\nBitte prüfe den Pfad und wähle die Datei erneut aus.");

        return (input, ValidateOutputPath(outputPath));
    }

    public static string ValidateOutputPath(string outputPath)
    {
        if (string.IsNullOrWhiteSpace(outputPath))
            throw new ConversionException("Bitte wähle einen Speicherort für das Word-Dokument aus.");
        string output;
        try { output = Path.GetFullPath(outputPath.Trim()); }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            throw new ConversionException("Der Zielpfad ist ungültig oder zu lang. Bitte wähle den Speicherort über „Browse …“ aus.", ex);
        }
        if (!string.Equals(Path.GetExtension(output), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new ConversionException("Der Name der Ausgabedatei muss auf .docx enden.");
        if (!Directory.Exists(Path.GetDirectoryName(output)))
            throw new ConversionException("Der Zielordner existiert nicht. Bitte wähle über „Browse …“ einen vorhandenen Ordner aus.");
        if (Directory.Exists(output))
            throw new ConversionException("Der Ausgabepfad bezeichnet einen Ordner. Bitte gib einen Dateinamen mit .docx an.");

        return output;
    }

    public async Task<ConversionResult> ConvertAsync(string inputPath, string outputPath,
        bool overwrite = false, CancellationToken cancellationToken = default)
    {
        var (input, output) = ValidatePaths(inputPath, outputPath);
        if (File.Exists(output) && !overwrite)
            throw new DestinationExistsException(output);

        string markdown;
        try
        {
            await using var stream = new FileStream(input, FileMode.Open, FileAccess.Read, FileShare.Read);
            if (stream.Length > MaxInputBytes)
                throw new ConversionException("Die Markdown-Datei ist grösser als 20 MB. Bitte teile sie vor der Konvertierung in kleinere Dateien auf.");
            using var reader = new StreamReader(stream, new UTF8Encoding(false, true), detectEncodingFromByteOrderMarks: true);
            markdown = await reader.ReadToEndAsync(cancellationToken);
        }
        catch (DecoderFallbackException ex)
        {
            throw new ConversionException("Die Zeichenkodierung konnte nicht gelesen werden. Bitte speichere die Markdown-Datei als UTF-8 und versuche es erneut.", ex);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new ConversionException("Die Markdown-Datei konnte nicht gelesen werden. Prüfe deine Zugriffsrechte und ob die Datei von einem anderen Programm gesperrt ist.", ex);
        }

        if (string.IsNullOrWhiteSpace(markdown))
            throw new ConversionException("Die Markdown-Datei ist leer. Füge zuerst etwas Inhalt hinzu.");

        return await WriteDocumentAsync(markdown, output, Path.GetDirectoryName(input), overwrite, cancellationToken);
    }

    /// <summary>Converts in-memory Markdown without creating an intermediate source file.</summary>
    public Task<ConversionResult> ConvertTextAsync(string markdown, string outputPath,
        bool overwrite = false, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ValidateMarkdownText(markdown);
        var output = ValidateOutputPath(outputPath);
        if (File.Exists(output) && !overwrite)
            throw new DestinationExistsException(output);
        return WriteDocumentAsync(markdown, output, null, overwrite, cancellationToken);
    }

    public static void ValidateMarkdownText(string? markdown)
    {
        if (string.IsNullOrWhiteSpace(markdown))
            throw new ConversionException("Es ist noch kein Markdown-Text vorhanden. Kopiere Text und füge ihn aus der Zwischenablage ein.");
        if (Encoding.UTF8.GetByteCount(markdown) > MaxInputBytes)
            throw new ConversionException("Der Markdown-Text ist grösser als 20 MB. Bitte füge einen kleineren Ausschnitt ein.");
    }

    private static async Task<ConversionResult> WriteDocumentAsync(string markdown, string output,
        string? sourceDirectory, bool overwrite, CancellationToken cancellationToken)
    {
        // A temporary file in the destination directory keeps the final move on the same volume.
        var temporaryPath = Path.Combine(Path.GetDirectoryName(output)!, $".markdown-docx-{Guid.NewGuid():N}.tmp");
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var document = Markdown.Parse(markdown, Pipeline);
            var images = document.Descendants<LinkInline>().Where(link => link.IsImage).ToArray();
            var hasRelativeImages = false;
            if (sourceDirectory is null)
            {
                foreach (var image in images.Where(image => !Uri.TryCreate(image.Url, UriKind.Absolute, out _)))
                {
                    // Clipboard text has no source folder. Keep the alt text, without guessing a file location.
                    image.IsImage = false;
                    image.Url = null;
                    hasRelativeImages = true;
                }
            }
            var html = document.ToHtml(Pipeline);
            string? warning = null;

            using (var package = WordprocessingDocument.Create(temporaryPath, WordprocessingDocumentType.Document))
            {
                var main = package.AddMainDocumentPart();
                main.Document = new Document(new Body());
                var settings = main.AddNewPart<DocumentSettingsPart>();
                settings.Settings = new Settings(new Compatibility(new CompatibilitySetting
                {
                    Name = CompatSettingNameValues.CompatibilityMode,
                    Uri = "http://schemas.microsoft.com/office/word",
                    Val = "15"
                }));
                main.AddNewPart<StyleDefinitionsPart>().Styles = WordDocumentDesign.CreateStyles();

                var resources = new DefaultWebRequest(ImageClient);
                if (sourceDirectory is not null)
                    resources.BaseImageUrl = new Uri(sourceDirectory + Path.DirectorySeparatorChar);
                var converter = new HtmlConverter(main, resources)
                {
                    // Markdown heading numbers are content, not a request for a Word outline list.
                    SupportsHeadingNumbering = false,
                    ContinueNumbering = false
                };
                await converter.ParseBody(html, cancellationToken);
                WordDocumentDesign.FormatContent(main);

                if (main.Document.Descendants<Drawing>().Count() < images.Length)
                    warning = hasRelativeImages
                        ? "Mindestens ein Bild konnte nicht eingebettet werden. Relative Bildpfade haben bei Text aus der Zwischenablage keinen Quellordner. Verwende vollständige Bildadressen oder konvertiere eine Markdown-Datei mit zugehörigen Bildern."
                        : "Mindestens ein Bild konnte nicht eingebettet werden. Prüfe die Bildpfade, das Bildformat und bei Online-Bildern die Internetverbindung.";

                main.Document.Body!.Append(new SectionProperties(
                    new PageSize { Width = 11906, Height = 16838 },
                    new PageMargin { Top = 1134, Right = 1134, Bottom = 1134, Left = 1134, Header = 567, Footer = 567, Gutter = 0 }));
                main.Document.Save();
            }

            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                File.Move(temporaryPath, output, overwrite);
            }
            catch (IOException) when (!overwrite && File.Exists(output))
            {
                throw new DestinationExistsException(output);
            }

            return new ConversionResult(output, new FileInfo(output).Length, warning);
        }
        catch (OperationCanceledException) { throw; }
        catch (DestinationExistsException) { throw; }
        catch (UnauthorizedAccessException ex)
        {
            throw new ConversionException("Der Zugriff auf die Zieldatei wurde verweigert. Schliesse eine bereits geöffnete Datei in Word und prüfe Schreibschutz und Zugriffsrechte. Du kannst auch einen anderen Zielordner wählen.", ex);
        }
        catch (IOException ex)
        {
            throw new ConversionException("Das Word-Dokument konnte nicht gespeichert werden. Schliesse eine bereits geöffnete Zieldatei in Word und prüfe Speicherplatz und Zugriffsrechte.", ex);
        }
        catch (Exception ex) when (ex is not ConversionException)
        {
            throw new ConversionException("Der Inhalt konnte nicht vollständig in ein Word-Dokument umgewandelt werden. Prüfe insbesondere eingebettete Bilder und den Markdown-Text.", ex);
        }
        finally
        {
            // Cleanup must never hide the original error or turn a successful conversion into an error.
            try { File.Delete(temporaryPath); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
