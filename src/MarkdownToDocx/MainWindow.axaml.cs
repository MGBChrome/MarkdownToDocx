using System.Diagnostics;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Platform.Storage;
using Avalonia.Interactivity;
using Avalonia.Input.Platform;
using MarkdownToDocx.Core;

namespace MarkdownToDocx;

public partial class MainWindow : Window
{
    private readonly MainViewModel _model = new();
    private readonly MarkdownConverter _converter = new();
    private CancellationTokenSource? _conversionCancellation;

    private static readonly FilePickerFileType MarkdownType = new("Markdown-Dateien")
    {
        Patterns = ["*.md", "*.markdown", "*.mdown"],
        MimeTypes = ["text/markdown", "text/plain"]
    };
    private static readonly FilePickerFileType WordType = new("Word-Dokument") { Patterns = ["*.docx"] };

    public MainWindow()
    {
        InitializeComponent();
        DataContext = _model;
        Closing += (_, _) => _conversionCancellation?.Cancel();
    }

    private async void Paste_Click(object? sender, RoutedEventArgs e)
    {
        if (!_model.IsIdle) return;
        _model.IsReadingClipboard = true;
        var pasted = false;
        try
        {
            var clipboard = Clipboard ?? throw new ConversionException("Die Zwischenablage ist derzeit nicht verfügbar. Du kannst Text auch direkt in die Vorschau eingeben.");
            pasted = _model.TryPasteMarkdown(await clipboard.TryGetTextAsync());
        }
        catch (Exception ex)
        {
            ShowError(ex, "Die Zwischenablage konnte nicht gelesen werden. Kopiere den Text erneut und versuche es noch einmal, oder füge ihn mit Strg+V in die Vorschau ein. Deine bisherige Vorschau bleibt erhalten.");
        }
        finally { _model.IsReadingClipboard = false; }
        if (pasted)
        {
            MarkdownPreviewBox.Focus();
            MarkdownPreviewBox.CaretIndex = 0;
        }
    }

    private void ClearPreview_Click(object? sender, RoutedEventArgs e)
    {
        if (!_model.CanClearPreview) return;
        _model.MarkdownText = "";
        _model.ShowStatus("Vorschau geleert.", "Füge neuen Markdown-Text ein, um ein weiteres Word-Dokument zu erstellen.");
        MarkdownPreviewBox.Focus();
    }

    private async void BrowseInput_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "Markdown-Datei auswählen",
                AllowMultiple = false,
                FileTypeFilter = [MarkdownType, FilePickerFileTypes.All],
                SuggestedStartLocation = await GetFolderAsync(_model.InputPath)
            });
            if (files.Count == 0) return;
            using var file = files[0];
            _model.InputPath = file.TryGetLocalPath()
                ?? throw new ConversionException("Bitte wähle eine lokal verfügbare Markdown-Datei aus.");
        }
        catch (Exception ex) { ShowError(ex, "Die Dateiauswahl konnte nicht geöffnet werden. Du kannst den vollständigen Pfad auch direkt eingeben."); }
    }

    private async void BrowseOutput_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            using var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Word-Dokument speichern unter",
                SuggestedFileName = string.IsNullOrWhiteSpace(_model.OutputPath) ? "Dokument.docx" : Path.GetFileName(_model.OutputPath),
                DefaultExtension = "docx",
                FileTypeChoices = [WordType],
                SuggestedStartLocation = await GetFolderAsync(_model.OutputPath),
                // Confirm immediately before conversion, including paths typed directly into the UI.
                ShowOverwritePrompt = false
            });
            if (file is null) return;
            var path = file.TryGetLocalPath()
                ?? throw new ConversionException("Bitte wähle einen lokal verfügbaren Speicherort aus.");
            if (_model.IsFileSource) _model.UseSourceFolder = false;
            _model.OutputPath = path;
        }
        catch (Exception ex) { ShowError(ex, "Der Speicherdialog konnte nicht geöffnet werden. Gib den Zielpfad direkt ein; deaktiviere bei einer Dateiquelle gegebenenfalls die Ordneroption."); }
    }

    private async Task<IStorageFolder?> GetFolderAsync(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return null;
        try
        {
            var folder = Path.GetDirectoryName(Path.GetFullPath(path));
            return Directory.Exists(folder) ? await StorageProvider.TryGetFolderFromPathAsync(folder) : null;
        }
        catch (ArgumentException) { return null; }
        catch (NotSupportedException) { return null; }
        catch (PathTooLongException) { return null; }
    }

    private async void Convert_Click(object? sender, RoutedEventArgs e)
    {
        if (!_model.CanConvert) return;
        _model.IsBusy = true;
        _model.CompletedPath = null;
        try
        {
            var isClipboard = _model.IsClipboardSource;
            var markdown = _model.MarkdownText;
            var input = _model.InputPath;
            var output = isClipboard
                ? MarkdownConverter.ValidateOutputPath(_model.OutputPath)
                : MarkdownConverter.ValidatePaths(input, _model.OutputPath).Output;
            if (isClipboard) MarkdownConverter.ValidateMarkdownText(markdown);
            var overwrite = File.Exists(output);
            if (overwrite && !await ConfirmOverwriteAsync(output))
            {
                _model.ShowStatus("Datei bleibt unverändert.", "Wähle einen anderen Namen oder Speicherort, um ein neues Dokument zu erstellen.");
                return;
            }

            _conversionCancellation = new CancellationTokenSource();
            var cancellationToken = _conversionCancellation.Token;
            _model.ShowStatus("Dein Word-Dokument entsteht …", "Inhalt und Formatierung werden übernommen. Bei Online-Bildern kann es etwas länger dauern.");
            var result = await Task.Run(() => isClipboard
                ? _converter.ConvertTextAsync(markdown, output, overwrite, cancellationToken)
                : _converter.ConvertAsync(input, output, overwrite, cancellationToken), cancellationToken);
            _model.CompletedPath = result.OutputPath;
            var size = Math.Max(1, (long)Math.Ceiling(result.Bytes / 1024d));
            _model.ShowStatus(result.Warning is null ? "Fertig! Dein Word-Dokument ist bereit." : "Dokument erstellt – bitte Bilder prüfen.",
                $"{result.OutputPath}\n{size:N0} KB · " + (result.Warning ?? "Du kannst die Datei jetzt direkt öffnen."),
                result.Warning is null ? "success" : "warning");
        }
        catch (OperationCanceledException)
        {
            _model.ShowStatus("Konvertierung abgebrochen.", "Es wurde keine neue Word-Datei gespeichert. Du kannst die Konvertierung erneut starten.");
        }
        catch (DestinationExistsException)
        {
            _model.ShowStatus("Die Zieldatei existiert inzwischen.", "Starte die Konvertierung erneut, um das Überschreiben zu bestätigen, oder wähle einen anderen Dateinamen.", "warning");
        }
        catch (Exception ex) { ShowError(ex, "Die Konvertierung ist unerwartet fehlgeschlagen. Bitte versuche es erneut oder wähle eine andere Datei."); }
        finally
        {
            _conversionCancellation?.Dispose();
            _conversionCancellation = null;
            _model.IsBusy = false;
        }
    }

    private void Cancel_Click(object? sender, RoutedEventArgs e) => _conversionCancellation?.Cancel();

    private void Open_Click(object? sender, RoutedEventArgs e)
    {
        if (_model.CompletedPath is not { } path) return;
        try
        {
            if (!File.Exists(path))
            {
                _model.CompletedPath = null;
                throw new ConversionException("Das erstellte Dokument wurde verschoben oder gelöscht. Bitte konvertiere die Datei erneut.");
            }
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            ShowError(ex, "Das Dokument wurde gespeichert, konnte aber nicht geöffnet werden. Richte eine Standardanwendung für .docx ein, z. B. Microsoft Word oder LibreOffice.\n" + path);
        }
    }

    private async Task<bool> ConfirmOverwriteAsync(string output)
    {
        var dialog = new Window
        {
            Title = "Vorhandene Datei ersetzen?",
            Width = 520, SizeToContent = SizeToContent.Height, CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ShowInTaskbar = false, Background = Avalonia.Media.Brushes.White
        };
        var cancel = new Button { Content = "Abbrechen", IsCancel = true };
        var replace = new Button { Content = "Datei ersetzen", Classes = { "primary" } };
        cancel.Click += (_, _) => dialog.Close(false);
        replace.Click += (_, _) => dialog.Close(true);
        dialog.Content = new StackPanel
        {
            Margin = new Thickness(28), Spacing = 16,
            Children =
            {
                new TextBlock { Text = "Vorhandene Datei ersetzen?", FontSize = 21, FontWeight = Avalonia.Media.FontWeight.SemiBold },
                new TextBlock { Text = "An diesem Speicherort gibt es bereits ein Dokument. Beim Fortfahren wird es durch die neue Word-Datei ersetzt.", TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                new SelectableTextBlock { Text = output, TextWrapping = Avalonia.Media.TextWrapping.Wrap, FontSize = 12 },
                new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Spacing = 10, Children = { cancel, replace } }
            }
        };
        dialog.Opened += (_, _) => cancel.Focus();
        return await dialog.ShowDialog<bool>(this);
    }

    private void ShowError(Exception exception, string fallback)
    {
        Trace.TraceError(exception.ToString());
        _model.ShowStatus("Das hat leider nicht geklappt.", exception is ConversionException ? exception.Message : fallback, "error");
    }
}
