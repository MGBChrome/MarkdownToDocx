using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text;
using Avalonia.Media;
using MarkdownToDocx.Core;

namespace MarkdownToDocx;

public sealed class MainViewModel : INotifyPropertyChanged
{
    private string _inputPath = "";
    private string _outputPath = "";
    private string _clipboardOutputPath;
    private readonly string _documentsFolder;
    private string _markdownText = "";
    private int _sourceIndex;
    private bool _clipboardOutputCustomized;
    private bool _isReadingClipboard;
    private int _markdownBytes;
    private bool _useSourceFolder = true;
    private bool _isBusy;
    private string? _completedPath;
    private string _statusTitle = "Bereit, wenn du es bist.";
    private string _statusMessage = "Wähle eine Markdown-Datei aus oder füge Text aus der Zwischenablage ein.";
    private string _statusSymbol = "i";
    private IBrush _statusBackground = Brush.Parse("#EEF3FC");
    private IBrush _statusForeground = Brush.Parse("#38547B");

    public event PropertyChangedEventHandler? PropertyChanged;

    public MainViewModel() : this(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)) { }

    public MainViewModel(string documentsFolder)
    {
        _documentsFolder = string.IsNullOrWhiteSpace(documentsFolder)
            ? Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) : documentsFolder;
        _clipboardOutputPath = Path.Combine(_documentsFolder, "Dokument.docx");
    }

    public int SourceIndex
    {
        get => _sourceIndex;
        set
        {
            if (value is not (0 or 1) || !Set(ref _sourceIndex, value)) return;
            Changed(nameof(IsFileSource));
            Changed(nameof(IsClipboardSource));
            Changed(nameof(OutputPath));
            Changed(nameof(IsOutputReadOnly));
            Changed(nameof(OutputHint));
            Changed(nameof(CanConvert));
            InvalidateResult();
        }
    }

    public bool IsFileSource => SourceIndex == 0;
    public bool IsClipboardSource => !IsFileSource;
    public bool IsOutputReadOnly => IsFileSource && UseSourceFolder;
    public string OutputHint => IsFileSource
        ? "Gleicher Dateiname, neue Endung: .docx. Über Browse … wählst du einen anderen Ort."
        : "Vorschlag: Dokumente-Ordner und erste Überschrift als Dateiname. Name und Ort kannst du ändern.";

    public string MarkdownText
    {
        get => _markdownText;
        set
        {
            if (!Set(ref _markdownText, value ?? "")) return;
            _markdownBytes = Encoding.UTF8.GetByteCount(_markdownText);
            if (!_clipboardOutputCustomized)
            {
                _clipboardOutputPath = Path.Combine(_documentsFolder, SuggestFileName(_markdownText));
                if (IsClipboardSource) Changed(nameof(OutputPath));
            }
            Changed(nameof(PreviewSummary));
            Changed(nameof(CanConvert));
            Changed(nameof(CanClearPreview));
            InvalidateResult();
            if (_markdownBytes > MarkdownConverter.MaxInputBytes)
                ShowStatus("Der Text ist zu gross.", "Bitte kürze den Markdown-Text auf maximal 20 MB.", "warning");
        }
    }

    public string PreviewSummary => $"{MarkdownText.Length:N0} Zeichen · bearbeitbarer Markdown-Text";

    public bool TryPasteMarkdown(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            ShowStatus("Kein Text in der Zwischenablage.", "Die Zwischenablage ist leer oder enthält nur ein Bild, eine Datei oder andere Daten. Kopiere Markdown-Text, z. B. über „Kopieren“ in ChatGPT. Deine bisherige Vorschau bleibt erhalten.", "warning");
            return false;
        }
        try { MarkdownConverter.ValidateMarkdownText(text); }
        catch (ConversionException ex)
        {
            ShowStatus("Der Text konnte nicht eingefügt werden.", ex.Message + " Deine bisherige Vorschau bleibt erhalten.", "warning");
            return false;
        }
        MarkdownText = text;
        ShowStatus("Text eingefügt. Bereit für Word.", "Prüfe die Vorschau und den Speicherort. Mit „In Word konvertieren“ erstellst du dein Dokument.");
        return true;
    }

    public string InputPath
    {
        get => _inputPath;
        set
        {
            if (!Set(ref _inputPath, value)) return;
            UpdateDefaultOutput();
            InvalidateResult();
            Changed(nameof(CanConvert));
            Changed(nameof(SourceHint));
        }
    }

    public string OutputPath
    {
        get => IsFileSource ? _outputPath : _clipboardOutputPath;
        set
        {
            if (IsFileSource)
            {
                if (!Set(ref _outputPath, value)) return;
            }
            else
            {
                if (!Set(ref _clipboardOutputPath, value)) return;
                _clipboardOutputCustomized = true;
            }
            InvalidateResult();
            Changed(nameof(CanConvert));
        }
    }

    public bool UseSourceFolder
    {
        get => _useSourceFolder;
        set
        {
            if (!Set(ref _useSourceFolder, value)) return;
            UpdateDefaultOutput();
            Changed(nameof(IsOutputReadOnly));
            InvalidateResult();
        }
    }

    public bool IsBusy
    {
        get => _isBusy;
        set
        {
            if (!Set(ref _isBusy, value)) return;
            Changed(nameof(IsIdle));
            Changed(nameof(CanConvert));
            Changed(nameof(CanOpen));
            Changed(nameof(ConvertLabel));
            Changed(nameof(CanClearPreview));
        }
    }

    public bool IsReadingClipboard
    {
        get => _isReadingClipboard;
        set
        {
            if (!Set(ref _isReadingClipboard, value)) return;
            Changed(nameof(IsIdle));
            Changed(nameof(CanConvert));
            Changed(nameof(PasteLabel));
            Changed(nameof(CanOpen));
            Changed(nameof(CanClearPreview));
        }
    }

    public string PasteLabel => IsReadingClipboard ? "Wird eingefügt …" : "Aus Zwischenablage einfügen";
    public bool IsIdle => !IsBusy && !IsReadingClipboard;
    public bool CanClearPreview => IsIdle && MarkdownText.Length > 0;
    public bool CanConvert => IsIdle && !string.IsNullOrWhiteSpace(OutputPath)
        && (IsFileSource ? !string.IsNullOrWhiteSpace(InputPath)
            : !string.IsNullOrWhiteSpace(MarkdownText) && _markdownBytes <= MarkdownConverter.MaxInputBytes);
    public bool CanOpen => IsIdle && CompletedPath is not null;
    public string ConvertLabel => IsBusy ? "Wird konvertiert …" : "In Word konvertieren";
    public string SourceHint => string.IsNullOrWhiteSpace(InputPath)
        ? "Unterstützt .md, .markdown und .mdown · bis 20 MB"
        : "Überschriften, Listen, Tabellen, Links und Code werden übernommen.";

    public string? CompletedPath
    {
        get => _completedPath;
        set { if (Set(ref _completedPath, value)) Changed(nameof(CanOpen)); }
    }

    public string StatusTitle { get => _statusTitle; private set => Set(ref _statusTitle, value); }
    public string StatusMessage { get => _statusMessage; private set => Set(ref _statusMessage, value); }
    public string StatusSymbol { get => _statusSymbol; private set => Set(ref _statusSymbol, value); }
    public IBrush StatusBackground { get => _statusBackground; private set => Set(ref _statusBackground, value); }
    public IBrush StatusForeground { get => _statusForeground; private set => Set(ref _statusForeground, value); }

    public void ShowStatus(string title, string message, string kind = "info")
    {
        StatusTitle = title;
        StatusMessage = message;
        StatusSymbol = kind switch { "success" => "✓", "error" => "!", "warning" => "!", _ => "i" };
        StatusBackground = Brush.Parse(kind switch { "success" => "#EAF6EF", "error" => "#FFF0EF", "warning" => "#FFF5E3", _ => "#EEF3FC" });
        StatusForeground = Brush.Parse(kind switch { "success" => "#236344", "error" => "#A33333", "warning" => "#805B12", _ => "#38547B" });
    }

    private void UpdateDefaultOutput()
    {
        if (!UseSourceFolder) return;
        try { _outputPath = string.IsNullOrWhiteSpace(InputPath) ? "" : MarkdownConverter.GetDefaultOutputPath(InputPath.Trim()); }
        catch (ArgumentException) { _outputPath = ""; }
        if (IsFileSource)
        {
            Changed(nameof(OutputPath));
            Changed(nameof(CanConvert));
        }
    }

    private static string SuggestFileName(string markdown)
    {
        using var reader = new StringReader(markdown);
        string? line;
        do { line = reader.ReadLine()?.Trim(); } while (line is not null && line.Length == 0);
        if (line is null) return "Dokument.docx";
        var headingLength = line.TakeWhile(c => c == '#').Count();
        if (headingLength is < 1 or > 6 || line.Length <= headingLength || !char.IsWhiteSpace(line[headingLength]))
            return "Dokument.docx";
        var title = line[headingLength..].Trim().TrimEnd('#').Trim();
        title = new string(title.Where(c => !char.IsControl(c) && !"<>:\"/\\|?*`".Contains(c)).Take(80).ToArray()).Trim().TrimEnd('.');
        if (title.Length > 0 && char.IsHighSurrogate(title[^1])) title = title[..^1];
        if (string.IsNullOrWhiteSpace(title)) return "Dokument.docx";
        var stem = title.Split('.')[0].ToUpperInvariant();
        if (new[] { "CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$" }.Contains(stem)
            || (stem.Length == 4 && (stem.StartsWith("COM") || stem.StartsWith("LPT")) && "123456789¹²³".Contains(stem[3])))
            title = "Dokument-" + title;
        return title + ".docx";
    }

    private void InvalidateResult()
    {
        CompletedPath = null;
        ShowStatus("Bereit für dein Dokument.", "Prüfe Quelle und Speicherort und starte die Konvertierung.");
    }

    private bool Set<T>(ref T field, T value, [CallerMemberName] string? property = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        Changed(property);
        return true;
    }

    private void Changed([CallerMemberName] string? property = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
}
