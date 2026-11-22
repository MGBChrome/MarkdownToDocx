using System.Diagnostics;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

[assembly: AvaloniaTestApplication(typeof(MarkdownToDocx.UiTests.TestAppBuilder))]

namespace MarkdownToDocx.UiTests;

public static class TestAppBuilder
{
    public static AppBuilder BuildAvaloniaApp() => AppBuilder.Configure<App>()
        .WithInterFont().UseHeadless(new AvaloniaHeadlessPlatformOptions());
}

public sealed class ClipboardUiTests
{
    private static T Control<T>(MainWindow window, string name) where T : Control => window.FindControl<T>(name)!;
    private static void Click(MainWindow window, string name) => Control<Button>(window, name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
    private static MainViewModel SelectClipboard(MainWindow window)
    {
        window.Show();
        Control<ListBox>(window, "SourceSelector").SelectedIndex = 1;
        var model = Assert.IsType<MainViewModel>(window.DataContext);
        Assert.True(model.IsClipboardSource);
        return model;
    }

    [AvaloniaTheory]
    [InlineData("# Zum Leeren\n\nEin Absatz.")]
    [InlineData(" \r\n\t")]
    public async Task ClearPreviewResetsActionsAndAllowsPastingAgain(string previewText)
    {
        var window = new MainWindow();
        try
        {
            var model = SelectClipboard(window);
            var clear = Control<Button>(window, "ClearPreviewButton");
            Assert.False(clear.IsEnabled);
            model.InputPath = "Dateiquelle.md";
            model.OutputPath = "Mein Ziel.docx";
            Control<TextBox>(window, "MarkdownPreviewBox").Text = previewText;
            model.CompletedPath = "Mein Ziel.docx";
            await window.Clipboard!.SetTextAsync("# Erneut einfügen");
            Assert.True(clear.IsEnabled);
            model.IsBusy = true;
            Assert.False(clear.IsEnabled);
            model.IsBusy = false;
            model.IsReadingClipboard = true;
            Assert.False(clear.IsEnabled);
            model.IsReadingClipboard = false;
            Assert.True(clear.IsEnabled);

            Click(window, "ClearPreviewButton");

            var preview = Control<TextBox>(window, "MarkdownPreviewBox");
            Assert.Equal("", preview.Text);
            Assert.Equal("", model.MarkdownText);
            Assert.True(preview.IsFocused);
            Assert.False(clear.IsEnabled);
            Assert.False(Control<Button>(window, "ConvertButton").IsEnabled);
            Assert.False(Control<Button>(window, "OpenButton").IsEnabled);
            Assert.Equal("Vorschau geleert.", model.StatusTitle);
            Assert.Equal("Mein Ziel.docx", model.OutputPath);
            Assert.Equal("Dateiquelle.md", model.InputPath);
            Assert.True(model.UseSourceFolder);
            Assert.Equal("# Erneut einfügen", await window.Clipboard!.TryGetTextAsync());

            Click(window, "PasteButton");
            await WaitUntil(() => !model.IsReadingClipboard);
            Assert.Equal("# Erneut einfügen", preview.Text);
            Assert.True(clear.IsEnabled);
            Assert.True(Control<Button>(window, "ConvertButton").IsEnabled);
        }
        finally { window.Close(); }
    }

    [AvaloniaFact]
    public async Task PasteEditAndConvertUsesPreviewAndShowsSuccess()
    {
        var folder = Directory.CreateTempSubdirectory("MarkdownToDocx.UiTests-");
        var window = new MainWindow();
        try
        {
            var model = SelectClipboard(window);
            Assert.False(Control<Button>(window, "ConvertButton").IsEnabled);
            Assert.False(Control<CheckBox>(window, "SameFolderCheckBox").IsVisible);
            Assert.False(Control<TextBox>(window, "OutputPathBox").IsReadOnly);
            model.OutputPath = Path.Combine(folder.FullName, "Vorschau.docx");
            await window.Clipboard!.SetTextAsync("# Aus der Zwischenablage\n\nOriginal");
            Click(window, "PasteButton");
            await WaitUntil(() => !model.IsReadingClipboard);
            var preview = Control<TextBox>(window, "MarkdownPreviewBox");
            Assert.Equal("# Aus der Zwischenablage\n\nOriginal", preview.Text);
            Assert.True(preview.IsFocused);
            preview.Text = "# 1. Bearbeitet\n\n**Änderung** in der Vorschau.";
            Assert.Equal(preview.Text, model.MarkdownText);
            Assert.True(Control<Button>(window, "ConvertButton").IsEnabled);
            Click(window, "ConvertButton");
            Assert.True(model.IsBusy);
            Assert.False(Control<Button>(window, "PasteButton").IsEnabled);
            Assert.False(preview.IsEnabled);
            Assert.False(Control<ListBox>(window, "SourceSelector").IsEnabled);
            await WaitUntil(() => !model.IsBusy);
            Assert.True(Control<Button>(window, "OpenButton").IsEnabled, model.StatusMessage);
            Assert.Contains("Fertig", model.StatusTitle);
            using var document = WordprocessingDocument.Open(model.OutputPath, false);
            Assert.Contains("1. Bearbeitet", document.MainDocumentPart!.Document!.InnerText);
            Assert.Contains("Änderung", document.MainDocumentPart.Document!.InnerText);
            Assert.DoesNotContain("Original", document.MainDocumentPart.Document!.InnerText);
            Assert.Single(folder.GetFiles());
            preview.Text += " Neue Änderung.";
            Assert.False(Control<Button>(window, "OpenButton").IsEnabled);
        }
        finally { window.Close(); folder.Delete(true); }
    }

    [AvaloniaTheory]
    [InlineData("empty")]
    [InlineData("bitmap")]
    [InlineData("file")]
    [InlineData("custom")]
    public async Task NonTextClipboardIsHandledWithoutReadingForeignDataOrLosingPreview(string kind)
    {
        var window = new MainWindow();
        try
        {
            var model = SelectClipboard(window);
            model.MarkdownText = "# Vorhandene Vorschau";
            var data = new DataTransfer();
            using var bitmap = new WriteableBitmap(new PixelSize(1, 1), new Vector(96, 96));
            var foreignDataRead = false;
            switch (kind)
            {
                case "bitmap":
                    var imageItem = new DataTransferItem();
                    imageItem.SetBitmap(bitmap);
                    data.Add(imageItem);
                    break;
                case "file":
                    data.Add(DataTransferItem.Create(DataFormat.File, () =>
                    {
                        foreignDataRead = true;
                        throw new InvalidOperationException("A copied file must never be read as Markdown text.");
                    }));
                    break;
                case "custom":
                    data.Add(DataTransferItem.Create(DataFormat.CreateBytesApplicationFormat("MarkdownToDocxTestData"), new byte[] { 1, 2, 3 }));
                    break;
            }
            await window.Clipboard!.SetDataAsync(data);
            Click(window, "PasteButton");
            await WaitUntil(() => !model.IsReadingClipboard);
            Assert.False(foreignDataRead);
            Assert.Equal("# Vorhandene Vorschau", Control<TextBox>(window, "MarkdownPreviewBox").Text);
            Assert.Contains("Kein Text", model.StatusTitle);
            Assert.True(Control<Button>(window, "PasteButton").IsEnabled);
            // The next valid paste still works after the unsupported format.
            await window.Clipboard.SetTextAsync("# Danach gültig");
            Click(window, "PasteButton");
            await WaitUntil(() => !model.IsReadingClipboard);
            Assert.Equal("# Danach gültig", model.MarkdownText);
        }
        finally { window.Close(); }
    }

    [AvaloniaFact]
    public async Task MixedFormatsUseOnlyTextAndEmptyPreviewCannotConvert()
    {
        var window = new MainWindow();
        try
        {
            var model = SelectClipboard(window);
            await window.Clipboard!.SetDataAsync(new DataTransfer());
            Click(window, "PasteButton");
            await WaitUntil(() => !model.IsReadingClipboard);
            Assert.False(Control<Button>(window, "ConvertButton").IsEnabled);
            var data = new DataTransfer();
            using var bitmap = new WriteableBitmap(new PixelSize(1, 1), new Vector(96, 96));
            var item = DataTransferItem.CreateText("# Nur Text verwenden");
            item.SetBitmap(bitmap);
            data.Add(item);
            await window.Clipboard.SetDataAsync(data);
            Click(window, "PasteButton");
            await WaitUntil(() => !model.IsReadingClipboard);
            Assert.Equal("# Nur Text verwenden", model.MarkdownText);
            Assert.True(Control<Button>(window, "ConvertButton").IsEnabled);
            Control<TextBox>(window, "MarkdownPreviewBox").Text = " \r\n ";
            Assert.False(Control<Button>(window, "ConvertButton").IsEnabled);
        }
        finally { window.Close(); }
    }

    [AvaloniaFact]
    public async Task FailingClipboardProviderShowsErrorAndKeepsPreview()
    {
        var window = new MainWindow();
        try
        {
            var model = SelectClipboard(window);
            model.MarkdownText = "Nicht verlieren";
            var data = new DataTransfer();
            data.Add(DataTransferItem.Create(DataFormat.Text, () => throw new InvalidOperationException("Clipboard unavailable")));
            await window.Clipboard!.SetDataAsync(data);
            Click(window, "PasteButton");
            await WaitUntil(() => !model.IsReadingClipboard);
            Assert.Equal("Nicht verlieren", model.MarkdownText);
            Assert.Contains("nicht gelesen", model.StatusMessage);
            Assert.True(model.IsIdle);
        }
        finally { window.Close(); }
    }

    [AvaloniaFact]
    public async Task CancelOverwriteKeepsFileAndPreview()
    {
        var folder = Directory.CreateTempSubdirectory("MarkdownToDocx.UiTests-");
        var window = new MainWindow();
        try
        {
            var model = SelectClipboard(window);
            model.MarkdownText = "# Neuer Inhalt";
            model.OutputPath = Path.Combine(folder.FullName, "Vorhanden.docx");
            await File.WriteAllTextAsync(model.OutputPath, "Original", TestContext.Current.CancellationToken);
            Click(window, "ConvertButton");
            await WaitUntil(() => window.OwnedWindows.Count == 1);
            var dialog = window.OwnedWindows.Single();
            var cancel = dialog.GetVisualDescendants().OfType<Button>().Single(button => button.IsCancel);
            cancel.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await WaitUntil(() => !model.IsBusy);
            Assert.Equal("Original", await File.ReadAllTextAsync(model.OutputPath, TestContext.Current.CancellationToken));
            Assert.Equal("# Neuer Inhalt", model.MarkdownText);
            Assert.False(model.CanOpen);
            Assert.Contains("unverändert", model.StatusTitle);
        }
        finally { window.Close(); folder.Delete(true); }
    }

    [AvaloniaFact]
    public void SourceSwitchingUpdatesVisibilityAndPreservesBothSources()
    {
        var window = new MainWindow();
        try
        {
            var model = SelectClipboard(window);
            Control<TextBox>(window, "MarkdownPreviewBox").Text = "# Behalten";
            Control<TextBox>(window, "OutputPathBox").Text = "Individuell.docx";
            Control<ListBox>(window, "SourceSelector").SelectedIndex = 0;
            Assert.True(Control<StackPanel>(window, "FileSourcePanel").IsVisible);
            Assert.False(Control<StackPanel>(window, "ClipboardSourcePanel").IsVisible);
            Control<TextBox>(window, "InputPathBox").Text = "Quelle.md";
            Assert.Equal("Quelle.docx", Control<TextBox>(window, "OutputPathBox").Text);
            Assert.True(Control<TextBox>(window, "OutputPathBox").IsReadOnly);
            Control<ListBox>(window, "SourceSelector").SelectedIndex = 1;
            Assert.Equal("# Behalten", model.MarkdownText);
            Assert.Equal("Individuell.docx", model.OutputPath);
            Assert.False(Control<TextBox>(window, "OutputPathBox").IsReadOnly);
        }
        finally { window.Close(); }
    }

    private static async Task WaitUntil(Func<bool> predicate)
    {
        var timer = Stopwatch.StartNew();
        while (!predicate())
        {
            Assert.True(timer.Elapsed < TimeSpan.FromSeconds(10), "UI operation did not complete within ten seconds.");
            await Task.Delay(10, TestContext.Current.CancellationToken);
        }
    }
}
