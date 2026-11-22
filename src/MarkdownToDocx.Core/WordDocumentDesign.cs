using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToDocx.Core;

/// <summary>A small, explicit Word style sheet; never relies on Word's theme defaults.</summary>
internal static class WordDocumentDesign
{
    private const string BodyFont = "Arial";

    public static Styles CreateStyles()
    {
        var styles = new Styles(new DocDefaults(
            new RunPropertiesDefault(new RunPropertiesBaseStyle(
                Fonts(), new Color { Val = "202020" }, new FontSize { Val = "22" })),
            new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(
                new WidowControl(), Spacing(after: 140)))));

        styles.Append(new Style
        {
            Type = StyleValues.Paragraph, StyleId = "Normal", Default = true,
            StyleName = new StyleName { Val = "Normal" },
            StyleParagraphProperties = new StyleParagraphProperties(new WidowControl(), Spacing(after: 140)),
            StyleRunProperties = new StyleRunProperties(Fonts(), new Color { Val = "202020" }, new FontSize { Val = "22" })
        });

        // Use semantic heading styles so Word navigation and accessibility still work.
        int[] sizes = [40, 28, 23, 22, 22, 22];
        for (var level = 1; level <= 6; level++)
        {
            var heading = ParagraphStyle($"Heading{level}", $"heading {level}");
            heading.StyleParagraphProperties = new StyleParagraphProperties(
                new KeepNext(), new KeepLines(),
                Spacing(before: level == 1 ? 0 : level == 2 ? 360 : 280,
                    after: level == 1 ? 240 : 140, line: 280),
                new OutlineLevel { Val = level - 1 });
            heading.StyleRunProperties = new StyleRunProperties(
                Fonts(), new Bold(), new BoldComplexScript(), new Color { Val = "000000" },
                new FontSize { Val = sizes[level - 1].ToString() },
                new FontSizeComplexScript { Val = sizes[level - 1].ToString() });
            styles.Append(heading);
        }

        var list = ParagraphStyle("ListParagraph", "List Paragraph");
        list.StyleParagraphProperties = new StyleParagraphProperties(new WidowControl(), Spacing(after: 40));
        styles.Append(list);

        var leadIn = ParagraphStyle("LeadIn", "Lead In");
        leadIn.StyleParagraphProperties = new StyleParagraphProperties(
            new KeepNext(), new KeepLines(), Spacing(before: 140, after: 80));
        styles.Append(leadIn);

        var source = ParagraphStyle("SourceNote", "Source Note");
        source.StyleParagraphProperties = new StyleParagraphProperties(Spacing(after: 140, line: 240));
        source.StyleRunProperties = new StyleRunProperties(new Color { Val = "666666" }, new FontSize { Val = "18" });
        styles.Append(source);

        foreach (var name in new[] { "Quote", "IntenseQuote" })
        {
            var quote = ParagraphStyle(name, name);
            quote.StyleParagraphProperties = new StyleParagraphProperties(
                Spacing(before: 80, after: 140), new Indentation { Left = "360", Right = "360" });
            quote.StyleRunProperties = new StyleRunProperties(new Italic(), new Color { Val = "444444" });
            styles.Append(quote);
        }

        styles.Append(new Style
        {
            Type = StyleValues.Character, StyleId = "Hyperlink",
            StyleName = new StyleName { Val = "Hyperlink" },
            StyleRunProperties = new StyleRunProperties(new Color { Val = "2459A6" }, new Underline { Val = UnderlineValues.Single })
        });
        return styles;
    }

    public static void FormatContent(MainDocumentPart main)
    {
        CompactExportSourceNotes(main);

        foreach (var paragraph in main.Document!.Descendants<Paragraph>())
        {
            var properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
            var styleId = properties.ParagraphStyleId?.Val?.Value;
            if (styleId == "ListParagraph")
            {
                // Give the end of a list breathing room without separating each item too much.
                if (paragraph.NextSibling() is not Paragraph next || next.ParagraphProperties?.ParagraphStyleId?.Val != "ListParagraph")
                    properties.SpacingBetweenLines = Spacing(after: 160);
            }
            else if (styleId is null or "Normal" && IsBoldLeadIn(paragraph))
            {
                properties.ParagraphStyleId = new ParagraphStyleId { Val = "LeadIn" };
            }
        }

        if (main.NumberingDefinitionsPart?.Numbering is not { } numbering) return;
        foreach (var level in numbering.Descendants<Level>())
        {
            var depth = level.LevelIndex?.Value ?? 0;
            level.PreviousParagraphProperties ??= new PreviousParagraphProperties();
            level.PreviousParagraphProperties.Indentation = new Indentation
            {
                Left = ((depth + 1) * 360).ToString(), Hanging = "180"
            };
            if (level.NumberingFormat?.Val?.Value == NumberFormatValues.Bullet)
            {
                level.LevelText = new LevelText { Val = "•" };
                level.NumberingSymbolRunProperties = new NumberingSymbolRunProperties(Fonts());
            }
        }
    }

    private static bool IsBoldLeadIn(Paragraph paragraph)
    {
        if (paragraph.ParagraphProperties?.NumberingProperties is not null || paragraph.InnerText.Length > 160)
            return false;
        var runs = paragraph.Descendants<Run>().Where(run => !string.IsNullOrWhiteSpace(run.InnerText)).ToArray();
        return runs.Length > 0 && runs.All(run => run.RunProperties?.Bold is { } bold && (bold.Val?.Value ?? true));
    }

    private static void CompactExportSourceNotes(MainDocumentPart main)
    {
        var links = main.HyperlinkRelationships.ToDictionary(link => link.Id, link => link.Uri);
        foreach (var paragraph in main.Document!.Descendants<Paragraph>().ToArray())
        {
            // Only a standalone link literally named "image" to Google's favicon endpoint
            // followed by a short plain source caption is an export artifact. Real links/images stay intact.
            if (paragraph.InnerText.Trim() != "image" || paragraph.ParagraphProperties?.NumberingProperties is not null)
                continue;
            var hyperlinks = paragraph.Elements<Hyperlink>().ToArray();
            if (hyperlinks.Length != 1 || hyperlinks[0].Id?.Value is not { } id || !links.TryGetValue(id, out var uri) || !IsFaviconUrl(uri))
                continue;
            if (paragraph.NextSibling() is not Paragraph caption || !IsSourceCaption(caption))
                continue;

            caption.ParagraphProperties ??= new ParagraphProperties();
            caption.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = "SourceNote" };
            if (paragraph.PreviousSibling() is Paragraph previous)
            {
                previous.ParagraphProperties ??= new ParagraphProperties();
                previous.ParagraphProperties.KeepNext = new KeepNext();
                previous.ParagraphProperties.SpacingBetweenLines = Spacing(after: 40);
                if (previous.ParagraphProperties.NumberingProperties?.NumberingLevelReference?.Val?.Value is { } depth)
                    caption.ParagraphProperties.Indentation = new Indentation { Left = ((depth + 1) * 360).ToString() };
            }
            paragraph.Remove();
        }
    }

    private static bool IsFaviconUrl(Uri uri) => uri.IsAbsoluteUri
        && uri.Scheme == Uri.UriSchemeHttps
        && uri.Host is "www.google.com" or "google.com"
        && uri.AbsolutePath == "/s2/favicons"
        && uri.Query.TrimStart('?').Split('&').Any(part => part.StartsWith("domain=", StringComparison.Ordinal));

    private static bool IsSourceCaption(Paragraph paragraph)
    {
        var text = paragraph.InnerText.Trim();
        return text.Length is > 0 and <= 80 && !text.EndsWith('.') && !text.EndsWith('!') && !text.EndsWith('?')
            && paragraph.ParagraphProperties?.NumberingProperties is null
            && paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value is null or "Normal"
            && !paragraph.Elements<Hyperlink>().Any()
            && !paragraph.Descendants<Drawing>().Any()
            && paragraph.Elements<Run>().All(run => run.RunProperties?.Bold is null && run.RunProperties?.Italic is null);
    }

    private static Style ParagraphStyle(string id, string name) => new()
    {
        Type = StyleValues.Paragraph, StyleId = id, StyleName = new StyleName { Val = name },
        BasedOn = new BasedOn { Val = "Normal" }, NextParagraphStyle = new NextParagraphStyle { Val = "Normal" },
        PrimaryStyle = new PrimaryStyle()
    };

    private static RunFonts Fonts() => new() { Ascii = BodyFont, HighAnsi = BodyFont, ComplexScript = BodyFont };

    private static SpacingBetweenLines Spacing(int before = 0, int after = 0, int line = 300) => new()
    {
        Before = before.ToString(), After = after.ToString(), Line = line.ToString(), LineRule = LineSpacingRuleValues.Auto
    };
}
