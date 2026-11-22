# Markdown to Word

Eine kleine Desktop-Anwendung in C# mit Avalonia. Sie wandelt Markdown-Dateien oder Markdown-Text aus der Zwischenablage in editierbare Word-Dokumente um. Die Oberfläche ist deutschsprachig.

## Starten unter Windows

Die fertig veröffentlichte Anwendung liegt unter:

`artifacts/win-x64/MarkdownToDocx.exe`

Die EXE per Doppelklick starten. Diese Windows-x64-Ausgabe enthält die .NET-Laufzeit; eine separate .NET-Installation ist nicht erforderlich. Beim Weitergeben **den gesamten Ordner** verwenden, nicht nur die EXE.

1. Unter **Markdown-Datei → Browse …** eine Datei auswählen oder ihren vollständigen Pfad einfügen.
2. Standardmässig entsteht im selben Ordner eine Datei mit gleichem Namen und der Endung `.docx`. Mit dem zweiten **Browse …** lassen sich Ordner und Dateiname ändern. Für eine manuelle Eingabe zuerst die Option „Im selben Ordner …“ deaktivieren.
3. **In Word konvertieren** anklicken. Bestehende Dateien werden erst nach Bestätigung ersetzt.
4. Nach erfolgreicher Konvertierung **Word-Datei öffnen** anklicken. Dafür muss eine Standardanwendung für `.docx`, beispielsweise Word oder LibreOffice, eingerichtet sein.

Eine Beispieldatei liegt in [samples/Beispiel.md](samples/Beispiel.md).

### Direkt aus ChatGPT oder der Zwischenablage

1. In ChatGPT den Markdown-Text über **Kopieren** kopieren.
2. Unter **Deine Quelle → Zwischenablage** auf **Aus Zwischenablage einfügen** klicken.
3. Den Text in der **Markdown-Vorschau** prüfen und bei Bedarf bearbeiten. Auch direktes Einfügen mit **Strg+V** ist möglich. Mit **Clear** direkt neben dem Einfügeknopf lässt sich die Vorschau leeren. Die Zwischenablage und ein selbst gewählter Speicherort bleiben erhalten. Die Vorschau zeigt den Markdown-Quelltext, keine Word-Seitenansicht.
4. **In Word konvertieren** klicken, anschliessend **Word-Datei öffnen**.

Als Speicherort wird der Dokumente-Ordner vorgeschlagen. Beginnt der Text mit einer Markdown-Überschrift (`#` bis `######`), wird daraus ein geeigneter Dateiname erzeugt; sonst heisst der Vorschlag `Dokument.docx`. Über **Browse …** oder das Pfadfeld können Name und Ort geändert werden. Ein selbst gewähltes Ziel bleibt bei weiterem Einfügen erhalten. Bestehende Dateien werden erst nach Bestätigung ersetzt.

Text, Dateiauswahl und die jeweiligen Zielpfade bleiben beim Quellenwechsel innerhalb der geöffneten Anwendung erhalten. Eine leere Zwischenablage, Bilder, kopierte Dateien und andere Formate ohne Text erzeugen einen Hinweis und lassen die vorhandene Vorschau unverändert. Enthält die Zwischenablage mehrere Formate einschliesslich Text, wird nur der Text verwendet. Zugriffsfehler werden abgefangen. Es wird keine Markdown-Zwischendatei angelegt. Beim Schliessen der Anwendung wird der Vorschautext nicht dauerhaft gespeichert.

## Entwicklung

Voraussetzung: **.NET 10 SDK**. Zum ersten Wiederherstellen der NuGet-Pakete ist Internetzugang erforderlich.

```powershell
dotnet restore MarkdownToDocx.slnx
dotnet build MarkdownToDocx.slnx -c Release
dotnet test MarkdownToDocx.slnx -c Release
dotnet run --project src/MarkdownToDocx
```

Eigenständige Windows-Ausgabe erstellen:

```powershell
dotnet publish src/MarkdownToDocx -c Release -r win-x64 --self-contained true -o artifacts/win-x64
```

## Umfang

- Eingaben: `.md`, `.markdown`, `.mdown`, maximal 20 MB; UTF-8 sowie durch BOM erkannte Unicode-Kodierungen.
- Zwischenablage: reiner Markdown-Text, maximal 20 MB in UTF-8; dieselbe Word-Formatierung wie bei Dateien.
- Überschriften, fett/kursiv/durchgestrichen, Absätze, Aufzählungen und nummerierte Listen, Zitate, Tabellen, Links und Codeblöcke.
- A4-Seiten mit 2-cm-Rändern, Arial 11 pt, schwarzer Überschriftenhierarchie und abgestimmten Absatz- und Listenabständen. Überschriften bleiben als echte Word-Formatvorlagen editierbar und in der Navigation sichtbar.
- Nummern in Markdown-Überschriften bleiben wörtlich erhalten: `### 1. Frage` wird nicht in eine automatische Kapitelnummer wie `1.1.1.` umgedeutet. Getrennte nummerierte Listen beginnen neu; explizite Startwerte werden respektiert.
- Eindeutige ChatGPT-Exportreste aus einem alleinstehenden `image`-Link zum Google-Favicon-Dienst und einer kurzen Quellenbezeichnung werden als kompakte graue Quellenzeile dargestellt. Die Quellenbezeichnung (auch `+2`) bleibt erhalten. Normale Links und Bilder bleiben erhalten. Aus fehlenden Artikel-URLs werden keine vermeintlichen Quellenlinks erzeugt.
- Relative Bildpfade werden vom Ordner der Markdown-Datei aus aufgelöst. Bei Zwischenablage-Text gibt es keinen Quellordner: relative Bilder werden mit Hinweis ausgelassen, ihr Alternativtext bleibt erhalten. Hier vollständige Bildadressen oder eine Markdown-Datei mit zugehörigen Bildern verwenden. Gängige Bildformate wie PNG und JPEG werden eingebettet. Nicht übernommene Bilder lösen einen sichtbaren Hinweis aus.
- Online-Bilder werden von ihrer angegebenen Adresse geladen und benötigen Internetzugang. Die Markdown-Datei wird nicht an einen Konvertierungsdienst hochgeladen.
- Raw-HTML wird als Text behandelt. Mermaid-Diagramme, LaTeX-Formeln, komplexes HTML/CSS und eine pixelgleiche Nachbildung eines Markdown-Editors gehören nicht zum Umfang. Code wird ohne Syntaxfarben übernommen.
- Word wird zur Erstellung nicht benötigt. Die Oberfläche ist mit Avalonia grundsätzlich plattformübergreifend; die bereitgestellte Ausgabe und der manuelle Anwendungstest beziehen sich auf Windows x64.

## Aufbau und Bibliotheken

- `src/MarkdownToDocx`: Avalonia-Oberfläche, kleiner ViewModel-Zustand und native Dateidialoge. Kein DI-Container und kein zusätzliches MVVM-Framework.
- `src/MarkdownToDocx.Core`: Konvertierung und Dateivalidierung, unabhängig von der Oberfläche. `WordDocumentDesign.cs` definiert das Dokumentlayout und die begrenzte Bereinigung von Quellen-Exportresten.
- `tests/MarkdownToDocx.Tests`: Integrationstests für DOCX-Inhalte und OpenXML-Validität, Bilder, Kodierung, Dateisperren, Überschreiben, Abbruch und UI-Zustand.
- `tests/MarkdownToDocx.UiTests`: automatisierte Avalonia-Headless-Tests mit xUnit v3 für Einfügen, Vorschau-Bearbeitung, Konvertierung, Quellenwechsel, Überschreibabbruch sowie leere, bildhaltige, dateihaltige und fremde Zwischenablageformate und Lesefehler. Die bestehenden Kern-Tests verwenden weiterhin xUnit v2.

Für die Zwischenablage-Erweiterung bestehen 54 automatisierte Testfälle, davon elf GUI-Testfälle. Dazu gehören das Leeren der Vorschau, deaktivierte Aktionen bei leerer Vorschau und während laufender Vorgänge sowie erneutes Einfügen nach dem Leeren. Zusätzlich wurde die veröffentlichte Windows-Anwendung mit der echten Zwischenablage getestet: Text einfügen, Ziel über den nativen Speicherdialog wählen, konvertieren und das Ergebnis über den Öffnen-Knopf in Microsoft Word prüfen. Auch ein Bildschirmbild und eine im Windows-Dateidialog kopierte Datei wurden als Zwischenablageinhalt geprüft; beide erzeugen einen Hinweis und lassen die Vorschau erhalten.

Die Konvertierung besteht aus zwei Bibliotheksschritten: **Markdown → HTML mit Markdig → DOCX mit HtmlToOpenXml**. Die DOCX-Datei wird zuerst vollständig in eine temporäre Datei im Zielordner geschrieben. Erst bei Erfolg wird sie an den endgültigen Platz verschoben; Fehler und Abbrüche vor diesem Schritt lassen eine bestehende Zieldatei unverändert.

Verwendete Pakete:

| Paket | Version | Zweck | Lizenz |
| --- | --- | --- | --- |
| [Avalonia](https://www.nuget.org/packages/Avalonia/12.1.6) | 12.1.6 | Desktop-Oberfläche und Fluent-Theme | MIT |
| [Markdig](https://www.nuget.org/packages/Markdig/1.6.0) | 1.6.0 | Markdown-Parser | BSD-2-Clause |
| [HtmlToOpenXml](https://www.nuget.org/packages/HtmlToOpenXml/4.0.0) | 4.0.0 | Word-Erzeugung; nutzt Open XML SDK | MIT |

Die Tests prüfen die erzeugte Word-Struktur gegen das OpenXML-Schema für Office 2013. Aktuelle API-Referenzen: [Avalonia-Dateidialoge](https://docs.avaloniaui.net/docs/services/file-dialogs), [Markdig](https://github.com/xoofx/markdig), [HtmlToOpenXml](https://github.com/onizet/html2openxml).

## Visuelle Dokumentprüfung

Für die lokale Entwicklung wurde LibreOffice **26.8.3.1** separat eingerichtet:

`%LOCALAPPDATA%\CodexTools\LibreOffice\26.8.3\program\soffice.exe`

Die offizielle, signierte Windows-x64-MSI wurde mit `msiexec /a` in diesen Ordner entpackt. Die Prüfsumme wurde vor der Einrichtung mit dem Windows-Paketkatalog abgeglichen. Diese separate Installation übernimmt keine Dateizuordnungen. Für den Renderer wird der Programmordner nur im jeweiligen Prozess dem `PATH` hinzugefügt. LibreOffice gehört nicht zu den Abhängigkeiten der ausgelieferten Avalonia-Anwendung.

Der Codex-Dokumentrenderer `render_docx.py` kann damit DOCX-Dateien über ein temporäres LibreOffice-Profil als PDF und PNG-Seiten rendern. Zusätzlich wurde die Ausgabe mit Microsoft Word geprüft. Seitenumbrüche können sich je nach Word-/LibreOffice-Version und installierten Schriftarten leicht unterscheiden; eine pixelgleiche Kopie der ChatGPT-Webansicht ist nicht das Ziel.
