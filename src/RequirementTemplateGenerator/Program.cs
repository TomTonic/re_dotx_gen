using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace RequirementTemplateGenerator;

/// <summary>
/// Generates a Word template (.dotx) with "Requirement" styles that support hierarchical numbering.
/// </summary>
public class Program
{
    // Constants for formatting
    private const string FontName = "Arial";
    private const string FontSizeH = "24"; // 12pt = 24 half-points
    private const string FontSize = "22"; // 11pt = 22 half-points
    private const int TextIndentTwips = 1701; // 3cm = 1701 twips (567 twips per cm)
    private const int HeadingLevels = 5;
    private const int RequirementLevels = 8;

    // Note types: (german, english, text)
    private static readonly (string German, string English, string Text)[] NoteTypes = new[]
    {
        ("Hinweis", "Note", "Hinweis: "),
        ("Beispiel", "Example", "Beispiel: "),
        ("Erläuterung/Begründung", "Rationale", "Erläuterung/Begründung: "),
        ("Referenz(en)", "References", "Referenz(en): "),
        ("Ableitung zu", "DerivedFrom", "Ableitung zu: ")
    };

    public static void Main(string[] args)
    {
        string outputPath = args.Length > 0 ? args[0] : "RequirementTemplate.dotx";

        Console.WriteLine($"Generating Word template: {outputPath}");
        try
        {
            GenerateTemplate(outputPath);
            Console.WriteLine($"Template generated successfully: {outputPath}");
        }
        catch (UnauthorizedAccessException ua)
        {
            Console.Error.WriteLine($"Access denied writing to '{outputPath}': {ua.Message}");
            var fallback = Path.Combine(Path.GetTempPath(), Path.GetFileName(outputPath));
            Console.WriteLine($"Attempting fallback output path: {fallback}");
            GenerateTemplate(fallback);
            Console.WriteLine($"Template generated successfully: {fallback}");
        }
        catch (IOException io) when (io.Message != null && io.Message.Contains("Permission denied"))
        {
            Console.Error.WriteLine($"I/O error writing to '{outputPath}': {io.Message}");
            var fallback = Path.Combine(Path.GetTempPath(), Path.GetFileName(outputPath));
            Console.WriteLine($"Attempting fallback output path: {fallback}");
            GenerateTemplate(fallback);
            Console.WriteLine($"Template generated successfully: {fallback}");
        }
    }

    /// <summary>
    /// Generates the Word template with all required styles and numbering definitions.
    /// </summary>
    public static void GenerateTemplate(string filePath)
    {
        // Determine document type based on file extension
        var docType = Path.GetExtension(filePath).ToLowerInvariant() == ".docx"
            ? WordprocessingDocumentType.Document
            : WordprocessingDocumentType.Template;

        using (var document = WordprocessingDocument.Create(filePath, docType))
        {
            // Add the main document part
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            // Add styles part
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = CreateStyles();

            // Add numbering part
            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = CreateNumbering();

            // Add document settings (compatibility)
            AddDocumentSettings(mainPart);

            // Add sample content to demonstrate the styles
            AddSampleContent(mainPart.Document.Body!);

            document.Save();
        }

        // Post-process the generated .dotx to ensure correct content types and relationship targets
        DotxFixer.FixDotx(filePath);
    }

    /// <summary>
    /// Creates all the styles for the document (Normal, Headings, Requirements).
    /// </summary>
    private static Styles CreateStyles()
    {
        var styles = new Styles();

        // Add document defaults
        styles.Append(CreateDocDefaults());

        // Create singular styles
        CreateSingularStyles(styles);

        // Add Requirement styles (1-8)
        for (int level = 1; level <= RequirementLevels; level++)
        {
            styles.Append(CreateRequirementStyle(level));
        }

        // Add H styles (H1-H5, not built-in Heading styles)
        for (int level = 1; level <= HeadingLevels; level++)
        {
            styles.Append(CreateHeadingStyle(level));
        }

        return styles;
    }

    /// <summary>
    /// Creates document defaults with 11pt font and single line spacing.
    /// </summary>
    private static DocDefaults CreateDocDefaults()
    {
        return new DocDefaults(
            new RunPropertiesDefault(
                new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = FontName, HighAnsi = FontName, ComplexScript = FontName },
                    new FontSize { Val = FontSize },
                    new FontSizeComplexScript { Val = FontSize }
                )
            ),
            new ParagraphPropertiesDefault(
                new ParagraphPropertiesBaseStyle(
                    new SpacingBetweenLines { LineRule = LineSpacingRuleValues.Auto }
                )
            )
        );
    }

    private static void CreateSingularStyles(Styles styles)
    {
        // Normal style - required base style so references to it work
        var normalStyle = new Style
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal"
        };
        normalStyle.Append(new StyleName { Val = "Normal" });
        normalStyle.Append(new StyleRunProperties(
            new RunFonts { Ascii = FontName, HighAnsi = FontName, ComplexScript = FontName },
            new FontSize { Val = FontSize },
            new FontSizeComplexScript { Val = FontSize }
        ));
        styles.Append(normalStyle);

        // Anonymous paragraph style
        var anonymousParaStyle = new Style
        {
            Type = StyleValues.Paragraph,
            CustomStyle = true,
            StyleId = "REAnonymousPara"
        };
        anonymousParaStyle.Append(new StyleName { Val = "RE Ergänzung - Absatz" });
        anonymousParaStyle.Append(new BasedOn { Val = "Normal" });
        anonymousParaStyle.Append(new PrimaryStyle());
        anonymousParaStyle.Append(new StyleParagraphProperties(
            new Indentation { Left = TextIndentTwips.ToString() },
            new SpacingBetweenLines { Before = "60" },
            new Justification { Val = JustificationValues.Both },
            new NextParagraphStyle { Val = "REAnonymousPara" }
        ));
        anonymousParaStyle.Append(new StyleRunProperties(
            new FontSize { Val = "16" }, // 8pt
            new FontSizeComplexScript { Val = "16" } // 8pt,
            ,
            new Languages { Val = "de-DE" }
            //new Italic()
        ));
        styles.Append(anonymousParaStyle);

        // Create note styles
        for (int i = 0; i < NoteTypes.Length; i++)
        {
            var (german, english, text) = NoteTypes[i];
            var noteStyle = new Style
            {
                Type = StyleValues.Paragraph,
                CustomStyle = true,
                StyleId = $"RE{english}"
            };
            noteStyle.Append(new StyleName { Val = $"RE Ergänzung - {german}" });
            noteStyle.Append(new BasedOn { Val = "REAnonymousPara" });
            noteStyle.Append(new PrimaryStyle());
            noteStyle.Append(new StyleParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = 10 + i } // 10 for first, 11 for second, etc.
                ),
                new SpacingBetweenLines { Before = "0" },
                new OutlineLevel { Val = 9 }, // Not in outline
                new NextParagraphStyle { Val = "REAnonymousPara" }
            ));
            //noteStyle.Append(new StyleRunProperties(
            //    new Italic()
            //));
            styles.Append(noteStyle);
        }
    }

    /// <summary>
    /// Creates a H style for the specified level (H1-H5).
    /// </summary>
    private static Style CreateHeadingStyle(int level)
    {
        var style = new Style
        {
            Type = StyleValues.Paragraph,
            CustomStyle = true,
            StyleId = $"REHeading{level}"
        };

        style.Append(new StyleName { Val = $"RE Überschrift {level}" });
        style.Append(new BasedOn { Val = level == 1 ? "Normal" : $"REHeading{level - 1}" });
        style.Append(new PrimaryStyle());

        style.Append(new StyleParagraphProperties(
            new NumberingProperties(
                new NumberingLevelReference { Val = level - 1 },
                new NumberingId { Val = 1 }
            ),
            new OutlineLevel { Val = 0 }, // Always top level in outline
            new SpacingBetweenLines { Before = "360", After = "120" }, // 12pt before, 4pt after headings   new Indentation { Left = TextIndentTwips.ToString(), Hanging = TextIndentTwips.ToString() },
            new Tabs(
                new TabStop { Val = TabStopValues.Left, Position = TextIndentTwips, Leader = TabStopLeaderCharValues.Dot }
            ),
            new NextParagraphStyle { Val = $"RERequirement{level + 1}" }
        ));

        if (level == 1) // other levels inherit their style from level 1
        {
            style.Append(new StyleRunProperties(
                new Bold(),
                new FontSize { Val = FontSizeH },
                new FontSizeComplexScript { Val = FontSizeH }
            ));
        }

        return style;
    }

    /// <summary>
    /// Creates a Requirement style for the specified level.
    /// Requirements use numbering that continues from heading levels.
    /// </summary>
    private static Style CreateRequirementStyle(int level)
    {
        // Map RequirementN to ilvl N (0-based), clamped to Word's 9-level max
        int numberingLevel = Math.Min(level, 9) - 1;

        var style = new Style
        {
            Type = StyleValues.Paragraph,
            CustomStyle = true,
            StyleId = $"RERequirement{level}"
        };

        style.Append(new StyleName { Val = $"RE Anforderung {level}" });
        style.Append(new BasedOn { Val = level == 1 ? "Normal" : $"RERequirement{level - 1}" });
        style.Append(new PrimaryStyle());

        style.Append(new StyleParagraphProperties(
            new NumberingProperties(
                new NumberingLevelReference { Val = numberingLevel },
                new NumberingId { Val = 2 }
            ),
            new OutlineLevel { Val = numberingLevel },
            new SpacingBetweenLines { Before = "180" }, // 6pt spacing before requirements
            new Justification { Val = JustificationValues.Both },
            new Indentation { Left = TextIndentTwips.ToString(), Hanging = TextIndentTwips.ToString() },
            new Tabs(
                new TabStop { Val = TabStopValues.Left, Position = TextIndentTwips, Leader = TabStopLeaderCharValues.Dot }
            ),
            new NextParagraphStyle { Val = $"RERequirement{level}" }
        ));
        if (level == 1) // other levels inherit their style from level 1
        {
            style.Append(new StyleRunProperties(
                new FontSize { Val = FontSize },
                new FontSizeComplexScript { Val = FontSize }
                ,
                new Languages { Val = "de-DE" }
            ));
        }

        return style;
    }

    /// <summary>
    /// Creates the numbering definitions for Headings, Requirements, and Notes.
    /// Separate abstract numberings for Headings and Requirements to avoid conflicts.
    /// </summary>
    private static Numbering CreateNumbering()
    {
        var numbering = new Numbering();

        // Create abstract numbering for Headings
        var headingAbstractNum = new AbstractNum { AbstractNumberId = 1 };
        headingAbstractNum.Append(new MultiLevelType { Val = MultiLevelValues.Multilevel });

        for (int i = 0; i < HeadingLevels; i++)
        {
            var level = CreateHeadingNumberingLevel(i);
            headingAbstractNum.Append(level);
        }

        numbering.Append(headingAbstractNum);

        // Create abstract numbering for Requirements
        var reqAbstractNum = new AbstractNum { AbstractNumberId = 2 };
        reqAbstractNum.Append(new MultiLevelType { Val = MultiLevelValues.Multilevel });

        int reqLevels = Math.Min(RequirementLevels, 9);
        for (int i = 0; i < reqLevels; i++)
        {
            var level = CreateRequirementNumberingLevel(i);
            reqAbstractNum.Append(level);
        }

        numbering.Append(reqAbstractNum);

        // Create abstract numbering for each note type
        for (int i = 0; i < NoteTypes.Length; i++)
        {
            var (german, english, text) = NoteTypes[i];
            var noteAbstractNum = new AbstractNum { AbstractNumberId = i + 3 }; // Start from 3
            noteAbstractNum.Append(new MultiLevelType { Val = MultiLevelValues.SingleLevel });
            var noteLevel = new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.None }, // No numbering, just text
                new LevelText { Val = text },
                new LevelSuffix { Val = LevelSuffixValues.Nothing }, // No suffix, text is the "bullet"
                new LevelJustification { Val = LevelJustificationValues.Left },
                new PreviousParagraphProperties(
                    new SpacingBetweenLines { Before = "0", After = "0" },
                    new OutlineLevel { Val = 9 }, // Not in outline
                    new Indentation { Left = TextIndentTwips.ToString(), Hanging = "0" }
                    // new Tabs(
                    // new TabStop { Val = TabStopValues.Left, Position = TextIndentTwips, Leader = TabStopLeaderCharValues.Dot }
                    //)
                ),
                new NumberingSymbolRunProperties(
                    //new Bold()
                    new FontSize { Val = "16" }, // 8pt
                    new FontSizeComplexScript { Val = "16" } // 8pt       )
                )
            )
            {
                LevelIndex = 0
            };
            noteAbstractNum.Append(noteLevel);
            numbering.Append(noteAbstractNum);
        }

        // Create the numbering instance for Headings
        var numInstance1 = new NumberingInstance(
            new AbstractNumId { Val = 1 }
        )
        {
            NumberID = 1
        };

        // Create the numbering instance for Requirements
        var numInstance2 = new NumberingInstance(
            new AbstractNumId { Val = 2 }
        )
        {
            NumberID = 2
        };

        numbering.Append(numInstance1);
        numbering.Append(numInstance2);

        // Create numbering instances for note types
        for (int i = 0; i < NoteTypes.Length; i++)
        {
            var numInstance = new NumberingInstance(
                new AbstractNumId { Val = i + 3 }
            )
            {
                NumberID = 10 + i
            };
            numbering.Append(numInstance);
        }

        return numbering;
    }

    /// <summary>
    /// Adds a DocumentSettingsPart containing compatibility settings (compatibilityMode).
    /// </summary>
    private static void AddDocumentSettings(MainDocumentPart mainPart)
    {
        if (mainPart == null) return;

        var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
        // <w:compat><w:compatSetting w:name="compatibilityMode" w:val="15"/></w:compat>
        settingsPart.Settings = new Settings(
            new Compatibility(
                new CompatibilitySetting { Name = CompatSettingNameValues.CompatibilityMode, Val = "15" }
            )
        );
        settingsPart.Settings.Save();
    }

    /// <summary>
    /// Creates a numbering level for Headings.
    /// </summary>
    private static Level CreateHeadingNumberingLevel(int levelIndex)
    {
        // Build the level text format (e.g., "1.", "1.1", "1.1.1", etc.)
        var levelTextBuilder = new StringBuilder();
        for (int j = 0; j <= levelIndex; j++)
        {
            if (j > 0)
                levelTextBuilder.Append('.');
            levelTextBuilder.Append($"%{j + 1}");
        }

        string levelText = (levelIndex == 0 ? levelTextBuilder.ToString() + "." : levelTextBuilder.ToString()) + " ";

        var level = new Level(
            new StartNumberingValue { Val = 1 },
            new NumberingFormat { Val = NumberFormatValues.Decimal },
            new LevelText { Val = levelText },
            new LevelSuffix { Val = LevelSuffixValues.Nothing }, // No suffix for headings
            new LevelJustification { Val = LevelJustificationValues.Left },
            new PreviousParagraphProperties(
                new Indentation { Left = TextIndentTwips.ToString(), Hanging = TextIndentTwips.ToString() },
                new Tabs(
                    new TabStop { Val = TabStopValues.Left, Position = TextIndentTwips, Leader = TabStopLeaderCharValues.Dot }
                )
            )
        )
        {
            LevelIndex = levelIndex
        };

        return level;
    }

    /// <summary>
    /// Creates a numbering level for Requirements.
    /// </summary>
    private static Level CreateRequirementNumberingLevel(int levelIndex)
    {
        // Build the level text format (e.g., "1.", "1.1", "1.1.1", etc.)
        var levelTextBuilder = new StringBuilder();
        for (int j = 0; j <= levelIndex; j++)
        {
            if (j > 0)
                levelTextBuilder.Append('.');
            levelTextBuilder.Append($"%{j + 1}");
        }

        string levelText = levelIndex == 0 ? levelTextBuilder.ToString() + "." : levelTextBuilder.ToString();

        var level = new Level(
            new StartNumberingValue { Val = 1 },
            new NumberingFormat { Val = NumberFormatValues.Decimal },
            new LevelText { Val = levelText },
            new LevelSuffix { Val = LevelSuffixValues.Tab }, // Tab for requirements
            new LevelJustification { Val = LevelJustificationValues.Left },
            new PreviousParagraphProperties(
                new Indentation { Left = TextIndentTwips.ToString(), Hanging = TextIndentTwips.ToString() },
                new Tabs(
                    new TabStop { Val = TabStopValues.Left, Position = TextIndentTwips, Leader = TabStopLeaderCharValues.Dot }
                )
            )
        )
        {
            LevelIndex = levelIndex
        };

        return level;
    }

    /// <summary>
    /// Adds sample content to demonstrate the styles and anchors.
    /// </summary>
    private static void AddSampleContent(Body body)
    {
        // Add a title
        body.Append(CreateTitleParagraph("Anforderungsdokument-Vorlage"));

        // H1 + Requirement1 examples
        body.Append(CreateHeadingParagraph(1, "Einleitung"));
        body.Append(CreateRequirementParagraph(1, "Das System MUSS eine Benutzerverwaltung bereitstellen."));
        body.Append(CreateRequirementParagraph(1, "Das System MUSS Sicherheitsmechanismen implementieren."));

        // Nested requirement
        body.Append(CreateRequirementParagraph(2, "Das System MUSS Benutzerrollen unterstützen."));
        body.Append(CreateRequirementParagraph(3, "Das System MUSS Administratorrechte vergeben können."));

        // H2 + Requirement2 example
        body.Append(CreateHeadingParagraph(2, "Hintergrund"));
        body.Append(CreateRequirementParagraph(2, "Das System SOLL skalierbar sein."));
        body.Append(CreateNoteParagraph("RENote", "Die Skalierbarkeit bezieht sich auf die Fähigkeit des Systems, mit wachsenden Benutzerzahlen und Datenmengen umzugehen. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat."));

        // H1 + Requirement1 again
        body.Append(CreateHeadingParagraph(1, "Funktionale Anforderungen"));
        body.Append(CreateRequirementParagraph(1, "Das System MUSS Benutzerauthentifizierung bereitstellen."));
        body.Append(CreateRequirementParagraph(2, "Benutzer MÜSSEN sich mit Benutzername und Passwort anmelden können."));
        body.Append(CreateRequirementParagraph(2, "Benutzer MÜSSEN ihr Passwort per E-Mail zurücksetzen können."));
        body.Append(CreateRequirementParagraph(1, "Das System MUSS Datenexport unterstützen."));
        body.Append(CreateNoteParagraph("REExample", "Ein Benutzer meldet sich mit dem Benutzernamen 'max.mustermann' und dem Passwort 'geheim123' an. Nach erfolgreicher Authentifizierung erhält er Zugriff auf seine persönlichen Daten. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur."));

        // H2 + Requirement2
        body.Append(CreateHeadingParagraph(2, "Leistungsanforderungen"));
        body.Append(CreateRequirementParagraph(2, "Die Antwortzeit DARF 2 Sekunden nicht überschreiten."));
        body.Append(CreateNoteParagraph("RERationale", "Die Antwortzeit von maximal 2 Sekunden ist notwendig, um eine gute Benutzererfahrung zu gewährleisten. Studien zeigen, dass Benutzer bei längeren Wartezeiten ungeduldig werden und Das System möglicherweise verlassen. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum. Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium."));

        // Demonstrate all heading levels with matching RequirementN
        body.Append(CreateHeadingParagraph(1, "Verschachtelungsbeispiel"));
        body.Append(CreateHeadingParagraph(2, "Ebene 2 Überschrift"));
        body.Append(CreateHeadingParagraph(3, "Ebene 3 Überschrift"));
        body.Append(CreateHeadingParagraph(4, "Ebene 4 Überschrift"));
        body.Append(CreateHeadingParagraph(5, "Ebene 5 Überschrift"));

        // Matching RequirementN under each HN
        body.Append(CreateRequirementParagraph(1, "Das System MUSS modulare Architektur aufweisen."));
        body.Append(CreateNoteParagraph("RENote", "Die modulare Architektur ermöglicht eine bessere Wartbarkeit und Erweiterbarkeit des Systems. Lorem ipsum dolor sit amet, consectetur adipiscing elit."));
        body.Append(CreateAnonParagraph(2, "Zweitens, anonymer Absatz für Ebene 2."));
        body.Append(CreateRequirementParagraph(3, "Module MÜSSEN unabhängig testbar sein."));
        body.Append(CreateRequirementParagraph(4, "Schnittstellen MÜSSEN dokumentiert werden."));
        body.Append(CreateRequirementParagraph(5, "Die Dokumentation MUSS aktuell gehalten werden."));
        body.Append(CreateNoteParagraph("RENote", "Zweiter Hinweis für Ebene 5. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut enim ad minim veniam."));
        body.Append(CreateRequirementParagraph(6, "Testfälle MÜSSEN automatisiert werden."));
        body.Append(CreateRequirementParagraph(7, "Berichte MÜSSEN generiert werden können."));
        body.Append(CreateRequirementParagraph(8, "Die Berichte MÜSSEN exportierbar sein."));

        // Add examples for all note types
        body.Append(CreateNoteParagraph("REReferences", "ISO/IEC 25010:2011 - Systems and software engineering — Systems and software Quality Requirements and Evaluation (SQuaRE) — System and software quality models. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua."));
        body.Append(CreateNoteParagraph("REDerivedFrom", "Kundenanforderung KA-001 vom 15.03.2024. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat."));
    }

    /// <summary>
    /// Creates a title paragraph.
    /// </summary>
    private static Paragraph CreateTitleParagraph(string text)
    {
        return new Paragraph(
            new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "400" }
            ),
            new Run(
                new RunProperties(
                    new Bold(),
                    new FontSize { Val = "48" },
                    new FontSizeComplexScript { Val = "48" }
                ),
                new Text(text)
            )
        );
    }

    /// <summary>
    /// Creates a heading paragraph with a bookmark anchor.
    /// </summary>
    private static Paragraph CreateHeadingParagraph(int level, string text)
    {
        var para = new Paragraph();
        var pPr = new ParagraphProperties(
            new ParagraphStyleId { Val = $"REHeading{level}" }
        );
        para.Append(pPr);
        para.Append(new Run(new Text(text)));
        return para;
    }

    /// <summary>
    /// Creates a requirement paragraph with a bookmark anchor.
    /// </summary>
    private static Paragraph CreateRequirementParagraph(int level, string text)
    {
        var para = new Paragraph();
        var pPr = new ParagraphProperties(
            new ParagraphStyleId { Val = $"RERequirement{level}" }
        );
        para.Append(pPr);
        para.Append(new Run(new Text(text)));
        return para;
    }

    private static Paragraph CreateNoteParagraph(string styleId, string text)
    {
        var para = new Paragraph();
        var pPr = new ParagraphProperties(
            new ParagraphStyleId { Val = styleId }
        );
        para.Append(pPr);
        para.Append(new Run(new Text(text)));
        return para;
    }


    private static Paragraph CreateAnonParagraph(int level, string text)
    {
        var para = new Paragraph();
        var pPr = new ParagraphProperties(
            new ParagraphStyleId { Val = $"REAnonymousPara" }
        );
        para.Append(pPr);
        para.Append(new Run(new Text(text)));
        return para;
    }
}
