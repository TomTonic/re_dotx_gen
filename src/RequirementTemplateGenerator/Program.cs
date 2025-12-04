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
            new SpacingBetweenLines { Before = "120" },
            new NextParagraphStyle { Val = "REAnonymousPara" }
        ));
        anonymousParaStyle.Append(new StyleRunProperties(
            new Italic()
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
            noteStyle.Append(new BasedOn { Val = "Normal" });
            noteStyle.Append(new PrimaryStyle());
            noteStyle.Append(new StyleParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = 10 + i } // 10 for first, 11 for second, etc.
                ),
                new SpacingBetweenLines { Before = "120" },
                new NextParagraphStyle { Val = "REAnonymousPara" }
            ));
            noteStyle.Append(new StyleRunProperties(
                new Italic()
            ));
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
            new SpacingBetweenLines { Before = "360", After = "120" }, // 6pt spacing before and after headings
            new Indentation { Left = TextIndentTwips.ToString(), Hanging = TextIndentTwips.ToString() },
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
                    new Indentation { Left = TextIndentTwips.ToString(), Hanging = "0" },
                    new Tabs(
                        new TabStop { Val = TabStopValues.Left, Position = TextIndentTwips, Leader = TabStopLeaderCharValues.Dot }
                    )
                ),
                new NumberingSymbolRunProperties(
                    new Bold()
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
        body.Append(CreateTitleParagraph("Requirement Document Template"));

        // H1 + Requirement1 examples
        body.Append(CreateHeadingParagraph(1, "Introduction"));
        body.Append(CreateRequirementParagraph(1, "Requirement for H1: top-level requirement."));
        body.Append(CreateRequirementParagraph(1, "Requirement for H1: another top-level requirement."));

        // Nested requirement
        body.Append(CreateRequirementParagraph(2, "This is a nested requirement (level 2)."));
        body.Append(CreateRequirementParagraph(3, "This is a deeply nested requirement (level 3)."));

        // H2 + Requirement2 example
        body.Append(CreateHeadingParagraph(2, "Background"));
        body.Append(CreateRequirementParagraph(2, "Requirement for H2."));

        // H1 + Requirement1 again
        body.Append(CreateHeadingParagraph(1, "Functional Requirements"));
        body.Append(CreateRequirementParagraph(1, "The system shall provide user authentication."));
        body.Append(CreateRequirementParagraph(2, "Users shall be able to log in with username and password."));
        body.Append(CreateRequirementParagraph(2, "Users shall be able to reset their password via email."));
        body.Append(CreateRequirementParagraph(1, "The system shall support data export."));

        // H2 + Requirement2
        body.Append(CreateHeadingParagraph(2, "Performance Requirements"));
        body.Append(CreateRequirementParagraph(2, "Response time shall be under 2 seconds."));

        // Demonstrate all heading levels with matching RequirementN
        body.Append(CreateHeadingParagraph(1, "Deep Nesting Example"));
        body.Append(CreateHeadingParagraph(2, "Level 2 Heading"));
        body.Append(CreateHeadingParagraph(3, "Level 3 Heading"));
        body.Append(CreateHeadingParagraph(4, "Level 4 Heading"));
        body.Append(CreateHeadingParagraph(5, "Level 5 Heading"));

        // Matching RequirementN under each HN
        body.Append(CreateRequirementParagraph(1, "Requirement for H1 under deep example."));
        body.Append(CreateNoteParagraph("RENote", "Some hints on Level 2."));
        body.Append(CreateAnonParagraph(2, "Second, anonymous for H2."));
        body.Append(CreateRequirementParagraph(3, "Requirement for H3."));
        body.Append(CreateRequirementParagraph(4, "Requirement for H4."));
        body.Append(CreateRequirementParagraph(5, "Requirement at level 5."));
        body.Append(CreateNoteParagraph("RENote", "Second Note for level 5."));
        body.Append(CreateRequirementParagraph(6, "Requirement at level 6."));
        body.Append(CreateRequirementParagraph(7, "Requirement at level 7."));
        body.Append(CreateRequirementParagraph(8, "Requirement at level 8."));

        // Add examples for all note types
        for (int i = 0; i < NoteTypes.Length; i++)
        {
            var (german, english, _) = NoteTypes[i];
            body.Append(CreateNoteParagraph($"RE{english}", $"Example for {german}."));
        }

        // Add instructions
        body.Append(new Paragraph(new Run(new Text(" "))));
        body.Append(new Paragraph(
            new Run(
                new RunProperties(new Italic()),
                new Text("Instructions: Apply the 'Heading X' styles for section headers and 'Requirement X' styles for requirements. Use Tab to adjust indent levels.")
            )
        ));

        // Add cross-reference example (plain text; bookmarking removed)
        body.Append(new Paragraph(
            new Run(new Text("Example cross-reference: See Requirement 2.1 for authentication requirements."))
        ));
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
