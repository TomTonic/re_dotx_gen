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
    private const string FontSize = "22"; // 11pt = 22 half-points
    private const int TextIndentTwips = 1134; // 2cm = 1134 twips (20 twips per point, 567 twips per cm)
    private const int HeadingLevels = 5;
    private const int RequirementLevels = 8;

    public static void Main(string[] args)
    {
        string outputPath = args.Length > 0 ? args[0] : "RequirementTemplate.dotx";
        
        Console.WriteLine($"Generating Word template: {outputPath}");
        GenerateTemplate(outputPath);
        Console.WriteLine($"Template generated successfully: {outputPath}");
    }

    /// <summary>
    /// Generates the Word template with all required styles and numbering definitions.
    /// </summary>
    public static void GenerateTemplate(string filePath)
    {
        using var document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Template);

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

    /// <summary>
    /// Creates all the styles for the document (Normal, Headings, Requirements).
    /// </summary>
    private static Styles CreateStyles()
    {
        var styles = new Styles();

        // Add document defaults
        styles.Append(CreateDocDefaults());

        // Add Normal style (base for all)
        styles.Append(CreateNormalStyle());

        // Add Heading styles (1-5)
        for (int level = 1; level <= HeadingLevels; level++)
        {
            styles.Append(CreateHeadingStyle(level));
        }

        // Add Requirement styles (1-8)
        for (int level = 1; level <= RequirementLevels; level++)
        {
            styles.Append(CreateRequirementStyle(level));
        }

        return styles;
    }

    /// <summary>
    /// Creates document defaults with 11pt font.
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
                    new SpacingBetweenLines { After = "200", Line = "276", LineRule = LineSpacingRuleValues.Auto }
                )
            )
        );
    }

    /// <summary>
    /// Creates the Normal base style.
    /// </summary>
    private static Style CreateNormalStyle()
    {
        return new Style(
            new StyleName { Val = "Normal" },
            new PrimaryStyle(),
            new StyleRunProperties(
                new RunFonts { Ascii = FontName, HighAnsi = FontName, ComplexScript = FontName },
                new FontSize { Val = FontSize },
                new FontSizeComplexScript { Val = FontSize }
            )
        )
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true
        };
    }

    /// <summary>
    /// Creates a Heading style for the specified level.
    /// </summary>
    private static Style CreateHeadingStyle(int level)
    {
        var style = new Style(
            new StyleName { Val = $"Heading {level}" },
            new BasedOn { Val = "Normal" },
            new NextParagraphStyle { Val = "Normal" },
            new PrimaryStyle(),
            new StyleParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = level - 1 },
                    new NumberingId { Val = 1 }
                ),
                new OutlineLevel { Val = level - 1 },
                new SpacingBetweenLines { Before = level == 1 ? "480" : "240", After = "120" },
                new Indentation { Left = "0", Hanging = "0" },
                new Tabs(
                    new TabStop { Val = TabStopValues.Left, Position = TextIndentTwips, Leader = TabStopLeaderCharValues.Dot }
                )
            ),
            new StyleRunProperties(
                new Bold(),
                new FontSize { Val = level switch { 1 => "32", 2 => "28", 3 => "26", 4 => "24", _ => FontSize } },
                new FontSizeComplexScript { Val = level switch { 1 => "32", 2 => "28", 3 => "26", 4 => "24", _ => FontSize } }
            )
        )
        {
            Type = StyleValues.Paragraph,
            StyleId = $"Heading{level}"
        };

        return style;
    }

    /// <summary>
    /// Creates a Requirement style for the specified level.
    /// Requirements use numbering that continues from heading levels.
    /// </summary>
    private static Style CreateRequirementStyle(int level)
    {
        // Requirement styles use levels HeadingLevels + level - 1 in the numbering
        // This allows them to continue from heading numbering
        int numberingLevel = HeadingLevels + level - 1;

        var style = new Style(
            new StyleName { Val = $"Requirement {level}" },
            new BasedOn { Val = "Normal" },
            new NextParagraphStyle { Val = $"Requirement{level}" },
            new PrimaryStyle(),
            new StyleParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = numberingLevel },
                    new NumberingId { Val = 1 }
                ),
                new SpacingBetweenLines { Before = "60", After = "60" },
                // No left indent for the number, but text starts at 2cm
                new Indentation { Left = "0", Hanging = "0" },
                new Tabs(
                    new TabStop { Val = TabStopValues.Left, Position = TextIndentTwips, Leader = TabStopLeaderCharValues.Dot }
                )
            ),
            new StyleRunProperties(
                new RunFonts { Ascii = FontName, HighAnsi = FontName, ComplexScript = FontName },
                new FontSize { Val = FontSize },
                new FontSizeComplexScript { Val = FontSize }
            )
        )
        {
            Type = StyleValues.Paragraph,
            StyleId = $"Requirement{level}"
        };

        return style;
    }

    /// <summary>
    /// Creates the numbering definitions for both Headings and Requirements.
    /// This creates a single abstract numbering that supports 13 levels total (5 headings + 8 requirements).
    /// </summary>
    private static Numbering CreateNumbering()
    {
        var numbering = new Numbering();

        // Create the abstract numbering definition
        var abstractNum = new AbstractNum { AbstractNumberId = 1 };
        abstractNum.Append(new MultiLevelType { Val = MultiLevelValues.Multilevel });

        int totalLevels = HeadingLevels + RequirementLevels;

        for (int i = 0; i < totalLevels; i++)
        {
            var level = CreateNumberingLevel(i, totalLevels);
            abstractNum.Append(level);
        }

        numbering.Append(abstractNum);

        // Create the numbering instance
        var numInstance = new NumberingInstance(
            new AbstractNumId { Val = 1 }
        )
        {
            NumberID = 1
        };

        numbering.Append(numInstance);

        return numbering;
    }

    /// <summary>
    /// Creates a numbering level with the appropriate format.
    /// </summary>
    private static Level CreateNumberingLevel(int levelIndex, int totalLevels)
    {
        // Build the level text format (e.g., "1.", "1.1", "1.1.1", etc.)
        var levelTextBuilder = new StringBuilder();
        for (int j = 0; j <= levelIndex; j++)
        {
            if (j > 0)
                levelTextBuilder.Append('.');
            levelTextBuilder.Append($"%{j + 1}");
        }

        // Add trailing period only for single-level (first heading)
        string levelText = levelIndex == 0 ? levelTextBuilder.ToString() + "." : levelTextBuilder.ToString();

        var level = new Level(
            new StartNumberingValue { Val = 1 },
            new NumberingFormat { Val = NumberFormatValues.Decimal },
            new LevelText { Val = levelText },
            new LevelJustification { Val = LevelJustificationValues.Left },
            new PreviousParagraphProperties(
                // Number starts at position 0, text starts at 2cm with dot leader
                new Indentation { Left = TextIndentTwips.ToString(), Hanging = TextIndentTwips.ToString() },
                new Tabs(
                    new TabStop { Val = TabStopValues.Left, Position = TextIndentTwips, Leader = TabStopLeaderCharValues.Dot }
                )
            ),
            new NumberingSymbolRunProperties(
                new RunFonts { Ascii = FontName, HighAnsi = FontName, ComplexScript = FontName },
                new FontSize { Val = FontSize },
                new FontSizeComplexScript { Val = FontSize }
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

        // Heading 1
        body.Append(CreateHeadingParagraph(1, "Introduction", "heading_1"));

        // Requirement under Heading 1
        body.Append(CreateRequirementParagraph(1, "This is a top-level requirement under the introduction.", "req_1_1"));
        body.Append(CreateRequirementParagraph(1, "This is another top-level requirement.", "req_1_2"));

        // Nested requirement
        body.Append(CreateRequirementParagraph(2, "This is a nested requirement (level 2).", "req_1_2_1"));
        body.Append(CreateRequirementParagraph(3, "This is a deeply nested requirement (level 3).", "req_1_2_1_1"));

        // Heading 2 under Heading 1
        body.Append(CreateHeadingParagraph(2, "Background", "heading_1_1"));
        body.Append(CreateRequirementParagraph(1, "Background requirement.", "req_1_1_1"));

        // Another top-level Heading
        body.Append(CreateHeadingParagraph(1, "Functional Requirements", "heading_2"));
        body.Append(CreateRequirementParagraph(1, "The system shall provide user authentication.", "req_2_1"));
        body.Append(CreateRequirementParagraph(2, "Users shall be able to log in with username and password.", "req_2_1_1"));
        body.Append(CreateRequirementParagraph(2, "Users shall be able to reset their password via email.", "req_2_1_2"));
        body.Append(CreateRequirementParagraph(1, "The system shall support data export.", "req_2_2"));

        // Heading 2
        body.Append(CreateHeadingParagraph(2, "Performance Requirements", "heading_2_1"));
        body.Append(CreateRequirementParagraph(1, "Response time shall be under 2 seconds.", "req_2_1_perf_1"));

        // Demonstrate all heading levels
        body.Append(CreateHeadingParagraph(1, "Deep Nesting Example", "heading_3"));
        body.Append(CreateHeadingParagraph(2, "Level 2 Heading", "heading_3_1"));
        body.Append(CreateHeadingParagraph(3, "Level 3 Heading", "heading_3_1_1"));
        body.Append(CreateHeadingParagraph(4, "Level 4 Heading", "heading_3_1_1_1"));
        body.Append(CreateHeadingParagraph(5, "Level 5 Heading", "heading_3_1_1_1_1"));

        // Requirements at deep levels
        body.Append(CreateRequirementParagraph(1, "Requirement at level 1 under deep heading.", "deep_req_1"));
        body.Append(CreateRequirementParagraph(2, "Requirement at level 2.", "deep_req_2"));
        body.Append(CreateRequirementParagraph(3, "Requirement at level 3.", "deep_req_3"));
        body.Append(CreateRequirementParagraph(4, "Requirement at level 4.", "deep_req_4"));
        body.Append(CreateRequirementParagraph(5, "Requirement at level 5.", "deep_req_5"));
        body.Append(CreateRequirementParagraph(6, "Requirement at level 6.", "deep_req_6"));
        body.Append(CreateRequirementParagraph(7, "Requirement at level 7.", "deep_req_7"));
        body.Append(CreateRequirementParagraph(8, "Requirement at level 8.", "deep_req_8"));

        // Add instructions
        body.Append(new Paragraph(new Run(new Text(" "))));
        body.Append(new Paragraph(
            new Run(
                new RunProperties(new Italic()),
                new Text("Instructions: Apply the 'Heading X' styles for section headers and 'Requirement X' styles for requirements. Use Tab to adjust indent levels.")
            )
        ));

        // Add hyperlink example
        body.Append(new Paragraph(
            new Run(new Text("Example cross-reference: See ")),
            CreateHyperlinkToBookmark("req_2_1", "Requirement 2.1"),
            new Run(new Text(" for authentication requirements."))
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
    private static Paragraph CreateHeadingParagraph(int level, string text, string bookmarkId)
    {
        var para = new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId { Val = $"Heading{level}" }
            )
        );

        // Add bookmark start
        string bookmarkIdNum = GetBookmarkId(bookmarkId);
        para.Append(new BookmarkStart { Id = bookmarkIdNum, Name = bookmarkId });

        // Add the text with a tab to position it at 2cm
        para.Append(new Run(
            new TabChar(),
            new Text(text)
        ));

        // Add bookmark end
        para.Append(new BookmarkEnd { Id = bookmarkIdNum });

        return para;
    }

    /// <summary>
    /// Creates a requirement paragraph with a bookmark anchor.
    /// </summary>
    private static Paragraph CreateRequirementParagraph(int level, string text, string bookmarkId)
    {
        var para = new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId { Val = $"Requirement{level}" }
            )
        );

        // Add bookmark start
        string bookmarkIdNum = GetBookmarkId(bookmarkId);
        para.Append(new BookmarkStart { Id = bookmarkIdNum, Name = bookmarkId });

        // Add the text with a tab to position it at 2cm
        para.Append(new Run(
            new TabChar(),
            new Text(text)
        ));

        // Add bookmark end
        para.Append(new BookmarkEnd { Id = bookmarkIdNum });

        return para;
    }

    /// <summary>
    /// Creates a hyperlink to a bookmark.
    /// </summary>
    private static Hyperlink CreateHyperlinkToBookmark(string bookmarkName, string displayText)
    {
        return new Hyperlink(
            new Run(
                new RunProperties(
                    new Underline { Val = UnderlineValues.Single },
                    new Color { Val = "0000FF" }
                ),
                new Text(displayText)
            )
        )
        {
            Anchor = bookmarkName
        };
    }

    /// <summary>
    /// Generates a numeric bookmark ID from a string name.
    /// </summary>
    private static readonly Dictionary<string, string> _bookmarkIds = new();
    private static int _nextBookmarkId = 0;

    private static string GetBookmarkId(string name)
    {
        if (!_bookmarkIds.TryGetValue(name, out string? id))
        {
            id = (_nextBookmarkId++).ToString();
            _bookmarkIds[name] = id;
        }
        return id;
    }
}
