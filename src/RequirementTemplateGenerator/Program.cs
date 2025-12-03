using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.IO;
using System.IO.Compression;

namespace RequirementTemplateGenerator;

/// <summary>
/// Generates a Word template (.dotx) with "Requirement" styles that support hierarchical numbering.
/// </summary>
public class Program
{
    // Constants for formatting
    private const string FontName = "Arial";
    private const string FontSize = "22"; // 11pt = 22 half-points
    private const int TextIndentTwips = 1701; // 3cm = 1701 twips (567 twips per cm)
    private const int HeadingLevels = 5;
    private const int RequirementLevels = 8;

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
    /// Creates a H style for the specified level (H1-H5).
    /// </summary>
    private static Style CreateHeadingStyle(int level)
    {
        var style = new Style
        {
            Type = StyleValues.Paragraph,
            CustomStyle = true,
            StyleId = $"H{level}"
        };

        style.Append(new StyleName { Val = $"H{level}" });
        style.Append(new BasedOn { Val = level == 1 ? "Normal" : $"H{level - 1}" });
        style.Append(new PrimaryStyle());

        style.Append(new StyleParagraphProperties(
            new NumberingProperties(
                new NumberingLevelReference { Val = level - 1 },
                new NumberingId { Val = 1 }
            ),
            new OutlineLevel { Val = level - 1 },
            // new SpacingBetweenLines { Before = level == 1 ? "480" : "240", After = "120" },
            new Indentation { Left = TextIndentTwips.ToString(), Hanging = TextIndentTwips.ToString() },
            new Tabs(
                new TabStop { Val = TabStopValues.Left, Position = TextIndentTwips, Leader = TabStopLeaderCharValues.Dot }
            ),
            new NextParagraphStyle { Val = $"Requirement{level+1}" }
        ));

        style.Append(new StyleRunProperties(
            new RunFonts { Ascii = FontName, HighAnsi = FontName, ComplexScript = FontName },
            new Bold(),
            new FontSize { Val = level switch { 1 => "32", 2 => "28", 3 => "26", 4 => "24", _ => FontSize } },
            new FontSizeComplexScript { Val = level switch { 1 => "32", 2 => "28", 3 => "26", 4 => "24", _ => FontSize } }
        ));

        return style;
    }

    /// <summary>
    /// Creates a Requirement style for the specified level.
    /// Requirements use numbering that continues from heading levels.
    /// </summary>
    private static Style CreateRequirementStyle(int level)
    {
        // Map RequirementN to ilvl N (0-based), clamped to Word's 9-level max
        int numberingLevel = Math.Min(level, 8);

        var style = new Style
        {
            Type = StyleValues.Paragraph,
            CustomStyle = true,
            StyleId = $"Requirement{level}"
        };

        style.Append(new StyleName { Val = $"Requirement {level}" });
        style.Append(new BasedOn { Val = level == 1 ? "Normal" : $"Requirement{level - 1}" });
        style.Append(new PrimaryStyle());

        style.Append(new StyleParagraphProperties(
            new NumberingProperties(
                new NumberingLevelReference { Val = numberingLevel },
                new NumberingId { Val = 1 }
            ),
            new OutlineLevel { Val = numberingLevel },
            // new SpacingBetweenLines { Before = "60", After = "60" },
            new Indentation { Left = TextIndentTwips.ToString(), Hanging = TextIndentTwips.ToString() },
            new Tabs(
                new TabStop { Val = TabStopValues.Left, Position = TextIndentTwips, Leader = TabStopLeaderCharValues.Dot }
            ),
            new NextParagraphStyle { Val = $"Requirement{level}" }
        ));

        style.Append(new StyleRunProperties(
            new RunFonts { Ascii = FontName, HighAnsi = FontName, ComplexScript = FontName },
            new FontSize { Val = FontSize },
            new FontSizeComplexScript { Val = FontSize }
        ));

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

        // Word supports a maximum of 9 multilevel list levels (0-8).
        // Cap the generated levels to avoid Word repair prompts.
        int totalLevels = Math.Min(HeadingLevels + RequirementLevels, 9);

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
            new LevelSuffix { Val = LevelSuffixValues.Tab },
            new LevelJustification { Val = LevelJustificationValues.Left },
            // Note: Binding pStyle here requires the OpenXML type for w:pStyle under w:lvl, which isn't directly exposed.
            // Removing to restore build; will rely on numPr in styles and outlineLvl.
            new PreviousParagraphProperties(
                // Numbers and text aligned: Left = text indent (3cm), Hanging = text indent
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
        body.Append(CreateRequirementParagraph(2, "Requirement for H2."));
        body.Append(CreateRequirementParagraph(3, "Requirement for H3."));
        body.Append(CreateRequirementParagraph(4, "Requirement for H4."));
        body.Append(CreateRequirementParagraph(5, "Requirement at level 5."));
        body.Append(CreateRequirementParagraph(6, "Requirement at level 6."));
        body.Append(CreateRequirementParagraph(7, "Requirement at level 7."));
        body.Append(CreateRequirementParagraph(8, "Requirement at level 8."));

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
            new ParagraphStyleId { Val = $"H{level}" }
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
            new ParagraphStyleId { Val = $"Requirement{level}" }
        );
        para.Append(pPr);
        para.Append(new Run(new Text(text)));
        return para;
    }

}
