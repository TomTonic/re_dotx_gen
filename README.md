# Requirement Template Generator

A C# program that generates a Microsoft Word template (.dotx) with "Requirement" and "Heading" styles designed for creating hierarchically numbered requirements documents.

## Features

- **Hierarchical Numbering**: Supports multi-level numbering (1., 1.1, 1.2.1, 2.4.3.7, etc.)
- **5 Heading Levels**: Heading 1 through Heading 5 styles with sequential numbering
- **8 Requirement Levels**: Requirement 1 through Requirement 8 styles that continue from heading numbers
- **Consistent Formatting**:
  - Numbers start at the left margin (no indentation)
  - Text starts uniformly at 2 cm
  - Dotted leader line between number and text
  - 11 pt font size throughout (Arial)
- **Bookmark Anchors**: Each heading and requirement can have a bookmark for cross-referencing
- **Hyperlinks**: Create internal links to any bookmarked heading or requirement

## Building

```bash
cd src/RequirementTemplateGenerator
dotnet build
```

## Usage

```bash
# Generate template with default name (RequirementTemplate.dotx)
dotnet run

# Generate template with custom filename
dotnet run -- path/to/output.dotx
```

## Generated Styles

### Heading Styles
- **Heading 1**: Top-level section (e.g., "1. Introduction")
- **Heading 2**: Sub-section (e.g., "1.1 Background")
- **Heading 3-5**: Deeper sub-sections

### Requirement Styles
- **Requirement 1**: First-level requirement under current heading (e.g., "1.1.1")
- **Requirement 2-8**: Nested requirements (e.g., "1.1.1.1", "1.1.1.1.1", etc.)

## How It Works

The template uses a single multi-level numbering definition with 13 levels:
- Levels 0-4: Heading styles (5 levels)
- Levels 5-12: Requirement styles (8 levels)

This design ensures that requirements automatically continue numbering from the heading above. For example:
- If heading is "2.3", the first requirement below it will be "2.3.1"
- Nested requirements continue as "2.3.1.1", "2.3.1.2", etc.

## Using the Template in Word

1. Open the generated `.dotx` file to create a new document based on the template
2. Apply styles from the Styles gallery:
   - Use "Heading 1-5" for section headers
   - Use "Requirement 1-8" for requirements
3. Press Tab after a number to insert the dotted leader and position text at 2 cm
4. To create cross-references:
   - Add a bookmark to the target paragraph
   - Insert a hyperlink pointing to the bookmark name

## Requirements

- .NET 10.0 or later
- DocumentFormat.OpenXml NuGet package (automatically restored during build)

## License

BSD 3-Clause License - see [LICENSE](LICENSE) file for details.