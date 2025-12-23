# docx2md

A professional DOCX to Markdown conversion tool with transparency, control, and traceability.

## Overview

`docx2md` converts Microsoft Word (.docx) documents to high-quality Markdown while preserving document structure, intent, and readability. Unlike batch converters, this tool prioritizes transparency by exposing how each document element is interpreted and translated.

## Features

- **Segment-based conversion** - Document decomposed into logical segments (headings, paragraphs, lists, tables, images)
- **Deterministic output** - Same input + same settings = identical output
- **Comprehensive diagnostics** - All unsupported features are explicitly flagged
- **Configurable conversion** - Control heading detection, formatting, and output style
- **Image extraction** - Embedded images are extracted and linked in Markdown
- **Quality checks** - Validates heading hierarchy, list continuity, and more

## Architecture

The project consists of four main components:

- **Docx2Md.Core** - Core library with parsing, conversion, and export logic
- **Docx2Md.Cli** - Command-line interface for batch conversion
- **Docx2Md.UI** - Cross-platform desktop UI with three-pane workbench (Avalonia-based)
- **Docx2Md.Tests** - Unit tests for core functionality

## Getting Started

### Prerequisites

- .NET 10.0 SDK or later

### Building

```bash
dotnet build
```

### Running Tests

```bash
dotnet test
```

### Using the CLI

```bash
# Convert a DOCX file to Markdown
dotnet run --project src/Docx2Md.Cli/Docx2Md.Cli.csproj document.docx

# Specify output file
dotnet run --project src/Docx2Md.Cli/Docx2Md.Cli.csproj document.docx output.md
```

## Usage

### Desktop UI

The desktop application provides a three-pane workbench interface as specified in the PRD:

```bash
# Run the UI application
dotnet run --project src/Docx2Md.UI/Docx2Md.UI.csproj
```

**Features:**
- **Left Pane**: DOCX preview (read-only view of original document)
- **Middle Pane**: Segment inspector with detailed view of each document segment
  - View segment type, style, content, and diagnostics
  - Toggle segment inclusion/exclusion
  - Override segment properties
- **Right Pane**: Markdown preview with toggle between rendered and raw view
- **Menu Bar**: File operations (Open, Export), View toggles, and Settings
- Cross-platform support (Windows, macOS, Linux) via Avalonia

### Command Line

```bash
docx2md <input.docx> [output.md]
```

**Arguments:**
- `input.docx` - Path to input DOCX file (required)
- `output.md` - Path to output Markdown file (optional, defaults to input filename with .md extension)

**Examples:**

```bash
# Convert with default output name
docx2md document.docx

# Specify output name
docx2md document.docx converted.md
```

### Programmatic API

```csharp
using Docx2Md.Core;
using Docx2Md.Core.Models;

// Create converter with default settings
var converter = new Docx2MdConverter();

// Convert file
var document = converter.ConvertFile("input.docx", "output.md");

// Inspect diagnostics
foreach (var diagnostic in document.GetAllDiagnostics())
{
    Console.WriteLine($"{diagnostic.Level}: {diagnostic.Message}");
}
```

### Custom Settings

```csharp
var settings = new ConversionSettings
{
    EnableStyleBasedHeadingDetection = true,
    InferHeadingsFromFormatting = true,
    ConvertUnderlineToEmphasis = false,
    ImageOutputFolder = "assets/images",
    GenerateDiagnosticReport = true
};

var converter = new Docx2MdConverter(settings);
```

## Supported Features

### Document Elements

- ✅ Headings (style-based and inferred)
- ✅ Paragraphs with basic formatting
- ✅ Lists (bulleted and numbered)
- ✅ Tables (with limitations)
- ✅ Images (extracted and linked)
- ✅ Page breaks (converted to horizontal rules)

### Formatting

- ✅ Bold, italic, underline detection
- ✅ Heading hierarchy
- ✅ List nesting
- ⚠️ Complex table layouts (simplified)

### Diagnostics

The tool detects and reports:
- Headers/footers (ignored)
- Comments (ignored)
- Track changes (ignored)
- Merged cells in tables (lossy)
- Heading hierarchy violations
- Missing image alt text

## Output

The converter generates:

1. **Markdown file** - Clean, GitHub-flavored Markdown
2. **Images folder** - Extracted embedded images
3. **Diagnostics report** - Optional Markdown or JSON report of conversion issues

## Configuration Options

| Setting | Default | Description |
|---------|---------|-------------|
| `EnableStyleBasedHeadingDetection` | true | Use paragraph styles to detect headings |
| `InferHeadingsFromFormatting` | true | Infer headings from bold/large text |
| `ConvertUnderlineToEmphasis` | false | Convert underline to italic |
| `EmitHtmlForUnsupported` | false | Output HTML for unsupported features |
| `AppendLostAndFoundSection` | true | Add excluded content to end of document |
| `MaxHeadingLevel` | 6 | Maximum heading level (1-6) |
| `ImageOutputFolder` | "images" | Folder for extracted images |
| `GenerateDiagnosticReport` | true | Generate diagnostics file |
| `PreserveLineBreaks` | true | Preserve line breaks in paragraphs |

## Architecture Details

### Segment Model

Each document element is represented as a `Segment` with:
- **Unique ID** - For tracking and reference
- **Order index** - Position in document
- **Type** - Heading, Paragraph, ListItem, Table, Image, etc.
- **Metadata** - Style name, outline level, formatting
- **Content** - Extracted text
- **Markdown output** - Converted Markdown
- **Diagnostics** - Warnings and errors
- **Overrides** - User customizations (for UI scenarios)

### Conversion Pipeline

1. **Parse** - Extract segments from DOCX using DocumentFormat.OpenXml
2. **Convert** - Apply rule-based conversion to generate Markdown
3. **Export** - Write Markdown, images, and diagnostics to disk

## Limitations

The following Word features are not supported or have limitations:

- ❌ Headers and footers (detected but not converted)
- ❌ Text boxes and shapes (detected but not converted)
- ❌ SmartArt (detected but not converted)
- ❌ Track changes (detected but not converted)
- ❌ Comments (detected but not converted)
- ❌ Multi-column layouts (detected but not converted)
- ❌ Floating images (detected but not converted)
- ❌ Complex table merging (simplified conversion)
- ❌ Round-trip conversion (Markdown → DOCX not supported)

All unsupported features are explicitly flagged in diagnostics.

## Project Structure

```
docx2md/
├── src/
│   ├── Docx2Md.Core/          # Core library
│   │   ├── Models/             # Domain models
│   │   ├── Parsing/            # DOCX parser
│   │   ├── Conversion/         # Markdown converter
│   │   ├── Diagnostics/        # Diagnostic codes
│   │   └── Export/             # Export functionality
│   ├── Docx2Md.Cli/            # CLI application
│   └── Docx2Md.UI/             # Desktop UI (Avalonia)
│       ├── Views/              # UI views
│       ├── ViewModels/         # View models (MVVM)
│       └── Assets/             # UI assets
├── tests/
│   └── Docx2Md.Tests/          # Unit tests
├── PRD.md                      # Product Requirements Document
└── README.md                   # This file
```

## Contributing

This project implements the specifications defined in [PRD.md](PRD.md).

## License

See [LICENSE](LICENSE) for details.

## Future Enhancements

Planned features (see PRD.md for details):
- ✅ Cross-platform desktop UI with three-pane workbench (Avalonia) - **COMPLETED**
- ✅ Interactive segment inspector - **COMPLETED**
- ✅ Live Markdown preview - **COMPLETED**
- ✅ Per-segment manual overrides - **COMPLETED**
- Advanced DOCX rendering in preview pane
- Enhanced segment highlighting and synchronization
- Advanced table conversion
- Enhanced list numbering support
- Custom conversion rule plugins