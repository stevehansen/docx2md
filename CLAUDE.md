# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**docx2md** is a DOCX to Markdown conversion tool with a three-pane desktop UI. It uses a segment-first model where documents are decomposed into atomic segments (headings, paragraphs, lists, tables, images) for transparent, controllable conversion.

## Build & Run Commands

```bash
# Build all projects
dotnet build

# Run tests
dotnet test

# Run CLI (converts document.docx to Markdown)
dotnet run --project src/Docx2Md.Cli/Docx2Md.Cli.csproj document.docx [output.md]

# Run desktop UI
dotnet run --project src/Docx2Md.UI/Docx2Md.UI.csproj
```

## Architecture

### Project Structure

- **Docx2Md.Core** - Core library with parsing, conversion, and export logic (no UI dependencies)
- **Docx2Md.Cli** - Command-line interface for batch conversion
- **Docx2Md.UI** - Avalonia cross-platform desktop UI with three-pane workbench
- **Docx2Md.Tests** - xUnit test suite

### Conversion Pipeline

```
DOCX → DocxParser → DocumentModel (Segments) → MarkdownConverter → MarkdownExporter → .md file
```

1. **DocxParser** (`Parsing/DocxParser.cs`) - Extracts content from DOCX using DocumentFormat.OpenXml
2. **MarkdownConverter** (`Conversion/MarkdownConverter.cs`) - Transforms segments to Markdown
3. **MarkdownExporter** (`Export/MarkdownExporter.cs`) - Writes outputs and extracts images

### Core Data Model

**Segment** (`Models/Segment.cs`) - The atomic unit of the document:
- `Type`: Heading, Paragraph, ListItem, Table, Image, PageBreak, SectionBreak, Unknown
- `Content`: Original text from Word
- `MarkdownOutput`: Converted Markdown
- `Diagnostics`: Issues found during conversion
- `Overrides`: User customizations (heading level, segment type, manual Markdown)

**DocumentModel** (`Models/DocumentModel.cs`) - Container for all segments with aggregate diagnostics

### UI Architecture (MVVM)

The desktop app uses Avalonia with CommunityToolkit.Mvvm:
- **MainWindowViewModel** - Uses `[ObservableProperty]` and `[RelayCommand]` attributes
- Three-pane layout: DOCX preview | Segment inspector | Markdown preview

## Technology Stack

- .NET 10.0 SDK
- DocumentFormat.OpenXml 3.2.0 (DOCX parsing)
- Avalonia 11.3.10 + SukiUI 6.0.3 (cross-platform UI)
- CommunityToolkit.Mvvm 8.2.1 (MVVM framework)
- xUnit 2.9.3 (testing)

## Key Patterns

- **Segment-first**: All document content is represented as segments before conversion
- **Deterministic**: Same input + settings = identical output (no ML)
- **Diagnostic codes**: Centralized in `DiagnosticCodes.cs` for consistent error/warning reporting
- **ConversionSettings**: Configurable conversion behavior (heading detection, formatting options)

## Known Limitations

- No headers/footers conversion
- No text boxes, shapes, or SmartArt
- Complex table cell merging is simplified
- No floating images support
