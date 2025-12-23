# DOCX to Markdown UI Guide

This guide provides detailed information about the desktop UI implementation for the DOCX to Markdown converter.

## Overview

The UI is built using **Avalonia**, a cross-platform .NET UI framework that runs on Windows, macOS, and Linux. This implementation follows the specifications outlined in the PRD (Product Requirements Document).

## Architecture

### Technology Stack

- **Avalonia 11.3.10**: Cross-platform UI framework
- **CommunityToolkit.Mvvm 8.2.1**: MVVM pattern implementation with source generators
- **Markdown.Avalonia 11.0.2**: Markdown rendering in preview pane
- **Docx2Md.Core**: Core conversion library

### Design Pattern

The UI uses the **Model-View-ViewModel (MVVM)** pattern:

- **Views**: XAML files defining the UI layout (`MainWindow.axaml`)
- **ViewModels**: C# classes containing UI logic and state (`MainWindowViewModel.cs`)
- **Models**: Domain models from `Docx2Md.Core` (Segment, DocumentModel, etc.)

## Three-Pane Workbench Layout

As specified in the PRD (Section 4), the UI provides three synchronized panes:

### Left Pane: DOCX Preview

- **Purpose**: Display read-only preview of the original DOCX document
- **Features**:
  - Document title display
  - Text representation of document content
  - Segment-based preview
- **Current Implementation**: Simple text preview showing all segments
- **Future Enhancement**: Rich visual rendering with segment highlighting

### Middle Pane: Segment Inspector

- **Purpose**: Provide detailed view and control over each document segment
- **Features**:
  - DataGrid displaying all segments in order
  - Columns:
    - `#`: Order index
    - `Type`: Segment type (Heading, Paragraph, ListItem, etc.)
    - `Style`: Source Word style name
    - `Content`: Extracted text content
    - `Markdown`: Generated Markdown output
    - `Diagnostics`: Count of warnings/errors for this segment
    - `Exclude`: Checkbox to exclude segment from output
  - Row selection synchronizes with other panes
  - Editable checkboxes for segment exclusion

### Right Pane: Markdown Preview

- **Purpose**: Show rendered or raw Markdown output
- **Features**:
  - Toggle button to switch between rendered and raw view
  - Rendered view: Uses Markdown.Avalonia for rich rendering
  - Raw view: Plain text with monospace font for editing/inspection
  - Live updates when segments change

## Menu Bar

### File Menu

- **Open DOCX...** (Ctrl+O)
  - Opens file picker to select a DOCX file
  - Loads and parses the document
  - Displays segments in inspector
  - Updates all panes

- **Export Markdown...** (Ctrl+E)
  - Opens save file dialog
  - Exports current Markdown to selected location
  - Includes all non-excluded segments
  - Only enabled when a document is loaded

- **Exit**
  - Closes the application using proper Avalonia shutdown

### View Menu

- **Show DOCX Preview**: Toggle left pane visibility
- **Show Segment Inspector**: Toggle middle pane visibility
- **Show Markdown Preview**: Toggle right pane visibility

### Settings Menu

Configuration options that affect conversion (matches PRD Section 8):

- **Enable Style-Based Heading Detection**: Use paragraph styles to identify headings
- **Infer Headings From Formatting**: Detect headings from bold/large text
- **Convert Underline to Emphasis**: Map underline formatting to italic
- **Generate Diagnostic Report**: Include diagnostics in export

**Note**: Settings changes automatically reconvert the current document if one is loaded.

## Status Bar

Located at the bottom of the window, displays:
- Current operation status
- Document load status
- Number of segments and diagnostics
- Error messages

## Getting Started

### Running the Application

```bash
cd /path/to/docx2md
dotnet run --project src/Docx2Md.UI/Docx2Md.UI.csproj
```

### Sample Document

On startup, the application loads a sample document with:
- Sample heading
- Paragraph text
- Sub-heading
- List item with diagnostic

This allows immediate exploration of the UI without needing a DOCX file.

### Opening a Document

1. Click **File → Open DOCX...** or press **Ctrl+O**
2. Select a `.docx` file from your system
3. The document will be parsed and displayed in all three panes
4. Segments appear in the inspector with diagnostics

### Inspecting Segments

1. Click on any row in the Segment Inspector (middle pane)
2. View the full content and Markdown output
3. Check/uncheck the "Exclude" column to include/exclude segments
4. View diagnostic messages for issues or warnings

### Exporting Markdown

1. Ensure a document is loaded
2. Click **File → Export Markdown...** or press **Ctrl+E**
3. Choose a location and filename
4. The Markdown file will be saved with all non-excluded segments

## Per-Segment Features

Each segment in the inspector supports:

- **Exclusion**: Checkbox to exclude from final output
- **Type Override**: Change segment type (via `OverrideType` property)
- **Heading Level Override**: Force specific heading level (via `OverrideHeadingLevel` property)
- **Manual Markdown**: Replace auto-generated Markdown (via `ManualMarkdownOverride` property)

**Current Implementation**: Exclusion checkbox is functional. Other overrides are supported in the data model but don't yet have UI controls.

## Diagnostics

Diagnostics appear in the segment inspector showing:
- Info, Warning, or Error level
- Diagnostic code (e.g., `LIST_DETECTED`, `HEADING_INFERRED`)
- Human-readable message
- Optional details

Examples from PRD Section 7.2:
- `HEADER_FOOTER_IGNORED`: Headers/footers detected but not converted
- `TABLE_MERGED_CELLS_LOSSY`: Complex table structure simplified
- `HEADING_INFERRED`: Heading detected from formatting rather than style
- `LIST_CONTINUITY_WARNING`: List numbering may be inconsistent

## Keyboard Shortcuts

- **Ctrl+O**: Open DOCX file
- **Ctrl+E**: Export Markdown
- **Tab**: Navigate between UI elements
- **Arrow Keys**: Navigate segment list

## Cross-Platform Support

The UI runs on:
- **Windows**: Full support, native file dialogs
- **macOS**: Full support, native file dialogs
- **Linux**: Full support, GTK-based file dialogs

## Customization

### Changing Settings

Settings can be toggled via the Settings menu. Changes take effect immediately:
- If a document is loaded, it will be reconverted with new settings
- Future document loads will use the new settings

### Hiding Panes

Use the View menu to hide panes you don't need:
- Hide DOCX preview to focus on conversion
- Hide Markdown preview to focus on segment inspection
- Keep only the panes relevant to your workflow

## Known Limitations

1. **DOCX Preview**: Currently shows simple text representation, not full visual fidelity
2. **Segment Highlighting**: Selection doesn't yet highlight in DOCX/Markdown previews
3. **Manual Overrides**: UI controls for heading level and type overrides not yet implemented
4. **Drag-and-Drop**: Not yet supported for opening files
5. **Zoom Controls**: Not yet implemented for previews

## Future Enhancements

Planned improvements:
- Rich DOCX rendering in left pane
- Segment highlighting synchronized across all panes
- UI controls for all per-segment override options
- Drag-and-drop file opening
- Export settings (image folder, diagnostic format)
- Recent files list
- Settings persistence
- Keyboard-only workflow support

## Troubleshooting

### Application Won't Start

- Ensure .NET 10.0 SDK is installed
- Check that all NuGet packages are restored: `dotnet restore`
- Verify no port conflicts if running multiple instances

### File Dialog Doesn't Appear

- On Linux, ensure GTK dependencies are installed
- On macOS, check file system permissions

### Markdown Preview Shows Raw Text

- Ensure Markdown.Avalonia package is installed
- Check that "Raw" toggle is not enabled
- Verify Markdown content is valid

### Segments Not Updating

- Verify document was loaded successfully (check status bar)
- Try closing and reopening the document
- Check for errors in the diagnostics column

## Contributing

When contributing to the UI:

1. Follow MVVM pattern strictly
2. Use `ObservableProperty` from CommunityToolkit.Mvvm
3. Keep ViewModels testable (inject dependencies)
4. Update this guide when adding features
5. Test on multiple platforms when possible

## References

- [PRD.md](PRD.md): Product Requirements Document
- [Avalonia Documentation](https://docs.avaloniaui.net/)
- [CommunityToolkit.Mvvm Docs](https://learn.microsoft.com/en-us/dotnet/communitytoolkit/mvvm/)
- [Markdown.Avalonia GitHub](https://github.com/whistyun/Markdown.Avalonia)
