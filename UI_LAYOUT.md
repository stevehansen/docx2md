# UI Layout Diagram

```
┌─────────────────────────────────────────────────────────────────────────────────────┐
│ DOCX → Markdown Translation Workbench                                        [─][□][×] │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ File   View   Settings                                                              │
├─────────────────────────────────────────────────────────────────────────────────────┤
│                                                                                     │
│  ┌─────────────────┬─┬───────────────────────────────┬─┬───────────────────────┐  │
│  │                 │ │                               │ │                       │  │
│  │  DOCX Preview   │ │   Segment Inspector           │ │  Markdown Preview     │  │
│  │  ┌────────────┐ │ │   ┌──────────────────────┐   │ │  ┌─────────────────┐  │  │
│  │  │Read-only   │ │ │   │ # │Type │Style│...  │   │ │  │ Rendered │ Raw  │  │  │
│  │  │view of     │ │ │   ├───┼─────┼─────┼─────┤   │ │  ├─────────────────┤  │  │
│  │  │original    │ │ │   │ 0 │Head │H1   │...  │   │ │  │                 │  │  │
│  │  │DOCX        │ │ │   │ 1 │Para │Norm │...  │   │ │  │ # Sample Doc    │  │  │
│  │  │document    │ │ │   │ 2 │Head │H2   │...  │   │ │  │                 │  │  │
│  │  │            │ │ │   │ 3 │List │List │...  │   │ │  │ This is a       │  │  │
│  │  │Shows:      │ │ │   └──────────────────────┘   │ │  │ sample para...  │  │  │
│  │  │• Title     │ │ │                               │ │  │                 │  │  │
│  │  │• Segments  │ │ │   Columns:                    │ │  │ ## Features     │  │  │
│  │  │• Content   │ │ │   • Order Index               │ │  │                 │  │  │
│  │  │            │ │ │   • Type                      │ │  │ - Three-pane    │  │  │
│  │  │            │ │ │   • Style                     │ │  │   workbench     │  │  │
│  │  │            │ │ │   • Content                   │ │  │                 │  │  │
│  │  │            │ │ │   • Markdown                  │ │  │                 │  │  │
│  │  │            │ │ │   • Diagnostics               │ │  │                 │  │  │
│  │  │            │ │ │   • Exclude ☑                 │ │  │                 │  │  │
│  │  │            │ │ │                               │ │  │                 │  │  │
│  │  │            │ │ │   Features:                   │ │  │                 │  │  │
│  │  │            │ │ │   • Row selection             │ │  │                 │  │  │
│  │  │            │ │ │   • Editable exclusion        │ │  │                 │  │  │
│  │  │            │ │ │   • Diagnostics view          │ │  │                 │  │  │
│  │  │            │ │ │   • Segment overrides         │ │  │                 │  │  │
│  │  └────────────┘ │ │                               │ │  └─────────────────┘  │  │
│  └─────────────────┴─┴───────────────────────────────┴─┴───────────────────────┘  │
│                                                                                     │
├─────────────────────────────────────────────────────────────────────────────────────┤
│ Status: Loaded 4 segments. 1 diagnostics.                                          │
└─────────────────────────────────────────────────────────────────────────────────────┘
```

## Component Breakdown

### Menu Bar
- **File**: Open, Export, Exit
- **View**: Toggle pane visibility
- **Settings**: Conversion options (checkboxes for various settings)

### Left Pane (DOCX Preview)
- Displays original document content
- Shows document title
- Lists segments in order
- Read-only view

### Middle Pane (Segment Inspector)
- DataGrid with all document segments
- Interactive column headers
- Sortable columns
- Editable exclusion checkboxes
- Displays diagnostic counts
- Row selection syncs other panes

### Right Pane (Markdown Preview)
- Toggle button for Rendered/Raw view
- Rendered mode: Rich Markdown display
- Raw mode: Editable text with monospace font
- Scrollable content
- Live updates

### Status Bar
- Shows current operation
- Document load status
- Segment and diagnostic counts
- Error messages

## Interaction Flow

```
User Opens File
      ↓
File Dialog → Select DOCX
      ↓
Parse Document
      ↓
├─→ Update DOCX Preview (Left)
├─→ Populate Segment Inspector (Middle)
└─→ Render Markdown Preview (Right)
      ↓
User Selects Segment in Inspector
      ↓
├─→ Highlight in DOCX Preview*
└─→ Scroll to in Markdown Preview*
      
      (* Future enhancement)
```

## Settings Integration

```
User Changes Setting (e.g., "Infer Headings")
      ↓
OnSettingsChanged event
      ↓
If document loaded:
      ↓
├─→ Recreate converter with new settings
├─→ Reconvert document
└─→ Update all panes
      ↓
Status: "Document reconverted with new settings."
```

## Data Flow

```
DOCX File
    ↓
DocxParser (Core)
    ↓
DocumentModel
    ↓
MarkdownConverter (Core)
    ↓
Segments with Markdown
    ↓
ObservableCollection<Segment>
    ↓
├─→ DataGrid Binding (Inspector)
├─→ Text Binding (DOCX Preview)
└─→ Markdown Binding (Preview)
```
