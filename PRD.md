# Product Requirements Document (PRD)
## Project: DOCX → Markdown Translation Workbench

---

## 1. Purpose

The purpose of this project is to deliver a professional, semi-automated tool that converts Microsoft Word (`.docx`) documents into high-quality Markdown while preserving document structure, intent, and readability.

Unlike batch converters, this tool prioritizes **transparency, control, and traceability** by exposing how each DOCX element is interpreted and translated.

---

## 2. Goals & Non-Goals

### 2.1 Goals
- Enable accurate DOCX → Markdown conversion with human-in-the-loop correction
- Visualize conversion decisions per document segment
- Clearly surface unsupported or lossy features
- Produce clean, deterministic Markdown suitable for documentation, Git, and static site generators

### 2.2 Non-Goals
- Full visual fidelity reproduction of Word documents
- Round-trip Markdown → DOCX conversion
- Support for all Word layout features (e.g., floating shapes, complex sections)
- Real-time collaborative editing

---

## 3. Target Users

- Software developers
- Technical writers
- Documentation engineers
- Compliance and standards teams
- Knowledge-base maintainers

Users are assumed to understand Markdown and document structure concepts.

---

## 4. High-Level User Experience

The application uses a **three-pane workbench layout**:

| Pane | Description |
|-----|-------------|
| Left | Original DOCX preview (read-only) |
| Middle | Segment-level translation inspector |
| Right | Rendered Markdown preview |

All panes are synchronized by document position.

---

## 5. Core Concepts

### 5.1 Segment-First Model

A DOCX document is decomposed into an ordered list of **segments**, each representing a block-level construct such as:

- Heading
- Paragraph
- List item
- Table
- Image
- Page or section break

Each segment acts as the atomic unit of conversion, inspection, and override.

---

## 6. Functional Requirements

### 6.1 Document Import
- Load `.docx` files from disk
- Extract document content in logical order (paragraphs + tables interleaved)
- Extract embedded images and relationships
- Detect headers, footers, footnotes, endnotes, comments

### 6.2 Segment Classification
Each segment MUST include:
- Unique identifier
- Document order index
- Segment type
- Source metadata (style, outline level, numbering info)
- Extracted content
- Conversion diagnostics

### 6.3 Conversion Engine
- Rule-based (not ML-based)
- Deterministic output
- Style-aware heading detection
- Proper list nesting and numbering
- Table conversion to GitHub-flavored Markdown
- Image extraction with relative links

### 6.4 Unsupported Feature Detection
The system MUST detect and flag:
- Headers / footers
- Text boxes / shapes
- SmartArt
- Track changes
- Comments
- Section breaks
- Multi-column layouts
- Floating images
- Word fields (TOC, page numbers)

Each unsupported feature MUST:
- Generate a diagnostic entry
- Be visible in the segment inspector
- Never silently fail

### 6.5 Segment Inspector (Middle Pane)
- Tabular view of all segments
- Display:
  - Segment type
  - Source style
  - Extracted snippet
  - Markdown output snippet
  - Diagnostics
- Row selection synchronizes left and right panes
- Per-segment overrides:
  - Force heading level
  - Change segment type
  - Exclude from output
  - Manual markdown edit

### 6.6 Markdown Preview (Right Pane)
- Rendered Markdown preview
- Raw Markdown view toggle
- Uses same token source as exporter
- Live update on overrides

### 6.7 DOCX Preview (Left Pane)
- Read-only
- Page-oriented
- Best-effort visual fidelity
- Highlight approximate region corresponding to selected segment

---

## 7. Diagnostics & Quality Controls

### 7.1 Diagnostic Types
- Info
- Warning
- Error

### 7.2 Example Diagnostics
- `HEADER_FOOTER_IGNORED`
- `TABLE_MERGED_CELLS_LOSSY`
- `HEADING_INFERRED`
- `LIST_CONTINUITY_WARNING`
- `INLINE_FORMATTING_LOSS`

### 7.3 Quality Checks
- Heading hierarchy validation
- List continuity validation
- Table shape consistency
- Missing image reference detection
- Excessive style ambiguity warnings

---

## 8. Automation Controls

User-configurable options:
- Enable / disable style-based heading detection
- Infer headings from formatting
- Convert underline to emphasis
- Emit HTML for unsupported constructs
- Append “Lost & Found” appendix section

Settings apply globally but may be overridden per segment.

---

## 9. Export

- Export final Markdown as a single `.md` file
- Export extracted images to a configurable folder
- Optional export of diagnostics report (Markdown or JSON)

---

## 10. Non-Functional Requirements

### 10.1 Performance
- Load documents up to 500 pages
- Segment extraction under 2 seconds for typical documents

### 10.2 Determinism
- Same input + same settings MUST produce identical output

### 10.3 Extensibility
- Conversion rules must be pluggable
- Segment types must be extensible
- Diagnostics catalog must be centralized

### 10.4 Offline-First
- No cloud dependency
- No telemetry by default

---

## 11. Technical Constraints

- Desktop application
- Cross-platform via Avalonia (.NET 10.0)
  - Windows 10/11 (native Win32 backend)
  - macOS (native Cocoa backend)
  - Linux (X11/Wayland backend)
- No dependency on Microsoft Word runtime
- Markdown renderer must support:
  - Tables (GitHub-flavored Markdown)
  - Fenced code blocks
  - Images
  - Extensions via pipeline configuration
- UI Framework: Avalonia 11.x with SukiUI theming
- MVVM Pattern: CommunityToolkit.Mvvm with source generators

---

## 12. Out of Scope (Explicit)

- Bidirectional editing (DOCX ↔ Markdown)
- Collaborative editing
- Visual Markdown editing
- PDF input
- Automatic semantic rewriting

---

## 13. Success Criteria

- Users can confidently trace every Markdown line back to its DOCX origin
- Unsupported features are always explicit
- Markdown output is immediately usable in Git-based workflows
- Manual correction effort is reduced by at least 70% compared to raw converters

---

## 14. Risks & Mitigations

| Risk | Mitigation |
|----|----|
| DOCX layout ambiguity | Segment-first model + diagnostics |
| User distrust of automation | Full transparency + overrides |
| Markdown flavor differences | Configurable renderer pipeline |
| Scope creep | Explicit unsupported list |

---

## 15. Implementation Status

### Implemented (v1.0)
- Three-pane workbench layout
- DOCX parsing with segment extraction
- Style-based heading detection
- List item detection (bullets and numbered)
- Full GFM table conversion:
  - Extracts all cell contents
  - Handles merged cells (GridSpan) with warnings
  - Proper header row and separator generation
  - Escapes pipe characters in cell content
- Image extraction and conversion:
  - Extracts embedded images with relationship ID linking
  - Alt text extraction from DocProperties
  - Correct file extension based on content type
  - Consistent filenames between preview and export
- Empty paragraph collapsing (Word spacing doesn't affect Markdown output)
- Markdown conversion and export
- Diagnostic report generation
- Unsupported feature detection (headers/footers, comments, track changes, etc.)
- Full override UI controls:
  - Heading level override (H1-H6)
  - Segment type override
  - Manual markdown override
  - Exclude from output
  - Live preview updates when overrides change
- Theme-aware Markdown preview (tables, headings, text adapt to light/dark mode)
- Visual segment highlighting across panes:
  - Click-to-select segments in DOCX preview
  - Click-to-select segments in raw Markdown view
  - Synchronized highlighting with segment inspector
  - Semi-transparent highlight overlay for selected segment
- Settings persistence between sessions:
  - Saves to %APPDATA%/docx2md/settings.json
  - Persists UI state (view toggles, raw/rendered mode)
  - Persists conversion settings
  - Debounced auto-save on changes
- Recent files list:
  - Tracks last 10 opened documents
  - File menu submenu with recent files
  - Handles missing files gracefully
  - Most recently opened appears first

### Implemented (v1.1 - Conversion Quality & Workflow)
- Hyperlink preservation:
  - Extracts Word hyperlinks with relationship ID resolution
  - Converts to Markdown `[text](url)` format
  - Supports internal bookmark links (`#anchor`)
- Inline formatting preservation:
  - Bold (`**text**`), italic (`*text*`), strikethrough (`~~text~~`)
  - Per-run formatting tracking (not just paragraph-level)
- Code block detection:
  - Detects monospace fonts (Courier New, Consolas, etc.)
  - Inline code wrapped in backticks
  - Full monospace paragraphs converted to fenced code blocks
  - Configurable monospace font list
- Footnote/endnote conversion:
  - Extracts footnote/endnote definitions from DOCX
  - Converts references to `[^1]` format
  - Appends definitions at end of document
- Custom style mappings:
  - JSON config file at %APPDATA%/docx2md/style-mappings.json
  - Map Word styles to: Heading, CodeBlock, Blockquote, Exclude
  - Custom prefix/suffix support
- Front matter templates:
  - Built-in Hugo and Jekyll templates
  - YAML and TOML format support
  - Dynamic field values from document metadata
- Project files (.docx2md):
  - Save/load segment overrides and settings
  - Content hash for detecting document changes
  - Preserves overrides across re-imports
  - Full UI integration:
    - File menu: Save Project (Ctrl+Shift+S), Open Project (Ctrl+Shift+O)
    - Dirty indicator (*) in title bar for unsaved changes
    - Title bar shows current document name

### Implemented (v1.2 - UI Polish & List Numbering)
- Segment overrides panel compact layout (segment list takes more vertical space)
- Auto-scroll to selected segment in DOCX and Markdown preview panes
- Click-to-select segments in rendered Markdown preview
- Project file save dialog defaults to source DOCX filename with .docx2md extension
- Preserve original numbered list numbering from DOCX (uses actual sequence numbers)
- Recent files list includes both DOCX documents and project files (.docx2md)

### Planned (Future)

#### UI & Productivity
- Rich DOCX visual preview
- Drag-and-drop file opening
- Keyboard shortcuts for segment navigation, view toggles, and override application
- Copy to clipboard without full export
- Search/filter segments by type, content, or diagnostic status
- Bulk override operations (apply same override to multiple selected segments)
- Undo/redo for override changes
- Document statistics panel (word count, segment breakdown, diagnostic summary)
- Diff view (side-by-side comparison of original text vs Markdown output)

#### Known Limitations
- Nested list visual rendering in Markdown preview is approximate (each segment renders independently, losing nesting context; markdown output is correct)

#### UI for New Features
- Style mappings configuration dialog (currently JSON file only)
- Front matter template editor

---

## 16. Resolved Questions

| Question | Decision | Rationale |
|----------|----------|-----------|
| DOCX preview strategy | Text approximation (v1), rich rendering (future) | Simpler implementation; visual rendering deferred |
| Default Markdown flavor | GitHub-flavored Markdown (GFM) | GFM tables are widely supported; aligns with Git workflows |
| Override persistence | Project files (.docx2md) save overrides (v1.1) | Content hash enables safe re-import; overrides restored from project file |

---
