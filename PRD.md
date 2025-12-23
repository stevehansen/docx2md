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
- Table detection with merged cell warnings
- Image extraction
- Markdown conversion and export
- Diagnostic report generation
- Unsupported feature detection (headers/footers, comments, track changes, etc.)

### Planned (Future)
- Rich DOCX visual preview
- Visual segment highlighting across panes
- Full override UI controls (heading level, type change)
- Settings persistence between sessions
- Drag-and-drop file opening
- Recent files list

---

## 16. Resolved Questions

| Question | Decision | Rationale |
|----------|----------|-----------|
| DOCX preview strategy | Text approximation (v1), rich rendering (future) | Simpler implementation; visual rendering deferred |
| Default Markdown flavor | GitHub-flavored Markdown (GFM) | GFM tables are widely supported; aligns with Git workflows |
| Override persistence | Not persisted across re-imports (v1) | Future enhancement; requires session/project file format |

---
