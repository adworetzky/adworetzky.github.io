# PowerPoint ↔ Excel Cell Linking Add-in Plan

## Goal
Create a simple PowerPoint add-in that links text in a slide to a specific Excel cell in a different workbook file, so users can refresh slide text when source values change.

## Desired user experience
1. User opens a PowerPoint presentation and launches the add-in task pane.
2. User selects a text box (or text range) on a slide.
3. User picks an Excel source file and a cell address (for example: `C:\Reports\Q4.xlsx`, `Summary!B12`).
4. Add-in stores the link metadata in the presentation.
5. User clicks **Refresh Links** and linked text updates from Excel.
6. User sees clear status for each link (updated, missing file, invalid sheet/cell, permission issue).

## Non-goals for v1
- Rich chart/table linking.
- Automatic background syncing without user action.
- Cross-platform parity (start with Windows desktop PowerPoint + Excel).
- Advanced formatting propagation from Excel.

## Technical approach (v1)
### 1) Platform choice
- Build as an Office.js PowerPoint task pane add-in.
- Use Office.js APIs to identify selected shape/text and store link metadata.
- Use a backend helper service (or local companion process) to read external Excel files, because direct local file access from Office.js is limited.

### 2) Link model
Store each link record with:
- `linkId` (GUID)
- `presentationId`
- `slideId`
- `shapeId`
- `textRangeHint` (optional for partial text replacement)
- `excelFilePath` (or cloud URL)
- `worksheetName`
- `cellAddress`
- `lastValue`
- `lastRefreshUtc`
- `status`

### 3) Binding strategy in PowerPoint
- On link creation:
  - Capture selected shape ID and full text.
  - Insert a lightweight token marker if partial replacement is needed (e.g., `{{LINK:abc123}}`).
- On refresh:
  - Resolve shape by ID.
  - Replace token or entire text with fetched cell value.

### 4) Excel read strategy
- Preferred v1: Windows-only companion service (Node/.NET) with file system access and Excel parsing library (or Excel COM automation if needed).
- Alternative v1.1: Support OneDrive/SharePoint files through Microsoft Graph endpoints.

### 5) Storage
- Persist link definitions in PowerPoint document settings (for portability with the deck).
- Cache latest refresh values in the same store for offline visibility.

## Proposed architecture
1. **Task pane UI (PowerPoint add-in)**
   - Link manager list
   - “Create Link” flow
   - Refresh controls + status panel
2. **Add-in runtime layer**
   - Slide/shape selection handling
   - Metadata serialization/deserialization
   - Apply text updates
3. **Data connector service**
   - Validates file path/sheet/cell
   - Reads current cell value
   - Returns typed value + formatting hints
4. **Optional local cache/log**
   - Refresh history and diagnostics for troubleshooting

## Milestone plan
### Milestone 0 — Discovery (2–3 days)
- Confirm Office.js PowerPoint APIs for shape selection and text mutation.
- Validate document-level settings size/limits.
- Spike local Excel file access options.
- Produce risk log and recommended stack.

### Milestone 1 — Minimal end-to-end prototype (1 week)
- Hard-code a single link in code.
- Pull value from one Excel file/cell.
- Update selected text box on command.
- Verify refresh loop in PowerPoint desktop.

### Milestone 2 — Link creation + persistence (1–2 weeks)
- UI to create/edit/delete links.
- Save/load links in presentation settings.
- Link list with health states.

### Milestone 3 — Robust refresh + errors (1 week)
- Batch refresh all links.
- Error categorization (file missing, no access, bad address, type mismatch).
- Last refreshed timestamp + clear per-link status.

### Milestone 4 — v1 hardening (1 week)
- Logging and diagnostics export.
- Basic security review (path handling, permissions, data exposure).
- Packaging + deployment docs.

## Risks and mitigations
- **Office.js API constraints**: prototype early with selection/text operations.
- **Local file access restrictions**: keep connector as explicit companion service.
- **Shape identity drift** (copy/paste slide elements): store fallback hints and relink flow.
- **Trust/security concerns**: signed companion app, explicit user consent, least privilege.

## Acceptance criteria for v1
- User can link at least one Excel cell to one PowerPoint text shape.
- Link survives save/reopen of the presentation.
- Manual refresh updates text correctly.
- Add-in shows actionable errors when refresh fails.
- Setup instructions allow a new user to run the workflow on Windows.

## Next-step backlog (post-v1)
- Multi-cell templating (`Revenue: {B12}, Margin: {B13}`).
- Scheduled refresh.
- OneDrive/SharePoint-native source picker.
- Formatting rules (number, currency, percent, date).
- Link conflict handling for collaborative edits.
