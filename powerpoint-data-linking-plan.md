# New Repository Plan: PowerPoint ↔ Excel Cell Linking Add-in

## Decision
Yes — this should be a **new repository** instead of living in this website repo.

### Why a separate repo
- Different runtime and tooling (Office add-in + service) than this static site.
- Independent CI/CD, release cadence, and issue tracking.
- Security-sensitive local file access requires isolated review and permissions model.

## Proposed repository
- **Name:** `ppt-excel-link-addin`
- **Purpose:** Think-cell-style linking of PowerPoint text to specific Excel cells in external files.
- **Target v1:** Windows desktop PowerPoint with manual refresh.

## Product goal
Allow a presenter to bind a text shape in PowerPoint to a specific Excel cell (`Workbook + Sheet + Cell`) and refresh it on demand.

## v1 scope
### In scope
1. Create/edit/delete links from a task pane.
2. Store links inside the `.pptx` (document settings/custom XML).
3. Refresh one or all links on user command.
4. Clear per-link status/errors.

### Out of scope
- Charts/table embedding.
- Background auto-sync.
- Full cross-platform parity.
- Complex formatting inheritance.

## High-level architecture
1. **Office.js PowerPoint add-in (task pane)**
   - UX for link management.
   - Shape selection and text replacement.
2. **Connector service (local companion)**
   - Reads local Excel files safely.
   - Returns value + type metadata.
3. **Shared contract package**
   - DTOs for link definitions, refresh requests, and error payloads.

## Recommended repo layout
```text
ppt-excel-link-addin/
  README.md
  docs/
    architecture.md
    security.md
    roadmap.md
  addin/
    manifest.xml
    package.json
    src/
      taskpane/
      commands/
      services/
  connector/
    package.json
    src/
      api/
      excel/
      validation/
  packages/
    contracts/
      src/
  .github/
    workflows/
      ci.yml
```

## Link data model (v1)
```json
{
  "linkId": "uuid",
  "slideId": "string",
  "shapeId": "string",
  "excelFile": "C:/path/report.xlsx",
  "worksheet": "Summary",
  "cell": "B12",
  "lastValue": "123.45",
  "lastRefreshedUtc": "2026-01-01T10:00:00Z",
  "status": "ok|warning|error",
  "errorCode": "optional"
}
```

## User flow
1. User opens deck and launches add-in.
2. User selects a text shape.
3. User chooses workbook path + sheet + cell.
4. Add-in validates and saves link metadata.
5. User clicks **Refresh**.
6. Add-in updates text and shows status.

## Milestones
### M0 (2–3 days): Feasibility
- Validate shape selection/text update with Office.js.
- Validate safe storage mechanism for link metadata.
- Validate connector’s ability to read cell values.

### M1 (1 week): End-to-end spike
- One link, one shape, one cell.
- Manual refresh updates text.

### M2 (1–2 weeks): Link manager
- CRUD for links.
- Persist/reload in presentation.
- Status display.

### M3 (1 week): Hardening
- Error taxonomy and retries.
- Logging and diagnostics export.
- Packaging and install docs.

## Security requirements
- Explicit user consent before accessing local files.
- Path allow-list and strict path normalization.
- Signed connector binaries for distribution.
- No silent background refresh in v1.

## Initial backlog for new repo bootstrap
1. Initialize new git repo: `ppt-excel-link-addin`.
2. Add Office add-in scaffold in `addin/`.
3. Add connector HTTP service scaffold in `connector/`.
4. Add shared `contracts` package.
5. Add CI workflow (lint/test/build).
6. Add architecture + threat model docs.

## Definition of done for planning
- New repo approved by stakeholders.
- Layout + architecture agreed.
- Milestone estimates accepted.
- Security constraints documented before coding starts.
