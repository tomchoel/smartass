# Design: Headless Excel Generation with Smart View HsGetValue Formulas

**Date:** 2026-10-27  
**Branch:** `spike/headless-hsgetvalue`  
**Status:** Approved  
**Topic:** Headless workbook generation with native Smart View refreshability  

---

## 1. Context & Motivation

Today, the SmartAss MCP engine (`smartview_mcp`) builds Excel workbooks by driving desktop Excel via COM (`pywin32` STA thread). While this guarantees that sheets are recognized as live Smart View ad-hoc grids, it incurs notable trade-offs:
- **Cold-start latency:** 15–30 seconds for hidden Excel launch, SSO dialog clearance, and provider connection.
- **Desktop conflict vulnerabilities:** Process collisions if PowerPoint holds the shared WebView2 cache, or if the user is in cell-edit mode (`0x800AC472`).
- **Resource serialization:** Desktop Excel window locks (`user_excel_<pid>.lock`) restrict parallel generation.

By adopting a **hybrid strategy**, we pull data sub-second over HTTP (via `SvHttpSource`) and assemble `.xlsx` workbooks purely headlessly using Python (`openpyxl`), embedding Oracle Smart View `=HsGetValue(...)` formulas into the data cells alongside pre-populated cached numbers.

This design specifies the proof-of-concept spike to validate whether this approach delivers instant (<1s) generation while preserving the user's ability to inspect source numbers and natively click **Smart View → Refresh** in Excel.

---

## 2. Architecture & Data Flow

```
       [ Oracle EPM Cloud ]
                 │
                 │ 1. Sub-second retrieve (0.2s)
                 ▼
  [ SvHttpSource (HTTP Provider) ]
                 │
                 │ 2. Extracted cell values & POV coordinates
                 ▼
    [ Headless Generator (openpyxl) ]
         ├── Injects =HsGetValue(...) formulas
         ├── Injects cached .value (so sheet opens pre-populated)
         └── Stamps MIP sensitivity label XML
                 │
                 │ 3. Generates .xlsx directly to disk (< 0.5s)
                 ▼
     [ User's Desktop Excel ] 
         ├── Opens cleanly with cached numbers (instant)
         └── User clicks "Smart View -> Refresh" -> Excel pulls fresh data
```

### Components
1. **Data Ingestion (`SvHttpSource` / `probe_sv_http` logic):**
   - Retrieves FinMax actuals for Wholesale (`un.SY.2WHS`) and Technology (`un.SY.2TEC`) EBITDA for August 2026 and September 2026.
   - Extracts numeric values and full 12-dimension POV coordinates.
2. **Headless Generator (`scripts/spike_hsgetvalue.py`):**
   - Creates the workbook headlessly using `openpyxl`.
   - Constructs `=HsGetValue(...)` formulas mapping dimensions to members.
   - Assigns evaluated numeric floats into cell values and configures cached value XML.
   - Injects Microsoft Information Protection (MIP) sensitivity labels matching Telenor's *"Internal"* label.
3. **Verification Harness (`engine/tests/live/test_spike_hsgetvalue.py`):**
   - Validates headless speed (<1.0s), XML structure, and formula presence.
   - Automates live Excel opening and Smart View refresh verification.

---

## 3. Formula Syntax & Coordinates

### Smart View `HsGetValue` Syntax
Oracle Smart View expects comma-separated dimension-member pairs:
```excel
=HsGetValue("FinMax", "Accounts#ac.EBITDA", "Entities#un.SY.2WHS", "Scenarios#Actual", "Years#FY26", "Period#Aug", "Currencies#NOK", "Functions#fu.NONE", "Products#pd.NONE", "Projects#pj.NONE", "Reserves#re.NONE", "Systems#sy.NONE", "GLSources#gl.NONE", "Counterparts#cp.NONE", "LegalEntities#le.NONE")
```

### Dimension Coverage
All 12 FinMax POV dimensions plus Years and Period are explicitly pinned to eliminate ambiguous POV drift:
- **Accounts:** `ac.EBITDA`
- **Entities:** `un.SY.2WHS`, `un.SY.2TEC`
- **Period / Years:** `Aug`, `Sep`, `FY26`
- **Scenarios / Currencies:** `Actual`, `NOK`
- **Default dimensions:** `fu.NONE`, `pd.NONE`, `pj.NONE`, `re.NONE`, `sy.NONE`, `gl.NONE`, `cp.NONE`, `le.NONE`

### Openpyxl Formula + Cached Value Strategy
In standard OpenXML spreadsheets, formula cells store both the expression `<f>` and the calculated result `<v>`:
```xml
<c r="C5" t="n">
    <f>HsGetValue("FinMax", ...)</f>
    <v>1994151357.55</v>
</c>
```
To avoid `#NAME?` or empty cells before Smart View connects, the generator populates both nodes:
- Cell formula assigned to string `=HsGetValue(...)`.
- Cell value `<v>` assigned to the retrieved float value.
- Summary rows (e.g. Total EBITDA) use standard Excel formulas (`=SUM(C5:C6)`).

---

## 4. Packaging, Styling & Sensitivity Labels

### Presentation & Styling
- **Header:** Title ("FinMax Actuals - HsGetValue Spike"), generation timestamp, and scope metadata.
- **Formatting:** Number format `#,##0` or `#,##0.0` (MNOK display).
- **Font & Layout:** Clean Telenor corporate font styling (Segoe UI / Aptos), bold section headers, and thin borders.

### Sensitivity Labels (MIP / Microsoft Purview)
- Telenor's tenant requires an *"Internal"* sensitivity label to avoid blocking file save operations.
- The spike stamps custom document properties in `docProps/custom.xml`:
  - `LabelId`: `f604d2c9-1577-460e-b668-57374a0216c3`
  - `SiteId`: `1676489c-5c72-46b7-ba63-9ab90c4aad44`
  - `AssignmentMethod`: `1`

### Destination Path
- Output file path: `%LOCALAPPDATA%\smartass\analyser\spike_hsgetvalue.xlsx` (kept local to avoid OneDrive sync contention).

---

## 5. Verification Strategy & Success Gates

### Gate 1: Headless Performance (< 1.0s)
Workbook generation runs entirely in Python memory and writes to disk in under 1 second without spawning any Excel process.

### Gate 2: Clean File Open & Static Readability
When opened in desktop Excel:
- No file repair or corruption prompts appear.
- Data cells show their formatted numbers immediately (no blank cells or initial `#NAME?` errors).
- Formula bar reveals the underlying `=HsGetValue(...)` expression.

### Gate 3: Native Smart View Refresh
When the user connects to Smart View and clicks **Smart View → Refresh**:
- Smart View evaluates the formulas against FinMax.
- Cell values refresh accurately and match the reference figures.
- Excel `=SUM(...)` formulas recalculate smoothly.

---

## 6. Risks & Fallbacks

| Risk | Mitigation |
|---|---|
| Smart View connection name `"FinMax"` not found | Test explicit URL connection or fallback to `"none"` with active sheet binding. |
| Openpyxl overwrites cached `<v>` on formula assignment | Use custom OpenXML cell injection if openpyxl's high-level API suppresses the value tag. |
| Tenant blocks unsigned `.xlsx` files | Validate custom MIP property injection or test template cloning (`.xltx`). |
