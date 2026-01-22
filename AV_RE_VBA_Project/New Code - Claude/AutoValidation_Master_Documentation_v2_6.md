# Auto-Validation System - Master Documentation

**Project:** VBA Excel Auto-Validation Framework  
**Version:** 2.6 (Header-Based Architecture Complete)  
**Last Updated:** 2026-01-21  
**Status:** Phase 2 Complete - Header-Based Architecture Fully Implemented

---

## 📋 Change Log

### [2.6.0] - 2026-01-21 - Header-Based Architecture Complete
**Status:** Complete

**Major Milestone: Complete Transition from Column Letters to Header Names**

All modules have been updated to use header-based cell lookups instead of column letters. This is a significant architectural improvement that makes the system resilient to column insertions/deletions and improves maintainability.

**Key Changes by Module:**

#### AV_Core v2.6
- ✅ `GetAutoValidationMap()` now reads `"ReviewSheet Column Header"` instead of `"ReviewSheet Column Letter"`
- ✅ `GetDDMValidationColumns()` now reads `"ReviewSheet Column Name"` for header-based lookup
- ✅ `GetValidMenuValues()` uses header-based table lookup instead of column letters
- ✅ **NEW:** `GetCellByHeader(TargetTable, rowNum, headerName)` - Central helper for header-based cell access
- ✅ All dictionary assignments use proper `Set` keyword (Error #450 fix retained)

#### AV_Engine v2.6
- ✅ `ENGINE_VERSION = "2.6"` constant for version tracking
- ✅ **NEW:** `CurrentTargetTable As ListObject` - Public variable for cross-module table access
- ✅ `ValidateSingleRow()` uses `GetCellByHeader()` for all cell lookups
- ✅ `RunAutoCheckDataValidation()` converted to header-based cell access
- ✅ `PrintHeaderDiagnostics()` outputs mapping verification to Immediate Window
- ✅ All `Range(colLetter & rowNum)` patterns replaced with header lookups

#### AV_Validators v2.6
- ✅ `GetSiblingCell()` now uses `"ReviewSheet Column Header"` instead of `"ReviewSheet Column Letter"`
- ✅ Uses `AV_Engine.CurrentTargetTable` for table reference
- ✅ Falls back to `AV_Core.GetCellByHeader()` for all lookups
- ✅ Passes `AutoValMap` to all validation rule calls for cached header access

#### AV_Format v2.6
- ✅ `WriteSystemTagToDropColumn()` accepts `ListObject` and header names
- ✅ `AddValidationFeedback()` uses `AV_Engine.CurrentTargetTable` for lookups
- ✅ All drop column operations use header-based access
- ✅ Source cell formatting uses header-based lookup

#### AV_ValidationRules v2.5
- ✅ All validators pass `AutoValMap` to `GetSiblingCell()` for header resolution
- ✅ `ValidatePairedFields()` uses header-based sibling cell lookup
- ✅ `Validate_HeatPairs()` uses header-based cell access
- ✅ All `AddValidationFeedback()` calls work with header-based architecture

**Breaking Changes:** None - maintains backward compatibility

**Configuration Table Changes:**
- `AutoValidationCommentPrefixMappingTable` now requires `"ReviewSheet Column Header"` column (header name, not letter)
- `AutoCheckDataValidationTable` uses `"ReviewSheet Column Name"` for target column identification

---

### [2.5.0] - 2026-01-21 - Critical Bug Fixes & Performance Optimization
**Status:** Complete

**Critical Fixes:**
- ✅ **Resolved Error #450** - Fixed Dictionary object assignment syntax (`Set dict(key) = item`)
- ✅ **14.6x Performance Improvement** - Validation time reduced from 467 seconds to 32 seconds

**Optimization Techniques Applied:**
- Column index caching in ValidateSingleRow
- Single Application.EnableEvents toggle per target (not per row)
- DoEvents every 10 rows instead of every row
- Cached validation tables

---

### [2.4.0] - 2026-01-20 - Table-Based Engine Overhaul
**Status:** Complete

- Auto-detects target table range from ListObject.DataBodyRange
- Uses ValidationTargets table for multi-table support
- Added diagnostic output comparing mapped headers vs actual table headers

---

### [2.3.0] - 2026-01-19 - Debug System & Logging
**Status:** Complete

- Dual-mode debug system (user progress vs developer debug)
- GlobalDebugOptions table controls system-wide debug output
- ValidationTrackerForm always shows progress for users

---

### [2.0.0 - 2.2.0] - Architecture Consolidation
**Status:** Complete

- Consolidated 15-20 modules into 6 core modules
- Created AV_Constants and AV_DataAccess modules
- Table caching implemented

---

## 🎯 Project Vision & Goals

### Primary Objectives
1. **Maintain Functionality** - All existing validations work correctly
2. **Header-Based Configuration** - No hardcoded column letters anywhere
3. **Performance** - <35 seconds for ~2000 rows
4. **Reusability** - Support multiple validation targets with different structures

### Achieved in v2.6
- ✅ Complete elimination of column letter references in core validation logic
- ✅ Table-driven configuration throughout
- ✅ Cross-module table sharing via `AV_Engine.CurrentTargetTable`
- ✅ Centralized cell lookup via `AV_Core.GetCellByHeader()`

---

## 🏗️ Architecture (v2.6)

### Module Structure

```
┌─────────────────────────────────────────────────────────┐
│                   AV_Core.bas v2.6                      │
│  Configuration + State + Caching + Cell Lookup         │
│  KEY FUNCTIONS:                                         │
│  - GetCellByHeader(table, row, headerName) ← NEW       │
│  - GetAutoValidationMap() - reads "ReviewSheet Column  │
│    Header" not "ReviewSheet Column Letter"             │
│  - GetDDMValidationColumns() - header-based            │
│  - GetValidMenuValues() - header-based table lookup    │
└─────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────┐
│                  AV_Engine.bas v2.6                     │
│  Orchestration & Execution                              │
│  KEY CHANGES:                                           │
│  - CurrentTargetTable As ListObject ← PUBLIC           │
│  - ValidateSingleRow() uses GetCellByHeader()          │
│  - RunAutoCheckDataValidation() header-based           │
│  - PrintHeaderDiagnostics() for mapping verification   │
└─────────────────────────────────────────────────────────┘
                            │
                    ┌───────┴───────┐
                    ▼               ▼
┌─────────────────────────┐  ┌──────────────────────────┐
│  AV_Validators.bas v2.6 │  │   AV_Format.bas v2.6     │
│  Routing Layer          │  │   Formatting & Feedback  │
│  KEY CHANGES:           │  │  KEY CHANGES:            │
│  - GetSiblingCell() now │  │  - WriteSystemTagTo      │
│    uses header names    │  │    DropColumn() accepts  │
│  - Uses CurrentTarget   │  │    ListObject + headers  │
│    Table from AV_Engine │  │  - Uses CurrentTarget    │
└─────────────────────────┘  │    Table for lookups     │
            │                └──────────────────────────┘
            ▼
┌─────────────────────────┐
│ AV_ValidationRules v2.5 │
│ Business Logic          │
│  - All validators pass  │
│    AutoValMap for       │
│    header resolution    │
│  - Sibling cell lookups │
│    use header names     │
└─────────────────────────┘
            │
            ▼
┌─────────────────────────┐
│     AV_UI.bas v2.1      │
│  User Interface         │
│  - ShowValidationTrackerForm │
│  - AppendUserLog        │
│  - CancelValidation     │
└─────────────────────────┘
```

### Key Data Flow (v2.6)

```
1. RunFullValidationMaster()
   │
   ├─> Set CurrentTargetTable = tblTarget  ← TABLE SHARED GLOBALLY
   │
   └─> FOR EACH row:
       │
       └─> ValidateSingleRow()
           │
           ├─> Get header name from AutoValMap("ColumnRef")
           │
           └─> GetCellByHeader(CurrentTargetTable, rowNum, headerName)
               │
               └─> Application.Run("Validate_Column_X")
                   │
                   └─> GetSiblingCell(cell, sheet, "OtherFunc", AutoValMap)
                       │
                       └─> Uses CurrentTargetTable + GetCellByHeader()
```

---

## 📊 Configuration Tables Reference

### Critical Configuration Changes for v2.6

#### AutoValidationCommentPrefixMappingTable
**IMPORTANT:** This table now uses `"ReviewSheet Column Header"` instead of `"ReviewSheet Column Letter"`

| Column | Type | Description | Example |
|--------|------|-------------|---------|
| Dev Function Names | String | Function identifier | "Electricity" |
| Drop in Column | String | **Header name** for messages column | "National Review Comments" |
| Prefix to message | String | English message prefix | "Electricity" |
| (FR) Prefix to message | String | French message prefix | "Électricité" |
| RuleTableName | String | Associated rule table | "ElectricityPairValidation" |
| AutoValidate | Boolean | TRUE to enable | TRUE |
| **ReviewSheet Column Header** | String | **Header name** of column to validate | "Electricity" |

**Current Mappings (9 validators):**

| Dev Function Names | ReviewSheet Column Header | RuleTableName |
|-------------------|---------------------------|---------------|
| GIWQuantity | Gender Inclusive Washrooms - Quantity | GIWValidationTable |
| GIWIncluded | Gender Inclusive Washrooms - Included | GIWValidationTable |
| Electricity | Electricity | ElectricityPairValidation |
| Electricity_Metered | Electricity Metered | ElectricityPairValidation |
| Plumbing | Plumbing | PlumbingPairValidation |
| Water_Metered | Water Metered | PlumbingPairValidation |
| Heat_Source | Heat Source | HeatSourcePairValidation |
| Heat_Metered |  Heat Metered | HeatSourcePairValidation |
| Construction_Date | Construction Date | (none) |

**Note:** The " Heat Metered" header has a leading space in the actual table - this must match exactly.

---

#### AutoCheckDataValidationTable (Simple Validations)
**Uses `"ReviewSheet Column Name"` for header-based lookup**

| Column | Type | Description |
|--------|------|-------------|
| Column Name | String | English display name |
| Column Name (FR) | String | French display name |
| **ReviewSheet Column Name** | String | **Header name** in target table |
| MenuField Column (EN) | String | Valid values column header (EN) |
| MenuField Column (FR) | String | Valid values column header (FR) |
| AutoCheck | Boolean | TRUE to enable |
| AutoComment Column | String | Drop column header for messages |

---

### ValidationTargets Table

| Column | Type | Description |
|--------|------|-------------|
| TableName | String | ListObject name |
| Enabled | Boolean | TRUE to validate |
| Mode | String | "Both", "Trigger", or "Bulk" |
| Key Column (Header Name) | String | Key column header (e.g., "AO ID") |

**Current Configuration:**
```
| TableName              | Enabled | Mode | Key Column (Header Name) |
|------------------------|---------|------|--------------------------|
| REP2DSMDraft_Buildings | TRUE    | Both | AO ID                    |
| REP2DSMDraft_Works     | TRUE    | Both | AO ID                    |
| REP2DSMDraft_Other     | TRUE    | Both | AO ID                    |
```

---

### Validation Rule Tables

All rule tables remain unchanged from v2.5:

1. **GIWValidationTable** - GIW Included/Quantity rules
2. **ElectricityPairValidation** - Electricity/Metered valid pairs
3. **PlumbingPairValidation** - Plumbing/Water Metered pairs
4. **HeatSourcePairValidation** - Heat Source/Metered pairs
5. **HeatSourceANYRefTable** - Heat ANY mapping

---

## 🔄 Cell Lookup Architecture (v2.6)

### The GetCellByHeader Pattern

All cell lookups in v2.6 use this central function:

```vba
' AV_Core.GetCellByHeader
Public Function GetCellByHeader(TargetTable As ListObject, rowNum As Long, headerName As String) As Range
    ' Returns cell at intersection of header column and worksheet row
    
    If TargetTable Is Nothing Then Exit Function
    If Len(headerName) = 0 Then Exit Function
    
    ' Get column index by header name
    Dim colIndex As Long
    On Error Resume Next
    colIndex = TargetTable.ListColumns(headerName).Index
    On Error GoTo 0
    
    If colIndex = 0 Then Exit Function
    
    ' Convert worksheet row to table row
    Dim tableRow As Long
    tableRow = rowNum - TargetTable.DataBodyRange.Row + 1
    
    If tableRow < 1 Or tableRow > TargetTable.ListRows.Count Then Exit Function
    
    Set GetCellByHeader = TargetTable.DataBodyRange(tableRow, colIndex)
End Function
```

### Cross-Module Table Sharing

```vba
' AV_Engine sets the table reference
Public CurrentTargetTable As ListObject

' In ProcessValidationTarget:
Set CurrentTargetTable = tblTarget

' Other modules use it:
' AV_Validators.GetSiblingCell:
Set GetSiblingCell = AV_Core.GetCellByHeader(AV_Engine.CurrentTargetTable, cell.Row, headerName)

' AV_Format.AddValidationFeedback:
WriteSystemTagToDropColumn AV_Engine.CurrentTargetTable, dropColHeader, targetRow, ...
```

---

## 📈 Performance Metrics

### Validation Execution Times

| Version | Time | Rows | Key Improvement |
|---------|------|------|-----------------|
| v1.0 (Original) | ~467s | ~2000 | Baseline |
| v2.5 (Error #450 fix) | ~32s | ~2000 | 14.6x faster |
| v2.6 (Header-based) | ~32s | ~2000 | Maintains performance |

### Why Header-Based is Better

| Aspect | Column Letters (Old) | Header Names (v2.6) |
|--------|---------------------|---------------------|
| Column insertion | **Breaks** all references | No impact |
| Column deletion | **Breaks** all references | No impact |
| Debugging | Hard to trace "M" | Clear "Electricity" |
| Configuration | Fragile | Robust |
| Maintenance | High effort | Low effort |

---

## 💾 Complete Module Inventory (v2.6)

| Module | Version | Lines | Status | Key Changes in v2.6 |
|--------|---------|-------|--------|---------------------|
| AV_Core | 2.6 | ~630 | ✅ Complete | GetCellByHeader(), header-based map loading |
| AV_Engine | 2.6 | ~500 | ✅ Complete | CurrentTargetTable, header-based ValidateSingleRow |
| AV_Format | 2.6 | ~560 | ✅ Complete | Header-based WriteSystemTagToDropColumn |
| AV_Validators | 2.6 | ~150 | ✅ Complete | Header-based GetSiblingCell |
| AV_ValidationRules | 2.5 | ~730 | ✅ Complete | Works with header-based architecture |
| AV_UI | 2.1 | ~140 | ✅ Stable | No changes needed |
| AV_DataAccess | 2.2 | ~350 | ✅ Stable | Supplementary functions |
| ValidationTrackerForm | 2.1 | ~150 | ✅ Stable | No changes needed |

**Total:** ~3,200+ lines across 8 components

---

## 🧪 Testing Checklist (v2.6)

### Version Verification
```vba
Sub CheckVersions()
    Debug.Print "AV_Engine version: " & AV_Engine.ENGINE_VERSION
    ' Expected: 2.6
End Sub
```

### Header Mapping Verification
When validation runs, the Immediate Window shows:
```
-----------------------------------------------
DIAGNOSTIC: Header Mapping Check
-----------------------------------------------
OK: Validate_Column_Electricity -> 'Electricity' found at index 22
OK: Validate_Column_Construction_Date -> 'Construction Date' found at index 7
MISSING: Validate_Column_Heat_Metered -> ' Heat Metered' NOT in table
```

### Functional Tests

- [ ] Run full validation on REP2DSMDraft_Buildings
- [ ] Verify GIW Quantity/Included pair validation works
- [ ] Verify auto-corrections apply correctly
- [ ] Verify SYS_TAG messages appear in drop column
- [ ] Verify formatting (Error, Autocorrect, Default) applies
- [ ] Verify key column priority formatting
- [ ] Test with French language option

### Regression Tests

- [ ] No Error #450 on GetAutoValidationMap
- [ ] No column letter references in validation logic
- [ ] ~32 seconds for ~2000 rows
- [ ] Cancel button works

---

## 🐛 Known Issues & Technical Debt

### Resolved in v2.6

| Issue | Status | Resolution |
|-------|--------|------------|
| Column letter brittleness | ✅ RESOLVED | Header-based lookups |
| GetSiblingCell using column letters | ✅ RESOLVED | Uses AutoValMap + GetCellByHeader |
| Error #450 Dictionary assignment | ✅ RESOLVED | Added `Set` keyword |
| Slow validation (~467s) | ✅ RESOLVED | Caching + event optimization |

### Remaining Low-Priority Items

| Issue | Priority | Status |
|-------|----------|--------|
| Heat Metered leading space in header | Low | Document only - must match config |
| EN/FR language switching | Low | Functional but not fully tested |
| ForceValidationTable full implementation | Low | Basic filtering works |

---

## 🗺️ Development Roadmap

### Phase 2: Complete ✅
- ✅ Table-based configuration
- ✅ Centralized constants
- ✅ Data access layer
- ✅ Performance optimization (14.6x)
- ✅ **Header-based architecture (v2.6)**

### Phase 3: Planned 📋
- Production testing with real data
- EN/FR language verification
- ForceValidationTable enhancement
- Additional validation rules as needed

### Phase 4: Future 📅
- Generic validation engine
- Multi-workbook support
- Validation function registry

---

## 📚 Quick Reference

### Key Code Locations (v2.6)

| Function | Location | Description |
|----------|----------|-------------|
| Main Entry | AV_Engine.RunFullValidationMaster() | Starts validation |
| Cell Lookup | AV_Core.GetCellByHeader() | Header-based cell access |
| Current Table | AV_Engine.CurrentTargetTable | Shared table reference |
| Sibling Cell | AV_Validators.GetSiblingCell() | Finds pair cell by header |
| Feedback | AV_Format.AddValidationFeedback() | Writes messages |
| Tag Writing | AV_Format.WriteSystemTagToDropColumn() | SYS_TAG management |

### Configuration Table Dependencies

```
ValidationTargets
    └─> Target Table (ListObject)
        └─> AutoValidationCommentPrefixMappingTable
            ├─> ReviewSheet Column Header → Target column
            ├─> Drop in Column → Message column  
            └─> RuleTableName → Validation rules
                └─> ElectricityPairValidation
                └─> GIWValidationTable
                └─> etc.
```

---

## Appendix A: Migration Notes

### Migrating from v2.5 to v2.6

1. **Update AutoValidationCommentPrefixMappingTable**
   - Ensure `"ReviewSheet Column Header"` column exists with header names (not letters)

2. **Update AutoCheckDataValidationTable**
   - Ensure `"ReviewSheet Column Name"` contains header names

3. **Import Updated Modules**
   - AV_Core v2.6
   - AV_Engine v2.6
   - AV_Validators v2.6
   - AV_Format v2.6

4. **Verify Headers Match**
   - Run validation and check Immediate Window for MISSING headers
   - Fix any mismatches (case-sensitive, including leading/trailing spaces)

---

## Appendix B: Error Reference

| Error | Cause | Solution |
|-------|-------|----------|
| #450 | Missing `Set` for object assignment | Use `Set dict(key) = item` |
| #91 | Object variable not set | Check object Is Not Nothing |
| Column not found | Header mismatch | Verify exact header name in table |
| GetSiblingCell returns Nothing | Header not in AutoValMap | Check "ReviewSheet Column Header" |

---

**END OF MASTER DOCUMENTATION v2.6**

*This is a living document. Update the changelog at the top when making significant changes.*
