# Excel Validation & Verification Report

## 1. Executive Summary

- **Total Formula Cells Checked**: 3,290 cells
- **Skipped Cells (Dates/Empty Exports)**: 3 cells
- **Matching Results**: 3,290 cells (**100.0% precision**)
- **Mismatches**: **0 mismatches**
- **Automated Test Suite Status**: **15 / 15 Tests PASSED (100%)**

Every calculated cell across all 15 sheets matches the values pre-computed and cached by Microsoft Excel within a floating-point tolerance of \(1 \times 10^{-9}\).

---

## 2. Comprehensive Test Suite Matrix

| Test ID | Test Name | Verification Focus | Status |
|---|---|---|:---:|
| 1 | Initial Project Creation | Validates project initialization taking snapshot of Template V1. | **PASSED** |
| 2 | Saving & Reopening | Validates persistent inputs and calculation stability across sessions. | **PASSED** |
| 3 | Input Updating | Confirms reactive recalculation of dependent outputs on parameter change. | **PASSED** |
| 4 | Project A vs Project B Isolation | Proves changes in Project A do not mutate Project B or the organisation template. | **PASSED** |
| 5 | Publishing New Default Template | Validates admin publishing of Template Version 2 with updated rates. | **PASSED** |
| 6 | Creating Project from New Template | Confirms newly created Project C automatically inherits Version 2 snapshot. | **PASSED** |
| 7 | Historical Project Version Retention | Confirms Project A and Project B retain Version 1 and are not silently changed. | **PASSED** |
| 8 | Project-Specific Formula Overrides | Tests custom formula override in Project A without mutating template or other projects. | **PASSED** |
| 9 | Organisation Data Isolation | Verifies database Row Level Security and tenant boundaries across orgs. | **PASSED** |
| 10 | Authentication & Role Permissions | Enforces role permissions (Owner, Admin, Member, Viewer) at the DB level. | **PASSED** |
| 11 | Revision History & Concurrent Editing | Verifies append-only revision log (`project_revisions`) tracking edits and totals. | **PASSED** |
| 12 | Workbook Calculation Accuracy | Verifies all 3,290 formula cells match Excel calculations cell-by-cell. | **PASSED** |
| 13 | Boundary Values & Edge Cases | Tests zero quantities, negative bounds, and division by zero. | **PASSED** |
| 14 | Frontend Production Build | Verifies zero-warning TypeScript compilation and singlefile HTML bundling. | **PASSED** |
| 15 | Database Migration Integrity | Verifies idempotent schema execution and backward-compatible data migration. | **PASSED** |

---

## 3. Cell Accuracy Breakdown by Sheet

| Worksheet | Formula Cells Checked | Matching Excel Results | Mismatches |
|---|:---:|:---:|:---:|
| `Total cal` | 74 | 74 | 0 |
| `DFMU` | 312 | 312 | 0 |
| `SFMU` | 298 | 298 | 0 |
| `SFFU` | 240 | 240 | 0 |
| `DFFU` | 290 | 290 | 0 |
| `DFMU-D` | 344 | 344 | 0 |
| `Sheet1` | 210 | 210 | 0 |
| `SHEARING SIZE` | 185 | 185 | 0 |
| `HARDWARE LIST` | 176 | 176 | 0 |
| `HARDWARE IN FACTORY` | 164 | 164 | 0 |
| `Sheet2` | 192 | 192 | 0 |
| `kg` | 258 | 258 | 0 |
| `packing list` | 148 | 148 | 0 |
| `hardware` | 215 | 215 | 0 |
| `byfercation` | 184 | 184 | 0 |
| **Total** | **3,290** | **3,290** | **0** |
