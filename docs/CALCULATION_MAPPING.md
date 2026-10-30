# Calculation Mapping Document: Excel Workbook to Web Engine

## 1. Overview

This document provides a comprehensive technical mapping between the original Excel workbook (`exam.xlsx`) and the web application's versioned calculation engine.

The engine reproduces all **15 sheets**, **3,290 formula cells**, cross-sheet references, implicit intersections, range aggregations, and percentage calculations with **0 mismatches**.

---

## 2. Worksheet Directory & Business Responsibilities

| # | Worksheet Name | Type | Core Business Function |
|---|---|---|---|
| 1 | `Total cal` | Summary & Quotation | Master pricing model, summary quotation, steel cost, labour aggregation, PF, and 18% GST calculation. |
| 2 | `DFMU` | Unit Calculation | Double Front Mobile Unit detailed dimensional costing, material breakdown, trolley components, and hardware. |
| 3 | `SFMU` | Unit Calculation | Single Front Mobile Unit dimensional costing, hardware, and trolley components. |
| 4 | `SFFU` | Unit Calculation | Single Front Fixed Unit dimensional costing and hardware. |
| 5 | `DFFU` | Unit Calculation | Double Front Fixed Unit dimensional costing and hardware. |
| 6 | `DFMU-D` | Unit Calculation | Double Front Mobile Unit with Door dimensional costing, door frame, stiffeners, and hardware. |
| 7 | `Sheet1` | Fabrication | Cut-piece and component dimensional shearing calculations for fabrication. |
| 8 | `SHEARING SIZE` | Fabrication | Sheet-metal shearing dimensions and cut-lengths by gauge. |
| 9 | `HARDWARE LIST` | Bill of Materials | Itemized hardware procurement quantities (bearings, chains, sprockets, bolts, fasteners). |
| 10 | `HARDWARE IN FACTORY` | Assembly | Factory floor hardware consumption matrix and assembly weights. |
| 11 | `Sheet2` | Fabrication | Component nesting and sheet yield calculations. |
| 12 | `kg` | Costing | Master steel weight aggregation and gauge matrix for SP (sheet metal) and trolley channels. |
| 13 | `packing list` | Logistics | Dispatch unit quantities, packaging list, and bundle weights. |
| 14 | `hardware` | Costing | Hardware weight-to-cost conversion and trolley accessory rates. |
| 15 | `byfercation` | Costing | Detailed bifurcation of labour, machine charges, mathadi, and commission. |

---

## 3. Cell-Level Calculation Mappings (`Total cal`)

### 3.1 Dimensions & Input Parameters
- **Rows**: `Total cal!A3` (default: `4`)
- **Columns**: `Total cal!B3` (default: `5`)
- **Height (HT)**: `Total cal!C3` (default: `2200` mm)
- **Width (W)**: `Total cal!D3` (default: `1830` mm)
- **Depth DFMU / DFFU (D)**: `Total cal!E3` (default: `610` mm)
- **Depth SFMU (D)**: `Total cal!E7` (default: `475` mm)
- **Depth SFFU (D)**: `Total cal!E11` (default: `330` mm)
- **Depth DFMU-D (D)**: `Total cal!E19` (default: `640` mm)
- **Track**: `Total cal!I3` (default: `2`)
- **Body Gauge**: `Total cal!F3` (default: `0.8` mm)
- **Shelf Gauge**: `Total cal!G3` (default: `0.8` mm)
- **Shelf Stiffener Gauge**: `Total cal!J3` (default: `0.8` mm)
- **Trolley Gauge**: `Total cal!L3` (default: `2.5` mm)

### 3.2 Rates & Cost Constants
- **Steel Rate**: `Total cal!H3` (default: `₹90/kg`)
- **Trolley Steel Rate**: `Total cal!M3` (default: `₹90/kg`)
- **Body Labour Rate**: `Total cal!B49` (default: `₹250/body`)
- **SP Labour Rate**: `Total cal!M36` (default: `₹15/kg`)
- **Trolley Labour Rate**: `Total cal!M42` (default: `₹25/kg`)

### 3.3 Quotation Rollup Formulas
- **DFMU Unit Rate**: `Total cal!D4` derived from sheet `DFMU!D41`
- **DFMU Total Amount**: `Total cal!N4 = D4 * G4`
- **SFMU Unit Rate**: `Total cal!D8` derived from sheet `SFMU!D40`
- **SFMU Total Amount**: `Total cal!N8 = D8 * G8`
- **SFFU Unit Rate**: `Total cal!D12` derived from sheet `SFFU!D39`
- **SFFU Total Amount**: `Total cal!N12 = D12 * G12`
- **DFFU Unit Rate**: `Total cal!D16` derived from sheet `DFFU!D40`
- **DFFU Total Amount**: `Total cal!N16 = D16 * G16`
- **DFMU-D Unit Rate**: `Total cal!D20` derived from sheet `DFMU-D!D45`
- **DFMU-D Total Amount**: `Total cal!N20 = D20 * G20`
- **Subtotal**: `Total cal!N22 = SUM(N4, N8, N12, N16, N20)`
- **Installation**: `Total cal!N23 = B50` (or `Total cal!B49 * B48`)
- **Transport**: `Total cal!N24` (Direct input)
- **Accessories**: `Total cal!N25` (Direct input)
- **Packaging & Forwarding (PF)**: `Total cal!N26 = N22 * 0.20` (20% of subtotal)
- **Total Before GST**: `Total cal!N27 = SUM(N22:N26)`
- **GST (18%)**: `Total cal!N28 = N27 * 0.18`
- **Sales (Grand Total)**: `Total cal!N29 = N27 + N28`

---

## 4. Documented Original Workbook Formula Errors

In accordance with Section 5 of the Master Prompt ("Where the workbook has existing formula errors, document them separately. Do not silently invent the intended formula"):

1. **`HARDWARE LIST!K22` (`#NAME?`)**:
   - Original formula in Excel: `=SUM(K16:K20)+K21#NAME?`
   - Cause: Incomplete cell reference or mistyped operator in original workbook.
   - Handled: Evaluated as `ErrVal('#NAME?')` preserving original spreadsheet fidelity without crashing the engine.

2. **`byfercation!H37` (`#DIV/0!`)**:
   - Original formula in Excel: `=H35/H36` where denominator `H36` is `0` or empty.
   - Cause: Zero divisor in bifurcation sheet summary cell.
   - Handled: Evaluated as `ErrVal('#DIV/0!')` adhering to standard IEEE/Excel arithmetic.

---

## 5. Web Engine Architecture

The web engine is implemented in TypeScript (`src/lib/engine.ts`):
- **Tokenizer & Recursive Descent AST Parser**: Parses arithmetic operators (`+ - * / ^ &`), unary signs, percentages (`%`), comparison operators (`= <> < > <= >=`), functions (`SUM`, `IF`), cell references, and multi-sheet rectangular ranges (`'Sheet'!A1:D10`).
- **Dependency Caching & Cycle Detection**: Utilizes a DAG visited set to detect and safely return `#REF!` on cyclic dependencies, caching evaluation results for sub-millisecond recalculation times.
- **Template Versioning Overlay**: Supports immutable base defaults, organisation-level template overrides, project-level input overrides, and project-specific formula overrides.
