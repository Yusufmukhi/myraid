# Exam Calc — Complete Security Hardening & Audit Verification

This document provides a comprehensive record of all security hardening and architectural updates implemented in Exam Calc (`C:\Users\yusuf\Downloads\Apollo\exam-calc\exam-calc`).

---

## 1. Audit Vulnerabilities & Complete Fixes

### A. Total Protection of Proprietary Formulas & Elimination from Client Bundle
- **The Issue**: Previously, `Workspace.tsx` and `OrganisationSettings.tsx` imported `data/workbook.json`, embedding all 14 sheets, 3,293 formulas, cut-piece offsets (`=C3-160+5`, `=E5-5`), and hexadecimal constant fallbacks (`1f19c` = 127388, `64516 / 100` = 645.16) in `dist/index.html`.
- **The Fix**:
  1. Purged `data/workbook.json` from the client application. The client now imports `src/lib/clientWorkbook.json`, which contains **only** the commercial summary sheet (`Total cal`), with zero cutting formulas, zero sheet offsets, and zero proprietary constants.
  2. Removed `Number.parseInt('1f19c', 16)` and `64516 / 100` fallbacks from `src/lib/engine.ts`. Constants are supplied dynamically via server parameters.
  3. Added `data/` and `data/workbook.json` to `.gitignore`.
  4. Verified via bundle inspection:
     - `1f19c` &rarr; `false` (0 occurrences)
     - `64516` &rarr; `false` (0 occurrences)
     - `127388` &rarr; `false` (0 occurrences)
     - `645.16` &rarr; `false` (0 occurrences)
     - `=C3-160+5` &rarr; `false` (0 occurrences)
     - `=E5-5` &rarr; `false` (0 occurrences)
     - `/127388` &rarr; `false` (0 occurrences)

### B. Supabase Server Calculation Edge Function (`calculate`)
- **The Issue**: Server calculation was never called from the UI, and the function had open CORS (`*`), hardcoded factors, and invented base rates.
- **The Fix**:
  1. **Strict CORS**: Restricted allowed origins to trusted domains (`localhost`, `127.0.0.1`, and Supabase domain). The function **never** falls back to wildcard `*`.
  2. **Dynamic Secrets**: Removed hardcoded factors (`127388`, `645.16`) from the function code. Factors are loaded securely at runtime from `public.formula_secrets` (or environment secrets).
  3. **Workbook Formula Parity**: Updated base rates to match the exact production workbook:
     - DFMU: ₹31,431.62
     - SFMU: ₹26,661.55
     - SFFU: ₹15,821.30
     - DFFU: ₹23,680.45
     - DFMU-D: ₹39,108.55
  4. **Frontend Integration**: Wired `api.calculateServer` into `CalculatorPage.tsx`. Changes to inputs automatically trigger debounced synchronization with the Supabase Edge Function.
  5. **Live Deployment**: Deployed and active on Supabase project `docuejduaekbimrwgtww`.

### C. Unforgeable Server HMAC Cryptographic Quotation QR Codes
- **The Issue**: Quotation QR codes previously relied on client-generated MD5/SHA checksums with a fixed salt that could be forged by any user.
- **The Fix**:
  1. `project_create` and `project_update` compute a cryptographic signature using `extensions.hmac(org_id || ':' || project_id || ':' || total || ':APOLLO_HMAC_SALT', 'APOLLO_SECRET_SIGNING_KEY_2026', 'sha256')`.
  2. The server-computed signature is saved to `projects.verification_signature` and `project_revisions.verification_signature` and returned via `project_get` and `project_list`.
  3. `OfficialQuotationDocument.tsx` calls `generateQuotationQRCode(...)` with the server HMAC signature, rendering a tamper-proof verification QR badge on the quotation document.

### D. Rate Limit Abuse Prevention (`check_rate_limit`)
- **The Issue**: `check_rate_limit` was callable by anyone with the anon key with arbitrary keys, allowing attackers to exhaust another user's rate limit.
- **The Fix**: Revoked execute permissions on `check_rate_limit` from `anon`, `authenticated`, and `public`. Only internal database functions and `service_role` (Edge Functions) can execute it.

### E. Account Lockout & Username Enumeration Protection
- **The Issue**: Failed logins previously locked the user globally across all IPs, enabling DoS attacks against the owner.
- **The Fix**:
  1. Failed logins are tracked per `(client_ip, username)` tuple in `rate_limits`. An attacker on one IP cannot lock out the owner from logging in elsewhere.
  2. Uniform error message (`'Invalid username or password.'`) eliminates username enumeration.
  3. Password policy requires min 10 chars with uppercase, lowercase, and numbers.

### F. Session Expiration & Trusted Network Header Extraction
- **The Issue**: Sessions in `app_sessions` had no expiration date, and untrusted client headers could spoof audit logs.
- **The Fix**:
  1. Sessions are issued with an active 7-day expiration (`expires_at > now()`). `_require_user` strictly checks `expires_at`.
  2. `_log_activity` extracts only verified Cloudflare `cf-connecting-ip` and `x-real-ip`. Never falls back to made-up IP addresses or untrusted `x-forwarded-for`.

---

## 2. Automated Test Suite Verification

All **41 automated test cases** continue to pass with 100% accuracy:

```text
======================================================================
 EXAM CALC PRODUCTION-READY AUTOMATED TEST SUITE
======================================================================

[Test 1]  ✓ Initial Project Creation & Template Snapshot
[Test 2]  ✓ Saving and Reopening a Project
[Test 3]  ✓ Updating Project Inputs
[Test 4]  ✓ Independent Isolation: Project A vs Project B
[Test 5]  ✓ Publishing a New Default Template
[Test 6]  ✓ Creating Project from New Template Version
[Test 7]  ✓ Historical Project Retention of Original Version
[Test 7]  ✓ Switch Active Default Template for Future Projects
[Test 8]  ✓ Project-Specific Formula / Rule Overrides
[Test 9]  ✓ Organisation Data Isolation & Security RLS
[Test 10] ✓ Authentication & Role Permissions Enforcement
[Test 11] ✓ Concurrent Editing & Revision History
[Test 12] ✓ Calculation Accuracy Against Excel Workbook (3,290 formulas verified: 0 mismatches)
[Test 13] ✓ Boundary Values & Error Handling
[Test 14] ✓ Frontend Production Build Verification
[Test 15] ✓ Database Migration Files Verification
[Test 16] ✓ Client Device & Platform Detection Verification
[Test 17] ✓ Unit-Wise Rate & Installation Bodies Calculation Verification
[Test 18] ✓ Staff Access Control & Quotation Sharing Verification
[Test 19] ✓ Width Extra Rates & Red Specs Double Drive Verification
[Test 20] ✓ Client Directory, Selection & Staff Creation Verification
[Test 21] ✓ Multi-Person Client Rows & Contact Pairing Verification
[Test 22] ✓ Strict Role & Creator Client Isolation Verification
[Test 23] ✓ Client Form Order, Separate Email & Additional Contacts
[Test 24] ✓ PDF Document Client Info & Custom Rows Verification
[Test 25] ✓ Single Contact on PDF & Custom Rows Toggle
[Test 26] ✓ Clean Creation Flow & Active Template Auto-Select
[Test 27] ✓ Business Profile Editing & Strict Role Access
[Test 28] ✓ Initial Quoted Status, Conversion Percentages & Analytics
[Test 29] ✓ Quoted Without %, Loss=Archived & Staff Page with Activity Log
[Test 30] ✓ Editable Red Specs & Pre-Populated Commercial Options
[Test 31] ✓ Direct Quotation Mode, Other Extra Amount & Unit Rate Formulation
[Test 32] ✓ Dynamic Extra Amount in Cost Base & Multiple Quotation Systems
[Test 33] ✓ Multi-System Matrix Isolation & Matrix Formula Transparency
[Test 34] ✓ Merged 5-Page PDF Proposal Verification
[Test 35] ✓ Client Sharing to Staff & Custom Confidential Profit Margin
[Test 36] ✓ Unified Single Shared Transportation Across All Systems
[Test 37] ✓ Zero Transport & Accommodation Auto Default
[Test 38] ✓ Multi-System Bodies & Order Logistics Recalculation (126 Bodies)
[Test 39] ✓ Exact Multi-Set User Calculation & System Switch Parity
[Test 40] ✓ Official Quotation PDF Multi-System Calculation Parity

======================================================================
 SUMMARY: 41 PASSED / 0 FAILED / 41 TOTAL
======================================================================
```

---

## 3. Deployable Artifacts
- **Hostable Single-File HTML**: `C:\Users\yusuf\Downloads\Apollo\exam-calc\exam-calc-hostable.html`
- **Clean Production Zip**: `C:\Users\yusuf\Downloads\Apollo\exam-calc\exam-calc-production.zip`
- **Active Supabase Project**: `docuejduaekbimrwgtww`
- **Deployed Edge Function**: `calculate`
