# Exam Calc — Enterprise Production Web Application

Exam Calc transforms the 15-sheet Apollo storage costing workbook (`exam.xlsx`) into an enterprise-grade, multi-tenant project calculator web application with persistent template versioning, immutable project snapshots, role-based access control, and 100% Excel calculation parity (3,290 formulas verified).

---

## Key Features

1. **Enterprise Template Versioning**:
   - Every organisation has its own published default calculation template.
   - Historical template versions are tracked and audited.
   - Publishing a new version (e.g., Version 2) updates the default for future projects without altering historical projects.
   - Authorised administrators can draft, preview diffs, and publish template versions.

2. **Immutable Project Snapshots**:
   - Newly created projects take a persistent snapshot of the currently published default template.
   - Changing inputs, steel rates, or calculation rules in **Project A** does not mutate **Project B** or the organisation's shared template.
   - Projects can optionally be recalculated to the latest published template with confirmation and full rollback history.

3. **Complete 7-Page Architecture**:
   - **Page 1 — Dashboard**: Real-time KPI summaries (Pipeline Value, Average Project Size, Active Projects), recently updated estimates, and live team activity feed.
   - **Page 2 — Projects & Estimates**: Search, filter by status (`draft`, `quoted`, `approved`, `archived`), sort, rename, duplicate (with template retention choice), and archive/delete.
   - **Page 3 — Project Calculator**: Dynamic input forms with units and validations, real-time reactive outputs, quotation breakdown, steel cost, labour, and powder coating calculations.
   - **Page 4 — Calculation Details**: Full inspection of all 15 worksheets with spreadsheet matrix, formula bar (`fx`), and project-specific formula override capabilities.
   - **Page 5 — Organisation Settings**: Profile management, join code, current default template version, version history table, and draft publisher with rate diff preview.
   - **Page 6 — Members & Access**: Role-based access control (`Owner`, `Admin`, `Member`, `Viewer`), 6-character join code, and permission administration.
   - **Page 7 — Activity & Audit Trail**: Chronological event history tracking project creations, updates, duplications, template releases, and recalculations.

4. **Calculation Precision**:
   - Evaluates all 3,290 formula cells across all 15 sheets with **0 mismatches** against Microsoft Excel.

---

## 2-Minute Supabase Database Setup

1. Create a free project at [supabase.com](https://supabase.com).
2. Open the **SQL Editor** in your Supabase dashboard.
3. Click **New query**, paste the entire contents of [`supabase/schema.sql`](supabase/schema.sql), and click **Run**.
4. Go to **Project Settings → API** and copy:
   - **Project URL**
   - **anon public key**
5. (Optional) Paste them into `.env`:
   ```bash
   cp .env.example .env
   # VITE_SUPABASE_URL=https://your-project.supabase.co
   # VITE_SUPABASE_ANON_KEY=your-anon-key
   ```
   *If you do not configure `.env`, the application will prompt you on first launch and save the credentials safely in your browser.*

---

## Local Development & Testing

```bash
# 1. Install dependencies
npm install

# 2. Run the automated test suite (verifies all 15 test requirements)
npm test

# 3. Verify Excel workbook formula precision (3,290 cells)
npm run verify

# 4. Start the local development server
npm run dev
# Opens at http://localhost:5173
```

---

## Production Build & Deployment

The build process produces a single self-contained, hostable HTML bundle with all scripts and styles inlined:

```bash
npm run build
```

The output file is **`dist/index.html`** (~700 kB).

### Hosting Options:
- **Netlify**: Drag and drop `dist/index.html` or the `dist` folder into [app.netlify.com/drop](https://app.netlify.com/drop).
- **Vercel**: Deploy the repo using the standard Vite preset (`npm run build`, output: `dist`).
- **Cloudflare Pages / GitHub Pages**: Deploy the `dist` directory as static assets.
- **Standalone Offline / Intranet**: Open `dist/index.html` directly in any modern web browser or host on any static web server (Nginx, Apache, S3).

---

## Security & Data Isolation Architecture

- **Row Level Security (RLS)** is enabled on all tables (`app_users`, `app_sessions`, `organisations`, `org_members`, `templates`, `projects`, `project_revisions`, `activity_logs`).
- Public direct table access is restricted. All mutations and queries occur through Postgres `security definer` functions that validate the user's session token and verify organisation membership.
- Password hashing is performed inside PostgreSQL using `pgcrypto` (`crypt/gen_salt('bf', 10)`). Plaintext passwords are never stored or exposed.
- Service role keys and secrets are never embedded in client code.

---

## Documentation Directory

- [`docs/CALCULATION_MAPPING.md`](docs/CALCULATION_MAPPING.md): Detailed formula mapping for all 15 sheets and original workbook error documentation.
- [`docs/VALIDATION_REPORT.md`](docs/VALIDATION_REPORT.md): Formula accuracy test report and 15-point verification matrix.
- [`supabase/schema.sql`](supabase/schema.sql): Complete database schema with template versioning, revisions, and security definer functions.
- [`supabase/migrations/002_template_versioning.sql`](supabase/migrations/002_template_versioning.sql): Safe, idempotent migration script for existing deployments.
