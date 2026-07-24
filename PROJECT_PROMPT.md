# Panoramic Ops PWA — Project Instructions

You are helping Aidan Salisbury build and maintain the **Panoramic Ops PWA** — a mobile-first progressive web app for Panoramic Building LLC that manages purchase orders, time tracking, PTO, payroll, and field tools.

**Before any app work: invoke the `panoramic-app-builder` skill.** Then read `panoramic-ops-context.json` in the working folder for full technical context.

---

## The Stack

| Layer | Detail |
|---|---|
| Frontend | Single-file PWA — `index.html` (~5,600 lines, vanilla JS + inline CSS) |
| Backend | Google Apps Script (GAS) — `PO_Manager_Code.gs` |
| Hosting | GitHub Pages — `ops.panoramicbuildingllc.com` (custom domain via `CNAME` file; legacy URL `panoramic01.github.io/po-manager/` still works) |
| GAS URL | `https://script.google.com/macros/s/AKfycbwvX4xv5kDofWEldpSRJP8O2yeVVgMcYVQsgtKmuFyKBZWTj0vfhFIIZAJp2UQob_XJ-Q/exec` |
| Data | Google Sheets (HR, PO Database, Time Tracking) |
| Auth | Email + password checked against HR sheet via GAS |
| Integrations | Asana (PTO/tasks), Google Drive (photo uploads) |

---

## Rules That Must Never Be Broken

1. **index.html is ~5,600 lines — never use the Edit tool on it.** It truncates large files. Always use Python scripts (read full file → str.replace → write back). After every write verify `</html>` exists: `python3 -c "h=open('index.html').read(); print(h.count('</html>'), repr(h[-60:]))"`
2. **sw.js** — edit via `sed -i` only. Bump the cache version (`po-manager-vN`) on every push that changes JS or CSS.
3. **GAS V8 rules:**
   - No nested `function foo(){}` inside blocks — use `var foo = function(){}` instead
   - ASCII only in string literals — no em dashes (—), curly quotes, or Unicode punctuation
   - Always verify brace balance after edits: `python3 -c "t=open('PO_Manager_Code.gs').read(); print(t.count('{'), t.count('}'))"`
   - Always redeploy after changes: Deploy → Manage Deployments → pencil → New version → Deploy
4. **gasCall()** — must have NO `Content-Type` header (GAS rejects OPTIONS preflight).
5. **Git on PowerShell** — no `&&` chaining. Separate lines: `git add`, `git commit`, `git push`.
6. **JS syntax** — always run `node --check` on extracted script before pushing:
   ```bash
   python3 -c "h=open('index.html').read(); js=h[h.find('<script>')+8:h.rfind('</script>')]; open('/tmp/c.js','w').write(js)"
   node --check /tmp/c.js
   ```
7. **Tab CSS** — tab divs must NOT have inline `style="display:none"`. The `.tab` CSS class handles hiding; `.tab.active` shows it.

---

## gasCall() — Callback Style (NOT promise chains)

The entire codebase uses callbacks. Do not write `.then()` chains — they won't work.

```javascript
// CORRECT
gasCall('action', { key: value }, function(data) {
  // success — data is the GAS return value
}, function(err) {
  // failure — network or parse error
});

// WRONG — do not use
gasCall('action', {}).then(function(data) { ... });
```

---

## Key Utility Functions (already exist — do not reinvent)

| Function | Purpose |
|---|---|
| `esc(str)` | HTML-escape dynamic content before injecting into innerHTML |
| `showToast(msg)` | Show user feedback toast notification |
| `sd(id, show)` | Show/hide element by ID — `sd('some-id', true/false)` |
| `fmtMoney(n)` | Format number as currency string |
| `today()` | Returns current date string |
| `gasCall(action, payload, onSuccess, onFailure)` | GAS fetch wrapper |
| `hideSplash()` | Dismiss the navy loading splash screen |
| `showToast(msg)` | Brief toast notification |
| `showOtherPanel(name)` | Navigate to a named Other tab panel |
| `showOtherHome()` | Return to Other tab home grid |
| `showTabDirect(name)` | Switch to a main nav tab by name |

---

## App Structure

**Nav tabs:** Dashboard · Log (tab-create) · Other (tab-other) · Account (tab-account)

**Other tab panels:** pricing, contacts, reconcile, job-cost, missing-inv, vendor-spend, material-report, sop, pto, timeclock, emp-mgr (admin), pto-overview (admin), payroll (admin)

**ALL_PANELS array** must be kept in sync in both `showOtherHome()` and `showOtherPanel()` forEach loops. A missing name causes the panel to stay visible as a ghost on other pages.

---

## Google Sheets Column Reference

**HR sheet:**
| Col | Field |
|-----|-------|
| A | Name |
| B | Email |
| C | Phone |
| D | Role |
| E | Password |
| F | Allotted PTO |
| G | Used PTO |
| H | Remaining PTO |

**Time Tracking sheet:** A=Name, B=Email, C=Date, D=ClockIn, E=ClockOut, F=Hours

---

## loadAccount() Field Path Reference

These exact paths caused bugs when wrong — use these:

```javascript
d.balance.allotted    // NOT d.allotted
d.balance.used
d.balance.remaining
d.myRequests          // NOT d.history
r.dates               // NOT r.startDate / r.endDate
r.status              // lowercase: 'approved', 'pending', 'denied' — NOT 'Approved'
```

---

## Pay Periods

Semi-monthly (NOT bi-weekly):
- Day 1–15 → period is 1st–15th of that month
- Day 16–end → period is 16th–last day of that month

Function: `getPeriodBounds(date)` in GAS returns `{ start, end }`.

---

## Roles

Values: `admin`, `office`, `site_manager`, `runner`, `accountant`

- Admin only: emp-mgr, pto-overview, payroll panels; pricing; reconcile; job-cost; missing-inv; vendor-spend; material-report
- Office + admin: contacts, New PO button, create tab, other tab
- All roles: account tab, sop panel

---

## Features Already Built

PO management, photo upload to Drive (5 per PO), exterior quality check (Prewalk/Daily/Closeout Walk checklists, Pass/Flag/N/A per item, submitter auto-attributed from signed-in account, stored purely as Asana subtasks — no separate sheet), job cleanup, SOP viewer, pricing lookup, contacts, reconciliation, job cost, missing invoices, vendor spend chart, material report, PTO requests (Asana-backed), time clock (clock in/out, hours this period), account tab (profile edit, time clock, PTO summary), employee manager (admin), PTO overview (admin), payroll summary (admin), back-button navigation (History API), service worker caching.

## Features Planned

- **Office Notes** — saves to Notes sheet by default; optional per-note Asana subtask (under date-based parent task in a chosen project); flag per note to also save to Notes sheet
- **Announcements** — admin posts a once-per-version message; employees dismiss via localStorage
- **Jobs Registry** — new "Jobs" tab in the PO Database spreadsheet (Job Name, GC, Asana GID, QuickBooks ID, Status) as the single source of truth linking a job across systems. "Add Job" admin panel lets office link an existing Asana task via search (option A — link only, doesn't create the Asana task). QuickBooks ID stays blank until QuickBooks OAuth integration exists. New PO eventually switches from free-text job entry to a searchable dropdown against this registry, once it has enough real jobs in it. Future option to revisit: create QuickBooks project + Asana task + registry link all in one action instead of linking only to pre-existing ones.
- **Site Walk Checklists (remaining)** — OSHA/Safety and Cleanup walk checklists for job sites. Quality's three walk types (Prewalk/Daily/Closeout) are now built using a walk-type selector + shared trade sections; OSHA/Safety and Cleanup still need their real checklist content supplied before building, but can reuse the same pattern (job picker, walk-type selector, per-category checklist, Pass/Flag/N/A, submit to Asana).

---

## Asana

- PAT stored in GAS Script Properties only — never in frontend code or chat
- PTO project GID: `1210392177822419`
- Sections: "New Requests", "Approved", "Denied"
- All Asana calls go through GAS — frontend never calls Asana directly

---

## Known Bug Patterns (check these before diagnosing)

| Symptom | Root cause |
|---|---|
| App stuck on navy splash forever | index.html truncated — missing `</script></body></html>` |
| Tab not showing when clicked | Tab div has inline `style="display:none"` overriding CSS class |
| Panel visible on wrong page | Panel name missing from ALL_PANELS forEach |
| Field shows `--` or undefined | Field path mismatch — check what GAS actually returns with `console.log` |
| Employee name shows as email | `getRoleByEmail()` not returning `name` field, or Time Tracking sheet has email in name col |
| Photo slot stuck on "Loading…" | iOS backgrounded JS mid-upload (FileReader suspended), or HEIC format (img.onerror not handled) |
| GAS change has no effect | Forgot to redeploy — push to GitHub does NOT update GAS |
| git index.lock error | Lock on Windows machine: `find .git -name 'index.lock' -exec rm {} \;` |
| node --check syntax error | Truncated function — check brace balance, look for cut-off lines near new additions |

---

## Deployment Checklist

1. Edit GAS → verify brace balance → redeploy
2. Edit index.html via Python script → verify `</html>` at end → `node --check`
3. Bump SW: `sed -i 's/po-manager-vN/po-manager-vM/' sw.js`
4. `git add -A` → `git commit -m "..."` → `git push` (separate lines in PowerShell)
5. Hard refresh app: pull-to-refresh on mobile or Ctrl+Shift+R desktop

---

## Skills Installed

- `panoramic-app-builder` — **invoke this first for any app work**
- `panoramic-sop` — builds SOP PDFs from screenshots
- `material-report` — fills Excel cost report from Drive invoices
- `po-received` — fills Google Form for PO logging
- `morning-assistant` — daily planning with Google Calendar + Tasks
- `ynab-financial-planner` — YNAB budget review
