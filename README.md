# Rewrite Card (Excel side panel)

Adds a **Rewrite Card** button to the Home tab in Excel. Click it, then click any policy row in
CalFairPlanRewrites.xlsx (Compare, Quote Options, Work Queue or No Option). The panel shows that
client's FAIR Plan renewal, current DIC and every priced rewrite option, color-coded:

- green = cheaper than what the client pays today (FAIR Plan renewal + current DIC)
- red = more expensive
- blue = a new DIC where the client has none

The panel reads the open workbook only. This repository holds code, not customer data.

## Files

| File | What it is |
|---|---|
| `taskpane.html`, `taskpane.css`, `taskpane.js` | The panel and its Excel hookup |
| `card.js` | Builds the card from a Quote Options row (no Excel code, testable on its own) |
| `manifest.xml` | The add-in file you load into Excel / Microsoft 365 |
| `assets/` | Icons |

## How it finds the data

- Reads row 1 (carrier groups) and row 2 (field names) of **Quote Options**, so columns can move
  as long as the header names stay the same.
- Matches the selected row by its **Policy #** value, so sorting and filtering are fine.
- Status and Agent Direction come from the **Status** and **Agent Direction** columns of the `Compare` table.
- Click **Refresh** at the bottom of the panel after big changes (new columns, rows re-sorted on Quote Options).

## One-time GitHub setup

Repo Settings → Pages → Build and deployment → Source: **Deploy from a branch**, Branch: `main`, folder `/ (root)` → Save.

## Install for yourself (test)

Excel for the web: Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in → pick `manifest.xml`.

## Roll out to the team

Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps → Office Add-in →
upload `manifest.xml` → assign users → Deploy. It can take a few hours to show up in everyone's Excel.

## Updating

Change files here and the panel updates for everyone on next open. Only re-deploy `manifest.xml`
if the manifest itself changes (bump `<Version>`).
