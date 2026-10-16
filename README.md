# Bakery Business Manager Pro (BakeSuite Pro PWA)

A standalone, 1-click installable Progressive Web App (PWA) converted from your BakeSuite Pro Google Sheets ecosystem.

## 🚀 Features

- **100% Offline Standalone PWA:** Installs directly to desktop or mobile via Chrome/Edge/Safari with full offline support via `sw.js`.
- **Single-File Overwrite Auto-Save:** Connects directly to `Bakery_Master_Data.json` on your hard drive via the **File System Access API**. Background saves debounced to 1 second whenever orders, inventory, or recipes are modified — **no duplicate file downloads**.
- **Locked Business Formulas:** All financial logic (COGS, margins, batch multipliers, inventory thresholds, tier servings, dowel counts) is locked in JavaScript to prevent formula tampering.
- **10 Integrated Bakery Modules:**
  1. **Executive Dashboard:** Real-time revenue, COGS, profit, unpaid receivables, pipeline breakdown, upcoming deliveries.
  2. **Orders & CRM:** Order status tracking, balance due calculation, days-until-event countdown, search & filtering.
  3. **Recipe Costing & Scaler:** Ingredient-level cost breakdown, labor hours & rates, overhead %, target markup, and dynamic batch production scaler (0.5x, 1x, 2x, 4x).
  4. **Master Inventory:** Live stock management with reorder threshold alerts and total valuation.
  5. **Cake Portion & Multi-Tier Planner:** Up to 5-tier architecture, dowel structural calculator, batter volume requirements, and pan surface area converter.
  6. **Weekly Prep & Grocery List:** Production week aggregator comparing recipe needs against live stock with shortage costs.
  7. **Invoice Generator:** Printable / PDF-ready customer invoices with payment methods (Zelle, Venmo, ACH).
  8. **Cottage Food Box Labels:** Regulatory compliant cottage food packaging labels with allergen warnings.
  9. **Contract Agreement:** Custom event contract with non-refundable deposit terms, temperature liability, and signature blocks.
  10. **Settings & Setup:** Full business profile customization and pan dimension matrices.

## 📁 Deployment to GitHub Pages

1. Create a new GitHub repository (e.g. `bakery-business-manager`).
2. Push all files from this directory (`index.html`, `styles.css`, `app.js`, `manifest.json`, `sw.js`, `icon.png`, `icon-192.png`) to the `main` branch.
3. In GitHub repo settings, go to **Pages** -> Source: **Deploy from branch (main / root)**.
4. Your PWA is live with HTTPS! Open the link in Chrome/Edge and click `🖥️ Install App to PC`.
