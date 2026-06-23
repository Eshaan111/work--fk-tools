---
name: Precise Ledger
colors:
  surface: '#f7f9fb'
  surface-dim: '#d8dadc'
  surface-bright: '#f7f9fb'
  surface-container-lowest: '#ffffff'
  surface-container-low: '#f2f4f6'
  surface-container: '#eceef0'
  surface-container-high: '#e6e8ea'
  surface-container-highest: '#e0e3e5'
  on-surface: '#191c1e'
  on-surface-variant: '#45474c'
  inverse-surface: '#2d3133'
  inverse-on-surface: '#eff1f3'
  outline: '#75777d'
  outline-variant: '#c5c6cd'
  surface-tint: '#545f73'
  primary: '#091426'
  on-primary: '#ffffff'
  primary-container: '#1e293b'
  on-primary-container: '#8590a6'
  inverse-primary: '#bcc7de'
  secondary: '#515f74'
  on-secondary: '#ffffff'
  secondary-container: '#d5e3fd'
  on-secondary-container: '#57657b'
  tertiary: '#001624'
  on-tertiary: '#ffffff'
  tertiary-container: '#002c42'
  on-tertiary-container: '#0099d9'
  error: '#ba1a1a'
  on-error: '#ffffff'
  error-container: '#ffdad6'
  on-error-container: '#93000a'
  primary-fixed: '#d8e3fb'
  primary-fixed-dim: '#bcc7de'
  on-primary-fixed: '#111c2d'
  on-primary-fixed-variant: '#3c475a'
  secondary-fixed: '#d5e3fd'
  secondary-fixed-dim: '#b9c7e0'
  on-secondary-fixed: '#0d1c2f'
  on-secondary-fixed-variant: '#3a485c'
  tertiary-fixed: '#c9e6ff'
  tertiary-fixed-dim: '#89ceff'
  on-tertiary-fixed: '#001e2f'
  on-tertiary-fixed-variant: '#004c6e'
  background: '#f7f9fb'
  on-background: '#191c1e'
  surface-variant: '#e0e3e5'
  success-emerald: '#10B981'
  error-rose: '#F43F5E'
  warning-amber: '#F59E0B'
  locked-gray: '#94A3B8'
  chart-line: '#6366F1'
typography:
  headline-lg:
    fontFamily: Hanken Grotesk
    fontSize: 28px
    fontWeight: '700'
    lineHeight: 36px
    letterSpacing: -0.02em
  headline-md:
    fontFamily: Hanken Grotesk
    fontSize: 20px
    fontWeight: '600'
    lineHeight: 28px
  body-lg:
    fontFamily: Inter
    fontSize: 16px
    fontWeight: '400'
    lineHeight: 24px
  body-md:
    fontFamily: Inter
    fontSize: 14px
    fontWeight: '400'
    lineHeight: 20px
  body-sm:
    fontFamily: Inter
    fontSize: 12px
    fontWeight: '400'
    lineHeight: 18px
  label-md:
    fontFamily: JetBrains Mono
    fontSize: 13px
    fontWeight: '500'
    lineHeight: 16px
  label-sm:
    fontFamily: JetBrains Mono
    fontSize: 11px
    fontWeight: '500'
    lineHeight: 14px
    letterSpacing: 0.05em
  data-table:
    fontFamily: Inter
    fontSize: 13px
    fontWeight: '400'
    lineHeight: 16px
  kpi-value:
    fontFamily: Hanken Grotesk
    fontSize: 24px
    fontWeight: '700'
    lineHeight: 32px
rounded:
  sm: 0.125rem
  DEFAULT: 0.25rem
  md: 0.375rem
  lg: 0.5rem
  xl: 0.75rem
  full: 9999px
spacing:
  unit: 4px
  container-padding: 24px
  gutter: 16px
  row-height-dense: 32px
  row-height-standard: 44px
  sidebar-width: 280px
---

## Brand & Style

The design system is engineered for high-utility enterprise environments where data density and accuracy are paramount. The brand personality is **authoritative, surgical, and efficient**, catering to logistics and finance professionals who manage high-volume settlement data.

The chosen aesthetic is **Corporate / Modern**, leaning into a highly structured, functional layout. It prioritizes information hierarchy through a rigorous grid system and clear typographic distinctions. The interface avoids unnecessary decoration, using subtle tonal layering and crisp borders to define workspaces. The emotional response should be one of "controlled complexity"—giving the user confidence that they can manipulate thousands of rows with total precision.

- **Focus:** Data density, legibility, and state clarity.
- **Visual Cues:** Systematic alignment, subtle elevation for interactive surfaces, and high-contrast accents for critical statuses (Success/Error).

## Colors

The palette is anchored by a reliable navy (`#1E293B`), representing stability and institutional trust. This primary color is used for the header and core structural elements to ground the application.

- **Primary & Secondary:** Used for structural hierarchy, primary actions, and typography.
- **Accents:** Semantic colors are strictly reserved for data status. `success-emerald` is used for "Active" listings and "Accept" decisions. `error-rose` is used for "Inactive" listings, "Reject" decisions, and "Auto-Flags." 
- **Neutral:** A cool-toned gray stack (`#F8FAFC` to `#94A3B8`) provides the canvas for the dense data tables, ensuring the UI remains "quiet" while the data is "loud."
- **Data Visualization:** A distinct indigo (`#6366F1`) is used for the settlement chart to separate analytical visualization from functional UI controls.

## Typography

This design system uses a tri-font system to categorize information:
1. **Hanken Grotesk** is used for headlines and KPI values. Its sharp, contemporary geometry provides a professional, "Tech-SaaS" feel for high-level metrics.
2. **Inter** is the workhorse for all body text and data tables. It is chosen for its exceptional legibility at small sizes and high x-height, essential for dense spreadsheet-style views.
3. **JetBrains Mono** is used for labels, SKU IDs, and technical metadata. The monospaced nature helps users quickly compare alphanumeric strings (like SKUs and Price Caps) where character alignment is crucial.

For mobile-specific views, `headline-lg` should scale down to 22px to maintain visual balance on smaller devices.

## Layout & Spacing

The layout follows a **Fixed-Fluid Hybrid** model. Navigation and toolbars are fixed to the viewport edges, while the data table and settlement chart occupy a fluid central area that expands to utilize all available screen real estate.

- **Rhythm:** A 4px baseline grid governs all spacing. 
- **Grid:** A 12-column system is used for top-level KPI cards and filter blocks. On desktop, KPI cards span 2 or 3 columns depending on density.
- **Data Density:** The primary listing table uses a "Dense" vertical rhythm (32px rows) to maximize visible data. 
- **Breakpoints:**
  - **Desktop (1280px+):** Full 12-column grid, persistent sidebar for Change Logs.
  - **Tablet (768px - 1279px):** Filter strip collapses into a drawer; Change Log moves to a bottom sheet.
  - **Mobile (<767px):** Single column view. KPI cards become a horizontal scrollable carousel. Data table converts to "Card" view per row.

## Elevation & Depth

To maintain a "Professional SaaS" aesthetic, depth is communicated through **Tonal Layering** and **Low-Contrast Outlines** rather than heavy shadows.

- **Surface 1 (Base):** `#F8FAFC` (Neutral) for the application background.
- **Surface 2 (Containers):** `#FFFFFF` with a 1px border of `#E2E8F0`. This is used for KPI cards, the Listing Table, and the Filter Strip.
- **Surface 3 (Overlays):** Modals and dropdowns use a subtle ambient shadow (0px 4px 12px, 5% opacity primary color) to indicate temporary focus.
- **Interaction:** Hovering over a table row or a chart point triggers a subtle background shift to `#F1F5F9` (Primary 50).
- **Frozen State:** Locked rows use a stippled or slightly desaturated background to visually "recede" behind active data.

## Shapes

The design system uses a **Soft** shape language (`roundedness: 1`). 

- **Standard Elements:** Buttons, Input Fields, and Checkboxes use a 4px (0.25rem) corner radius. This provides a modern touch without sacrificing the "precise" feel of an enterprise tool.
- **Large Elements:** Cards and Modals use 8px (0.5rem) to distinguish them from data rows.
- **Interactive Pills:** Status indicators (Active/Inactive) use 12px (0.75rem) to stand out as distinct, clickable tokens within the angular table grid.

## Components

### Buttons
- **Primary:** Solid `#1E293B` with white text. High emphasis.
- **Secondary:** Outlined `#1E293B` with a 1px border. 
- **Destructive/Error:** Solid `#F43F5E` for actions like "Unfreeze All" or "Reset Selection."

### Status Chips
- **Active:** Emerald background (10% opacity) with Emerald text.
- **Inactive/Error:** Rose background (10% opacity) with Rose text.
- **Flagged:** Amber background with an icon prefix to denote "Reason."

### Input Fields
- Use `Inter` for values and `JetBrains Mono` for labels. 
- Focused state uses a 1px `tertiary_color` (Sky Blue) border and a 2px outer glow.

### KPI Cards
- Large `hankenGrotesk` values. 
- Include a small sparkline or trend indicator if relevant to the "Settlement Chart" logic.

### Listing Table
- Header row is sticky with a `#F1F5F9` background.
- Cell borders are minimal (horizontal only) to keep the eye moving across the row.
- **Locked Rows:** Apply a lock icon prefix to the SKU and reduce text opacity to 60%.

### Change Log Cards
- Vertical list of small containers.
- Use a color-coded left border (Emerald for compute, Rose for deletes/resets, Blue for bulk edits).