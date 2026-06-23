# Rate Insight Dashboard Features

This document explains the behavior of `Rate_insight.py` in detail.

## Overview

`Rate_insight.py` is a PyQt5 desktop dashboard for reviewing settlement data from Excel/CSV files, classifying listings, detecting risky rows, bulk editing settlement values, applying offer logic, and exporting a processed file.

It is designed around a single loaded dataset and supports two working modes:

- `Normal Mode`
- `Offer Mode`

The mode is selected automatically from the columns found in the loaded file, and can be toggled only if both layouts are supported by the current file.

## Supported File Types

The dashboard can load:

- `.xlsx`
- `.xls`
- `.csv`

### Input mode detection

The app checks for these column sets:

#### Normal Mode columns

- `Seller SKU Id`
- `Product Title`
- `Bank Settlement`

#### Offer Mode columns

- `SKU ID`
- `FSN`
- `Selling Price(Rs)`

If neither layout is detected, the file is rejected as unknown.

## Main Screen Layout

The dashboard is divided into these areas:

1. Header action bar
2. Filter strip
3. KPI cards
4. Offer controls
5. Size tools
6. Freeze and bulk edit controls
7. Settlement chart
8. Summary panel
9. Listing table
10. Change log

## Header Actions

### Load Excel

Opens a file picker and loads the source file into memory.

When a file is loaded, the app:

- reads the file into a DataFrame
- trims column names
- resets row index
- creates internal helper columns
- detects the account name from the filename
- determines the working mode
- classifies listings
- detects sizes
- applies auto flags
- sets up filters
- refreshes the dashboard

### Export File

Exports the working dataset back to disk.

Before export, the app:

- sorts rows by original row order
- removes internal columns:
  - `__orig_index`
  - `__locked`
  - `__flag_exemptions`

Export behavior:

- CSV stays CSV
- XLSX stays XLSX
- XLS is exported as XLSX if needed

### Undo Last Move

Reverts the most recent tracked operation.

Undo stores:

- a deep copy of the full DataFrame
- current size overrides

Typical actions that create undo state:

- compute discount
- apply decision
- save size
- upload sizes
- freeze
- unfreeze all
- bulk edit
- set status
- manual table edits

### Offer Mode Toggle

Switches between normal and offer column mappings if both are supported by the loaded file.

The toggle changes:

- which SKU column is used
- which title column is used
- which settlement column is used
- whether offer-only controls are shown

## Filter Strip

The filter strip controls which rows are visible and editable.

### Search

Searches case-insensitively across:

- SKU column
- title column

### Listing Filter

Filters by `Listing Type`.

Possible values are generated from current data.

### Jeans Filter

Filters by `Jeans Type`.

Possible values are generated from current data.

### Size Filter

Filters by detected size.

Standard sizes:

- `26`
- `28`
- `30`
- `32`
- `34`
- `36`

If present, `UNDETECTED` also appears.

### Status Filter

Filters by `Listing Status`.

Supported display values:

- `ACTIVE`
- `INACTIVE`

### Settlement Max Slider

Only keeps rows whose settlement value is less than or equal to the slider value.

### Important filtering rule

Locked rows are excluded from normal working filters. Most actions operate only on:

- unlocked rows
- rows with valid settlement values
- rows matching the current filters
- rows inside the currently selected chart range, when range selection is active

## KPI Cards

The top metric cards summarize the current state.

### Total Loaded

Total rows in the loaded file.

### Visible Listings

Rows remaining after filters are applied, excluding chart-range selection.

### Exported Rows

Rows that are currently inside the active working selection.

This means:

- if no chart range is selected, this usually matches visible rows
- if a chart range is selected, this becomes the count inside that range

### Active Listings

Count of `ACTIVE` rows in the visible filtered set.

### Inactive Listings

Count of `INACTIVE` rows in the visible filtered set.

If `Listing Status` is missing, the status cards show `-`.

## Automatic Classification

When a file is loaded, the app enriches the data with additional derived fields.

### Listing Type

Derived from the title field.

If the title contains one of these owner keywords:

- `Starvielle`
- `Genz Vane`
- `INDIVANE`
- `FADEVIELLE`
- `FLEECRANE`

Then the row becomes:

- `Owner`

Otherwise:

- `Latched`

### Jeans Type

Derived from SKU and title text.

Current mapping rules:

- contains `white` in SKU -> `WHITE`
- contains `ice` or `blue` in SKU -> `ICE`
- contains `beige` or `cream` in SKU -> `BEIGE`
- contains `baggy` in SKU -> `BLACK-BAGGY`
- contains `black` in SKU and `relaxed` in title -> `BLACK-BAGGY`
- contains `black` in SKU -> `BLACK-PLAIN`
- otherwise -> `MIX`

### Size Detection

Size is derived from SKU using:

1. stored manual override, if present
2. special `_39` ending mapped to `32`
3. pattern matching such as:
   - `-26-`
   - `_26_`
   - `_26`
4. fallback to `UNDETECTED`

## Offer Controls

These controls are visible only in offer mode.

### Inputs

- `y%`
- `x%`
- `Cap Rs`

### Threshold Inputs

One threshold input exists for each jeans category:

- `ICE`
- `BEIGE`
- `WHITE`
- `BLACK-BAGGY`
- `BLACK-PLAIN`
- `MIX`

### Compute Discount

Applies discount logic to the current working rows.

Formula:

1. `base = y * settlement`
2. `discount = min(x * base, cap)`
3. `final_price = settlement - discount`

Outputs written:

- `Discount`
- `Final Price`

### Apply Decision

Uses `Final Price` and the jeans-type threshold to assign:

- `ACCEPT`
- `REJECT`

Rule:

- if `Final Price >= threshold` -> `ACCEPT`
- else -> `REJECT`

### Combined background logic

There is also an `apply_offer_logic()` helper that can compute:

- `Discount`
- `Final Price`
- `Decision`

in one pass.

## Size Tools

These tools help fix or manage size detection.

### Save Size

Manually stores a size override for one SKU.

Behavior:

- requires a loaded file
- requires a SKU input
- stores the mapping in `size_overrides.json`
- applies the new size immediately if matching rows are in the current working selection
- re-runs size flags and stock sync

### Download Undetected

Exports a template of all currently visible `UNDETECTED` SKUs.

Export file structure:

- `sku`
- `size`

This is intended to be filled in manually and re-uploaded.

### Upload Sizes

Imports size mappings from a file.

Accepted columns:

- `sku`
- `size`

Validation rules:

- column names are normalized to lowercase
- blank SKUs are ignored
- only valid size values are accepted
- duplicate SKUs keep the last occurrence

After import:

- overrides are saved to `size_overrides.json`
- matching live rows are updated
- flags are re-applied
- stock counts are synced
- a change is logged

## Flag System

The dashboard can automatically flag rows and force them inactive.

### Flag config source

Flag rules come from `flag_rules.xlsx`.

If that file does not exist, the app creates a default one.

### Supported flag categories

#### Title Keyword flags

If a configured title keyword appears in the title, the row is flagged.

#### SKU flags

If a configured SKU phrase appears in the SKU, the row is flagged.

#### Size flags

If the detected size is in the configured inactive-size list, the row is flagged.

### What a flag does

When a row is flagged:

- `Auto Flag` is populated with reasons
- `Listing Status` is forced to `INACTIVE` if that column exists
- stock count is later synced to `0`

### Flag detail storage

The app internally tracks:

- `__flag_size_reason`
- `__flag_title_reason`
- `__flag_sku_reason`
- `__auto_flag`

### Flag summary popup

When flagged rows are detected during certain actions, the app shows a popup summarizing the reasons.

### Flag exemptions

If a user explicitly chooses to activate flagged rows, exemptions are saved in `__flag_exemptions`.

Exemption tokens can be:

- `size:<value>`
- `title:<keyword>`
- `sku:<keyword>`

This lets rows remain active even if they still match the configured rule.

## Status Management

### Set Status

Bulk changes visible selected rows to:

- `ACTIVE`
- `INACTIVE`

### Activation safety behavior

If the target is `ACTIVE`, the app checks whether any selected rows are flagged.

The user is prompted with a choice:

- activate flagged rows
- keep flagged rows inactive

If flagged rows are activated, exemptions are stored so they do not get auto-blocked again for the same reasons.

### Stock count synchronization

If `Your Stock Count` exists:

- `ACTIVE` rows are set to `250`
- `INACTIVE` rows are set to `0`

This sync also runs in other workflows, including:

- load/classification
- size override changes
- uploaded size sheets
- manual SKU edits
- manual status edits

## Freeze System

Freeze is used to lock rows so they are excluded from most operations.

### Freeze

The user chooses:

- a column
- a text value

Rows are locked when the selected column contains the entered text.

Effect:

- `__locked = True`
- locked rows drop out of normal filtered operations

### Unfreeze All

Clears all row locks by setting `__locked = False`.

## Bulk Edit

Bulk edit changes settlement values for the current working set.

### Supported modes

- `Add`
- `Multiply`
- `Replace`

### Optional cap logic for multiply

When using `Multiply`, an optional cap can be applied:

- `Min` cap
- `Max` cap

Behavior:

- `Min` uses `np.minimum(result, cap)`
- `Max` uses `np.maximum(result, cap)`

### Target rows

Bulk edit applies only to rows that are:

- unlocked
- visible under filters
- inside selected chart range, if any

## Summary Panel

The summary panel shows:

- loaded / visible / export counts
- account name
- chart range summary
- selection statistics
- active/inactive size matrix table

### Account

The account label is derived from the input filename.

### Selection range

Shows:

- `None` if no chart range is active
- numeric low/high bounds if a range is selected

### Selection statistics

When a chart range is selected, the app shows:

- count
- mean
- median
- mode

### Active/Inactive matrix table

This table shows, for `ACTIVE` and `INACTIVE` rows:

- row count
- unique sizes present

## Listing Table

The main table shows the currently active working dataset.

### Table contents

It includes:

- original source columns
- derived columns
- internal helper columns unless explicitly filtered out elsewhere

### Editable behavior

The table supports direct editing.

Each edited cell updates the underlying DataFrame.

### Special edit handling

#### Editing SKU column

When the SKU field is edited:

- size is re-detected
- flags are re-applied
- stock is synced
- flag popup may appear

#### Editing Listing Status

When status is edited:

- value is normalized to uppercase
- activating rows triggers flag resolution flow
- stock is synced

#### Editing other fields

The value is written directly and logged.

## Settlement Chart

The chart visualizes settlement distribution.

### Chart data

The chart uses settlement values from the visible filtered set, excluding locked rows.

### Visual behavior

- x-axis: settlement values
- y-axis: frequency
- plotted as line + points

### Hover

Hovering a point shows:

- settlement value
- frequency

### Drag range selection

Dragging across the chart selects a settlement range.

This updates:

- export row count
- table rows
- selection stats
- selection summary
- change log

### Reset Selection

Clears the active chart range and restores the full visible filtered set.

## Change Log

The change log stores the latest tracked actions as styled cards.

Each entry includes:

- action name
- affected row count
- before min/max settlement
- after min/max settlement
- extra notes
- dataset context summary

### Context summary can include

- jeans type
- listing type
- size
- listing status

## Internal Helper Columns

The app creates internal columns for tracking state:

- `__orig_index`  
  Original row order for stable export.

- `__locked`  
  Whether the row is frozen out of normal operations.

- `__flag_exemptions`  
  Stored exemptions allowing flagged rows to remain active.

It may also create derived columns such as:

- `Listing Type`
- `Jeans Type`
- `Size`
- `Auto Flag`
- `Discount`
- `Final Price`
- `Decision`

## Data Cleaning Rules

Settlement values are cleaned by removing:

- commas
- `Rs`
- malformed rupee text fragments

Then the values are converted to numeric.

## Operational Rules

Most actions work on the current filtered mask.

The mask generally requires:

- row is not locked
- settlement is numeric
- settlement is within slider maximum
- row matches search/filter controls
- row matches status filter if enabled
- row lies inside selected chart range when selection is active

## Error Handling and Guardrails

The app warns or blocks when:

- file format is unsupported
- required offer inputs are invalid
- no file is loaded
- no rows are eligible for an action
- upload files are missing required columns
- upload files contain no valid size rows
- decision is attempted before computing final price
- status changes are attempted without `Listing Status`

## Files Used by the Dashboard

### `size_overrides.json`

Stores manual SKU-to-size overrides.

### `flag_rules.xlsx`

Stores flag rule configuration.

### `undetected_sizes.xlsx`

Used as a size-correction template export.

## Typical Workflow

1. Load a settlement file.
2. Let the app classify rows and apply initial flags.
3. Use filters to narrow the working set.
4. Fix size issues with manual overrides or upload sheet.
5. Review active/inactive distribution.
6. Optionally freeze rows that should not be touched.
7. In offer mode:
   - compute discount
   - apply decision
8. Use chart range selection to target a pricing band.
9. Apply bulk edits or status changes.
10. Review the change log.
11. Export the processed file.

## Known Design Characteristics

- The app operates entirely in-memory until export.
- Undo is single-step, not multi-level.
- Many operations are intentionally restricted to unlocked rows.
- Flagging is aggressive by design and can force rows inactive automatically.
- Manual activation of flagged rows is possible, but only through explicit override flow.

