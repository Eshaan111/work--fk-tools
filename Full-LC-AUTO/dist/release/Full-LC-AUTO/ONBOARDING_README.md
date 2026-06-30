# Full LC Auto Onboarding

## What This Does

`setup.py` is the onboarding/admin app.

Use it to define:
- a new laptop
- one or more seller profiles
- one or more brands inside each profile
- one or more verticals
- one or more kinds inside each vertical
- scaffold source for each vertical

It can then:
- generate or update `config.json`
- export a reusable `setup-package.json`
- create starter scaffold files for the new verticals

`main_ui_themed.py` is the operator/run app.

Use it after onboarding to:
- choose laptop, profile, vertical, kind, brand, run count
- run image folder insight
- run preflight checks
- start a listing batch

## Verified Dry Run

A code-level dry run was executed on June 30, 2026 with an off-screen Tk session.

The dry run created:
- laptop: `DRYRUN_LAPTOP`
- profile: `demo_profile`
- brands: `Demo Brand`, `Backup Brand`
- vertical: `cargo_pants`
- kinds: `Cargo`, `Linen`
- scaffold source: `trouser`

The dry run successfully verified:
- config payload generation
- scaffold generation into clean temp folders
- setup package export data
- setup package reload into a new setup session

Sample generated scaffold targets included:
- `assets_root/Price-Stock-Shipping-inputs.json`
- `data_inputs_root/cargo_pants/Product-Description-inputs.xlsx`
- `data_inputs_root/cargo_pants/Additional-Description-inputs.xlsx`
- `data_inputs_root/common/Variants-excel.xlsx`
- `flow_root/cargo_pants_flipkart`

## Onboarding Flow

1. Run `Full-LC-AUTO/setup.py`.
2. Enter the laptop name.
3. Choose the snapshot directory.
4. Set the shared roots if this laptop stores project files in a different location:
   - `run_helpers`
   - latest error file
   - success record workbook
   - `json_LC_creation`
   - `data inputs`
   - `assets`
5. Add one or more accounts.
6. For each account, enter:
   - profile name
   - alias
   - Firefox profile folder
   - one or more brands with brand codes
7. Add one or more verticals.
8. For each vertical, enter:
   - vertical name
   - default kind
   - default size
   - scaffold source
   - one or more kinds with image folders
9. Keep `Create starter scaffold files` enabled unless you already prepared files manually.
10. Click `Preview Output`.
11. Click `Generate Files`.
12. If this machine should become active immediately, enable `Also replace config.json` before generating.

## Setup Package Flow

Use setup packages when moving a configured seller setup to another laptop.

Export:
1. Open `setup.py`.
2. Fill or load the onboarding form.
3. Click `Export Setup Package`.
4. Save the JSON file.

Import:
1. Open `setup.py` on the target laptop.
2. Click `Import Setup Package`.
3. Select the exported JSON file.
4. Review paths such as Firefox profiles, snapshots, and project roots.
5. Generate files or replace `config.json` as needed.

## Operator Flow

1. Run `Full-LC-AUTO/main_ui_themed.py`.
2. Select:
   - laptop
   - profile
   - flow target
   - kind
   - size
   - brand
   - run count
3. Optionally run `Refresh Insight`.
4. Run `Preflight Check`.
5. Fix any reported issues.
6. Click `Start Batch`.

## What Preflight Checks

The preflight check validates the current selection before a live run.

It checks:
- Firefox profile path
- image directory
- price/stock/shipping Excel and JSON
- product description Excel and JSON
- additional description Excel and JSON
- variants workbook
- expected worksheet names
- resolved listing URL
- JSON flow or legacy flow availability
- current insight workbook availability

## Notes

- `setup.py` now keeps only onboarding information. It does not manage runtime choices like listing URL, primary surface, or run count for a live operator session.
- `main_ui_themed.py` is the runtime UI and should be used for day-to-day listing operations.
- Scaffold generation gives a strong starting point, but a brand-new vertical may still need field-mapping review depending on how different it is from its scaffold source.
