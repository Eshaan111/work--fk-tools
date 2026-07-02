# Add Product Guide

This guide explains how to upgrade the app when you want to add:

1. A new vertical, such as `shirt`, `jogger`, or `kurta`
2. A new kind inside an existing vertical, such as a new jeans wash or fit

The system is now mostly config-driven. In most cases, you do not need to edit Python code.

## Mental Model

There are 4 layers:

1. `shared.products`
This defines product-level behavior such as default sizes, sheet names, and optional legacy flow steps.

2. `laptops.<LAPTOP>.verticals`
This defines which verticals and kinds are available on that laptop, and which image folder each kind uses.

3. `data inputs/...`
This holds the Excel files used by the bot for that product.

4. `json_LC_creation/...`
This holds the JSON flow files that tell the bot how to fill the listing UI.

If the new product uses the same page behavior and same field structure, you usually only add config and files.

## Quick Decision

Add a new `kind` when:

- The product family is the same
- The form fields are the same
- Only the images or naming differ

Add a new `vertical` when:

- The product family is different
- The Excel sheet names differ
- The listing flow or field set differs

## Folder And File Rules

The current routing logic uses these shared patterns from `config.json`:

```json
"flow_directory_pattern": "{product_type}_{surface}",
"product_input_patterns": {
  "product_description_excel": "{product_type}/Product-Description-inputs{surface_suffix}.xlsx",
  "additional_description_excel": "{product_type}/Additional-Description-inputs{surface_suffix}.xlsx",
  "product_description_flow_json": "02_product_description.json",
  "additional_description_flow_json": "01_additional_description.json"
}
```

That means:

- Excel files live under `data inputs/<product_type>/`
- Flow folders live under `json_LC_creation/<product_type>_<surface>/`
- For `flipkart`, the Excel name usually has no suffix
- For `shopsy`, the Excel name usually uses `-Shopsy`

Examples:

- `data inputs/jeans/Product-Description-inputs.xlsx`
- `data inputs/jeans/Additional-Description-inputs.xlsx`
- `data inputs/trouser/Product-Description-inputs-Shopsy.xlsx`
- `json_LC_creation/jeans_flipkart/02_product_description.json`
- `json_LC_creation/jeans_flipkart/01_additional_description.json`

## Option A: Add A New Kind Inside An Existing Vertical

Example: add `Blue-Baggy` under `jeans`

### Step 1: Add the kind under each laptop that should support it

In `config.json`, add it inside:

```json
"laptops": {
  "VAIO": {
    "verticals": {
      "jeans": {
        "kinds": {
          "1": { "kind": "Beige", "image_directory": "..." },
          "2": { "kind": "Ice", "image_directory": "..." },
          "6": {
            "kind": "Blue-Baggy",
            "image_directory": "C:/your-path/LISTING IMAGES AUTOMATED/BLUE-BAGGY"
          }
        }
      }
    }
  }
}
```

Notes:

- The key values like `"1"`, `"2"`, `"6"` are just ordering keys
- Keep them numeric strings
- The UI reads and sorts them numerically

### Step 2: Add the image folder

Create the image folder path you referenced in config.

Example:

```text
C:/your-path/LISTING IMAGES AUTOMATED/BLUE-BAGGY
```

### Step 3: Reuse existing Excel and JSON flows if behavior is the same

If `Blue-Baggy` uses the same jeans inputs and same jeans field flow:

- no new Python code
- no new product Excel file
- no new flow JSON file

Only the kind entry and image directory are required.

### Step 4: Restart the app and verify

The startup UI should now show:

1. `Vertical`: `jeans`
2. `Surface`: only surfaces that exist for jeans
3. `Kind`: now includes `Blue-Baggy`

## Option B: Add A Completely New Vertical

Example: add `shirt`

### Step 1: Add the product definition in `shared.products`

Add a new block under:

```json
"shared": {
  "products": {
    "shirt": {
      "default_kind_by_surface": {
        "default": "Casual-Shirt"
      },
      "default_size_by_surface": {
        "flipkart": "M",
        "shopsy": "M",
        "default": "M"
      },
      "sheet_names": {
        "product_description_by_surface": {
          "flipkart": "Shirt Product Inputs",
          "shopsy": "Shirt Product Inputs"
        },
        "additional_description_by_surface": {
          "flipkart": "Shirt Addl Desc Inputs",
          "shopsy": "Shirt Addl Desc Inputs"
        }
      }
    }
  }
}
```

Add `legacy_flow_steps` only if this product must run a legacy page flow instead of a JSON flow.

### Step 2: Add the variants sheet mapping if the product has variants

If the product uses the common variants Excel:

```json
"common_inputs": {
  "variants": {
    "sheet_name_by_product_type": {
      "jeans": "Jeans Variant Inputs",
      "shirt": "Shirt Variant Inputs"
    }
  }
}
```

If the product does not use variants, you can leave this out unless the flow requires it.

### Step 3: Add the vertical to each laptop that should use it

Example:

```json
"laptops": {
  "VAIO": {
    "verticals": {
      "shirt": {
        "kinds": {
          "1": {
            "kind": "Casual-Shirt",
            "image_directory": "C:/your-path/LISTING IMAGES AUTOMATED/CASUAL-SHIRT"
          },
          "2": {
            "kind": "Formal-Shirt",
            "image_directory": "C:/your-path/LISTING IMAGES AUTOMATED/FORMAL-SHIRT"
          }
        }
      }
    }
  }
}
```

Repeat for `ASUS` or any other laptop if needed.

### Step 4: Create the Excel files

Create:

- `data inputs/shirt/Product-Description-inputs.xlsx`
- `data inputs/shirt/Additional-Description-inputs.xlsx`

If `shopsy` needs different files, also create:

- `data inputs/shirt/Product-Description-inputs-Shopsy.xlsx`
- `data inputs/shirt/Additional-Description-inputs-Shopsy.xlsx`

### Step 5: Create the JSON flow folder

Create:

- `json_LC_creation/shirt_flipkart/01_additional_description.json`
- `json_LC_creation/shirt_flipkart/02_product_description.json`

If `shopsy` has different flow behavior, also create:

- `json_LC_creation/shirt_shopsy/01_additional_description.json`
- `json_LC_creation/shirt_shopsy/02_product_description.json`

### Step 6: Decide whether to reuse or create a new flow

Reuse an existing flow if:

- The same fields exist
- The same tab order exists
- The same listing page behavior exists

Create a new flow JSON if:

- Field labels differ
- Dropdown behavior differs
- Required fields differ
- Variant handling differs

### Step 7: Test in the UI

After restart:

1. Choose the laptop
2. Choose the profile
3. Choose `shirt` as the vertical
4. Choose surface
5. Confirm the kinds appear
6. Run one test listing

## Surfaces And Availability

The UI now treats `vertical` and `surface` separately.

Expected behavior:

1. User chooses a vertical first
2. The UI only shows surfaces that are valid for that vertical
3. The UI only shows kinds configured for that vertical on the active laptop

So if a surface does not appear, check:

- the product exists in `shared.products`
- the flow folder exists for that surface
- the expected Excel files exist for that surface

## What Usually Needs No Code Change

These changes should stay config-only:

- adding a new kind
- changing image folder paths
- adding a new laptop
- adding a new profile
- adding brands inside an existing profile structure
- adding product input files
- adding JSON flow files with the same schema

## What Usually Does Need Code Change

You may need Python changes only if:

- the website introduces a new widget type
- the flow needs a brand-new tab or page step
- the product has a field type not supported by the current JSON flow runner
- the success/verification logic must change

## Safe Upgrade Checklist

Before calling a new product ready, verify:

1. `config.json` has the new product or kind
2. image directories exist
3. Excel paths exist
4. JSON flow paths exist
5. sheet names in config match the real workbook sheet names
6. the startup UI shows the new option
7. one dry run completes
8. batch result JSON is created
9. no unexpected lock or launch errors appear

## Smallest Possible Changes

### To add only a new kind

Change only:

- `config.json`
- image folder on disk

### To add a new vertical with existing flow style

Change:

- `config.json`
- `data inputs/<vertical>/...`
- `json_LC_creation/<vertical>_<surface>/...`

### To add a new vertical with new UI behavior

Change:

- `config.json`
- Excel files
- JSON flow files
- Python code only if current flow primitives cannot support it

## Recommended Product Strategy

For easier upgradability:

1. Keep adding new sellable variations as `kinds` whenever possible
2. Create a new `vertical` only when the listing form meaningfully changes
3. Reuse flow JSON patterns between products whenever the website structure matches
4. Keep folder names simple and stable so setup remains easy for future customers

## Related Files

- `config.json`
- `main_ui_themed.py`
- `setup.py`
- `data inputs/`
- `json_LC_creation/`
- `onboarding-steps.md`
