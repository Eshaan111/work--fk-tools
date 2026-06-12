from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from selenium.common.exceptions import WebDriverException

from main import BotConfig, build_firefox_driver, log_event, prompt_for_profile


OUTPUT_DIR = Path("bot-page-scans")
DEFAULT_SURFACE = "flipkart"
DEFAULT_PRODUCT_TYPE = "jeans"
INPUT_TYPE_CHOICES = {
    "1": "text",
    "2": "number",
    "3": "combobox",
    "4": "tag_input",
    "5": "tag_input_commit",
    "6": "skip",
}


FIELD_SCAN_SCRIPT = r"""
function isVisible(element) {
    if (!element) {
        return false;
    }
    const style = window.getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    return style.display !== "none" &&
        style.visibility !== "hidden" &&
        rect.width > 0 &&
        rect.height > 0;
}

function cleanText(value) {
    return (value || "").replace(/\s+/g, " ").trim();
}

function getDirectText(element) {
    const clone = element.cloneNode(true);
    for (const child of Array.from(clone.children)) {
        child.remove();
    }
    return cleanText(clone.textContent);
}

function nearestFieldContainer(labelElement) {
    let current = labelElement;
    for (let depth = 0; current && depth < 8; depth += 1) {
        const controls = current.querySelectorAll(
            "input, textarea, button[role='combobox'], [role='combobox'], " +
            "[data-testid='trigger-multi-select'], .rti--container"
        );
        if (controls.length > 0) {
            return current;
        }
        current = current.parentElement;
    }
    return labelElement.parentElement || labelElement;
}

function detectInputType(container) {
    const textArea = Array.from(container.querySelectorAll("textarea")).find(isVisible);
    if (textArea) {
        return ["text", "nearest textarea"];
    }

    const reactTags = Array.from(container.querySelectorAll(".rti--container")).find(isVisible);
    if (reactTags) {
        return ["tag_input_commit", "nearest tag-style text input"];
    }

    const multiSelect = Array.from(
        container.querySelectorAll("[data-testid='trigger-multi-select']")
    ).find(isVisible);
    if (multiSelect) {
        return ["tag_input", "nearest multi-select combobox / checkbox-tree selector"];
    }

    const comboboxes = Array.from(
        container.querySelectorAll("button[role='combobox'], [role='combobox']")
    ).filter(isVisible);
    if (comboboxes.length > 0) {
        return ["combobox", "nearest single-select combobox"];
    }

    const input = Array.from(
        container.querySelectorAll("input:not([type='hidden']):not([type='radio']):not([type='checkbox'])")
    ).find(isVisible);
    if (input) {
        const type = (input.getAttribute("type") || "text").toLowerCase();
        if (type === "number") {
            return ["number", "nearest input[type='number']"];
        }
        return ["text", "nearest input[type='text']"];
    }

    return ["skip", "no editable control detected"];
}

function isLikelyLabel(element) {
    if (!isVisible(element)) {
        return false;
    }

    const text = cleanText(element.textContent);
    if (!text || text.length > 80) {
        return false;
    }

    const tagName = element.tagName.toLowerCase();
    if (["script", "style", "button", "input", "textarea", "option"].includes(tagName)) {
        return false;
    }

    const container = nearestFieldContainer(element);
    return container && container !== element && isVisible(container);
}

const labelCandidates = Array.from(
    document.querySelectorAll("label, span, div, p")
).filter(isLikelyLabel);

const seen = new Set();
const fields = [];

for (const labelElement of labelCandidates) {
    const labelText = cleanText(labelElement.textContent).replace(/\*$/, "").trim();
    if (!labelText || seen.has(labelText.toLowerCase())) {
        continue;
    }

    const container = nearestFieldContainer(labelElement);
    const [inputType, locatorTarget] = detectInputType(container);
    if (inputType === "skip" && !container.querySelector("input, textarea, button, [role='combobox']")) {
        continue;
    }

    seen.add(labelText.toLowerCase());
    const rect = labelElement.getBoundingClientRect();
    fields.push({
        label: labelText,
        required: /\*/.test(cleanText(labelElement.textContent)),
        input_type: inputType,
        locator_hint: `label text '${labelText}' -> ${locatorTarget}`,
        top: Math.round(rect.top + window.scrollY),
        left: Math.round(rect.left + window.scrollX)
    });
}

fields.sort((a, b) => (a.top - b.top) || (a.left - b.left));
return fields.map((field, index) => ({
    order: index + 1,
    label: field.label,
    required: field.required,
    input_type: field.input_type,
    locator_hint: field.locator_hint
}));
"""


def prompt_with_default(prompt: str, default: str) -> str:
    value = input(f"{prompt} [{default}]: ").strip()
    return value or default


def choose_input_type(current_input_type: str) -> str:
    print("Choose input type:")
    for key, value in INPUT_TYPE_CHOICES.items():
        marker = " (detected)" if value == current_input_type else ""
        print(f"{key}. {value}{marker}")

    while True:
        choice = input(f"Input type [{current_input_type}]: ").strip()
        if not choice:
            return current_input_type
        if choice in INPUT_TYPE_CHOICES:
            return INPUT_TYPE_CHOICES[choice]
        if choice in INPUT_TYPE_CHOICES.values():
            return choice
        print("Please choose a listed number or input type name.")


def review_detected_fields(fields: list[dict[str, Any]]) -> list[dict[str, Any]]:
    reviewed_fields: list[dict[str, Any]] = []
    next_order = 1

    for field in fields:
        print()
        print(
            f"{field['label']} detected as {field['input_type']}. "
            "Keep? [Y/n/e/s]"
        )
        print("Y = keep, n = drop, e = edit type/label, s = mark skip")
        choice = input("> ").strip().lower()

        if choice in {"n", "no"}:
            continue

        reviewed_field = dict(field)
        if choice in {"s", "skip"}:
            reviewed_field["input_type"] = "skip"
            reviewed_field["locator_hint"] = (
                f"skip for now; detected from label text '{reviewed_field['label']}'"
            )
        elif choice in {"e", "edit"}:
            new_label = input(f"Label [{reviewed_field['label']}]: ").strip()
            if new_label:
                reviewed_field["label"] = new_label
            reviewed_field["input_type"] = choose_input_type(str(reviewed_field["input_type"]))
            reviewed_field["required"] = (
                input(f"Required? [y/N] ({reviewed_field['required']}): ").strip().lower()
                in {"y", "yes"}
            )
            reviewed_field["locator_hint"] = (
                f"label text '{reviewed_field['label']}' -> reviewed as {reviewed_field['input_type']}"
            )

        reviewed_field["order"] = next_order
        next_order += 1
        reviewed_fields.append(reviewed_field)

    return reviewed_fields


def build_mapping_payload(
    surface: str,
    product_type: str,
    phase: str,
    fields: list[dict[str, Any]],
) -> dict[str, Any]:
    return {
        "surface": surface,
        "product_type": product_type,
        "phase": phase,
        "scope_note": (
            f"Generated by bot-scans-page.py for {surface} {product_type}. "
            "Review before using in main automation."
        ),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "locator_strategy": (
            "Prefer locating the field label text first, then target the nearest "
            "input, textarea, single-select combobox, multi-select combobox, or tag-style field."
        ),
        "fields": fields,
    }


def save_mapping(payload: dict[str, Any], output_path: Path | None = None) -> Path:
    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = "_".join(
            str(payload[key]).strip().lower().replace(" ", "-")
            for key in ("surface", "product_type", "phase")
        )
        output_path = OUTPUT_DIR / f"{timestamp}_{safe_name}.json"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(payload, indent=2, ensure_ascii=True) + "\n",
        encoding="utf-8",
    )
    return output_path


def main() -> None:
    print("--- BOT PAGE FIELD SCANNER ---")
    surface = prompt_with_default("Surface", "flipkart")
    product_type = prompt_with_default("Product type", "jeans")
    phase = prompt_with_default("Page/phase name", "Additional Description")
    output_path_value = input(
        "Output JSON path [blank = bot-page-scans timestamped file]: "
    ).strip().strip('"')
    output_path = Path(output_path_value) if output_path_value else None

    profile_name = prompt_for_profile()
    config = BotConfig(profile_name=profile_name)
    driver = build_firefox_driver(config)

    try:
        print()
        print("Firefox is open. Navigate to the exact Flipkart page/tab you want to scan.")
        input("Press Enter here when the page is ready...")
        fields = driver.execute_script(FIELD_SCAN_SCRIPT)
        if not fields:
            raise SystemExit("No fields were detected on the current page.")

        print(f"Detected {len(fields)} possible field(s).")
        reviewed_fields = review_detected_fields(fields)
        payload = build_mapping_payload(surface, product_type, phase, reviewed_fields)
        saved_path = save_mapping(payload, output_path)
        log_event("SCAN", f"Saved {len(reviewed_fields)} reviewed field(s) to: {saved_path}")
    except WebDriverException as exc:
        raise SystemExit(f"Could not scan the current browser page: {exc}") from exc
    finally:
        should_close = input("Close browser? [y/N]: ").strip().lower()
        if should_close in {"y", "yes"}:
            driver.quit()


if __name__ == "__main__":
    main()
