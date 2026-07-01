from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from selenium.common.exceptions import NoSuchWindowException, WebDriverException

from main import BotConfig, build_firefox_driver, log_event, prompt_for_profile


DEFAULT_SURFACE = "flipkart"
DEFAULT_PRODUCT_TYPE = "jeans"
DEFAULT_FLOW_NAME = "listing_creation"
DEFAULT_BRAND_NAME = "STARVIELLE"
OUTPUT_ROOT = Path("json_LC_creation")

FIELD_PAGE_STEPS = (
    {
        "order": 2,
        "step_id": "additional_description",
        "spec_file": "01_additional_description.json",
        "page_name": "Additional Description",
        "field_fill_mode": "additional_description",
        "data_source": {
            "type": "excel_row",
            "workbook_attr": "additional_description_excel",
            "worksheet": "Jeans Addl Desc Inputs",
            "match_by": {"kind": "$listing.kind", "size": "$listing.size"},
        },
        "source_snapshot": "Full-LC-AUTO/assets/Additional Description.htm",
        "locator_strategy": (
            "Prefer locating the field label text first, then target the nearest input, "
            "single-select combobox, multi-select combobox, textarea, or tag-style field "
            "inside the same attribute wrapper."
        ),
        "log_stage": "ADDL",
        "log_message": "Starting Additional Description field fill from Excel + JSON mapping...",
        "filled_checkpoint_label": "Additional Description fields filled",
        "verification": {
            "enabled": True,
            "mode": "tab_cycle",
            "cycle_label": "Additional Description page",
            "tab_sequence": ["Additional Description", "Product Description"],
            "success_conditions": [
                {"type": "toast", "value": "Changes saved!"},
                {"type": "tab_progress_nonzero"},
            ],
            "checkpoint_label": "Changes saved detected after Additional Description click cycle",
        },
    },
    {
        "order": 3,
        "step_id": "product_description",
        "spec_file": "02_product_description.json",
        "page_name": "Product Description",
        "field_fill_mode": "product_description",
        "data_source": {
            "type": "excel_row",
            "workbook_attr": "product_description_excel",
            "worksheet": "Jeans Product Inputs",
            "match_by": {"kind": "$listing.kind", "size": "$listing.size"},
        },
        "source_snapshot": "Full-LC-AUTO/product description page.htm",
        "locator_strategy": (
            "Prefer locating the field label text first, then target the nearest input, "
            "combobox, or tag-style field within the same attribute wrapper."
        ),
        "log_stage": "DESC",
        "log_message": "Starting Product Description field fill from Excel + JSON mapping...",
        "filled_checkpoint_label": "Product Description fields filled",
        "verification": {
            "enabled": True,
            "mode": "tab_cycle",
            "cycle_label": "Product Description page",
            "tab_sequence": ["Product Description", "Price, Stock and Shipping Information"],
            "success_conditions": [
                {"type": "toast", "value": "Changes saved!"},
                {"type": "tab_progress_nonzero"},
            ],
            "checkpoint_label": "Changes saved detected after Product Description click cycle",
        },
    },
    {
        "order": 4,
        "step_id": "price_stock_shipping",
        "spec_file": "03_price_stock_shipping.json",
        "page_name": "Price, Stock and Shipping Information",
        "field_fill_mode": "price_stock_shipping",
        "data_source": {
            "type": "excel_row",
            "workbook_attr": "price_stock_shipping_excel",
            "match_by": {"kind": "$listing.kind", "size": "$listing.size"},
        },
        "source_snapshot": "Full-LC-AUTO/snapshots/PHASE 1.html",
        "locator_strategy": (
            "Prefer locating the field label text first, then target the nearest input/combobox "
            "within the same attribute wrapper."
        ),
        "filled_checkpoint_label": "Price/Stock/Shipping fields filled",
        "snapshot_name": "PHASE 1.html",
        "save_to_context": {
            "price_stock_shipping.generated.Seller SKU ID": "Seller SKU ID"
        },
        "verification": {
            "enabled": True,
            "mode": "tab_cycle",
            "cycle_label": "SKU page",
            "tab_sequence": ["Price, Stock and Shipping Information", "Image addition"],
            "success_conditions": [
                {"type": "toast", "value": "Changes saved!"},
                {"type": "tab_progress_nonzero"},
            ],
            "checkpoint_label": "Changes saved detected after SKU-page click cycle",
        },
    },
)

IMAGE_STEP = {
    "order": 5,
    "step_id": "images",
    "spec_file": "04_images.json",
    "page_name": "Image addition",
    "handler": "image_upload_page",
    "checkpoint_label": "Images tab opened",
    "brand_name": DEFAULT_BRAND_NAME,
    "image_source": {
        "type": "selected_listing_image_directory",
        "brand_name_from": "$context.brand_name",
    },
    "selection_rules": {
        "folder_strategy": "choose_next_available_for_brand",
        "ordered_images": True,
        "max_images": 5,
    },
    "upload_targets": {
        "slot_ids": ["thumbnail_0", "thumbnail_1", "thumbnail_2", "thumbnail_3", "thumbnail_4"],
    },
    "filled_checkpoint_label": "Images uploaded",
    "verification": {
        "enabled": True,
        "mode": "tab_cycle",
        "cycle_label": "Images page",
        "tab_sequence": ["Image addition", "Variant addition"],
        "success_conditions": [{"type": "toast", "value": "Changes saved!"}],
        "checkpoint_label": "Changes saved detected after Images page click cycle",
    },
}

VARIANT_STEP = {
    "order": 6,
    "step_id": "variants",
    "spec_file": "05_variants.json",
    "page_name": "Variant addition",
    "handler": "variant_page",
    "checkpoint_label": "Variant tab opened",
    "data_source": {
        "type": "excel_row",
        "workbook_attr": "variants_excel",
        "worksheet": "Jeans Variant Inputs",
        "match_by": {"kind": "$listing.kind", "size": "$listing.size"},
    },
    "excel_config_attr": "variants_excel",
    "log_stage": "VARIANT",
    "log_message": "Starting Variant page creation loop from Excel mapping...",
    "filled_checkpoint_label": "Variant rows created",
    "source_sku_from": "$context.price_stock_shipping.generated.Seller SKU ID",
    "variant_creation": {
        "qualifier_column": "variant_qualifier",
        "sizes_column_prefix": "size_variant",
        "source_size_from": "$listing.size",
    },
    "copy_paste": {"enabled": True, "rewrite_target_skus": True},
}

FIELD_DISCOVERY_SCRIPT = r"""
function cleanText(value) {
    return (value || "").replace(/\s+/g, " ").trim();
}

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

function detectInputType(container) {
    const textArea = Array.from(container.querySelectorAll("textarea")).find(isVisible);
    if (textArea) {
        return ["text", "nearest textarea", textArea];
    }

    const reactTags = Array.from(container.querySelectorAll(".rti--container")).find(isVisible);
    if (reactTags) {
        return ["tag_input_commit", "nearest tag-style text input", reactTags];
    }

    const multiSelect = Array.from(
        container.querySelectorAll("[data-testid='trigger-multi-select']")
    ).find(isVisible);
    if (multiSelect) {
        return ["tag_input", "nearest multi-select combobox / checkbox-tree selector", multiSelect];
    }

    const combobox = Array.from(
        container.querySelectorAll("button[role='combobox'], [role='combobox']")
    ).find(isVisible);
    if (combobox) {
        return ["combobox", "nearest single-select combobox", combobox];
    }

    const input = Array.from(
        container.querySelectorAll("input:not([type='hidden']):not([type='radio']):not([type='checkbox'])")
    ).find(isVisible);
    if (input) {
        const type = (input.getAttribute("type") || "text").toLowerCase();
        if (type === "number") {
            return ["number", "nearest input[type='number']", input];
        }
        return ["text", "nearest input[type='text']", input];
    }

    return ["skip", "no editable control detected", null];
}

function detectInputTypeFromControl(controlElement) {
    if (!controlElement || !isVisible(controlElement)) {
        return ["skip", "no visible editable control", null];
    }

    const tagName = controlElement.tagName.toLowerCase();
    const testId = controlElement.getAttribute("data-testid") || "";
    const role = controlElement.getAttribute("role") || "";

    if (tagName === "textarea") {
        return ["text", "textarea control", controlElement];
    }

    if (testId === "trigger-single-select" || controlElement.id === "trigger-single-select") {
        return ["combobox", "single-select combobox", controlElement];
    }

    if (testId === "trigger-multi-select" || controlElement.id === "trigger-multi-select") {
        return ["tag_input", "multi-select combobox / checkbox-tree selector", controlElement];
    }

    if (controlElement.matches(".rti--container, .rti--container *")) {
        const container = controlElement.closest(".rti--container");
        return ["tag_input_commit", "tag-style text input", container || controlElement];
    }

    if (
        controlElement.matches("[data-testid='trigger-multi-select']") ||
        controlElement.closest("[data-testid='trigger-multi-select']")
    ) {
        const trigger = controlElement.closest("[data-testid='trigger-multi-select']") || controlElement;
        return ["tag_input", "multi-select combobox / checkbox-tree selector", trigger];
    }

    if (
        controlElement.matches("button[role='combobox'], [role='combobox']") ||
        controlElement.closest("button[role='combobox'], [role='combobox']")
    ) {
        const combo = controlElement.closest("button[role='combobox'], [role='combobox']") || controlElement;
        return ["combobox", "single-select combobox", combo];
    }

    if (tagName === "input") {
        if (controlElement.closest(".rti--container")) {
            const container = controlElement.closest(".rti--container");
            return ["tag_input_commit", "tag-style text input", container || controlElement];
        }

        const type = (controlElement.getAttribute("type") || "text").toLowerCase();
        if (["hidden", "radio", "checkbox"].includes(type)) {
            return ["skip", "unsupported input type", null];
        }
        if (type === "number") {
            return ["number", "input[type='number']", controlElement];
        }
        if (type === "search") {
            return ["text", "input[type='search']", controlElement];
        }
        return ["text", "input[type='text']", controlElement];
    }

    if (role === "textbox") {
        return ["text", "textbox role control", controlElement];
    }

    return ["skip", "unsupported control element", null];
}

function buildTabInfo() {
    const selected = document.querySelector(
        "button[role='tab'][aria-selected='true'], button[role='tab'][data-state='active']"
    );
    if (!selected) {
        return null;
    }

    const label = cleanText(selected.textContent);
    return {
        label: label,
        locator_candidates: [
            { type: "label_text", value: label },
            {
                type: "xpath",
                value: `//button[@role='tab' and .//span[contains(normalize-space(), ${JSON.stringify(label)})]]`
            }
        ],
        dom_context: {
            tag: selected.tagName.toLowerCase(),
            role: selected.getAttribute("role"),
            aria_selected: selected.getAttribute("aria-selected"),
            class_name: selected.className || ""
        }
    };
}

function buildElementPath(element) {
    if (!element || !(element instanceof Element)) {
        return "";
    }

    const parts = [];
    let current = element;
    while (current && current.nodeType === Node.ELEMENT_NODE && parts.length < 8) {
        let part = current.tagName.toLowerCase();
        const currentId = current.id || "";
        const hasUniqueId = currentId && document.querySelectorAll(`#${CSS.escape(currentId)}`).length === 1;
        if (hasUniqueId) {
            part += `#${currentId}`;
            parts.unshift(part);
            break;
        }

        const testId = current.getAttribute("data-testid");
        if (testId) {
            part += `[data-testid="${testId}"]`;
        } else {
            const parent = current.parentElement;
            if (parent) {
                const siblings = Array.from(parent.children).filter(
                    sibling => sibling.tagName === current.tagName
                );
                if (siblings.length > 1) {
                    part += `:nth-of-type(${siblings.indexOf(current) + 1})`;
                }
            }
        }

        parts.unshift(part);
        current = current.parentElement;
    }

    return parts.join(" > ");
}

function findFieldWrapper(controlElement) {
    let current = controlElement.parentElement || controlElement;
    for (let depth = 0; current && depth < 8; depth += 1) {
        if (!(current instanceof Element)) {
            return controlElement.parentElement || controlElement;
        }

        if (
            current === controlElement ||
            current.matches("button, input, textarea, [role='combobox'], [data-testid='trigger-single-select'], [data-testid='trigger-multi-select'], .rti--container")
        ) {
            current = current.parentElement;
            continue;
        }

        const text = cleanText(current.textContent);
        const hasEnoughContext = text.length > 0 && text.length <= 300;
        const containsPeerLabel = current.querySelector("label, span, div, p");
        if (hasEnoughContext && containsPeerLabel) {
            return current;
        }
        current = current.parentElement;
    }
    return controlElement.parentElement || controlElement;
}

function scoreLabelCandidate(element, controlElement) {
    const text = cleanText(element.textContent);
    const normalized = text.replace(/\*$/, "").trim();
    if (!normalized || normalized.length > 80) {
        return null;
    }
    if (["select", "search", "resize", "choose", "add"].includes(normalized.toLowerCase())) {
        return null;
    }
    if (element.querySelector("input, textarea, button[role='combobox'], [role='combobox'], [data-testid='trigger-multi-select'], .rti--container")) {
        return null;
    }

    const rect = element.getBoundingClientRect();
    const controlRect = controlElement.getBoundingClientRect();
    const verticalDistance = Math.abs(rect.top - controlRect.top);
    const horizontalDistance = Math.abs(rect.left - controlRect.left);
    const isLikelyLeftLabel = rect.right <= controlRect.left + 80;
    const isLikelyTopLabel = rect.bottom <= controlRect.top + 40;
    const overlapsControlRow = Math.abs(rect.top - controlRect.top) <= 28;
    let score = verticalDistance + horizontalDistance;
    if (isLikelyLeftLabel) {
        score -= 35;
    }
    if (isLikelyTopLabel) {
        score -= 28;
    }
    if (overlapsControlRow) {
        score -= 18;
    }

    return {
        text: normalized,
        rawText: text,
        score: score,
        element: element
    };
}

function collectLabelCandidates(root, controlElement) {
    if (!root || !(root instanceof Element)) {
        return [];
    }

    const scored = [];
    const labelElements = Array.from(root.querySelectorAll("label, span, div, p"));
    for (const element of labelElements) {
        if (!isVisible(element) || element === controlElement || element.contains(controlElement)) {
            continue;
        }
        const candidate = scoreLabelCandidate(element, controlElement);
        if (candidate) {
            scored.push(candidate);
        }
    }
    return scored;
}

function collectExternalLabelCandidates(wrapper, controlElement) {
    const candidates = [];
    const root = wrapper.parentElement;
    if (!root) {
        return candidates;
    }

    const wrapperRect = wrapper.getBoundingClientRect();
    const controlRect = controlElement.getBoundingClientRect();
    const nearbyNodes = Array.from(root.children).filter(isVisible);

    for (const node of nearbyNodes) {
        if (node === wrapper || node.contains(wrapper) || wrapper.contains(node)) {
            continue;
        }

        const nodeRect = node.getBoundingClientRect();
        const isNearRow = Math.abs(nodeRect.top - controlRect.top) <= 50;
        const isNearAbove = nodeRect.bottom <= wrapperRect.top + 20 && Math.abs(nodeRect.left - wrapperRect.left) <= 220;
        if (!isNearRow && !isNearAbove) {
            continue;
        }

        candidates.push(...collectLabelCandidates(node, controlElement));
        const selfCandidate = scoreLabelCandidate(node, controlElement);
        if (selfCandidate) {
            candidates.push(selfCandidate);
        }
    }

    return candidates;
}

function labelFromAttributes(controlElement) {
    const attributeCandidates = [
        controlElement.getAttribute("aria-label"),
        controlElement.getAttribute("placeholder"),
        controlElement.getAttribute("name")
    ];

    const labelledBy = controlElement.getAttribute("aria-labelledby");
    if (labelledBy) {
        for (const id of labelledBy.split(/\s+/)) {
            const labelElement = document.getElementById(id);
            if (labelElement) {
                attributeCandidates.unshift(cleanText(labelElement.textContent));
            }
        }
    }

    const controlId = controlElement.getAttribute("id");
    if (controlId) {
        const labelElement = document.querySelector(`label[for="${controlId}"]`);
        if (labelElement) {
            attributeCandidates.unshift(cleanText(labelElement.textContent));
        }
    }

    for (const candidate of attributeCandidates) {
        const text = cleanText(candidate || "");
        if (text && !["select", "search", "resize", "choose", "add"].includes(text.toLowerCase())) {
            return text;
        }
    }

    return "";
}

function findBestLabel(wrapper, controlElement) {
    const candidates = collectLabelCandidates(wrapper, controlElement);
    candidates.push(...collectExternalLabelCandidates(wrapper, controlElement));
    const ancestorCandidates = [];
    let current = wrapper.parentElement;
    for (let depth = 0; current && depth < 4; depth += 1) {
        ancestorCandidates.push(...collectLabelCandidates(current, controlElement));
        current = current.parentElement;
    }

    candidates.push(...ancestorCandidates);
    if (candidates.length === 0) {
        return null;
    }

    const uniqueByTextAndPath = new Map();
    for (const candidate of candidates) {
        const key = `${candidate.text}::${buildElementPath(candidate.element)}`;
        if (!uniqueByTextAndPath.has(key) || uniqueByTextAndPath.get(key).score > candidate.score) {
            uniqueByTextAndPath.set(key, candidate);
        }
    }

    const deduped = Array.from(uniqueByTextAndPath.values());
    deduped.sort((a, b) => a.score - b.score);
    return deduped.length > 0 ? deduped[0] : null;
}

function extractCandidateFields() {
    const candidates = [];
    const seenKeys = new Set();
    const controlCandidates = Array.from(document.querySelectorAll(
        "textarea, input:not([type='hidden']):not([type='radio']):not([type='checkbox']), " +
        "button[role='combobox'], [role='combobox'], " +
        "[data-testid='trigger-single-select'], [data-testid='trigger-multi-select'], .rti--container"
    )).filter(isVisible);

    for (const rawControl of controlCandidates) {
        const [inputType, locatorTarget, controlElement] = detectInputTypeFromControl(rawControl);
        if (inputType === "skip" || !controlElement) {
            continue;
        }

        const controlPath = buildElementPath(controlElement);
        if (!controlPath || seenKeys.has(controlPath)) {
            continue;
        }

        const wrapper = findFieldWrapper(controlElement);
        const labelCandidate = findBestLabel(wrapper, controlElement);

        let labelText = labelCandidate ? labelCandidate.text : "";
        let required = labelCandidate ? /\*/.test(labelCandidate.rawText) : false;

        if (!labelText) {
            labelText = labelFromAttributes(controlElement);
        }

        if (!labelText) {
            continue;
        }

        const rect = wrapper.getBoundingClientRect();
        seenKeys.add(controlPath);
        candidates.push({
            order: 0,
            label: labelText,
            required: required,
            input_type: inputType,
            locator_hint: `label text '${labelText}' -> ${locatorTarget}`,
            locator_candidates: [{ type: "label_text", value: labelText }],
            dom_context: {
                wrapper_path: buildElementPath(wrapper),
                control_path: controlPath,
                control_tag: controlElement.tagName.toLowerCase(),
                control_role: controlElement.getAttribute("role") || "",
                control_type: controlElement.getAttribute("type") || "",
                control_testid: controlElement.getAttribute("data-testid") || ""
            },
            top: Math.round(rect.top + window.scrollY),
            left: Math.round(rect.left + window.scrollX)
        });
    }

    candidates.sort((a, b) => (a.top - b.top) || (a.left - b.left));
    candidates.forEach((candidate, index) => {
        candidate.order = index + 1;
        delete candidate.top;
        delete candidate.left;
    });

    return {
        active_tab: buildTabInfo(),
        candidates: candidates
    };
}

return extractCandidateFields();
"""

FOCUS_FIELD_SCRIPT = r"""
const controlPath = arguments[0];

function clearHighlights() {
    document.querySelectorAll("[data-json-flow-recorder-highlight='1']").forEach(element => {
        element.style.outline = element.dataset.jsonFlowRecorderPrevOutline || "";
        element.style.outlineOffset = element.dataset.jsonFlowRecorderPrevOutlineOffset || "";
        delete element.dataset.jsonFlowRecorderPrevOutline;
        delete element.dataset.jsonFlowRecorderPrevOutlineOffset;
        delete element.dataset.jsonFlowRecorderHighlight;
    });
}

function findByPath(path) {
    if (!path) {
        return null;
    }
    if (path.includes("#")) {
        const idPart = path.split("#")[1].split(/[ >]/)[0];
        const direct = document.getElementById(idPart);
        if (direct) {
            return direct;
        }
    }
    const selector = path
        .replace(/ > /g, " > ")
        .replace(/\[data-testid="([^"]+)"\]/g, '[data-testid="$1"]');
    try {
        return document.querySelector(selector);
    } catch (error) {
        return null;
    }
}

clearHighlights();
const target = findByPath(controlPath);
if (!target) {
    return false;
}

target.scrollIntoView({block: "center", inline: "nearest"});
target.dataset.jsonFlowRecorderPrevOutline = target.style.outline || "";
target.dataset.jsonFlowRecorderPrevOutlineOffset = target.style.outlineOffset || "";
target.dataset.jsonFlowRecorderHighlight = "1";
target.style.outline = "3px solid #ff7a00";
target.style.outlineOffset = "2px";
if (typeof target.focus === "function") {
    target.focus({preventScroll: true});
}
return true;
"""


def prompt_with_default(prompt: str, default: str) -> str:
    value = input(f"{prompt} [{default}]: ").strip()
    return value or default


def save_json(payload: dict[str, Any], output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")
    return output_path


def ensure_browser_context_available(driver, action_label: str) -> None:
    try:
        handles = driver.window_handles
    except NoSuchWindowException as exc:
        raise SystemExit(
            "The Selenium-controlled Firefox window was closed before the recorder could continue. "
            f"Keep the recorder-opened Firefox window alive while preparing the page ({action_label})."
        ) from exc
    except WebDriverException as exc:
        if "Browsing context has been discarded" in str(exc):
            raise SystemExit(
                "Firefox discarded the Selenium browsing context before recording could continue. "
                "Use the Firefox window opened by the recorder itself, and avoid closing/replacing that window."
            ) from exc
        raise

    if not handles:
        raise SystemExit(
            "No Firefox window is attached to the recorder right now. "
            "Keep the recorder-opened Firefox window alive while preparing the page."
        )


def discover_page_candidates(driver) -> dict[str, Any]:
    ensure_browser_context_available(driver, "discover page candidates")
    payload = driver.execute_script(FIELD_DISCOVERY_SCRIPT)
    if not isinstance(payload, dict):
        raise SystemExit("Could not inspect the current Firefox page.")
    return payload


def focus_candidate(driver, candidate: dict[str, Any]) -> bool:
    ensure_browser_context_available(driver, "focus candidate")
    control_path = candidate.get("dom_context", {}).get("control_path", "")
    return bool(driver.execute_script(FOCUS_FIELD_SCRIPT, control_path))


def prompt_for_tab_capture(expected_label: str) -> None:
    print()
    print(f"Click the '{expected_label}' tab in Firefox, then press Enter here.")
    input("Press Enter after the tab is active... ")


def warn_if_unexpected_tab(active_tab: dict[str, Any] | None, expected_label: str) -> None:
    if not active_tab:
        print(f"Warning: could not detect an active tab. Expected '{expected_label}'.")
        return
    actual_label = str(active_tab.get("label", "")).strip()
    if actual_label != expected_label:
        print(f"Warning: expected active tab '{expected_label}', but detected '{actual_label}'.")


def normalize_tab_payload(tab_payload: dict[str, Any], fallback_label: str) -> dict[str, Any]:
    label = str(tab_payload.get("label") or fallback_label).strip()
    locator_candidates = tab_payload.get("locator_candidates")
    if not isinstance(locator_candidates, list) or not locator_candidates:
        locator_candidates = [
            {"type": "label_text", "value": label},
            {
                "type": "xpath",
                "value": f"//button[@role='tab' and .//span[contains(normalize-space(), {json.dumps(label)})]]",
            },
        ]
    return {
        "label": label,
        "locator_candidates": locator_candidates,
        "dom_context": tab_payload.get("dom_context", {}),
    }


def build_flow_manifest(surface: str, product_type: str) -> dict[str, Any]:
    return {
        "surface": surface,
        "product_type": product_type,
        "flow_name": DEFAULT_FLOW_NAME,
        "version": 1,
        "context": {"brand_name": DEFAULT_BRAND_NAME},
        "steps": [
            {"order": 1, "step_id": "open_listing", "handler": "navigation_step", "spec_file": "00_open_listing.json"},
            {"order": 2, "step_id": "additional_description", "handler": "field_fill_page", "spec_file": "01_additional_description.json"},
            {"order": 3, "step_id": "product_description", "handler": "field_fill_page", "spec_file": "02_product_description.json"},
            {"order": 4, "step_id": "price_stock_shipping", "handler": "field_fill_page", "spec_file": "03_price_stock_shipping.json"},
            {"order": 5, "step_id": "images", "handler": "image_upload_page", "spec_file": "04_images.json"},
            {"order": 6, "step_id": "variants", "handler": "variant_page", "spec_file": "05_variants.json"},
            {"order": 7, "step_id": "save_and_exit", "handler": "navigation_step", "spec_file": "06_save_and_exit.json"},
        ],
    }


def build_open_listing_spec() -> dict[str, Any]:
    return {
        "order": 1,
        "step_name": "open_listing",
        "handler": "navigation_step",
        "page_name": "Open listing flow",
        "actions": [
            {"type": "open_listing_page", "url": "$config.listing_url", "checkpoint_label": "Listing page opened"},
            {"type": "dismiss_optional_ad_popup", "timeout_seconds": 5, "checkpoint_label": "Optional popup handling complete"},
            {"type": "fill_brand_name", "brand_name": "$context.brand_name", "checkpoint_label": "Brand entered"},
            {"type": "click_create_new_listing", "checkpoint_label": "Create new listing clicked"},
            {"type": "click_optional_continue", "timeout_seconds": 5, "checkpoint_label": "Optional continue handling complete"},
        ],
    }


def build_save_and_exit_spec() -> dict[str, Any]:
    return {
        "order": 7,
        "step_name": "save_and_exit",
        "handler": "navigation_step",
        "page_name": "Save listing and exit",
        "actions": [
            {"type": "click_save_and_go_back", "checkpoint_label": "Save & Go Back clicked"}
        ],
    }


def build_field_page_spec(
    step_template: dict[str, Any],
    tab_payload: dict[str, Any],
    fields: list[dict[str, Any]],
) -> dict[str, Any]:
    checkpoint_label = (
        "Selling info tab opened"
        if step_template["step_id"] == "price_stock_shipping"
        else f"{step_template['page_name']} tab opened"
    )
    return {
        "order": step_template["order"],
        "step_name": step_template["step_id"],
        "handler": "field_fill_page",
        "field_fill_mode": step_template["field_fill_mode"],
        "page_name": step_template["page_name"],
        "tab": normalize_tab_payload(tab_payload, step_template["page_name"]),
        "checkpoint_label": checkpoint_label,
        "data_source": step_template["data_source"],
        "source_snapshot": step_template["source_snapshot"],
        "locator_strategy": step_template["locator_strategy"],
        "log_stage": step_template.get("log_stage"),
        "log_message": step_template.get("log_message"),
        "filled_checkpoint_label": step_template["filled_checkpoint_label"],
        "snapshot_name": step_template.get("snapshot_name"),
        "save_to_context": step_template.get("save_to_context"),
        "verification": step_template["verification"],
        "fields": fields,
        "recording_metadata": {
            "recorded_via": "record-json-flow.py",
            "capture_mode": "bot-focused y-n walkthrough",
        },
    }


def build_image_page_spec(tab_payload: dict[str, Any]) -> dict[str, Any]:
    return {
        "order": IMAGE_STEP["order"],
        "step_name": IMAGE_STEP["step_id"],
        "handler": IMAGE_STEP["handler"],
        "page_name": IMAGE_STEP["page_name"],
        "tab": normalize_tab_payload(tab_payload, IMAGE_STEP["page_name"]),
        "checkpoint_label": IMAGE_STEP["checkpoint_label"],
        "brand_name": IMAGE_STEP["brand_name"],
        "image_source": IMAGE_STEP["image_source"],
        "selection_rules": IMAGE_STEP["selection_rules"],
        "upload_targets": IMAGE_STEP["upload_targets"],
        "filled_checkpoint_label": IMAGE_STEP["filled_checkpoint_label"],
        "verification": IMAGE_STEP["verification"],
        "recording_metadata": {
            "recorded_via": "record-json-flow.py",
            "capture_mode": "tab-anchor-only",
            "note": "Image page behavior remains Python-driven.",
        },
    }


def build_variant_page_spec(tab_payload: dict[str, Any]) -> dict[str, Any]:
    return {
        "order": VARIANT_STEP["order"],
        "step_name": VARIANT_STEP["step_id"],
        "handler": VARIANT_STEP["handler"],
        "page_name": VARIANT_STEP["page_name"],
        "tab": normalize_tab_payload(tab_payload, VARIANT_STEP["page_name"]),
        "checkpoint_label": VARIANT_STEP["checkpoint_label"],
        "data_source": VARIANT_STEP["data_source"],
        "excel_config_attr": VARIANT_STEP["excel_config_attr"],
        "log_stage": VARIANT_STEP["log_stage"],
        "log_message": VARIANT_STEP["log_message"],
        "filled_checkpoint_label": VARIANT_STEP["filled_checkpoint_label"],
        "source_sku_from": VARIANT_STEP["source_sku_from"],
        "variant_creation": VARIANT_STEP["variant_creation"],
        "copy_paste": VARIANT_STEP["copy_paste"],
        "recording_metadata": {
            "recorded_via": "record-json-flow.py",
            "capture_mode": "tab-anchor-only",
            "note": "Variant behavior remains Python-driven.",
        },
    }


def walk_candidates(driver, page_name: str, candidates: list[dict[str, Any]]) -> list[dict[str, Any]]:
    kept_fields: list[dict[str, Any]] = []
    print()
    print(f"Reviewing detected editable controls on '{page_name}'.")
    print("Press `y` to keep, `n` to skip, `q` to stop this page.")

    for index, candidate in enumerate(candidates, start=1):
        focused = focus_candidate(driver, candidate)
        print()
        print(f"{index}/{len(candidates)}")
        print(f"Label: {candidate['label']}")
        print(f"Detected type: {candidate['input_type']}")
        print(f"Locator hint: {candidate['locator_hint']}")
        if not focused:
            print("Warning: could not visually focus/highlight this control in Firefox.")

        while True:
            choice = input("[y/n/q]: ").strip().lower()
            if choice in {"y", "n", "q"}:
                break
            print("Use `y` to keep, `n` to skip, or `q` to stop this page.")

        if choice == "q":
            break
        if choice == "y":
            kept_field = dict(candidate)
            kept_field["order"] = len(kept_fields) + 1
            kept_fields.append(kept_field)

    return kept_fields


def main() -> None:
    print("--- JSON FLOW RECORDER ---")
    surface = prompt_with_default("Surface", DEFAULT_SURFACE)
    product_type = prompt_with_default("Product type", DEFAULT_PRODUCT_TYPE)
    flow_folder_name = prompt_with_default("Flow folder", f"{product_type}_{surface}")
    output_directory = OUTPUT_ROOT / flow_folder_name

    profile_name = prompt_for_profile()
    config = BotConfig(profile_name=profile_name)
    driver = build_firefox_driver(config)

    try:
        print()
        print("Firefox is open.")
        print("Prepare the listing flow manually until the page tabs are available.")
        input("Press Enter here once the seller UI is ready for recording... ")

        save_json(build_flow_manifest(surface, product_type), output_directory / "flow.json")
        save_json(build_open_listing_spec(), output_directory / "00_open_listing.json")
        save_json(build_save_and_exit_spec(), output_directory / "06_save_and_exit.json")

        for step_template in FIELD_PAGE_STEPS:
            prompt_for_tab_capture(step_template["page_name"])
            discovery = discover_page_candidates(driver)
            active_tab = discovery.get("active_tab") if isinstance(discovery, dict) else None
            warn_if_unexpected_tab(active_tab if isinstance(active_tab, dict) else None, step_template["page_name"])
            candidates = discovery.get("candidates", [])
            if not isinstance(candidates, list) or not candidates:
                raise SystemExit(f"No editable controls were detected on {step_template['page_name']}.")
            print(f"Detected {len(candidates)} candidate control(s) on {step_template['page_name']}.")
            kept_fields = walk_candidates(driver, step_template["page_name"], candidates)
            if not kept_fields:
                raise SystemExit(f"No fields were kept for {step_template['page_name']}.")
            page_payload = build_field_page_spec(
                step_template,
                active_tab if isinstance(active_tab, dict) else {},
                kept_fields,
            )
            saved_path = save_json(page_payload, output_directory / step_template["spec_file"])
            log_event("RECORDER", f"Saved {step_template['page_name']} spec to: {saved_path}")

        prompt_for_tab_capture(IMAGE_STEP["page_name"])
        image_discovery = discover_page_candidates(driver)
        image_tab = image_discovery.get("active_tab") if isinstance(image_discovery, dict) else None
        warn_if_unexpected_tab(image_tab if isinstance(image_tab, dict) else None, IMAGE_STEP["page_name"])
        image_path = save_json(
            build_image_page_spec(image_tab if isinstance(image_tab, dict) else {}),
            output_directory / IMAGE_STEP["spec_file"],
        )
        log_event("RECORDER", f"Saved Image page spec to: {image_path}")

        prompt_for_tab_capture(VARIANT_STEP["page_name"])
        variant_discovery = discover_page_candidates(driver)
        variant_tab = variant_discovery.get("active_tab") if isinstance(variant_discovery, dict) else None
        warn_if_unexpected_tab(variant_tab if isinstance(variant_tab, dict) else None, VARIANT_STEP["page_name"])
        variant_path = save_json(
            build_variant_page_spec(variant_tab if isinstance(variant_tab, dict) else {}),
            output_directory / VARIANT_STEP["spec_file"],
        )
        log_event("RECORDER", f"Saved Variant page spec to: {variant_path}")

        print()
        print(f"Flow JSON recorder finished. Files saved under: {output_directory}")
    except KeyboardInterrupt:
        raise SystemExit("Recorder stopped by user.")
    except WebDriverException as exc:
        raise SystemExit(f"Could not inspect the current Firefox page: {exc}") from exc
    finally:
        try:
            should_close = input("Close browser? [y/N]: ").strip().lower()
        except KeyboardInterrupt:
            should_close = "n"
        if should_close in {"y", "yes"}:
            driver.quit()


if __name__ == "__main__":
    main()
