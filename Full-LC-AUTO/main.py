from __future__ import annotations

import json
import os
import random
import string
import threading
from datetime import datetime
from dataclasses import dataclass
from pathlib import Path
from time import sleep

import msvcrt
from openpyxl import load_workbook
import pyautogui
from pynput.mouse import Button as PynputButton
from pynput.mouse import Controller as MouseController

from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


DEFAULT_LISTING_URL = (
    "https://seller.flipkart.com/index.html#dashboard/addListings/single"
    "?vertical=jean&vid=667"
)
DEFAULT_BRAND_NAME = "STARVIELLE"
DEFAULT_IMAGE_DIRECTORY = Path(
    r"C:\work-mom\HOSERY\SHORTS\CHATGPT\Lead_Permutations_Output"
)
DEFAULT_SNAPSHOT_DIRECTORY = Path(r"C:\work-mom\Code-Tools\Full-LC-AUTO\snapshots")
DEFAULT_PRICE_STOCK_SHIPPING_EXCEL = Path(
    r"C:\work-mom\Code-Tools\Full-LC-AUTO\data inputs\Price-Stock-Shipping-inputs.xlsx"
)
DEFAULT_PRICE_STOCK_SHIPPING_JSON = Path(
    r"C:\work-mom\Code-Tools\Full-LC-AUTO\assets\Price-Stock-Shipping-inputs.json"
)
DEFAULT_PRODUCT_DESCRIPTION_EXCEL = Path(
    r"C:\work-mom\Code-Tools\Full-LC-AUTO\data inputs\Product-Description-inputs-Shorts.xlsx"
)
DEFAULT_PRODUCT_DESCRIPTION_JSON = Path(
    r"C:\work-mom\Code-Tools\Full-LC-AUTO\assets\Product-Description-inputs-Shorts.json"
)
PHASE_ONE_SNAPSHOT_NAME = "PHASE 1.html"
DEFAULT_TARGET_KIND = "Shorts"
DEFAULT_TARGET_SIZE = "XL"
IMAGE_SLOT_IDS = [
    "thumbnail_0",
    "thumbnail_1",
    "thumbnail_2",
    "thumbnail_3",
    "thumbnail_4",
]
PRODUCT_DESCRIPTION_TAB_CLICK_POINT = (770, 380)
SKU_PAGE_ALTERNATE_CLICK_POINT = (518, 379)
BRAND_CODE_MAP = {
    "STAR": "STARVIELLE",
    "GENZ": "GENZ VANE",
    "IND": "INDIVANE",
    "FADE": "FADEVIELLE",
    "FLEE": "FLEECRANE",
}
BRAND_NAME_TO_CODE = {
    normalize_name: code for code, normalize_name in (
        (code, " ".join(name.strip().upper().split())) for code, name in BRAND_CODE_MAP.items()
    )
}
FIREFOX_PROFILES = {
    "seema": Path(r"C:\Users\ESHAAN\Documents\Firefox-Profiles\ekyb3fej.Seema"),
    "prabhu": Path(r"C:\Users\ESHAAN\Documents\Firefox-Profiles\7kkhlz7p.prabhu-bt"),
}
PROFILE_ALIASES = {
    "s": "seema",
    "seema": "seema",
    "p": "prabhu",
    "prabhu": "prabhu",
}


def log_event(stage: str, message: str) -> None:
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] [{stage}] {message}")


@dataclass(slots=True)
class BotConfig:
    listing_url: str = DEFAULT_LISTING_URL
    image_directory: Path = Path(
        os.getenv("FLIPKART_IMAGE_DIR", str(DEFAULT_IMAGE_DIRECTORY))
    ).expanduser()
    snapshot_directory: Path = Path(
        os.getenv("FLIPKART_SNAPSHOT_DIR", str(DEFAULT_SNAPSHOT_DIRECTORY))
    ).expanduser()
    price_stock_shipping_excel: Path = Path(
        os.getenv("PRICE_STOCK_SHIPPING_EXCEL", str(DEFAULT_PRICE_STOCK_SHIPPING_EXCEL))
    ).expanduser()
    price_stock_shipping_json: Path = Path(
        os.getenv("PRICE_STOCK_SHIPPING_JSON", str(DEFAULT_PRICE_STOCK_SHIPPING_JSON))
    ).expanduser()
    product_description_excel: Path = Path(
        os.getenv("PRODUCT_DESCRIPTION_EXCEL", str(DEFAULT_PRODUCT_DESCRIPTION_EXCEL))
    ).expanduser()
    product_description_json: Path = Path(
        os.getenv("PRODUCT_DESCRIPTION_JSON", str(DEFAULT_PRODUCT_DESCRIPTION_JSON))
    ).expanduser()
    data_directory: Path = Path(os.getenv("FLIPKART_DATA_DIR", "")).expanduser()
    firefox_binary: str | None = os.getenv("FIREFOX_BINARY")
    geckodriver_path: str | None = os.getenv("GECKODRIVER_PATH")
    profile_name: str = "seema"
    headless: bool = os.getenv("HEADLESS", "0") == "1"

    @property
    def firefox_profile_path(self) -> Path:
        if self.profile_name not in FIREFOX_PROFILES:
            available_profiles = ", ".join(sorted(FIREFOX_PROFILES))
            raise ValueError(
                f"Unknown Firefox profile '{self.profile_name}'. "
                f"Choose one of: {available_profiles}."
            )
        return resolve_profile_path(self.profile_name)


@dataclass(slots=True)
class ImageFolder:
    folder_path: Path
    folder_number: int
    exhausted_brands: set[str]
    image_paths: list[Path]


@dataclass(slots=True)
class ProductInputRow:
    kind: str
    size: str
    values: dict[str, str]


@dataclass(slots=True)
class FieldDefinition:
    order: int
    label: str
    required: bool
    input_type: str
    locator_hint: str


@dataclass(slots=True)
class FillResult:
    generated_values: dict[str, str]
    skipped_fields: set[str]


class PauseController:
    def __init__(self) -> None:
        self.pause_requested = False
        self._stop_event = threading.Event()
        self._listener_thread: threading.Thread | None = None

    def start(self) -> None:
        if self._listener_thread is not None:
            return

        self._listener_thread = threading.Thread(target=self._listen_for_spacebar, daemon=True)
        self._listener_thread.start()

    def stop(self) -> None:
        self._stop_event.set()

    def pause_if_requested(
        self,
        stage_name: str,
        driver: webdriver.Firefox,
        config: BotConfig,
    ) -> None:
        if not self.pause_requested:
            return

        self.pause_requested = False
        snapshot_path = save_html_snapshot(driver, config.snapshot_directory, stage_name)
        log_event("PAUSE", f"Saved HTML snapshot: {snapshot_path}")
        input(f"[PAUSED] {stage_name}. Press Enter to continue...")

    def _listen_for_spacebar(self) -> None:
        while not self._stop_event.is_set():
            if msvcrt.kbhit():
                pressed_key = msvcrt.getwch()
                if pressed_key == " ":
                    self.pause_requested = True
                    log_event("PAUSE", "Pause requested. The bot will pause at the next safe step.")
            sleep(0.1)


def checkpoint_pause(
    pause_controller: PauseController,
    stage_name: str,
    driver: webdriver.Firefox,
    config: BotConfig,
) -> None:
    pause_controller.pause_if_requested(stage_name, driver, config)


def resolve_profile_name(selected_value: str) -> str:
    normalized_value = selected_value.strip().lower()
    if normalized_value in PROFILE_ALIASES:
        return PROFILE_ALIASES[normalized_value]

    available_profiles = ", ".join(sorted(FIREFOX_PROFILES))
    raise ValueError(
        f"Unknown Firefox profile '{selected_value}'. Choose one of: {available_profiles}, s, p."
    )


def normalize_brand_name(brand_name: str) -> str:
    return " ".join(brand_name.strip().upper().split())


def normalize_field_label(label: str) -> str:
    cleaned = label.replace("*", " ")
    return " ".join(cleaned.strip().lower().split())


def parse_folder_number(folder_name: str) -> int:
    first_token = folder_name.split("-", maxsplit=1)[0].strip()
    if not first_token.isdigit():
        raise ValueError(f"Folder name must start with a number: {folder_name}")
    return int(first_token)


def parse_exhausted_brands(folder_name: str) -> set[str]:
    tokens = [part.strip().upper() for part in folder_name.split("-")[1:] if part.strip()]
    exhausted_brands: set[str] = set()

    for token in tokens:
        if token not in BRAND_CODE_MAP:
            raise ValueError(f"Unknown brand code '{token}' in folder '{folder_name}'")
        exhausted_brands.add(BRAND_CODE_MAP[token])

    return exhausted_brands


def collect_ordered_images(folder_path: Path) -> list[Path]:
    image_candidates = [path for path in folder_path.iterdir() if path.is_file()]

    def image_sort_key(path: Path) -> tuple[int, str]:
        stem = path.stem.strip()
        if stem.isdigit():
            return (int(stem), path.name.lower())
        return (10**9, path.name.lower())

    return sorted(image_candidates, key=image_sort_key)


def load_image_folders(image_root: Path) -> list[ImageFolder]:
    if not image_root.exists():
        raise ValueError(f"Image directory does not exist: {image_root}")
    if not image_root.is_dir():
        raise ValueError(f"Image directory is not a folder: {image_root}")

    image_folders: list[ImageFolder] = []
    for folder_path in image_root.iterdir():
        if not folder_path.is_dir():
            continue

        image_folders.append(
            ImageFolder(
                folder_path=folder_path,
                folder_number=parse_folder_number(folder_path.name),
                exhausted_brands=parse_exhausted_brands(folder_path.name),
                image_paths=collect_ordered_images(folder_path),
            )
        )

    return sorted(image_folders, key=lambda folder: folder.folder_number)


def choose_image_folder_for_brand(image_root: Path, brand_name: str) -> ImageFolder | None:
    normalized_brand = normalize_brand_name(brand_name)

    for image_folder in load_image_folders(image_root):
        if normalized_brand not in image_folder.exhausted_brands:
            return image_folder

    return None


def get_brand_code(brand_name: str) -> str:
    normalized_brand = normalize_brand_name(brand_name)
    if normalized_brand not in BRAND_NAME_TO_CODE:
        raise ValueError(f"No short code configured for brand: {brand_name}")
    return BRAND_NAME_TO_CODE[normalized_brand]


def build_exhausted_folder_name(image_folder: ImageFolder, brand_name: str) -> str:
    folder_codes = {
        get_brand_code(exhausted_brand) for exhausted_brand in image_folder.exhausted_brands
    }
    folder_codes.add(get_brand_code(brand_name))
    sorted_codes = sorted(folder_codes)
    return "-".join([str(image_folder.folder_number), *sorted_codes])


def mark_image_folder_exhausted(image_folder: ImageFolder, brand_name: str) -> Path:
    new_folder_name = build_exhausted_folder_name(image_folder, brand_name)
    new_folder_path = image_folder.folder_path.with_name(new_folder_name)

    if new_folder_path == image_folder.folder_path:
        log_event(
            "IMAGES",
            f"Image folder already marked for {brand_name}: {image_folder.folder_path.name}",
        )
        return image_folder.folder_path

    if new_folder_path.exists():
        raise ValueError(f"Cannot rename folder because target already exists: {new_folder_path}")

    image_folder.folder_path.rename(new_folder_path)
    image_folder.folder_path = new_folder_path
    image_folder.exhausted_brands.add(normalize_brand_name(brand_name))
    log_event("IMAGES", f"Renamed image folder to: {new_folder_path.name}")
    return new_folder_path


def resolve_profile_path(profile_name: str) -> Path:
    configured_path = FIREFOX_PROFILES[profile_name]

    if configured_path.exists():
        return configured_path

    log_event(
        "PROFILE",
        f"Saved path for profile '{profile_name}' was not found: {configured_path}",
    )
    manual_path = input(
        f"Paste the actual Firefox profile directory for '{profile_name}': "
    ).strip().strip('"')

    if not manual_path:
        raise ValueError(f"No profile directory was provided for '{profile_name}'.")

    resolved_path = Path(manual_path).expanduser()
    if not resolved_path.exists():
        raise ValueError(f"Profile directory does not exist: {resolved_path}")

    FIREFOX_PROFILES[profile_name] = resolved_path
    return resolved_path


def prompt_for_profile() -> str:
    env_profile = os.getenv("FIREFOX_PROFILE", "prabhu")
    prompt = f"Choose Firefox profile - seema (s) or prabhu (p) [{env_profile}]: "
    selected_value = input(prompt).strip()
    return resolve_profile_name(selected_value or env_profile)


def build_firefox_driver(config: BotConfig) -> webdriver.Firefox:
    options = FirefoxOptions()

    if config.firefox_binary:
        options.binary_location = config.firefox_binary

    if config.headless:
        options.add_argument("-headless")

    # Use the existing profile as a template so Selenium runs against a copied session
    # instead of writing WebDriver preferences into the original profile directory.
    options.profile = FirefoxProfile(str(config.firefox_profile_path))

    if config.geckodriver_path:
        service = FirefoxService(executable_path=config.geckodriver_path)
        return webdriver.Firefox(service=service, options=options)

    # Selenium Manager can resolve the driver automatically in recent Selenium versions.
    return webdriver.Firefox(options=options)


def open_listing_page(driver: webdriver.Firefox, url: str) -> None:
    driver.maximize_window()
    driver.get(url)


def save_html_snapshot(driver: webdriver.Firefox, snapshot_directory: Path, stage_name: str) -> Path:
    snapshot_directory.mkdir(parents=True, exist_ok=True)
    safe_stage_name = "".join(
        character if character.isalnum() else "_" for character in stage_name.strip().lower()
    ).strip("_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    snapshot_path = snapshot_directory / f"{timestamp}_{safe_stage_name}.html"
    snapshot_path.write_text(driver.page_source, encoding="utf-8")
    return snapshot_path


def save_named_html_snapshot(
    driver: webdriver.Firefox,
    snapshot_directory: Path,
    file_name: str,
) -> Path:
    snapshot_directory.mkdir(parents=True, exist_ok=True)
    snapshot_path = snapshot_directory / file_name
    snapshot_path.write_text(driver.page_source, encoding="utf-8")
    return snapshot_path


def load_product_input_row(
    workbook_path: Path,
    target_kind: str,
    target_size: str,
    worksheet_name: str | None = None,
) -> ProductInputRow:
    if not workbook_path.exists():
        raise ValueError(f"Excel file was not found: {workbook_path}")

    workbook = load_workbook(workbook_path, data_only=True)
    if worksheet_name and worksheet_name in workbook.sheetnames:
        worksheet = workbook[worksheet_name]
    else:
        preferred_sheet = next(
            (sheet_name for sheet_name in workbook.sheetnames if "product inputs" in sheet_name.lower()),
            workbook.sheetnames[0],
        )
        worksheet = workbook[preferred_sheet]
    headers = [worksheet.cell(1, column).value for column in range(1, worksheet.max_column + 1)]

    normalized_target_kind = target_kind.strip().lower()
    normalized_target_size = target_size.strip().lower()

    for row_index in range(2, worksheet.max_row + 1):
        row_values = {
            str(headers[column_index - 1]): worksheet.cell(row_index, column_index).value
            for column_index in range(1, worksheet.max_column + 1)
            if headers[column_index - 1]
        }
        row_kind = str(row_values.get("kind", "")).strip()
        row_size = str(row_values.get("size", "")).strip()

        if row_kind.lower() == normalized_target_kind and row_size.lower() == normalized_target_size:
            return ProductInputRow(
                kind=row_kind,
                size=row_size,
                values={key: "" if value is None else str(value) for key, value in row_values.items()},
            )

    raise ValueError(
        f"No Excel row found for kind='{target_kind}' and size='{target_size}' in {workbook_path}"
    )


def load_field_definitions(json_path: Path) -> list[FieldDefinition]:
    if not json_path.exists():
        raise ValueError(f"Price/Stock/Shipping JSON file was not found: {json_path}")

    payload = json.loads(json_path.read_text(encoding="utf-8"))
    return [FieldDefinition(**field) for field in payload["fields"]]


def xpath_literal(value: str) -> str:
    if '"' not in value:
        return f'"{value}"'
    if "'" not in value:
        return f"'{value}'"
    parts = value.split('"')
    return 'concat(' + ', \'"\', '.join(f'"{part}"' for part in parts) + ')'


def generate_sku_suffix(length: int = 7) -> str:
    alphabet = string.ascii_uppercase + string.digits
    return "".join(random.choices(alphabet, k=length))


def normalize_field_value(value: str) -> str:
    normalized = str(value).strip()
    if not normalized:
        return ""

    try:
        numeric_value = float(normalized)
    except ValueError:
        return " ".join(normalized.split()).lower()

    if numeric_value.is_integer():
        return str(int(numeric_value))
    return ("%f" % numeric_value).rstrip("0").rstrip(".")


def find_matching_label_element(
    driver: webdriver.Firefox,
    field_label: str,
) -> WebElement | None:
    normalized_target = normalize_field_label(field_label)
    label_elements = driver.find_elements(
        By.XPATH,
        "//div[contains(@class,'AttributeItemLabelName')]",
    )
    for label_element in label_elements:
        normalized_text = normalize_field_label(label_element.text)
        if (
            normalized_text == normalized_target
            or normalized_target in normalized_text
            or normalized_text in normalized_target
        ):
            return label_element
    return None


def locate_field_label_element(
    driver: webdriver.Firefox,
    field_label: str,
    timeout_per_scroll: float = 3,
) -> WebElement:
    immediate_match = find_matching_label_element(driver, field_label)
    if immediate_match is not None:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", immediate_match)
        return immediate_match

    def wait_callback(_: webdriver.Firefox) -> WebElement | None:
        return find_matching_label_element(driver, field_label)

    for scroll_fraction in (0.35, 0.6, 0.85, 1.0):
        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight * arguments[0]);",
            scroll_fraction,
        )
        try:
            label_element = WebDriverWait(driver, timeout_per_scroll).until(wait_callback)
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", label_element)
            return label_element
        except TimeoutException:
            continue

    raise TimeoutException(f"Could not find label element for field: {field_label}")


def get_field_wrapper(
    driver: webdriver.Firefox,
    field_label: str,
    timeout_per_scroll: float = 3,
) -> WebElement:
    label_element = locate_field_label_element(
        driver,
        field_label,
        timeout_per_scroll=timeout_per_scroll,
    )
    wrapper = driver.execute_script(
        """
        const label = arguments[0];
        const wrapperSelectors = [
            ".styles__FocusWrapper-sc-7uiywl-3",
            ".styles__EditAttributeItemWrapper-sc-gni56x-0",
            ".styles__AttributeItemFieldWrapper-sc-ske8mu-0",
        ];
        for (const selector of wrapperSelectors) {
            const candidate = label.closest(selector);
            if (candidate) {
                return candidate;
            }
        }
        let current = label.parentElement;
        while (current) {
            if (
                current.querySelector("input, [role='combobox'], #trigger-single-select")
                && current.textContent.includes(label.textContent)
            ) {
                return current;
            }
            current = current.parentElement;
        }
        return label.parentElement;
        """,
        label_element,
    )
    if wrapper is None:
        raise TimeoutException(f"Could not find field wrapper for label: {field_label}")
    return wrapper


def get_editable_field_element(
    driver: webdriver.Firefox,
    field_label: str,
    prefer_combobox: bool,
    timeout_per_scroll: float = 3,
) -> WebElement:
    label_element = locate_field_label_element(
        driver,
        field_label,
        timeout_per_scroll=timeout_per_scroll,
    )
    field_element = driver.execute_script(
        """
        const label = arguments[0];
        const preferCombobox = arguments[1];
        const wrapperSelectors = [
            ".styles__AttributeWrapper-sc-ske8mu-5",
            ".styles__FocusWrapper-sc-7uiywl-3",
            ".styles__EditAttributeItemWrapper-sc-gni56x-0",
            ".styles__AttributeItemFieldWrapper-sc-ske8mu-0",
        ];

        function isEditableInput(element) {
            if (!element || element.tagName !== "INPUT") {
                return false;
            }
            if (element.type === "radio" || element.readOnly || element.disabled) {
                return false;
            }
            return true;
        }

        function findControl(root) {
            if (!root) {
                return null;
            }
            if (preferCombobox) {
                return root.querySelector("button[role='combobox'], [role='combobox']");
            }
            const inputs = Array.from(root.querySelectorAll("input"));
            return inputs.find(isEditableInput) || null;
        }

        for (const selector of wrapperSelectors) {
            const wrapper = label.closest(selector);
            const control = findControl(wrapper);
            if (control) {
                return control;
            }
        }

        let current = label.parentElement;
        while (current) {
            const control = findControl(current);
            if (control) {
                return control;
            }
            current = current.parentElement;
        }
        return null;
        """,
        label_element,
        prefer_combobox,
    )
    if field_element is None:
        raise TimeoutException(f"Could not find editable element for label: {field_label}")
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", field_element)
    return field_element


def set_input_value(
    driver: webdriver.Firefox,
    input_element: WebElement,
    field_value: str,
    field_label: str | None = None,
) -> None:
    def reacquire_input() -> WebElement:
        if not field_label:
            return input_element
        return get_editable_field_element(
            driver,
            field_label,
            prefer_combobox=False,
            timeout_per_scroll=1,
        )

    def handle_length_field_bug() -> bool:
        if field_label != "Length" or not field_value:
            return False

        length_input = reacquire_input()
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", length_input)
        driver.execute_script("arguments[0].focus();", length_input)
        sleep(0.05)
        length_input.send_keys(Keys.CONTROL, "a")
        sleep(0.05)
        length_input.send_keys(Keys.DELETE)
        sleep(0.05)
        length_input.send_keys(field_value[0])
        log_event(
            "FORM",
            "Applied Length-field bug workaround: typed the first character to trigger the blur behavior.",
        )
        sleep(0.15)
        ActionChains(driver).send_keys(Keys.TAB).perform()
        sleep(0.15)

        length_input = reacquire_input()
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", length_input)
        driver.execute_script("arguments[0].focus();", length_input)
        sleep(0.05)
        length_input.send_keys(Keys.CONTROL, "a")
        sleep(0.05)
        length_input.send_keys(Keys.DELETE)
        sleep(0.05)
        length_input.send_keys(field_value)
        sleep(0.1)
        current_value = (length_input.get_attribute("value") or "").strip()
        if current_value == field_value:
            log_event(
                "FORM",
                f"Length workaround completed successfully with final value: {field_value}",
            )
            return True
        return False

    if handle_length_field_bug():
        return

    active_input = input_element
    for attempt in range(3):
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", active_input)
            driver.execute_script("arguments[0].focus();", active_input)
            sleep(0.05)
            active_input.send_keys(Keys.CONTROL, "a")
            sleep(0.05)
            active_input.send_keys(Keys.DELETE)
            sleep(0.05)

            for character in field_value:
                try:
                    active_input.send_keys(character)
                except WebDriverException:
                    active_input = reacquire_input()
                    driver.execute_script("arguments[0].focus();", active_input)
                    sleep(0.02)
                    current_partial = (active_input.get_attribute("value") or "").strip()
                    remaining_value = field_value[len(current_partial):]
                    if remaining_value:
                        active_input.send_keys(remaining_value)
                    break
                sleep(0.02)

            current_value = (active_input.get_attribute("value") or "").strip()
            if current_value != field_value:
                active_input = reacquire_input()
                active_input.send_keys(Keys.CONTROL, "a")
                sleep(0.05)
                active_input.send_keys(Keys.DELETE)
                sleep(0.05)
                active_input.send_keys(field_value)
                current_value = (active_input.get_attribute("value") or "").strip()

            if current_value == field_value:
                return
        except WebDriverException:
            if attempt == 2:
                raise
            sleep(0.15)
            active_input = reacquire_input()

    current_value = (active_input.get_attribute("value") or "").strip()
    raise TimeoutException(
        f"Typed value did not stick. Expected '{field_value}', got '{current_value}'."
    )


def get_field_current_value(
    driver: webdriver.Firefox,
    field_label: str,
    timeout_per_scroll: float = 0.35,
) -> str:
    input_element = get_editable_field_element(
        driver,
        field_label,
        prefer_combobox=False,
        timeout_per_scroll=timeout_per_scroll,
    )
    return (input_element.get_attribute("value") or "").strip()


def normalize_field_value(value: str) -> str:
    return " ".join((value or "").strip().split()).lower()


def field_values_match(current_value: str, expected_value: str) -> bool:
    return normalize_field_value(current_value) == normalize_field_value(expected_value)


def get_combobox_current_value(
    driver: webdriver.Firefox,
    field_label: str,
    timeout_per_scroll: float = 0.35,
) -> str:
    combobox = get_editable_field_element(
        driver,
        field_label,
        prefer_combobox=True,
        timeout_per_scroll=timeout_per_scroll,
    )
    button_text = combobox.text.strip()
    if button_text:
        return button_text
    return (combobox.get_attribute("value") or "").strip()


def fill_text_or_number_field(
    driver: webdriver.Firefox,
    field_label: str,
    field_value: str,
    timeout_per_scroll: float = 3,
) -> None:
    current_value = get_field_current_value(
        driver,
        field_label,
        timeout_per_scroll=min(timeout_per_scroll, 1),
    )
    if field_values_match(current_value, field_value):
        log_event("FORM", f"Skipping {field_label}: already filled with expected value '{current_value}'.")
        return

    input_element = get_editable_field_element(
        driver,
        field_label,
        prefer_combobox=False,
        timeout_per_scroll=timeout_per_scroll,
    )
    set_input_value(driver, input_element, field_value, field_label=field_label)
    log_event("FORM", f"Filled {field_label}: {field_value}")


def click_field_by_mouse(
    driver: webdriver.Firefox,
    field_label: str,
    timeout_per_scroll: float = 1,
) -> None:
    input_element = get_editable_field_element(
        driver,
        field_label,
        prefer_combobox=False,
        timeout_per_scroll=timeout_per_scroll,
    )
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_element)
    click_element_resilient(driver, input_element)
    log_event("FORM", f"Clicked {field_label} field by mouse.")


def click_screen_point_multiple_times(
    x: int,
    y: int,
    click_count: int = 4,
    pause_between_clicks: float = 0.2,
) -> None:
    mouse = MouseController()
    original_position = mouse.position
    log_event(
        "MOUSE",
        f"Preparing screen-pixel click sequence at ({x}, {y}) for {click_count} click(s).",
    )
    pyautogui.moveTo(x, y, duration=0.15)
    mouse.position = (x, y)
    for click_index in range(1, click_count + 1):
        mouse.click(PynputButton.left, 1)
        log_event(
            "MOUSE",
            f"Clicked screen pixel ({x}, {y}) [{click_index}/{click_count}].",
        )
        sleep(pause_between_clicks)
    mouse.position = original_position
    log_event(
        "MOUSE",
        f"Restored mouse position to ({int(original_position[0])}, {int(original_position[1])}).",
    )


def get_element_screen_center(driver: webdriver.Firefox, element: WebElement) -> tuple[int, int]:
    geometry = driver.execute_script(
        """
        const element = arguments[0];
        const rect = element.getBoundingClientRect();
        return {
            left: rect.left,
            top: rect.top,
            width: rect.width,
            height: rect.height,
            screenX: window.screenX,
            screenY: window.screenY,
            outerWidth: window.outerWidth,
            outerHeight: window.outerHeight,
            innerWidth: window.innerWidth,
            innerHeight: window.innerHeight,
            devicePixelRatio: window.devicePixelRatio || 1
        };
        """,
        element,
    )
    device_pixel_ratio = float(geometry["devicePixelRatio"] or 1)
    horizontal_chrome = max(
        0.0,
        (float(geometry["outerWidth"]) - float(geometry["innerWidth"])) / 2.0,
    )
    vertical_chrome = max(
        0.0,
        float(geometry["outerHeight"]) - float(geometry["innerHeight"]) - horizontal_chrome,
    )
    center_x_css = float(geometry["screenX"]) + horizontal_chrome + float(geometry["left"]) + (
        float(geometry["width"]) / 2.0
    )
    center_y_css = float(geometry["screenY"]) + vertical_chrome + float(geometry["top"]) + (
        float(geometry["height"]) / 2.0
    )
    return (
        int(round(center_x_css * device_pixel_ratio)),
        int(round(center_y_css * device_pixel_ratio)),
    )


def click_element_via_autogui(driver: webdriver.Firefox, element: WebElement, label: str) -> None:
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});",
        element,
    )
    sleep(0.1)
    screen_x, screen_y = get_element_screen_center(driver, element)
    log_event(
        "MOUSE",
        f"Clicking {label} via screen coordinates at ({screen_x}, {screen_y}).",
    )
    click_screen_point_multiple_times(
        screen_x,
        screen_y,
        click_count=1,
        pause_between_clicks=0.05,
    )


def fill_combobox_field(
    driver: webdriver.Firefox,
    field_label: str,
    field_value: str,
    timeout_per_scroll: float = 3,
) -> None:
    current_value = get_combobox_current_value(
        driver,
        field_label,
        timeout_per_scroll=min(timeout_per_scroll, 1),
    )
    if field_values_match(current_value, field_value):
        log_event("FORM", f"Skipping {field_label}: already selected as '{current_value}'.")
        return

    combobox = get_editable_field_element(
        driver,
        field_label,
        prefer_combobox=True,
        timeout_per_scroll=timeout_per_scroll,
    )
    select_combobox_option(driver, combobox, field_value, field_label)
    log_event("FORM", f"Selected {field_label}: {field_value}")


def click_element_resilient(driver: webdriver.Firefox, element: WebElement) -> None:
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});",
        element,
    )
    try:
        element.click()
        return
    except WebDriverException:
        pass

    driver.execute_script("arguments[0].click();", element)


def click_element_without_js(driver: webdriver.Firefox, element: WebElement) -> None:
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});",
        element,
    )
    try:
        element.click()
        return
    except WebDriverException:
        pass

    ActionChains(driver).move_to_element(element).pause(0.05).click().perform()


def get_visible_dropdown_container(driver: webdriver.Firefox) -> WebElement:
    dropdown_candidates = driver.find_elements(
        By.CSS_SELECTOR,
        "[data-testid='content-single-select'], [data-testid='content-multi-select'], "
        ".styles__DropdownContent-sc-zkytp-1, .styles__DropdownContent-sc-lf8o9y-2",
    )
    for candidate in dropdown_candidates:
        if candidate.is_displayed():
            return candidate
    raise TimeoutException("Could not find visible dropdown container after opening combobox.")


def close_dropdown_with_escape(driver: webdriver.Firefox, field_label: str) -> None:
    log_event("FORM", f"Closing dropdown for {field_label} with Escape.")
    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    sleep(0.15)


def select_combobox_option(
    driver: webdriver.Firefox,
    combobox: WebElement,
    field_value: str,
    field_label: str,
) -> None:
    click_element_without_js(driver, combobox)
    sleep(0.15)
    dropdown_container = get_visible_dropdown_container(driver)

    search_inputs = dropdown_container.find_elements(
        By.CSS_SELECTOR,
        "input[aria-label='Search'], input[placeholder='Select'], input[type='search'], input[type='text']",
    )
    visible_search_input = next(
        (element for element in search_inputs if element.is_displayed() and element.is_enabled()),
        None,
    )
    if visible_search_input is not None:
        log_event("FORM", f"Typing into dropdown search for {field_label}: {field_value}")
        driver.execute_script("arguments[0].focus();", visible_search_input)
        sleep(0.05)
        visible_search_input.send_keys(Keys.CONTROL, "a")
        sleep(0.05)
        visible_search_input.send_keys(Keys.DELETE)
        sleep(0.05)
        for character in field_value:
            visible_search_input.send_keys(character)
            sleep(0.02)
        sleep(0.25)

    normalized_target = field_value.strip().lower()
    matching_label = None

    for _ in range(10):
        option_labels = dropdown_container.find_elements(
            By.CSS_SELECTOR,
            "label[for], .style__LabelWrapper-sc-n7qfg8-0, [role='radio'] label",
        )
        for option_label in option_labels:
            try:
                if not option_label.is_displayed():
                    continue
                option_text = " ".join(option_label.text.strip().split()).lower()
            except WebDriverException:
                continue
            if option_text == normalized_target:
                matching_label = option_label
                break
        if matching_label is not None:
            break
        sleep(0.1)

    if matching_label is None:
        option_labels = dropdown_container.find_elements(
            By.CSS_SELECTOR,
            "label[for], .style__LabelWrapper-sc-n7qfg8-0, [role='radio'] label",
        )
        available_options = []
        for option_label in option_labels:
            try:
                if option_label.is_displayed() and option_label.text.strip():
                    available_options.append(" ".join(option_label.text.strip().split()))
            except WebDriverException:
                continue
        raise TimeoutException(
            f"Could not find dropdown option '{field_value}' for {field_label}. "
            f"Available options: {available_options}"
        )

    try:
        click_element_via_autogui(driver, matching_label, f"{field_label} option '{field_value}'")
    except Exception:
        click_element_without_js(driver, matching_label)
    sleep(0.25)
    close_dropdown_with_escape(driver, field_label)


def get_tag_input_wrapper(driver: webdriver.Firefox, field_label: str) -> WebElement:
    wrapper = get_field_wrapper(driver, field_label, timeout_per_scroll=1)
    tag_wrapper = driver.execute_script(
        """
        const wrapper = arguments[0];
        return (
            wrapper.querySelector('.rti--container') ||
            wrapper.querySelector('.multi-select-field-wrapper') ||
            wrapper.querySelector("[data-testid='trigger-multi-select']")?.closest('.multi-select-field-wrapper') ||
            wrapper.querySelector("[data-testid='trigger-multi-select']") ||
            wrapper
        );
        """,
        wrapper,
    )
    if tag_wrapper is None:
        raise TimeoutException(f"Could not find tag input wrapper for field: {field_label}")
    return tag_wrapper


def get_tag_input_current_value(driver: webdriver.Firefox, field_label: str) -> str:
    tag_wrapper = get_tag_input_wrapper(driver, field_label)
    values = driver.execute_script(
        """
        const wrapper = arguments[0];
        const helperText = wrapper.parentElement?.querySelector('.styles__MultiSelectHelperText-sc-ske8mu-16');
        if (helperText && helperText.textContent.trim()) {
            return helperText.textContent
                .split(',')
                .map((value) => value.trim())
                .filter(Boolean);
        }

        const checkedCheckboxValues = Array.from(
            wrapper.querySelectorAll("input[type='checkbox']:checked")
        )
            .map((element) => element.value || element.getAttribute('data-label') || '')
            .filter(Boolean);
        if (checkedCheckboxValues.length) {
            return checkedCheckboxValues;
        }

        return Array.from(wrapper.querySelectorAll('[role="tab"][label], [role="tab"][value]'))
            .map((element) => element.getAttribute('label') || element.getAttribute('value') || element.textContent.trim())
            .filter(Boolean);
        """,
        tag_wrapper,
    )
    return ", ".join(values)


def fill_tag_input_field(
    driver: webdriver.Firefox,
    field_label: str,
    field_value: str,
) -> None:
    desired_values = [value.strip() for value in field_value.split(",") if value.strip()]
    current_value = get_tag_input_current_value(driver, field_label)
    current_values = [value.strip() for value in current_value.split(",") if value.strip()]
    if current_values == desired_values:
        log_event("FORM", f"Skipping {field_label}: already filled with expected value(s) '{field_value}'.")
        return

    tag_wrapper = get_tag_input_wrapper(driver, field_label)
    multi_select_combobox = driver.execute_script(
        """
        const wrapper = arguments[0];
        return wrapper.matches("[data-testid='trigger-multi-select'], button[role='combobox']")
            ? wrapper
            : wrapper.querySelector("[data-testid='trigger-multi-select'], button[role='combobox']");
        """,
        tag_wrapper,
    )
    if multi_select_combobox is not None:
        click_element_without_js(driver, multi_select_combobox)
        sleep(0.2)
        dropdown_container = get_visible_dropdown_container(driver)
        search_input = next(
            (
                element
                for element in dropdown_container.find_elements(
                    By.CSS_SELECTOR,
                    "input[id*='checkbox-tree'][type='text'], input[placeholder='Search Paramter'], input[aria-label='Search']",
                )
                if element.is_displayed() and element.is_enabled()
            ),
            None,
        )
        if search_input is None:
            raise TimeoutException(f"Could not find Brand Color search input for field: {field_label}")

        selected_labels = dropdown_container.find_elements(
            By.CSS_SELECTOR,
            "input[type='checkbox']:checked",
        )
        selected_map: dict[str, WebElement] = {}
        for selected_label in selected_labels:
            try:
                selected_value = (
                    selected_label.get_attribute("value")
                    or selected_label.get_attribute("data-label")
                    or ""
                ).strip()
            except WebDriverException:
                continue
            if selected_value:
                selected_map[normalize_field_value(selected_value)] = selected_label

        for current_selected_value, selected_input in selected_map.items():
            if current_selected_value in {normalize_field_value(value) for value in desired_values}:
                continue
            checkbox_id = selected_input.get_attribute("id")
            if not checkbox_id:
                continue
            matching_label = dropdown_container.find_element(By.CSS_SELECTOR, f"label[for='{checkbox_id}']")
            try:
                click_element_via_autogui(
                    driver,
                    matching_label,
                    f"{field_label} option '{matching_label.text.strip()}'",
                )
            except Exception:
                click_element_without_js(driver, matching_label)
            sleep(0.15)

        for desired_value in desired_values:
            log_event("FORM", f"Typing into multi-select search for {field_label}: {desired_value}")
            driver.execute_script("arguments[0].focus();", search_input)
            sleep(0.05)
            search_input.send_keys(Keys.CONTROL, "a")
            sleep(0.05)
            search_input.send_keys(Keys.DELETE)
            sleep(0.05)
            for character in desired_value:
                search_input.send_keys(character)
                sleep(0.02)
            sleep(0.25)

            normalized_target = normalize_field_value(desired_value)
            matching_label = None
            for _ in range(10):
                option_labels = dropdown_container.find_elements(
                    By.CSS_SELECTOR,
                    "label[for], .style__LabelWrapper-sc-n7qfg8-0",
                )
                for option_label in option_labels:
                    try:
                        if not option_label.is_displayed():
                            continue
                        option_text = normalize_field_value(option_label.text)
                    except WebDriverException:
                        continue
                    if option_text == normalized_target:
                        matching_label = option_label
                        break
                if matching_label is not None:
                    break
                sleep(0.1)

            if matching_label is None:
                raise TimeoutException(
                    f"Could not find multi-select option '{desired_value}' for {field_label}."
                )

            checkbox_id = matching_label.get_attribute("for")
            if checkbox_id:
                checkbox = dropdown_container.find_element(By.ID, checkbox_id)
                if checkbox.is_selected():
                    continue
            try:
                click_element_via_autogui(driver, matching_label, f"{field_label} option '{desired_value}'")
            except Exception:
                click_element_without_js(driver, matching_label)
            sleep(0.15)

        close_dropdown_with_escape(driver, field_label)
        log_event("FORM", f"Filled {field_label}: {field_value}")
        return

    input_element = driver.execute_script(
        """
        const wrapper = arguments[0];
        return wrapper.querySelector('.rti--input, input');
        """,
        tag_wrapper,
    )
    if input_element is None:
        raise TimeoutException(f"Could not find tag input element for field: {field_label}")

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_element)
    input_element.click()

    existing_count = len(current_values)
    for _ in range(existing_count):
        input_element.send_keys(Keys.BACKSPACE)
        sleep(0.05)

    for value in desired_values:
        input_element.send_keys(value)
        input_element.send_keys(Keys.ENTER)
        sleep(0.05)

    log_event("FORM", f"Filled {field_label}: {field_value}")


def fill_price_stock_shipping_fields(
    driver: webdriver.Firefox,
    field_definitions: list[FieldDefinition],
    product_input_row: ProductInputRow,
) -> FillResult:
    generated_values: dict[str, str] = {}
    skipped_fields: set[str] = set()
    handling_fee_fields_unavailable = False

    for field in field_definitions:
        if handling_fee_fields_unavailable and field.label in {
            "Zonal handling fee",
            "National handling fee",
        }:
            skipped_fields.add(field.label)
            log_event(
                "PRICE",
                f"Skipping {field.label}: handling fee fields are not present in current listing state.",
            )
            continue

        raw_value = product_input_row.values.get(field.label, "").strip()
        if not raw_value:
            log_event("PRICE", f"Skipping {field.label}: no Excel value provided.")
            continue

        if field.label == "Seller SKU ID":
            raw_value = f"{raw_value}{generate_sku_suffix()}"
            generated_values[field.label] = raw_value
            log_event("PRICE", f"Generated Seller SKU ID value: {raw_value}")

        timeout_per_scroll = 3 if field.required else 0.25
        try:
            if field.input_type == "combobox":
                fill_combobox_field(
                    driver,
                    field.label,
                    raw_value,
                    timeout_per_scroll=timeout_per_scroll,
                )
            else:
                fill_text_or_number_field(
                    driver,
                    field.label,
                    raw_value,
                    timeout_per_scroll=timeout_per_scroll,
                )
        except TimeoutException:
            if field.required:
                raise
            if field.label == "Local handling fee":
                handling_fee_fields_unavailable = True
                skipped_fields.update(
                    {"Local handling fee", "Zonal handling fee", "National handling fee"}
                )
                log_event(
                    "PRICE",
                    "Skipping Local handling fee: field not present in current listing state. "
                    "Assuming Zonal and National handling fee fields are also absent.",
                )
                continue
            skipped_fields.add(field.label)
            log_event("PRICE", f"Skipping {field.label}: field not present in current listing state.")

    generated_sku_value = generated_values.get("Seller SKU ID", "")
    if generated_sku_value:
        try:
            current_sku_value = get_field_current_value(driver, "Seller SKU ID")
            if current_sku_value != generated_sku_value:
                log_event(
                    "PRICE",
                    "Seller SKU ID changed after later field updates. "
                    f"Re-applying generated value: {generated_sku_value}",
                )
                fill_text_or_number_field(driver, "Seller SKU ID", generated_sku_value)
            else:
                log_event("PRICE", f"Seller SKU ID persisted: {generated_sku_value}")
        except TimeoutException:
            log_event("PRICE", "Could not re-check Seller SKU ID at the end of fill.")

    return FillResult(generated_values=generated_values, skipped_fields=skipped_fields)


def build_expected_field_values(
    field_definitions: list[FieldDefinition],
    product_input_row: ProductInputRow,
    fill_result: FillResult,
) -> dict[str, str]:
    expected_values: dict[str, str] = {}
    for field in field_definitions:
        if field.label in fill_result.skipped_fields:
            continue

        if field.label in fill_result.generated_values:
            expected_values[field.label] = fill_result.generated_values[field.label]
            continue

        raw_value = product_input_row.values.get(field.label, "").strip()
        if raw_value:
            expected_values[field.label] = raw_value
    return expected_values


def verify_and_refill_price_stock_shipping_fields(
    driver: webdriver.Firefox,
    field_definitions: list[FieldDefinition],
    expected_values: dict[str, str],
    skipped_fields: set[str],
    max_passes: int = 2,
) -> None:
    for verification_pass in range(1, max_passes + 1):
        mismatches: list[tuple[FieldDefinition, str, str]] = []

        for field in field_definitions:
            if field.label in skipped_fields:
                continue
            expected_value = expected_values.get(field.label, "").strip()
            if not expected_value:
                continue

            try:
                if field.input_type == "combobox":
                    current_value = get_combobox_current_value(
                        driver,
                        field.label,
                        timeout_per_scroll=0.2 if not field.required else 0.35,
                    )
                else:
                    current_value = get_field_current_value(
                        driver,
                        field.label,
                        timeout_per_scroll=0.2 if not field.required else 0.35,
                    )
            except TimeoutException:
                if field.required:
                    raise
                continue

            if normalize_field_value(current_value) != normalize_field_value(expected_value):
                mismatches.append((field, expected_value, current_value))

        if not mismatches:
            log_event(
                "VERIFY",
                f"Verification pass {verification_pass}: all filled values matched expected data.",
            )
            return

        log_event(
            "VERIFY",
            f"Verification pass {verification_pass}: found {len(mismatches)} mismatch(es). "
            "Refilling changed fields.",
        )
        for field, expected_value, current_value in mismatches:
            log_event(
                "VERIFY",
                f"Re-filling {field.label}: expected '{expected_value}', current '{current_value}'.",
            )
            if field.input_type == "combobox":
                fill_combobox_field(driver, field.label, expected_value, timeout_per_scroll=1)
            else:
                fill_text_or_number_field(driver, field.label, expected_value, timeout_per_scroll=1)

    log_event("VERIFY", "Verification finished after maximum refill passes.")


def dismiss_optional_ad_popup(driver: webdriver.Firefox, timeout_seconds: int = 5) -> None:
    close_button_locator = (
        By.XPATH,
        "//button[@data-testid='button' and normalize-space()='Close']",
    )

    try:
        close_button = WebDriverWait(driver, timeout_seconds).until(
            EC.element_to_be_clickable(close_button_locator)
        )
    except TimeoutException:
        log_event("PAGE", "Optional ad popup did not appear.")
        return

    close_button.click()
    log_event("PAGE", "Optional ad popup closed.")


def wait_for_clickable(
    driver: webdriver.Firefox,
    locator: tuple[str, str],
    timeout_seconds: int = 15,
) -> WebElement:
    return WebDriverWait(driver, timeout_seconds).until(EC.element_to_be_clickable(locator))


def fill_brand_name(driver: webdriver.Firefox, brand_name: str) -> None:
    brand_input_locator = (By.CSS_SELECTOR, "input[placeholder='Enter Brand Name']")
    check_brand_button_locator = (
        By.XPATH,
        "//button[@data-testid='button' and normalize-space()='Check Brand']",
    )

    brand_input = wait_for_clickable(driver, brand_input_locator)
    brand_input.clear()
    brand_input.send_keys(brand_name)
    log_event("LISTING", f"Entered brand name: {brand_name}")

    check_brand_button = wait_for_clickable(driver, check_brand_button_locator)
    check_brand_button.click()
    log_event("LISTING", "Clicked Check Brand.")


def click_create_new_listing(driver: webdriver.Firefox) -> None:
    create_listing_button_locator = (
        By.XPATH,
        "//button[@data-testid='button' and normalize-space()='Create new listing']",
    )
    create_listing_button = wait_for_clickable(driver, create_listing_button_locator)
    create_listing_button.click()
    log_event("LISTING", "Clicked Create new listing.")


def click_optional_continue(driver: webdriver.Firefox, timeout_seconds: int = 5) -> None:
    continue_button_locator = (
        By.XPATH,
        "//button[@type='button' and .//span[normalize-space()='Continue']]",
    )

    try:
        continue_button = wait_for_clickable(driver, continue_button_locator, timeout_seconds)
    except TimeoutException:
        log_event("LISTING", "Optional Continue button did not appear.")
        return

    continue_button.click()
    log_event("LISTING", "Clicked optional Continue button.")


def preview_selected_image_folder(config: BotConfig, brand_name: str) -> ImageFolder | None:
    if not str(config.image_directory):
        log_event("IMAGES", "Image directory not configured yet, skipping image folder selection.")
        return None

    selected_folder = choose_image_folder_for_brand(config.image_directory, brand_name)
    if selected_folder is None:
        log_event("IMAGES", f"No image folder is available for brand: {brand_name}")
        return None

    log_event(
        "IMAGES",
        f"Selected image folder: {selected_folder.folder_path.name} "
        f"with {len(selected_folder.image_paths)} image(s)",
    )
    for image_path in selected_folder.image_paths:
        log_event("IMAGES", f"Queued image file: {image_path.name}")
    log_event(
        "IMAGES",
        "Folder name after successful upload would become: "
        f"{build_exhausted_folder_name(selected_folder, brand_name)}",
    )
    return selected_folder


def click_image_slot(driver: webdriver.Firefox, slot_id: str) -> None:
    image_slot = wait_for_clickable(driver, (By.ID, slot_id))
    image_slot.click()
    log_event("IMAGES", f"Selected image slot: {slot_id}")


def upload_image_to_selected_slot(driver: webdriver.Firefox, image_path: Path) -> None:
    upload_input = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, "upload-image"))
    )
    upload_input.send_keys(str(image_path))
    WebDriverWait(driver, 10).until(lambda _: upload_input.get_attribute("value"))
    sleep(0.5)
    log_event("IMAGES", f"Uploaded image file into active slot: {image_path.name}")


def is_image_slot_uploaded(driver: webdriver.Firefox, slot_id: str) -> bool:
    slot_element = driver.find_element(By.ID, slot_id)
    has_thumbnail = bool(slot_element.find_elements(By.CSS_SELECTOR, "img.styles__Img-sc-1o2k4cf-0"))
    has_check_icon = bool(slot_element.find_elements(By.CSS_SELECTOR, "i.fa-check"))
    has_plus_icon = bool(slot_element.find_elements(By.CSS_SELECTOR, "i.fa-plus"))
    return has_thumbnail and has_check_icon and not has_plus_icon


def wait_for_uploaded_image_slots(
    driver: webdriver.Firefox,
    slot_ids: list[str],
    timeout_seconds: int = 45,
) -> None:
    def all_slots_uploaded(_: webdriver.Firefox) -> bool:
        return all(is_image_slot_uploaded(driver, slot_id) for slot_id in slot_ids)

    try:
        WebDriverWait(driver, timeout_seconds).until(all_slots_uploaded)
    except TimeoutException as error:
        incomplete_slots = [
            slot_id for slot_id in slot_ids if not is_image_slot_uploaded(driver, slot_id)
        ]
        raise TimeoutException(
            "Image upload verification timed out. Incomplete slot(s): "
            + ", ".join(incomplete_slots)
        ) from error

    log_event("IMAGES", f"Verified uploaded image slot(s): {', '.join(slot_ids)}")


def upload_image_folder(
    driver: webdriver.Firefox,
    image_folder: ImageFolder,
    brand_name: str,
    pause_controller: PauseController,
    config: BotConfig,
) -> None:
    if not image_folder.image_paths:
        raise ValueError(f"No images found in folder: {image_folder.folder_path}")

    upload_count = min(len(image_folder.image_paths), len(IMAGE_SLOT_IDS))
    for index, (slot_id, image_path) in enumerate(
        zip(IMAGE_SLOT_IDS[:upload_count], image_folder.image_paths),
        start=1,
    ):
        checkpoint_pause(pause_controller, f"Before upload slot {index}", driver, config)
        click_image_slot(driver, slot_id)
        upload_image_to_selected_slot(driver, image_path)
        checkpoint_pause(pause_controller, f"After upload slot {index}", driver, config)

    uploaded_slot_ids = IMAGE_SLOT_IDS[:upload_count]
    wait_for_uploaded_image_slots(driver, uploaded_slot_ids)
    checkpoint_pause(pause_controller, "After verifying uploaded image slots", driver, config)

    if len(image_folder.image_paths) > len(IMAGE_SLOT_IDS):
        log_event(
            "IMAGES",
            f"Only uploaded the first {len(IMAGE_SLOT_IDS)} image(s); "
            f"{len(image_folder.image_paths) - len(IMAGE_SLOT_IDS)} extra image(s) were skipped.",
        )

    mark_image_folder_exhausted(image_folder, brand_name)


def open_selling_info_tab(driver: webdriver.Firefox) -> None:
    selling_info_tab_locator = (
        By.XPATH,
        "//button[@role='tab' and .//span[contains(normalize-space(), 'Price, Stock and Shipping Information')]]",
    )
    selling_info_tab = wait_for_clickable(driver, selling_info_tab_locator)
    selling_info_tab.click()
    log_event("NAV", "Opened Price, Stock and Shipping Information tab.")


def open_product_description_tab(driver: webdriver.Firefox) -> None:
    product_description_tab_locator = (
        By.XPATH,
        "//button[@role='tab' and .//span[contains(normalize-space(), 'Product Description')]]",
    )
    product_description_tab = wait_for_clickable(driver, product_description_tab_locator)
    product_description_tab.click()
    log_event("NAV", "Opened Product Description tab.")


def fill_size_qualifier_field(driver: webdriver.Firefox, field_value: str) -> None:
    size_wrapper = get_field_wrapper(driver, "Size", timeout_per_scroll=1)
    qualifier_combobox = driver.execute_script(
        """
        const wrapper = arguments[0];
        const comboboxes = Array.from(
            wrapper.querySelectorAll("button[role='combobox'], [role='combobox']")
        ).filter((element) => {
            if (!element) {
                return false;
            }
            const style = window.getComputedStyle(element);
            return style.display !== 'none' && style.visibility !== 'hidden';
        });
        if (comboboxes.length < 2) {
            return comboboxes.length === 1 ? comboboxes[0] : null;
        }
        return comboboxes[1];
        """,
        size_wrapper,
    )
    if qualifier_combobox is None:
        raise TimeoutException("Could not find Size Qualifier combobox inside the Size field wrapper.")
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", qualifier_combobox)
    log_event(
        "DESC",
        f"Preparing second Size dropdown selection (current button like 'Number') -> {field_value}",
    )
    select_combobox_option(driver, qualifier_combobox, field_value, "Size Qualifier")
    log_event("DESC", f"Selected second Size dropdown (Size Qualifier): {field_value}")


def fill_product_description_fields(
    driver: webdriver.Firefox,
    field_definitions: list[FieldDefinition],
    product_input_row: ProductInputRow,
) -> FillResult:
    generated_values: dict[str, str] = {}
    skipped_fields: set[str] = set()

    size_qualifier_value = product_input_row.values.get("Size Qualifier", "").strip()
    main_size_value = (
        product_input_row.values.get("Size2", "").strip()
        or product_input_row.values.get("Size", "").strip()
    )

    for field in field_definitions:
        raw_value = product_input_row.values.get(field.label, "").strip()

        if field.label == "Size":
            if size_qualifier_value:
                fill_size_qualifier_field(driver, size_qualifier_value)
            else:
                log_event("DESC", "Skipping Size Qualifier: no Excel value provided.")

            if not main_size_value:
                log_event("DESC", "Skipping Size: no Excel value provided in Size2/Size columns.")
                continue

            log_event("DESC", f"Preparing main Size dropdown selection after qualifier -> {main_size_value}")
            fill_combobox_field(driver, field.label, main_size_value, timeout_per_scroll=1)
            continue

        if not raw_value:
            log_event("DESC", f"Skipping {field.label}: no Excel value provided.")
            continue

        if field.label == "Style Code":
            raw_value = f"{raw_value}{generate_sku_suffix()}"
            generated_values[field.label] = raw_value
            log_event("DESC", f"Generated Style Code value: {raw_value}")

        if field.input_type == "combobox":
            fill_combobox_field(driver, field.label, raw_value, timeout_per_scroll=1)
        elif field.input_type == "tag_input":
            fill_tag_input_field(driver, field.label, raw_value)
        else:
            fill_text_or_number_field(driver, field.label, raw_value, timeout_per_scroll=1)

    return FillResult(generated_values=generated_values, skipped_fields=skipped_fields)


def wait_for_changes_saved_toast(
    driver: webdriver.Firefox,
    pause_controller: PauseController,
    config: BotConfig,
    timeout_seconds: int = 20,
) -> None:
    deadline = datetime.now().timestamp() + timeout_seconds
    while datetime.now().timestamp() < deadline:
        checkpoint_pause(pause_controller, "Waiting for 'Changes saved!' toast", driver, config)
        if has_changes_saved_toast(driver):
            log_event("TOAST", "Detected success toast: Changes saved!")
            return
        sleep(0.25)

    raise TimeoutException("Timed out waiting for 'Changes saved!' toast after opening selling info tab.")


def has_changes_saved_toast(driver: webdriver.Firefox) -> bool:
    success_toast_locator = (
        By.XPATH,
        "//div[contains(@class,'toast-title-container')]//b[normalize-space()='Changes saved!']",
    )
    return bool(driver.find_elements(*success_toast_locator))


def cycle_save_clicks_until_toast(
    driver: webdriver.Firefox,
    pause_controller: PauseController,
    config: BotConfig,
    cycle_label: str = "page",
    timeout_seconds: int = 45,
) -> None:
    deadline = datetime.now().timestamp() + timeout_seconds
    click_points = [
        PRODUCT_DESCRIPTION_TAB_CLICK_POINT,
        SKU_PAGE_ALTERNATE_CLICK_POINT,
    ]
    click_index = 0

    log_event(
        "MOUSE",
        f"Starting alternating {cycle_label} click cycle until Selenium detects 'Changes saved!'.",
    )
    while datetime.now().timestamp() < deadline:
        checkpoint_pause(
            pause_controller,
            f"Cycling {cycle_label} clicks until save toast",
            driver,
            config,
        )
        if has_changes_saved_toast(driver):
            log_event("TOAST", f"Detected success toast during {cycle_label} click cycle.")
            return

        current_point = click_points[click_index % len(click_points)]
        click_screen_point_multiple_times(
            current_point[0],
            current_point[1],
            click_count=1,
            pause_between_clicks=0.05,
        )
        click_index += 1
        sleep(1)

    raise TimeoutException(
        f"Timed out while cycling {cycle_label} clicks waiting for the 'Changes saved!' toast."
    )


def print_runtime_context(config: BotConfig) -> None:
    log_event("BOOT", "Flipkart lister bot starting...")
    log_event("BOOT", f"Target URL: {config.listing_url}")
    log_event("BOOT", f"Firefox profile: {config.profile_name}")
    log_event("BOOT", f"Firefox profile path: {config.firefox_profile_path}")
    log_event(
        "BOOT",
        f"Image directory: {config.image_directory if str(config.image_directory) else 'not configured yet'}",
    )
    log_event(
        "BOOT",
        f"Data directory: {config.data_directory if str(config.data_directory) else 'not configured yet'}",
    )
    log_event("BOOT", f"Price/Stock/Shipping Excel: {config.price_stock_shipping_excel}")
    log_event("BOOT", f"Price/Stock/Shipping JSON: {config.price_stock_shipping_json}")
    log_event("BOOT", f"Product Description Excel: {config.product_description_excel}")
    log_event("BOOT", f"Product Description JSON: {config.product_description_json}")
    log_event("BOOT", f"Snapshot directory: {config.snapshot_directory}")
    log_event("BOOT", "Pause control: press Space in this terminal to pause at the next safe step.")


def main() -> None:
    try:
        config = BotConfig(profile_name=prompt_for_profile())
        print_runtime_context(config)
    except ValueError as exc:
        raise SystemExit(str(exc)) from exc

    try:
        log_event("BOOT", "Launching Firefox WebDriver...")
        driver = build_firefox_driver(config)
        log_event("BOOT", "Firefox WebDriver launched successfully.")
    except WebDriverException as exc:
        raise SystemExit(
            "Could not start Firefox WebDriver. Make sure Firefox is installed, the selected "
            "profile is valid, and that the same profile is not already open in another Firefox "
            "window. You can also set GECKODRIVER_PATH/FIREFOX_BINARY if needed.\n"
            f"Original error: {exc}"
        ) from exc

    pause_controller = PauseController()
    pause_controller.start()

    log_event("NAV", f"Opening listing page: {config.listing_url}")
    open_listing_page(driver, config.listing_url)
    log_event("NAV", "Listing page opened in Firefox.")
    checkpoint_pause(pause_controller, "Listing page opened", driver, config)
    dismiss_optional_ad_popup(driver)
    checkpoint_pause(pause_controller, "Optional popup handling complete", driver, config)
    fill_brand_name(driver, DEFAULT_BRAND_NAME)
    checkpoint_pause(pause_controller, "Brand entered", driver, config)
    click_create_new_listing(driver)
    checkpoint_pause(pause_controller, "Create new listing clicked", driver, config)
    click_optional_continue(driver)
    checkpoint_pause(pause_controller, "Optional continue handling complete", driver, config)
    selected_image_folder = preview_selected_image_folder(config, DEFAULT_BRAND_NAME)
    if selected_image_folder is not None:
        checkpoint_pause(pause_controller, "Image folder selected", driver, config)
        upload_image_folder(driver, selected_image_folder, DEFAULT_BRAND_NAME, pause_controller, config)
        checkpoint_pause(pause_controller, "Images uploaded", driver, config)
    open_selling_info_tab(driver)
    checkpoint_pause(pause_controller, "Selling info tab opened", driver, config)
    wait_for_changes_saved_toast(driver, pause_controller, config)
    sleep(2)
    phase_one_snapshot_path = save_named_html_snapshot(
        driver,
        config.snapshot_directory,
        PHASE_ONE_SNAPSHOT_NAME,
    )
    log_event("SNAPSHOT", f"Saved phase snapshot: {phase_one_snapshot_path}")
    product_input_row = load_product_input_row(
        config.price_stock_shipping_excel,
        DEFAULT_TARGET_KIND,
        DEFAULT_TARGET_SIZE,
    )
    field_definitions = load_field_definitions(config.price_stock_shipping_json)
    log_event(
        "DATA",
        f"Loaded Price/Stock/Shipping row for filling: kind={product_input_row.kind}, "
        f"size={product_input_row.size}",
    )
    checkpoint_pause(pause_controller, "Changes saved toast detected", driver, config)
    price_fill_result = fill_price_stock_shipping_fields(
        driver,
        field_definitions,
        product_input_row,
    )
    expected_field_values = build_expected_field_values(
        field_definitions,
        product_input_row,
        price_fill_result,
    )
    verify_and_refill_price_stock_shipping_fields(
        driver,
        field_definitions,
        expected_field_values,
        price_fill_result.skipped_fields,
    )
    checkpoint_pause(pause_controller, "Price/Stock/Shipping fields filled", driver, config)
    log_event(
        "VERIFY",
        "Price/Stock/Shipping verification completed. Starting save-detection click cycle.",
    )
    cycle_save_clicks_until_toast(driver, pause_controller, config, cycle_label="SKU page")
    checkpoint_pause(pause_controller, "Changes saved detected after SKU-page click cycle", driver, config)
    log_event(
        "MOUSE",
        "Clicking Product Description tab point immediately after detecting save toast.",
    )
    click_screen_point_multiple_times(
        PRODUCT_DESCRIPTION_TAB_CLICK_POINT[0],
        PRODUCT_DESCRIPTION_TAB_CLICK_POINT[1],
        click_count=1,
        pause_between_clicks=0.05,
    )
    log_event("NAV", "Waiting 2 seconds after Product Description tab click before field filling...")
    sleep(2)
    checkpoint_pause(pause_controller, "Product Description tab opened by screen click", driver, config)
    product_description_row = load_product_input_row(
        config.product_description_excel,
        DEFAULT_TARGET_KIND,
        DEFAULT_TARGET_SIZE,
        worksheet_name="Shorts Product Inputs",
    )
    product_description_field_definitions = load_field_definitions(config.product_description_json)
    log_event(
        "DATA",
        f"Loaded Product Description row: kind={product_description_row.kind}, "
        f"size={product_description_row.size}",
    )
    log_event("DESC", "Starting Product Description field fill from Excel + JSON mapping...")
    fill_product_description_fields(
        driver,
        product_description_field_definitions,
        product_description_row,
    )
    checkpoint_pause(pause_controller, "Product Description fields filled", driver, config)
    log_event(
        "VERIFY",
        "Product Description filling completed. Starting save-detection click cycle.",
    )
    cycle_save_clicks_until_toast(
        driver,
        pause_controller,
        config,
        cycle_label="Product Description page",
    )
    checkpoint_pause(
        pause_controller,
        "Changes saved detected after Product Description click cycle",
        driver,
        config,
    )
    log_event(
        "DONE",
        "Firefox opened under Selenium control. Close the browser window to stop the session.",
    )

    try:
        input("Press Enter here after you are done with the browser session...")
    finally:
        pause_controller.stop()
        driver.quit()


if __name__ == "__main__":
    main()
