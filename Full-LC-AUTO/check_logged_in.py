from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path
import sys
import tkinter as tk
from typing import Any

from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from main import (
    BotConfig,
    DEFAULT_LISTING_URL,
    build_firefox_driver,
    log_event,
    prompt_for_profile,
)


@dataclass(slots=True)
class LoginCheckResult:
    status: str
    reason: str
    profile_name: str
    requested_url: str
    final_url: str
    page_title: str
    opened_requested_url: bool
    opened_login_page: bool
    login_popup_opened: bool
    username_filled: bool
    next_clicked: bool
    password_filled: bool
    login_clicked: bool
    gmail_tab_opened: bool
    login_confirmation_acknowledged: bool


LOGIN_URL_KEYWORDS = ("login", "signin", "sign-in", "auth")
LOGIN_PAGE_TITLES = {
    "become an online seller in india | flipkart seller hub",
}
LOGIN_PAGE_TEXT_MARKERS = (
    "become an online seller in india",
    "flipkart seller hub",
)
LOGIN_PAGE_BUTTON_LOCATORS = (
    (
        By.XPATH,
        "//button[@data-testid='button' and normalize-space(.)='Login']",
    ),
    (
        By.XPATH,
        "//button[contains(@class,'ButtonStyle') and normalize-space(.)='Login']",
    ),
)
LOGIN_MODAL_LOCATORS = (
    (By.XPATH, "//h4[contains(@class,'modal-title') and normalize-space(.)='Login']"),
    (By.XPATH, "//input[@name='username' and @placeholder='Username or phone number or email']"),
    (By.XPATH, "//section[contains(@class,'modal-body-section')]"),
)
USERNAME_INPUT_LOCATORS = (
    (
        By.XPATH,
        "//input[@name='username' and @placeholder='Username or phone number or email']",
    ),
    (
        By.XPATH,
        "//input[contains(@class,'login') and @data-testid='test-input']",
    ),
)
NEXT_BUTTON_LOCATORS = (
    (
        By.XPATH,
        "//button[@data-testid='button' and .//span[normalize-space(.)='Next'] and not(@disabled)]",
    ),
    (
        By.XPATH,
        "//button[.//span[normalize-space(.)='Next'] and not(@disabled)]",
    ),
)
LOGIN_SUBMIT_BUTTON_LOCATORS = (
    (
        By.XPATH,
        "//button[@data-testid='button' and .//span[normalize-space(.)='Login'] and not(@disabled)]",
    ),
    (
        By.XPATH,
        "//button[.//span[normalize-space(.)='Login'] and not(@disabled)]",
    ),
)
PASSWORD_INPUT_LOCATORS = (
    (
        By.XPATH,
        "//input[@name='password' and @placeholder='Enter password']",
    ),
    (
        By.XPATH,
        "//input[@type='password' and contains(@class,'password')]",
    ),
)
PROFILE_EMAILS = {
    "prabhu": "rajeshbansal201973@gmail.com",
    "seema": "seemabansal091976@gmail.com",
}
PROJECT_ROOT = Path(__file__).resolve().parent
LOGGED_OUT_LOCATORS = (
    (By.XPATH, "//input[@type='password']"),
    (By.XPATH, "//input[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'password')]"),
    (By.XPATH, "//input[contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'password')]"),
    (By.XPATH, "//input[contains(translate(@name,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'otp')]"),
    (By.XPATH, "//input[contains(translate(@placeholder,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'otp')]"),
    (By.XPATH, "//button[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'login')]"),
    (By.XPATH, "//button[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'sign in')]"),
)
LOGGED_IN_LOCATORS = (
    (By.XPATH, "//*[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'dashboard')]"),
    (By.XPATH, "//*[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'add listings')]"),
    (By.XPATH, "//*[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'catalog')]"),
    (By.XPATH, "//*[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'orders')]"),
)


def prompt_with_default(prompt: str, default: str) -> str:
    selected_value = input(f"{prompt} [{default}]: ").strip()
    return selected_value or default


def wait_for_page_settle(driver, timeout_seconds: int = 20) -> None:
    WebDriverWait(driver, timeout_seconds).until(
        lambda current_driver: current_driver.execute_script("return document.readyState") == "complete"
    )


def first_matching_locator(driver, locators: tuple[tuple[str, str], ...], timeout_seconds: int = 2):
    for locator in locators:
        try:
            return WebDriverWait(driver, timeout_seconds).until(
                EC.presence_of_element_located(locator)
            )
        except TimeoutException:
            continue
    return None


def normalize_url(url: str) -> str:
    return url.strip().rstrip("/")


def is_login_page(driver) -> bool:
    current_url = (driver.current_url or "").lower()
    title = " ".join((driver.title or "").strip().lower().split())
    page_source = (driver.page_source or "").lower()

    if any(keyword in current_url for keyword in LOGIN_URL_KEYWORDS):
        return True
    if title in LOGIN_PAGE_TITLES:
        return True
    return all(marker in page_source for marker in LOGIN_PAGE_TEXT_MARKERS)


def did_open_requested_url(requested_url: str, final_url: str) -> bool:
    normalized_requested = normalize_url(requested_url)
    normalized_final = normalize_url(final_url)
    return (
        normalized_final == normalized_requested
        or normalized_final.startswith(normalized_requested + "#")
        or normalized_final.startswith(normalized_requested + "?")
    )


def click_when_clickable(driver, locators: tuple[tuple[str, str], ...], timeout_seconds: int = 10) -> bool:
    for locator in locators:
        try:
            element = WebDriverWait(driver, timeout_seconds).until(
                EC.element_to_be_clickable(locator)
            )
            driver.execute_script("arguments[0].click();", element)
            return True
        except TimeoutException:
            continue
    return False


def is_login_popup_open(driver) -> bool:
    return first_matching_locator(driver, LOGIN_MODAL_LOCATORS, timeout_seconds=2) is not None


def open_login_popup(driver, timeout_seconds: int = 10) -> bool:
    if is_login_popup_open(driver):
        return True
    if not click_when_clickable(driver, LOGIN_PAGE_BUTTON_LOCATORS, timeout_seconds=timeout_seconds):
        return False
    return is_login_popup_open(driver)


def fill_username_for_profile(driver, profile_name: str, timeout_seconds: int = 10) -> bool:
    profile_email = PROFILE_EMAILS.get(profile_name.strip().lower())
    if not profile_email:
        raise ValueError(f"No login email configured for profile '{profile_name}'.")

    input_element = first_matching_locator(driver, USERNAME_INPUT_LOCATORS, timeout_seconds=timeout_seconds)
    if input_element is None:
        return False

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_element)
    input_element.click()
    input_element.send_keys(Keys.CONTROL, "a")
    input_element.send_keys(Keys.DELETE)
    input_element.send_keys(profile_email)
    return (input_element.get_attribute("value") or "").strip().lower() == profile_email.lower()


def click_next_button(driver, timeout_seconds: int = 10) -> bool:
    return click_when_clickable(driver, NEXT_BUTTON_LOCATORS, timeout_seconds=timeout_seconds)


def click_login_submit_button(driver, timeout_seconds: int = 10) -> bool:
    return click_when_clickable(driver, LOGIN_SUBMIT_BUTTON_LOCATORS, timeout_seconds=timeout_seconds)


def load_dotenv_values(env_path: Path | None = None) -> dict[str, str]:
    resolved_path = env_path or (PROJECT_ROOT / ".env")
    if not resolved_path.exists():
        return {}

    values: dict[str, str] = {}
    for raw_line in resolved_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key:
            values[key] = value
    return values


def get_profile_password(profile_name: str) -> str:
    normalized_profile = profile_name.strip().lower()
    env_key = f"{normalized_profile.upper()}_PASSWORD"
    password = os.getenv(env_key)
    if password:
        return password

    dotenv_values = load_dotenv_values()
    password = dotenv_values.get(env_key)
    if password:
        return password

    raise ValueError(f"No password found for profile '{profile_name}'. Expected env key {env_key}.")


def fill_password_for_profile(driver, profile_name: str, timeout_seconds: int = 10) -> bool:
    password_value = get_profile_password(profile_name)
    input_element = first_matching_locator(driver, PASSWORD_INPUT_LOCATORS, timeout_seconds=timeout_seconds)
    if input_element is None:
        return False

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_element)
    input_element.click()
    input_element.send_keys(Keys.CONTROL, "a")
    input_element.send_keys(Keys.DELETE)
    input_element.send_keys(password_value)
    return (input_element.get_attribute("value") or "") == password_value


def open_gmail_in_new_tab(driver, timeout_seconds: int = 20) -> bool:
    existing_handles = list(driver.window_handles)
    driver.switch_to.new_window("tab")
    WebDriverWait(driver, timeout_seconds).until(
        lambda current_driver: len(current_driver.window_handles) == len(existing_handles) + 1
    )
    driver.get("https://www.gmail.com")
    wait_for_page_settle(driver, timeout_seconds=timeout_seconds)
    current_url = (driver.current_url or "").lower()
    return "gmail.com" in current_url or "accounts.google.com" in current_url


def show_logged_in_confirmation_dialog() -> bool:
    acknowledged = False
    root = tk.Tk()
    root.title("Login Required")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    def handle_ok() -> None:
        nonlocal acknowledged
        acknowledged = True
        root.destroy()

    frame = tk.Frame(root, padx=20, pady=16)
    frame.pack()
    tk.Label(frame, text="Please click OK when logged in.").pack(pady=(0, 12))
    tk.Button(frame, text="OK", width=12, command=handle_ok).pack()
    root.protocol("WM_DELETE_WINDOW", handle_ok)
    root.mainloop()
    return acknowledged


def infer_login_state(driver) -> LoginCheckResult:
    current_url = (driver.current_url or "").lower()
    title = (driver.title or "").strip()
    opened_login_page = is_login_page(driver)

    if opened_login_page:
        return LoginCheckResult(
            "logged_out",
            "browser opened the saved-login-page equivalent instead of the requested seller page",
            "",
            "",
            driver.current_url,
            title,
            False,
            True,
            False,
            False,
            False,
            False,
            False,
            False,
            False,
        )

    logged_out_element = first_matching_locator(driver, LOGGED_OUT_LOCATORS)
    if logged_out_element is not None:
        return LoginCheckResult(
            "logged_out",
            f"detected login control: <{logged_out_element.tag_name}>",
            "",
            "",
            driver.current_url,
            title,
            False,
            False,
            False,
            False,
            False,
            False,
            False,
            False,
            False,
        )

    logged_in_element = first_matching_locator(driver, LOGGED_IN_LOCATORS)
    if logged_in_element is not None:
        matched_text = " ".join(logged_in_element.text.split()) or logged_in_element.tag_name
        return LoginCheckResult(
            "logged_in",
            f"detected seller page content: {matched_text}",
            "",
            "",
            driver.current_url,
            title,
            False,
            False,
            False,
            False,
            False,
            False,
            False,
            False,
            False,
        )

    if "seller.flipkart.com" in current_url:
        return LoginCheckResult(
            "likely_logged_in",
            f"seller domain opened without obvious auth redirect (title: {title or 'n/a'})",
            "",
            "",
            driver.current_url,
            title,
            False,
            False,
            False,
            False,
            False,
            False,
            False,
            False,
            False,
        )

    return LoginCheckResult(
        "unknown",
        f"could not confidently classify page (url: {driver.current_url}, title: {title or 'n/a'})",
        "",
        "",
        driver.current_url,
        title,
        False,
        False,
        False,
        False,
        False,
        False,
        False,
        False,
        False,
    )


def check_logged_in(
    target_url: str,
    profile_name: str = "prabhu",
    *,
    close_browser: bool = True,
    prompt_after_gmail: bool = True,
) -> LoginCheckResult:
    config = BotConfig(profile_name=profile_name)
    driver = build_firefox_driver(config)

    try:
        log_event("BOOT", f"Opening URL for profile '{profile_name}': {target_url}")
        driver.maximize_window()
        driver.get(target_url)
        wait_for_page_settle(driver)

        result = infer_login_state(driver)
        final_url = driver.current_url
        opened_requested_url = did_open_requested_url(target_url, final_url)
        opened_login_page = is_login_page(driver)
        login_popup_opened = False
        username_filled = False
        next_clicked = False
        password_filled = False
        login_clicked = False
        gmail_tab_opened = False
        login_confirmation_acknowledged = False
        if opened_login_page:
            login_popup_opened = open_login_popup(driver)
            if login_popup_opened:
                username_filled = fill_username_for_profile(driver, profile_name)
                if username_filled:
                    next_clicked = click_next_button(driver)
                    status = "opened_login_page"
                    if next_clicked:
                        password_filled = fill_password_for_profile(driver, profile_name)
                        if password_filled:
                            login_clicked = click_login_submit_button(driver)
                            if login_clicked:
                                gmail_tab_opened = open_gmail_in_new_tab(driver)
                                if gmail_tab_opened:
                                    if prompt_after_gmail:
                                        login_confirmation_acknowledged = show_logged_in_confirmation_dialog()
                                    reason = (
                                        "requested URL redirected to the login page, the full login flow was advanced, "
                                        "and Gmail opened in a new tab"
                                    )
                                else:
                                    reason = "requested URL redirected to the login page, login was clicked, but Gmail did not open in the new tab"
                            else:
                                reason = "requested URL redirected to the login page, password was filled, but the Login button was not clicked"
                        else:
                            reason = "requested URL redirected to the login page, login popup opened, username was filled, Next was clicked, but password was not filled"
                    else:
                        reason = "requested URL redirected to the login page, login popup opened, username was filled, but Next was not clicked"
                else:
                    status = "opened_login_page"
                    reason = "requested URL redirected to the login page, login popup opened, but username was not filled"
            else:
                status = "opened_login_page"
                reason = "requested URL redirected to the login page, but the login popup did not open"
        elif opened_requested_url:
            status = "opened_requested_url"
            reason = "requested URL opened successfully"
        else:
            status = result.status
            reason = result.reason
        return LoginCheckResult(
            status=status,
            reason=reason,
            profile_name=profile_name,
            requested_url=target_url,
            final_url=final_url,
            page_title=driver.title,
            opened_requested_url=opened_requested_url,
            opened_login_page=opened_login_page,
            login_popup_opened=login_popup_opened,
            username_filled=username_filled,
            next_clicked=next_clicked,
            password_filled=password_filled,
            login_clicked=login_clicked,
            gmail_tab_opened=gmail_tab_opened,
            login_confirmation_acknowledged=login_confirmation_acknowledged,
        )
    except WebDriverException as exc:
        raise RuntimeError(f"Could not open the browser or target URL: {exc}") from exc
    finally:
        if close_browser:
            driver.quit()


def check_logged_in_as_dict(
    target_url: str,
    profile_name: str = "prabhu",
    *,
    close_browser: bool = True,
    prompt_after_gmail: bool = True,
) -> dict[str, Any]:
    result = check_logged_in(
        target_url,
        profile_name,
        close_browser=close_browser,
        prompt_after_gmail=prompt_after_gmail,
    )
    return {
        "status": result.status,
        "reason": result.reason,
        "profile_name": result.profile_name,
        "requested_url": result.requested_url,
        "final_url": result.final_url,
        "page_title": result.page_title,
        "opened_requested_url": result.opened_requested_url,
        "opened_login_page": result.opened_login_page,
        "login_popup_opened": result.login_popup_opened,
        "username_filled": result.username_filled,
        "next_clicked": result.next_clicked,
        "password_filled": result.password_filled,
        "login_clicked": result.login_clicked,
        "gmail_tab_opened": result.gmail_tab_opened,
        "login_confirmation_acknowledged": result.login_confirmation_acknowledged,
    }


def parse_cli_args(argv: list[str]) -> tuple[str | None, str | None, bool]:
    profile_name: str | None = None
    target_url: str | None = None
    close_browser = True

    index = 0
    while index < len(argv):
        arg = argv[index]
        if arg == "--profile" and index + 1 < len(argv):
            profile_name = argv[index + 1]
            index += 2
            continue
        if arg == "--url" and index + 1 < len(argv):
            target_url = argv[index + 1]
            index += 2
            continue
        if arg == "--keep-browser-open":
            close_browser = False
            index += 1
            continue
        raise SystemExit(f"Unknown or incomplete argument: {arg}")

    return profile_name, target_url, close_browser


def main() -> None:
    cli_profile_name, cli_target_url, close_browser = parse_cli_args(sys.argv[1:])
    print("--- CHECK LOGGED IN ---")
    if cli_profile_name is None:
        profile_name = prompt_for_profile()
    else:
        profile_name = cli_profile_name

    if cli_target_url is None:
        target_url = prompt_with_default("Enter desired URL", DEFAULT_LISTING_URL)
        should_close = input("Close browser? [y/N]: ").strip().lower()
        close_browser = should_close in {"y", "yes"}
    else:
        target_url = cli_target_url

    try:
        result = check_logged_in(
            target_url,
            profile_name,
            close_browser=close_browser,
            prompt_after_gmail=True,
        )
    except RuntimeError as exc:
        raise SystemExit(str(exc)) from exc

    print()
    print(f"Profile: {result.profile_name}")
    print(f"Requested URL: {result.requested_url}")
    print(f"Final URL: {result.final_url}")
    print(f"Title: {result.page_title}")
    print(f"Opened requested URL: {result.opened_requested_url}")
    print(f"Opened login page: {result.opened_login_page}")
    print(f"Login popup opened: {result.login_popup_opened}")
    print(f"Username filled: {result.username_filled}")
    print(f"Next clicked: {result.next_clicked}")
    print(f"Password filled: {result.password_filled}")
    print(f"Login clicked: {result.login_clicked}")
    print(f"Gmail tab opened: {result.gmail_tab_opened}")
    print(f"Login confirmation acknowledged: {result.login_confirmation_acknowledged}")
    print(f"Status: {result.status}")
    print(f"Reason: {result.reason}")


if __name__ == "__main__":
    main()
