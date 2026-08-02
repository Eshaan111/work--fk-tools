"""Tab-based Firefox runner for Full LC Auto.

This entry point reuses one Firefox process while consecutive queued work uses
the same laptop/profile.  Every listing run still starts in a fresh tab.  The
original ``main_ui_themed.py`` remains the restart-Firefox-per-run version.
"""

from __future__ import annotations

import gc
import traceback
from dataclasses import dataclass
from functools import wraps
from pathlib import Path
from typing import Callable

import main_ui_themed as base


_MISSING_LABEL_ERROR_PARTS = (
    "could not find label element for field:",
    "could not find field wrapper for label:",
    "could not find editable element for label:",
    "could not find tag input wrapper for field:",
    "could not find tag input element for field:",
)
_missing_label_skip_installed = False


def _is_missing_label_error(error: base.TimeoutException) -> bool:
    message = str(error).casefold()
    return any(part in message for part in _MISSING_LABEL_ERROR_PARTS)


def _skip_missing_labels_in_filler(
    filler: Callable[..., base.FillResult],
    log_stage: str,
) -> Callable[..., base.FillResult]:
    """Run a page filler field-by-field so an absent label skips only that value."""

    @wraps(filler)
    def wrapped(driver, field_definitions, product_input_row) -> base.FillResult:
        generated_values: dict[str, str] = {}
        skipped_fields: set[str] = set()

        for field in field_definitions:
            try:
                field_result = filler(driver, [field], product_input_row)
            except base.TimeoutException as error:
                if not _is_missing_label_error(error):
                    raise
                skipped_fields.add(field.label)
                base.log_event(
                    log_stage,
                    f"Skipping {field.label}: its label is not present in the current listing state.",
                )
                continue

            generated_values.update(field_result.generated_values)
            skipped_fields.update(field_result.skipped_fields)

        return base.FillResult(
            generated_values=generated_values,
            skipped_fields=skipped_fields,
        )

    return wrapped


def install_missing_label_skip_behavior() -> None:
    """Apply missing-label skipping only to runs launched by this entry point."""

    global _missing_label_skip_installed
    if _missing_label_skip_installed:
        return

    base.fill_price_stock_shipping_fields = _skip_missing_labels_in_filler(
        base.fill_price_stock_shipping_fields,
        "PRICE",
    )
    base.fill_product_description_fields = _skip_missing_labels_in_filler(
        base.fill_product_description_fields,
        "DESC",
    )
    base.fill_additional_description_fields = _skip_missing_labels_in_filler(
        base.fill_additional_description_fields,
        "ADDL",
    )
    _missing_label_skip_installed = True


class TabRunControl(base.RunControl):
    """Keep Abort Current Run tab-scoped; Abort Queue still closes Firefox."""

    def request_abort_current_run(self) -> None:
        self.abort_current_run_event.set()
        base.log_event(
            "RUN",
            "Current-run cancellation flag set; the active tab will close at the next safe point.",
        )


@dataclass
class ProfileBrowserSession:
    run_control: base.RunControl | None = None
    driver: object | None = None
    session_key: tuple[str, str] | None = None
    anchor_handle: str | None = None
    current_run_handle: str | None = None
    launch_number: int = 0

    def _key_for(self, startup_selection: base.StartupSelection, config: base.BotConfig) -> tuple[str, str]:
        return (
            startup_selection.laptop_name.strip().upper(),
            str(config.firefox_profile_path.resolve()).casefold(),
        )

    def is_healthy(self) -> bool:
        if self.driver is None:
            return False
        try:
            return bool(self.driver.window_handles)
        except Exception:
            return False

    def ensure(self, startup_selection: base.StartupSelection, config: base.BotConfig):
        requested_key = self._key_for(startup_selection, config)
        if self.session_key != requested_key or not self.is_healthy():
            if self.driver is not None:
                reason = "profile changed" if self.session_key != requested_key else "browser session is no longer healthy"
                base.log_event("BROWSER", f"Restarting Firefox because the {reason}.")
                self.close()
            self._launch(config, requested_key)
        else:
            base.log_event(
                "BROWSER",
                f"Reusing Firefox for laptop={requested_key[0]}, profile={config.profile_name}.",
            )
        return self.driver

    def _launch(self, config: base.BotConfig, requested_key: tuple[str, str]) -> None:
        self.launch_number += 1
        log_path = base.get_geckodriver_log_path(self.launch_number)
        base.log_event(
            "BROWSER",
            f"Launching Firefox for profile {config.profile_name} (browser launch {self.launch_number}).",
        )
        driver = base.build_firefox_driver(config, geckodriver_log_path=log_path)
        self.driver = driver
        self.session_key = requested_key
        if self.run_control is not None:
            setattr(driver, "_full_lc_run_control", self.run_control)
            self.run_control.set_active_driver(driver)
        try:
            driver.maximize_window()
            driver.get("about:blank")
            self.anchor_handle = driver.current_window_handle
        except Exception:
            self.close()
            raise
        self.current_run_handle = None
        base.log_event("BROWSER", "Firefox launched; the original tab is reserved as the session anchor.")

    def open_run_tab(self, run_index: int, total_runs: int):
        if not self.is_healthy():
            raise base.WebDriverException("The shared Firefox session is not available.")
        self.driver.switch_to.new_window("tab")
        self.current_run_handle = self.driver.current_window_handle
        base.log_event("TAB", f"Opened a fresh working tab for run {run_index}/{total_runs}.")
        return self.driver

    def close_run_tab(self, run_index: int, total_runs: int) -> None:
        driver = self.driver
        run_handle = self.current_run_handle
        self.current_run_handle = None
        if driver is None or run_handle is None:
            return
        try:
            handles = driver.window_handles
            if run_handle in handles:
                driver.switch_to.window(run_handle)
                driver.close()
            remaining_handles = driver.window_handles
            target = self.anchor_handle if self.anchor_handle in remaining_handles else remaining_handles[0]
            self.anchor_handle = target
            driver.switch_to.window(target)
            base.log_event("TAB", f"Closed the working tab for run {run_index}/{total_runs}.")
        except Exception as exc:
            base.log_event("TAB", f"Could not close the working tab cleanly: {exc}")
            if not self.is_healthy():
                self.retire_broken_session()

    def retire_broken_session(self) -> None:
        if self.driver is None:
            return
        base.log_event("BROWSER", "Retiring an unusable shared Firefox session.")
        self.close()

    def close(self) -> None:
        driver = self.driver
        self.driver = None
        self.session_key = None
        self.anchor_handle = None
        self.current_run_handle = None
        if driver is None:
            return
        try:
            base.quit_webdriver_safely(driver)
        finally:
            base.untrack_webdriver(driver)
            if self.run_control is not None:
                self.run_control.clear_active_driver(driver)
        base.log_event("BROWSER", "Closed the shared Firefox session.")


def execute_listing_in_tab(
    session: ProfileBrowserSession,
    config: base.BotConfig,
    listing_selection: base.ListingSelection,
    json_flow_definition: base.FlowDefinition | None,
    run_index: int,
    total_runs: int,
    run_control: base.RunControl | None,
) -> base.JobSessionResult:
    base.set_current_run_label(f"run {run_index}/{total_runs}")
    result = base.JobSessionResult(run_index=run_index, total_runs=total_runs, succeeded=False)

    if run_control is not None:
        run_control.check_abort()

    try:
        driver = session.open_run_tab(run_index, total_runs)
    except Exception as exc:
        message = f"Run {run_index}/{total_runs} failed while opening a fresh tab: {exc}"
        base.write_latest_error(message)
        base.log_event("ERROR", message)
        result.error_message = message
        session.retire_broken_session()
        return result

    pause_controller = base.PauseController(run_control=run_control)
    pause_controller.start()
    try:
        if run_control is not None:
            run_control.check_abort()
        base.log_event("RUN", f"Starting run {run_index} of {total_runs} in a fresh tab.")

        if json_flow_definition is None:
            base.log_event("NAV", f"Opening listing page: {config.listing_url}")
            base.open_listing_page(driver, config.listing_url)
            base.log_event("NAV", "Listing page opened in Firefox.")
            base.checkpoint_pause(pause_controller, "Listing page opened", driver, config)
            base.dismiss_optional_ad_popup(driver)
            base.checkpoint_pause(pause_controller, "Optional popup handling complete", driver, config)
            base.fill_brand_name(driver, listing_selection.brand_name)
            base.checkpoint_pause(pause_controller, "Brand entered", driver, config)
            base.click_create_new_listing(driver)
            base.checkpoint_pause(pause_controller, "Create new listing clicked", driver, config)
            base.click_optional_continue(driver)
            base.checkpoint_pause(pause_controller, "Optional continue handling complete", driver, config)

        flow_state = base.run_listing_page_flow(driver, pause_controller, config, listing_selection)
        if base.flow_definition_saves_and_exits(json_flow_definition):
            base.log_event("DONE", "JSON flow handled the final listing action.")
        else:
            base.click_final_listing_action_button(driver, config)
            base.commit_pending_image_folder_exhaustion(flow_state)
        base.log_event("DONE", f"{listing_selection.product_type.title()} flow completed for run {run_index}/{total_runs}.")

        final_action = base.resolve_final_listing_action(config.final_listing_action)
        if final_action == "send_to_qc":
            base.wait_for_changes_saved_toast_appearances(
                driver, pause_controller, config, required_appearances=1
            )
            base.log_event(
                "TAB",
                f"Waiting {base.SUCCESS_CLOSE_DELAY_SECONDS} seconds after the Send to QC success toast before completing this tab.",
            )
        else:
            base.log_event(
                "TAB",
                f"Waiting {base.SUCCESS_CLOSE_DELAY_SECONDS} seconds before completing this tab.",
            )
        base.wait_before_browser_shutdown(base.SUCCESS_CLOSE_DELAY_SECONDS, run_control)
        result.succeeded = True
        return result
    except base.RunAbortRequested as exc:
        message = str(exc)
        base.write_latest_error(message)
        base.log_event("RUN", f"Worker acknowledged cancellation: {message}")
        result.error_message = message
        return result
    except Exception as exc:
        if run_control is not None and (
            run_control.should_abort_current_run() or run_control.should_abort_batch()
        ):
            message = "Batch aborted by user." if run_control.should_abort_batch() else "Current run aborted by user."
            base.write_latest_error(message)
            base.log_event("RUN", message)
            result.error_message = message
            return result
        try:
            snapshot = base.save_html_snapshot(
                driver, config.snapshot_directory, f"run {run_index} error before tab close"
            )
            result.snapshot_path = snapshot
            base.log_event("ERROR", f"Saved failure snapshot before closing tab: {snapshot}")
        except Exception as snapshot_error:
            base.log_event("ERROR", f"Could not save failure snapshot: {snapshot_error}")
        message = f"Run {run_index}/{total_runs} failed: {exc}"
        base.write_latest_error(message)
        base.log_event("ERROR", message)
        base.log_event("RUN", "Aborting this run; the next run will use a fresh tab.")
        result.error_message = message
        return result
    finally:
        pause_controller.stop()
        session.close_run_tab(run_index, total_runs)
        gc.collect()


def run_job_tab_based(
    startup_selection: base.StartupSelection,
    session: ProfileBrowserSession,
    run_control: base.RunControl | None = None,
    progress_callback: Callable[[base.JobSessionResult], None] | None = None,
) -> base.JobRunResult:
    install_missing_label_skip_behavior()
    config, flow_definition = base.build_bot_config(startup_selection)
    listing = startup_selection.listing_selection
    base.run_login_precheck(config, listing, flow_definition)
    base.print_runtime_context(config)
    base.log_event("BROWSER", "Browser mode: reuse Firefox and open a fresh tab per run.")

    completed = 0
    failed = 0
    results: list[base.JobSessionResult] = []

    try:
        session.ensure(startup_selection, config)
    except Exception as exc:
        message = f"Could not launch Firefox for profile {config.profile_name}: {exc}"
        base.write_latest_error(message)
        base.log_event("ERROR", message)
        for run_index in range(1, startup_selection.run_count + 1):
            item = base.JobSessionResult(
                run_index=run_index,
                total_runs=startup_selection.run_count,
                succeeded=False,
                error_message=message,
                launch_failed_before_browser=True,
            )
            results.append(item)
            failed += 1
            if progress_callback is not None:
                progress_callback(item)
    else:
        for run_index in range(1, startup_selection.run_count + 1):
            if run_control is not None and run_control.should_abort_batch():
                break
            if not session.is_healthy():
                try:
                    session.ensure(startup_selection, config)
                except Exception as exc:
                    message = f"Could not restart Firefox before run {run_index}: {exc}"
                    item = base.JobSessionResult(
                        run_index=run_index,
                        total_runs=startup_selection.run_count,
                        succeeded=False,
                        error_message=message,
                        launch_failed_before_browser=True,
                    )
                    results.append(item)
                    failed += 1
                    if progress_callback is not None:
                        progress_callback(item)
                    break

            item = execute_listing_in_tab(
                session,
                config,
                listing,
                flow_definition,
                run_index,
                startup_selection.run_count,
                run_control,
            )
            if item.succeeded:
                completed += 1
                try:
                    record_path = base.record_successful_run(listing, startup_selection.profile_name)
                    item.success_record_path = record_path
                    base.log_event("DONE", f"Recorded successful run in Excel: {record_path.name}")
                except Exception as exc:
                    item.success_record_error = str(exc)
                    base.log_event("ERROR", f"Could not update successful run record Excel: {exc}")
            else:
                failed += 1
            results.append(item)
            if progress_callback is not None:
                progress_callback(item)
            if run_control is not None:
                if run_control.should_abort_batch():
                    break
                if run_control.should_abort_current_run():
                    base.log_event("RUN", "Current run abort completed. Continuing in a fresh tab.")
                    run_control.finish_current_run()

    base.set_current_run_label("summary")
    base.log_event("DONE", f"Batch finished. Successful run(s): {completed}. Failed run(s): {failed}.")
    job_result = base.JobRunResult(
        config=config,
        listing_selection=listing,
        json_flow_definition=flow_definition,
        completed_runs=completed,
        failed_runs=failed,
        session_results=results,
    )
    result_path = base.write_job_run_result(job_result)
    base.log_event("DONE", f"Saved batch result JSON: {result_path.name}")
    return job_result


def run_queued_jobs_tab_based(
    startup_selections: list[base.StartupSelection],
    run_control: base.RunControl | None = None,
    progress_callback: Callable[[base.JobSessionResult], None] | None = None,
) -> list[base.JobRunResult]:
    queue_results: list[base.JobRunResult] = []
    session = ProfileBrowserSession(run_control=run_control)
    total_items = len(startup_selections)
    try:
        for queue_index, selection in enumerate(startup_selections, start=1):
            if run_control is not None and run_control.should_abort_batch():
                break
            listing = selection.listing_selection
            base.set_current_run_label(f"queue {queue_index}/{total_items}")
            base.log_event(
                "QUEUE",
                f"Starting tab-based item {queue_index}/{total_items}: "
                f"laptop={selection.laptop_name}, account={selection.profile_name}, "
                f"vertical={listing.product_type}, surface={listing.surface}, runs={selection.run_count}.",
            )
            try:
                base.set_active_laptop(selection.laptop_name)
                queue_results.append(
                    run_job_tab_based(selection, session, run_control, progress_callback)
                )
            except Exception as exc:
                message = f"Queue item {queue_index}/{total_items} failed: {exc}"
                base.write_latest_error(message)
                base.log_event("ERROR", message)
                base.log_event("ERROR", traceback.format_exc())
                for run_index in range(1, selection.run_count + 1):
                    if progress_callback is not None:
                        progress_callback(
                            base.JobSessionResult(
                                run_index=run_index,
                                total_runs=selection.run_count,
                                succeeded=False,
                                error_message=message,
                            )
                        )
            if run_control is not None and run_control.should_abort_batch():
                break
            base.log_event("QUEUE", f"Finished item {queue_index}/{total_items}.")
    finally:
        session.close()

    base.set_current_run_label("queue summary")
    base.log_event(
        "QUEUE",
        f"Tab-based queue finished. Completed {len(queue_results)} of {total_items} item(s).",
    )
    return queue_results


def main() -> None:
    # The shared monitor resolves these names from main_ui_themed at runtime.
    install_missing_label_skip_behavior()
    base.RunControl = TabRunControl
    base.run_queued_jobs = run_queued_jobs_tab_based
    base.main()


if __name__ == "__main__":
    main()
