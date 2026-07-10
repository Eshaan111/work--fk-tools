from __future__ import annotations

import json
import re
import shutil
import threading
import uuid
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from app_paths import get_app_root


FOLDER_NUMBER_RE = re.compile(r"^(\d+)(?:-|$)")


@dataclass(frozen=True)
class Destination:
    laptop: str
    vertical: str
    kind: str
    path: Path

    @property
    def label(self) -> str:
        return f"{self.vertical.replace('_', ' ').title()} — {self.kind}"


def load_destinations(config_path: Path) -> dict[str, list[Destination]]:
    with config_path.open("r", encoding="utf-8") as config_file:
        payload = json.load(config_file)

    result: dict[str, list[Destination]] = {}
    for laptop, laptop_config in payload.get("laptops", {}).items():
        destinations: list[Destination] = []
        seen: set[tuple[str, str, str]] = set()
        for vertical, vertical_config in laptop_config.get("verticals", {}).items():
            for kind_config in vertical_config.get("kinds", {}).values():
                kind = str(kind_config.get("kind", "")).strip()
                raw_path = str(kind_config.get("image_directory", "")).strip()
                if not kind or not raw_path:
                    continue
                path = Path(raw_path).expanduser()
                identity = (vertical.casefold(), kind.casefold(), str(path).casefold())
                if identity in seen:
                    continue
                seen.add(identity)
                destinations.append(Destination(laptop, vertical, kind, path))
        result[laptop] = destinations
    return result


def highest_folder_number(destination: Path) -> int:
    highest = -1
    for child in destination.iterdir():
        if not child.is_dir():
            continue
        match = FOLDER_NUMBER_RE.match(child.name)
        if match:
            highest = max(highest, int(match.group(1)))
    return highest


def copy_numbered_folders(sources: list[Path], destination: Path) -> list[Path]:
    if not destination.is_dir():
        raise FileNotFoundError(f"Destination folder does not exist:\n{destination}")
    if not sources:
        raise ValueError("Select at least one source folder.")

    resolved_destination = destination.resolve()
    for source in sources:
        if not source.is_dir():
            raise FileNotFoundError(f"Source folder does not exist:\n{source}")
        resolved_source = source.resolve()
        if resolved_source == resolved_destination or resolved_destination in resolved_source.parents:
            raise ValueError(f"A source folder cannot be the destination or contain it:\n{source}")

    next_number = highest_folder_number(destination) + 1
    copied: list[Path] = []
    for source in sources:
        while (destination / str(next_number)).exists():
            next_number += 1
        target = destination / str(next_number)
        temporary = destination / f".copying-{next_number}-{uuid.uuid4().hex}"
        try:
            shutil.copytree(source, temporary, copy_function=shutil.copy2)
            temporary.rename(target)
        except Exception:
            if temporary.exists():
                shutil.rmtree(temporary, ignore_errors=True)
            raise
        copied.append(target)
        next_number += 1
    return copied


class CopyUtilityApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Listing Image Folder Copier")
        self.root.minsize(720, 510)
        self.destinations = load_destinations(get_app_root() / "config.json")
        if not self.destinations:
            raise ValueError("No laptop image destinations were found in config.json.")
        self.sources: list[Path] = []

        self.laptop_var = tk.StringVar()
        self.destination_var = tk.StringVar()
        self.path_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Select source folders and a destination.")
        self._build_ui()

        laptop_names = list(self.destinations)
        preferred = "ASUS" if "ASUS" in laptop_names else laptop_names[0]
        self.laptop_var.set(preferred)
        self._refresh_destinations()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Copy folders into Listing Images", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(frame, text="New folders are numbered after the highest existing folder number; gaps are not filled.").pack(anchor="w", pady=(2, 16))

        controls = ttk.Frame(frame)
        controls.pack(fill="x")
        ttk.Label(controls, text="Laptop").grid(row=0, column=0, sticky="w")
        self.laptop_box = ttk.Combobox(controls, textvariable=self.laptop_var, values=list(self.destinations), state="readonly", width=17)
        self.laptop_box.grid(row=1, column=0, sticky="ew", padx=(0, 12))
        self.laptop_box.bind("<<ComboboxSelected>>", lambda _event: self._refresh_destinations())
        ttk.Label(controls, text="Folder kind").grid(row=0, column=1, sticky="w")
        self.destination_box = ttk.Combobox(controls, textvariable=self.destination_var, state="readonly", width=40)
        self.destination_box.grid(row=1, column=1, sticky="ew")
        self.destination_box.bind("<<ComboboxSelected>>", lambda _event: self._show_destination_path())
        controls.columnconfigure(1, weight=1)

        ttk.Label(frame, textvariable=self.path_var, foreground="#555555", wraplength=680).pack(anchor="w", pady=(8, 16))

        source_header = ttk.Frame(frame)
        source_header.pack(fill="x")
        ttk.Label(source_header, text="Source folders", font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Button(source_header, text="Add folders…", command=self._add_folders).pack(side="right")
        ttk.Button(source_header, text="Remove selected", command=self._remove_selected).pack(side="right", padx=8)

        self.source_list = tk.Listbox(frame, selectmode="extended", height=12)
        self.source_list.pack(fill="both", expand=True, pady=(8, 12))

        footer = ttk.Frame(frame)
        footer.pack(fill="x")
        ttk.Label(footer, textvariable=self.status_var).pack(side="left", fill="x", expand=True)
        self.copy_button = ttk.Button(footer, text="Copy and number folders", command=self._start_copy)
        self.copy_button.pack(side="right")

    def _current_destinations(self) -> list[Destination]:
        return self.destinations.get(self.laptop_var.get(), [])

    def _refresh_destinations(self) -> None:
        options = self._current_destinations()
        labels = [destination.label for destination in options]
        self.destination_box["values"] = labels
        self.destination_var.set(labels[0] if labels else "")
        self._show_destination_path()

    def _selected_destination(self) -> Destination:
        index = self.destination_box.current()
        options = self._current_destinations()
        if index < 0 or index >= len(options):
            raise ValueError("Select a destination folder kind.")
        return options[index]

    def _show_destination_path(self) -> None:
        try:
            self.path_var.set(str(self._selected_destination().path))
        except ValueError:
            self.path_var.set("No destination configured.")

    def _add_folders(self) -> None:
        while True:
            selected = filedialog.askdirectory(title="Select a source folder (Cancel when finished)", mustexist=True)
            if not selected:
                break
            path = Path(selected)
            if path not in self.sources:
                self.sources.append(path)
                self.source_list.insert("end", str(path))
            if not messagebox.askyesno("Add another?", "Do you want to select another source folder?"):
                break
        self.status_var.set(f"{len(self.sources)} source folder(s) selected.")

    def _remove_selected(self) -> None:
        for index in reversed(self.source_list.curselection()):
            del self.sources[index]
            self.source_list.delete(index)
        self.status_var.set(f"{len(self.sources)} source folder(s) selected.")

    def _start_copy(self) -> None:
        try:
            destination = self._selected_destination()
            if not self.sources:
                raise ValueError("Select at least one source folder.")
        except Exception as exc:
            messagebox.showerror("Cannot copy", str(exc))
            return

        sources = list(self.sources)
        if not messagebox.askyesno("Confirm copy", f"Copy {len(sources)} folder(s) to:\n\n{destination.path}\n\nThe original folders will remain unchanged."):
            return
        self.copy_button.configure(state="disabled")
        self.status_var.set("Copying folders…")
        threading.Thread(target=self._copy_worker, args=(sources, destination.path), daemon=True).start()

    def _copy_worker(self, sources: list[Path], destination: Path) -> None:
        try:
            copied = copy_numbered_folders(sources, destination)
        except Exception as exc:
            self.root.after(0, self._copy_failed, str(exc))
        else:
            self.root.after(0, self._copy_finished, copied)

    def _copy_failed(self, error: str) -> None:
        self.copy_button.configure(state="normal")
        self.status_var.set("Copy failed.")
        messagebox.showerror("Copy failed", error)

    def _copy_finished(self, copied: list[Path]) -> None:
        self.copy_button.configure(state="normal")
        self.sources.clear()
        self.source_list.delete(0, "end")
        names = ", ".join(path.name for path in copied)
        self.status_var.set(f"Copied successfully as: {names}")
        messagebox.showinfo("Copy complete", f"Copied {len(copied)} folder(s) successfully.\n\nNew folder names: {names}")


def main() -> None:
    root = tk.Tk()
    try:
        CopyUtilityApp(root)
    except Exception as exc:
        root.withdraw()
        messagebox.showerror("Startup error", str(exc))
        root.destroy()
        return
    root.mainloop()


if __name__ == "__main__":
    main()
