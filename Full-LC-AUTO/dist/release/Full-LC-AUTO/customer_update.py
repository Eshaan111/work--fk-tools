from __future__ import annotations

import hashlib
import json
import shutil
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from app_paths import get_app_root

APP_ROOT = get_app_root()
UPDATES_ROOT = APP_ROOT / "customer_updates"
BUNDLES_ROOT = UPDATES_ROOT / "bundles"
BACKUPS_ROOT = UPDATES_ROOT / "backups"
REPORTS_ROOT = UPDATES_ROOT / "reports"
MANIFEST_NAME = "manifest.json"

ALLOWED_SUFFIXES = {
    ".json",
    ".xlsx",
    ".xlsm",
    ".xls",
    ".csv",
    ".tsv",
}


def now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file_handle:
        for chunk in iter(lambda: file_handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def ensure_directory(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def normalize_target(path: Path) -> Path:
    resolved = path.resolve()
    app_root_resolved = APP_ROOT.resolve()
    if resolved == app_root_resolved:
        raise ValueError("Update target cannot be the app root itself.")
    try:
        resolved.relative_to(app_root_resolved)
    except ValueError as exc:
        raise ValueError(f"Target must stay inside app folder: {path}") from exc
    return resolved


def to_portable(path: Path) -> str:
    return path.as_posix()


def collect_supported_files(path: Path) -> list[Path]:
    normalized = normalize_target(path)
    if normalized.is_file():
        if normalized.suffix.lower() not in ALLOWED_SUFFIXES:
            raise ValueError(f"Unsupported file type: {normalized.name}")
        return [normalized]
    if not normalized.is_dir():
        raise ValueError(f"Path does not exist: {normalized}")
    files = [
        child.resolve()
        for child in normalized.rglob("*")
        if child.is_file() and child.suffix.lower() in ALLOWED_SUFFIXES
    ]
    if not files:
        raise ValueError(f"No supported files found in: {normalized}")
    return sorted(files)


def build_manifest_entry(file_path: Path) -> dict[str, object]:
    relative_path = file_path.relative_to(APP_ROOT.resolve())
    return {
        "target": to_portable(relative_path),
        "archive_path": f"payload/{to_portable(relative_path)}",
        "size": file_path.stat().st_size,
        "sha256": sha256_file(file_path),
        "modified_at": datetime.fromtimestamp(file_path.stat().st_mtime).isoformat(timespec="seconds"),
    }


def create_update_bundle(bundle_path: Path, source_paths: list[Path], note: str, author: str) -> tuple[Path, dict[str, object]]:
    ensure_directory(bundle_path.parent)
    unique_files: dict[Path, dict[str, object]] = {}
    for source_path in source_paths:
        for file_path in collect_supported_files(source_path):
            unique_files[file_path] = build_manifest_entry(file_path)

    manifest = {
        "bundle_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "created_by": author.strip(),
        "note": note.strip(),
        "file_count": len(unique_files),
        "entries": [entry for _, entry in sorted(unique_files.items(), key=lambda item: to_portable(item[0].relative_to(APP_ROOT.resolve())))],
    }

    with zipfile.ZipFile(bundle_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(MANIFEST_NAME, json.dumps(manifest, indent=2))
        for file_path, entry in unique_files.items():
            archive.write(file_path, arcname=str(entry["archive_path"]))

    return bundle_path, manifest


@dataclass
class ApplyResult:
    bundle_path: Path
    report_path: Path
    backup_root: Path
    applied_targets: list[str]


def apply_update_bundle(bundle_path: Path) -> ApplyResult:
    ensure_directory(BACKUPS_ROOT)
    ensure_directory(REPORTS_ROOT)
    bundle_path = bundle_path.resolve()
    stamp = now_stamp()
    backup_root = BACKUPS_ROOT / stamp
    report_path = REPORTS_ROOT / f"apply_report_{stamp}.txt"
    applied_targets: list[str] = []

    with zipfile.ZipFile(bundle_path, "r") as archive:
        manifest = json.loads(archive.read(MANIFEST_NAME).decode("utf-8"))
        entries = manifest.get("entries", [])
        if not isinstance(entries, list) or not entries:
            raise ValueError("Bundle manifest does not contain any entries.")

        report_lines = [
            f"Bundle: {bundle_path}",
            f"Applied at: {datetime.now().isoformat(timespec='seconds')}",
            f"Created at: {manifest.get('created_at', '')}",
            f"Created by: {manifest.get('created_by', '')}",
            f"Note: {manifest.get('note', '')}",
            "",
            "Applied files:",
        ]

        for entry in entries:
            target_value = str(entry["target"])
            archive_path = str(entry["archive_path"])
            expected_hash = str(entry["sha256"])
            target_path = normalize_target(APP_ROOT / Path(target_value))
            ensure_directory(target_path.parent)

            if target_path.exists():
                backup_path = backup_root / Path(target_value)
                ensure_directory(backup_path.parent)
                shutil.copy2(target_path, backup_path)

            with archive.open(archive_path) as source, target_path.open("wb") as destination:
                shutil.copyfileobj(source, destination)

            actual_hash = sha256_file(target_path)
            if actual_hash != expected_hash:
                raise ValueError(f"Hash verification failed for {target_value}")

            applied_targets.append(target_value)
            report_lines.append(f"- {target_value}")

        report_lines.extend(["", f"Backup root: {backup_root}"])
        report_path.write_text("\n".join(report_lines) + "\n", encoding="utf-8")

    return ApplyResult(
        bundle_path=bundle_path,
        report_path=report_path,
        backup_root=backup_root,
        applied_targets=applied_targets,
    )


class UpdateCenterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Full LC Auto Update Center")
        self.root.geometry("980x700")

        self.author_var = tk.StringVar()
        self.note_var = tk.StringVar()
        self.bundle_var = tk.StringVar()
        self.selected_paths: list[Path] = []

        self.build_ui()
        self.refresh_selected_paths()

    def build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        outer = ttk.Frame(self.root, padding=16)
        outer.grid(sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(3, weight=1)

        ttk.Label(outer, text="Customer Update Bundles", font=("Segoe UI", 18, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(
            outer,
            text="Create update zip files for customer JSON and Excel changes, or apply a received update with automatic backups.",
            wraplength=900,
        ).grid(row=1, column=0, sticky="w", pady=(6, 16))

        meta = ttk.LabelFrame(outer, text="Bundle Info", padding=12)
        meta.grid(row=2, column=0, sticky="ew")
        meta.columnconfigure(1, weight=1)
        ttk.Label(meta, text="Author").grid(row=0, column=0, sticky="w")
        ttk.Entry(meta, textvariable=self.author_var).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ttk.Label(meta, text="Note").grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(meta, textvariable=self.note_var).grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(10, 0))

        body = ttk.Panedwindow(outer, orient="horizontal")
        body.grid(row=3, column=0, sticky="nsew", pady=(16, 0))

        left = ttk.Frame(body, padding=8)
        right = ttk.Frame(body, padding=8)
        body.add(left, weight=3)
        body.add(right, weight=2)

        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)
        ttk.Label(left, text="Selected Targets", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w")
        self.selected_box = tk.Text(left, height=20, wrap="word")
        self.selected_box.grid(row=1, column=0, sticky="nsew", pady=(8, 12))

        left_actions = ttk.Frame(left)
        left_actions.grid(row=2, column=0, sticky="ew")
        ttk.Button(left_actions, text="Add Files", command=self.add_files).pack(side="left")
        ttk.Button(left_actions, text="Add Folder", command=self.add_folder).pack(side="left", padx=(8, 0))
        ttk.Button(left_actions, text="Clear", command=self.clear_selected).pack(side="left", padx=(8, 0))
        ttk.Button(left_actions, text="Create Bundle", command=self.create_bundle).pack(side="right")

        right.columnconfigure(0, weight=1)
        right.rowconfigure(4, weight=1)
        ttk.Label(right, text="Apply Update", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(right, text="Bundle Zip").grid(row=1, column=0, sticky="w", pady=(12, 0))
        bundle_row = ttk.Frame(right)
        bundle_row.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        bundle_row.columnconfigure(0, weight=1)
        ttk.Entry(bundle_row, textvariable=self.bundle_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(bundle_row, text="Browse", command=self.browse_bundle).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(right, text="Apply Bundle", command=self.apply_bundle).grid(row=3, column=0, sticky="ew", pady=(14, 0))

        info = tk.Text(right, height=18, wrap="word")
        info.grid(row=4, column=0, sticky="nsew", pady=(16, 0))
        info.insert(
            "1.0",
            "\n".join(
                [
                    f"App root: {APP_ROOT}",
                    "",
                    "Recommended remote workflow:",
                    "1. You change the customer's JSON or Excel files in your admin copy.",
                    "2. Create an update bundle zip.",
                    "3. Send that zip to the customer.",
                    "4. Customer opens Update Center and applies it.",
                    "5. The tool backs up previous files before replacing them.",
                ]
            ),
        )
        info.configure(state="disabled")

    def refresh_selected_paths(self) -> None:
        self.selected_box.configure(state="normal")
        self.selected_box.delete("1.0", "end")
        if not self.selected_paths:
            self.selected_box.insert("1.0", "No files or folders selected yet.")
        else:
            self.selected_box.insert("1.0", "\n".join(str(path) for path in self.selected_paths))
        self.selected_box.configure(state="disabled")

    def add_files(self) -> None:
        targets = filedialog.askopenfilenames(
            parent=self.root,
            title="Select JSON or Excel files",
            filetypes=[
                ("Supported files", "*.json *.xlsx *.xlsm *.xls *.csv *.tsv"),
                ("All files", "*.*"),
            ],
            initialdir=str(APP_ROOT),
        )
        for target in targets:
            path = Path(target)
            if path not in self.selected_paths:
                self.selected_paths.append(path)
        self.refresh_selected_paths()

    def add_folder(self) -> None:
        target = filedialog.askdirectory(parent=self.root, title="Select folder", initialdir=str(APP_ROOT))
        if target:
            path = Path(target)
            if path not in self.selected_paths:
                self.selected_paths.append(path)
        self.refresh_selected_paths()

    def clear_selected(self) -> None:
        self.selected_paths.clear()
        self.refresh_selected_paths()

    def create_bundle(self) -> None:
        if not self.selected_paths:
            messagebox.showerror("No targets", "Select at least one file or folder first.", parent=self.root)
            return
        ensure_directory(BUNDLES_ROOT)
        suggested = BUNDLES_ROOT / f"customer_update_{now_stamp()}.zip"
        target = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save update bundle",
            defaultextension=".zip",
            initialfile=suggested.name,
            initialdir=str(BUNDLES_ROOT),
            filetypes=[("Zip files", "*.zip")],
        )
        if not target:
            return
        try:
            bundle_path, manifest = create_update_bundle(
                Path(target),
                self.selected_paths,
                self.note_var.get(),
                self.author_var.get(),
            )
        except Exception as exc:
            messagebox.showerror("Bundle creation failed", str(exc), parent=self.root)
            return
        messagebox.showinfo(
            "Bundle created",
            f"Saved bundle:\n{bundle_path}\n\nFiles included: {manifest['file_count']}",
            parent=self.root,
        )

    def browse_bundle(self) -> None:
        target = filedialog.askopenfilename(
            parent=self.root,
            title="Select update bundle",
            initialdir=str(BUNDLES_ROOT if BUNDLES_ROOT.exists() else APP_ROOT),
            filetypes=[("Zip files", "*.zip")],
        )
        if target:
            self.bundle_var.set(target)

    def apply_bundle(self) -> None:
        bundle_value = self.bundle_var.get().strip()
        if not bundle_value:
            messagebox.showerror("No bundle", "Choose a bundle zip first.", parent=self.root)
            return
        try:
            result = apply_update_bundle(Path(bundle_value))
        except Exception as exc:
            messagebox.showerror("Apply failed", str(exc), parent=self.root)
            return
        messagebox.showinfo(
            "Update applied",
            f"Updated {len(result.applied_targets)} file(s).\n\nBackup: {result.backup_root}\nReport: {result.report_path}",
            parent=self.root,
        )


def main() -> None:
    root = tk.Tk()
    UpdateCenterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
