#!/usr/bin/env python3
"""
Windows folder-picker utility.

Workflow:
1. Pick the SOURCE parent folder.
2. Pick the DESTINATION parent folder.
3. Copy all immediately contained numbered source folders into the destination.
4. Continue numbering after the highest numbered destination folder.
5. Rename each successfully copied source folder to:
       <original-number>-copied

Only folders whose names contain digits only are processed.
"""

from __future__ import annotations

import shutil
import sys  
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox


def get_numbered_directories(parent: Path) -> list[Path]:
    """Return immediate child directories whose names contain only digits."""
    return sorted(
        (
            item
            for item in parent.iterdir()
            if item.is_dir() and item.name.isdigit()
        ),
        key=lambda path: int(path.name),
    )


def get_last_destination_number(destination_parent: Path) -> int:
    """Return the highest numbered destination directory, or 0 if none exist."""
    numbered_directories = get_numbered_directories(destination_parent)

    if not numbered_directories:
        return 0

    return max(int(directory.name) for directory in numbered_directories)


def get_available_copied_name(source_directory: Path) -> Path:
    """
    Return an available '<original>-copied' folder path.

    If it already exists, use:
        <original>-copied-2
        <original>-copied-3
        ...
    """
    base_name = f"{source_directory.name}-copied"
    candidate = source_directory.with_name(base_name)

    counter = 2
    while candidate.exists():
        candidate = source_directory.with_name(f"{base_name}-{counter}")
        counter += 1

    return candidate


def choose_folder(title: str) -> Path | None:
    """
    Open a native Windows folder picker.

    tkinter.filedialog.askdirectory selects directories only.
    It does not allow selecting individual files.
    """
    selected = filedialog.askdirectory(
        title=title,
        mustexist=True,
    )

    if not selected:
        return None

    return Path(selected).resolve()


def copy_and_renumber(
    source_parent: Path,
    destination_parent: Path,
) -> tuple[int, list[str]]:
    if source_parent == destination_parent:
        raise ValueError(
            "The source and destination parent folders must be different."
        )

    source_directories = get_numbered_directories(source_parent)

    if not source_directories:
        raise ValueError(
            "The selected source folder contains no numbered subfolders."
        )

    last_destination_number = get_last_destination_number(destination_parent)
    next_destination_number = last_destination_number + 1

    completed = 0
    operation_log: list[str] = []

    for source_directory in source_directories:
        destination_directory = (
            destination_parent / str(next_destination_number)
        )

        if destination_directory.exists():
            raise FileExistsError(
                f"Destination folder already exists: {destination_directory}"
            )

        renamed_source_directory = get_available_copied_name(source_directory)

        try:
            shutil.copytree(source_directory, destination_directory)
        except Exception:
            # Remove an incomplete copy when possible.
            if destination_directory.exists():
                shutil.rmtree(destination_directory, ignore_errors=True)
            raise

        try:
            source_directory.rename(renamed_source_directory)
        except Exception as error:
            raise RuntimeError(
                "The folder was copied successfully, but the original source "
                f"folder could not be renamed.\n\nCopied folder:\n"
                f"{destination_directory}\n\nSource folder:\n"
                f"{source_directory}\n\nReason: {error}"
            ) from error

        operation_log.append(
            f"{source_directory.name}  →  "
            f"{destination_directory.name}  →  "
            f"{renamed_source_directory.name}"
        )

        completed += 1
        next_destination_number += 1

    return completed, operation_log


def main() -> int:
    root = tk.Tk()
    root.withdraw()
    root.update()

    messagebox.showinfo(
        "Select Source Folder",
        "First, select the SOURCE parent folder.\n\n"
        "This folder should contain numbered subfolders such as "
        "1, 2, 3, 4, and so on.",
        parent=root,
    )

    source_parent = choose_folder("Select SOURCE Parent Folder")

    if source_parent is None:
        messagebox.showinfo(
            "Cancelled",
            "No source folder was selected.",
            parent=root,
        )
        root.destroy()
        return 0

    messagebox.showinfo(
        "Select Destination Folder",
        "Now select the DESTINATION parent folder.\n\n"
        "Copied folders will be added after its highest existing number.",
        parent=root,
    )

    destination_parent = choose_folder("Select DESTINATION Parent Folder")

    if destination_parent is None:
        messagebox.showinfo(
            "Cancelled",
            "No destination folder was selected.",
            parent=root,
        )
        root.destroy()
        return 0

    try:
        source_directories = get_numbered_directories(source_parent)

        if not source_directories:
            raise ValueError(
                "The selected source folder contains no numbered subfolders."
            )

        last_number = get_last_destination_number(destination_parent)
        first_new_number = last_number + 1
        final_new_number = first_new_number + len(source_directories) - 1

        preview_lines = [
            f"{folder.name}  →  {first_new_number + index}"
            for index, folder in enumerate(source_directories)
        ]

        preview = "\n".join(preview_lines)

        confirmed = messagebox.askyesno(
            "Confirm Folder Copy",
            f"SOURCE:\n{source_parent}\n\n"
            f"DESTINATION:\n{destination_parent}\n\n"
            f"Numbered source folders found: {len(source_directories)}\n"
            f"New destination numbers: "
            f"{first_new_number} to {final_new_number}\n\n"
            f"Planned operations:\n{preview}\n\n"
            "After each successful copy, the original source folder will be "
            "renamed to '<number>-copied'.\n\n"
            "Continue?",
            parent=root,
        )

        if not confirmed:
            messagebox.showinfo(
                "Cancelled",
                "No folders were copied or renamed.",
                parent=root,
            )
            root.destroy()
            return 0

        completed, operation_log = copy_and_renumber(
            source_parent,
            destination_parent,
        )

        messagebox.showinfo(
            "Completed",
            f"Successfully processed {completed} folder(s).\n\n"
            + "\n".join(operation_log),
            parent=root,
        )

    except Exception as error:
        messagebox.showerror(
            "Operation Failed",
            str(error),
            parent=root,
        )
        root.destroy()
        return 1

    root.destroy()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
