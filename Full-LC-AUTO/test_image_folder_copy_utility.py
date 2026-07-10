from pathlib import Path

from image_folder_copy_utility import copy_numbered_folders, highest_folder_number


def test_highest_number_ignores_gaps_and_reads_brand_suffixes(tmp_path: Path) -> None:
    for name in ("0", "1-FB-B", "3", "6-KOT-P", "notes"):
        (tmp_path / name).mkdir()
    assert highest_folder_number(tmp_path) == 6


def test_copy_multiple_folders_consecutively(tmp_path: Path) -> None:
    destination = tmp_path / "destination"
    destination.mkdir()
    (destination / "0").mkdir()
    (destination / "2-FB-B").mkdir()
    first = tmp_path / "first"
    second = tmp_path / "second"
    first.mkdir()
    second.mkdir()
    (first / "one.jpg").write_bytes(b"one")
    (second / "two.jpg").write_bytes(b"two")

    copied = copy_numbered_folders([first, second], destination)

    assert [path.name for path in copied] == ["3", "4"]
    assert (destination / "3" / "one.jpg").read_bytes() == b"one"
    assert (destination / "4" / "two.jpg").read_bytes() == b"two"
    assert first.exists() and second.exists()
