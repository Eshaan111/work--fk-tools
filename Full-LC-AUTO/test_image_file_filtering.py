from pathlib import Path

from main_ui_themed import collect_ordered_images


def test_collect_ordered_images_ignores_non_image_files(tmp_path: Path) -> None:
    for name in ("1.png", "2.JPG", "3.jpeg", "4.webp", "5.bmp"):
        (tmp_path / name).write_bytes(b"image")
    for name in ("metadata.json", "notes.txt", "archive.zip", "image.svg"):
        (tmp_path / name).write_text("not an uploadable image", encoding="utf-8")

    assert [path.name for path in collect_ordered_images(tmp_path)] == [
        "1.png",
        "2.JPG",
        "3.jpeg",
        "4.webp",
        "5.bmp",
    ]
