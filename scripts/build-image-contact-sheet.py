from __future__ import annotations

import argparse
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps


def build_contact_sheet(
    input_dir: Path,
    output_path: Path,
    *,
    columns: int = 3,
    thumbnail_width: int = 360,
) -> None:
    paths = sorted(
        input_dir.glob("page-*.png"),
        key=lambda path: int(path.stem.rsplit("-", 1)[-1]),
    )
    if not paths:
        raise ValueError(f"No page-*.png files found in {input_dir}")

    gap = 24
    label_height = 36
    font = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 20)
    thumbnails: list[Image.Image] = []
    for path in paths:
        with Image.open(path) as source:
            image = source.convert("RGB")
            height = round(image.height * thumbnail_width / image.width)
            thumbnails.append(ImageOps.contain(image, (thumbnail_width, height)))

    cell_height = max(image.height for image in thumbnails) + label_height
    rows = math.ceil(len(thumbnails) / columns)
    sheet = Image.new(
        "RGB",
        (
            columns * thumbnail_width + (columns + 1) * gap,
            rows * cell_height + (rows + 1) * gap,
        ),
        (225, 229, 235),
    )
    draw = ImageDraw.Draw(sheet)
    for index, image in enumerate(thumbnails):
        column = index % columns
        row = index // columns
        x = gap + column * (thumbnail_width + gap)
        label_y = gap + row * cell_height
        y = label_y + label_height
        draw.text((x, label_y), f"第 {index + 1} 页", font=font, fill=(20, 30, 45))
        sheet.paste(image, (x, y))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(output_path, format="PNG", optimize=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a contact sheet from page-N.png previews.")
    parser.add_argument("input_dir", type=Path)
    parser.add_argument("output_path", type=Path)
    parser.add_argument("--columns", type=int, default=3)
    parser.add_argument("--thumbnail-width", type=int, default=360)
    args = parser.parse_args()
    build_contact_sheet(
        args.input_dir.resolve(),
        args.output_path.resolve(),
        columns=max(1, args.columns),
        thumbnail_width=max(120, args.thumbnail_width),
    )
    print(args.output_path.resolve())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
