"""Place photos into a photo frame of a template image, one output per photo.

Core logic only, no Streamlit dependency. The frame is an axis-aligned box in
template pixel coordinates; each photo is scaled to cover (or fit inside) the
box, cropped, optionally given rounded corners, and pasted onto a copy of the
template.
"""

import argparse
import math
import sys
from dataclasses import dataclass
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageOps

try:  # iPhone photos are often HEIC; support them when pillow-heif is installed.
    from pillow_heif import register_heif_opener

    register_heif_opener()
    HEIC_SUPPORTED = True
except ImportError:
    HEIC_SUPPORTED = False

PHOTO_EXTENSIONS = ["jpg", "jpeg", "png", "webp"] + (["heic", "heif"] if HEIC_SUPPORTED else [])

FIT_COVER = "cover"      # fill the whole frame, crop the overflow
FIT_CONTAIN = "contain"  # show the whole photo, leave empty bands
FITS = (FIT_COVER, FIT_CONTAIN)

# Minimum per-channel jump between neighbouring pixels to count as an edge.
# Low enough to catch a pale placeholder (light sky) on a white background.
EDGE_THRESHOLD = 20
DETECT_MAX_SIDE = 1000


@dataclass(frozen=True)
class Box:
    """Frame rectangle in template pixels; right/bottom are exclusive."""

    left: int
    top: int
    right: int
    bottom: int

    @property
    def width(self) -> int:
        return self.right - self.left

    @property
    def height(self) -> int:
        return self.bottom - self.top

    def validate(self, template_size: tuple[int, int]) -> None:
        w, h = template_size
        if not (0 <= self.left < self.right <= w and 0 <= self.top < self.bottom <= h):
            raise ValueError(f"Frame {self} is outside the template ({w}x{h})")

    @classmethod
    def parse(cls, text: str) -> "Box":
        """Parse ``left,top,right,bottom``."""
        parts = [int(p) for p in text.split(",")]
        if len(parts) != 4:
            raise ValueError("Box must be 'left,top,right,bottom'")
        return cls(*parts)


def _runs(mask_row: np.ndarray, min_len: int) -> list[tuple[int, int]]:
    """[start, end) of each run of True at least ``min_len`` long."""
    padded = np.concatenate(([False], mask_row, [False]))
    diff = np.diff(padded.astype(np.int8))
    starts = np.flatnonzero(diff == 1)
    ends = np.flatnonzero(diff == -1)
    return [(int(a), int(b)) for a, b in zip(starts, ends) if b - a >= min_len]


def _iou(a: Box, b: Box) -> float:
    ix = max(0, min(a.right, b.right) - max(a.left, b.left))
    iy = max(0, min(a.bottom, b.bottom) - max(a.top, b.top))
    inter = ix * iy
    union = a.width * a.height + b.width * b.height - inter
    return inter / union if union else 0.0


def _mostly_inside(inner: Box, outer: Box) -> bool:
    """True if at least half of ``inner`` lies within a larger ``outer``."""
    if outer.width * outer.height <= inner.width * inner.height:
        return False
    ix = max(0, min(inner.right, outer.right) - max(inner.left, outer.left))
    iy = max(0, min(inner.bottom, outer.bottom) - max(inner.top, outer.top))
    return ix * iy >= 0.5 * inner.width * inner.height


def detect_frames(template: Image.Image, max_results: int = 5) -> list[Box]:
    """Guess photo placeholders: rectangles bounded by long, straight, sharp edges.

    Works for placeholders with solid edges (a sample photo or a flat colour
    block); dashed decorative borders are ignored because their edges are
    broken. Rounded corners are tolerated. Largest first.
    """
    # Big Canva exports are slow to scan; detect on a reduced copy.
    scale = min(1.0, DETECT_MAX_SIDE / max(template.size))
    small = template
    if scale < 1.0:
        small = template.resize(
            (round(template.width * scale), round(template.height * scale)),
            Image.Resampling.BILINEAR,
        )
    a = np.asarray(small.convert("RGB"), dtype=np.int16)
    h, w = a.shape[:2]
    min_len = max(20, int(min(w, h) * 0.12))
    tol = max(4, int(min(w, h) * 0.03))  # corner radius / antialias slack

    # Edge at index i = jump between pixels i-2 and i (two pixels apart so a
    # one-pixel antialiased border still reads as one sharp edge).
    edge_y = np.zeros((h, w), bool)
    edge_y[2:] = np.abs(a[2:] - a[:-2]).max(axis=2) > EDGE_THRESHOLD
    edge_x = np.zeros((h, w), bool)
    edge_x[:, 2:] = np.abs(a[:, 2:] - a[:, :-2]).max(axis=2) > EDGE_THRESHOLD
    # A side may wander by a pixel along its length; read columns two wide.
    edge_x[:, :-1] |= edge_x[:, 1:]

    h_runs = [(y, x0, x1) for y in range(h) for x0, x1 in _runs(edge_y[y], min_len)]
    col_cum = np.vstack([np.zeros((1, w), np.int32), np.cumsum(edge_x, axis=0)])

    def find_side(target: int, top: int, bottom: int) -> int | None:
        """x nearest ``target`` whose column is an edge along most of top..bottom.

        Coverage rather than one unbroken run, so a side that fades towards
        the background colour for a stretch still counts.
        """
        t, b = top + tol, bottom - tol
        lo, hi = max(0, target - 2 * tol), min(w - 1, target + 2 * tol)
        coverage = (col_cum[b, lo:hi + 1] - col_cum[t, lo:hi + 1]) / (b - t)
        xs = [lo + i for i in np.flatnonzero(coverage >= 0.7)]
        return min(xs, key=lambda x: abs(x - target)) if xs else None

    found: list[Box] = []
    for i, (yt, a0, a1) in enumerate(h_runs):
        for yb, b0, b1 in h_runs[i + 1:]:
            if yb - yt < min_len or abs(a0 - b0) > tol or abs(a1 - b1) > tol:
                continue
            x0, x1 = min(a0, b0), max(a1, b1)
            left = find_side(x0, yt, yb)
            right = find_side(x1, yt, yb)
            if left is None or right is None or right - left < min_len:
                continue
            box = Box(left, yt, right, yb)
            if all(_iou(box, f) < 0.8 for f in found):
                found.append(box)

    # Drop rectangles sitting inside another one (straight lines in a photo).
    found = [b for b in found if not any(_mostly_inside(b, o) for o in found)]
    found.sort(key=lambda b: b.width * b.height, reverse=True)
    if scale < 1.0:
        found = [
            Box(*(min(round(v / scale), lim) for v, lim in zip(
                (b.left, b.top, b.right, b.bottom),
                (template.width, template.height) * 2,
            )))
            for b in found
        ]
    return found[:max_results]


def load_image(path_or_file) -> Image.Image:
    """Open an image and apply its EXIF rotation (phone photos)."""
    img = Image.open(path_or_file)
    img = ImageOps.exif_transpose(img)
    img.load()
    return img


def fit_photo(
    photo: Image.Image,
    size: tuple[int, int],
    fit: str = FIT_COVER,
    align: tuple[float, float] = (0.5, 0.5),
) -> Image.Image:
    """Resize ``photo`` to exactly ``size`` (RGBA).

    ``align`` is the (x, y) anchor in 0..1: for cover it picks which part of
    the overflow is kept (0 = left/top, 1 = right/bottom); for contain it
    places the photo inside the transparent canvas.
    """
    photo = photo.convert("RGBA")
    if fit == FIT_COVER:
        return ImageOps.fit(photo, size, Image.Resampling.LANCZOS, centering=align)
    if fit == FIT_CONTAIN:
        scaled = ImageOps.contain(photo, size, Image.Resampling.LANCZOS)
        canvas = Image.new("RGBA", size, (0, 0, 0, 0))
        x = round((size[0] - scaled.width) * align[0])
        y = round((size[1] - scaled.height) * align[1])
        canvas.paste(scaled, (x, y))
        return canvas
    raise ValueError(f"Unknown fit {fit!r}; expected one of {FITS}")


def compose(
    template: Image.Image,
    photo: Image.Image,
    box: Box,
    fit: str = FIT_COVER,
    align: tuple[float, float] = (0.5, 0.5),
    radius: int = 0,
) -> Image.Image:
    """Return a copy of ``template`` with ``photo`` placed in ``box``."""
    box.validate(template.size)
    out = template.convert("RGBA")
    tile = fit_photo(photo, (box.width, box.height), fit, align)
    mask = tile.getchannel("A")
    if radius > 0:
        mask = rounded_mask(mask, radius)
    out.paste(tile, (box.left, box.top), mask)
    return out


Point = tuple[float, float]


def order_quad(points: list[Point]) -> list[Point]:
    """Sort 4 clicked points into top-left, top-right, bottom-right, bottom-left."""
    if len(points) != 4:
        raise ValueError("A frame needs exactly 4 corners")
    cx = sum(p[0] for p in points) / 4
    cy = sum(p[1] for p in points) / 4
    # Clockwise on screen (y grows downwards), starting anywhere...
    pts = sorted(points, key=lambda p: math.atan2(p[1] - cy, p[0] - cx))
    # ...then start from the corner nearest the top-left.
    start = min(range(4), key=lambda i: pts[i][0] + pts[i][1])
    return pts[start:] + pts[:start]


def _perspective_coeffs(src: list[Point], dst: list[Point]) -> list[float]:
    """Coefficients for ``Image.transform(PERSPECTIVE)`` mapping dst points to src."""
    rows, rhs = [], []
    for (x, y), (u, v) in zip(dst, src):
        rows.append([x, y, 1, 0, 0, 0, -u * x, -u * y])
        rows.append([0, 0, 0, x, y, 1, -v * x, -v * y])
        rhs += [u, v]
    return np.linalg.solve(np.array(rows, float), np.array(rhs, float)).tolist()


def compose_quad(
    template: Image.Image,
    photo: Image.Image,
    quad: list[Point],
    fit: str = FIT_COVER,
    align: tuple[float, float] = (0.5, 0.5),
    radius: int = 0,
) -> Image.Image:
    """Place ``photo`` into a 4-corner frame, e.g. a tilted photo in the template.

    The photo is fitted to the frame's size, given rounded corners, then
    warped onto the quadrilateral (rotation, and any perspective, included).
    """
    tl, tr, br, bl = order_quad(quad)
    width = round((math.dist(tl, tr) + math.dist(bl, br)) / 2)
    height = round((math.dist(tl, bl) + math.dist(tr, br)) / 2)
    tile = fit_photo(photo, (width, height), fit, align)
    if radius > 0:
        tile.putalpha(rounded_mask(tile.getchannel("A"), radius))

    # Warp only over the frame's bounding box, then paste it there.
    xs, ys = [p[0] for p in (tl, tr, br, bl)], [p[1] for p in (tl, tr, br, bl)]
    left, top = math.floor(min(xs)), math.floor(min(ys))
    right, bottom = math.ceil(max(xs)), math.ceil(max(ys))
    local = [(x - left, y - top) for x, y in (tl, tr, br, bl)]
    corners = [(0, 0), (width, 0), (width, height), (0, height)]
    warped = tile.transform(
        (right - left, bottom - top), Image.Transform.PERSPECTIVE,
        _perspective_coeffs(corners, local), Image.Resampling.BICUBIC,
    )
    out = template.convert("RGBA")
    out.alpha_composite(warped, (left, top))
    return out


def rounded_mask(mask: Image.Image, radius: int) -> Image.Image:
    """Clear the corners of an ``L`` mask outside a rounded rectangle."""
    rounded = Image.new("L", mask.size, 0)
    ImageDraw.Draw(rounded).rounded_rectangle(
        (0, 0, mask.width - 1, mask.height - 1), radius=radius, fill=255
    )
    return Image.composite(mask, rounded, rounded)


def save_image(img: Image.Image, path: Path) -> Path:
    """Save as PNG or JPEG depending on the suffix (JPEG gets a white background)."""
    if path.suffix.lower() in (".jpg", ".jpeg"):
        bg = Image.new("RGB", img.size, "white")
        bg.paste(img, mask=img.getchannel("A") if img.mode == "RGBA" else None)
        bg.save(path, quality=95)
    else:
        img.save(path)
    return path


def save_pdf(images: list[Image.Image], path: Path) -> Path:
    """Save all images as one multi-page PDF (one page per photo)."""
    pages = []
    for img in images:
        page = Image.new("RGB", img.size, "white")
        page.paste(img, mask=img.getchannel("A") if img.mode == "RGBA" else None)
        pages.append(page)
    pages[0].save(path, save_all=True, append_images=pages[1:], resolution=150)
    return path


def output_name(photo_path: Path, ext: str) -> str:
    return f"{photo_path.stem}.{ext}"


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        description="Place each photo into the frame of a template image."
    )
    parser.add_argument("-t", "--template", required=True, type=Path, help="Template image")
    parser.add_argument(
        "-b", "--box", type=Box.parse,
        help="Frame in template pixels: left,top,right,bottom (see --detect)",
    )
    parser.add_argument(
        "--detect", action="store_true", help="Print detected frames and exit"
    )
    parser.add_argument(
        "-q", "--quad",
        type=lambda t: [tuple(map(float, p.split(","))) for p in t.split(";")],
        help="Tilted frame as 4 corners 'x,y;x,y;x,y;x,y' (any order); overrides --box",
    )
    parser.add_argument("-o", "--output", type=Path, default=Path("output"), help="Output dir")
    parser.add_argument("-f", "--format", choices=["png", "jpg"], default="jpg")
    parser.add_argument("--fit", choices=FITS, default=FIT_COVER)
    parser.add_argument("--align-x", type=float, default=0.5, help="0=left .. 1=right")
    parser.add_argument("--align-y", type=float, default=0.5, help="0=top .. 1=bottom")
    parser.add_argument("--radius", type=int, default=0, help="Corner radius in pixels")
    parser.add_argument("--pdf", action="store_true", help="Also write all pages to one PDF")
    parser.add_argument("photos", nargs="*", type=Path)
    args = parser.parse_args(argv)

    template = load_image(args.template)
    if args.detect:
        for i, b in enumerate(detect_frames(template), 1):
            print(f"{i}: {b.left},{b.top},{b.right},{b.bottom}  ({b.width}x{b.height})")
        return
    if (args.box is None and args.quad is None) or not args.photos:
        parser.error("--box or --quad, and at least one photo, are required (or use --detect)")
    if args.quad is None:
        args.box.validate(template.size)
    args.output.mkdir(parents=True, exist_ok=True)

    results = []
    for photo_path in args.photos:
        align = (args.align_x, args.align_y)
        if args.quad:
            img = compose_quad(template, load_image(photo_path), args.quad,
                               args.fit, align, args.radius)
        else:
            img = compose(template, load_image(photo_path), args.box,
                          args.fit, align, args.radius)
        out = save_image(img, args.output / output_name(photo_path, args.format))
        results.append(img)
        print(f"OK {photo_path.name} -> {out}")

    if args.pdf:
        print(f"PDF -> {save_pdf(results, args.output / 'tat_ca.pdf')}")


if __name__ == "__main__":
    sys.exit(main())
