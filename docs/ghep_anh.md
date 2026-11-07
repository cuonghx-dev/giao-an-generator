# Ghép ảnh hàng loạt

Places each photo into the photo frame of a template image and produces one
page per photo — a lightweight replacement for Canva's "Bulk create" when only
the photo changes between pages.

## Input

- **Template**: the design exported from Canva as PNG/JPG, with the photo
  frame left as a placeholder (Canva's sample sky/grass picture or a flat
  colour block). Export at the size you want to print; the output keeps the
  template's resolution.
- **Photos**: JPG, PNG, WEBP, and HEIC (iPhone) when `pillow-heif` is
  installed. EXIF rotation is applied.

## Web UI

Page `/ghep-anh`:

1. Upload the template.
2. Pick the frame. The tool detects rectangles with straight, sharp edges and
   numbers them on the preview; pick one, or choose "Tự chỉnh" and drag the
   sliders. Fine-tune any choice in "Chỉnh vị trí khung".
   Then pick the corner style: "Vuông" (square) or "Bo góc" (rounded) with a
   radius slider; adjust it until the photo's corners match the frame in the
   preview.
   For a **tilted frame** (e.g. a reference design that already holds a
   rotated photo), choose "Khung nghiêng (bấm 4 góc)" and click the 4
   corners of that photo on the template, in any order ("↺ Chọn lại" to
   start over). Each new photo is fitted to the frame's size, rounded if
   asked, and warped onto those 4 corners, so it takes exactly the old
   photo's place and tilt.
3. Upload the photos. The first one is previewed in the frame.
4. Options: fill the frame (crop, default) or show the whole photo; which part
   to keep when cropping (top/centre/bottom, left/centre/right); JPG or PNG;
   an extra multi-page PDF for printing.
5. Generate, then download a zip (one file per photo, named after the photo)
   and/or the PDF.

## CLI

```bash
# Find the frame coordinates
python3 -m tools.ghep_anh -t template.png --detect
# 1: 125,190,468,406  (343x216)
# 2: 64,484,293,795  (229x311)

# Place every photo into frame 2
python3 -m tools.ghep_anh -t template.png -b 64,484,293,795 -o out --pdf photos/*.jpg

# Tilted frame: its 4 corners, any order
python3 -m tools.ghep_anh -t template.png -q "840,1050;1390,1120;1300,1830;740,1740" \
    --radius 25 -o out photos/*.jpg
```

Options: `-q/--quad` (tilted frame, overrides `-b`), `-f png|jpg`,
`--fit cover|contain`, `--align-x/--align-y 0..1`, `--radius PX`, `--pdf`.

## Frame detection

`detect_frames` looks for pairs of long horizontal edges with matching
extents, then for vertical edges near their ends that cover most of the
height. Dashed decorative borders are ignored (their edges are broken), and
rectangles lying mostly inside a bigger one (lines inside a photo) are
dropped. Large templates are scanned at 1000 px on the long side and the
result is scaled back, so a detected box can be off by a pixel or two — check
the preview and nudge the sliders if needed.
