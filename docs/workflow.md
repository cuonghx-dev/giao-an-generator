# Giáo Án Generator - Workflow Documentation

## Overview

`generate_giao_an_v2.py` builds lesson plan files from two inputs:

- **Schedule file** (`.xlsx`) — one sheet per class, a Mon–Fri timetable
  (e.g. `Tuần 1_Thế giới rực rỡ sắc màu-Tuần 2 tháng 7.xlsx`)
- **Content file** (`.xlsx`) — a single "KH chi tiết" workbook holding the
  detailed content of every activity of every week

```
Tuần 1....xlsx                      KH chi tiết.xlsx
  (one sheet per class)               (all weeks, all activities)
          |                                    |
          v                                    v
   Parse the class timetable          Parse the week's activity blocks
   - dates, theme, teacher            - one block per activity
   - 6 time slots per day             - split "Hoạt động" into phases
   - skip ESL/routine slots             + Đánh giá + Phân hóa
          |                                    |
          +------------------+-----------------+
                             v
              Match schedule entry -> activity
              (category first, then fuzzy name)
                             |
                             v
              One sheet per weekday: header,
              summary table, detail table, footer
                             |
                             v
                 output/giao_an_<Lớp>.xlsx
```

---

## Step 1: Parse the content workbook

**Function:** `parse_content(content_path, week)`

The workbook has a single sheet laid out top-to-bottom by week, and
left-to-right by activity:

```
row 1   TUẦN 1                                       <- week header (merged A:X)
row 2   HOẠT ĐỘNG CHÍNH                              <- category header (merged A:X)
row 3   HOẠT ĐỘNG 1: ... | Hoạt động 2: ... | ...    <- 5 activity titles
row 4   GV SOẠN | Tiêu chuẩn | Mục tiêu bài | ...    <- field labels
row 5   <values>                                      <- one data row per block
row 6   TRÒ CHUYỆN ĐẦU NGÀY                          <- next category
...
```

Blocks sit side by side and each spans about 6 columns (A:F, G:L, M:R, S:Y,
Z:AE). The exact spans are read from the merged ranges of the title row, so a
block that is one column wider or narrower still parses.

### How rows are classified

| Row | Detected by |
|-----|-------------|
| Week header | Column A matches `TUẦN <n>` and the rest of the row is empty |
| Category header | Column A has a value and the rest of the row is empty |
| Field labels | Column A is `GV SOẠN` or `Giáo viên soạn` |
| Activity titles | The row above the field labels |
| Values | The row below the field labels |

Field labels are inconsistent between sections (the same column is
`Tiêu chuẩn` in one section and `Giáo viên chuẩn bị` in another), so values
are mapped **by label**, not by column position:

| Label in the content file | Output column |
|---------------------------|---------------|
| `Tiêu chuẩn` | B — CHUẨN/MỤC TIÊU CHƯƠNG |
| `Mục tiêu bài` | C — MỤC TIÊU BÀI, and the summary table |
| `Tiêu chí thành công` | D — TIÊU CHÍ THÀNH CÔNG |
| `Tài liệu học tập` / `Đồ dùng` | E — TÀI LIỆU HỌC TẬP, CHUẨN BỊ |
| `Hoạt động` | F, H, I — see below |
| `Giáo viên chuẩn bị` | E, when the block has no materials column |
| `GV SOẠN` / `Giáo viên soạn` | teacher name (not written to the output) |

Blocks with an empty title or no `Hoạt động` text are skipped — the content
file reserves 5 blocks per category even when only 2 are filled in.

### Splitting the "Hoạt động" field

One cell holds the whole lesson: numbered phases with the `Đánh giá` and
`Phân hóa` sections written in wherever the author put them. `split_steps()`
pulls the three parts back out.

**Sections** start at a line beginning with `Đánh giá` or `Phân hóa`
(any case, `Phân hoá` too) and run until the next section or the next phase.
The heading may run straight into its first sentence
(`Đánh giá: Trẻ có tập trung lắng nghe...`).

**Phases** — each becomes one row of the activity block in the output. A
numbered line starts a new phase only when it also mentions a phase keyword
(`mở đầu`, `gây hứng thú`, `khởi động`, `ổn định`, `hoạt động chính`,
`giảng dạy chính`, `giới thiệu hoạt động`, `làm mẫu`, `kết thúc`, `củng cố`).

Without that keyword rule, a morning-chat lesson written as
`1. Chào hỏi` / `2. Điểm danh` / `3. Hoạt động trọng tâm` / `4. Kết thúc`
would produce four rows instead of the intended two.

---

## Step 2: Parse the schedule

**Function:** `parse_schedule(schedule_path, class_name)`

Each class sheet is a Mon–Fri timetable. Days map to columns C–G; the six
teaching slots are row pairs `(8,9) (10,11) (12,13) (14,15) (21,22) (23,24)`,
where the first row is the activity type and the second is its name.

| Field | Source | Example |
|-------|--------|---------|
| Dates | Row 2 col B | `Từ ngày 13/7 đến 17/7/2026` |
| Theme | Row 3 col B | `Chủ đề/Theme: Thế giới rực rỡ sắc màu (Tuần 1)` |
| Week number | Row 3 col B, the `(Tuần n)` part | `1` |
| Teacher | Row 3 col G, first name of the list | `Cao Hương` |

The workbook also carries sheets that are not timetables (evidence matrices,
review notes). `get_all_classes()` keeps only sheets with a week line in B2.

### Skip rules

A slot is dropped when its type or name contains a foreign-teacher marker
(`Story time`, `ESL Class`), generic play (`Chơi tự do`, `Softplay`,
`phòng Lego`, `phòng Cosplay`), or a routine (`Đón trẻ`, `Ăn trưa`, `Ngủ`,
`Vệ sinh`, `Uống sữa`, `Tái hiện`, `Dọn dẹp`), or when its type is not one of
the known categories.

A type cell may carry a qualifier on a second line —
`Hoạt động tập thể\n  Vivokids/thể chất`. Only the first line names the
category, so `clean_type()` keeps that and drops the rest.

### Category mapping

| Schedule type | Content category |
|---------------|------------------|
| Trò chuyện đầu ngày | TRÒ CHUYỆN ĐẦU NGÀY |
| Hoạt động chính | HOẠT ĐỘNG CHÍNH |
| Hoạt động Tiếng Việt | TIẾNG VIỆT |
| Chơi theo chủ đề, Hoạt động vui chơi | CHƠI THEO CHỦ ĐỀ |
| Hoạt động ngoài trời | HOẠT ĐỘNG NGOÀI TRỜI |
| Hoạt động tập thể, Hoạt động thể chất | HOẠT ĐỘNG THỂ CHẤT |

---

## Step 3: Match schedule entries to activities

**Function:** `find_activity(schedule_entry, activities)`

Names differ between the two files (prefixes, quotes, line breaks), so the
lookup runs twice: first against activities of the entry's own category, then
against every activity of the week. Within each pool it tries a normalized
exact match on each candidate name, then word-overlap scoring with a
threshold of 60.

Candidate names are the raw schedule name, its normalized form, and the same
two for each line of a multi-line cell — that is what lets
`TC& SK: Hoạt động: Chuyến du ngoạn quanh trường học` find
`CHUYẾN DU NGOẠN QUANH TRƯỜNG HỌC`.

An entry with no match still gets a one-row placeholder holding its type and
name, for the teacher to fill in.

---

## Step 4: Build the output

One sheet per weekday, named `Thứ <n>_<DD.MM.YYYY>`:

```
Row 1      GIÁO ÁN                         (Cambria 18pt bold, height 48)
Row 2      Dành cho thứ n, ngày DD/MM/YYYY (gray #5A5A5A bold)
Row 3      Theme                           (white on red #CC0000)
Row 4      Cơ sở: ...
Row 5      Giáo viên thực hiện: ...
Row 6      Lớp: ...
Row 7      Các hoạt động trong ngày:       (green #B6D7A8)
Row 8      STT | Loại hoạt động | Tên hoạt động | Mục tiêu hoạt động (D:I)
Rows 9-N   One summary row per activity
Rows N+1,2 Detail header, 2 rows          (navy #1F3864, F:G merged on row 1)
Rows N+3+  One block per activity
Last row   Suy ngẫm về dạy & học:
```

### Detail block

A block is as tall as the activity has phases. Column F holds one phase per
row; every other column is merged down the whole block.

| Column | Content |
|--------|---------|
| A | `<Loại hoạt động>\n<Tên hoạt động>` |
| B | CHUẨN/MỤC TIÊU CHƯƠNG |
| C | MỤC TIÊU BÀI/ MỤC TIÊU HOẠT ĐỘNG |
| D | TIÊU CHÍ THÀNH CÔNG |
| E | TÀI LIỆU HỌC TẬP, CHUẨN BỊ |
| F | One teaching phase |
| G | Mục tiêu hướng tới — `MT1, MT2, ...`, one per objective, max 4 |
| H | ĐÁNH GIÁ |
| I | PHÂN HÓA |

Row heights are left unset so Excel auto-fits the wrapped text on open.

---

## Usage

```bash
# List the classes in a schedule file
python3 generate_giao_an_v2.py -s "Tuần 1.xlsx" --list-classes

# One class
python3 generate_giao_an_v2.py \
  -s "examples_new/Tuần 1_Thế giới rực rỡ sắc màu-Tuần 2 tháng 7.xlsx" \
  -c "examples_new/KH chi tiết.xlsx" \
  --class "Vega 6"

# Every class, into a chosen directory
python3 generate_giao_an_v2.py \
  -s "Tuần 1.xlsx" -c "KH chi tiết.xlsx" --all -o output/
```

### CLI arguments

| Argument | Short | Required | Description |
|----------|-------|----------|-------------|
| `--schedule` | `-s` | Yes | Weekly schedule file (`.xlsx`) |
| `--content` | `-c` | To generate | "KH chi tiết" content file (`.xlsx`) |
| `--class` | `-C` | | Class name, e.g. `Vega 6` |
| `--all` | `-a` | | Generate for every class |
| `--list-classes` | `-l` | | List classes and exit |
| `--week` | `-w` | | Week to read from the content file; default: the `(Tuần n)` of the class theme |
| `--location` | | | School name printed on the plan |
| `--output` | `-o` | | Output directory (default `output/`) |
