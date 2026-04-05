# Giáo Án Generator - Workflow Documentation

## Overview

The `generate_giao_an.py` script automatically generates lesson plan files (`giao_an_output.xlsx`) by combining two data sources:

- **Schedule file** (`.xlsx`) — Weekly schedule per class (which activities happen on which day)
- **Content files** (`.zip` or directory of `.xlsx`) — 8 content template files with detailed lesson content

```
schedule.xlsx                    content.zip (or directory/)
  (Weekly Schedule)                  (Activity Content)
           |                                  |
           v                                  v
    Step 1: Parse schedule       Step 2: Parse content files
    for the given class          Find activity blocks by
    - date, theme, teacher       column A merged cells
    - 6 time slots per day       Extract values + merges
    - skip ESL/routine           for each activity block
           |                                  |
           v                                  v
    +---------------------------------------------------------+
    | Step 3: Match schedule activities -> content blocks      |
    |   4-level fuzzy matching (exact, normalized, substring,  |
    |   word-overlap)                                          |
    +---------------------------------------------------------+
                              |
                              v
    +---------------------------------------------------------+
    | Step 4: Build giao_an_output.xlsx                        |
    |   5 sheets (Mon-Fri), each with:                         |
    |   - Header (title, date, theme, teacher, class)          |
    |   - Summary table of activities                          |
    |   - Detail table with full lesson content                |
    |   - Footer for teacher reflection                        |
    +---------------------------------------------------------+
```

---

## Step 1: Parse Content Files

**Function:** `parse_all_content_files()`

Reads all 8 `.xlsx` files from `noi_dung_ke_hoach_giang_day/`:

| File | Activities |
|------|-----------|
| Trò chuyện đầu ngày.xlsx | 5: Ngôi sao của tuần, Đọc truyện, Bài hát của tuần, Chiếc túi vần điệu, Khoảnh khắc ấn tượng |
| Hoạt động chính.xlsx | 2: So sánh kích thước, Nào mình cùng chinh phục thử thách |
| Hoạt động Tiếng Việt.xlsx | 2: Thơ Bé làm bao nhiêu nghề, Truyện Bác gấu đen làm bánh |
| Hoạt động tập thể.xlsx | 2: Đoàn tàu tí xíu, Lính cứu hoả |
| Hoạt động ngoài trời.xlsx | 2: Xếp tháp, Bạn nhảy xa bao nhiêu |
| Hoạt động vui chơi.xlsx | 1: Thủ công và biểu đạt trực quan |
| Hoạt động dự án.xlsx | 1: Trạm yêu thích |
| Dã ngoại sự kiện và gia đình.xlsx | 1: Chuyến du ngoạn |

### How activity blocks are detected

Each content file has a standard structure:
- Rows 1-7: Template header (title, date placeholder, theme, etc.)
- Rows 8-9 (or 14-15): Column headers for the 9-column detail table
- Rows 10+: Activity data blocks

An **activity block** is identified by a **merged cell range on column A**. For example, in `Hoạt động chính.xlsx`:

```
A10:A12 merged  ->  "So sánh kích thước"          (3-row block)
A15:A17 merged  ->  "Nào mình cùng chinh phục..."  (3-row block)
```

Header values like `'HOẠT ĐỘNG'`, `'STT'`, `'GIÁO ÁN'` are skipped.

### What is extracted per block

For each activity block (e.g., rows 10-12):

| Data | Description |
|------|-------------|
| `name` | Activity name from column A (first row of merged range) |
| `num_rows` | Number of rows in the block (e.g., 3) |
| `rows` | Cell values for columns A-I, all rows |
| `merges` | All merge patterns within the block, stored as **relative offsets** (so the block can be placed at any target row in the output) |
| `objectives_summary` | All column C values joined with `\n` (used for the summary table) |

### Formula handling

Some cells contain formulas (e.g., `=H10`). The script loads each file twice:
1. With `data_only=True` to get computed values
2. Without `data_only` to access formulas as fallback

For simple cell references like `=H10`, the script resolves them by reading the referenced cell.

---

## Step 2: Parse Schedule

**Function:** `parse_schedule(class_name)`

Reads `ke_hoach_giang_day.xlsx` and finds the sheet for the given class.

### Schedule file structure

Each class sheet (e.g., "Vega 12") is a Mon-Fri timetable:

```
Row 1:     "KẾ HOẠCH GIẢNG DẠY" (title)
Row 2:     Week/date range + class name
Row 3:     Theme + teacher names
Row 4:     Day headers: Thời gian | | Thứ hai | Thứ ba | Thứ tư | Thứ năm | Thứ sáu
Rows 5-27: Time slots with activities per day
```

Days map to columns: Mon=C, Tue=D, Wed=E, Thu=F, Fri=G.

### Extracted metadata

| Field | Source | Example |
|-------|--------|---------|
| Date range | Row 2, regex | `06/4 đến 10/4/2026` |
| Theme | Row 3, col B | `Cuộc sống ở thành phố` |
| Teacher | Row 3, col G | `Cao Hương` |

### Time slots scanned

Each slot has 2 rows: **type row** (activity category) and **detail row** (specific name):

| Time Slot | Type Row | Detail Row |
|-----------|----------|------------|
| 8h30-9h00 | Row 8 | Row 9 |
| 9h00-9h30 | Row 10 | Row 11 |
| 9h30-10h00 | Row 12 | Row 13 |
| 10h00-10h30 | Row 14 | Row 15 |
| 14h30-15h00 | Row 21 | Row 22 |
| 15h00-15h30 | Row 23 | Row 24 |

### Filtering (skip rules)

Activities are excluded if any of these keywords appear in the type or detail:

- **Foreign teacher:** `Story time`, `ESL Class`, `ESL 1/2/3`
- **Generic play:** `Chơi tự do`, `Vui chơi Softplay`, `phòng Lego`, `phòng Cosplay`, `phòng thư viện`
- **Routine:** `Đón trẻ`, `Thể dục sáng`, `Ăn sáng/trưa/chiều`, `Ngủ`, `Vệ sinh`, `Uống sữa`, `Tái hiện`, `Chơi tự chọn`, `Dọn dẹp`

Additionally, the activity type must exist in `ACTIVITY_FILE_MAP` (meaning it's a known Vietnamese-teacher structured activity).

---

## Step 3: Match Activities to Content

**Function:** `find_content(activity, all_content)`

Schedule activity names often differ from content file names (different prefixes, quotes, line breaks). The matching uses **4 levels of fuzzy matching**:

### Level 1: Exact NFC match

Compares Unicode NFC-normalized strings directly.

```
Schedule: "Ngôi sao của tuần"
Content:  "Ngôi sao của tuần"
-> MATCH
```

### Level 2: Normalized match

Strips common prefixes (`Truyện: `, `Kỹ năng toán học\n`, `Câu chuyện của tháng\n`, etc.) and quotes before comparing.

```
Schedule: "Kỹ năng toán học\nSo sánh kích thước"
  -> normalize -> "So sánh kích thước"
Content:  "So sánh kích thước"
-> MATCH
```

### Level 3: Substring match

Checks if one normalized name contains the other (minimum 4 characters).

```
Schedule: "Bài hát của tuần"
Content:  "Bài hát của tuần: Tớ có thể nhảy lò cò/ I can hop"
-> MATCH (schedule name is substring of content name)
```

### Level 4: Word-overlap scoring

Counts shared words between two names. Requires score >= 60.

```
Schedule: "Chiếc túi bài hát/vần điệu"
Content:  "Chiếc túi vần điệu"
-> Shared words: "Chiếc", "túi", "vần", "điệu" (4 words x 20 = 80)
-> MATCH
```

### Unmatched activities

Some activities in the schedule don't have pre-built content (e.g., "Khám phá thành phố", "Vẽ thành phố bằng mực chanh"). These get a 1-row placeholder in the output. The teacher can manually fill in the details.

---

## Step 4: Build Output Sheets

**Function:** `build_sheet(ws, schedule, day, matched_activities)`

For each weekday (Mon-Fri), creates one sheet with this structure:

### Sheet layout

```
Row 1:      "GIÁO ÁN"                    (Cambria 28pt, center, medium border)
Row 2:      Date line                      (Cambria 10pt bold gray #5A5A5A, center)
Row 3:      Theme/Chapter                  (white text on red #CC0000, center)
Row 4:      Location "Cơ sở: Times City"   (Cambria 10pt bold, left)
Row 5:      Teacher name                   (Cambria 10pt bold, left)
Row 6:      Class name                     (Cambria 10pt bold, left)
Row 7:      "Các hoạt động trong ngày:"    (bold, green #B6D7A8, center)
Row 8:      Summary header                 (STT | Loại | Tên | Mục tiêu D:I merged)
Row 9-N:    Summary rows                   (one per activity)
Row N+1:    Detail header row 1            (navy #1F3864, white bold, F:G merged)
Row N+2:    Detail header row 2            ("Hoạt động" | "Mục tiêu hướng tới")
Row N+3+:   Activity content blocks        (from content files, thin borders)
Last 2:     Footer "Suy ngẫm về dạy & học:" (Cambria 11pt bold, medium border)
```

### Summary table columns

| Column | Content | Source |
|--------|---------|--------|
| A | STT (1, 2, 3, ...) | Auto-numbered |
| B | Activity type | From schedule type row |
| C | Activity name | From schedule detail row |
| D-I (merged) | Objectives | Column C values from content file, joined with `\n` |

### Detail table columns (9 columns)

| Column | Header | Description |
|--------|--------|-------------|
| A | HOẠT ĐỘNG | Activity name |
| B | CHUẨN/MỤC TIÊU CHƯƠNG | Chapter standards/objectives |
| C | MỤC TIÊU BÀI | Lesson objectives (may span multiple rows) |
| D | TIÊU CHÍ THÀNH CÔNG | Success criteria |
| E | TÀI LIỆU HỌC TẬP | Learning materials |
| F | CÁC BƯỚC TỔ CHỨC (Hoạt động) | Teaching steps (3 phases: Mở đầu, Giảng dạy chính, Kết thúc) |
| G | Mục tiêu hướng tới | Target objectives per step (e.g., MT1, MT2) |
| H | ĐÁNH GIÁ | Assessment methods |
| I | PHÂN HÓA | Differentiation strategies |

### Content block vs Placeholder

**With content file match:** Full block copied from content file — all cell values and merge patterns are reproduced with consistent styling (Cambria 10pt, thin borders, left/top aligned, wrap text).

**Without content file match:** 1-row placeholder with activity type + name in column A. Remaining columns are empty for the teacher to fill manually.

### Sheet properties

| Property | Value |
|----------|-------|
| Orientation | Landscape |
| Column A width | 25.71 |
| Column J width | 8.71 |
| Default column width | 14.43 |
| Default row height | 15.0 |
| Row 1 height | 48.0 |
| Data row heights | 12.75 |
| Page margins | left=0.7, right=0.7, top=0.75, bottom=0.75 |

---

## Activity Type Mapping

Maps schedule activity types to content files:

| Schedule Type | Content File | Notes |
|---------------|-------------|-------|
| Trò chuyện đầu ngày | Trò chuyện đầu ngày.xlsx | 5 morning chat activities |
| Hoạt động chính | Hoạt động chính.xlsx | Main activities (math, PE, etc.) |
| Hoạt động Tiếng Việt | Hoạt động Tiếng Việt.xlsx | Vietnamese language activities |
| Hoạt động tập thể | Hoạt động tập thể.xlsx | Group activities |
| Hoạt động thể chất | Hoạt động tập thể.xlsx | Physical activities (same file) |
| Hoạt động ngoài trời | Hoạt động ngoài trời.xlsx | Outdoor activities |
| Hoạt động vui chơi | Hoạt động vui chơi.xlsx | Play activities |
| Chơi theo chủ đề | Hoạt động vui chơi.xlsx | Themed play (same file) |
| Hoạt động Dự án | Hoạt động dự án.xlsx | Project activities |
| Dã ngoại/Sự kiện gia đình | Dã ngoại sự kiện và gia đình.xlsx | Field trips/family events |

---

## Usage

```bash
# List available classes in a schedule file
python3 generate_giao_an.py -s schedule.xlsx --list-classes

# Generate lesson plan from zip content
python3 generate_giao_an.py \
  --schedule schedule.xlsx \
  --content noi_dung.zip \
  --class "Vega 12"

# Generate from directory content with custom output path
python3 generate_giao_an.py \
  -s schedule.xlsx \
  -c noi_dung_dir/ \
  -C "Luna 1" \
  -o giao_an_luna1.xlsx

# Using examples
python3 generate_giao_an.py \
  -s examples/ke_hoach_giang_day.xlsx \
  -c examples/noi_dung_ke_hoach_giang_day.zip \
  -C "Vega 12"
```

### CLI Arguments

| Argument | Short | Required | Description |
|----------|-------|----------|-------------|
| `--schedule` | `-s` | Yes | Path to schedule file (.xlsx) |
| `--content` | `-c` | For generate | Path to content files (.zip or directory) |
| `--class` | `-C` | For generate | Class name (e.g., "Vega 12") |
| `--output` | `-o` | No | Output path (default: `giao_an_output.xlsx`) |
| `--list-classes` | `-l` | No | List available classes and exit |

### Content input formats

The `--content` argument accepts either:
- **`.zip` file** — containing `.xlsx` files (Mac `__MACOSX` entries are auto-skipped)
- **Directory** — containing `.xlsx` files directly
