---
name: pptx
description: "Use this skill any time a .pptx file is involved in any way — as input, output, or both. This includes: creating slide decks, pitch decks, or presentations; reading, parsing, or extracting text from any .pptx file (even if the extracted content will be used elsewhere, like in an email or summary); editing, modifying, or updating existing presentations; combining or splitting slide files; working with templates, layouts, speaker notes, or comments. Trigger whenever the user mentions \"deck,\" \"slides,\" \"presentation,\" or references a .pptx filename, regardless of what they plan to do with the content afterward. If a .pptx file needs to be opened, created, or touched, use this skill."
license: Proprietary. LICENSE.txt has complete terms
---

# PPTX Skill

## Quick Reference

| Task | Guide |
|------|-------|
| Read/analyze content | `python -m markitdown presentation.pptx` |
| Edit or create from template | Read [editing.md](editing.md) |
| Create from scratch | Read [pptxgenjs.md](pptxgenjs.md) |

---

## Reading Content

```bash
# Text extraction
python -m markitdown presentation.pptx

# Visual overview
python scripts/thumbnail.py presentation.pptx

# Raw XML
python scripts/office/unpack.py presentation.pptx unpacked/
```

---

## Editing Workflow

**Read [editing.md](editing.md) for full details.**

1. Analyze template with `thumbnail.py`
2. Unpack → manipulate slides → edit content → clean → pack

---

## Creating from Scratch

**Read [pptxgenjs.md](pptxgenjs.md) for full details.**

Use when no template or reference presentation is available.

---

## Design System (MANDATORY)

**You MUST follow this specific design system for all presentation generation and editing tasks.**
This overrides any generic design advice.

### Presentation Metadata
- **Theme Concept**: "Digital Empowerment & Minimalist Trust" (俐落知性 / Cathay CUBE APP留白美學 / 數位賦能)
- **Style**: Professional, cool (專業冷靜), with warm accents (暖色系點綴) and minimalist aesthetics (留白美學).
- **Aspect Ratio**: 16:9
- **Language**: Traditional Chinese (Taiwan)

### Color Palette (60-30-10 Rule)

| Role | Weight | Color Name | Hex | Usage |
|------|--------|------------|-----|-------|
| **Primary** | 60% | Cathay Dark Blue | `#0843AD` | Backgrounds, Headers, Main Charts |
| | | Tech Blue | `#186AFF` | |
| | | Data Teal | `#36CBDA` | |
| **Neutral** | 30% | Text Black | `#262626` | Body Text, Borders |
| | | Subtitle Gray | `#A5A5A5` | Annotations, Secondary Text |
| | | Canvas White | `#FFFFFF` | Backgrounds |
| | | Section Gray | `#F2F2F2` | Section Separators, Fill |
| **Accent** | 10% | Energy Yellow | `#FFD600` | Highlights, KPIs |
| | | Action Orange | `#F75801` | Call-to-Actions, Important dots/lines |

**Gradients:**
1. `#B2D9FC` to `#186AFF` to `#0843AD` (Top-Left to Bottom-Right)
2. `#CDF2F6` to `#82DFE8` to `#36CBDA`

### Typography

- **English Font**: Arial
- **Chinese Font**: Microsoft JhengHei (微軟正黑體)
- **Hierarchy**:
  - **Slide Title**: 30pt, Bold, `#0843AD` (Cathay Dark Blue)
  - **Section Header**: 28pt, Bold, `#262626`
  - **Subtitle**: 22pt, Regular, `#186AFF` (Tech Blue)
  - **Body Text**: 14pt, Regular, `#262626` (Line Height: 1.5)
  - **Annotation**: 12pt, `#A5A5A5`

### Layout Rules

- **Whitespace**: High (Minimum 30% of slide area MUST be empty).
- **Grid System**: 12-column modular grid for charts and content blocks.
- **Visual Style**:
  - Use rounded corners for images and content blocks (simulating App UI).
  - Avoid heavy borders; use whitespace or light gray fills (`#F2F2F2`) to separate sections.
  - Icons should be outlined (stroke style) in Dark Blue or Gray.

### Instructions for Generation

1. **Tone**: Maintain a professional, corporate yet innovative tone.
2. **Charts**: Ensure all charts are clear, flat design, informative and use the specified palette.
  - Bar Charts: Use Tech Blue (`#186AFF`) for current data, Light Blue (`#B2D9FC`) for historical.
  - KPI Highlights: Large font size (40pt+) in Dark Blue, accented with Orange line or dot.
3. **Imagery**: Images should be high-quality, representing "Connectivity", "Scale", and "future".
  - Abstract tech imagery (fiber optics, building abstract) or Gradient Blue backgrounds are preferred.
4. **Forbidden**: Do not use Serif fonts. Do not use generic, low-effort layouts.

---

## QA (Required)

**Assume there are problems. Your job is to find them.**

Your first render is almost never correct. Approach QA as a bug hunt, not a confirmation step. If you found zero issues on first inspection, you weren't looking hard enough.

### Content QA

```bash
python -m markitdown output.pptx
```

Check for missing content, typos, wrong order.

**When using templates, check for leftover placeholder text:**

```bash
python -m markitdown output.pptx | grep -iE "xxxx|lorem|ipsum|this.*(page|slide).*layout"
```

If grep returns results, fix them before declaring success.

### Visual QA

**⚠️ USE SUBAGENTS** — even for 2-3 slides. You've been staring at the code and will see what you expect, not what's there. Subagents have fresh eyes.

Convert slides to images (see [Converting to Images](#converting-to-images)), then use this prompt:

```
Visually inspect these slides. Assume there are issues — find them.

Look for:
- Overlapping elements (text through shapes, lines through words, stacked elements)
- Text overflow or cut off at edges/box boundaries
- Decorative lines positioned for single-line text but title wrapped to two lines
- Source citations or footers colliding with content above
- Elements too close (< 0.3" gaps) or cards/sections nearly touching
- Uneven gaps (large empty area in one place, cramped in another)
- Insufficient margin from slide edges (< 0.5")
- Columns or similar elements not aligned consistently
- Low-contrast text (e.g., light gray text on cream-colored background)
- Low-contrast icons (e.g., dark icons on dark backgrounds without a contrasting circle)
- Text boxes too narrow causing excessive wrapping
- Leftover placeholder content

For each slide, list issues or areas of concern, even if minor.

Read and analyze these images:
1. /path/to/slide-01.jpg (Expected: [brief description])
2. /path/to/slide-02.jpg (Expected: [brief description])

Report ALL issues found, including minor ones.
```

### Verification Loop

1. Generate slides → Convert to images → Inspect
2. **List issues found** (if none found, look again more critically)
3. Fix issues
4. **Re-verify affected slides** — one fix often creates another problem
5. Repeat until a full pass reveals no new issues

**Do not declare success until you've completed at least one fix-and-verify cycle.**

---

## Converting to Images

Convert presentations to individual slide images for visual inspection:

```bash
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

This creates `slide-01.jpg`, `slide-02.jpg`, etc.

To re-render specific slides after fixes:

```bash
pdftoppm -jpeg -r 150 -f N -l N output.pdf slide-fixed
```

---

## Dependencies

- `pip install "markitdown[pptx]"` - text extraction
- `pip install Pillow` - thumbnail grids
- `npm install -g pptxgenjs` - creating from scratch
- LibreOffice (`soffice`) - PDF conversion (auto-configured for sandboxed environments via `scripts/office/soffice.py`)
- Poppler (`pdftoppm`) - PDF to images
