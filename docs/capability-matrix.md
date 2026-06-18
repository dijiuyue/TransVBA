# TransVBA Capability Matrix

Status: 2026-05-16

## Legend
- ✅ **Complete** — Feature is implemented, tested, and verified
- ⚠️ **Partial** — Feature exists but has known limitations
- ❌ **Missing** — Feature is not implemented
- 🔧 **Risky** — Feature appears to work but has structural risks

## Core Formatting

| Capability | Status | Notes |
|---|---|---|
| Body text formatting | ✅ | Font, size, alignment, spacing, indentation |
| Normal style setup | ⚠️ | eastAsia font set on first run, not on style itself |
| ASCII font unification | ✅ | Times New Roman across all runs |
| .doc → .docx conversion | ✅ | COM-based, with proper error handling |

## Title Detection & Formatting

| Capability | Status | Notes |
|---|---|---|
| Arabic numeric titles (1, 1.1, 1.1.1, 1.1.1.1) | ✅ | Pattern-based detection for levels 1-4 |
| Chinese number titles (一、二、三) | ➖ | Intentionally skipped; only Arabic numeric headings are auto-detected |
| Word multilevel list headings | ⚠️ | COM resolver enabled by default; docx fallback is limited (no list text resolution) |
| Compound paragraph splitting | ✅ | Splits concatenated titles |
| Title formatting (4 auto-detected levels) | ✅ | Outline level, font, size, bold, alignment, spacing |
| Respect existing outline levels | ✅ | User-set outline levels preserved |
| Heading style recognition | ✅ | Word built-in heading styles recognized |
| List item detection (1), （1）, a.) | ➖ | Intentionally skipped to avoid false positives |

## COM / List Resolver

| Capability | Status | Notes |
|---|---|---|
| DocxListResolver (python-docx) | ⚠️ | Returns level from numPr, but list text always None |
| ComListResolver (pywin32) | ✅ | Full VBA-equivalent ListLevelNumber + ListString |
| auto_select() | ✅ | Prefers COM, graceful fallback to docx with warning |
| COM resolver enabled by default | ✅ | prefer_com_resolver=true in both templates |
| COM→python-docx paragraph mapping | ✅ | Element-based mapping with index fallback |
| COM/temp file cleanup on error | ✅ | try/finally ensures Word process and temp files cleaned up |

## Table Formatting

| Capability | Status | Notes |
|---|---|---|
| Table caption detection | ✅ | Pattern-based ("表 x.x-x") |
| Table caption formatting | ✅ | Font, size, bold, alignment, spacing |
| Table body formatting | ✅ | Borders, row height, cell fonts |
| Table centering | ✅ | Horizontal alignment |
| Auto-fit mode | ✅ | Window/content/fixed |
| Shape/TextFrame captions | ⚠️ | Detected but not formatted (python-docx limitation) |
| Standalone caption formatting | ✅ | Captions not near tables also formatted |

## Figure Formatting

| Capability | Status | Notes |
|---|---|---|
| Figure caption detection | ✅ | Pattern-based ("图 x.x-x") |
| Figure caption formatting | ✅ | Font, size, bold, alignment, spacing |
| Standalone caption formatting | ✅ | All matching captions formatted |

## TOC (Table of Contents)

| Capability | Status | Notes |
|---|---|---|
| TOC entry detection | ✅ | Text pattern + style-based |
| TOC entry formatting | ✅ | Font, size, bold, spacing by level |
| Custom TOC style support | ✅ | Any style with "toc" in name |
| TOC title formatting | ✅ | "目录" detection |

## Cover & Appendix

| Capability | Status | Notes |
|---|---|---|
| Cover title formatting | ✅ | Centered, 二号 font, within first N paragraphs |
| Appendix title detection | ✅ | "附件N" pattern |
| Appendix title formatting | ✅ | 小四 bold |
| Appendix body formatting | ✅ | 小五 single spacing |
| Appendix boundary detection | ✅ | Stops at next heading |

## Header Formatting

| Capability | Status | Notes |
|---|---|---|
| Header font/size/bold | ✅ | All sections |
| Header line spacing | ✅ | Per settings |
| "Rev." spacing normalization | ✅ | Single space after "Rev." |

## Validation / Format Check

| Capability | Status | Notes |
|---|---|---|
| Chinese font check | ✅ | 宋体 verification |
| ASCII font check | ✅ | Times New Roman verification |
| Bracket check | ✅ | Fullwidth bracket verification |
| Period check | ✅ | List item period verification |
| Forbidden words check | ✅ | Configurable word list |
| Table font size check | ✅ | |
| Table row height check | ✅ | |
| Cover title size check | ✅ | 二号 verification |
| Appendix body size check | ✅ | 小五 verification |
| Grid alignment check | ✅ | Level 1 snapToGrid verification |
| Appendix colon check | ✅ | Fullwidth colon verification |
| Figure/table space check | ✅ | Single space verification |
| Figure position check | ✅ | Caption above/below verification |
| chairman_number check | ✅ | Implemented in _check_chairman_number |

## GUI

| Capability | Status | Notes |
|---|---|---|
| File open | ✅ | |
| Template switching | ✅ | |
| Settings panels (body, title, etc.) | ✅ | |
| Apply formatting | ✅ | With progress and elapsed time |
| Format validation | ✅ | |
| Save/load user settings | ✅ | %APPDATA%\TransVBA |
| COM resolver toggle | ✅ | Auto-enabled, fallback warning in status bar |
| Warning display | ✅ | Warnings surfaced in completion dialog and status bar |
| Preset load/save | ✅ | Save/load named presets in presets/ directory |
| COM availability indicator | ✅ | Green/red label in advanced settings |
| Output path collision avoidance | ✅ | Timestamp-based naming when file exists |

## Test Infrastructure

| Capability | Status | Notes |
|---|---|---|
| pytest runnable | ⚠️ | Tests defined; requires `pip install -e ".[dev]"` in current environment |
| COM test isolation | ✅ | word_com marker, auto-skip when Word unavailable |
| Golden fixtures | ❌ | Directory created, no golden files yet |
| word_com marker | ✅ | Configured in pyproject.toml |
