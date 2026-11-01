# sv-topic-screener — Multi-modal Short-Video Topic Screener

[English](README.md) | [简体中文](README.zh-CN.md)

A single-file, clean-room Python tool that decides whether each short video in an Excel
list is about a given topic. It collects three kinds of evidence in a fixed order — the
**title/topic text**, a **speech transcript** from faster-whisper, and **on-screen text**
read by OpenCV frame sampling plus RapidOCR — and writes the verdicts back to a new
workbook. The built-in demonstration topic is automotive driver-assist ("智驾" / ADAS), and
every term list is configurable. It never scrapes a platform and contains no third-party
proprietary code: you supply the videos yourself.

---

## Overview

The workflow it targets is a manual screening queue. A marketing or research team exports
a list of short videos from some platform's own dashboard into Excel, then a human opens
each one to decide "is this about driver-assist?" That is slow and inconsistent. This tool
replaces the first pass with deterministic keyword matching, and leaves only the genuinely
ambiguous rows for a human.

The input is an `.xlsx` file whose first worksheet has a header row containing a title
column, a link column and an optional platform column (many Chinese and English header
aliases are accepted, see [Input columns](#input-columns-and-aliases)). The videos
themselves must already be on local disk, named either by the media id extracted from the
URL or by the Excel row number. If a video is absent, the tool cannot look at its audio or
pixels and says so instead of guessing.

Every row ends in exactly one of four outcomes: **related via text**, **related via media**,
**needs manual review**, or **not related**. A row that matches any `exclude_terms` entry in
its title, transcript or OCR text is always **not related**, with status `excluded`, even if
it also matches a topic term. Rows that hit a topic term in the title are decided without
ever touching a video file, so a text-only run over a large sheet is fast.

### Clean-room and legal scope

- No platform scraping, no logins, no cookies, no private APIs. The repository contains no
  URL fetcher at all; `find_media` only looks for files that are already in a directory you
  pass in (`screener.py:75-86`).
- No third-party proprietary code or model weights are vendored. The dependencies in
  `requirements.txt` are public packages, and each one resolves its own model files at
  runtime.
- Responsibility for having the right to process a given video rests with whoever runs the
  tool. Nothing here bypasses a platform's terms of service because nothing here contacts a
  platform.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Architecture](#architecture)
- [Key flows](#key-flows)
- [Data model](#data-model)
- [Directory layout](#directory-layout)
- [Getting started](#getting-started)
- [Commands](#commands)
- [Configuration](#configuration)
- [CLI](#cli)
- [Troubleshooting](#troubleshooting)
- [Design decisions & limitations](#design-decisions--limitations)
- [FAQ](#faq)
- [License](#license)

---

## Features

| Area | Behaviour |
|---|---|
| Input | Excel `.xlsx`, first worksheet only, row 1 treated as the header (`screener.py:56-72`) |
| Column detection | Case-insensitive exact match against three alias lists (`screener.py:27-29`, `screener.py:48-53`) |
| Text check | Case-insensitive substring search of the title against the topic terms (`screener.py:155-157`) |
| Exclusion | Any `exclude_terms` match in the title, transcript or OCR text wins over topic hits and yields `不相关` / `excluded` (`screener.py:181-183`, `screener.py:200-202`) |
| Media check | faster-whisper transcript + 6-frame OpenCV/RapidOCR sample, merged (`screener.py:194-204`) |
| Media lookup | `<media-id>.mp4` from the URL, falling back to `<excel-row>.mp4` (`screener.py:75-86`) |
| Short-circuit | A title hit returns before any video is opened (`screener.py:181-187`) |
| Failure handling | ASR/OCR exceptions and a 0-frame OCR probe become `待人工复核` with `media_error` / `media_gap` plus a note, never `不相关` (`screener.py:194-210`) |
| Verdicts | `相关` / `待人工复核` / `不相关` with a machine-readable status column (`screener.py:173-211`) |
| Traceability | `序号` keeps the source sheet row; `视频文件` records the matched filename (`screener.py:272-275`) |
| Output | New workbook: one result sheet plus a `统计` sheet that always lists all three verdicts (`screener.py:280-295`) |
| Example generator | `--init-example` writes a two-row demo workbook (`screener.py:214-222`) |
| Rules override | A JSON file shallow-merged over the built-in defaults; a missing or invalid file is a hard error (`screener.py:257-263`) |
| Text-only mode | `--no-media` skips ASR and OCR entirely (`screener.py:245`, `screener.py:188-189`) |

There is no concurrency, no caching to disk, no database, and no network code.

---

## Architecture

Everything except the two setup scripts is one module. The boxes below are functions in
`screener.py`; model objects are module-level singletons so they survive across rows.

```mermaid
flowchart TD
    IN["Excel input (.xlsx)<br/>first worksheet, row 1 = header"]
    READ["read_rows() + _col()<br/>alias resolution"]
    RULES["rule merge<br/>DEFAULT_RULES then --rules JSON"]
    DEC["decide()"]
    EX{"exclude term<br/>in title?"}
    TXT["hits(title, smart_terms<br/>+ brand_terms if count_brand_as_topic)"]
    T1{"title hit?"}
    NM{"--no-media?"}
    FM["find_media()<br/>media-id.mp4, then row.mp4"]
    V1{"file exists<br/>and size above 0?"}
    ASR["transcribe()<br/>faster-whisper small, cpu, int8"]
    OCR["ocr_frames()<br/>6 frames, OpenCV + RapidOCR"]
    MEX{"exclude term in<br/>transcript or OCR?"}
    MH{"topic term in transcript<br/>or OCR text?"}
    ERR{"ASR/OCR error<br/>or OCR gap?"}
    RX["不相关 / excluded"]
    R1["相关 / text<br/>done_text"]
    R2["待人工复核 / no text hit<br/>boundary_no_media"]
    R3["待人工复核 / no local video<br/>boundary_no_media"]
    R4["相关 / media<br/>done_media"]
    R5["不相关 / media missed<br/>done_media"]
    R6["待人工复核 / media failure<br/>media_error or media_gap"]
    OUT["Result workbook<br/>筛选结果 + 统计"]

    IN --> READ --> DEC
    RULES --> DEC
    DEC --> EX
    EX -- yes --> RX
    EX -- no --> TXT --> T1
    T1 -- yes --> R1
    T1 -- no --> NM
    NM -- yes --> R2
    NM -- no --> FM --> V1
    V1 -- no --> R3
    V1 -- yes --> ASR
    V1 -- yes --> OCR
    ASR --> MEX
    OCR --> MEX
    MEX -- yes --> RX
    MEX -- no --> MH
    MH -- yes --> R4
    MH -- no --> ERR
    ERR -- yes --> R6
    ERR -- no --> R5
    R1 --> OUT
    R2 --> OUT
    R3 --> OUT
    R4 --> OUT
    R5 --> OUT
    R6 --> OUT
    RX --> OUT
```

*(source: `screener.py:173-211`, `screener.py:239-301`)*

| Layer | Path | Responsibility |
|---|---|---|
| CLI | `screener.py:239-301` | argparse, default output name, rules loading, per-row loop, workbook write, console summary |
| Input reader | `screener.py:56-72` | Load the first worksheet, resolve the title/url/platform columns, build `{row, title, url, platform}` records |
| Header matcher | `screener.py:48-53` | Strip + lowercase a header cell, return the first index matching an alias |
| Rules | `screener.py:35-45`, `rules.example.json` | Built-in term dictionaries and the overridable example |
| Rules loader | `screener.py:225-236` | Parse `--rules` JSON, reject a missing path, invalid JSON or a non-object |
| Matcher | `screener.py:155-157` | `hits(text, terms)` — lowercase substring containment |
| Decision | `screener.py:173-211` | The precedence documented below, including exclusion and media-failure handling |
| Media lookup | `screener.py:75-86` | URL regex, two filename candidates, non-empty-file check |
| ASR | `screener.py:92-106` | Lazy `WhisperModel` singleton, VAD on, beam size 1; returns text plus an error note |
| OCR | `screener.py:112-152` | Frame-count probe, uniform sampling, RapidOCR singleton, error and zero-frame gap reporting |
| Output writer | `screener.py:270-300` | Result rows, `统计` sheet, progress every 50 rows, console distribution |
| Setup | `setup.ps1`, `setup.bat` | Create `.venv` and `pip install -r requirements.txt` (Windows) |

---

## Key flows

### 1. One row, with a local video

This is the full path. The two media signals are independent: either one matching is
enough, and both are stored in the output even when the row is rejected.

```mermaid
sequenceDiagram
    participant X as Input workbook
    participant M as main loop
    participant D as decide
    participant FS as media dir
    participant W as faster-whisper
    participant O as OpenCV + RapidOCR
    participant OB as Output workbook

    X->>M: read_rows() gives {row, title, url, platform}
    M->>D: decide(rec, rules, media_dir, do_media)
    D->>D: hits(title, exclude_terms)
    alt exclude term in title
        D-->>M: 不相关, 排除词(<term>), excluded
    else no exclude term, hits(title, smart_terms [+ brand if enabled])
        D-->>M: 相关, 文本(标题/话题), done_text
    else no hit and do_media is false
        D-->>M: 待人工复核, 文本未命中, boundary_no_media
    else no hit and do_media is true
        D->>FS: find_media(): media-id.mp4 then row.mp4
        alt no file, or file size is 0
            D-->>M: 待人工复核, 无本地视频, boundary_no_media
        else file found
            D->>W: transcribe(video)
            W-->>D: transcript plus an error note on failure
            D->>O: ocr_frames(video, n_frames=6)
            O-->>D: OCR text plus error and gap notes
            alt exclude term in transcript or OCR
                D-->>M: 不相关, 排除词(<term>), excluded
            else any topic term in transcript or OCR
                D-->>M: 相关, 媒体(口播/画面), done_media
            else ASR/OCR failed or OCR was skipped
                D-->>M: 待人工复核, 媒体处理异常, media_error or media_gap
            else fully processed, no term anywhere
                D-->>M: 不相关, 媒体均未命中, done_media
            end
        end
    end
    M->>OB: append row with source row, matched filename, hits[:8], snippets, note
    M->>OB: 统计 sheet: all three verdicts, 合计, 耗时秒
```

*(source: `screener.py:173-211`, `screener.py:270-276`)*

### 2. Batch startup, rule merge and model caching

The models are built on first use and reused for the rest of the run. `--init-example`
returns before any input file is opened, and a missing `--input` is an argparse error
(exit code 2) rather than a silent no-op.

```mermaid
sequenceDiagram
    participant U as User
    participant A as argparse
    participant R as rules
    participant X as Excel
    participant L as row loop
    participant C as model singletons

    U->>A: screener.py --input ... --media-dir ...
    alt --init-example
        A->>X: write example workbook (default example_input.xlsx)
        A-->>U: exit 0
    else no --input
        A-->>U: argparse error, exit 2
    end
    alt --rules path missing, invalid JSON, or not a JSON object
        A-->>U: error message, exit 2
    end
    A->>R: copy DEFAULT_RULES, then update() from --rules JSON
    R-->>A: effective rules (lists replaced wholesale)
    A->>X: read_rows() on worksheet 1
    X-->>A: records
    loop each record
        A->>L: decide(...)
        L->>C: first ASR or OCR call constructs the singleton
        C-->>L: reused for every later row
    end
    A->>X: save stem + _筛选结果.xlsx
    A-->>U: record count, elapsed time, verdict distribution, output path
```

*(source: `screener.py:249-264`, `screener.py:100-103`, `screener.py:122-129`, `screener.py:296-300`)*

### 3. Media lookup and sampling detail

Worth knowing before naming your files, because a stale file silently wins.

```mermaid
sequenceDiagram
    participant D as media lookup
    participant U as URL string
    participant F as media dir
    participant CV as OpenCV

    D->>U: regex video|note|explore|sph + slash + 6 or more id chars
    U-->>D: media id or nothing
    D->>F: try media-id.mp4 first
    F-->>D: exists and size above 0 means accepted
    D->>F: otherwise try excel-row.mp4
    F-->>D: same check
    D->>CV: VideoCapture, read CAP_PROP_FRAME_COUNT
    CV-->>D: total frames
    Note over D: 0 frames becomes a media_gap note, not a silent skip
    D->>CV: seek to int(total * (k + 0.5) / 6) for k = 0..5
    CV-->>D: 6 frames at the centre of six equal slices
```

The frame index formula `int(total * (k + 0.5) / n_frames)` samples slice centres rather
than the very first and last frames, which are often black or contain intros.
*(source: `screener.py:82-85`, `screener.py:124-132`)*

---

## Data model

The "database" is two Excel workbooks. Input columns are located by alias, so their order
does not matter and unread fields are ignored. Output columns are always written in the
same fixed order.

```mermaid
erDiagram
    INPUT_SHEET {
        string title "标题 / video title - matched against topic terms"
        string url "链接 - used only to extract a media id"
        string platform "平台 - copied through to the result"
        string other "any other columns - ignored"
    }
    RESULT_SHEET {
        int xuhao "序号 - source sheet row number (header is row 1)"
        string pingtai "平台"
        string biaoti "标题"
        string lianjie "链接"
        string panding "判定 - 相关 / 待人工复核 / 不相关"
        string laiyuan "命中来源"
        string mingzhongci "命中词 - at most 8, joined with 、"
        string koubo "口播片段 - first 200 characters"
        string huamian "画面片段 - first 200 characters"
        string shipinwenjian "视频文件 - matched local filename, empty if no video was looked up"
        string beizhu "备注 - ASR/OCR error or OCR gap note"
        string zhuangtai "状态 - done_text / done_media / excluded / media_error / media_gap / boundary_no_media"
    }
    STATS_SHEET {
        string metric "判定 column: all three verdicts, then 合计, then 耗时秒"
        int value "条数 column: count (zero included), total rows, or elapsed seconds"
    }
    INPUT_SHEET ||--o{ RESULT_SHEET : "one row in, one row out"
```

*(source: `screener.py:56-72`, `screener.py:270-295`)*

### Input columns and aliases

Header matching is an **exact** comparison after `str().strip().lower()`, not a substring
or fuzzy match (`screener.py:48-53`). A header like `视频标题（必填）` therefore will not be
recognised as a title column. All three columns are optional; a missing one yields empty
strings for every row.

| Field | Accepted header names | Notes |
|---|---|---|
| Title | `标题`, `视频标题`, `短视频标题`, `作品标题`, `笔记标题`, `title` | The only field matched against topic terms |
| URL | `链接`, `视频链接`, `作品链接`, `笔记链接`, `url`, `url链接` | Only used to extract a media id; never opened |
| Platform | `平台`, `发布平台`, `媒体平台`, `platform` | Copied to the output, not matched |

### Output columns

| # | Column | Content | Source |
|---|---|---|---|
| 1 | `序号` | Source sheet row number: the header is row 1, so the first data row is `2` | `screener.py:272-273` |
| 2 | `平台` | Input platform value | `screener.py:273` |
| 3 | `标题` | Input title value | `screener.py:273` |
| 4 | `链接` | Input URL value | `screener.py:273` |
| 5 | `判定` | `相关` / `待人工复核` / `不相关` | `screener.py:181-211` |
| 6 | `命中来源` | `文本(标题/话题)`, `排除词(<term>)`, `文本未命中`, `无本地视频`, `媒体(口播/画面)`, `媒体处理异常`, `媒体均未命中` | `screener.py:181-211` |
| 7 | `命中词` | Matched terms joined with `、`, **first 8 only** | `screener.py:274` |
| 8 | `口播片段` | First 200 characters of the transcript, empty in text-only or no-media rows | `screener.py:206-211` |
| 9 | `画面片段` | First 200 characters of the OCR text, same caveat | `screener.py:206-211` |
| 10 | `视频文件` | Matched local filename (`<media-id>.mp4` or `<data index, 1-based>.mp4`); empty when no video was looked up | `screener.py:275` |
| 11 | `备注` | Media-processing note: ASR/OCR failure text or the OCR zero-frame gap | `screener.py:275` |
| 12 | `状态` | `done_text`, `boundary_no_media`, `excluded`, `done_media`, `media_error`, `media_gap` | `screener.py:181-211` |

The `统计` sheet holds two columns, `判定` and `条数`. It always lists all three verdicts in
the fixed order `相关`, `待人工复核`, `不相关`, including rows whose count is zero, then a
`合计` row with the row count, then a `耗时秒` row whose number is the elapsed seconds
rounded to one decimal (`screener.py:288-295`).

### Verdict precedence

This is the exact order inside `decide()`. The first matching branch returns and the rest
are never evaluated.

| Step | Condition | `判定` | `命中来源` | `命中词` | `状态` | Source |
|---|---|---|---|---|---|---|
| 1 | Title contains any exclude term | `不相关` | `排除词(<term>)` | matched exclude terms | `excluded` | `screener.py:181-183` |
| 2 | No exclude term, title contains any topic term | `相关` | `文本(标题/话题)` | matched terms | `done_text` | `screener.py:185-187` |
| 3 | No text hit, `do_media` is false (`--no-media`) | `待人工复核` | `文本未命中` | empty | `boundary_no_media` | `screener.py:188-189` |
| 4 | No text hit, media enabled, no usable video file | `待人工复核` | `无本地视频` | empty | `boundary_no_media` | `screener.py:190-192` |
| 5 | Video found, an exclude term appears in transcript **or** OCR text | `不相关` | `排除词(<term>)` | matched exclude terms | `excluded` | `screener.py:200-202` |
| 6 | Video found, no exclude term, a topic term appears in transcript **or** OCR text | `相关` | `媒体(口播/画面)` | deduplicated terms | `done_media` | `screener.py:204-206` |
| 7 | Video found, no term, but ASR/OCR failed or OCR was skipped (0 frames) | `待人工复核` | `媒体处理异常` | empty | `media_error` / `media_gap` | `screener.py:207-210` |
| 8 | Video found, fully processed, no term in either signal | `不相关` | `媒体均未命中` | empty | `done_media` | `screener.py:211` |

Consequences worth stating plainly:

- Exclusion comes first in both stages. A title that contains an exclude term is `不相关`
  even when it also contains a topic term, and the same rule applies to transcript/OCR text.
- Step 3 precedes step 4. Under `--no-media` a row never reaches `find_media`, so its
  `命中来源` is always `文本未命中` even when a matching video happens to exist.
- Step 2 precedes the media path. A title hit costs no decode at all, so on a sheet where
  most titles are on-topic the media path runs for only the remainder.
- A processing failure never yields `不相关` by itself. `不相关` requires a complete run with
  no exclude term and no topic hit; otherwise the row is `待人工复核` with a note.

### Per-row status lifecycle

The `状态` column takes six values, and all of them are terminal — no row is revisited or
re-queued by the tool.

```mermaid
stateDiagram-v2
    [*] --> TextCheck
    TextCheck --> excluded: exclude term in title
    TextCheck --> done_text: topic hit
    TextCheck --> boundary_no_media: no hit, do_media false
    TextCheck --> MediaLookup: no hit, do_media true
    MediaLookup --> boundary_no_media: no usable video
    MediaLookup --> ASR_OCR: video found
    ASR_OCR --> excluded: exclude term in transcript or OCR
    ASR_OCR --> done_media: topic hit
    ASR_OCR --> media_error: ASR/OCR failed, no hit
    ASR_OCR --> media_gap: OCR frames unavailable, no hit
    ASR_OCR --> done_media: fully processed, no hit
    done_text --> [*]
    done_media --> [*]
    excluded --> [*]
    media_error --> [*]
    media_gap --> [*]
    boundary_no_media --> [*]
```

---

## Directory layout

```text
sv-topic-screener/
├─ screener.py          # the whole tool: reader, matcher, decision, ASR, OCR, writer
├─ rules.example.json   # overridable term dictionaries for the ADAS demo topic
├─ requirements.txt     # runtime dependencies (ASR, OCR, Excel)
├─ setup.ps1            # creates .venv and pip-installs requirements (Windows)
├─ setup.bat            # cmd wrapper that calls setup.ps1 with UTF-8 console
├─ README.md            # this file (English)
├─ README.zh-CN.md      # Simplified Chinese translation
└─ .gitignore           # venv, media, model cache, generated workbooks
```

`.gitignore` excludes `.venv/`, `__pycache__/`, `*.pyc`, `*.mp4`, `media/`, `hf_cache/`,
`*_筛选结果.xlsx` and `example_input.xlsx` — i.e. videos, model caches and generated
workbooks are never committed.

---

## Getting started

### Prerequisites

- **Python 3.11+**. `setup.ps1` prints a 3.11+ requirement when `python` is not on `PATH`
  (`setup.ps1:5`); the script does not enforce a version number.
- Network access on first setup: `pip` must download wheels, and the first ASR run
  downloads the Whisper model.
- **No system ffmpeg is required.** faster-whisper decodes through PyAV, which arrives as
  the `av` wheel with faster-whisper (`screener.py:100`).

### Install (Windows, scripted)

```powershell
git clone https://github.com/Ethan-upgrade/sv-topic-screener.git
cd sv-topic-screener
.\setup.bat
```

`setup.bat` sets `chcp 65001` and calls `setup.ps1`, which creates `.venv` and runs
`pip install --upgrade pip` followed by `pip install -r requirements.txt` (`setup.ps1:6-10`).

### Install (manual, any OS)

```bash
python -m venv .venv
# Windows
.venv\Scripts\python.exe -m pip install -r requirements.txt
# macOS / Linux
.venv/bin/python -m pip install -r requirements.txt
```

### Dependencies

| Package | Role |
|---|---|
| `faster-whisper` | ASR engine, `WhisperModel` (`screener.py:100-103`); installs its own `ctranslate2`, `av` and `onnxruntime` backends |
| `opencv-python-headless` | `cv2.VideoCapture`, frame count and seeking (`screener.py:122-131`) |
| `rapidocr` | OCR engine, `RapidOCR()` (`screener.py:127-129`); installs its own `onnxruntime` |
| `openpyxl` | Reading and writing `.xlsx` (`screener.py:25`) |

`requirements.txt` declares only packages that `screener.py` imports directly. Inference
backends and numeric/image helpers are installed transitively by the two engines.

### First run

```powershell
.venv\Scripts\python.exe screener.py --init-example
.venv\Scripts\python.exe screener.py --input example_input.xlsx
.venv\Scripts\python.exe screener.py --input example_input.xlsx --media-dir .\media
```

The second command is text-only in effect (no `--media-dir`), so every row that misses on
text lands in `待人工复核`. The third reads `media\<media-id>.mp4` or `media\<data index, 1-based>.mp4`.

---

## Commands

| Command | Does |
|---|---|
| `setup.bat` / `powershell -File setup.ps1` | Create `.venv`, upgrade pip, install requirements |
| `python screener.py --init-example` | Write a 2-row demo workbook to `example_input.xlsx` and exit 0 |
| `python screener.py --init-example --input demo.xlsx` | Write the demo workbook to `demo.xlsx` instead |
| `python screener.py --input data.xlsx` | Text-only decision; no video files are opened |
| `python screener.py --input data.xlsx --no-media` | Same as above, made explicit even if a media dir is given |
| `python screener.py --input data.xlsx --media-dir videos` | Text first, then ASR + OCR for the remaining rows |
| `python screener.py --input data.xlsx --rules my-rules.json` | Use custom term dictionaries |
| `python screener.py --input data.xlsx --out result.xlsx` | Choose the output path |

There is no test suite, no linter configuration and no CI workflow in the repository.

---

## Configuration

### CLI flags

| Flag | Default | Behaviour |
|---|---|---|
| `--input` | empty | Input workbook. Required unless `--init-example`; missing it is an argparse error (exit 2) |
| `--out` | empty | Output path. When empty, `<input-stem>_筛选结果.xlsx` beside the input (`screener.py:256`) |
| `--media-dir` | empty | Directory holding the videos. Empty means every non-text row becomes `待人工复核` |
| `--rules` | empty | JSON rules file, shallow-merged over the defaults. A missing path, invalid JSON or a non-object exits 2 with an error message |
| `--no-media` | off | Forces text-only mode; the media directory is not consulted |
| `--init-example` | off | Writes the demo workbook and returns before reading `--input` |

### Rules JSON

`rules` starts as a copy of `DEFAULT_RULES` and is then updated key by key, so unknown keys
pass through, omitted keys keep their defaults, and **lists are replaced entirely** rather
than appended (`screener.py:35-45`, `screener.py:257-263`). The file itself is validated
before the merge: a missing path, invalid JSON or a top-level non-object prints an error and
exits 2 (`screener.py:225-236`).

| Key | Type | Default | Effect |
|---|---|---|---|
| `topic` | string | `"智驾"` | Printed in the banner only; not used in matching |
| `smart_terms` | string[] | 23 terms (`智驾`, `智能驾驶`, `辅助驾驶`, `自动驾驶`, `高阶智驾`, `领航辅助`, `领航`, `自动泊车`, `记忆泊车`, `代客泊车`, `泊车`, `激光雷达`, `世界模型`, `NOA`, `NGP`, `NCA`, `ACC`, `AEB`, `APA`, `AVP`, `LCC`, `LKA`, `TJA`) | Matched against title, transcript and OCR text |
| `brand_terms` | string[] | `["Momenta"]` | Added to the matching term list for the title, transcript and OCR **only** when `count_brand_as_topic` is true |
| `exclude_terms` | string[] | `["纯话题标签", "无关样例词"]` | Checked first in both stages. Any match in the title, transcript or OCR text forces `不相关` with status `excluded`, even when a topic term also matches; the matched term is recorded in `命中来源` and `命中词` |
| `count_brand_as_topic` | bool | `false` | Adds `brand_terms` to the term list used by **both** the text and media stages |

### Environment variables

| Variable | Default | Effect |
|---|---|---|
| `WHISPER_MODEL` | `small` | Model name passed to `WhisperModel`; read once, when the ASR singleton is first built (`screener.py:102`) |

### Hard-coded tuning values

| Setting | Value | Where |
|---|---|---|
| ASR device | `cpu` | `screener.py:102` |
| ASR compute type | `int8` | `screener.py:102` |
| VAD filter | on | `screener.py:103` |
| Beam size | `1` | `screener.py:103` |
| ASR language | not specified, so faster-whisper auto-detects | `screener.py:103` |
| OCR frames per video | `6` (`n_frames`) | `screener.py:112` |
| Transcript / OCR snippet kept in output | 200 characters each | `screener.py:206-211` |
| Matched terms shown per row | 8 | `screener.py:274` |
| Media filename extension | `.mp4` only | `screener.py:82` |
| Progress print interval | every 50 rows | `screener.py:277` |

---

## CLI

The tool is a single argparse program; there is no server and no HTTP API.

```text
python screener.py [--input PATH] [--out PATH] [--media-dir PATH]
                   [--rules PATH] [--no-media] [--init-example]
```

| Option | Type | Required | Notes |
|---|---|---|---|
| `--input` | path | yes, unless `--init-example` | `.xlsx`; only worksheet 1 is read |
| `--out` | path | no | Overwritten without prompting if it exists |
| `--media-dir` | path | no | Non-recursive: files must sit directly in this directory |
| `--rules` | path | no | A missing path, invalid JSON or a non-object prints an error and exits 2 |
| `--no-media` | flag | no | Short-circuits before `find_media` |
| `--init-example` | flag | no | Ignores `--out`, `--media-dir`, `--rules` |

Exit codes: `0` on a successful run or a generated example, `2` from argparse when
`--input` is missing and from `main()` when `--rules` cannot be loaded.

### Expected media filenames

Given a URL containing a segment like `/video/1000000000000000001`, the id is
`1000000000000000001` and the first candidate is `1000000000000000001.mp4`. The second
candidate is `<excel row number>.mp4`, so the first data row (row 2 of the sheet) maps to
`2.mp4`. The id candidate is tried first; a zero-byte file is treated as absent
(`screener.py:82-85`).

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Every row is `待人工复核` with `无本地视频` | `--media-dir` points at the wrong folder, or filenames match neither candidate | Use `<media-id>.mp4` or `<data index, 1-based>.mp4` directly inside the given directory |
| A video exists but is ignored | Extension is not `.mp4`, or the file is zero bytes | Re-encode to `.mp4`; the size check is strict |
| The wrong file is analysed | A URL-derived `<media-id>.mp4` exists but is stale; it is preferred over `<data index, 1-based>.mp4` | Delete the stale file or name by row explicitly |
| All titles come out empty | Header is not on row 1, or the header text is not an exact alias | Move the header to row 1; rename `视频标题（必填）` to `视频标题` |
| Run stops with `规则文件不存在` / `规则文件无法解析` / `必须是 JSON 对象` | The `--rules` path is missing, the JSON is invalid, or the top level is not an object | Fix the path or JSON; the tool exits 2 and writes no output |
| Editing `smart_terms` had no effect | The rules list replaces the default list wholesale; a typo in a key creates a new unused key | Print `rules` or check spelling against the key table |
| A row is `不相关` with `排除词(<term>)` unexpectedly | An `exclude_terms` entry matched as a substring in the title, transcript or OCR text | Narrow or remove that exclude term; `命中词` shows what matched |
| Rows show `待人工复核` / `媒体处理异常` with status `media_error` | ASR or OCR raised an exception; the `备注` column names the failing stage | Test `faster_whisper` and `RapidOCR` standalone; check the install |
| Rows show `待人工复核` / `媒体处理异常` with status `media_gap` | `CAP_PROP_FRAME_COUNT` returned 0, so OCR had no frames to sample | Remux the file; some streamed containers lack a frame count |
| First run appears to hang | The Whisper model is being downloaded into the Hugging Face cache | Wait it out once; later runs reuse the cache |
| Output cannot be written | The target `.xlsx` is open in Excel | Close it and rerun |
| Console shows garbled Chinese on Windows | The shell code page is not UTF-8 | `chcp 65001`, as `setup.bat` does |
| `统计` lists a verdict with `0` | Expected: all three verdicts are always emitted in a fixed order | No action |

---

## Design decisions & limitations

- **No scraping, by design.** The tool is decoupled from every platform. Getting the videos
  onto disk is the operator's job, which keeps the tool out of terms-of-service territory
  and makes reruns reproducible.
- **Only the title is text evidence.** The URL, the platform and any other columns are never
  matched against topic terms. A topic tag that lives only in a tag column or in the URL is
  invisible to the text check (`screener.py:174`).
- **Substring matching, not tokenisation.** `hits()` is lowercase containment, so short
  Latin abbreviations are risky: `ACC` matches `access` and `accessories`, `APA` matches
  `apart`. Precision depends on the term list being chosen carefully.
- **Exclusion is substring-based and text-only.** `exclude_terms` is checked in the title,
  transcript and OCR text, not in the URL, platform or any other input column. A term that
  is a substring of a longer word can exclude a row unintentionally, just as `smart_terms`
  can match by accident (`screener.py:181-183`, `screener.py:200-202`).
- **Brand terms follow one flag.** `brand_terms` join the term list for the title,
  transcript and OCR only when `count_brand_as_topic` is true, so the two stages cannot
  disagree (`screener.py:175-177`, `screener.py:204`).
- **Failures are reported, not retried.** An ASR/OCR exception or a 0-frame OCR probe makes
  the row `待人工复核` with status `media_error` / `media_gap` and a note in `备注`
  (`screener.py:194-210`). The tool does not retry, switch decoders or take a second
  sample; a human has to inspect or re-run the file.
- **Six uniform frames.** There is no scene detection and no OCR over time; burned-in text
  that appears only briefly can be missed entirely. A container that reports zero frames
  yields `media_gap` instead of a silent miss, but the gap is still not filled.
- **Small CPU model.** `small` on CPU with `int8` and `beam_size=1` is a speed choice at the
  cost of transcription accuracy, especially for accented or noisy audio. No language is
  pinned, so detection is per-clip.
- **Fully sequential.** Rows are processed one at a time; the ASR and OCR singletons make
  this cheap in memory but there is no parallelism.
- **The output is a fresh workbook.** Original columns and formulas are not carried over,
  but `序号` keeps the source sheet row and `视频文件` names the analysed file, so every
  result traces back to its input row (`screener.py:272-275`).
- **Header matching is strict.** Exact equality after trim and lowercase means decorated
  headers miss. See [Input columns](#input-columns-and-aliases).
- **No tests, no CI.** There is no automated coverage of `decide()`, which is the part most
  worth testing since it is pure and side-effect free apart from file lookups.

---

## FAQ

**Does it download the videos itself?** No. `find_media` only checks for local files
(`screener.py:75-86`). There is no HTTP client anywhere in the repository.

**Do I need ffmpeg installed?** No. faster-whisper decodes via PyAV (`av` wheel); OpenCV
decodes its own frames.

**Do I need a GPU?** No. The device is hard-coded to `cpu` and the compute type to `int8`
(`screener.py:102`).

**Can I use a bigger ASR model?** Yes — set `WHISPER_MODEL`, e.g. `medium` or
`large-v3`. It is read only when the singleton is built, so it applies from the next process
start, not mid-run.

**Can I screen a topic other than cars?** Yes. Replace `smart_terms` (and optionally
`brand_terms`) via `--rules`. `topic` is cosmetic and only appears in the console banner.

**What happens to a row with a video that cannot be decoded?** It is reported as
`待人工复核` with status `media_error` (or `media_gap` when the container reports zero
frames), and the failure is written to `备注`. A decoding problem can no longer look like a
clean `不相关` result (`screener.py:194-210`).

**Why is a brand-only clip `不相关` when `count_brand_as_topic` is false?** Because the
flag now governs both stages: with `false`, `brand_terms` are not used for the title,
transcript or OCR (`screener.py:175-177`, `screener.py:204`). Set the flag to `true` to
count brands as topic evidence everywhere, or edit `brand_terms` to change the list.

**How do I force a row out even if it mentions a topic term?** Put a distinctive phrase in
`exclude_terms`. Any match in the title, transcript or OCR text yields `不相关` with status
`excluded` and records the matched term in `命中来源` / `命中词`.

**Can two videos have the same media id?** Not by accident: the candidate filename comes
from the URL, so identical URLs resolve to the same file. Use the `<data index, 1-based>.mp4` convention if
you need per-row files for duplicated links.

**Does it run on macOS or Linux?** `screener.py` is pure Python and portable. Only
`setup.bat` and `setup.ps1` are Windows-specific.

---

## License

No license file is included in this repository, and GitHub reports no license for it.
Treat the code as all rights reserved unless the author states otherwise.

## 可选：自动下载（yt-dlp）
```
python screener.py --input 数据.xlsx --download
python screener.py --input 数据.xlsx --download --cookies cookies.txt --max-download-mb 300
```
- 仅对 **yt-dlp 支持的站点**有效（YouTube/TikTok 等）。
- 抖音/小红书/视频号等强反爬站点常失败或需登录 Cookie（本工具不内置平台逆向）。
- 下载失败的行标记为「待人工复核」。