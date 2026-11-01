
## Optional: auto-download (yt-dlp)
```
python screener.py --input data.xlsx --download
python screener.py --input data.xlsx --download --cookies cookies.txt --max-download-mb 300
```
- Works only for **sites supported by yt-dlp** (YouTube/TikTok, etc.).
- Anti-bot sites (Douyin/Xiaohongshu/WeChat Channels) often fail or need login cookies (no platform reverse-engineering included).
- Rows that fail to download are marked `needs review`.

## YouTube & login cookies (optional)
- **YouTube requires a yt-dlp JS runtime**: install Deno (recommended) or Node, then pass `--js-runtime deno` (or `node`).
- Login/region-restricted content: use `--cookies cookies.txt`, or `--cookies-from-browser chrome` (chrome/edge/firefox supported).
- Example:
```
python screener.py --input data.xlsx --download --js-runtime node --cookies-from-browser chrome
```
- Still unsupported: Douyin / Xiaohongshu / WeChat Channels (anti-bot; platform reverse-engineering is not included).
- If YouTube returns 403 (anti-bot), use `--cookies-from-browser chrome` (or `--cookies cookies.txt`).

## Local video directory (`--media-dir`) — detailed usage

### 1) Example directory layout
```
videos/
├─ 1.mp4        # data row #1 (1-based, header excluded)
├─ 2.mp4
├─ 3.mkv        # mp4 / mkv / webm / mov supported
└─ ...
```

### 2) Naming rules (either one; content ID takes priority)
1. **`<data-index>.mp4`** — 1-based **data order**: first data row = `1.mp4`, second = `2.mp4` (this is **not** the Excel row number: header is row 1, first data row is row 2, but the file is still `1.mp4`).
2. **`<content-id>.mp4`** — the ID extracted from the video URL, e.g. `https://www.douyin.com/video/7345678901234567890` -> `7345678901234567890.mp4`.

> If neither is found, the row is marked `needs review` (status `boundary_no_media`); it never aborts the run.

### 3) Batch preparation tips
- Download manually in a browser or export officially, then rename per the rules above;
- Put all files in one folder and pass `--media-dir <folder>`;
- Keep each file < ~300MB; on CPU, whisper(small) + 6-frame OCR takes ~20–40s per video.

### 4) Run
```
python screener.py --input data.xlsx --media-dir D:\videos
```

### 5) Output status
| Status | Meaning |
|---|---|
| `done_text` | title/tags hit -> relevant (video not used) |
| `done_media` | ASR + on-screen OCR applied |
| `boundary_no_media` | text miss and no matching video -> needs review |

### 6) Tips
- Avoid spaces/special characters in filenames (`1.mp4` is safest);
- For duplicate URLs, data-index naming is clearer;
- To see which rows need videos, first run `--no-media`; rows with `boundary_no_media` are the ones to fill.
