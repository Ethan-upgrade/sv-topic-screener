
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
