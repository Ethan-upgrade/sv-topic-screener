
## Optional: auto-download (yt-dlp)
```
python screener.py --input data.xlsx --download
python screener.py --input data.xlsx --download --cookies cookies.txt --max-download-mb 300
```
- Works only for **sites supported by yt-dlp** (YouTube/TikTok, etc.).
- Anti-bot sites (Douyin/Xiaohongshu/WeChat Channels) often fail or need login cookies (no platform reverse-engineering included).
- Rows that fail to download are marked `needs review`.