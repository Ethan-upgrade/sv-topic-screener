# 多模态短视频内容筛选框架（sv-topic-screener）

用**公开开源库**实现「标题 + 口播语音转写 + 画面文字 OCR」三路证据判定，
判断短视频是否属于某个主题（内置示例：汽车 **智驾/辅助驾驶**，可自定义）。

> 本项目为 **clean-room 通用实现**：不抓取任何平台、不内置任何第三方专有代码；
> 语音/画面素材由使用者**自行合法获取**并在本地提供。

## 功能
- 输入 Excel（列：`平台` / `视频链接` / `短视频标题`，列名支持常见别名）。
- 判定逻辑（可在 `rules.example.json` 自定义主题词表）：
  1. **文本**：标题/话题命中主题词 → 「相关」；
  2. 未命中 → 若有本地视频（`--media-dir` 下放 `<内容ID>.mp4` 或 `<行号>.mp4`）：
     - **语音转写**（faster-whisper，自带解码，无需系统 ffmpeg）
     - **画面 OCR**（OpenCV 抽帧 + RapidOCR）
     - 语音/画面命中 → 「相关」，否则「不相关」；
  3. 无视频 → 「待人工复核」。
- 输出：逐条结果 Excel（判定/命中来源/命中词/口播片段/画面片段/状态 + 统计页）。

## 安装（Windows）
1. 装 **Python 3.11+**（加入 PATH）。
2. 双击 **`setup.bat`**（自动建 `.venv` 并安装依赖；首次转写会自动下载 whisper small 模型 ~460MB，需联网）。

## 使用
```
.venv\Scripts\python.exe screener.py --init-example                 # 生成示例 Excel
.venv\Scripts\python.exe screener.py --input 数据.xlsx              # 仅文字判定（最快）
.venv\Scripts\python.exe screener.py --input 数据.xlsx --media-dir 视频目录   # 文字+语音+画面
.venv\Scripts\python.exe screener.py --input 数据.xlsx --rules rules.example.json
```
结果文件：`<原名>_筛选结果.xlsx`。

## 自定义主题
编辑 `rules.example.json`（`topic` / `smart_terms` / `brand_terms` / `count_brand_as_topic`），
换成任意主题（如“露营”“母婴”“数码”）即可复用。

## 依赖
faster-whisper / ctranslate2 / av / RapidOCR / onnxruntime / OpenCV / openpyxl。

## 合规
本工具仅用于分析**你有权访问/处理**的内容；请自行合法获取音视频素材，遵守各平台条款与当地法律。
