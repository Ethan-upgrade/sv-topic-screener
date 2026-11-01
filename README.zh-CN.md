# sv-topic-screener — 多模态短视频主题筛选器

[English](README.md) | [简体中文](README.zh-CN.md)

一个单文件的 Python 工具，用公开开源库判断 Excel 清单中的每条短视频是否属于某个主题。
它按固定顺序收集三类证据——**标题/话题文本**、faster-whisper 转写的**口播语音**、以及
OpenCV 抽帧配 RapidOCR 读出的**画面文字**——再把判定结果写成新的工作簿。内置示例主题是
汽车智驾（ADAS），全部词表都可配置。工具不抓取任何平台，也不含任何第三方专有代码，
视频素材由使用者自行提供。

---

## 概述

它针对的是一条人工筛选队列。运营或研究同学先从平台自己的后台导出短视频清单到 Excel，
再由人工逐条打开判断"这条是不是讲智驾的"。这件事慢且标准不一。本工具用确定性的关键词
匹配替掉第一遍筛选，只把真正模糊的行留给人工。

输入是一个 `.xlsx`，第一个工作表的表头行里要有标题列、链接列，以及可选的平台列（支持
多种中英文列名别名，见[输入列与别名](#输入列与别名)）。视频本身必须已经在本地磁盘上，
文件名用 URL 中提取的媒体 ID，或直接用 Excel 行号。找不到视频时，工具不会去猜音频和
画面，而是明确标记为待复核。

每一行最终只会落入四种结果之一：**文本相关**、**媒体相关**、**待人工复核**、**不相关**。
只要标题、口播或画面文本命中任一 `exclude_terms`，该行一律判为 **不相关**（状态
`excluded`），即便同时命中主题词。标题命中主题词的行根本不会打开视频文件，因此在大表上
做纯文本筛选很快。

### Clean-room 与合规边界

- 不抓取平台、不登录、不用 Cookie、不调私有接口。仓库里完全没有 URL 抓取代码，
  `find_media` 只在你传入的目录里找已存在的文件（`screener.py:75-86`）。
- 不内置任何第三方专有代码或模型权重。`requirements.txt` 里都是公开包，模型文件由各自
  的包在运行时解析。
- 是否有权处理某个视频，由运行工具的人负责。这里没有任何绕过平台条款的行为，
  因为这个工具根本不接触平台。

---

## 目录

- [概述](#概述)
- [功能](#功能)
- [架构](#架构)
- [关键流程](#关键流程)
- [数据模型](#数据模型)
- [目录结构](#目录结构)
- [快速开始](#快速开始)
- [命令](#命令)
- [配置](#配置)
- [命令行](#命令行)
- [排错](#排错)
- [设计取舍与局限](#设计取舍与局限)
- [常见问题](#常见问题)
- [许可证](#许可证)

---

## 功能

| 方面 | 行为 |
|---|---|
| 输入 | Excel `.xlsx`，只读第一个工作表，第 1 行视为表头（`screener.py:56-72`） |
| 列识别 | 对三组别名做忽略大小写的精确匹配（`screener.py:27-29`、`screener.py:48-53`） |
| 文本判定 | 标题对主题词做忽略大小写的子串匹配（`screener.py:155-157`） |
| 排除词 | 标题、口播或画面文本命中任一 `exclude_terms` 即压过主题命中，判为 `不相关` / `excluded`（`screener.py:181-183`、`screener.py:200-202`） |
| 媒体判定 | faster-whisper 转写 + OpenCV 抽 6 帧 + RapidOCR，两路结果合并（`screener.py:194-204`） |
| 视频定位 | 先按 URL 提取的 `<媒体ID>.mp4`，再退回 `<Excel行号>.mp4`（`screener.py:75-86`） |
| 短路 | 标题命中即返回，不打开任何视频（`screener.py:181-187`） |
| 失败处理 | ASR/OCR 异常与 0 帧 OCR 探针判为 `待人工复核`，状态 `media_error` / `media_gap` 并带备注，不再误判为 `不相关`（`screener.py:194-210`） |
| 判定结果 | `相关` / `待人工复核` / `不相关`，另有机器可读的 `状态` 列（`screener.py:173-211`） |
| 可追溯 | `序号` 保留源表行号；`视频文件` 记录实际匹配到的文件名（`screener.py:272-275`） |
| 输出 | 新工作簿：结果页 + `统计` 页；统计页始终列出全部三种判定（`screener.py:280-295`） |
| 示例生成 | `--init-example` 写出两行示例工作簿（`screener.py:214-222`） |
| 规则覆盖 | JSON 文件浅合并覆盖内置默认值；路径缺失或 JSON 非法直接报错（`screener.py:257-263`） |
| 纯文本模式 | `--no-media` 完全跳过语音与画面（`screener.py:245`、`screener.py:188-189`） |

没有并发、没有磁盘缓存、没有数据库，也没有任何网络代码。

---

## 架构

除两个安装脚本外，全部逻辑都在一个模块里。下图中的方框是 `screener.py` 里的函数，
模型对象是模块级单例，因此能在多行之间复用。

```mermaid
flowchart TD
    IN["Excel 输入 (.xlsx)<br/>第一个工作表，第 1 行为表头"]
    READ["read_rows() + _col()<br/>列名别名解析"]
    RULES["规则合并<br/>DEFAULT_RULES 再叠加 --rules JSON"]
    DEC["decide()"]
    EX{"标题命中排除词?"}
    TXT["hits(title, smart_terms<br/>count_brand_as_topic 为真时加 brand_terms)"]
    T1{"标题命中?"}
    NM{"--no-media?"}
    FM["find_media()<br/>先媒体ID.mp4，再行号.mp4"]
    V1{"文件存在<br/>且大小大于 0?"}
    ASR["transcribe()<br/>faster-whisper small / cpu / int8"]
    OCR["ocr_frames()<br/>抽 6 帧，OpenCV + RapidOCR"]
    MEX{"口播或画面命中排除词?"}
    MH{"口播或画面命中主题词?"}
    ERR{"ASR/OCR 异常<br/>或 OCR 0 帧?"}
    RX["不相关 / excluded"]
    R1["相关 / 文本<br/>done_text"]
    R2["待人工复核 / 文本未命中<br/>boundary_no_media"]
    R3["待人工复核 / 无本地视频<br/>boundary_no_media"]
    R4["相关 / 媒体<br/>done_media"]
    R5["不相关 / 媒体均未命中<br/>done_media"]
    R6["待人工复核 / 媒体处理异常<br/>media_error 或 media_gap"]
    OUT["结果工作簿<br/>筛选结果 + 统计"]

    IN --> READ --> DEC
    RULES --> DEC
    DEC --> EX
    EX -- 是 --> RX
    EX -- 否 --> TXT --> T1
    T1 -- 是 --> R1
    T1 -- 否 --> NM
    NM -- 是 --> R2
    NM -- 否 --> FM --> V1
    V1 -- 否 --> R3
    V1 -- 是 --> ASR
    V1 -- 是 --> OCR
    ASR --> MEX
    OCR --> MEX
    MEX -- 是 --> RX
    MEX -- 否 --> MH
    MH -- 是 --> R4
    MH -- 否 --> ERR
    ERR -- 是 --> R6
    ERR -- 否 --> R5
    R1 --> OUT
    R2 --> OUT
    R3 --> OUT
    R4 --> OUT
    R5 --> OUT
    R6 --> OUT
    RX --> OUT
```

*（来源：`screener.py:173-211`、`screener.py:239-301`）*

| 层 | 路径 | 职责 |
|---|---|---|
| 命令行 | `screener.py:239-301` | argparse、默认输出名、规则加载、逐行循环、写工作簿、控制台摘要 |
| 输入读取 | `screener.py:56-72` | 读第一个工作表、解析标题/链接/平台三列、组装 `{row, title, url, platform}` |
| 表头匹配 | `screener.py:48-53` | 对表头单元格去空白并转小写，返回首个命中的列下标 |
| 规则 | `screener.py:35-45`、`rules.example.json` | 内置词表与可覆盖的示例配置 |
| 规则加载 | `screener.py:225-236` | 解析 `--rules` JSON，拒绝缺失路径、非法 JSON 或非对象 |
| 匹配器 | `screener.py:155-157` | `hits(text, terms)`：转小写后的子串包含判断 |
| 判定 | `screener.py:173-211` | 下文记录的优先级链，含排除词与媒体失败处理 |
| 媒体定位 | `screener.py:75-86` | URL 正则、两个候选文件名、非空文件校验 |
| 语音 | `screener.py:92-106` | 懒加载 `WhisperModel` 单例，开 VAD，beam size 1；返回文本与错误说明 |
| 画面 OCR | `screener.py:112-152` | 帧数探测、均匀抽帧、`RapidOCR` 单例、错误与 0 帧缺口上报 |
| 输出写入 | `screener.py:270-300` | 结果行、`统计` 页、每 50 行打印进度、控制台分布 |
| 安装 | `setup.ps1`、`setup.bat` | 创建 `.venv` 并 `pip install -r requirements.txt`（Windows） |

---

## 关键流程

### 1. 单条记录（有本地视频）

这是完整路径。两路媒体证据彼此独立：任意一路命中即可，且无论最终是否判定为相关，
两段文本都会写进输出。

```mermaid
sequenceDiagram
    participant X as 输入工作簿
    participant M as main 循环
    participant D as decide
    participant FS as 视频目录
    participant W as faster-whisper
    participant O as OpenCV + RapidOCR
    participant OB as 输出工作簿

    X->>M: read_rows() 得到 {row, title, url, platform}
    M->>D: decide(rec, rules, media_dir, do_media)
    D->>D: hits(title, exclude_terms)
    alt 标题命中排除词
        D-->>M: 不相关, 排除词(<term>), excluded
    else 无排除词，hits(title, smart_terms [+ 品牌词，若启用])
        D-->>M: 相关, 文本(标题/话题), done_text
    else 未命中且 do_media 为假
        D-->>M: 待人工复核, 文本未命中, boundary_no_media
    else 未命中且 do_media 为真
        D->>FS: find_media(): 先媒体ID.mp4 再行号.mp4
        alt 没有文件，或文件大小为 0
            D-->>M: 待人工复核, 无本地视频, boundary_no_media
        else 找到文件
            D->>W: transcribe(video)
            W-->>D: 转写文本；失败时附带错误说明
            D->>O: ocr_frames(video, n_frames=6)
            O-->>D: OCR 文本；附带错误与 0 帧缺口说明
            alt 口播或画面命中排除词
                D-->>M: 不相关, 排除词(<term>), excluded
            else 口播或画面命中主题词
                D-->>M: 相关, 媒体(口播/画面), done_media
            else ASR/OCR 失败或 OCR 被跳过
                D-->>M: 待人工复核, 媒体处理异常, media_error 或 media_gap
            else 两路都处理完成且没有命中
                D-->>M: 不相关, 媒体均未命中, done_media
            end
        end
    end
    M->>OB: 追加一行：源行号、匹配文件名、命中词[:8]、片段、备注
    M->>OB: 统计页：全部三种判定、合计、耗时秒
```

*（来源：`screener.py:173-211`、`screener.py:270-276`）*

### 2. 批量启动、规则合并与模型复用

模型在首次使用时构建，之后整个运行期复用。`--init-example` 在任何输入文件被打开之前
就返回；缺少 `--input` 是 argparse 报错（退出码 2），而不是静默空跑。

```mermaid
sequenceDiagram
    participant U as 用户
    participant A as argparse
    participant R as 规则
    participant X as Excel
    participant L as 逐行循环
    participant C as 模型单例

    U->>A: screener.py --input ... --media-dir ...
    alt 传了 --init-example
        A->>X: 写示例工作簿（默认 example_input.xlsx）
        A-->>U: 退出 0
    else 没传 --input
        A-->>U: argparse 报错，退出 2
    end
    alt --rules 路径缺失、JSON 非法或顶层不是对象
        A-->>U: 错误信息，退出 2
    end
    A->>R: 复制 DEFAULT_RULES，再用 --rules 的 JSON update()
    R-->>A: 生效规则（列表整体替换）
    A->>X: read_rows() 读取工作表 1
    X-->>A: 记录列表
    loop 每条记录
        A->>L: decide(...)
        L->>C: 首次语音或 OCR 调用构造单例
        C-->>L: 之后所有行复用
    end
    A->>X: 保存 原文件名 + _筛选结果.xlsx
    A-->>U: 记录数、耗时、判定分布、输出路径
```

*（来源：`screener.py:249-264`、`screener.py:100-103`、`screener.py:122-129`、`screener.py:296-300`）*

### 3. 视频定位与抽帧细节

命名文件之前值得先看这一段，因为一个陈旧文件会静默胜出。

```mermaid
sequenceDiagram
    participant D as 媒体定位
    participant U as URL 字符串
    participant F as 视频目录
    participant CV as OpenCV

    D->>U: 正则 video|note|explore|sph + 斜杠 + 至少 6 位 ID 字符
    U-->>D: 媒体 ID 或空
    D->>F: 先试 媒体ID.mp4
    F-->>D: 存在且大小大于 0 才算接受
    D->>F: 否则试 Excel行号.mp4
    F-->>D: 同样的校验
    D->>CV: VideoCapture，读取 CAP_PROP_FRAME_COUNT
    CV-->>D: 总帧数
    Note over D: 总帧数为 0 时记为 media_gap 缺口，不再静默跳过
    D->>CV: 对 k = 0..5 定位到 int(total * (k + 0.5) / 6)
    CV-->>D: 六等分段中心处的 6 帧
```

帧号公式 `int(total * (k + 0.5) / n_frames)` 取的是每一段的中心而不是首尾帧，因为首尾
往往是黑屏或片头。*（来源：`screener.py:82-85`、`screener.py:124-132`）*

---

## 数据模型

这里的"数据库"就是两个 Excel 工作簿。输入列按别名定位，因此列顺序无所谓，没被读取的列
直接忽略；输出列则始终按固定顺序写出。

```mermaid
erDiagram
    INPUT_SHEET {
        string title "标题 - 与主题词比对"
        string url "链接 - 只用来提取媒体 ID"
        string platform "平台 - 原样带到结果"
        string other "其余列 - 忽略"
    }
    RESULT_SHEET {
        int xuhao "序号 - 源表行号（表头是第 1 行）"
        string pingtai "平台"
        string biaoti "标题"
        string lianjie "链接"
        string panding "判定 - 相关 / 待人工复核 / 不相关"
        string laiyuan "命中来源"
        string mingzhongci "命中词 - 最多 8 个，用顿号连接"
        string koubo "口播片段 - 前 200 字符"
        string huamian "画面片段 - 前 200 字符"
        string shipinwenjian "视频文件 - 实际匹配到的本地文件名；未查视频时为空"
        string beizhu "备注 - ASR/OCR 错误或 OCR 0 帧缺口说明"
        string zhuangtai "状态 - done_text / done_media / excluded / media_error / media_gap / boundary_no_media"
    }
    STATS_SHEET {
        string metric "判定列：全部三种判定、合计、耗时秒"
        int value "条数列：条数（含 0）、总行数、或秒数"
    }
    INPUT_SHEET ||--o{ RESULT_SHEET : "一行进，一行出"
```

*（来源：`screener.py:56-72`、`screener.py:270-295`）*

### 输入列与别名

表头匹配是 `str().strip().lower()` 之后的**精确**比较，不是子串匹配，也不是模糊匹配
（`screener.py:48-53`）。所以 `视频标题（必填）` 这种表头不会被认成标题列。三列都是
可选的；缺哪列，对应字段在每一行都是空串。

| 字段 | 接受的表头名 | 说明 |
|---|---|---|
| 标题 | `标题`、`视频标题`、`短视频标题`、`作品标题`、`笔记标题`、`title` | 唯一参与主题词比对的字段 |
| 链接 | `链接`、`视频链接`、`作品链接`、`笔记链接`、`url`、`url链接` | 只用于提取媒体 ID，从不被打开 |
| 平台 | `平台`、`发布平台`、`媒体平台`、`platform` | 原样写入输出，不参与匹配 |

### 输出列

| # | 列 | 内容 | 来源 |
|---|---|---|---|
| 1 | `序号` | 源表行号：表头是第 1 行，因此第一条数据是 `2` | `screener.py:272-273` |
| 2 | `平台` | 输入的平台值 | `screener.py:273` |
| 3 | `标题` | 输入的标题值 | `screener.py:273` |
| 4 | `链接` | 输入的链接值 | `screener.py:273` |
| 5 | `判定` | `相关` / `待人工复核` / `不相关` | `screener.py:181-211` |
| 6 | `命中来源` | `文本(标题/话题)`、`排除词(<term>)`、`文本未命中`、`无本地视频`、`媒体(口播/画面)`、`媒体处理异常`、`媒体均未命中` | `screener.py:181-211` |
| 7 | `命中词` | 命中词用 `、` 连接，**只取前 8 个** | `screener.py:274` |
| 8 | `口播片段` | 转写文本前 200 字符；纯文本模式或无视频时为空 | `screener.py:206-211` |
| 9 | `画面片段` | OCR 文本前 200 字符，同上 | `screener.py:206-211` |
| 10 | `视频文件` | 实际匹配到的本地文件名（`<媒体ID>.mp4` 或 `<行号>.mp4`）；未查视频时为空 | `screener.py:275` |
| 11 | `备注` | 媒体处理备注：ASR/OCR 失败信息或 OCR 0 帧缺口 | `screener.py:275` |
| 12 | `状态` | `done_text`、`boundary_no_media`、`excluded`、`done_media`、`media_error`、`media_gap` | `screener.py:181-211` |

`统计` 页只有两列：`判定` 和 `条数`。它始终按固定顺序列出全部三种判定——`相关`、
`待人工复核`、`不相关`，数量为 0 也照常输出；接着是 `合计` 行（总行数），最后是
`耗时秒` 行（耗时保留一位小数，`screener.py:288-295`）。

### 判定优先级

下面是 `decide()` 内部的真实顺序。先命中的分支直接返回，后面的分支不再求值。

| 步骤 | 条件 | `判定` | `命中来源` | `命中词` | `状态` | 来源 |
|---|---|---|---|---|---|---|
| 1 | 标题包含任意排除词 | `不相关` | `排除词(<term>)` | 命中的排除词 | `excluded` | `screener.py:181-183` |
| 2 | 无排除词，标题包含任意主题词 | `相关` | `文本(标题/话题)` | 命中的词 | `done_text` | `screener.py:185-187` |
| 3 | 标题未命中，且 `do_media` 为假（`--no-media`） | `待人工复核` | `文本未命中` | 空 | `boundary_no_media` | `screener.py:188-189` |
| 4 | 标题未命中，启用了媒体，但没有可用视频文件 | `待人工复核` | `无本地视频` | 空 | `boundary_no_media` | `screener.py:190-192` |
| 5 | 找到视频，口播**或**画面文本出现任意排除词 | `不相关` | `排除词(<term>)` | 命中的排除词 | `excluded` | `screener.py:200-202` |
| 6 | 找到视频，无排除词，口播**或**画面文本出现任意主题词 | `相关` | `媒体(口播/画面)` | 去重后的词 | `done_media` | `screener.py:204-206` |
| 7 | 找到视频，没有命中词，但 ASR/OCR 失败或 OCR 被跳过（0 帧） | `待人工复核` | `媒体处理异常` | 空 | `media_error` / `media_gap` | `screener.py:207-210` |
| 8 | 找到视频，两路都处理完成且都没命中 | `不相关` | `媒体均未命中` | 空 | `done_media` | `screener.py:211` |

几个值得说清的后果：

- 排除词在两个阶段都最先判断。标题同时含排除词和主题词时判 `不相关`；口播/画面文本
  同理。
- 步骤 3 在步骤 4 之前。开了 `--no-media` 时，行根本不会走到 `find_media`，所以
  `命中来源` 永远是 `文本未命中`，即便磁盘上恰好有对应视频。
- 步骤 2 在媒体路径之前。标题命中完全不需要解码，所以标题大多相关的表里，媒体路径只跑
  剩下的那一部分。
- 处理失败本身不会得出 `不相关`。只有完整处理、既无排除词也无主题命中时才判
  `不相关`；否则是 `待人工复核` 并带备注。

### 单行状态流转

`状态` 列有六个取值，且都是终态——工具不会回头重跑或重新排队任何一行。

```mermaid
stateDiagram-v2
    [*] --> TextCheck
    TextCheck --> excluded: 标题命中排除词
    TextCheck --> done_text: 标题命中主题词
    TextCheck --> boundary_no_media: 未命中且 do_media 为假
    TextCheck --> MediaLookup: 未命中且 do_media 为真
    MediaLookup --> boundary_no_media: 没有可用视频
    MediaLookup --> ASR_OCR: 找到视频
    ASR_OCR --> excluded: 口播或画面命中排除词
    ASR_OCR --> done_media: 命中主题词
    ASR_OCR --> media_error: ASR/OCR 失败且未命中
    ASR_OCR --> media_gap: OCR 0 帧且未命中
    ASR_OCR --> done_media: 完整处理且未命中
    done_text --> [*]
    done_media --> [*]
    excluded --> [*]
    media_error --> [*]
    media_gap --> [*]
    boundary_no_media --> [*]
```

---

## 目录结构

```text
sv-topic-screener/
├─ screener.py          # 全部逻辑：读取、匹配、判定、语音、OCR、写出
├─ rules.example.json   # 可覆盖的词表（智驾示例主题）
├─ requirements.txt     # 运行时依赖（语音、OCR、Excel）
├─ setup.ps1            # 创建 .venv 并安装依赖（Windows）
├─ setup.bat            # cmd 包装，切 UTF-8 代码页后调用 setup.ps1
├─ README.md            # 英文版
├─ README.zh-CN.md      # 本文件（简体中文）
└─ .gitignore           # 虚拟环境、视频、模型缓存、生成的工作簿
```

`.gitignore` 排除了 `.venv/`、`__pycache__/`、`*.pyc`、`*.mp4`、`media/`、`hf_cache/`、
`*_筛选结果.xlsx` 和 `example_input.xlsx`——也就是视频、模型缓存和生成的工作簿都不会
进版本库。

---

## 快速开始

### 环境要求

- **Python 3.11+**。`setup.ps1` 在 `python` 不在 `PATH` 上时会提示需要 3.11+
  （`setup.ps1:5`），脚本本身并不校验版本号。
- 首次安装需要联网：`pip` 要下载 wheel；首次跑语音时还要下载 Whisper 模型。
- **不需要系统 ffmpeg**。faster-whisper 通过 PyAV 解码，`av` 随 faster-whisper 一起安装
  （`screener.py:100`）。

### 安装（Windows，脚本方式）

```powershell
git clone https://github.com/Ethan-upgrade/sv-topic-screener.git
cd sv-topic-screener
.\setup.bat
```

`setup.bat` 先 `chcp 65001` 再调用 `setup.ps1`；后者创建 `.venv`，执行
`pip install --upgrade pip`，然后 `pip install -r requirements.txt`（`setup.ps1:6-10`）。

### 安装（手动，任意系统）

```bash
python -m venv .venv
# Windows
.venv\Scripts\python.exe -m pip install -r requirements.txt
# macOS / Linux
.venv/bin/python -m pip install -r requirements.txt
```

### 依赖说明

| 包 | 作用 |
|---|---|
| `faster-whisper` | 语音引擎，`WhisperModel`（`screener.py:100-103`）；自带 `ctranslate2`、`av`、`onnxruntime` 后端 |
| `opencv-python-headless` | `cv2.VideoCapture`、帧数与定位（`screener.py:122-131`） |
| `rapidocr` | OCR 引擎，`RapidOCR()`（`screener.py:127-129`）；自带 `onnxruntime` |
| `openpyxl` | 读写 `.xlsx`（`screener.py:25`） |

`requirements.txt` 只声明 `screener.py` 直接导入的包；推理后端与数值/图像辅助库由两个
引擎传递安装。

### 首次运行

```powershell
.venv\Scripts\python.exe screener.py --init-example
.venv\Scripts\python.exe screener.py --input example_input.xlsx
.venv\Scripts\python.exe screener.py --input example_input.xlsx --media-dir .\media
```

第二条实际上只做文本判定（没给 `--media-dir`），所以没命中的行都会落到 `待人工复核`。
第三条会去找 `media\<媒体ID>.mp4` 或 `media\<行号>.mp4`。

---

## 命令

| 命令 | 作用 |
|---|---|
| `setup.bat` / `powershell -File setup.ps1` | 创建 `.venv`、升级 pip、安装依赖 |
| `python screener.py --init-example` | 在 `example_input.xlsx` 写两行示例并退出 0 |
| `python screener.py --init-example --input demo.xlsx` | 示例改写到 `demo.xlsx` |
| `python screener.py --input data.xlsx` | 纯文本判定，不打开任何视频 |
| `python screener.py --input data.xlsx --no-media` | 与上一条相同，即便给了视频目录也强制纯文本 |
| `python screener.py --input data.xlsx --media-dir videos` | 先文本，剩余行再做语音 + 画面 |
| `python screener.py --input data.xlsx --rules my-rules.json` | 使用自定义词表 |
| `python screener.py --input data.xlsx --out result.xlsx` | 指定输出路径 |

仓库里没有测试套件、没有 lint 配置，也没有 CI 工作流。

---

## 配置

### 命令行参数

| 参数 | 默认值 | 行为 |
|---|---|---|
| `--input` | 空 | 输入工作簿。除 `--init-example` 外必填，缺失即 argparse 报错（退出 2） |
| `--out` | 空 | 输出路径。为空时在输入文件旁生成 `<原文件名>_筛选结果.xlsx`（`screener.py:256`） |
| `--media-dir` | 空 | 存放视频的目录。为空则所有未命中文本的行都变成 `待人工复核` |
| `--rules` | 空 | JSON 规则文件，浅合并覆盖默认值。路径缺失、JSON 非法或顶层不是对象时报错并以 2 退出 |
| `--no-media` | 关 | 强制纯文本模式，不查视频目录 |
| `--init-example` | 关 | 写示例工作簿，在读取 `--input` 之前就返回 |

### 规则 JSON

`rules` 先复制 `DEFAULT_RULES`，再逐键 update。因此未知键会原样保留，省略的键保留默认值，
而**列表是整体替换**，不是追加（`screener.py:35-45`、`screener.py:257-263`）。规则文件
在合并前会先校验：路径缺失、JSON 非法或顶层不是对象都会打印错误并以 2 退出
（`screener.py:225-236`）。

| 键 | 类型 | 默认值 | 作用 |
|---|---|---|---|
| `topic` | 字符串 | `"智驾"` | 只出现在控制台横幅里，不参与匹配 |
| `smart_terms` | 字符串数组 | 23 个词（`智驾`、`智能驾驶`、`辅助驾驶`、`自动驾驶`、`高阶智驾`、`领航辅助`、`领航`、`自动泊车`、`记忆泊车`、`代客泊车`、`泊车`、`激光雷达`、`世界模型`、`NOA`、`NGP`、`NCA`、`ACC`、`AEB`、`APA`、`AVP`、`LCC`、`LKA`、`TJA`） | 同时用于标题、口播和画面 |
| `brand_terms` | 字符串数组 | `["Momenta"]` | **只有** `count_brand_as_topic` 为真时才加入标题、口播和画面的匹配词表 |
| `exclude_terms` | 字符串数组 | `["纯话题标签", "无关样例词"]` | 两个阶段都最先检查。标题、口播或画面文本命中任意一项即判 `不相关` / `excluded`，即使同时命中主题词；命中词会写进 `命中来源` 和 `命中词` |
| `count_brand_as_topic` | 布尔 | `false` | 把 `brand_terms` 加入**文本与媒体两个阶段**共用的匹配词表 |

### 环境变量

| 变量 | 默认值 | 作用 |
|---|---|---|
| `WHISPER_MODEL` | `small` | 传给 `WhisperModel` 的模型名；只在 ASR 单例首次构建时读取（`screener.py:102`） |

### 写死的调参值

| 设置 | 值 | 位置 |
|---|---|---|
| 语音设备 | `cpu` | `screener.py:102` |
| 计算类型 | `int8` | `screener.py:102` |
| VAD 过滤 | 开启 | `screener.py:103` |
| beam size | `1` | `screener.py:103` |
| 语音语言 | 未指定，由 faster-whisper 自动检测 | `screener.py:103` |
| 每条视频 OCR 帧数 | `6`（`n_frames`） | `screener.py:112` |
| 写入输出的文本片段长度 | 各 200 字符 | `screener.py:206-211` |
| 每行展示的命中词 | 8 个 | `screener.py:274` |
| 视频文件扩展名 | 仅 `.mp4` | `screener.py:82` |
| 进度打印间隔 | 每 50 行 | `screener.py:277` |

---

## 命令行

这是一个单文件 argparse 程序，没有服务端，也没有 HTTP 接口。

```text
python screener.py [--input PATH] [--out PATH] [--media-dir PATH]
                   [--rules PATH] [--no-media] [--init-example]
```

| 选项 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `--input` | 路径 | 是（`--init-example` 除外） | `.xlsx`；只读第一个工作表 |
| `--out` | 路径 | 否 | 已存在时直接覆盖，不提示 |
| `--media-dir` | 路径 | 否 | 不递归：文件必须直接放在这个目录里 |
| `--rules` | 路径 | 否 | 路径缺失、JSON 非法或顶层不是对象都会打印错误并以 2 退出 |
| `--no-media` | 开关 | 否 | 在 `find_media` 之前短路 |
| `--init-example` | 开关 | 否 | 忽略 `--out`、`--media-dir`、`--rules` |

退出码：正常运行或生成示例返回 `0`；缺少 `--input` 时由 argparse 返回 `2`，`--rules`
无法加载时由 `main()` 返回 `2`。

### 视频文件命名约定

如果 URL 里有类似 `/video/1000000000000000001` 的片段，媒体 ID 就是
`1000000000000000001`，第一个候选文件是 `1000000000000000001.mp4`。第二个候选是
`<Excel行号>.mp4`，所以第一条数据行（表里第 2 行）对应 `2.mp4`。先试 ID 候选；
零字节文件视为不存在（`screener.py:82-85`）。

---

## 排错

| 现象 | 原因 | 处理 |
|---|---|---|
| 所有行都是 `待人工复核` + `无本地视频` | `--media-dir` 指错了目录，或文件名两个候选都不匹配 | 把 `<媒体ID>.mp4` 或 `<行号>.mp4` 直接放进该目录 |
| 明明有视频却被忽略 | 扩展名不是 `.mp4`，或文件是 0 字节 | 转成 `.mp4`；大小检查是严格的 |
| 分析的不是想要的那个文件 | URL 提取出的 `<媒体ID>.mp4` 存在但已经过期，它优先于 `<行号>.mp4` | 删掉过期文件，或统一用行号命名 |
| 所有标题都是空的 | 表头不在第 1 行，或表头文字不是精确别名 | 把表头移到第 1 行；把 `视频标题（必填）` 改成 `视频标题` |
| 运行直接报 `规则文件不存在` / `规则文件无法解析` / `必须是 JSON 对象` | `--rules` 路径缺失、JSON 非法，或顶层不是对象 | 修正路径或 JSON；工具以 2 退出且不写输出 |
| 改了 `smart_terms` 没变化 | 列表是整体替换；键名写错会变成一个没用的新键 | 打印 `rules` 或对照键名表检查拼写 |
| 某行意外变成 `不相关` / `排除词(<term>)` | 标题、口播或画面文本子串命中了某个 `exclude_terms` | 收窄或删掉该排除词；`命中词` 会显示命中的是哪个 |
| 行显示 `待人工复核` / `媒体处理异常`，状态 `media_error` | ASR 或 OCR 抛了异常，`备注` 列会写明是哪一路失败 | 单独测试 `faster_whisper` 和 `RapidOCR`，检查安装 |
| 行显示 `待人工复核` / `媒体处理异常`，状态 `media_gap` | `CAP_PROP_FRAME_COUNT` 返回 0，OCR 没有可抽的帧 | 重新封装文件；部分流式容器不提供帧数 |
| 首次运行像是卡住了 | Whisper 模型正在下载到 Hugging Face 缓存 | 等一次即可，之后复用缓存 |
| 输出写不进去 | 目标 `.xlsx` 正在 Excel 里打开 | 关掉再跑 |
| Windows 控制台中文乱码 | 终端代码页不是 UTF-8 | 执行 `chcp 65001`，`setup.bat` 就是这么做的 |
| `统计` 页出现数量为 0 的判定 | 属预期：三种判定始终按固定顺序全部输出 | 无需处理 |

---

## 设计取舍与局限

- **有意不抓取。** 工具与所有平台解耦，把视频落到磁盘是使用者的工作。这让它远离平台
  条款问题，也让复跑可复现。
- **文本证据只有标题。** 链接、平台以及其它任何列都不参与主题词比对。只存在于标签列或
  URL 里的话题标签，对文本判定完全不可见（`screener.py:174`）。
- **子串匹配，不做分词。** `hits()` 是小写包含判断，短的拉丁缩写很容易误伤：`ACC` 会
  命中 `access` 和 `accessories`，`APA` 会命中 `apart`。精度取决于词表挑得够不够细。
- **排除词是子串匹配，且只看文本。** `exclude_terms` 只在标题、口播和画面文本里检查，
  不看链接、平台或其它输入列。和 `smart_terms` 一样，短词可能误伤更长的单词
  （`screener.py:181-183`、`screener.py:200-202`）。
- **品牌词统一由一个开关控制。** 只有 `count_brand_as_topic` 为真时，`brand_terms` 才会
  加入标题、口播和画面的匹配词表，两个阶段不会再各判各的（`screener.py:175-177`、
  `screener.py:204`）。
- **失败会报出来，但不会重试。** ASR/OCR 异常或 OCR 0 帧会让该行变成 `待人工复核`，
  状态为 `media_error` / `media_gap`，原因写进 `备注`（`screener.py:194-210`）。工具不会
  重试、切换解码器或二次抽帧，需要人工检查或重跑文件。
- **只有 6 帧且均匀分布。** 没有镜头检测，也不做跨时间 OCR；只闪一下的画面文字可能被
  完全漏掉。容器报 0 帧时会得到 `media_gap` 而不是静默漏判，但缺口本身仍然没有被补上。
- **CPU 小模型。** `small` + `int8` + `beam_size=1` 是用精度换速度的选择，带口音或嘈杂
  音频上尤其明显。没有固定语言，逐条自动检测。
- **完全串行。** 逐行处理；语音与 OCR 单例让内存开销可控，但没有并行。
- **输出是全新工作簿。** 原始列和公式不会保留，但 `序号` 保留源表行号，`视频文件`
  记录实际分析的文件名，因此每条结果都能回溯到输入行（`screener.py:272-275`）。
- **表头匹配很严格。** 去空白转小写后的精确相等，带装饰的表头会漏。见
  [输入列与别名](#输入列与别名)。
- **没有测试、没有 CI。** `decide()` 是最值得测的部分（除了文件查找外基本是纯函数），
  但目前没有任何自动化覆盖。

---

## 常见问题

**它会自己下载视频吗？** 不会。`find_media` 只检查本地文件（`screener.py:75-86`），
仓库里没有任何 HTTP 客户端。

**需要安装 ffmpeg 吗？** 不需要。faster-whisper 通过 PyAV（`av` wheel）解码，OpenCV 自己
解码抽帧。

**需要 GPU 吗？** 不需要。设备写死为 `cpu`，计算类型写死为 `int8`（`screener.py:102`）。

**能换更大的语音模型吗？** 可以，设置 `WHISPER_MODEL`，例如 `medium` 或 `large-v3`。
它只在单例构建时读取，所以下次启动进程才生效，不能中途切换。

**除了汽车，能筛别的主题吗？** 能。用 `--rules` 替换 `smart_terms`（以及可选的
`brand_terms`）。`topic` 只用于控制台横幅，不影响匹配。

**视频解不出来时会怎样？** 会判为 `待人工复核`，状态是 `media_error`（容器报 0 帧时是
`media_gap`），失败原因写进 `备注`。解码问题不会再伪装成"干净的 `不相关` 结果"
（`screener.py:194-210`）。

**为什么 `count_brand_as_topic` 为假时，只提品牌的视频不是 `相关`？** 因为该开关现在
同时管文本和媒体两个阶段：为 `false` 时，`brand_terms` 不参与标题、口播和画面匹配
（`screener.py:175-177`、`screener.py:204`）。想让品牌词在所有环节都算主题证据，把开关
设为 `true`；想换品牌词，改 `brand_terms`。

**某行明明提到主题词，怎么强制排除？** 把有辨识度的短语写进 `exclude_terms`。标题、
口播或画面文本命中即判 `不相关` / `excluded`，命中的词会记录在 `命中来源` 和 `命中词`。

**两条视频会用到同一个媒体 ID 吗？** 不会凭空发生：候选文件名来自 URL，相同 URL 自然
指向同一个文件。如果重复链接需要各自独立的视频，就用 `<行号>.mp4` 的约定。

**能在 macOS 或 Linux 上跑吗？** `screener.py` 是纯 Python，可移植。只有 `setup.bat` 和
`setup.ps1` 是 Windows 专用。

---

## 许可证

仓库中没有许可证文件，GitHub 也显示该项目没有 license。除非作者另有说明，否则视为
保留所有权利。

## YouTube 与登录 Cookie（可选）
- **YouTube 需要 yt-dlp 的 JS 运行时**：安装 Deno（推荐）或 Node，运行时传 `--js-runtime deno`（或 `node`）。
- 需要登录/地区限制的内容：用 `--cookies cookies.txt`，或 `--cookies-from-browser chrome`（支持 chrome/edge/firefox 等）。
- 示例：
```
python screener.py --input 数据.xlsx --download --js-runtime node --cookies-from-browser chrome
```
- 仍不支持：抖音/小红书/视频号等强反爬站点（需平台逆向，本公开版不内置）。
- 若 YouTube 返回 403（反爬），用 `--cookies-from-browser chrome`（或 `--cookies cookies.txt`）带登录态。

## 本地视频目录（--media-dir）详细用法

### 1) 目录结构示例
```
视频目录/
├─ 1.mp4        # 第 1 条数据（数据序号从 1 开始，表头不算）
├─ 2.mp4
├─ 3.mkv        # 支持 mp4 / mkv / webm / mov
└─ ...
```

### 2) 命名规则（二选一，优先用「内容ID」）
1. **`<数据序号>.mp4`**：按 Excel 的**数据行顺序**编号——第 1 条数据 = `1.mp4`，第 2 条 = `2.mp4`（**不是 Excel 行号**：表头是第 1 行，第一条数据在 Excel 第 2 行，但文件名仍是 `1.mp4`）。
2. **`<内容ID>.mp4`**：从视频链接里提取的 ID。例：`https://www.douyin.com/video/7345678901234567890` → `7345678901234567890.mp4`。

> 两种都找不到时，该行判为「待人工复核」（状态 `boundary_no_media`），不会报错中断。

### 3) 批量准备建议
- 用浏览器手动下载 / 平台官方导出，然后按上面规则重命名（可用批量重命名工具）；
- 全部放同一个文件夹，运行时 `--media-dir 该文件夹`；
- 单条建议 < 300MB；CPU 上 whisper(small) 转写 + 6 帧 OCR，视频越长越慢（约 20–40 秒/条）。

### 4) 运行示例
```
python screener.py --input 数据.xlsx --media-dir D:\videos
```

### 5) 输出状态说明
| 状态 | 含义 |
|---|---|
| `done_text` | 标题/话题命中，直接判相关（未用视频） |
| `done_media` | 已做语音转写 + 画面OCR 后判定 |
| `boundary_no_media` | 文本未命中且找不到对应视频 → 待人工复核 |

### 6) 目录结构小贴士
- 视频文件名不要带空格/特殊字符（`1.mp4` 最稳）；
- 若同一批里有重复链接，按数据序号命名更直观；
- 想先看哪些行需要视频：先跑 `--no-media`，`boundary_no_media` 的就是要补视频的行。
