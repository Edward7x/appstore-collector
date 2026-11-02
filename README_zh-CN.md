# App Store 竞品资料采集工具

[English](README.md) | 简体中文

采集 App Store 美国区和中国区的公开应用资料、介绍、图标与截图，生成包含图片预览的七工作表 Excel，同时保存 CSV、JSON 和原始素材，供竞品研究使用。

运行环境：Python 3.10 或更高版本，能够访问 Apple Lookup API 与图片 CDN。Excel 导出依赖 openpyxl 和 Pillow，无需登录 App Store 账号。

## 目录

- [快速开始](#快速开始)
- [使用案例](#使用案例)
- [参数说明](#参数说明)
- [输出说明](#输出说明)
- [常见问题](#常见问题)
- [已知限制](#已知限制)
- [开发维护](#开发维护)

## 快速开始

### 1. 检查 Python 并安装依赖

```powershell
python --version
python -m pip install -r requirements.txt
```

Windows 中若找不到 `python`，可改用 `py -3`；macOS、Linux 通常使用 `python3`。安装依赖和运行采集时应使用同一个 Python 解释器。

### 2. 准备竞品清单

编辑 `apps.txt`，每行填一个 App Store 链接或者 App ID，例如：

```text
# 每行一个应用
333903271
https://apps.apple.com/us/app/id284882215
```

### 3. 执行采集任务

在当前目录下打开 PowerShell，执行：

```powershell
py -3 collector.py --product-name "收藏助手" --input apps.txt --output output
```

程序会分别采集美国区和中国区的数据，并在 `output` 文件夹中创建归档。

### 4. 尝试发现额外图片 [可选]

```powershell
py -3 collector.py --product-name "收藏助手" --input apps.txt --output output --web-fallback
```

它会在正常采集的基础上补充公开网页中的候选图片，但这些额外图片未必全部是宣传截图。

### 采集状态含义对照表

| 状态 | 含义 |
| --- | --- |
| `complete` | 本次采集完成，没有媒体下载失败；不代表 API 提供了商店中的全部素材 |
| `partial` | 已获取元数据，但部分媒体下载失败 |
| `not_available` | Lookup 未在该地区返回应用 |
| `error` | 请求、缓存读取或其他采集步骤失败 |

## 使用案例

以下示例假定已安装依赖，且 `apps-readlater.txt` 已填入目标应用。

### 只采集中国区

```powershell
python collector.py --product-name "收藏助手" --input apps-readlater.txt --output "output/收藏助手" --regions cn
```

此参数只限制本次采集请求。如果该目录中已有同一批 App ID 的美国区缓存，汇总仍可能包含这些记录。需要仅包含新选地区的报告时，使用新目录。

### 更新已有研究

```powershell
python collector.py --product-name "收藏助手" --input apps-readlater.txt --output "output/收藏助手" --refresh
```

### 独立保存一个日期批次

```powershell
python collector.py --product-name "收藏助手" --input apps-readlater.txt --output "output/收藏助手/2026-11-02"
```

### 调整请求间隔、重试与超时

```powershell
python collector.py --product-name "收藏助手" --input apps-readlater.txt --output "output/收藏助手" --interval 5 --retries 5 --timeout 30
```

### 尝试补充公开网页图片

```powershell
python collector.py --product-name "收藏助手" --input apps-readlater.txt --output "output/收藏助手" --web-fallback
```

补充图片标记为 `web-unclassified`，可能包含无关素材和重复图片，使用前需人工核对。

## 参数说明

运行 `python collector.py --help` 查看命令行帮助。

| 参数 | 默认值 | 说明 |
| --- | --- | --- |
| `--product-name` | 必填 | 本次研究的产品名，用于工作簿命名 |
| `--input` | `apps.txt` | UTF-8 输入文件，支持 UTF-8 BOM |
| `--output` | `output` | 输出目录 |
| `--excel-language` | `zh-CN` | Excel 界面语言，可选 `zh-CN` 或 `en`；原始内容不翻译 |
| `--regions` | `us cn` | 选择一个或两个地区 |
| `--interval` | `3.5` | Lookup 请求启动间隔，单位秒，最小为 3 |
| `--retries` | `4` | 每次请求的最大重试次数，最小为 0 |
| `--timeout` | `25` | 请求超时秒数，必须大于 0 |
| `--refresh` | 关闭 | 重新获取元数据和媒体 |
| `--web-fallback` | 关闭 | 从公开 HTML 发现额外的未分类图片 |
| `--verbose` | 关闭 | 输出调试日志 |

间隔只作用于 Lookup 请求，不限制 CDN 下载。408、429、部分 5xx 等临时 HTTP 错误会退避重试，并处理 `Retry-After`；网络失败也会退避重试。永久性失败会记录在结果中。

## 输出说明

```
output/
├── 产品名_AppStore竞品宣传图与介绍_yyyy-MM-dd.xlsx
├── index.csv
├── index.json
│
├── <App ID>/
│   ├── us/
│   │   ├── metadata.json
│   │   ├── description.md
│   │   ├── manifest.json
│   │   ├── icon/
│   │   └── screenshots/
│   │       ├── iphone/
│   │       └── ipad/
│   │
│   └── cn/
│       └── ...
│
└── <App ID>/
    ├── us/
    └── cn/
```

### 采集文档 Excel 命名规范

Excel 文件名固定为：

```text
产品名_AppStore竞品宣传图与介绍_yyyy-MM-dd.xlsx
```

例如：`收藏助手_AppStore竞品宣传图与介绍_2026-11-02.xlsx`。

- 日期按北京时间（UTC+8）在命令启动时确定，单次运行跨日不改变文件名。`fetched_at` 等记录时间采用 UTC。
- `fetched_at` 表示元数据获取时间，复用缓存时保持不变；清单中的 `last_attempt` 和 `last_success` 单独记录采集运行时间。历史缓存沿用原有 `last_success` 作为近似的元数据获取时间；使用 `--refresh` 可建立新的准确时间。
- 同产品同一天再次运行会更新同名工作簿；不同日期生成不同工作簿，但同目录中的 CSV、JSON 和素材仍会更新。
- 产品名不能为空，不能以句点结尾，也不能包含 `<>:"/\|?*` 或控制字符。
- 运行前关闭同名工作簿。需要使用本地原图链接时，保留原始目录。
- 已有手动命名的 `index.xlsx` 不会被命令行工具自动改名。
- 不同产品使用独立目录；需要保留完整历史批次时，再按日期建立子目录。


导出英文工作簿时，添加 `--excel-language en`：

```powershell
python collector.py --product-name "收藏助手" --input apps.txt --output "output/收藏助手-en" --excel-language en
```

中文使用 `--excel-language zh-CN`，也是默认值。语言设置切换工作表名称、表头、生成的说明、图片占位提示及导航文字；应用名称、介绍、更新说明、原始字段与记录中的状态和错误保留原文。文件命名沿用现有规则。同产品同日期在同一目录中切换语言会覆盖同名工作簿，需要保留两版时使用不同输出目录。

### 工作簿布局

| 工作表 | 内容 |
| --- | --- |
| 研究总览 | 每个应用 / 地区一行，展示比较指标、状态和预览跳转 |
| 宣传图预览 | 彩色标题条、图标和应用信息；优先展示前三张 iPhone 图，无 iPhone 图时使用其他图；右侧提供导航 |
| 全部宣传图 | 每组三张，展示清单中的非图标素材，标记设备与来源 |
| 应用介绍 | 应用介绍与更新说明按每段最多 900 字符拆分 |
| 宣传图索引 | 设备、来源、本地可用状态、文件链接、原图地址、哈希与错误 |
| 来源数据 | 原始汇总字段，支持筛选和自动换行 |
| 采集说明 | 范围、统计口径、限制和导航说明 |

缩略图按比例嵌入，不修改原始文件。缺失或无法读取的图片显示占位说明。本地文件链接依赖原目录；网络链接依赖网络和源站可用性。

## 常见问题

### **提示缺少 `--product-name`，怎么办？**

补上研究名称，例如 `--product-name "收藏助手"`。程序不会根据输入应用自动推断名称。

### **为什么同一应用只在一个地区有结果？**

Lookup 可能没有在另一地区返回该应用。查看商店页面和采集状态，程序不会用其他地区的数据替代。

### 重新运行与断点续传

当采集到一半意外退出时，重新执行原命令即可。

```powershell
py -3 collector.py --product-name "收藏助手" --input apps.txt --output output
```

程序会跳过已成功保存的文件，继续尝试未完成的下载。如果希望重新拉取产品最新信息和宣传素材，可以加上：

```powershell
py -3 collector.py --product-name "收藏助手" --input apps.txt --output output --refresh
```

默认 API 请求间隔为 3.5 秒，约每分钟 17 次调用，低于 Apple 文档所述约每分钟 20 次的参考限制。

### **为什么宣传图少于商店页面上的数量？**

公开 API 可能只返回部分素材。`--web-fallback` 可以尝试补充，但无法保证完整，也不能保证识别准确。

### **生成工作簿时出现 `PermissionError`，怎么办？**

关闭 Excel 或其他查看器中的同名文件后重试。如果文件未打开，检查输出目录的写入权限。已完成的下载可在下次运行时复用。

### **能否完全不联网，只重新导出 Excel？**

目前没有仅导出的命令行选项。普通重跑会复用缓存，但仍可能重试失败项，或查询没有缓存元数据的记录。如果最新 Lookup 结果为不可用，重跑会重新查询，不会将历史元数据直接视为当前数据。历史 JSON 和素材仍保留在磁盘上，报告可能在最新状态旁展示历史内容。

### **只选择一个地区，为什么报告仍有另一个地区？**

汇总会读取输入 ID 对应的美国区和中国区已有清单，`--regions` 只控制本次采集请求。使用独立输出目录可避免旧地区缓存进入报告。

### **总览宣传图数量为什么与本地可用图片不同？**

总览统计清单中标记成功的非图标文件，可能包含历史文件；索引另外检查本地文件是否存在。手动移动或删除素材后，应结合两处信息判断。

### **可以只发送 Excel 文件吗？**

嵌入的预览图可以随工作簿查看；高清本地链接需要原始目录。分享前确认内容使用权限。

### **不同日期生成工作簿，是否等于备份完整数据？**

只会生成新的工作簿文件名。元数据和素材仍使用原目录。要保留独立数据，请使用新的输出目录或复制整个研究目录。

## 已知限制

公开 Lookup API 不保证提供所有设备尺寸、本地化素材、自定义产品页面或视频。公开网页补充采用启发式识别，不能可靠区分所有素材类型。URL 去重仅在同一应用 / 地区内进行。

程序不访问私有接口、不采集凭据、不绕过访问控制，也不访问第三方应用的 App Store Connect。采集时应遵守源站限流要求及素材权利约定。

Excel 单元格最多保存 32,767 个字符，完整原始内容见 JSON 和 Markdown。文件采用临时文件替换写入，但整次采集不是覆盖整个输出目录的事务。每完成一个应用 / 地区后会重建汇总。

## 开发维护

| 文件 | 职责 |
| --- | --- |
| `collector.py` | 网络请求、缓存、命令行与汇总生成 |
| `excel_export.py` | 工作簿布局与图片嵌入 |
| `requirements.txt` | 运行依赖 |
| `tests/` | 离线单元测试 |
| `AGENTS.md` | 项目协作与验证约定 |
| `.gitignore` | 忽略输出、工作簿、缓存、环境和本地配置 |

运行离线测试，无需启动真实采集：

```powershell
python -m unittest discover -s tests -v
```

修改导出功能后，另行打开样例工作簿检查工作表顺序、图片与内部链接。修改参数、依赖、命名或布局时同步维护两版 README。新增忽略规则不会自动取消已跟踪文件的 Git 跟踪状态。
