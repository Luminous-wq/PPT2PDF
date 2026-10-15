# PPT2PDF — Windows PowerPoint 转 PDF 与自动裁白边工具

一个面向 **Windows + Microsoft PowerPoint** 的轻量级 PPT/PPTX 导出工具：将演示文稿中的**指定单页、多页或全部页面**导出为 PDF，并自动检测内容边界、去除四周白边。项目还提供可选的 **LLM Agent 命令行交互**，可通过自然语言指定文件、页码与裁剪规则。

## 功能简介

- **指定页导出**：支持第 `28` 页、`3,5-7` 页或 `all`（全部页面）。
- **自动裁白边**：识别页面中的非白色内容并裁剪多余空白，可设置保留边距（单位：pt）。
- **两种裁剪模式**：
  - `per_page`：每页独立裁剪，尽可能减少白边（默认）。
  - `union`：多页采用统一的内容边界，输出页面尺寸一致。
- **尽量保留矢量质量**：临时渲染图像仅用于检测内容范围；最终从原始 PDF 页面复制裁剪区域，而不是将整页截图保存为 PDF。
- **可选的自然语言 Agent**：例如输入“把 `method.pptx` 第 3、5-7 页导出为 PDF，统一尺寸，保留 3pt 白边”，Agent 根据指令调用本地工具。

> **适用场景**：论文框架图、实验结果图、流程图等使用 PowerPoint 绘制的插图导出，尤其适合后续插入 LaTeX 文档。

## 运行环境

- **操作系统**：Windows 10 / 11
- **软件要求**：本机安装 **Microsoft PowerPoint**（使用 PowerPoint COM 接口导出 PDF）
- **Python**：建议 Python 3.10 或更新版本
- **Python 依赖**：`pywin32`、`PyMuPDF`、`numpy`；Agent 模式另需 `openai`、`python-dotenv`

> 本项目依赖 Windows COM 与本机 PowerPoint，不适用于直接在 Linux 或 macOS 上运行当前实现。仅安装 Python 依赖，不能代替安装 PowerPoint。

## 项目结构

```text
PPT2PDF/
├── main.py                 # Agent 命令行入口
├── agent.py                # LLM / Function Calling / Tool 执行循环
├── prompt.txt              # Agent 系统提示词
├── tools/
│   ├── __init__.py
│   └── ppt_tools.py        # PPT 转 PDF、选页、自动裁剪
├── requirements.txt        # Python 依赖
├── .env.example            # 环境变量示例（可自行创建）
└── .gitignore
```

## 安装

### 1. 下载项目并进入目录

```powershell
cd C:\path\to\PPT2PDF
```

### 2. 安装依赖

```powershell
python -m pip install -r requirements.txt
```

如有python基本环境，可直接

```powershell
python -m pip install pywin32 PyMuPDF numpy openai python-dotenv
```

## 使用方式一：直接调用 Python 函数（不需要 API Key）

在 `tools/ppt_tools.py` 中提供 `export_ppt_to_cropped_pdf()` 时，可直接使用：

```python
from tools.ppt_tools import export_ppt_to_cropped_pdf

result = export_ppt_to_cropped_pdf(
    ppt_path=r"C:\path\to\method.pptx",
    pages="3,5-7",
    output_path=r"C:\path\to\method_cropped.pdf",
    margin=3,
    crop_mode="per_page"
)

print(result)
```

参数说明：

| 参数          | 说明                                                      | 示例 / 默认值               |
| ------------- | --------------------------------------------------------- | --------------------------- |
| `ppt_path`    | PPT/PPTX 文件路径（绝对或相对路径）                       | `r"C:\path\to\method.pptx"` |
| `pages`       | 要导出的页码，**从 1 开始**                               | `"28"`、`"3,5-7"`、`"all"`  |
| `output_path` | 输出 PDF 路径；`None` 时在 PPT 同目录生成 `*_cropped.pdf` | `None`                      |
| `margin`      | 裁剪后保留的边距（pt）                                    | `6`                         |
| `crop_mode`   | `per_page`（逐页）或 `union`（统一边界）                  | `"per_page"`                |
| `dpi`         | 用于检测内容范围的临时渲染精度                            | `200`                       |

## 使用方式二：Agent模式（需要模型 API）

### 1. 配置 API

在项目根目录创建 `.env` 文件。
目前项目使用的OpenRouter, 注册就有一个api，每日50次调度，日常研究绘图使用足以。
```env
OPENAI_API_KEY=sk-你的密钥
OPENAI_MODEL=openrouter/free
```
使用 OpenAI API 时：

```env
OPENAI_API_KEY=your-api-key-here
OPENAI_MODEL=your-supported-model
```
后续改为self.client = OpenAI()，为空即可。

### 2. 启动 Agent

```powershell
python main.py
```

输入示例：

```text
请输入 > 把 C:\path\to\method.pptx 第 28 页导出成 PDF，裁掉白边，保留 1pt

请输入 > 把 C:\path\to\method.pptx 第 3、5-7 页导出来，每页尺寸保持一致

请输入 > 把这个 PPT 改成导出第 9 页
```

## 裁剪模式说明

- **`per_page`**：分别计算每页内容的外接矩形；不同页可能得到不同大小的 PDF 页面。适合独立插图。
- **`union`**：在同一 PPT 页面坐标系中，计算所选各页内容边界的并集，再对所有页使用同一裁剪区域。适合连续的多页示意图。此模式假定各页使用相同的页面尺寸与坐标系。

## 注意事项与局限

1. **背景问题**：当前算法依据“非白色像素”定位内容。如果页面有深色背景、铺满整页的图片或底色，可能无法裁掉预期的边缘。
2. **接近白色的图形**：过浅的线条和元素可能被当作白色背景；可调整 `white_threshold` 或 `dpi`。
3. **矢量与内容支持**：最终裁剪会尽量保留原始 PDF 的矢量绘制内容，但不保证保留原文件中的链接、注释等交互元素。
4. **`union` 的前提**：适用于来自相同版式和坐标系的页面，不适合混用不同尺寸的 PDF 页面。
5. **PowerPoint 依赖**：必须能在 Windows 上正常启动 Microsoft PowerPoint；转换失败时检查安装情况与文件权限。
6. **Agent 成本**：直接调用 Python 函数不需要大模型 API；Agent 模式可能产生 API 费用或受到免费额度限制。
