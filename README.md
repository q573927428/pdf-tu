# PDF Catalog Builder

用于扫描年级目录中的 PDF，提取文件标题，按需将每个 PDF 的前 5 页转换为带水印图片，并输出 CSV/XLSX 表格。生成的目录和文件名默认使用英文/ASCII。

项目方案见 [docs/方案设计.md](docs/方案设计.md)。

## 初始化环境

```powershell
python -m venv .venv

先激活虚拟环境
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e . --no-build-isolation

//或者使用
.\.venv\Scripts\python.exe -m pip install -e . --no-build-isolation
```

## 预期命令

```powershell
# 仅扫描 PDF，生成目录表格，暂不转换图片
pdf-catalog scan --config config.yaml
# 将 PDF 页面转换为图片，并更新目录表格
pdf-catalog convert --config config.yaml
# 执行完整流程：扫描 PDF、转换图片并输出目录表格
pdf-catalog run --config config.yaml
# 仅处理第 11~20 个 PDF（序号按稳定排序，从 1 开始，包含首尾）
pdf-catalog run --config config.yaml --start 11 --end 20
# 仅为第 11~20 个 PDF 生成文案和封面（需先配置豆包 API）
pdf-catalog run --config config.yaml --ai-start 11 --ai-end 20 --generate-copy --generate-cover
# AI 独立阶段：只生成文案，不渲染 PDF 图片
pdf-catalog ai --config config.yaml --start 1 --end 10 --generate-copy
# 启动本地服务，让 HTML 中的“生成文案/生成封面”按钮可以直接工作
pdf-catalog serve --config config.yaml
```

在 Windows 中也可以直接双击项目根目录的 `乐发发资料.bat`（`start_server.bat` 为兼容入口）：脚本会使用 `.venv` 中的 Python 启动本地服务，等待服务就绪后自动打开 `pdf_catalog.html`。如使用自定义端口或 HTML 文件名，请同步修改脚本顶部的 `PAGE` 地址。

表格新增“生成文案”和“封面图链接”两列。豆包接口采用 Ark 的 OpenAI 兼容地址，配置 `ai.enabled/api_key`，并分别设置 `copy_model`（文本模型）和 `image_model`（图片生成模型，如 Seedream）后再显式传入生成参数；两者留空时会兼容使用旧的 `ai.model`。文案和封面图共用同一个基础 `endpoint`，但分别访问 `/chat/completions` 和 `/images/generations`。生成的封面会自动下载到对应 PDF 的 `output/images/.../<序号>/cover.png`；XLSX 中“封面图链接”及“图 1”至“图 5”列均保留图片路径文字、不嵌入图片，HTML 中则显示可点击的图片缩略图且不显示路径，CSV 仍保留相对路径。`--ai-start/--ai-end/--ai-limit` 只约束 AI 生成范围，`--start/--end` 约束本次扫描范围。

项目根目录提供 `static/` 静态资源目录。生成封面时可在 `ai.reference_image` 填写参考图 URL，或填写本地路径（例如 `./static/reference.png`）；程序会将其作为图片生成接口的 `image` 参数，实现图文生图/参考图编辑。留空则按纯文生图处理。

前期测试可在配置中将 `processing.max_pdfs` 设为 `20`，或临时执行 `pdf-catalog run --config config.yaml --limit 5`；正式全量运行时设置为 `null` 或使用 `--limit none`。

命令已实现，运行后会在 `output/` 生成 `images/`、`pdf_catalog.xlsx`、`pdf_catalog.csv`、`errors.csv` 和 `run.log`。图片路径相对于输出目录，重复运行时会根据源文件属性、渲染参数和水印配置复用缓存。损坏或加密 PDF 会记录到 `errors.csv`，不会中断批次。

同时会生成 `pdf_catalog.html` 单页表格，可直接用浏览器打开；支持关键词筛选、图片预览、链接跳转和打印（打印时自动收缩表格）。当“封面图链接”或“生成文案”为空时，HTML 对应单元格会显示按钮。先运行 `pdf-catalog serve --config config.yaml`，再打开 HTML，点击按钮即可为当前 PDF 调用 AI 并更新目录文件；服务默认监听 `127.0.0.1:8765`，可用 `--host`、`--port` 修改。如需修改 HTML 文件名，可在 `table.html` 配置项中覆盖默认值。

也可以直接从源码运行：

```powershell
$env:PYTHONPATH = "src"
python -m pdf_catalog.cli run --config config.yaml --limit 5
```

## 切换年级

一年级、二年级、三年级分别复制一份配置文件，修改 `source_root`、`output_root` 和 `grade`。例如二年级：

```yaml
source_root: "F:\\ipkaishi\\zhixiaoman\\二年级.zip\\二年级"
output_root: "./output-二年级"
grade: "二年级"
semester: "全册"
```

当目录包含 `上册/下册` 时，程序会自动识别学期，并将其后的目录写入“科目”和“分类”列。幼升小没有册别目录时使用默认学期 `全册`。不同年级建议使用独立的 `output_root`。
