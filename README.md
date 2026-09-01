# PDF Catalog Builder

用于扫描年级目录中的 PDF，提取文件标题，按需将每个 PDF 的前 5 页转换为带水印图片，并输出 CSV/XLSX 表格。生成的目录和文件名默认使用英文/ASCII。

项目方案见 [docs/方案设计.md](docs/方案设计.md)。

## 初始化环境

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e .
```

## 预期命令

```powershell
pdf-catalog scan --config config.example.yaml
pdf-catalog convert --config config.example.yaml
pdf-catalog run --config config.example.yaml
```

前期测试可在配置中将 `processing.max_pdfs` 设为 `20`，或临时执行 `pdf-catalog run --config config.example.yaml --limit 5`；正式全量运行时设置为 `null` 或使用 `--limit none`。

命令已实现，运行后会在 `output/` 生成 `images/`、`pdf_catalog.xlsx`、`pdf_catalog.csv`、`errors.csv` 和 `run.log`。图片路径相对于输出目录，重复运行时会根据源文件属性、渲染参数和水印配置复用缓存。损坏或加密 PDF 会记录到 `errors.csv`，不会中断批次。

也可以直接从源码运行：

```powershell
$env:PYTHONPATH = "src"
python -m pdf_catalog.cli run --config config.example.yaml --limit 5
```
