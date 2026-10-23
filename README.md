# PDF → 可编辑 Word / PowerPoint

将 PDF、截图和图片版 PPT 转为可编辑 DOCX/PPTX，优先保留文字、页序和位置。适合课件、讲义以及原生文字和截图混合的 PDF。

- 原生 PDF 文字直接重建，保留文字内容、颜色和坐标。
- RapidOCR + ONNX Runtime 在本机补充识别图片和轮廓字。
- 保留照片、表格线、背景和图示；清除已重建文字在底图中的字形，避免重复显示。
- 默认一页对应一个 Word 页面或一张幻灯片。
- 分别检查分页、文字保留、文字越界和图形保留，并明确 OCR 仍需校对。

这是一套本地转换工具，也可作为 AI agent skill 使用，入口为 [SKILL.md](SKILL.md)。

## 安装

主要验证环境：Windows、64 位 **Python 3.12**。PPT 输入和渲染检查需要安装 [LibreOffice](https://www.libreoffice.org/download/download-libreoffice/)。其他系统与 Python 版本尚未验证。

```powershell
git clone https://github.com/elkanhaalexander-dot/pdf-to-docx-pptx.git
cd pdf-to-docx-pptx
python scripts/install_runtime.py
```

安装器将运行依赖放在仓库的 `vendor/ocr/`，模型放在 `vendor/ocr_models/`。模型并行下载，校对官方大小和 SHA-256。安装临时文件和缓存位于仓库内部，完成后清理；把仓库放在 E 盘即可让这些文件留在 E 盘。安装需要联网。

安装后转换使用本地模型，不需要云 OCR 或 API 密钥。仓库只存放源码和依赖/模型清单，不提交运行库、模型、个人文档或转换结果。

```powershell
# 只检查已安装的依赖和模型，不下载
python scripts/install_runtime.py --check-only
```

## 转换

```powershell
python scripts/pdf2docx_editable.py 'input.pdf' 'output.docx'
python scripts/pdf2pptx.py 'input.pdf' 'output.pptx'

# 截图、图片和图片版 PPT 使用相同入口
python scripts/pdf2pptx.py 'screenshot.png' 'editable.pptx'
python scripts/pdf2pptx.py 'image-slides.pptx' 'editable-slides.pptx'
```

输出旁生成 `.conversion.json`，记录源/输出哈希、逐页文字、OCR 区域和待校对候选。修改生成文档后，应更新记录并重新验证。

## 检查结果

```powershell
python scripts/verify.py 'input.pdf' 'output.docx' --render
python scripts/verify.py 'input.pdf' 'output.pptx' --render --json 'checks.json'
```

不加 `--render` 只检查文件结构和文字，不能确认实际分页或可见效果。实际渲染通过也不等于 OCR 识别完全正确。请对照原图校对，尤其检查数字、专名、小字和图形附近的短字符。

已在一个 30 页混合图文 PDF 上检查页数、原生文字保留和 LibreOffice 渲染，并测试截图及两页图片版 PPT。该测试范围不代表任意 PDF 都能无误转换。

## 常用选项

| 选项 | 用途 |
|---|---|
| 默认设置 | 原生文字 + 残余内容 OCR，保留图形底图 |
| `--ocr-all-regions` | 也识别角落的小字、标识和页脚；角落含正文或需编辑页码时使用 |
| `--ocr off` | 只重建原生文字，图片中的文字仍为图像 |
| `--ocr-backend windows` | 改用 Windows 自带 OCR，需要中文/英文语言包 |
| `--visuals none` | 明确要求纯文字时丢弃图形 |

LibreOffice 不在默认位置时，可设置 `PDF_CONVERT_SOFFICE` 为 `soffice.com` 或 `soffice` 的路径。

## 实际限制

文字以原生文本框形式编辑。图片表格暂不重建为真实 Word 单元格；照片、背景、标识、表格线和箭头仍为位图。替代字体会改变描边、阴影和字形细节。默认角落小标识保留为图形，需要编辑时使用 `--ocr-all-regions`。

PPT 的整套幻灯片只有一个画布尺寸，混合尺寸 PDF 会明确报错。Word 使用固定位置文本框，适合保留页面；大幅增删文字需要手动调整文本框。

第三方依赖和模型由各自项目提供，其许可证以对应项目为准：[PyMuPDF](https://github.com/pymupdf/PyMuPDF)、[RapidOCR](https://github.com/RapidAI/RapidOCR)、[ONNX Runtime](https://github.com/microsoft/onnxruntime)、[python-docx](https://github.com/python-openxml/python-docx)、[python-pptx](https://github.com/scanny/python-pptx)。固定版本和模型来源见 `vendor/ocr-runtime.json` 与 `vendor/ocr_models/models.json`。
