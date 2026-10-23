---
name: pdf-to-docx-pptx
description: 将 PDF、截图和图片版 PPT 转成逐页对应的可编辑 Word 或 PowerPoint。优先保留原生文字，用离线 RapidOCR 补充图片和轮廓字，保留图形底图，并分开验证文字、分页、可见内容与 OCR 校对状态。
---

# PDF、图片与幻灯片转可编辑 Office 文件

默认目标是文字可编辑、页序对应、位置接近原件。图片有助于保留内容，但整页图片不能算可编辑转换成功。只有明确要求图片版 Word 时才调用 `scripts/pdf2docx_image.py`。

## 环境与入口

首次使用按 [README.md](README.md) 安装本地运行依赖。主要验证环境为 Windows、64 位 CPython 3.12、LibreOffice。使用安装依赖的同一个 Python 解释器；无需全局安装依赖。模型加载前检查 SHA-256，转换不自动下载模型、不需要 API 密钥。

```powershell
python scripts/pdf2docx_editable.py 'input.pdf' 'output.docx'
python scripts/pdf2pptx.py 'input.pdf' 'output.pptx'
```

两个入口支持 PDF、PNG/JPEG 等图片、PPT/PPTX。图片版 PPT 先由本机 LibreOffice 渲染，再重建文字。PPT 使用一个画布尺寸；源 PDF 页尺寸不一致时停止，不静默拉伸。

先按原坐标重建原生文字，再从移除原生文字后的页面识别残余文字，避免重复识别。将图片中的表格文字、轮廓字和截图文字恢复为可编辑文本框，清除底图中的对应字形，其余图形保留。相同残余页面在本次转换内复用 OCR 结果。

默认把左上角、右下角的小字当标识保留为图形。先看源页；如果角落包含正文、页码或需要修改的文字，加 `--ocr-all-regions` 并核对结果。不要把角落文字自动等同于品牌。

输出同名 `.conversion.json`，记录输入/输出哈希、逐页原生和 OCR 文字、识别区域、低置信度候选及静态图形范围。保留记录供验证。

## 验证与校对

```powershell
python scripts/verify.py 'input.pdf' 'output.docx' --render
python scripts/verify.py 'input.pdf' 'output.pptx' --render
```

验证器分别检查逐页原生文字、结构页数、实际渲染分页、渲染文字缺失/越界，以及文字区域外的可见内容。`--json 'checks.json'` 保存检查结果。

机械检查不能证明 OCR 正确。`MECHANICAL_CHECKS_PASS_OCR_REVIEW_REQUIRED` 表示仍需校对。不能用原生文字零缺失代替图片或轮廓字完整性检查。

逐项比较 OCR 文字与源图，重点看无文字层页面、轮廓标题、评分表、数字、否定词、专名、箭头附近和小字。低于 0.90 的候选优先检查；高置信度也不能免除校对。修正后重新渲染，确认修改可见、页序不变、图形未丢失。不确定的内容逐页说明，不声称全部正确。

## 编辑范围与其他模式

OCR 字体替代会改变描边、阴影和细节。表格恢复为文本框加图形底图，并非真实 Word 单元格。图片、标识和箭头保留为位图，并非原生形状。

- `--ocr off`：只重建原生文字，其余保留为图形。
- `--ocr-backend windows`：选用 Windows 离线 OCR，需要相应语言包；识别失败必须报告，不能当作空白页。
- `--visuals none`：仅在明确要求纯文字时丢弃图形。
- `--profile flow --ocr off`：旧原生文字重排模式，不保证图形、轮廓字或分页。

遵从指定输出路径，不覆盖源文件。每次渲染使用输出目录下的独立临时目录和 LibreOffice 配置，只有新 PDF 存在才算成功。不得终止全部 Office 进程；超时只停止本次进程树。详见 [quick-reference.md](references/quick-reference.md)。
