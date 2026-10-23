# 选项与故障排查

| 选项 | 实际结果 |
|---|---|
| 默认 `--ocr auto --visuals preserve` | 原生文字 + 残余图像/轮廓字 OCR 文本框，保留图形底图 |
| `--ocr off` | 只有原生文字可编辑 |
| `--ocr-backend windows` | Windows PowerShell 5.1 和对应语言包；不调用云 OCR |
| `--visuals none` | 明确丢弃图形的文字版 |
| `verify.py --render` | 实际转回 PDF，检查分页、文字和图形保留 |

## 记录与校对

`文件.docx.conversion.json` 或 `文件.pptx.conversion.json` 与文件哈希绑定：

- `pages[].native_text` 不包含图片中的文字。
- `ocr_text`、`ocr_regions` 是 OCR 恢复的文字与位置。
- `review` 是低置信度候选；置信度不是正确率。
- `static_peripheral_regions` 是作为图形保留的边缘标志/页脚，不算可编辑 OCR 文字。
- `background` 表示是否保留底图。

图片表格默认恢复为文本框加底图。需要真实单元格、公式结构或原生箭头时，另做结构重建。

OCR 修订在原生文本框中进行。DOCX 文字在 `w:txbxContent/w:t`，PPTX 在原生文本框。修订后更新输出哈希并记录修订说明；重新渲染验证。不能只改验证结论或放宽阈值来消除失败。

## 常见问题

| 现象 | 处理 |
|---|---|
| 依赖导入失败 | 使用指定 Python，核对 vendor/ocr 的 Python 与 Windows 架构 |
| 模型缺失/哈希不符 | 停止，依据 vendor/ocr_models/models.json 核对官方来源，不自动降级图片版 |
| Windows OCR WinRT 失败 | 报告实际错误；在已授权环境重试或选 RapidOCR |
| OCR 把箭头当短横线 | 对照源图修正，不直接当正文写入 |
| 文字裁剪/重叠 | 检查字号、行框、锚点和可用宽度，改后重新渲染 |
| 旧 PDF 冒充成功 | 新输出目录、独立 LibreOffice 配置；不使用旧文件作成功证据 |
| OCR 通顺但与图片不符 | 以原图为准，不根据语义猜测补写 |

保留原生文字优于全页 OCR。图片/轮廓字才进入补充识别流程；反复识别原生文字会丢失本来可靠的拼写和颜色。
