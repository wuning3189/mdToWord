# Markdown 文档导出工具

一个纯前端、本地运行的 Markdown → DOCX/PDF 转换工具。文档内容不会上传到服务器。

## 方案说明

本项目先把 Markdown 解析成统一的语义结构，再分别输出：

- **Word**：使用 `docx` 生成原生 `.docx`。标题、段落、列表、表格、代码块、引用、链接、图片和分页符均转换为可编辑的 Word 元素。
- **PDF**：使用与页面预览相同的 HTML/CSS，通过 `html2pdf.js` 按 A4 版式生成 PDF。
- **安全预览**：原始 HTML 在进入预览前由 DOMPurify 清理。

这种方案无需安装 Pandoc、LibreOffice 或调用本机 Word，适合当前 Windows 环境，也便于部署为静态网页。

## 运行

需要 Node.js 16 或更高版本。当前项目为兼容现有 Node.js 16 环境固定了 Vite 4 和 Marked 9；长期使用建议升级到仍在维护的 Node.js LTS。

```bash
npm install
npm run dev
```

生产构建：

```bash
npm run build
```

构建产物位于 `dist`。

## 使用

1. 点击“导入 MD”选择 Markdown 文件，也可以直接在左侧编辑。
2. 如果文档引用本地图片，点击“添加图片”并选择图片文件。
3. 在右侧检查 A4 预览。
4. 点击“导出 Word”或“导出 PDF”。

本地图片建议使用相对路径，例如：

```markdown
![产品界面](images/product.png)
```

选择资源时，工具会优先按完整相对路径匹配，然后按文件名匹配。浏览器不会自动读取 Markdown 所在目录，必须手动添加本地图片。
## 支持的格式

| Markdown 内容      | Word  |  PDF  | 处理方式                            |
| ------------------ | :---: | :---: | ----------------------------------- |
| 1—6 级标题         |   ✓   |   ✓   | Word 原生标题样式，可用于导航和目录 |
| 粗体、斜体、删除线 |   ✓   |   ✓   | 原生行内格式                        |
| 行内代码、代码块   |   ✓   |   ✓   | 等宽字体、底色、保留换行            |
| 有序/无序/嵌套列表 |   ✓   |   ✓   | Word 原生编号与项目符号             |
| 任务列表           |   ✓   |   ✓   | 转换为 `☑` / `☐`                    |
| GFM 管道表格       |   ✓   |   ✓   | 表头底色、边框、对齐、表头跨页重复  |
| 引用               |   ✓   |   ✓   | 专用缩进和边框样式                  |
| 超链接             |   ✓   |   ✓   | HTTP、HTTPS、邮件链接可点击         |
| 图片               |   ✓   |   ✓   | 等比缩放；Word 中统一转成 PNG       |
| 分隔线             |   ✓   |   ✓   | 原生边框/水平线                     |
| 显式分页           |   ✓   |   ✓   | 使用 `<!-- pagebreak -->`           |

## 已知边界

- Markdown 没有统一定义“所有样式”。本项目以 CommonMark 与 GFM 常用语法为首版范围。
- 原始 HTML 会经过安全清理。复杂 HTML、`style` 脚本化布局、表格合并单元格不会完整转换到 Word。
- Mermaid、数学公式、脚注和自动目录尚未实现。
- PDF 由浏览器画布生成，视觉效果接近预览，但正文在部分 PDF 阅读器中可能不能像原生排版 PDF 那样逐字选择。
- 网络图片必须支持 CORS；否则浏览器无法把图片写入导出文件。
- GIF 写入 Word 时保留首帧；SVG、WebP 等格式写入 Word 前转换为 PNG。
- 超宽表格会压缩到页面宽度。复杂表格建议拆分，或者后续增加横向页面规则。
- Word 中文字体优先使用微软雅黑，代码优先使用 Consolas。目标机器缺少字体时，Word 会自动替换。

## 验收建议

使用 `example.md` 分别导出 DOCX 和 PDF，并检查：

- 标题层级及 Word 导航窗格；
- 三级嵌套列表和任务状态；
- 表头、对齐和跨页表现；
- 中英文代码换行；
- 引用、链接和显式分页；
- 本地图片与网络图片；
- 中文字体和 A4 页边距。

## 主要依赖

- [Marked](https://marked.js.org/)：Markdown 解析
- [docx](https://docx.js.org/)：浏览器端生成 DOCX
- [DOMPurify](https://github.com/cure53/DOMPurify)：预览 HTML 清理
- [html2pdf.js](https://github.com/eKoopmans/html2pdf.js)：浏览器端生成 PDF
