# Markdown 导出格式示例

这是一份用于检查 **Word 与 PDF 导出效果** 的示例文档。正文支持 *斜体*、~~删除线~~、`行内代码` 与 [可点击链接](https://example.com)。

> 引用内容会使用独立的缩进、底色和边框样式。
>
> 引用内也可以包含多段文字。

## 标题与列表

### 无序列表

- 一级项目
  - 二级项目
    - 三级项目
- 支持较长的列表内容，导出后仍会保持清晰的缩进关系。

### 有序列表与任务

1. 导入 Markdown 文档
2. 检查右侧 A4 预览
3. 导出 Word 或 PDF

- [x] 标题、段落与行内格式
- [x] 表格、引用与代码块
- [ ] 添加自己的图片资源

## 表格

| 功能     | Word  |  PDF  |                 说明 |
| :------- | :---: | :---: | -------------------: |
| 标题层级 |   ✓   |   ✓   |               1—6 级 |
| 表格     |   ✓   |   ✓   |         GFM 管道表格 |
| 代码块   |   ✓   |   ✓   |       保留空格与换行 |
| 图片     |   ✓   |   ✓   | 支持本地资源及网络图 |

## 代码块

```javascript
function markdownToDocument(source) {
  const document = parse(source);
  return exportWithStyles(document);
}

console.log(markdownToDocument("# 你好，Markdown"));
```
## 图片

使用工具栏中的“添加图片”选择 Markdown 引用的本地图片。建议把 Markdown 与图片放在同一目录，并使用相对路径：

```markdown
![图片说明](example.png)
```

网络图片需要图片服务器允许跨域访问，才能写入 Word 或 PDF。

<!-- pagebreak -->

## 新的一页

在 Markdown 中写入以下注释，可以让 Word 与 PDF 都从新的一页开始：

```markdown
<!-- pagebreak -->
```

### 支持范围

- CommonMark 基础语法
- GFM 表格、删除线和任务列表
- 六级标题、引用、分隔线及围栏代码块
- PNG、JPEG、GIF、WebP、SVG 图片会统一转换成 Word 兼容的 PNG
- 原始 HTML 仅提取安全文本，不保证复杂 HTML 布局的导出效果

---

文档结束。