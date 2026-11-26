(function(){
"use strict";
const sampleMarkdown = "# Markdown 导出格式示例\r\n\r\n这是一份用于检查 **Word 与 PDF 导出效果** 的示例文档。正文支持 *斜体*、~~删除线~~、`行内代码` 与 [可点击链接](https://example.com)。\r\n\r\n> 引用内容会使用独立的缩进、底色和边框样式。\r\n>\r\n> 引用内也可以包含多段文字。\r\n\r\n## 标题与列表\r\n\r\n### 无序列表\r\n\r\n- 一级项目\r\n  - 二级项目\r\n    - 三级项目\r\n- 支持较长的列表内容，导出后仍会保持清晰的缩进关系。\r\n\r\n### 有序列表与任务\r\n\r\n1. 导入 Markdown 文档\r\n2. 检查右侧 A4 预览\r\n3. 导出 Word 或 PDF\r\n\r\n- [x] 标题、段落与行内格式\r\n- [x] 表格、引用与代码块\r\n- [ ] 添加自己的图片资源\r\n\r\n## 表格\r\n\r\n| 功能     | Word  |  PDF  |                 说明 |\r\n| :------- | :---: | :---: | -------------------: |\r\n| 标题层级 |   ✓   |   ✓   |               1—6 级 |\r\n| 表格     |   ✓   |   ✓   |         GFM 管道表格 |\r\n| 代码块   |   ✓   |   ✓   |       保留空格与换行 |\r\n| 图片     |   ✓   |   ✓   | 支持本地资源及网络图 |\r\n\r\n## 代码块\r\n\r\n```javascript\r\nfunction markdownToDocument(source) {\r\n  const document = parse(source);\r\n  return exportWithStyles(document);\r\n}\r\n\r\nconsole.log(markdownToDocument(\"# 你好，Markdown\"));\r\n```\r\n## 图片\r\n\r\n使用工具栏中的“添加图片”选择 Markdown 引用的本地图片。建议把 Markdown 与图片放在同一目录，并使用相对路径：\r\n\r\n```markdown\r\n![图片说明](example.png)\r\n```\r\n\r\n网络图片需要图片服务器允许跨域访问，才能写入 Word 或 PDF。\r\n\r\n<!-- pagebreak -->\r\n\r\n## 新的一页\r\n\r\n在 Markdown 中写入以下注释，可以让 Word 与 PDF 都从新的一页开始：\r\n\r\n```markdown\r\n<!-- pagebreak -->\r\n```\r\n\r\n### 支持范围\r\n\r\n- CommonMark 基础语法\r\n- GFM 表格、删除线和任务列表\r\n- 六级标题、引用、分隔线及围栏代码块\r\n- PNG、JPEG、GIF、WebP、SVG 图片会统一转换成 Word 兼容的 PNG\r\n- 原始 HTML 仅提取安全文本，不保证复杂 HTML 布局的导出效果\r\n\r\n---\r\n\r\n文档结束。";
const { AlignmentType, BorderStyle, Document, ExternalHyperlink, Footer, HeadingLevel, ImageRun, LevelFormat, Packer, PageBreak, PageNumber, Paragraph, ShadingType, Table, TableCell, TableRow, TextRun, UnderlineType, VerticalAlign, WidthType } = window.docx;
const html2pdf = window.html2pdf;
const marked = window.marked;
const DOMPurify = window.DOMPurify;
const M = (function(){

const PAGE_BREAK = "<div class=\"page-break\" aria-label=\"分页符\"></div>";

marked.setOptions({ gfm: true, breaks: false });

function normalizeMarkdown(markdown) {
  return markdown.replace(/<!--\s*pagebreak\s*-->/gi, PAGE_BREAK);
}

function parseMarkdown(markdown) {
  return marked.lexer(normalizeMarkdown(markdown));
}

function renderMarkdown(markdown) {
  const unsafeHtml = marked.parse(normalizeMarkdown(markdown));
  const html = DOMPurify.sanitize(unsafeHtml, {
    ADD_ATTR: ["target"],
    USE_PROFILES: { html: true },
  });
  const template = document.createElement("template");
  template.innerHTML = html;
  template.content.querySelectorAll("a").forEach((link) => {
    link.target = "_blank";
    link.rel = "noopener noreferrer";
  });
  template.content.querySelectorAll("img").forEach((image) => {
    image.loading = "lazy";
    image.alt ||= "Markdown 图片";
  });
  return template.innerHTML;
}

function plainText(tokens = []) {
  return tokens.map((token) => {
    if (token.type === "br") return "\n";
    if (token.type === "image") return token.text || "图片";
    if (token.tokens) return plainText(token.tokens);
    return token.text || token.raw || "";
  }).join("");
}
return { parseMarkdown, renderMarkdown, plainText };
})();
const { parseMarkdown, renderMarkdown, plainText } = M;
const D = (function(){

const FONT_CN = "Microsoft YaHei";
const FONT_LATIN = "Aptos";
const FONT_CODE = "Consolas";
const PAGE_WIDTH_PX = 600;
const headingLevels = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4,
  5: HeadingLevel.HEADING_5,
  6: HeadingLevel.HEADING_6,
};

function runFont(name = FONT_LATIN) {
  return { ascii: name, hAnsi: name, eastAsia: FONT_CN };
}

function safeLink(url = "") {
  return /^(https?:|mailto:)/i.test(url) ? url : null;
}
function normalizeAssetPath(value = "") {
  return decodeURIComponent(value.split(/[?#]/)[0])
    .replace(/\\/g, "/")
    .replace(/^\.\//, "");
}

function findAsset(source, assets) {
  const path = normalizeAssetPath(source);
  const baseName = path.split("/").pop();
  return assets.get(path) || assets.get(baseName);
}

async function sourceToBlob(source, assets) {
  const localFile = findAsset(source, assets);
  if (localFile) return localFile;
  if (!/^(https?:|data:|blob:)/i.test(source)) {
    throw new Error(`未找到本地图片：${source}`);
  }
  const response = await fetch(source);
  if (!response.ok) throw new Error(`图片加载失败：${source}`);
  return response.blob();
}

function loadBitmap(blob) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const image = new Image();
    image.onload = () => {
      URL.revokeObjectURL(url);
      resolve(image);
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("浏览器无法解析图片"));
    };
    image.src = url;
  });
}

async function createImageRun(source, alt, assets) {
  const image = await loadBitmap(await sourceToBlob(source, assets));
  const dataScale = Math.min(1, 1800 / Math.max(image.naturalWidth, image.naturalHeight));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(image.naturalWidth * dataScale));
  canvas.height = Math.max(1, Math.round(image.naturalHeight * dataScale));
  canvas.getContext("2d").drawImage(image, 0, 0, canvas.width, canvas.height);
  const pngBlob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
  if (!pngBlob) throw new Error("图片转换失败");
  const displayScale = Math.min(1, PAGE_WIDTH_PX / canvas.width, 760 / canvas.height);
  return new ImageRun({
    type: "png",
    data: await pngBlob.arrayBuffer(),
    transformation: {
      width: Math.round(canvas.width * displayScale),
      height: Math.round(canvas.height * displayScale),
    },
    altText: { title: alt || "图片", description: alt || "Markdown 图片", name: alt || "image" },
  });
}

function textOptions(state, text) {
  return {
    text,
    bold: state.bold,
    italics: state.italics,
    strike: state.strike,
    font: runFont(state.code ? FONT_CODE : FONT_LATIN),
    shading: state.code ? { type: ShadingType.CLEAR, fill: "EEF1F4" } : undefined,
    size: state.code ? 20 : 22,
  };
}
async function inlineRuns(tokens = [], assets, state = {}) {
  const children = [];
  for (const token of tokens) {
    if (token.type === "strong" || token.type === "em" || token.type === "del") {
      const next = {
        ...state,
        bold: state.bold || token.type === "strong",
        italics: state.italics || token.type === "em",
        strike: state.strike || token.type === "del",
      };
      children.push(...await inlineRuns(token.tokens, assets, next));
    } else if (token.type === "codespan") {
      children.push(new TextRun(textOptions({ ...state, code: true }, token.text)));
    } else if (token.type === "br") {
      children.push(new TextRun({ break: 1 }));
    } else if (token.type === "link") {
      const text = plainText(token.tokens) || token.href;
      const link = safeLink(token.href);
      if (link) {
        children.push(new ExternalHyperlink({
          link,
          children: [new TextRun({
            ...textOptions(state, text),
            color: "0969DA",
            underline: { type: UnderlineType.SINGLE },
          })],
        }));
      } else {
        children.push(new TextRun(textOptions(state, text)));
      }
    } else if (token.type === "image") {
      try {
        children.push(await createImageRun(token.href, token.text, assets));
      } catch (error) {
        children.push(new TextRun({
          ...textOptions({ ...state, italics: true }, `[图片未加载：${token.text || token.href}]`),
          color: "9A3412",
        }));
      }
    } else if (token.tokens) {
      children.push(...await inlineRuns(token.tokens, assets, state));
    } else if (token.type === "html") {
      const container = document.createElement("span");
      container.innerHTML = token.text || token.raw || "";
      if (container.textContent) children.push(new TextRun(textOptions(state, container.textContent)));
    } else {
      const text = token.text ?? token.raw ?? "";
      if (text) children.push(new TextRun(textOptions(state, text)));
    }
  }
  return children.length ? children : [new TextRun("")];
}

async function paragraphFromToken(token, context, options = {}) {
  const tokens = token.tokens || [{ type: "text", text: token.text || "" }];
  return new Paragraph({
    children: await inlineRuns(tokens, context.assets),
    style: options.style,
    heading: options.heading,
    numbering: options.numbering,
    bullet: options.bullet,
    alignment: options.alignment,
    indent: options.indent,
    keepNext: options.keepNext,
  });
}
function createOrderedNumbering(context, start = 1) {
  const reference = `ordered-${++context.numberingIndex}`;
  context.numbering.push({
    reference,
    levels: [{
      level: 0,
      format: LevelFormat.DECIMAL,
      text: "%1.",
      alignment: AlignmentType.LEFT,
      start: Number(start) || 1,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } },
    }],
  });
  return reference;
}

async function convertList(list, context) {
  const result = [];
  const reference = list.ordered ? createOrderedNumbering(context, list.start) : null;
  for (const item of list.items) {
    const normalTokens = item.tokens.filter((token) => token.type !== "list");
    const nestedLists = item.tokens.filter((token) => token.type === "list");
    const first = normalTokens.shift() || { type: "text", text: "", tokens: [] };
    const prefix = item.task ? `${item.checked ? "☑" : "☐"} ` : "";
    const firstTokens = first.tokens || [{ type: "text", text: first.text || "" }];
    if (prefix) firstTokens.unshift({ type: "text", text: prefix });
    result.push(await paragraphFromToken({ ...first, tokens: firstTokens }, context, {
      style: context.quote ? "Quote" : undefined,
      numbering: list.ordered ? { reference, level: 0 } : undefined,
      bullet: list.ordered ? undefined : { level: Math.min(context.listLevel, 8) },
    }));
    for (const token of normalTokens) {
      result.push(...await convertBlocks([token], {
        ...context,
        continuationIndent: 720 + context.listLevel * 360,
      }));
    }
    for (const nested of nestedLists) {
      result.push(...await convertList(nested, { ...context, listLevel: context.listLevel + 1 }));
    }
  }
  return result;
}

function tableAlignment(value) {
  if (value === "center") return AlignmentType.CENTER;
  if (value === "right") return AlignmentType.RIGHT;
  return AlignmentType.LEFT;
}

async function createTableCell(cell, column, token, context, header = false) {
  const tokens = cell.tokens || [{ type: "text", text: cell.text || "" }];
  return new TableCell({
    children: [new Paragraph({
      children: await inlineRuns(tokens, context.assets),
      alignment: tableAlignment(token.align?.[column]),
    })],
    verticalAlign: VerticalAlign.CENTER,
    shading: header ? { type: ShadingType.CLEAR, fill: "E9EEF5" } : undefined,
    margins: { top: 90, bottom: 90, left: 110, right: 110 },
  });
}

async function convertTable(token, context) {
  const headerCells = await Promise.all(token.header.map((cell, index) =>
    createTableCell(cell, index, token, context, true)));
  const rows = [new TableRow({ children: headerCells, tableHeader: true, cantSplit: true })];
  for (const row of token.rows) {
    rows.push(new TableRow({
      children: await Promise.all(row.map((cell, index) =>
        createTableCell(cell, index, token, context))),
      cantSplit: true,
    }));
  }
  const border = { style: BorderStyle.SINGLE, size: 4, color: "C8D1DC" };
  return new Table({
    rows,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border },
  });
}
async function convertBlocks(tokens, context) {
  const output = [];
  for (const token of tokens) {
    const common = {
      style: context.quote ? "Quote" : undefined,
      indent: context.continuationIndent ? { left: context.continuationIndent } : undefined,
    };
    if (token.type === "space" || token.type === "def") continue;
    if (token.type === "heading") {
      output.push(await paragraphFromToken(token, context, {
        heading: headingLevels[token.depth],
        keepNext: true,
      }));
    } else if (token.type === "paragraph" || token.type === "text") {
      output.push(await paragraphFromToken(token, context, common));
    } else if (token.type === "blockquote") {
      output.push(...await convertBlocks(token.tokens, { ...context, quote: true }));
    } else if (token.type === "list") {
      output.push(...await convertList(token, context));
    } else if (token.type === "table") {
      output.push(await convertTable(token, context));
      output.push(new Paragraph({ children: [] }));
    } else if (token.type === "code") {
      if (token.lang) {
        output.push(new Paragraph({
          style: "CodeLabel",
          children: [new TextRun({ text: token.lang, font: runFont(FONT_CODE), size: 18 })],
          keepNext: true,
        }));
      }
      const lines = token.text.split("\n");
      output.push(new Paragraph({
        style: "CodeBlock",
        children: lines.flatMap((line, index) => [
          new TextRun({
            text: line || " ",
            break: index ? 1 : undefined,
            font: runFont(FONT_CODE),
            size: 19,
          }),
        ]),
      }));
    } else if (token.type === "hr") {
      output.push(new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 5, color: "CBD5E1" } },
        spacing: { before: 120, after: 180 },
      }));
    } else if (token.type === "html" && /class=["']page-break/i.test(token.raw || token.text)) {
      output.push(new Paragraph({ children: [new PageBreak()] }));
    } else if (token.type === "html") {
      const container = document.createElement("div");
      container.innerHTML = token.text || token.raw || "";
      const text = container.textContent.trim();
      if (text) output.push(new Paragraph({ children: [new TextRun(text)] }));
    }
  }
  return output;
}

function headingStyle(id, name, size, color, before, after, outlineLevel) {
  return {
    id,
    name,
    basedOn: "Normal",
    next: "Normal",
    quickFormat: true,
    run: { bold: true, size, color, font: runFont(FONT_LATIN) },
    paragraph: {
      spacing: { before, after },
      keepNext: true,
      keepLines: true,
      outlineLevel,
    },
  };
}

function documentStyles() {
  return {
    default: {
      document: {
        run: { font: runFont(FONT_LATIN), size: 22, color: "24292F" },
        paragraph: { spacing: { line: 360, after: 120 } },
      },
    },
    paragraphStyles: [
      headingStyle("Heading1", "Heading 1", 36, "172033", 320, 180, 0),
      headingStyle("Heading2", "Heading 2", 30, "172033", 280, 140, 1),
      headingStyle("Heading3", "Heading 3", 26, "25324A", 240, 120, 2),
      headingStyle("Heading4", "Heading 4", 23, "25324A", 200, 100, 3),
      headingStyle("Heading5", "Heading 5", 22, "334155", 180, 80, 4),
      headingStyle("Heading6", "Heading 6", 21, "475569", 160, 80, 5),
      {
        id: "Quote", name: "Quote", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { color: "526076", italics: true, font: runFont(FONT_LATIN) },
        paragraph: {
          indent: { left: 440, right: 120 },
          spacing: { before: 80, after: 120 },
          border: { left: { style: BorderStyle.SINGLE, size: 14, color: "5B7CFA", space: 12 } },
        },
      },
      {
        id: "CodeBlock", name: "Code Block", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: runFont(FONT_CODE), size: 19, color: "182230" },
        paragraph: {
          spacing: { before: 0, after: 160, line: 280 },
          indent: { left: 180, right: 180 },
          shading: { type: ShadingType.CLEAR, fill: "F4F6F8" },
          border: { left: { style: BorderStyle.SINGLE, size: 8, color: "B8C4D4", space: 8 } },
        },
      },
      {
        id: "CodeLabel", name: "Code Label", basedOn: "Normal", next: "CodeBlock",
        run: { font: runFont(FONT_CODE), size: 18, color: "FFFFFF", bold: true },
        paragraph: { spacing: { before: 120, after: 0 }, shading: { type: ShadingType.CLEAR, fill: "344258" }, indent: { left: 120 } },
      },
    ],
  };
}
function cleanFileName(name) {
  return (name || "Markdown文档").replace(/\.(md|markdown)$/i, "").replace(/[<>:"/\\|?*]/g, "_");
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function createWordDocument(markdown, options = {}) {
  const context = {
    assets: options.assets || new Map(),
    numbering: [],
    numberingIndex: 0,
    listLevel: 0,
    quote: false,
  };
  const tokens = parseMarkdown(markdown);
  const children = await convertBlocks(tokens, context);
  if (!children.length) children.push(new Paragraph({ children: [new TextRun("")] }));

  const footer = new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: "第 ", color: "718096", size: 18 }),
        new TextRun({ children: [PageNumber.CURRENT], color: "718096", size: 18 }),
        new TextRun({ text: " 页", color: "718096", size: 18 }),
      ],
    })],
  });
  const title = cleanFileName(options.fileName);
  return {
    title,
    document: new Document({
      creator: "Markdown 文档导出工具",
      title,
      description: "由 Markdown 自动转换生成",
      styles: documentStyles(),
      numbering: { config: context.numbering },
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1134, right: 1276, bottom: 1134, left: 1276, footer: 600 },
          },
        },
        footers: { default: footer },
        children,
      }],
    }),
  };
}

async function exportWord(markdown, options = {}) {
  const result = await createWordDocument(markdown, options);
  downloadBlob(await Packer.toBlob(result.document), `${result.title}.docx`);
}


return { exportWord, cleanFileName };
})();
const { exportWord, cleanFileName } = D;
const P = (function(){

async function waitForImages(root) {
  const images = [...root.querySelectorAll("img")];
  await Promise.all(images.map((image) => {
    if (image.complete) return Promise.resolve();
    return new Promise((resolve) => {
      image.addEventListener("load", resolve, { once: true });
      image.addEventListener("error", resolve, { once: true });
    });
  }));
}

async function exportPdf(preview, fileName) {
  const host = document.createElement("div");
  host.className = "pdf-export-host";
  const content = preview.cloneNode(true);
  content.classList.add("pdf-document");
  host.append(content);
  document.body.append(host);

  try {
    await document.fonts?.ready;
    await waitForImages(content);
    await html2pdf().set({
      margin: [12, 12, 14, 12],
      filename: `${cleanFileName(fileName)}.pdf`,
      image: { type: "jpeg", quality: 0.97 },
      html2canvas: {
        scale: 2,
        useCORS: true,
        allowTaint: false,
        backgroundColor: "#ffffff",
        logging: false,
      },
      jsPDF: { unit: "mm", format: "a4", orientation: "portrait" },
      pagebreak: {
        mode: ["css", "legacy"],
        before: ".page-break",
        avoid: ["pre", "blockquote", "table", "img", "h1", "h2", "h3"],
      },
    }).from(content).save();
  } finally {
    host.remove();
  }
}
return { exportPdf };
})();
const { exportPdf } = P;
(function(){

const editor = document.querySelector("#editor");
const preview = document.querySelector("#preview");
const status = document.querySelector("#status");
const documentName = document.querySelector("#documentName");
const mdFile = document.querySelector("#mdFile");
const assetFiles = document.querySelector("#assetFiles");
const wordButton = document.querySelector("#exportDocx");
const pdfButton = document.querySelector("#exportPdf");

const assets = new Map();
const objectUrls = new Set();
let fileName = "Markdown格式示例.md";
let renderTimer;

function normalizePath(value = "") {
  return decodeURIComponent(value.split(/[?#]/)[0])
    .replace(/\\/g, "/")
    .replace(/^\.\//, "");
}

function setStatus(message, error = false) {
  status.textContent = message;
  status.classList.toggle("error", error);
}

function addAssets(files) {
  for (const file of files) {
    const relative = normalizePath(file.webkitRelativePath || file.name);
    assets.set(relative, file);
    assets.set(file.name, file);
  }
}
function findAsset(source) {
  const path = normalizePath(source);
  return assets.get(path) || assets.get(path.split("/").pop());
}

function releaseObjectUrls() {
  objectUrls.forEach((url) => URL.revokeObjectURL(url));
  objectUrls.clear();
}

function hydratePreviewImages() {
  releaseObjectUrls();
  preview.querySelectorAll("img").forEach((image) => {
    const file = findAsset(image.getAttribute("src"));
    if (!file) return;
    const url = URL.createObjectURL(file);
    objectUrls.add(url);
    image.src = url;
  });
}

function updatePreview() {
  try {
    preview.innerHTML = renderMarkdown(editor.value);
    hydratePreviewImages();
    const characters = editor.value.length;
    const words = editor.value.trim() ? editor.value.trim().split(/\s+/).length : 0;
    const imageCount = new Set(assets.values()).size;
    setStatus(`${characters.toLocaleString()} 字符 · ${words.toLocaleString()} 词 · ${imageCount} 个图片资源`);
  } catch (error) {
    setStatus(`预览失败：${error.message}`, true);
  }
}

function schedulePreview() {
  window.clearTimeout(renderTimer);
  renderTimer = window.setTimeout(updatePreview, 120);
}

async function runExport(button, label, task) {
  const original = button.textContent;
  button.disabled = true;
  button.textContent = label;
  setStatus(label);
  try {
    await task();
    setStatus("导出已完成，文件已保存到浏览器下载目录");
  } catch (error) {
    console.error(error);
    setStatus(`导出失败：${error.message}`, true);
  } finally {
    button.disabled = false;
    button.textContent = original;
  }
}

editor.addEventListener("input", schedulePreview);
mdFile.addEventListener("change", async () => {
  const file = mdFile.files[0];
  if (!file) return;
  fileName = file.name;
  documentName.textContent = fileName;
  editor.value = await file.text();
  updatePreview();
  mdFile.value = "";
});
assetFiles.addEventListener("change", () => {
  addAssets(assetFiles.files);
  updatePreview();
  assetFiles.value = "";
});

wordButton.addEventListener("click", () => runExport(wordButton, "正在生成 Word…", () =>
  exportWord(editor.value, { fileName, assets })));

pdfButton.addEventListener("click", () => runExport(pdfButton, "正在生成 PDF…", () =>
  exportPdf(preview, fileName)));

window.addEventListener("beforeunload", releaseObjectUrls);

editor.value = sampleMarkdown;
documentName.textContent = fileName;
updatePreview();
})();
})();
