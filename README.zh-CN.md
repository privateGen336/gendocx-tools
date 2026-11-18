# GenDocX Tools

**六个纯浏览器文档工具：文件不上传、不落盘、不需要账号。**

图片拼成 Word、Markdown 转成有版式的 `.docx`、PDF 合并签名、给 PDF/Word 加水印、对比两版合同、直接编辑 `.docx` —— 打开 Network 面板操作，它一直是安静的。

[English](./README.md) · [在线版](https://gendocx.com/zh/tools)

![浏览器窗口中的六种文档工具](./docs/images/browser-document-tools.jpg)

---

## 为什么做这个

大多数「免费在线转换」网站会把文件传到你看不见的服务器上。一份签好的合同、一张证件扫描件、一份内部备忘，传过去之后发生了什么，你没有任何办法验证。

这套工具反过来做：解码、处理、重新打包，全在当前标签页里。

- 图片用 `createImageBitmap` 解码，由 `docx` 库打包成 OOXML。
- PDF 交给 Web Worker 里的 `pdf.js` 解析，用 `pdf-lib` 修改。
- DOCX 对比跑在 `docxodus`（.NET 引擎编译成 WebAssembly）的 worker 里。
- 唯一的网络请求，是加载页面本身的静态文件。

代码是 MIT 许可。可以读、可以审计、可以 fork，也可以直接放进自己的产品。

## 包含哪些工具

### 图片转 Word

- 支持 JPG、PNG、WebP、GIF、AVIF、BMP，一次最多 50 张
- 拖拽排序、旋转、删除，改设置时分页位置实时预览
- 纸张（A4 / Letter）、方向（自动或锁定）、每页张数（1 / 2 / 4）
- 边距（无 / 1 cm / 2 cm / 上下分别设置）、质量（适应页面或保留原像素）、JPEG 或 PNG 编码
- 输出是真正的 `.docx`，Word、Pages、WPS、Google Docs、LibreOffice 都能打开

### PDF 工具箱

- **合并**：多份 PDF 按指定顺序合成一份
- **拆分**：按可视化拆分线，或每 N 页切一份
- **提取**、**删除**：在页面墙上点选要处理的页
- **旋转**、**重排**：拖拽调整页面顺序
- **扁平化**：源文件允许时压平文档
- **签名**：手写签名后放到任意一页
- 撤销 / 重做、提交前缩略图预览、多文件结果打包 ZIP

### PDF / DOCX 水印

- `.pdf` 和 `.docx` 都支持
- 文字水印或图片水印，实心或描边，压在内容上方或下方
- 九宫格锚点、尺寸与不透明度可调
- 支持动态变量：`{date}`、`{time}`、`{datetime}`、`{year}`、`{filename}`、`{pages}`
- 能扫描已有文件，并移除你自己加过的水印

「去水印」只对你自己拥有的文档有意义，别拿它去冒用别人的成果。

### Word 文档对比

- 放入原始版和修改版两个 `.docx`
- 逐条列出版式修订及其类型：插入、删除、格式变更、移动
- 在预览里逐条跳转定位
- 导出保留修订标记的 redline `.docx`

### Markdown 转 Word

- 粘贴 Markdown 或打开 `.md` 文件，选一种版式，下载 `.docx`
- 标题、表格、代码块、列表、链接和图片都是 Word 原生元素，不是截图
- LaTeX `$...$` 与 `$$...$$` 转成可编辑的 Word 公式
- 标准 / 紧凑 / 报告 / 公文四套版式；公文按 GB/T 9704 的版心、仿宋三号正文和首行缩进两字符
- 预览渲染的是生成后的 `.docx` 本身，也就是你下载下来的那个文件

### Word 编辑器

- 打开 `.docx` 就地编辑：文字、表格、列表、图片
- 打印或导出结果
- 基于 `@docx-editor.dev` 引擎（Apache-2.0），按需懒加载

## 本地运行

```bash
npm install
npm run dev        # http://localhost:5173
```

```bash
npm run build      # 静态产物输出到 dist/
npm run preview    # 预览生产构建
```

Node 20.19+（或 22.12+）。

## 部署到任何地方

`npm run build` 产出纯静态站点，不需要 Node 运行时、不需要服务端函数。`base` 设为 `./`，放在子路径下也能跑：GitHub Pages、Cloudflare Pages、Netlify、Vercel（静态）、对象存储、nginx、Caddy。

## 「不上传」是怎么成立的

不是关掉了上传开关，而是代码里没有上传链路：

- `src/lib/generate-docx.ts` 接收 `File[]`，返回 `Blob`
- `src/lib/pdf-ops.ts` 接收 `Uint8Array`，返回 `Uint8Array`
- `src/lib/watermark/*` 在内存里重写文件字节
- `src/lib/docx-compare.worker.ts` 把对比引擎放在主线程之外

界面上唯一的外部 URL，是指回在线产品的那个链接。

## 在线版

不想装任何东西的话，同样的工具在 GenDocX 上可以直接用。同一套代码，独立工具这部分同样「文件不离开浏览器」。

| 工具 | 源码 | 在线版 |
|---|---|---|
| 图片转 Word | `src/components/ImageToWord.tsx` | [图片拼成 Word](https://gendocx.com/zh/tools/images-to-docx) |
| Markdown 转 Word | `src/components/MarkdownToWord.tsx` | [Markdown 转 Word](https://gendocx.com/zh/tools/markdown-to-docx) |
| PDF 工具箱 | `src/components/PdfToolbox.tsx` | [PDF 合并、拆分、签名](https://gendocx.com/zh/tools/pdf) |
| 水印 | `src/components/WatermarkTool.tsx` | [给 PDF 或 Word 加水印](https://gendocx.com/zh/tools/watermark) |
| 文档对比 | `src/components/WordCompare.tsx` | [对比两版 Word 文档](https://gendocx.com/zh/tools/compare) |
| 编辑器 | `src/components/WordEditor.tsx` | [在线编辑 Word 文档](https://gendocx.com/zh/tools/editor) |

## 和 GenDocX 是什么关系

这六个工具是 [GenDocX](https://gendocx.com) 的浏览器端核心。GenDocX 是一个文档工作台：在同一个窗口里编辑 Word 文件，同时和 AI agent 对话。agent 能读你上传的文件、产出 `.docx`、再交回来给你审阅，不需要另开一个转换站。

区别在范围，不在可信度：

- **这个仓库**是独立的工具箱。一个文件进，一个文件出，不要账号，什么都不存。
- **[gendocx.com](https://gendocx.com)** 在工具外面加了一层工作区：会话、文件历史、AI agent、模板。想看那是什么样子，可以从[工具页](https://gendocx.com/zh/tools)进，或者读[文档](https://gendocx.com/zh/docs)。

如果你只是想把一批扫描件转一次，clone 下来用完就不用再管。如果转完之后还要继续处理这份文档，那是主站要解决的事。

## 技术栈

React 19 · TypeScript · Vite · `docx` · `pdf-lib` · `pdfjs-dist` · `docxodus` · `@docx-editor.dev` · `fflate` · `lucide-react`

## 常见问题

**文件会被上传吗？**
不会。代码里没有上传链路。文件在本地读取，用 `URL.createObjectURL` 输出结果。

**能离线使用吗？**
可以。构建一次，用本地静态服务打开 `dist/`，页面加载后不再请求任何资源。

**扫描版 PDF 能用吗？**
合并、拆分、旋转、删除、重排、加水印都可以。不做 OCR，扫描的文字不可搜索。

**加密 PDF 能打开吗？**
不能。加密文件会被拒绝，而不是处理出半成品。

**输出的是真正的 .docx 吗？**
是。由 `docx` 库生成的标准 OOXML，不是把 HTML 改后缀。

**为什么包体看起来不小？**
`docxodus` 和编辑器引擎是 WebAssembly，只有打开「文档对比」或「编辑器」时才加载；另外三个工具用的是更小的包。

**可以商用吗？**
可以，MIT 许可。保留版权声明即可，不强制署名。

## 许可

MIT © GenDocX
