# GenDocX Tools

**纯浏览器文档工具：多张图片拼成 Word，PDF 合并 / 拆分 / 旋转 / 签名，文件不上传。**

所有处理都在当前标签页里完成。没有服务器、没有接口调用、不需要登录、没有统计埋点。打开 DevTools 的 Network 面板，一边操作一边看——它不会发请求。

[English](./README.md)

---

## 为什么做这个

大多数「免费在线转换」网站会把文件传到你看不见的服务器上。对于一份签好的合同或一张证件扫描件，这是实打实的风险，而且你没法验证传过去之后发生了什么。

这套工具反过来做：整条处理链路都在浏览器里跑。

- 图片用 `createImageBitmap` 解码，用 `docx` 排版，在内存里打包成 `.docx`。
- PDF 由 `pdf.js` 在 Web Worker 里解析，用 `pdf-lib` 修改，最后以 `Blob` 下载。
- 唯一的网络请求，是加载这个页面本身的静态资源。

代码是 MIT 许可，你可以读、可以审计、可以 fork，也可以直接放进自己的产品里。

## 包含哪些工具

### 图片转 Word

- 支持 JPG、PNG、WebP、GIF、AVIF、BMP，一次最多 50 张
- 拖拽排序、旋转、删除，分页位置实时预览
- 纸张（A4 / Letter / Legal）、方向（自动或锁定）、每页张数
- 边距（无 / 窄 / 标准 / 宽）、质量（适应页面或保留原像素）、JPEG / PNG 编码
- 输出是真正的 `.docx`，Word、Pages、WPS、Google Docs、LibreOffice 都能打开

### PDF 工具箱

- **合并**：多份 PDF 按你指定的顺序合成一份
- **拆分**：按可视化拆分线，或每 N 页切一份
- **提取**、**删除**：在页面墙上点选要处理的页
- **旋转**、**重排**：拖拽调整页面顺序
- **扁平化**：在源文件允许的情况下压平文档
- **签名**：手写签名后放到任意一页
- 撤销 / 重做、提交前缩略图预览、多文件结果打包成 ZIP 下载

加密 PDF 会被直接拒绝：解密需要密码，静默输出一个损坏的文件比明确说「不行」更糟。

## 本地运行

```bash
npm install
npm run dev        # http://localhost:5173
```

```bash
npm run build      # 静态产物输出到 dist/
npm run preview    # 本地预览生产构建
```

环境要求：Node 20.19+（或 22.12+）。

## 部署到任何地方

`npm run build` 产出的是纯静态站点，不需要 Node 运行时。`base` 已设为 `./`，放在子路径下也能跑：

- GitHub Pages
- Cloudflare Pages / Netlify / Vercel（静态模式）
- 对象存储、nginx、Caddy，甚至一个 U 盘

## 「不上传」是怎么保证的

不是「关掉了上传开关」，而是代码里根本没有上传逻辑：

- `src/lib/generate-docx.ts` 接收一组 `File` 对象，返回 `Blob`
- `src/lib/pdf-ops.ts` 接收 `Uint8Array`，返回 `Uint8Array`
- `src/lib/pdf-preview.ts` 通过 `pdf.js` 把页面渲染到 `<canvas>`

界面上唯一的外部 URL，是指回在线产品的那个链接。

## 技术栈

React 19 · TypeScript · Vite · `docx` · `pdf-lib` · `pdfjs-dist` · `fflate` · `lucide-react`

## 和 GenDocX 的关系

这些工具是 [GenDocX](https://gendocx.com) 的浏览器端核心。GenDocX 是一个文档工作台，可以在同一个界面里编辑 Word 文件、同时和 AI agent 对话。

如果你不想自己跑，同样的工具在 **[gendocx.com/tools](https://gendocx.com/tools)** 可以直接用——不用安装；独立工具的这部分，文件同样不会离开你的浏览器。

## 常见问题

**文件会被上传吗？**
不会。代码里没有上传链路。工具在本地读取 `File`，用 `URL.createObjectURL` 输出结果。

**能离线使用吗？**
可以。构建一次，用本地静态服务打开 `dist/`，页面加载完成后不再请求任何东西。

**扫描版 PDF 能用吗？**
合并、拆分、旋转、删除、重排都可以。但不做 OCR，扫描的文字不可搜索，也不能当文本编辑。

**加密 PDF 能打开吗？**
不能。加密文件会被拒绝，而不是处理出一份半成品。

**输出的是真正的 .docx 吗？**
是。由 `docx` 库生成的标准 OOXML，不是把 HTML 改后缀。

**可以商用吗？**
可以，MIT 许可。保留版权声明即可，不强制署名。

## 许可

MIT © GenDocX
