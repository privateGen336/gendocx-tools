import { de, deepMerge, en, fr, zhCN } from "@docx-editor.dev/i18n";
import type { Translations } from "@docx-editor.dev/i18n";

/**
 * 上游 @docx-editor.dev/i18n 提供 en / de / fr / zh-CN 等，但没有 es。
 * 站点支持 de / en / es / fr / zh 五种语言，为了让编辑器界面与页面文案一致，
 * 这里补一份西语高频词条（菜单、格式栏、样式、缩放等最常露出的区域）。
 *
 * `Translations` 是 DeepPartial：适配器会自动把缺失的键回退到英文，
 * 所以只需覆盖可见度最高的一批，不必复制全部 666 条。
 */
const esEditor = {
  _lang: "es",
  common: {
    cancel: "Cancelar",
    apply: "Aplicar",
    close: "Cerrar",
    comment: "Comentar",
    dismiss: "Descartar",
  },
  toolbar: {
    ariaLabel: "Barra de formato",
    file: "Archivo",
    format: "Formato",
    insert: "Insertar",
    review: "Revisar",
    open: "Abrir",
    save: "Guardar",
    print: "Imprimir",
    pageSetup: "Configurar página",
    image: "Imagen",
    table: "Tabla",
    insertTable: "Insertar tabla",
    break: "Salto",
  },
  titleBar: {
    untitled: "Sin título",
    documentNameAriaLabel: "Nombre del documento",
    menuBarAriaLabel: "Barra de menú",
  },
  styles: {
    selectAriaLabel: "Seleccionar estilo de párrafo",
    normalText: "Texto normal",
    title: "Título",
    subtitle: "Subtítulo",
    heading1: "Título 1",
    heading2: "Título 2",
    heading3: "Título 3",
  },
  alignment: {
    alignLeft: "Alinear a la izquierda",
    center: "Centrar",
    alignRight: "Alinear a la derecha",
    justify: "Justificar",
  },
  lists: {
    ariaLabel: "Formato de lista",
    bulletList: "Lista con viñetas",
    numberedList: "Lista numerada",
    decreaseIndent: "Reducir sangría",
    increaseIndent: "Aumentar sangría",
  },
  font: {
    selectAriaLabel: "Seleccionar fuente",
    sansSerif: "Sans Serif",
    serif: "Serif",
    monospace: "Monoespacio",
    documentFonts: "Fuentes del documento",
  },
  fontSize: {
    decrease: "Reducir tamaño de fuente",
    increase: "Aumentar tamaño de fuente",
    label: "Tamaño de fuente",
    listLabel: "Tamaños de fuente",
  },
  loading: { label: "Cargando" },
  zoom: {
    zoomOut: "Reducir zoom",
    zoomIn: "Aumentar zoom",
    zoomLevel: "Nivel de zoom",
    automatic: "Automático",
    fitWidth: "Ajustar al ancho",
  },
  editingMode: {
    label: "Modo de edición",
    editing: "Edición",
    suggesting: "Sugerencias",
    viewing: "Visualización",
  },
  formattingBar: {
    undo: "Deshacer",
    redo: "Rehacer",
    bold: "Negrita",
    italic: "Cursiva",
    underline: "Subrayado",
    strikethrough: "Tachado",
    fontColor: "Color de fuente",
    highlightColor: "Color de resaltado",
    insertLink: "Insertar enlace",
    superscript: "Superíndice",
    subscript: "Subíndice",
  },
} satisfies Translations;

/**
 * 上游 zh-CN 包比 en 落后一版：`toolbar.review`、段落对话框、右键菜单等
 * 共 100 余条是 `null`，适配器会静默回退英文——菜单里就出现了 "Review"。
 *
 * `deepMerge` 会跳过 null 并按层合并，所以这里只补缺的那批，其余中文仍来自 zhCN。
 * 术语对齐 Microsoft Word 简体中文版（审阅 / 格式刷 / 制表位 / 孤行控制）。
 * 未覆盖 `collaboration*`：本产品不启用协作功能，那批文案不会出现在界面上。
 */
const zhOverrides = {
  toolbar: {
    review: "审阅",
    insertTable: "插入表格",
  },
  formattingBar: {
    groups: { format: "格式" },
    formatPainter: "格式刷",
    formatPainterShortcut: "格式刷 (Ctrl+Alt+C / Ctrl+Alt+V)",
  },
  lineSpacing: {
    options: "行距选项…",
  },
  zoom: {
    automatic: "自动",
    fitWidth: "适应页宽",
  },
  comments: {
    resolve: "解决",
    reopen: "重新打开",
  },
  revisions: {
    paragraphMarkDeleted: "已删除的段落标记",
    paragraphMarkMovedFrom: "段落标记从此处移走",
    paragraphMarkMovedTo: "段落标记移到此位置",
  },
  reviewers: {
    markupOptions: "标记选项",
  },
  review: {
    resolved: "已解决",
    showResolvedComment: "显示已解决的批注",
  },
  contextMenu: {
    pasteWithoutFormatting: "只粘贴文本",
    pasteWithoutFormattingShortcut: "Ctrl+Shift+V",
    copyFormatting: "复制格式",
    copyFormattingShortcut: "Ctrl+Alt+C",
    pasteFormatting: "粘贴格式",
    pasteFormattingShortcut: "Ctrl+Alt+V",
  },
  equationPopup: {
    title: "编辑公式",
    inputLabel: "线性格式",
    placeholder: "输入公式",
    apply: "应用",
    delete: "删除",
    refused: "无法应用该公式，请检查语法后重试。",
  },
  disabledReason: {
    editorNotReady: "编辑器尚未就绪",
    notWired: "该控件尚未连接到编辑器命令",
    viewing: "文档处于查看模式",
    trackedOnly: "此文档仅允许以修订形式编辑",
    noDocument: "尚未加载文档",
    noReviewAuthors: "此文档没有审阅作者",
    noContentControl: "未选中内容控件",
    noCopiedFormatting: "请先复制格式",
    noDrawing: "未选中图片",
    drawingLocked: "该图片已锁定",
    commentSelection: "请先选中要添加批注的内容",
    mergeUnsupported: "不支持合并单元格",
    splitUnsupported: "不支持拆分单元格",
    imageSuggesting: "建议模式下不能修改图片",
    invalidValue: "该值不适用于此命令",
    unavailable: "此命令当前不可用",
    bodyOnlySectionBreak: "分节符需插入到文档正文中。",
    bodyOnlyToc: "目录需插入到文档正文中。",
    bodyOnlyNote: "脚注和尾注需从文档正文中添加。",
    furnitureOnlyPageField: "请先打开页眉或页脚，再插入页码。",
    suggestingSectionBreak: "请关闭修订建议后再插入分节符：它会改变下一节的起始位置，无法作为建议记录。",
    tableCellSectionBreak: "分节符不能放在表格内。",
    lockedSectionBreak: "下一节被锁定或链接的内容控件占用，此分节符无法修改它。",
    lockedContentBreak: "此处文本被内容控件锁定，无法在此分节。",
    nothingToCopyFormatting: "请选中文本以复制其格式",
    nothingToPaint: "请选中文本以应用格式",
  },
  dialogs: {
    paragraph: {
      title: "段落",
      general: "常规",
      alignment: "对齐方式",
      alignLeft: "左对齐",
      alignCenter: "居中",
      alignRight: "右对齐",
      alignJustify: "两端对齐",
      indentation: "缩进",
      beforeText: "左侧",
      afterText: "右侧",
      special: "特殊格式",
      specialNone: "（无）",
      specialFirstLine: "首行缩进",
      specialHanging: "悬挂缩进",
      by: "度量值",
      spacing: "间距",
      spaceBefore: "段前",
      spaceAfter: "段后",
      lineSpacing: "行距",
      ruleMultiple: "多倍行距",
      ruleAtLeast: "最小值",
      ruleExactly: "固定值",
      at: "设置值",
      contextualSpacing: "不在相同样式的段落之间添加空格",
      pagination: "换行和分页",
      keepNext: "与下段同页",
      keepLines: "段中不分页",
      widowControl: "孤行控制",
      pageBreakBefore: "段前分页",
      unitInches: "英寸",
      unitPoints: "磅",
      ok: "确定",
      cancel: "取消",
      tabStops: "制表位",
      tabPosition: "制表位位置",
      tabAlignment: "对齐方式",
      tabLeader: "前导符",
      tabAdd: "设置",
      tabClearAll: "全部清除",
      tabRemove: "清除",
      tabNone: "无",
      tabAlignLeft: "左对齐",
      tabAlignCenter: "居中",
      tabAlignRight: "右对齐",
      tabAlignDecimal: "小数点对齐",
      tabAlignBar: "竖线对齐",
      tabLeaderDot: "点线",
      tabLeaderHyphen: "虚线",
      tabLeaderUnderscore: "实线",
      tabEmpty: "尚未设置制表位。",
      tabPositionLabel: "{position} 英寸",
      tabRemoveAt: "清除 {position} 英寸处的制表位",
      refused: "这些值无法在此应用。",
      tabMixed: "所选段落的制表位不一致。",
      mixed: "多种值",
      notShown: "此处不显示",
    },
  },
} satisfies Translations;

/** zhCN 打底，只覆盖上游仍为 null 的词条。 */
const zhEditor = deepMerge(zhCN, zhOverrides) as Translations;

export const editorLocales: Record<string, Translations> = {
  zh: zhEditor,
  en,
  de,
  fr,
  es: esEditor,
};

export const defaultEditorLocale: Translations = en;
