const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.author = 'Trae';
pres.title = '新型信息化工程生产范式';
pres.subject = '本地智能体IDE+云服务器+分布式代码版本控制+自进化智能体+通用大模型';

// ============================================================
// DIMENSIONS
// ============================================================
pres.layout = 'LAYOUT_16x9';
const SLIDE_W = 10;
const SLIDE_H = 5.625;
const MARGIN = 0.5;
const CONTENT_X = MARGIN;
const CONTENT_Y = MARGIN;
const CONTENT_W = SLIDE_W - (2 * MARGIN);
const CONTENT_H = SLIDE_H - (2 * MARGIN);
const CENTER_X = SLIDE_W / 2;
const CENTER_Y = SLIDE_H / 2;

// ============================================================
// PALETTE
// ============================================================
const C = {
  primary: "1E2761",
  secondary: "CADCFC",
  accent: "F97316",
  bg: "F8FAFC",
  card: "FFFFFF",
  text: "1F2937",
  text2: "4B5563",
  text3: "9CA3AF",
  good: "059669",
  bad: "DC2626",
  warn: "D97706",
  dark: "0F172A",
  ice: "E0E7FF",
};

// ============================================================
// SHADOW FACTORY
// ============================================================
const makeShadow = () => ({
  type: "outer", blur: 12, offset: 4, angle: 135,
  color: "000000", opacity: 0.28
});
const makeHeroShadow = () => ({
  type: "outer", blur: 16, offset: 6, angle: 135,
  color: "000000", opacity: 0.32
});

// ============================================================
// CONTAINER SYSTEM
// ============================================================
function calculateScaledImageOpts(opts) {
  const { path, w: targetW, h: targetH, x = 0, y = 0, mode = 'cover', ...rest } = opts;
  if (!path || !targetW || !targetH) return opts;
  return { path, x, y, w: targetW, h: targetH, sizing: { type: mode, w: targetW, h: targetH }, ...rest };
}

function createVirtualNode(type, data, parentX = 0, parentY = 0) {
  const opts = data.opts || {};
  const node = { type, data, absX: parentX + (opts.x || 0), absY: parentY + (opts.y || 0), w: opts.w || 0, h: opts.h || 0, children: [] };
  node.addShape = function(shapeType, opts = {}) { const child = createVirtualNode('shape', { shapeType, opts }, node.absX, node.absY); node.children.push(child); return child; };
  node.addText = function(text, opts = {}) {
    const safeOpts = { fit: "shrink", ...opts };
    const bulletRe = /^(?:[\u2022\u2023\u25E6\u2043\u2219\u00B7\u25CF\u25CB\u2013\u2014]\s*|\-\s+)/;
    if (Array.isArray(text)) { text = text.map(item => { if (item && item.options && item.options.bullet && typeof item.text === 'string') { return { ...item, text: item.text.replace(bulletRe, '') }; } return item; }); }
    const child = createVirtualNode('text', { text, opts: safeOpts }, node.absX, node.absY); node.children.push(child); return child;
  };
  node.addImage = function(opts = {}) { const scaledOpts = calculateScaledImageOpts(opts); const child = createVirtualNode('image', { opts: scaledOpts }, node.absX, node.absY); node.children.push(child); return child; };
  node.addTable = function(tableData, opts = {}) { const child = createVirtualNode('table', { tableData, opts }, node.absX, node.absY); node.children.push(child); return child; };
  return node;
}

function flattenNode(node, realSlide, pres) {
  const absOpts = { ...node.data.opts, x: node.absX, y: node.absY };
  if (node.type === 'shape') realSlide.addShape(node.data.shapeType, absOpts);
  else if (node.type === 'text') realSlide.addText(node.data.text, absOpts);
  else if (node.type === 'image') realSlide.addImage(absOpts);
  else if (node.type === 'table') realSlide.addTable(node.data.tableData, absOpts);
  node.children.forEach(child => flattenNode(child, realSlide, pres));
}

const originalAddSlide = pres.addSlide.bind(pres);
pres.addSlide = function(options) {
  const realSlide = originalAddSlide(options);
  const virtualSlide = {
    children: [], _realSlide: realSlide,
    set background(val) { realSlide.background = val; },
    get background() { return realSlide.background; },
    addShape: function(shapeType, opts = {}) { const node = createVirtualNode('shape', { shapeType, opts }, 0, 0); this.children.push(node); return node; },
    addText: function(text, opts = {}) { const safeOpts = { fit: "shrink", ...opts }; const node = createVirtualNode('text', { text, opts: safeOpts }, 0, 0); this.children.push(node); return node; },
    addImage: function(opts = {}) { const scaledOpts = calculateScaledImageOpts(opts); const node = createVirtualNode('image', { opts: scaledOpts }, 0, 0); this.children.push(node); return node; },
    addTable: function(tableData, opts = {}) { const node = createVirtualNode('table', { tableData, opts }, 0, 0); this.children.push(node); return node; },
    addChart: function(chartType, data, opts = {}) { realSlide.addChart(chartType, data, opts); },
    render: function() { this.children.forEach(child => flattenNode(child, realSlide, pres)); }
  };
  return virtualSlide;
};

// ============================================================
// HELPERS (enlarged fonts)
// ============================================================
function cardGrid(cols, rows, opts = {}) {
  const { startX = CONTENT_X, startY = 1.5, areaW = CONTENT_W, areaH = CONTENT_H - 1.0, gapX = 0.25, gapY = 0.25, padding = 0.3 } = opts;
  const cardW = (areaW - gapX * (cols - 1)) / cols;
  const cardH = (areaH - gapY * (rows - 1)) / rows;
  const cards = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      cards.push({ x: startX + c * (cardW + gapX), y: startY + r * (cardH + gapY), w: cardW, h: cardH, padding });
    }
  }
  return { cards, cardW, cardH };
}

function addEyebrow(slide, text, y = 0.35, color = C.accent) {
  slide.addText(text, {
    x: CONTENT_X, y, w: CONTENT_W, h: 0.35,
    fontSize: 14, fontFace: "Calibri", color, bold: true,
    charSpacing: 0.5, uppercase: true
  });
}

function addSlideTitle(slide, text, y = 0.7) {
  slide.addText(text, {
    x: CONTENT_X, y, w: CONTENT_W, h: 0.6,
    fontSize: 34, fontFace: "Georgia", color: C.text, bold: true,
    charSpacing: 1.5
  });
}

function addLede(slide, text, y = 1.35) {
  slide.addText(text, {
    x: CONTENT_X, y, w: CONTENT_W * 0.85, h: 0.45,
    fontSize: 18, fontFace: "Calibri", color: C.text2
  });
}

function addFooter(slide, text, align = "right") {
  slide.addText(text, {
    x: CONTENT_X, y: SLIDE_H - 0.4, w: CONTENT_W, h: 0.28,
    fontSize: 12, fontFace: "Calibri", color: C.text3, align
  });
}

// ============================================================
// SLIDE 1: COVER
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { path: "images/bg-cover_16x9.jpg" };
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: SLIDE_W, h: SLIDE_H,
    fill: { color: C.dark, transparency: 45 }
  });
  slide.addText("信息工程 · 范式重构", {
    x: MARGIN, y: 0.35, w: 4, h: 0.35,
    fontSize: 14, fontFace: "Calibri", color: "FFFFFF", bold: true,
    charSpacing: 0.5, uppercase: true
  });
  slide.addText("2026", {
    x: SLIDE_W - MARGIN - 1, y: 0.35, w: 1, h: 0.35,
    fontSize: 14, fontFace: "Calibri", color: "FFFFFF", align: "right"
  });
  slide.addText("新型信息化工程\n生产范式", {
    x: MARGIN, y: 1.7, w: 7, h: 1.8,
    fontSize: 48, fontFace: "Georgia", color: "FFFFFF", bold: true,
    charSpacing: 2.5
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: MARGIN, y: 3.6, w: 1.2, h: 0.06,
    fill: { color: C.accent }
  });
  slide.addText("本地智能体IDE · 云服务器 · 分布式代码版本控制 · 自进化智能体 · 通用大模型 —— 五位一体的融合范式", {
    x: MARGIN, y: 3.8, w: 7.5, h: 0.6,
    fontSize: 16, fontFace: "Calibri", color: "E2E8F0"
  });
  slide.addText("新型生产范式展望", {
    x: MARGIN, y: SLIDE_H - 0.45, w: 4, h: 0.28,
    fontSize: 12, fontFace: "Calibri", color: "94A3B8"
  });
  slide.render();
}

// ============================================================
// SLIDE 2: TABLE OF CONTENTS (3 items + image)
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addEyebrow(slide, "Contents · 目录", 0.35, C.primary);
  addSlideTitle(slide, "本次分享的三个部分", 0.7);

  const tocItems = [
    { num: "01", title: "时代背景", desc: "软件工程范式的演进历程与传统范式的瓶颈" },
    { num: "02", title: "五大核心支柱", desc: "逐一拆解新型范式的五大构成要素与能力" },
    { num: "03", title: "协同生产范式", desc: "端到端工作流与三层协同架构如何运转" },
  ];

  // Left side: 3 stacked cards
  const leftW = CONTENT_W * 0.52;
  const cardH = 1.2;
  const gap = 0.12;
  const startY = 1.35;
  tocItems.forEach((item, i) => {
    const y = startY + i * (cardH + gap);
    let card = slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: CONTENT_X, y, w: leftW, h: cardH,
      fill: { color: C.card }, rectRadius: 0.1, shadow: makeShadow()
    });
    card.addText(item.num, {
      x: 0.25, y: 0.15, w: 0.7, h: 0.5,
      fontSize: 30, fontFace: "Calibri", color: C.primary, bold: true
    });
    card.addText(item.title, {
      x: 1.05, y: 0.15, w: leftW - 1.3, h: 0.4,
      fontSize: 22, fontFace: "Georgia", color: C.text, bold: true
    });
    card.addText(item.desc, {
      x: 1.05, y: 0.6, w: leftW - 1.3, h: 0.5,
      fontSize: 14, fontFace: "Calibri", color: C.text2
    });
  });

  // Right side: coding background image
  const imgX = CONTENT_X + leftW + 0.35;
  const imgW = CONTENT_W - leftW - 0.35;
  const imgH = cardH * 3 + gap * 2;
  slide.addImage({
    path: "images/bg-coding_16x9.jpg",
    x: imgX, y: startY, w: imgW, h: imgH,
    sizing: { type: "cover", w: imgW, h: imgH }
  });

  addFooter(slide, "新型信息化工程生产范式");
  slide.render();
}

// ============================================================
// SLIDE 3: SECTION 01
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.primary };
  slide.addText("Section · 01", {
    x: CENTER_X - 2, y: 1.5, w: 4, h: 0.4,
    fontSize: 15, fontFace: "Calibri", color: C.accent, bold: true,
    align: "center", charSpacing: 0.5, uppercase: true
  });
  slide.addText("时代背景：范式演进", {
    x: CENTER_X - 4, y: 1.95, w: 8, h: 1.0,
    fontSize: 44, fontFace: "Georgia", color: "FFFFFF", bold: true,
    align: "center", charSpacing: 2
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: CENTER_X - 0.5, y: 3.1, w: 1, h: 0.04,
    fill: { color: C.accent }
  });
  slide.addText("软件工程生产范式，正经历从「人力密集」到「智能体密集」的根本性跃迁。", {
    x: CENTER_X - 4, y: 3.3, w: 8, h: 0.5,
    fontSize: 18, fontFace: "Calibri", color: C.secondary, align: "center"
  });
  slide.render();
}

// ============================================================
// SLIDE 4: 四次范式跃迁
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addEyebrow(slide, "Evolution · 范式跃迁", 0.35, C.primary);
  addSlideTitle(slide, "四次范式跃迁", 0.7);
  addLede(slide, "每一次跃迁，都把「人」从重复劳动中解放一阶。", 1.3);

  const steps = [
    { year: "1968", title: "工程化", desc: "结构化方法、瀑布模型，软件工程成为独立学科" },
    { year: "2001", title: "敏捷化", desc: "迭代交付、以人为本，响应变化重于遵循计划" },
    { year: "2009", title: "云原生化", desc: "DevOps、微服务、CI/CD，交付走向自动化" },
    { year: "2024+", title: "智能体化", desc: "LLM 驱动智能体直接编码与决策，人机协同成主流", hl: true },
  ];
  const nodeW = 1.8, nodeH = 0.8, hGap = 0.4;
  const totalW = steps.length * nodeW + (steps.length - 1) * hGap;
  const startX = (SLIDE_W - totalW) / 2;
  const nodeY = 2.1;

  steps.forEach((s, i) => {
    const x = startX + i * (nodeW + hGap);
    const fillColor = s.hl ? C.accent : C.primary;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: nodeY, w: nodeW, h: nodeH,
      fill: { color: fillColor }, rectRadius: 0.12, shadow: makeShadow()
    });
    slide.addText(s.year, {
      x, y: nodeY + 0.05, w: nodeW, h: nodeH * 0.4,
      align: "center", valign: "middle",
      fontSize: 14, fontFace: "Calibri", color: "FFFFFF", bold: true
    });
    slide.addText(s.title, {
      x, y: nodeY + nodeH * 0.4, w: nodeW, h: nodeH * 0.5,
      align: "center", valign: "middle",
      fontSize: 16, fontFace: "Georgia", color: "FFFFFF", bold: true
    });
    if (i < steps.length - 1) {
      slide.addShape(pres.shapes.LINE, {
        x: x + nodeW, y: nodeY + nodeH / 2, w: hGap, h: 0,
        line: { color: "94A3B8", width: 2, endArrowType: "triangle" }
      });
    }
    slide.addText(s.desc, {
      x: x - 0.1, y: nodeY + nodeH + 0.25, w: nodeW + 0.2, h: 1.0,
      fontSize: 13, fontFace: "Calibri", color: C.text2, align: "center"
    });
  });
  addFooter(slide, "01 · 时代背景");
  slide.render();
}

// (Slide 5 removed — traditional bottlenecks slide not in latest HTML)

// ============================================================
// SLIDE 6: SECTION 02
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.primary };
  slide.addText("Section · 02", {
    x: CENTER_X - 2, y: 1.5, w: 4, h: 0.4,
    fontSize: 15, fontFace: "Calibri", color: C.accent, bold: true,
    align: "center", charSpacing: 0.5, uppercase: true
  });
  slide.addText("五大核心支柱", {
    x: CENTER_X - 3.5, y: 1.95, w: 7, h: 1.0,
    fontSize: 44, fontFace: "Georgia", color: "FFFFFF", bold: true,
    align: "center", charSpacing: 2
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: CENTER_X - 0.5, y: 3.1, w: 1, h: 0.04,
    fill: { color: C.accent }
  });
  slide.addText("五类能力各司其职，又彼此咬合，共同构成新型生产范式的底座。", {
    x: CENTER_X - 4, y: 3.3, w: 8, h: 0.5,
    fontSize: 18, fontFace: "Calibri", color: C.secondary, align: "center"
  });
  slide.render();
}

// ============================================================
// SLIDE 7: 五位一体总览
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addEyebrow(slide, "Overview · 五位一体", 0.35, C.primary);
  addSlideTitle(slide, "五位一体的能力底座", 0.7);
  addLede(slide, "五大支柱分别覆盖「交互、算力、资产、进化、智能」五个维度，缺一不可。", 1.3);

  const pillars = [
    { tag: "P1", title: "本地智能体IDE", desc: "人机交互的现场" },
    { tag: "P2", title: "云服务器", desc: "弹性算力与协作枢纽" },
    { tag: "P3", title: "分布式版本控制", desc: "工程知识资产基线", hl: true },
    { tag: "P4", title: "自进化智能体", desc: "持续变强的执行大脑" },
    { tag: "P5", title: "通用大模型", desc: "通用智能供给底座" },
  ];
  const cardW = (CONTENT_W - 0.3 * 4) / 5;
  const cardH = 2.6;
  const startY = 2.0;
  pillars.forEach((p, i) => {
    const x = CONTENT_X + i * (cardW + 0.3);
    const fillColor = p.hl ? C.primary : C.card;
    const textColor = p.hl ? "FFFFFF" : C.text;
    const descColor = p.hl ? "E2E8F0" : C.text3;
    let card = slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: startY, w: cardW, h: cardH,
      fill: { color: fillColor }, rectRadius: 0.1,
      shadow: p.hl ? makeHeroShadow() : makeShadow()
    });
    card.addText(p.tag, {
      x: 0.15, y: 0.2, w: cardW - 0.3, h: 0.3,
      fontSize: 13, fontFace: "Calibri", color: p.hl ? C.accent : C.primary,
      bold: true, charSpacing: 0.5
    });
    card.addText(p.title, {
      x: 0.15, y: 0.6, w: cardW - 0.3, h: 0.5,
      fontSize: 16, fontFace: "Georgia", color: textColor, bold: true
    });
    card.addText(p.desc, {
      x: 0.15, y: 1.2, w: cardW - 0.3, h: 0.8,
      fontSize: 13, fontFace: "Calibri", color: descColor
    });
  });
  slide.addText("交互 → 算力 → 资产 → 进化 → 智能，形成闭环", {
    x: CONTENT_X, y: SLIDE_H - 0.95, w: CONTENT_W, h: 0.3,
    fontSize: 14, fontFace: "Calibri", color: C.text3, align: "center"
  });
  addFooter(slide, "02 · 五大核心支柱");
  slide.render();
}

// ============================================================
// PILLAR DETAIL HELPER (enlarged)
// ============================================================
function addPillarSlide(tag, title, definition, value, capabilities) {
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addEyebrow(slide, tag, 0.35, C.primary);
  addSlideTitle(slide, title, 0.7);

  // Left: definition + value
  const leftW = CONTENT_W * 0.42;
  slide.addText(definition, {
    x: CONTENT_X, y: 1.35, w: leftW, h: 1.2,
    fontSize: 16, fontFace: "Calibri", color: C.text2
  });
  // Value box
  slide.addShape(pres.shapes.RECTANGLE, {
    x: CONTENT_X, y: 2.6, w: leftW, h: 1.4,
    fill: { color: C.card }, shadow: makeShadow()
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: CONTENT_X, y: 2.6, w: 0.06, h: 1.4,
    fill: { color: C.accent }
  });
  slide.addText("核心价值", {
    x: CONTENT_X + 0.25, y: 2.7, w: leftW - 0.5, h: 0.3,
    fontSize: 13, fontFace: "Calibri", color: C.accent, bold: true,
    charSpacing: 0.5, uppercase: true
  });
  slide.addText(value, {
    x: CONTENT_X + 0.25, y: 3.05, w: leftW - 0.5, h: 0.8,
    fontSize: 15, fontFace: "Calibri", color: C.text2
  });

  // Right: 2x2 capability cards
  const capW = CONTENT_W * 0.52;
  const capStartX = CONTENT_X + CONTENT_W * 0.46;
  const capCardW = (capW - 0.2) / 2;
  const capCardH = 1.7;
  capabilities.forEach((cap, i) => {
    const cx = capStartX + (i % 2) * (capCardW + 0.2);
    const cy = 1.35 + Math.floor(i / 2) * (capCardH + 0.2);
    let card = slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: cx, y: cy, w: capCardW, h: capCardH,
      fill: { color: C.card }, rectRadius: 0.1, shadow: makeShadow()
    });
    card.addText(cap.title, {
      x: 0.2, y: 0.2, w: capCardW - 0.4, h: 0.4,
      fontSize: 18, fontFace: "Georgia", color: C.text, bold: true
    });
    card.addText(cap.desc, {
      x: 0.2, y: 0.65, w: capCardW - 0.4, h: 0.9,
      fontSize: 14, fontFace: "Calibri", color: C.text3
    });
  });

  addFooter(slide, "02 · 五大核心支柱");
  slide.render();
}

// ============================================================
// SLIDES 8-12: PILLAR DETAILS
// ============================================================
addPillarSlide(
  "P1 · 本地智能体IDE", "人机交互的现场",
  "在开发者本地运行、以智能体为核心的集成开发环境。智能体可直接读写工程、操作终端、自主规划并完成端到端任务。",
  "低延迟交互、私有数据不出域、深度理解工程上下文。",
  [
    { title: "自然语言编程", desc: "用需求描述驱动代码生成与修改" },
    { title: "多文件上下文", desc: "跨文件感知工程结构与依赖关系" },
    { title: "本地工具调用", desc: "直接操作终端、构建、测试与调试" },
    { title: "数据本地留存", desc: "敏感代码不出域，保障隐私合规" },
  ]
);

addPillarSlide(
  "P2 · 云服务器", "弹性算力与协作枢纽",
  "承载弹性算力、CI/CD 流水线与模型推理的云端基础设施。按需供给算力、统一运行环境、全局可达。",
  "按需扩缩、免运维、支撑远程协同与大规模推理。",
  [
    { title: "弹性算力调度", desc: "按负载自动扩缩，峰值无忧" },
    { title: "CI/CD 自动交付", desc: "构建、测试、部署全链路自动化" },
    { title: "模型推理服务", desc: "统一承载大模型与微服务调用" },
    { title: "远程开发环境", desc: "环境一致性，随地接入工程" },
  ]
);

addPillarSlide(
  "P3 · 分布式版本控制", "工程知识的资产基线",
  "以 Git 为代表的去中心化代码与知识资产管理体系。每一次变更可追溯、可回滚、可并行，是团队协作的事实基线。",
  "团队协作基线、工程知识可复用可审计、容灾与回滚保障。",
  [
    { title: "分支并行开发", desc: "多特性并行，互不阻塞" },
    { title: "变更追溯审计", desc: "谁、何时、为何改动一目了然" },
    { title: "工程知识沉淀", desc: "把经验固化为可复用资产" },
    { title: "一键回滚容灾", desc: "故障可快速回到已知良好状态" },
  ]
);

addPillarSlide(
  "P4 · 自进化智能体", "持续变强的执行大脑",
  "能依据执行反馈，持续优化自身策略、工具与记忆的智能体。越用越强，最终形成组织级智能资产。",
  "降低重复劳动、累积组织智能、逼近「零维护」自治。",
  [
    { title: "经验记忆累积", desc: "把成功路径沉淀为可复用记忆" },
    { title: "工具自生成", desc: "按需创造并复用新工具" },
    { title: "错误自修复", desc: "自诊断失败并迭代修正" },
    { title: "策略迭代优化", desc: "闭环学习，持续提升成功率" },
  ]
);

addPillarSlide(
  "P5 · 通用大模型", "通用智能的供给底座",
  "具备通用理解、推理与生成能力的基础大模型，是整个范式的智能供给底座，为智能体提供「大脑」。",
  "通用智能即取即用、降低专用模型成本、跨域能力迁移。",
  [
    { title: "代码生成补全", desc: "从片段到模块的端到端生成" },
    { title: "文档与用例", desc: "自动撰写文档、测试与用例" },
    { title: "复杂推理决策", desc: "多步推理与任务分解" },
    { title: "多模态理解", desc: "图文、界面、日志一并解析" },
  ]
);

// ============================================================
// SLIDE 13: SECTION 03
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.primary };
  slide.addText("Section · 03", {
    x: CENTER_X - 2, y: 1.5, w: 4, h: 0.4,
    fontSize: 15, fontFace: "Calibri", color: C.accent, bold: true,
    align: "center", charSpacing: 0.5, uppercase: true
  });
  slide.addText("协同生产范式", {
    x: CENTER_X - 3.5, y: 1.95, w: 7, h: 1.0,
    fontSize: 44, fontFace: "Georgia", color: "FFFFFF", bold: true,
    align: "center", charSpacing: 2
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: CENTER_X - 0.5, y: 3.1, w: 1, h: 0.04,
    fill: { color: C.accent }
  });
  slide.addText("五大支柱并非简单堆叠，而是通过数据流与反馈闭环，组合成一条端到端的生产链路。", {
    x: CENTER_X - 4, y: 3.3, w: 8, h: 0.5,
    fontSize: 18, fontFace: "Calibri", color: C.secondary, align: "center"
  });
  slide.render();
}

// ============================================================
// SLIDE 14: 端到端工作流
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addEyebrow(slide, "Pipeline · 端到端链路", 0.35, C.primary);
  addSlideTitle(slide, "一条端到端的生产链路", 0.7);
  addLede(slide, "从需求到上线再回流，五支柱串联成闭环，反馈以「分钟」计。", 1.3);

  const nodes = [
    { label: "需求理解", sub: "大模型解析意图" },
    { label: "智能体编码", sub: "本地IDE执行" },
    { label: "版本提交", sub: "分布式分支协作" },
    { label: "云端流水线", sub: "CI/CD 自动化" },
    { label: "部署上线", sub: "弹性算力承载" },
    { label: "自进化反馈", sub: "闭环回流需求", hl: true },
  ];
  const nodeW = 1.25, nodeH = 1.0, hGap = 0.2;
  const totalW = nodes.length * nodeW + (nodes.length - 1) * hGap;
  const startX = (SLIDE_W - totalW) / 2;
  const nodeY = 2.2;

  nodes.forEach((n, i) => {
    const x = startX + i * (nodeW + hGap);
    const fillColor = n.hl ? C.accent : C.primary;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: nodeY, w: nodeW, h: nodeH,
      fill: { color: fillColor }, rectRadius: 0.1, shadow: n.hl ? makeHeroShadow() : makeShadow()
    });
    slide.addText(n.label, {
      x, y: nodeY + 0.1, w: nodeW, h: 0.4,
      align: "center", valign: "middle",
      fontSize: 14, fontFace: "Georgia", color: "FFFFFF", bold: true
    });
    slide.addText(n.sub, {
      x, y: nodeY + 0.5, w: nodeW, h: 0.35,
      align: "center", valign: "middle",
      fontSize: 11, fontFace: "Calibri", color: "E2E8F0"
    });
    if (i < nodes.length - 1) {
      slide.addShape(pres.shapes.LINE, {
        x: x + nodeW, y: nodeY + nodeH / 2, w: hGap, h: 0,
        line: { color: "94A3B8", width: 1.5, endArrowType: "triangle" }
      });
    }
  });

  slide.addText("自进化反馈回流至需求理解与编码，形成持续优化的闭环", {
    x: CONTENT_X, y: SLIDE_H - 1.0, w: CONTENT_W, h: 0.3,
    fontSize: 14, fontFace: "Calibri", color: C.text3, align: "center"
  });
  addFooter(slide, "03 · 协同生产范式");
  slide.render();
}

// ============================================================
// SLIDE 15: 三层协同架构
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addEyebrow(slide, "Architecture · 三层协同", 0.35, C.primary);
  addSlideTitle(slide, "三层协同架构", 0.7);
  addLede(slide, "五大支柱映射为三层架构：智能体编排、协作资产、算力模型。", 1.3);

  const layers = [
    { name: "编排层\nAgent", color: "1E3A5F", cells: ["本地智能体IDE", "自进化智能体", "任务规划调度"] },
    { name: "资产层\nCollab", color: C.primary, cells: ["分布式版本控制", "工程知识库", "CI/CD 流水线"], hl: true },
    { name: "底座层\nInfra", color: "1E3A5F", cells: ["云服务器", "通用大模型", "推理与网关"] },
  ];
  const layerH = 0.95, cellW = 2.4, cellGap = 0.2, layerGap = 0.3;
  const labelW = 1.0;
  const labelGap = 0.25;
  const startX = CONTENT_X + labelW + labelGap;
  const startY = 2.0;

  layers.forEach((layer, li) => {
    const y = startY + li * (layerH + layerGap);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: CONTENT_X, y, w: labelW, h: layerH,
      fill: { color: layer.color }, rectRadius: 0.08
    });
    slide.addText(layer.name, {
      x: CONTENT_X, y, w: labelW, h: layerH,
      align: "center", valign: "middle",
      fontSize: 12, fontFace: "Calibri", color: "FFFFFF", bold: true
    });
    layer.cells.forEach((cell, ci) => {
      const cx = startX + ci * (cellW + cellGap);
      const fillColor = layer.hl ? C.card : "F1F5F9";
      const borderColor = layer.hl ? C.accent : "E2E8F0";
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: cx, y, w: cellW, h: layerH,
        fill: { color: fillColor }, rectRadius: 0.08,
        line: { color: borderColor, width: layer.hl ? 2 : 1 },
        shadow: layer.hl ? makeShadow() : undefined
      });
      slide.addText(cell, {
        x: cx, y, w: cellW, h: layerH,
        align: "center", valign: "middle",
        fontSize: 14, fontFace: "Calibri", color: C.text
      });
    });
  });
  addFooter(slide, "03 · 协同生产范式");
  slide.render();
}

// ============================================================
// SLIDE 16: 写在最后 (Section divider)
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.primary };
  slide.addText("写在最后", {
    x: CENTER_X - 2, y: 1.5, w: 4, h: 0.4,
    fontSize: 15, fontFace: "Calibri", color: C.accent, bold: true,
    align: "center", charSpacing: 0.5, uppercase: true
  });
  slide.addText("价值与落地", {
    x: CENTER_X - 3, y: 1.95, w: 6, h: 1.0,
    fontSize: 44, fontFace: "Georgia", color: "FFFFFF", bold: true,
    align: "center", charSpacing: 2
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x: CENTER_X - 0.5, y: 3.1, w: 1, h: 0.04,
    fill: { color: C.accent }
  });
  slide.addText("从效率收益、实施路径到风险应对，把范式从「概念」落到「工程」。", {
    x: CENTER_X - 4, y: 3.3, w: 8, h: 0.5,
    fontSize: 18, fontFace: "Calibri", color: C.secondary, align: "center"
  });
  slide.render();
}

// ============================================================
// SLIDE 17: BS系统技术栈与服务架构
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addEyebrow(slide, "Architecture · 技术路线", 0.3, C.primary);
  addSlideTitle(slide, "BS系统技术栈与服务架构", 0.6);

  // 6 layers of tech stack
  const layers = [
    { name: "用户层", hl: false, cells: ["🌐 Web浏览器", "📱 移动端H5", "🖥️ 桌面客户端", "🔧 管理后台"] },
    { name: "代理层", hl: true, cells: ["⚖️ Nginx负载均衡", "🔒 HTTPS/SSL", "🚪 API网关", "🛡️ WAF防火墙"] },
    { name: "应用服务", hl: false, cells: ["📦 用户服务", "📦 业务服务", "📦 权限服务", "📦 文件服务", "📦 消息服务"] },
    { name: "数据中台", hl: true, cells: ["📊 数据采集", "📊 数据治理", "📊 数据分析", "📊 数据可视化"] },
    { name: "中间件", hl: false, cells: ["⚡ Redis缓存", "💌 RabbitMQ", "🔍 ElasticSearch", "📋 Nacos"] },
    { name: "数据库", hl: true, cells: ["🗄️ MySQL集群", "🗄️ MongoDB", "🗄️ ClickHouse", "🗄️ MinIO"] },
  ];
  const labelW = 0.85;
  const labelGap = 0.15;
  const cellGap = 0.08;
  const layerGap = 0.08;
  const layerH = 0.42;
  const cellAreaX = CONTENT_X + labelW + labelGap;
  const cellAreaW = CONTENT_W - labelW - labelGap;
  const startY = 1.35;

  layers.forEach((layer, li) => {
    const y = startY + li * (layerH + layerGap);
    // Label
    const labelColor = layer.hl ? C.accent : C.primary;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: CONTENT_X, y, w: labelW, h: layerH,
      fill: { color: labelColor }, rectRadius: 0.06
    });
    slide.addText(layer.name, {
      x: CONTENT_X, y, w: labelW, h: layerH,
      align: "center", valign: "middle",
      fontSize: 11, fontFace: "Calibri", color: "FFFFFF", bold: true
    });
    // Cells
    const numCells = layer.cells.length;
    const cellW = (cellAreaW - cellGap * (numCells - 1)) / numCells;
    layer.cells.forEach((cell, ci) => {
      const cx = cellAreaX + ci * (cellW + cellGap);
      const fillColor = layer.hl ? "FFF5EB" : C.card;
      const borderColor = layer.hl ? C.accent : "E2E8F0";
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: cx, y, w: cellW, h: layerH,
        fill: { color: fillColor }, rectRadius: 0.05,
        line: { color: borderColor, width: layer.hl ? 1.5 : 1 },
        shadow: makeShadow()
      });
      slide.addText(cell, {
        x: cx, y, w: cellW, h: layerH,
        align: "center", valign: "middle",
        fontSize: 10, fontFace: "Calibri", color: C.text
      });
    });
  });

  // Roles flow at bottom
  const roles = ["产品经理", "需求分析", "UI/UX设计", "开发", "测试", "运维"];
  const roleY = startY + 6 * (layerH + layerGap) + 0.1;
  const roleW = 1.1;
  const roleGap = 0.25;
  const arrowW = 0.3;
  const rolesTotalW = roles.length * roleW + (roles.length - 1) * roleGap;
  const rolesStartX = (SLIDE_W - rolesTotalW) / 2;
  roles.forEach((role, i) => {
    const rx = rolesStartX + i * (roleW + roleGap);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: rx, y: roleY, w: roleW, h: 0.35,
      fill: { color: C.card },
      line: { color: C.accent, width: 1.5 },
      rectRadius: 0.17
    });
    slide.addText(role, {
      x: rx, y: roleY, w: roleW, h: 0.35,
      align: "center", valign: "middle",
      fontSize: 10, fontFace: "Calibri", color: C.accent, bold: true
    });
    if (i < roles.length - 1) {
      slide.addText("→", {
        x: rx + roleW, y: roleY, w: roleGap, h: 0.35,
        align: "center", valign: "middle",
        fontSize: 12, fontFace: "Calibri", color: C.text3
      });
    }
  });

  addFooter(slide, "价值与落地");
  slide.render();
}

// ============================================================
// SLIDE 18: 从0开始的 Vibe Coding SOP
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addEyebrow(slide, "SOP · 从0开始 Vibe Coding", 0.3, C.primary);
  addSlideTitle(slide, "从0开始的 Vibe Coding SOP", 0.6);

  const steps = [
    { num: "01", title: "需求拆解", desc: "用自然语言描述产品目标，让智能体拆解为可执行的功能模块。", warn: "⚠ 需求必须具体可验证，避免模糊描述。" },
    { num: "02", title: "架构约定", desc: "明确技术栈、目录结构与代码规范，写入项目文档供智能体遵循。", warn: "⚠ 不先约定架构，智能体会自行发挥导致技术栈混乱。" },
    { num: "03", title: "增量编码", desc: "按模块逐步生成，每次只聚焦一个功能点，小步快跑。", warn: "⚠ 一次性生成全部代码，出错后难以定位和回滚。" },
    { num: "04", title: "即时验证", desc: "每生成一个模块立即运行测试，确保可用再继续下一步。", warn: "⚠ 跳过验证堆积到最后，bug成本指数级上升。" },
    { num: "05", title: "版本提交", desc: "每个可用模块即时commit，写清变更说明，保留回滚能力。", warn: "⚠ 不提交版本，智能体改坏了无法回退到可用状态。" },
    { num: "06", title: "人工把关", desc: "关键逻辑和安全相关代码必须人工Review，不可全权交给智能体。", warn: "⚠ 盲目信任智能体输出，安全漏洞与逻辑错误会被掩盖。" },
  ];

  const cols = 2;
  const cardGapX = 0.2;
  const cardGapY = 0.12;
  const cardW = (CONTENT_W - cardGapX) / cols;
  const cardH = 1.15;
  const startY = 1.3;

  steps.forEach((s, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = CONTENT_X + col * (cardW + cardGapX);
    const y = startY + row * (cardH + cardGapY);

    let card = slide.addShape(pres.shapes.RECTANGLE, {
      x, y, w: cardW, h: cardH,
      fill: { color: C.card }, shadow: makeShadow()
    });
    // Left accent bar
    card.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 0.06, h: cardH,
      fill: { color: C.accent }
    });
    // Number
    card.addText(s.num, {
      x: 0.2, y: 0.12, w: 0.55, h: 0.4,
      fontSize: 20, fontFace: "Calibri", color: C.accent, bold: true
    });
    // Title
    card.addText(s.title, {
      x: 0.8, y: 0.1, w: cardW - 1.0, h: 0.3,
      fontSize: 15, fontFace: "Georgia", color: C.text, bold: true
    });
    // Description
    card.addText(s.desc, {
      x: 0.8, y: 0.42, w: cardW - 1.0, h: 0.3,
      fontSize: 10, fontFace: "Calibri", color: C.text2
    });
    // Warning
    card.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y: 0.78, w: cardW - 1.0, h: 0.32,
      fill: { color: "FEF2F2" }
    });
    card.addText(s.warn, {
      x: 0.85, y: 0.78, w: cardW - 1.1, h: 0.32,
      fontSize: 8, fontFace: "Calibri", color: C.bad, valign: "middle"
    });
  });

  addFooter(slide, "价值与落地");
  slide.render();
}

// ============================================================
// SLIDE 19: 实施路线
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addEyebrow(slide, "Roadmap · 实施路径", 0.35, C.primary);
  addSlideTitle(slide, "四阶段实施路线", 0.7);

  const phases = [
    { tag: "NOW · 试点", title: "单点引入", items: ["引入智能体IDE", "接入通用大模型", "建立试点项目"], hl: true },
    { tag: "NEXT · 集成", title: "链路打通", items: ["接入云服务器", "版本控制规范化", "打通CI/CD"] },
    { tag: "LATER · 自治", title: "闭环进化", items: ["部署自进化智能体", "构建知识沉淀", "形成反馈闭环"] },
    { tag: "VISION · 范式化", title: "规模复用", items: ["组织级智能体平台", "智能资产规模化", "持续自运营"] },
  ];
  const colW = (CONTENT_W - 0.2 * 3) / 4;
  const startY = 1.5;
  const colH = 3.0;
  phases.forEach((p, i) => {
    const x = CONTENT_X + i * (colW + 0.2);
    const fillColor = p.hl ? C.primary : C.card;
    const textColor = p.hl ? "FFFFFF" : C.text;
    const tagColor = p.hl ? C.accent : C.primary;
    const itemColor = p.hl ? "E2E8F0" : C.text2;

    let card = slide.addShape(pres.shapes.RECTANGLE, {
      x, y: startY, w: colW, h: colH,
      fill: { color: fillColor }, shadow: makeShadow()
    });
    card.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.15, y: 0.15, w: colW - 0.3, h: 0.4,
      fill: { color: p.hl ? "2A3A6B" : C.ice }, rectRadius: 0.15
    });
    card.addText(p.tag, {
      x: 0.15, y: 0.15, w: colW - 0.3, h: 0.4,
      align: "center", valign: "middle",
      fontSize: 11, fontFace: "Calibri", color: tagColor, bold: true
    });
    card.addText(p.title, {
      x: 0.15, y: 0.65, w: colW - 0.3, h: 0.4,
      fontSize: 17, fontFace: "Georgia", color: textColor, bold: true
    });
    p.items.forEach((item, ii) => {
      card.addText(item, {
        x: 0.15, y: 1.1 + ii * 0.45, w: colW - 0.3, h: 0.4,
        fontSize: 13, fontFace: "Calibri", color: itemColor, bullet: true
      });
    });
  });
  addFooter(slide, "价值与落地");
  slide.render();
}

// ============================================================
// SLIDE 20: 挑战与应对
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addEyebrow(slide, "Challenges · 挑战与应对", 0.35, C.primary);
  addSlideTitle(slide, "三类核心挑战及应对", 0.7);

  const challenges = [
    { num: "①", title: "智能体可靠性 · 幻觉与错误决策", desc: "应对：人机协同把关 + 版本控制回滚兜底，关键决策保留人工确认节点。", color: C.warn, bg: "FFFBEB" },
    { num: "②", title: "数据安全与合规 · 代码与知识上云", desc: "应对：本地IDE保留敏感数据 + 分级授权 + 私有化模型部署，建立数据出境边界。", color: C.bad, bg: "FEF2F2" },
    { num: "③", title: "工程文化转型 · 从「写代码」到「编排智能体」", desc: "应对：建立智能体协作规范与度量体系，重塑工程师角色与能力模型。", color: C.good, bg: "F0FDF4" },
  ];
  const cardH = 1.1;
  const gap = 0.15;
  const startY = 1.4;
  challenges.forEach((ch, i) => {
    const y = startY + i * (cardH + gap);
    let card = slide.addShape(pres.shapes.RECTANGLE, {
      x: CONTENT_X, y, w: CONTENT_W, h: cardH,
      fill: { color: ch.bg }, shadow: makeShadow()
    });
    card.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 0.06, h: cardH,
      fill: { color: ch.color }
    });
    card.addText(ch.num + " " + ch.title, {
      x: 0.3, y: 0.15, w: CONTENT_W - 0.6, h: 0.4,
      fontSize: 18, fontFace: "Georgia", color: C.text, bold: true
    });
    card.addText(ch.desc, {
      x: 0.3, y: 0.6, w: CONTENT_W - 0.6, h: 0.4,
      fontSize: 14, fontFace: "Calibri", color: C.text2
    });
  });
  addFooter(slide, "价值与落地");
  slide.render();
}

// ============================================================
// SLIDE 21: 总结
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.bg };
  addEyebrow(slide, "Takeaways · 总结", 0.35, C.primary);
  addSlideTitle(slide, "三句话总结", 0.7);

  const takeaways = [
    { num: "01", title: "范式本质", desc: "从「人力驱动」转向「智能体驱动」，人负责编排与决策，智能体负责执行与进化。" },
    { num: "02", title: "五位一体", desc: "五大支柱咬合形成数据闭环，让生产越跑越快、越用越强，反馈以分钟计。" },
    { num: "03", title: "落地关键", desc: "以版本控制为基线、以自进化为引擎，逐步构建组织级智能资产，而非一步到位。" },
  ];
  const cardH = 1.0;
  const gap = 0.2;
  const startY = 1.5;
  takeaways.forEach((tk, i) => {
    const y = startY + i * (cardH + gap);
    let card = slide.addShape(pres.shapes.RECTANGLE, {
      x: CONTENT_X, y, w: CONTENT_W, h: cardH,
      fill: { color: C.card }, shadow: makeShadow()
    });
    card.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 0.05, h: cardH,
      fill: { color: C.accent }
    });
    card.addText(tk.num, {
      x: 0.3, y: 0.15, w: 0.6, h: 0.6,
      fontSize: 26, fontFace: "Calibri", color: C.accent, bold: true
    });
    card.addText(tk.title, {
      x: 1.0, y: 0.15, w: 2, h: 0.35,
      fontSize: 18, fontFace: "Georgia", color: C.text, bold: true
    });
    card.addText(tk.desc, {
      x: 1.0, y: 0.5, w: CONTENT_W - 1.4, h: 0.45,
      fontSize: 14, fontFace: "Calibri", color: C.text2
    });
  });
  addFooter(slide, "价值与落地");
  slide.render();
}

// ============================================================
// SLIDE 22: Thanks
// ============================================================
{
  let slide = pres.addSlide();
  slide.background = { color: C.dark };
  slide.addText("Thanks", {
    x: CENTER_X - 4, y: 1.5, w: 8, h: 1.5,
    fontSize: 60, fontFace: "Georgia", color: "FFFFFF", bold: true,
    align: "center", charSpacing: 3
  });
  slide.addText("让智能体成为生产力，让工程师回归创造。", {
    x: CENTER_X - 4, y: 3.2, w: 8, h: 0.5,
    fontSize: 18, fontFace: "Georgia", color: C.secondary, align: "center", italic: true
  });
  slide.addText("新型信息化工程生产范式  ·  2026", {
    x: CENTER_X - 4, y: 4.0, w: 8, h: 0.3,
    fontSize: 12, fontFace: "Calibri", color: "64748B", align: "center"
  });
  slide.render();
}

// ============================================================
// WRITE
// ============================================================
pres.writeFile({ fileName: "新型信息化工程生产范式.pptx" });
console.log("PPTX generated successfully!");
