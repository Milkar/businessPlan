// Vitoma BP deck: cover, contents, part 1, team
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const THEME = {
  name: "Vitoma",
  headFontFace: "Microsoft YaHei",
  bodyFontFace: "Microsoft YaHei",
  colors: {
    dk1: "1B2A33", lt1: "FFFFFF", dk2: "0B3C5D", lt2: "EEF4F4",
    accent1: "0B3C5D", accent2: "1FA39A", accent3: "7FC8C0",
    accent4: "F2A541", accent5: "5B6B73", accent6: "D9E7E6",
    hlink: "1FA39A", folHlink: "0B3C5D",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Vitoma 商业计划书";
pres.company = "Vitoma";
const C = pres.SchemeColor;
const X0 = 0.6, W = 12.13;

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "COVER",
  background: { color: HEX.dk2 },
  objects: [],
});
pres.defineSlideMaster({
  title: "SECTION",
  background: { color: HEX.dk2 },
  objects: [
    { placeholder: { options: { name: "num", type: "body", x: X0, y: 2.0, w: 3, h: 1.3, fontSize: 72, bold: true, align: "left",
        color: C.accent3, margin: 0, valign: "bottom" }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: X0, y: 3.35, w: W, h: 0.9, fontSize: 40, bold: true, align: "left",
        color: C.background1, margin: 0, valign: "middle" }, text: "" } },
    { placeholder: { options: { name: "sub", type: "body", x: X0, y: 4.25, w: W, h: 0.6, fontSize: 20, align: "left",
        color: C.accent3, margin: 0, valign: "top" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "TITLE_ONLY",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: X0, y: 0.35, w: W, h: 0.75,
        fontSize: 28, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: "Vitoma｜AIVC 虚拟类器官", options: { x: X0, y: 7.05, w: 6, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.4, h: 0.3, fontSize: 10, color: C.accent5 },
});

const tb = (s, text, o) => s.addText(text, Object.assign({ margin: 0, isTextBox: true }, o));
const card = (s, x, y, w, h, name, fill) => s.addShape(pres.shapes.ROUNDED_RECTANGLE,
  { x, y, w, h, rectRadius: 0.08, fill: { color: fill || C.background2 }, line: { color: fill || C.background2 }, objectName: name });
const numDot = (s, x, y, n, d, color) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: color || C.accent2 }, line: { color: color || C.accent2 }, objectName: `dot-${n}-${x}-${y}` });
  tb(s, String(n), { x, y, w: d, h: d, align: "center", valign: "middle", fontSize: d > 0.6 ? 20 : 14, bold: true, color: C.background1 });
};

// ---------- 1 cover ----------
pres.addSection({ title: "封面" });
let s = pres.addSlide({ masterName: "COVER", sectionTitle: "封面" });
tb(s, "Vitoma", { x: X0, y: 1.2, w: W, h: 1.2, fontSize: 66, bold: true, color: C.background1, fontFace: "Arial" });
tb(s, "Virtual Intelligence Twin & Organism Modeling Architecture", { x: X0, y: 2.4, w: W, h: 0.5, fontSize: 20, color: C.accent3, fontFace: "Arial" });
tb(s, "让每一次治疗，先在患者自己的类器官上预演", { x: X0, y: 3.35, w: W, h: 0.8, fontSize: 34, bold: true, color: C.background1 });
tb(s, "Rehearse every treatment on the patient's own organoids — before it begins.", { x: X0, y: 4.15, w: W, h: 0.45, fontSize: 16, italic: true, color: C.accent3, fontFace: "Arial" });
tb(s, "AIVC 虚拟类器官｜让复发与耐药，在发生之前被看见", { x: X0, y: 5.1, w: W, h: 0.45, fontSize: 18, color: C.background1 });
tb(s, "本轮融资 5000 万元人民币（覆盖 12 个月）｜2026 年 10 月", { x: X0, y: 6.3, w: W, h: 0.4, fontSize: 14, color: C.accent3 });
s.addNotes("封面。Vitoma 的首字母拆写：Virtual Intelligence Twin & Organism Modeling Architecture。");

// ---------- 2 contents ----------
pres.addSection({ title: "目录" });
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "目录" });
s.addText("目录", { placeholder: "title" });
const toc = [
  ["01", "为什么要做", "市场需求与现有缺口"],
  ["02", "为什么我们能做成", "核心缺口 · 技术优势 · 护城河 · 团队"],
  ["03", "做成之后", "社会价值 · 商业模式 · 融资预算与 Plan B · 投资回报"],
];
toc.forEach(([n, t, sub], i) => {
  const y = 1.55 + i * 1.65;
  card(s, X0, y, W, 1.35, `toc-${n}`);
  tb(s, n, { x: X0 + 0.4, y, w: 1.4, h: 1.35, fontSize: 44, bold: true, color: C.accent2, valign: "middle", fontFace: "Arial" });
  tb(s, t, { x: X0 + 2.0, y: y + 0.22, w: 9.5, h: 0.5, fontSize: 24, bold: true, color: C.text2, valign: "middle" });
  tb(s, sub, { x: X0 + 2.0, y: y + 0.75, w: 9.5, h: 0.4, fontSize: 15, color: C.accent5, valign: "middle" });
});

// ---------- 3 part 1 divider ----------
pres.addSection({ title: "第一部分 为什么要做" });
const P1 = "第一部分 为什么要做";
s = pres.addSlide({ masterName: "SECTION", sectionTitle: P1 });
s.addText("01", { placeholder: "num" });
s.addText("为什么要做", { placeholder: "title" });
s.addText("市场需求与现有缺口", { placeholder: "sub" });

// ---------- 4 1.1 industry cost ----------
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: P1 });
s.addText("药物研发的「试错」已不可承受", { placeholder: "title" });
const stats = [["10+ 年", "创新药从发现到上市的平均耗时"], ["$20 亿+", "平均研发花费，且仍在上升"], ["临床后期", "失败暴露的阶段——代价最高"]];
stats.forEach(([v, l], i) => {
  const x = X0 + i * (3.9 + 0.215);
  card(s, x, 1.35, 3.9, 1.85, `stat-${i}`);
  tb(s, v, { x: x + 0.3, y: 1.5, w: 3.3, h: 0.95, fontSize: 44, bold: true, color: C.accent2, valign: "middle", fontFace: "Arial" });
  tb(s, l, { x: x + 0.3, y: 2.45, w: 3.3, h: 0.55, fontSize: 15, color: C.text1, valign: "top" });
});
tb(s, "为什么失败到后期才暴露？早期模型预测不了患者之间的差异，更追不上肿瘤随时间的演化", { x: X0, y: 3.5, w: W, h: 0.45, fontSize: 16, bold: true, color: C.text2 });
const chain = [["细胞模型", "体外有效"], ["动物模型", "动物有效"], ["患者身上", "无效甚至有害"]];
chain.forEach(([a, b], i) => {
  const x = X0 + i * 4.115;
  const last = i === 2;
  card(s, x, 4.1, 3.6, 1.05, `chain-${i}`, last ? HEX.accent4 : HEX.lt2);
  tb(s, a, { x: x + 0.25, y: 4.18, w: 3.1, h: 0.42, fontSize: 15, bold: true, color: last ? C.background1 : C.text2, valign: "middle" });
  tb(s, b, { x: x + 0.25, y: 4.6, w: 3.1, h: 0.42, fontSize: 15, color: last ? C.background1 : C.text1, valign: "middle" });
  if (!last) tb(s, "→", { x: x + 3.6, y: 4.1, w: 0.515, h: 1.05, fontSize: 24, bold: true, color: C.accent2, align: "center", valign: "middle" });
});
card(s, X0, 5.45, W, 1.05, "contradiction", HEX.dk2);
tb(s, [
  { text: "根本矛盾  ", options: { bold: true, color: C.accent3 } },
  { text: "候选药物太多、患者差异太大、生命状态随时间持续变化，真实实验永远做不过来；耐药与复发要等到患者治疗失败后才被看见。", options: { color: C.background1 } },
], { x: X0 + 0.3, y: 5.5, w: W - 0.6, h: 0.95, fontSize: 15, valign: "middle" });
tb(s, "来源：行业公开数据（具体出处与年份待补充）", { x: X0, y: 6.62, w: 8, h: 0.3, fontSize: 10, color: C.accent5 });

// ---------- 5 1.2 clinical gaps ----------
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: P1 });
s.addText("治疗决策缺少及时、患者特异性的功能证据", { placeholder: "title" });
const gaps = [
  ["分子图谱是静态快照", "基因组能提示脆弱性，但疗效还取决于细胞状态、生态、治疗史与可塑性；单次检测只代表采样那一刻"],
  ["反馈滞后", "影像与临床结局只在患者已经接受治疗之后，才说明发生了什么"],
  ["工具割裂", "类器官药筛停在「敏感 / 不敏感」，很少转化为经过校准的患者纵向预测"],
];
gaps.forEach(([t, d], i) => {
  const x = X0 + i * (3.9 + 0.215);
  card(s, x, 1.35, 3.9, 2.3, `gap-${i}`);
  numDot(s, x + 0.3, 1.6, i + 1, 0.55);
  tb(s, t, { x: x + 1.0, y: 1.6, w: 2.7, h: 0.55, fontSize: 17, bold: true, color: C.text2, valign: "middle" });
  tb(s, d, { x: x + 0.3, y: 2.35, w: 3.3, h: 1.6, fontSize: 14, color: C.text1, valign: "top" });
});
card(s, X0, 3.95, W, 1.25, "evidence", HEX.dk2);
tb(s, [
  { text: "外部证据  ", options: { bold: true, color: C.accent3 } },
  { text: "Peng et al., ", options: { color: C.background1 } },
  { text: "Cell Stem Cell", options: { italic: true, color: C.background1 } },
  { text: " 2025：患者来源脑肿瘤类器官（IPTO）高保真保留肿瘤生态系统，并在前瞻性临床研究中预测患者治疗响应（PFS / OS，p = 0.034 / 0.033）", options: { color: C.background1 } },
], { x: X0 + 0.3, y: 4.0, w: W - 0.6, h: 1.15, fontSize: 15, valign: "middle" });
tb(s, [
  { text: "同时要诚实  ", options: { bold: true, color: C.accent2 } },
  { text: "已有前瞻研究（如结直肠癌）表明，类器官的预测表现取决于治疗类别与实施情境——需要的是「校准 + 前瞻验证」，而不只是药筛。", options: { color: C.text1 } },
], { x: X0, y: 5.45, w: W, h: 0.8, fontSize: 15, valign: "middle" });

// ---------- 6 1.3 why now ----------
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: P1 });
s.addText("为什么是现在：四项条件同步成熟", { placeholder: "title" });
const conds = [
  ["类器官技术", "IPTO 高保真保留肿瘤生态系统，个体差异与微环境同时在线，可纵向追踪演化"],
  ["成像与测量", "长时程活细胞成像、高内涵成像、单细胞 / 空间组学，把细胞状态变成可计算的多模态数据"],
  ["数据与算力", "多模态数据基础设施成熟，大规模训练成本显著下降"],
  ["AI 方法", "自监督预训练 + 扰动建模，让 AI 能学习「干预之后进入什么新状态」"],
];
conds.forEach(([t, d], i) => {
  const col = i % 2, row = Math.floor(i / 2);
  const x = X0 + col * (5.915 + 0.3), y = 1.35 + row * (1.75 + 0.3);
  card(s, x, y, 5.915, 1.75, `cond-${i}`);
  numDot(s, x + 0.3, y + 0.3, i + 1, 0.55);
  tb(s, t, { x: x + 1.05, y: y + 0.3, w: 4.6, h: 0.55, fontSize: 18, bold: true, color: C.text2, valign: "middle" });
  tb(s, d, { x: x + 1.05, y: y + 0.88, w: 4.6, h: 0.8, fontSize: 14, color: C.text1, valign: "top" });
});
card(s, X0, 5.55, W, 1.0, "summary", HEX.dk2);
tb(s, [
  { text: "第一部分小结  ", options: { bold: true, color: C.accent3 } },
  { text: "治疗决策需要「事前预测」，而今天的工具只能在治疗失败之后给出答案。技术条件已经成熟——缺的是什么、我们如何补上，见第二部分。", options: { color: C.background1 } },
], { x: X0 + 0.3, y: 5.6, w: W - 0.6, h: 0.9, fontSize: 15, valign: "middle" });

// ---------- 7 team ----------
pres.addSection({ title: "第二部分 为什么我们能做成" });
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "第二部分 为什么我们能做成" });
s.addText("团队：科学 × 数据 × 模型 × 运营，核心能力实现自持闭环", { placeholder: "title" });
const members = [
  { name: "卓悦 博士", role: "CEO 首席执行官", lines: [
      "创立 Verqura 实验室管理系统，打通样本、实验与数据管理",
      "精通 iPSC 脑类器官构建、CRISPR 基因组编辑与标记",
      "海德堡大学 / DKFZ 发育生物学博士；美国密歇根大学安娜堡分校学士",
      "Cell Stem Cell 2025、npj Precis Oncol 2024 作者；擅长单细胞组学与图像 AI" ] },
  { name: "田伟利 博士", role: "CSO 首席科学官", lines: [
      "DKFZ / 海德堡大学分子神经遗传学资深科学家",
      "深耕神经调控与脑类器官方向，长期研究胶质母细胞瘤复发与耐药",
      "霍英东杰出青年学者；主持欧中重大科研项目",
      "Cell Stem Cell 2025、npj Precis Oncol 2024 核心成员" ] },
  { name: "王中杰 博士", role: "CTO 首席技术官", lines: [
      "巴斯德研究所博士后；曾任职哈佛医学院",
      "精通大规模计算与机器学习预测建模，近 10 年多组学分析经验",
      "擅长搭建端到端可复用分析流程，支撑临床与生物医药项目的数据挖掘与决策输出" ] },
  { name: "兰晓静", role: "COO 首席运营官", lines: [
      "海德堡大学经济学硕士",
      "浙江大学管理学学士，竺可桢学院创业与创新管理强化班",
      "FRM 持证；国际金融机构风控合规与跨境尽调经验" ] },
];
const GAP = 0.3, CW = (W - 3 * GAP) / 4, CY = 1.3, CH = 3.15;
members.forEach((m, i) => {
  const x = X0 + i * (CW + GAP);
  card(s, x, CY, CW, CH, `member-${i}`);
  tb(s, m.name, { x: x + 0.22, y: CY + 0.15, w: CW - 0.4, h: 0.42, fontSize: 18, bold: true, color: C.text1, valign: "middle" });
  tb(s, m.role, { x: x + 0.22, y: CY + 0.56, w: CW - 0.4, h: 0.34, fontSize: 14, bold: true, color: C.accent2, valign: "middle" });
  tb(s, m.lines.map((t, k) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: k < m.lines.length - 1 } })),
    { x: x + 0.18, y: CY + 1.0, w: CW - 0.3, h: CH - 1.08, fontSize: 11, color: C.text1, valign: "top", paraSpaceAfter: 4 });
});
const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 } } });
s.addTable([
  [hdr("职务"), hdr("成员"), hdr("核心职责")],
  ["CEO 首席执行官", "卓悦 博士", "公司战略与融资；实验数字化与数据标准（Verqura）；关键合作伙伴关系"],
  ["CSO 首席科学官", "田伟利 博士", "科学路线与类器官体系；临床合作与三级验证设计；疾病机制研究"],
  ["CTO 首席技术官", "王中杰 博士", "AIVC 预测模型；大规模计算与多模态 / 自监督 / 因果建模；数据平台与算力架构"],
  ["COO 首席运营官", "兰晓静", "运营与项目管理；财务与合规（人类遗传资源、数据安全）；商务拓展"],
], { x: X0, y: 4.65, w: W, colW: [2.1, 1.8, 8.23], fontSize: 12, color: HEX.dk1, valign: "middle", rowH: 0.38,
  margin: [0, 0.1, 0, 0.1], border: { type: "solid", pt: 0.75, color: "D9E7E6" }, fill: { color: "FFFFFF" }, objectName: "roles-table" });
tb(s, [
  { text: "团队即壁垒：", options: { bold: true, color: C.accent2 } },
  { text: "患者采样 → 类器官建立 → 数据构建 → 模型迭代 → 前瞻验证，全链路在团队内部闭环", options: { color: C.text1 } },
], { x: X0, y: 6.65, w: W, h: 0.32, fontSize: 12 });
s.addNotes("四位成员同属创始团队。第二部分其他页面待卓悦博士、田伟利博士补充科学技术内容后制作。");

(async () => {
  await pres.writeFile({ fileName: process.argv[2] });
  await applyTheme(process.argv[2], THEME);
  console.log("written", process.argv[2]);
})();
