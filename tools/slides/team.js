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
pres.title = "Vitoma 团队";
const C = pres.SchemeColor;

pres.defineSlideMaster({
  title: "TITLE_ONLY",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.75,
        fontSize: 30, bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "" } },
    { text: { text: "Vitoma｜AIVC 虚拟类器官", options: { x: 0.6, y: 7.05, w: 6, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.4, h: 0.3, fontSize: 10, color: C.accent5 },
});

pres.addSection({ title: "团队" });
const s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "团队" });
s.addText("团队：从类器官到 AI，飞轮的每一环都有人负责", { placeholder: "title" });

const members = [
  { ini: "卓", name: "卓悦 博士", role: "创始人 & CEO", lines: [
      "Verqura 实验室管理系统创始人",
      "海德堡大学 / DKFZ 发育生物学博士；密歇根大学学士",
      "Cell Stem Cell 2025、npj Precis Oncol 2024 作者",
      "类器官 · 单细胞多组学 · 图像 AI（YOLO / Cellpose）" ] },
  { ini: "田", name: "田伟利 博士", role: "CSO 首席科学官", lines: [
      "DKFZ / 海德堡大学分子神经遗传学资深科学家",
      "霍英东杰出青年学者；主持欧中重大科研项目",
      "Cell Stem Cell 2025、npj Precis Oncol 2024 核心成员",
      "长期研究胶质母细胞瘤复发与耐药" ] },
  { ini: "王", name: "王中杰 博士", role: "CTO 首席技术官", lines: [
      "巴斯德研究所博士后；曾任职哈佛医学院",
      "近 10 年多组学分析经验",
      "多模态建模、自监督预训练、因果建模",
      "擅长端到端可复用分析流程" ] },
  { ini: "兰", name: "兰晓静", role: "COO 首席运营官", lines: [
      "海德堡大学经济学硕士；浙江大学管理学学士",
      "FRM 持证",
      "国际金融机构风控与合规经验",
      "跨境尽职调查经验" ] },
];

const X0 = 0.6, GAP = 0.3, CW = (12.13 - 3 * GAP) / 4, CY = 1.3, CH = 2.95;
members.forEach((m, i) => {
  const x = X0 + i * (CW + GAP);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: CY, w: CW, h: CH, rectRadius: 0.08,
    fill: { color: C.background2 }, line: { color: C.background2 }, objectName: `card-${i}` });
  // avatar (photo placeholder)
  s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: CY + 0.2, w: 0.7, h: 0.7,
    fill: { color: i === 0 ? C.accent2 : C.accent1 }, line: { color: "FFFFFF", width: 1 }, objectName: `avatar-${i}` });
  s.addText(m.ini, { x: x + 0.2, y: CY + 0.2, w: 0.7, h: 0.7, align: "center", valign: "middle",
    fontSize: 20, bold: true, color: C.background1, margin: 0, isTextBox: true });
  s.addText(m.name, { x: x + 1.05, y: CY + 0.2, w: CW - 1.2, h: 0.38, fontSize: 16, bold: true,
    color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  s.addText(m.role, { x: x + 1.05, y: CY + 0.56, w: CW - 1.2, h: 0.34, fontSize: 13, bold: true,
    color: i === 0 ? C.accent2 : C.accent1, margin: 0, valign: "middle", isTextBox: true });
  s.addText(m.lines.map((t, k) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: k < m.lines.length - 1 } })),
    { x: x + 0.18, y: CY + 1.05, w: CW - 0.26, h: CH - 1.15, fontSize: 11, color: C.text1,
      valign: "top", margin: 0, paraSpaceAfter: 5, isTextBox: true });
});

// roles & responsibilities table
const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 } } });
const rows = [
  [hdr("职务"), hdr("成员"), hdr("核心职责"), hdr("对应飞轮环节")],
  ["创始人 & CEO", "卓悦 博士", "公司战略与融资；实验数字化与数据标准（Verqura）；关键合作伙伴关系", "数据生产体系"],
  ["CSO 首席科学官", "田伟利 博士", "科学路线与类器官体系；临床合作与三级验证设计；疾病机制研究", "患者样本 · 类器官 · 验证"],
  ["CTO 首席技术官", "王中杰 博士", "AIVC 预测模型；多模态 / 自监督 / 因果建模；数据平台与算力架构", "模型 · 锁定预测 · 评分"],
  ["COO 首席运营官", "兰晓静", "运营与项目管理；财务与合规（人类遗传资源、数据安全）；商务拓展", "临床合作 · 合规 · 商务"],
];
s.addTable(rows, { x: X0, y: 4.5, w: 12.13, colW: [1.9, 1.5, 6.03, 2.7],
  fontSize: 12, color: HEX.dk1, valign: "middle", rowH: 0.4, margin: [0, 0.1, 0, 0.1],
  border: { type: "solid", pt: 0.75, color: "D9E7E6" }, fill: { color: "FFFFFF" }, objectName: "roles-table" });

s.addText([
  { text: "团队即壁垒：", options: { bold: true, color: C.accent2 } },
  { text: "患者采样 → 类器官建立 → 数据构建 → 模型迭代 → 前瞻验证，全链路在团队内部闭环", options: { color: C.text1 } },
], { x: X0, y: 6.6, w: 12.13, h: 0.35, fontSize: 12, margin: 0, isTextBox: true });

s.addNotes("头像为占位，正式版替换为成员照片。卓悦博士的介绍来自简历（2026-08）；其他成员沿用原 BP。");

(async () => {
  await pres.writeFile({ fileName: process.argv[2] });
  await applyTheme(process.argv[2], THEME);
  console.log("written", process.argv[2]);
})();
