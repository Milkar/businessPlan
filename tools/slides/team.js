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
  { name: "卓悦 博士", role: "CEO 首席执行官", lines: [
      "创立 Verqura 实验室管理系统，打通样本、实验与数据管理",
      "精通 iPSC 脑类器官构建、CRISPR 基因组编辑与标记",
      "海德堡大学 / DKFZ 发育生物学博士，获亥姆霍兹研究生院奖学金",
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
      "浙江大学管理学学士，就读竺可桢学院创业与创新管理强化班",
      "FRM 持证；国际金融机构风控合规与跨境尽调经验" ] },
];

const X0 = 0.6, GAP = 0.3, CW = (12.13 - 3 * GAP) / 4, CY = 1.2, CH = 3.2;
members.forEach((m, i) => {
  const x = X0 + i * (CW + GAP);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: CY, w: CW, h: CH, rectRadius: 0.08,
    fill: { color: C.background2 }, line: { color: C.background2 }, objectName: `card-${i}` });
  s.addText(m.name, { x: x + 0.22, y: CY + 0.15, w: CW - 0.4, h: 0.42, fontSize: 18, bold: true,
    color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  s.addText(m.role, { x: x + 0.22, y: CY + 0.56, w: CW - 0.4, h: 0.34, fontSize: 14, bold: true,
    color: C.accent2, margin: 0, valign: "middle", isTextBox: true });
  s.addText(m.lines.map((t, k) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: k < m.lines.length - 1 } })),
    { x: x + 0.18, y: CY + 1.0, w: CW - 0.3, h: CH - 1.08, fontSize: 11, color: C.text1,
      valign: "top", margin: 0, paraSpaceAfter: 5, isTextBox: true });
});

// roles & responsibilities table
const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 } } });
const rows = [
  [hdr("职务"), hdr("成员"), hdr("核心职责")],
  ["CEO 首席执行官", "卓悦 博士", "公司战略与融资；实验数字化与数据标准（Verqura）；关键合作伙伴关系"],
  ["CSO 首席科学官", "田伟利 博士", "科学路线与类器官体系；临床合作与三级验证设计；疾病机制研究"],
  ["CTO 首席技术官", "王中杰 博士", "AIVC 预测模型；大规模计算与多模态 / 自监督 / 因果建模；数据平台与算力架构"],
  ["COO 首席运营官", "兰晓静", "运营与项目管理；财务与合规（人类遗传资源、数据安全）；商务拓展"],
];
s.addTable(rows, { x: X0, y: 4.65, w: 12.13, colW: [2.1, 1.8, 8.23],
  fontSize: 12, color: HEX.dk1, valign: "middle", rowH: 0.38, margin: [0, 0.1, 0, 0.1],
  border: { type: "solid", pt: 0.75, color: "D9E7E6" }, fill: { color: "FFFFFF" }, objectName: "roles-table" });

s.addText([
  { text: "团队即壁垒：", options: { bold: true, color: C.accent2 } },
  { text: "患者采样 → 类器官建立 → 数据构建 → 模型迭代 → 前瞻验证，全链路在团队内部闭环", options: { color: C.text1 } },
], { x: X0, y: 6.65, w: 12.13, h: 0.32, fontSize: 12, margin: 0, isTextBox: true });

s.addNotes("四位成员同属创始团队。卓悦博士的介绍来自简历（2026-08）与用户补充；田伟利、王中杰、兰晓静的介绍来自原 BP 与用户补充。");

(async () => {
  await pres.writeFile({ fileName: process.argv[2] });
  await applyTheme(process.argv[2], THEME);
  console.log("written", process.argv[2]);
})();
