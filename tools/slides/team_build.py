import re,sys
from xml.sax.saxutils import escape
p=sys.argv[1]; s=open(p,encoding='utf-8').read()
def span(tag,nm):
    i=s.index(f'name="{nm}"'); a=s.rfind(f'<p:{tag}>',0,i); b=s.index(f'</p:{tag}>',i)+len(f'</p:{tag}>'); return a,b
for tag,nm in [('sp','文本框 66'),('graphicFrame','表格 71'),('sp','任意多边形 72'),('sp','文本框 73')]:
    a,b=span(tag,nm); s=s[:a]+s[b:]
s=s.replace('<a:t>Source: AIVC 商业故事 BP §5</a:t>','<a:t>Source: 团队成员简历</a:t>')
E=914400; F='<a:latin typeface="微软雅黑"/><a:ea typeface="微软雅黑"/><a:cs typeface="微软雅黑"/>'
nid=[300]
def nextid(): nid[0]+=1; return nid[0]
def emu(v): return int(round(v*E))
def rpr(sz,color,b=False):
    bb=' b="1"' if b else ''
    return f'<a:rPr lang="zh-CN" altLang="en-US" sz="{sz}"{bb}><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>{F}</a:rPr>'
def textbox(name,x,y,w,h,paras,anchor='t'):
    i=nextid()
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="{anchor}"><a:noAutofit/></a:bodyPr><a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')
def para(text,sz,color,b=False,algn='l',bullet=False,after=0,bold_lead=None):
    ppr=f'<a:pPr algn="{algn}"'+(' marL="171450" indent="-171450"' if bullet else '')+'>'
    ppr+='<a:lnSpc><a:spcPct val="112000"/></a:lnSpc>'+(f'<a:spcAft><a:spcPts val="{after}"/></a:spcAft>' if after else '')
    ppr+=('<a:buClr><a:srgbClr val="08A6F6"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="•"/>' if bullet else '<a:buNone/>')+'</a:pPr>'
    runs=''
    if bold_lead: runs+=f'<a:r>{rpr(sz,"061F32",True)}<a:t>{escape(bold_lead)}</a:t></a:r>'
    if text: runs+=f'<a:r>{rpr(sz,color,b)}<a:t>{escape(text)}</a:t></a:r>'
    return f'<a:p>{ppr}{runs}</a:p>'
def photo(name,x,y,d):
    i=nextid()
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(d)}" cy="{emu(d)}"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="E7ECF2"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="99BDED"/></a:solidFill></a:ln></p:spPr>'
            f'<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>{para("照片",1200,"8496B0",algn="ctr")}</p:txBody></p:sp>')
people=[
 ('卓悦 博士','CEO 首席执行官','湿实验 × 数据 × AI，全链路打通',[
   ('Verqura 实验室管理系统创始人','：打通样本、实验与数据管理'),
   ('Cell Stem Cell 2025、npj Precision Oncology 2024 作者','（类器官预测患者治疗响应）'),
   ('精通 iPSC 脑类器官构建','，引进 Pașca 团队分化方案；CRISPR 基因组编辑'),
   ('单细胞 RNA / ATAC 测序分析','；训练 YOLO / Cellpose 图像 AI 模型'),
   ('海德堡大学 / DKFZ 发育生物学博士','；密歇根大学安娜堡分校学士；亥姆霍兹研究生院奖学金')]),
 ('田伟利 博士','CSO 首席科学官','神经科学 × 脑类器官 × 产业转化',[
   ('霍英东杰出青年学者','；中山大学香港高等研究所研究员'),
   ('DKFZ 资深科学家、课题负责人','；海德堡大学医院开展患者来源肿瘤类器官指导的个体化胶质母细胞瘤治疗研究'),
   ('第一作者成果 Science（修回中）、Cell（投稿中）','；Cell Stem Cell 2025 等 19 篇论文'),
   ('参与 ERC、DFG、HFSP、NIH BRAIN Initiative','及国家 973 / 863 等重大项目'),
   ('Neusician 首席神经科学官','；慕尼黑大学 / 亥姆霍兹中心神经生物学博士')]),
 ('王中杰 博士','CTO 首席技术官','生物信息 × 机器学习预测建模',[
   ('慕尼黑工业大学 / 慕尼黑亥姆霍兹中心生物信息学博士','；厦门大学硕士'),
   ('约 10 年生物信息与多组学分析经验','；曾任职哈佛医学院，负责临床试验数据与多组学预测建模'),
   ('预测模型 AUC > 0.95','（高维数据分类）；临床多组学预测 AUC > 0.80'),
   ('多模态 + 纵向数据整合','：测序、转录、代谢、表型与临床元数据'),
   ('HPC / SLURM 规模化分析流程','，单批覆盖数百至数千样本')]),
 ('兰晓静','COO 首席运营官','跨境金融风控 × 科创运营',[
   ('德国商业银行（Commerzbank）高级合规官','：统筹集团全球分支风控合规，向董事会提交风险研判'),
   ('曾任 Sopra Steria 合规官','、中国银行法兰克福分行反洗钱分析师'),
   ('歌德商学院高管教育中心高级项目经理','：搭建中德企业家与科创创始人交流平台'),
   ('FRM、ESG Investing 认证','；CFA 一级；精通英语、德语'),
   ('海德堡大学经济学硕士','；浙江大学管理学学士（竺可桢学院创新与创业管理强化班）')]),
]
X0=0.58; TW=12.17; G=0.3; W=(TW-3*G)/4; D=1.25
shapes=[]
for k,(nm,role,tag,bul) in enumerate(people):
    x=X0+k*(W+G)
    shapes.append(photo(f'照片 {k+1}',x+(W-D)/2,1.12,D))
    shapes.append(textbox(f'姓名 {k+1}',x,2.45,W,0.36,[para(nm,1600,'061F32',True,'ctr')],'ctr'))
    shapes.append(textbox(f'职位 {k+1}',x,2.8,W,0.3,[para(role,1300,'08A6F6',True,'ctr')],'ctr'))
    shapes.append(textbox(f'亮点 {k+1}',x,3.12,W,0.28,[para(tag,1050,'44546A',False,'ctr')],'ctr'))
    shapes.append(textbox(f'经历 {k+1}',x+0.08,3.55,W-0.1,3.35,[para(t,1000,'44546A',bullet=True,after=300,bold_lead=bl) for bl,t in bul]))
i=s.index('name="文本框 77"'); a=s.rfind('<p:sp>',0,i)
s=s[:a]+''.join(shapes)+s[a:]
open(p,'w',encoding='utf-8').write(s)
