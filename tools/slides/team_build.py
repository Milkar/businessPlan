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
def para(text,sz,color,b=False,algn='l',bullet=False,after=0):
    ppr=f'<a:pPr algn="{algn}"'+(' marL="171450" indent="-171450"' if bullet else '')+'>'
    ppr+='<a:lnSpc><a:spcPct val="115000"/></a:lnSpc>'+(f'<a:spcAft><a:spcPts val="{after}"/></a:spcAft>' if after else '')
    ppr+=('<a:buClr><a:srgbClr val="08A6F6"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="•"/>' if bullet else '<a:buNone/>')+'</a:pPr>'
    return f'<a:p>{ppr}<a:r>{rpr(sz,color,b)}<a:t>{escape(text)}</a:t></a:r></a:p>'
def photo(name,x,y,d):
    i=nextid()
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{i}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(d)}" cy="{emu(d)}"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="E7ECF2"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="99BDED"/></a:solidFill></a:ln></p:spPr>'
            f'<p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>{para("照片",1200,"8496B0",algn="ctr")}</p:txBody></p:sp>')
people=[
 ('卓悦 博士','CEO 首席执行官',['海德堡大学 / DKFZ 发育生物学博士','密歇根大学安娜堡分校学士','Verqura 实验室管理系统创始人','精通 iPSC 脑类器官构建、CRISPR 基因组编辑与标记','单细胞测序与 AI 图像分析（YOLO / Cellpose）','Cell Stem Cell 2025、npj Precision Oncology 2024 作者']),
 ('田伟利 博士','CSO 首席科学官',['DKFZ / 海德堡大学分子神经遗传学资深科学家','霍英东杰出青年学者','主持欧中重大科研项目','深耕神经调控与脑类器官方向，长期研究胶质母细胞瘤复发与耐药','Cell Stem Cell 2025、npj Precision Oncology 2024 核心成员']),
 ('王中杰 博士','CTO 首席技术官',['慕尼黑大学博士','曾任职哈佛医学院','近 10 年生物信息与多组学数据分析经验','精通大规模计算与机器学习预测建模','擅长搭建端到端可复用分析流程，支撑临床与生物医药项目的数据挖掘与决策输出']),
 ('兰晓静','COO 首席运营官',['海德堡大学经济学硕士','浙江大学管理学学士（竺可桢学院创业与创新管理强化班）','FRM 持证人','多年国际金融机构风控合规与跨境尽调经验','精通跨境财务统筹、合规体系搭建与硬科技项目运营']),
]
X0=0.58; TW=12.17; G=0.3; W=(TW-3*G)/4; D=1.55
shapes=[]
for k,(nm,role,bul) in enumerate(people):
    x=X0+k*(W+G)
    shapes.append(photo(f'照片 {k+1}',x+(W-D)/2,1.2,D))
    shapes.append(textbox(f'姓名 {k+1}',x,2.9,W,0.38,[para(nm,1600,'061F32',True,'ctr')],'ctr'))
    shapes.append(textbox(f'职位 {k+1}',x,3.28,W,0.32,[para(role,1300,'08A6F6',True,'ctr')],'ctr'))
    shapes.append(textbox(f'经历 {k+1}',x+0.1,3.8,W-0.15,3.05,[para(t,1100,'44546A',bullet=True,after=400) for t in bul]))
i=s.index('name="文本框 77"'); a=s.rfind('<p:sp>',0,i)
s=s[:a]+''.join(shapes)+s[a:]
open(p,'w',encoding='utf-8').write(s)
