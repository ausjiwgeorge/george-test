# -*- coding: utf-8 -*-
"""
生成面向 CFO 的演示 PPT: Finance & Controlling 转型建议 (10-15 分钟)
运行:  python3 build_ppt.py
输出:  CFO_Finance_Controlling_Transformation.pptx  (含每页演讲备注)
依赖:  python-pptx  (pip install python-pptx)
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ---------- 主题配色 ----------
NAVY   = RGBColor(0x0B, 0x2A, 0x4A)   # 深蓝(标题/主色)
BLUE   = RGBColor(0x14, 0x6C, 0xB4)   # 亮蓝(强调)
TEAL   = RGBColor(0x1F, 0x9E, 0x8F)   # 青绿(正向)
RED    = RGBColor(0xC0, 0x39, 0x2B)   # 红(痛点/风险)
GREY   = RGBColor(0x44, 0x4A, 0x52)   # 正文灰
LIGHT  = RGBColor(0xEC, 0xF1, 0xF6)   # 浅底
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
GOLD   = RGBColor(0xE1, 0xA5, 0x2B)

FONT = "微软雅黑"      # 中文字体, 无则由 PowerPoint 自动替换
FONT_EN = "Calibri"

prs = Presentation()
prs.slide_width  = Inches(13.333)   # 16:9
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
SW, SH = prs.slide_width, prs.slide_height


# ---------- 工具函数 ----------
def add_slide():
    return prs.slides.add_slide(BLANK)

def rect(slide, x, y, w, h, fill, line=None):
    from pptx.enum.shapes import MSO_SHAPE
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(1)
    sp.shadow.inherit = False
    return sp

def txt(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        space_after=6, line_spacing=1.05):
    """runs: list of paragraphs; each para = list of (text,size,color,bold) tuples"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    tf.vertical_anchor = anchor
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_after = Pt(space_after)
        p.line_spacing = line_spacing
        for (t, sz, col, bold) in para:
            r = p.add_run(); r.text = t
            r.font.size = Pt(sz); r.font.color.rgb = col; r.font.bold = bold
            r.font.name = FONT
    return tb

def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text

def header(slide, kicker, title, idx=None):
    """通用内容页页眉"""
    rect(slide, 0, 0, SW, Inches(1.15), NAVY)
    rect(slide, 0, Inches(1.15), SW, Pt(4), GOLD)
    txt(slide, Inches(0.6), Inches(0.12), Inches(11.5), Inches(0.35),
        [[(kicker, 12, GOLD, True)]])
    txt(slide, Inches(0.6), Inches(0.42), Inches(11.9), Inches(0.7),
        [[(title, 26, WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE)
    if idx:
        txt(slide, Inches(12.3), Inches(0.12), Inches(0.8), Inches(0.35),
            [[(idx, 12, GOLD, True)]], align=PP_ALIGN.RIGHT)

def bullets(slide, x, y, w, h, items, size=15, gap=8):
    """items: list of (text, level, color) ; level 0 顶级, 1 次级"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    for i, (t, lvl, col) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap); p.line_spacing = 1.05
        bullet = "▪  " if lvl == 0 else "–  "
        r = p.add_run(); r.text = ("" if lvl == 0 else "      ") + bullet + t
        r.font.size = Pt(size if lvl == 0 else size-1)
        r.font.color.rgb = col; r.font.bold = (lvl == 0)
        r.font.name = FONT
    return tb


# =====================================================================
# 1) 封面
# =====================================================================
s = add_slide()
rect(s, 0, 0, SW, SH, NAVY)
rect(s, 0, Inches(4.55), SW, Pt(4), GOLD)
txt(s, Inches(0.9), Inches(1.5), Inches(11.5), Inches(0.5),
    [[("Controller · Business Intelligence & Reporting", 16, GOLD, True)]])
txt(s, Inches(0.9), Inches(2.2), Inches(11.6), Inches(2.2),
    [[("从「报表工厂」到「业务伙伴」", 44, WHITE, True)],
     [("Finance & Controlling 转型建议", 30, RGBColor(0xC9,0xDA,0xEA), True)]],
    line_spacing=1.1)
txt(s, Inches(0.9), Inches(4.8), Inches(11.6), Inches(1.2),
    [[("核心主张:  让财务少花时间做报表,多花时间帮您做决策 —— ", 16, WHITE, False),
      ("更快、更准、更前瞻,且低成本、对齐集团。", 16, GOLD, True)]],
    line_spacing=1.2)
txt(s, Inches(0.9), Inches(6.5), Inches(11.6), Inches(0.5),
    [[("汇报对象: CFO    |    时长: 10–15 分钟", 13, RGBColor(0x9F,0xB4,0xC9), False)]])
notes(s, "开场白: 感谢给我这个机会。今天我用 12 分钟讲三件事——我怎么看现状、我会先做什么、以及一条低成本、对齐集团的落地路线。我的目标只有一句话: 让财务少做报表、多帮您做决策。")

# =====================================================================
# 2) 现状诊断
# =====================================================================
s = add_slide()
header(s, "1 · 现状诊断  DIAGNOSIS", "问题不在「人不努力」,而在「没有平台和标准」", "1")
cols = [
    ("月报第4天才出,月结太慢", "根因: SAP导出→Excel手工加工→手工合并,无自动化管道"),
    ("7个实体口径不一致", "根因: 缺统一KPI/科目定义与主数据治理,无「单一事实来源」"),
    ("每个追问都要重来一遍", "根因: 报告是静态Excel,无法自助下钻"),
    ("预测/计划与实际对不上", "根因: 计划、预测、实际分散多张表,无统一数据模型"),
]
x0 = Inches(0.6); y0 = Inches(1.5); cw = Inches(6.0); ch = Inches(1.25); gapx=Inches(0.25); gapy=Inches(0.2)
for i,(head,root) in enumerate(cols):
    cx = x0 + (cw+gapx) * (i%2)
    cy = y0 + (ch+gapy) * (i//2)
    rect(s, cx, cy, cw, ch, LIGHT)
    rect(s, cx, cy, Pt(6), ch, RED)
    txt(s, cx+Inches(0.2), cy+Inches(0.12), cw-Inches(0.35), Inches(0.5),
        [[("⚠ " + head, 15, RED, True)]])
    txt(s, cx+Inches(0.2), cy+Inches(0.58), cw-Inches(0.35), Inches(0.6),
        [[(root, 12.5, GREY, False)]], line_spacing=1.05)
txt(s, Inches(0.6), Inches(4.55), Inches(12), Inches(0.6),
    [[("一句话总结: ", 15, NAVY, True),
      ("我们现在——慢、口径乱、被动、且向后看。", 15, GREY, False)]])
notes(s, "先说结论: 我把管理层的四个抱怨,翻译成了四个根因。注意——右边红字都是'系统与标准'问题,不是'人不努力'。尤其'口径不一致'本质是人和流程问题: 每个实体对同一个指标的定义不一样。抓住根因,后面的路线图才对症。")

# =====================================================================
# 3) 五条原则
# =====================================================================
s = add_slide()
header(s, "2 · 方法论  PRINCIPLES", "五条贯穿始终的原则", "2")
principles = [
    ("① 单一事实来源", "SAP 是唯一系统记录,所有报告从同一数据模型出"),
    ("② 先定义,后自动化", "KPI口径先统一,否则自动化的是错误"),
    ("③ 人做分析,机器做搬运", "把手工合并交给系统,人专注洞察"),
    ("④ 自助式", "管理层自己点开下钻,追问不再产生手工活"),
    ("⑤ 务实小步快跑", "先用现有许可证做速赢,用成果换预算"),
]
y = Inches(1.55)
for head, desc in principles:
    rect(s, Inches(0.6), y, Inches(3.6), Inches(0.82), NAVY)
    txt(s, Inches(0.75), y+Inches(0.14), Inches(3.4), Inches(0.6),
        [[(head, 16, WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE)
    rect(s, Inches(4.35), y, Inches(8.35), Inches(0.82), LIGHT)
    txt(s, Inches(4.55), y+Inches(0.14), Inches(8.0), Inches(0.6),
        [[(desc, 14, GREY, False)]], anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.98)
notes(s, "这五条是我的工作方法招牌。最想强调第②条: 先定义后自动化——如果口径还没统一就上工具,只会更快地生产不一致的数字。第⑤条是给您的定心丸: 我不会一上来就要大预算大项目,而是先用公司已有的工具做出成果,再用成果申请下一步投入。")

# =====================================================================
# 4) 愿景
# =====================================================================
s = add_slide()
header(s, "3 · 愿景  VISION", "从「报表工厂」到「业务决策伙伴」", "3")
rect(s, Inches(0.8), Inches(2.2), Inches(5.0), Inches(3.0), LIGHT)
rect(s, Inches(0.8), Inches(2.2), Inches(5.0), Inches(0.7), RED)
txt(s, Inches(0.8), Inches(2.3), Inches(5.0), Inches(0.55),
    [[("今天 · 报表工厂", 18, WHITE, True)]], align=PP_ALIGN.CENTER)
bullets(s, Inches(1.1), Inches(3.1), Inches(4.5), Inches(2.0),
    [("首版报告 第4天", 0, GREY), ("大量手工合并", 0, GREY),
     ("口径不一 / 被动", 0, GREY), ("向后看", 0, GREY)], size=15, gap=10)

# 箭头
txt(s, Inches(5.9), Inches(3.3), Inches(1.5), Inches(1.0),
    [[("➜", 54, BLUE, True)]], align=PP_ALIGN.CENTER)

rect(s, Inches(7.5), Inches(2.2), Inches(5.0), Inches(3.0), LIGHT)
rect(s, Inches(7.5), Inches(2.2), Inches(5.0), Inches(0.7), TEAL)
txt(s, Inches(7.5), Inches(2.3), Inches(5.0), Inches(0.55),
    [[("目标 · 业务伙伴", 18, WHITE, True)]], align=PP_ALIGN.CENTER)
bullets(s, Inches(7.8), Inches(3.1), Inches(4.5), Inches(2.0),
    [("准实时 · 单一真相", 0, NAVY), ("自动化管道", 0, NAVY),
     ("自助下钻 / 主动", 0, NAVY), ("前瞻 · 滚动预测", 0, NAVY)], size=15, gap=10)
notes(s, "一张图说明方向: 左边是现在,右边是目标。转型不是买个工具,而是把财务的角色从'搬数字'变成'讲业务、看未来'。这也回应了管理层要的三个词: 更快、更可靠、更前瞻。")

# =====================================================================
# 5) 优先级
# =====================================================================
s = add_slide()
header(s, "4 · 优先级  PRIORITIES", "先做什么,为什么", "4")
items = [
    ("最先做:统一 KPI/科目定义 (数据字典)", 0, NAVY),
    ("零成本、最高价值;是自动化与实时报告的地基。地基不统一,自动化只会更快出错。", 1, GREY),
    ("紧接着:自动化现有合并 (Power Query)", 0, NAVY),
    ("立刻见效——首版报告 第4天 → 第2天,不需IT、不花钱。", 1, GREY),
    ("然后:中央数据模型 = 单一事实来源", 0, NAVY),
    ("各实体本地科目映射到集团科目,消灭'谁的表覆盖谁'。", 1, GREY),
    ("最后:整合计划与预测 + 前瞻分析", 0, NAVY),
    ("实际/预算/预测同一视图;滚动预测、情景模拟。", 1, GREY),
]
bullets(s, Inches(0.7), Inches(1.6), Inches(12), Inches(5), items, size=16, gap=9)
notes(s, "排序逻辑: 按'价值高 + 见效快 + 依赖少'来排。第一件事是统一定义,因为它零成本却是一切的地基;第二件是自动化现有 Excel,90 天内就能让您看到报告变快。先解决'快和准',再谈'前瞻'。")

# =====================================================================
# 6) 路线图 (表格)
# =====================================================================
s = add_slide()
header(s, "5 · 落地路线图  ROADMAP", "稳住 → 标准化 → 自动化 → 智能化 (四阶段)", "5")
rows = [
    ("阶段", "关键动作", "可交付 / 里程碑"),
    ("0 诊断 (~30天)", "画流程图、盘点报告与许可证、锁定核心KPI", "《现状评估+速赢清单》"),
    ("1 稳住 (1–3月)", "建KPI数据字典+7实体签字;Power Query自动合并", "首版报告 第4天→第2天"),
    ("2 标准化 (3–6月)", "中央数据模型;本地科目→集团科目映射;主数据治理", "单一事实来源+自助下钻"),
    ("3 自动化 (6–12月)", "全链路自动;每日/准实时;整合计划与预测", "首版→第1天/准实时"),
    ("4 智能化 (12月+)", "滚动预测、驱动因子、情景模拟、异常预警", "前瞻+预测性分析"),
]
tbl_shape = s.shapes.add_table(len(rows), 3, Inches(0.6), Inches(1.5),
                               Inches(12.1), Inches(4.9))
table = tbl_shape.table
table.columns[0].width = Inches(2.5)
table.columns[1].width = Inches(6.1)
table.columns[2].width = Inches(3.5)
for r, row in enumerate(rows):
    table.rows[r].height = Inches(0.8)
    for c, val in enumerate(row):
        cell = table.cell(r, c)
        cell.margin_left = Inches(0.12); cell.margin_right = Inches(0.1)
        cell.margin_top = Inches(0.05); cell.margin_bottom = Inches(0.05)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; run = p.add_run(); run.text = val
        run.font.name = FONT
        if r == 0:
            cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
            run.font.color.rgb = WHITE; run.font.bold = True; run.font.size = Pt(14)
        else:
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT if r % 2 else WHITE
            run.font.color.rgb = GREY; run.font.size = Pt(12.5)
            if c == 0: run.font.bold = True; run.font.color.rgb = BLUE
notes(s, "这是整份提案的骨架。四个阶段,每个都有能量化的里程碑——重点看'首版报告出具时间'这条主线: 第4天→第2天→第1天→准实时。我不会一次性推倒重来,而是每阶段交付看得见的成果。")

# =====================================================================
# 7) 平台架构
# =====================================================================
s = add_slide()
header(s, "6 · 智能共享平台  PLATFORM", "Intelligent Shared Platform —— 分层架构", "6")
layers = [
    ("消费层", "BI仪表盘 · 自助分析 · 计划与预测 · 高级分析/AI", BLUE),
    ("语义 / KPI 层", "统一KPI定义 (数据字典落进模型)", TEAL),
    ("中央数据层", "数据仓库 = 单一事实来源 (统一主数据 · 集团科目映射)", NAVY),
    ("数据集成层", "自动抽取/转换 ELT (Power Query→dataflows→SAP Datasphere/BW)", RGBColor(0x5A,0x6B,0x7B)),
    ("源系统层", "SAP (ERP · 唯一系统记录) + 其他源", GREY),
]
y = Inches(1.5)
for name, desc, col in layers:
    rect(s, Inches(2.2), y, Inches(3.0), Inches(0.82), col)
    txt(s, Inches(2.3), y+Inches(0.13), Inches(2.8), Inches(0.6),
        [[(name, 15, WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    rect(s, Inches(5.4), y, Inches(6.9), Inches(0.82), LIGHT)
    txt(s, Inches(5.6), y+Inches(0.13), Inches(6.6), Inches(0.6),
        [[(desc, 13, GREY, False)]], anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.95)
# 治理竖条
rect(s, Inches(0.6), Inches(1.5), Inches(1.4), Inches(4.55), GOLD)
txt(s, Inches(0.62), Inches(1.5), Inches(1.36), Inches(4.55),
    [[("治理贯穿", 15, WHITE, True)],
     [("数据所有权", 11, WHITE, False)],
     [("KPI标准", 11, WHITE, False)],
     [("安全权限", 11, WHITE, False)],
     [("变更管理", 11, WHITE, False)]],
    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, space_after=8)
notes(s, "平台一层层建: 底层 SAP 是唯一记录,往上是自动集成、中央仓库(单一真相)、KPI语义层,最上面才是看板和预测。左侧金色竖条是贯穿所有层的治理——所有权、标准、权限、变更。路线图就是从下往上、一层层把它建起来。")

# =====================================================================
# 8) 案例A: 口径统一 (before/after 表)
# =====================================================================
s = add_slide()
header(s, "7 · 案例A  ONE TRUTH", "口径统一实例:同一「净销售额」,7个实体不再各算各的", "7")
txt(s, Inches(0.6), Inches(1.35), Inches(12), Inches(0.5),
    [[("做法: KPI数据字典锁死4个维度(含税/加减项/科目映射/确认时点)+7实体书面签字+写进模型", 12.5, NAVY, True)]])
rows = [
    ("实体", "原上报(万)", "统一口径后(万)", "差异原因"),
    ("医疗-华东", "1000", "940", "补扣现金折扣、退货"),
    ("医疗-华北", "950", "910", "剔除运费收入"),
    ("生产实体", "1050", "900", "剔除增值税(含税拆分)"),
    ("服务实体", "900", "930", "补入漏计的服务收入"),
    ("贸易实体", "980", "940", "发货口径→开票口径"),
]
tbl = s.shapes.add_table(len(rows), 4, Inches(0.6), Inches(1.9),
                         Inches(12.1), Inches(3.6)).table
tbl.columns[0].width=Inches(2.6); tbl.columns[1].width=Inches(2.6)
tbl.columns[2].width=Inches(3.0); tbl.columns[3].width=Inches(3.9)
for r,row in enumerate(rows):
    for c,val in enumerate(row):
        cell=tbl.cell(r,c); cell.vertical_anchor=MSO_ANCHOR.MIDDLE
        cell.margin_left=Inches(0.12)
        p=cell.text_frame.paragraphs[0]; run=p.add_run(); run.text=val
        run.font.name=FONT
        if r==0:
            cell.fill.solid(); cell.fill.fore_color.rgb=NAVY
            run.font.color.rgb=WHITE; run.font.bold=True; run.font.size=Pt(13)
        else:
            cell.fill.solid(); cell.fill.fore_color.rgb=LIGHT if r%2 else WHITE
            run.font.size=Pt(12.5)
            run.font.color.rgb = TEAL if c==2 else GREY
            if c==2: run.font.bold=True
txt(s, Inches(0.6), Inches(5.7), Inches(12), Inches(1.0),
    [[("价值: ", 14, NAVY, True),
      ("差异从「算法差异」变成「业务差异」——追问从'这数对不对'变成'这生意为什么好',不再手工核对。", 14, GREY, False)]],
    line_spacing=1.1)
notes(s, "这是最有说服力的一页。同一个'净销售额',七个实体原来算出七个数,差几百万。数据字典把四个维度锁死、让各实体签字、并写进模型后重算——右边绿色是统一口径结果。关键价值: 现在实体间的差异都是真实的业务差异,您的追问不会再触发一轮手工对账。")

# =====================================================================
# 9) 案例B: 映射拆分管道
# =====================================================================
s = add_slide()
header(s, "8 · 案例B  PIPELINE", "映射拆分管道:脏数据进 → 干净表出 (配置表驱动)", "8")
# 三层流程
steps = [
    ("① Staging 贴源", "SAP原样导出\n(含税/运费混入/退货为正)", RED),
    ("② 转换层", "JOIN配置表→拆税÷1.13\n拆服务·剔运费·打加减号", BLUE),
    ("③ 干净事实表", "长格式\n集团科目+EUR", TEAL),
]
x = Inches(0.6)
for i,(h,d,col) in enumerate(steps):
    rect(s, x, Inches(1.6), Inches(3.7), Inches(1.7), col)
    txt(s, x+Inches(0.2), Inches(1.75), Inches(3.3), Inches(0.5),
        [[(h, 16, WHITE, True)]])
    txt(s, x+Inches(0.2), Inches(2.3), Inches(3.3), Inches(0.9),
        [[(line, 12.5, WHITE, False)] for line in d.split("\n")], line_spacing=1.05)
    if i<2:
        txt(s, x+Inches(3.75), Inches(2.0), Inches(0.5), Inches(0.8),
            [[("➜", 30, NAVY, True)]], align=PP_ALIGN.CENTER)
    x += Inches(4.15)
# E4 结果
txt(s, Inches(0.6), Inches(3.6), Inches(12), Inches(0.45),
    [[("E4 生产实体 (含税13%+运费混入) 结果:", 14, NAVY, True)]])
res = [
    ("原始直接加总(错)", "11,300,000 + 113,000 = 11,413,000", RED),
    ("统一口径净销售(对)", "10,000,000 − 1,000,000 − 200,000 − 50,000 = 8,750,000 CNY", TEAL),
    ("内置校验", "含税还原✓  总额守恒✓  未映射告警✓", NAVY),
]
y=Inches(4.05)
for h,d,col in res:
    rect(s, Inches(0.6), y, Inches(3.2), Inches(0.62), col)
    txt(s, Inches(0.72), y+Inches(0.11), Inches(3.0), Inches(0.45),
        [[(h, 12.5, WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE)
    rect(s, Inches(3.9), y, Inches(8.8), Inches(0.62), LIGHT)
    txt(s, Inches(4.05), y+Inches(0.11), Inches(8.5), Inches(0.45),
        [[(d, 12.5, GREY, False)]], anchor=MSO_ANCHOR.MIDDLE)
    y+=Inches(0.72)
txt(s, Inches(0.6), Inches(6.4), Inches(12), Inches(0.5),
    [[("规则放配置表、逻辑放管道:改口径只改一行、全实体复用、每个数字可审计追溯。", 13, NAVY, True)]])
notes(s, "上一页讲'为什么',这页讲'怎么做到'。三层管道: 脏数据进,自动拆税、拆服务、剔运费、打加减号,吐出一张干净的长表。以 E4 为例: 直接加总会虚高 260 万,统一口径后是 875 万。关键在'配置表驱动'——业务改口径只改表里一行,不用找 IT,而且每个数字都能追回 SAP 凭证。这套逻辑我已经写成可运行的 SQL 和 Excel 脚本。")

# =====================================================================
# 10) 实时报告
# =====================================================================
s = add_slide()
header(s, "9 · 实时报告  REAL-TIME", "如何实现「实时」报告", "9")
txt(s, Inches(0.6), Inches(1.4), Inches(12), Inches(0.7),
    [[("成熟观点: ", 15, GOLD, True),
      ("管理决策要的是「right-time」(对的时间)——多数是每日/准实时,而非逐秒流式。把成本花在刀刃上。", 15, GREY, False)]],
    line_spacing=1.15)
items = [
    ("自动定时抽取 + 增量刷新,替代人工导出 (先做到每日刷新)", 0, NAVY),
    ("预建数据模型,报告直接读模型:秒级打开、可下钻", 0, NAVY),
    ("关键指标(现金/订单)用 DirectQuery / 近实时连接单独处理", 0, NAVY),
    ("SAP取数方式(由轻到重):", 0, NAVY),
    ("OData / CDS View  →  连接器抽表  →  复用集团 SAP BW / SAP Analytics Cloud", 1, GREY),
]
bullets(s, Inches(0.7), Inches(2.5), Inches(12), Inches(4), items, size=16, gap=12)
notes(s, "回答'如何实时'这个问题,先给您一个成熟判断: 别追求逐秒流式,那贵且没必要。管理报告要的是'对的时间'——大多数每日刷新就够。做法是: 自动抽取代替手工导出、预建模型让报告秒开可下钻;只有现金、订单这类关键指标才单独做近实时。取数方式按集团现有资产由轻到重选。")

# =====================================================================
# 11) 预算/IT/集团
# =====================================================================
s = add_slide()
header(s, "10 · 约束下落地  CONSTRAINTS", "有限预算 + 有限IT + 对齐集团,怎么做到", "10")
items = [
    ("先用「已经付过钱」的工具", 0, NAVY),
    ("Microsoft 365 的 Power Query / Power BI / Power Automate,几乎零增量成本跑完阶段1–2", 1, GREY),
    ("零成本高价值的先做", 0, NAVY),
    ("KPI数据字典、流程标准化不花钱,却解决口径乱", 1, GREY),
    ("低代码、少定制", 0, NAVY),
    ("减少对IT依赖:业务主导 + IT把关(治理护栏)", 1, GREY),
    ("复用集团资产", 0, NAVY),
    ("对接集团统一科目表 / SAP BW / SAC——既省钱又天然合规", 1, GREY),
    ("用成果换预算", 0, NAVY),
    ("先交付90天速赢,用'第4天→第2天、省X工时'去申请下一阶段投入", 1, GREY),
]
bullets(s, Inches(0.7), Inches(1.55), Inches(12), Inches(5.5), items, size=15, gap=7)
notes(s, "这页专门打消您对'又要花大钱'的顾虑。核心思路: 先用公司已经买过的微软工具,先做不花钱但高价值的标准化,尽量低代码少依赖 IT,能复用集团的就复用。最重要的一条——用 90 天的成果去换下一步的预算,让投入始终跟着价值走。")

# =====================================================================
# 12) 角色演进
# =====================================================================
s = add_slide()
header(s, "11 · 持续进化  GROWTH", "角色在演进,我如何持续跟上", "11")
items = [
    ("建立「财务 + 数据」的 T 型能力", 0, NAVY),
    ("精进数据建模、BI、分析(Power BI / SQL / SAC 认证)", 1, GREY),
    ("关注趋势", 0, NAVY),
    ("FP&A自动化、分析、AI在财务的应用;加入集团Controlling网络与外部社区", 1, GREY),
    ("建立长期伙伴关系", 0, NAVY),
    ("与IT、业务并肩,而非单打独斗", 1, GREY),
    ("保持「产品思维」", 0, NAVY),
    ("把报告当产品持续迭代,而非一次性交付", 1, GREY),
]
bullets(s, Inches(0.7), Inches(1.7), Inches(12), Inches(5), items, size=16, gap=10)
notes(s, "这个岗位在快速演进,我的应对是: 持续把自己打造成'既懂财务又懂数据'的 T 型人才,紧跟 FP&A 自动化和 AI 趋势,和 IT、业务建立长期伙伴关系,并且用产品思维持续迭代报告——它永远是 v1、v2、v3,不是交付一次就结束。")

# =====================================================================
# 13) 前90天 + 行动请求
# =====================================================================
s = add_slide()
header(s, "12 · 前90天  QUICK WINS", "前90天承诺 + 行动请求", "12")
rows = [
    ("30 天", "现状诊断 + 统一KPI数据字典(各实体确认)"),
    ("60 天", "自动化月度合并,首版报告 第4天 → 第2天"),
    ("90 天", "上线第一个自助管理仪表盘(单一版本真相)"),
]
y=Inches(1.6)
for d,t in rows:
    rect(s, Inches(0.7), y, Inches(1.8), Inches(0.95), TEAL)
    txt(s, Inches(0.7), y+Inches(0.22), Inches(1.8), Inches(0.5),
        [[(d, 20, WHITE, True)]], align=PP_ALIGN.CENTER)
    rect(s, Inches(2.7), y, Inches(10.0), Inches(0.95), LIGHT)
    txt(s, Inches(2.9), y+Inches(0.22), Inches(9.6), Inches(0.5),
        [[(t, 15, GREY, False)]], anchor=MSO_ANCHOR.MIDDLE)
    y+=Inches(1.1)
rect(s, Inches(0.7), Inches(5.15), Inches(12.0), Inches(1.5), NAVY)
txt(s, Inches(0.95), Inches(5.35), Inches(11.6), Inches(1.1),
    [[("行动请求", 16, GOLD, True)],
     [("请授权我牵头 KPI 口径对齐,并指定各实体一名数据对接人。", 18, WHITE, True)]],
    line_spacing=1.2, space_after=8)
notes(s, "最后给您一个低风险、可验证的承诺: 30天、60天、90天各交付什么,清清楚楚。我今天只有一个请求——请授权我牵头 KPI 口径对齐,并让每个实体指定一名数据对接人。这一步不花钱,却能启动整个转型。谢谢您,我很期待加入团队。")

# =====================================================================
# 14) 结束
# =====================================================================
s = add_slide()
rect(s, 0, 0, SW, SH, NAVY)
rect(s, Inches(0.9), Inches(3.35), Inches(4.0), Pt(4), GOLD)
txt(s, Inches(0.9), Inches(2.4), Inches(11.6), Inches(1.0),
    [[("让财务少做报表,多帮您做决策。", 34, WHITE, True)]])
txt(s, Inches(0.9), Inches(3.7), Inches(11.6), Inches(1.2),
    [[("更快 · 更准 · 更前瞻    |    低成本 · 对齐集团", 20, GOLD, True)]])
txt(s, Inches(0.9), Inches(6.4), Inches(11.6), Inches(0.6),
    [[("谢谢  ·  期待您的问题", 16, RGBColor(0x9F,0xB4,0xC9), False)]])
notes(s, "收尾: 用一句话把主张再钉一遍——让财务少做报表、多帮您做决策。预留时间回答问题。准备好三个可能的追问: 数据质量怎么保证、各实体不配合怎么办、和集团项目会不会冲突。")

prs.save("CFO_Finance_Controlling_Transformation.pptx")
print("已生成: CFO_Finance_Controlling_Transformation.pptx  共", len(prs.slides.__iter__.__self__._sldIdLst), "页")
