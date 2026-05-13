#!/usr/bin/env python3
"""风电/光伏项目评估报告 Word 生成脚本
使用方法：修改 REPORT_DATA 字典后直接运行
依赖：python-docx (pip install python-docx --break-system-packages)
Logo：~/.hermes/skills/openclaw-imports/wind-power-analysis/assets/jianeng_logo_header.png
"""
import os, sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ───────────────────────────────────────────────
# >>> 配置（每个项目修改此处）<<<
# ───────────────────────────────────────────────
LOGO = os.path.expanduser("~/.hermes/skills/openclaw-imports/wind-power-analysis/assets/jianeng_logo_header.png")
REPORT_DATA = {
    "project_name": "九江",       # 项目名称（如"九江"/"南昌"）
    "capacity_mw": 100,           # 装机容量 MW
    "annual_hours": 2150,         # 理论利用小时
    "curtailment_rate": 0.03,     # 限电率
    "mechanism_price": 0.365,     # 机制电价 含税
    "mechanism_ratio": 0.80,      # 机制电量比例
    "mechanism_years": 10,        # 执行期限
    "mkt_long_price": 0.38,       # 中长期电价 含税
    "mkt_spot_price": 0.32,       # 现货电价 含税
    "t1_limit": 5.39,             # 投资边界 元/W
    "t4_limit": 6.31,             # 出售边界 元/W
    "t1_irr": 7.83,               # 投资边界IRR %
    "t4_irr": 6.00,               # 出售边界IRR %
    "t4_equity_irr": 13.10,       # 资本金IRR %
    "t1_dscr": 1.2000,            # 投资边界DSCR
    "t4_dscr": 1.2929,            # 出售边界DSCR
    "t1_lcoe": 0.2761,            # LCOE 元/kWh
    "t4_lcoe": 0.3084,
    "curtailment_source": "冬雪说《江西省风电限电率分析》2026-05-08",
    "curtailment_detail": "赣北1%-3%，取上限保守",
}
TODAY = "20260511"
# ───────────────────────────────────────────────

COLOR_DB = RGBColor(0x1B, 0x3A, 0x5C)

def h1(doc, text):
    h = doc.add_heading(text, level=1)
    for r in h.runs:
        r.font.color.rgb = COLOR_DB; r.font.name = '黑体'
        r._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体'); r.font.size = Pt(15)

def h2(doc, text):
    h = doc.add_heading(text, level=2)
    for r in h.runs:
        r.font.color.rgb = COLOR_DB; r.font.name = '黑体'
        r._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体'); r.font.size = Pt(13)

def para(doc, text, bold=False, sz=11, align=None):
    p = doc.add_paragraph()
    if align: p.alignment = align
    r = p.add_run(text)
    r.font.name = '仿宋'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')
    r.font.size = Pt(sz); r.bold = bold

def table(doc, headers, rows):
    t = doc.add_table(rows=1+len(rows), cols=len(headers))
    t.style = 'Table Grid'
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]; c.text = h
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.bold = True; r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
                r.font.name = '黑体'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="1B3A5C" w:val="clear"/>')
        c._element.get_or_add_tcPr().append(shd)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            c = t.rows[ri+1].cells[ci]; c.text = str(val)
            for p in c.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.size = Pt(9); r.font.name = '仿宋'
                    r._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')

def add_header(doc, logo_path):
    """江能 Logo 页眉，2cm 宽居右"""
    for s in doc.sections:
        hdr = s.header
        hdr.is_linked_to_previous = False
        p = hdr.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run()
        run.add_picture(logo_path, width=Cm(2))

def add_disclaimer(doc):
    """免责声明 — 必须包含"""
    h1(doc, "免责声明")
    text = (
        "本报告由 Yvonne（基于 DeepSeek Pro 大语言模型）辅助生成，仅供项目投资决策参考，不构成任何形式的投资建议或承诺。"
        "报告中所引用的电力市场数据（机制电价、中长期交易电价、现货电价、限电率、利用小时数等）均来自公开渠道，"
        "部分电价参数在缺乏直接交易数据的情况下采用保守假设推算，可能与实际成交价格存在偏差。"
        "实际项目投资决策应结合以下因素综合判断：\n\n"
        "① 场址实测测风数据（至少一个完整年度）；"
        "② 电网接入批复及送出工程条件；"
        "③ EPC 招标实际报价；"
        "④ 项目所在地最新的机制电价竞价结果；"
        "⑤ 所在省电力交易中心公布的全年现货与中长期交易结算数据。\n\n"
        "本报告中的财务模型基于特定假设（融资利率 4%、经营期 20 年、等额本金还款等），"
        "不同融资结构、利率环境及政策变化可能导致测算结果显著偏离。"
        "报告中的"投资边界"和"出售边界"为理论测算阈值，不代表项目实际可实现的交易价格或融资条件。\n\n"
        "江能新能源及报告编制方不对因使用本报告而产生的任何直接或间接损失承担责任。"
        "未经授权，不得转载或用于商业用途。"
    )
    para(doc, text, sz=9)

def build(doc, R):
    """构建报告正文"""
    # 标题
    t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(f"江能新能源 · {R['project_name']}100MW风电项目\n投资评估报告")
    r.font.size = Pt(16); r.font.bold = True; r.font.color.rgb = COLOR_DB
    r.font.name = '黑体'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

    st = doc.add_paragraph(); st.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = st.add_run(f"Yvonne (DeepSeek Pro) · {TODAY}")
    r.font.size = Pt(9); r.font.color.rgb = RGBColor(0x88,0x88,0x88)
    r.font.name = '仿宋'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋')

    gen_mwh = R['capacity_mw'] * R['annual_hours'] * (1 - R['curtailment_rate'])

    # 一、项目概况
    h1(doc, "一、项目概况")
    table(doc, ["参数", "数值", "来源"], [
        ["装机容量", f"{R['capacity_mw']} MW", "用户提供"],
        ["理论利用小时", f"{R['annual_hours']} h", "龙源/大唐2025年数据+补偿"],
        ["限电率", f"{R['curtailment_rate']*100:.1f}%", R['curtailment_detail']],
        ["年净发电量", f"{gen_mwh:,.0f} MWh", f"{R['capacity_mw']}MW×{R['annual_hours']}h×97%"],
    ])

    # 二、电价参数
    h1(doc, "二、电价参数")
    table(doc, ["参数", "数值", "来源"], [
        ["竞价上限", "0.38", "赣发改价管〔2025〕718号"],
        ["实际中标价(风电)", f"{R['mechanism_price']}", "国网江西2026第二轮竞价"],
        ["机制电量比例", f"{R['mechanism_ratio']*100:.0f}%", "江西省竞价实施细则第11条"],
        ["执行期限", f"{R['mechanism_years']}年", "江西省实施细则"],
        ["现货均价(含税)", f"{R['mkt_spot_price']}", "消纳友好省保守估计"],
        ["中长期均价(含税)", f"{R['mkt_long_price']}", "煤电基准价折让"],
    ])

    # 三、测算结果
    h1(doc, "三、测算结果")
    h2(doc, "投资边界（100%融资，DSCR≥1.2）")
    table(doc, ["指标", "数值"], [
        ["最高单瓦投资", f"{R['t1_limit']}元/W"],
        ["总投资", f"{R['t1_limit']*10:.2f}亿元"],
        ["全投资IRR", f"{R['t1_irr']}%"],
        ["最小DSCR", f"{R['t1_dscr']:.4f} ✓"],
        ["LCOE", f"{R['t1_lcoe']:.4f}元/kWh"],
    ])

    h2(doc, "出售边界（80%融资，IRR≥6% + 资本金IRR≥8%）")
    table(doc, ["指标", "数值"], [
        ["目标单瓦投资", f"{R['t4_limit']}元/W"],
        ["总投资", f"{R['t4_limit']*10:.2f}亿元"],
        ["全投资IRR", f"{R['t4_irr']}% ✓"],
        ["税后资本金IRR", f"{R['t4_equity_irr']}% ✓"],
        ["最小DSCR", f"{R['t4_dscr']:.4f}"],
        ["LCOE", f"{R['t4_lcoe']:.4f}元/kWh"],
    ])

    # 免责声明
    add_disclaimer(doc)

def generate(R):
    doc = Document()
    for s in doc.sections:
        s.top_margin = Cm(2.5); s.bottom_margin = Cm(2.2)
        s.left_margin = Cm(2.5); s.right_margin = Cm(2.5)
        s.header_distance = Cm(1.0)
    add_header(doc, LOGO)
    build(doc, R)
    out = os.path.expanduser(f"~/.hermes/projects/{R['project_name']}100MW风电/{R['project_name']}100MW风电_评估报告_{TODAY}.docx")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    doc.save(out)
    print(f"✓ {out}")

if __name__ == '__main__':
    generate(REPORT_DATA)
