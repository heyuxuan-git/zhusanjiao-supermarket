#!/usr/bin/env python3
"""Generate 珠三角超市 App 功能清单.docx from structured data."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = "/Users/hin/XiaomiMiMoProjects/商超+充电桩/珠三角超市-App功能清单.docx"

INK = RGBColor(0x1F, 0x1F, 0x1F)
MUTED = RGBColor(0x59, 0x59, 0x59)
NAVY = RGBColor(0x1F, 0x3A, 0x5F)
HDR_BG = "1F3A5F"
ALT_BG = "F5F7FA"
P0_BG = "E8F5E9"
P1_BG = "FFF8E1"
P2_BG = "F5F5F5"


def set_run_font(run, name="Calibri", east="PingFang SC", size=11, bold=False, color=INK):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:eastAsia"), east)


def tune_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.font.color.rgb = INK
    normal.paragraph_format.line_spacing = 1.15
    normal.paragraph_format.space_after = Pt(6)
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), "PingFang SC")

    for n, size, color in [(1, 18, NAVY), (2, 14, NAVY), (3, 12, INK)]:
        s = doc.styles[f"Heading {n}"]
        s.font.name = "Calibri Light" if n < 3 else "Calibri"
        s.font.size = Pt(size)
        s.font.bold = True
        s.font.color.rgb = color
        s.paragraph_format.space_before = Pt(14 - 2 * n)
        s.paragraph_format.space_after = Pt(6)
        rPr = s.element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.append(rFonts)
        rFonts.set(qn("w:eastAsia"), "PingFang SC")

    title = doc.styles["Title"]
    title.font.name = "Calibri Light"
    title.font.size = Pt(26)
    title.font.bold = True
    title.font.color.rgb = NAVY
    rPr = title.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), "PingFang SC")


def setup_page(doc):
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21.0), Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(2.2)
    section.left_margin = section.right_margin = Cm(2.2)


def shade_cell(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_cell_text(cell, text, *, bold=False, size=10, color=INK, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.1
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)


def add_table(doc, headers, rows, col_widths=None, priority_col=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_text(cell, h, bold=True, size=10, color=RGBColor(0xFF, 0xFF, 0xFF),
                      align=WD_ALIGN_PARAGRAPH.CENTER)
        shade_cell(cell, HDR_BG)

    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            cell = table.rows[r_idx + 1].cells[c_idx]
            align = WD_ALIGN_PARAGRAPH.CENTER if c_idx in (0, len(headers) - 1) else WD_ALIGN_PARAGRAPH.LEFT
            set_cell_text(cell, str(val), size=9.5, align=align)
            if r_idx % 2 == 1:
                shade_cell(cell, ALT_BG)
            if priority_col is not None and c_idx == priority_col:
                p = str(val)
                if p.startswith("P0"):
                    shade_cell(cell, P0_BG)
                elif p.startswith("P1"):
                    shade_cell(cell, P1_BG)
                elif p.startswith("P2"):
                    shade_cell(cell, P2_BG)
                set_cell_text(cell, p, bold=True, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER)

    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph()
    return table


def add_meta_table(doc):
    rows = [
        ("产品暂定名", "珠三角超市"),
        ("文档版本", "v0.4"),
        ("范围说明", "本版仅到店自提，配送暂不上线"),
        ("配套材料", "功能清单.md · PRD.md · 报价模块清单.md · index.html 线框原型"),
        ("在线预览", "https://heyuxuan-git.github.io/zhusanjiao-supermarket/"),
        ("优先级说明", "P0 首版必做 / P1 可二期 / P2 观望"),
    ]
    add_table(doc, ["字段", "内容"], rows, col_widths=[4, 12])


def build():
    doc = Document()
    setup_page(doc)
    tune_styles(doc)

    p = doc.add_paragraph(style="Title")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run("珠三角超市 App · 功能清单")
    set_run_font(run, name="Calibri Light", size=26, bold=True, color=NAVY)

    sub = doc.add_paragraph()
    run = sub.add_run("半直营连锁商超 · 广告积分 · 驿能充电 · 加油核销")
    set_run_font(run, size=11, color=MUTED)

    add_meta_table(doc)

    doc.add_heading("一、产品一句话", level=1)
    p = doc.add_paragraph()
    for text, bold in [
        ("商超下单到店自提", True),
        (" 为交易底座，", False),
        ("广告任务产积分", True),
        ("，积分兑券/权益/活动，并入 ", False),
        ("驿能充电", True),
        (" 与 ", False),
        ("加油券核销", True),
        ("。", False),
    ]:
        r = p.add_run(text)
        set_run_font(r, bold=bold, size=11)

    p = doc.add_paragraph()
    r = p.add_run("配送到家功能暂不纳入本版清单。")
    set_run_font(r, size=10, color=MUTED, bold=True)

    doc.add_heading("二、底部导航", level=1)
    add_table(doc, ["Tab", "定位"], [
        ("首页", "流量入口：搜索、活动 Banner、最近门店、金刚区、推荐"),
        ("购物", "交易主战场：选门店、到店自提、分类加购、购物车"),
        ("积分任务", "积分获取与消耗入口"),
        ("能源", "充电 + 加油积分"),
        ("我的", "资产、订单、地址、设置"),
    ], col_widths=[3.5, 12.5])

    doc.add_heading("三、按页面功能明细", level=1)

    sections = [
        ("1. 首页", [
            ("H1", "搜索入口", "整行搜索框，进搜索页", "P0"),
            ("H2", "顶部 Banner", "活动 / 新品运营位，可多帧轮播", "P0"),
            ("H3", "最近门店条", "展示距用户最近门店：名称、距离、营业时间", "P0"),
            ("H4", "门店条 → 购物", "点击进入购物页再选门店", "P0"),
            ("H5", "金刚区", "购物、积分任务、充电、加油、兑换、活动、赚积分、我的", "P0"),
            ("H6", "今日推荐", "商品双列卡片，可加购", "P0"),
            ("H7", "副广 Feed 位", "信息流广告展示位", "P0"),
            ("H8", "附近活动", "活动列表入口", "P1"),
        ]),
        ("2. 购物", [
            ("S1", "搜索", "站内商品搜索", "P0"),
            ("S2", "切换门店", "门店列表含距离、营业、有货率；影响价格库存", "P0"),
            ("S3", "到店自提", "本版仅自提；配送暂不上线", "P0"),
            ("S4", "价格提示", "切换门店时提示价格可能变化", "P0"),
            ("S5", "分类侧栏", "二级分类浏览", "P0"),
            ("S6", "商品列表", "横向商品卡：图、价、加购", "P0"),
            ("S7", "购物车 FAB + 抽屉", "悬浮购物车、改数量、去结算", "P0"),
            ("S8", "购物车页", "清空、勾选、金额明细", "P0"),
        ]),
        ("3. 商品 / 搜索", [
            ("P1", "搜索页", "热词、历史、结果列表", "P0"),
            ("P2", "商品详情", "主图、价格、库存、门店、加购/立即购买", "P0"),
            ("P3", "加购", "加入购物车并同步角标", "P0"),
        ]),
        ("4. 结算与订单", [
            ("O1", "确认订单", "自提取货门店", "P0"),
            ("O2", "优惠券抵扣", "选择可用券", "P0"),
            ("O3", "积分抵扣", "积分抵现金，规则可配", "P0"),
            ("O4", "支付", "微信/支付宝", "P0"),
            ("O5", "支付结果", "提货码 + 获得积分", "P0"),
            ("O6", "订单列表", "按状态/类型筛选", "P0"),
            ("O7", "订单详情", "进度、商品、门店、客服", "P0"),
            ("O8", "订单状态机", "待支付→备货→待提货→完成", "P0"),
            ("O9", "售后退款", "首版可人工", "P2"),
        ]),
        ("5. 积分任务", [
            ("T1", "积分余额卡", "总分、今日获得", "P0"),
            ("T2", "激励视频任务", "看完整广告发分，接入联盟 SDK", "P0"),
            ("T3", "浏览副广任务", "停留时长任务", "P0"),
            ("T4", "每日签到", "连续奖励可扩展", "P0"),
            ("T5", "成长/首单任务", "如首单满额奖励", "P1"),
            ("T6", "广告模拟播放层", "框架示意；上线换真 SDK", "P0"),
            ("T7", "积分明细", "获取/消耗流水", "P0"),
            ("T8", "防刷与日上限", "服务端幂等、冷却、上限", "P0"),
            ("T9", "广告任务墙", "副广位资源说明页", "P1"),
        ]),
        ("6. 兑换与权益", [
            ("E1", "兑换商城首页", "积分横幅、分类 Tabs、商品卡", "P0"),
            ("E2", "商超券兑换", "满减券/无门槛券等", "P0"),
            ("E3", "加油券兑换", "积分换加油券", "P0"),
            ("E4", "虚拟权益", "视频会员、话费、咖啡券等", "P1"),
            ("E5", "券包", "已领券/加油券展示与使用引导", "P0"),
            ("E6", "权益履约", "第三方发货/券码下发", "P1"),
            ("E7", "积分不足引导", "引导去看广告/做任务", "P0"),
        ]),
        ("7. 线下活动", [
            ("A1", "活动列表", "时间、地点、积分报名价、余位", "P1"),
            ("A2", "活动详情", "说明、报名", "P1"),
            ("A3", "积分报名", "扣积分得核销码", "P1"),
            ("A4", "到店核销", "码核销", "P1"),
        ]),
        ("8. 能源 · 充电", [
            ("C1", "能源首页", "充电/加油双入口、进行中订单", "P0"),
            ("C2", "找桩列表", "地图+列表、空闲枪、筛选", "P0"),
            ("C3", "电站详情", "电价、空闲、导航", "P0"),
            ("C4", "启动充电", "扫码/选枪", "P0"),
            ("C5", "充电中", "功率、电量、费用实时", "P0"),
            ("C6", "结束与账单", "结算、积分抵扣", "P0"),
            ("C7", "充电订单", "历史账单", "P0"),
            ("C8", "与小程序打通", "API/嵌入/跳转策略待定", "报价必问"),
        ]),
        ("9. 能源 · 加油", [
            ("F1", "加油站列表", "附近、油价、距离", "P0"),
            ("F2", "积分兑加油券", "兑换入券包", "P0"),
            ("F3", "出示核销码", "到店核销", "P0"),
            ("F4", "订单/对账回传", "视接口", "P1"),
        ]),
        ("10. 我的", [
            ("M1", "身份卡", "头像、昵称、手机、会员等级、编辑", "P0"),
            ("M2", "当前门店行", "显示门店，可切换", "P0"),
            ("M3", "资产四格", "积分 / 优惠券 / 加油券 / 权益", "P0"),
            ("M4", "订单四宫格", "待付款/备货/待提货 + 角标", "P0"),
            ("M5", "全部订单", "列表与筛选", "P0"),
            ("M6", "线下活动", "报名记录", "P1"),
            ("M7", "收货地址", "增删改、默认址", "P0"),
            ("M8", "兑换商城入口", "—", "P0"),
            ("M9", "联系客服", "订单/售后/充电", "P1"),
            ("M10", "广告与积分说明", "规则与隐私入口", "P1"),
            ("M11", "设置", "账号安全、通知、隐私、关于、退出", "P0"),
        ]),
        ("11. 全局 / 账号", [
            ("G1", "手机号登录注册", "验证码", "P0"),
            ("G2", "微信授权登录", "可选", "P1"),
            ("G3", "多门店选择", "全局生效，影响价格与库存", "P0"),
            ("G4", "消息通知", "订单/活动/积分", "P1"),
            ("G5", "合规", "隐私政策、注销、个性化广告开关", "P0/P1"),
        ]),
    ]

    for title, items in sections:
        doc.add_heading(title, level=2)
        note = None
        if title.startswith("1."):
            note = "不做：首页门店切换、配送切换、购物车浮层。"
        add_table(doc, ["ID", "功能", "说明", "优先级"],
                  items,
                  col_widths=[1.6, 3.6, 8.4, 2.4],
                  priority_col=3)
        if note:
            p = doc.add_paragraph()
            r = p.add_run(note)
            set_run_font(r, size=9.5, color=MUTED, bold=True)

    doc.add_heading("四、管理后台", level=1)
    p = doc.add_paragraph()
    r = p.add_run("独立拆工，勿并入 App 工时。")
    set_run_font(r, size=10, color=MUTED, bold=True)
    add_table(doc, ["ID", "功能", "优先级"], [
        ("B1", "门店管理：自提点、营业时间", "P0"),
        ("B2", "商品/分类/库存：对接 ERP 或手动", "P0"),
        ("B3", "订单工作台：接单、备货、自提核销", "P0"),
        ("B4", "积分与券策略", "P0"),
        ("B5", "广告任务配置与发分风控", "P0"),
        ("B6", "活动与权益管理", "P1"),
        ("B7", "数据看板：GMV、积分、广告、充电", "P1"),
        ("B8", "角色权限", "P1"),
    ], col_widths=[1.6, 12, 2.4], priority_col=2)

    doc.add_heading("五、核心闭环", level=1)
    steps = [
        "看广告 / 做任务 → 得积分",
        "兑换：商超券 / 加油券 / 权益 / 活动报名",
        "使用：商超下单抵扣 · 充电抵扣 · 加油到店核销",
        "交易底座：多门店 + 到店自提，价格因店可能不同",
    ]
    for s in steps:
        doc.add_paragraph(s, style="List Number")

    doc.add_heading("六、第三方接口确认", level=1)
    p = doc.add_paragraph()
    r = p.add_run("报价关键项，需甲方确认。")
    set_run_font(r, size=10, color=MUTED, bold=True)
    add_table(doc, ["#", "依赖", "待确认"], [
        ("1", "ERP/POS/库存", "有无接口文档"),
        ("2", "支付", "微信/支付宝"),
        ("3", "广告联盟", "穿山甲/优量汇等"),
        ("4", "驿能充电", "API / H5 / 跳转；用户订单迁移？"),
        ("5", "加油站", "字段与核销方式"),
        ("6", "虚拟权益供货", "渠道与结算"),
        ("7", "地图/短信", "SDK 与通道"),
    ], col_widths=[1.2, 4.5, 10.3])

    footer = doc.add_paragraph()
    r = footer.add_run("完整拆工与人日模板见《报价模块清单》；范围与验收见《PRD》。在线原型：https://heyuxuan-git.github.io/zhusanjiao-supermarket/")
    set_run_font(r, size=9, color=MUTED)

    doc.save(OUT)
    print("saved", OUT)


if __name__ == "__main__":
    build()
