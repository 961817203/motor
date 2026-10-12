# -*- coding: utf-8 -*-
"""生成《运动控制向导组态过程说明》Word 文档"""
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# ---------- 基础样式 ----------
normal = doc.styles['Normal']
normal.font.name = '宋体'
normal.font.size = Pt(10.5)
normal.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

def set_run(run, ea='宋体', ascii_font='宋体', size=10.5, bold=False, color=None, mono=False):
    run.font.name = ascii_font
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), ea)
    if mono:
        rfonts.set(qn('w:ascii'), ascii_font)
        rfonts.set(qn('w:hAnsi'), ascii_font)

def shade_cell(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)

def shade_paragraph(par, fill):
    pPr = par._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), fill)
    pPr.append(shd)

def h1(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run(r, size=16, bold=True, color=(0x1F, 0x3B, 0x73))
    p.space_before = Pt(12)
    p.space_after = Pt(6)
    return p

def h2(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run(r, size=13, bold=True, color=(0x1F, 0x3B, 0x73))
    p.space_before = Pt(8)
    p.space_after = Pt(4)
    return p

def para(text, bold=False, size=10.5):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run(r, size=size, bold=bold)
    return p

def bullet(text):
    p = doc.add_paragraph(style='List Bullet')
    r = p.add_run(text)
    set_run(r)
    return p

def code_block(lines):
    for ln in lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        r = p.add_run(ln if ln else ' ')
        set_run(r, ea='宋体', ascii_font='Consolas', size=9, mono=True)
        shade_paragraph(p, 'F2F2F2')

def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ''
        r = hdr[i].paragraphs[0].add_run(h)
        set_run(r, bold=True, size=10)
        shade_cell(hdr[i], 'DEEAF6')
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ''
            r = cells[i].paragraphs[0].add_run(str(v))
            set_run(r, size=9.5)
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    return t

def placeholder(text):
    import os, re
    m = re.match(r'截图(\d+)', text)
    if m:
        path = r'e:\两轴步进电机控制\截图\截图' + m.group(1) + '.png'
        if os.path.exists(path):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(path, width=Cm(15))
            cap = doc.add_paragraph()
            cr = cap.add_run('▲ ' + text)
            set_run(cr, size=8.5, color=(0x60, 0x60, 0x60))
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            return p
    p = doc.add_paragraph()
    r = p.add_run('【附图占位】' + text)
    set_run(r, size=9.5, bold=True, color=(0xB0, 0x6A, 0x00))
    shade_paragraph(p, 'FFF2CC')
    return p

# ---------- 封面标题 ----------
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = title.add_run('运动控制向导组态过程说明')
set_run(r, size=22, bold=True, color=(0x1F, 0x3B, 0x73))
sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sub.add_run('西门子 S7-200 SMART ST60 · 两轴（X/Y）步进电机 · 组态步骤与生成文件手册')
set_run(r, size=12, color=(0x40, 0x40, 0x40))

para('')
para('说明：本手册以「文字步骤 + 示意图」呈现向导组态全过程。因无法直接抓取软件界面截图，方框图/流程图均为示意，页面名称与按钮文字以你安装的 STEP 7-Micro/WIN SMART 实际版本为准，组态时请对照软件逐项核对。', bold=True, size=10)

# ---------- 一、开始前准备 ----------
h1('一、开始前准备')
bullet('CPU 必须是「晶体管输出型」（型号末尾 ST，如 6ES7288-1ST60-0AA2），继电器型 CPU 无高速脉冲输出，无法驱动步进电机。')
bullet('本项目组态两个运动轴：轴0（Axis0）= X 轴，轴1（Axis1）= Y 轴，二者均为「脉冲 + 方向」开环控制。')
bullet('组态前先确定机械/速度参数：每转脉冲数 = 200 × 驱动器细分；每转位移 = π × 滚筒直径 × 减速比（见方案文档第七节机械参数待填表）。')
bullet('建议流程：先完成向导组态 → 用「运动控制面板」试运行、核对接线 → 再编写 AXISx_* 程序。')

para('向导入口路径：')
code_block([
    'STEP 7-Micro/WIN SMART',
    ' |-- 菜单栏「工具(Tools)」',
    ' |      |-- 「运动(Motion)」  -->  打开「运动向导」',
    ' |-- 或：左侧项目树「工具」文件夹下双击「运动」',
])
placeholder('截图0：软件主界面与向导入口（工具 → 运动）')

# ---------- 二、组态步骤 ----------
h1('二、向导组态步骤（逐页详解）')
para('向导总流程：')
code_block([
    '[选轴/命名] --> [测量系统] --> [速度/动态] --> [回原点] --> [脉冲输出] --> [库存储区] --> [生成组件]',
    '  轴0=X 轴      每转脉冲数      最大/启停速度     参考点模式     脉冲+方向      建议地址        生成子程序',
    '  轴1=Y 轴      每转位移        加减速时间        方向/边沿      引脚映射      记录勿占用      +数据块+符号',
])

h2('步骤1：选择要组态的轴并命名')
placeholder('截图1：向导「选择运动轴 / 命名」页面')
table(
    ['设置项', '本项目填写'],
    [
        ['勾选轴', '轴0（Axis0）、轴1（Axis1）均勾选'],
        ['轴0 名称', 'X轴（可自定义，影响符号名 AXIS0_*）'],
        ['轴1 名称', 'Y轴（可自定义，影响符号名 AXIS1_*）'],
    ],
    widths=[5.0, 11.0],
)
para('说明：每个轴会生成一组独立的 AXISx_* 子程序，命名后符号表会以该编号区分两轴。', size=9.5)

h2('步骤2：测量系统（测量单位）')
placeholder('截图2：向导「测量系统」页面')
para('把「每转脉冲数」与「每转移动距离」填对，程序里才能直接用 mm 给定位置/速度。')
table(
    ['设置项', '本项目填写 / 公式'],
    [
        ['测量单位', '工程单位 mm（或按需要选择脉冲）'],
        ['电机每转脉冲数', '= 200 × 细分（例：细分 8 → 1600 脉冲/转）'],
        ['每转移动距离(mm)', '= π × 滚筒直径 × 减速比（例：φ63.7 无减速 → ≈200mm/转）'],
    ],
    widths=[5.0, 11.0],
)

h2('步骤3：速度与动态特性')
placeholder('截图3：向导「速度与动态特性」页面')
table(
    ['设置项', '本项目推荐值', '说明'],
    [
        ['最大速度 MAX_SPEED', '20000 脉冲/s', '≤ 100kHz，留 20% 余量'],
        ['启停速度 SS_SPEED', '200 脉冲/s', '约最大速度的 5%'],
        ['加速时间 / 减速时间', '300 ms', '惯量大可取 500~1000ms'],
        ['急停减速度', '同减速时间', '立即停止时的减速度'],
    ],
    widths=[4.8, 3.4, 7.8],
)

h2('步骤4：回原点（参考点）')
placeholder('截图4：向导「回原点 / 参考点」页面')
para('本项目原点传感器只有 RPS、无编码器 ZP 零脉冲，因此只使用 RPS。')
table(
    ['设置项', '本项目填写'],
    [
        ['参考点查找模式', '仅使用 RPS（参考点开关），不使用 ZP 零脉冲'],
        ['查找方向', '根据地传感器安装位置选择（向传感器一侧）'],
        ['逼近方向', '与查找方向一致或相反（按实际回位方向）'],
        ['RPS 有效边沿', '按传感器遮挡到/离开方式选择'],
        ['参考点偏移 RP Offset', '0（或按机械原点补差）'],
        ['限位 LMT+/LMT-', '不组态（本项目无限位开关）'],
    ],
    widths=[5.0, 11.0],
)

h2('步骤5：脉冲输出方式与输出点')
placeholder('截图5：向导「脉冲输出 / 输出点」页面')
para('选择「脉冲 + 方向」后，向导自动映射高速脉冲点（固定义，不可改）。')
table(
    ['轴', '脉冲 P0', '方向 P1', '使能 DIS'],
    [
        ['轴0（X轴）', 'Q0.0', 'Q0.2', 'Q0.4'],
        ['轴1（Y轴）', 'Q0.1', 'Q0.7', 'Q0.5'],
    ],
    widths=[3.5, 3.5, 3.5, 3.5],
)
para('注意：Q0.3、Q0.6 为轴2（Axis2）预留的高速/使能点，本项目不使用，程序里也不要挪用。', size=9.5)

h2('步骤6：库存储区分配')
placeholder('截图6：向导「库存储区分配」页面')
bullet('点「建议地址」由软件自动分配一段 V 存储区给运动库，记录下该起始地址。')
bullet('该段 V 区被向导独占，主程序中其它逻辑绝不能再占用（否则报错误 3 / 组态错误）。')
bullet('若改为手动指定，需确保与程序里已有的 V 地址（VB/VD 变量）不冲突。')

h2('步骤7：生成组件')
placeholder('截图7：向导「生成组件」完成页面')
para('点击「生成」，向导创建运动控制子程序、符号表条目、数据块，并把运动配置写入系统块。之后需「编译」并「下载」到 PLC。')

# ---------- 三、生成组件 ----------
h1('三、向导生成的组件与文件名称')
para('生成后，每个轴会创建一组 AXISx_* 子程序（x=0 为 X 轴、x=1 为 Y 轴），共 11 个：')
table(
    ['子程序', '功能', '本项目是否使用'],
    [
        ['AXISx_CTRL', '轴使能 / 初始化，每个扫描周期必调', '是（网络4/5）'],
        ['AXISx_MAN', '手动模式（点动 / 连续运行）', '是（网络14/15）'],
        ['AXISx_GOTO', '定位到指定位置（绝对/相对/单速连续）', '是（网络17）'],
        ['AXISx_RUN', '运行速度包络（多段速度曲线）', '否（可选扩展）'],
        ['AXISx_RSEEK', '找参考点 / 回原点', '是（网络9/10）'],
        ['AXISx_LDPOS', '装载当前位置值（软设零点）', '可选（建零点备用）'],
        ['AXISx_LDOFF', '装载参考点偏移', '否'],
        ['AXISx_SRATE', '设置加减速倍率 / 速率', '可选'],
        ['AXISx_DIS', '输出 / 关闭 DIS 使能信号', '可选（用输出使能时）'],
        ['AXISx_CFG', '重新载入轴组态', '可选（改参数后）'],
        ['AXISx_CACHE', '高速脉冲数据缓存（优化）', '无需手动调用'],
    ],
    widths=[4.2, 8.0, 3.8],
)
para('生成后的项目树结构（示意）：')
code_block([
    '项目树（生成后）',
    ' |-- 程序块',
    ' |     +-- 调用子程序（运动向导生成）',
    ' |          |-- AXIS0_CTRL / AXIS1_CTRL',
    ' |          |-- AXIS0_MAN  / AXIS1_MAN',
    ' |          |-- AXIS0_GOTO / AXIS1_GOTO',
    ' |          |-- AXIS0_RUN  / AXIS1_RUN',
    ' |          |-- AXIS0_RSEEK/ AXIS1_RSEEK',
    ' |          |-- AXIS0_LDPOS/ AXIS1_LDPOS',
    ' |          |-- AXIS0_LDOFF/ AXIS1_LDOFF',
    ' |          |-- AXIS0_SRATE/ AXIS1_SRATE',
    ' |          |-- AXIS0_DIS  / AXIS1_DIS',
    ' |          |-- AXIS0_CFG  / AXIS1_CFG',
    ' |          +-- AXIS0_CACHE/ AXIS1_CACHE',
    ' |-- 符号表（新增运动轴状态/配置变量）',
    ' |-- 数据块（运动组态数据）',
    ' +-- 库存储区（PTO 运动库占用的 V 区）',
])

# ---------- 四、指令盒引脚 ----------
h1('四、本项目用到的指令盒引脚速查')
para('在 LAD 视图从「指令树 → 调用子程序」拖入对应 AXISx_* 指令盒后，按以下引脚接线（X 轴用 AXIS0_*、Y 轴用 AXIS1_*）。')

h2('AXISx_CTRL（网络4/5）')
table(
    ['引脚', '类型', '本项目接线（X 轴 / Y 轴）'],
    [
        ['EN', '输入', 'SM0.0（始终使能调用）'],
        ['MOD_EN', '输入', 'M0.7（运动允许）'],
        ['Done', '输出', 'M1.0 / M1.2'],
        ['Error', '输出', 'VB100 / VB102'],
        ['C_Pos', '输出', 'VD104 / VD204'],
        ['C_Speed', '输出', 'VD108 / VD208'],
        ['C_Dir', '输出', 'M1.1 / M1.3'],
    ],
    widths=[3.0, 2.5, 6.5],
)

h2('AXISx_RSEEK（网络9/10）')
table(
    ['引脚', '类型', '本项目接线（X 轴 / Y 轴）'],
    [
        ['EN', '输入', 'SM0.0'],
        ['START', '输入', 'M2.1 / M2.2（沿触发）'],
        ['Done', '输出', 'M4.0 / M4.1'],
        ['Error', '输出', 'VB110 / VB120'],
    ],
    widths=[3.0, 2.5, 6.5],
)

h2('AXISx_MAN（网络14/15）')
table(
    ['引脚', '类型', '本项目接线（X 轴 / Y 轴）'],
    [
        ['EN', '输入', 'SM0.0'],
        ['RUN', '输入', 'M0.3（备用，接 0）'],
        ['JOG_P', '输入', 'I0.6 / I1.0'],
        ['JOG_N', '输入', 'I0.7 / I1.1'],
        ['Speed', '输入', 'VD130 / VD230'],
        ['Dir', '输入', 'M0.4（备用，接 0）'],
        ['Error', '输出', 'VB140 / VB240'],
        ['C_Pos', '输出', 'VD144 / VD244'],
        ['C_Speed', '输出', 'VD148 / VD248'],
        ['C_Dir', '输出', 'M0.5 / M0.6'],
    ],
    widths=[3.0, 2.5, 6.5],
)

h2('AXISx_GOTO（网络17，以 X 轴为例）')
table(
    ['引脚', '类型', '本项目接线'],
    [
        ['EN', '输入', 'SM0.0'],
        ['START', '输入', 'M3.0（沿触发）'],
        ['Pos', '输入', 'VD300（目标位置）'],
        ['Speed', '输入', 'VD310（定位速度）'],
        ['Mode', '输入', '0=绝对，1=相对'],
        ['Abort', '输入', 'M3.1（中止）'],
        ['Done', '输出', 'M3.2'],
        ['Error', '输出', 'VB150'],
        ['C_Pos', '输出', 'VD154'],
        ['C_Speed', '输出', 'VD158'],
    ],
    widths=[3.0, 2.5, 6.5],
)

# ---------- 五、参数速查 ----------
h1('五、本项目组态参数速查（汇总）')
table(
    ['页面', '关键项', '推荐值 / 填写'],
    [
        ['选轴', '轴0 / 轴1', 'X 轴 / Y 轴'],
        ['测量系统', '每转脉冲数', '200 × 细分'],
        ['测量系统', '每转位移', 'π × 滚筒直径 × 减速比'],
        ['速度', '最大速度', '20000 脉冲/s'],
        ['速度', '启停速度', '200 脉冲/s'],
        ['速度', '加减速时间', '300 ms'],
        ['回原点', '模式', '仅用 RPS，不用 ZP'],
        ['回原点', '限位', '不组态'],
        ['输出', '方式', '脉冲 + 方向'],
        ['输出', '引脚', 'X: Q0.0/Q0.2/Q0.4；Y: Q0.1/Q0.7/Q0.5'],
        ['库存储区', '起始地址', '点「建议地址」自动分配并记录'],
    ],
    widths=[4.0, 4.5, 7.5],
)

# ---------- 六、注意事项 ----------
h1('六、注意事项')
bullet('组态完成后务必「编译」并「下载」到 PLC，运动配置才生效（仅在向导里点「生成」不下载无效）。')
bullet('先用运动控制面板手动点动/回原点测试，确认接线与方向无误后，再编写 AxISx_* 程序。')
bullet('同一轴同一时刻只能激活一条运动指令（MAN / RSEEK / GOTO 不可同时），否则报错误 128。')
bullet('绝对定位 GOTO(Mode=0) 前必须先回原点或 LDPOS 建零点，否则报错误 22。')
bullet('改测量系统/速度参数时，轴须处于未使能状态，改完重新编译下载，否则报错误 2。')
bullet('库存储区分配的 V 地址不要被主程序其它变量占用，否则报错误 3。')

out = r'e:\两轴步进电机控制\运动向导组态过程说明.docx'
doc.save(out)
print('已生成:', out)