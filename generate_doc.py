# -*- coding: utf-8 -*-
"""生成《两轴步进电机控制方案》Word 文档"""
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

# ---------- 封面标题 ----------
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = title.add_run('两轴步进电机控制方案')
set_run(r, size=22, bold=True, color=(0x1F, 0x3B, 0x73))
sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sub.add_run('西门子 S7-200 SMART ST60 · 现场接线与编程手册')
set_run(r, size=12, color=(0x40, 0x40, 0x40))

# ---------- 一、硬件概况 ----------
h1('一、硬件概况')
table(
    ['项目', '参数'],
    [
        ['PLC', '西门子 S7-200 SMART CPU ST60（6ES7288-1ST60-0AA2，DC/DC/DC，36DI / 24DO 晶体管输出）'],
        ['输入寻址', 'I0.0 ~ I3.7（32 点）+ I4.0 ~ I4.3（4 点）= 36 点'],
        ['输出寻址', 'Q0.0 ~ Q2.7 = 24 点；其中 Q0.0 / Q0.1 / Q0.3 为高速脉冲输出（用于运动轴）'],
        ['步进电机', 'SUMTOR 57HS11242B4D8SC：NEMA23、两相、步距角 1.8°（200 脉冲/转）、相电流 4.2A、保持转矩约 3N·m'],
        ['步进驱动器', '需配套步进驱动器（推荐 DM556，供电 24~48VDC，电流拨到 4.2A 附近）'],
        ['制动器', '电机自带永磁制动器（通电释放、失电抱闸，24VDC）'],
        ['原点传感器', 'ZM-Y45，NPN 型'],
    ],
    widths=[3.5, 12.5],
)
para('关键结论（来自西门子官方运动轴输出点固定分配）：ST60 的 3 个高速脉冲点只能分配给 Q0.0、Q0.1、Q0.3，方向/使能信号占用其它普通输出点。', bold=True)

# ---------- 二、I/O 分配表 ----------
h1('二、I/O 分配表')

h2('2.1 输入分配（DI）')
table(
    ['地址', '符号', '说明', '备注'],
    [
        ['I0.0', '启动按钮', '常开(NO)', '系统启动'],
        ['I0.1', '停止按钮', '常闭(NC)', '停止'],
        ['I0.2', '急停按钮', '常闭(NC)', '急停（软件+硬件双重）'],
        ['I0.3', '回原点按钮', '常开(NO)', '触发 X/Y 寻原点'],
        ['I0.4', 'X原点传感器', 'NPN ZM-Y45', '接 X 轴 RPS（参考点）'],
        ['I0.5', 'Y原点传感器', 'NPN ZM-Y45', '接 Y 轴 RPS（参考点）'],
        ['I0.6', 'X正向点动', '常开(NO)', '手动正转'],
        ['I0.7', 'X反向点动', '常开(NO)', '手动反转'],
        ['I1.0', 'Y正向点动', '常开(NO)', '手动正转'],
        ['I1.1', 'Y反向点动', '常开(NO)', '手动反转'],
        ['I1.2', '定位触发(示例)', '常开(NO)', '触发一次定位'],
        ['I1.3 ~ I4.3', '备用', '—', '留作扩展'],
    ],
    widths=[3.2, 3.2, 3.2, 6.4],
)

h2('2.2 输出分配（DO，晶体管源型）')
table(
    ['地址', '符号', '说明'],
    [
        ['Q0.0', 'X轴脉冲 P0', '运动轴 Axis0 脉冲 → 驱动器 X PUL'],
        ['Q0.2', 'X轴方向 P1', 'Axis0 方向 → 驱动器 X DIR'],
        ['Q0.4', 'X轴使能 DIS', 'Axis0 使能 → 驱动器 X ENA（可选）'],
        ['Q0.1', 'Y轴脉冲 P0', '运动轴 Axis1 脉冲 → 驱动器 Y PUL'],
        ['Q0.7', 'Y轴方向 P1', 'Axis1 方向 → 驱动器 Y DIR'],
        ['Q0.5', 'Y轴使能 DIS', 'Axis1 使能 → 驱动器 Y ENA（可选）'],
        ['Q1.0', 'X制动释放', '24V 中间继电器 → X 制动器（通电释放）'],
        ['Q1.1', 'Y制动释放', '24V 中间继电器 → Y 制动器（通电释放）'],
        ['Q0.3 / Q0.6', '保留', '高速输出（Axis2 占用），不要挪用'],
        ['Q1.2 ~ Q2.7', '备用', '留作扩展/指示灯'],
    ],
    widths=[3.2, 3.6, 9.2],
)
para('轴映射依据（脉冲+方向模式）：Axis0 = Q0.0/Q0.2/Q0.4，Axis1 = Q0.1/Q0.7/Q0.5。这是向导固定分配，写程序前在向导里选好轴即可。', bold=True)

# ---------- 三、接线要点 ----------
h1('三、接线要点')
bullet('CPU 供电：L+ 接 +24V，M 接 0V。')
bullet('NPN 传感器（ZM-Y45 三线制）：棕=+24V，蓝=0V，黑=信号到 I 点；同时把输入公共端 1M 接 +24V（NPN 型必须这样接，与 PNP 相反）。')
bullet('驱动器脉冲/方向：Q0.0→PUL、Q0.2→DIR、Q0.4→ENA（X 轴）；Q0.1→PUL、Q0.7→DIR、Q0.5→ENA（Y 轴）。驱动器输入端选 24V 电平（DM556 支持）。PLC 源型输出与驱动器共 0V（M 与驱动器 COM 同地）。')
bullet('制动器不直接接 PLC 点：抱闸线圈电流大，用 Q1.0/Q1.1 驱动 24V 中间继电器，再由继电器触点给制动器供电（通电释放）。')
bullet('急停双重保护：急停常闭触点除进 I0.2 外，建议再串入驱动器使能回路或抱闸电源回路做硬件切断。')

# ---------- 四、运动控制向导组态 ----------
h1('四、运动控制向导组态（关键页）')
para('在 STEP 7-Micro/WIN SMART 的「工具 → 运动向导」里建两个轴：轴0=X轴、轴1=Y轴。')
table(
    ['向导页', '设置项', 'X轴(轴0)', 'Y轴(轴1)'],
    [
        ['运动轴', '轴号/名称', 'Axis0 / X轴', 'Axis1 / Y轴'],
        ['输出点', '输出方式', '脉冲+方向', '脉冲+方向'],
        ['', '脉冲 P0', 'Q0.0（自动）', 'Q0.1（自动）'],
        ['', '方向 P1', 'Q0.2（自动）', 'Q0.7（自动）'],
        ['', '使能 DIS', 'Q0.4（自动）', 'Q0.5（自动）'],
        ['测量系统', '每转脉冲数', '200 × 驱动器细分数', '同左'],
        ['', '每转负载位移', '按传送带实际算', '按实际算'],
        ['速度', '最大速度 MAX_SPEED', '示例 20000 脉冲/s', '同左'],
        ['', '启停速度 SS_SPEED', '示例 200 脉冲/s', '同左'],
        ['', '加/减速时间', '示例 300 ms', '同左'],
        ['点动', '点动速度/加减速', '4000 脉冲/s，200 ms', '同左'],
        ['参考点', '查找模式', '仅用 RPS，不用 ZP', '同左'],
        ['', '搜索/逼近方向', '按传感器安装方向定', '按传感器安装方向定'],
        ['限位', 'LMT+/LMT-', '不组态（无限位开关）', '不组态'],
    ],
    widths=[2.6, 3.6, 5.2, 4.6],
)
para('每转位移换算：传送带每转移动量 = 主动滚筒周长 × 减速比（mm）。例如滚筒直径 63.7mm、无减速比 → ≈200mm/转。填入实际值，PLC 自动换算成脉冲。')

# ---------- 五、主程序 ----------
h1('五、主程序（梯形图 / 可转 LAD 的 STL）')
para('向导编译后会生成 AXIS0_*、AXIS1_* 等一系列子程序，主程序 OB1 调用它们。')

h2('网络1 首次扫描初始化（赋点动/定位速度）')
code_block([
    'LD   SM0.1',
    'MOVD 4000,  VD130     // X 点动速度（脉冲/s）',
    'MOVD 4000,  VD230     // Y 点动速度',
    'MOVD 20000, VD310     // X 定位速度',
    'MOVD 20000, VD410     // Y 定位速度',
])

h2('网络2 运行标志自锁（启动/停止/急停）')
code_block([
    'LD   I0.0      // 启动（常开）',
    'O    M0.0      // 自锁',
    'A    I0.1      // 停止（常闭按钮，触点用常开）',
    'A    I0.2      // 急停（常闭按钮，触点用常开）',
    '=    M0.0      // 运行标志',
])

h2('网络3 抱闸释放与运动允许（先松闸→延时→再允许运动）')
code_block([
    'LD   M0.0      // 运行标志',
    '=    Q1.0      // X 抱闸释放（通电释放）',
    '=    Q1.1      // Y 抱闸释放',
    'TON  T38, +3   // 抱闸松开确认延时 100ms×3=300ms',
    '',
    'LD   T38       // 延时到（抱闸已机械松开）',
    '=    M0.7      // 运动允许标志',
])

h2('网络4 X轴控制使能（MOD_EN 受运动允许 M0.7 门控）')
code_block([
    'AXIS0_CTRL 盒：',
    '  EN = SM0.0，MOD_EN = M0.7',
    '  Done→M1.0，Error→VB100，C_Pos→VD104，C_Speed→VD108，C_Dir→M1.1',
])

h2('网络5 Y轴控制使能（MOD_EN 受运动允许 M0.7 门控）')
code_block([
    'AXIS1_CTRL 盒：',
    '  EN = SM0.0，MOD_EN = M0.7',
    '  Done→M1.2，Error→VB102，C_Pos→VD204，C_Speed→VD208，C_Dir→M1.3',
])

h2('回原点顺序状态机（X 完成后才触发 Y）')
para('状态/标志定义：', bold=True)
table(
    ['符号', '含义', '说明'],
    [
        ['M2.0', '回原点请求', '自锁，回原点完成后自动复位'],
        ['M2.1', 'X 轴回原点阶段', '作为 AXIS0_RSEEK 的 START（上升沿触发）'],
        ['M2.2', 'Y 轴回原点阶段', '作为 AXIS1_RSEEK 的 START（上升沿触发）'],
        ['M2.3', '回原点完成指示', '可接指示灯/HMI，下次请求时清零'],
        ['M2.4', '回原点故障', '寻原点出错或超时置位，新请求时自动清零'],
        ['M4.0', 'X 寻原点完成', 'AXIS0_RSEEK 的 Done 输出'],
        ['M4.1', 'Y 寻原点完成', 'AXIS1_RSEEK 的 Done 输出'],
    ],
    widths=[2.6, 4.4, 9.0],
)
para('阶段转换：空闲 → X 寻原点 → Y 寻原点 → 完成（再次按下 I0.3 可重跑）。')

h2('网络6 回原点请求（上升沿、运行中、空闲时才启动）')
code_block([
    'LD   I0.3      // 回原点按钮',
    'EU             // 上升沿',
    'A    M0.7      // 运动允许（抱闸已松开）',
    'AN   M2.0      // 不在回原点中（空闲）',
    'S    M2.0, 1   // 置位请求',
    'R    M2.3, 1   // 清上次完成标志',
    'R    M2.4, 1   // 清回原点故障标志',
    'S    M2.1, 1   // 进入 X 阶段 → AXIS0_RSEEK 启动',
])

h2('网络7 状态推进：X 完成 → 启动 Y')
code_block([
    'LD   M4.0      // X 寻原点完成',
    'A    M2.1      // 且仍处 X 阶段',
    'R    M2.1, 1   // 退出 X 阶段',
    'S    M2.2, 1   // 进入 Y 阶段 → AXIS1_RSEEK 启动',
])

h2('网络8 状态推进：Y 完成 → 结束')
code_block([
    'LD   M4.1      // Y 寻原点完成',
    'A    M2.2      // 且仍处 Y 阶段',
    'R    M2.2, 1   // 退出 Y 阶段',
    'R    M2.0, 1   // 复位请求',
    'S    M2.3, 1   // 置位完成指示',
])

h2('网络9 X轴寻原点  AXIS0_RSEEK')
code_block([
    'AXIS0_RSEEK 盒：',
    '  EN = SM0.0，START = M2.1',
    '  Done→M4.0，Error→VB110',
])

h2('网络10 Y轴寻原点  AXIS1_RSEEK')
code_block([
    'AXIS1_RSEEK 盒：',
    '  EN = SM0.0，START = M2.2',
    '  Done→M4.1，Error→VB120',
])

h2('网络11 回原点出错复位（X/Y 任一 Error 非零）')
code_block([
    'LD   M2.0      // 回原点请求中',
    'AB<> VB110, 0  // 且 X 错误非零',
    'LD   M2.2      // 或：Y 阶段',
    'AB<> VB120, 0  // 且 Y 错误非零',
    'OLD            // 两个分支取“或”',
    'S    M2.4, 1   // 置位故障标志',
    'R    M2.0, 1   // 复位请求',
    'R    M2.1, 1   // 退出 X 阶段',
    'R    M2.2, 1   // 退出 Y 阶段',
])

h2('网络12 回原点超时定时（T37，5s）')
code_block([
    'LD   M2.0      // 回原点进行中',
    'TON  T37, +50  // 100ms 时基 × 50 = 5s 超时',
])

h2('网络13 超时故障复位')
code_block([
    'LD   T37       // 定时到（超时）',
    'S    M2.4, 1   // 置位故障标志',
    'R    M2.0, 1   // 复位请求',
    'R    M2.1, 1   // 退出 X 阶段',
    'R    M2.2, 1   // 退出 Y 阶段',
])
para('故障标志 M2.4 可接报警输出（如：LD M2.4 = Q1.2）；重新按回原点按钮 I0.3 会自动清零 M2.4；超时时间通过修改 T37 设定值调整。')

h2('网络14 X轴手动点动  AXIS0_MAN')
code_block([
    'AXIS0_MAN 盒：',
    '  EN = SM0.0，RUN = M0.3(0)，JOG_P = I0.6，JOG_N = I0.7，Speed = VD130，Dir = M0.4',
    '  Error→VB140，C_Pos→VD144，C_Speed→VD148，C_Dir→M0.5',
])

h2('网络15 Y轴手动点动  AXIS1_MAN')
code_block([
    'AXIS1_MAN 盒：',
    '  EN = SM0.0，RUN = M0.3(0)，JOG_P = I1.0，JOG_N = I1.1，Speed = VD230，Dir = M0.4',
    '  Error→VB240，C_Pos→VD244，C_Speed→VD248，C_Dir→M0.6',
])

h2('网络16 定位触发示例（受运动允许 M0.7 门控）')
code_block([
    'LD   I1.2',
    'EU',
    'A    M0.7',
    '=    M3.0',
])

h2('网络17 X轴定位到目标位置  AXIS0_GOTO（Mode=0 绝对 / 1 相对）')
code_block([
    'AXIS0_GOTO 盒（X 轴）：',
    '  EN = SM0.0，START = M3.0，Pos = VD300，Speed = VD310，Mode = 0，Abort = M3.1',
    '  Done→M3.2，Error→VB150，C_Pos→VD154，C_Speed→VD158',
    '',
    'AXIS1_GOTO 盒（Y 轴）：',
    '  EN = SM0.0，START = M3.3，Pos = VD400，Speed = VD410，Mode = 0，Abort = M3.4',
    '  Done→M3.5，Error→VB250，C_Pos→VD254，C_Speed→VD258',
])

# ---------- 六、注意事项 ----------
h1('六、注意事项')
bullet('先回原点再用绝对定位：绝对定位（Mode=0）前必须已完成寻原点，否则报「零位置未知」，建议用相对定位（Mode=1）或先跑 RSEEK。')
bullet('抱闸时序已联动：上电运行先释放抱闸，经 T38（300ms）确认机械松开后才置 M0.7「运动允许」，并作为各轴使能 MOD_EN 及回原点/定位的触发门槛；停止或急停时抱闸立即失电抱紧（急停再由硬件串入抱闸电源兜底）。若实际松开时间更长，改 T38 设定值即可。')
bullet('驱动器电流/细分：DM556 电流拨到 4.2A，细分按「每转脉冲数」设定统一（如 8 细分=1600 脉冲/转，则测量系统「每转脉冲数」填 1600）。')
bullet('输出点保护：Q0.0/Q0.1/Q0.3 为脉冲专用，程序里不要用普通线圈去写这几点；Q0.2/Q0.4/Q0.7/Q0.5 被轴占用后也不要重复使用。')
bullet('NPN 共端：所有用到 NPN 传感器的输入组，公共端 M 必须统一接 +24V，不能与 PNP 混接。')
bullet('回原点异常保护：已内置两种保护——X/Y 寻原点出错（VB110/VB120 非零）与超时（T37，默认 5s）。触发时自动复位状态机并置位故障标志 M2.4（可接报警灯），重新按回原点按钮自动清零；超时时长可改 T37 设定值。')

# ---------- 七、现场机械参数待填表 ----------
h1('七、现场机械参数待填表')
para('测量系统与速度参数需现场确认后填写（可打印后手写）。')
table(
    ['参数项', '填写值', '说明 / 计算公式'],
    [
        ['主动滚筒直径 D (mm)', '', '传送带主动辊外径'],
        ['减速比 i', '1', '无减速机填 1'],
        ['每转负载位移 (mm)', '', '= π × D × i（例：φ63.7mm、i=1 → ≈200mm）'],
        ['驱动器细分 (脉冲/转)', '', '例：8 细分 = 1600 脉冲/转'],
        ['每转脉冲数', '', '= 200 × 细分（1.8° 电机）'],
        ['最大速度 MAX_SPEED', '', '建议 ≤ 20000 脉冲/s，留 20% 余量'],
        ['启停速度 SS_SPEED', '', '约最大速度的 5%'],
        ['加速 / 减速时间 (ms)', '', '惯量大取 500~1000'],
        ['点动速度 JOG', '', '调试用低速，例 4000 脉冲/s'],
        ['回原点搜索方向', '', '负向 / 正向（按传感器安装位置）'],
        ['回原点超时 (s)', '5', '改 T37 设定值'],
    ],
    widths=[5.2, 3.0, 7.8],
)
para('换算关系：PLC 按「每转脉冲数」与「每转负载位移」自动换算，程序里直接以 mm 为单位给定位置/速度。')

# ---------- 八、现场调试对照（抱闸时序 + 变量汇总） ----------
h1('八、现场调试对照（抱闸时序 + 变量汇总）')

h2('8.1 抱闸 ↔ 电机转动时序')
para('上电/启动（松闸在前、运动在后）：', bold=True)
code_block([
    '启动 I0.0 = 1',
    '  → M0.0 运行 = 1',
    '  → Q1.0 / Q1.1 = 1   抱闸通电释放（机械松开 100~300ms）',
    '  → T38 计时 300ms',
    '  → M0.7 运动允许 = 1',
    '  → 各轴 MOD_EN 使能 + 放行回原点/点动/定位',
    '  → 才能发脉冲，电机转动',
])
para('停止/急停（失电即抱紧）：', bold=True)
code_block([
    '停止 I0.1 或 急停 I0.2',
    '  → M0.0 = 0',
    '  → Q1.0 / Q1.1 = 0   抱闸立即失电抱紧',
    '  → M0.7 = 0   轴失能',
    '  （急停另由硬件常闭串入抱闸电源，双重保护）',
])
para('调试检查点（对照信号状态）：', bold=True)
table(
    ['阶段', 'M0.0', 'Q1.0/Q1.1', 'T38', 'M0.7', '轴状态'],
    [
        ['未启动', '0', '0', '0', '0', '未使能'],
        ['启动后 ＜300ms', '1', '1', '计时中', '0', '未使能（松闸中）'],
        ['启动后 ≥300ms', '1', '1', '到', '1', '已使能，可运动'],
        ['停止 / 急停', '0', '0', '0', '0', '失能，抱紧'],
    ],
    widths=[3.4, 1.4, 2.2, 1.8, 1.4, 3.6],
)

h2('8.2 变量 / 标志汇总表')
para('M 位（布尔）：', bold=True)
table(
    ['地址', '含义', '地址', '含义'],
    [
        ['M0.0', '运行标志', 'M2.0', '回原点请求'],
        ['M0.3', '连续运行 RUN（备用）', 'M2.1', 'X 回原点阶段'],
        ['M0.4', '手动方向（备用）', 'M2.2', 'Y 回原点阶段'],
        ['M0.5', 'X 手动方向（输出）', 'M2.3', '回原点完成'],
        ['M0.6', 'Y 手动方向（输出）', 'M2.4', '回原点故障'],
        ['M0.7', '运动允许（抱闸已松开）', 'M3.0', '定位启动脉冲'],
        ['M1.0', 'X CTRL Done', 'M3.1', 'X GOTO Abort'],
        ['M1.1', 'X CTRL C_Dir', 'M3.2', 'X GOTO Done'],
        ['M1.2', 'Y CTRL Done', 'M3.3', 'Y GOTO START'],
        ['M1.3', 'Y CTRL C_Dir', 'M3.4', 'Y GOTO Abort'],
        ['M3.5', 'Y GOTO Done', 'M4.0', 'X 寻原点完成'],
        ['M4.1', 'Y 寻原点完成', '', ''],
    ],
    widths=[1.6, 5.4, 1.6, 5.4],
)
para('定时器：', bold=True)
table(
    ['地址', '含义', '地址', '含义'],
    [
        ['T37', '回原点超时（100ms×50=5s）', 'T38', '抱闸松开确认延时（100ms×3=300ms）'],
    ],
    widths=[1.6, 5.4, 1.6, 5.4],
)
para('错误字节 VB：', bold=True)
table(
    ['地址', '含义', '地址', '含义'],
    [
        ['VB100', 'X CTRL 错误', 'VB110', 'X RSEEK 错误'],
        ['VB102', 'Y CTRL 错误', 'VB120', 'Y RSEEK 错误'],
        ['VB140', 'X MAN 错误', 'VB150', 'X GOTO 错误'],
        ['VB240', 'Y MAN 错误', 'VB250', 'Y GOTO 错误'],
    ],
    widths=[1.6, 5.4, 1.6, 5.4],
)
para('数据 VD：', bold=True)
table(
    ['地址', '含义', '地址', '含义'],
    [
        ['VD104', 'X 当前位', 'VD108', 'X 当前速度'],
        ['VD204', 'Y 当前位', 'VD208', 'Y 当前速度'],
        ['VD130', 'X 点动速度', 'VD230', 'Y 点动速度'],
        ['VD300', 'X 定位目标位', 'VD310', 'X 定位速度'],
        ['VD400', 'Y 定位目标位', 'VD410', 'Y 定位速度'],
        ['VD144', 'X 手动当前位', 'VD148', 'X 手动当前速度'],
        ['VD154', 'X GOTO 当前位', 'VD158', 'X GOTO 当前速度'],
        ['VD254', 'Y GOTO 当前位', 'VD258', 'Y GOTO 当前速度'],
        ['VD244', 'Y 手动当前位', 'VD248', 'Y 手动当前速度'],
    ],
    widths=[1.6, 5.4, 1.6, 5.4],
)

h2('8.3 常见故障排查表')
table(
    ['故障现象', '排查信号 / 检查点', '常见原因与处理'],
    [
        ['电机不转', 'M0.0、M0.7、T38、Q1.0/Q1.1；Q0.0/Q0.1', '抱闸未松：M0.7=0(时序未到或急停未复位)；急停 I0.1/停止 I0.2 未复位；驱动器未供电或报警灯亮；脉冲线 Q0.0/Q0.1 接错；运动未触发(M2.0/M3.0 未置)'],
        ['电机转、方向反', 'Q0.2/Q0.7（方向输出）', '方向输出与驱动器 DIR 接反，或向导“反转方向”未勾；对调 DIR 或改方向参数'],
        ['丢步 / 位置偏移', '当前位 VD104/204；最大速度、加减速', '速度/加减速过高；驱动器电流过小(应拨 4.2A)；抱闸未全松开(T38 过短)；负载卡滞；驱动器供电不足'],
        ['抖动 / 异响 / 共振', '加减速时间、细分设置', '加减速过陡；细分太低；进入低速共振区；负载刚性差'],
        ['寻原点失败', 'I0.4/I0.5、M2.4、VB110/VB120、T37', '传感器无信号(查 NPN 接线与 1M 接 +24V)；搜索方向反；挡块没对准/被遮；超时(T37)；轴未使能(M0.7=0)'],
        ['原点找到但位置偏', '每转位移、每转脉冲数', '测量系统参数填错(滚筒直径/细分)；挡块位置不一致'],
        ['急停无效', 'I0.2', '急停接成常开或未串入硬件抱闸回路；程序用错触点类型'],
        ['抱闸发烫', 'Q1.0/Q1.1', '长期通电正常温升；过高则查松闸是否不完全或供电超额定'],
        ['运行中突停/报错', 'VB100/VB102 错误字节', '读错误码定位(超时/限位/参数/驱动器过流)并复位故障标志'],
        ['点动无反应', 'I0.6/I0.7/I1.0/I1.1、M0.7、VD130/230', '运动允许 M0.7 未置；点动速度未赋；MAN 的 JOG_P/JOG_N 接错'],
    ],
    widths=[2.8, 5.0, 6.2],
)

h2('8.4 运动控制错误码对照（读 VB100/VB102/VB110/VB120/VB140/VB240/VB150/VB250）')
para('各错误字节对应指令：VB100=X CTRL、VB102=Y CTRL、VB110=X RSEEK、VB120=Y RSEEK、VB140=X MAN、VB240=Y MAN、VB150=X GOTO、VB250=Y GOTO。错误码即以上字节的当前值。')
table(
    ['错误码', '含义', '本项目常见情形与处理'],
    [
        ['0', '无错误', '正常执行'],
        ['1', '由用户打断', 'GOTO 的 Abort(M3.1) 或运行中使能 M0.7/CTRL 被断开；检查 Abort 与停止逻辑，勿误触发'],
        ['2', '组态错误', '测量系统参数在轴使能时被修改；改向导参数需先停止、重新编译下载'],
        ['3', '非法命令 / 组态块错误', '库存储区被其余程序占用、向导未点选“参考点”、CPU 非晶体管型；核对向导与存储区分配'],
        ['22', '零位置未知', '绝对定位 GOTO(Mode=0) 前未建零点；先执行回原点或在向导/程序里用 LDPOS 建零点'],
        ['128', '该轴无法处理此指令', '同一轴同一时刻只能激活一条运动指令(MAN/RSEEK/GOTO 不可同时)；START 必须用沿触发、且 CTRL 保持使能'],
    ],
    widths=[1.8, 5.0, 7.2],
)
para('特别注意（“找不到原点”通常不落入上面某个错误码）：当 RSEEK 因原点传感器无信号而找不到参考点时，往往不立即报错，而是长时间不完成（Done 不复位）。本方案已用 T37(5s) 超时置位故障标志 M2.4 兜底，此时排查 I0.4/I0.5 传感器信号、搜索方向与 1M 接线。', size=9.5)
para('完整错误码表以 STEP 7-Micro/WIN SMART 的 F1 帮助（运动控制 → 错误代码）为准。', size=9.5)

# ---------- 九、附录：STL 程序清单（照抄版） ----------
h1('九、附录：STL 程序清单（照抄版）')
para('纯 STL 网络可在 STL 视图逐行输入后切回 LAD；AXISx_* 运动指令需用向导生成的指令盒插入 LAD 并按引脚接线。', size=9.5)

para('可在 STL 视图输入的网络：', bold=True)
code_block([
    '// 网络1 首次扫描初始化',
    'LD   SM0.1',
    'MOVD 4000,  VD130',
    'MOVD 4000,  VD230',
    'MOVD 20000, VD310',
    'MOVD 20000, VD410',
    '',
    '// 网络2 运行标志自锁',
    'LD   I0.0',
    'O    M0.0',
    'A    I0.1',
    'A    I0.2',
    '=    M0.0',
    '',
    '// 网络3 抱闸释放与运动允许',
    'LD   M0.0',
    '=    Q1.0',
    '=    Q1.1',
    'TON  T38, +3',
    '',
    'LD   T38',
    '=    M0.7',
    '',
    '// 网络6 回原点请求',
    'LD   I0.3',
    'EU',
    'A    M0.7',
    'AN   M2.0',
    'S    M2.0, 1',
    'R    M2.3, 1',
    'R    M2.4, 1',
    'S    M2.1, 1',
    '',
    '// 网络7 X完成→Y',
    'LD   M4.0',
    'A    M2.1',
    'R    M2.1, 1',
    'S    M2.2, 1',
    '',
    '// 网络8 Y完成→结束',
    'LD   M4.1',
    'A    M2.2',
    'R    M2.2, 1',
    'R    M2.0, 1',
    'S    M2.3, 1',
    '',
    '// 网络11 出错复位',
    'LD   M2.0',
    'AB<> VB110, 0',
    'LD   M2.2',
    'AB<> VB120, 0',
    'OLD',
    'S    M2.4, 1',
    'R    M2.0, 1',
    'R    M2.1, 1',
    'R    M2.2, 1',
    '',
    '// 网络12 超时定时',
    'LD   M2.0',
    'TON  T37, +50',
    '',
    '// 网络13 超时复位',
    'LD   T37',
    'S    M2.4, 1',
    'R    M2.0, 1',
    'R    M2.1, 1',
    'R    M2.2, 1',
    '',
    '// 网络16 定位触发',
    'LD   I1.2',
    'EU',
    'A    M0.7',
    '=    M3.0',
])

para('运动指令盒引脚（在 LAD 视图拖入指令盒后接线）：', bold=True)
code_block([
    '// 网络4  AXIS0_CTRL    EN=SM0.0  MOD_EN=M0.7;  Done→M1.0  Error→VB100  C_Pos→VD104  C_Speed→VD108  C_Dir→M1.1',
    '// 网络5  AXIS1_CTRL    EN=SM0.0  MOD_EN=M0.7;  Done→M1.2  Error→VB102  C_Pos→VD204  C_Speed→VD208  C_Dir→M1.3',
    '// 网络9  AXIS0_RSEEK   EN=SM0.0  START=M2.1;   Done→M4.0  Error→VB110',
    '// 网络10 AXIS1_RSEEK   EN=SM0.0  START=M2.2;   Done→M4.1  Error→VB120',
    '// 网络14 AXIS0_MAN     EN=SM0.0  RUN=M0.3  JOG_P=I0.6  JOG_N=I0.7  Speed=VD130  Dir=M0.4;  Error→VB140  C_Pos→VD144  C_Speed→VD148  C_Dir→M0.5',
    '// 网络15 AXIS1_MAN     EN=SM0.0  RUN=M0.3  JOG_P=I1.0  JOG_N=I1.1  Speed=VD230  Dir=M0.4;  Error→VB240  C_Pos→VD244  C_Speed→VD248  C_Dir→M0.6',
    '// 网络17 AXIS0_GOTO    EN=SM0.0  START=M3.0  Pos=VD300  Speed=VD310  Mode=0  Abort=M3.1;  Done→M3.2  Error→VB150  C_Pos→VD154  C_Speed→VD158',
    '// 网络18 AXIS1_GOTO    EN=SM0.0  START=M3.3  Pos=VD400  Speed=VD410  Mode=0  Abort=M3.4;  Done→M3.5  Error→VB250  C_Pos→VD254  C_Speed→VD258',
])

out = r'e:\两轴步进电机控制\两轴步进电机控制方案.docx'
doc.save(out)
print('已生成:', out)