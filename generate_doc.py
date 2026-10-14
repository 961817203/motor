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


{
  "name": "@scipen/mini-claude-code",
  "version": "0.8.1",
  "lockfileVersion": 3,
  "requires": true,
  "packages": {
    "": {
      "name": "@scipen/mini-claude-code",
      "version": "0.8.1",
      "license": "MIT",
      "dependencies": {
        "@anthropic-ai/sdk": "^0.64.0",
        "@inquirer/prompts": "^5.5.0",
        "@modelcontextprotocol/sdk": "^1.18.2",
        "chalk": "^5.6.2",
        "docx": "^9.5.1",
        "ora": "^9.0.0"
      },
      "bin": {
        "mini-claude-code": "dist/index.js"
      },
      "devDependencies": {
        "@types/node": "^20.0.0",
        "ts-node": "^10.9.0",
        "typescript": "^5.0.0"
      },
      "engines": {
        "node": ">=16.0.0"
      }
    },
    "node_modules/@anthropic-ai/sdk": {
      "version": "0.64.0",
      "resolved": "https://registry.npmmirror.com/@anthropic-ai/sdk/-/sdk-0.64.0.tgz",
      "integrity": "sha512-K2JaWzPiIA4/ghD7bMlyM/4JvM3k7sKdHFmXWAKqqpVL9ee1H7pguRqygxTgK+SrzK86ZMf1x3L+RK4IrDtxjA==",
      "dependencies": {
        "json-schema-to-ts": "^3.1.1"
      },
      "bin": {
        "anthropic-ai-sdk": "bin/cli"
      },
      "peerDependencies": {
        "zod": "^3.25.0 || ^4.0.0"
      },
      "peerDependenciesMeta": {
        "zod": {
          "optional": true
        }
      }
    },
    "node_modules/@babel/runtime": {
      "version": "7.28.4",
      "resolved": "https://registry.npmmirror.com/@babel/runtime/-/runtime-7.28.4.tgz",
      "integrity": "sha512-Q/N6JNWvIvPnLDvjlE1OUBLPQHH6l3CltCEsHIujp45zQUSSh8K+gHnaEX45yAT1nyngnINhvWtzN+Nb9D8RAQ==",
      "engines": {
        "node": ">=6.9.0"
      }
    },
    "node_modules/@cspotcode/source-map-support": {
      "version": "0.8.1",
      "resolved": "https://registry.npmmirror.com/@cspotcode/source-map-support/-/source-map-support-0.8.1.tgz",
      "integrity": "sha512-IchNf6dN4tHoMFIn/7OE8LWZ19Y6q/67Bmf6vnGREv8RSbBVb9LPJxEcnwrcwX6ixSvaiGoomAUvu4YSxXrVgw==",
      "dev": true,
      "dependencies": {
        "@jridgewell/trace-mapping": "0.3.9"
      },
      "engines": {
        "node": ">=12"
      }
    },
    "node_modules/@inquirer/checkbox": {
      "version": "2.5.0",
      "resolved": "https://registry.npmmirror.com/@inquirer/checkbox/-/checkbox-2.5.0.tgz",
      "integrity": "sha512-sMgdETOfi2dUHT8r7TT1BTKOwNvdDGFDXYWtQ2J69SvlYNntk9I/gJe7r5yvMwwsuKnYbuRs3pNhx4tgNck5aA==",
      "license": "MIT",
      "dependencies": {
        "@inquirer/core": "^9.1.0",
        "@inquirer/figures": "^1.0.5",
        "@inquirer/type": "^1.5.3",
        "ansi-escapes": "^4.3.2",
        "yoctocolors-cjs": "^2.1.2"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@inquirer/confirm": {
      "version": "3.2.0",
      "resolved": "https://registry.npmmirror.com/@inquirer/confirm/-/confirm-3.2.0.tgz",
      "integrity": "sha512-oOIwPs0Dvq5220Z8lGL/6LHRTEr9TgLHmiI99Rj1PJ1p1czTys+olrgBqZk4E2qC0YTzeHprxSQmoHioVdJ7Lw==",
      "license": "MIT",
      "dependencies": {
        "@inquirer/core": "^9.1.0",
        "@inquirer/type": "^1.5.3"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@inquirer/core": {
      "version": "9.2.1",
      "resolved": "https://registry.npmmirror.com/@inquirer/core/-/core-9.2.1.tgz",
      "integrity": "sha512-F2VBt7W/mwqEU4bL0RnHNZmC/OxzNx9cOYxHqnXX3MP6ruYvZUZAW9imgN9+h/uBT/oP8Gh888J2OZSbjSeWcg==",
      "license": "MIT",
      "dependencies": {
        "@inquirer/figures": "^1.0.6",
        "@inquirer/type": "^2.0.0",
        "@types/mute-stream": "^0.0.4",
        "@types/node": "^22.5.5",
        "@types/wrap-ansi": "^3.0.0",
        "ansi-escapes": "^4.3.2",
        "cli-width": "^4.1.0",
        "mute-stream": "^1.0.0",
        "signal-exit": "^4.1.0",
        "strip-ansi": "^6.0.1",
        "wrap-ansi": "^6.2.0",
        "yoctocolors-cjs": "^2.1.2"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@inquirer/core/node_modules/@inquirer/type": {
      "version": "2.0.0",
      "resolved": "https://registry.npmmirror.com/@inquirer/type/-/type-2.0.0.tgz",
      "integrity": "sha512-XvJRx+2KR3YXyYtPUUy+qd9i7p+GO9Ko6VIIpWlBrpWwXDv8WLFeHTxz35CfQFUiBMLXlGHhGzys7lqit9gWag==",
      "license": "MIT",
      "dependencies": {
        "mute-stream": "^1.0.0"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@inquirer/core/node_modules/@types/node": {
      "version": "22.18.8",
      "resolved": "https://registry.npmmirror.com/@types/node/-/node-22.18.8.tgz",
      "integrity": "sha512-pAZSHMiagDR7cARo/cch1f3rXy0AEXwsVsVH09FcyeJVAzCnGgmYis7P3JidtTUjyadhTeSo8TgRPswstghDaw==",
      "license": "MIT",
      "dependencies": {
        "undici-types": "~6.21.0"
      }
    },
    "node_modules/@inquirer/editor": {
      "version": "2.2.0",
      "resolved": "https://registry.npmmirror.com/@inquirer/editor/-/editor-2.2.0.tgz",
      "integrity": "sha512-9KHOpJ+dIL5SZli8lJ6xdaYLPPzB8xB9GZItg39MBybzhxA16vxmszmQFrRwbOA918WA2rvu8xhDEg/p6LXKbw==",
      "license": "MIT",
      "dependencies": {
        "@inquirer/core": "^9.1.0",
        "@inquirer/type": "^1.5.3",
        "external-editor": "^3.1.0"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@inquirer/expand": {
      "version": "2.3.0",
      "resolved": "https://registry.npmmirror.com/@inquirer/expand/-/expand-2.3.0.tgz",
      "integrity": "sha512-qnJsUcOGCSG1e5DTOErmv2BPQqrtT6uzqn1vI/aYGiPKq+FgslGZmtdnXbhuI7IlT7OByDoEEqdnhUnVR2hhLw==",
      "license": "MIT",
      "dependencies": {
        "@inquirer/core": "^9.1.0",
        "@inquirer/type": "^1.5.3",
        "yoctocolors-cjs": "^2.1.2"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@inquirer/figures": {
      "version": "1.0.13",
      "resolved": "https://registry.npmmirror.com/@inquirer/figures/-/figures-1.0.13.tgz",
      "integrity": "sha512-lGPVU3yO9ZNqA7vTYz26jny41lE7yoQansmqdMLBEfqaGsmdg7V3W9mK9Pvb5IL4EVZ9GnSDGMO/cJXud5dMaw==",
      "license": "MIT",
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@inquirer/input": {
      "version": "2.3.0",
      "resolved": "https://registry.npmmirror.com/@inquirer/input/-/input-2.3.0.tgz",
      "integrity": "sha512-XfnpCStx2xgh1LIRqPXrTNEEByqQWoxsWYzNRSEUxJ5c6EQlhMogJ3vHKu8aXuTacebtaZzMAHwEL0kAflKOBw==",
      "license": "MIT",
      "dependencies": {
        "@inquirer/core": "^9.1.0",
        "@inquirer/type": "^1.5.3"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@inquirer/number": {
      "version": "1.1.0",
      "resolved": "https://registry.npmmirror.com/@inquirer/number/-/number-1.1.0.tgz",
      "integrity": "sha512-ilUnia/GZUtfSZy3YEErXLJ2Sljo/mf9fiKc08n18DdwdmDbOzRcTv65H1jjDvlsAuvdFXf4Sa/aL7iw/NanVA==",
      "license": "MIT",
      "dependencies": {
        "@inquirer/core": "^9.1.0",
        "@inquirer/type": "^1.5.3"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@inquirer/password": {
      "version": "2.2.0",
      "resolved": "https://registry.npmmirror.com/@inquirer/password/-/password-2.2.0.tgz",
      "integrity": "sha512-5otqIpgsPYIshqhgtEwSspBQE40etouR8VIxzpJkv9i0dVHIpyhiivbkH9/dGiMLdyamT54YRdGJLfl8TFnLHg==",
      "license": "MIT",
      "dependencies": {
        "@inquirer/core": "^9.1.0",
        "@inquirer/type": "^1.5.3",
        "ansi-escapes": "^4.3.2"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@inquirer/prompts": {
      "version": "5.5.0",
      "resolved": "https://registry.npmmirror.com/@inquirer/prompts/-/prompts-5.5.0.tgz",
      "integrity": "sha512-BHDeL0catgHdcHbSFFUddNzvx/imzJMft+tWDPwTm3hfu8/tApk1HrooNngB2Mb4qY+KaRWF+iZqoVUPeslEog==",
      "license": "MIT",
      "dependencies": {
        "@inquirer/checkbox": "^2.5.0",
        "@inquirer/confirm": "^3.2.0",
        "@inquirer/editor": "^2.2.0",
        "@inquirer/expand": "^2.3.0",
        "@inquirer/input": "^2.3.0",
        "@inquirer/number": "^1.1.0",
        "@inquirer/password": "^2.2.0",
        "@inquirer/rawlist": "^2.3.0",
        "@inquirer/search": "^1.1.0",
        "@inquirer/select": "^2.5.0"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@inquirer/rawlist": {
      "version": "2.3.0",
      "resolved": "https://registry.npmmirror.com/@inquirer/rawlist/-/rawlist-2.3.0.tgz",
      "integrity": "sha512-zzfNuINhFF7OLAtGHfhwOW2TlYJyli7lOUoJUXw/uyklcwalV6WRXBXtFIicN8rTRK1XTiPWB4UY+YuW8dsnLQ==",
      "license": "MIT",
      "dependencies": {
        "@inquirer/core": "^9.1.0",
        "@inquirer/type": "^1.5.3",
        "yoctocolors-cjs": "^2.1.2"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@inquirer/search": {
      "version": "1.1.0",
      "resolved": "https://registry.npmmirror.com/@inquirer/search/-/search-1.1.0.tgz",
      "integrity": "sha512-h+/5LSj51dx7hp5xOn4QFnUaKeARwUCLs6mIhtkJ0JYPBLmEYjdHSYh7I6GrLg9LwpJ3xeX0FZgAG1q0QdCpVQ==",
      "license": "MIT",
      "dependencies": {
        "@inquirer/core": "^9.1.0",
        "@inquirer/figures": "^1.0.5",
        "@inquirer/type": "^1.5.3",
        "yoctocolors-cjs": "^2.1.2"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@inquirer/select": {
      "version": "2.5.0",
      "resolved": "https://registry.npmmirror.com/@inquirer/select/-/select-2.5.0.tgz",
      "integrity": "sha512-YmDobTItPP3WcEI86GvPo+T2sRHkxxOq/kXmsBjHS5BVXUgvgZ5AfJjkvQvZr03T81NnI3KrrRuMzeuYUQRFOA==",
      "license": "MIT",
      "dependencies": {
        "@inquirer/core": "^9.1.0",
        "@inquirer/figures": "^1.0.5",
        "@inquirer/type": "^1.5.3",
        "ansi-escapes": "^4.3.2",
        "yoctocolors-cjs": "^2.1.2"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@inquirer/type": {
      "version": "1.5.5",
      "resolved": "https://registry.npmmirror.com/@inquirer/type/-/type-1.5.5.tgz",
      "integrity": "sha512-MzICLu4yS7V8AA61sANROZ9vT1H3ooca5dSmI1FjZkzq7o/koMsRfQSzRtFo+F3Ao4Sf1C0bpLKejpKB/+j6MA==",
      "license": "MIT",
      "dependencies": {
        "mute-stream": "^1.0.0"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@jridgewell/resolve-uri": {
      "version": "3.1.2",
      "resolved": "https://registry.npmmirror.com/@jridgewell/resolve-uri/-/resolve-uri-3.1.2.tgz",
      "integrity": "sha512-bRISgCIjP20/tbWSPWMEi54QVPRZExkuD9lJL+UIxUKtwVJA8wW1Trb1jMs1RFXo1CBTNZ/5hpC9QvmKWdopKw==",
      "dev": true,
      "engines": {
        "node": ">=6.0.0"
      }
    },
    "node_modules/@jridgewell/sourcemap-codec": {
      "version": "1.5.5",
      "resolved": "https://registry.npmmirror.com/@jridgewell/sourcemap-codec/-/sourcemap-codec-1.5.5.tgz",
      "integrity": "sha512-cYQ9310grqxueWbl+WuIUIaiUaDcj7WOq5fVhEljNVgRfOUhY9fy2zTvfoqWsnebh8Sl70VScFbICvJnLKB0Og==",
      "dev": true
    },
    "node_modules/@jridgewell/trace-mapping": {
      "version": "0.3.9",
      "resolved": "https://registry.npmmirror.com/@jridgewell/trace-mapping/-/trace-mapping-0.3.9.tgz",
      "integrity": "sha512-3Belt6tdc8bPgAtbcmdtNJlirVoTmEb5e2gC94PnkwEW9jI6CAHUeoG85tjWP5WquqfavoMtMwiG4P926ZKKuQ==",
      "dev": true,
      "dependencies": {
        "@jridgewell/resolve-uri": "^3.0.3",
        "@jridgewell/sourcemap-codec": "^1.4.10"
      }
    },
    "node_modules/@modelcontextprotocol/sdk": {
      "version": "1.18.2",
      "resolved": "https://registry.npmmirror.com/@modelcontextprotocol/sdk/-/sdk-1.18.2.tgz",
      "integrity": "sha512-beedclIvFcCnPrYgHsylqiYJVJ/CI47Vyc4tY8no1/Li/O8U4BTlJfy6ZwxkYwx+Mx10nrgwSVrA7VBbhh4slg==",
      "license": "MIT",
      "dependencies": {
        "ajv": "^6.12.6",
        "content-type": "^1.0.5",
        "cors": "^2.8.5",
        "cross-spawn": "^7.0.5",
        "eventsource": "^3.0.2",
        "eventsource-parser": "^3.0.0",
        "express": "^5.0.1",
        "express-rate-limit": "^7.5.0",
        "pkce-challenge": "^5.0.0",
        "raw-body": "^3.0.0",
        "zod": "^3.23.8",
        "zod-to-json-schema": "^3.24.1"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@tsconfig/node10": {
      "version": "1.0.11",
      "resolved": "https://registry.npmmirror.com/@tsconfig/node10/-/node10-1.0.11.tgz",
      "integrity": "sha512-DcRjDCujK/kCk/cUe8Xz8ZSpm8mS3mNNpta+jGCA6USEDfktlNvm1+IuZ9eTcDbNk41BHwpHHeW+N1lKCz4zOw==",
      "dev": true
    },
    "node_modules/@tsconfig/node12": {
      "version": "1.0.11",
      "resolved": "https://registry.npmmirror.com/@tsconfig/node12/-/node12-1.0.11.tgz",
      "integrity": "sha512-cqefuRsh12pWyGsIoBKJA9luFu3mRxCA+ORZvA4ktLSzIuCUtWVxGIuXigEwO5/ywWFMZ2QEGKWvkZG1zDMTag==",
      "dev": true
    },
    "node_modules/@tsconfig/node14": {
      "version": "1.0.3",
      "resolved": "https://registry.npmmirror.com/@tsconfig/node14/-/node14-1.0.3.tgz",
      "integrity": "sha512-ysT8mhdixWK6Hw3i1V2AeRqZ5WfXg1G43mqoYlM2nc6388Fq5jcXyr5mRsqViLx/GJYdoL0bfXD8nmF+Zn/Iow==",
      "dev": true
    },
    "node_modules/@tsconfig/node16": {
      "version": "1.0.4",
      "resolved": "https://registry.npmmirror.com/@tsconfig/node16/-/node16-1.0.4.tgz",
      "integrity": "sha512-vxhUy4J8lyeyinH7Azl1pdd43GJhZH/tP2weN8TntQblOY+A0XbT8DJk1/oCPuOOyg/Ja757rG0CgHcWC8OfMA==",
      "dev": true
    },
    "node_modules/@types/mute-stream": {
      "version": "0.0.4",
      "resolved": "https://registry.npmmirror.com/@types/mute-stream/-/mute-stream-0.0.4.tgz",
      "integrity": "sha512-CPM9nzrCPPJHQNA9keH9CVkVI+WR5kMa+7XEs5jcGQ0VoAGnLv242w8lIVgwAEfmE4oufJRaTc9PNLQl0ioAow==",
      "license": "MIT",
      "dependencies": {
        "@types/node": "*"
      }
    },
    "node_modules/@types/node": {
      "version": "20.19.17",
      "resolved": "https://registry.npmmirror.com/@types/node/-/node-20.19.17.tgz",
      "integrity": "sha512-gfehUI8N1z92kygssiuWvLiwcbOB3IRktR6hTDgJlXMYh5OvkPSRmgfoBUmfZt+vhwJtX7v1Yw4KvvAf7c5QKQ==",
      "dependencies": {
        "undici-types": "~6.21.0"
      }
    },
    "node_modules/@types/wrap-ansi": {
      "version": "3.0.0",
      "resolved": "https://registry.npmmirror.com/@types/wrap-ansi/-/wrap-ansi-3.0.0.tgz",
      "integrity": "sha512-ltIpx+kM7g/MLRZfkbL7EsCEjfzCcScLpkg37eXEtx5kmrAKBkTJwd1GIAjDSL8wTpM6Hzn5YO4pSb91BEwu1g==",
      "license": "MIT"
    },
    "node_modules/accepts": {
      "version": "2.0.0",
      "resolved": "https://registry.npmmirror.com/accepts/-/accepts-2.0.0.tgz",
      "integrity": "sha512-5cvg6CtKwfgdmVqY1WIiXKc3Q1bkRqGLi+2W/6ao+6Y7gu/RCwRuAhGEzh5B4KlszSuTLgZYuqFqo5bImjNKng==",
      "license": "MIT",
      "dependencies": {
        "mime-types": "^3.0.0",
        "negotiator": "^1.0.0"
      },
      "engines": {
        "node": ">= 0.6"
      }
    },
    "node_modules/acorn": {
      "version": "8.15.0",
      "resolved": "https://registry.npmmirror.com/acorn/-/acorn-8.15.0.tgz",
      "integrity": "sha512-NZyJarBfL7nWwIq+FDL6Zp/yHEhePMNnnJ0y3qfieCrmNvYct8uvtiV41UvlSe6apAfk0fY1FbWx+NwfmpvtTg==",
      "dev": true,
      "bin": {
        "acorn": "bin/acorn"
      },
      "engines": {
        "node": ">=0.4.0"
      }
    },
    "node_modules/acorn-walk": {
      "version": "8.3.4",
      "resolved": "https://registry.npmmirror.com/acorn-walk/-/acorn-walk-8.3.4.tgz",
      "integrity": "sha512-ueEepnujpqee2o5aIYnvHU6C0A42MNdsIDeqy5BydrkuC5R1ZuUFnm27EeFJGoEHJQgn3uleRvmTXaJgfXbt4g==",
      "dev": true,
      "dependencies": {
        "acorn": "^8.11.0"
      },
      "engines": {
        "node": ">=0.4.0"
      }
    },
    "node_modules/ajv": {
      "version": "6.12.6",
      "resolved": "https://registry.npmmirror.com/ajv/-/ajv-6.12.6.tgz",
      "integrity": "sha512-j3fVLgvTo527anyYyJOGTYJbG+vnnQYvE0m5mmkc1TK+nxAppkCLMIL0aZ4dblVCNoGShhm+kzE4ZUykBoMg4g==",
      "license": "MIT",
      "dependencies": {
        "fast-deep-equal": "^3.1.1",
        "fast-json-stable-stringify": "^2.0.0",
        "json-schema-traverse": "^0.4.1",
        "uri-js": "^4.2.2"
      },
      "funding": {
        "type": "github",
        "url": "https://github.com/sponsors/epoberezkin"
      }
    },
    "node_modules/ansi-escapes": {
      "version": "4.3.2",
      "resolved": "https://registry.npmmirror.com/ansi-escapes/-/ansi-escapes-4.3.2.tgz",
      "integrity": "sha512-gKXj5ALrKWQLsYG9jlTRmR/xKluxHV+Z9QEwNIgCfM1/uwPMCuzVVnh5mwTd+OuBZcwSIMbqssNWRm1lE51QaQ==",
      "license": "MIT",
      "dependencies": {
        "type-fest": "^0.21.3"
      },
      "engines": {
        "node": ">=8"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/ansi-regex": {
      "version": "5.0.1",
      "resolved": "https://registry.npmmirror.com/ansi-regex/-/ansi-regex-5.0.1.tgz",
      "integrity": "sha512-quJQXlTSUGL2LH9SUXo8VwsY4soanhgo6LNSm84E1LBcE8s3O0wpdiRzyR9z/ZZJMlMWv37qOOb9pdJlMUEKFQ==",
      "license": "MIT",
      "engines": {
        "node": ">=8"
      }
    },
    "node_modules/ansi-styles": {
      "version": "4.3.0",
      "resolved": "https://registry.npmmirror.com/ansi-styles/-/ansi-styles-4.3.0.tgz",
      "integrity": "sha512-zbB9rCJAT1rbjiVDb2hqKFHNYLxgtk8NURxZ3IZwD3F6NtxbXZQCnnSi1Lkx+IDohdPlFp222wVALIheZJQSEg==",
      "license": "MIT",
      "dependencies": {
        "color-convert": "^2.0.1"
      },
      "engines": {
        "node": ">=8"
      },
      "funding": {
        "url": "https://github.com/chalk/ansi-styles?sponsor=1"
      }
    },
    "node_modules/arg": {
      "version": "4.1.3",
      "resolved": "https://registry.npmmirror.com/arg/-/arg-4.1.3.tgz",
      "integrity": "sha512-58S9QDqG0Xx27YwPSt9fJxivjYl432YCwfDMfZ+71RAqUrZef7LrKQZ3LHLOwCS4FLNBplP533Zx895SeOCHvA==",
      "dev": true
    },
    "node_modules/body-parser": {
      "version": "2.2.0",
      "resolved": "https://registry.npmmirror.com/body-parser/-/body-parser-2.2.0.tgz",
      "integrity": "sha512-02qvAaxv8tp7fBa/mw1ga98OGm+eCbqzJOKoRt70sLmfEEi+jyBYVTDGfCL/k06/4EMk/z01gCe7HoCH/f2LTg==",
      "license": "MIT",
      "dependencies": {
        "bytes": "^3.1.2",
        "content-type": "^1.0.5",
        "debug": "^4.4.0",
        "http-errors": "^2.0.0",
        "iconv-lite": "^0.6.3",
        "on-finished": "^2.4.1",
        "qs": "^6.14.0",
        "raw-body": "^3.0.0",
        "type-is": "^2.0.0"
      },
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/bytes": {
      "version": "3.1.2",
      "resolved": "https://registry.npmmirror.com/bytes/-/bytes-3.1.2.tgz",
      "integrity": "sha512-/Nf7TyzTx6S3yRJObOAV7956r8cr2+Oj8AC5dt8wSP3BQAoeX58NoHyCU8P8zGkNXStjTSi6fzO6F0pBdcYbEg==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.8"
      }
    },
    "node_modules/call-bind-apply-helpers": {
      "version": "1.0.2",
      "resolved": "https://registry.npmmirror.com/call-bind-apply-helpers/-/call-bind-apply-helpers-1.0.2.tgz",
      "integrity": "sha512-Sp1ablJ0ivDkSzjcaJdxEunN5/XvksFJ2sMBFfq6x0ryhQV/2b/KwFe21cMpmHtPOSij8K99/wSfoEuTObmuMQ==",
      "license": "MIT",
      "dependencies": {
        "es-errors": "^1.3.0",
        "function-bind": "^1.1.2"
      },
      "engines": {
        "node": ">= 0.4"
      }
    },
    "node_modules/call-bound": {
      "version": "1.0.4",
      "resolved": "https://registry.npmmirror.com/call-bound/-/call-bound-1.0.4.tgz",
      "integrity": "sha512-+ys997U96po4Kx/ABpBCqhA9EuxJaQWDQg7295H4hBphv3IZg0boBKuwYpt4YXp6MZ5AmZQnU/tyMTlRpaSejg==",
      "license": "MIT",
      "dependencies": {
        "call-bind-apply-helpers": "^1.0.2",
        "get-intrinsic": "^1.3.0"
      },
      "engines": {
        "node": ">= 0.4"
      },
      "funding": {
        "url": "https://github.com/sponsors/ljharb"
      }
    },
    "node_modules/chalk": {
      "version": "5.6.2",
      "resolved": "https://registry.npmmirror.com/chalk/-/chalk-5.6.2.tgz",
      "integrity": "sha512-7NzBL0rN6fMUW+f7A6Io4h40qQlG+xGmtMxfbnH/K7TAtt8JQWVQK+6g0UXKMeVJoyV5EkkNsErQ8pVD3bLHbA==",
      "license": "MIT",
      "engines": {
        "node": "^12.17.0 || ^14.13 || >=16.0.0"
      },
      "funding": {
        "url": "https://github.com/chalk/chalk?sponsor=1"
      }
    },
    "node_modules/chardet": {
      "version": "0.7.0",
      "resolved": "https://registry.npmmirror.com/chardet/-/chardet-0.7.0.tgz",
      "integrity": "sha512-mT8iDcrh03qDGRRmoA2hmBJnxpllMR+0/0qlzjqZES6NdiWDcZkCNAk4rPFZ9Q85r27unkiNNg8ZOiwZXBHwcA==",
      "license": "MIT"
    },
    "node_modules/cli-cursor": {
      "version": "5.0.0",
      "resolved": "https://registry.npmmirror.com/cli-cursor/-/cli-cursor-5.0.0.tgz",
      "integrity": "sha512-aCj4O5wKyszjMmDT4tZj93kxyydN/K5zPWSCe6/0AV/AA1pqe5ZBIw0a2ZfPQV7lL5/yb5HsUreJ6UFAF1tEQw==",
      "license": "MIT",
      "dependencies": {
        "restore-cursor": "^5.0.0"
      },
      "engines": {
        "node": ">=18"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/cli-spinners": {
      "version": "3.3.0",
      "resolved": "https://registry.npmmirror.com/cli-spinners/-/cli-spinners-3.3.0.tgz",
      "integrity": "sha512-/+40ljC3ONVnYIttjMWrlL51nItDAbBrq2upN8BPyvGU/2n5Oxw3tbNwORCaNuNqLJnxGqOfjUuhsv7l5Q4IsQ==",
      "license": "MIT",
      "engines": {
        "node": ">=18.20"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/cli-width": {
      "version": "4.1.0",
      "resolved": "https://registry.npmmirror.com/cli-width/-/cli-width-4.1.0.tgz",
      "integrity": "sha512-ouuZd4/dm2Sw5Gmqy6bGyNNNe1qt9RpmxveLSO7KcgsTnU7RXfsw+/bukWGo1abgBiMAic068rclZsO4IWmmxQ==",
      "license": "ISC",
      "engines": {
        "node": ">= 12"
      }
    },
    "node_modules/color-convert": {
      "version": "2.0.1",
      "resolved": "https://registry.npmmirror.com/color-convert/-/color-convert-2.0.1.tgz",
      "integrity": "sha512-RRECPsj7iu/xb5oKYcsFHSppFNnsj/52OVTRKb4zP5onXwVF3zVmmToNcOfGC+CRDpfK/U584fMg38ZHCaElKQ==",
      "license": "MIT",
      "dependencies": {
        "color-name": "~1.1.4"
      },
      "engines": {
        "node": ">=7.0.0"
      }
    },
    "node_modules/color-name": {
      "version": "1.1.4",
      "resolved": "https://registry.npmmirror.com/color-name/-/color-name-1.1.4.tgz",
      "integrity": "sha512-dOy+3AuW3a2wNbZHIuMZpTcgjGuLU/uBL/ubcZF9OXbDo8ff4O8yVp5Bf0efS8uEoYo5q4Fx7dY9OgQGXgAsQA==",
      "license": "MIT"
    },
    "node_modules/content-disposition": {
      "version": "1.0.0",
      "resolved": "https://registry.npmmirror.com/content-disposition/-/content-disposition-1.0.0.tgz",
      "integrity": "sha512-Au9nRL8VNUut/XSzbQA38+M78dzP4D+eqg3gfJHMIHHYa3bg067xj1KxMUWj+VULbiZMowKngFFbKczUrNJ1mg==",
      "license": "MIT",
      "dependencies": {
        "safe-buffer": "5.2.1"
      },
      "engines": {
        "node": ">= 0.6"
      }
    },
    "node_modules/content-type": {
      "version": "1.0.5",
      "resolved": "https://registry.npmmirror.com/content-type/-/content-type-1.0.5.tgz",
      "integrity": "sha512-nTjqfcBFEipKdXCv4YDQWCfmcLZKm81ldF0pAopTvyrFGVbcR6P/VAAd5G7N+0tTr8QqiU0tFadD6FK4NtJwOA==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.6"
      }
    },
    "node_modules/cookie": {
      "version": "0.7.2",
      "resolved": "https://registry.npmmirror.com/cookie/-/cookie-0.7.2.tgz",
      "integrity": "sha512-yki5XnKuf750l50uGTllt6kKILY4nQ1eNIQatoXEByZ5dWgnKqbnqmTrBE5B4N7lrMJKQ2ytWMiTO2o0v6Ew/w==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.6"
      }
    },
    "node_modules/cookie-signature": {
      "version": "1.2.2",
      "resolved": "https://registry.npmmirror.com/cookie-signature/-/cookie-signature-1.2.2.tgz",
      "integrity": "sha512-D76uU73ulSXrD1UXF4KE2TMxVVwhsnCgfAyTg9k8P6KGZjlXKrOLe4dJQKI3Bxi5wjesZoFXJWElNWBjPZMbhg==",
      "license": "MIT",
      "engines": {
        "node": ">=6.6.0"
      }
    },
    "node_modules/core-util-is": {
      "version": "1.0.3",
      "resolved": "https://registry.npmmirror.com/core-util-is/-/core-util-is-1.0.3.tgz",
      "integrity": "sha512-ZQBvi1DcpJ4GDqanjucZ2Hj3wEO5pZDS89BWbkcrvdxksJorwUDDZamX9ldFkp9aw2lmBDLgkObEA4DWNJ9FYQ==",
      "license": "MIT"
    },
    "node_modules/cors": {
      "version": "2.8.5",
      "resolved": "https://registry.npmmirror.com/cors/-/cors-2.8.5.tgz",
      "integrity": "sha512-KIHbLJqu73RGr/hnbrO9uBeixNGuvSQjul/jdFvS/KFSIH1hWVd1ng7zOHx+YrEfInLG7q4n6GHQ9cDtxv/P6g==",
      "license": "MIT",
      "dependencies": {
        "object-assign": "^4",
        "vary": "^1"
      },
      "engines": {
        "node": ">= 0.10"
      }
    },
    "node_modules/create-require": {
      "version": "1.1.1",
      "resolved": "https://registry.npmmirror.com/create-require/-/create-require-1.1.1.tgz",
      "integrity": "sha512-dcKFX3jn0MpIaXjisoRvexIJVEKzaq7z2rZKxf+MSr9TkdmHmsU4m2lcLojrj/FHl8mk5VxMmYA+ftRkP/3oKQ==",
      "dev": true
    },
    "node_modules/cross-spawn": {
      "version": "7.0.6",
      "resolved": "https://registry.npmmirror.com/cross-spawn/-/cross-spawn-7.0.6.tgz",
      "integrity": "sha512-uV2QOWP2nWzsy2aMp8aRibhi9dlzF5Hgh5SHaB9OiTGEyDTiJJyx0uy51QXdyWbtAHNua4XJzUKca3OzKUd3vA==",
      "license": "MIT",
      "dependencies": {
        "path-key": "^3.1.0",
        "shebang-command": "^2.0.0",
        "which": "^2.0.1"
      },
      "engines": {
        "node": ">= 8"
      }
    },
    "node_modules/debug": {
      "version": "4.4.3",
      "resolved": "https://registry.npmmirror.com/debug/-/debug-4.4.3.tgz",
      "integrity": "sha512-RGwwWnwQvkVfavKVt22FGLw+xYSdzARwm0ru6DhTVA3umU5hZc28V3kO4stgYryrTlLpuvgI9GiijltAjNbcqA==",
      "license": "MIT",
      "dependencies": {
        "ms": "^2.1.3"
      },
      "engines": {
        "node": ">=6.0"
      },
      "peerDependenciesMeta": {
        "supports-color": {
          "optional": true
        }
      }
    },
    "node_modules/depd": {
      "version": "2.0.0",
      "resolved": "https://registry.npmmirror.com/depd/-/depd-2.0.0.tgz",
      "integrity": "sha512-g7nH6P6dyDioJogAAGprGpCtVImJhpPk/roCzdb3fIh61/s/nPsfR6onyMwkCAR/OlC3yBC0lESvUoQEAssIrw==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.8"
      }
    },
    "node_modules/diff": {
      "version": "4.0.2",
      "resolved": "https://registry.npmmirror.com/diff/-/diff-4.0.2.tgz",
      "integrity": "sha512-58lmxKSA4BNyLz+HHMUzlOEpg09FV+ev6ZMe3vJihgdxzgcwZ8VoEEPmALCZG9LmqfVoNMMKpttIYTVG6uDY7A==",
      "dev": true,
      "engines": {
        "node": ">=0.3.1"
      }
    },
    "node_modules/docx": {
      "version": "9.5.1",
      "resolved": "https://registry.npmmirror.com/docx/-/docx-9.5.1.tgz",
      "integrity": "sha512-ABDI7JEirFD2+bHhOBlsGZxaG1UgZb2M/QMKhLSDGgVNhxDesTCDcP+qoDnDGjZ4EOXTRfUjUgwHVuZ6VSTfWQ==",
      "license": "MIT",
      "dependencies": {
        "@types/node": "^24.0.1",
        "hash.js": "^1.1.7",
        "jszip": "^3.10.1",
        "nanoid": "^5.1.3",
        "xml": "^1.0.1",
        "xml-js": "^1.6.8"
      },
      "engines": {
        "node": ">=10"
      }
    },
    "node_modules/docx/node_modules/@types/node": {
      "version": "24.10.0",
      "resolved": "https://registry.npmmirror.com/@types/node/-/node-24.10.0.tgz",
      "integrity": "sha512-qzQZRBqkFsYyaSWXuEHc2WR9c0a0CXwiE5FWUvn7ZM+vdy1uZLfCunD38UzhuB7YN/J11ndbDBcTmOdxJo9Q7A==",
      "license": "MIT",
      "dependencies": {
        "undici-types": "~7.16.0"
      }
    },
    "node_modules/docx/node_modules/undici-types": {
      "version": "7.16.0",
      "resolved": "https://registry.npmmirror.com/undici-types/-/undici-types-7.16.0.tgz",
      "integrity": "sha512-Zz+aZWSj8LE6zoxD+xrjh4VfkIG8Ya6LvYkZqtUQGJPZjYl53ypCaUwWqo7eI0x66KBGeRo+mlBEkMSeSZ38Nw==",
      "license": "MIT"
    },
    "node_modules/dunder-proto": {
      "version": "1.0.1",
      "resolved": "https://registry.npmmirror.com/dunder-proto/-/dunder-proto-1.0.1.tgz",
      "integrity": "sha512-KIN/nDJBQRcXw0MLVhZE9iQHmG68qAVIBg9CqmUYjmQIhgij9U5MFvrqkUL5FbtyyzZuOeOt0zdeRe4UY7ct+A==",
      "license": "MIT",
      "dependencies": {
        "call-bind-apply-helpers": "^1.0.1",
        "es-errors": "^1.3.0",
        "gopd": "^1.2.0"
      },
      "engines": {
        "node": ">= 0.4"
      }
    },
    "node_modules/ee-first": {
      "version": "1.1.1",
      "resolved": "https://registry.npmmirror.com/ee-first/-/ee-first-1.1.1.tgz",
      "integrity": "sha512-WMwm9LhRUo+WUaRN+vRuETqG89IgZphVSNkdFgeb6sS/E4OrDIN7t48CAewSHXc6C8lefD8KKfr5vY61brQlow==",
      "license": "MIT"
    },
    "node_modules/emoji-regex": {
      "version": "8.0.0",
      "resolved": "https://registry.npmmirror.com/emoji-regex/-/emoji-regex-8.0.0.tgz",
      "integrity": "sha512-MSjYzcWNOA0ewAHpz0MxpYFvwg6yjy1NG3xteoqz644VCo/RPgnr1/GGt+ic3iJTzQ8Eu3TdM14SawnVUmGE6A==",
      "license": "MIT"
    },
    "node_modules/encodeurl": {
      "version": "2.0.0",
      "resolved": "https://registry.npmmirror.com/encodeurl/-/encodeurl-2.0.0.tgz",
      "integrity": "sha512-Q0n9HRi4m6JuGIV1eFlmvJB7ZEVxu93IrMyiMsGC0lrMJMWzRgx6WGquyfQgZVb31vhGgXnfmPNNXmxnOkRBrg==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.8"
      }
    },
    "node_modules/es-define-property": {
      "version": "1.0.1",
      "resolved": "https://registry.npmmirror.com/es-define-property/-/es-define-property-1.0.1.tgz",
      "integrity": "sha512-e3nRfgfUZ4rNGL232gUgX06QNyyez04KdjFrF+LTRoOXmrOgFKDg4BCdsjW8EnT69eqdYGmRpJwiPVYNrCaW3g==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.4"
      }
    },
    "node_modules/es-errors": {
      "version": "1.3.0",
      "resolved": "https://registry.npmmirror.com/es-errors/-/es-errors-1.3.0.tgz",
      "integrity": "sha512-Zf5H2Kxt2xjTvbJvP2ZWLEICxA6j+hAmMzIlypy4xcBg1vKVnx89Wy0GbS+kf5cwCVFFzdCFh2XSCFNULS6csw==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.4"
      }
    },
    "node_modules/es-object-atoms": {
      "version": "1.1.1",
      "resolved": "https://registry.npmmirror.com/es-object-atoms/-/es-object-atoms-1.1.1.tgz",
      "integrity": "sha512-FGgH2h8zKNim9ljj7dankFPcICIK9Cp5bm+c2gQSYePhpaG5+esrLODihIorn+Pe6FGJzWhXQotPv73jTaldXA==",
      "license": "MIT",
      "dependencies": {
        "es-errors": "^1.3.0"
      },
      "engines": {
        "node": ">= 0.4"
      }
    },
    "node_modules/escape-html": {
      "version": "1.0.3",
      "resolved": "https://registry.npmmirror.com/escape-html/-/escape-html-1.0.3.tgz",
      "integrity": "sha512-NiSupZ4OeuGwr68lGIeym/ksIZMJodUGOSCZ/FSnTxcrekbvqrgdUxlJOMpijaKZVjAJrWrGs/6Jy8OMuyj9ow==",
      "license": "MIT"
    },
    "node_modules/etag": {
      "version": "1.8.1",
      "resolved": "https://registry.npmmirror.com/etag/-/etag-1.8.1.tgz",
      "integrity": "sha512-aIL5Fx7mawVa300al2BnEE4iNvo1qETxLrPI/o05L7z6go7fCw1J6EQmbK4FmJ2AS7kgVF/KEZWufBfdClMcPg==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.6"
      }
    },
    "node_modules/eventsource": {
      "version": "3.0.7",
      "resolved": "https://registry.npmmirror.com/eventsource/-/eventsource-3.0.7.tgz",
      "integrity": "sha512-CRT1WTyuQoD771GW56XEZFQ/ZoSfWid1alKGDYMmkt2yl8UXrVR4pspqWNEcqKvVIzg6PAltWjxcSSPrboA4iA==",
      "license": "MIT",
      "dependencies": {
        "eventsource-parser": "^3.0.1"
      },
      "engines": {
        "node": ">=18.0.0"
      }
    },
    "node_modules/eventsource-parser": {
      "version": "3.0.6",
      "resolved": "https://registry.npmmirror.com/eventsource-parser/-/eventsource-parser-3.0.6.tgz",
      "integrity": "sha512-Vo1ab+QXPzZ4tCa8SwIHJFaSzy4R6SHf7BY79rFBDf0idraZWAkYrDjDj8uWaSm3S2TK+hJ7/t1CEmZ7jXw+pg==",
      "license": "MIT",
      "engines": {
        "node": ">=18.0.0"
      }
    },
    "node_modules/express": {
      "version": "5.1.0",
      "resolved": "https://registry.npmmirror.com/express/-/express-5.1.0.tgz",
      "integrity": "sha512-DT9ck5YIRU+8GYzzU5kT3eHGA5iL+1Zd0EutOmTE9Dtk+Tvuzd23VBU+ec7HPNSTxXYO55gPV/hq4pSBJDjFpA==",
      "license": "MIT",
      "dependencies": {
        "accepts": "^2.0.0",
        "body-parser": "^2.2.0",
        "content-disposition": "^1.0.0",
        "content-type": "^1.0.5",
        "cookie": "^0.7.1",
        "cookie-signature": "^1.2.1",
        "debug": "^4.4.0",
        "encodeurl": "^2.0.0",
        "escape-html": "^1.0.3",
        "etag": "^1.8.1",
        "finalhandler": "^2.1.0",
        "fresh": "^2.0.0",
        "http-errors": "^2.0.0",
        "merge-descriptors": "^2.0.0",
        "mime-types": "^3.0.0",
        "on-finished": "^2.4.1",
        "once": "^1.4.0",
        "parseurl": "^1.3.3",
        "proxy-addr": "^2.0.7",
        "qs": "^6.14.0",
        "range-parser": "^1.2.1",
        "router": "^2.2.0",
        "send": "^1.1.0",
        "serve-static": "^2.2.0",
        "statuses": "^2.0.1",
        "type-is": "^2.0.1",
        "vary": "^1.1.2"
      },
      "engines": {
        "node": ">= 18"
      },
      "funding": {
        "type": "opencollective",
        "url": "https://opencollective.com/express"
      }
    },
    "node_modules/express-rate-limit": {
      "version": "7.5.1",
      "resolved": "https://registry.npmmirror.com/express-rate-limit/-/express-rate-limit-7.5.1.tgz",
      "integrity": "sha512-7iN8iPMDzOMHPUYllBEsQdWVB6fPDMPqwjBaFrgr4Jgr/+okjvzAy+UHlYYL/Vs0OsOrMkwS6PJDkFlJwoxUnw==",
      "license": "MIT",
      "engines": {
        "node": ">= 16"
      },
      "funding": {
        "url": "https://github.com/sponsors/express-rate-limit"
      },
      "peerDependencies": {
        "express": ">= 4.11"
      }
    },
    "node_modules/external-editor": {
      "version": "3.1.0",
      "resolved": "https://registry.npmmirror.com/external-editor/-/external-editor-3.1.0.tgz",
      "integrity": "sha512-hMQ4CX1p1izmuLYyZqLMO/qGNw10wSv9QDCPfzXfyFrOaCSSoRfqE1Kf1s5an66J5JZC62NewG+mK49jOCtQew==",
      "license": "MIT",
      "dependencies": {
        "chardet": "^0.7.0",
        "iconv-lite": "^0.4.24",
        "tmp": "^0.0.33"
      },
      "engines": {
        "node": ">=4"
      }
    },
    "node_modules/external-editor/node_modules/iconv-lite": {
      "version": "0.4.24",
      "resolved": "https://registry.npmmirror.com/iconv-lite/-/iconv-lite-0.4.24.tgz",
      "integrity": "sha512-v3MXnZAcvnywkTUEZomIActle7RXXeedOR31wwl7VlyoXO4Qi9arvSenNQWne1TcRwhCL1HwLI21bEqdpj8/rA==",
      "license": "MIT",
      "dependencies": {
        "safer-buffer": ">= 2.1.2 < 3"
      },
      "engines": {
        "node": ">=0.10.0"
      }
    },
    "node_modules/fast-deep-equal": {
      "version": "3.1.3",
      "resolved": "https://registry.npmmirror.com/fast-deep-equal/-/fast-deep-equal-3.1.3.tgz",
      "integrity": "sha512-f3qQ9oQy9j2AhBe/H9VC91wLmKBCCU/gDOnKNAYG5hswO7BLKj09Hc5HYNz9cGI++xlpDCIgDaitVs03ATR84Q==",
      "license": "MIT"
    },
    "node_modules/fast-json-stable-stringify": {
      "version": "2.1.0",
      "resolved": "https://registry.npmmirror.com/fast-json-stable-stringify/-/fast-json-stable-stringify-2.1.0.tgz",
      "integrity": "sha512-lhd/wF+Lk98HZoTCtlVraHtfh5XYijIjalXck7saUtuanSDyLMxnHhSXEDJqHxD7msR8D0uCmqlkwjCV8xvwHw==",
      "license": "MIT"
    },
    "node_modules/finalhandler": {
      "version": "2.1.0",
      "resolved": "https://registry.npmmirror.com/finalhandler/-/finalhandler-2.1.0.tgz",
      "integrity": "sha512-/t88Ty3d5JWQbWYgaOGCCYfXRwV1+be02WqYYlL6h0lEiUAMPM8o8qKGO01YIkOHzka2up08wvgYD0mDiI+q3Q==",
      "license": "MIT",
      "dependencies": {
        "debug": "^4.4.0",
        "encodeurl": "^2.0.0",
        "escape-html": "^1.0.3",
        "on-finished": "^2.4.1",
        "parseurl": "^1.3.3",
        "statuses": "^2.0.1"
      },
      "engines": {
        "node": ">= 0.8"
      }
    },
    "node_modules/forwarded": {
      "version": "0.2.0",
      "resolved": "https://registry.npmmirror.com/forwarded/-/forwarded-0.2.0.tgz",
      "integrity": "sha512-buRG0fpBtRHSTCOASe6hD258tEubFoRLb4ZNA6NxMVHNw2gOcwHo9wyablzMzOA5z9xA9L1KNjk/Nt6MT9aYow==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.6"
      }
    },
    "node_modules/fresh": {
      "version": "2.0.0",
      "resolved": "https://registry.npmmirror.com/fresh/-/fresh-2.0.0.tgz",
      "integrity": "sha512-Rx/WycZ60HOaqLKAi6cHRKKI7zxWbJ31MhntmtwMoaTeF7XFH9hhBp8vITaMidfljRQ6eYWCKkaTK+ykVJHP2A==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.8"
      }
    },
    "node_modules/function-bind": {
      "version": "1.1.2",
      "resolved": "https://registry.npmmirror.com/function-bind/-/function-bind-1.1.2.tgz",
      "integrity": "sha512-7XHNxH7qX9xG5mIwxkhumTox/MIRNcOgDrxWsMt2pAr23WHp6MrRlN7FBSFpCpr+oVO0F744iUgR82nJMfG2SA==",
      "license": "MIT",
      "funding": {
        "url": "https://github.com/sponsors/ljharb"
      }
    },
    "node_modules/get-east-asian-width": {
      "version": "1.4.0",
      "resolved": "https://registry.npmmirror.com/get-east-asian-width/-/get-east-asian-width-1.4.0.tgz",
      "integrity": "sha512-QZjmEOC+IT1uk6Rx0sX22V6uHWVwbdbxf1faPqJ1QhLdGgsRGCZoyaQBm/piRdJy/D2um6hM1UP7ZEeQ4EkP+Q==",
      "license": "MIT",
      "engines": {
        "node": ">=18"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/get-intrinsic": {
      "version": "1.3.0",
      "resolved": "https://registry.npmmirror.com/get-intrinsic/-/get-intrinsic-1.3.0.tgz",
      "integrity": "sha512-9fSjSaos/fRIVIp+xSJlE6lfwhES7LNtKaCBIamHsjr2na1BiABJPo0mOjjz8GJDURarmCPGqaiVg5mfjb98CQ==",
      "license": "MIT",
      "dependencies": {
        "call-bind-apply-helpers": "^1.0.2",
        "es-define-property": "^1.0.1",
        "es-errors": "^1.3.0",
        "es-object-atoms": "^1.1.1",
        "function-bind": "^1.1.2",
        "get-proto": "^1.0.1",
        "gopd": "^1.2.0",
        "has-symbols": "^1.1.0",
        "hasown": "^2.0.2",
        "math-intrinsics": "^1.1.0"
      },
      "engines": {
        "node": ">= 0.4"
      },
      "funding": {
        "url": "https://github.com/sponsors/ljharb"
      }
    },
    "node_modules/get-proto": {
      "version": "1.0.1",
      "resolved": "https://registry.npmmirror.com/get-proto/-/get-proto-1.0.1.tgz",
      "integrity": "sha512-sTSfBjoXBp89JvIKIefqw7U2CCebsc74kiY6awiGogKtoSGbgjYE/G/+l9sF3MWFPNc9IcoOC4ODfKHfxFmp0g==",
      "license": "MIT",
      "dependencies": {
        "dunder-proto": "^1.0.1",
        "es-object-atoms": "^1.0.0"
      },
      "engines": {
        "node": ">= 0.4"
      }
    },
    "node_modules/gopd": {
      "version": "1.2.0",
      "resolved": "https://registry.npmmirror.com/gopd/-/gopd-1.2.0.tgz",
      "integrity": "sha512-ZUKRh6/kUFoAiTAtTYPZJ3hw9wNxx+BIBOijnlG9PnrJsCcSjs1wyyD6vJpaYtgnzDrKYRSqf3OO6Rfa93xsRg==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.4"
      },
      "funding": {
        "url": "https://github.com/sponsors/ljharb"
      }
    },
    "node_modules/has-symbols": {
      "version": "1.1.0",
      "resolved": "https://registry.npmmirror.com/has-symbols/-/has-symbols-1.1.0.tgz",
      "integrity": "sha512-1cDNdwJ2Jaohmb3sg4OmKaMBwuC48sYni5HUw2DvsC8LjGTLK9h+eb1X6RyuOHe4hT0ULCW68iomhjUoKUqlPQ==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.4"
      },
      "funding": {
        "url": "https://github.com/sponsors/ljharb"
      }
    },
    "node_modules/hash.js": {
      "version": "1.1.7",
      "resolved": "https://registry.npmmirror.com/hash.js/-/hash.js-1.1.7.tgz",
      "integrity": "sha512-taOaskGt4z4SOANNseOviYDvjEJinIkRgmp7LbKP2YTTmVxWBl87s/uzK9r+44BclBSp2X7K1hqeNfz9JbBeXA==",
      "license": "MIT",
      "dependencies": {
        "inherits": "^2.0.3",
        "minimalistic-assert": "^1.0.1"
      }
    },
    "node_modules/hasown": {
      "version": "2.0.2",
      "resolved": "https://registry.npmmirror.com/hasown/-/hasown-2.0.2.tgz",
      "integrity": "sha512-0hJU9SCPvmMzIBdZFqNPXWa6dqh7WdH0cII9y+CyS8rG3nL48Bclra9HmKhVVUHyPWNH5Y7xDwAB7bfgSjkUMQ==",
      "license": "MIT",
      "dependencies": {
        "function-bind": "^1.1.2"
      },
      "engines": {
        "node": ">= 0.4"
      }
    },
    "node_modules/http-errors": {
      "version": "2.0.0",
      "resolved": "https://registry.npmmirror.com/http-errors/-/http-errors-2.0.0.tgz",
      "integrity": "sha512-FtwrG/euBzaEjYeRqOgly7G0qviiXoJWnvEH2Z1plBdXgbyjv34pHTSb9zoeHMyDy33+DWy5Wt9Wo+TURtOYSQ==",
      "license": "MIT",
      "dependencies": {
        "depd": "2.0.0",
        "inherits": "2.0.4",
        "setprototypeof": "1.2.0",
        "statuses": "2.0.1",
        "toidentifier": "1.0.1"
      },
      "engines": {
        "node": ">= 0.8"
      }
    },
    "node_modules/http-errors/node_modules/statuses": {
      "version": "2.0.1",
      "resolved": "https://registry.npmmirror.com/statuses/-/statuses-2.0.1.tgz",
      "integrity": "sha512-RwNA9Z/7PrK06rYLIzFMlaF+l73iwpzsqRIFgbMLbTcLD6cOao82TaWefPXQvB2fOC4AjuYSEndS7N/mTCbkdQ==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.8"
      }
    },
    "node_modules/iconv-lite": {
      "version": "0.6.3",
      "resolved": "https://registry.npmmirror.com/iconv-lite/-/iconv-lite-0.6.3.tgz",
      "integrity": "sha512-4fCk79wshMdzMp2rH06qWrJE4iolqLhCUH+OiuIgU++RB0+94NlDL81atO7GX55uUKueo0txHNtvEyI6D7WdMw==",
      "license": "MIT",
      "dependencies": {
        "safer-buffer": ">= 2.1.2 < 3.0.0"
      },
      "engines": {
        "node": ">=0.10.0"
      }
    },
    "node_modules/immediate": {
      "version": "3.0.6",
      "resolved": "https://registry.npmmirror.com/immediate/-/immediate-3.0.6.tgz",
      "integrity": "sha512-XXOFtyqDjNDAQxVfYxuF7g9Il/IbWmmlQg2MYKOH8ExIT1qg6xc4zyS3HaEEATgs1btfzxq15ciUiY7gjSXRGQ==",
      "license": "MIT"
    },
    "node_modules/inherits": {
      "version": "2.0.4",
      "resolved": "https://registry.npmmirror.com/inherits/-/inherits-2.0.4.tgz",
      "integrity": "sha512-k/vGaX4/Yla3WzyMCvTQOXYeIHvqOKtnqBduzTHpzpQZzAskKMhZ2K+EnBiSM9zGSoIFeMpXKxa4dYeZIQqewQ==",
      "license": "ISC"
    },
    "node_modules/ipaddr.js": {
      "version": "1.9.1",
      "resolved": "https://registry.npmmirror.com/ipaddr.js/-/ipaddr.js-1.9.1.tgz",
      "integrity": "sha512-0KI/607xoxSToH7GjN1FfSbLoU0+btTicjsQSWQlh/hZykN8KpmMf7uYwPW3R+akZ6R/w18ZlXSHBYXiYUPO3g==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.10"
      }
    },
    "node_modules/is-fullwidth-code-point": {
      "version": "3.0.0",
      "resolved": "https://registry.npmmirror.com/is-fullwidth-code-point/-/is-fullwidth-code-point-3.0.0.tgz",
      "integrity": "sha512-zymm5+u+sCsSWyD9qNaejV3DFvhCKclKdizYaJUuHA83RLjb7nSuGnddCHGv0hk+KY7BMAlsWeK4Ueg6EV6XQg==",
      "license": "MIT",
      "engines": {
        "node": ">=8"
      }
    },
    "node_modules/is-interactive": {
      "version": "2.0.0",
      "resolved": "https://registry.npmmirror.com/is-interactive/-/is-interactive-2.0.0.tgz",
      "integrity": "sha512-qP1vozQRI+BMOPcjFzrjXuQvdak2pHNUMZoeG2eRbiSqyvbEf/wQtEOTOX1guk6E3t36RkaqiSt8A/6YElNxLQ==",
      "license": "MIT",
      "engines": {
        "node": ">=12"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/is-promise": {
      "version": "4.0.0",
      "resolved": "https://registry.npmmirror.com/is-promise/-/is-promise-4.0.0.tgz",
      "integrity": "sha512-hvpoI6korhJMnej285dSg6nu1+e6uxs7zG3BYAm5byqDsgJNWwxzM6z6iZiAgQR4TJ30JmBTOwqZUw3WlyH3AQ==",
      "license": "MIT"
    },
    "node_modules/is-unicode-supported": {
      "version": "2.1.0",
      "resolved": "https://registry.npmmirror.com/is-unicode-supported/-/is-unicode-supported-2.1.0.tgz",
      "integrity": "sha512-mE00Gnza5EEB3Ds0HfMyllZzbBrmLOX3vfWoj9A9PEnTfratQ/BcaJOuMhnkhjXvb2+FkY3VuHqtAGpTPmglFQ==",
      "license": "MIT",
      "engines": {
        "node": ">=18"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/isarray": {
      "version": "1.0.0",
      "resolved": "https://registry.npmmirror.com/isarray/-/isarray-1.0.0.tgz",
      "integrity": "sha512-VLghIWNM6ELQzo7zwmcg0NmTVyWKYjvIeM83yjp0wRDTmUnrM678fQbcKBo6n2CJEF0szoG//ytg+TKla89ALQ==",
      "license": "MIT"
    },
    "node_modules/isexe": {
      "version": "2.0.0",
      "resolved": "https://registry.npmmirror.com/isexe/-/isexe-2.0.0.tgz",
      "integrity": "sha512-RHxMLp9lnKHGHRng9QFhRCMbYAcVpn69smSGcq3f36xjgVVWThj4qqLbTLlq7Ssj8B+fIQ1EuCEGI2lKsyQeIw==",
      "license": "ISC"
    },
    "node_modules/json-schema-to-ts": {
      "version": "3.1.1",
      "resolved": "https://registry.npmmirror.com/json-schema-to-ts/-/json-schema-to-ts-3.1.1.tgz",
      "integrity": "sha512-+DWg8jCJG2TEnpy7kOm/7/AxaYoaRbjVB4LFZLySZlWn8exGs3A4OLJR966cVvU26N7X9TWxl+Jsw7dzAqKT6g==",
      "dependencies": {
        "@babel/runtime": "^7.18.3",
        "ts-algebra": "^2.0.0"
      },
      "engines": {
        "node": ">=16"
      }
    },
    "node_modules/json-schema-traverse": {
      "version": "0.4.1",
      "resolved": "https://registry.npmmirror.com/json-schema-traverse/-/json-schema-traverse-0.4.1.tgz",
      "integrity": "sha512-xbbCH5dCYU5T8LcEhhuh7HJ88HXuW3qsI3Y0zOZFKfZEHcpWiHU/Jxzk629Brsab/mMiHQti9wMP+845RPe3Vg==",
      "license": "MIT"
    },
    "node_modules/jszip": {
      "version": "3.10.1",
      "resolved": "https://registry.npmmirror.com/jszip/-/jszip-3.10.1.tgz",
      "integrity": "sha512-xXDvecyTpGLrqFrvkrUSoxxfJI5AH7U8zxxtVclpsUtMCq4JQ290LY8AW5c7Ggnr/Y/oK+bQMbqK2qmtk3pN4g==",
      "license": "(MIT OR GPL-3.0-or-later)",
      "dependencies": {
        "lie": "~3.3.0",
        "pako": "~1.0.2",
        "readable-stream": "~2.3.6",
        "setimmediate": "^1.0.5"
      }
    },
    "node_modules/lie": {
      "version": "3.3.0",
      "resolved": "https://registry.npmmirror.com/lie/-/lie-3.3.0.tgz",
      "integrity": "sha512-UaiMJzeWRlEujzAuw5LokY1L5ecNQYZKfmyZ9L7wDHb/p5etKaxXhohBcrw0EYby+G/NA52vRSN4N39dxHAIwQ==",
      "license": "MIT",
      "dependencies": {
        "immediate": "~3.0.5"
      }
    },
    "node_modules/log-symbols": {
      "version": "7.0.1",
      "resolved": "https://registry.npmmirror.com/log-symbols/-/log-symbols-7.0.1.tgz",
      "integrity": "sha512-ja1E3yCr9i/0hmBVaM0bfwDjnGy8I/s6PP4DFp+yP+a+mrHO4Rm7DtmnqROTUkHIkqffC84YY7AeqX6oFk0WFg==",
      "license": "MIT",
      "dependencies": {
        "is-unicode-supported": "^2.0.0",
        "yoctocolors": "^2.1.1"
      },
      "engines": {
        "node": ">=18"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/make-error": {
      "version": "1.3.6",
      "resolved": "https://registry.npmmirror.com/make-error/-/make-error-1.3.6.tgz",
      "integrity": "sha512-s8UhlNe7vPKomQhC1qFelMokr/Sc3AgNbso3n74mVPA5LTZwkB9NlXf4XPamLxJE8h0gh73rM94xvwRT2CVInw==",
      "dev": true
    },
    "node_modules/math-intrinsics": {
      "version": "1.1.0",
      "resolved": "https://registry.npmmirror.com/math-intrinsics/-/math-intrinsics-1.1.0.tgz",
      "integrity": "sha512-/IXtbwEk5HTPyEwyKX6hGkYXxM9nbj64B+ilVJnC/R6B0pH5G4V3b0pVbL7DBj4tkhBAppbQUlf6F6Xl9LHu1g==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.4"
      }
    },
    "node_modules/media-typer": {
      "version": "1.1.0",
      "resolved": "https://registry.npmmirror.com/media-typer/-/media-typer-1.1.0.tgz",
      "integrity": "sha512-aisnrDP4GNe06UcKFnV5bfMNPBUw4jsLGaWwWfnH3v02GnBuXX2MCVn5RbrWo0j3pczUilYblq7fQ7Nw2t5XKw==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.8"
      }
    },
    "node_modules/merge-descriptors": {
      "version": "2.0.0",
      "resolved": "https://registry.npmmirror.com/merge-descriptors/-/merge-descriptors-2.0.0.tgz",
      "integrity": "sha512-Snk314V5ayFLhp3fkUREub6WtjBfPdCPY1Ln8/8munuLuiYhsABgBVWsozAG+MWMbVEvcdcpbi9R7ww22l9Q3g==",
      "license": "MIT",
      "engines": {
        "node": ">=18"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/mime-db": {
      "version": "1.54.0",
      "resolved": "https://registry.npmmirror.com/mime-db/-/mime-db-1.54.0.tgz",
      "integrity": "sha512-aU5EJuIN2WDemCcAp2vFBfp/m4EAhWJnUNSSw0ixs7/kXbd6Pg64EmwJkNdFhB8aWt1sH2CTXrLxo/iAGV3oPQ==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.6"
      }
    },
    "node_modules/mime-types": {
      "version": "3.0.1",
      "resolved": "https://registry.npmmirror.com/mime-types/-/mime-types-3.0.1.tgz",
      "integrity": "sha512-xRc4oEhT6eaBpU1XF7AjpOFD+xQmXNB5OVKwp4tqCuBpHLS/ZbBDrc07mYTDqVMg6PfxUjjNp85O6Cd2Z/5HWA==",
      "license": "MIT",
      "dependencies": {
        "mime-db": "^1.54.0"
      },
      "engines": {
        "node": ">= 0.6"
      }
    },
    "node_modules/mimic-function": {
      "version": "5.0.1",
      "resolved": "https://registry.npmmirror.com/mimic-function/-/mimic-function-5.0.1.tgz",
      "integrity": "sha512-VP79XUPxV2CigYP3jWwAUFSku2aKqBH7uTAapFWCBqutsbmDo96KY5o8uh6U+/YSIn5OxJnXp73beVkpqMIGhA==",
      "license": "MIT",
      "engines": {
        "node": ">=18"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/minimalistic-assert": {
      "version": "1.0.1",
      "resolved": "https://registry.npmmirror.com/minimalistic-assert/-/minimalistic-assert-1.0.1.tgz",
      "integrity": "sha512-UtJcAD4yEaGtjPezWuO9wC4nwUnVH/8/Im3yEHQP4b67cXlD/Qr9hdITCU1xDbSEXg2XKNaP8jsReV7vQd00/A==",
      "license": "ISC"
    },
    "node_modules/ms": {
      "version": "2.1.3",
      "resolved": "https://registry.npmmirror.com/ms/-/ms-2.1.3.tgz",
      "integrity": "sha512-6FlzubTLZG3J2a/NVCAleEhjzq5oxgHyaCU9yYXvcLsvoVaHJq/s5xXI6/XXP6tz7R9xAOtHnSO/tXtF3WRTlA==",
      "license": "MIT"
    },
    "node_modules/mute-stream": {
      "version": "1.0.0",
      "resolved": "https://registry.npmmirror.com/mute-stream/-/mute-stream-1.0.0.tgz",
      "integrity": "sha512-avsJQhyd+680gKXyG/sQc0nXaC6rBkPOfyHYcFb9+hdkqQkR9bdnkJ0AMZhke0oesPqIO+mFFJ+IdBc7mst4IA==",
      "license": "ISC",
      "engines": {
        "node": "^14.17.0 || ^16.13.0 || >=18.0.0"
      }
    },
    "node_modules/nanoid": {
      "version": "5.1.6",
      "resolved": "https://registry.npmmirror.com/nanoid/-/nanoid-5.1.6.tgz",
      "integrity": "sha512-c7+7RQ+dMB5dPwwCp4ee1/iV/q2P6aK1mTZcfr1BTuVlyW9hJYiMPybJCcnBlQtuSmTIWNeazm/zqNoZSSElBg==",
      "funding": [
        {
          "type": "github",
          "url": "https://github.com/sponsors/ai"
        }
      ],
      "license": "MIT",
      "bin": {
        "nanoid": "bin/nanoid.js"
      },
      "engines": {
        "node": "^18 || >=20"
      }
    },
    "node_modules/negotiator": {
      "version": "1.0.0",
      "resolved": "https://registry.npmmirror.com/negotiator/-/negotiator-1.0.0.tgz",
      "integrity": "sha512-8Ofs/AUQh8MaEcrlq5xOX0CQ9ypTF5dl78mjlMNfOK08fzpgTHQRQPBxcPlEtIw0yRpws+Zo/3r+5WRby7u3Gg==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.6"
      }
    },
    "node_modules/object-assign": {
      "version": "4.1.1",
      "resolved": "https://registry.npmmirror.com/object-assign/-/object-assign-4.1.1.tgz",
      "integrity": "sha512-rJgTQnkUnH1sFw8yT6VSU3zD3sWmu6sZhIseY8VX+GRu3P6F7Fu+JNDoXfklElbLJSnc3FUQHVe4cU5hj+BcUg==",
      "license": "MIT",
      "engines": {
        "node": ">=0.10.0"
      }
    },
    "node_modules/object-inspect": {
      "version": "1.13.4",
      "resolved": "https://registry.npmmirror.com/object-inspect/-/object-inspect-1.13.4.tgz",
      "integrity": "sha512-W67iLl4J2EXEGTbfeHCffrjDfitvLANg0UlX3wFUUSTx92KXRFegMHUVgSqE+wvhAbi4WqjGg9czysTV2Epbew==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.4"
      },
      "funding": {
        "url": "https://github.com/sponsors/ljharb"
      }
    },
    "node_modules/on-finished": {
      "version": "2.4.1",
      "resolved": "https://registry.npmmirror.com/on-finished/-/on-finished-2.4.1.tgz",
      "integrity": "sha512-oVlzkg3ENAhCk2zdv7IJwd/QUD4z2RxRwpkcGY8psCVcCYZNq4wYnVWALHM+brtuJjePWiYF/ClmuDr8Ch5+kg==",
      "license": "MIT",
      "dependencies": {
        "ee-first": "1.1.1"
      },
      "engines": {
        "node": ">= 0.8"
      }
    },
    "node_modules/once": {
      "version": "1.4.0",
      "resolved": "https://registry.npmmirror.com/once/-/once-1.4.0.tgz",
      "integrity": "sha512-lNaJgI+2Q5URQBkccEKHTQOPaXdUxnZZElQTZY0MFUAuaEqe1E+Nyvgdz/aIyNi6Z9MzO5dv1H8n58/GELp3+w==",
      "license": "ISC",
      "dependencies": {
        "wrappy": "1"
      }
    },
    "node_modules/onetime": {
      "version": "7.0.0",
      "resolved": "https://registry.npmmirror.com/onetime/-/onetime-7.0.0.tgz",
      "integrity": "sha512-VXJjc87FScF88uafS3JllDgvAm+c/Slfz06lorj2uAY34rlUu0Nt+v8wreiImcrgAjjIHp1rXpTDlLOGw29WwQ==",
      "license": "MIT",
      "dependencies": {
        "mimic-function": "^5.0.0"
      },
      "engines": {
        "node": ">=18"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/ora": {
      "version": "9.0.0",
      "resolved": "https://registry.npmmirror.com/ora/-/ora-9.0.0.tgz",
      "integrity": "sha512-m0pg2zscbYgWbqRR6ABga5c3sZdEon7bSgjnlXC64kxtxLOyjRcbbUkLj7HFyy/FTD+P2xdBWu8snGhYI0jc4A==",
      "license": "MIT",
      "dependencies": {
        "chalk": "^5.6.2",
        "cli-cursor": "^5.0.0",
        "cli-spinners": "^3.2.0",
        "is-interactive": "^2.0.0",
        "is-unicode-supported": "^2.1.0",
        "log-symbols": "^7.0.1",
        "stdin-discarder": "^0.2.2",
        "string-width": "^8.1.0",
        "strip-ansi": "^7.1.2"
      },
      "engines": {
        "node": ">=20"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/ora/node_modules/ansi-regex": {
      "version": "6.2.2",
      "resolved": "https://registry.npmmirror.com/ansi-regex/-/ansi-regex-6.2.2.tgz",
      "integrity": "sha512-Bq3SmSpyFHaWjPk8If9yc6svM8c56dB5BAtW4Qbw5jHTwwXXcTLoRMkpDJp6VL0XzlWaCHTXrkFURMYmD0sLqg==",
      "license": "MIT",
      "engines": {
        "node": ">=12"
      },
      "funding": {
        "url": "https://github.com/chalk/ansi-regex?sponsor=1"
      }
    },
    "node_modules/ora/node_modules/strip-ansi": {
      "version": "7.1.2",
      "resolved": "https://registry.npmmirror.com/strip-ansi/-/strip-ansi-7.1.2.tgz",
      "integrity": "sha512-gmBGslpoQJtgnMAvOVqGZpEz9dyoKTCzy2nfz/n8aIFhN/jCE/rCmcxabB6jOOHV+0WNnylOxaxBQPSvcWklhA==",
      "license": "MIT",
      "dependencies": {
        "ansi-regex": "^6.0.1"
      },
      "engines": {
        "node": ">=12"
      },
      "funding": {
        "url": "https://github.com/chalk/strip-ansi?sponsor=1"
      }
    },
    "node_modules/os-tmpdir": {
      "version": "1.0.2",
      "resolved": "https://registry.npmmirror.com/os-tmpdir/-/os-tmpdir-1.0.2.tgz",
      "integrity": "sha512-D2FR03Vir7FIu45XBY20mTb+/ZSWB00sjU9jdQXt83gDrI4Ztz5Fs7/yy74g2N5SVQY4xY1qDr4rNddwYRVX0g==",
      "license": "MIT",
      "engines": {
        "node": ">=0.10.0"
      }
    },
    "node_modules/pako": {
      "version": "1.0.11",
      "resolved": "https://registry.npmmirror.com/pako/-/pako-1.0.11.tgz",
      "integrity": "sha512-4hLB8Py4zZce5s4yd9XzopqwVv/yGNhV1Bl8NTmCq1763HeK2+EwVTv+leGeL13Dnh2wfbqowVPXCIO0z4taYw==",
      "license": "(MIT AND Zlib)"
    },
    "node_modules/parseurl": {
      "version": "1.3.3",
      "resolved": "https://registry.npmmirror.com/parseurl/-/parseurl-1.3.3.tgz",
      "integrity": "sha512-CiyeOxFT/JZyN5m0z9PfXw4SCBJ6Sygz1Dpl0wqjlhDEGGBP1GnsUVEL0p63hoG1fcj3fHynXi9NYO4nWOL+qQ==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.8"
      }
    },
    "node_modules/path-key": {
      "version": "3.1.1",
      "resolved": "https://registry.npmmirror.com/path-key/-/path-key-3.1.1.tgz",
      "integrity": "sha512-ojmeN0qd+y0jszEtoY48r0Peq5dwMEkIlCOu6Q5f41lfkswXuKtYrhgoTpLnyIcHm24Uhqx+5Tqm2InSwLhE6Q==",
      "license": "MIT",
      "engines": {
        "node": ">=8"
      }
    },
    "node_modules/path-to-regexp": {
      "version": "8.3.0",
      "resolved": "https://registry.npmmirror.com/path-to-regexp/-/path-to-regexp-8.3.0.tgz",
      "integrity": "sha512-7jdwVIRtsP8MYpdXSwOS0YdD0Du+qOoF/AEPIt88PcCFrZCzx41oxku1jD88hZBwbNUIEfpqvuhjFaMAqMTWnA==",
      "license": "MIT",
      "funding": {
        "type": "opencollective",
        "url": "https://opencollective.com/express"
      }
    },
    "node_modules/pkce-challenge": {
      "version": "5.0.0",
      "resolved": "https://registry.npmmirror.com/pkce-challenge/-/pkce-challenge-5.0.0.tgz",
      "integrity": "sha512-ueGLflrrnvwB3xuo/uGob5pd5FN7l0MsLf0Z87o/UQmRtwjvfylfc9MurIxRAWywCYTgrvpXBcqjV4OfCYGCIQ==",
      "license": "MIT",
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/process-nextick-args": {
      "version": "2.0.1",
      "resolved": "https://registry.npmmirror.com/process-nextick-args/-/process-nextick-args-2.0.1.tgz",
      "integrity": "sha512-3ouUOpQhtgrbOa17J7+uxOTpITYWaGP7/AhoR3+A+/1e9skrzelGi/dXzEYyvbxubEF6Wn2ypscTKiKJFFn1ag==",
      "license": "MIT"
    },
    "node_modules/proxy-addr": {
      "version": "2.0.7",
      "resolved": "https://registry.npmmirror.com/proxy-addr/-/proxy-addr-2.0.7.tgz",
      "integrity": "sha512-llQsMLSUDUPT44jdrU/O37qlnifitDP+ZwrmmZcoSKyLKvtZxpyV0n2/bD/N4tBAAZ/gJEdZU7KMraoK1+XYAg==",
      "license": "MIT",
      "dependencies": {
        "forwarded": "0.2.0",
        "ipaddr.js": "1.9.1"
      },
      "engines": {
        "node": ">= 0.10"
      }
    },
    "node_modules/punycode": {
      "version": "2.3.1",
      "resolved": "https://registry.npmmirror.com/punycode/-/punycode-2.3.1.tgz",
      "integrity": "sha512-vYt7UD1U9Wg6138shLtLOvdAu+8DsC/ilFtEVHcH+wydcSpNE20AfSOduf6MkRFahL5FY7X1oU7nKVZFtfq8Fg==",
      "license": "MIT",
      "engines": {
        "node": ">=6"
      }
    },
    "node_modules/qs": {
      "version": "6.14.0",
      "resolved": "https://registry.npmmirror.com/qs/-/qs-6.14.0.tgz",
      "integrity": "sha512-YWWTjgABSKcvs/nWBi9PycY/JiPJqOD4JA6o9Sej2AtvSGarXxKC3OQSk4pAarbdQlKAh5D4FCQkJNkW+GAn3w==",
      "license": "BSD-3-Clause",
      "dependencies": {
        "side-channel": "^1.1.0"
      },
      "engines": {
        "node": ">=0.6"
      },
      "funding": {
        "url": "https://github.com/sponsors/ljharb"
      }
    },
    "node_modules/range-parser": {
      "version": "1.2.1",
      "resolved": "https://registry.npmmirror.com/range-parser/-/range-parser-1.2.1.tgz",
      "integrity": "sha512-Hrgsx+orqoygnmhFbKaHE6c296J+HTAQXoxEF6gNupROmmGJRoyzfG3ccAveqCBrwr/2yxQ5BVd/GTl5agOwSg==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.6"
      }
    },
    "node_modules/raw-body": {
      "version": "3.0.1",
      "resolved": "https://registry.npmmirror.com/raw-body/-/raw-body-3.0.1.tgz",
      "integrity": "sha512-9G8cA+tuMS75+6G/TzW8OtLzmBDMo8p1JRxN5AZ+LAp8uxGA8V8GZm4GQ4/N5QNQEnLmg6SS7wyuSmbKepiKqA==",
      "license": "MIT",
      "dependencies": {
        "bytes": "3.1.2",
        "http-errors": "2.0.0",
        "iconv-lite": "0.7.0",
        "unpipe": "1.0.0"
      },
      "engines": {
        "node": ">= 0.10"
      }
    },
    "node_modules/raw-body/node_modules/iconv-lite": {
      "version": "0.7.0",
      "resolved": "https://registry.npmmirror.com/iconv-lite/-/iconv-lite-0.7.0.tgz",
      "integrity": "sha512-cf6L2Ds3h57VVmkZe+Pn+5APsT7FpqJtEhhieDCvrE2MK5Qk9MyffgQyuxQTm6BChfeZNtcOLHp9IcWRVcIcBQ==",
      "license": "MIT",
      "dependencies": {
        "safer-buffer": ">= 2.1.2 < 3.0.0"
      },
      "engines": {
        "node": ">=0.10.0"
      },
      "funding": {
        "type": "opencollective",
        "url": "https://opencollective.com/express"
      }
    },
    "node_modules/readable-stream": {
      "version": "2.3.8",
      "resolved": "https://registry.npmmirror.com/readable-stream/-/readable-stream-2.3.8.tgz",
      "integrity": "sha512-8p0AUk4XODgIewSi0l8Epjs+EVnWiK7NoDIEGU0HhE7+ZyY8D1IMY7odu5lRrFXGg71L15KG8QrPmum45RTtdA==",
      "license": "MIT",
      "dependencies": {
        "core-util-is": "~1.0.0",
        "inherits": "~2.0.3",
        "isarray": "~1.0.0",
        "process-nextick-args": "~2.0.0",
        "safe-buffer": "~5.1.1",
        "string_decoder": "~1.1.1",
        "util-deprecate": "~1.0.1"
      }
    },
    "node_modules/readable-stream/node_modules/safe-buffer": {
      "version": "5.1.2",
      "resolved": "https://registry.npmmirror.com/safe-buffer/-/safe-buffer-5.1.2.tgz",
      "integrity": "sha512-Gd2UZBJDkXlY7GbJxfsE8/nvKkUEU1G38c1siN6QP6a9PT9MmHB8GnpscSmMJSoF8LOIrt8ud/wPtojys4G6+g==",
      "license": "MIT"
    },
    "node_modules/restore-cursor": {
      "version": "5.1.0",
      "resolved": "https://registry.npmmirror.com/restore-cursor/-/restore-cursor-5.1.0.tgz",
      "integrity": "sha512-oMA2dcrw6u0YfxJQXm342bFKX/E4sG9rbTzO9ptUcR/e8A33cHuvStiYOwH7fszkZlZ1z/ta9AAoPk2F4qIOHA==",
      "license": "MIT",
      "dependencies": {
        "onetime": "^7.0.0",
        "signal-exit": "^4.1.0"
      },
      "engines": {
        "node": ">=18"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/router": {
      "version": "2.2.0",
      "resolved": "https://registry.npmmirror.com/router/-/router-2.2.0.tgz",
      "integrity": "sha512-nLTrUKm2UyiL7rlhapu/Zl45FwNgkZGaCpZbIHajDYgwlJCOzLSk+cIPAnsEqV955GjILJnKbdQC1nVPz+gAYQ==",
      "license": "MIT",
      "dependencies": {
        "debug": "^4.4.0",
        "depd": "^2.0.0",
        "is-promise": "^4.0.0",
        "parseurl": "^1.3.3",
        "path-to-regexp": "^8.0.0"
      },
      "engines": {
        "node": ">= 18"
      }
    },
    "node_modules/safe-buffer": {
      "version": "5.2.1",
      "resolved": "https://registry.npmmirror.com/safe-buffer/-/safe-buffer-5.2.1.tgz",
      "integrity": "sha512-rp3So07KcdmmKbGvgaNxQSJr7bGVSVk5S9Eq1F+ppbRo70+YeaDxkw5Dd8NPN+GD6bjnYm2VuPuCXmpuYvmCXQ==",
      "funding": [
        {
          "type": "github",
          "url": "https://github.com/sponsors/feross"
        },
        {
          "type": "patreon",
          "url": "https://www.patreon.com/feross"
        },
        {
          "type": "consulting",
          "url": "https://feross.org/support"
        }
      ],
      "license": "MIT"
    },
    "node_modules/safer-buffer": {
      "version": "2.1.2",
      "resolved": "https://registry.npmmirror.com/safer-buffer/-/safer-buffer-2.1.2.tgz",
      "integrity": "sha512-YZo3K82SD7Riyi0E1EQPojLz7kpepnSQI9IyPbHHg1XXXevb5dJI7tpyN2ADxGcQbHG7vcyRHk0cbwqcQriUtg==",
      "license": "MIT"
    },
    "node_modules/sax": {
      "version": "1.4.3",
      "resolved": "https://registry.npmmirror.com/sax/-/sax-1.4.3.tgz",
      "integrity": "sha512-yqYn1JhPczigF94DMS+shiDMjDowYO6y9+wB/4WgO0Y19jWYk0lQ4tuG5KI7kj4FTp1wxPj5IFfcrz/s1c3jjQ==",
      "license": "BlueOak-1.0.0"
    },
    "node_modules/send": {
      "version": "1.2.0",
      "resolved": "https://registry.npmmirror.com/send/-/send-1.2.0.tgz",
      "integrity": "sha512-uaW0WwXKpL9blXE2o0bRhoL2EGXIrZxQ2ZQ4mgcfoBxdFmQold+qWsD2jLrfZ0trjKL6vOw0j//eAwcALFjKSw==",
      "license": "MIT",
      "dependencies": {
        "debug": "^4.3.5",
        "encodeurl": "^2.0.0",
        "escape-html": "^1.0.3",
        "etag": "^1.8.1",
        "fresh": "^2.0.0",
        "http-errors": "^2.0.0",
        "mime-types": "^3.0.1",
        "ms": "^2.1.3",
        "on-finished": "^2.4.1",
        "range-parser": "^1.2.1",
        "statuses": "^2.0.1"
      },
      "engines": {
        "node": ">= 18"
      }
    },
    "node_modules/serve-static": {
      "version": "2.2.0",
      "resolved": "https://registry.npmmirror.com/serve-static/-/serve-static-2.2.0.tgz",
      "integrity": "sha512-61g9pCh0Vnh7IutZjtLGGpTA355+OPn2TyDv/6ivP2h/AdAVX9azsoxmg2/M6nZeQZNYBEwIcsne1mJd9oQItQ==",
      "license": "MIT",
      "dependencies": {
        "encodeurl": "^2.0.0",
        "escape-html": "^1.0.3",
        "parseurl": "^1.3.3",
        "send": "^1.2.0"
      },
      "engines": {
        "node": ">= 18"
      }
    },
    "node_modules/setimmediate": {
      "version": "1.0.5",
      "resolved": "https://registry.npmmirror.com/setimmediate/-/setimmediate-1.0.5.tgz",
      "integrity": "sha512-MATJdZp8sLqDl/68LfQmbP8zKPLQNV6BIZoIgrscFDQ+RsvK/BxeDQOgyxKKoh0y/8h3BqVFnCqQ/gd+reiIXA==",
      "license": "MIT"
    },
    "node_modules/setprototypeof": {
      "version": "1.2.0",
      "resolved": "https://registry.npmmirror.com/setprototypeof/-/setprototypeof-1.2.0.tgz",
      "integrity": "sha512-E5LDX7Wrp85Kil5bhZv46j8jOeboKq5JMmYM3gVGdGH8xFpPWXUMsNrlODCrkoxMEeNi/XZIwuRvY4XNwYMJpw==",
      "license": "ISC"
    },
    "node_modules/shebang-command": {
      "version": "2.0.0",
      "resolved": "https://registry.npmmirror.com/shebang-command/-/shebang-command-2.0.0.tgz",
      "integrity": "sha512-kHxr2zZpYtdmrN1qDjrrX/Z1rR1kG8Dx+gkpK1G4eXmvXswmcE1hTWBWYUzlraYw1/yZp6YuDY77YtvbN0dmDA==",
      "license": "MIT",
      "dependencies": {
        "shebang-regex": "^3.0.0"
      },
      "engines": {
        "node": ">=8"
      }
    },
    "node_modules/shebang-regex": {
      "version": "3.0.0",
      "resolved": "https://registry.npmmirror.com/shebang-regex/-/shebang-regex-3.0.0.tgz",
      "integrity": "sha512-7++dFhtcx3353uBaq8DDR4NuxBetBzC7ZQOhmTQInHEd6bSrXdiEyzCvG07Z44UYdLShWUyXt5M/yhz8ekcb1A==",
      "license": "MIT",
      "engines": {
        "node": ">=8"
      }
    },
    "node_modules/side-channel": {
      "version": "1.1.0",
      "resolved": "https://registry.npmmirror.com/side-channel/-/side-channel-1.1.0.tgz",
      "integrity": "sha512-ZX99e6tRweoUXqR+VBrslhda51Nh5MTQwou5tnUDgbtyM0dBgmhEDtWGP/xbKn6hqfPRHujUNwz5fy/wbbhnpw==",
      "license": "MIT",
      "dependencies": {
        "es-errors": "^1.3.0",
        "object-inspect": "^1.13.3",
        "side-channel-list": "^1.0.0",
        "side-channel-map": "^1.0.1",
        "side-channel-weakmap": "^1.0.2"
      },
      "engines": {
        "node": ">= 0.4"
      },
      "funding": {
        "url": "https://github.com/sponsors/ljharb"
      }
    },
    "node_modules/side-channel-list": {
      "version": "1.0.0",
      "resolved": "https://registry.npmmirror.com/side-channel-list/-/side-channel-list-1.0.0.tgz",
      "integrity": "sha512-FCLHtRD/gnpCiCHEiJLOwdmFP+wzCmDEkc9y7NsYxeF4u7Btsn1ZuwgwJGxImImHicJArLP4R0yX4c2KCrMrTA==",
      "license": "MIT",
      "dependencies": {
        "es-errors": "^1.3.0",
        "object-inspect": "^1.13.3"
      },
      "engines": {
        "node": ">= 0.4"
      },
      "funding": {
        "url": "https://github.com/sponsors/ljharb"
      }
    },
    "node_modules/side-channel-map": {
      "version": "1.0.1",
      "resolved": "https://registry.npmmirror.com/side-channel-map/-/side-channel-map-1.0.1.tgz",
      "integrity": "sha512-VCjCNfgMsby3tTdo02nbjtM/ewra6jPHmpThenkTYh8pG9ucZ/1P8So4u4FGBek/BjpOVsDCMoLA/iuBKIFXRA==",
      "license": "MIT",
      "dependencies": {
        "call-bound": "^1.0.2",
        "es-errors": "^1.3.0",
        "get-intrinsic": "^1.2.5",
        "object-inspect": "^1.13.3"
      },
      "engines": {
        "node": ">= 0.4"
      },
      "funding": {
        "url": "https://github.com/sponsors/ljharb"
      }
    },
    "node_modules/side-channel-weakmap": {
      "version": "1.0.2",
      "resolved": "https://registry.npmmirror.com/side-channel-weakmap/-/side-channel-weakmap-1.0.2.tgz",
      "integrity": "sha512-WPS/HvHQTYnHisLo9McqBHOJk2FkHO/tlpvldyrnem4aeQp4hai3gythswg6p01oSoTl58rcpiFAjF2br2Ak2A==",
      "license": "MIT",
      "dependencies": {
        "call-bound": "^1.0.2",
        "es-errors": "^1.3.0",
        "get-intrinsic": "^1.2.5",
        "object-inspect": "^1.13.3",
        "side-channel-map": "^1.0.1"
      },
      "engines": {
        "node": ">= 0.4"
      },
      "funding": {
        "url": "https://github.com/sponsors/ljharb"
      }
    },
    "node_modules/signal-exit": {
      "version": "4.1.0",
      "resolved": "https://registry.npmmirror.com/signal-exit/-/signal-exit-4.1.0.tgz",
      "integrity": "sha512-bzyZ1e88w9O1iNJbKnOlvYTrWPDl46O1bG0D3XInv+9tkPrxrN8jUUTiFlDkkmKWgn1M6CfIA13SuGqOa9Korw==",
      "license": "ISC",
      "engines": {
        "node": ">=14"
      },
      "funding": {
        "url": "https://github.com/sponsors/isaacs"
      }
    },
    "node_modules/statuses": {
      "version": "2.0.2",
      "resolved": "https://registry.npmmirror.com/statuses/-/statuses-2.0.2.tgz",
      "integrity": "sha512-DvEy55V3DB7uknRo+4iOGT5fP1slR8wQohVdknigZPMpMstaKJQWhwiYBACJE3Ul2pTnATihhBYnRhZQHGBiRw==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.8"
      }
    },
    "node_modules/stdin-discarder": {
      "version": "0.2.2",
      "resolved": "https://registry.npmmirror.com/stdin-discarder/-/stdin-discarder-0.2.2.tgz",
      "integrity": "sha512-UhDfHmA92YAlNnCfhmq0VeNL5bDbiZGg7sZ2IvPsXubGkiNa9EC+tUTsjBRsYUAz87btI6/1wf4XoVvQ3uRnmQ==",
      "license": "MIT",
      "engines": {
        "node": ">=18"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/string_decoder": {
      "version": "1.1.1",
      "resolved": "https://registry.npmmirror.com/string_decoder/-/string_decoder-1.1.1.tgz",
      "integrity": "sha512-n/ShnvDi6FHbbVfviro+WojiFzv+s8MPMHBczVePfUpDJLwoLT0ht1l4YwBCbi8pJAveEEdnkHyPyTP/mzRfwg==",
      "license": "MIT",
      "dependencies": {
        "safe-buffer": "~5.1.0"
      }
    },
    "node_modules/string_decoder/node_modules/safe-buffer": {
      "version": "5.1.2",
      "resolved": "https://registry.npmmirror.com/safe-buffer/-/safe-buffer-5.1.2.tgz",
      "integrity": "sha512-Gd2UZBJDkXlY7GbJxfsE8/nvKkUEU1G38c1siN6QP6a9PT9MmHB8GnpscSmMJSoF8LOIrt8ud/wPtojys4G6+g==",
      "license": "MIT"
    },
    "node_modules/string-width": {
      "version": "8.1.0",
      "resolved": "https://registry.npmmirror.com/string-width/-/string-width-8.1.0.tgz",
      "integrity": "sha512-Kxl3KJGb/gxkaUMOjRsQ8IrXiGW75O4E3RPjFIINOVH8AMl2SQ/yWdTzWwF3FevIX9LcMAjJW+GRwAlAbTSXdg==",
      "license": "MIT",
      "dependencies": {
        "get-east-asian-width": "^1.3.0",
        "strip-ansi": "^7.1.0"
      },
      "engines": {
        "node": ">=20"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/string-width/node_modules/ansi-regex": {
      "version": "6.2.2",
      "resolved": "https://registry.npmmirror.com/ansi-regex/-/ansi-regex-6.2.2.tgz",
      "integrity": "sha512-Bq3SmSpyFHaWjPk8If9yc6svM8c56dB5BAtW4Qbw5jHTwwXXcTLoRMkpDJp6VL0XzlWaCHTXrkFURMYmD0sLqg==",
      "license": "MIT",
      "engines": {
        "node": ">=12"
      },
      "funding": {
        "url": "https://github.com/chalk/ansi-regex?sponsor=1"
      }
    },
    "node_modules/string-width/node_modules/strip-ansi": {
      "version": "7.1.2",
      "resolved": "https://registry.npmmirror.com/strip-ansi/-/strip-ansi-7.1.2.tgz",
      "integrity": "sha512-gmBGslpoQJtgnMAvOVqGZpEz9dyoKTCzy2nfz/n8aIFhN/jCE/rCmcxabB6jOOHV+0WNnylOxaxBQPSvcWklhA==",
      "license": "MIT",
      "dependencies": {
        "ansi-regex": "^6.0.1"
      },
      "engines": {
        "node": ">=12"
      },
      "funding": {
        "url": "https://github.com/chalk/strip-ansi?sponsor=1"
      }
    },
    "node_modules/strip-ansi": {
      "version": "6.0.1",
      "resolved": "https://registry.npmmirror.com/strip-ansi/-/strip-ansi-6.0.1.tgz",
      "integrity": "sha512-Y38VPSHcqkFrCpFnQ9vuSXmquuv5oXOKpGeT6aGrr3o3Gc9AlVa6JBfUSOCnbxGGZF+/0ooI7KrPuUSztUdU5A==",
      "license": "MIT",
      "dependencies": {
        "ansi-regex": "^5.0.1"
      },
      "engines": {
        "node": ">=8"
      }
    },
    "node_modules/tmp": {
      "version": "0.0.33",
      "resolved": "https://registry.npmmirror.com/tmp/-/tmp-0.0.33.tgz",
      "integrity": "sha512-jRCJlojKnZ3addtTOjdIqoRuPEKBvNXcGYqzO6zWZX8KfKEpnGY5jfggJQ3EjKuu8D4bJRr0y+cYJFmYbImXGw==",
      "license": "MIT",
      "dependencies": {
        "os-tmpdir": "~1.0.2"
      },
      "engines": {
        "node": ">=0.6.0"
      }
    },
    "node_modules/toidentifier": {
      "version": "1.0.1",
      "resolved": "https://registry.npmmirror.com/toidentifier/-/toidentifier-1.0.1.tgz",
      "integrity": "sha512-o5sSPKEkg/DIQNmH43V0/uerLrpzVedkUh8tGNvaeXpfpuwjKenlSox/2O/BTlZUtEe+JG7s5YhEz608PlAHRA==",
      "license": "MIT",
      "engines": {
        "node": ">=0.6"
      }
    },
    "node_modules/ts-algebra": {
      "version": "2.0.0",
      "resolved": "https://registry.npmmirror.com/ts-algebra/-/ts-algebra-2.0.0.tgz",
      "integrity": "sha512-FPAhNPFMrkwz76P7cdjdmiShwMynZYN6SgOujD1urY4oNm80Ou9oMdmbR45LotcKOXoy7wSmHkRFE6Mxbrhefw=="
    },
    "node_modules/ts-node": {
      "version": "10.9.2",
      "resolved": "https://registry.npmmirror.com/ts-node/-/ts-node-10.9.2.tgz",
      "integrity": "sha512-f0FFpIdcHgn8zcPSbf1dRevwt047YMnaiJM3u2w2RewrB+fob/zePZcrOyQoLMMO7aBIddLcQIEK5dYjkLnGrQ==",
      "dev": true,
      "dependencies": {
        "@cspotcode/source-map-support": "^0.8.0",
        "@tsconfig/node10": "^1.0.7",
        "@tsconfig/node12": "^1.0.7",
        "@tsconfig/node14": "^1.0.0",
        "@tsconfig/node16": "^1.0.2",
        "acorn": "^8.4.1",
        "acorn-walk": "^8.1.1",
        "arg": "^4.1.0",
        "create-require": "^1.1.0",
        "diff": "^4.0.1",
        "make-error": "^1.1.1",
        "v8-compile-cache-lib": "^3.0.1",
        "yn": "3.1.1"
      },
      "bin": {
        "ts-node": "dist/bin.js",
        "ts-node-cwd": "dist/bin-cwd.js",
        "ts-node-esm": "dist/bin-esm.js",
        "ts-node-script": "dist/bin-script.js",
        "ts-node-transpile-only": "dist/bin-transpile.js",
        "ts-script": "dist/bin-script-deprecated.js"
      },
      "peerDependencies": {
        "@swc/core": ">=1.2.50",
        "@swc/wasm": ">=1.2.50",
        "@types/node": "*",
        "typescript": ">=2.7"
      },
      "peerDependenciesMeta": {
        "@swc/core": {
          "optional": true
        },
        "@swc/wasm": {
          "optional": true
        }
      }
    },
    "node_modules/type-fest": {
      "version": "0.21.3",
      "resolved": "https://registry.npmmirror.com/type-fest/-/type-fest-0.21.3.tgz",
      "integrity": "sha512-t0rzBq87m3fVcduHDUFhKmyyX+9eo6WQjZvf51Ea/M0Q7+T374Jp1aUiyUl0GKxp8M/OETVHSDvmkyPgvX+X2w==",
      "license": "(MIT OR CC0-1.0)",
      "engines": {
        "node": ">=10"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/type-is": {
      "version": "2.0.1",
      "resolved": "https://registry.npmmirror.com/type-is/-/type-is-2.0.1.tgz",
      "integrity": "sha512-OZs6gsjF4vMp32qrCbiVSkrFmXtG/AZhY3t0iAMrMBiAZyV9oALtXO8hsrHbMXF9x6L3grlFuwW2oAz7cav+Gw==",
      "license": "MIT",
      "dependencies": {
        "content-type": "^1.0.5",
        "media-typer": "^1.1.0",
        "mime-types": "^3.0.0"
      },
      "engines": {
        "node": ">= 0.6"
      }
    },
    "node_modules/typescript": {
      "version": "5.9.2",
      "resolved": "https://registry.npmmirror.com/typescript/-/typescript-5.9.2.tgz",
      "integrity": "sha512-CWBzXQrc/qOkhidw1OzBTQuYRbfyxDXJMVJ1XNwUHGROVmuaeiEm3OslpZ1RV96d7SKKjZKrSJu3+t/xlw3R9A==",
      "dev": true,
      "bin": {
        "tsc": "bin/tsc",
        "tsserver": "bin/tsserver"
      },
      "engines": {
        "node": ">=14.17"
      }
    },
    "node_modules/undici-types": {
      "version": "6.21.0",
      "resolved": "https://registry.npmmirror.com/undici-types/-/undici-types-6.21.0.tgz",
      "integrity": "sha512-iwDZqg0QAGrg9Rav5H4n0M64c3mkR59cJ6wQp+7C4nI0gsmExaedaYLNO44eT4AtBBwjbTiGPMlt2Md0T9H9JQ=="
    },
    "node_modules/unpipe": {
      "version": "1.0.0",
      "resolved": "https://registry.npmmirror.com/unpipe/-/unpipe-1.0.0.tgz",
      "integrity": "sha512-pjy2bYhSsufwWlKwPc+l3cN7+wuJlK6uz0YdJEOlQDbl6jo/YlPi4mb8agUkVC8BF7V8NuzeyPNqRksA3hztKQ==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.8"
      }
    },
    "node_modules/uri-js": {
      "version": "4.4.1",
      "resolved": "https://registry.npmmirror.com/uri-js/-/uri-js-4.4.1.tgz",
      "integrity": "sha512-7rKUyy33Q1yc98pQ1DAmLtwX109F7TIfWlW1Ydo8Wl1ii1SeHieeh0HHfPeL2fMXK6z0s8ecKs9frCuLJvndBg==",
      "license": "BSD-2-Clause",
      "dependencies": {
        "punycode": "^2.1.0"
      }
    },
    "node_modules/util-deprecate": {
      "version": "1.0.2",
      "resolved": "https://registry.npmmirror.com/util-deprecate/-/util-deprecate-1.0.2.tgz",
      "integrity": "sha512-EPD5q1uXyFxJpCrLnCc1nHnq3gOa6DZBocAIiI2TaSCA7VCJ1UJDMagCzIkXNsUYfD1daK//LTEQ8xiIbrHtcw==",
      "license": "MIT"
    },
    "node_modules/v8-compile-cache-lib": {
      "version": "3.0.1",
      "resolved": "https://registry.npmmirror.com/v8-compile-cache-lib/-/v8-compile-cache-lib-3.0.1.tgz",
      "integrity": "sha512-wa7YjyUGfNZngI/vtK0UHAN+lgDCxBPCylVXGp0zu59Fz5aiGtNXaq3DhIov063MorB+VfufLh3JlF2KdTK3xg==",
      "dev": true
    },
    "node_modules/vary": {
      "version": "1.1.2",
      "resolved": "https://registry.npmmirror.com/vary/-/vary-1.1.2.tgz",
      "integrity": "sha512-BNGbWLfd0eUPabhkXUVm0j8uuvREyTh5ovRa/dyow/BqAbZJyC+5fU+IzQOzmAKzYqYRAISoRhdQr3eIZ/PXqg==",
      "license": "MIT",
      "engines": {
        "node": ">= 0.8"
      }
    },
    "node_modules/which": {
      "version": "2.0.2",
      "resolved": "https://registry.npmmirror.com/which/-/which-2.0.2.tgz",
      "integrity": "sha512-BLI3Tl1TW3Pvl70l3yq3Y64i+awpwXqsGBYWkkqMtnbXgrMD+yj7rhW0kuEDxzJaYXGjEW5ogapKNMEKNMjibA==",
      "license": "ISC",
      "dependencies": {
        "isexe": "^2.0.0"
      },
      "bin": {
        "node-which": "bin/node-which"
      },
      "engines": {
        "node": ">= 8"
      }
    },
    "node_modules/wrap-ansi": {
      "version": "6.2.0",
      "resolved": "https://registry.npmmirror.com/wrap-ansi/-/wrap-ansi-6.2.0.tgz",
      "integrity": "sha512-r6lPcBGxZXlIcymEu7InxDMhdW0KDxpLgoFLcguasxCaJ/SOIZwINatK9KY/tf+ZrlywOKU0UDj3ATXUBfxJXA==",
      "license": "MIT",
      "dependencies": {
        "ansi-styles": "^4.0.0",
        "string-width": "^4.1.0",
        "strip-ansi": "^6.0.0"
      },
      "engines": {
        "node": ">=8"
      }
    },
    "node_modules/wrap-ansi/node_modules/string-width": {
      "version": "4.2.3",
      "resolved": "https://registry.npmmirror.com/string-width/-/string-width-4.2.3.tgz",
      "integrity": "sha512-wKyQRQpjJ0sIp62ErSZdGsjMJWsap5oRNihHhu6G7JVO/9jIB6UyevL+tXuOqrng8j/cxKTWyWUwvSTriiZz/g==",
      "license": "MIT",
      "dependencies": {
        "emoji-regex": "^8.0.0",
        "is-fullwidth-code-point": "^3.0.0",
        "strip-ansi": "^6.0.1"
      },
      "engines": {
        "node": ">=8"
      }
    },
    "node_modules/wrappy": {
      "version": "1.0.2",
      "resolved": "https://registry.npmmirror.com/wrappy/-/wrappy-1.0.2.tgz",
      "integrity": "sha512-l4Sp/DRseor9wL6EvV2+TuQn63dMkPjZ/sp9XkghTEbV9KlPS1xUsZ3u7/IQO4wxtcFB4bgpQPRcR3QCvezPcQ==",
      "license": "ISC"
    },
    "node_modules/xml": {
      "version": "1.0.1",
      "resolved": "https://registry.npmmirror.com/xml/-/xml-1.0.1.tgz",
      "integrity": "sha512-huCv9IH9Tcf95zuYCsQraZtWnJvBtLVE0QHMOs8bWyZAFZNDcYjsPq1nEx8jKA9y+Beo9v+7OBPRisQTjinQMw==",
      "license": "MIT"
    },
    "node_modules/xml-js": {
      "version": "1.6.11",
      "resolved": "https://registry.npmmirror.com/xml-js/-/xml-js-1.6.11.tgz",
      "integrity": "sha512-7rVi2KMfwfWFl+GpPg6m80IVMWXLRjO+PxTq7V2CDhoGak0wzYzFgUY2m4XJ47OGdXd8eLE8EmwfAmdjw7lC1g==",
      "license": "MIT",
      "dependencies": {
        "sax": "^1.2.4"
      },
      "bin": {
        "xml-js": "bin/cli.js"
      }
    },
    "node_modules/yn": {
      "version": "3.1.1",
      "resolved": "https://registry.npmmirror.com/yn/-/yn-3.1.1.tgz",
      "integrity": "sha512-Ux4ygGWsu2c7isFWe8Yu1YluJmqVhxqK2cLXNQA5AcC3QfbGNpM7fu0Y8b/z16pXLnFxZYvWhd3fhBY9DLmC6Q==",
      "dev": true,
      "engines": {
        "node": ">=6"
      }
    },
    "node_modules/yoctocolors": {
      "version": "2.1.2",
      "resolved": "https://registry.npmmirror.com/yoctocolors/-/yoctocolors-2.1.2.tgz",
      "integrity": "sha512-CzhO+pFNo8ajLM2d2IW/R93ipy99LWjtwblvC1RsoSUMZgyLbYFr221TnSNT7GjGdYui6P459mw9JH/g/zW2ug==",
      "license": "MIT",
      "engines": {
        "node": ">=18"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/yoctocolors-cjs": {
      "version": "2.1.3",
      "resolved": "https://registry.npmmirror.com/yoctocolors-cjs/-/yoctocolors-cjs-2.1.3.tgz",
      "integrity": "sha512-U/PBtDf35ff0D8X8D0jfdzHYEPFxAI7jJlxZXwCSez5M3190m+QobIfh+sWDWSHMCWWJN2AWamkegn6vr6YBTw==",
      "license": "MIT",
      "engines": {
        "node": ">=18"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/zod": {
      "version": "3.25.76",
      "resolved": "https://registry.npmmirror.com/zod/-/zod-3.25.76.tgz",
      "integrity": "sha512-gzUt/qt81nXsFGKIFcC3YnfEAx5NkunCfnDlvuBSSFS02bcXu4Lmea0AFIUwbLWxWPx3d9p8S5QoaujKcNQxcQ==",
      "license": "MIT",
      "funding": {
        "url": "https://github.com/sponsors/colinhacks"
      }
    },
    "node_modules/zod-to-json-schema": {
      "version": "3.24.6",
      "resolved": "https://registry.npmmirror.com/zod-to-json-schema/-/zod-to-json-schema-3.24.6.tgz",
      "integrity": "sha512-h/z3PKvcTcTetyjl1fkj79MHNEjm+HpD6NXheWjzOekY7kV+lwDYnHw+ivHkijnCSMz1yJaWBD9vu/Fcmk+vEg==",
      "license": "ISC",
      "peerDependencies": {
        "zod": "^3.24.1"
      }
    }
  }
}
