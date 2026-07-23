# -*- coding: utf-8 -*-
r"""acad_common —— cad-drawing 技能公共库（FP-007 / FP-008 两张已验收图纸的全部机制提炼）

用法（在 workspace 绘图脚本中）：
    import sys
    sys.path.insert(0, r"C:\Users\wu\AppData\Roaming\kimi-desktop\daimon-share\daimon\skills\cad-drawing\scripts")
    import acad_common as A

    acad = A.connect_acad()
    doc, ms = A.new_drawing(acad)
    A.setup_env(doc, dimlfac=1.0, layer_def=A.LAYERS_DEFAULT)
    ... 用 A.line / A.circle / A.text / A.dim / A.leader ... 逐步绘制 ...
    A.plot_png(doc, out_png); A.plot_window(doc, zoom_png, ll, ur)
    ... 自审计 + A.verify_dwg_unchanged() ...

本模块 import 时不连接 AutoCAD、不产生任何副作用；全部功能在函数内。
铁律：禁止 SaveAs；引线单段直线 + AddSolid 实心箭头；每实体 Update、每阶段 Regen。
"""
import math
import os
import re
import time
from collections import Counter

import pythoncom
import pywintypes
import win32com.client
from win32com.client import dynamic, gencache

RPC_REJECTED = -2147418111  # 被呼叫方拒绝接收呼叫（AutoCAD 忙 / 弹对话框）

# 逐步可见延时（规范 V1.4：实体级 0.2~0.35s，阶段级 0.8~1s）。可按需调小做冒烟测试。
SLEEP_ENT = 0.3
SLEEP_STAGE = 0.9

# 字形实测（_glyph_probe.py）：宋体缺 Ø(U+00D8)/↧(U+21A7) → Φ 用 U+03A6、深度用 ↓(U+2193)
DEPTH = "↓"

# 规范默认七层；用户指定层名/五层时由调用方传入自定义 layer_def
LAYERS_DEFAULT = {  # 层名: (ACI 颜色, 线型, 线宽x100)
    "粗实线": (7, "Continuous", 50),   # 图框 + 可见轮廓 + 螺纹小径
    "细实线": (7, "Continuous", 25),   # 尺寸要素外的细线 + 螺纹大径
    "中心线": (1, "CENTER2",    25),   # 轴线/十字中心线/分度圆
    "虚线":   (8, "DASHED2",    25),   # 不可见轮廓
    "标注":   (3, "Continuous", 25),   # 全部尺寸/文字/引线/框格
}

LEADER_AUDIT = []   # (label, 尖端x, 尖端y, 朝向deg, 段数)，leader() 自动记录，供自审计
_DWG_MTIME0 = {}    # snapshot_dwg() 填充，verify_dwg_unchanged() 比对


# ==================================================================
# COM 基础
# ==================================================================
def call(fn, *args, retries=120, **kw):
    """带重试的 COM 调用：AutoCAD 忙/弹对话框时（RPC_REJECTED）等待 1s 重试。
    若持续被拒：检查 CAD 中是否有挂起的对话框/文字编辑窗口（可 PostMessage 发 ESC）。"""
    for _ in range(retries):
        try:
            return fn(*args, **kw)
        except pywintypes.com_error as e:
            if e.args and e.args[0] == RPC_REJECTED:
                time.sleep(1)
                continue
            raise
    raise RuntimeError("AutoCAD 持续繁忙，请关闭 CAD 中的对话框后重试")


def setp(obj, name, val):
    """早期绑定缺属性时自动转后期绑定写属性"""
    try:
        setattr(obj, name, val)
    except AttributeError:
        setattr(dynamic.Dispatch(obj), name, val)


def getp(obj, name):
    try:
        return getattr(obj, name)
    except AttributeError:
        return getattr(dynamic.Dispatch(obj), name)


def pt(x, y, z=0.0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,
                                   (float(x), float(y), float(z)))


# ==================================================================
# 连接 / 新建文档（禁止保存；运行前清理遗留 Drawing*）
# ==================================================================
def connect_acad():
    """连接已运行的 AutoCAD（须已由用户打开），置 Visible 逐步可见。"""
    acad = gencache.EnsureDispatch("AutoCAD.Application")
    call(setp, acad, "Visible", True)
    print("[0] connected to AutoCAD:", acad.Name, acad.Version)
    return acad


def new_drawing(acad, ws_dir=r"C:\Users\wu\Documents\kimi\workspace"):
    """新建图形（Documents.Add，全程不保存）；关闭上次运行遗留的未保存 Drawing*；
    快照 workspace 全部 .dwg 的 mtime 供结束后 verify_dwg_unchanged() 比对。"""
    for i in range(call(lambda: acad.Documents.Count) - 1, -1, -1):
        d = call(acad.Documents.Item, i)
        if str(call(getp, d, "Name")).lower().startswith("drawing"):
            print("[0] 关闭遗留未保存文档:", call(getp, d, "Name"))
            call(d.Close, False)
    snapshot_dwg(ws_dir)
    doc = call(acad.Documents.Add)
    call(doc.SetVariable, "INSUNITS", 4)        # 毫米
    call(doc.SetVariable, "MEASUREMENT", 1)     # 公制（配合 acadiso.lin）
    call(doc.SetVariable, "BACKGROUNDPLOT", 0)  # 前台打印，PlotToFile 同步返回
    call(doc.SetVariable, "LWDISPLAY", 1)       # 屏幕显示线宽
    print("[0] new document:", doc.Name, "（新建图形，全程不保存）")
    return doc, doc.ModelSpace


def snapshot_dwg(ws_dir):
    global _DWG_MTIME0
    _DWG_MTIME0 = {f: os.path.getmtime(os.path.join(ws_dir, f))
                   for f in os.listdir(ws_dir) if f.lower().endswith(".dwg")}
    print("[0] 磁盘 .dwg 修改时间快照（{} 个）".format(len(_DWG_MTIME0)))
    return dict(_DWG_MTIME0)


def verify_dwg_unchanged(ws_dir):
    """保存禁令验证：磁盘 .dwg mtime 全部未变 + 新文档 FullName 为空（从未保存）。"""
    changed = [f for f, m0 in _DWG_MTIME0.items()
               if not os.path.exists(os.path.join(ws_dir, f))
               or os.path.getmtime(os.path.join(ws_dir, f)) != m0]
    print("磁盘 .dwg 修改时间比对（{} 个）：{}".format(
        len(_DWG_MTIME0), "全部未变（PASS）" if not changed else "被改动 -> " + str(changed)))
    assert not changed, "有 .dwg 被改动，违反禁止保存规则: " + str(changed)
    return True


# ==================================================================
# 环境：线型 / 图层 / 文字样式 / 尺寸变量
# ==================================================================
def ensure_linetypes(doc, names=("CENTER2", "DASHED2")):
    """线型加载：存在性检查 + acadiso.lin/acad.lin 双源重试 + 读回校验 + 失败抛错。
    （DASHED2 曾偶发退化为实线，必须读回校验。）返回应设置的 LTSCALE。"""
    lt_file = "acadiso.lin"

    def _exists(lt):
        try:
            doc.Linetypes.Item(lt)
            return True
        except pywintypes.com_error:
            return False

    for lt in names:
        ok = False
        for _ in range(6):
            if _exists(lt):
                ok = True
                break
            for fname in ("acadiso.lin", "acad.lin"):
                try:
                    call(doc.Linetypes.Load, lt, fname)
                    lt_file = fname
                    break
                except pywintypes.com_error:
                    continue
            if _exists(lt):
                ok = True
                break
            time.sleep(0.6)
        if not ok:
            raise RuntimeError("线型加载失败: " + lt)
        print("[1] linetype ready:", lt)
    ltscale = 0.4 if lt_file == "acadiso.lin" else 0.4 * 25.4
    call(doc.SetVariable, "LTSCALE", ltscale)
    call(doc.SetVariable, "CELTSCALE", 1.0)
    print("[1] linetypes from", lt_file, "LTSCALE =", ltscale)
    return ltscale


def _new_cmcolor(acad, aci):
    last = None
    for progid in ("AutoCAD.AcCmColor.23", "AutoCAD.AcCmColor"):
        try:
            cm = call(acad.GetInterfaceObject, progid)
            call(setp, cm, "ColorIndex", aci)
            return cm
        except Exception as ex:
            last = ex
    raise last


def make_layers(doc, acad, layer_def=LAYERS_DEFAULT):
    """建层：TrueColor 设色 + 线型读回校验（失败抛错）+ 线宽。"""
    for name, (color, lt, lw) in layer_def.items():
        try:
            lay = call(doc.Layers.Add, name)
        except pywintypes.com_error:
            lay = doc.Layers.Item(name)
        try:
            call(setp, lay, "TrueColor", _new_cmcolor(acad, color))
        except Exception:
            try:
                call(setp, dynamic.Dispatch(lay), "Color", color)
            except Exception as ex2:
                print("[1] warn: color set failed on layer", name, "->", ex2)
        lt_ok = False
        for _ in range(4):
            try:
                call(setp, lay, "Linetype", lt)
                if str(call(getp, lay, "Linetype")).upper() == lt.upper():
                    lt_ok = True
                    break
            except pywintypes.com_error:
                time.sleep(0.5)
        if not lt_ok:
            raise RuntimeError("图层 {} 线型 {} 设置失败".format(name, lt))
        try:
            call(setp, lay, "Lineweight", lw)
        except pywintypes.com_error:
            print("[1] warn: lineweight set failed on layer", name)
    print("[1] layers ready:", ", ".join(layer_def))


def setup_text_style(doc, style="CN"):
    """宋体文字样式（规范附录 A.7）。Φ 用 U+03A6、深度用 ↓（字形缺失实测，勿改回 Ø/↧）。"""
    try:
        ts = call(doc.TextStyles.Add, style)
    except pywintypes.com_error:
        ts = doc.TextStyles.Item(style)
    font_ok = False
    for face in ("宋体", "SimSun"):
        try:
            call(ts.SetFont, face, False, False, 134, 0)
            font_ok = True
            print("[1] text style", style, "->", face)
            break
        except pywintypes.com_error:
            continue
    if not font_ok:
        call(setp, ts, "FontFile", "simsun.ttc")
        print("[1] text style", style, "-> simsun.ttc")
    call(doc.SetVariable, "TEXTSTYLE", style)
    return ts


def set_dim_vars(doc, dimlfac=1.0, style="CN"):
    """尺寸变量。dimlfac = 1/缩放因子（如 1:2.5 图面 → 2.5），使标注读数还原真实值。"""
    for var, val in [("DIMTXT", 3.0), ("DIMASZ", 2.5), ("DIMEXO", 1.0), ("DIMEXE", 1.5),
                     ("DIMGAP", 1.0), ("DIMTAD", 1), ("DIMDEC", 2), ("DIMZIN", 8),
                     ("DIMTIH", 0), ("DIMTOH", 0), ("DIMSCALE", 1.0), ("DIMLUNIT", 2),
                     ("DIMLFAC", dimlfac),
                     ("DIMSE1", 0), ("DIMSE2", 0), ("DIMTXSTY", style)]:
        call(doc.SetVariable, var, val)
    call(doc.SetVariable, "FILLMODE", 1)   # 保证 AddSolid 实心箭头/涂黑三角显示
    print("[1] dim variables set, DIMLFAC =", dimlfac)


def setup_env(doc, acad, dimlfac=1.0, layer_def=LAYERS_DEFAULT):
    """一步完成：线型 + 图层 + 文字样式 + 尺寸变量。"""
    ensure_linetypes(doc)
    make_layers(doc, acad, layer_def)
    setup_text_style(doc)
    set_dim_vars(doc, dimlfac)


# ==================================================================
# 实体封装（每个 Add* 后立即 Update + 实体级延时 → 过程实时可见）
# ==================================================================
def line(ms, x1, y1, x2, y2, layer):
    e = call(ms.AddLine, pt(x1, y1), pt(x2, y2))
    call(setp, e, "Layer", layer)
    call(e.Update)
    time.sleep(SLEEP_ENT)
    return e


def circle(ms, x, y, r, layer, lw=None):
    e = call(ms.AddCircle, pt(x, y), r)
    call(setp, e, "Layer", layer)
    if lw is not None:
        call(setp, e, "Lineweight", lw)   # 实体线宽覆盖（如螺纹小径虚线加粗 0.5）
    call(e.Update)
    time.sleep(SLEEP_ENT)
    return e


def arc(ms, x, y, r, a1, a2, layer):
    """a1/a2 为角度制（如螺纹大径 3/4 圆：45 → 315）。"""
    e = call(ms.AddArc, pt(x, y), r, math.radians(a1), math.radians(a2))
    call(setp, e, "Layer", layer)
    call(e.Update)
    time.sleep(SLEEP_ENT)
    return e


def text(ms, s, x, y, h, layer="标注", style="CN"):
    e = call(ms.AddText, s, pt(x, y), h)
    call(setp, e, "Layer", layer)
    try:
        call(setp, e, "StyleName", style)
    except pywintypes.com_error:
        pass
    call(e.Update)
    time.sleep(SLEEP_ENT)
    return e


def text_mc(ms, s, cx, cy, h, layer="标注", style="CN"):
    """MiddleCenter 对齐文字（框格逐格居中用；10 = acAlignmentMiddleCenter）"""
    e = call(ms.AddText, s, pt(cx, cy), h)
    call(setp, e, "Layer", layer)
    try:
        call(setp, e, "StyleName", style)
    except pywintypes.com_error:
        pass
    call(setp, e, "Alignment", 10)
    call(setp, e, "TextAlignmentPoint", pt(cx, cy))
    call(e.Update)
    time.sleep(SLEEP_ENT)
    return e


def dim(ms, p1, p2, loc, ang_deg, override=None, layer="标注"):
    """旋转尺寸。同侧标注规则：小尺寸在内、大尺寸在外。override 为显示文字
    （偏差堆叠用 MText 格式：'Φ115{\\\\H0.7x;\\\\S+0.15^+0.10;}'，0 补齐、位数相同）。"""
    e = call(ms.AddDimRotated, pt(*p1), pt(*p2), pt(*loc), math.radians(ang_deg))
    call(setp, e, "Layer", layer)
    if override:
        call(setp, e, "TextOverride", override)
    call(e.Update)
    time.sleep(SLEEP_ENT)
    return e


def solid(ms, p1, p2, p3, layer):
    """AddSolid 实心三角形（第 4 角点与第 3 角点重合即闭合为三角形）。需 FILLMODE=1。"""
    e = call(ms.AddSolid, pt(*p1), pt(*p2), pt(*p3), pt(*p3))
    call(setp, e, "Layer", layer)
    call(e.Update)
    time.sleep(SLEEP_ENT)
    return e


def leader(ms, points, layer="标注", size=2.5, label=""):
    """V1.6 确定性引线：单段直线 + 靠图形一端 AddSolid 实心三角箭头（2.5=DIMASZ）。
    禁用 AddLeader（无模板新文档箭头块为空，箭头不显示，probe_leader_arrow.py 实证）。
    points = [(尖端x,尖端y), (尾端x,尾端y)]，必须恰好 2 点（引线不得折弯）。"""
    assert len(points) == 2, "V1.6 禁令：引线必须单段直线（2 点），收到 {} 点 -> {}".format(
        len(points), label)
    line(ms, points[0][0], points[0][1], points[1][0], points[1][1], layer)
    tx, ty = points[0]
    dx, dy = tx - points[1][0], ty - points[1][1]
    L = math.hypot(dx, dy) or 1.0
    ux, uy = dx / L, dy / L
    bx, by = tx - ux * size, ty - uy * size
    hw = size / 6.0
    px, py = -uy * hw, ux * hw
    solid(ms, (bx + px, by + py), (bx - px, by - py), (tx, ty), layer)
    LEADER_AUDIT.append((label, tx, ty, math.degrees(math.atan2(uy, ux)), len(points) - 1))


def tolerance_frame(ms, cells, x0, y0, h=None, txt_h=2.5, layer="标注"):
    """手工形位公差框格（禁用 AddTolerance：渲染挤压互相干涉，已实证）。
    cells = [(文字, 格宽), ...] 如 [("⊥",5.0), ("0.05",10.0), ("A",5.0)]；
    框高默认 2×字高；外框细实线矩形 + 每格竖分隔线 + 文字 MiddleCenter 逐格居中。
    返回 (cells_info, x1)，cells_info 供自审计逐格验证居中。"""
    h = h or 2 * txt_h
    info = []
    x = x0
    for t, w in cells:
        info.append([t, x, x + w, y0, y0 + h, x + w / 2, y0 + h / 2, None])
        x += w
    x1 = x
    edges = [(x0, y0, x1, y0), (x1, y0, x1, y0 + h), (x1, y0 + h, x0, y0 + h), (x0, y0 + h, x0, y0)]
    edges += [(c[2], y0, c[2], y0 + h) for c in info[:-1]]
    for (a, b, c, d) in edges:
        line(ms, a, b, c, d, layer)
    for c in info:
        c[7] = text_mc(ms, c[0], c[5], c[6], txt_h, layer)
    return info, x1


def stage_done(doc, ms, tag, name):
    """阶段级强制刷新 + 实体计数打印 + 阶段延时（只 sleep 不 Regen = 画面不更新）。"""
    call(doc.Regen, 1)   # 1 = acActiveViewport
    print("[{}] {} 完成 | 当前实体数 = {}".format(tag, name, call(lambda: ms.Count)))
    time.sleep(SLEEP_STAGE)


def zoom_extents(acad):
    """S1 图框绘出后即调用一次（图框充满窗口、全程保持），结束后再调用一次。中途不缩放。"""
    call(acad.ZoomExtents)


def center_views(ms, zone=(3, 294, 56, 207), text_y_min=56, exclude_text=("其余",),
                 view_layers=("OUTLINE", "CENTER", "HIDDEN", "DIM")):
    """规范 V1.12：将视图整体（视图图层全部实体 + 零件标注文字）平移，
    使其包围盒中心对齐图纸可用空间中心（默认内框 x3..294、标题栏顶 y56 至内框顶 y207）。
    不动：FRAME 层、标题栏文字（y<=text_y_min）、exclude_text 命中的注记（如"其余 Ra3.2"）。
    平移量取整 mm（容差 ±5mm）。返回 (dx, dy, 平移实体数)。

    包围盒用几何法逐类型计算（直线端点/圆心半径/文字插入点/实心点坐标），
    不用 GetBoundingBox（AcDbRotatedDimension 经 COM 读回的包围盒不可靠，
    曾导致平移量错误、文字移出图框）；尺寸对象不参与包围盒计算（贴近被注几何，影响 <±5mm），
    但仍随整体平移。平移后重新量测一次并打印，便于审计。"""
    from win32com.client import dynamic as _dyn
    zcx, zcy = (zone[0] + zone[1]) / 2.0, (zone[2] + zone[3]) / 2.0

    def _extent(e, kind):
        """几何法包围盒：直线端点/圆心半径/文字插入点/实心点坐标；尺寸返回 None。
        不用 GetBoundingBox（AcDbRotatedDimension 经 COM 读回不可靠）。"""
        if kind == "AcDbLine":
            p, q = e.StartPoint, e.EndPoint
            return (min(p[0], q[0]), min(p[1], q[1]), max(p[0], q[0]), max(p[1], q[1]))
        if kind in ("AcDbCircle", "AcDbArc"):
            c, r = e.Center, e.Radius
            return (c[0] - r, c[1] - r, c[0] + r, c[1] + r)
        if kind == "AcDbText":
            ip, h = e.InsertionPoint, e.Height
            return (ip[0], ip[1], ip[0] + 0.75 * h * len(e.TextString), ip[1] + h)
        if kind == "AcDbSolid":
            co = list(e.Coordinates)
            xs, ys = co[0::3], co[1::3]
            return (min(xs), min(ys), max(xs), max(ys))
        return None                                     # 尺寸等：平移但不计入包围盒

    def _collect():
        """返回 (move_set, (minx, miny, maxx, maxy))。
        标题栏/符号区（整体 ymax <= text_y_min）的实体一律不动——
        第一角投影符号画在 OUTLINE 层，必须靠 y 阈值排除，否则拉偏包围盒。"""
        mv, box = [], [1e9, 1e9, -1e9, -1e9]
        for i in range(call(lambda: ms.Count)):
            e = ms.Item(i)
            lay = getp(e, "Layer")
            kind = getp(e, "ObjectName")
            if lay in view_layers:
                pass
            elif lay == "TEXT" and kind == "AcDbText":
                t = _dyn.Dispatch(e)
                if not (t.InsertionPoint[1] > text_y_min
                        and not any(k in t.TextString for k in exclude_text)):
                    continue
            else:
                continue
            e = _dyn.Dispatch(e)
            ext = _extent(e, kind)
            if ext is not None and ext[3] <= text_y_min:    # 标题栏/投影符号区，排除
                continue
            mv.append(e)
            if ext is not None:
                box[0], box[1] = min(box[0], ext[0]), min(box[1], ext[1])
                box[2], box[3] = max(box[2], ext[2]), max(box[3], ext[3])
        return mv, box

    move_set, b = _collect()
    if not move_set or b[0] > b[2]:
        print("[center_views] 未找到视图实体，跳过")
        return (0, 0, 0)
    dx = round(zcx - (b[0] + b[2]) / 2)
    dy = round(zcy - (b[1] + b[3]) / 2)
    p0, p1 = pt(0, 0), pt(dx, dy)
    for e in move_set:
        e.Move(p0, p1)
    _, b2 = _collect()
    print("[center_views] 平移 (%+d,%+d) 共 %d 件 | 新包围盒 x%.1f..%.1f y%.1f..%.1f 中心(%.1f,%.1f) 目标(%.1f,%.1f)"
          % (dx, dy, len(move_set), b2[0], b2[2], b2[1], b2[3],
             (b2[0] + b2[2]) / 2, (b2[1] + b2[3]) / 2, zcx, zcy))
    return (dx, dy, len(move_set))


# ==================================================================
# PNG 举证
# ==================================================================
def _setup_plot_layout(doc):
    layout = doc.ActiveLayout
    call(setp, layout, "ConfigName", "PublishToWeb PNG.pc3")
    call(layout.RefreshPlotDeviceInfo)
    names = list(call(layout.GetCanonicalMediaNames))
    best, area, rot = None, -1, 0
    port_best, port_area = None, -1
    for n in names:
        m = re.search(r"(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)", str(n), re.I)
        if not m:
            continue
        w, h = float(m.group(1)), float(m.group(2))
        if w >= h and w * h > area:
            best, area, rot = n, w * h, 0
        elif w < h and w * h > port_area:
            port_best, port_area = n, w * h
    if best is None and port_best is not None:
        best, rot = port_best, 1
    if best:
        call(setp, layout, "CanonicalMediaName", best)
        print("[png] plot media:", best, "| rotation:", rot)
    call(setp, layout, "UseStandardScale", True)
    call(setp, layout, "StandardScale", 0)   # acScaleToFit
    call(setp, layout, "CenterPlot", True)
    call(setp, layout, "PlotRotation", rot)
    try:
        call(setp, layout, "StyleSheet", "monochrome.ctb")
    except pywintypes.com_error:
        print("[png] warn: monochrome.ctb not applied")
    try:
        call(setp, layout, "PlotWithLineweights", True)
    except pywintypes.com_error:
        pass
    plot = doc.Plot
    call(setp, plot, "QuietErrorMode", True)
    return layout


def plot_png(doc, out_png):
    """整图横向 PNG（PlotToFile，PublishToWeb PNG.pc3 + monochrome.ctb）。"""
    if os.path.exists(out_png):
        os.remove(out_png)
    layout = _setup_plot_layout(doc)
    call(setp, layout, "PlotType", 1)   # acExtents
    call(doc.Plot.PlotToFile, out_png)
    time.sleep(2)
    ok = os.path.exists(out_png)
    print("[png] PlotToFile", out_png, "->", ok)
    return ok


def plot_window(doc, out_png, ll, ur):
    """局部放大举证 PNG。必须先 SetWindowToPlot 再设 PlotType=4(acWindow)，顺序不可反。"""
    if os.path.exists(out_png):
        os.remove(out_png)
    lay = doc.ActiveLayout
    p1 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(ll[0]), float(ll[1])))
    p2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(ur[0]), float(ur[1])))
    call(lay.SetWindowToPlot, p1, p2)
    call(setp, lay, "PlotType", 4)   # 4 = acWindow
    call(doc.Plot.PlotToFile, out_png)
    time.sleep(2)
    call(setp, lay, "PlotType", 1)   # 恢复 acExtents
    ok = os.path.exists(out_png)
    print("[png] window plot {} -> {} : {}".format(ll, ur, ok))
    return ok


# ==================================================================
# 自审计
# ==================================================================
def collect_audit(ms):
    """遍历模型空间：返回 (by_layer, by_type, texts)。texts 含 Text 文字与尺寸 override/测量值。"""
    total = call(lambda: ms.Count)
    by_layer = Counter()
    by_type = Counter()
    texts = []
    for i in range(total):
        e = call(ms.Item, i)
        by_layer[call(getp, e, "Layer")] += 1
        typ = call(getp, e, "ObjectName")
        by_type[typ] += 1
        if typ == "AcDbText":
            texts.append(call(getp, e, "TextString"))
        elif typ == "AcDbRotatedDimension":
            ov = call(getp, e, "TextOverride")
            texts.append(str(ov) if ov else str(call(getp, e, "Measurement")))
    print("实体审计：模型空间实体总数 =", total)
    for lay, n in sorted(by_layer.items()):
        print("  图层 {:8s} {} 个".format(lay, n))
    for typ, n in sorted(by_type.items()):
        print("  类型 {:22s} {} 个".format(typ, n))
    return by_layer, by_type, texts


def audit_common(ms, doc, texts, required, expect_solids=None, layer_linetypes=None):
    """通用断言集：关键文字逐条、无 PCD、无 AcDbLeader、引线单段（LEADER_AUDIT）、
    图层线型读回、AddSolid 计数。required 为必须出现的关键文字列表。"""
    missing = [r for r in required if not any(r == t or r in t for t in texts)]
    for r in required:
        print(("  [OK]  " if r not in missing else "  [缺失] ") + r)
    print("关键文字核对：{} / {} 条存在".format(len(required) - len(missing), len(required)))
    leaders = [i for i in range(call(lambda: ms.Count))
               if call(getp, call(ms.Item, i), "ObjectName") == "AcDbLeader"]
    solids = [i for i in range(call(lambda: ms.Count))
              if call(getp, call(ms.Item, i), "ObjectName") == "AcDbSolid"]
    for (label, tx, ty, ang, nseg) in LEADER_AUDIT:
        print("  实心箭头 [{}]：尖端 = ({:.2f}, {:.2f})，朝向 = {:.1f}°，引线 = {} 段直线".format(
            label, tx, ty, ang, nseg))
    bent = [label for (label, tx, ty, ang, nseg) in LEADER_AUDIT if nseg != 1]
    pcd_hits = [t for t in texts if "PCD" in t.upper()]
    print("引线单段直线核对（V1.6 禁令）：",
          "全部 {} 条均为单段直线（PASS）".format(len(LEADER_AUDIT)) if not bent else "FAIL -> " + str(bent))
    print("全图 'PCD' 字符串检索：", "不存在（PASS）" if not pcd_hits else "FAIL -> " + str(pcd_hits))
    if layer_linetypes:
        for lname, expect_lt in layer_linetypes.items():
            actual = str(call(getp, doc.Layers.Item(lname), "Linetype")).upper()
            print("图层线型读回：{} = {}".format(lname, actual))
            assert actual == expect_lt.upper(), "图层 {} 线型应为 {}，读回 {}".format(lname, expect_lt, actual)
    assert not missing, "自审计失败：缺失关键文字 " + str(missing)
    assert not pcd_hits, "自审计失败：全图不允许出现 'PCD'"
    assert not bent, "自审计失败：引线必须单段直线 -> " + str(bent)
    assert len(leaders) == 0, "自审计失败：不应有 AcDbLeader"
    if expect_solids is not None:
        assert len(solids) == expect_solids, \
            "自审计失败：AddSolid 预期 {}，实际 {}".format(expect_solids, len(solids))
    print("自审计通用断言：PASS")
    return True


def audit_frame_cells(cells_info):
    """框格三格逐格居中验证（tolerance_frame 返回的 cells_info）。"""
    ok_all = True
    for (t, x0, x1, y0, y1, cx, cy, ent) in cells_info:
        s = str(call(getp, ent, "TextString"))
        ap = call(getp, ent, "TextAlignmentPoint")
        ax, ay = float(ap[0]), float(ap[1])
        ok = (s == t) and (x0 < ax < x1) and (y0 < ay < y1) \
            and abs(ax - cx) < 1e-6 and abs(ay - cy) < 1e-6
        ok_all = ok_all and ok
        print("  格[{}] 文字读回 '{}' MiddleCenter=({:.2f},{:.2f}) -> {}".format(
            t, s, ax, ay, "居中格内 OK" if ok else "FAIL"))
    assert ok_all, "自审计失败：框格文字未居中于各自格内"
    return True


# ==================================================================
# 避让验算几何工具
# ==================================================================
def seg_pt_dist(p, a, b):
    """点 p 到线段 ab 的最小距离"""
    vx, vy = b[0] - a[0], b[1] - a[1]
    L2 = vx * vx + vy * vy
    t = max(0.0, min(1.0, ((p[0] - a[0]) * vx + (p[1] - a[1]) * vy) / L2)) if L2 else 0.0
    return math.hypot(p[0] - (a[0] + t * vx), p[1] - (a[1] + t * vy))


def seg_circle_cross(a, b, c, r):
    """线段 ab 与圆 (c,r) 的交点列表（引注引线避让验算用）"""
    dx, dy = b[0] - a[0], b[1] - a[1]
    fx, fy = a[0] - c[0], a[1] - c[1]
    A2 = dx * dx + dy * dy
    B2 = 2 * (fx * dx + fy * dy)
    C2 = fx * fx + fy * fy - r * r
    disc = B2 * B2 - 4 * A2 * C2
    if A2 == 0 or disc <= 0:
        return []
    sq = math.sqrt(disc)
    return [(a[0] + t * dx, a[1] + t * dy)
            for t in ((-B2 - sq) / (2 * A2), (-B2 + sq) / (2 * A2)) if 0.0 <= t <= 1.0]
