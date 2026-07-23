# -*- coding: utf-8 -*-
"""cad-drawing 技能公共库冒烟测试：验证 acad_common 全部核心函数可用。
注意：不调用 new_drawing()（它会关闭遗留 Drawing*，会误关用户打开的未保存图纸），
改为手动 Documents.Add，结束后 Close(False) 自清，不留任何文档。"""
import os
import sys
import time

SKILL_SCRIPTS = r"C:\Users\wu\AppData\Roaming\kimi-desktop\daimon-share\daimon\skills\cad-drawing\scripts"
sys.path.insert(0, SKILL_SCRIPTS)
import acad_common as A

A.SLEEP_ENT = 0.05      # 冒烟测试加速（正式绘图保持默认 0.3/0.9）
A.SLEEP_STAGE = 0.1

WS = r"C:\Users\wu\Documents\kimi\workspace"
OUT_PNG = os.path.join(WS, "_smoke_cad_drawing.png")
OUT_ZOOM = os.path.join(WS, "_smoke_cad_drawing_zoom.png")

acad = A.connect_acad()
A.snapshot_dwg(WS)

doc = A.call(acad.Documents.Add)
A.call(doc.SetVariable, "INSUNITS", 4)
A.call(doc.SetVariable, "MEASUREMENT", 1)
A.call(doc.SetVariable, "BACKGROUNDPLOT", 0)
A.call(doc.SetVariable, "LWDISPLAY", 1)
ms = doc.ModelSpace
print("[smoke] test document:", doc.Name)

try:
    A.setup_env(doc, acad, dimlfac=1.0, layer_def=A.LAYERS_DEFAULT)

    # 图框 + ZOOM
    for (x1, y1, x2, y2) in [(0, 0, 297, 0), (297, 0, 297, 210), (297, 210, 0, 210), (0, 210, 0, 0),
                             (3, 3, 294, 3), (294, 3, 294, 207), (294, 207, 3, 207), (3, 207, 3, 3)]:
        A.line(ms, x1, y1, x2, y2, "粗实线")
    A.stage_done(doc, ms, "S1", "图框")
    A.zoom_extents(acad)
    time.sleep(0.5)

    # 视图：圆 + 中心线 + 虚线圆 + 圆弧
    A.circle(ms, 100, 130, 40, "粗实线")
    A.circle(ms, 100, 130, 20, "虚线")
    A.arc(ms, 100, 130, 30, 45, 315, "细实线")
    A.line(ms, 50, 130, 150, 130, "中心线")
    A.line(ms, 100, 80, 100, 180, "中心线")
    A.stage_done(doc, ms, "S2", "视图")

    # 尺寸（含偏差堆叠 override）+ 文字
    A.dim(ms, (60, 130), (140, 130), (100, 175), 0, "80")
    A.dim(ms, (140, 90), (140, 170), (152, 130), 90, "Φ80{\\H0.7x;\\S+0.15^+0.10;}")
    A.stage_done(doc, ms, "S3", "尺寸")

    # 引线 + 引注 + 三格框格 + 基准三角
    A.leader(ms, [(80, 158.28), (65, 170)], label="冒烟测试引注")
    A.text(ms, "冒烟测试 Φ60 " + A.DEPTH + "35", 30, 168.5, 2.5)
    cells, x1 = A.tolerance_frame(ms, [("⊥", 5.0), ("0.05", 10.0), ("A", 5.0)], 160, 140)
    A.solid(ms, (172.5, 107), (175.5, 107), (174, 110), "标注")
    A.text(ms, "A", 172.9, 99.2, 3.0)
    A.stage_done(doc, ms, "S4", "引注/框格/基准")

    # PNG 举证
    png_ok = A.plot_png(doc, OUT_PNG)
    zoom_ok = A.plot_window(doc, OUT_ZOOM, (40, 80), (190, 190))

    # 自审计
    by_layer, by_type, texts = A.collect_audit(ms)
    REQUIRED = ["冒烟测试 Φ60 " + A.DEPTH + "35", "80", "Φ80", "+0.15", "+0.10", "0.05", "A", "⊥"]
    A.audit_common(ms, doc, texts, REQUIRED,
                   expect_solids=2,   # 1 引线箭头 + 1 基准三角
                   layer_linetypes={"虚线": "DASHED2", "中心线": "CENTER2"})
    A.audit_frame_cells(cells)
    A.verify_dwg_unchanged(WS)
    assert png_ok and zoom_ok, "PNG 举证失败"
    print("[smoke] ALL PASS | png:", png_ok, "| zoom:", zoom_ok)
finally:
    A.call(doc.Close, False)   # 自清：不保存、不留文档
    print("[smoke] test document closed (not saved)")

# 清理测试 PNG
for f in (OUT_PNG, OUT_ZOOM):
    if os.path.exists(f):
        os.remove(f)
print("[smoke] done, png files cleaned")
