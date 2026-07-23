# -*- coding: utf-8 -*-
"""draw_template.py —— cad-drawing 技能绘图脚本骨架（复制到 workspace 改名后填充几何）

复制本文件为 workspace 下 draw_<图号>_acad.py，然后：
  1) 在 docstring 写清零件、视图方案、比例评估（V4.3）、避让验算结论；
  2) 按几何分析替换 S3~S8 的坐标（不要凭空写数）；
  3) 更新 REQUIRED 关键文字清单与断言预期值；
  4) 运行：python draw_<图号>_acad.py（AutoCAD 须已打开）。
阶段顺序（S1→S9 固定）：图框 → 标题栏 → 视图轮廓 → 中心线 → （其余视图/虚线）
  → 尺寸标注 → 引注/框格/基准 → 技术要求 → ZOOM → PNG 举证 → 自审计 → 保存禁令验证。
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))  # 与本文件同目录即可 import
import acad_common as A

WS = r"C:\Users\wu\Documents\kimi\workspace"
OUT_PNG = os.path.join(WS, "<图号>_DWG校验.png")
OUT_PNG_ZOOM = os.path.join(WS, "<图号>_放大举证.png")

T0 = time.time()

# ---- 0. 连接 + 新建文档（不保存；自动清理遗留 Drawing*、快照 .dwg mtime） ----
acad = A.connect_acad()
doc, ms = A.new_drawing(acad, WS)

# ---- 1. 环境（dimlfac = 1/缩放因子，如比例 1:2 → 2.0；用户指定层名时传自定义 layer_def） ----
A.setup_env(doc, acad, dimlfac=1.0, layer_def=A.LAYERS_DEFAULT)

# ---- S1. 图框（附录 A.1：A4 外框 + 内框） ----
print("=" * 60)
print("[S1] 阶段：图框...")
for (x1, y1, x2, y2) in [(0, 0, 297, 0), (297, 0, 297, 210), (297, 210, 0, 210), (0, 210, 0, 0),
                         (3, 3, 294, 3), (294, 3, 294, 207), (294, 207, 3, 207), (3, 207, 3, 3)]:
    A.line(ms, x1, y1, x2, y2, "粗实线")
A.stage_done(doc, ms, "S1", "图框")
A.zoom_extents(acad)                      # V4.4/V1.5：图框充满窗口，全程保持该视野
time.sleep(1.0)

# ---- S2. 标题栏（附录 A.2~A.4 坐标 1:1；文字按零件信息替换） ----
print("[S2] 阶段：标题栏...")
for (x1, y1, x2, y2) in [(82, 3, 82, 56), (82, 28, 294, 28), (82, 42, 294, 42),
                         (82, 56, 294, 56), (112, 3, 112, 42), (162, 3, 162, 28),
                         (222, 3, 222, 42), (257, 3, 257, 28)]:
    A.line(ms, x1, y1, x2, y2, "粗实线")
A.line(ms, 87.5, 18.25, 87.5, 22.75, "粗实线")      # 第一角投影符号
A.line(ms, 96.5, 17, 96.5, 24, "粗实线")
A.line(ms, 87.5, 22.75, 96.5, 24, "粗实线")
A.line(ms, 87.5, 18.25, 96.5, 17, "粗实线")
A.circle(ms, 103.5, 20.5, 3.5, "粗实线")
A.circle(ms, 103.5, 20.5, 1.75, "粗实线")
A.text(ms, "一角法", 91, 9, 2.5)
A.text(ms, "老吴ai学习有限公司", 152, 46.5, 5.0)
A.text(ms, "零件名称", 116, 33, 3.0)
A.text(ms, "<零件名>", 158, 33, 4.0)
A.text(ms, "图号", 226, 33, 3.0)
A.text_mc(ms, "<图号>", 264, 35, 4.0)              # V1.8：图号值 MiddleCenter 居中于值区 (264,35)
A.text(ms, "比例", 118, 19, 3.0)
A.text(ms, "<比例>", 140, 19, 4.0)                   # V4.3 评估值
A.text(ms, "单位", 118, 10, 3.0)
A.text(ms, "mm", 140, 10, 4.0)
A.text(ms, "材料", 166, 19, 3.0)
A.text(ms, "<材料>", 190, 19, 4.0)
A.text(ms, "表面处理", 166, 10, 3.0)
A.text(ms, "<处理>", 190, 10, 3.5)
A.text(ms, "数量", 226, 19, 3.0)
A.text(ms, "<数量>", 246, 19, 4.0)
A.text(ms, "图样", 226, 10, 3.0)
A.text(ms, "A4", 246, 10, 4.0)
A.text(ms, "设计/制图/审核", 259, 19.5, 3.0)
A.text(ms, "老吴ai学习", 260, 11, 3.0)
A.stage_done(doc, ms, "S2", "标题栏")

# ---- S3~S6. 视图（按几何分析填充：轮廓 → 中心线 → 其余视图/隐藏虚线） ----
# 例：A.circle(ms, cx, cy, r, "粗实线")  A.line(ms, ...)  A.arc(ms, x, y, r, 45, 315, "虚线")
# 回转体侧视图隐藏特征用 "虚线" 层；分度圆/轴线用 "中心线" 层。
print("[S3~S6] 阶段：视图绘制（按几何分析填充）...")
# TODO: 视图实体
A.stage_done(doc, ms, "S3~S6", "视图")

# ---- S7. 尺寸标注（同侧小内大外；偏差堆叠 'Φ115{\\H0.7x;\\S+0.15^+0.10;}'） ----
print("[S7] 阶段：尺寸标注...")
# TODO: A.dim(ms, p1, p2, loc, ang_deg, override)
A.stage_done(doc, ms, "S7", "尺寸标注")

# ---- S8. 引注 + 形位公差框格 + 基准符号（引线一律 A.leader，单段直线+实心箭头） ----
print("[S8] 阶段：引注/框格/基准...")
# TODO: A.leader(ms, [(尖端), (尾端)], label="...") + A.text(...)
# 形位公差框格：cells_info, x1 = A.tolerance_frame(ms, [("⊥",5.0),("0.05",10.0),("A",5.0)], x0, y0)
# 基准符号：A.solid 涂黑三角 + A.line 细线 + 方框 + 字母
A.stage_done(doc, ms, "S8", "引注/框格/基准")

# ---- S9. 技术要求 + 未注公差表 + 右上 其余 Ra3.2（附录 A.5） ----
print("[S9] 阶段：技术要求...")
for (x1, y1, x2, y2) in [(15, 3, 67, 3), (67, 3, 67, 17), (67, 17, 15, 17), (15, 17, 15, 3)]:
    A.line(ms, x1, y1, x2, y2, "粗实线")
A.text(ms, "未注线性尺寸公差", 17, 13.5, 2.2)
A.text(ms, "0.5~6:±0.05  6~30:±0.1", 17, 9.5, 2.2)
A.text(ms, "30~120:±0.15  120~315:±0.2", 17, 5.5, 2.2)
A.text(ms, "技术要求", 15, 44, 3.0)
A.text(ms, "1. 零件表面不可刮花、碰伤；", 15, 39, 2.5)
A.text(ms, "2. 去除所有毛刺、锐边；", 15, 34.5, 2.5)
A.text(ms, "3. 未注倒角 C0.5；", 15, 30, 2.5)
A.text(ms, "其余 Ra3.2", 252, 198, 3.0)
A.stage_done(doc, ms, "S9", "技术要求")

print("=" * 60)
print("[绘制] 全部阶段完成，逐步绘制总耗时 = {:.1f} 秒".format(time.time() - T0))

# ---- 10. 结束 ZOOM（中途不缩放）；不执行任何 SaveAs ----
A.zoom_extents(acad)
time.sleep(1.5)

# ---- 11. PNG 举证：整图 + 局部放大（先 SetWindowToPlot 再 PlotType=4，库内已保证） ----
A.plot_png(doc, OUT_PNG)
A.plot_window(doc, OUT_PNG_ZOOM, (140, 60), (250, 200))   # TODO: 改为关键区域坐标

# ---- 12. 自审计（断言不过 = 未交付） ----
by_layer, by_type, texts = A.collect_audit(ms)
REQUIRED = ["<零件名>", "<图号>", "<材料>"]   # TODO: 全部关键标注文字
A.audit_common(ms, doc, texts, REQUIRED,
               expect_solids=None,              # TODO: 引线箭头数+基准三角数
               layer_linetypes={"虚线": "DASHED2", "中心线": "CENTER2"})
# A.audit_frame_cells(cells_info)             # 有形位框格时启用

# ---- 13. 保存禁令验证 ----
A.verify_dwg_unchanged(WS)
try:
    full_name = doc.FullName
except Exception:
    full_name = "<读取失败>"
print("新文档状态: Name =", doc.Name, "| FullName =", repr(full_name), "（空 = 从未保存）")
print("=" * 60)
