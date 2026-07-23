---
name: AI制图规范V1.5
description: 在 AutoCAD 中按《CAD制图的AI规范》实时逐步绘制机械工程图（DWG，COM/pywin32 驱动）。凡用户请求带有"制图""绘图""工程图""按制图规范出图""CAD出图""画零件图/图纸"等，或要求生成 DWG 图纸、按指定尺寸/孔系/公差出机械零件图时，必须使用本技能。覆盖平板类、回转体类零件的多视图表达、尺寸/引注/形位公差/基准标注、A4 图框标题栏、逐步可见绘制、COM 自审计与 PNG 举证验收全流程。内置《CAD制图的AI规范 V1.12》全文于 references/，开箱即用。
---

# CAD 工程图绘制（AutoCAD 实时出图）

## 0. 铁律（每次绘图前默念）

1. **先读规范，再动笔**。必读两份文件（单次会话读一次即可）：
   - `CAD制图的AI规范V1.1.md`（内容版本以文档标题为准，当前 V1.12；文件名沿用 V1.1 是用户指定，不要改名）
   - `制图规范存档说明.md`
   读取顺序：**优先读工作区 `C:\Users\wu\Documents\kimi\workspace\` 下的版本（用户持续更新）；
   工作区没有时，读技能包内置副本 `references/` 下的同名文件**（打包分发时规范随技能一起走）。
2. **用户指令优先于规范默认**。用户指定的句式/层名/表达方式（如"（均布）"、五层中文层名、侧视图不剖）逐字照抄，并在最终汇报中说明与规范默认的差异。用户未指定的才按规范默认执行。
3. **禁止任何保存操作**：不 SaveAs、不保存新文档、绝不触碰磁盘已有 .dwg（运行前后比对 mtime）。图纸画完保持打开，由用户手动保存。
4. **逐步可见**：每个实体 Add* 后 `Update()`，每阶段 `Regen`；实体级延时 0.3s、阶段级 0.9s。只 sleep 不刷新画面＝用户看不到过程，属于失败。不改 AutoCAD 窗口状态；S1 图框绘出后即 `ZoomExtents`，结束再做一次。
5. **引线**：单段直线 + `AddSolid` 实心三角箭头（尺寸 2.5 = DIMASZ，FILLMODE=1）。**禁用 `AddLeader`**（无模板新文档箭头块为空，箭头不显示，已实证）。引线不得折弯（V1.6 禁令）。
6. **自审计不过＝未交付**：COM 实体审计 + 整图/局部放大 PNG 举证 + 亲自目视验收 PNG 后，才能向用户汇报完成。

## 1. 流程定式（9 步）

1. **读规范两文件**（见铁律 1），提取本次适用条款。
2. **几何分析**：孔位坐标表（阵列角、分度圆半径）、避让验算（引注/引线不穿孔、不压图）、视图布置规划。写进脚本 docstring。
3. **比例评估前置**（规范 V4.3）：含标注的总高 < 150mm 可用高、要素净距 ≥ 15mm，否则降比例，取国标比例系列。算出缩放因子，`DIMLFAC = 1/因子` 使标注读数还原真实值。
4. **写/改脚本**：复制 `scripts/draw_template.py` 为 workspace 下的新脚本（如 `draw_fp009_acad.py`），`sys.path` 加入本技能 `scripts/` 目录，`import acad_common`。几何与标注坐标全部来自第 2、3 步的分析，不凭空写数。
5. **实时运行**：托管 Python `python draw_xxx.py`（pywin32 已装）。AutoCAD 须已打开；若 COM 被拒（RPC_REJECTED 持续），检查 CAD 是否有遗留对话框/文字编辑窗口挂起（可 PostMessage 发 ESC）。
6. **COM 自审计**：脚本尾部断言（实体计数、关键文字逐条、无 PCD、图层线型读回、引线单段、框格居中、.dwg mtime 未变、新文档 FullName 为空）。
7. **PNG 举证**：整图横向 PNG（PlotToFile + PublishToWeb PNG.pc3 + monochrome.ctb）+ 关键区域局部放大 PNG（先 `SetWindowToPlot` 再 `PlotType=4`）。
8. **亲自目视验收**：用读图工具查看 PNG，核对：尺寸小内大外、引注互不干涉、箭头指向正确、虚线/中心线线型正确、框格文字居中。
9. **汇报**：图号、比例、视图方案、标注清单、用户指令与规范默认的差异、审计结论、PNG 路径。提醒用户图纸未保存（手动 Ctrl+S）。

## 2. 公共库 acad_common（scripts/acad_common.py）

已封装全部验证过的机制，**不要重写这些代码，直接 import**：

- 连接/文档：`connect_acad()`、`new_drawing(acad)`（关遗留 Drawing*、设 INSUNITS/LWDISPLAY 等）、`snapshot_dwg(dir)`/`verify_dwg_unchanged()`
- 环境：`ensure_linetypes(doc)`（CENTER2/DASHED2，存在性检查+重试+读回校验+失败抛错，LTSCALE 自动适配 acadiso.lin/acad.lin）、`setup_text_style(doc)`（宋体 CN 样式）、`set_dim_vars(doc, dimlfac)`、`make_layers(doc, layer_def)`（TrueColor + 线型读回校验）
- 实体（每个自带 Update+延时，过程实时可见）：`line / circle / arc / text / text_mc / dim / solid / leader`（leader 强制单段直线+实心箭头并自动记录审计）
- 形位公差：`tolerance_frame(...)` 手工三格框格（⊥｜值｜基准，文字 MiddleCenter 逐格居中）——**禁用 AddTolerance**（渲染挤压互相干涉，已实证）
- 阶段/视野：`stage_done(...)`（Regen+实体计数+阶段延时）、`zoom_extents(acad)`
- 举证：`plot_png(doc, path)` 整图、`plot_window(doc, path, ll, ur)` 局部放大
- 审计：`collect_audit(ms)` 返回分层/分类型计数与全部文字

调用方式与顺序见 `scripts/draw_template.py`（可运行的骨架：9 阶段结构 + 自审计 + 举证 + 保存禁令验证）。

## 3. 固定实现方案（已被两张图验证，直接沿用）

| 项 | 方案 |
|---|---|
| 字体 | 宋体（TextStyle "CN"，SetFont("宋体")，回退 SimSun/simsun.ttc） |
| 直径符号 | `Φ` 用 U+03A6（Ø U+00D8 宋体缺字形）；正文叙述可用 Ø |
| 深度符号 | ↧ U+21A7 宋体缺字形 → 用 `↓` U+2193，汇报中说明 |
| 线型 | CENTER2 / DASHED2（acadiso.lin），LTSCALE=0.4（acad.lin 时 ×25.4） |
| 箭头 | AddSolid 实心三角，长 2.5、半宽 2.5/6；基准涂黑三角同法 |
| 偏差堆叠 | 尺寸文字 MText 堆叠 `{\\H0.7x;\\S+0.15^+0.10;}`，小数点对齐、位数相同、0 补齐 |
| 形位框格 | 手工画：外框细实线矩形（高=2×字高）、竖分隔线、文字逐格 MiddleCenter 居中 |
| 图框 | A4 按规范附录 A 坐标 1:1：外框 (0,0)-(297,210)、内框 (3,3)-(294,207)、标题栏全套、一角法符号、技术要求 3 条、未注公差表、右上"其余 Ra3.2" |

## 4. 标注规则速查（规范核心条款）

- 同侧尺寸：**小尺寸在内、大尺寸在外**。
- 均布孔句式：规范默认 `"EQS 6-M5 完全贯穿"` 式（EQS 在前、合并标注、引注离图形适当距离）；用户指定句式时照抄用户原文。
- 倒角引注（如 4-C2）**紧邻**被注倒角：引线取短（一般不超过约 10mm），文字在倒角外侧就近空旷处（V1.7），但不得干涉其它标注。
- 直径尺寸 45° 引出；引注互不干涉（过近则移位或旋转到空旷方向）。
- 分度圆只注 `"Φ80"` 式直径，不写 "PCD"（除非用户指定）。
- 腰槽（长圆槽）竖向定位尺寸：以两端半圆中心线之一为基准，不注槽中间中心（V1.9）；半圆中心线保持短画不延长，用 DIM 层细实线自中心线端部接画至尺寸界线（引出基线加长，V1.10）。
- 非两边对称的零件不画长对称中心线，仅保留孔/槽中心标记；仍有对称轴的方向保留该向中心线（V1.10）。
- 中心对称线超出零件轮廓以 3~5mm 为限，不得超出过多（V1.11）。
- 非两边对称方向的特征位置尺寸必须标注到零件外边缘，不得只注特征间距（如竖向孔除孔距 58 外补注下边缘→孔 6），同侧仍小内大外（V1.11）。
- 视图整体（含标注）须居于图纸可用空间中间：绘图最后调用 `acad_common.center_views(ms)`（S10 阶段），几何法包围盒（不用 GetBoundingBox），标题栏/投影符号区（y≤56）实体自动排除（V1.12）。
- 偏差：小数点对齐、上下偏差位数相同、0 补齐、字号 0.7×。

## 5. 已知坑（详见 references/pitfalls.md）

COM 被拒（ESC 解挂）、AddLeader 空箭头、AddTolerance 挤压、DASHED2 偶发退化为实线（读回校验兜底）、Ø/↧ 缺字形、局部放大必须先 SetWindowToPlot、只 sleep 不刷新、遗留 Drawing* 累积（new_drawing 已自动清理）、早期绑定属性缺失（setp/getp 动态回退已封装）。

## 6. 环境锚点

- AutoCAD 2020，COM ProgID `AutoCAD.Application`（GetInterfaceObject 用 `AutoCAD.AcCmColor.23`）
- 托管 Python 3.12 已装 pywin32；工作目录 `C:\Users\wu\Documents\kimi\workspace`
- 已验收样例脚本（workspace）：`draw_fp007_acad.py`（平板类）、`draw_fp008_acad.py`（回转体法兰，侧视图不剖 + 五层中文层名 + 框格三格分置）
