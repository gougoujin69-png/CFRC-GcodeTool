
# -*- coding: utf-8 -*-
"""
PathCutGcode_GUI_v5.py
--------------------------------------------------
- 每路径独立 Z；
- O/A/E/C/B/D 剪丝逻辑；
- D->E 后直接进入下一路径（不再抬到安全Z）；
- **换行符可选**：CRLF/LF；保存文件时使用 newline='' 保持行尾一致；
- **仅针对你最新要求的两点改动**：
  1) 从 O 点执行剪丝宏之后，到 C 点“快速挤出”之前，所有移动不再带 E 值；
  2) C 点处“快速挤出”改为宏命令 `FEED65`，不再输出 `G1 F.. E..`；
     若 M82（绝对E），内部仅同步 e_state += (L + extra)，以保持后续绝对E连续。
"""
import math
from dataclasses import dataclass, field
from typing import List, Tuple, Optional

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

Point = Tuple[float, float]


# --------------------------- 数据模型 ---------------------------

@dataclass
class PathData:
    name: str = "Path"
    points: List[Point] = field(default_factory=list)
    z: float = 0.20  # 本路径层高Z


@dataclass
class GenSettings:
    # 模式/单位
    use_absolute_xy: bool = True
    use_mm: bool = True

    # 坐标变换
    scale: float = 1.0
    offset_x: float = 0.0
    offset_y: float = 0.0
    flip_y: bool = False
    workarea_height: float = 0.0

    # 进给
    f_travel: float = 4800.0           # 空移
    f_print_normal: float = 1800.0     # 常规打印速度

    # E 挤出
    enable_extrude: bool = True
    e_per_mm: float = 0.0              # 每毫米挤出量
    use_abs_e: bool = True             # M82 / M83
    start_e: float = 0.0
    e_feed: float = 4800.0             # E 改变时的进给
    retract_len: float = 0.0           # 可选：到 A 后回抽（本版在 O→C 禁止 E，因此会跳过）
    prime_len: float = 0.0             # 可选：到 E 后回推

    # 剪丝流程参数
    L_back_from_A: float = 15.0        # L：从 A 回退到 O 的距离
    macro_at_O: str = "M280"           # O 点剪丝宏（可多行）

    safe_lift_dz: float = 3.0          # A 抬升到 B 的 ΔZ

    C_back_from_E: float = 10.0        # 从 E 反向到 C 的距离
    D_back_from_E: float = 2.0         # 从 E 反向到 D 的短距离

    extra_len_at_C: float = 0.60       # 用于同步绝对 E 的多出量（FEED65 宏替代快速挤出）
    purge_feed: float = 4800.0         # 兼容保留（不再用于输出）

    approach_feed: float = 3600.0      # C→D 逼近速度
    undershoot_dz: float = 0.05        # D 点低于 Z2 的 Δz

    pause_ms_at_D: int = 500           # D 点暂停
    slow_to_E_feed: float = 600.0      # D→E 缓慢打印速度

    # 文件换行
    line_ending: str = "CRLF"          # CRLF 或 LF

    # 新建路径默认Z
    default_path_z: float = 0.20


# --------------------------- 工具函数 ---------------------------

def fmt(v: float) -> str:
    return f"{v:.3f}"

def hypot2(ax: float, ay: float, bx: float, by: float) -> float:
    return math.hypot(bx-ax, by-ay)

def transform_xy(x: float, y: float, s: GenSettings) -> Point:
    tx = s.offset_x + s.scale * x
    if s.flip_y:
        ty = s.offset_y + (s.workarea_height - s.scale * y)
    else:
        ty = s.offset_y + s.scale * y
    return (tx, ty)

def join_lines(lines: List[str], ending: str) -> str:
    sep = "\r\n" if str(ending).upper() == "CRLF" else "\n"
    return sep.join(line.rstrip("\r\n") for line in lines) + sep

def point_on_polyline_from_end(tpoints: List[Point], retreat: float) -> Tuple[Point, int, float]:
    """沿折线末端反向量距 retreat，返回：(坐标, 段起点索引 i, 段内比例 t)。"""
    if not tpoints: 
        return (0.0, 0.0), 0, 0.0
    need = max(0.0, retreat)
    for i in range(len(tpoints)-2, -1, -1):
        p1 = tpoints[i]; p2 = tpoints[i+1]
        d = hypot2(p1[0], p1[1], p2[0], p2[1])
        if d >= need:
            if d == 0: return p2, i, 1.0
            t = (d - need) / d
            ox = p1[0] + t * (p2[0] - p1[0])
            oy = p1[1] + t * (p2[1] - p1[1])
            return (ox, oy), i, t
        else:
            need -= d
    return tpoints[0], 0, 0.0

def direction_at_start(tpoints: List[Point]) -> Point:
    """路径起点方向单位向量（E->第二点）。"""
    for i in range(len(tpoints)-1):
        p0, p1 = tpoints[i], tpoints[i+1]
        d = hypot2(p0[0], p0[1], p1[0], p1[1])
        if d > 1e-9:
            return ((p1[0]-p0[0])/d, (p1[1]-p0[1])/d)
    return (1.0, 0.0)  # 默认X正向

def add_move_no_e(lines: List[str], *, x: Optional[float]=None, y: Optional[float]=None,
                  z: Optional[float]=None, f: Optional[float]=None, comment: str=""):
    parts = ["G1"]
    if x is not None: parts.append(f"X{fmt(x)}")
    if y is not None: parts.append(f"Y{fmt(y)}")
    if z is not None: parts.append(f"Z{fmt(z)}")
    if f is not None: parts.append(f"F{fmt(f)}")
    if comment: parts.append(f"; {comment}")
    lines.append(" ".join(parts))

def add_e_move(lines: List[str], x: float, y: float, z: Optional[float], seg_len: float,
               s: GenSettings, e_state: dict, feed: float, comment: str=""):
    """带 E 的打印段。"""
    if s.enable_extrude and s.e_per_mm > 0:
        de = seg_len * s.e_per_mm
        if s.use_abs_e:
            e_state['E'] += de
            if z is None:
                lines.append(f"G1 X{fmt(x)} Y{fmt(y)} E{fmt(e_state['E'])} F{fmt(feed)}{(' ; '+comment) if comment else ''}")
            else:
                lines.append(f"G1 X{fmt(x)} Y{fmt(y)} Z{fmt(z)} E{fmt(e_state['E'])} F{fmt(feed)}{(' ; '+comment) if comment else ''}")
        else:
            if z is None:
                lines.append(f"G1 X{fmt(x)} Y{fmt(y)} E{fmt(de)} F{fmt(feed)}{(' ; '+comment) if comment else ''}")
            else:
                lines.append(f"G1 X{fmt(x)} Y{fmt(y)} Z{fmt(z)} E{fmt(de)} F{fmt(feed)}{(' ; '+comment) if comment else ''}")
    else:
        add_move_no_e(lines, x=x, y=y, z=z, f=feed, comment=comment)


# --------------------------- 生成 G-code ---------------------------

def generate_gcode(paths: List[PathData], s: GenSettings) -> str:
    lines: List[str] = []

    # 头部
    lines.append("G90" if s.use_absolute_xy else "G91")
    lines.append("G21" if s.use_mm else "G20")
    e_state = {'E': 0.0}
    if s.enable_extrude:
        lines.append("M82" if s.use_abs_e else "M83")
        if s.use_abs_e:
            e_state['E'] = s.start_e
            lines.append(f"G92 E{fmt(e_state['E'])}")
    lines.append("")

    carry_E_for_next: Optional[Tuple[float,float,float]] = None  # (Ex,Ey,Ez) 若上条已驻留在本条起点

    n = len(paths)
    for idx in range(n):
        path = paths[idx]
        if len(path.points) < 2:
            continue
        z1 = path.z
        tpts1 = [transform_xy(x,y,s) for (x,y) in path.points]
        A = tpts1[-1]
        O, seg_i, t_seg = point_on_polyline_from_end(tpts1, s.L_back_from_A)

        # 下一路径
        has_next = idx+1 < n and len(paths[idx+1].points) >= 2
        if has_next:
            path2 = paths[idx+1]
            z2 = path2.z
            tpts2 = [transform_xy(x,y,s) for (x,y) in path2.points]
            E = tpts2[0]
            ux, uy = direction_at_start(tpts2)
            C = (E[0] - ux*s.C_back_from_E, E[1] - uy*s.C_back_from_E)
            D = (E[0] - ux*s.D_back_from_E, E[1] - uy*s.D_back_from_E)
        else:
            z2 = None; E = None; C = None; D = None

        lines.append(f"; ---- Path {idx+1}: {path.name} (Z={fmt(z1)}) ----")
        info = f";   O=({fmt(O[0])},{fmt(O[1])})  A=({fmt(A[0])},{fmt(A[1])})"
        if has_next:
            info += f"  C=({fmt(C[0])},{fmt(C[1])})  D=({fmt(D[0])},{fmt(D[1])})  E=({fmt(E[0])},{fmt(E[1])}) (Z2={fmt(z2)})"
        lines.append(info)

        P0 = tpts1[0]

        # —— 若上一条已把喷头停在本条 E（即 P0），跳过安全Z起刀 ——
        use_carry = False
        if carry_E_for_next is not None:
            ex, ey, ez = carry_E_for_next
            if abs(P0[0]-ex) < 1e-6 and abs(P0[1]-ey) < 1e-6:
                use_carry = True
        carry_E_for_next = None

        if use_carry:
            lines.append(f"G1 Z{fmt(z1)} F{fmt(s.f_print_normal)}")
            lines.append(f"G1 X{fmt(P0[0])} Y{fmt(P0[1])} F{fmt(s.f_print_normal)}")
        else:
            Bz_abs = z1 + s.safe_lift_dz
            lines.append(f"G1 Z{fmt(Bz_abs)} F{fmt(s.e_feed)}")
            lines.append(f"G1 X{fmt(P0[0])} Y{fmt(P0[1])} F{fmt(s.f_travel)}")
            lines.append(f"G1 Z{fmt(z1)} F{fmt(s.e_feed)}")

        # ---- 沿路径打印；在 O 点执行剪丝；O 之后至 C 前全部禁止 E ----
        last = P0
        postO_noE = False
        for j in range(1, len(tpts1)):
            cur = tpts1[j]
            if j-1 == seg_i:
                # last -> O（打印，带 E）
                seg_len = hypot2(last[0], last[1], O[0], O[1])
                add_e_move(lines, O[0], O[1], None, seg_len, s, e_state, s.f_print_normal, "to O")
                # O 点宏
                lines.append("; --- Macro at O ---")
                for ln in (s.macro_at_O or "M280").splitlines():
                    ln = ln.strip()
                    if ln:
                        lines.append(ln)
                # O->cur（禁止 E）
                add_move_no_e(lines, x=cur[0], y=cur[1], f=s.f_print_normal, comment="O->seg end (no E)")
                postO_noE = True
            else:
                seg_len = hypot2(last[0], last[1], cur[0], cur[1])
                if postO_noE:
                    add_move_no_e(lines, x=cur[0], y=cur[1], f=s.f_print_normal, comment="post-O (no E)")
                else:
                    add_e_move(lines, cur[0], cur[1], None, seg_len, s, e_state, s.f_print_normal, "")
            last = cur

        lines.append("; --- Reached A ---")

        # O→C 区间内禁止 E：跳过回抽
        # if s.enable_extrude and s.retract_len > 0: pass

        # 抬到 B（无 E）
        lines.append(f"G1 Z{fmt(z1 + s.safe_lift_dz)} F{fmt(s.e_feed)} ; lift to B")

        if not has_next:
            lines.append("")
            continue

        # B→C（仍安全Z，无 E）
        lines.append(f"G1 X{fmt(C[0])} Y{fmt(C[1])} F{fmt(s.f_travel)} ; move to C (safe Z)")

        # C 点：FEED65 宏（替代快速挤出）；若 M82 同步 e_state
        lines.append("FEED65 ; purge at C via macro")
        purge_len = s.L_back_from_A + s.extra_len_at_C
        if s.enable_extrude and s.use_abs_e and s.e_per_mm >= 0:
            e_state['E'] += purge_len  # 同步绝对 E 坐标（不输出 G1）

        # C->D 下降（无 E）
        ZD = max(0.0, z2 - s.undershoot_dz)
        lines.append(f"G1 X{fmt(D[0])} Y{fmt(D[1])} Z{fmt(ZD)} F{fmt(s.approach_feed)} ; C->D descend")

        # D 暂停
        if s.pause_ms_at_D > 0:
            lines.append(f"G4 P{s.pause_ms_at_D} ; dwell at D")

        # D->E 慢速打印，并升至 z2（恢复 E）
        seg_len_DE = hypot2(D[0], D[1], E[0], E[1])
        add_e_move(lines, E[0], E[1], z2, seg_len_DE, s, e_state, s.slow_to_E_feed, "D->E slow print")

        # 可选 prime（已允许 E）
        if s.enable_extrude and s.prime_len > 0:
            if s.use_abs_e:
                e_state['E'] += s.prime_len
                lines.append(f"G1 E{fmt(e_state['E'])} F{fmt(s.e_feed)} ; prime at E")
            else:
                lines.append(f"G1 E{fmt(s.prime_len)} F{fmt(s.e_feed)} ; prime at E")

        # 告知下一条：已在 E 且 Z=z2（跳过安全Z起刀）
        carry_E_for_next = (E[0], E[1], z2)

        lines.append("")

    return join_lines(lines, s.line_ending)


# --------------------------- GUI ---------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("路径编辑 + G-code 生成（v5：O→C 禁 E，C 用 FEED65）")
        self.geometry("1350x860")
        self.paths: List[PathData] = []
        self._build_ui()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=2)
        self.columnconfigure(2, weight=2)

        # 左：路径列表
        lf = ttk.LabelFrame(self, text="路径列表（每条路径可设置 Z）")
        lf.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        lf.rowconfigure(2, weight=1)
        for c in range(4): lf.columnconfigure(c, weight=1)

        self.lb = tk.Listbox(lf, exportselection=False)
        self.lb.grid(row=2, column=0, columnspan=4, sticky="nsew")
        self.lb.bind("<<ListboxSelect>>", self._on_select)

        ttk.Button(lf, text="新建路径", command=self.add_path).grid(row=0, column=0, sticky="ew", padx=2, pady=4)
        ttk.Button(lf, text="删除路径", command=self.del_path).grid(row=0, column=1, sticky="ew", padx=2, pady=4)
        ttk.Button(lf, text="上移", command=lambda: self.move_path(-1)).grid(row=0, column=2, sticky="ew", padx=2, pady=4)
        ttk.Button(lf, text="下移", command=lambda: self.move_path(+1)).grid(row=0, column=3, sticky="ew", padx=2, pady=4)

        ttk.Label(lf, text="路径名").grid(row=1, column=0, sticky="e")
        self.path_name_var = tk.StringVar()
        ttk.Entry(lf, textvariable=self.path_name_var).grid(row=1, column=1, sticky="ew")

        ttk.Label(lf, text="该路径层高 Z").grid(row=1, column=2, sticky="e")
        self.path_z_var = tk.DoubleVar(value=0.20)
        ttk.Entry(lf, textvariable=self.path_z_var).grid(row=1, column=3, sticky="ew")
        ttk.Button(lf, text="保存路径名&Z", command=self.save_path_meta).grid(row=3, column=0, columnspan=4, sticky="ew", pady=4)

        # 中：点编辑
        mf = ttk.LabelFrame(self, text="点编辑（选中路径）")
        mf.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        for i in range(6): mf.columnconfigure(i, weight=1)
        mf.rowconfigure(1, weight=1)

        self.tree = ttk.Treeview(mf, columns=("#","X","Y"), show="headings", selectmode="browse")
        for c,w in zip(("#","X","Y"),[60,120,120]):
            self.tree.heading(c, text=c); self.tree.column(c, width=w, anchor="center")
        self.tree.grid(row=1, column=0, columnspan=6, sticky="nsew", pady=4)

        ttk.Button(mf, text="添加点", command=self.add_point).grid(row=2, column=0, sticky="ew", pady=2)
        ttk.Button(mf, text="插入点(在选中之前)", command=self.ins_point).grid(row=2, column=1, sticky="ew", pady=2)
        ttk.Button(mf, text="删除点", command=self.del_point).grid(row=2, column=2, sticky="ew", pady=2)
        ttk.Button(mf, text="上移点", command=lambda: self.move_point(-1)).grid(row=2, column=3, sticky="ew", pady=2)
        ttk.Button(mf, text="下移点", command=lambda: self.move_point(+1)).grid(row=2, column=4, sticky="ew", pady=2)
        ttk.Button(mf, text="从文本粘贴", command=self.paste_points).grid(row=2, column=5, sticky="ew", pady=2)

        ttk.Button(mf, text="导入CSV", command=self.import_csv).grid(row=3, column=0, sticky="ew", pady=2)
        ttk.Button(mf, text="导出CSV", command=self.export_csv).grid(row=3, column=1, sticky="ew", pady=2)
        ttk.Button(mf, text="保存工程", command=self.save_project).grid(row=3, column=2, sticky="ew", pady=2)
        ttk.Button(mf, text="加载工程", command=self.load_project).grid(row=3, column=3, sticky="ew", pady=2)
        ttk.Button(mf, text="预览G-code", command=self.preview).grid(row=3, column=4, sticky="ew", pady=2)

        # 右：设置
        rf = ttk.LabelFrame(self, text="生成设置 & 剪丝参数 & 文件换行")
        rf.grid(row=0, column=2, sticky="nsew", padx=8, pady=8)
        for r in range(50): rf.rowconfigure(r, weight=0)
        rf.rowconfigure(49, weight=1)
        rf.columnconfigure(1, weight=1)

        # 基本
        self.v_abs = tk.BooleanVar(value=True)
        self.v_mm  = tk.BooleanVar(value=True)
        ttk.Checkbutton(rf, text="绝对坐标 G90", variable=self.v_abs).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(rf, text="单位mm G21", variable=self.v_mm).grid(row=0, column=1, sticky="w")

        self.v_scale = tk.DoubleVar(value=1.0)
        self.v_offx = tk.DoubleVar(value=0.0)
        self.v_offy = tk.DoubleVar(value=0.0)
        self.v_h    = tk.DoubleVar(value=0.0)
        row = 1
        for text,var in [("scale", self.v_scale), ("offsetX", self.v_offx), ("offsetY", self.v_offy)]:
            ttk.Label(rf, text=text).grid(row=row, column=0, sticky="e")
            ttk.Entry(rf, textvariable=var).grid(row=row, column=1, sticky="ew"); row += 1
        self.v_flip = tk.BooleanVar(value=False)
        ttk.Checkbutton(rf, text="Y翻转(配合工作台高度)", variable=self.v_flip).grid(row=row, column=0, sticky="w")
        ttk.Label(rf, text="工作台高度").grid(row=row, column=1, sticky="e")
        ttk.Entry(rf, textvariable=self.v_h).grid(row=row, column=2, sticky="ew"); row += 1

        # 进给
        self.v_ftr   = tk.DoubleVar(value=4800)
        self.v_fnorm = tk.DoubleVar(value=1800)
        ttk.Label(rf, text="空移F").grid(row=row, column=0, sticky="e")
        ttk.Entry(rf, textvariable=self.v_ftr).grid(row=row, column=1, sticky="ew"); row += 1
        ttk.Label(rf, text="常规打印F").grid(row=row, column=0, sticky="e")
        ttk.Entry(rf, textvariable=self.v_fnorm).grid(row=row, column=1, sticky="ew"); row += 1

        # E
        self.v_enable_e = tk.BooleanVar(value=True)
        self.v_abs_e = tk.BooleanVar(value=True)
        self.v_e_per_mm = tk.DoubleVar(value=0.0)
        self.v_e_start = tk.DoubleVar(value=0.0)
        self.v_e_feed  = tk.DoubleVar(value=4800)
        self.v_ret     = tk.DoubleVar(value=0.0)
        self.v_prime   = tk.DoubleVar(value=0.0)
        ttk.Checkbutton(rf, text="启用E挤出", variable=self.v_enable_e).grid(row=row, column=0, sticky="w"); row += 1
        ttk.Checkbutton(rf, text="M82绝对E(取消=M83相对E)", variable=self.v_abs_e).grid(row=row, column=0, sticky="w"); row += 1
        for text,var in [("E/mm", self.v_e_per_mm), ("E起始(绝对E)", self.v_e_start),
                         ("E进给F", self.v_e_feed), ("回抽", self.v_ret), ("回推", self.v_prime)]:
            ttk.Label(rf, text=text).grid(row=row, column=0, sticky="e")
            ttk.Entry(rf, textvariable=var).grid(row=row, column=1, sticky="ew"); row += 1

        # 剪丝参数
        self.v_L = tk.DoubleVar(value=15.0)
        self.v_macro = tk.Text(rf, height=3)
        self.v_macro.insert("1.0", "M280")
        self.v_lift = tk.DoubleVar(value=3.0)
        self.v_Cback = tk.DoubleVar(value=10.0)
        self.v_Dback = tk.DoubleVar(value=2.0)
        self.v_extra = tk.DoubleVar(value=0.60)
        self.v_purgeF = tk.DoubleVar(value=4800)  # 保留但不使用
        self.v_appF   = tk.DoubleVar(value=3600)
        self.v_under  = tk.DoubleVar(value=0.05)
        self.v_pause  = tk.IntVar(value=500)
        self.v_slowE  = tk.DoubleVar(value=600.0)

        for text,widget in [("O←A 距离 L(mm)", self.v_L),
                            ("抬升 ΔZ 到 B(mm)", self.v_lift),
                            ("C 距离：E反向(mm)", self.v_Cback),
                            ("D 距离：E反向(mm)", self.v_Dback),
                            ("C 处额外量 extra(mm)", self.v_extra),
                            ("C→D 逼近F", self.v_appF),
                            ("D 点低于Z的 Δz", self.v_under),
                            ("D 点暂停ms", self.v_pause),
                            ("D→E 缓慢打印F", self.v_slowE)]:
            ttk.Label(rf, text=text).grid(row=row, column=0, sticky="e")
            ttk.Entry(rf, textvariable=widget).grid(row=row, column=1, sticky="ew"); row += 1

        ttk.Label(rf, text="O 点宏(多行)").grid(row=row, column=0, sticky="w"); row += 1
        self.v_macro.grid(row=row, column=0, columnspan=3, sticky="nsew"); row += 1

        # 行尾
        ttk.Label(rf, text="换行符").grid(row=row, column=0, sticky="e")
        self.v_line = tk.StringVar(value="CRLF")
        ttk.Combobox(rf, textvariable=self.v_line, values=["CRLF","LF"], state="readonly").grid(row=row, column=1, sticky="ew"); row += 1

        ttk.Button(rf, text="生成G-code并保存…", command=self.save_gcode).grid(row=row, column=0, columnspan=3, sticky="ew", pady=6)

        self.status = tk.StringVar(value="就绪")
        ttk.Label(self, textvariable=self.status, anchor="w").grid(row=1, column=0, columnspan=3, sticky="ew")

        # 初始化
        self.refresh_lb()

    # ---- 路径/点操作 ----
    def refresh_lb(self):
        self.lb.delete(0, tk.END)
        for i, p in enumerate(self.paths, 1):
            self.lb.insert(tk.END, f"{i}. {p.name} (Z={p.z:.3f}, {len(p.points)} pts)")

    def sel_idx(self) -> Optional[int]:
        try:
            (i,) = self.lb.curselection(); return int(i)
        except Exception:
            return None

    def _on_select(self, _=None):
        i = self.sel_idx()
        self.tree.delete(*self.tree.get_children())
        if i is None or not (0 <= i < len(self.paths)):
            self.path_name_var = getattr(self, 'path_name_var', tk.StringVar())
            self.path_z_var = getattr(self, 'path_z_var', tk.DoubleVar(value=0.20))
            self.path_name_var.set(""); self.path_z_var.set(0.20); return
        p = self.paths[i]
        self.path_name_var.set(p.name)
        self.path_z_var.set(p.z)
        for k,(x,y) in enumerate(p.points,1):
            self.tree.insert("", "end", values=(k, f"{x:.6f}", f"{y:.6f}"))

    def save_path_meta(self):
        i = self.sel_idx()
        if i is None: return
        name = getattr(self, 'path_name_var').get().strip() or f"Path{i+1}"
        z = float(getattr(self, 'path_z_var').get() or 0.20)
        self.paths[i].name = name
        self.paths[i].z = z
        self.refresh_lb()

    def add_path(self):
        name = simpledialog.askstring("新建路径","路径名称：", parent=self) or f"Path{len(self.paths)+1}"
        default_z = self.collect_settings().default_path_z
        self.paths.append(PathData(name=name, z=default_z))
        self.refresh_lb()

    def del_path(self):
        i = self.sel_idx()
        if i is None: return
        if messagebox.askyesno("确认","删除选中路径？"):
            self.paths.pop(i); self.refresh_lb(); self._on_select()

    def move_path(self, delta:int):
        i = self.sel_idx(); 
        if i is None: return
        j = i + delta
        if 0 <= j < len(self.paths):
            self.paths[i], self.paths[j] = self.paths[j], self.paths[i]
            self.refresh_lb(); self.lb.selection_set(j); self._on_select()

    def add_point(self):
        i = self.sel_idx()
        if i is None: return
        x = simpledialog.askfloat("添加点","X：", parent=self)
        y = simpledialog.askfloat("添加点","Y：", parent=self)
        if x is None or y is None: return
        self.paths[i].points.append((float(x), float(y)))
        self._on_select(); self.refresh_lb()

    def ins_point(self):
        i = self.sel_idx()
        if i is None: return
        sel = self.tree.selection()
        pos = self.tree.index(sel[0]) if sel else 0
        x = simpledialog.askfloat("插入点","X：", parent=self)
        y = simpledialog.askfloat("插入点","Y：", parent=self)
        if x is None or y is None: return
        self.paths[i].points.insert(pos, (float(x), float(y)))
        self._on_select(); self.refresh_lb()

    def del_point(self):
        i = self.sel_idx()
        if i is None: return
        sel = self.tree.selection()
        if not sel: return
        pos = self.tree.index(sel[0])
        if messagebox.askyesno("确认","删除该点？"):
            self.paths[i].points.pop(pos)
            self._on_select(); self.refresh_lb()

    def move_point(self, delta:int):
        i = self.sel_idx()
        if i is None: return
        sel = self.tree.selection()
        if not sel: return
        pos = self.tree.index(sel[0]); j = pos + delta
        pts = self.paths[i].points
        if 0 <= j < len(pts):
            pts[pos], pts[j] = pts[j], pts[pos]
            self._on_select(); self.refresh_lb()

    def paste_points(self):
        i = self.sel_idx()
        if i is None: return
        text = simpledialog.askstring("从文本粘贴", "每行: x,y 或 x y", parent=self)
        if not text: return
        added = 0
        for ln in text.splitlines():
            ln = ln.strip().replace(",", " ")
            if not ln: continue
            sp = [p for p in ln.split() if p]
            if len(sp) >= 2:
                try:
                    x, y = float(sp[0]), float(sp[1])
                    self.paths[i].points.append((x,y)); added += 1
                except: pass
        self._on_select(); self.refresh_lb()
        self.status.set(f"新增 {added} 个点")

    # ---- 导入/导出/工程 ----
    def import_csv(self):
        import csv, os
        fp = filedialog.askopenfilename(title="导入CSV", filetypes=[("CSV","*.csv"),("All","*.*")])
        if not fp: return
        loaded: List[PathData] = []; by_name = {}
        try:
            with open(fp, "r", encoding="utf-8-sig", newline="") as f:
                rdr = csv.reader(f)
                # 支持 2列: x,y；3列: name,x,y；4列: name,x,y,z
                for row in rdr:
                    vals = [c.strip() for c in row if str(c).strip() != ""]
                    if not vals: continue
                    if len(vals) == 2:
                        if not loaded:
                            default_z = self.collect_settings().default_path_z
                            loaded.append(PathData(name=os.path.basename(fp), points=[], z=default_z))
                        x,y = float(vals[0]), float(vals[1])
                        loaded[0].points.append((x,y))
                    elif len(vals) == 3:
                        name, x, y = vals[0], float(vals[1]), float(vals[2])
                        if name not in by_name:
                            default_z = self.collect_settings().default_path_z
                            p = PathData(name=name, points=[], z=default_z)
                            by_name[name] = p; loaded.append(p)
                        by_name[name].points.append((x,y))
                    elif len(vals) >= 4:
                        name, x, y, z = vals[0], float(vals[1]), float(vals[2]), float(vals[3])
                        if name not in by_name:
                            p = PathData(name=name, points=[], z=z)
                            by_name[name] = p; loaded.append(p)
                        else:
                            by_name[name].z = z
                        by_name[name].points.append((x,y))
        except Exception as e:
            messagebox.showerror("错误", f"读取CSV失败：{e}"); return
        if not loaded:
            messagebox.showwarning("提示","CSV中没有有效数据"); return
        if messagebox.askyesno("导入", f"导入 {len(loaded)} 条路径并替换当前工程？"):
            self.paths = loaded
        else:
            self.paths.extend(loaded)
        self.refresh_lb(); self._on_select()
        self.status.set(f"已导入 CSV：{len(loaded)} 条路径")

    def export_csv(self):
        import csv
        if not self.paths:
            messagebox.showwarning("提示","无路径可导出"); return
        fp = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv")])
        if not fp: return
        with open(fp, "w", encoding="utf-8", newline="") as f:
            wr = csv.writer(f); wr.writerow(["name","x","y","z"])
            for p in self.paths:
                for x,y in p.points:
                    wr.writerow([p.name, x, y, p.z])
        self.status.set(f"CSV 已导出：{fp}")

    def save_project(self):
        import json
        fp = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON","*.json")])
        if not fp: return
        data = {
            "paths":[{"name":p.name, "z":p.z, "points":p.points} for p in self.paths],
            "settings": self.collect_settings().__dict__,
        }
        with open(fp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self.status.set(f"工程已保存：{fp}")

    def load_project(self):
        import json
        fp = filedialog.askopenfilename(filetypes=[("JSON","*.json"),("All","*.*")])
        if not fp: return
        try:
            with open(fp, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.paths = [PathData(name=p.get("name","Path"), z=float(p.get("z",0.20)), points=[tuple(pt) for pt in p.get("points",[])]) for p in data.get("paths",[])]
            self.apply_settings(data.get("settings", {}))
            self.refresh_lb(); self._on_select()
            self.status.set(f"工程已加载：{fp}")
        except Exception as e:
            messagebox.showerror("错误", f"读取失败：{e}")

    # ---- 预览/生成 ----
    def collect_settings(self) -> GenSettings:
        s = GenSettings()
        s.use_absolute_xy = self.v_abs.get()
        s.use_mm = self.v_mm.get()
        s.scale = float(self.v_scale.get() or 1.0)
        s.offset_x = float(self.v_offx.get() or 0.0)
        s.offset_y = float(self.v_offy.get() or 0.0)
        s.workarea_height = float(self.v_h.get() or 0.0)
        s.flip_y = self.v_flip.get()

        s.f_travel = float(self.v_ftr.get() or 4800)
        s.f_print_normal = float(self.v_fnorm.get() or 1800)

        s.enable_extrude = self.v_enable_e.get()
        s.use_abs_e = self.v_abs_e.get()
        s.e_per_mm = float(self.v_e_per_mm.get() or 0.0)
        s.start_e = float(self.v_e_start.get() or 0.0)
        s.e_feed  = float(self.v_e_feed.get() or 4800)
        s.retract_len = float(self.v_ret.get() or 0.0)
        s.prime_len   = float(self.v_prime.get() or 0.0)

        s.L_back_from_A = float(self.v_L.get() or 0.0)
        s.macro_at_O = self.v_macro.get("1.0","end").strip() or "M280"
        s.safe_lift_dz = float(self.v_lift.get() or 0.0)
        s.C_back_from_E = float(self.v_Cback.get() or 0.0)
        s.D_back_from_E = float(self.v_Dback.get() or 0.0)
        s.extra_len_at_C = float(self.v_extra.get() or 0.0)
        s.purge_feed = float(self.v_purgeF.get() or 0.0)  # 兼容保留
        s.approach_feed = float(self.v_appF.get() or 0.0)
        s.undershoot_dz = float(self.v_under.get() or 0.0)
        s.pause_ms_at_D = int(self.v_pause.get() or 0)
        s.slow_to_E_feed = float(self.v_slowE.get() or 0.0)

        s.line_ending = self.v_line.get() or "CRLF"
        return s

    def apply_settings(self, d: dict):
        self.v_abs.set(d.get("use_absolute_xy", True))
        self.v_mm.set(d.get("use_mm", True))
        self.v_scale.set(d.get("scale", 1.0))
        self.v_offx.set(d.get("offset_x", 0.0))
        self.v_offy.set(d.get("offset_y", 0.0))
        self.v_h.set(d.get("workarea_height", 0.0))
        self.v_flip.set(d.get("flip_y", False))

        self.v_ftr.set(d.get("f_travel", 4800))
        self.v_fnorm.set(d.get("f_print_normal", 1800))

        self.v_enable_e.set(d.get("enable_extrude", True))
        self.v_abs_e.set(d.get("use_abs_e", True))
        self.v_e_per_mm.set(d.get("e_per_mm", 0.0))
        self.v_e_start.set(d.get("start_e", 0.0))
        self.v_e_feed.set(d.get("e_feed", 4800))
        self.v_ret.set(d.get("retract_len", 0.0))
        self.v_prime.set(d.get("prime_len", 0.0))

        self.v_L.set(d.get("L_back_from_A", 15.0))
        self.v_lift.set(d.get("safe_lift_dz", 3.0))
        self.v_Cback.set(d.get("C_back_from_E", 10.0))
        self.v_Dback.set(d.get("D_back_from_E", 2.0))
        self.v_extra.set(d.get("extra_len_at_C", 0.60))
        self.v_purgeF.set(d.get("purge_feed", 4800))
        self.v_appF.set(d.get("approach_feed", 3600))
        self.v_under.set(d.get("undershoot_dz", 0.05))
        self.v_pause.set(d.get("pause_ms_at_D", 500))
        self.v_slowE.set(d.get("slow_to_E_feed", 600.0))
        self.v_line.set(d.get("line_ending", "CRLF"))

    def preview(self):
        if not self.paths or not any(p.points for p in self.paths):
            messagebox.showwarning("提示","请先添加路径/点"); return
        g = generate_gcode(self.paths, self.collect_settings())
        win = tk.Toplevel(self); win.title("G-code 预览"); win.geometry("1000x680")
        txt = tk.Text(win, wrap="none"); txt.pack(fill="both", expand=True)
        txt.insert("1.0", g)

    def save_gcode(self):
        if not self.paths or not any(p.points for p in self.paths):
            messagebox.showwarning("提示","请先添加路径/点"); return
        g = generate_gcode(self.paths, self.collect_settings())
        fp = filedialog.asksaveasfilename(defaultextension=".gcode", filetypes=[("G-code","*.gcode;*.nc;*.txt")])
        if not fp: return
        with open(fp, "w", encoding="utf-8", newline="") as f:  # newline='' 保留行尾
            f.write(g)
        self.status.set(f"G-code 已保存：{fp}")


def main():
    App().mainloop()

if __name__ == "__main__":
    main()
