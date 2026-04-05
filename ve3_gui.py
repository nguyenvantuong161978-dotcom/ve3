#!/usr/bin/env python3
"""VE3 Studio — compact, professional GUI."""

import sys, os, shutil, threading, time as _time, json
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from datetime import datetime
from typing import Dict

VE3_DIR = Path(__file__).parent
sys.path.insert(0, str(VE3_DIR))

try:
    import customtkinter as ctk
except ImportError:
    os.system(f'"{sys.executable}" -m pip install customtkinter')
    import customtkinter as ctk
from PIL import Image

ctk.set_appearance_mode("light")

# ── palette ──────────────────────────────────────────────
AC = "#C00"           # accent red
AC2 = "#A00"
SB = "#1E1E1E"        # sidebar
SB2 = "#2D2D2D"
SB3 = "#3A3A3A"
BG = "#FAFAFA"
CD = "#FFF"           # card
BD = "#DDD"           # border
BD2 = "#EEE"
EN = "#F5F5F5"        # entry bg
T1 = "#111"           # text primary
T2 = "#555"
T3 = "#999"
OK = "#1B8"           # green
OK2 = "#169"
ER = "#D22"
RN = "#17C"           # running blue
TW, TH = 110, 74     # thumb
SW = 175              # sidebar width

BADGES = {
    "pending": (T3, "#F0F0F0",  "Đợi"),
    "running": ("#0D6EFD", "#E7F1FF", "Đang tạo"),
    "done":    ("#198754", "#D1E7DD", "Xong"),
    "error":   ("#DC3545", "#F8D7DA", "Lỗi"),
    "skip":    (T3, "#F0F0F0",  "Bỏ qua"),
}

def _thumb(p, w=TW, h=TH):
    try:
        if p and Path(p).exists():
            i = Image.open(str(p))
            return ctk.CTkImage(light_image=i, dark_image=i, size=(w, h))
    except Exception: pass
    return None

def _ph(w=TW, h=TH):
    i = Image.new("RGB", (w*2, h*2), "#E8E8E8")
    return ctk.CTkImage(light_image=i, dark_image=i, size=(w, h))

def _ts(s):
    if s is None: return ""
    s = int(s)
    return f"{s//60}:{s%60:02d}" if s >= 60 else f"{s}s"

# ── badge ────────────────────────────────────────────────
class Badge(ctk.CTkLabel):
    def __init__(self, master, st="pending", **k):
        fg, bg, tx = BADGES.get(st, BADGES["pending"])
        super().__init__(master, text=tx, text_color=fg, fg_color=bg,
                         corner_radius=8, font=("", 10, "bold"), padx=7, pady=1, **k)
    def set(self, st):
        fg, bg, tx = BADGES.get(st, BADGES["pending"])
        self.configure(text=tx, text_color=fg, fg_color=bg)

# ── character card ───────────────────────────────────────
class CharCard(ctk.CTkFrame):
    def __init__(self, master, d, nv, on_regen=None, on_view=None, **k):
        super().__init__(master, fg_color=CD, corner_radius=8,
                         border_width=1, border_color=BD2, height=90, **k)
        self.cid = d["id"]; self.nv = nv
        self.on_regen = on_regen; self.on_view = on_view
        self.grid_columnconfigure(1, weight=1)
        self.grid_propagate(False)

        # img
        self.img = ctk.CTkLabel(self, text="", width=TW, height=TH,
                                 fg_color="#ECECEC", corner_radius=4, cursor="hand2")
        self.img.grid(row=0, column=0, rowspan=2, padx=(6,4), pady=6)
        self.img.bind("<Button-1>", lambda e: self._view())
        self._load_img()

        # row 0: id·name·role  badge  time  server  regen
        r0 = ctk.CTkFrame(self, fg_color="transparent")
        r0.grid(row=0, column=1, sticky="ew", padx=(0,6), pady=(6,0))
        r0.grid_columnconfigure(0, weight=1)

        role = d.get("role",""); name = d.get("name","")
        t = self.cid
        if name: t += f" · {name}"
        if role: t += f" · {role}"
        ctk.CTkLabel(r0, text=t, font=("",12,"bold"), text_color=T1,
                     anchor="w").grid(row=0, column=0, sticky="w")

        st = (d.get("status") or "pending").lower()
        self.badge = Badge(r0, st if st in BADGES else "pending")
        self.badge.grid(row=0, column=1, padx=3)

        self.lbl_t = ctk.CTkLabel(r0, text="", font=("",9), text_color=T3)
        self.lbl_t.grid(row=0, column=2, padx=2)
        self.lbl_s = ctk.CTkLabel(r0, text="", font=("",9), text_color=T3)
        self.lbl_s.grid(row=0, column=3, padx=2)

        ctk.CTkButton(r0, text="Tạo lại", width=54, height=22, corner_radius=4,
                      fg_color="#EBEBEB", hover_color="#DDD", text_color=T2,
                      font=("",10), command=self._regen).grid(row=0, column=4, padx=(3,0))

        # row 1: prompt
        self.pb = ctk.CTkTextbox(self, height=36, font=("",11), fg_color=EN,
                                  border_color=BD2, border_width=1, corner_radius=4, wrap="word")
        self.pb.grid(row=1, column=1, sticky="ew", padx=(0,6), pady=(2,6))
        p = d.get("english_prompt") or d.get("vietnamese_prompt") or ""
        if p: self.pb.insert("1.0", p)

    def _regen(self):
        if self.on_regen: self.on_regen(self.cid, self.get_prompt())
    def _view(self):
        p = self.nv / f"{self.cid}.png"
        if p.exists() and self.on_view: self.on_view(p, self.cid)
    def _load_img(self):
        p = self.nv / f"{self.cid}.png"
        t = _thumb(p)
        if t: self.img.configure(image=t, text="", fg_color="transparent"); self.img._r = t
        else:
            ph = _ph(); self.img.configure(image=ph, text="", fg_color="#ECECEC"); self.img._r = ph
    def set_status(self, st, ex=None):
        self.badge.set(st)
        c = {"running": RN, "done": OK, "error": ER}.get(st)
        self.configure(border_color=c or BD2, border_width=2 if c else 1)
        if st == "done": self._load_img()
        ex = ex or {}
        if "elapsed" in ex: self.lbl_t.configure(text=_ts(ex["elapsed"]))
        if "server" in ex: self.lbl_s.configure(text=f'{ex["server"]}(q={ex.get("queue","?")})')
        if "queue_pos" in ex and ex["queue_pos"] is not None:
            self.lbl_s.configure(text=f'pos={ex["queue_pos"]}')
        if st == "running" and "elapsed" not in ex and "queue_pos" not in ex:
            self.lbl_t.configure(text="...")
    def get_prompt(self):
        return self.pb.get("1.0", "end-1c").strip()

# ── scene card ───────────────────────────────────────────
class SceneCard(ctk.CTkFrame):
    def __init__(self, master, d, idir, on_regen=None, on_regen_vid=None, on_view=None, **k):
        super().__init__(master, fg_color=CD, corner_radius=8,
                         border_width=1, border_color=BD2, height=110, **k)
        self.sid = d["scene_id"]; self.idir = idir
        self.on_regen = on_regen; self.on_regen_vid = on_regen_vid; self.on_view = on_view
        self.grid_columnconfigure(1, weight=1)
        self.grid_propagate(False)

        # Ảnh preview
        self.img = ctk.CTkLabel(self, text="", width=TW, height=TH,
                                 fg_color="#ECECEC", corner_radius=4, cursor="hand2")
        self.img.grid(row=0, column=0, rowspan=3, padx=(6,4), pady=6)
        self.img.bind("<Button-1>", lambda e: self._view())
        self._load_img()

        # Row 0: Scene ID + SRT + badge ảnh + thời gian + server + nút tạo lại ảnh
        r0 = ctk.CTkFrame(self, fg_color="transparent")
        r0.grid(row=0, column=1, sticky="ew", padx=(0,6), pady=(6,0))
        r0.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(r0, text=f"S{self.sid:03d}", font=("",12,"bold"),
                     text_color=T1).grid(row=0, column=0, sticky="w")
        srt = d.get("srt_text","")
        if srt:
            ctk.CTkLabel(r0, text=srt[:45]+("…" if len(srt)>45 else ""),
                         font=("",10), text_color=T3, anchor="w"
                         ).grid(row=0, column=1, sticky="w", padx=4)

        st = (d.get("status_img") or "pending").lower()
        self.badge = Badge(r0, st if st in BADGES else "pending")
        self.badge.grid(row=0, column=2, padx=2)

        self.lbl_t = ctk.CTkLabel(r0, text="", font=("",9), text_color=T3)
        self.lbl_t.grid(row=0, column=3, padx=2)
        self.lbl_s = ctk.CTkLabel(r0, text="", font=("",9), text_color=T3)
        self.lbl_s.grid(row=0, column=4, padx=2)

        ctk.CTkButton(r0, text="Tạo ảnh", width=54, height=22, corner_radius=4,
                      fg_color="#EBEBEB", hover_color="#DDD", text_color=T2,
                      font=("",10), command=self._regen).grid(row=0, column=5, padx=(2,0))

        # Row 0b: Video badge + nút tạo video
        stv = (d.get("status_vid") or "pending").lower()
        self.badge_vid = Badge(r0, stv if stv in BADGES else "pending")
        self.badge_vid.grid(row=0, column=6, padx=2)

        ctk.CTkLabel(r0, text="vid", font=("",8), text_color=T3).grid(row=0, column=7)

        self.lbl_tv = ctk.CTkLabel(r0, text="", font=("",9), text_color=T3)
        self.lbl_tv.grid(row=0, column=8, padx=1)

        ctk.CTkButton(r0, text="Tạo video", width=62, height=22, corner_radius=4,
                      fg_color="#E0E7FF", hover_color="#C7D2FE", text_color="#3730A3",
                      font=("",10), command=self._regen_vid).grid(row=0, column=9, padx=(2,0))

        # Row 1: Prompt ảnh
        r1 = ctk.CTkFrame(self, fg_color="transparent")
        r1.grid(row=1, column=1, sticky="ew", padx=(0,6), pady=(2,1))
        r1.grid_columnconfigure(0, weight=1)

        self.pb = ctk.CTkTextbox(r1, height=30, font=("",10), fg_color=EN,
                                  border_color=BD2, border_width=1, corner_radius=4, wrap="word")
        self.pb.grid(row=0, column=0, sticky="ew")
        p = d.get("img_prompt","")
        if p: self.pb.insert("1.0", p)

        # Row 2: Video prompt (editable) + refs
        r2 = ctk.CTkFrame(self, fg_color="transparent")
        r2.grid(row=2, column=1, sticky="ew", padx=(0,6), pady=(1,6))
        r2.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(r2, text="Video:", font=("",9), text_color=T3).grid(row=0, column=0, sticky="w")
        self.vp = ctk.CTkTextbox(r2, height=22, font=("",10), fg_color="#F0F0FF",
                                  border_color="#D0D0E8", border_width=1, corner_radius=4, wrap="word")
        self.vp.grid(row=1, column=0, sticky="ew")
        vp = d.get("video_prompt","") or ""
        if vp: self.vp.insert("1.0", vp)

    def _regen(self):
        if self.on_regen: self.on_regen(self.sid, self.get_prompt())
    def _regen_vid(self):
        if self.on_regen_vid: self.on_regen_vid(self.sid, self.get_video_prompt())
    def _view(self):
        p = self.idir / f"scene_{self.sid:03d}.png"
        if p.exists() and self.on_view: self.on_view(p, f"Scene {self.sid:03d}")
    def _load_img(self):
        p = self.idir / f"scene_{self.sid:03d}.png"
        t = _thumb(p)
        if t: self.img.configure(image=t, text="", fg_color="transparent"); self.img._r = t
        else:
            ph = _ph(); self.img.configure(image=ph, text="", fg_color="#ECECEC"); self.img._r = ph
    def set_status(self, st, ex=None):
        ex = ex or {}
        is_vid = ex.get("phase") == "video"
        if is_vid:
            self.badge_vid.set(st)
            if "elapsed" in ex: self.lbl_tv.configure(text=_ts(ex["elapsed"]))
            if st == "running" and "elapsed" not in ex: self.lbl_tv.configure(text="...")
        else:
            self.badge.set(st)
            c = {"running": RN, "done": OK, "error": ER}.get(st)
            self.configure(border_color=c or BD2, border_width=2 if c else 1)
            if st == "done": self._load_img()
            if "elapsed" in ex: self.lbl_t.configure(text=_ts(ex["elapsed"]))
            if "server" in ex: self.lbl_s.configure(text=f'{ex["server"]}(q={ex.get("queue","?")})')
            if "queue_pos" in ex and ex["queue_pos"] is not None:
                self.lbl_s.configure(text=f'pos={ex["queue_pos"]}')
            if st == "running" and "elapsed" not in ex and "queue_pos" not in ex:
                self.lbl_t.configure(text="...")
    def get_prompt(self):
        return self.pb.get("1.0", "end-1c").strip()
    def get_video_prompt(self):
        return self.vp.get("1.0", "end-1c").strip()

# ── image viewer ─────────────────────────────────────────
class ImageViewer(ctk.CTkToplevel):
    def __init__(self, master, path, title=""):
        super().__init__(master)
        self.title(title or Path(path).name)
        self.geometry("820x620"); self.configure(fg_color="#111")
        self.transient(master); self.grab_set()
        try:
            i = Image.open(str(path))
            r = min(800/i.width, 600/i.height)
            ci = ctk.CTkImage(light_image=i, dark_image=i, size=(int(i.width*r), int(i.height*r)))
            l = ctk.CTkLabel(self, image=ci, text=""); l.pack(expand=True); l._r = ci
        except Exception as e:
            ctk.CTkLabel(self, text=str(e), text_color="#FFF").pack(expand=True)

# ── HOME PAGE ────────────────────────────────────────────
class HomePage(ctk.CTkScrollableFrame):
    def __init__(self, master, app, **k):
        super().__init__(master, fg_color=BG, **k)
        self.app = app
        self.grid_columnconfigure(0, weight=1)
        self._mk_file()
        self._mk_server()
        self._mk_progress()
        self._mk_log()

    def _card(self, row, title):
        c = ctk.CTkFrame(self, fg_color=CD, corner_radius=8, border_width=1, border_color=BD)
        c.grid(row=row, column=0, sticky="ew", padx=10, pady=4)
        c.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(c, text=title, font=("",13,"bold"), text_color=T1,
                     anchor="w").grid(row=0, column=0, padx=10, pady=(8,4), sticky="w", columnspan=4)
        return c

    def _mk_file(self):
        c = self._card(0, "Dự án")
        r = ctk.CTkFrame(c, fg_color="transparent")
        r.grid(row=1, column=0, padx=10, pady=(0,4), sticky="w", columnspan=4)
        for txt, cmd, clr in [
            ("Tải Excel lên", self.app.upload_excel, AC),
            ("Tạo từ SRT", self.app.create_from_srt, AC),
            ("Tải mẫu", self.app.download_template, "#777"),
        ]:
            ctk.CTkButton(r, text=txt, width=100, height=28, fg_color=clr,
                          hover_color=AC2 if clr==AC else "#555", text_color="#FFF",
                          corner_radius=5, font=("",11), command=cmd
                          ).pack(side="left", padx=(0,4))

        self.lbl_file = ctk.CTkLabel(c, text="Chưa chọn file — tải Excel hoặc tạo từ SRT", text_color=T3, font=("",11))
        self.lbl_file.grid(row=2, column=0, padx=10, pady=(0,8), sticky="w", columnspan=4)

    def _mk_server(self):
        c = self._card(1, "Token & Mã dự án")
        tf = ctk.CTkFrame(c, fg_color="transparent")
        tf.grid(row=1, column=0, padx=10, pady=(0,8), sticky="ew", columnspan=4)
        tf.grid_columnconfigure(1, weight=1)
        tf.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(tf, text="Token", font=("",11), text_color=T2
                     ).grid(row=0, column=0, padx=(0,4))
        self.ent_token = ctk.CTkEntry(tf, placeholder_text="ya29.xxx… (không cần chữ Bearer)",
                                       show="•", height=28, corner_radius=4,
                                       font=("Consolas",10), fg_color=EN, border_color=BD)
        self.ent_token.grid(row=0, column=1, sticky="ew", padx=(0,8))

        ctk.CTkLabel(tf, text="ProjID", font=("",11), text_color=T2
                     ).grid(row=0, column=2, padx=(0,4))
        self.ent_projid = ctk.CTkEntry(tf, placeholder_text="project-id",
                                        height=28, corner_radius=4,
                                        font=("Consolas",10), fg_color=EN, border_color=BD)
        self.ent_projid.grid(row=0, column=3, sticky="ew")

        self.lbl_sync = ctk.CTkLabel(c, text="", font=("",9), text_color=T3)
        self.lbl_sync.grid(row=2, column=0, padx=10, pady=(0,6), sticky="w", columnspan=4)

    def load_server_config(self):
        pass

    def update_server_status(self, infos):
        pass

    def _mk_progress(self):
        c = self._card(2, "Tiến độ")
        f = ctk.CTkFrame(c, fg_color="transparent")
        f.grid(row=1, column=0, padx=10, pady=(0,4), sticky="ew", columnspan=4)
        f.grid_columnconfigure(1, weight=1)

        for i, (lbl, color, atp, atl) in enumerate([
            ("Nhân vật", AC, "pb_refs", "lbl_refs"),
            ("Ảnh cảnh", OK, "pb_scenes", "lbl_scenes"),
            ("Video", RN, "pb_vids", "lbl_vids"),
        ]):
            ctk.CTkLabel(f, text=lbl, font=("",11), text_color=T2
                         ).grid(row=i, column=0, padx=(0,6), sticky="e")
            pb = ctk.CTkProgressBar(f, progress_color=color, height=10,
                                     corner_radius=5, fg_color="#EBEBEB")
            pb.grid(row=i, column=1, sticky="ew", pady=2); pb.set(0)
            setattr(self, atp, pb)
            lb = ctk.CTkLabel(f, text="0/0", font=("",10,"bold"), text_color=T2, width=50)
            lb.grid(row=i, column=2, padx=(4,0)); setattr(self, atl, lb)

        self.lbl_cur = ctk.CTkLabel(c, text="", font=("",10), text_color=RN)
        self.lbl_cur.grid(row=2, column=0, padx=10, pady=(0,6), sticky="w", columnspan=4)

    def _mk_log(self):
        c = self._card(3, "Nhật ký")
        self.log_box = ctk.CTkTextbox(c, height=170, font=("Consolas",10),
                                       fg_color="#1A1A1A", text_color="#CCC",
                                       corner_radius=4, wrap="word")
        self.log_box.grid(row=1, column=0, padx=8, pady=(0,8), sticky="ew", columnspan=4)
        self.log_box.configure(state="disabled")

    # ── public ──
    def set_config(self, cfg):
        self.ent_token.delete(0,"end")
        v = cfg.get("flow_bearer_token","")
        if v: self.ent_token.insert(0, v)
        self.ent_projid.delete(0,"end")
        v = cfg.get("flow_project_id","")
        if v: self.ent_projid.insert(0, v)

    def get_token(self):
        t = self.ent_token.get().strip()
        if t.lower().startswith("bearer "): t = t[7:].strip(); self.ent_token.delete(0,"end"); self.ent_token.insert(0, t)
        return t

    def get_project_id(self): return self.ent_projid.get().strip()

    def fill_from_excel(self, wb):
        up = []
        for key, ent in [("flow_bearer_token", self.ent_token), ("flow_project_id", self.ent_projid)]:
            v = wb.get_config_value(key) or ""
            if v:
                if key=="flow_bearer_token" and v.lower().startswith("bearer "): v = v[7:].strip()
                if not ent.get().strip(): ent.delete(0,"end"); ent.insert(0, v); up.append(key.split("_")[-1])
        if up: self.lbl_sync.configure(text=f"Loaded {', '.join(up)} từ Excel", text_color=OK)

    def sync_to_excel(self, wb):
        t = self.get_token()
        if t: wb.set_config_value("flow_bearer_token", t)
        p = self.get_project_id()
        if p: wb.set_config_value("flow_project_id", p)

    def update_progress(self, phase, cur, tot):
        if phase == "refs": self.pb_refs.set(cur/max(tot,1)); self.lbl_refs.configure(text=f"{cur}/{tot}")
        elif phase == "scenes": self.pb_scenes.set(cur/max(tot,1)); self.lbl_scenes.configure(text=f"{cur}/{tot}")
        elif phase == "videos": self.pb_vids.set(cur/max(tot,1)); self.lbl_vids.configure(text=f"{cur}/{tot}")

    def log(self, msg, level="INFO"):
        ts = datetime.now().strftime("%H:%M:%S")
        ic = {"SUCCESS":"✓","ERROR":"✗","WARN":"!"}.get(level," ")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{ts}] {ic} {msg}\n")
        self.log_box.see("end"); self.log_box.configure(state="disabled")


# ── GENERATE PAGE ────────────────────────────────────────
class GeneratePage(ctk.CTkFrame):
    def __init__(self, master, app, **k):
        super().__init__(master, fg_color=BG, **k)
        self.app = app
        self.cc: Dict[str, CharCard] = {}
        self.sc: Dict[int, SceneCard] = {}
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(4, weight=2)

        # chars header
        ch = ctk.CTkFrame(self, fg_color="transparent")
        ch.grid(row=0, column=0, sticky="ew", padx=10, pady=(8,2))
        ch.grid_columnconfigure(0, weight=1)
        self.lbl_c = ctk.CTkLabel(ch, text="Nhân vật (0)", font=("",13,"bold"), text_color=T1)
        self.lbl_c.grid(row=0, column=0, sticky="w")
        ctk.CTkButton(ch, text="Lưu", width=50, height=24, fg_color=AC, hover_color=AC2,
                      text_color="#FFF", corner_radius=4, font=("",10),
                      command=app.save_characters).grid(row=0, column=1)

        self.cs = ctk.CTkScrollableFrame(self, fg_color=BG, height=180)
        self.cs.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0,2))
        self.cs.grid_columnconfigure(0, weight=1)

        # divider
        ctk.CTkFrame(self, fg_color=BD, height=1).grid(row=2, column=0, sticky="ew", padx=10, pady=2)

        # scenes header
        sh = ctk.CTkFrame(self, fg_color="transparent")
        sh.grid(row=3, column=0, sticky="ew", padx=10, pady=(2,2))
        sh.grid_columnconfigure(0, weight=1)
        self.lbl_s = ctk.CTkLabel(sh, text="Cảnh (0)", font=("",13,"bold"), text_color=T1)
        self.lbl_s.grid(row=0, column=0, sticky="w")
        ctk.CTkButton(sh, text="Lưu", width=50, height=24, fg_color=AC, hover_color=AC2,
                      text_color="#FFF", corner_radius=4, font=("",10),
                      command=app.save_scenes).grid(row=0, column=1)

        self.ss = ctk.CTkScrollableFrame(self, fg_color=BG)
        self.ss.grid(row=4, column=0, sticky="nsew", padx=6, pady=(0,6))
        self.ss.grid_columnconfigure(0, weight=1)

    def load_chars(self, data, nv):
        for w in self.cs.winfo_children(): w.destroy()
        self.cc.clear()
        for i, d in enumerate(data):
            c = CharCard(self.cs, d, nv, on_regen=self.app.regen_character, on_view=self.app.view_image)
            c.grid(row=i, column=0, sticky="ew", pady=2, padx=2); self.cc[d["id"]] = c
        self.lbl_c.configure(text=f"Nhân vật ({len(data)})")

    def load_scenes(self, data, idir):
        for w in self.ss.winfo_children(): w.destroy()
        self.sc.clear()
        for i, d in enumerate(data):
            c = SceneCard(self.ss, d, idir, on_regen=self.app.regen_scene,
                         on_regen_vid=self.app.regen_video, on_view=self.app.view_image)
            c.grid(row=i, column=0, sticky="ew", pady=2, padx=2); self.sc[d["scene_id"]] = c
        n = len([s for s in data if s.get("img_prompt")])
        self.lbl_s.configure(text=f"Cảnh ({n})")

    def update_char(self, cid, st, ex=None):
        if cid in self.cc: self.cc[cid].set_status(st, ex)
    def update_scene(self, sid, st, ex=None):
        sid = int(sid) if isinstance(sid,str) else sid
        if sid in self.sc: self.sc[sid].set_status(st, ex)


# ── SETTINGS PAGE ────────────────────────────────────────
class SettingsPage(ctk.CTkScrollableFrame):
    def __init__(self, master, app, **k):
        super().__init__(master, fg_color=BG, **k)
        self.app = app
        self.grid_columnconfigure(0, weight=1)

        # ── Servers ──
        sc = ctk.CTkFrame(self, fg_color=CD, corner_radius=8, border_width=1, border_color=BD)
        sc.grid(row=0, column=0, sticky="ew", padx=10, pady=(10,4))
        sc.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(sc, text="Danh sách Server", font=("",13,"bold"), text_color=T1
                     ).grid(row=0, column=0, padx=10, pady=(8,4), sticky="w", columnspan=6)

        # Add row
        ar = ctk.CTkFrame(sc, fg_color="transparent")
        ar.grid(row=1, column=0, padx=10, pady=(0,4), sticky="ew", columnspan=6)
        ar.grid_columnconfigure(0, weight=1)

        self.ent_url = ctk.CTkEntry(ar, placeholder_text="http://192.168.x.x:5000",
                                     height=28, corner_radius=4, font=("Consolas",10),
                                     fg_color=EN, border_color=BD)
        self.ent_url.grid(row=0, column=0, sticky="ew", padx=(0,3))
        self.ent_nm = ctk.CTkEntry(ar, placeholder_text="Tên", width=70,
                                    height=28, corner_radius=4, font=("",10),
                                    fg_color=EN, border_color=BD)
        self.ent_nm.grid(row=0, column=1, padx=(0,3))
        ctk.CTkButton(ar, text="+", width=28, height=28, corner_radius=4,
                      fg_color=OK, hover_color=OK2, text_color="#FFF", font=("",13,"bold"),
                      command=self._add).grid(row=0, column=2, padx=(0,3))
        ctk.CTkButton(ar, text="Kiểm tra", width=60, height=28, corner_radius=4,
                      fg_color=RN, hover_color="#1565C0", text_color="#FFF", font=("",10),
                      command=app.test_all_servers).grid(row=0, column=3)

        # Server list
        self.sv_frame = ctk.CTkFrame(sc, fg_color="transparent")
        self.sv_frame.grid(row=2, column=0, padx=10, pady=(2,8), sticky="ew", columnspan=6)
        self.sv_frame.grid_columnconfigure(1, weight=1)
        self.sv_rows = []

        # ── Generation ──
        gc = ctk.CTkFrame(self, fg_color=CD, corner_radius=8, border_width=1, border_color=BD)
        gc.grid(row=1, column=0, sticky="ew", padx=10, pady=4)
        gc.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(gc, text="Cấu hình tạo ảnh", font=("",13,"bold"), text_color=T1
                     ).grid(row=0, column=0, padx=10, pady=(8,6), sticky="w", columnspan=3)

        ctk.CTkLabel(gc, text="Số prompt gửi cùng lúc:", font=("",11), text_color=T2
                     ).grid(row=1, column=0, padx=(10,6), sticky="e")
        self.ent_conc = ctk.CTkEntry(gc, width=60, height=28, corner_radius=4,
                                      font=("",11), fg_color=EN, border_color=BD)
        self.ent_conc.grid(row=1, column=1, sticky="w")
        ctk.CTkLabel(gc, text="(1=tuần tự, 3-5=gửi song song nhanh hơn)", font=("",9),
                     text_color=T3).grid(row=1, column=2, padx=6, sticky="w")

        ctk.CTkLabel(gc, text="Số lần thử lại:", font=("",11), text_color=T2
                     ).grid(row=2, column=0, padx=(10,6), sticky="e")
        self.ent_retry = ctk.CTkEntry(gc, width=60, height=28, corner_radius=4,
                                       font=("",11), fg_color=EN, border_color=BD)
        self.ent_retry.grid(row=2, column=1, sticky="w")

        ctk.CTkLabel(gc, text="Tỷ lệ khung hình:", font=("",11), text_color=T2
                     ).grid(row=3, column=0, padx=(10,6), sticky="e")
        self.opt_ar = ctk.CTkOptionMenu(gc, values=["landscape","portrait","square"],
                                         width=120, height=28, corner_radius=4,
                                         fg_color=EN, button_color=BD, text_color=T1,
                                         font=("",11))
        self.opt_ar.grid(row=3, column=1, sticky="w", pady=(0,8))

        # Save button
        ctk.CTkButton(gc, text="Lưu cài đặt", width=120, height=30,
                      fg_color=AC, hover_color=AC2, text_color="#FFF",
                      font=("",11,"bold"), corner_radius=6,
                      command=self._save).grid(row=4, column=0, columnspan=3, padx=10, pady=(4,10))

        self.lbl_saved = ctk.CTkLabel(gc, text="", font=("",9), text_color=OK)
        self.lbl_saved.grid(row=5, column=0, columnspan=3, padx=10, pady=(0,6))

        # ── Update ──
        self._build_update_section()

    def _add(self):
        url = self.ent_url.get().strip()
        if not url: return
        if not url.startswith("http"): url = "http://" + url
        nm = self.ent_nm.get().strip() or f"Sv-{len(self.sv_rows)+1}"
        cfg = self.app.config_data
        if "local_server_list" not in cfg:
            old = cfg.get("local_server_url","")
            cfg["local_server_list"] = []
            if old: cfg["local_server_list"].append({"url":old,"name":"Sv-1","enabled":True})
        cfg["local_server_list"].append({"url":url,"name":nm,"enabled":True})
        cfg["local_server_url"] = url
        self.app._save_config()
        self.ent_url.delete(0,"end"); self.ent_nm.delete(0,"end")
        self._render()
        self.app.test_all_servers()

    def _rm(self, i):
        sl = self.app.config_data.get("local_server_list",[])
        if 0<=i<len(sl): sl.pop(i); self.app._save_config(); self._render()

    def _toggle(self, i):
        sl = self.app.config_data.get("local_server_list",[])
        if 0<=i<len(sl): sl[i]["enabled"]=not sl[i].get("enabled",True); self.app._save_config(); self._render()

    def _render(self):
        for w in self.sv_frame.winfo_children(): w.destroy()
        self.sv_rows.clear()
        sl = self.app.config_data.get("local_server_list",[])
        if not sl:
            u = self.app.config_data.get("local_server_url","")
            if u: sl = [{"url":u,"name":"Sv-1","enabled":True}]
        if not sl:
            ctk.CTkLabel(self.sv_frame, text="— chưa có server —", font=("",10),
                         text_color=T3).grid(row=0, column=0, columnspan=6, pady=2)
            return
        for i, s in enumerate(sl):
            url = s["url"] if isinstance(s,dict) else s
            nm = s.get("name",f"Sv-{i+1}") if isinstance(s,dict) else f"Sv-{i+1}"
            en = s.get("enabled",True) if isinstance(s,dict) else True

            dot = ctk.CTkLabel(self.sv_frame, text="●" if en else "○",
                               text_color=T3, font=("",10))
            dot.grid(row=i, column=0, padx=(0,2))
            ctk.CTkLabel(self.sv_frame, text=nm, font=("",11,"bold"),
                         text_color=T1 if en else T3).grid(row=i, column=1, sticky="w")
            ctk.CTkLabel(self.sv_frame, text=url, font=("Consolas",9),
                         text_color=T3 if en else "#CCC").grid(row=i, column=2, sticky="w", padx=4)
            info = ctk.CTkLabel(self.sv_frame, text="", font=("",9), text_color=T3)
            info.grid(row=i, column=3, padx=4)
            tgl = "ON" if en else "OFF"
            ctk.CTkButton(self.sv_frame, text=tgl, width=30, height=18, corner_radius=3,
                          fg_color=OK if en else "#BBB", hover_color=BD, text_color="#FFF",
                          font=("",8,"bold"),
                          command=lambda x=i: self._toggle(x)).grid(row=i, column=4, padx=1)
            ctk.CTkButton(self.sv_frame, text="✕", width=20, height=18, corner_radius=3,
                          fg_color="#F5D5D5", hover_color=ER, text_color=ER, font=("",9,"bold"),
                          command=lambda x=i: self._rm(x)).grid(row=i, column=5, padx=(1,0))
            self.sv_rows.append({"dot":dot,"info":info,"url":url})

    def update_server_status(self, infos):
        m = {s["url"].rstrip("/"): s for s in infos}
        for r in self.sv_rows:
            si = m.get(r["url"].rstrip("/"))
            if si:
                if si.get("available"):
                    r["dot"].configure(text="●", text_color=OK)
                    r["info"].configure(text=f'q={si.get("queue_size","?")}', text_color=OK)
                else:
                    r["dot"].configure(text="○", text_color=ER)
                    r["info"].configure(text="offline", text_color=ER)

    def load_config(self, cfg):
        self._render()
        self.ent_conc.delete(0,"end"); self.ent_conc.insert(0, str(cfg.get("max_concurrent",1)))
        self.ent_retry.delete(0,"end"); self.ent_retry.insert(0, str(cfg.get("retry_count",3)))
        ar = cfg.get("flow_aspect_ratio","landscape")
        self.opt_ar.set(ar)

    def _save(self):
        cfg = self.app.config_data
        try: cfg["max_concurrent"] = max(1, int(self.ent_conc.get().strip() or "1"))
        except: cfg["max_concurrent"] = 1
        try: cfg["retry_count"] = max(1, int(self.ent_retry.get().strip() or "3"))
        except: cfg["retry_count"] = 3
        cfg["flow_aspect_ratio"] = self.opt_ar.get()
        self.app._save_config()
        self.lbl_saved.configure(text="✓ Đã lưu!")
        self.after(2000, lambda: self.lbl_saved.configure(text=""))

    def _build_update_section(self):
        """Nút cập nhật tool từ GitHub."""
        uc = ctk.CTkFrame(self, fg_color=CD, corner_radius=8, border_width=1, border_color=BD)
        uc.grid(row=2, column=0, sticky="ew", padx=10, pady=4)
        uc.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(uc, text="Cập nhật phần mềm", font=("",13,"bold"), text_color=T1
                     ).grid(row=0, column=0, padx=10, pady=(8,4), sticky="w", columnspan=2)

        self.lbl_update = ctk.CTkLabel(uc, text="", font=("",10), text_color=T3)
        self.lbl_update.grid(row=1, column=0, padx=10, sticky="w")

        ctk.CTkButton(uc, text="Kiểm tra cập nhật", width=130, height=30,
                      fg_color=RN, hover_color="#1565C0", text_color="#FFF",
                      font=("",11), corner_radius=6,
                      command=self._check_update).grid(row=1, column=1, padx=10, pady=(0,8))

    def _check_update(self):
        """Tải code mới từ GitHub zip, cập nhật file .py và modules/."""
        self.lbl_update.configure(text="Đang tải cập nhật...", text_color=RN)

        def _do():
            import requests, zipfile, io, shutil
            ZIP_URL = "https://github.com/nguyenvantuong161978-dotcom/ve3/archive/refs/heads/main.zip"
            try:
                r = requests.get(ZIP_URL, timeout=30)
                if r.status_code != 200:
                    self.after(0, lambda: self.lbl_update.configure(
                        text=f"Lỗi tải: HTTP {r.status_code}", text_color=ER))
                    return

                z = zipfile.ZipFile(io.BytesIO(r.content))
                # Zip có folder gốc "ve3-main/"
                prefix = z.namelist()[0].split("/")[0] + "/"

                updated = []
                # Chỉ cập nhật file code, KHÔNG ghi đè config/settings.yaml, PROJECTS/, templates/
                for name in z.namelist():
                    rel = name[len(prefix):]
                    if not rel or name.endswith("/"):
                        continue
                    # Chỉ cập nhật .py, .bat, requirements.txt
                    if rel.endswith(".py") or rel.endswith(".bat") or rel == "requirements.txt":
                        dest = VE3_DIR / rel
                        dest.parent.mkdir(parents=True, exist_ok=True)
                        with z.open(name) as src, open(dest, "wb") as dst:
                            dst.write(src.read())
                        updated.append(rel)
                    # Cập nhật settings.example.yaml (không đè settings.yaml)
                    elif rel == "config/settings.example.yaml":
                        dest = VE3_DIR / rel
                        dest.parent.mkdir(parents=True, exist_ok=True)
                        with z.open(name) as src, open(dest, "wb") as dst:
                            dst.write(src.read())
                        updated.append(rel)

                msg = f"Đã cập nhật {len(updated)} file. Khởi động lại tool để áp dụng."
                self.after(0, lambda: self.lbl_update.configure(text=msg, text_color=OK))
                self.after(0, lambda: self.app._log(f"Update: {len(updated)} files — {', '.join(updated[:5])}{'...' if len(updated)>5 else ''}", "SUCCESS"))

            except Exception as e:
                self.after(0, lambda: self.lbl_update.configure(
                    text=f"Lỗi: {e}", text_color=ER))

        threading.Thread(target=_do, daemon=True).start()


# ── MAIN APP ─────────────────────────────────────────────
class VE3App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("VE3 Studio"); self.geometry("1200x780"); self.minsize(900,600)
        self.config_data = {}; self.worker = None; self.worker_thread = None
        self.excel_path = None; self.project_dir = None; self.wb = None
        self._t0 = None
        self._load_config(); self._build(); self.after(400, self._boot)

    def _load_config(self):
        try:
            import yaml
            p = VE3_DIR / "config" / "settings.yaml"
            if p.exists():
                with open(p,"r",encoding="utf-8") as f: self.config_data = yaml.safe_load(f) or {}
        except: self.config_data = {}

    def _save_config(self):
        try:
            import yaml
            with open(VE3_DIR/"config"/"settings.yaml","w",encoding="utf-8") as f:
                yaml.dump(self.config_data, f, default_flow_style=False, allow_unicode=True)
        except: pass

    def _build(self):
        self.grid_columnconfigure(1, weight=1); self.grid_rowconfigure(0, weight=1)

        # sidebar
        sb = ctk.CTkFrame(self, width=SW, fg_color=SB, corner_radius=0)
        sb.grid(row=0, column=0, sticky="ns"); sb.grid_rowconfigure(3, weight=1); sb.grid_propagate(False)

        lf = ctk.CTkFrame(sb, fg_color="transparent")
        lf.grid(row=0, column=0, padx=12, pady=(16,20))
        ctk.CTkLabel(lf, text="▶", font=("",18,"bold"), text_color=AC).pack(side="left")
        ctk.CTkLabel(lf, text=" VE3", font=("",16,"bold"), text_color="#FFF").pack(side="left")

        self.nav = {}
        for i, (k, t) in enumerate([("home","Tổng quan"), ("gen","Tạo ảnh")]):
            b = ctk.CTkButton(sb, text=t, width=SW-16, height=34, fg_color="transparent",
                              hover_color=SB3, text_color="#999", anchor="w", corner_radius=6,
                              font=("",12), command=lambda x=k: self.show(x))
            b.grid(row=i+1, column=0, padx=8, pady=1); self.nav[k] = b

        self.btn_go = ctk.CTkButton(sb, text="▶  CHẠY", width=SW-16, height=46,
                                     fg_color="#2E7D32", hover_color="#1B5E20", text_color="#FFFFFF",
                                     font=("",16,"bold"), corner_radius=8, command=self.start_worker)
        self.btn_go.grid(row=4, column=0, padx=8, pady=(6,3))

        self.btn_st = ctk.CTkButton(sb, text="⏹  DỪNG", width=SW-16, height=46,
                                     fg_color="#555", hover_color="#333", text_color="#999",
                                     font=("",16,"bold"), corner_radius=8,
                                     command=self.stop_worker, state="disabled")
        self.btn_st.grid(row=5, column=0, padx=8, pady=(0,3))

        self.lbl_tm = ctk.CTkLabel(sb, text="", font=("",10), text_color="#666")
        self.lbl_tm.grid(row=6, column=0, padx=8)

        ctk.CTkButton(sb, text="Mở thư mục", width=SW-16, height=28,
                      fg_color=SB2, hover_color=SB3, text_color="#888",
                      font=("",10), corner_radius=6,
                      command=self.open_folder).grid(row=7, column=0, padx=8, pady=(2,4))

        # Settings button at bottom
        cfg_btn = ctk.CTkButton(sb, text="⚙ Cài đặt", width=SW-16, height=28,
                                fg_color="transparent", hover_color=SB3, text_color="#777",
                                font=("",11), corner_radius=6, anchor="w",
                                command=lambda: self.show("cfg"))
        cfg_btn.grid(row=8, column=0, padx=8, pady=(0,14))
        self.nav["cfg"] = cfg_btn

        # main
        self.mf = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        self.mf.grid(row=0, column=1, sticky="nsew")
        self.mf.grid_columnconfigure(0, weight=1); self.mf.grid_rowconfigure(0, weight=1)

        self.pages = {
            "home": HomePage(self.mf, self),
            "gen": GeneratePage(self.mf, self),
            "cfg": SettingsPage(self.mf, self),
        }
        self.pages["home"].set_config(self.config_data)
        self.pages["cfg"].load_config(self.config_data)
        self.show("home")

    def show(self, k):
        for p in self.pages.values(): p.grid_forget()
        self.pages[k].grid(row=0, column=0, sticky="nsew")
        for n, b in self.nav.items():
            if n==k: b.configure(fg_color=AC, text_color="#FFF", hover_color=AC2)
            else: b.configure(fg_color="transparent", text_color="#999", hover_color=SB3)

    def _boot(self):
        self.pages["home"].load_server_config()
        self.pages["cfg"]._render()
        self.test_all_servers()

    # ── servers ──
    def _get_svs(self):
        out = []
        sl = self.config_data.get("local_server_list",[])
        if sl:
            for s in sl:
                if isinstance(s,str): out.append({"url":s,"name":s})
                elif isinstance(s,dict) and s.get("enabled",True): out.append(s)
        else:
            u = self.config_data.get("local_server_url","")
            if u: out.append({"url":u,"name":"Sv-1"})
        return out

    def test_all_servers(self):
        svs = self._get_svs()
        if not svs: return
        def _t():
            import requests; res = []
            for s in svs:
                u = s["url"].rstrip("/"); nm = s.get("name",u)
                try:
                    r = requests.get(f"{u}/api/status", timeout=8)
                    if r.status_code==200:
                        d = r.json()
                        res.append({"name":nm,"url":u,"available":True,
                                    "queue_size":d.get("queue_size",0),
                                    "chrome_ready":d.get("chrome_ready",False)})
                    else: res.append({"name":nm,"url":u,"available":False,"queue_size":0,"chrome_ready":False})
                except: res.append({"name":nm,"url":u,"available":False,"queue_size":0,"chrome_ready":False})
            self.after(0, lambda: self.pages["home"].update_server_status(res))
            self.after(0, lambda: self.pages["cfg"].update_server_status(res))
            ok = sum(1 for r in res if r["available"])
            self.after(0, lambda: self._log(f"Servers: {ok}/{len(res)} online", "SUCCESS" if ok else "WARN"))
        threading.Thread(target=_t, daemon=True).start()

    # ── file ──
    def upload_excel(self):
        p = filedialog.askopenfilename(title="Excel", filetypes=[("Excel","*.xlsx"),("All","*.*")])
        if p: self._load_excel(Path(p))

    def create_from_srt(self):
        p = filedialog.askopenfilename(title="SRT", filetypes=[("SRT","*.srt"),("All","*.*")])
        if not p: return
        if not self.config_data.get("deepseek_api_key") and not self.config_data.get("deepseek_api_keys"):
            messagebox.showwarning("Lỗi","Cần deepseek_api_key"); return
        sp = Path(p); code = sp.stem
        pd = VE3_DIR/"PROJECTS"/code; pd.mkdir(parents=True,exist_ok=True)
        dest = pd/sp.name
        if str(sp)!=str(dest): shutil.copy2(str(sp),str(dest))
        self._log(f"SRT → Excel: {sp.name}")
        def _r():
            try:
                from modules.progressive_prompts import ProgressivePromptsGenerator
                ep = pd/f"{code}_prompts.xlsx"
                g = ProgressivePromptsGenerator(srt_path=str(dest),output_path=str(ep),config=self.config_data,
                    log_func=lambda m,l="INFO": self.after(0, lambda: self._log(m,l)))
                g.run_all_steps()
                self.after(0, lambda: self._load_excel(ep))
            except Exception as e: self.after(0, lambda: self._log(f"Error: {e}","ERROR"))
        threading.Thread(target=_r, daemon=True).start()

    def download_template(self):
        src = VE3_DIR/"templates"/"template.xlsx"
        if not src.exists():
            from create_template import create_template; create_template(str(src))
        d = filedialog.asksaveasfilename(title="Save",defaultextension=".xlsx",initialfile="template.xlsx",
                                          filetypes=[("Excel","*.xlsx")])
        if d: shutil.copy2(str(src),d); messagebox.showinfo("OK",f"Saved: {d}")

    def _load_excel(self, path):
        try:
            from modules.excel_manager import PromptWorkbook
            code = path.stem.replace("_prompts","")
            pd = VE3_DIR/"PROJECTS"/code; pd.mkdir(parents=True,exist_ok=True)
            dest = pd/path.name
            if str(path.resolve())!=str(dest.resolve()): shutil.copy2(str(path),str(dest))
            wb = PromptWorkbook(str(dest)); wb.load_or_create()
            self.wb = wb; self.excel_path = dest; self.project_dir = pd
            nv = pd/"nv"; img = pd/"img"; nv.mkdir(exist_ok=True); img.mkdir(exist_ok=True)
            self.pages["home"].fill_from_excel(wb)

            chars = wb.get_characters()
            cd = [c.to_dict() if hasattr(c,'to_dict') else {"id":c.id,"name":c.name,"role":c.role,
                  "english_prompt":c.english_prompt,"vietnamese_prompt":getattr(c,'vietnamese_prompt',''),
                  "status":c.status,"is_child":c.is_child,"media_id":getattr(c,'media_id','')} for c in chars]
            scenes = wb.get_scenes()
            sd = [{"scene_id":s.scene_id,"srt_text":getattr(s,'srt_text',''),"img_prompt":s.img_prompt,
                   "video_prompt":getattr(s,'video_prompt','') or '',
                   "characters_used":getattr(s,'characters_used',''),"location_used":getattr(s,'location_used',''),
                   "reference_files":getattr(s,'reference_files',''),
                   "status_img":getattr(s,'status_img',''),"status_vid":getattr(s,'status_vid','')} for s in scenes]

            self.pages["gen"].load_chars(cd, nv)
            self.pages["gen"].load_scenes(sd, img)
            nc = len(cd); ns = len([s for s in sd if s.get("img_prompt")])
            self.pages["home"].lbl_file.configure(text=f"{path.name}  ·  {nc} chars  ·  {ns} scenes", text_color=T1)
            self._log(f"Loaded {path.name} — {nc} chars, {ns} scenes","SUCCESS")
        except Exception as e:
            self._log(f"Excel error: {e}","ERROR"); messagebox.showerror("Lỗi",str(e))

    # ── save ──
    def save_characters(self):
        if not self.wb: return
        for cid, c in self.pages["gen"].cc.items(): self.wb.update_character(cid, english_prompt=c.get_prompt())
        self.pages["home"].sync_to_excel(self.wb); self.wb.safe_save()
        self._log(f"Saved {len(self.pages['gen'].cc)} characters","SUCCESS")

    def save_scenes(self):
        if not self.wb: return
        for sid, c in self.pages["gen"].sc.items():
            self.wb.update_scene(sid, img_prompt=c.get_prompt(), video_prompt=c.get_video_prompt())
        self.pages["home"].sync_to_excel(self.wb); self.wb.safe_save()
        self._log(f"Saved {len(self.pages['gen'].sc)} scenes","SUCCESS")

    def view_image(self, p, t=""):
        if p and Path(p).exists(): ImageViewer(self, p, t)

    # ── token ──
    def _build_cfg(self):
        t = self.pages["home"].get_token()
        if not t: messagebox.showwarning("Token","Nhập Token!"); return None
        if not t.startswith("ya29."): messagebox.showwarning("Token","Token phải bắt đầu ya29."); return None
        c = dict(self.config_data)
        c["flow_bearer_token"] = t
        c["flow_project_id"] = self.pages["home"].get_project_id()
        # Load concurrent setting
        try: c["max_concurrent"] = max(1, int(self.pages["cfg"].ent_conc.get().strip() or "1"))
        except: c["max_concurrent"] = 1
        return c

    # ── regen ──
    def regen_character(self, cid, prompt):
        if not self.project_dir or not self.wb: messagebox.showwarning("Lỗi","Chưa có project!"); return
        if not prompt: messagebox.showwarning("Lỗi","Prompt trống!"); return
        cfg = self._build_cfg()
        if not cfg: return
        self.wb.update_character(cid, english_prompt=prompt); self.wb.safe_save()
        self.pages["gen"].update_char(cid, "running"); self._log(f"Regen {cid}...")
        ip = self.project_dir/"nv"/f"{cid}.png"
        def _r():
            from ve3_worker import VE3Worker
            try:
                w = VE3Worker(project_dir=str(self.project_dir), config=cfg,
                              log_func=lambda m,l="INFO": self.after(0, lambda: self._log(m,l)))
                t0 = _time.time(); ok, med, si = w._submit_image(prompt, ip)
                el = round(_time.time()-t0,1); ex = {"elapsed":el, **si}
                if ok:
                    self.wb.update_character(cid, status="done", media_id=med or ""); self.wb.safe_save()
                    self.after(0, lambda: self._reload_wb())
                    self.after(0, lambda: self.pages["gen"].update_char(cid, "done", ex))
                    self.after(0, lambda: self._log(f"{cid} done ({el}s)","SUCCESS"))
                else:
                    self.after(0, lambda: self.pages["gen"].update_char(cid, "error", ex))
                    self.after(0, lambda: self._log(f"{cid} failed","ERROR"))
            except Exception as e:
                self.after(0, lambda: self.pages["gen"].update_char(cid, "error"))
                self.after(0, lambda: self._log(f"Error: {e}","ERROR"))
        threading.Thread(target=_r, daemon=True).start()

    def regen_scene(self, sid, prompt):
        if not self.project_dir or not self.wb: messagebox.showwarning("Lỗi","Chưa có project!"); return
        if not prompt: messagebox.showwarning("Lỗi","Prompt trống!"); return
        cfg = self._build_cfg()
        if not cfg: return
        self.wb.update_scene(sid, img_prompt=prompt); self.wb.safe_save()
        self.pages["gen"].update_scene(sid, "running"); self._log(f"Regen scene {sid}...")
        ip = self.project_dir/"img"/f"scene_{sid:03d}.png"
        def _r():
            from ve3_worker import VE3Worker
            try:
                w = VE3Worker(project_dir=str(self.project_dir), config=cfg,
                              log_func=lambda m,l="INFO": self.after(0, lambda: self._log(m,l)))
                mids = w._load_media_ids(self.wb); scenes = self.wb.get_scenes()
                so = next((s for s in scenes if s.scene_id==sid), None)
                refs = w._build_references(so, mids) if so else []
                t0 = _time.time(); ok, med, si = w._submit_image(prompt, ip, refs)
                el = round(_time.time()-t0,1); ex = {"elapsed":el, **si}
                if ok:
                    self.wb.update_scene(sid, status_img="done", media_id=med or ""); self.wb.safe_save()
                    self.after(0, lambda: self._reload_wb())
                    self.after(0, lambda: self.pages["gen"].update_scene(sid, "done", ex))
                    self.after(0, lambda: self._log(f"Scene {sid} done ({el}s)","SUCCESS"))
                else:
                    self.wb.update_scene(sid, status_img="error"); self.wb.safe_save()
                    self.after(0, lambda: self.pages["gen"].update_scene(sid, "error", ex))
                    self.after(0, lambda: self._log(f"Scene {sid} failed","ERROR"))
            except Exception as e:
                self.after(0, lambda: self.pages["gen"].update_scene(sid, "error"))
                self.after(0, lambda: self._log(f"Error: {e}","ERROR"))
        threading.Thread(target=_r, daemon=True).start()

    def _reload_wb(self):
        """Reload workbook từ file để lấy data mới nhất (media_id, status...)."""
        if self.excel_path and self.excel_path.exists():
            from modules.excel_manager import PromptWorkbook
            self.wb = PromptWorkbook(str(self.excel_path))
            self.wb.load_or_create()

    def regen_video(self, sid, video_prompt):
        """Tạo lại video cho 1 scene (Image-to-Video)."""
        if not self.project_dir or not self.wb:
            messagebox.showwarning("Lỗi","Chưa có project!"); return
        if not video_prompt:
            messagebox.showwarning("Lỗi","Video prompt trống!"); return
        cfg = self._build_cfg()
        if not cfg: return

        # Reload workbook để lấy media_id mới nhất
        self._reload_wb()

        # Lấy media_id của ảnh scene
        scenes = self.wb.get_scenes()
        scene_obj = next((s for s in scenes if s.scene_id == sid), None)
        if not scene_obj:
            messagebox.showwarning("Lỗi", f"Không tìm thấy scene {sid}"); return
        media_id = getattr(scene_obj, 'media_id', '') or ''
        if not media_id:
            messagebox.showwarning("Lỗi",
                f"Scene {sid} chưa có media_id.\n"
                "Cần tạo ảnh trước (bấm 'Tạo ảnh') rồi mới tạo video được."); return

        self.wb.update_scene(sid, video_prompt=video_prompt); self.wb.safe_save()
        self.pages["gen"].update_scene(sid, "running", {"phase": "video"})
        self._log(f"Tạo video scene {sid}...")

        vid_path = self.project_dir / "vid" / f"scene_{sid:03d}.mp4"

        def _r():
            from ve3_worker import VE3Worker
            try:
                w = VE3Worker(project_dir=str(self.project_dir), config=cfg,
                              log_func=lambda m,l="INFO": self.after(0, lambda: self._log(m,l)))
                t0 = _time.time()
                ok, si = w._submit_video(video_prompt, vid_path, media_id)
                el = round(_time.time()-t0, 1)
                ex = {"elapsed": el, "phase": "video", **si}
                if ok:
                    self.wb.update_scene(sid, status_vid="done", video_path=str(vid_path))
                    self.wb.safe_save()
                    self.after(0, lambda: self.pages["gen"].update_scene(sid, "done", ex))
                    self.after(0, lambda: self._log(f"Video scene {sid} xong ({el}s)","SUCCESS"))
                else:
                    self.wb.update_scene(sid, status_vid="error"); self.wb.safe_save()
                    self.after(0, lambda: self.pages["gen"].update_scene(sid, "error", ex))
                    self.after(0, lambda: self._log(f"Video scene {sid} lỗi","ERROR"))
            except Exception as e:
                self.after(0, lambda: self.pages["gen"].update_scene(sid, "error", {"phase":"video"}))
                self.after(0, lambda: self._log(f"Lỗi: {e}","ERROR"))
        threading.Thread(target=_r, daemon=True).start()

    # ── full run ──
    def start_worker(self):
        if not self.excel_path: messagebox.showwarning("Lỗi","Upload Excel trước!"); return
        cfg = self._build_cfg()
        if not cfg: return
        if not cfg.get("local_server_url") and not cfg.get("local_server_list"):
            messagebox.showwarning("Lỗi","Thêm server trước!"); return
        self.config_data.update({"flow_bearer_token":cfg["flow_bearer_token"],"flow_project_id":cfg["flow_project_id"]})
        self._save_config(); self.save_characters(); self.save_scenes()

        h = self.pages["home"]
        h.pb_refs.set(0); h.pb_scenes.set(0); h.pb_vids.set(0)
        h.lbl_refs.configure(text="0/0"); h.lbl_scenes.configure(text="0/0"); h.lbl_vids.configure(text="0/0")
        h.lbl_cur.configure(text="")
        h.log_box.configure(state="normal"); h.log_box.delete("1.0","end"); h.log_box.configure(state="disabled")

        # Nút CHẠY mờ, nút DỪNG sáng đỏ
        self.btn_go.configure(state="disabled", fg_color="#555", text_color="#999")
        self.btn_st.configure(state="normal", fg_color="#D32F2F", text_color="#FFFFFF")
        self._t0 = _time.time(); self._tick()

        from ve3_worker import VE3Worker
        self.worker = VE3Worker(
            project_dir=str(self.project_dir), config=cfg,
            log_func=lambda m,l="INFO": self.after(0, lambda: self._log(m,l)),
            progress_func=lambda *a,**kw: self.after(0, lambda: self._prog(*a,**kw)),
            on_item_status=lambda *a,**kw: self.after(0, lambda: self._item(*a,**kw)))
        def _r():
            res = self.worker.run(); self.after(0, lambda: self._done(res))
        self.worker_thread = threading.Thread(target=_r, daemon=True)
        self.worker_thread.start(); self._log("Started!"); self.show("home")

    def stop_worker(self):
        if self.worker: self.worker.stop(); self._log("Stopping…","WARN")

    def _tick(self):
        if self._t0 and self.btn_st.cget("state")!="disabled":
            self.lbl_tm.configure(text=_ts(_time.time()-self._t0)); self.after(1000, self._tick)

    def _prog(self, ph, cur, tot, det=""):
        self.pages["home"].update_progress(ph, cur, tot)
        if det: self.pages["home"].lbl_cur.configure(text=f"→ {det}")

    def _item(self, tp, id, st, path=None, ex=None):
        g = self.pages["gen"]
        if tp=="char": g.update_char(id, st, ex)
        elif tp=="scene": g.update_scene(id, st, ex)

    def _done(self, r):
        # Nút CHẠY sáng xanh lại, nút DỪNG mờ
        self.btn_go.configure(state="normal", fg_color="#2E7D32", text_color="#FFFFFF")
        self.btn_st.configure(state="disabled", fg_color="#555", text_color="#999")
        self.pages["home"].lbl_cur.configure(text="")
        # Reload workbook để GUI có media_id mới nhất
        self._reload_wb()
        tt = f" ({_ts(_time.time()-self._t0)})" if self._t0 else ""
        self._t0 = None
        if r.get("success"): self._log(f"Done: {r['completed']}/{r['total']}{tt}","SUCCESS")
        else:
            e = "; ".join(r.get("errors",[])); self._log(f"End: {r['completed']}/{r['total']}{tt} {e}","ERROR" if e else "WARN")

    def open_folder(self):
        t = self.project_dir if self.project_dir and self.project_dir.exists() else VE3_DIR/"PROJECTS"
        t.mkdir(parents=True,exist_ok=True); os.startfile(str(t))

    def _log(self, m, l="INFO"): self.pages["home"].log(m, l)

def main():
    # Ẩn cửa sổ console trên Windows
    try:
        import ctypes
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    except Exception:
        pass
    app = VE3App(); app.mainloop()

if __name__ == "__main__":
    main()
