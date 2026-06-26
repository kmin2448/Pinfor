#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
개인정보 검출 스캐너 — GUI
pii_scanner.py 의 검출 엔진을 그대로 사용합니다. 두 파일을 같은 폴더에 두세요.

디자인: 폴더 동기화 프로그램과 동일한 결 — 차분한 세이지 그린 팔레트,
        Malgun Gothic 폰트, 얇은 테두리의 플랫 카드.

필요:
    pip install customtkinter python-docx openpyxl pdfplumber
    (hwp 지원 시 olefile 추가)

실행:
    python pii_scanner_gui.py
"""

import os
import sys
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

# ── 엔진 import (같은 폴더의 pii_scanner.py) ──────────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pii_scanner import (
    DETECTORS, walk_folder, mask, write_csv, write_html, SEV_NAME,
    quarantine_files, delete_files,
)

# ───────────────────────────────────────────────────────────
# 팔레트 — 라이트 그레이 배경 + 흰색 카드(드롭 섀도) · 그린 포인트 유지
# ───────────────────────────────────────────────────────────
BG_BASE     = "#edeef1"   # 페이지 배경 (라이트 쿨 그레이)
PANEL       = "#ffffff"   # 카드 배경 (흰색)
PANEL_IN    = "#f6f7f9"   # 결과 스크롤 영역 (아주 옅은 회색)
INPUT_BG    = "#f4f5f7"   # 입력칸
CARD_IN     = "#ffffff"   # 결과 파일 카드 (흰색)
SHADOW      = "#d6d9df"   # 카드 드롭 섀도 톤

ACCENT      = "#6fa288"   # 세이지 그린 (선택/진행 표시 · 포인트 유지)
ACCENT_HV   = "#5d8f76"
SOFT        = "#dcebdd"   # 연한 그린 버튼 채움
SOFT_HV     = "#cce0ce"
SOFT_TX     = "#3f6b54"   # 연한 그린 버튼 글자
SOFT_BD     = "#bcd6bf"

LIGHT       = "#eef0f3"   # 보조 버튼 (찾아보기 등) · 중립 그레이
LIGHT_HV    = "#e1e4e9"

TEXT        = "#39433c"
MUTED       = "#8a909a"
BORDER      = "#e7e9ee"
INPUT_BD    = "#dfe2e7"

SEV_COLOR   = {3: "#cf6f5f", 2: "#cf9a5a", 1: "#b59a4e"}
DANGER_SOFT = "#f0ddd8"   # 삭제 버튼 (차분한 레드 톤)
DANGER_HV   = "#e7cbc4"
DANGER_TX   = "#a8493b"
DANGER_BD   = "#e0bdb4"

# 폰트 — 한글이 깔끔하게 보이도록 Malgun Gothic 고정 (없으면 시스템 기본)
FONT_FAMILY = "Malgun Gothic"


def F(size=12, weight="normal"):
    """공통 폰트 헬퍼."""
    return ctk.CTkFont(family=FONT_FAMILY, size=size, weight=weight)


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("green")


class PIIScannerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("개인정보 검출 스캐너")
        self.geometry("860x720")
        self.minsize(760, 640)
        self.configure(fg_color=BG_BASE)

        self.results = []
        self.skipped = []
        self.type_vars = {}
        self.reveal_var = ctk.BooleanVar(value=False)
        self.file_vars = {}          # 파일경로 -> BooleanVar (격리/삭제 선택)
        self.scanned_folder = ""     # 마지막으로 검사한 폴더 (격리 기준)

        self._build_header()
        self._build_folder_row()
        self._build_options()
        self._build_action_row()
        self._build_results_area()
        self._build_footer()

    # ── 공통 패널 헬퍼 ──────────────────────────────────────
    def _shadow_card(self, parent, radius=16, fg=PANEL):
        """흰 카드 + 우하단 드롭 섀도를 만든다.
        같은 grid 셀에 섀도 프레임과 카드를 겹쳐, 카드를 좌상단으로
        살짝 올려 섀도가 우·하단으로 비치게 한다.
        반환: (holder=배치용 프레임, card=내용 담을 프레임)"""
        holder = ctk.CTkFrame(parent, fg_color="transparent")
        holder.grid_rowconfigure(0, weight=1)
        holder.grid_columnconfigure(0, weight=1)
        shadow = ctk.CTkFrame(holder, fg_color=SHADOW, corner_radius=radius)
        shadow.grid(row=0, column=0, sticky="nsew", padx=(2, 0), pady=(3, 0))
        card = ctk.CTkFrame(holder, fg_color=fg, corner_radius=radius,
                            border_width=0)
        card.grid(row=0, column=0, sticky="nsew", padx=(0, 3), pady=(0, 4))
        return holder, card

    # ── 헤더 ────────────────────────────────────────────────
    def _build_header(self):
        head = ctk.CTkFrame(self, fg_color="transparent")
        head.pack(fill="x", padx=24, pady=(22, 8))
        ctk.CTkLabel(head, text="🔒  개인정보 검출 스캐너",
                     font=F(23, "bold"), text_color=TEXT).pack(side="left")
        ctk.CTkLabel(head, text="로컬 전용 · 외부 전송 없음",
                     font=F(12), text_color=MUTED).pack(side="right", pady=8)

    # ── 폴더 선택 ──────────────────────────────────────────
    def _build_folder_row(self):
        holder, card = self._shadow_card(self)
        holder.pack(fill="x", padx=24, pady=8)
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=14)

        ctk.CTkLabel(inner, text="검사 폴더", font=F(13, "bold"),
                     text_color=TEXT).pack(anchor="w", pady=(0, 6))
        row = ctk.CTkFrame(inner, fg_color="transparent")
        row.pack(fill="x")
        self.folder_entry = ctk.CTkEntry(
            row, placeholder_text="검사할 폴더를 선택하세요…",
            font=F(12), fg_color=INPUT_BG, border_color=INPUT_BD, border_width=1,
            text_color=TEXT, corner_radius=10, height=40)
        self.folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(row, text="폴더 찾기", width=110, height=40, corner_radius=10,
                      font=F(12), fg_color=LIGHT, hover_color=LIGHT_HV, text_color=TEXT,
                      border_width=1, border_color=BORDER,
                      command=self._browse).pack(side="left")

    def _browse(self):
        path = filedialog.askdirectory(title="검사할 폴더 선택")
        if path:
            self.folder_entry.delete(0, "end")
            self.folder_entry.insert(0, path)

    # ── 옵션 (검출 유형 / 마스킹) ──────────────────────────
    def _build_options(self):
        holder, card = self._shadow_card(self)
        holder.pack(fill="x", padx=24, pady=8)
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=14)

        ctk.CTkLabel(inner, text="검출 유형", font=F(13, "bold"),
                     text_color=TEXT).pack(anchor="w", pady=(0, 8))

        grid = ctk.CTkFrame(inner, fg_color="transparent")
        grid.pack(fill="x")
        for i, (key, det) in enumerate(DETECTORS.items()):
            var = ctk.BooleanVar(value=True)
            self.type_vars[key] = var
            cb = ctk.CTkCheckBox(
                grid, text=det["label"], variable=var,
                font=F(12), text_color=TEXT,
                fg_color=ACCENT, hover_color=ACCENT_HV,
                checkmark_color="#ffffff", border_color=INPUT_BD, corner_radius=5,
                border_width=2)
            cb.grid(row=i // 3, column=i % 3, sticky="w", padx=8, pady=6)

        bottom = ctk.CTkFrame(inner, fg_color="transparent")
        bottom.pack(fill="x", pady=(10, 0))
        ctk.CTkSwitch(
            bottom, text="실제 값 표시 (마스킹 해제 · 주의)",
            variable=self.reveal_var, command=self._reveal_warn,
            font=F(12), text_color=TEXT,
            progress_color=ACCENT, button_color="#ffffff",
            fg_color=INPUT_BD).pack(side="left")

    def _reveal_warn(self):
        if self.reveal_var.get():
            ok = messagebox.askokcancel(
                "주의",
                "마스킹을 해제하면 리포트와 화면에 실제 개인정보가\n"
                "그대로 표시됩니다. 계속하시겠습니까?")
            if not ok:
                self.reveal_var.set(False)

    # ── 실행 버튼 + 진행률 ─────────────────────────────────
    def _build_action_row(self):
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x", padx=24, pady=(4, 8))

        self.scan_btn = ctk.CTkButton(
            row, text="검사 시작", height=46, width=160, corner_radius=12,
            font=F(15, "bold"),
            fg_color=SOFT, hover_color=SOFT_HV, text_color=SOFT_TX,
            border_width=1, border_color=SOFT_BD,
            command=self._start_scan)
        self.scan_btn.pack(side="left")

        self.progress = ctk.CTkProgressBar(
            row, height=10, corner_radius=5, progress_color=ACCENT,
            fg_color="#dfe2e7")
        self.progress.pack(side="left", fill="x", expand=True, padx=16)
        self.progress.set(0)

        self.status = ctk.CTkLabel(row, text="대기 중", text_color=MUTED,
                                   font=F(12), width=160, anchor="e")
        self.status.pack(side="right")

    # ── 결과 영역 ──────────────────────────────────────────
    def _build_results_area(self):
        holder, card = self._shadow_card(self)
        holder.pack(fill="both", expand=True, padx=24, pady=8)
        self.results_frame = ctk.CTkScrollableFrame(
            card, fg_color=PANEL_IN, corner_radius=12,
            scrollbar_button_color=INPUT_BD, scrollbar_button_hover_color=MUTED)
        self.results_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self._placeholder()

    def _placeholder(self):
        for w in self.results_frame.winfo_children():
            w.destroy()
        ctk.CTkLabel(self.results_frame,
                     text="폴더를 선택하고 검사를 시작하세요.",
                     text_color=MUTED, font=F(13)).pack(pady=40)

    # ── 푸터 (대상폴더 + 선택/이동/삭제 + 요약 + 내보내기) ──
    def _build_footer(self):
        # 0행: 이동(격리) 대상 폴더 선택
        dest = ctk.CTkFrame(self, fg_color="transparent")
        dest.pack(fill="x", padx=24, pady=(2, 2))
        ctk.CTkLabel(dest, text="이동 폴더", font=F(12, "bold"),
                     text_color=TEXT).pack(side="left", padx=(0, 8))
        self.dest_entry = ctk.CTkEntry(
            dest, placeholder_text="이동할 별도 폴더를 선택하세요 (비우면 검사 폴더 안에 '_PII_격리_날짜' 자동 생성)",
            font=F(12), fg_color=INPUT_BG, border_color=INPUT_BD, border_width=1,
            text_color=TEXT, corner_radius=10, height=36)
        self.dest_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkButton(dest, text="찾아보기", width=90, height=36, corner_radius=10,
                      font=F(12), fg_color=LIGHT, hover_color=LIGHT_HV, text_color=TEXT,
                      border_width=1, border_color=BORDER,
                      command=self._browse_dest).pack(side="left")

        # 1행: 선택 제어 + 이동/삭제
        act = ctk.CTkFrame(self, fg_color="transparent")
        act.pack(fill="x", padx=24, pady=(4, 2))

        self.select_all_var = ctk.BooleanVar(value=False)
        self.select_all_btn = ctk.CTkButton(
            act, text="전체 선택", width=96, height=34, corner_radius=10,
            font=F(12, "bold"), fg_color=LIGHT, hover_color=LIGHT_HV, text_color=TEXT,
            border_width=1, border_color=BORDER,
            state="disabled", command=self._toggle_all)
        self.select_all_btn.pack(side="left")
        self.select_info = ctk.CTkLabel(act, text="", text_color=MUTED, font=F(12))
        self.select_info.pack(side="left", padx=10)

        self.delete_btn = ctk.CTkButton(
            act, text="선택 삭제", width=120, height=38, corner_radius=10,
            font=F(13, "bold"),
            fg_color=DANGER_SOFT, hover_color=DANGER_HV, text_color=DANGER_TX,
            border_width=1, border_color=DANGER_BD,
            state="disabled", command=self._delete_selected)
        self.delete_btn.pack(side="right", padx=(8, 0))
        self.quar_btn = ctk.CTkButton(
            act, text="선택 파일 이동", width=130, height=38, corner_radius=10,
            font=F(13, "bold"),
            fg_color=SOFT, hover_color=SOFT_HV, text_color=SOFT_TX,
            border_width=1, border_color=SOFT_BD,
            state="disabled", command=self._quarantine_selected)
        self.quar_btn.pack(side="right", padx=(8, 0))

        # 2행: 요약 + 내보내기
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(fill="x", padx=24, pady=(2, 18))
        self.summary = ctk.CTkLabel(row, text="", text_color=TEXT, font=F(13, "bold"))
        self.summary.pack(side="left")

        self.html_btn = ctk.CTkButton(
            row, text="HTML 저장", width=110, height=38, corner_radius=10,
            font=F(13), fg_color=LIGHT, hover_color=LIGHT_HV, text_color=TEXT,
            border_width=1, border_color=BORDER, state="disabled",
            command=lambda: self._export("html"))
        self.html_btn.pack(side="right", padx=(8, 0))
        self.csv_btn = ctk.CTkButton(
            row, text="CSV 저장", width=110, height=38, corner_radius=10,
            font=F(13), fg_color=LIGHT, hover_color=LIGHT_HV, text_color=TEXT,
            border_width=1, border_color=BORDER, state="disabled",
            command=lambda: self._export("csv"))
        self.csv_btn.pack(side="right", padx=(8, 0))

    # ── 스캔 실행 (스레드) ─────────────────────────────────
    def _start_scan(self):
        folder = self.folder_entry.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("오류", "유효한 폴더를 선택하세요.")
            return
        active = [k for k, v in self.type_vars.items() if v.get()]
        if not active:
            messagebox.showerror("오류", "검출 유형을 하나 이상 선택하세요.")
            return

        self.scan_btn.configure(state="disabled", text="검사 중…")
        self.csv_btn.configure(state="disabled")
        self.html_btn.configure(state="disabled")
        self.quar_btn.configure(state="disabled")
        self.delete_btn.configure(state="disabled")
        self.select_all_btn.configure(state="disabled", text="전체 선택")
        self.select_all_var.set(False)
        self.select_info.configure(text="")
        self.file_vars = {}
        self.scanned_folder = folder
        self.progress.set(0)
        for w in self.results_frame.winfo_children():
            w.destroy()

        t = threading.Thread(target=self._scan_worker, args=(folder, active), daemon=True)
        t.start()

    def _scan_worker(self, folder, active):
        def progress_cb(idx, total, path):
            frac = idx / total if total else 1
            name = os.path.basename(path)
            self.after(0, lambda: self._update_progress(frac, idx, total, name))
        try:
            results, skipped = walk_folder(folder, active, progress=progress_cb)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("오류", f"검사 중 오류: {e}"))
            self.after(0, lambda: self.scan_btn.configure(state="normal", text="검사 시작"))
            return
        self.results, self.skipped = results, skipped
        self.after(0, self._render_results)

    def _update_progress(self, frac, idx, total, name):
        self.progress.set(frac)
        self.status.configure(text=f"{idx}/{total}  {name[:18]}")

    # ── 결과 렌더링 ────────────────────────────────────────
    def _render_results(self):
        self.progress.set(1)
        self.scan_btn.configure(state="normal", text="검사 시작")
        for w in self.results_frame.winfo_children():
            w.destroy()

        total = sum(len(f) for _, f in self.results)
        self.status.configure(text="완료")
        self.summary.configure(
            text=f"파일 {len(self.results)}개 · 검출 {total}건 · 건너뜀 {len(self.skipped)}개")

        if not self.results:
            ctk.CTkLabel(self.results_frame, text="✅ 검출된 개인정보가 없습니다.",
                         text_color=ACCENT, font=F(14, "bold")).pack(pady=40)
            return

        self.csv_btn.configure(state="normal")
        self.html_btn.configure(state="normal")
        self.select_all_btn.configure(state="normal", text="전체 선택")
        self.file_vars = {}
        reveal = self.reveal_var.get()

        ordered = sorted(self.results,
                         key=lambda x: -max(f["severity"] for f in x[1]))
        for path, findings in ordered:
            self._render_file_card(path, findings, reveal)
        self._update_select_info()

    def _render_file_card(self, path, findings, reveal):
        holder, card = self._shadow_card(self.results_frame, radius=12, fg=CARD_IN)
        holder.pack(fill="x", padx=6, pady=6)

        head = ctk.CTkFrame(card, fg_color="transparent")
        head.pack(fill="x", padx=14, pady=(12, 6))
        var = ctk.BooleanVar(value=False)
        self.file_vars[path] = var
        ctk.CTkCheckBox(
            head, text="", variable=var, width=24, command=self._update_select_info,
            fg_color=ACCENT, hover_color=ACCENT_HV, checkmark_color="#ffffff",
            border_color=INPUT_BD, corner_radius=5, border_width=2).pack(side="left")
        # 폴더 열기 버튼
        ctk.CTkButton(
            head, text="📂 폴더", width=66, height=28, corner_radius=8,
            font=F(11), fg_color=LIGHT, hover_color=LIGHT_HV, text_color=TEXT,
            border_width=1, border_color=BORDER,
            command=lambda p=path: self._open_folder(p)).pack(side="right")
        # 파일명: 클릭하면 기본 프로그램으로 열림
        name_lbl = ctk.CTkLabel(head, text="📄  " + path, anchor="w",
                                text_color=ACCENT_HV, font=F(13, "bold"),
                                cursor="hand2", wraplength=600, justify="left")
        name_lbl.pack(side="left", fill="x", expand=True)
        name_lbl.bind("<Button-1>", lambda e, p=path: self._open_file(p))
        name_lbl.bind("<Enter>", lambda e, w=name_lbl: w.configure(text_color=ACCENT))
        name_lbl.bind("<Leave>", lambda e, w=name_lbl: w.configure(text_color=ACCENT_HV))

        by_type = {}
        for f in findings:
            by_type.setdefault(f["type"], []).append(f)
        for typ, items in sorted(by_type.items(), key=lambda x: -x[1][0]["severity"]):
            sev = items[0]["severity"]
            line = ctk.CTkFrame(card, fg_color="transparent")
            line.pack(fill="x", padx=14, pady=2)
            ctk.CTkLabel(line, text="●", text_color=SEV_COLOR[sev],
                         font=F(13), width=16).pack(side="left")
            ctk.CTkLabel(line, text=f"{SEV_NAME[sev]} · {typ} · {len(items)}건",
                         text_color=TEXT, font=F(12), anchor="w").pack(side="left")
            sample = items[0]
            shown = sample["value"] if reveal else mask(sample["value"])
            ctk.CTkLabel(line, text=f"  예) {sample['line']}: {shown}",
                         text_color=MUTED, font=F(11), anchor="w").pack(side="left")
        ctk.CTkFrame(card, fg_color="transparent", height=6).pack()

    # ── 내보내기 ──────────────────────────────────────────
    def _export(self, kind):
        if not self.results:
            return
        reveal = self.reveal_var.get()
        if kind == "csv":
            path = filedialog.asksaveasfilename(
                defaultextension=".csv", filetypes=[("CSV", "*.csv")],
                initialfile="개인정보_검출결과.csv")
            if path:
                write_csv(self.results, path, reveal)
                messagebox.showinfo("완료", f"CSV 저장됨:\n{path}")
        else:
            path = filedialog.asksaveasfilename(
                defaultextension=".html", filetypes=[("HTML", "*.html")],
                initialfile="개인정보_검출결과.html")
            if path:
                write_html(self.results, self.skipped, path, reveal)
                messagebox.showinfo("완료", f"HTML 저장됨:\n{path}")

    # ── 파일/폴더 열기 ─────────────────────────────────────
    def _open_path(self, target):
        """OS 기본 연결 프로그램으로 파일/폴더를 연다 (크로스 플랫폼)."""
        try:
            if sys.platform.startswith("win"):
                os.startfile(target)                      # Windows
            elif sys.platform == "darwin":
                subprocess.Popen(["open", target])        # macOS
            else:
                subprocess.Popen(["xdg-open", target])    # Linux
            return True
        except Exception as e:
            messagebox.showerror("열기 실패", f"열 수 없습니다:\n{target}\n\n{e}")
            return False

    def _open_file(self, path):
        if not os.path.exists(path):
            messagebox.showwarning(
                "파일 없음",
                f"파일을 찾을 수 없습니다.\n이미 이동/삭제되었을 수 있습니다.\n\n{path}")
            return
        self._open_path(path)

    def _open_folder(self, path):
        """파일이 들어 있는 폴더를 연다 (가능하면 해당 파일을 선택)."""
        folder = os.path.dirname(path) or "."
        if not os.path.isdir(folder):
            messagebox.showwarning("폴더 없음", f"폴더를 찾을 수 없습니다.\n\n{folder}")
            return
        # Windows 탐색기는 파일을 선택한 채로 열 수 있음
        if sys.platform.startswith("win") and os.path.exists(path):
            try:
                subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
                return
            except Exception:
                pass
        self._open_path(folder)

    # ── 대상 폴더 선택 ─────────────────────────────────────
    def _browse_dest(self):
        path = filedialog.askdirectory(title="이동할 별도 폴더 선택")
        if path:
            self.dest_entry.delete(0, "end")
            self.dest_entry.insert(0, path)

    # ── 선택 / 이동 / 삭제 ─────────────────────────────────
    def _toggle_all(self):
        """전체 선택 ↔ 선택 해제 토글 (버튼)."""
        if not self.file_vars:
            return
        select = not self.select_all_var.get()   # 현재 상태의 반대로
        self.select_all_var.set(select)
        for var in self.file_vars.values():
            var.set(select)
        self.select_all_btn.configure(text="선택 해제" if select else "전체 선택")
        self._update_select_info()

    def _selected_paths(self):
        return [p for p, v in self.file_vars.items() if v.get()]

    def _update_select_info(self):
        n = len(self._selected_paths())
        total = len(self.file_vars)
        self.select_info.configure(text=f"{n}/{total}개 선택됨" if n else "")
        state = "normal" if n else "disabled"
        self.quar_btn.configure(state=state)
        self.delete_btn.configure(state=state)
        # 개별 체크 변화에 따라 전체선택 버튼 라벨 동기화
        if total and n == total:
            self.select_all_var.set(True)
            self.select_all_btn.configure(text="선택 해제")
        else:
            self.select_all_var.set(False)
            self.select_all_btn.configure(text="전체 선택")

    def _remove_from_results(self, paths):
        """결과 목록·선택 상태에서 처리된 파일 제거 후 다시 렌더링."""
        gone = set(paths)
        self.results = [(p, f) for p, f in self.results if p not in gone]
        self.select_all_var.set(False)
        self._render_results()

    def _quarantine_selected(self):
        paths = self._selected_paths()
        if not paths:
            return

        dest = self.dest_entry.get().strip()
        if dest:
            # 지정한 별도 폴더로 이동 (없으면 생성)
            try:
                os.makedirs(dest, exist_ok=True)
            except Exception as e:
                messagebox.showerror(
                    "오류", f"이동 폴더를 만들 수 없습니다:\n{dest}\n\n{e}")
                return
            if os.path.abspath(dest) == os.path.abspath(self.scanned_folder):
                messagebox.showerror(
                    "오류", "이동 폴더가 검사 폴더와 같습니다.\n다른 폴더를 선택하세요.")
                return
            detail = f"지정한 폴더로 이동됩니다:\n{dest}"
        else:
            detail = ("검사 폴더 안에 '_PII_격리_(날짜시각)' 폴더를\n"
                      "자동으로 만들어 그곳으로 이동됩니다.")

        if not messagebox.askokcancel(
            "이동 확인",
            f"선택한 {len(paths)}개 파일을 이동합니다.\n\n{detail}\n\n"
            f"· 원본 위치에서는 사라집니다(하위 폴더 구조 유지).\n"
            f"· 복구가 필요하면 이동된 폴더에서 되돌릴 수 있습니다.\n\n"
            f"계속하시겠습니까?"):
            return

        moved, errors, qroot = quarantine_files(
            paths, self.scanned_folder, quarantine_root=(dest or None))
        self._remove_from_results([src for src, _ in moved])
        msg = f"{len(moved)}개 파일을 이동했습니다.\n\n이동 위치:\n{qroot}"
        if errors:
            msg += f"\n\n실패 {len(errors)}개:\n" + "\n".join(
                f"· {os.path.basename(p)}: {e}" for p, e in errors[:5])
        messagebox.showinfo("이동 완료", msg)

    def _delete_selected(self):
        paths = self._selected_paths()
        if not paths:
            return
        # 1차 확인
        if not messagebox.askokcancel(
            "삭제 확인",
            f"선택한 {len(paths)}개 파일을 삭제합니다.\n"
            f"send2trash 가 설치돼 있으면 휴지통으로(복구 가능),\n"
            f"없으면 영구 삭제됩니다.\n\n계속하시겠습니까?"):
            return
        # 2차 확인 (되돌릴 수 없는 작업이므로 한 번 더)
        if not messagebox.askyesno(
            "최종 확인",
            "정말로 삭제하시겠습니까?\n이 작업은 되돌리기 어려울 수 있습니다.",
            icon="warning"):
            return
        deleted, errors, to_trash = delete_files(paths, use_trash=True)
        self._remove_from_results(deleted)
        where = "휴지통으로 이동" if to_trash else "영구 삭제"
        msg = f"{len(deleted)}개 파일을 {where}했습니다."
        if errors:
            msg += f"\n\n실패 {len(errors)}개:\n" + "\n".join(
                f"· {os.path.basename(p)}: {e}" for p, e in errors[:5])
        messagebox.showinfo("삭제 완료", msg)


if __name__ == "__main__":
    app = PIIScannerApp()
    app.mainloop()
