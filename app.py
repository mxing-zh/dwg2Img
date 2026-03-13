from __future__ import annotations

import shutil
import sys
import threading
import webbrowser
from multiprocessing import freeze_support
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from converter import ConvertConfig, auto_workers_details, batch_convert

ODA_DOWNLOAD_URL = "https://www.opendesign.com/guestfiles/oda_file_converter"

LAYOUT_MODE_LABELS = {
    "自动（优先布局，回退模型）": "auto",
    "模型空间": "model",
    "指定布局（按下方布局名）": "layout",
}
COLOR_MODE_LABELS = {
    "黑白": "bw",
    "保留原色": "original",
}

NOTE_COLOR = "#777777"
NOTE_FONT = ("Microsoft YaHei UI", 9)


def detect_oda_converter() -> Path | None:
    """Try to auto-detect ODA File Converter executable path."""
    candidates: list[Path] = []

    for cmd in ("ODAFileConverter", "ODAFileConverter.exe"):
        found = shutil.which(cmd)
        if found:
            return Path(found)

    if sys.platform.startswith("win"):
        program_files = [
            Path("C:/Program Files/ODA"),
            Path("C:/Program Files (x86)/ODA"),
        ]
        for base in program_files:
            if base.exists():
                candidates.extend(base.glob("ODAFileConverter*/ODAFileConverter.exe"))
    elif sys.platform == "darwin":
        candidates.extend(
            [
                Path("/Applications/ODAFileConverter.app/Contents/MacOS/ODAFileConverter"),
                Path("/usr/local/bin/ODAFileConverter"),
            ]
        )
    else:
        candidates.extend(
            [
                Path("/usr/bin/ODAFileConverter"),
                Path("/usr/local/bin/ODAFileConverter"),
                Path.home() / ".local/bin/ODAFileConverter",
            ]
        )

    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.withdraw()
        self.title("DWG 批量转图片工具")
        self.geometry("920x700")
        self.minsize(880, 660)

        recommended_workers, worker_hint = auto_workers_details()

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.oda_var = tk.StringVar()
        self.format_var = tk.StringVar(value="png")
        self.dpi_var = tk.StringVar(value="96")
        self.mirror_var = tk.BooleanVar(value=True)
        self.layout_mode_var = tk.StringVar(value="自动（优先布局，回退模型）")
        self.layout_name_var = tk.StringVar()
        self.color_mode_var = tk.StringVar(value="黑白")
        self.max_workers_var = tk.StringVar(value=str(recommended_workers))
        self.cluster_gap_var = tk.StringVar(value="8.0")
        self.view_padding_percent_var = tk.StringVar(value="10")
        self.worker_hint_var = tk.StringVar(value=worker_hint)
        self.oda_workers_var = self.max_workers_var
        self.batch_size_var = tk.StringVar(value="0")

        self._build_ui()
        self._init_oda_path()
        self._center_on_screen()
        self.deiconify()

    def _build_ui(self) -> None:
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        self._path_row(frm, 0, "DWG 根目录：", self.input_var, self._choose_input)
        self._path_row(frm, 1, "输出根目录：", self.output_var, self._choose_output)
        self._path_row(frm, 2, "ODA 转换器：", self.oda_var, self._choose_oda_file)

        mode_frame = ttk.LabelFrame(frm, text="输出模式", padding=8)
        mode_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(6, 6))
        ttk.Checkbutton(mode_frame, text="保留原目录结构", variable=self.mirror_var).grid(row=0, column=0, sticky="w")

        option_frame = ttk.LabelFrame(frm, text="渲染选项", padding=10)
        option_frame.grid(row=4, column=0, columnspan=3, sticky="ew")
        for idx in (1, 4, 7, 10):
            option_frame.columnconfigure(idx, weight=1)

        self._option_field(
            option_frame,
            row=0,
            label_col=0,
            field_col=1,
            label="格式：",
            widget=self._make_combobox(option_frame, self.format_var, ["png", "jpg"], width=7),
        )
        self._option_field(
            option_frame,
            row=0,
            label_col=3,
            field_col=4,
            label="颜色模式：",
            widget=self._make_combobox(option_frame, self.color_mode_var, list(COLOR_MODE_LABELS.keys()), width=7),
        )
        self._option_field(
            option_frame,
            row=0,
            label_col=6,
            field_col=7,
            label="聚合阈值倍率：",
            widget=self._make_entry(option_frame, self.cluster_gap_var, width=10),
        )
        self._option_field(
            option_frame,
            row=0,
            label_col=9,
            field_col=10,
            label="布局模式：",
            widget=self._make_combobox(option_frame, self.layout_mode_var, list(LAYOUT_MODE_LABELS.keys()), width=28),
        )

        self._option_field(
            option_frame,
            row=1,
            label_col=0,
            field_col=1,
            label="DPI：",
            widget=self._make_entry(option_frame, self.dpi_var, width=10),
        )
        self._option_field(
            option_frame,
            row=1,
            label_col=3,
            field_col=4,
            label="渲染并发：",
            widget=self._make_entry(option_frame, self.max_workers_var, width=10),
        )
        padding_widget = ttk.Frame(option_frame)
        self._make_entry(padding_widget, self.view_padding_percent_var, width=10).grid(row=0, column=0, sticky="w")
        ttk.Label(padding_widget, text="%").grid(row=0, column=1, sticky="w", padx=(2, 0))
        self._option_field(
            option_frame,
            row=1,
            label_col=6,
            field_col=7,
            label="四周外扩比例：",
            widget=padding_widget,
        )
        self._option_field(
            option_frame,
            row=1,
            label_col=9,
            field_col=10,
            label="指定布局名：",
            widget=self._make_entry(option_frame, self.layout_name_var, width=31),
        )

        for sep_col in (2, 5, 8):
            ttk.Separator(option_frame, orient="vertical").grid(
                row=0,
                column=sep_col,
                rowspan=2,
                sticky="ns",
                padx=12,
                pady=(6, 0),
            )

        note_group = ttk.LabelFrame(frm, text="选项说明", padding=8)
        note_group.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        note_frame = ttk.Frame(note_group)
        note_frame.pack(fill=tk.X, expand=True)
        note_frame.columnconfigure(1, weight=1)
        self._note_label(note_frame, "1. 聚合阈值倍率：无单位，越大越容易把较远元素并为一组。").grid(row=0, column=0, columnspan=2, sticky="w")
        self._note_label(note_frame, "2. 渲染并发建议：").grid(row=1, column=0, sticky="w")
        self._note_label(note_frame, textvariable=self.worker_hint_var).grid(row=1, column=1, sticky="w")
        self._note_label(note_frame, "3. 四周外扩比例：默认 10，表示按最终有效范围向四周等比例外扩。").grid(row=2, column=0, columnspan=2, sticky="w")
        self._note_label(note_frame, "4. 指定布局名：仅在“指定布局”模式下优先生效。").grid(row=3, column=0, columnspan=2, sticky="w")

        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(frm, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(10, 4))
        self.progress_label = ttk.Label(frm, text="进度: 0%")
        self.progress_label.grid(row=7, column=0, columnspan=3, sticky="w")

        self.start_btn = ttk.Button(frm, text="开始批量转换", command=self._start)
        self.start_btn.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(6, 8))

        self.log_text = tk.Text(frm, height=18)
        self.log_text.grid(row=9, column=0, columnspan=3, sticky="nsew")

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(9, weight=1)

    def _option_field(
        self,
        parent: ttk.Frame,
        *,
        row: int,
        label_col: int,
        field_col: int,
        label: str,
        widget,
        field_span: int = 1,
    ) -> None:
        left_gap = 0 if label_col == 0 else 6
        ttk.Label(parent, text=label, anchor="e").grid(
            row=row,
            column=label_col,
            sticky="e",
            pady=(6, 0),
            padx=(left_gap, 6),
        )
        widget.grid(row=row, column=field_col, columnspan=field_span, sticky="w", pady=(6, 0))

    def _make_entry(self, parent, variable: tk.StringVar, width: int, justify: str = "center") -> ttk.Entry:
        return ttk.Entry(parent, textvariable=variable, width=width, justify=justify)

    def _make_combobox(self, parent, variable: tk.StringVar, values: list[str], width: int) -> ttk.Combobox:
        return ttk.Combobox(parent, textvariable=variable, values=values, width=width, state="readonly", justify="center")

    def _note_label(self, parent, text: str | None = None, textvariable: tk.StringVar | None = None) -> tk.Label:
        return tk.Label(parent, text=text, textvariable=textvariable, fg=NOTE_COLOR, font=NOTE_FONT)

    def _center_on_screen(self) -> None:
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = max((screen_w - width) // 2, 0)
        y = max((screen_h - height) // 2, 0)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _init_oda_path(self) -> None:
        detected = detect_oda_converter()
        if detected:
            self.oda_var.set(str(detected))
            self._append_log(f"已自动检测到 ODA 转换器：{detected}")
        else:
            self._append_log("未检测到 ODA File Converter，请先安装后使用。")
            self._append_log(f"下载地址：{ODA_DOWNLOAD_URL}")

    def _path_row(self, parent: ttk.Frame, row: int, label: str, var: tk.StringVar, handler) -> None:
        ttk.Label(parent, text=label, anchor="e").grid(row=row, column=0, sticky="e", pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(parent, text="选择", command=handler).grid(row=row, column=2, sticky="e", pady=4)

    def _choose_input(self) -> None:
        p = filedialog.askdirectory(title="选择 DWG 根目录")
        if p:
            self.input_var.set(p)

    def _choose_output(self) -> None:
        p = filedialog.askdirectory(title="选择输出根目录")
        if p:
            self.output_var.set(p)

    def _choose_oda_file(self) -> None:
        p = filedialog.askopenfilename(title="选择 ODAFileConverter 可执行文件")
        if p:
            self.oda_var.set(p)

    def _append_log(self, message: str) -> None:
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    def _update_progress(self, message: str) -> None:
        if not message.startswith("进度:"):
            return
        try:
            ratio_text = message.rsplit("(", 1)[1].rstrip(")")
            ratio = float(ratio_text.strip("%"))
            self.progress_var.set(ratio)
            self.progress_label.configure(text=f"进度: {ratio:.1f}%")
        except Exception:
            return

    def _start(self) -> None:
        try:
            dpi = int(self.dpi_var.get())
            max_workers = int(self.max_workers_var.get())
            cluster_gap_scale = float(self.cluster_gap_var.get())
            view_padding_percent = float(self.view_padding_percent_var.get())
        except ValueError:
            messagebox.showerror("参数错误", "DPI/并发/聚合阈值/外扩比例 必须是数值")
            return

        if dpi <= 0:
            messagebox.showerror("参数错误", "DPI 必须大于 0")
            return
        if max_workers <= 0:
            messagebox.showerror("参数错误", "渲染并发必须大于 0")
            return
        if cluster_gap_scale <= 0:
            messagebox.showerror("参数错误", "聚合阈值倍率 必须大于 0")
            return
        if view_padding_percent < 0:
            messagebox.showerror("参数错误", "四周外扩比例 必须大于等于 0")
            return

        cfg = ConvertConfig(
            input_root=Path(self.input_var.get().strip()),
            output_root=Path(self.output_var.get().strip()),
            image_format=self.format_var.get(),
            dpi=dpi,
            mirror_structure=self.mirror_var.get(),
            oda_converter=Path(self.oda_var.get().strip()) if self.oda_var.get().strip() else None,
            layout_mode=LAYOUT_MODE_LABELS[self.layout_mode_var.get()],
            preferred_layout=self.layout_name_var.get().strip() or None,
            color_mode=COLOR_MODE_LABELS[self.color_mode_var.get()],
            max_workers=max_workers,
            cluster_gap_scale=cluster_gap_scale,
            view_padding_ratio=view_padding_percent / 100.0,
        )

        if not cfg.input_root.exists() or not cfg.output_root.exists():
            messagebox.showerror("参数错误", "请先选择存在的输入/输出目录")
            return

        if cfg.oda_converter is None or not cfg.oda_converter.exists():
            go_download = messagebox.askyesno(
                "缺少 ODA File Converter",
                "未检测到可用的 ODA File Converter 可执行文件。\n\n"
                f"请先安装后再进行转换。\n下载地址：{ODA_DOWNLOAD_URL}\n\n"
                "是否现在打开下载页面？",
            )
            if go_download:
                webbrowser.open(ODA_DOWNLOAD_URL)
            return

        self.start_btn.configure(state="disabled")
        self.log_text.delete("1.0", tk.END)
        self.progress_var.set(0.0)
        self.progress_label.configure(text="进度: 0%")

        def worker() -> None:
            try:
                for msg in batch_convert(cfg):
                    self.after(0, self._append_log, msg)
                    self.after(0, self._update_progress, msg)
                self.after(0, lambda: messagebox.showinfo("完成", "转换任务完成"))
            except Exception as exc:
                self.after(0, lambda: messagebox.showerror("转换失败", str(exc)))
            finally:
                self.after(0, lambda: self.start_btn.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    freeze_support()
    App().mainloop()
