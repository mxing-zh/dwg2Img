from __future__ import annotations

import threading
import shutil
import sys
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from converter import ConvertConfig, auto_workers, batch_convert

ODA_DOWNLOAD_URL = "https://www.opendesign.com/guestfiles/oda_file_converter"


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
        self.title("DWG 批量转图片工具")
        self.geometry("920x700")

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.oda_var = tk.StringVar()
        self.format_var = tk.StringVar(value="png")
        self.dpi_var = tk.StringVar(value="96")
        self.mirror_var = tk.BooleanVar(value=True)
        self.layout_mode_var = tk.StringVar(value="auto")
        self.layout_name_var = tk.StringVar()
        self.color_mode_var = tk.StringVar(value="bw")
        self.max_workers_var = tk.StringVar(value="0")
        self.oda_workers_var = tk.StringVar(value="0")
        self.batch_size_var = tk.StringVar(value="200")

        self._build_ui()
        self._init_oda_path()

    def _build_ui(self) -> None:
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        self._path_row(frm, 0, "DWG 根目录", self.input_var, self._choose_input)
        self._path_row(frm, 1, "输出根目录", self.output_var, self._choose_output)
        self._path_row(frm, 2, "ODA转换器", self.oda_var, self._choose_oda_file)

        mode_frame = ttk.LabelFrame(frm, text="输出模式", padding=10)
        mode_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 8))
        ttk.Checkbutton(
            mode_frame,
            text="保留原目录结构",
            variable=self.mirror_var,
        ).grid(row=0, column=0, sticky="w")

        option_frame = ttk.LabelFrame(frm, text="渲染选项", padding=10)
        option_frame.grid(row=4, column=0, columnspan=3, sticky="ew")
        option_frame.columnconfigure(1, weight=1)
        option_frame.columnconfigure(3, weight=1)

        ttk.Label(option_frame, text="格式").grid(row=0, column=0, sticky="w")
        ttk.Combobox(option_frame, textvariable=self.format_var, values=["png", "jpg"], width=8, state="readonly").grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(option_frame, text="DPI(默认96)").grid(row=0, column=2, sticky="e")
        ttk.Entry(option_frame, textvariable=self.dpi_var, width=10).grid(row=0, column=3, sticky="w", padx=6)

        ttk.Label(option_frame, text="布局模式").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Combobox(
            option_frame,
            textvariable=self.layout_mode_var,
            values=["auto", "model", "layout"],
            width=12,
            state="readonly",
        ).grid(row=1, column=1, sticky="w", padx=6, pady=(8, 0))

        ttk.Label(option_frame, text="指定布局名(可选)").grid(row=1, column=2, sticky="e", pady=(8, 0))
        ttk.Entry(option_frame, textvariable=self.layout_name_var, width=18).grid(row=1, column=3, sticky="w", padx=6, pady=(8, 0))

        ttk.Label(option_frame, text="颜色模式").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Combobox(
            option_frame,
            textvariable=self.color_mode_var,
            values=["bw", "original"],
            width=12,
            state="readonly",
        ).grid(row=2, column=1, sticky="w", padx=6, pady=(8, 0))
        ttk.Label(option_frame, text="bw=黑白, original=保留原色").grid(row=2, column=2, columnspan=2, sticky="w", pady=(8, 0))

        ttk.Label(option_frame, text=f"渲染并发(0=自动, 建议{auto_workers()})").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(option_frame, textvariable=self.max_workers_var, width=10).grid(row=3, column=1, sticky="w", padx=6, pady=(8, 0))
        ttk.Label(option_frame, text="ODA并发(0=自动)").grid(row=3, column=2, sticky="e", pady=(8, 0))
        ttk.Entry(option_frame, textvariable=self.oda_workers_var, width=10).grid(row=3, column=3, sticky="w", padx=6, pady=(8, 0))

        ttk.Label(option_frame, text="ODA批大小").grid(row=4, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(option_frame, textvariable=self.batch_size_var, width=10).grid(row=4, column=1, sticky="w", padx=6, pady=(8, 0))

        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ttk.Progressbar(frm, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(10, 4))
        self.progress_label = ttk.Label(frm, text="进度: 0%")
        self.progress_label.grid(row=6, column=0, columnspan=3, sticky="w")

        self.start_btn = ttk.Button(frm, text="开始批量转换", command=self._start)
        self.start_btn.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(6, 8))

        self.log_text = tk.Text(frm, height=20)
        self.log_text.grid(row=8, column=0, columnspan=3, sticky="nsew")

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(8, weight=1)

    def _init_oda_path(self) -> None:
        detected = detect_oda_converter()
        if detected:
            self.oda_var.set(str(detected))
            self._append_log(f"已自动检测到 ODA 转换器：{detected}")
        else:
            self._append_log("未检测到 ODA File Converter，请先安装后使用。")
            self._append_log(f"下载地址：{ODA_DOWNLOAD_URL}")

    def _path_row(self, parent: ttk.Frame, row: int, label: str, var: tk.StringVar, handler) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(parent, text="选择", command=handler).grid(row=row, column=2, sticky="e", pady=4)

    def _choose_input(self) -> None:
        p = filedialog.askdirectory(title="选择DWG根目录")
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
            oda_workers = int(self.oda_workers_var.get())
            batch_size = int(self.batch_size_var.get())
        except ValueError:
            messagebox.showerror("参数错误", "DPI/并发/批大小 必须是整数")
            return

        if dpi <= 0:
            messagebox.showerror("参数错误", "DPI 必须大于 0")
            return
        if max_workers < 0 or oda_workers < 0:
            messagebox.showerror("参数错误", "并发数不能小于 0")
            return
        if batch_size <= 0:
            messagebox.showerror("参数错误", "批大小必须大于 0")
            return

        cfg = ConvertConfig(
            input_root=Path(self.input_var.get().strip()),
            output_root=Path(self.output_var.get().strip()),
            image_format=self.format_var.get(),
            dpi=dpi,
            mirror_structure=self.mirror_var.get(),
            oda_converter=Path(self.oda_var.get().strip()) if self.oda_var.get().strip() else None,
            layout_mode=self.layout_mode_var.get(),
            preferred_layout=self.layout_name_var.get().strip() or None,
            color_mode=self.color_mode_var.get(),
            max_workers=max_workers,
            oda_workers=oda_workers,
            oda_batch_size=batch_size,
        )

        if not cfg.input_root.exists() or not cfg.output_root.exists():
            messagebox.showerror("参数错误", "请先选择存在的输入/输出目录")
            return

        same_root = cfg.input_root.resolve() == cfg.output_root.resolve()
        if same_root:
            if cfg.mirror_structure:
                prompt = (
                    "检测到输出目录与源目录相同，且已勾选“保留原目录结构”。\n\n"
                    "这会在各个源文件所在目录直接生成同名图片（png/jpg），\n"
                    "源文件与输出文件会混放在一起。\n\n"
                    "建议改用全新的输出目录，以免后续管理混乱。\n"
                    "是否仍继续转换？"
                )
            else:
                prompt = (
                    "检测到输出目录与源目录相同，且未勾选“保留原目录结构”。\n\n"
                    "这会把所有结果输出到源根目录，可能产生重名覆盖，\n"
                    "源文件与输出文件会混放在一起。\n\n"
                    "建议改用全新的输出目录。\n"
                    "是否仍继续转换？"
                )

            if not messagebox.askyesno("输出目录建议", prompt):
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
    App().mainloop()
