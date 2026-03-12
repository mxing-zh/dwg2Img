from __future__ import annotations

import threading
import shutil
import sys
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from converter import ConvertConfig, batch_convert

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
        self.single_dir_var = tk.StringVar()
        self.oda_var = tk.StringVar()
        self.format_var = tk.StringVar(value="png")
        self.dpi_var = tk.StringVar(value="96")
        self.mode_var = tk.StringVar(value="mirror")
        self.layout_mode_var = tk.StringVar(value="auto")
        self.layout_name_var = tk.StringVar()

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
        mode_frame.columnconfigure(1, weight=1)

        ttk.Radiobutton(
            mode_frame,
            text="保持源目录结构（推荐）",
            variable=self.mode_var,
            value="mirror",
            command=self._on_mode_change,
        ).grid(row=0, column=0, columnspan=3, sticky="w")

        ttk.Radiobutton(
            mode_frame,
            text="全部输出到单目录",
            variable=self.mode_var,
            value="single",
            command=self._on_mode_change,
        ).grid(row=1, column=0, sticky="w")

        self.single_entry = ttk.Entry(mode_frame, textvariable=self.single_dir_var)
        self.single_entry.grid(row=1, column=1, sticky="ew", padx=6)
        self.single_btn = ttk.Button(mode_frame, text="选择", command=self._choose_single)
        self.single_btn.grid(row=1, column=2, sticky="e")

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

        self.start_btn = ttk.Button(frm, text="开始批量转换", command=self._start)
        self.start_btn.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(10, 8))

        self.log_text = tk.Text(frm, height=20)
        self.log_text.grid(row=6, column=0, columnspan=3, sticky="nsew")

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(6, weight=1)
        self._on_mode_change()

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

    def _choose_single(self) -> None:
        p = filedialog.askdirectory(title="选择单目录输出")
        if p:
            self.single_dir_var.set(p)

    def _choose_oda_file(self) -> None:
        p = filedialog.askopenfilename(title="选择 ODAFileConverter 可执行文件")
        if p:
            self.oda_var.set(p)

    def _on_mode_change(self) -> None:
        enabled = self.mode_var.get() == "single"
        state = "normal" if enabled else "disabled"
        self.single_entry.configure(state=state)
        self.single_btn.configure(state=state)

    def _append_log(self, message: str) -> None:
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    def _start(self) -> None:
        try:
            dpi = int(self.dpi_var.get())
        except ValueError:
            messagebox.showerror("参数错误", "DPI 必须是整数")
            return

        if dpi <= 0:
            messagebox.showerror("参数错误", "DPI 必须大于 0")
            return

        mode = self.mode_var.get()
        single_dir = Path(self.single_dir_var.get()) if mode == "single" and self.single_dir_var.get().strip() else None

        cfg = ConvertConfig(
            input_root=Path(self.input_var.get().strip()),
            output_root=Path(self.output_var.get().strip()),
            image_format=self.format_var.get(),
            dpi=dpi,
            single_output_dir=single_dir,
            mirror_structure=(mode == "mirror"),
            oda_converter=Path(self.oda_var.get().strip()) if self.oda_var.get().strip() else None,
            layout_mode=self.layout_mode_var.get(),
            preferred_layout=self.layout_name_var.get().strip() or None,
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

        if cfg.single_output_dir and not cfg.single_output_dir.exists():
            messagebox.showerror("参数错误", "单目录输出路径不存在")
            return

        self.start_btn.configure(state="disabled")
        self.log_text.delete("1.0", tk.END)

        def worker() -> None:
            try:
                for msg in batch_convert(cfg):
                    self.after(0, self._append_log, msg)
                self.after(0, lambda: messagebox.showinfo("完成", "转换任务完成"))
            except Exception as exc:
                self.after(0, lambda: messagebox.showerror("转换失败", str(exc)))
            finally:
                self.after(0, lambda: self.start_btn.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    App().mainloop()
