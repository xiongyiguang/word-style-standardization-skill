#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import sys
import threading
import traceback
from pathlib import Path
from tkinter import BooleanVar, StringVar, Tk, filedialog, messagebox, ttk

from normalize_word_style import default_output_paths, normalize_docx


def app_root() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parents[1]


def bundled_template() -> Path:
    template = app_root() / "assets" / "标准样式Word文档.docx"
    if template.exists():
        return template
    return app_root() / "assets" / "standard_template.docx"


class WordStyleStandardizerApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("Word 样式规整工具")
        self.root.geometry("760x420")
        self.root.minsize(720, 380)

        self.input_path = StringVar()
        self.template_path = StringVar(value=str(bundled_template()))
        self.output_path = StringVar()
        self.report_path = StringVar()
        self.use_default_names = BooleanVar(value=True)
        self.status = StringVar(value="选择 Word 文件后点击开始规整。")

        self._build_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        title = ttk.Label(frame, text="Word 样式规整工具", font=("Microsoft YaHei UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 16))

        self._path_row(frame, 1, "源 Word 文件", self.input_path, self.choose_input)
        self._path_row(frame, 2, "标准模板", self.template_path, self.choose_template)

        default_check = ttk.Checkbutton(
            frame,
            text="按源文件名自动生成输出文件和校验报告",
            variable=self.use_default_names,
            command=self.toggle_output_fields,
        )
        default_check.grid(row=3, column=1, sticky="w", pady=(8, 4))

        self.output_entry = self._path_row(frame, 4, "输出文件", self.output_path, self.choose_output)
        self.report_entry = self._path_row(frame, 5, "校验报告", self.report_path, self.choose_report)
        self.toggle_output_fields()

        actions = ttk.Frame(frame)
        actions.grid(row=6, column=1, sticky="w", pady=(18, 10))
        self.run_button = ttk.Button(actions, text="开始规整", command=self.start)
        self.run_button.pack(side="left")
        ttk.Button(actions, text="打开输出目录", command=self.open_output_dir).pack(side="left", padx=(10, 0))

        status_frame = ttk.LabelFrame(frame, text="状态")
        status_frame.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
        status_frame.columnconfigure(0, weight=1)
        self.status_label = ttk.Label(status_frame, textvariable=self.status, wraplength=690, justify="left")
        self.status_label.grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.progress = ttk.Progressbar(status_frame, mode="indeterminate")
        self.progress.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))

    def _path_row(self, parent: ttk.Frame, row: int, label: str, value: StringVar, command) -> ttk.Entry:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=6)
        entry = ttk.Entry(parent, textvariable=value)
        entry.grid(row=row, column=1, sticky="ew", pady=6)
        ttk.Button(parent, text="浏览", command=command).grid(row=row, column=2, sticky="w", padx=(10, 0), pady=6)
        return entry

    def choose_input(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if not path:
            return
        self.input_path.set(path)
        if self.use_default_names.get():
            output, report = default_output_paths(Path(path))
            self.output_path.set(str(Path(path).with_name(output.name)))
            self.report_path.set(str(Path(path).with_name(report.name)))

    def choose_template(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Word 模板文档", "*.docx"), ("所有文件", "*.*")])
        if path:
            self.template_path.set(path)

    def choose_output(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        )
        if path:
            self.output_path.set(path)

    def choose_report(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".md",
            filetypes=[("Markdown 报告", "*.md"), ("文本文件", "*.txt"), ("所有文件", "*.*")],
        )
        if path:
            self.report_path.set(path)

    def toggle_output_fields(self) -> None:
        state = "disabled" if self.use_default_names.get() else "normal"
        self.output_entry.configure(state=state)
        self.report_entry.configure(state=state)
        if self.use_default_names.get() and self.input_path.get():
            output, report = default_output_paths(Path(self.input_path.get()))
            self.output_path.set(str(Path(self.input_path.get()).with_name(output.name)))
            self.report_path.set(str(Path(self.input_path.get()).with_name(report.name)))

    def validate(self) -> tuple[Path, Path, Path | None, Path | None]:
        input_file = Path(self.input_path.get().strip())
        template_file = Path(self.template_path.get().strip())
        if not input_file.exists():
            raise ValueError("请选择存在的源 Word 文件。")
        if input_file.suffix.lower() != ".docx":
            raise ValueError("源文件必须是 .docx 文件。")
        if not template_file.exists():
            raise ValueError("标准模板不存在。")

        output_file = None
        report_file = None
        if self.use_default_names.get():
            output, report = default_output_paths(input_file)
            output_file = input_file.with_name(output.name)
            report_file = input_file.with_name(report.name)
        else:
            output_text = self.output_path.get().strip()
            report_text = self.report_path.get().strip()
            if not output_text:
                raise ValueError("请选择输出文件路径。")
            if not report_text:
                raise ValueError("请选择校验报告路径。")
            output_file = Path(output_text)
            report_file = Path(report_text)
        return input_file, template_file, output_file, report_file

    def start(self) -> None:
        try:
            params = self.validate()
        except Exception as exc:
            messagebox.showerror("无法开始", str(exc))
            return

        self.run_button.configure(state="disabled")
        self.progress.start(12)
        self.status.set("正在规整 Word 样式，请稍候...")
        threading.Thread(target=self.run, args=params, daemon=True).start()

    def run(self, input_file: Path, template_file: Path, output_file: Path | None, report_file: Path | None) -> None:
        try:
            out, report, before, after = normalize_docx(input_file, template_file, output_file, report_file)
        except Exception:
            error = traceback.format_exc()
            self.root.after(0, self.finish_error, error)
            return
        self.root.after(0, self.finish_success, out, report, before, after)

    def finish_success(self, output_file: Path, report_file: Path, before: dict[str, int], after: dict[str, int]) -> None:
        self.progress.stop()
        self.run_button.configure(state="normal")
        self.output_path.set(str(output_file))
        self.report_path.set(str(report_file))
        self.status.set(
            "处理完成。\n"
            f"输出文件：{output_file}\n"
            f"校验报告：{report_file}\n"
            f"段落：{before['paragraphs']} -> {after['paragraphs']}，"
            f"媒体：{before['media_files']} -> {after['media_files']}，"
            f"图片关系：{before['image_relationships']} -> {after['image_relationships']}。"
        )
        messagebox.showinfo("处理完成", f"已生成：\n{output_file}\n\n校验报告：\n{report_file}")

    def finish_error(self, error: str) -> None:
        self.progress.stop()
        self.run_button.configure(state="normal")
        self.status.set("处理失败。请查看错误信息。")
        messagebox.showerror("处理失败", error)

    def open_output_dir(self) -> None:
        target = self.output_path.get().strip() or self.input_path.get().strip()
        if not target:
            messagebox.showinfo("提示", "还没有可打开的输出目录。")
            return
        folder = Path(target)
        if folder.suffix:
            folder = folder.parent
        if folder.exists():
            os.startfile(folder)
        else:
            messagebox.showinfo("提示", "输出目录尚不存在。")


def main() -> None:
    root = Tk()
    WordStyleStandardizerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
