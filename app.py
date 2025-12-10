import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict
import threading
import os
from PIL import Image, ImageTk

import document_gen as dg


class DocGenApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("文书生成系统")
        self.geometry("1620x900")
        try:
            self.state("zoomed")
        except Exception:
            pass
        self.minsize(900, 600)
        self.selected_template = None
        self.entries: Dict[str, tk.Entry] = {}
        self.style_var = tk.StringVar(value="formal")
        self.status_var = tk.StringVar(value="就绪")
        self.desc_var = tk.StringVar(value="")
        self.paned = None
        self._pane_mins = (220, 300, 260)
        self.view_mode = tk.StringVar(value="text")
        self.theme_var = tk.StringVar(value="light")
        self.font_var = tk.StringVar(value="手写-马善政")
        self._img = None
        self._history_items = []
        self._build_ui()

    def _build_ui(self):
        paned = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)
        self.paned = paned

        left = ttk.Frame(paned, width=300)
        center = ttk.Frame(paned)
        right = ttk.Frame(paned, width=420)

        paned.add(left, weight=1)
        paned.add(center, weight=3)
        paned.add(right, weight=2)

        paned.bind("<Configure>", lambda e: self._clamp_sashes())
        paned.bind("<ButtonRelease-1>", lambda e: self._clamp_sashes())
        self.after(200, self._clamp_sashes)

        self._build_left(left)
        self._build_center(center)
        self._build_right(right)

        status = ttk.Frame(self)
        status.pack(fill=tk.X)
        ttk.Label(status, textvariable=self.status_var, anchor="w").pack(side=tk.LEFT, padx=10, pady=4)
        self.progress = ttk.Progressbar(status, mode="indeterminate", length=160)
        self.progress.pack(side=tk.RIGHT, padx=10, pady=4)
        self._refresh_templates()
        self._history_refresh()
        self.bind_all("<Control-g>", lambda e: self._generate())
        self.bind_all("<Control-s>", lambda e: self._save_txt())

    def _build_left(self, frame: ttk.Frame):
        head = ttk.Frame(frame)
        head.pack(fill=tk.X, pady=(10, 4))
        ttk.Label(head, text="模板").pack(side=tk.LEFT, padx=(10, 4))
        self.count_var = tk.StringVar(value="0")
        ttk.Label(head, textvariable=self.count_var).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        sbar = ttk.Entry(frame, textvariable=self.search_var)
        sbar.pack(fill=tk.X, padx=10)
        sbar.bind("<KeyRelease>", lambda e: self._refresh_templates())
        self.tpl_var = tk.StringVar()
        self.tpl_cb = ttk.Combobox(frame, textvariable=self.tpl_var, state="readonly")
        self.tpl_cb.pack(fill=tk.X, padx=10, pady=10)
        self.tpl_cb.bind("<<ComboboxSelected>>", lambda e: self._on_template_select())
        ttl = ttk.Label(frame, textvariable=self.desc_var, wraplength=280)
        ttl.pack(fill=tk.X, padx=10, pady=(0, 10))

    def _build_center(self, frame: ttk.Frame):
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill=tk.X)
        left_grp = ttk.Frame(toolbar)
        right_grp = ttk.Frame(toolbar)
        left_grp.pack(side=tk.LEFT, fill=tk.X, expand=True)
        right_grp.pack(side=tk.RIGHT)
        ttk.Label(left_grp, text="风格").pack(side=tk.LEFT, padx=(10, 2))
        style_cb = ttk.Combobox(left_grp, textvariable=self.style_var, values=["formal", "neutral", "strict"], state="readonly", width=10)
        style_cb.pack(side=tk.LEFT)
        style_cb.bind("<<ComboboxSelected>>", lambda e: self._generate())
        ttk.Button(left_grp, text="智能填充", command=self._smart_fill).pack(side=tk.LEFT, padx=10)
        ttk.Label(left_grp, text="字体").pack(side=tk.LEFT, padx=(10, 2))
        font_cb = ttk.Combobox(left_grp, textvariable=self.font_var, values=["手写-马善政", "手写-芝蔓行", "手写-龙藏", "宋体", "楷体", "黑体"], state="readonly", width=12)
        font_cb.pack(side=tk.LEFT)
        ttk.Button(left_grp, text="生成", command=self._generate).pack(side=tk.LEFT)
        ttk.Button(left_grp, text="清空", command=self._clear_form).pack(side=tk.LEFT, padx=8)
        ttk.Button(left_grp, text="复制预览", command=self._copy_preview).pack(side=tk.LEFT)
        ttk.Button(right_grp, text="切换文字/图片", command=self._toggle_view).pack(side=tk.LEFT, padx=10)
        ttk.Button(right_grp, text="切换主题", command=self._toggle_theme).pack(side=tk.LEFT)
        ttk.Separator(frame, orient="horizontal").pack(fill=tk.X, pady=(4, 4))

        self.preview_text = tk.Text(frame, wrap=tk.WORD)
        self.preview_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.preview_image = ttk.Label(frame)
        self.preview_image.pack_forget()

    def _build_right(self, frame: ttk.Frame):
        ttk.Label(frame, text="字段").pack(anchor=tk.W, padx=10, pady=(10, 4))
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        self.form_container = ttk.Frame(canvas)
        self.form_container.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        window = canvas.create_window((0, 0), window=self.form_container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=10)
        scrollbar.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10), pady=10)

        bbar = ttk.Frame(frame)
        bbar.pack(fill=tk.X, padx=10, pady=(0, 10))
        ttk.Button(bbar, text="保存TXT", command=self._save_txt).pack(side=tk.LEFT)
        ttk.Button(bbar, text="保存DOCX", command=self._save_docx).pack(side=tk.LEFT, padx=8)
        ttk.Label(frame, text="历史记录").pack(anchor=tk.W, padx=10)
        self.hist_var = tk.StringVar()
        self.hist_cb = ttk.Combobox(frame, textvariable=self.hist_var, state="readonly")
        self.hist_cb.pack(fill=tk.X, padx=10, pady=(0, 10))
        ttk.Button(frame, text="恢复", command=self._restore_history).pack(anchor=tk.W, padx=10)

    def _refresh_templates(self):
        q = (self.search_var.get() or "").strip().lower()
        items = dg.template_list()
        values = []
        for item in items:
            name = f"{item['name']} ({item['id']})"
            blob = (item["id"] + item["name"] + item["description"]).lower()
            if q and q not in blob:
                continue
            values.append(name)
        self.tpl_cb["values"] = values
        self.count_var.set(str(len(values)))
        if values:
            self.tpl_cb.current(0)
            self._on_template_select()
        else:
            self.tpl_cb.set("")
            self.selected_template = None
            self.desc_var.set("")
            self._clear_form()

    def _on_template_select(self):
        text = self.tpl_cb.get()
        if not text:
            return
        tid = text.split("(")[-1][:-1]
        self.selected_template = tid
        desc = ""
        for it in dg.template_list():
            if it["id"] == tid:
                desc = it["description"]
                break
        self.desc_var.set(desc)
        self._build_form_fields()

    def _clamp_sashes(self):
        if not self.paned:
            return
        total = self.paned.winfo_width()
        if total <= 1:
            return
        ml, mc, mr = self._pane_mins
        if total < (ml + mc + mr + 10):
            # fallback positions when too small
            try:
                self.paned.sashpos(0, ml)
                self.paned.sashpos(1, max(ml + mc, total - mr))
            except Exception:
                pass
            return
        try:
            p0 = self.paned.sashpos(0)
            p1 = self.paned.sashpos(1)
            min0 = ml
            max0 = total - mc - mr
            np0 = min(max(p0, min0), max0)
            min1 = ml + mc
            max1 = total - mr
            np1 = min(max(p1, min1), max1)
            if np0 != p0:
                self.paned.sashpos(0, np0)
            if np1 != p1 or np1 <= np0:
                self.paned.sashpos(1, max(np1, np0 + 50))
        except Exception:
            pass

    def _build_form_fields(self):
        for child in list(self.form_container.children.values()):
            child.destroy()
        self.entries.clear()
        fields = dg.template_fields(self.selected_template or "")
        for i, f in enumerate(fields):
            lab = ttk.Label(self.form_container, text=f)
            ent = ttk.Entry(self.form_container)
            lab.grid(row=i, column=0, sticky="w", padx=4, pady=4)
            ent.grid(row=i, column=1, sticky="ew", padx=4, pady=4)
            self.form_container.grid_columnconfigure(1, weight=1)
            self.entries[f] = ent

    def _collect_data(self) -> Dict[str, str]:
        r: Dict[str, str] = {}
        for k, e in self.entries.items():
            r[k] = e.get().strip()
        return r

    def _update_preview_text(self, text: str):
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert(tk.END, text)
        if self.view_mode.get() == "text":
            self.preview_image.pack_forget()
            self.preview_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def _smart_fill(self):
        if not self.selected_template:
            return
        data = self._collect_data()
        text = dg.generate_document(self.selected_template, data, self.style_var.get())
        self._update_preview_text(text)
        self.status_var.set(f"{self.selected_template} 预览 {len(text)} 字符")

    def _clear_form(self):
        for e in self.entries.values():
            e.delete(0, tk.END)
        self.preview_text.delete("1.0", tk.END)
        self.status_var.set("已清空")

    def _copy_preview(self):
        txt = self.preview_text.get("1.0", tk.END)
        self.clipboard_clear()
        self.clipboard_append(txt)
        messagebox.showinfo("已复制", "预览内容已复制到剪贴板")

    def _save_txt(self):
        if not self.selected_template:
            return
        data = self._collect_data()
        style = self.style_var.get()
        text = dg.generate_document(self.selected_template, data, style)
        path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text", "*.txt")])
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        try:
            dg.record_training(self.selected_template, data)
        except Exception:
            pass
        messagebox.showinfo("已保存", path)

    def _save_docx(self):
        try:
            import docx
        except Exception:
            messagebox.showwarning("未安装", "未检测到python-docx，已为您保存TXT。")
            self._save_txt()
            return
        if not self.selected_template:
            return
        data = self._collect_data()
        style = self.style_var.get()
        text = dg.generate_document(self.selected_template, data, style)
        path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if not path:
            return
        doc = docx.Document()
        for line in text.splitlines():
            doc.add_paragraph(line)
        doc.save(path)
        try:
            dg.record_training(self.selected_template, data)
        except Exception:
            pass
        messagebox.showinfo("已保存", path)

    def _generate(self):
        if not self.selected_template:
            return
        data = self._collect_data()
        style = self.style_var.get()
        self._set_busy(True)
        def run():
            try:
                text = dg.generate_document(self.selected_template, data, style)
                out_dir = os.path.join(os.getcwd(), "handwrite_output")
                os.makedirs(out_dir, exist_ok=True)
                img_path = dg.generate_handwriting_image(text, os.path.join(out_dir, f"{self.selected_template}_{int(__import__('time').time())}.png"), {"font_name": self.font_var.get()})
                dg.add_history(self.selected_template, data, text, img_path)
                self._history_refresh()
                self._update_preview_text(text)
                if self.view_mode.get() == "image":
                    self._show_image(img_path)
                self.status_var.set(f"已生成 {len(text)} 字符，图片：{img_path}")
            except Exception as e:
                messagebox.showerror("错误", str(e))
            finally:
                self.after(0, lambda: self._set_busy(False))
        threading.Thread(target=run, daemon=True).start()

    def _set_busy(self, busy: bool):
        if busy:
            try:
                self.progress.start(12)
            except Exception:
                pass
            self.status_var.set("生成中…")
            try:
                self.config(cursor="watch")
            except Exception:
                pass
        else:
            try:
                self.progress.stop()
            except Exception:
                pass
            try:
                self.config(cursor="")
            except Exception:
                pass

    def _show_image(self, path: str):
        try:
            im = Image.open(path)
            w = self.preview_text.winfo_width() or 1200
            h = self.preview_text.winfo_height() or 600
            im.thumbnail((w - 40, h - 40))
            self._img = ImageTk.PhotoImage(im)
            self.preview_image.configure(image=self._img)
            self.preview_text.pack_forget()
            self.preview_image.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        except Exception as e:
            messagebox.showerror("显示错误", str(e))

    def _toggle_view(self):
        if self.view_mode.get() == "text":
            self.view_mode.set("image")
            items = dg.list_history()
            if items:
                p = items[0].get("image_path")
                if p and os.path.exists(p):
                    self._show_image(p)
        else:
            self.view_mode.set("text")
            self.preview_image.pack_forget()
            self.preview_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def _toggle_theme(self):
        cur = self.theme_var.get()
        nxt = "dark" if cur == "light" else "light"
        self.theme_var.set(nxt)
        if nxt == "dark":
            self.configure(bg="#1e1e1e")
            try:
                self.preview_text.configure(bg="#1e1e1e", fg="#e5e5e5")
            except Exception:
                pass
        else:
            self.configure(bg="#f5f5f5")
            try:
                self.preview_text.configure(bg="#ffffff", fg="#000000")
            except Exception:
                pass

    def _history_refresh(self):
        items = dg.list_history()
        self._history_items = items
        values = []
        for it in items:
            ts = it.get("ts")
            tid = it.get("template_id")
            values.append(f"{tid}-{ts}")
        self.hist_cb["values"] = values
        if values:
            self.hist_cb.current(0)

    def _restore_history(self):
        idx = self.hist_cb.current()
        if idx is None or idx < 0:
            return
        if idx >= len(self._history_items):
            return
        it = self._history_items[idx]
        data = it.get("data", {})
        for k, e in self.entries.items():
            e.delete(0, tk.END)
            v = data.get(k) or ""
            e.insert(0, v)
        text = it.get("text", "")
        self._update_preview_text(text)
        p = it.get("image_path")
        if p and os.path.exists(p) and self.view_mode.get() == "image":
            self._show_image(p)


def main():
    app = DocGenApp()
    app.mainloop()


if __name__ == "__main__":
    main()
