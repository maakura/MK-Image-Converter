import sys
import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
from PIL import Image, ImageTk, ImageDraw, ImageFont
from tkinterdnd2 import DND_FILES, TkinterDnD

# 文字のぼやけ対策
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except: pass

CONFIG_FILE = "preset_config.json"

# 多言語辞書
MESSAGES = {
    "JP": {
        "title": "MK Image Converter Expert",
        "save_preset": "設定を保存",
        "config": "設定",
        "drop_area": " 1. 画像をドラッグ＆ドロップ ",
        "select_file": "ファイルを選択",
        "list_label": "変換対象リスト (クリックで拡大プレビュー)",
        "edit_conf": " 2. 変換・加工設定 ",
        "size": "サイズ(px):",
        "aspect": "比率維持",
        "wm_text": "透かし(文字):",
        "wm_size": "サイズ:",
        "wm_pos": "位置:",
        "wm_color": "色",
        "wm_preview": "透かし確認",
        "save_name": "保存名:",
        "run": "一括変換スタート",
        "status_wait": "待機中...",
        "status_done": "完了！",
        "msg_done": "すべての変換が完了しました。",
        "lang_select": "言語選択 (Language)",
    },
    "EN": {
        "title": "MK Image Converter Expert",
        "save_preset": "Save Preset",
        "config": "Settings",
        "drop_area": " 1. Drag & Drop Images Here ",
        "select_file": "Select Files",
        "list_label": "Image List (Click to preview)",
        "edit_conf": " 2. Process Settings ",
        "size": "Size(px):",
        "aspect": "Keep Aspect",
        "wm_text": "Watermark:",
        "wm_size": "Size:",
        "wm_pos": "Pos:",
        "wm_color": "Color",
        "wm_preview": "Preview WM",
        "save_name": "Filename:",
        "run": "Start Conversion",
        "status_wait": "Waiting...",
        "status_done": "Done!",
        "msg_done": "All conversions completed.",
        "lang_select": "Select Language",
    }
}

class ImageConverter:
    def __init__(self, root):
        self.root = root
        self.image_items = []
        self.wm_color = "#ffffff"
        self.preset = self.load_preset()
        self.lang = self.preset.get("lang", "JP")

        self.create_widgets()
        self.apply_preset()
        
        # 起動時にREADMEとISSを生成
        try:
            self.generate_readme()
            self.generate_installer_script()
        except Exception as e:
            print(f"Generation error: {e}")

    def load_preset(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except: pass
        return {"lang": "JP", "width": "1280", "height": "720", "keep_aspect": True}

    def save_preset(self):
        current_config = {
            "lang": self.lang,
            "width": self.ent_width.get(),
            "height": self.ent_height.get(),
            "keep_aspect": self.keep_aspect.get(),
            "filename": self.ent_filename.get(),
            "ext_index": self.combo_ext.current(),
            "wm_text": self.wm_text.get(),
            "wm_size": self.wm_size.get(),
            "wm_pos": self.wm_pos.get(),
            "wm_color": self.wm_color
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(current_config, f, indent=4)
        messagebox.showinfo("OK", "Preset Saved / 設定を保存しました")

    def apply_preset(self):
        p = self.preset
        self.update_ui_text()
        self.ent_width.delete(0, tk.END); self.ent_width.insert(0, p.get("width", "1280"))
        self.ent_height.delete(0, tk.END); self.ent_height.insert(0, p.get("height", "720"))
        self.keep_aspect.set(p.get("keep_aspect", True))
        self.ent_filename.delete(0, tk.END); self.ent_filename.insert(0, p.get("filename", "{name}_{n}"))
        self.combo_ext.current(p.get("ext_index", 0))
        self.wm_text.delete(0, tk.END); self.wm_text.insert(0, p.get("wm_text", ""))
        self.wm_size.delete(0, tk.END); self.wm_size.insert(0, p.get("wm_size", "40"))
        self.wm_pos.set(p.get("wm_pos", "右下"))
        self.wm_color = p.get("wm_color", "#ffffff")

    def update_ui_text(self):
        m = MESSAGES[self.lang]
        self.root.title(m["title"])
        self.lbl_title.config(text=m["title"])
        self.btn_save_pre.config(text=m["save_preset"])
        self.btn_config.config(text=m["config"])
        self.drop_frame.config(text=m["drop_area"])
        self.btn_select.config(text=m["select_file"])
        self.lbl_list.config(text=m["list_label"])
        self.conf_frame.config(text=m["edit_conf"])
        self.lbl_size.config(text=m["size"])
        self.chk_aspect.config(text=m["aspect"])
        self.lbl_wm_t.config(text=m["wm_text"])
        self.lbl_wm_s.config(text=m["wm_size"])
        self.lbl_wm_p.config(text=m["wm_pos"])
        self.btn_wm_c.config(text=m["wm_color"])
        self.btn_wm_pre.config(text=m["wm_preview"])
        self.lbl_save_n.config(text=m["save_name"])
        self.btn_run.config(text=m["run"])
        self.lbl_status.config(text=m["status_wait"])

    def open_config_window(self):
        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.geometry("300x150")
        tk.Label(win, text=MESSAGES[self.lang]["lang_select"]).pack(pady=10)
        lang_var = tk.StringVar(value=self.lang)
        combo = ttk.Combobox(win, textvariable=lang_var, values=["JP", "EN"], state="readonly")
        combo.pack(pady=5)
        def change_lang():
            self.lang = lang_var.get()
            self.update_ui_text()
            win.destroy()
        tk.Button(win, text="Apply / 適用", command=change_lang).pack(pady=10)

    def create_widgets(self):
        header = tk.Frame(self.root, bg="#f5f5f5")
        header.pack(fill="x", padx=20, pady=10)
        self.lbl_title = tk.Label(header, font=("Arial", 16, "bold"), bg="#f5f5f5")
        self.lbl_title.pack(side="left")
        self.btn_config = tk.Button(header, command=self.open_config_window)
        self.btn_config.pack(side="right", padx=5)
        self.btn_save_pre = tk.Button(header, command=self.save_preset, bg="#607D8B", fg="white")
        self.btn_save_pre.pack(side="right")

        self.drop_frame = tk.LabelFrame(self.root, bg="#e3f2fd", relief="groove")
        self.drop_frame.pack(pady=5, padx=20, fill="x")
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', lambda e: self.add_images_to_list(self.root.tk.splitlist(e.data)))
        self.btn_select = tk.Button(self.drop_frame, command=self.select_images)
        self.btn_select.pack(pady=10)

        self.lbl_list = tk.Label(self.root, bg="#f5f5f5", font=("Arial", 9))
        self.lbl_list.pack(anchor="w", padx=25)
        container = tk.Frame(self.root, bg="white", bd=1, relief="sunken")
        container.pack(pady=5, padx=20, fill="both", expand=True)
        self.canvas = tk.Canvas(container, bg="white", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        self.scroll_frame = tk.Frame(self.canvas, bg="white")
        self.scroll_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.conf_frame = tk.LabelFrame(self.root, padx=15, pady=10)
        self.conf_frame.pack(pady=10, padx=20, fill="x")
        self.lbl_size = tk.Label(self.conf_frame)
        self.lbl_size.grid(row=0, column=0, sticky="w")
        self.ent_width = tk.Entry(self.conf_frame, width=7); self.ent_width.grid(row=0, column=1)
        self.ent_height = tk.Entry(self.conf_frame, width=7); self.ent_height.grid(row=0, column=3)
        self.keep_aspect = tk.BooleanVar()
        self.chk_aspect = tk.Checkbutton(self.conf_frame, variable=self.keep_aspect)
        self.chk_aspect.grid(row=0, column=4)

        self.lbl_wm_t = tk.Label(self.conf_frame)
        self.lbl_wm_t.grid(row=1, column=0, sticky="w", pady=5)
        self.wm_text = tk.Entry(self.conf_frame, width=15); self.wm_text.grid(row=1, column=1, columnspan=2, sticky="we")
        self.lbl_wm_s = tk.Label(self.conf_frame)
        self.lbl_wm_s.grid(row=1, column=3)
        self.wm_size = tk.Entry(self.conf_frame, width=5); self.wm_size.grid(row=1, column=4)
        
        self.lbl_wm_p = tk.Label(self.conf_frame)
        self.lbl_wm_p.grid(row=2, column=0, sticky="w")
        self.wm_pos = ttk.Combobox(self.conf_frame, values=["左上", "右上", "左下", "右下", "中央"], width=7, state="readonly")
        self.wm_pos.grid(row=2, column=1)
        self.btn_wm_c = tk.Button(self.conf_frame, command=self.choose_wm_color, width=5)
        self.btn_wm_c.grid(row=2, column=2)
        self.btn_wm_pre = tk.Button(self.conf_frame, command=self.preview_watermark, bg="#FF9800", fg="white")
        self.btn_wm_pre.grid(row=2, column=3, columnspan=2, sticky="we")

        self.lbl_save_n = tk.Label(self.conf_frame)
        self.lbl_save_n.grid(row=3, column=0, pady=10, sticky="w")
        self.ent_filename = tk.Entry(self.conf_frame, width=15); self.ent_filename.grid(row=3, column=1, columnspan=2, sticky="we")
        self.combo_ext = ttk.Combobox(self.conf_frame, values=["PNG", "JPEG", "JPG", "GIF", "WEBP"], width=7, state="readonly")
        self.combo_ext.grid(row=3, column=3, columnspan=2)

        self.progress = ttk.Progressbar(self.root, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=30, pady=5)
        self.lbl_status = tk.Label(self.root, bg="#f5f5f5")
        self.lbl_status.pack()

        self.btn_run = tk.Button(self.root, command=self.convert_images, bg="#2196F3", fg="white", font=("bold"), height=2)
        self.btn_run.pack(pady=10, fill="x", padx=100)

    def choose_wm_color(self):
        color = colorchooser.askcolor()[1]
        if color: self.wm_color = color

    def select_images(self):
        paths = filedialog.askopenfilenames(filetypes=[("Images", "*.jpg *.jpeg *.png *.bmp *.webp *.gif")])
        if paths: self.add_images_to_list(paths)

    def add_images_to_list(self, paths):
        for path in paths:
            if any(item["path"] == path for item in self.image_items): continue
            try:
                img = Image.open(path); img.thumbnail((50, 50))
                photo = ImageTk.PhotoImage(img); var = tk.BooleanVar(value=True)
                row = tk.Frame(self.scroll_frame, bg="white", pady=2)
                row.pack(fill="x", expand=True)
                tk.Checkbutton(row, variable=var, bg="white").pack(side="left")
                lbl = tk.Label(row, image=photo, bg="white"); lbl.image = photo; lbl.pack(side="left")
                btn = tk.Button(row, text=os.path.basename(path), relief="flat", bg="white", command=lambda p=path: self.full_preview(p))
                btn.pack(side="left", padx=10)
                self.image_items.append({"path": path, "var": var, "row": row})
            except: pass

    def full_preview(self, path):
        top = tk.Toplevel(self.root); img = Image.open(path); img.thumbnail((800, 600))
        ph = ImageTk.PhotoImage(img); l = tk.Label(top, image=ph); l.image = ph; l.pack()

    def preview_watermark(self):
        if not self.image_items: return
        path = self.image_items[0]["path"]
        with Image.open(path) as img:
            img = self.apply_edits(img)
            top = tk.Toplevel(self.root); img.thumbnail((600, 400)); ph = ImageTk.PhotoImage(img); l = tk.Label(top, image=ph); l.image = ph; l.pack()

    def apply_edits(self, img):
        try:
            w, h = int(self.ent_width.get()), int(self.ent_height.get())
            if self.keep_aspect.get(): h = int(w * (img.height / img.width))
            img = img.resize((w, h), Image.LANCZOS)
        except: pass
        txt = self.wm_text.get()
        if txt:
            draw = ImageDraw.Draw(img)
            try: font = ImageFont.truetype("arial.ttf", int(self.wm_size.get()))
            except: font = ImageFont.load_default()
            bbox = draw.textbbox((0, 0), txt, font=font)
            tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
            pos_map = {"左上": (10, 10), "右上": (img.width-tw-10, 10), "左下": (10, img.height-th-10), "右下": (img.width-tw-10, img.height-th-10), "中央": ((img.width-tw)//2, (img.height-th)//2)}
            draw.text(pos_map.get(self.wm_pos.get(), (10, 10)), txt, fill=self.wm_color, font=font)
        return img

    def convert_images(self):
        targets = [it for it in self.image_items if it["var"].get()]
        if not targets: return
        save_dir = filedialog.askdirectory()
        if not save_dir: return
        self.progress["maximum"] = len(targets)
        ext = self.combo_ext.get().lower()
        for i, item in enumerate(targets):
            self.lbl_status.config(text=f"{MESSAGES[self.lang]['run']}... {i+1}/{len(targets)}")
            self.progress["value"] = i + 1
            self.root.update_idletasks()
            with Image.open(item["path"]) as img:
                img = self.apply_edits(img)
                orig = os.path.splitext(os.path.basename(item["path"]))[0]
                name = self.ent_filename.get().replace("{name}", orig).replace("{n}", f"{i+1:03}")
                if ext in ["jpeg", "jpg"] and img.mode == "RGBA": img = img.convert("RGB")
                img.save(os.path.join(save_dir, f"{name}.{ext}"), ext.upper() if ext != "jpg" else "JPEG")
        self.lbl_status.config(text=MESSAGES[self.lang]["status_done"])
        os.startfile(save_dir)
        messagebox.showinfo("OK", MESSAGES[self.lang]["msg_done"])

    def generate_readme(self):
        content = """# MK Image Converter Expert - README

## 概要 / Overview
このソフトは、複数の画像を直感的な操作で一括変換・加工するためのツールです。
This software is a tool for batch converting and processing multiple images.

## 免責事項 / Disclaimer
本ソフトウェアを使用したことによって生じたいかなる損害についても、作者は一切の責任を負いません。
The author shall not be held responsible for any damages resulting from the use of this software.

## 著作権 / Copyright
© 2025 maaku. All rights reserved.
"""
        with open("README.txt", "w", encoding="utf-8") as f:
            f.write(content)

    def generate_installer_script(self):
        # Inno Setup用スクリプト（同じ階層のexeを使用するように設定）
        iss_content = f"""[Setup]
AppName=MK Image Converter Expert
AppVersion=1.0
DefaultDirName={{pf}}\\MKImageConverter
DefaultGroupName=MK Image Converter
OutputDir=.
OutputBaseFilename=MK_Converter_Installer
SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes

[Files]
Source: "imageconverter.exe"; DestDir: "{{app}}"; Flags: ignoreversion
Source: "README.txt"; DestDir: "{{app}}"; Flags: ignoreversion

[Icons]
Name: "{{group}}\\MK Image Converter"; Filename: "{{app}}\\imageconverter.exe"
Name: "{{commondesktop}}\\MK Image Converter"; Filename: "{{app}}\\imageconverter.exe"

[Run]
Filename: "{{app}}\\imageconverter.exe"; Description: "Launch MK Image Converter"; Flags: nowait postinstall skipifsilent
"""
        with open("installer_script.iss", "w", encoding="utf-8") as f:
            f.write(iss_content)

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ImageConverter(root)
    root.mainloop()