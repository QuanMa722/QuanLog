# -*- coding: utf-8 -*-

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, date
from ttkbootstrap.dialogs import Messagebox
from ttkbootstrap.constants import *
from tkinter import filedialog
from PIL import Image, ImageTk
import ttkbootstrap as ttkbs
import subprocess
import threading
import traceback
import tkinter
import shutil
import queue
import time
import cv2
import os
import re


class QuanLog:
    def __init__(self, root):
        self.root = root
        self.root.title("QuanLog")
        self.root.geometry("1400x700")
        self.root.resizable(False, False)

        self.selected_files = []
        self.current_dir = os.getcwd()
        self.is_working = False

        self.video_play_locked = False
        self.stop_video_flag = False
        self.current_playing_video = None
        self.is_playing_video = False
        self.video_lock = threading.Lock()

        now = datetime.now()
        self.date_parts = {
            'year': now.strftime('%Y'),
            'month': now.strftime('%m'),
            'day': now.strftime('%d'),
            'date_str': now.strftime('%Y%m%d')
        }

        self.current_year = self.date_parts['year']
        self.last_year = str(int(self.current_year) - 1)

        self.date_path = os.path.join(
            self.current_dir,
            self.date_parts['year'],
            self.date_parts['month'],
            self.date_parts['day']
        )

        self.text_save_path = os.path.join(
            self.date_path,
            f"{self.date_parts['date_str']}.txt"
        )

        self.log_file_path = os.path.join(
            self.current_dir,
            self.date_parts['year'],
            f"{self.date_parts['year']}.log"
        )

        self.FILE_TYPES = {
            "direct_open": (
                ".pdf", ".doc", ".docx", ".xls", ".xlsx",
                ".ppt", ".pptx", ".wps", ".et", ".dps"
            ),
            "video": (
                ".mp4", ".avi", ".mov", ".mkv",
                ".flv", ".wmv", ".webm"
            ),
            "text": (
                ".txt", ".py", ".json", ".md", ".ini",
                ".conf", ".log", ".csv", ".java", ".go",
                ".html", ".htm", ".css", ".js", ".ts",
                ".php", ".c", ".cpp", ".h", ".hpp",
                ".sh", ".bat", ".xml", ".yaml", ".yml"
            ),
            "image": (
                ".png", ".jpg", ".jpeg", ".bmp",
                ".gif", ".tiff", ".ico", ".webp"
            )
        }

        self.ninety_days_ago = date.today() - timedelta(days=90)
        cpu_count = os.cpu_count() or 4
        max_workers = max(4, min(cpu_count // 2, 16))
        self.thread_pool = ThreadPoolExecutor(max_workers=max_workers)

        self.log_queue = queue.Queue()
        self._stop_event = threading.Event()

        self._init_ui()
        self._start_log_consumer()
        self.center_window()
        self._load_files_thread()
        self.file_tree.bind("<Control-c>", self.copy_selected_file_to_clipboard)

    def center_window(self):
        self.root.update()
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        win_w = self.root.winfo_width()
        win_h = self.root.winfo_height()
        x = (screen_w - win_w) // 2
        y = (screen_h - win_h) // 2
        self.root.geometry(f"+{x}+{y}")

    def _truncate_filename(self, filename):
        max_width = 180
        char_width = 7
        display_width = 0

        for char in filename:
            display_width += char_width * 2 if '\u4e00' <= char <= '\u9fff' else char_width

        if display_width <= max_width:
            return filename

        name, ext = os.path.splitext(filename)
        if len(name) <= 8:
            return filename

        ext_width = sum(
            char_width * 2 if '\u4e00' <= c <= '\u9fff' else char_width
            for c in ext
        )

        avail_chars = max(6, (max_width - ext_width - 20) // char_width)
        keep_chars = max(3, min(len(name) // 2, avail_chars // 2, 10))

        return f"{name[:keep_chars]}...{name[-keep_chars:]}{ext}"

    def get_unique_file_path(self, target_dir, original_filename, is_text_log=False):
        if is_text_log:
            return os.path.join(target_dir, original_filename)

        name, ext = os.path.splitext(original_filename)
        target_path = os.path.join(target_dir, original_filename)
        count = 1

        while os.path.exists(target_path):
            target_path = os.path.join(target_dir, f"{name}-{count}{ext}")
            count += 1

        return target_path

    def _start_log_consumer(self):
        def consume():
            while not self._stop_event.is_set():
                try:
                    content = self.log_queue.get(timeout=0.1)
                    self.root.after(0, self._update_log_ui, content)

                except queue.Empty:
                    continue

                except Exception:
                    try:
                        self.log_queue.put("Log consumer exception:\n" + traceback.format_exc())

                    except Exception:
                        pass
                    break

        threading.Thread(target=consume, daemon=True).start()

    def _update_log_ui(self, content):
        time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_line = f"[{time_str}] {content}\n"
        try:
            self.log_text.config(state=NORMAL)
            self.log_text.insert("1.0", log_line)
            self.log_text.see("1.0")
            self.log_text.config(state=DISABLED)

        except Exception:
            print("Failed to update log UI:", traceback.format_exc())

    def write_log(self, log_type, **kwargs):
        log_mapping = {
            "text": ("TEXT_SAVE", "Text saved: {content}", "content"),
            "file": ("FILE_UPLOAD", "File saved: {filename} ({size}KB)", "filename, size"),
            "delete": ("FILE_DELETE", "File deleted: {filename}", "filename"),
            "copy": ("FILE_COPY", "Copied: {filename}", "filename")
        }
        if log_type not in log_mapping:
            return

        label, tmpl, _ = log_mapping[log_type]
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            os.makedirs(os.path.dirname(self.log_file_path), exist_ok=True)
            if log_type == "text":
                log_line = f"time: {now_str}, type: {label}, content: {kwargs.get('content', '')}\n"
                ui = tmpl.format(content=kwargs.get('content', ''))

            elif log_type == "file":
                fname = kwargs.get("filename", "")
                size = kwargs.get("size", 0)
                log_line = f"time: {now_str}, type: {label}, filename: {fname}, size: {size}KB\n"
                ui = tmpl.format(filename=fname, size=size)

            else:
                fname = kwargs.get("filename", "")
                log_line = f"time: {now_str}, type: {label}, filename: {fname}\n"
                ui = tmpl.format(filename=fname)

            with open(self.log_file_path, "a", encoding="utf-8") as f:
                f.write(log_line)
            self.log_queue.put(ui)

        except Exception:
            try:
                self.log_queue.put("write_log exception:\n" + traceback.format_exc())

            except Exception:
                pass

    def _init_ui(self):
        main_frame = ttkbs.Frame(self.root, padding=2)
        main_frame.pack(fill=BOTH, expand=True)

        left_frame = ttkbs.Frame(main_frame, width=350)
        left_frame.pack(side=LEFT, fill=BOTH, padx=(0, 0))
        left_frame.pack_propagate(False)

        core_frame = ttkbs.LabelFrame(left_frame, text="Text & File Input", font=("Times New Roman", 10, "bold"))
        core_frame.pack(fill=BOTH, expand=True, padx=0, pady=0, ipady=0)

        self.text_input = tkinter.Text(core_frame, height=14, width=30, wrap=WORD, font=("Times New Roman", 14))
        self.text_input.pack(fill=BOTH, expand=True, padx=5, pady=5)
        self.text_input.edit_modified(False)
        self.text_input.bind("<<Modified>>", self._on_text_change)
        self.text_input.bind("<MouseWheel>", self._on_mouse_wheel)

        btn_frame = ttkbs.Frame(core_frame)
        btn_frame.pack(fill=X, padx=5, pady=5)

        self.select_btn = ttkbs.Button(
            btn_frame, text="Files", command=self.select_files, bootstyle=PRIMARY, width=15
        )
        self.upload_btn = ttkbs.Button(
            btn_frame, text="Upload", command=self._upload_thread, bootstyle=SUCCESS,
            width=15, state=DISABLED
        )
        self.select_btn.pack(side=LEFT, padx=5, expand=True)
        self.upload_btn.pack(side=RIGHT, padx=5, expand=True)

        log_frame = ttkbs.LabelFrame(left_frame, text="Operation Log", font=("Times New Roman", 10, "bold"))
        log_frame.pack(fill=BOTH, expand=True, padx=0, pady=0, ipady=0)

        self.log_text = tkinter.Text(log_frame, height=10, width=30, wrap=WORD, font=("Times New Roman", 14))
        self.log_text.pack(fill=BOTH, expand=True, padx=5, pady=5)
        self.log_text.config(state=DISABLED)
        self.log_text.bind("<MouseWheel>", self._on_mouse_wheel)

        mid_frame = ttkbs.Frame(main_frame, width=500)
        mid_frame.pack(side=LEFT, fill=BOTH, expand=True)
        mid_frame.pack_propagate(False)

        self.preview_frame = ttkbs.Frame(mid_frame, padding=5)
        self.preview_frame.pack(fill=BOTH, expand=True)
        self.preview_text = tkinter.Text(self.preview_frame, wrap=WORD, state=DISABLED, font=("Times New Roman", 15))
        self.preview_scrollbar = ttkbs.Scrollbar(self.preview_frame, orient=VERTICAL, command=self.preview_text.yview)
        self.preview_text.config(yscrollcommand=self.preview_scrollbar.set)
        self.preview_text.pack(fill=BOTH, expand=True, side=LEFT)
        self.preview_scrollbar.pack(side=RIGHT, fill=Y)
        self.preview_text.bind("<MouseWheel>", self._on_mouse_wheel)

        right_frame = ttkbs.Frame(main_frame, width=200)
        right_frame.pack(side=RIGHT, fill=BOTH, expand=True)
        right_frame.pack_propagate(False)

        search_frame = ttkbs.Frame(right_frame, padding=(2, 2, 2, 2))
        search_frame.pack(fill=X, side=TOP)

        ttkbs.Label(search_frame, text="Search:", font=("Times New Roman", 10, "bold")).pack(
            side=LEFT, padx=(0, 2)
        )

        self.search_entry = ttkbs.Entry(search_frame, font=("Times New Roman", 10))
        self.search_entry.pack(fill=X, expand=True, side=LEFT)
        self.search_entry.bind("<KeyRelease>", self.filter_file_list)

        file_list_frame = ttkbs.Frame(right_frame)
        file_list_frame.pack(fill=BOTH, expand=True, side=BOTTOM)

        self.file_tree = ttkbs.Treeview(file_list_frame, show="tree")
        self.file_tree.pack(fill=BOTH, expand=True)
        self.file_tree.bind("<<TreeviewSelect>>", self.on_file_click)
        self.file_tree.bind("<Double-1>", self.on_file_double_click)
        self.file_tree.bind("<Delete>", self.on_delete_file)

    @staticmethod
    def _on_mouse_wheel(event):
        event.widget.yview_scroll(-int(event.delta / 120), "units")
        return "break"

    def _on_text_change(self, event):
        if self.text_input.edit_modified():
            txt = self.text_input.get("1.0", END).strip()

            if not self.is_working:
                self.upload_btn.config(state=NORMAL if (txt or self.selected_files) else DISABLED)
            self.text_input.edit_modified(False)

    def select_files(self):
        if self.is_working:
            Messagebox.show_info("Info", "Task is running, please wait!")
            return

        files = filedialog.askopenfilenames(
            title="Select Files",
            filetypes=[("All Files", "*.*"), ("Text Files", "*.txt"), ("Excel Files", "*.xlsx;*.xls;*.csv")]
        )

        try:
            self.selected_files = [f for f in set(files) if os.path.isfile(f)]
            self.log_queue.put(f"{len(self.selected_files)} files selected")
            txt = self.text_input.get("1.0", END).strip()
            self.upload_btn.config(state=NORMAL if (txt or self.selected_files) else DISABLED)

        except Exception:
            self.log_queue.put("select_files exception:\n" + traceback.format_exc())

    def _load_files_worker(self):
        cache = {}
        try:
            if not os.path.exists(self.current_dir):
                self.root.after(0, self._update_file_tree, cache)
                return

            for year in os.listdir(self.current_dir):
                yp = os.path.join(self.current_dir, year)
                if not self._is_valid_dir(yp, year, 4):
                    continue

                for month in os.listdir(yp):
                    mp = os.path.join(yp, month)
                    if not self._is_valid_dir(mp, month, 2):
                        continue

                    for day in os.listdir(mp):
                        dp = os.path.join(mp, day)
                        if not self._is_valid_dir(dp, day, 2):
                            continue
                        date_str = f"{year}-{month}-{day}"

                        try:
                            if datetime.strptime(date_str, "%Y-%m-%d").date() < self.ninety_days_ago:
                                continue

                        except Exception:
                            continue

                        files = []
                        for f in os.listdir(dp):
                            fp = os.path.join(dp, f)
                            if os.path.isfile(fp) and not f.endswith(".log"):
                                files.append(fp)
                        if files:
                            cache[date_str] = files

        except Exception:
            self.log_queue.put("_load_files_worker exception:\n" + traceback.format_exc())

        finally:
            self.root.after(0, self._update_file_tree, cache)

    @staticmethod
    def _is_valid_dir(path, name, length):
        return os.path.isdir(path) and name.isdigit() and len(name) == length

    def _update_file_tree(self, cache):
        try:
            self.file_tree.delete(*self.file_tree.get_children())

            for date_str in sorted(cache.keys(), reverse=True):
                node = self.file_tree.insert("", END, text=f"📅 {date_str}", open=True)

                for fp in sorted(cache[date_str]):
                    name = self._truncate_filename(os.path.basename(fp))
                    self.file_tree.insert(node, END, text=name, values=(fp,))

        except Exception:
            self.log_queue.put("_update_file_tree exception:\n" + traceback.format_exc())

    def _load_files_thread(self):
        try:
            self.thread_pool.submit(self._load_files_worker)

        except Exception:
            self.log_queue.put("_load_files_thread exception:\n" + traceback.format_exc())

    def filter_file_list(self, event):
        kw = self.search_entry.get().strip().lower()
        if not kw:
            self._load_files_thread()
            return
        res = {}

        try:
            self._fast_search(kw, res)
            self.file_tree.delete(*self.file_tree.get_children())
            if not res:
                self.file_tree.insert("", END, text="No matching files found")
                return

            for d in sorted(res.keys(), reverse=True):
                node = self.file_tree.insert("", END, text=f"📅 {d}", open=True)
                for fp in sorted(res[d]):
                    self.file_tree.insert(
                        node, END, text=self._truncate_filename(os.path.basename(fp)), values=(fp,)
                    )

        except Exception:
            self.log_queue.put("filter_file_list exception:\n" + traceback.format_exc())

    def _fast_search(self, keyword, result):
        target_years = {self.current_year, self.last_year}
        root = self.current_dir
        if not os.path.exists(root):
            return

        for year_dir in os.listdir(root):
            if year_dir not in target_years:
                continue

            year_path = os.path.join(root, year_dir)
            if not os.path.isdir(year_path):
                continue

            for month_dir in os.listdir(year_path):
                if not month_dir.isdigit() or len(month_dir) != 2:
                    continue

                month_path = os.path.join(year_path, month_dir)
                if not os.path.isdir(month_path):
                    continue

                for day_dir in os.listdir(month_path):
                    if not day_dir.isdigit() or len(day_dir) != 2:
                        continue

                    day_path = os.path.join(month_path, day_dir)
                    if not os.path.isdir(day_path):
                        continue
                    date_key = f"{year_dir}-{month_dir}-{day_dir}"

                    try:
                        files = []
                        for fname in os.listdir(day_path):
                            fpath = os.path.join(day_path, fname)

                            if os.path.isfile(fpath) and not fname.endswith(".log") and keyword in fname.lower():
                                files.append(fpath)
                        if files:
                            result[date_key] = files

                    except Exception:
                        continue

    def _upload_worker(self):
        try:
            self.root.after(0, lambda: (
                setattr(self, 'is_working', True),
                self.upload_btn.config(state=DISABLED),
                self.select_btn.config(state=DISABLED)
            ))
            os.makedirs(self.date_path, exist_ok=True)
            text = self.text_input.get("1.0", END).strip()

            if text:
                target = self.get_unique_file_path(
                    os.path.dirname(self.text_save_path),
                    os.path.basename(self.text_save_path),
                    True
                )

                with open(target, "a", encoding="utf-8") as f:
                    f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n{text}\n\n")
                self.write_log("text", content=text)

            if self.selected_files:
                for src in self.selected_files:
                    try:
                        if not os.path.isfile(src):
                            continue
                        fname = os.path.basename(src)
                        is_text_log = bool(fname.endswith('.txt') and re.match(r'^\d{8}\.txt$', fname))
                        target = self.get_unique_file_path(self.date_path, fname, is_text_log)
                        shutil.copy2(src, target)
                        size = round(os.path.getsize(src) / 1024, 2)
                        self.write_log("file", filename=os.path.basename(target), size=size)

                    except Exception:
                        self.log_queue.put("_upload_worker file copy exception:\n" + traceback.format_exc())

        except Exception:
            self.log_queue.put("_upload_worker exception:\n" + traceback.format_exc())

        finally:
            self.root.after(0, lambda: (
                self._load_files_thread(),
                self._reset_state(),
                self.select_btn.config(state=NORMAL),
                setattr(self, 'is_working', False)
            ))

    def _upload_thread(self):
        if self.is_working:
            Messagebox.show_info("Info", "Task is running!")
            return

        try:
            self.thread_pool.submit(self._upload_worker)

        except Exception:
            self.log_queue.put("_upload_thread exception:\n" + traceback.format_exc())

    def _reset_state(self):
        try:
            self.selected_files.clear()
            self.text_input.delete("1.0", END)
            self.upload_btn.config(state=DISABLED)
            self.text_input.edit_modified(False)

        except Exception:
            self.log_queue.put("_reset_state exception:\n" + traceback.format_exc())

    def on_delete_file(self, event):
        if self.is_working:
            return
        sel = self.file_tree.selection()

        if not sel:
            return
        item = self.file_tree.item(sel[0])

        if item["text"].startswith("📅"):
            return
        fp = item["values"][0] if item["values"] else os.path.join(self.current_dir, item["text"])

        if fp == self.current_playing_video:
            self.safe_stop_video()

        try:
            self.root.after(200, self.thread_pool.submit, self._delete_worker, fp, sel[0])

        except Exception:
            self.log_queue.put("on_delete_file exception:\n" + traceback.format_exc())

    def _delete_worker(self, fp, tree_item):
        for i in range(3):
            try:
                if os.path.exists(fp):
                    os.remove(fp)
                self.root.after(0, lambda: (
                    self.file_tree.delete(tree_item),
                    self.switch_preview("text"),
                    self._load_files_thread()
                ))
                self.write_log("delete", filename=os.path.basename(fp))
                return

            except PermissionError:
                self.log_queue.put(f"Delete retry {i + 1}/3: File is in use")
                time.sleep(0.1)

            except Exception:
                self.log_queue.put("_delete_worker exception:\n" + traceback.format_exc())
                time.sleep(0.1)

        self.log_queue.put("Delete failed: File is in use by another program")

    def copy_selected_file_to_clipboard(self, event=None):
        sel = self.file_tree.selection()
        if not sel:
            return

        item = self.file_tree.item(sel[0])
        if item["text"].startswith("📅"):
            return

        fp = item["values"][0] if item["values"] else None
        if not fp or not os.path.isfile(fp):
            return

        try:
            subprocess.run(
                f'powershell -Command "Set-Clipboard -Path \'{fp.replace(chr(39), chr(39) * 2)}\'"',
                shell=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW
            )
            self.write_log("copy", filename=os.path.basename(fp))

        except Exception:
            self.log_queue.put("copy_selected_file_to_clipboard exception:\n" + traceback.format_exc())

    def on_file_click(self, event):
        sel = self.file_tree.selection()
        if not sel:
            return

        item = self.file_tree.item(sel[0])
        if item["text"].startswith("📅"):
            return

        fp = item["values"][0] if item["values"] else os.path.join(self.current_dir, item["text"])
        try:
            if os.path.splitext(fp)[-1].lower() not in self.FILE_TYPES["direct_open"]:
                self._preview_file(fp)

        except Exception:
            self.log_queue.put("on_file_click exception:\n" + traceback.format_exc())

    def on_file_double_click(self, event):
        sel = self.file_tree.selection()
        if not sel:
            return

        item = self.file_tree.item(sel[0])
        if item["text"].startswith("📅"):
            return

        fp = item["values"][0] if item["values"] else os.path.join(self.current_dir, item["text"])
        ext = os.path.splitext(fp)[-1].lower()
        try:
            if ext in self.FILE_TYPES["video"]:
                self._preview_file(fp)

            elif ext in self.FILE_TYPES["direct_open"]:
                self.thread_pool.submit(self._open_file, fp)

            else:
                self._preview_file(fp)

        except Exception:
            self.log_queue.put("on_file_double_click exception:\n" + traceback.format_exc())

    def _open_file(self, fp):
        try:
            if os.path.exists(fp):
                os.startfile(fp)
                self.log_queue.put(f"File opened: {os.path.basename(fp)}")

        except Exception:
            self.log_queue.put("_open_file exception:\n" + traceback.format_exc())

    def _preview_file(self, fp):
        if not os.path.isfile(fp):
            return

        ext = os.path.splitext(fp)[-1].lower()
        try:
            if ext in self.FILE_TYPES["text"]:
                self.switch_preview("text")
                self.thread_pool.submit(self._preview_text, fp)

            elif ext in self.FILE_TYPES["image"]:
                self.switch_preview("media")
                self.thread_pool.submit(self._preview_image, fp)

            elif ext in self.FILE_TYPES["video"]:
                if self.current_playing_video != fp or not self.is_playing_video:
                    self.safe_stop_video()
                    self.switch_preview("media")
                    self._play_video(fp)

        except Exception:
            self.log_queue.put("_preview_file exception:\n" + traceback.format_exc())

    def switch_preview(self, typ):
        try:
            self.safe_stop_video()
            self.preview_text.pack_forget()
            self.preview_scrollbar.pack_forget()
            self.preview_text.config(state=NORMAL)
            self.preview_text.delete("1.0", END)
            self.preview_text.config(state=DISABLED)

            if typ == "text":
                self.preview_text.pack(fill=BOTH, expand=True, side=LEFT)
                self.preview_scrollbar.pack(side=RIGHT, fill=Y)

            elif typ == "media":
                self.preview_media_label = ttkbs.Label(self.preview_frame, anchor=CENTER)
                self.preview_media_label.pack(fill=BOTH, expand=True, side=LEFT)

        except Exception:
            self.log_queue.put("switch_preview exception:\n" + traceback.format_exc())

    def _preview_text(self, fp):
        content = ""
        try:
            for enc in ["utf-8", "gbk", "gb2312", "utf-8-sig"]:
                try:
                    with open(fp, "r", encoding=enc) as f:
                        lines = f.readlines()
                        if fp.lower().endswith(".log"):
                            res = []
                            for line in reversed(lines):
                                line = line.strip()
                                if line:
                                    res.append(f"{self._format_log_line(line)}\n\n")
                            content = "".join(res)
                        else:
                            content = "".join(lines)
                    break

                except Exception:
                    continue

        except Exception:
            self.log_queue.put("_preview_text exception:\n" + traceback.format_exc())
            content = ""

        self.root.after(0, self._update_text_preview, content)

    def _update_text_preview(self, content):
        try:
            self.preview_text.config(state=NORMAL)
            self.preview_text.delete("1.0", END)
            if content:
                self.preview_text.insert("1.0", content)
            self.preview_text.config(state=DISABLED)

        except Exception:
            self.log_queue.put("_update_text_preview exception:\n" + traceback.format_exc())

    @staticmethod
    def _format_log_line(line):
        parts = re.split(r',\s*', line)
        res = []
        for p in parts:
            if ':' in p:
                k, v = p.split(':', 1)
                res.append(f"{k.strip()}: {v.strip()}")

            else:
                res.append(p.strip())

        return "\n".join(res)

    def _preview_image(self, fp):
        try:
            img = Image.open(fp)
            w = max(self.preview_frame.winfo_width() - 20, 480)
            h = max(self.preview_frame.winfo_height() - 10, 480)
            scale = min(w / img.width, h / img.height, 1)
            img = img.resize((int(img.width * scale), int(img.height * scale)), Image.Resampling.LANCZOS)
            self.root.after(0, self._update_image_preview, img)

        except Exception:
            self.log_queue.put("_preview_image exception:\n" + traceback.format_exc())

    def _update_image_preview(self, img):
        try:
            tkimg = ImageTk.PhotoImage(img)
            self.preview_media_label.config(image=tkimg)
            self.preview_media_label.image = tkimg

        except Exception:
            self.log_queue.put("_update_image_preview exception:\n" + traceback.format_exc())

    def safe_stop_video(self):
        with self.video_lock:
            self.stop_video_flag = True
            self.is_playing_video = False

        if hasattr(self, "preview_media_label") and self.preview_media_label:
            try:
                self.preview_media_label.config(image="")
                self.preview_media_label.destroy()

            except Exception:
                self.log_queue.put("safe_stop_video preview_media_label destroy exception:\n" + traceback.format_exc())
            self.preview_media_label = None

        try:
            cv2.destroyAllWindows()
            cv2.waitKey(1)

        except Exception:
            self.log_queue.put("safe_stop_video cv2.destroyAllWindows exception:\n" + traceback.format_exc())

        self.current_playing_video = None
        self.video_play_locked = False

    def _play_video(self, fp):
        self.stop_video_flag = False
        self.current_playing_video = fp
        self.video_play_locked = True
        try:
            threading.Thread(target=self._video_worker, args=(fp,), daemon=True).start()

        except Exception:
            self.log_queue.put("_play_video thread start exception:\n" + traceback.format_exc())

    def _get_scaled_frame_size(self, fw, fh, mw, mh):
        try:
            s = min(mw / fw, mh / fh, 1.0) if fw and fh else 1.0
            return int(fw * s), int(fh * s)

        except Exception:
            self.log_queue.put("_get_scaled_frame_size exception:\n" + traceback.format_exc())
            return fw, fh

    def _video_worker(self, fp):
        with self.video_lock:
            self.is_playing_video = True
        self.log_queue.put(f"Playing: {os.path.basename(fp)}")
        cap = None

        try:
            cap = cv2.VideoCapture(fp)
            if not cap.isOpened():
                self.log_queue.put(f"Failed to open video: {os.path.basename(fp)}")
                self.safe_stop_video()
                return
            fps = cap.get(cv2.CAP_PROP_FPS) or 0.0

            try:
                delay = 1 / max(fps, 15)

            except Exception:
                delay = 1 / 15.0

            w, h = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH)), int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
            mw, mh = max(self.preview_frame.winfo_width() - 20, 1), max(self.preview_frame.winfo_height() - 10, 1)
            dw, dh = self._get_scaled_frame_size(w, h, mw, mh)

            while not self.stop_video_flag and self.is_playing_video:
                if self.current_playing_video != fp:
                    break

                ret, frame = cap.read()
                if not ret:
                    break

                try:
                    frame = cv2.resize(frame, (dw, dh), interpolation=cv2.INTER_AREA)
                    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    img = ImageTk.PhotoImage(Image.fromarray(frame))
                    self.root.after(0, self._update_video_frame, img)

                except Exception:
                    self.log_queue.put("_video_worker frame processing exception:\n" + traceback.format_exc())
                time.sleep(delay)

        except Exception:
            self.log_queue.put("_video_worker exception:\n" + traceback.format_exc())

        finally:
            try:
                if cap is not None:
                    cap.release()

            except Exception:
                self.log_queue.put("_video_worker cap.release exception:\n" + traceback.format_exc())

            try:
                cv2.destroyAllWindows()

            except Exception:
                pass

            with self.video_lock:
                self.is_playing_video = False
            self.video_play_locked = False
            self.log_queue.put(f"Stopped: {os.path.basename(fp)}")

    def _update_video_frame(self, img):
        try:
            if hasattr(self, "preview_media_label") and self.preview_media_label and not self.stop_video_flag:
                self.preview_media_label.config(image=img)
                self.preview_media_label.image = img

        except Exception:
            self.log_queue.put("_update_video_frame exception:\n" + traceback.format_exc())

    def stop(self):
        try:
            self._stop_event.set()
            self.safe_stop_video()
            try:
                self.thread_pool.shutdown(wait=False)

            except Exception:
                pass

        except Exception:
            print("stop exception:", traceback.format_exc())

    def __del__(self):
        try:
            self.stop()

        except Exception:
            pass


if __name__ == "__main__":
    app = ttkbs.Window(themename="flatly")
    gui = QuanLog(app)

    try:
        app.mainloop()

    finally:
        try:
            gui.stop()

        except Exception:
            pass
