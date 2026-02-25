# -*- coding: utf-8 -*-

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, date
from ttkbootstrap.dialogs import Messagebox
from ttkbootstrap.constants import *
from tkinter import filedialog
from PIL import Image, ImageTk
import ttkbootstrap as ttkbs
import threading
import tkinter
import shutil
import queue
import time
import cv2
import os
import re


# V1.1
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

        now = datetime.now()
        self.date_parts = {
            'year': now.strftime('%Y'),
            'month': now.strftime('%m'),
            'day': now.strftime('%d'),
            'date_str': now.strftime('%Y%m%d')
        }
        self.date_path = os.path.join(self.current_dir, self.date_parts['year'], self.date_parts['month'],
                                      self.date_parts['day'])
        self.text_save_path = os.path.join(self.date_path, f"{self.date_parts['date_str']}.txt")
        self.log_file_path = os.path.join(self.current_dir, self.date_parts['year'], f"{self.date_parts['year']}.log")

        self.FILE_TYPES = {
            "direct_open": (".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".wps", ".et", ".dps"),
            "video": (".mp4", ".avi", ".mov", ".mkv", ".flv", ".wmv", ".webm"),
            "text": (".txt", ".py", ".json", ".md", ".ini", ".conf", ".log", ".csv",
                     ".java", ".go", ".html", ".htm", ".css", ".js", ".ts", ".php",
                     ".c", ".cpp", ".h", ".hpp", ".sh", ".bat", ".xml", ".yaml", ".yml"),
            "image": (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff", ".ico", ".webp")
        }

        self.ninety_days_ago = date.today() - timedelta(days=90)
        self.thread_pool = ThreadPoolExecutor(max_workers=4)
        self.log_queue = queue.Queue()
        self.video_lock = threading.Lock()

        self._init_ui()
        self._start_log_consumer()
        self.center_window()
        self._load_files_thread()

    def center_window(self):
        self.root.update_idletasks()
        screen_w, screen_h = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        win_w, win_h = self.root.winfo_width(), self.root.winfo_height()
        x = (screen_w - win_w) // 2
        y = (screen_h - win_h) // 2
        self.root.geometry(f"+{x}+{y}")

    def _truncate_filename(self, filename, max_length=30):
        if len(filename) <= max_length:
            return filename

        name, ext = os.path.splitext(filename)
        ext_len = len(ext)

        if ext_len >= max_length:
            return filename[:max_length]

        avail_len = max_length - ext_len - 3
        if avail_len <= 0:
            return filename[:max_length]

        half1 = avail_len * 2 // 3
        half2 = avail_len - half1
        return f"{name[:half1]}......{name[-half2:]}{ext}"

    def get_unique_file_path(self, target_dir, original_filename, is_text_log=False):
        if is_text_log:
            return os.path.join(target_dir, original_filename)

        name, ext = os.path.splitext(original_filename)
        target_path = os.path.join(target_dir, original_filename)

        if not os.path.exists(target_path):
            return target_path

        count = 1
        while True:
            new_path = os.path.join(target_dir, f"{name}-{count}{ext}")
            if not os.path.exists(new_path):
                return new_path
            count += 1

    def _start_log_consumer(self):
        def consume_logs():
            while True:
                try:
                    content = self.log_queue.get(timeout=0.1)
                    self.root.after(0, lambda c=content: self._update_log_ui(c))
                except queue.Empty:
                    pass
                except:
                    break

        threading.Thread(target=consume_logs, daemon=True).start()

    def _update_log_ui(self, content):
        time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_line = f"[{time_str}] {content}\n"

        self.log_text.config(state=NORMAL)
        self.log_text.insert("1.0", log_line)
        self.log_text.see("1.0")
        self.log_text.config(state=DISABLED)

    def write_log(self, log_type, **kwargs):
        log_mapping = {
            "text": ("TEXT_SAVE", "Text saved: {content}", "content"),
            "file": ("FILE_UPLOAD", "File saved: {filename} ({size}KB)", "filename, size"),
            "delete": ("FILE_DELETE", "File deleted: {filename}", "filename")
        }

        if log_type not in log_mapping:
            return

        log_label, ui_template, params = log_mapping[log_type]
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        os.makedirs(os.path.dirname(self.log_file_path), exist_ok=True)

        if log_type == "text":
            content = kwargs.get("content", "")
            log_line = f"time: {current_time}, type: {log_label}, content: {content}\n"
            ui_content = ui_template.format(content=content)
        elif log_type == "file":
            filename = kwargs.get("filename", "")
            size = kwargs.get("size", 0)
            log_line = f"time: {current_time}, type: {log_label}, filename: {filename}, size: {size}KB\n"
            ui_content = ui_template.format(filename=filename, size=size)
        else:
            filename = kwargs.get("filename", "")
            log_line = f"time: {current_time}, type: {log_label}, filename: {filename}\n"
            ui_content = ui_template.format(filename=filename)

        with open(self.log_file_path, "a", encoding="utf-8") as f:
            f.write(log_line)
        self.log_queue.put(ui_content)

    def _init_ui(self):
        main_frame = self._create_frame(self.root, padding=5, fill=BOTH, expand=True)

        left_frame = self._create_frame(main_frame, width=350, fill=BOTH, expand=False, side=LEFT, padx=(0, 5))
        left_frame.pack_propagate(False)

        core_frame = self._create_label_frame(left_frame, "Text & File Input", padx=10, pady=10, ipady=5)

        self.text_input = tkinter.Text(
            core_frame, height=18, width=30, wrap=WORD, font=("Times New Roman", 10)
        )
        self.text_input.pack(fill=BOTH, expand=YES, padx=5, pady=5)
        self.text_input.edit_modified(False)
        self.text_input.bind("<<Modified>>", self._on_text_change)
        self.text_input.bind("<MouseWheel>", self._on_mouse_wheel)

        btn_frame = self._create_frame(core_frame, fill=X, padx=5, pady=5)
        self.select_btn = ttkbs.Button(
            btn_frame, text="Files", command=self.select_files, bootstyle=PRIMARY, width=15
        )
        self.upload_btn = ttkbs.Button(
            btn_frame, text="Upload", command=self._upload_thread, bootstyle=SUCCESS, width=15, state=DISABLED
        )
        self.select_btn.pack(side=LEFT, padx=5, expand=YES)
        self.upload_btn.pack(side=RIGHT, padx=5, expand=YES)

        log_frame = self._create_label_frame(left_frame, "Operation Log", padx=10, pady=10, ipady=5)
        self.log_text = tkinter.Text(
            log_frame, height=10, width=30, wrap=WORD, font=("Times New Roman", 12)
        )
        self.log_text.pack(fill=BOTH, expand=YES, padx=5, pady=5)
        self.log_text.config(state=DISABLED)
        self.log_text.bind("<MouseWheel>", self._on_mouse_wheel)

        mid_frame = self._create_frame(main_frame, width=500, fill=BOTH, expand=True, side=LEFT, padx=(0, 5))
        mid_frame.pack_propagate(False)

        self.preview_frame = self._create_frame(mid_frame, padding=10, fill=BOTH, expand=True)

        self.preview_text = tkinter.Text(
            self.preview_frame, wrap=WORD, state=DISABLED, font=("Times New Roman", 15)
        )
        self.preview_scrollbar = ttkbs.Scrollbar(
            self.preview_frame, orient=VERTICAL, command=self.preview_text.yview
        )
        self.preview_text.config(yscrollcommand=self.preview_scrollbar.set)
        self.preview_text.pack(fill=BOTH, expand=YES, side=LEFT)
        self.preview_scrollbar.pack(side=RIGHT, fill=Y)
        self.preview_text.bind("<MouseWheel>", self._on_mouse_wheel)

        right_frame = self._create_frame(main_frame, width=300, fill=BOTH, expand=True, side=RIGHT)
        right_frame.pack_propagate(False)

        search_frame = self._create_frame(right_frame, padding=(5, 5, 5, 0), fill=X, side=TOP)
        ttkbs.Label(
            search_frame, text="Search:", font=("Times New Roman", 10, "bold")
        ).pack(side=LEFT, padx=(0, 5))
        self.search_entry = ttkbs.Entry(search_frame, font=("Times New Roman", 10))
        self.search_entry.pack(fill=X, expand=YES, side=LEFT)
        self.search_entry.bind("<KeyRelease>", self.filter_file_list)

        file_list_frame = self._create_frame(right_frame, padding=5, fill=BOTH, expand=True, side=BOTTOM)
        self.file_tree = ttkbs.Treeview(file_list_frame, show="tree")
        self.file_tree.pack(fill=BOTH, expand=True)
        self.file_tree.bind("<<TreeviewSelect>>", self.on_file_click)
        self.file_tree.bind("<Double-1>", self.on_file_double_click)
        self.file_tree.bind("<Delete>", self.on_delete_file)

    def _create_frame(self, parent, **kwargs):
        frame = ttkbs.Frame(parent, **{k: v for k, v in kwargs.items() if k in ['width', 'height', 'padding']})
        frame.pack(**{k: v for k, v in kwargs.items() if k in ['fill', 'expand', 'side', 'padx', 'pady']})
        return frame

    def _create_label_frame(self, parent, text, **kwargs):
        frame = ttkbs.LabelFrame(
            parent, text=text, font=("Times New Roman", 10, "bold"),
            **{k: v for k, v in kwargs.items() if k in ['width', 'height', 'padding']}
        )
        frame.pack(**{k: v for k, v in kwargs.items() if k in ['fill', 'expand', 'side', 'padx', 'pady', 'ipady']})
        return frame

    def _on_mouse_wheel(self, event):
        event.widget.yview_scroll(-int(event.delta / 120), "units")
        return "break"

    def _on_text_change(self, event):
        if self.text_input.edit_modified():
            txt = self.text_input.get(1.0, END).strip()
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

        self.selected_files = [f for f in set(files) if os.path.exists(f)]
        self.log_queue.put(f"{len(self.selected_files)} files selected")

        txt = self.text_input.get(1.0, END).strip()
        self.upload_btn.config(state=NORMAL if (txt or self.selected_files) else DISABLED)

    def _load_files_worker(self):
        cache = {}
        root = self.current_dir

        self.root.after(0, lambda: self.file_tree.delete(*self.file_tree.get_children()))

        log_file_path = os.path.join(self.current_dir, self.date_parts['year'], f"{self.date_parts['year']}.log")
        if os.path.exists(log_file_path):
            display_name = self._truncate_filename(os.path.basename(log_file_path))
            self.file_tree.insert("", 0, text=display_name, values=(log_file_path,))

        if os.path.exists(root):
            for year in os.listdir(root):
                year_path = os.path.join(root, year)
                if not self._is_valid_dir(year_path, year, 4):
                    continue

                for month in os.listdir(year_path):
                    month_path = os.path.join(year_path, month)
                    if not self._is_valid_dir(month_path, month, 2):
                        continue

                    for day in os.listdir(month_path):
                        day_path = os.path.join(month_path, day)
                        if not self._is_valid_dir(day_path, day, 2):
                            continue

                        date_str = f"{year}-{month}-{day}"
                        try:
                            folder_date = datetime.strptime(date_str, "%Y-%m-%d").date()
                            if folder_date < self.ninety_days_ago:
                                continue
                        except:
                            continue

                        files = [os.path.join(day_path, f) for f in os.listdir(day_path)
                                 if os.path.isfile(os.path.join(day_path, f))]
                        if files:
                            cache[date_str] = files

        self.root.after(0, lambda: self._update_file_tree(cache))

    def _is_valid_dir(self, path, name, expected_len):
        return (os.path.isdir(path) and name.isdigit() and len(name) == expected_len)

    def _update_file_tree(self, cache):
        dates = sorted(cache.keys(), reverse=True)
        for date_str in dates:
            node = self.file_tree.insert("", END, text=f"📅 {date_str}", open=True)
            for file_path in sorted(cache[date_str]):
                display_name = self._truncate_filename(os.path.basename(file_path))
                self.file_tree.insert(node, END, text=display_name, values=(file_path,))

    def _load_files_thread(self):
        self.thread_pool.submit(self._load_files_worker)

    def filter_file_list(self, event):
        keyword = self.search_entry.get().strip().lower()
        self.file_tree.delete(*self.file_tree.get_children())

        if not keyword:
            self._load_files_thread()
            return

        result = {}
        self._search_files(self.current_dir, keyword, result)

        if not result:
            self.file_tree.insert("", END, text="No matching files found")
            return

        for date_str in sorted(result.keys(), reverse=True):
            node = self.file_tree.insert("", END, text=f"📅 {date_str}", open=True)
            for file_path in sorted(result[date_str]):
                display_name = self._truncate_filename(os.path.basename(file_path))
                self.file_tree.insert(node, END, text=display_name, values=(file_path,))

    def _search_files(self, folder, keyword, result):
        for item in os.listdir(folder):
            item_path = os.path.join(folder, item)
            if os.path.isdir(item_path) and not item.startswith('.'):
                self._search_files(item_path, keyword, result)
            elif os.path.isfile(item_path) and keyword in item.lower():
                match = re.search(r'(\d{4})[\\/](\d{2})[\\/](\d{2})', item_path)
                date_str = f"{match.group(1)}-{match.group(2)}-{match.group(3)}" if match else "Other Files"
                result.setdefault(date_str, []).append(item_path)

    def _upload_worker(self):
        self.root.after(0, lambda: [
            setattr(self, 'is_working', True),
            self.upload_btn.config(state=DISABLED),
            self.select_btn.config(state=DISABLED)
        ])

        input_text = self.text_input.get(1.0, END).strip()
        os.makedirs(self.date_path, exist_ok=True)

        if input_text:
            save_content = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n{input_text}\n\n"
            fname = os.path.basename(self.text_save_path)
            target_path = self.get_unique_file_path(os.path.dirname(self.text_save_path), fname, True)
            with open(target_path, "a", encoding="utf-8") as f:
                f.write(save_content)
            self.write_log("text", content=input_text)

        if self.selected_files:
            self.log_queue.put(f"Start saving {len(self.selected_files)} files")
            for idx, file_path in enumerate(self.selected_files, 1):
                filename = os.path.basename(file_path)
                is_txt_log = filename.endswith('.txt') and re.match(r'^\d{8}\.txt$', filename)
                target_path = self.get_unique_file_path(self.date_path, filename, is_txt_log)
                shutil.copy2(file_path, target_path)
                size = round(os.path.getsize(file_path) / 1024, 2)
                self.write_log("file", filename=os.path.basename(target_path), size=size)

        self.root.after(0, lambda: [
            self._load_files_thread(),
            self._reset_state(),
            self.select_btn.config(state=NORMAL),
            setattr(self, 'is_working', False)
        ])

    def _upload_thread(self):
        if self.is_working:
            Messagebox.show_info("Info", "Task is running!")
            return
        self.thread_pool.submit(self._upload_worker)

    def _reset_state(self):
        self.selected_files = []
        self.text_input.delete(1.0, END)
        self.upload_btn.config(state=DISABLED)
        self.text_input.edit_modified(False)

    def on_delete_file(self, event):
        if self.is_working:
            return

        selected_item = self.file_tree.selection()
        if not selected_item:
            return

        tree_item = selected_item[0]
        item = self.file_tree.item(tree_item)

        if item["text"].startswith("📅"):
            return

        file_path = item["values"][0] if item["values"] else os.path.join(self.current_dir, item["text"])

        if file_path == self.current_playing_video:
            self.safe_stop_video()

        self.root.after(200, lambda: self.thread_pool.submit(self._delete_worker, file_path, tree_item))

    def _delete_worker(self, file_path, tree_item):
        max_retry = 3
        deleted = False

        for retry in range(max_retry):
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    self.log_queue.put(f"Deleted: {os.path.basename(file_path)}")
                    self.root.after(0, lambda: [
                        self.file_tree.delete(tree_item),
                        self.switch_preview("text"),
                        self._load_files_thread(),
                        self.write_log("delete", filename=os.path.basename(file_path))
                    ])
                    deleted = True
                    break
            except PermissionError:
                self.log_queue.put(f"Delete retry {retry + 1}/{max_retry}: File is in use")
                time.sleep(0.1)

        if not deleted:
            self.log_queue.put("Delete failed: File is in use by another program")

    def on_file_click(self, event):
        selected_item = self.file_tree.selection()
        if not selected_item or self.file_tree.item(selected_item)["text"].startswith("📅"):
            return

        item = self.file_tree.item(selected_item)
        file_path = item["values"][0] if item["values"] else os.path.join(self.current_dir, item["text"])
        suffix = os.path.splitext(file_path)[-1].lower()

        if suffix not in self.FILE_TYPES["direct_open"]:
            self._preview_file(file_path)

    def on_file_double_click(self, event):
        selected_item = self.file_tree.selection()
        if not selected_item or self.file_tree.item(selected_item)["text"].startswith("📅"):
            return

        item = self.file_tree.item(selected_item)
        file_path = item["values"][0] if item["values"] else os.path.join(self.current_dir, item["text"])
        suffix = os.path.splitext(file_path)[-1].lower()

        if suffix in self.FILE_TYPES["video"]:
            return
        elif suffix in self.FILE_TYPES["direct_open"]:
            self.thread_pool.submit(self._open_file, file_path)
        else:
            self._preview_file(file_path)

    def _open_file(self, file_path):
        if os.path.exists(file_path):
            os.startfile(file_path)
            self.log_queue.put(f"File opened: {os.path.basename(file_path)}")

    def _preview_file(self, file_path):
        if not os.path.isfile(file_path):
            return

        suffix = os.path.splitext(file_path)[-1].lower()

        if suffix in self.FILE_TYPES["text"]:
            self.switch_preview("text")
            self.thread_pool.submit(self._preview_text, file_path)
        elif suffix in self.FILE_TYPES["image"]:
            self.switch_preview("media")
            self.thread_pool.submit(self._preview_image, file_path)
        elif suffix in self.FILE_TYPES["video"]:
            if self.current_playing_video != file_path or not self.is_playing_video:
                self.safe_stop_video()
                self.switch_preview("media")
                self._play_video(file_path)

    def switch_preview(self, preview_type):
        self.safe_stop_video()

        self.preview_text.pack_forget()
        self.preview_scrollbar.pack_forget()

        self.preview_text.config(state=NORMAL)
        self.preview_text.delete(1.0, END)
        self.preview_text.config(state=DISABLED)

        if preview_type == "text":
            self.preview_text.pack(fill=BOTH, expand=YES, side=LEFT)
            self.preview_scrollbar.pack(side=RIGHT, fill=Y)
        elif preview_type == "media":
            self.preview_media_label = ttkbs.Label(self.preview_frame)
            self.preview_media_label.pack(fill=BOTH, expand=YES, side=LEFT)

    def _preview_text(self, file_path):
        content = None
        encodings = ["utf-8", "gbk", "gb2312", "utf-8-sig"]

        for encoding in encodings:
            try:
                with open(file_path, "r", encoding=encoding) as f:
                    lines = f.readlines()
                    if file_path.lower().endswith(".log"):
                        lines = reversed(lines)
                        formatted_lines = []
                        for line in lines:
                            line = line.strip()
                            if line:
                                formatted_lines.append(f"{self._format_log_line(line)}\n\n")
                        content = "".join(formatted_lines)
                    else:
                        content = "".join(lines)
                break
            except:
                continue

        self.root.after(0, lambda: self._update_text_preview(content))

    def _update_text_preview(self, content):
        self.preview_text.config(state=NORMAL)
        self.preview_text.delete(1.0, END)
        if content:
            self.preview_text.insert(1.0, content)
        self.preview_text.config(state=DISABLED)

    def _format_log_line(self, log_line):
        parts = re.split(r',\s*', log_line)
        formatted = []

        for part in parts:
            if ':' in part:
                key, value = part.split(':', 1)
                formatted.append(f"{key.strip()}: {value.strip()}")
            else:
                formatted.append(part.strip())

        return "\n".join(formatted)

    def _preview_image(self, file_path):
        img = Image.open(file_path)

        self.preview_frame.update_idletasks()
        max_w = max(self.preview_frame.winfo_width() - 20, 480)
        max_h = max(self.preview_frame.winfo_height() - 10, 480)

        scale = min(max_w / img.width, max_h / img.height, 1)
        new_size = (int(img.width * scale), int(img.height * scale))

        img = img.resize(new_size, Image.Resampling.LANCZOS)

        self.root.after(0, lambda: self._update_image_preview(img))

    def _update_image_preview(self, img):
        img_tk = ImageTk.PhotoImage(img)
        self.preview_media_label.config(image=img_tk)
        self.preview_media_label.image = img_tk

    def safe_stop_video(self):
        with self.video_lock:
            self.stop_video_flag = True
            self.is_playing_video = False

        if hasattr(self, 'preview_media_label') and self.preview_media_label:
            try:
                self.preview_media_label.config(image='')
                self.preview_media_label.destroy()
            except:
                pass
            self.preview_media_label = None

        cv2.destroyAllWindows()
        cv2.waitKey(1)
        self.current_playing_video = None
        self.video_play_locked = False

    def _play_video(self, file_path):
        self.stop_video_flag = False
        self.current_playing_video = file_path
        self.video_play_locked = True

        video_thread = threading.Thread(target=self._video_worker, args=(file_path,), daemon=True)
        video_thread.start()

    def _get_scaled_frame_size(self, frame_width, frame_height, max_width, max_height):
        scale_ratio = min(max_width / frame_width, max_height / frame_height, 1.0)

        new_width = int(frame_width * scale_ratio)
        new_height = int(frame_height * scale_ratio)

        return new_width, new_height

    def _video_worker(self, file_path):
        with self.video_lock:
            self.is_playing_video = True

        self.log_queue.put(f"Playing: {os.path.basename(file_path)}")

        cap = cv2.VideoCapture(file_path)
        if not cap.isOpened():
            self.log_queue.put(f"Failed to open video: {os.path.basename(file_path)}")
            self.safe_stop_video()
            return

        video_width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        video_height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))

        fps = cap.get(cv2.CAP_PROP_FPS)
        frame_delay = 1000.0 / max(fps, 15)

        self.preview_frame.update_idletasks()
        max_display_width = self.preview_frame.winfo_width() - 20
        max_display_height = self.preview_frame.winfo_height() - 10

        display_width, display_height = self._get_scaled_frame_size(
            video_width, video_height, max_display_width, max_display_height
        )

        last_frame_time = time.time() * 1000

        while not self.stop_video_flag and self.is_playing_video:
            if self.current_playing_video != file_path:
                break

            ret, frame = cap.read()
            if not ret:
                break

            current_time = time.time() * 1000
            elapsed = current_time - last_frame_time
            if elapsed < frame_delay:
                time.sleep((frame_delay - elapsed) / 1000)
            last_frame_time = current_time

            frame = cv2.resize(frame, (display_width, display_height), interpolation=cv2.INTER_AREA)
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(frame)
            img_tk = ImageTk.PhotoImage(img)

            self.root.after(0, lambda img=img_tk: self._update_video_frame(img))

        cap.release()
        cv2.destroyAllWindows()

        with self.video_lock:
            self.is_playing_video = False

        if self.current_playing_video == file_path:
            self.video_play_locked = False

        self.log_queue.put(f"Stopped: {os.path.basename(file_path)}")

    def _update_video_frame(self, img_tk):
        if self.preview_media_label and not self.stop_video_flag:
            self.preview_media_label.config(image=img_tk)
            self.preview_media_label.image = img_tk

    def __del__(self):
        self.safe_stop_video()
        if hasattr(self, 'thread_pool'):
            self.thread_pool.shutdown(wait=False)


if __name__ == "__main__":
    app = ttkbs.Window(themename="flatly")
    gui = QuanLog(app)
    app.mainloop()
