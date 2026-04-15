import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
import pyperclip
import threading
import time
import json
import os
from datetime import datetime
import hashlib
import keyboard
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

# 设置外观模式和颜色主题
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class ClipboardManager:
    def __init__(self):
        if HAS_DND:
            self.root = TkinterDnD.Tk()
            ctk.set_appearance_mode("System")
            ctk.set_default_color_theme("blue")
        else:
            self.root = ctk.CTk()

        self.config_file = "config.json"
        self.config = {}
        self.load_config()

        self.root.title("剪切板管理器")
        self.root.geometry(self.config.get('window_geometry', '700x750'))
        self.always_on_top = True
        self.root.attributes('-topmost', self.always_on_top)
        
        self.clipboard_history = []
        self.max_history = 300
        self.data_file = "clipboard_history.json"
        self.current_clipboard = ""
        self.monitoring = False
        self.monitor_thread = None
        self.hotkey_listening = False
        self.hotkey_thread = None
        self.is_processing_paste = False
        self.quick_paste_mode = False
        self.last_pasted_index = -1
        self.create_widgets()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.bind("<Map>", self._on_window_map, add="+")
        self.apply_theme()

    def load_config(self):
        defaults = {
            'show_window': 'ctrl+alt+c',
            'quick_paste': 'f8',
            'max_history': 100
        }
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                loaded_config = json.load(f)
                loaded_config.pop('paste_next', None)
                loaded_config.pop('sequential_paste', None)
                self.config = {**defaults, **loaded_config}
        except (FileNotFoundError, json.JSONDecodeError):
            self.config = defaults
            self.save_config()
        self.max_history = int(self.config.get('max_history', 100))

    def save_config(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置错误: {e}")

    def create_widgets(self):
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(3, weight=1) # Tabview row

        # === 主界面容器 ===
        self.main_ui_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.main_ui_frame.grid(row=0, column=0, rowspan=6, sticky="nsew")
        self.main_ui_frame.grid_columnconfigure(0, weight=1)
        self.main_ui_frame.grid_rowconfigure(3, weight=1)

        # 1. 顶部控制区 (Buttons + Search)
        self.top_frame = ctk.CTkFrame(self.main_ui_frame, fg_color="transparent")
        self.top_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        self.top_frame.grid_columnconfigure(1, weight=1) # Search bar expands

        # 按钮组
        btn_frame = ctk.CTkFrame(self.top_frame, fg_color="transparent")
        btn_frame.grid(row=0, column=0, sticky="w")
        
        self.quick_paste_btn = ctk.CTkButton(btn_frame, text="⚡ 开启连贴", command=self.toggle_quick_paste_mode, width=100)
        self.quick_paste_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        delete_selected_btn = ctk.CTkButton(btn_frame, text="🗑️ 删除", command=self.delete_selected, width=80, fg_color="#D32F2F", hover_color="#B71C1C")
        delete_selected_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        clear_btn = ctk.CTkButton(btn_frame, text="🧹 清空", command=self.clear_history_prompt, width=80, fg_color="#E64A19", hover_color="#D84315")
        clear_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # 导入文本文件按钮
        import_btn = ctk.CTkButton(btn_frame, text="📁 导入", command=self.import_text_file, width=80, fg_color="#4CAF50", hover_color="#388E3C")
        import_btn.pack(side=tk.LEFT, padx=(0, 5))
        
                # 更显眼的设置按钮，使用紫色系提升可见性
        self.settings_btn = ctk.CTkButton(
            btn_frame,
            text="⚙️ 设置",
            command=self.open_settings_window,
            width=90,
            fg_color="#6A1B9A",  # 深紫色
            hover_color="#8E24AA",
        )
        self.settings_btn.pack(side=tk.LEFT, padx=(0, 5))
        # 为设置窗口添加快捷键 Ctrl+,（逗号）
        self.root.bind("<Control-comma>", lambda e: self.open_settings_window())
        

        # 迷你模式按钮
        mini_mode_btn = ctk.CTkButton(btn_frame, text="📱 迷你", command=self.enable_mini_mode, width=60, fg_color="#00897B", hover_color="#00695C")
        mini_mode_btn.pack(side=tk.LEFT, padx=(0, 5))

        # 搜索栏
        search_frame = ctk.CTkFrame(self.top_frame, fg_color="transparent")
        search_frame.grid(row=0, column=1, sticky="ew", padx=(10, 0))
        
        self.search_var = tk.StringVar()
        self.search_var.trace("w", lambda name, index, mode, sv=self.search_var: self.on_search_change())
        
        search_entry = ctk.CTkEntry(search_frame, textvariable=self.search_var, placeholder_text="🔍 搜索历史记录...", width=150)
        search_entry.pack(fill=tk.X, expand=True)

        # 2. 列表区 (Tabview)
        self.tabview = ctk.CTkTabview(self.main_ui_frame)
        self.tabview.grid(row=3, column=0, sticky="nsew", padx=10, pady=5)

        self.tabview.add("历史记录")
        self.tabview.add("已粘贴")

        self.tabview.tab("历史记录").grid_columnconfigure(0, weight=1)
        self.tabview.tab("历史记录").grid_rowconfigure(0, weight=1)
        self.tabview.tab("已粘贴").grid_columnconfigure(0, weight=1)
        self.tabview.tab("已粘贴").grid_rowconfigure(0, weight=1)

        self.history_list = ctk.CTkScrollableFrame(self.tabview.tab("历史记录"), fg_color="transparent")
        self.history_list.grid(row=0, column=0, sticky="nsew")
        self.history_list.grid_columnconfigure(0, weight=1)

        self.pasted_list = ctk.CTkScrollableFrame(self.tabview.tab("已粘贴"), fg_color="transparent")
        self.pasted_list.grid(row=0, column=0, sticky="nsew")
        self.pasted_list.grid_columnconfigure(0, weight=1)

        # 卡片行存储：{index: frame_widget}
        self._card_frames = {}
        self.selected_index = -1  # 当前选中的 clipboard_history 索引

        self.root.bind("<Control-Up>", lambda e: self.move_selected_items("up"))
        self.root.bind("<Control-Down>", lambda e: self.move_selected_items("down"))

        # 3. 详细内容区
        self.detail_frame = ctk.CTkFrame(self.main_ui_frame)
        self.detail_frame.grid(row=4, column=0, sticky="ew", padx=10, pady=5)
        self.detail_frame.grid_columnconfigure(0, weight=1)
        
        detail_header = ctk.CTkFrame(self.detail_frame, fg_color="transparent")
        detail_header.pack(fill="x", padx=10, pady=(5, 0))
        
        detail_label = ctk.CTkLabel(detail_header, text="📄 详细内容", font=("Arial", 12, "bold"))
        detail_label.pack(side="left")

        self.detail_text = ctk.CTkTextbox(self.detail_frame, height=100, wrap="word", state="disabled")
        self.detail_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 4. 状态栏
        self.status_var = tk.StringVar(value="正在初始化...")
        self.status_bar = ctk.CTkLabel(self.main_ui_frame, textvariable=self.status_var, anchor='w', height=28, fg_color=("gray90", "gray20"), padx=10)
        self.status_bar.grid(row=5, column=0, sticky="ew", padx=0, pady=0)

        self.root.bind("<space>", self.copy_selected_on_space)
        self.root.bind("<Return>", self.on_item_double_click)

        # 拖拽导入 txt 文件
        if HAS_DND:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.on_file_drop)
        # === 迷你模式界面 (默认隐藏) ===
        self.mini_ui_frame = ctk.CTkFrame(self.root, corner_radius=0)
        # 不立即 grid，切换时再 grid
        self.mini_ui_frame.grid_columnconfigure(0, weight=1)
        self.mini_ui_frame.grid_rowconfigure(0, weight=1)

        self.mini_content_label = ctk.CTkLabel(self.mini_ui_frame, text="无内容", anchor="w", padx=10, cursor="hand2")
        self.mini_content_label.grid(row=0, column=0, sticky="ew", padx=(5, 5))
        
        # 绑定点击标签复制功能
        self.mini_content_label.bind("<Button-1>", lambda e: self.copy_latest_in_mini())

        mini_btn_frame = ctk.CTkFrame(self.mini_ui_frame, fg_color="transparent")
        mini_btn_frame.grid(row=0, column=1, sticky="e", padx=5)

        # 按钮组：粘贴 | 最新 | 抓取 | 返回
        self.mini_paste_btn = ctk.CTkButton(mini_btn_frame, text="📋 粘贴", width=60, command=self.paste_from_mini, fg_color="#F57C00", hover_color="#E65100")
        self.mini_paste_btn.pack(side="left", padx=2)

        self.mini_top_btn = ctk.CTkButton(mini_btn_frame, text="�  重置", width=60, command=self.copy_latest_in_mini)
        self.mini_top_btn.pack(side="left", padx=2)

        # 改为“抓取”按钮，模拟 Ctrl+C
        self.mini_capture_btn = ctk.CTkButton(mini_btn_frame, text="✂️ 抓取", width=60, command=self.capture_selection_from_mini)
        self.mini_capture_btn.pack(side="left", padx=2)
        
        ctk.CTkButton(mini_btn_frame, text="🔙 返回", width=60, command=self.disable_mini_mode).pack(side="left", padx=2)

        # 拖拽移动窗口 (迷你模式下)
        self.mini_ui_frame.bind("<ButtonPress-1>", self.start_move)
        self.mini_ui_frame.bind("<ButtonRelease-1>", self.stop_move)
        self.mini_ui_frame.bind("<B1-Motion>", self.do_move)
        # Label 绑定 Button-1 按下记录位置，Button-1 释放时如果移动距离小则视为点击，否则视为拖拽。
        self.mini_content_label.bind("<ButtonPress-1>", self.start_move_or_click)
        self.mini_content_label.bind("<ButtonRelease-1>", self.stop_move_or_click)
        self.mini_content_label.bind("<B1-Motion>", self.do_move)

    def start_move(self, event):
        self.x = event.x
        self.y = event.y

    def stop_move(self, event):
        self.x = None
        self.y = None

    def start_move_or_click(self, event):
        self.x = event.x
        self.y = event.y
        self.click_start_time = time.time()

    def stop_move_or_click(self, event):
        # 如果移动距离很小且时间很短，视为点击
        if self.x is not None and abs(event.x - self.x) < 3 and (time.time() - self.click_start_time) < 0.3:
            self.copy_latest_in_mini()
        self.x = None
        self.y = None

    def do_move(self, event):
        if self.x is None: return
        deltax = event.x - self.x
        deltay = event.y - self.y
        x = self.root.winfo_x() + deltax
        y = self.root.winfo_y() + deltay
        self.root.geometry(f"+{x}+{y}")

    def enable_mini_mode(self):
        self.previous_geometry = self.root.geometry()
        self.main_ui_frame.grid_forget()
        self.mini_ui_frame.grid(row=0, column=0, sticky="nsew")
        
        self.root.geometry("520x60")
        self.root.resizable(False, False)
        self.update_mini_label()

    def disable_mini_mode(self):
        self.mini_ui_frame.grid_forget()
        self.main_ui_frame.grid(row=0, column=0, rowspan=6, sticky="nsew")
        
        if hasattr(self, 'previous_geometry'):
            self.root.geometry(self.previous_geometry)
        else:
            self.root.geometry("700x750")
        self.root.resizable(True, True)

    def capture_selection_from_mini(self):
        """隐藏窗口，模拟 Ctrl+C，然后恢复窗口"""
        self.root.withdraw()
        self.root.update()
        time.sleep(0.2) # 等待焦点切换
        try:
            keyboard.send('ctrl+c')
            time.sleep(0.1)
            # 按钮反馈
            self.mini_capture_btn.configure(text="✅ 已抓", fg_color="#2E7D32")
            self.root.after(1000, lambda: self.mini_capture_btn.configure(text="✂️ 抓取", fg_color=["#3B8ED0", "#1F6AA5"]))
        except Exception as e:
            print(f"Capture failed: {e}")
            self.mini_capture_btn.configure(text="❌ 失败", fg_color="#C62828")
            self.root.after(1000, lambda: self.mini_capture_btn.configure(text="✂️ 抓取", fg_color=["#3B8ED0", "#1F6AA5"]))
        finally:
            self.restore_window()

    def paste_from_mini(self):
        """隐藏窗口，执行连贴逻辑 (同 Ctrl+V)，然后恢复窗口"""
        self.root.withdraw()
        self.root.update()
        time.sleep(0.2) # 等待窗口隐藏和焦点切换
        try:
            # 直接调用连贴的核心逻辑
            self.on_ctrl_v_pressed()
        except Exception as e:
            print(f"Paste failed: {e}")
        finally:
            # 给一点时间让粘贴动作完成，再恢复窗口
            # on_ctrl_v_pressed 内部是异步处理后续逻辑的，所以这里只需等待按键发送
            time.sleep(0.2)
            self.restore_window()

    def copy_latest_in_mini(self):
        if self.clipboard_history:
            # 重置为从第一条开始粘贴（从旧到新的顺序）
            content = self.clipboard_history[0]['content']
            try:
                pyperclip.copy(content)
                self.current_clipboard = content
                # 重置索引，从第一条开始
                self.last_pasted_index = -1  # 设为-1，下次会从0开始
                
                # Label 反馈
                self.mini_content_label.configure(text="✅ 已重置为第一条!")
                self.root.after(1000, lambda: self.update_mini_label())
                
                # 注意：这里不要立即调用 prepare_first_unpasted_for_paste
                # 否则剪切板会被覆盖为“下一条”，导致用户无法粘贴刚才选中的“最新项”
                # 等用户粘贴了最新项后，_process_paste_after_action 会自动准备下一条
                
            except Exception as e:
                print(f"Copy failed: {e}")

    def update_mini_label(self):
        if self.clipboard_history:
            total = len(self.clipboard_history)
            unpasted = sum(1 for item in self.clipboard_history if not item.get('pasted', False))
            
            # 获取当前准备粘贴的内容（即 last_pasted_index - 1，如果刚重置则是最新项）
            # 逻辑上，mini mode 显示的应该是“当前剪切板里的内容”或者“即将粘贴的内容”
            # 这里我们显示当前剪切板内容的预览
            
            current_content = pyperclip.paste().strip().replace('\n', ' ')
            if len(current_content) > 15: 
                current_content = current_content[:15] + "..."
                
            display_text = f"[{unpasted}/{total}] {current_content}"
            self.mini_content_label.configure(text=display_text)
        else:
            self.mini_content_label.configure(text="无历史记录")

    def _make_card(self, parent, index, item, is_next):
        """创建单条卡片行"""
        is_pasted = item.get('pasted', False)

        if is_next:
            bg = ("#cce5ff", "#1a3a5c")  # 蓝色高亮：当前准备粘贴
        elif self.selected_index == index:
            bg = ("#d0e8ff", "#1f4060")
        else:
            bg = ("#f5f5f5", "#2b2b2b") if index % 2 == 0 else ("#ffffff", "#333333")

        frame = ctk.CTkFrame(parent, fg_color=bg, corner_radius=4)
        frame.grid(row=index, column=0, sticky="ew", padx=4, pady=2)
        frame.grid_columnconfigure(1, weight=1)

        # 序号标签
        num_label = ctk.CTkLabel(frame, text=f"{index + 1}", width=30,
                                  font=("Microsoft YaHei UI", 10),
                                  text_color=("gray40", "gray60"))
        num_label.grid(row=0, column=0, padx=(6, 2), pady=4)

        # 内容预览
        content_preview = item['content'].strip().replace('\r\n', ' ↵ ').replace('\n', ' ↵ ').replace('\r', ' ↵ ')
        if not content_preview:
            content_preview = "<空>"
        preview_text = content_preview[:200]

        content_label = ctk.CTkLabel(frame, text=preview_text, anchor="w",
                                      font=("Microsoft YaHei UI", 10),
                                      wraplength=0)
        content_label.grid(row=0, column=1, sticky="ew", padx=(2, 6), pady=4)

        # 当前准备粘贴标记
        if is_next:
            next_label = ctk.CTkLabel(frame, text="▶", width=20,
                                       font=("Arial", 12), text_color="#1565C0")
            next_label.grid(row=0, column=2, padx=(0, 4))

        # 绑定事件
        for widget in (frame, num_label, content_label):
            widget.bind("<Button-1>", lambda e, i=index: self._on_card_click(i))
            widget.bind("<Double-Button-1>", lambda e, i=index: self._on_card_double_click(i))
            widget.bind("<Button-3>", lambda e, i=index: self._on_card_right_click(e, i))

        return frame

    def _on_card_click(self, index):
        self.selected_index = index
        self._update_card_highlights()
        # 显示详细内容
        try:
            content = self.clipboard_history[index]['content']
            self.detail_text.configure(state="normal")
            self.detail_text.delete(1.0, tk.END)
            self.detail_text.insert(1.0, content)
            self.detail_text.configure(state="disabled")
        except IndexError:
            pass

    def _on_card_double_click(self, index):
        self.selected_index = index
        self.copy_selected_item()
        # 视觉反馈：短暂变绿
        if index in self._card_frames and self._card_frames[index].winfo_exists():
            self._card_frames[index].configure(fg_color=("#b9f6ca", "#1b5e20"))
            self.root.after(400, self._update_card_highlights)

    def _on_card_right_click(self, event, index):
        self.selected_index = index
        self._update_card_highlights()
        is_pasted = self.clipboard_history[index].get('pasted', False)
        menu = tk.Menu(self.root, tearoff=0)
        if not is_pasted:
            menu.add_command(label="📋 复制 (设为下一个粘贴项)", command=self.copy_selected_item)
            menu.add_command(label="✏️ 编辑", command=self.edit_selected_item)
            menu.add_separator()
            menu.add_command(label="⬆️ 上移 (Ctrl+Up)", command=lambda: self.move_selected_items("up"))
            menu.add_command(label="⬇️ 下移 (Ctrl+Down)", command=lambda: self.move_selected_items("down"))
        else:
            menu.add_command(label="↩️ 移回历史记录", command=self.mark_as_unpasted)
            menu.add_command(label="📋 重新复制 (设为下一个粘贴项)", command=self.copy_selected_item)
            menu.add_command(label="✏️ 编辑", command=self.edit_selected_item)
        menu.add_separator()
        menu.add_command(label="🗑️ 删除所选", command=self.delete_selected)
        menu.tk_popup(event.x_root, event.y_root)

    def edit_selected_item(self):
        if self.selected_index < 0 or self.selected_index >= len(self.clipboard_history):
            return
        index = self.selected_index
        original = self.clipboard_history[index]['content']

        dialog = ctk.CTkToplevel(self.root)
        dialog.title("编辑条目")
        dialog.geometry(self.config.get('edit_geometry', '500x350'))
        dialog.transient(self.root)
        dialog.grab_set()

        textbox = ctk.CTkTextbox(dialog, wrap="word", font=("Microsoft YaHei UI", 11))
        textbox.pack(fill="both", expand=True, padx=10, pady=(10, 5))
        textbox.insert("1.0", original)
        textbox.focus_set()

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=(0, 10))

        def _save_edit_geo_and_close(destroy_fn):
            self.config['edit_geometry'] = dialog.geometry()
            self.save_config()
            destroy_fn()

        def on_save():
            new_content = textbox.get("1.0", "end-1c")
            self.config['edit_geometry'] = dialog.geometry()
            if new_content != original:
                new_hash = hashlib.md5(new_content.encode('utf-8')).hexdigest()
                self.clipboard_history[index]['content'] = new_content
                self.clipboard_history[index]['hash'] = new_hash
                self.clipboard_history[index]['type'] = self.detect_content_type(new_content)
                self.refresh_all_trees()
                self.save_history()
                self.status_var.set("条目已更新。")
            self.save_config()
            dialog.destroy()

        ctk.CTkButton(btn_frame, text="💾 保存", command=on_save, width=100).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="❌ 取消", command=lambda: _save_edit_geo_and_close(dialog.destroy), width=80,
                      fg_color="transparent", border_width=1, text_color=("gray10", "gray90")).pack(side="left", padx=10)

    def monitor_clipboard(self):
        while self.monitoring:
            try:
                time.sleep(0.5)
                new_content = pyperclip.paste()
                if new_content and new_content != self.current_clipboard:
                    new_hash = hashlib.md5(new_content.encode('utf-8')).hexdigest()
                    if not any(item.get('hash') == new_hash for item in self.clipboard_history):
                        self.current_clipboard = new_content
                        self.root.after(0, self.add_to_history, new_content)
            except Exception:
                time.sleep(1)

    def apply_theme(self, theme_name=None):
        if theme_name is None:
            theme_name = self.config.get('theme', 'light')
        theme_name = theme_name.lower()
        
        if theme_name == 'light':
            ctk.set_appearance_mode('Light')
            self.root.attributes('-alpha', 1.0)
        elif theme_name == 'glass':
            ctk.set_appearance_mode('Light')
            self.root.attributes('-alpha', 0.9)
        else:
            ctk.set_appearance_mode('Light')
            self.root.attributes('-alpha', 1.0)
            
        self.config['theme'] = theme_name
        self.save_config()

    def open_settings_window(self):
        settings_win = ctk.CTkToplevel(self.root)
        settings_win.title("设置")
        settings_win.geometry(self.config.get('settings_geometry', '500x500'))
        settings_win.transient(self.root)
        settings_win.grab_set()
        
        main_frame = ctk.CTkFrame(settings_win)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 快捷键设置
        hotkey_frame = ctk.CTkFrame(main_frame)
        hotkey_frame.pack(fill="x", pady=5, padx=5)
        
        ctk.CTkLabel(hotkey_frame, text="⌨️ 快捷键设置", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)

        show_hotkey_var = tk.StringVar(value=self.config.get('show_window', ''))
        quick_paste_hotkey_var = tk.StringVar(value=self.config.get('quick_paste', ''))

        grid_frame = ctk.CTkFrame(hotkey_frame, fg_color="transparent")
        grid_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(grid_frame, text="快速连贴 (自动粘贴并准备):").grid(row=0, column=0, sticky='w', pady=5)
        ctk.CTkEntry(grid_frame, textvariable=quick_paste_hotkey_var).grid(row=0, column=1, sticky='ew', padx=10)
        
        ctk.CTkLabel(grid_frame, text="显示/隐藏窗口:").grid(row=1, column=0, sticky='w', pady=5)
        ctk.CTkEntry(grid_frame, textvariable=show_hotkey_var).grid(row=1, column=1, sticky='ew', padx=10)
        
        ctk.CTkLabel(grid_frame, text="(提示: 顺序粘贴已集成至 Ctrl+V)", text_color="gray").grid(row=2, column=0, columnspan=2, sticky='w', pady=(5, 0))
        grid_frame.columnconfigure(1, weight=1)

        # 主题设置
        theme_frame = ctk.CTkFrame(main_frame)
        theme_frame.pack(fill="x", pady=5, padx=5)
        ctk.CTkLabel(theme_frame, text="🎨 主题设置", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)
        
        theme_var = tk.StringVar(value=self.config.get('theme', 'light'))
        theme_option_menu = ctk.CTkOptionMenu(
            theme_frame,
            values=["light", "glass"],
            variable=theme_var,
            width=200
        )
        theme_option_menu.pack(padx=10, pady=5, anchor="w")

        # 常规设置
        general_frame = ctk.CTkFrame(main_frame)
        general_frame.pack(fill="x", pady=5, padx=5)
        ctk.CTkLabel(general_frame, text="🛠️ 常规设置", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)
        
        max_history_var = tk.IntVar(value=self.max_history)
        gen_grid = ctk.CTkFrame(general_frame, fg_color="transparent")
        gen_grid.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(gen_grid, text="最大历史条数:").pack(side="left")
        ctk.CTkEntry(gen_grid, textvariable=max_history_var, width=60).pack(side="left", padx=10)

        # 常用操作
        action_frame = ctk.CTkFrame(main_frame)
        action_frame.pack(fill="x", pady=5, padx=5)
        ctk.CTkLabel(action_frame, text="⚡ 常用操作", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=5)
        
        self.settings_monitor_btn = ctk.CTkButton(action_frame,
                                               text=f"切换监控状态 (当前: {'开' if self.monitoring else '关'})",
                                               command=self.toggle_monitoring)
        self.settings_monitor_btn.pack(side="left", padx=10, pady=10)
        
        self.settings_topmost_btn = ctk.CTkButton(action_frame,
                                               text=f"切换窗口置顶 (当前: {'开' if self.always_on_top else '关'})",
                                               command=self.toggle_topmost)
        self.settings_topmost_btn.pack(side="left", padx=10, pady=10)

        def apply_and_save_settings():
            self.config['show_window'] = show_hotkey_var.get().lower().strip()
            self.config['quick_paste'] = quick_paste_hotkey_var.get().lower().strip()
            try:
                self.config['max_history'] = int(max_history_var.get())
            except:
                pass
            self.max_history = self.config['max_history']
            self.config['settings_geometry'] = settings_win.geometry()
            self.apply_theme(theme_var.get())
            self.save_config()
            self.reregister_hotkeys()
            self.toggle_quick_paste_mode(update_ui_only=True)
            self.trim_history()
            settings_win.destroy()

        def cancel_settings():
            self.config['settings_geometry'] = settings_win.geometry()
            self.save_config()
            settings_win.destroy()

        save_cancel_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        save_cancel_frame.pack(pady=(20, 0))
        ctk.CTkButton(save_cancel_frame, text="💾 保存并关闭", command=apply_and_save_settings).pack(side="left", padx=10)
        ctk.CTkButton(save_cancel_frame, text="❌ 取消", command=cancel_settings, fg_color="transparent", border_width=1, text_color=("gray10", "gray90")).pack(side="left", padx=10)

    def trim_history(self):
        if len(self.clipboard_history) > self.max_history:
            self.clipboard_history = self.clipboard_history[-self.max_history:]
            self.refresh_all_trees()
            self.save_history()
            self.status_var.set(f"历史记录已根据新限制 ({self.max_history}条) 裁剪。")

    def toggle_monitoring(self):
        self.stop_monitoring() if self.monitoring else self.start_monitoring()
        if hasattr(self, 'settings_monitor_btn') and self.settings_monitor_btn.winfo_exists():
            self.settings_monitor_btn.configure(text=f"切换监控状态 (当前: {'开' if self.monitoring else '关'})")

    def toggle_topmost(self):
        self.always_on_top = not self.always_on_top
        self.root.attributes('-topmost', self.always_on_top)
        self.status_var.set("窗口已置顶" if self.always_on_top else "窗口置顶已取消")
        if hasattr(self, 'settings_topmost_btn') and self.settings_topmost_btn.winfo_exists():
            self.settings_topmost_btn.configure(text=f"切换窗口置顶 (当前: {'开' if self.always_on_top else '关'})")

    def _on_window_map(self, event):
        self.root.unbind("<Map>")
        self.root.update_idletasks()
        self.load_history()
        self.start_monitoring()
        self.start_hotkey_listener()



    def load_history(self):
        if not os.path.exists(self.data_file):
            self.refresh_all_trees()
            return
        try:
            with open(self.data_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                for item in data:
                    item.setdefault('pasted', False)
                    item.setdefault('saved', False)
                self.clipboard_history = data
        except Exception as e:
            print(f"加载历史记录错误: {e}")
            self.clipboard_history = []
        finally:
            self.trim_history()
            self.refresh_all_trees()
            self.prepare_first_unpasted_for_paste()
            if not next((item for item in self.clipboard_history if not item.get('pasted', False)), None):
                if self.clipboard_history:
                    pyperclip.copy(self.clipboard_history[-1]['content'])
                    self.status_var.set("历史记录已加载，无未粘贴项。")
                else:
                    self.status_var.set("历史记录为空。")

    def on_ctrl_v_pressed(self):
        if self.is_processing_paste:
            return
        self.is_processing_paste = True
        try:
            pasted_content_before_action = pyperclip.paste()
            if not pasted_content_before_action:
                return
            keyboard.remove_hotkey('ctrl+v')
            keyboard.send('ctrl+v')
            time.sleep(0.05)
            threading.Thread(target=self._process_paste_after_action, args=(pasted_content_before_action,)).start()
        finally:
            keyboard.add_hotkey('ctrl+v', self.on_ctrl_v_pressed, suppress=True)
            self.is_processing_paste = False

    def _process_paste_after_action(self, pasted_content):
        found_index = -1
        for i in range(len(self.clipboard_history)):
            item = self.clipboard_history[i]
            if item['content'] == pasted_content and not item.get('pasted', False):
                item['pasted'] = True
                found_index = i
                break

        if found_index == -1:
            for i in range(len(self.clipboard_history)):
                if self.clipboard_history[i]['content'] == pasted_content:
                    found_index = i
                    break

        if found_index != -1:
            self.last_pasted_index = found_index

        self.save_history()
        # 延迟一次性刷新，合并 refresh + prepare + select_next 避免多次重建闪烁
        self.root.after(100, self._post_paste_refresh)

    def reregister_hotkeys(self):
        try:
            keyboard.unhook_all()
            hotkeys = self.config
            keyboard.add_hotkey('ctrl+v', self.on_ctrl_v_pressed, suppress=True)

            if hotkeys.get('show_window'):
                keyboard.add_hotkey(hotkeys['show_window'], lambda: self.root.after(0, self.toggle_window_visibility))

            if self.quick_paste_mode:
                quick_paste_key = hotkeys.get('quick_paste')
                if quick_paste_key:
                    keyboard.add_hotkey(quick_paste_key, lambda: self.root.after(0, self.perform_quick_paste))

            if not self.quick_paste_mode:
                self.status_var.set("快捷键已更新。")
        except Exception as e:
            error_msg = f"注册快捷键失败: {e}. 请检查格式。"
            self.status_var.set(error_msg)
            messagebox.showerror("快捷键错误", error_msg)

    def toggle_quick_paste_mode(self, update_ui_only=False):
        if not update_ui_only:
            self.quick_paste_mode = not self.quick_paste_mode
        quick_paste_key = self.config.get('quick_paste', 'f8').upper()
        if self.quick_paste_mode:
            self.quick_paste_btn.configure(text=f"⚡ 关闭连贴 ({quick_paste_key})", fg_color="#F57C00", hover_color="#E65100")
            self.status_var.set(f"快速连贴已开启！按 {quick_paste_key} 自动粘贴。")
        else:
            self.quick_paste_btn.configure(text="⚡ 开启连贴", fg_color=["#3B8ED0", "#1F6AA5"], hover_color=["#36719F", "#144870"])
            self.status_var.set("快速连贴已关闭。") if not update_ui_only else None
        if not update_ui_only:
            self.reregister_hotkeys()

    def start_monitoring(self):
        if self.monitoring: return
        self.monitoring = True
        self.current_clipboard = pyperclip.paste()
        self.monitor_thread = threading.Thread(target=self.monitor_clipboard, daemon=True)
        self.monitor_thread.start()
        self.status_var.set("监控中...")

    def stop_monitoring(self):
        self.monitoring = False
        self.status_var.set("已停止监控")

    def perform_quick_paste(self):
        try:
            self.on_ctrl_v_pressed()
        except Exception as e:
            print(f"快速连贴执行错误: {e}")
            self.status_var.set("快速连贴出错！")

    def delete_selected(self):
        if self.selected_index < 0:
            self.status_var.set("请先选择要删除的项目")
            return
        if messagebox.askyesno("确认删除", "确定要删除所选的项目吗？"):
            del self.clipboard_history[self.selected_index]
            self.selected_index = -1
            self.last_pasted_index = -1
            self.refresh_all_trees()
            self.save_history()
            self.status_var.set("已删除 1 个项目")
            self.prepare_first_unpasted_for_paste()

    def copy_selected_on_space(self, event=None):
        if self.selected_index < 0:
            self.status_var.set("请先选择一个项目再按空格键复制")
            return
        self.copy_selected_item()

    def _post_paste_refresh(self):
        """粘贴后：移除已粘贴卡片，其余向上移，再把它追加到已粘贴列表"""
        pasted_index = self.last_pasted_index

        # 1. 销毁历史记录列表中对应的卡片
        if pasted_index in self._card_frames:
            frame = self._card_frames.pop(pasted_index)
            if frame.winfo_exists():
                frame.destroy()

        # 2. 重新排列历史记录列表中剩余卡片的 grid row（向上填补空位）
        history_cards = sorted(
            [(i, f) for i, f in self._card_frames.items()
             if f.winfo_exists() and not self.clipboard_history[i].get('pasted', False)],
            key=lambda x: x[0]
        )
        for row, (i, frame) in enumerate(history_cards):
            frame.grid(row=row, column=0, sticky="ew", padx=4, pady=2)

        # 3. 把已粘贴条目追加到已粘贴列表末尾
        if 0 <= pasted_index < len(self.clipboard_history):
            item = self.clipboard_history[pasted_index]
            pasted_row = len([f for f in self._card_frames.values()
                               if f.winfo_exists() and
                               self.clipboard_history.index(item) != pasted_index])
            # 直接数已粘贴列表现有行数
            pasted_row = len(self.pasted_list.winfo_children())
            card = self._make_card(self.pasted_list, pasted_index, item, is_next=False)
            card.grid(row=pasted_row, column=0, sticky="ew", padx=4, pady=2)
            self._card_frames[pasted_index] = card

        # 4. 准备下一条 + 更新高亮
        self.prepare_first_unpasted_for_paste()
        self._update_card_highlights()

        # 5. 更新状态栏
        unpasted = sum(1 for it in self.clipboard_history if not it.get('pasted', False))
        total = len(self.clipboard_history)
        self.status_var.set(f"就绪 | 历史: {unpasted} | 已粘贴: {total - unpasted}")

    def select_next_unpasted_item(self):
        next_index = self.last_pasted_index + 1
        for i in range(next_index, len(self.clipboard_history)):
            if not self.clipboard_history[i].get('pasted', False):
                self.selected_index = i
                break

    def prepare_first_unpasted_for_paste(self, new_item_content=None):
        if not self.clipboard_history:
            return

        # 找下一条未粘贴的条目
        if new_item_content is not None:
            # 新内容加入时，从头找第一条未粘贴的
            start = 0
        else:
            start = self.last_pasted_index + 1

        # 从 start 往后找第一条未粘贴的
        next_item = None
        next_index = -1
        for i in range(start, len(self.clipboard_history)):
            if not self.clipboard_history[i].get('pasted', False):
                next_item = self.clipboard_history[i]
                next_index = i
                break

        if next_item is None:
            # 历史记录里没有未粘贴的条目，停止
            self.status_var.set("✅ 所有条目已粘贴完毕，历史记录已清空待用。")
            self.mini_content_label.configure(text="✅ 所有条目已粘贴完毕")
            self.root.after(1500, self.update_mini_label)
            return

        pyperclip.copy(next_item['content'])
        self.current_clipboard = next_item['content']

        preview = next_item['content'].strip().replace('\n', ' ')[:30]
        self.status_var.set(f"已准备下一条 [{next_index + 1}/{len(self.clipboard_history)}]: {preview}...")
        self.update_mini_label()

    def add_to_history(self, content):
        content_hash = hashlib.md5(content.encode('utf-8')).hexdigest()
        item = {'content': content, 'timestamp': datetime.now().isoformat(), 'type': self.detect_content_type(content),
                'hash': content_hash, 'pasted': False, 'saved': False}
        self.clipboard_history.append(item)

        self.trim_history()
        self.refresh_all_trees(scroll_to_end=True)
        self.save_history()

    def on_search_change(self):
        self.refresh_all_trees(scroll_to_end=False)

    def refresh_all_trees(self, scroll_to_end=False):
        # 清空两个列表
        for widget in self.history_list.winfo_children():
            widget.destroy()
        for widget in self.pasted_list.winfo_children():
            widget.destroy()
        self._card_frames = {}

        search_term = self.search_var.get().lower().strip()

        # 找出下一条准备粘贴的索引
        next_index = -1
        for i in range(self.last_pasted_index + 1, len(self.clipboard_history)):
            if not self.clipboard_history[i].get('pasted', False):
                next_index = i
                break
        self._next_index = next_index

        history_row = 0
        pasted_row = 0
        history_count = pasted_count = 0

        for i, item in enumerate(self.clipboard_history):
            if search_term and search_term not in item['content'].lower():
                continue

            is_next = (i == next_index)
            if item.get('pasted', False):
                card = self._make_card(self.pasted_list, i, item, is_next=False)
                card.grid(row=pasted_row, column=0, sticky="ew", padx=4, pady=2)
                pasted_row += 1
                pasted_count += 1
            else:
                card = self._make_card(self.history_list, i, item, is_next=is_next)
                card.grid(row=history_row, column=0, sticky="ew", padx=4, pady=2)
                history_row += 1
                history_count += 1
            self._card_frames[i] = card

        self.status_var.set(f"就绪 | 历史: {history_count} | 已粘贴: {pasted_count}")
        self.update_mini_label()

    def _update_card_highlights(self):
        """只更新卡片背景色，不重建，避免闪烁"""
        next_index = -1
        for i in range(self.last_pasted_index + 1, len(self.clipboard_history)):
            if not self.clipboard_history[i].get('pasted', False):
                next_index = i
                break
        self._next_index = next_index

        for i, frame in self._card_frames.items():
            if not frame.winfo_exists():
                continue
            if i == next_index:
                bg = ("#cce5ff", "#1a3a5c")
            elif i == self.selected_index:
                bg = ("#d0e8ff", "#1f4060")
            else:
                bg = ("#f5f5f5", "#2b2b2b") if i % 2 == 0 else ("#ffffff", "#333333")
            frame.configure(fg_color=bg)

        unpasted = sum(1 for item in self.clipboard_history if not item.get('pasted', False))
        total = len(self.clipboard_history)
        self.status_var.set(f"就绪 | 历史: {unpasted} | 已粘贴: {total - unpasted}")
        self.update_mini_label()

    def mark_as_unpasted(self):
        if self.selected_index < 0:
            self.status_var.set("请先选择一个已粘贴的项目")
            return
        index = self.selected_index
        if 0 <= index < len(self.clipboard_history):
            self.clipboard_history[index]['pasted'] = False
            self.clipboard_history[index]['saved'] = False
            # 让 prepare 从这条开始找，确保移回的条目不被跳过
            self.last_pasted_index = index - 1
            self.refresh_all_trees()
            self.save_history()
            self.status_var.set("已将条目移回历史记录")
            self.prepare_first_unpasted_for_paste()

    def on_item_double_click(self, event=None):
        if self.selected_index >= 0:
            self.copy_selected_item()

    def show_item_detail(self, event=None):
        pass  # 卡片点击已在 _on_card_click 中处理

    def get_active_selection(self):
        """返回 (is_pasted, [selected_index]) 兼容旧接口"""
        if self.selected_index < 0 or self.selected_index >= len(self.clipboard_history):
            return None, []
        return None, [str(self.selected_index)]

    def copy_selected_item(self):
        if self.selected_index < 0 or self.selected_index >= len(self.clipboard_history):
            messagebox.showwarning("提示", "请先选择一个项目再进行复制。")
            return
        try:
            content = self.clipboard_history[self.selected_index]['content']
            pyperclip.copy(content)
            self.current_clipboard = content
            self.last_pasted_index = self.selected_index - 1  # 下次 prepare 会从 selected+1 开始找
            self.status_var.set(f"已手动选择: {content[:30]}... 按 Ctrl+V 粘贴。")
        except IndexError:
            self.status_var.set("选择的项目无效")

    def move_selected_items(self, direction):
        i = self.selected_index
        if i < 0 or i >= len(self.clipboard_history):
            self.status_var.set("请先选择一个项目以调整顺序。")
            return
        if self.clipboard_history[i].get('pasted', False):
            self.status_var.set("请在'历史记录'列表中选择项目以调整顺序。")
            return

        if direction == "up" and i > 0:
            self.clipboard_history[i], self.clipboard_history[i - 1] = self.clipboard_history[i - 1], self.clipboard_history[i]
            self.selected_index = i - 1
        elif direction == "down" and i < len(self.clipboard_history) - 1:
            self.clipboard_history[i], self.clipboard_history[i + 1] = self.clipboard_history[i + 1], self.clipboard_history[i]
            self.selected_index = i + 1
        else:
            return

        self.refresh_all_trees()
        self.save_history()
        self.prepare_first_unpasted_for_paste()
        self.status_var.set(f"已向{'上' if direction == 'up' else '下'}移动。")

    def show_context_menu(self, event):
        pass  # 右键菜单已移至卡片的 _on_card_right_click

    def clear_history_prompt(self):
        res = messagebox.askquestion("清空历史记录", "要清空所有记录吗？\n('是'清空所有, '否'仅清空已粘贴)",
                                     type=messagebox.YESNOCANCEL)
        if res == 'yes':
            if messagebox.askyesno("确认", "确定要清空所有记录吗？此操作无法撤销。"): self.clipboard_history.clear()
        elif res == 'no':
            if messagebox.askyesno("确认", "确定要清空所有已粘贴的记录吗？"): self.clipboard_history = [i for i in self.clipboard_history if not i.get('pasted', False)]
        else:
            return
        self.refresh_all_trees()
        pyperclip.copy('')
        self.detail_text.configure(state="normal")
        self.detail_text.delete(1.0, tk.END)
        self.detail_text.configure(state="disabled")
        self.save_history()

    def detect_content_type(self, content):
        return "🌐 URL" if content.startswith(('http://', 'https://')) else "🔢 数字" if content.isnumeric() else "📝 多行文本" if '\n' in content or '\r' in content else "📄 文本"

    def restore_window(self):
        try:
            self.root.deiconify()
            self.root.lift()
            if self.always_on_top:
                self.root.attributes('-topmost', True)
        except Exception as e:
            print(f"恢复窗口错误: {e}")

    def save_history(self):
        try:
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(self.clipboard_history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存历史记录错误: {e}")

    def toggle_window_visibility(self):
        if self.root.winfo_viewable():
            self.root.withdraw()
        else:
            self.restore_window()

    def start_hotkey_listener(self):
        if not self.hotkey_listening:
            self.hotkey_listening = True
            self.reregister_hotkeys()
            self.hotkey_thread = threading.Thread(target=keyboard.wait, daemon=True)
            self.hotkey_thread.start()

    def stop_hotkey_listener(self):
        self.hotkey_listening = False
        keyboard.unhook_all()

    def on_closing(self):
        self.config['window_geometry'] = self.root.geometry()
        self.save_config()
        self.stop_monitoring()
        self.stop_hotkey_listener()
        self.save_history()
        self.root.destroy()

    def auto_save_pasted_history(self):
        items_to_save = [item for item in self.clipboard_history if item.get('pasted', False) and not item.get('saved', False)]
        if not items_to_save:
            return

        try:
            now = datetime.now()
            date_folder = now.strftime("%Y%m%d")
            os.makedirs(date_folder, exist_ok=True)

            time_str = now.strftime("%H%M%S")
            filename = f"pasted_history_{time_str}.txt"
            filepath = os.path.join(date_folder, filename)

            processed_contents = []
            for item in items_to_save:
                lines = item['content'].strip().splitlines()
                non_empty_lines = [line for line in lines if line.strip()]
                processed_contents.append("\n".join(non_empty_lines))
            content_to_save = "\n\n".join(processed_contents)

            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content_to_save)

            for item in items_to_save:
                item['saved'] = True
            
            self.save_history()
            self.status_var.set(f"记录已自动保存到 {filepath}")

        except Exception as e:
            self.status_var.set(f"自动保存失败: {e}")
            messagebox.showerror("错误", f"自动保存文件时出错: {e}")

    def on_file_drop(self, event):
        """处理拖拽进来的文件"""
        # tkinterdnd2 返回的路径可能带花括号（多文件或含空格时）
        raw = event.data.strip()
        # 解析多个文件路径
        if raw.startswith('{'):
            paths = [p.strip('{}') for p in raw.split('} {')]
        else:
            paths = raw.split()

        txt_paths = [p for p in paths if p.lower().endswith('.txt')]
        if not txt_paths:
            messagebox.showwarning("提示", "仅支持拖入 .txt 文本文件")
            return

        for path in txt_paths:
            self._import_from_path(path)

    def import_text_file(self):
        """导入文本文件到剪切板历史记录"""
        file_path = filedialog.askopenfilename(
            title="选择要导入的文本文件",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            initialdir=os.getcwd()
        )
        if file_path:
            self._import_from_path(file_path)

    def _import_from_path(self, file_path):
        """从指定路径读取并导入文本文件（供按钮和拖拽共用）"""
        try:
            content = None
            for enc in ('utf-8', 'gbk', 'latin-1'):
                try:
                    with open(file_path, 'r', encoding=enc) as f:
                        content = f.read()
                    break
                except UnicodeDecodeError:
                    continue

            if not content or not content.strip():
                messagebox.showwarning("警告", "文件内容为空！")
                return

            # 用自定义对话框让用户选择导入选项
            result = {}

            dialog = ctk.CTkToplevel(self.root)
            dialog.title("导入选项")
            dialog.geometry("400x210")
            dialog.transient(self.root)
            dialog.grab_set()
            dialog.resizable(False, False)

            ctk.CTkLabel(dialog, text="导入方式", font=("Arial", 13, "bold")).pack(pady=(15, 5))

            split_mode_var = tk.StringVar(value="separator")
            skip_var = tk.BooleanVar(value=False)

            opt_frame = ctk.CTkFrame(dialog, fg_color="transparent")
            opt_frame.pack(fill="x", padx=20, pady=5)

            ctk.CTkRadioButton(opt_frame, text="按空行分割（条目内不能含空行）", variable=split_mode_var, value="blank_line").pack(anchor="w", pady=3)
            ctk.CTkRadioButton(opt_frame, text="按 ---- 行分割（条目内可含空行）", variable=split_mode_var, value="separator").pack(anchor="w", pady=3)

            ctk.CTkCheckBox(opt_frame, text="跳过已存在的重复内容", variable=skip_var).pack(anchor="w", pady=(8, 3))

            btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
            btn_frame.pack(pady=(10, 0))

            def on_confirm():
                result['split_mode'] = split_mode_var.get()
                result['skip'] = skip_var.get()
                result['confirmed'] = True
                dialog.destroy()

            def on_cancel():
                result['confirmed'] = False
                dialog.destroy()

            ctk.CTkButton(btn_frame, text="✅ 确认导入", command=on_confirm, width=120).pack(side="left", padx=10)
            ctk.CTkButton(btn_frame, text="❌ 取消", command=on_cancel, width=80,
                          fg_color="transparent", border_width=1, text_color=("gray10", "gray90")).pack(side="left", padx=10)

            dialog.wait_window()

            if not result.get('confirmed'):
                return

            split_mode = result['split_mode']
            skip_dup = result['skip']

            imported_count = skipped_count = 0
            if split_mode == "blank_line":
                sections = content.split('\n\n')
            else:  # separator
                import re
                sections = re.split(r'\n-{4,}\n', content)

            for section in sections:
                section = section.strip()
                if not section:
                    continue
                h = hashlib.md5(section.encode('utf-8')).hexdigest()
                is_dup = any(item.get('hash') == h for item in self.clipboard_history)
                if is_dup and skip_dup:
                    skipped_count += 1
                    continue
                self.clipboard_history.append({
                    'content': section, 'timestamp': datetime.now().isoformat(),
                    'type': self.detect_content_type(section), 'hash': h,
                    'pasted': False, 'saved': False
                })
                imported_count += 1

            if imported_count > 0:
                self.trim_history()
                self.refresh_all_trees(scroll_to_end=True)
                self.save_history()
                # 导入后从第一条开始准备，不重置 last_pasted_index 的方式
                self.last_pasted_index = -1
                self.prepare_first_unpasted_for_paste()
                filename = os.path.basename(file_path)
                msg = f"成功导入 {imported_count} 条"
                if skipped_count:
                    msg += f"，跳过重复 {skipped_count} 条"
                self.status_var.set(f"从 {filename} {msg}")
            else:
                self.status_var.set("没有新内容导入" + (f"（重复条目 {skipped_count} 条均已跳过）" if skipped_count else ""))

        except Exception as e:
            error_msg = f"导入文件时出错: {str(e)}"
            self.status_var.set(error_msg)
            messagebox.showerror("导入错误", error_msg)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    try:
        import pyperclip, keyboard
    except ImportError as e:
        messagebox.showerror("缺少依赖", f"错误: 缺少必要的库 -> {e.name}\n请运行: pip install {e.name}")
        exit()
    app = ClipboardManager()
    app.run()