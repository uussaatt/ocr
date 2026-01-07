import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyperclip
import threading
import time
import json
import os
from datetime import datetime
import hashlib
import keyboard

# 设置外观模式和颜色主题
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class ClipboardManager:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("剪切板管理器")
        self.root.geometry("700x750")
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
        self.quick_paste_mode = False
        self.last_pasted_index = -1
        self.config_file = "config.json"
        self.config = {}

        self.load_config()
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

        self.history_tree = self._setup_treeview(self.tabview.tab("历史记录"))
        self.pasted_tree = self._setup_treeview(self.tabview.tab("已粘贴"))

        self.history_tree.bind("<Control-Up>", lambda e: self.move_selected_items("up"))
        self.history_tree.bind("<Control-Down>", lambda e: self.move_selected_items("down"))

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

    def _setup_treeview(self, parent_frame):
        style = ttk.Style()
        style.theme_use("default")
        
        # 适配暗色/亮色模式
        is_dark = ctk.get_appearance_mode() == "Dark"
        bg_color = "#2b2b2b" if is_dark else "#ffffff"
        fg_color = "white" if is_dark else "black"
        field_bg = "#2b2b2b" if is_dark else "#ffffff"
        header_bg = "#565b5e" if is_dark else "#e1e1e1"
        header_fg = "white" if is_dark else "black"
        selected_bg = "#1f538d"

        style.configure("Treeview",
                        background=bg_color,
                        foreground=fg_color,
                        rowheight=30,
                        fieldbackground=field_bg,
                        borderwidth=0,
                        font=("Microsoft YaHei UI", 10))
        
        style.map('Treeview', background=[('selected', selected_bg)], foreground=[('selected', 'white')])
        
        style.configure("Treeview.Heading",
                        background=header_bg,
                        foreground=header_fg,
                        relief="flat",
                        font=("Microsoft YaHei UI", 10, "bold"))
        
        style.map("Treeview.Heading",
                  background=[('active', '#3484F0')])

        columns = ("时间", "类型", "内容预览")
        tree = ttk.Treeview(parent_frame, columns=columns, show="headings", selectmode="extended")
        
        tree.heading("时间", text="⏰ 时间")
        tree.heading("类型", text="🏷️ 类型")
        tree.heading("内容预览", text="📝 内容预览")
        
        # 优化列宽设置，让内容预览列自动填充剩余空间
        tree.column("时间", width=120, minwidth=100, stretch=False)
        tree.column("类型", width=80, minwidth=60, stretch=False)
        tree.column("内容预览", width=400, minwidth=250, stretch=True)
        
        scrollbar = ctk.CTkScrollbar(parent_frame, orientation="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        tree.bind("<Double-1>", self.on_item_double_click)
        tree.bind("<ButtonRelease-1>", self.show_item_detail)
        tree.bind("<Button-3>", self.show_context_menu)
        
        return tree

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
            theme_name = self.config.get('theme', 'dark')
        theme_name = theme_name.lower()
        
        if theme_name == 'dark':
            ctk.set_appearance_mode('Dark')
            self.root.attributes('-alpha', 1.0)
        elif theme_name == 'light':
            ctk.set_appearance_mode('Light')
            self.root.attributes('-alpha', 1.0)
        elif theme_name == 'glass':
            ctk.set_appearance_mode('Light')
            self.root.attributes('-alpha', 0.9)
        else:
            ctk.set_appearance_mode('Dark')
            self.root.attributes('-alpha', 1.0)
            
        self.config['theme'] = theme_name
        self.save_config()

    def open_settings_window(self):
        settings_win = ctk.CTkToplevel(self.root)
        settings_win.title("设置")
        settings_win.geometry("500x500")
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
        
        theme_var = tk.StringVar(value=self.config.get('theme', 'dark'))
        theme_option_menu = ctk.CTkOptionMenu(
            theme_frame,
            values=["dark", "light", "glass"],
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
            
            self.apply_theme(theme_var.get())
            
            self.save_config()
            self.reregister_hotkeys()
            self.toggle_quick_paste_mode(update_ui_only=True)
            self.trim_history()
            settings_win.destroy()

        save_cancel_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        save_cancel_frame.pack(pady=(20, 0))
        ctk.CTkButton(save_cancel_frame, text="💾 保存并关闭", command=apply_and_save_settings).pack(side="left", padx=10)
        ctk.CTkButton(save_cancel_frame, text="❌ 取消", command=settings_win.destroy, fg_color="transparent", border_width=1, text_color=("gray10", "gray90")).pack(side="left", padx=10)

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
        item_marked = False
        # 重新实现逆序查找，确保匹配到最新的那条（如果内容重复）
        for i in range(len(self.clipboard_history) - 1, -1, -1):
            item = self.clipboard_history[i]
            if item['content'] == pasted_content:
                if not item.get('pasted', False):
                    item['pasted'] = True
                    item_marked = True
                self.last_pasted_index = i
                break

        if item_marked or True: # 即使没有标记新粘贴（重贴旧的），也要准备下一条
            self.root.after(0, self.refresh_all_trees)
            self.save_history()
            self.root.after(10, self.prepare_first_unpasted_for_paste)
            # 粘贴成功后选中下一个待粘贴的条目
            self.root.after(50, self.select_next_unpasted_item)

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
        active_tree, _ = self.get_active_selection()
        selected_iids = active_tree.selection()
        if not selected_iids:
            self.status_var.set("请先选择要删除的项目")
            return
        try:
            items_to_delete_hashes = {self.clipboard_history[int(iid)]['hash'] for iid in selected_iids}
        except (ValueError, IndexError):
            self.status_var.set("选择项中包含无效的项目ID")
            return

        if messagebox.askyesno("确认删除", f"确定要删除所选的 {len(selected_iids)} 个项目吗？"):
            self.clipboard_history = [item for item in self.clipboard_history if item['hash'] not in items_to_delete_hashes]
            self.last_pasted_index = -1 # 删除后重置索引，防止错位
            self.refresh_all_trees()
            self.save_history()
            self.status_var.set(f"已删除 {len(selected_iids)} 个项目")
            self.prepare_first_unpasted_for_paste()

    def copy_selected_on_space(self, event=None):
        active_tree, selected_iids = self.get_active_selection()
        if not selected_iids:
            self.status_var.set("请先选择一个项目再按空格键复制")
            return
        self.copy_selected_item()

    def select_next_unpasted_item(self):
        """选中下一个待粘贴的条目"""
        try:
            # 确保当前在历史记录标签页
            self.tabview.set("历史记录")
            
            # 计算下一个要粘贴的索引（从旧到新的顺序）
            next_index = self.last_pasted_index + 1
            if next_index >= len(self.clipboard_history):
                next_index = 0
            
            # 在树视图中查找对应的项目
            for child in self.history_tree.get_children():
                if child == str(next_index):
                    # 清除之前的选择
                    self.history_tree.selection_remove(self.history_tree.selection())
                    # 选中新项目
                    self.history_tree.selection_set(child)
                    # 滚动到可见区域
                    self.history_tree.see(child)
                    # 显示详细内容
                    self.show_item_detail()
                    break
        except Exception as e:
            print(f"选中下一个条目时出错: {e}")

    def prepare_first_unpasted_for_paste(self, new_item_content=None):
        # 修改逻辑：从旧到新的顺序粘贴，配合显示顺序（最先复制的在上面）
        if not self.clipboard_history:
            return

        # 如果是新添加的内容，从第一个开始准备
        if new_item_content is not None:
            self.last_pasted_index = -1  # 重置为-1，下次会从0开始
            # 新添加内容时，准备第一条（最旧的）
            next_index = 0
        else:
            # 正常连贴流程，准备下一条（索引加1）
            next_index = self.last_pasted_index + 1
            
        is_finished_cycle = False
        # 循环：如果超出范围，回到开头（最旧的）
        if next_index >= len(self.clipboard_history):
            next_index = 0
            is_finished_cycle = True
        
        # 确保索引在有效范围内
        if next_index < 0:
            next_index = 0
            
        item = self.clipboard_history[next_index]
        
        pyperclip.copy(item['content'])
        self.current_clipboard = item['content']
        
        # 状态栏提示
        preview = item['content'].strip().replace('\n', ' ')[:30]
        if new_item_content is not None:
            self.status_var.set(f"新内容已添加，准备从第一条开始: {preview}...")
        else:
            self.status_var.set(f"已准备下一条: {preview}...")
        
        # 同步更新迷你模式的显示
        if is_finished_cycle and new_item_content is None:
             self.mini_content_label.configure(text="✅ 所有记录已粘贴完毕")
             # 延迟 1.5 秒后恢复显示内容预览，让用户看到提示
             self.root.after(1500, self.update_mini_label)
        else:
             self.update_mini_label()

    def add_to_history(self, content):
        content_hash = hashlib.md5(content.encode('utf-8')).hexdigest()
        item = {'content': content, 'timestamp': datetime.now().isoformat(), 'type': self.detect_content_type(content),
                'hash': content_hash, 'pasted': False, 'saved': False}
        self.clipboard_history.append(item)

        self.prepare_first_unpasted_for_paste(new_item_content=content)

        self.trim_history()
        self.refresh_all_trees(scroll_to_end=True)
        self.save_history()

    def on_search_change(self):
        self.refresh_all_trees(scroll_to_end=False)

    def refresh_all_trees(self, scroll_to_end=False):
        for tree in [self.history_tree, self.pasted_tree]:
            tree.delete(*tree.get_children())
        
        search_term = self.search_var.get().lower().strip()
        
        history_count = pasted_count = 0
        
        # 正序遍历，让最先复制的项显示在最上面
        for i in range(len(self.clipboard_history)):
            item = self.clipboard_history[i]
            # 搜索过滤
            if search_term and search_term not in item['content'].lower():
                continue

            ts = datetime.fromisoformat(item['timestamp']).strftime("%m-%d %H:%M:%S")
            content_preview = item['content'].strip()
            if not content_preview:
                content_preview = "<空>"
            else:
                content_preview = content_preview.replace('\r\n', ' ↵ ').replace('\n', ' ↵ ').replace('\r', ' ↵ ')
            
            preview = content_preview[:300]
            
            values = (ts, item['type'], preview)
            if item.get('pasted', False):
                self.pasted_tree.insert("", "end", iid=str(i), values=values)
                pasted_count += 1
            else:
                self.history_tree.insert("", "end", iid=str(i), values=values)
                history_count += 1
        
        # 更新状态栏统计
        self.status_var.set(f"就绪 | 历史: {history_count} | 已粘贴: {pasted_count}")
        
        # 同时更新迷你模式的标签（如果有新内容）
        self.update_mini_label()

        if scroll_to_end and self.history_tree.get_children():
            # 滚动到最后一个（最新的）条目
            last = self.history_tree.get_children()[-1]
            self.history_tree.see(last)
            self.history_tree.selection_set(last)

    def mark_as_unpasted(self):
        active_tree, selected_iids = self.get_active_selection()
        if not selected_iids or active_tree != self.pasted_tree:
            self.status_var.set("请在'已粘贴'列表中选择一个或多个项目")
            return

        try:
            count = 0
            for iid in selected_iids:
                index = int(iid)
                if 0 <= index < len(self.clipboard_history):
                    self.clipboard_history[index]['pasted'] = False
                    self.clipboard_history[index]['saved'] = False
                    count += 1

            if count > 0:
                self.refresh_all_trees()
                self.save_history()
                self.status_var.set(f"已将 {count} 个条目移回历史记录")
                self.prepare_first_unpasted_for_paste()

        except (ValueError, IndexError):
            self.status_var.set("错误：选择的项目无效")

    def on_item_double_click(self, event):
        active_tree, selected_iids = self.get_active_selection()
        if not selected_iids: return
        self.copy_selected_item()

    def show_item_detail(self, event=None):
        if event and event.widget.identify_region(event.x, event.y) == 'heading': return
        active_tree, selected_iids = self.get_active_selection()
        if not selected_iids: return
        sel = selected_iids[0]
        try:
            content = self.clipboard_history[int(sel)]['content']
            self.detail_text.configure(state="normal")
            self.detail_text.delete(1.0, tk.END)
            self.detail_text.insert(1.0, content)
            self.detail_text.configure(state="disabled")
        except (ValueError, IndexError):
            pass

    def get_active_selection(self):
        try:
            current_tab = self.tabview.get()
            active_tree = self.history_tree if current_tab == "历史记录" else self.pasted_tree
            selection = active_tree.selection()
            return active_tree, selection
        except Exception:
            return self.history_tree, ()

    def copy_selected_item(self):
        active_tree, selected_iids = self.get_active_selection()
        if not selected_iids:
            messagebox.showwarning("提示", "请先在主窗口选择一个项目再进行复制。")
            return

        iid = selected_iids[0]
        try:
            content = self.clipboard_history[int(iid)]['content']
            pyperclip.copy(content)
            self.current_clipboard = content
            # 手动选择时，更新 last_pasted_index，但不立即准备下一条
            # 这样用户粘贴这条后，会自动准备下一条
            self.last_pasted_index = int(iid)
            
            # 高亮显示当前选中的项目
            active_tree.selection_set(iid)
            active_tree.see(iid)
            
            self.status_var.set(f"已手动选择: {content[:30]}... 按 Ctrl+V 粘贴。")
        except (ValueError, IndexError):
            self.status_var.set("选择的项目无效")

    def move_selected_items(self, direction):
        active_tree, selected_iids = self.get_active_selection()
        if not selected_iids or active_tree != self.history_tree:
            self.status_var.set("请在'历史记录'列表中选择项目以调整顺序。")
            return

        try:
            indices = [int(iid) for iid in selected_iids]

            if direction == "up":
                indices.sort()
                for i in indices:
                    if i > 0:
                        self.clipboard_history[i], self.clipboard_history[i - 1] = self.clipboard_history[i - 1], self.clipboard_history[i]
            elif direction == "down":
                indices.sort(reverse=True)
                for i in indices:
                    if i < len(self.clipboard_history) - 1:
                        self.clipboard_history[i], self.clipboard_history[i + 1] = self.clipboard_history[i + 1], self.clipboard_history[i]

            offset = -1 if direction == "up" else 1
            new_iids_to_select = [str(i + offset) for i in indices]

            self.refresh_all_trees()
            self.save_history()

            for new_iid in new_iids_to_select:
                self.history_tree.selection_add(new_iid)
            if new_iids_to_select:
                self.history_tree.see(new_iids_to_select[0])

            self.prepare_first_unpasted_for_paste()
            self.status_var.set(f"已将 {len(indices)} 个项目向{'上' if direction == 'up' else '下'}移动。")

        except (ValueError, IndexError) as e:
            self.status_var.set(f"顺序调整失败: {e}")

    def show_context_menu(self, event):
        iid = event.widget.identify_row(event.y)
        if iid:
            if iid not in event.widget.selection():
                event.widget.selection_set(iid)

            menu = tk.Menu(self.root, tearoff=0)
            if event.widget == self.history_tree:
                menu.add_command(label="📋 复制 (设为下一个粘贴项)", command=self.copy_selected_item)
                menu.add_separator()
                menu.add_command(label="⬆️ 上移 (Ctrl+Up)", command=lambda: self.move_selected_items("up"))
                menu.add_command(label="⬇️ 下移 (Ctrl+Down)", command=lambda: self.move_selected_items("down"))
            else:
                menu.add_command(label="↩️ 移回历史记录", command=self.mark_as_unpasted)
                menu.add_command(label="📋 重新复制 (设为下一个粘贴项)", command=self.copy_selected_item)

            menu.add_separator()
            menu.add_command(label="🗑️ 删除所选", command=self.delete_selected)
            menu.tk_popup(event.x_root, event.y_root)

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

    def import_text_file(self):
        """导入文本文件到剪切板历史记录"""
        try:
            # 打开文件选择对话框
            file_path = filedialog.askopenfilename(
                title="选择要导入的文本文件",
                filetypes=[
                    ("文本文件", "*.txt"),
                    ("所有文件", "*.*")
                ],
                initialdir=os.getcwd()
            )
            
            if not file_path:
                return
            
            # 读取文件内容
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                # 如果UTF-8失败，尝试其他编码
                try:
                    with open(file_path, 'r', encoding='gbk') as f:
                        content = f.read()
                except UnicodeDecodeError:
                    with open(file_path, 'r', encoding='latin-1') as f:
                        content = f.read()
            
            if not content.strip():
                messagebox.showwarning("警告", "文件内容为空！")
                return
            
            # 询问导入方式
            import_choice = messagebox.askyesnocancel(
                "导入方式选择",
                "选择导入方式：\n\n"
                "是(Yes) - 按空行分割导入（空行分隔的内容作为独立条目）\n"
                "否(No) - 整个文件作为一个历史记录\n"
                "取消(Cancel) - 取消导入"
            )
            
            if import_choice is None:  # 用户点击取消
                return
            
            imported_count = 0
            
            if import_choice:  # 按空行分割导入
                # 按空行分割内容
                sections = content.split('\n\n')  # 双换行符分割
                
                # 反转顺序，从后往前导入
                for section in reversed(sections):
                    section = section.strip()
                    if section:  # 跳过空段落
                        # 检查是否已存在相同内容
                        section_hash = hashlib.md5(section.encode('utf-8')).hexdigest()
                        if not any(item.get('hash') == section_hash for item in self.clipboard_history):
                            item = {
                                'content': section,
                                'timestamp': datetime.now().isoformat(),
                                'type': self.detect_content_type(section),
                                'hash': section_hash,
                                'pasted': False,
                                'saved': False
                            }
                            self.clipboard_history.append(item)
                            imported_count += 1
            else:  # 整个文件作为一个记录
                content_hash = hashlib.md5(content.encode('utf-8')).hexdigest()
                if not any(item.get('hash') == content_hash for item in self.clipboard_history):
                    item = {
                        'content': content,
                        'timestamp': datetime.now().isoformat(),
                        'type': self.detect_content_type(content),
                        'hash': content_hash,
                        'pasted': False,
                        'saved': False
                    }
                    self.clipboard_history.append(item)
                    imported_count = 1
            
            if imported_count > 0:
                # 限制历史记录数量
                self.trim_history()
                
                # 刷新界面
                self.refresh_all_trees(scroll_to_end=True)
                
                # 保存历史记录
                self.save_history()
                
                # 准备第一个未粘贴的项目
                self.prepare_first_unpasted_for_paste()
                
                # 显示成功消息
                filename = os.path.basename(file_path)
                self.status_var.set(f"成功从 {filename} 导入 {imported_count} 条记录")
                messagebox.showinfo("导入成功", f"成功导入 {imported_count} 条记录！")
            else:
                messagebox.showinfo("导入完成", "没有新内容需要导入（可能已存在相同记录）")
                
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
