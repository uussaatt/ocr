import requests
import os
import base64
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, simpledialog, Menu, ttk
from pathlib import Path
from urllib.parse import quote_plus
from PIL import Image
import threading
import json
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.widgets import LassoSelector
from matplotlib.path import Path as MplPath
import re
import random
from matplotlib import font_manager

# 加载 .env 文件
env_path = Path(__file__).parent / '.env'
if env_path.exists():
    with open(env_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#') and '=' in line:
                key, value = line.split('=', 1)
                os.environ[key.strip()] = value.strip()

# 高精度识别的密钥
API_KEY = os.getenv("BAIDU_API_KEY", "")
SECRET_KEY = os.getenv("BAIDU_SECRET_KEY", "")

# 快速识别的密钥（如果没有配置，则使用高精度的密钥）
API_KEY_BASIC = os.getenv("BAIDU_API_KEY_BASIC", API_KEY)
SECRET_KEY_BASIC = os.getenv("BAIDU_SECRET_KEY_BASIC", SECRET_KEY)

# 通用识别的密钥（如果没有配置，则使用快速识别的密钥）
API_KEY_GENERAL = os.getenv("BAIDU_API_KEY_GENERAL", API_KEY_BASIC)
SECRET_KEY_GENERAL = os.getenv("BAIDU_SECRET_KEY_GENERAL", SECRET_KEY_BASIC)


# === 字体配置 (Windows 环境) ===
def configure_styles_force():
    plt.rcParams['axes.unicode_minus'] = False
    font_paths = [r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyh.ttf", r"C:\Windows\Fonts\simhei.ttf"]
    font_loaded = False
    for path in font_paths:
        if os.path.exists(path):
            try:
                font_manager.fontManager.addfont(path)
                font_name = font_manager.FontProperties(fname=path).get_name()
                plt.rcParams['font.sans-serif'] = [font_name]
                font_loaded = True
                break
            except:
                pass
    if not font_loaded:
        plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial Unicode MS']


configure_styles_force()


def get_access_token(use_basic=False, use_general=False):
    """
    使用 AK，SK 生成鉴权签名（Access Token）
    :param use_basic: 是否使用快速识别的密钥
    :param use_general: 是否使用通用识别的密钥
    :return: access_token，或是None(如果错误)
    """
    url = "https://aip.baidubce.com/oauth/2.0/token"
    
    if use_general:
        # 使用通用识别的密钥
        params = {"grant_type": "client_credentials", "client_id": API_KEY_GENERAL, "client_secret": SECRET_KEY_GENERAL}
    elif use_basic:
        # 使用快速识别的密钥
        params = {"grant_type": "client_credentials", "client_id": API_KEY_BASIC, "client_secret": SECRET_KEY_BASIC}
    else:
        # 使用高精度识别的密钥
        params = {"grant_type": "client_credentials", "client_id": API_KEY, "client_secret": SECRET_KEY}
    
    return str(requests.post(url, params=params).json().get("access_token"))


def get_file_content_as_base64(path, max_size=8192, max_file_size_mb=3.5):
    """将图片文件转换为 base64 编码，自动压缩大图片和大文件"""
    try:
        # 检查原始文件大小
        file_size = os.path.getsize(path)
        file_size_mb = file_size / (1024 * 1024)
        
        # 打开图片
        img = Image.open(path)
        width, height = img.size
        
        # 判断是否需要压缩（尺寸过大或文件过大）
        need_compress = (width > max_size or height > max_size or file_size_mb > max_file_size_mb)
        
        if need_compress:
            print(f"图片需要压缩: 尺寸({width}x{height}) 文件大小({file_size_mb:.1f}MB)")
            
            # 计算目标尺寸
            if width > max_size or height > max_size:
                # 按尺寸压缩
                scale = min(max_size / width, max_size / height)
                new_width = int(width * scale)
                new_height = int(height * scale)
            else:
                # 按文件大小压缩（保持尺寸，降低质量）
                new_width = width
                new_height = height
            
            # 压缩图片
            if new_width != width or new_height != height:
                img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                print(f"尺寸压缩: {width}x{height} → {new_width}x{new_height}")
            
            # 转换为字节流并调整质量
            import io
            img_byte_arr = io.BytesIO()
            
            # 根据文件大小动态调整质量
            quality = 85
            if file_size_mb > 10:
                quality = 60
            elif file_size_mb > 5:
                quality = 70
            elif file_size_mb > 3:
                quality = 80
            
            img.save(img_byte_arr, format='JPEG', quality=quality, optimize=True)
            compressed_data = img_byte_arr.getvalue()
            compressed_size_mb = len(compressed_data) / (1024 * 1024)
            
            print(f"压缩完成: {file_size_mb:.1f}MB → {compressed_size_mb:.1f}MB (质量:{quality})")
            
            # 如果压缩后仍然太大，进一步降低质量
            if compressed_size_mb > max_file_size_mb:
                for lower_quality in [50, 40, 30, 20]:
                    img_byte_arr = io.BytesIO()
                    img.save(img_byte_arr, format='JPEG', quality=lower_quality, optimize=True)
                    compressed_data = img_byte_arr.getvalue()
                    compressed_size_mb = len(compressed_data) / (1024 * 1024)
                    print(f"进一步压缩: 质量{lower_quality} → {compressed_size_mb:.1f}MB")
                    if compressed_size_mb <= max_file_size_mb:
                        break
            
            return base64.b64encode(compressed_data).decode("utf8")
        else:
            # 图片尺寸和文件大小都合适，直接读取
            print(f"图片无需压缩: 尺寸({width}x{height}) 文件大小({file_size_mb:.1f}MB)")
            with open(path, "rb") as f:
                return base64.b64encode(f.read()).decode("utf8")
    
    except Exception as e:
        print(f"处理图片时出错: {e}")
        # 如果出错，尝试使用原始方法
        try:
            with open(path, "rb") as f:
                return base64.b64encode(f.read()).decode("utf8")
        except:
            return None


def ocr_image(image_path):
    """对图片进行 OCR 识别（高精度版）"""
    url = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate?access_token=" + get_access_token()
    
    # 高精度识别使用较宽松的文件大小限制
    image_base64 = get_file_content_as_base64(image_path, max_size=8192, max_file_size_mb=3.8)
    
    if image_base64 is None:
        return {"error_msg": "图片处理失败", "error_code": -1}
    
    # 需要获取位置信息，所以不关闭 location
    payload = {
        'image': image_base64,
        'detect_direction': 'false',
        'paragraph': 'false',
        'probability': 'false',
        'char_probability': 'false',
        'multidirectional_recognize': 'false'
    }
    
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': 'application/json'
    }
    
    response = requests.post(url, headers=headers, data=payload)
    response.encoding = "utf-8"
    return response.json()


def ocr_image_basic(image_path):
    """对图片进行 OCR 识别（快速版 - accurate_basic）"""
    url = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic?access_token=" + get_access_token(use_basic=True)
    
    # 快速识别使用中等的文件大小限制
    image_base64 = get_file_content_as_base64(image_path, max_size=8100, max_file_size_mb=3.5)
    
    if image_base64 is None:
        return {"error_msg": "图片处理失败", "error_code": -1}
    
    # 使用字典格式的payload，和高精度识别保持一致
    payload = {
        'image': image_base64,
        'detect_direction': 'false',
        'paragraph': 'false',
        'probability': 'false',
        'multidirectional_recognize': 'false'
    }
    
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': 'application/json'
    }
    
    response = requests.post(url, headers=headers, data=payload)
    response.encoding = "utf-8"
    return response.json()


def ocr_image_general(image_path):
    """对图片进行 OCR 识别（通用版 - general）"""
    # 使用通用识别的密钥
    url = "https://aip.baidubce.com/rest/2.0/ocr/v1/general?access_token=" + get_access_token(use_general=True)
    
    # 通用识别使用较严格的文件大小限制
    image_base64 = get_file_content_as_base64(image_path, max_size=4096, max_file_size_mb=3.0)
    
    if image_base64 is None:
        return {"error_msg": "图片处理失败", "error_code": -1}
    
    # 通用识别的参数（按照你提供的代码格式）
    payload = {
        'image': image_base64,
        'detect_direction': 'false',
        'detect_language': 'false',
        'vertexes_location': 'false',
        'paragraph': 'false',
        'probability': 'false'
    }
    
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': 'application/json'
    }
    
    response = requests.post(url, headers=headers, data=payload)
    response.encoding = "utf-8"
    return response.json()



class DataStore:
    """统一数据存储管理器"""
    def __init__(self, filepath):
        self.filepath = filepath
        self.data = {
            'window_config': {},
            'stats': {},
            'history': [],
            'history_limit': 100,
            'size_limits': {},
            'font_config': {'font_size': 11},
            'popup_windows': {}
        }
        self.load()

    def load(self):
        if self.filepath.exists():
            try:
                with open(self.filepath, 'r', encoding='utf-8') as f:
                    saved = json.load(f)
                    # 深度合并或更新，这里简单更新顶层键
                    for k, v in saved.items():
                        self.data[k] = v
            except Exception as e:
                print(f"加载数据文件失败: {e}")

    def save(self):
        try:
            with open(self.filepath, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存数据文件失败: {e}")

    def get(self, key, default=None):
        return self.data.get(key, default)

    def set(self, key, value):
        self.data[key] = value
        self.save()

    def migrate_legacy_files(self, parent_dir):
        """从旧的分散文件迁移数据"""
        legacy_files = {
            'window_config': 'window_config.json',
            'stats': 'ocr_stats.json',
            'history': 'ocr_history.json',
            'history_limit': 'history_limit.json',
            'size_limits': 'size_limits.json',
            'font_config': 'font_config.json',
            'popup_windows': 'popup_windows.json'
        }
        
        migrated = False
        for key, filename in legacy_files.items():
            path = parent_dir / filename
            if path.exists():
                try:
                    with open(path, 'r', encoding='utf-8') as f:
                        content = json.load(f)
                        # 特殊处理 history_limit 格式
                        if key == 'history_limit' and isinstance(content, dict):
                            self.data[key] = content.get('limit', 100)
                        else:
                            self.data[key] = content
                    print(f"✓ 已迁移旧文件: {filename}")
                    migrated = True
                    
                    # 可选：重命名旧文件作为备份
                    # try:
                    #     path.rename(path.with_suffix('.json.bak'))
                    # except: pass
                except Exception as e:
                    print(f"迁移 {filename} 失败: {e}")
        
        if migrated:
            self.save()
            print("✓ 数据迁移完成，已保存到 ocr_data.json")


class OCRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OCR 文字识别 + 数据分类工具")
        

        # 数据存储初始化
        self.data_file = Path(__file__).parent / 'ocr_data.json'
        self.store = DataStore(self.data_file)
        
        # 如果数据文件不存在，尝试迁移旧数据
        if not self.data_file.exists():
            self.store.migrate_legacy_files(Path(__file__).parent)
        
        # 加载并应用窗口配置
        self.load_window_config()
        
        self.root.minsize(1200, 800)  # 设置最小尺寸，防止窗口过小
        
        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 统计数据
        self.stats = self.store.get('stats', {})
        
        # 历史记录
        self.history_limit = self.store.get('history_limit', 100)
        self.history_data = self.store.get('history', [])
        
        # 尺寸限制解锁状态
        self.size_limit_unlocked = False
        self.unlock_password = "000"  # 设置密码
        
        # 图片尺寸限制配置（可自定义）- 使用范围限制
        self.size_limits = {
            'accurate_min_width': 3500,    # 高精度最小宽度
            'accurate_min_height': 4000,   # 高精度最小高度
            'accurate_max_width': 15000,   # 高精度最大宽度
            'accurate_max_height': 15000,  # 高精度最大高度
            'basic_min_width': 0,          # 快速识别最小宽度
            'basic_min_height': 0,         # 快速识别最小高度
            'basic_max_width': 8100,       # 快速识别最大宽度
            'basic_max_height': 3000,      # 快速识别最大高度
            'general_min_width': 0,        # 通用识别最小宽度
            'general_min_height': 0,       # 通用识别最小高度
            'general_max_width': 8192,     # 通用识别最大宽度
            'general_max_height': 8192     # 通用识别最大高度
        }
        self.load_size_limits()
        
        # 数据分类相关属性
        self.current_font_size = 11  # 默认字号
        self.load_font_config()  # 加载保存的字号设置
        self.df = pd.DataFrame(columns=['Label', 'Y', 'X'])
        self.thresholds = []
        self.category_list = []
        self.marked_indices = set()
        self.custom_cat_names = {}
        self.drag_source_item = None
        self.enable_lasso_mode = tk.BooleanVar(value=False)
        self.color_cycle = ['#FF0000', '#00AA00', '#FF8C00', '#9400D3', '#0000FF', '#00CED1']
        self.lasso = None
        
        # 创建主界面
        self.setup_main_interface()
        
        # 启用拖放功能
        self._setup_drag_drop()

    def setup_main_interface(self):
        """设置主界面"""
        # 创建主标签页
        self.main_notebook = ttk.Notebook(self.root)
        self.main_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # OCR 标签页
        self.ocr_tab = tk.Frame(self.main_notebook)
        self.main_notebook.add(self.ocr_tab, text=" 🔍 OCR识别 ")
        
        # 数据分类标签页
        self.classifier_tab = tk.Frame(self.main_notebook)
        self.main_notebook.add(self.classifier_tab, text=" 📊 数据分类 ")
        
        # 设置 OCR 标签页
        self.setup_ocr_tab()
        
        # 设置数据分类标签页
        self.setup_classifier_tab()

    def setup_ocr_tab(self):
        """设置 OCR 标签页"""
        # ========== Ribbon 功能区 ==========
        ribbon_frame = tk.Frame(self.ocr_tab, bg="#f0f0f0", relief=tk.RAISED, bd=1)
        ribbon_frame.pack(fill=tk.X, padx=0, pady=0)
        
        # Ribbon 内容框架
        ribbon_content = tk.Frame(ribbon_frame, bg="#f0f0f0")
        ribbon_content.pack(fill=tk.X, padx=10, pady=8)
        
        # === 文件组 ===
        file_group = self._create_ribbon_group(ribbon_content, "文件")
        self.select_btn = self._create_ribbon_button(file_group, "📁\n选择图片", self.select_file, 
                                                      "#4CAF50", large=True)
        
        # === 识别组 ===
        ocr_group = self._create_ribbon_group(ribbon_content, "识别")
        self.ocr_btn = self._create_ribbon_button(ocr_group, "🔍\n高精度", self.perform_ocr, 
                                                   "#2196F3", state=tk.DISABLED)
        self.quick_ocr_btn = self._create_ribbon_button(ocr_group, "⚡\n快速", self.perform_quick_ocr, 
                                                         "#00BCD4", state=tk.DISABLED)
        self.general_ocr_btn = self._create_ribbon_button(ocr_group, "📄\n通用", self.perform_general_ocr, 
                                                           "#9C27B0", state=tk.DISABLED)
        
        # === 图片处理组 ===
        image_group = self._create_ribbon_group(ribbon_content, "图片处理")
        self.merge_btn = self._create_ribbon_button(image_group, "🖼️\n拼接", self.merge_images, "#FF9800")
        self.crop_merge_btn = self._create_ribbon_button(image_group, "✂️\n裁剪", self.crop_and_merge_direct, "#FF6F00")
        
        # === 结果操作组 ===
        result_group = self._create_ribbon_group(ribbon_content, "结果操作")
        self.copy_btn = self._create_ribbon_button(result_group, "📋\n复制", self.copy_text, 
                                                    "#607D8B", state=tk.DISABLED)
        self.add_zeros_btn = self._create_ribbon_button(result_group, "➕\n加|0|0", self.add_zeros_to_lines, 
                                                         "#9C27B0", state=tk.DISABLED)
        self.export_btn = self._create_ribbon_button(result_group, "💾\n导出", self.export_results, 
                                                      "#FF5722", state=tk.DISABLED)
        self.clear_btn = self._create_ribbon_button(result_group, "🗑️\n清空", self.clear_result, "#757575")
        
        # === 数据查看组 ===
        data_group = self._create_ribbon_group(ribbon_content, "数据")
        self.stats_btn = self._create_ribbon_button(data_group, "📊\n统计", self.show_stats, "#3F51B5")
        self.history_btn = self._create_ribbon_button(data_group, "📜\n历史", self.show_history, "#00897B")
        
        # === 设置组 ===
        settings_group = self._create_ribbon_group(ribbon_content, "设置")
        self.api_key_btn = self._create_ribbon_button(settings_group, "🔑\n密钥", self.show_api_key_settings, "#673AB7")
        self.unlock_btn = self._create_ribbon_button(settings_group, "🔓\n解锁", self.unlock_size_limit, "#E91E63")
        
        # 文件路径标签
        self.file_label = tk.Label(self.ocr_tab, text="未选择文件", fg="gray", wraplength=1350, bg="#fafafa", 
                                   pady=8, font=("Arial", 9))
        self.file_label.pack(fill=tk.X, padx=10, pady=(5, 0))
        
        # 进度条
        self.progress_frame = tk.Frame(self.ocr_tab)
        self.progress_frame.pack(fill=tk.X, padx=20, pady=5)
        
        self.progress_label = tk.Label(self.progress_frame, text="", fg="blue")
        self.progress_label.pack(side=tk.LEFT)
        
        # 提示信息
        acc_range = f"{self.size_limits['accurate_min_width']}~{self.size_limits['accurate_max_width']}x{self.size_limits['accurate_min_height']}~{self.size_limits['accurate_max_height']}"
        bas_range = f"{self.size_limits['basic_min_width']}~{self.size_limits['basic_max_width']}x{self.size_limits['basic_min_height']}~{self.size_limits['basic_max_height']}"
        gen_range = f"{self.size_limits['general_min_width']}~{self.size_limits['general_max_width']}x{self.size_limits['general_min_height']}~{self.size_limits['general_max_height']}"
        self.size_hint_label = tk.Label(self.progress_frame, text=f"💡 高精度({acc_range}) | 快速({bas_range}) | 通用({gen_range})", 
                fg="gray", font=("Arial", 9))
        self.size_hint_label.pack(side=tk.RIGHT, padx=10)
        
        # 结果显示区域
        result_label = tk.Label(self.ocr_tab, text="识别结果：", font=("Arial", 12, "bold"))
        result_label.pack(pady=(10, 5))
        
        self.result_text = scrolledtext.ScrolledText(self.ocr_tab, width=160, height=40, 
                                                      font=("Microsoft YaHei", 11))
        self.result_text.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
        
        # 添加右键菜单
        self.context_menu = tk.Menu(self.result_text, tearoff=0)
        self.context_menu.add_command(label="复制选中内容", command=self.copy_selected)
        self.context_menu.add_command(label="复制全部（文字+位置）", command=self.copy_all_text)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="全选", command=self.select_all)
        
        self.result_text.bind("<Button-3>", self.show_context_menu)
        
        self.image_paths = []  # 存储多个图片路径
        self.all_results = []  # 存储所有识别结果

    def setup_classifier_tab(self):
        """设置数据分类标签页"""
        # 左侧面板
        self.left_panel = tk.Frame(self.classifier_tab, width=420, bg="#f0f0f0")
        self.left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)
        
        # 右侧面板
        self.right_panel = tk.Frame(self.classifier_tab, bg="white")
        self.right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 右侧标签页
        self.classifier_notebook = ttk.Notebook(self.right_panel)
        self.classifier_notebook.pack(fill=tk.BOTH, expand=True)
        
        # 分类结果标签页
        self.tab_res = tk.Frame(self.classifier_notebook)
        self.classifier_notebook.add(self.tab_res, text=" 📊 分类结果与报告 ")
        
        # 交互绘图标签页
        self.tab_plt = tk.Frame(self.classifier_notebook, bg="white")
        self.classifier_notebook.add(self.tab_plt, text=" 📈 交互绘图区 ")
        
        # 初始化各个模块
        self.setup_left_panel()
        self.setup_results_tab()
        self.setup_plot_tab()
        self.apply_font_style()

    def setup_left_panel(self):
        """设置左侧控制面板"""
        # 0. 全局设置 (字号下拉框)
        settings_frame = tk.LabelFrame(self.left_panel, text="0. 全局字号设置", padx=10, pady=8, font=("", 10, "bold"),
                                       fg="purple")
        settings_frame.pack(fill=tk.X, pady=5)
        tk.Label(settings_frame, text="界面字号:").pack(side=tk.LEFT)
        self.combo_font = ttk.Combobox(settings_frame, values=[str(i) for i in range(8, 31)], width=5, state="readonly")
        self.combo_font.set(str(self.current_font_size))
        self.combo_font.pack(side=tk.LEFT, padx=5)
        self.combo_font.bind("<<ComboboxSelected>>", self.on_font_combo_change)

        # 1. 数据导入
        control_frame = tk.LabelFrame(self.left_panel, text="1. 数据导入", padx=10, pady=10)
        control_frame.pack(fill=tk.X, pady=5)
        self.text_input = tk.Text(control_frame, height=10, width=40, font=("Consolas", 10))
        self.text_input.pack(fill=tk.X, pady=5)
        tk.Button(control_frame, text="📋 粘贴并解析数据", command=self.load_from_text, bg="#e1f5fe",
                  font=("", 10, "bold")).pack(fill=tk.X)

        # 2. 交互模式
        mode_frame = tk.LabelFrame(self.left_panel, text="2. 绘图模式切换", padx=10, pady=10, fg="blue")
        mode_frame.pack(fill=tk.X, pady=10)
        tk.Radiobutton(mode_frame, text="🖱️ 直线模式 (左键加线/右键删线)", variable=self.enable_lasso_mode, value=False,
                       command=self.update_plot_view).pack(anchor="w")
        tk.Radiobutton(mode_frame, text="🎯 圈选模式 (画圈提取数据)", variable=self.enable_lasso_mode, value=True,
                       command=self.update_plot_view).pack(anchor="w")

        # 3. 操作
        op_frame = tk.LabelFrame(self.left_panel, text="3. 全局重置", padx=10, pady=10)
        op_frame.pack(fill=tk.X, pady=10)
        tk.Button(op_frame, text="🗑️ 清空所有数据及分类", command=self.reset_all, bg="#ffdddd").pack(fill=tk.X)

    def setup_results_tab(self):
        """设置分类结果标签页"""
        self.inner_nb = ttk.Notebook(self.tab_res)
        self.inner_nb.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- 分类树 ---
        self.tab_tree = tk.Frame(self.inner_nb)
        self.inner_nb.add(self.tab_tree, text="分类目录树")
        t_bar = tk.Frame(self.tab_tree, bg="#ddd")
        t_bar.pack(fill=tk.X, side=tk.TOP)
        tk.Button(t_bar, text="➕ 新增", command=self.open_add_data_dialog, bg="#ccffcc").pack(side=tk.LEFT, padx=2,
                                                                                              pady=2)
        tk.Button(t_bar, text="❌ 删除", command=self.delete_selected_data, bg="#ffcccc").pack(side=tk.LEFT, padx=2,
                                                                                              pady=2)
        tk.Label(t_bar, text="|").pack(side=tk.LEFT, padx=2)
        tk.Button(t_bar, text="↑ 上移", command=self.move_item_up).pack(side=tk.LEFT, padx=2)
        tk.Button(t_bar, text="↓ 下移", command=self.move_item_down).pack(side=tk.LEFT, padx=2)

        self.tree = ttk.Treeview(self.tab_tree, columns=('Label', 'Status', 'Index'), show='tree headings',
                                 displaycolumns=('Label', 'Status'))
        self.tree.heading('#0', text='分类目录');
        self.tree.heading('Label', text='名称');
        self.tree.heading('Status', text='标记')
        self.tree.column('Index', width=0, stretch=False)
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.tree.bind("<ButtonPress-1>", self.on_drag_start)
        self.tree.bind("<B1-Motion>", self.on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self.on_drag_release)
        self.tree.bind("<Button-3>", self.on_right_click)

        # --- 报告页 ---
        self.tab_report = tk.Frame(self.inner_nb)
        self.inner_nb.add(self.tab_report, text="文本报告")
        r_bar = tk.Frame(self.tab_report, bg="#ddd")
        r_bar.pack(fill=tk.X, side=tk.TOP)
        tk.Button(r_bar, text="💾 导出 TXT", command=self.export_txt_file, bg="#e1f5fe").pack(side=tk.LEFT, padx=5,
                                                                                             pady=2)
        tk.Button(r_bar, text="繁 -> 简", command=self.convert_to_simplified, bg="#fff0f5").pack(side=tk.LEFT, padx=2)
        tk.Button(r_bar, text="简 -> 繁", command=self.convert_to_traditional, bg="#fff0f5").pack(side=tk.LEFT, padx=2)
        self.report_text = tk.Text(self.tab_report);
        self.report_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def setup_plot_tab(self):
        """定义绘图标签页内容"""
        self.fig, self.ax = plt.subplots(figsize=(6, 6), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.tab_plt)
        self.canvas.mpl_connect('button_press_event', self.on_plot_click)

        # 添加 matplotlib 工具栏
        toolbar = NavigationToolbar2Tk(self.canvas, self.tab_plt)
        toolbar.update()

        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # ===============================================
    # 数据分类功能方法
    # ===============================================
    def move_item_up(self):
        """上移项目"""
        selected = self.tree.selection()
        for item in selected:
            parent = self.tree.parent(item)
            if parent:
                idx = self.tree.index(item)
                if idx > 0: self.tree.move(item, parent, idx - 1)
        self.generate_report_from_tree()

    def move_item_down(self):
        """下移项目"""
        selected = reversed(self.tree.selection())
        for item in selected:
            parent = self.tree.parent(item)
            if parent:
                idx = self.tree.index(item)
                siblings = self.tree.get_children(parent)
                if idx < len(siblings) - 1: self.tree.move(item, parent, idx + 1)
        self.generate_report_from_tree()

    def open_add_data_dialog(self):
        """打开新增数据对话框 (美化版)"""
        # 使用 create_popup_window 创建窗口，统一风格
        dialog = self.create_popup_window(self.root, "新增数据", "add_data_dialog", 420, 320)
        
        # 准备默认数据
        default_y, default_x, insert_pos = 0.0, 0.0, len(self.df)
        selected = self.tree.selection()
        if selected and self.tree.parent(selected[0]):
            vals = self.tree.item(selected[0], 'values')
            row_idx = int(vals[2])
            if row_idx in self.df.index:
                default_y, default_x = self.df.loc[row_idx, 'Y'] + 1, self.df.loc[row_idx, 'X']
                insert_pos = self.df.index.get_loc(row_idx) + 1

        # 1. 标题头
        tk.Label(dialog, text="➕ 添加新数据点", font=("Microsoft YaHei", 14, "bold"), fg="#333").pack(pady=(20, 15))

        # 2. 表单区域
        form_frame = tk.Frame(dialog)
        form_frame.pack(padx=40, pady=5, fill=tk.X)
        
        # 样式配置
        lbl_font = ("Microsoft YaHei", 10)
        ent_font = ("Arial", 11)
        
        # 名称
        tk.Label(form_frame, text="名称 Name:", font=lbl_font, fg="#555").grid(row=0, column=0, sticky="w", pady=8)
        n_ent = tk.Entry(form_frame, font=ent_font, bg="white", highlightthickness=1, relief="solid", bd=1)
        n_ent.config(highlightbackground="#ccc", highlightcolor="#2196F3") # Mac/Unix only potentially, but harmless
        n_ent.grid(row=0, column=1, sticky="ew", padx=(10, 0))
        n_ent.focus_set()
        
        # Y坐标
        tk.Label(form_frame, text="数值 Y:", font=lbl_font, fg="#555").grid(row=1, column=0, sticky="w", pady=8)
        y_ent = tk.Entry(form_frame, font=ent_font, bg="white", highlightthickness=1, relief="solid", bd=1)
        y_ent.insert(0, str(default_y))
        y_ent.grid(row=1, column=1, sticky="ew", padx=(10, 0))
        
        # X坐标
        tk.Label(form_frame, text="数值 X (可选):", font=lbl_font, fg="#555").grid(row=2, column=0, sticky="w", pady=8)
        x_ent = tk.Entry(form_frame, font=ent_font, bg="white", highlightthickness=1, relief="solid", bd=1)
        x_ent.insert(0, str(default_x))
        x_ent.grid(row=2, column=1, sticky="ew", padx=(10, 0))
        
        form_frame.columnconfigure(1, weight=1)

        # 3. 按钮区域
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=25, fill=tk.X)
        
        # 居中容器
        center_frame = tk.Frame(btn_frame)
        center_frame.pack()

        def save(event=None):
            name = n_ent.get().strip() or "未命名"
            try:
                y_val = float(y_ent.get())
                x_val = float(x_ent.get())
                
                row = pd.DataFrame([[name, y_val, x_val]], columns=['Label', 'Y', 'X'])
                self.df = pd.concat([self.df.iloc[:insert_pos], row, self.df.iloc[insert_pos:]]).reset_index(drop=True)
                self.category_list, self.marked_indices = [], set()
                self.refresh_all()
                dialog.destroy()
            except ValueError:
                messagebox.showerror("输入错误", "坐标值必须为数字！", parent=dialog)
            except Exception as e:
                messagebox.showerror("错误", f"添加失败: {e}", parent=dialog)

        save_btn = tk.Button(center_frame, text="保 存", command=save, bg="#4CAF50", fg="white", 
                            font=("Microsoft YaHei", 10, "bold"), width=12, padx=5, pady=3,
                            cursor="hand2", relief="raised", bd=0)
        save_btn.pack(side=tk.LEFT, padx=10)
        
        cancel_btn = tk.Button(center_frame, text="取 消", command=dialog.destroy, bg="#f5f5f5", fg="#333",
                              font=("Microsoft YaHei", 10), width=10, padx=5, pady=3,
                              cursor="hand2", relief="raised")
        cancel_btn.pack(side=tk.LEFT, padx=10)
        
        # 绑定回车键保存
        dialog.bind('<Return>', save)
        dialog.bind('<Escape>', lambda e: dialog.destroy())

    def on_drag_start(self, event):
        """开始拖拽"""
        item = self.tree.identify_row(event.y)
        if item and self.tree.parent(item): self.drag_source_item = item

    def on_drag_motion(self, event):
        """拖拽中"""
        target = self.tree.identify_row(event.y)
        if target: self.tree.selection_set(target)

    def on_drag_release(self, event):
        """结束拖拽"""
        if not self.drag_source_item: return
        target = self.tree.identify_row(event.y)
        if target and target != self.drag_source_item:
            dest_p = self.tree.parent(target) or target
            try:
                self.tree.move(self.drag_source_item, dest_p, self.tree.index(target))
                self.generate_report_from_tree()
            except:
                pass
        self.drag_source_item = None

    def on_plot_click(self, event):
        """绘图点击事件"""
        if event.inaxes != self.ax: return
        if not self.enable_lasso_mode.get():
            if event.button == 1:
                val = round(event.ydata, 1)
                if val not in self.thresholds: self.thresholds.append(val); self.thresholds.sort(); self.refresh_all()
            elif event.button == 3 and self.thresholds:
                closest = min(self.thresholds, key=lambda x: abs(x - event.ydata))
                if abs(closest - event.ydata) < (self.ax.get_ylim()[1] - self.ax.get_ylim()[0]) * 0.05:
                    self.thresholds.remove(closest);
                    self.refresh_all()

    def on_lasso_select(self, verts):
        """圈选事件"""
        if self.df.empty: return
        path = MplPath(verts)
        inside = path.contains_points(self.df[['X', 'Y']].values)
        new_idx = set(self.df.index[inside].tolist())
        if new_idx:
            for cat in self.category_list: cat['indices'] -= new_idx
            cat_id = len(self.category_list) + 1
            self.category_list.append({'name': f"圈选提取 {cat_id}", 'indices': new_idx,
                                       'color': self.color_cycle[(cat_id - 1) % len(self.color_cycle)]})
            self.refresh_all()

    def update_plot_view(self):
        """更新绘图视图"""
        self.ax.clear();
        self.ax.set_title("绘图交互区")
        if not self.df.empty:
            colors = ['#1f77b4'] * len(self.df);
            sizes = [60] * len(self.df)
            for i in self.df.index:
                if i in self.marked_indices:
                    colors[i], sizes[i] = 'red', 120
                else:
                    for cat in self.category_list:
                        if i in cat['indices']: colors[i], sizes[i] = cat['color'], 100; break
            self.ax.scatter(self.df['X'], self.df['Y'], c=colors, s=sizes, zorder=5)
            for idx, row in self.df.iterrows():
                m = idx in self.marked_indices
                self.ax.annotate(row['Label'], (row['X'], row['Y']), xytext=(0, 5), textcoords="offset points",
                                 ha='center', fontsize=9, color='red' if m else 'black',
                                 weight='bold' if m else 'normal')
        for y in self.thresholds: self.ax.axhline(y=y, color='blue', linestyle='--', alpha=0.5)
        if self.enable_lasso_mode.get():
            self.lasso = LassoSelector(self.ax, onselect=self.on_lasso_select, props={'color': 'red', 'linewidth': 1.5})
        else:
            if self.lasso: self.lasso.set_active(False); self.lasso = None
        self.canvas.draw()

    def classify_and_display(self):
        """分类并显示"""
        for i in self.tree.get_children(): self.tree.delete(i)
        if self.df.empty: return
        cat_idx = set()
        for i, cat in enumerate(self.category_list):
            if not cat['indices']: continue
            tag = f"tag_{cat['color']}"
            self.tree.tag_configure(tag, foreground=cat['color'], font=("", self.current_font_size, "bold"))
            pid = self.tree.insert("", "end", text=f"📂 {cat['name']}", open=True, tags=(tag,))
            for idx in sorted(list(cat['indices'])):
                m = idx in self.marked_indices
                self.tree.insert(pid, "end", values=(self.df.loc[idx, 'Label'], "✅ 标记" if m else "", idx),
                                 tags=('marked' if m else ''))
                cat_idx.add(idx)
        rem_df = self.df.drop(list(cat_idx))
        if not rem_df.empty:
            t_sorted = sorted(self.thresholds)
            line_cats = []
            if not t_sorted:
                line_cats.append(("数据区", rem_df))
            else:
                line_cats.append((f"低于 {t_sorted[0]}", rem_df[rem_df['Y'] < t_sorted[0]]))
                for i in range(len(t_sorted) - 1):
                    line_cats.append((f"{t_sorted[i]} ~ {t_sorted[i + 1]}",
                                      rem_df[(rem_df['Y'] >= t_sorted[i]) & (rem_df['Y'] < t_sorted[i + 1])]))
                line_cats.append((f"高于 {t_sorted[-1]}", rem_df[rem_df['Y'] >= t_sorted[-1]]))
            for name, sub in line_cats:
                if sub.empty: continue
                pid = self.tree.insert("", "end", text=f"📂 {self.custom_cat_names.get(name, name)}", open=True)
                for r_idx, r in sub.iterrows():
                    m = r_idx in self.marked_indices
                    self.tree.insert(pid, "end", values=(r['Label'], "✅ 标记" if m else "", r_idx),
                                     tags=('marked' if m else ''))
        self.generate_report_from_tree()

    def generate_report_from_tree(self):
        """从树生成报告"""
        self.report_text.delete("1.0", tk.END);
        content = ""
        for pid in self.tree.get_children(""):
            title = self.tree.item(pid, "text").replace("📂 ", "");
            children = self.tree.get_children(pid)
            if not children: continue
            content += f"【{title}】:\n"
            prev_m = None
            for i, cid in enumerate(children):
                vals = self.tree.item(cid, "values");
                name, idx = vals[0], int(vals[2]);
                curr_m = idx in self.marked_indices
                if curr_m:
                    if prev_m is False or prev_m is None: content += "\n"
                    content += f"{name}\n"
                else:
                    content += f"\n{name}\n\n"
                if curr_m:
                    next_m = False
                    if i < len(children) - 1: next_m = int(
                        self.tree.item(children[i + 1], "values")[2]) in self.marked_indices
                    if not next_m: content += "\n"
                prev_m = curr_m
            content += "\n"
        self.report_text.insert(tk.END, re.sub(r'\n{3,}', '\n\n', content).strip() + "\n")

    def on_font_combo_change(self, event):
        """字体大小改变"""
        self.current_font_size = int(self.combo_font.get())
        self.save_font_config()  # 保存字号设置
        self.apply_font_style()
        self.refresh_all()

    def apply_font_style(self):
        """应用字体样式"""
        s = self.current_font_size
        # 更新全局Treeview样式 (内容和标题)
        ttk.Style().configure("Treeview", font=("Microsoft YaHei", s), rowheight=int(s * 2.5))
        ttk.Style().configure("Treeview.Heading", font=("Microsoft YaHei", s, "bold"))
        
        # 更新特定标签样式
        self.tree.tag_configure('marked', foreground='red', font=("Microsoft YaHei", s, "bold"))
        self.report_text.configure(font=("Microsoft YaHei", s))

    def on_right_click(self, event):
        """右键点击事件"""
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            if self.tree.parent(iid):
                idx = int(self.tree.item(iid, 'values')[2])
                if idx in self.marked_indices:
                    self.marked_indices.remove(idx)
                else:
                    self.marked_indices.add(idx)
                self.refresh_all()
            else:
                old = self.tree.item(iid, "text").replace("📂 ", "")
                new = simpledialog.askstring("重命名", "分类名称:", initialvalue=old)
                if new:
                    idx = self.tree.get_children("").index(iid);
                    if idx < len(self.category_list):
                        self.category_list[idx]['name'] = new
                    else:
                        self.custom_cat_names[old] = new
                    self.refresh_all()

    def refresh_all(self):
        """刷新所有"""
        self.update_plot_view(); self.classify_and_display()

    def delete_selected_data(self):
        """删除选中数据"""
        items = self.tree.selection()
        indices = [int(self.tree.item(i, 'values')[2]) for i in items if self.tree.parent(i)]
        if indices and messagebox.askyesno("确认", "删除数据？"):
            self.df = self.df.drop(indices).reset_index(drop=True)
            self.category_list, self.marked_indices = [], set();
            self.refresh_all()

    def reset_all(self):
        """重置所有"""
        self.thresholds, self.category_list, self.marked_indices, self.custom_cat_names = [], [], set(), {};
        self.refresh_all()

    def load_from_text(self):
        """从文本加载数据"""
        try:
            txt = self.root.clipboard_get()
            if txt: self.text_input.delete("1.0", tk.END); self.text_input.insert(tk.END, txt)
        except:
            pass
        raw = self.text_input.get("1.0", tk.END).strip();
        data = []
        for line in raw.split('\n'):
            parts = re.split(r'[|\t,，]+', line.strip())
            if len(parts) >= 3:
                try:
                    data.append([parts[0].strip(), float(parts[1]), float(parts[2])])
                except:
                    continue
        if data:
            self.df = pd.DataFrame(data, columns=['Label', 'Y', 'X']);
            self.reset_all();
            self.main_notebook.select(self.classifier_tab)
            self.classifier_notebook.select(self.tab_plt)

    def convert_text(self, mode):
        """转换文本"""
        try:
            import opencc
            txt = self.report_text.get("1.0", tk.END).strip()
            if txt:
                converter = opencc.OpenCC(mode);
                self.report_text.delete("1.0", tk.END);
                self.report_text.insert(tk.END, converter.convert(txt))
        except ImportError:
            messagebox.showwarning("提示", "需要安装 opencc-python-reimplemented 库才能使用繁简转换功能")

    def convert_to_simplified(self):
        """转换为简体"""
        self.convert_text('t2s')

    def convert_to_traditional(self):
        """转换为繁体"""
        self.convert_text('s2t')

    def export_txt_file(self):
        """导出文本文件"""
        raw = self.report_text.get("1.0", tk.END);
        path = filedialog.asksaveasfilename(defaultextension=".txt")
        if path:
            filtered = [l for l in raw.splitlines() if not (l.strip().startswith("【") and "】" in l)]
            with open(path, "w", encoding="utf-8") as f: f.write("\n".join(filtered).strip())

    def _setup_drag_drop(self):
        """设置拖放功能"""
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD
            
            # 如果root不是TkinterDnD.Tk实例，则无法使用拖放
            # 这种情况下我们使用Windows原生的拖放API
            pass
        except ImportError:
            # 如果没有安装tkinterdnd2，使用Windows原生方法
            pass
        
        # 绑定拖放事件到主窗口和文件标签
        try:
            self.root.drop_target_register('DND_Files')
            self.root.dnd_bind('<<Drop>>', self._on_drop)
            
            if hasattr(self, 'file_label'):
                self.file_label.drop_target_register('DND_Files')
                self.file_label.dnd_bind('<<Drop>>', self._on_drop)
                
                # 添加拖放提示
                self.file_label.config(text="未选择文件 | 💡 可拖放图片到此处", fg="gray")
        except:
            # 如果拖放功能不可用，忽略错误
            pass
    
    def _on_drop(self, event):
        """处理拖放事件"""
        try:
            # 获取拖放的文件路径
            files = event.data
            print(f"拖放原始数据: {files}")  # 调试信息
            print(f"数据类型: {type(files)}")  # 调试信息
            
            # 处理不同格式的路径
            file_list = []
            if isinstance(files, str):
                files = files.strip()
                
                # 尝试多种解析方式
                if files.startswith('{'):
                    # 格式1: {C:/path/file1.jpg} {C:/path/file2.jpg}
                    import re
                    file_list = re.findall(r'\{([^}]+)\}', files)
                else:
                    # 尝试智能分割路径
                    # Windows路径格式: C:\path\file.jpg 或 C:/path/file.jpg
                    import re
                    # 匹配Windows路径模式 (盘符:\路径)
                    pattern = r'[A-Za-z]:[^\s]+'
                    matches = re.findall(pattern, files)
                    
                    if matches:
                        file_list = matches
                    elif ' ' in files and not os.path.exists(files):
                        # 简单空格分割
                        file_list = files.split()
                    else:
                        # 单个文件
                        file_list = [files]
            elif isinstance(files, tuple):
                # 元组格式
                file_list = list(files)
            else:
                file_list = [str(files)]
            
            print(f"解析后的文件列表: {file_list}")  # 调试信息
            
            # 清理路径
            cleaned_files = []
            for f in file_list:
                # 移除各种引号和空格
                f = f.strip().strip('{}').strip('"').strip("'").strip()
                
                # 尝试不同的路径格式
                # 1. 原始路径
                if os.path.exists(f):
                    cleaned_files.append(f)
                    print(f"✓ 找到文件: {f}")
                    continue
                
                # 2. 转换斜杠
                f_backslash = f.replace('/', '\\')
                if os.path.exists(f_backslash):
                    cleaned_files.append(f_backslash)
                    print(f"✓ 找到文件(转换后): {f_backslash}")
                    continue
                
                # 3. 转换为正斜杠
                f_slash = f.replace('\\', '/')
                if os.path.exists(f_slash):
                    cleaned_files.append(f_slash)
                    print(f"✓ 找到文件(转换后): {f_slash}")
                    continue
                
                print(f"✗ 文件不存在: {f}")
            
            if not cleaned_files:
                error_msg = f"未找到有效的文件！\n\n原始数据: {files}\n解析结果: {file_list}\n\n请确保拖放的是图片文件。"
                messagebox.showwarning("提示", error_msg)
                return
            
            # 过滤出图片文件
            image_extensions = ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp')
            image_files = [f for f in cleaned_files if f.lower().endswith(image_extensions)]
            
            if not image_files:
                messagebox.showwarning("提示", f"请拖放图片文件！\n\n找到 {len(cleaned_files)} 个文件，但都不是图片格式\n支持格式：JPG, PNG, BMP等")
                return
            
            # 单张图片直接选择
            if len(image_files) == 1:
                self.select_file_internal(image_files[0])
                self.progress_label.config(text=f"✓ 已通过拖放选择 1 个文件")
            else:
                # 多张图片，弹出选项菜单
                self._show_multi_image_options(image_files)
        
        except Exception as e:
            print(f"拖放处理错误: {e}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("错误", f"拖放文件失败：{str(e)}")
    
    def _show_multi_image_options(self, image_files):
        """显示多图片操作选项"""
        option_window = self.create_popup_window(self.root, "选择操作", "multi_image_options", 500, 400)
        
        tk.Label(option_window, text="🖼️ 检测到多张图片", 
                font=("Arial", 14, "bold")).pack(pady=20)
        
        tk.Label(option_window, text=f"已拖入 {len(image_files)} 张图片", 
                fg="blue", font=("Arial", 11)).pack(pady=5)
        
        tk.Label(option_window, text="请选择操作方式：", 
                font=("Arial", 10)).pack(pady=15)
        
        # 选项1：批量识别
        option1_frame = tk.Frame(option_window, relief=tk.RIDGE, borderwidth=2, bg="#E3F2FD")
        option1_frame.pack(pady=8, padx=30, fill=tk.X)
        
        tk.Label(option1_frame, text="1️⃣ 批量识别", 
                font=("Arial", 12, "bold"), bg="#E3F2FD").pack(pady=8)
        
        tk.Label(option1_frame, text="分别识别每张图片，适合处理多个独立文档", 
                fg="gray", font=("Arial", 9), bg="#E3F2FD").pack(pady=5)
        
        def batch_recognize():
            option_window.destroy()
            self.batch_select_files_internal(image_files)
            self.progress_label.config(text=f"✓ 已通过拖放选择 {len(image_files)} 个文件")
        
        tk.Button(option1_frame, text="批量识别", command=batch_recognize,
                 bg="#2196F3", fg="white", padx=20, pady=6, font=("Arial", 10)).pack(pady=8)
        
        # 选项2：拼接图片
        option2_frame = tk.Frame(option_window, relief=tk.RIDGE, borderwidth=2, bg="#FFF3E0")
        option2_frame.pack(pady=8, padx=30, fill=tk.X)
        
        tk.Label(option2_frame, text="2️⃣ 拼接图片", 
                font=("Arial", 12, "bold"), bg="#FFF3E0").pack(pady=8)
        
        tk.Label(option2_frame, text="将多张图片横向拼接成一张，然后识别", 
                fg="gray", font=("Arial", 9), bg="#FFF3E0").pack(pady=5)
        
        def merge_images_action():
            option_window.destroy()
            self._merge_images_from_drag(image_files)
        
        tk.Button(option2_frame, text="拼接图片", command=merge_images_action,
                 bg="#FF9800", fg="white", padx=20, pady=6, font=("Arial", 10)).pack(pady=8)
        
        # 选项3：裁剪拼接
        option3_frame = tk.Frame(option_window, relief=tk.RIDGE, borderwidth=2, bg="#E8F5E9")
        option3_frame.pack(pady=8, padx=30, fill=tk.X)
        
        tk.Label(option3_frame, text="3️⃣ 裁剪拼接", 
                font=("Arial", 12, "bold"), bg="#E8F5E9").pack(pady=8)
        
        tk.Label(option3_frame, text="手动框选区域后拼接，适合精确裁剪", 
                fg="gray", font=("Arial", 9), bg="#E8F5E9").pack(pady=5)
        
        def crop_merge_action():
            option_window.destroy()
            self._open_crop_window(image_files)
        
        tk.Button(option3_frame, text="裁剪拼接", command=crop_merge_action,
                 bg="#4CAF50", fg="white", padx=20, pady=6, font=("Arial", 10)).pack(pady=8)
        
        # 取消按钮
        tk.Button(option_window, text="取消", command=option_window.destroy,
                 bg="#757575", fg="white", padx=30, pady=8).pack(pady=15)
    
    def _merge_images_from_drag(self, file_paths):
        """从拖放触发的拼接图片功能"""
        try:
            # 加载所有图片
            images = []
            for path in file_paths:
                img = Image.open(path)
                images.append(img)
            
            # 计算拼接后的尺寸
            total_width = sum(img.width for img in images)
            max_height = max(img.height for img in images)
            
            # 创建拼接图片（从右到左）
            merged_image = Image.new('RGB', (total_width, max_height), 'white')
            
            x_offset = 0
            for img in reversed(images):
                y_offset = (max_height - img.height) // 2
                merged_image.paste(img, (x_offset, y_offset))
                x_offset += img.width
            
            # 询问是否保存
            save_choice = messagebox.askyesnocancel(
                "拼接完成",
                f"拼接完成！\n\n"
                f"图片数量: {len(images)}\n"
                f"拼接尺寸: {total_width}x{max_height}\n\n"
                f"是否保存拼接后的图片？\n\n"
                f"「是」= 保存图片并识别\n"
                f"「否」= 只识别不保存\n"
                f"「取消」= 取消操作"
            )
            
            if save_choice is None:  # 取消
                return
            
            # 保存到临时文件
            import tempfile
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, "merged_temp.jpg")
            merged_image.save(temp_path, format='JPEG', quality=90)
            
            # 如果选择保存
            if save_choice:
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".jpg",
                    filetypes=[("JPEG图片", "*.jpg"), ("PNG图片", "*.png"), ("所有文件", "*.*")],
                    initialfile=f"merged_{len(images)}images_{total_width}x{max_height}.jpg"
                )
                
                if save_path:
                    if save_path.lower().endswith('.png'):
                        merged_image.save(save_path, format='PNG')
                    else:
                        merged_image.save(save_path, format='JPEG', quality=95)
                    
                    self.progress_label.config(
                        text=f"✓ 拼接图片已保存到：{os.path.basename(save_path)}")
                    temp_path = save_path
            
            # 继续识别流程
            result = messagebox.askyesno("开始识别", 
                f"是否立即识别拼接后的图片？\n\n"
                f"拼接尺寸: {total_width}x{max_height}")
            
            if result:
                self.image_paths = [temp_path]
                self.file_label.config(
                    text=f"已选择: 拼接图片 ({len(images)}张) - {total_width}x{max_height}", 
                    fg="blue")
                
                # 检查尺寸并启用相应按钮
                width_in_accurate = self.size_limits["accurate_min_width"] <= total_width <= self.size_limits["accurate_max_width"]
                height_in_accurate = self.size_limits["accurate_min_height"] <= max_height <= self.size_limits["accurate_max_height"]
                meets_accurate = width_in_accurate and height_in_accurate
                
                width_in_basic = self.size_limits["basic_min_width"] <= total_width <= self.size_limits["basic_max_width"]
                height_in_basic = self.size_limits["basic_min_height"] <= max_height <= self.size_limits["basic_max_height"]
                meets_basic = width_in_basic and height_in_basic
                
                if meets_accurate:
                    self.ocr_btn.config(state=tk.NORMAL)
                else:
                    self.ocr_btn.config(state=tk.DISABLED)
                
                if meets_basic:
                    self.quick_ocr_btn.config(state=tk.NORMAL)
                else:
                    self.quick_ocr_btn.config(state=tk.DISABLED)
                
                self.progress_label.config(text="")
                
                # 选择识别方式
                if meets_accurate and meets_basic:
                    ocr_choice = messagebox.askyesno("选择识别方式",
                        f"是否使用高精度识别？\n\n"
                        f"「是」= 高精度识别\n"
                        f"「否」= 快速识别")
                    if ocr_choice:
                        self.root.after(500, self.perform_ocr)
                    else:
                        self.root.after(500, self.perform_quick_ocr)
                elif meets_accurate:
                    self.root.after(500, self.perform_ocr)
                elif meets_basic:
                    self.root.after(500, self.perform_quick_ocr)
                else:
                    messagebox.showwarning("警告", 
                        f"拼接后的图片尺寸不符合任何识别要求\n\n"
                        f"当前尺寸: {total_width}x{max_height}\n"
                        f"高精度要求: 宽≥{self.size_limits['accurate_min_width']} 且 高≥{self.size_limits['accurate_min_height']}\n"
                        f"快速识别要求: 宽<{self.size_limits['basic_max_width']} 且 高<{self.size_limits['basic_max_height']}")
        
        except Exception as e:
            messagebox.showerror("错误", f"拼接失败：{str(e)}")
    
    def _create_ribbon_group(self, parent, title):
        """创建Ribbon功能组"""
        group_frame = tk.Frame(parent, bg="#f0f0f0", relief=tk.FLAT, bd=0)
        group_frame.pack(side=tk.LEFT, padx=5, pady=5)
        
        # 按钮容器
        btn_container = tk.Frame(group_frame, bg="#f0f0f0")
        btn_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 3))
        
        # 组标题
        title_label = tk.Label(group_frame, text=title, bg="#f0f0f0", fg="#333", 
                              font=("Arial", 8), anchor=tk.CENTER)
        title_label.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 右侧分隔线
        separator = tk.Frame(parent, width=1, bg="#d0d0d0")
        separator.pack(side=tk.LEFT, fill=tk.Y, padx=3, pady=8)
        
        return btn_container
    
    def _create_ribbon_button(self, parent, text, command, color, large=False, state=tk.NORMAL):
        """创建Ribbon按钮"""
        if large:
            # 大按钮（单个）
            btn = tk.Button(parent, text=text, command=command, bg=color, fg="white",
                          font=("Arial", 9), width=10, height=3, relief=tk.RAISED, bd=1,
                          cursor="hand2", state=state)
            btn.pack(side=tk.LEFT, padx=3, pady=2)
        else:
            # 小按钮（多个）
            btn = tk.Button(parent, text=text, command=command, bg=color, fg="white",
                          font=("Arial", 8), width=8, height=3, relief=tk.RAISED, bd=1,
                          cursor="hand2", state=state)
            btn.pack(side=tk.LEFT, padx=2, pady=2)
        
        # 鼠标悬停效果
        def on_enter(e):
            if btn['state'] != tk.DISABLED:
                btn['relief'] = tk.RAISED
                btn['bd'] = 2
        
        def on_leave(e):
            btn['relief'] = tk.RAISED
            btn['bd'] = 1
        
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        
        return btn
    
    def unlock_size_limit(self):
        """解锁尺寸限制功能（提供两个选项）"""
        if self.size_limit_unlocked:
            # 已解锁，显示选项菜单
            self.show_unlock_menu()
            return
        
        # 创建密码输入窗口
        password_window = self.create_popup_window(self.root, "解锁尺寸限制", "unlock_password", 500, 350)
        
        tk.Label(password_window, text="🔓 解锁尺寸限制", 
                font=("Arial", 14, "bold")).pack(pady=20)
        
        tk.Label(password_window, text="解锁后可以：", 
                fg="gray", font=("Arial", 10)).pack(pady=5)
        
        tk.Label(password_window, text="1️⃣ 解除所有限制（任意尺寸使用高精度）", 
                fg="blue", font=("Arial", 9)).pack(pady=2)
        
        tk.Label(password_window, text="2️⃣ 修改尺寸范围（自定义限制）", 
                fg="blue", font=("Arial", 9)).pack(pady=2)
        
        tk.Label(password_window, text="请输入密码：", font=("Arial", 10)).pack(pady=15)
        password_entry = tk.Entry(password_window, show="*", font=("Arial", 12), width=20)
        password_entry.pack(pady=5)
        password_entry.focus_set()
        
        result_label = tk.Label(password_window, text="", fg="red")
        result_label.pack(pady=5)
        
        def check_password():
            entered_password = password_entry.get()
            if entered_password == self.unlock_password:
                self.size_limit_unlocked = True
                self.unlock_btn.config(text="🔓 已解锁", bg="#4CAF50")
                
                password_window.destroy()
                
                # 显示选项菜单
                self.show_unlock_menu()
                
                if self.image_paths:
                    if len(self.image_paths) == 1:
                        self.select_file_internal(self.image_paths[0])
                    else:
                        self.batch_select_files_internal(self.image_paths)
            else:
                result_label.config(text="❌ 密码错误，请重试")
                password_entry.delete(0, tk.END)
        
        btn_frame = tk.Frame(password_window)
        btn_frame.pack(pady=15)
        
        tk.Button(btn_frame, text="确定", command=check_password,
                 bg="#4CAF50", fg="white", padx=30, pady=8).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="取消", command=password_window.destroy,
                 bg="#757575", fg="white", padx=30, pady=8).pack(side=tk.LEFT, padx=5)
        
        password_entry.bind("<Return>", lambda e: check_password())
    
    def show_unlock_menu(self):
        """显示解锁后的选项菜单"""
        menu_window = self.create_popup_window(self.root, "尺寸限制管理", "size_limit_menu", 550, 500)
        
        tk.Label(menu_window, text="🔓 尺寸限制管理", 
                font=("Arial", 14, "bold")).pack(pady=20)
        
        tk.Label(menu_window, text="请选择操作：", 
                fg="gray", font=("Arial", 10)).pack(pady=10)
        
        # 选项1：解除所有限制
        option1_frame = tk.Frame(menu_window, relief=tk.RIDGE, borderwidth=2, bg="#E3F2FD")
        option1_frame.pack(pady=10, padx=30, fill=tk.X)
        
        tk.Label(option1_frame, text="1️⃣ 解除所有限制", 
                font=("Arial", 12, "bold"), bg="#E3F2FD").pack(pady=10)
        
        tk.Label(option1_frame, text="允许对任意尺寸的图片使用高精度识别\n不受尺寸范围限制", 
                fg="gray", font=("Arial", 9), bg="#E3F2FD").pack(pady=5)
        
        def remove_all_limits():
            # 设置为无限制模式
            if hasattr(self, 'size_hint_label'):
                bas_range = f"{self.size_limits['basic_min_width']}~{self.size_limits['basic_max_width']}x{self.size_limits['basic_min_height']}~{self.size_limits['basic_max_height']}"
                self.size_hint_label.config(text=f"💡 高精度(已解除限制) | 快速({bas_range})")
            else:
                # 兼容旧版本的更新方式
                for widget in self.progress_frame.winfo_children():
                    if isinstance(widget, tk.Label) and ("高精度" in widget.cget("text") or "已解锁" in widget.cget("text")):
                        bas_range = f"{self.size_limits['basic_min_width']}~{self.size_limits['basic_max_width']}x{self.size_limits['basic_min_height']}~{self.size_limits['basic_max_height']}"
                        widget.config(text=f"💡 高精度(已解除限制) | 快速({bas_range})")
            
            menu_window.destroy()
            messagebox.showinfo("成功", 
                "已解除所有尺寸限制！\n\n"
                "现在可以对任意尺寸的图片使用高精度识别")
            
            if self.image_paths:
                if len(self.image_paths) == 1:
                    self.select_file_internal(self.image_paths[0])
                else:
                    self.batch_select_files_internal(self.image_paths)
        
        tk.Button(option1_frame, text="解除所有限制", command=remove_all_limits,
                 bg="#2196F3", fg="white", padx=20, pady=8, font=("Arial", 10)).pack(pady=10)
        
        # 选项2：修改尺寸范围
        option2_frame = tk.Frame(menu_window, relief=tk.RIDGE, borderwidth=2, bg="#FFF3E0")
        option2_frame.pack(pady=10, padx=30, fill=tk.X)
        
        tk.Label(option2_frame, text="2️⃣ 修改尺寸范围", 
                font=("Arial", 12, "bold"), bg="#FFF3E0").pack(pady=10)
        
        tk.Label(option2_frame, text="自定义高精度和快速识别的尺寸范围\n更灵活地控制识别条件", 
                fg="gray", font=("Arial", 9), bg="#FFF3E0").pack(pady=5)
        
        def open_size_settings():
            menu_window.destroy()
            self.show_size_settings()
        
        tk.Button(option2_frame, text="修改尺寸范围", command=open_size_settings,
                 bg="#FF9800", fg="white", padx=20, pady=8, font=("Arial", 10)).pack(pady=10)
        
        # 关闭按钮
        tk.Button(menu_window, text="关闭", command=menu_window.destroy,
                 bg="#757575", fg="white", padx=30, pady=8).pack(pady=15)
    
    def select_file_internal(self, file_path):
        """内部方法：处理文件选择逻辑"""
        self.image_paths = [file_path]
        
        try:
            img = Image.open(file_path)
            width, height = img.size
            file_size = os.path.getsize(file_path)
            
            if file_size < 1024 * 1024:
                size_str = f"{file_size/1024:.1f}KB"
            else:
                size_str = f"{file_size/(1024*1024):.1f}MB"
            
            if self.size_limit_unlocked:
                meets_accurate_requirement = True
            else:
                # 高精度：宽度和高度都在范围内（两个都要满足）
                width_in_accurate_range = self.size_limits["accurate_min_width"] <= width <= self.size_limits["accurate_max_width"]
                height_in_accurate_range = self.size_limits["accurate_min_height"] <= height <= self.size_limits["accurate_max_height"]
                meets_accurate_requirement = width_in_accurate_range and height_in_accurate_range
            
            # 快速识别：宽度和高度都在范围内（两个都要满足）
            width_in_basic_range = self.size_limits["basic_min_width"] <= width <= self.size_limits["basic_max_width"]
            height_in_basic_range = self.size_limits["basic_min_height"] <= height <= self.size_limits["basic_max_height"]
            meets_basic_requirement = width_in_basic_range and height_in_basic_range
            
            # 通用识别：宽度和高度都在范围内（两个都要满足）
            width_in_general_range = self.size_limits["general_min_width"] <= width <= self.size_limits["general_max_width"]
            height_in_general_range = self.size_limits["general_min_height"] <= height <= self.size_limits["general_max_height"]
            meets_general_requirement = width_in_general_range and height_in_general_range
            
            # 统计符合的模式数量
            available_modes = []
            if meets_accurate_requirement:
                available_modes.append("高精度")
            if meets_basic_requirement:
                available_modes.append("快速")
            if meets_general_requirement:
                available_modes.append("通用")
            
            # 根据可用模式设置按钮状态和提示信息
            self.ocr_btn.config(state=tk.NORMAL if meets_accurate_requirement else tk.DISABLED)
            self.quick_ocr_btn.config(state=tk.NORMAL if meets_basic_requirement else tk.DISABLED)
            self.general_ocr_btn.config(state=tk.NORMAL if meets_general_requirement else tk.DISABLED)
            
            unlock_hint = " [已解锁]" if self.size_limit_unlocked and (width < self.size_limits["accurate_min_width"] or height < self.size_limits["accurate_min_height"]) else ""
            
            if len(available_modes) == 3:
                # 三种模式都可用
                info_text = f"已选择: {os.path.basename(file_path)} ({width}x{height}, {size_str}){unlock_hint}"
                self.file_label.config(text=info_text, fg="black")
                self.progress_label.config(text="")
            elif len(available_modes) == 2:
                # 两种模式可用
                modes_str = "、".join(available_modes)
                info_text = f"已选择: {os.path.basename(file_path)} ({width}x{height}, {size_str}){unlock_hint} ✓ 可用: {modes_str}"
                self.file_label.config(text=info_text, fg="blue")
                unavailable = [m for m in ["高精度", "快速", "通用"] if m not in available_modes]
                self.progress_label.config(text=f"💡 提示：{unavailable[0]}识别不可用，建议使用{modes_str}识别")
            elif len(available_modes) == 1:
                # 只有一种模式可用
                mode_str = available_modes[0]
                info_text = f"已选择: {os.path.basename(file_path)} ({width}x{height}, {size_str}){unlock_hint} ⚠️ 仅可用: {mode_str}"
                self.file_label.config(text=info_text, fg="orange")
                self.progress_label.config(text=f"💡 提示：该图片尺寸仅符合{mode_str}识别要求")
            else:
                # 没有可用模式
                info_text = f"已选择: {os.path.basename(file_path)} ({width}x{height}, {size_str}) ❌ 尺寸不符合任何识别要求"
                self.file_label.config(text=info_text, fg="red")
                self.progress_label.config(text="❌ 错误：图片尺寸不符合任何识别要求，请检查图片尺寸或点击「解锁限制」")
        except:
            self.file_label.config(text=f"已选择: {os.path.basename(file_path)}", fg="black")
            self.ocr_btn.config(state=tk.NORMAL)
            self.quick_ocr_btn.config(state=tk.NORMAL)
            self.general_ocr_btn.config(state=tk.NORMAL)
            self.progress_label.config(text="")
    
    def select_file(self):
        """选择图片文件（支持多选）"""
        file_paths = filedialog.askopenfilenames(
            title="选择图片（可多选）",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp"), ("所有文件", "*.*")]
        )
        if file_paths:
            if len(file_paths) == 1:
                self.select_file_internal(file_paths[0])
            else:
                self.batch_select_files_internal(list(file_paths))
    
    def batch_select_files_internal(self, file_paths):
        """内部方法：处理批量文件选择逻辑"""
        self.image_paths = file_paths
        count = len(self.image_paths)
        
        meets_accurate_count = 0
        meets_basic_count = 0
        meets_general_count = 0
        meets_all_count = 0
        meets_none_count = 0
        
        try:
            total_size = 0
            for path in self.image_paths:
                total_size += os.path.getsize(path)
                try:
                    img = Image.open(path)
                    width, height = img.size
                    
                    if self.size_limit_unlocked:
                        meets_accurate = True
                    else:
                        # 高精度：宽度和高度都在范围内
                        width_in_accurate = self.size_limits["accurate_min_width"] <= width <= self.size_limits["accurate_max_width"]
                        height_in_accurate = self.size_limits["accurate_min_height"] <= height <= self.size_limits["accurate_max_height"]
                        meets_accurate = width_in_accurate and height_in_accurate
                    
                    # 快速识别：宽度和高度都在范围内
                    width_in_basic = self.size_limits["basic_min_width"] <= width <= self.size_limits["basic_max_width"]
                    height_in_basic = self.size_limits["basic_min_height"] <= height <= self.size_limits["basic_max_height"]
                    meets_basic = width_in_basic and height_in_basic
                    
                    # 通用识别：宽度和高度都在范围内
                    width_in_general = self.size_limits["general_min_width"] <= width <= self.size_limits["general_max_width"]
                    height_in_general = self.size_limits["general_min_height"] <= height <= self.size_limits["general_max_height"]
                    meets_general = width_in_general and height_in_general
                    
                    # 统计各种组合
                    available_modes = 0
                    if meets_accurate:
                        meets_accurate_count += 1
                        available_modes += 1
                    if meets_basic:
                        meets_basic_count += 1
                        available_modes += 1
                    if meets_general:
                        meets_general_count += 1
                        available_modes += 1
                    
                    if available_modes == 3:
                        meets_all_count += 1
                    elif available_modes == 0:
                        meets_none_count += 1
                        
                except:
                    meets_none_count += 1
            
            if total_size < 1024 * 1024:
                size_str = f"{total_size/1024:.1f}KB"
            else:
                size_str = f"{total_size/(1024*1024):.1f}MB"
            
            info_parts = [f"已选择 {count} 个文件 (总大小: {size_str})"]
            if meets_all_count > 0:
                info_parts.append(f"全部可用: {meets_all_count}张")
            if meets_accurate_count > meets_all_count:
                info_parts.append(f"高精度: {meets_accurate_count}张")
            if meets_basic_count > meets_all_count:
                info_parts.append(f"快速: {meets_basic_count}张")
            if meets_general_count > meets_all_count:
                info_parts.append(f"通用: {meets_general_count}张")
            if meets_none_count > 0:
                info_parts.append(f"都不符合: {meets_none_count}张")
            
            info_text = " | ".join(info_parts)
            
            # 设置按钮状态
            self.ocr_btn.config(state=tk.NORMAL if meets_accurate_count > 0 else tk.DISABLED)
            self.quick_ocr_btn.config(state=tk.NORMAL if meets_basic_count > 0 else tk.DISABLED)
            self.general_ocr_btn.config(state=tk.NORMAL if meets_general_count > 0 else tk.DISABLED)
            
            # 根据可用模式数量设置提示信息
            available_mode_count = sum([1 for count in [meets_accurate_count, meets_basic_count, meets_general_count] if count > 0])
            
            if available_mode_count == 3:
                self.file_label.config(text=info_text, fg="black")
                if meets_none_count > 0:
                    self.progress_label.config(text=f"💡 提示：{meets_none_count}张图片不符合任何识别要求，将被跳过")
                else:
                    self.progress_label.config(text="")
            elif available_mode_count == 2:
                available_modes = []
                if meets_accurate_count > 0:
                    available_modes.append("高精度")
                if meets_basic_count > 0:
                    available_modes.append("快速")
                if meets_general_count > 0:
                    available_modes.append("通用")
                modes_str = "、".join(available_modes)
                self.file_label.config(text=info_text + f" ✓ 可用: {modes_str}", fg="blue")
                self.progress_label.config(text=f"💡 提示：部分图片可用{modes_str}识别")
            elif available_mode_count == 1:
                if meets_accurate_count > 0:
                    mode_str = "高精度"
                elif meets_basic_count > 0:
                    mode_str = "快速"
                else:
                    mode_str = "通用"
                self.file_label.config(text=info_text + f" ⚠️ 仅可用: {mode_str}", fg="orange")
                self.progress_label.config(text=f"💡 提示：所有图片仅符合{mode_str}识别要求")
            else:
                self.file_label.config(text=info_text + " ❌ 所有图片都不符合任何识别要求", fg="red")
                if self.size_limit_unlocked:
                    self.progress_label.config(text="❌ 错误：所有图片尺寸都不符合任何识别要求")
                else:
                    self.progress_label.config(text="❌ 错误：所有图片尺寸都不符合任何识别要求，可点击「解锁限制」")
        except:
            self.file_label.config(text=f"已选择 {count} 个文件", fg="black")
            self.ocr_btn.config(state=tk.NORMAL)
            self.quick_ocr_btn.config(state=tk.NORMAL)
            self.general_ocr_btn.config(state=tk.NORMAL)
            self.progress_label.config(text="")

    
    def perform_ocr(self):
        """执行 OCR 识别（支持批量）- 使用多线程避免卡顿"""
        if not self.image_paths:
            messagebox.showwarning("警告", "请先选择图片文件！")
            return
        
        if not API_KEY or not SECRET_KEY:
            messagebox.showerror("错误", "请先在 .env 文件中配置 API_KEY 和 SECRET_KEY！")
            return
        
        self.ocr_btn.config(state=tk.DISABLED)
        self.select_btn.config(state=tk.DISABLED)
        
        thread = threading.Thread(target=self._perform_ocr_thread, daemon=True)
        thread.start()
    
    def _perform_ocr_thread(self):
        """OCR识别线程（后台执行）"""
        try:
            self.root.after(0, lambda: self.result_text.delete(1.0, tk.END))
            self.all_results = []
            
            total = len(self.image_paths)
            
            for idx, image_path in enumerate(self.image_paths, 1):
                self.root.after(0, lambda i=idx, p=image_path: 
                    self.progress_label.config(text=f"正在处理: {i}/{total} - {os.path.basename(p)}"))
                
                self.root.after(0, lambda: self.result_text.insert(tk.END, f"\n{'='*80}\n"))
                self.root.after(0, lambda i=idx, p=image_path: 
                    self.result_text.insert(tk.END, f"文件 {i}/{total}: {os.path.basename(p)}\n"))
                self.root.after(0, lambda: self.result_text.insert(tk.END, f"{'='*80}\n"))
                
                try:
                    img = Image.open(image_path)
                    width, height = img.size
                    
                    unlock_status = " [已解锁]" if self.size_limit_unlocked else ""
                    self.root.after(0, lambda w=width, h=height, u=unlock_status: 
                        self.result_text.insert(tk.END, f"图片尺寸: {w}x{h}{u}\n"))
                    
                    # 检查是否符合高精度识别要求
                    if not self.size_limit_unlocked:
                        width_in_accurate = self.size_limits["accurate_min_width"] <= width <= self.size_limits["accurate_max_width"]
                        height_in_accurate = self.size_limits["accurate_min_height"] <= height <= self.size_limits["accurate_max_height"]
                        meets_accurate = width_in_accurate and height_in_accurate
                        
                        if not meets_accurate:
                            acc_w_range = f"{self.size_limits['accurate_min_width']}~{self.size_limits['accurate_max_width']}"
                            acc_h_range = f"{self.size_limits['accurate_min_height']}~{self.size_limits['accurate_max_height']}"
                            self.root.after(0, lambda w=width, h=height, wr=acc_w_range, hr=acc_h_range: 
                                self.result_text.insert(tk.END, 
                                    f"⚠️ 跳过：图片尺寸不符合要求\n"
                                    f"   当前尺寸: {w}x{h}\n"
                                    f"   要求：宽度({wr})且高度({hr})都要在范围内\n"
                                    f"   建议使用「快速识别」按钮或点击「解锁限制」\n"))
                            
                            self.all_results.append({
                                'file': os.path.basename(image_path),
                                'path': image_path,
                                'lines': [],
                                'count': 0,
                                'skipped': True,
                                'reason': f'图片尺寸不符合要求（{width}x{height}）'
                            })
                            
                            self.root.after(0, lambda: self.result_text.see(tk.END))
                            continue
                    
                except Exception as e:
                    self.root.after(0, lambda err=str(e): 
                        self.result_text.insert(tk.END, f"⚠️ 无法读取图片尺寸: {err}\n"))
                
                result = ocr_image(image_path)
                
                if "words_result" in result:
                    formatted_lines = []
                    for item in result["words_result"]:
                        words = item["words"]
                        location = item.get("location", {})
                        top = location.get("top", 0)
                        left = location.get("left", 0)
                        height = location.get("height", 0)
                        formatted_lines.append(f"{words}|{top}|{left}|{height}")
                    
                    recognized_text = "\n".join(formatted_lines)
                    self.root.after(0, lambda t=recognized_text: 
                        self.result_text.insert(tk.END, t + "\n"))
                    
                    self.all_results.append({
                        'file': os.path.basename(image_path),
                        'path': image_path,
                        'lines': formatted_lines,
                        'count': len(formatted_lines)
                    })
                    
                    self.root.after(0, lambda c=len(formatted_lines): 
                        self.result_text.insert(tk.END, f"\n✓ 识别成功：{c} 行文字\n"))
                else:
                    self.root.after(0, lambda r=result: 
                        self.result_text.insert(tk.END, f"✗ 识别失败：{r}\n"))
                    self.all_results.append({
                        'file': os.path.basename(image_path),
                        'path': image_path,
                        'lines': [],
                        'count': 0,
                        'error': str(result)
                    })
                
                self.root.after(0, lambda: self.result_text.see(tk.END))
                
                if idx < total:
                    import time
                    time.sleep(0.5)
            
            success_count = sum(1 for r in self.all_results if r['count'] > 0)
            skipped_count = sum(1 for r in self.all_results if r.get('skipped', False))
            failed_count = total - success_count - skipped_count
            total_lines = sum(r['count'] for r in self.all_results)
            
            if total > 0:
                self.record_ocr('accurate', success_count, failed_count, total_lines)
                if skipped_count > 0:
                    today = datetime.now().strftime("%Y-%m-%d")
                    if today in self.stats and 'accurate' in self.stats[today]:
                        self.stats[today]['accurate']['skipped'] += skipped_count
                        self.save_stats()
                
                # 添加到历史记录（在主线程中执行）
                results_copy = [r.copy() for r in self.all_results]
                self.root.after(0, lambda: self.add_to_history('高精度识别', results_copy))
            
            self.root.after(0, lambda: self.progress_label.config(text=f"✓ 完成！共处理 {total} 个文件"))
            self.root.after(0, lambda: self.export_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.copy_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.add_zeros_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.quick_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.general_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.select_btn.config(state=tk.NORMAL))
            
            status_msg = f"✓ 高精度识别完成！总:{total} 成功:{success_count}"
            if skipped_count > 0:
                status_msg += f" 跳过:{skipped_count}"
            if failed_count > 0:
                status_msg += f" 失败:{failed_count}"
            status_msg += f" | 文字行数:{total_lines}"
            if skipped_count > 0:
                status_msg += " | 💡跳过的图片可用快速识别"
            
            self.root.after(0, lambda m=status_msg: self.progress_label.config(text=m))
        
        except Exception as e:
            self.root.after(0, lambda: self.result_text.insert(tk.END, f"\n发生错误：{str(e)}\n"))
            self.root.after(0, lambda: messagebox.showerror("错误", f"发生错误：{str(e)}"))
            self.root.after(0, lambda: self.progress_label.config(text="✗ 处理失败"))
            self.root.after(0, lambda: self.ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.quick_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.general_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.select_btn.config(state=tk.NORMAL))

    
    def perform_quick_ocr(self):
        """执行快速 OCR 识别"""
        if not self.image_paths:
            messagebox.showwarning("警告", "请先选择图片文件！")
            return
        
        if not API_KEY or not SECRET_KEY:
            messagebox.showerror("错误", "请先在 .env 文件中配置 API_KEY 和 SECRET_KEY！")
            return
        
        self.ocr_btn.config(state=tk.DISABLED)
        self.quick_ocr_btn.config(state=tk.DISABLED)
        self.general_ocr_btn.config(state=tk.DISABLED)
        self.select_btn.config(state=tk.DISABLED)
        
        thread = threading.Thread(target=self._perform_quick_ocr_thread, daemon=True)
        thread.start()
    
    def _perform_quick_ocr_thread(self):
        """快速OCR识别线程"""
        try:
            self.root.after(0, lambda: self.result_text.delete(1.0, tk.END))
            self.all_results = []
            
            total = len(self.image_paths)
            
            for idx, image_path in enumerate(self.image_paths, 1):
                self.root.after(0, lambda i=idx, p=image_path: 
                    self.progress_label.config(text=f"快速识别中: {i}/{total} - {os.path.basename(p)}"))
                
                self.root.after(0, lambda: self.result_text.insert(tk.END, f"\n{'='*80}\n"))
                self.root.after(0, lambda i=idx, p=image_path: 
                    self.result_text.insert(tk.END, f"文件 {i}/{total}: {os.path.basename(p)}\n"))
                self.root.after(0, lambda: self.result_text.insert(tk.END, f"{'='*80}\n"))
                
                try:
                    img = Image.open(image_path)
                    width, height = img.size
                    
                    self.root.after(0, lambda w=width, h=height: 
                        self.result_text.insert(tk.END, f"图片尺寸: 宽{w} x 高{h}\n"))
                    
                    # 检查是否符合快速识别要求
                    width_in_basic = self.size_limits["basic_min_width"] <= width <= self.size_limits["basic_max_width"]
                    height_in_basic = self.size_limits["basic_min_height"] <= height <= self.size_limits["basic_max_height"]
                    meets_basic = width_in_basic and height_in_basic
                    
                    if not meets_basic:
                        bas_w_range = f"{self.size_limits['basic_min_width']}~{self.size_limits['basic_max_width']}"
                        bas_h_range = f"{self.size_limits['basic_min_height']}~{self.size_limits['basic_max_height']}"
                        self.root.after(0, lambda w=width, h=height, wr=bas_w_range, hr=bas_h_range: 
                            self.result_text.insert(tk.END, 
                                f"⚠️ 跳过：图片尺寸不符合要求\n"
                                f"   当前尺寸: 宽{w} x 高{h}\n"
                                f"   要求：宽度({wr})且高度({hr})都要在范围内\n"
                                f"   建议使用「高精度识别」按钮\n"))
                        
                        self.all_results.append({
                            'file': os.path.basename(image_path),
                            'path': image_path,
                            'lines': [],
                            'count': 0,
                            'skipped': True,
                            'reason': f'图片尺寸不符合要求（宽{width} x 高{height}）'
                        })
                        
                        self.root.after(0, lambda: self.result_text.see(tk.END))
                        continue
                    
                except Exception as e:
                    self.root.after(0, lambda err=str(e): 
                        self.result_text.insert(tk.END, f"⚠️ 无法读取图片尺寸: {err}\n"))
                
                result = ocr_image_basic(image_path)
                
                if "words_result" in result:
                    text_only_lines = []
                    for item in result["words_result"]:
                        words = item["words"]
                        text_only_lines.append(words)
                    
                    recognized_text = "\n".join(text_only_lines)
                    self.root.after(0, lambda t=recognized_text: 
                        self.result_text.insert(tk.END, t + "\n"))
                    
                    self.all_results.append({
                        'file': os.path.basename(image_path),
                        'path': image_path,
                        'lines': text_only_lines,
                        'count': len(text_only_lines)
                    })
                    
                    self.root.after(0, lambda c=len(text_only_lines): 
                        self.result_text.insert(tk.END, f"\n✓ 识别成功：{c} 行文字\n"))
                else:
                    self.root.after(0, lambda r=result: 
                        self.result_text.insert(tk.END, f"✗ 识别失败：{r}\n"))
                    self.all_results.append({
                        'file': os.path.basename(image_path),
                        'path': image_path,
                        'lines': [],
                        'count': 0,
                        'error': str(result)
                    })
                
                self.root.after(0, lambda: self.result_text.see(tk.END))
                
                if idx < total:
                    import time
                    time.sleep(0.5)
            
            success_count = sum(1 for r in self.all_results if r['count'] > 0)
            skipped_count = sum(1 for r in self.all_results if r.get('skipped', False))
            failed_count = total - success_count - skipped_count
            total_lines = sum(r['count'] for r in self.all_results)
            
            actual_processed = total - skipped_count
            if actual_processed > 0:
                self.record_ocr('basic', success_count, failed_count, total_lines)
                # 添加到历史记录（在主线程中执行）
                results_copy = [r.copy() for r in self.all_results]
                self.root.after(0, lambda: self.add_to_history('快速识别', results_copy))
            
            self.root.after(0, lambda: self.progress_label.config(text=f"✓ 完成！共处理 {total} 个文件"))
            self.root.after(0, lambda: self.export_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.copy_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.add_zeros_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.quick_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.general_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.select_btn.config(state=tk.NORMAL))
            
            status_msg = f"✓ 快速识别完成！总:{total} 成功:{success_count}"
            if skipped_count > 0:
                status_msg += f" 跳过:{skipped_count}"
            if failed_count > 0:
                status_msg += f" 失败:{failed_count}"
            status_msg += f" | 文字行数:{total_lines}"
            if skipped_count > 0:
                status_msg += " | 💡跳过的图片可用高精度识别"
            
            self.root.after(0, lambda m=status_msg: self.progress_label.config(text=m))
        
        except Exception as e:
            self.root.after(0, lambda: self.result_text.insert(tk.END, f"\n发生错误：{str(e)}\n"))
            self.root.after(0, lambda: messagebox.showerror("错误", f"发生错误：{str(e)}"))
            self.root.after(0, lambda: self.progress_label.config(text="✗ 处理失败"))
            self.root.after(0, lambda: self.ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.quick_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.general_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.select_btn.config(state=tk.NORMAL))

    def perform_general_ocr(self):
        """执行通用 OCR 识别"""
        if not self.image_paths:
            messagebox.showwarning("警告", "请先选择图片文件！")
            return
        
        if not API_KEY or not SECRET_KEY:
            messagebox.showerror("错误", "请先在 .env 文件中配置 API_KEY 和 SECRET_KEY！")
            return
        
        self.ocr_btn.config(state=tk.DISABLED)
        self.quick_ocr_btn.config(state=tk.DISABLED)
        self.general_ocr_btn.config(state=tk.DISABLED)
        self.select_btn.config(state=tk.DISABLED)
        
        thread = threading.Thread(target=self._perform_general_ocr_thread, daemon=True)
        thread.start()
    
    def _perform_general_ocr_thread(self):
        """通用OCR识别线程"""
        try:
            self.root.after(0, lambda: self.result_text.delete(1.0, tk.END))
            self.all_results = []
            
            total = len(self.image_paths)
            
            for idx, image_path in enumerate(self.image_paths, 1):
                self.root.after(0, lambda i=idx, p=image_path: 
                    self.progress_label.config(text=f"通用识别中: {i}/{total} - {os.path.basename(p)}"))
                
                self.root.after(0, lambda: self.result_text.insert(tk.END, f"\n{'='*80}\n"))
                self.root.after(0, lambda i=idx, p=image_path: 
                    self.result_text.insert(tk.END, f"文件 {i}/{total}: {os.path.basename(p)}\n"))
                self.root.after(0, lambda: self.result_text.insert(tk.END, f"{'='*80}\n"))
                
                try:
                    img = Image.open(image_path)
                    width, height = img.size
                    
                    self.root.after(0, lambda w=width, h=height: 
                        self.result_text.insert(tk.END, f"图片尺寸: 宽{w} x 高{h}\n"))
                    
                    # 检查是否符合通用识别要求
                    width_in_general = self.size_limits["general_min_width"] <= width <= self.size_limits["general_max_width"]
                    height_in_general = self.size_limits["general_min_height"] <= height <= self.size_limits["general_max_height"]
                    meets_general = width_in_general and height_in_general
                    
                    if not meets_general:
                        gen_w_range = f"{self.size_limits['general_min_width']}~{self.size_limits['general_max_width']}"
                        gen_h_range = f"{self.size_limits['general_min_height']}~{self.size_limits['general_max_height']}"
                        self.root.after(0, lambda w=width, h=height, wr=gen_w_range, hr=gen_h_range: 
                            self.result_text.insert(tk.END, 
                                f"⚠️ 跳过：图片尺寸不符合要求\n"
                                f"   当前尺寸: 宽{w} x 高{h}\n"
                                f"   要求：宽度({wr})且高度({hr})都要在范围内\n"
                                f"   建议使用其他识别模式\n"))
                        
                        self.all_results.append({
                            'file': os.path.basename(image_path),
                            'path': image_path,
                            'lines': [],
                            'count': 0,
                            'skipped': True,
                            'reason': f'图片尺寸不符合要求（宽{width} x 高{height}）'
                        })
                        
                        self.root.after(0, lambda: self.result_text.see(tk.END))
                        continue
                    
                except Exception as e:
                    self.root.after(0, lambda err=str(e): 
                        self.result_text.insert(tk.END, f"⚠️ 无法读取图片尺寸: {err}\n"))
                
                result = ocr_image_general(image_path)
                
                if "words_result" in result:
                    formatted_lines = []
                    for item in result["words_result"]:
                        words = item["words"]
                        location = item.get("location", {})
                        top = location.get("top", 0)
                        left = location.get("left", 0)
                        height = location.get("height", 0)
                        formatted_lines.append(f"{words}|{top}|{left}|{height}")
                    
                    recognized_text = "\n".join(formatted_lines)
                    self.root.after(0, lambda t=recognized_text: 
                        self.result_text.insert(tk.END, t + "\n"))
                    
                    self.all_results.append({
                        'file': os.path.basename(image_path),
                        'path': image_path,
                        'lines': formatted_lines,
                        'count': len(formatted_lines)
                    })
                    
                    self.root.after(0, lambda c=len(formatted_lines): 
                        self.result_text.insert(tk.END, f"\n✓ 识别成功：{c} 行文字\n"))
                else:
                    self.root.after(0, lambda r=result: 
                        self.result_text.insert(tk.END, f"✗ 识别失败：{r}\n"))
                    self.all_results.append({
                        'file': os.path.basename(image_path),
                        'path': image_path,
                        'lines': [],
                        'count': 0,
                        'error': str(result)
                    })
                
                self.root.after(0, lambda: self.result_text.see(tk.END))
                
                if idx < total:
                    import time
                    time.sleep(0.5)
            
            success_count = sum(1 for r in self.all_results if r['count'] > 0)
            skipped_count = sum(1 for r in self.all_results if r.get('skipped', False))
            failed_count = total - success_count - skipped_count
            total_lines = sum(r['count'] for r in self.all_results)
            
            actual_processed = total - skipped_count
            if actual_processed > 0:
                self.record_ocr('general', success_count, failed_count, total_lines)
                # 添加到历史记录（在主线程中执行）
                results_copy = [r.copy() for r in self.all_results]
                self.root.after(0, lambda: self.add_to_history('通用识别', results_copy))
            
            self.root.after(0, lambda: self.progress_label.config(text=f"✓ 完成！共处理 {total} 个文件"))
            self.root.after(0, lambda: self.export_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.copy_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.add_zeros_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.quick_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.general_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.select_btn.config(state=tk.NORMAL))
            
            status_msg = f"✓ 通用识别完成！总:{total} 成功:{success_count}"
            if skipped_count > 0:
                status_msg += f" 跳过:{skipped_count}"
            if failed_count > 0:
                status_msg += f" 失败:{failed_count}"
            status_msg += f" | 文字行数:{total_lines}"
            if skipped_count > 0:
                status_msg += " | 💡跳过的图片可用其他识别模式"
            
            self.root.after(0, lambda m=status_msg: self.progress_label.config(text=m))
        
        except Exception as e:
            self.root.after(0, lambda: self.result_text.insert(tk.END, f"\n发生错误：{str(e)}\n"))
            self.root.after(0, lambda: messagebox.showerror("错误", f"发生错误：{str(e)}"))
            self.root.after(0, lambda: self.progress_label.config(text="✗ 处理失败"))
            self.root.after(0, lambda: self.ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.quick_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.general_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.select_btn.config(state=tk.NORMAL))

    def perform_quick_ocr(self):
        """执行快速 OCR 识别"""
        if not self.image_paths:
            messagebox.showwarning("警告", "请先选择图片文件！")
            return
        
        if not API_KEY or not SECRET_KEY:
            messagebox.showerror("错误", "请先在 .env 文件中配置 API_KEY 和 SECRET_KEY！")
            return
        
        self.ocr_btn.config(state=tk.DISABLED)
        self.quick_ocr_btn.config(state=tk.DISABLED)
        self.general_ocr_btn.config(state=tk.DISABLED)
        self.select_btn.config(state=tk.DISABLED)
        
        thread = threading.Thread(target=self._perform_quick_ocr_thread, daemon=True)
        thread.start()
    
    def _perform_quick_ocr_thread(self):
        """快速OCR识别线程"""
        try:
            self.root.after(0, lambda: self.result_text.delete(1.0, tk.END))
            self.all_results = []
            
            total = len(self.image_paths)
            
            for idx, image_path in enumerate(self.image_paths, 1):
                self.root.after(0, lambda i=idx, p=image_path: 
                    self.progress_label.config(text=f"快速识别中: {i}/{total} - {os.path.basename(p)}"))
                
                self.root.after(0, lambda: self.result_text.insert(tk.END, f"\n{'='*80}\n"))
                self.root.after(0, lambda i=idx, p=image_path: 
                    self.result_text.insert(tk.END, f"文件 {i}/{total}: {os.path.basename(p)}\n"))
                self.root.after(0, lambda: self.result_text.insert(tk.END, f"{'='*80}\n"))
                
                try:
                    img = Image.open(image_path)
                    width, height = img.size
                    
                    self.root.after(0, lambda w=width, h=height: 
                        self.result_text.insert(tk.END, f"图片尺寸: 宽{w} x 高{h}\n"))
                    
                    # 检查是否符合快速识别要求
                    width_in_basic = self.size_limits["basic_min_width"] <= width <= self.size_limits["basic_max_width"]
                    height_in_basic = self.size_limits["basic_min_height"] <= height <= self.size_limits["basic_max_height"]
                    meets_basic = width_in_basic and height_in_basic
                    
                    if not meets_basic:
                        bas_w_range = f"{self.size_limits['basic_min_width']}~{self.size_limits['basic_max_width']}"
                        bas_h_range = f"{self.size_limits['basic_min_height']}~{self.size_limits['basic_max_height']}"
                        self.root.after(0, lambda w=width, h=height, wr=bas_w_range, hr=bas_h_range: 
                            self.result_text.insert(tk.END, 
                                f"⚠️ 跳过：图片尺寸不符合要求\n"
                                f"   当前尺寸: 宽{w} x 高{h}\n"
                                f"   要求：宽度({wr})且高度({hr})都要在范围内\n"
                                f"   建议使用「高精度识别」按钮\n"))
                        
                        self.all_results.append({
                            'file': os.path.basename(image_path),
                            'path': image_path,
                            'lines': [],
                            'count': 0,
                            'skipped': True,
                            'reason': f'图片尺寸不符合要求（宽{width} x 高{height}）'
                        })
                        
                        self.root.after(0, lambda: self.result_text.see(tk.END))
                        continue
                    
                except Exception as e:
                    self.root.after(0, lambda err=str(e): 
                        self.result_text.insert(tk.END, f"⚠️ 无法读取图片尺寸: {err}\n"))
                
                result = ocr_image_basic(image_path)
                
                if "words_result" in result:
                    text_only_lines = []
                    for item in result["words_result"]:
                        words = item["words"]
                        text_only_lines.append(words)
                    
                    recognized_text = "\n".join(text_only_lines)
                    self.root.after(0, lambda t=recognized_text: 
                        self.result_text.insert(tk.END, t + "\n"))
                    
                    self.all_results.append({
                        'file': os.path.basename(image_path),
                        'path': image_path,
                        'lines': text_only_lines,
                        'count': len(text_only_lines)
                    })
                    
                    self.root.after(0, lambda c=len(text_only_lines): 
                        self.result_text.insert(tk.END, f"\n✓ 识别成功：{c} 行文字\n"))
                else:
                    self.root.after(0, lambda r=result: 
                        self.result_text.insert(tk.END, f"✗ 识别失败：{r}\n"))
                    self.all_results.append({
                        'file': os.path.basename(image_path),
                        'path': image_path,
                        'lines': [],
                        'count': 0,
                        'error': str(result)
                    })
                
                self.root.after(0, lambda: self.result_text.see(tk.END))
                
                if idx < total:
                    import time
                    time.sleep(0.5)
            
            success_count = sum(1 for r in self.all_results if r['count'] > 0)
            skipped_count = sum(1 for r in self.all_results if r.get('skipped', False))
            failed_count = total - success_count - skipped_count
            total_lines = sum(r['count'] for r in self.all_results)
            
            actual_processed = total - skipped_count
            if actual_processed > 0:
                self.record_ocr('basic', success_count, failed_count, total_lines)
                # 添加到历史记录（在主线程中执行）
                results_copy = [r.copy() for r in self.all_results]
                self.root.after(0, lambda: self.add_to_history('快速识别', results_copy))
            
            self.root.after(0, lambda: self.progress_label.config(text=f"✓ 完成！共处理 {total} 个文件"))
            self.root.after(0, lambda: self.export_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.copy_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.add_zeros_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.quick_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.select_btn.config(state=tk.NORMAL))
            
            status_msg = f"✓ 快速识别完成！总:{total} 成功:{success_count}"
            if skipped_count > 0:
                status_msg += f" 跳过:{skipped_count}"
            if failed_count > 0:
                status_msg += f" 失败:{failed_count}"
            status_msg += f" | 文字行数:{total_lines}"
            if skipped_count > 0:
                status_msg += " | 💡跳过的图片可用高精度识别"
            
            self.root.after(0, lambda m=status_msg: self.progress_label.config(text=m))
        
        except Exception as e:
            self.root.after(0, lambda: self.result_text.insert(tk.END, f"\n发生错误：{str(e)}\n"))
            self.root.after(0, lambda: messagebox.showerror("错误", f"发生错误：{str(e)}"))
            self.root.after(0, lambda: self.progress_label.config(text="✗ 处理失败"))
            self.root.after(0, lambda: self.ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.quick_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.general_ocr_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.select_btn.config(state=tk.NORMAL))
    
    def clear_result(self):
        """清空结果"""
        self.result_text.delete(1.0, tk.END)
        self.all_results = []
        self.progress_label.config(text="")
        self.export_btn.config(state=tk.DISABLED)
        self.copy_btn.config(state=tk.DISABLED)
        self.add_zeros_btn.config(state=tk.DISABLED)
    
    def copy_text(self):
        """复制识别的文字到剪贴板"""
        if not self.all_results:
            messagebox.showwarning("警告", "没有可复制的文字！")
            return
        
        try:
            all_lines = []
            for result in self.all_results:
                all_lines.extend(result['lines'])
            
            text_to_copy = "\n".join(all_lines)
            
            self.root.clipboard_clear()
            self.root.clipboard_append(text_to_copy)
            self.root.update()
            
            line_count = len(all_lines)
            char_count = len(text_to_copy)
            
            has_position = any('|' in line for line in all_lines)
            
            if has_position:
                format_info = "格式: 文字|top|left|height"
            else:
                format_info = "格式: 纯文字"
            
            self.progress_label.config(
                text=f"✓ 已复制到剪贴板！{format_info} | {line_count}行 {char_count}字符")
        
        except Exception as e:
            messagebox.showerror("错误", f"复制失败：{str(e)}")
    
    def add_zeros_to_lines(self):
        """在纯文字行后面添加|0|0（带位置信息的不改变）"""
        if not self.all_results:
            messagebox.showwarning("警告", "没有可处理的文字！")
            return
        
        try:
            # 统计处理的行数
            total_lines = 0
            modified_lines = 0
            skipped_lines = 0
            
            # 遍历所有结果
            for result in self.all_results:
                if result['lines']:
                    new_lines = []
                    for line in result['lines']:
                        total_lines += 1
                        # 如果行中已经有|符号，说明是带位置信息的格式，不改变
                        if '|' in line:
                            new_lines.append(line)
                            skipped_lines += 1
                        else:
                            # 纯文字，直接添加|0|0
                            new_line = f"{line}|0|0"
                            new_lines.append(new_line)
                            modified_lines += 1
                    
                    # 更新结果
                    result['lines'] = new_lines
            
            # 更新显示
            self.result_text.delete(1.0, tk.END)
            for result in self.all_results:
                self.result_text.insert(tk.END, f"\n{'='*80}\n")
                self.result_text.insert(tk.END, f"文件: {result['file']}\n")
                self.result_text.insert(tk.END, f"{'='*80}\n")
                
                if result['lines']:
                    for line in result['lines']:
                        self.result_text.insert(tk.END, line + "\n")
                    self.result_text.insert(tk.END, f"\n✓ 已处理：{len(result['lines'])} 行\n")
                else:
                    self.result_text.insert(tk.END, "无内容\n")
            
            # 显示处理结果
            if modified_lines > 0:
                self.progress_label.config(
                    text=f"✓ 已添加|0|0！处理 {modified_lines} 行，跳过 {skipped_lines} 行（已有位置信息）")
                
                messagebox.showinfo("处理完成", 
                    f"已在纯文字行后面添加|0|0\n\n"
                    f"总行数: {total_lines} 行\n"
                    f"已处理: {modified_lines} 行（纯文字）\n"
                    f"已跳过: {skipped_lines} 行（带位置信息）")
            else:
                self.progress_label.config(
                    text=f"✓ 无需处理！所有 {total_lines} 行都已有位置信息")
                
                messagebox.showinfo("无需处理", 
                    f"所有行都已包含位置信息，无需添加|0|0\n\n"
                    f"总行数: {total_lines} 行")
        
        except Exception as e:
            messagebox.showerror("错误", f"处理失败：{str(e)}")
    
    def show_context_menu(self, event):
        """显示右键菜单"""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
    
    def copy_selected(self):
        """复制选中的文字"""
        try:
            selected_text = self.result_text.get(tk.SEL_FIRST, tk.SEL_LAST)
            if selected_text:
                self.root.clipboard_clear()
                self.root.clipboard_append(selected_text)
                self.root.update()
                self.progress_label.config(text=f"✓ 已复制 {len(selected_text)} 个字符")
        except tk.TclError:
            messagebox.showwarning("提示", "请先选中要复制的文字！")
    
    def copy_all_text(self):
        """复制全部文字和位置信息"""
        try:
            all_lines = []
            for result in self.all_results:
                all_lines.extend(result['lines'])
            
            if all_lines:
                text_to_copy = "\n".join(all_lines)
                self.root.clipboard_clear()
                self.root.clipboard_append(text_to_copy)
                self.root.update()
                
                line_count = len(all_lines)
                self.progress_label.config(text=f"✓ 已复制 {line_count} 行文字和位置信息")
            else:
                messagebox.showwarning("提示", "没有可复制的文字！")
        except Exception as e:
            messagebox.showerror("错误", f"复制失败：{str(e)}")
    
    def select_all(self):
        """全选文字"""
        self.result_text.tag_add(tk.SEL, "1.0", tk.END)
        self.result_text.mark_set(tk.INSERT, "1.0")
        self.result_text.see(tk.INSERT)
    
    def load_window_config(self):
        """加载主窗口配置"""
        try:
            config = self.store.get('window_config', {})
            if config:
                width = config.get('width', 1300)
                height = config.get('height', 900)
                x = config.get('x', None)
                y = config.get('y', None)
                
                # 应用窗口尺寸和位置
                if x is not None and y is not None:
                    self.root.geometry(f"{width}x{height}+{x}+{y}")
                else:
                    self.root.geometry(f"{width}x{height}")
                
                print(f"✓ 已加载窗口配置：{width}x{height}")
            else:
                # 默认尺寸
                self.root.geometry("1300x900")
                print("✓ 使用默认窗口尺寸")
        except Exception as e:
            print(f"⚠️ 加载窗口配置失败: {e}")
            self.root.geometry("1300x900")
    
    def save_window_config(self):
        """保存主窗口配置"""
        try:
            # 获取当前窗口尺寸和位置
            geometry = self.root.geometry()
            # 格式：widthxheight+x+y
            parts = geometry.replace('+', 'x').replace('-', 'x').split('x')
            
            if len(parts) >= 2:
                config = {
                    'width': int(parts[0]),
                    'height': int(parts[1])
                }
                
                # 保存位置（如果有）
                if len(parts) >= 4:
                    config['x'] = int(parts[2])
                    config['y'] = int(parts[3])
                
                self.store.set('window_config', config)
                print(f"✓ 已保存窗口配置：{config['width']}x{config['height']}")
        except Exception as e:
            print(f"⚠️ 保存窗口配置失败: {e}")
    
    def load_popup_config(self, window_name):
        """加载弹出窗口配置"""
        try:
            all_configs = self.store.get('popup_windows', {})
            return all_configs.get(window_name, None)
        except Exception as e:
            print(f"⚠️ 加载弹出窗口配置失败: {e}")
            return None
    
    def save_popup_config(self, window_name, window):
        """保存弹出窗口配置"""
        try:
            all_configs = self.store.get('popup_windows', {})
            
            # 获取窗口尺寸和位置
            geometry = window.geometry()
            parts = geometry.replace('+', 'x').replace('-', 'x').split('x')
            
            if len(parts) >= 2:
                config = {
                    'width': int(parts[0]),
                    'height': int(parts[1])
                }
                
                if len(parts) >= 4:
                    config['x'] = int(parts[2])
                    config['y'] = int(parts[3])
                
                # 更新配置
                all_configs[window_name] = config
                self.store.set('popup_windows', all_configs)
                
                print(f"✓ 已保存 {window_name} 窗口配置：{config['width']}x{config['height']}")
        except Exception as e:
            print(f"⚠️ 保存弹出窗口配置失败: {e}")
    
    def create_popup_window(self, parent, title, window_name, default_width=500, default_height=400):
        """创建带配置保存功能的弹出窗口"""
        popup = tk.Toplevel(parent)
        popup.title(title)
        popup.transient(parent)
        popup.grab_set()
        
        # 加载保存的配置
        config = self.load_popup_config(window_name)
        
        if config:
            width = config.get('width', default_width)
            height = config.get('height', default_height)
            x = config.get('x', None)
            y = config.get('y', None)
            
            if x is not None and y is not None:
                popup.geometry(f"{width}x{height}+{x}+{y}")
            else:
                popup.geometry(f"{width}x{height}")
                # 居中显示
                popup.update_idletasks()
                x = (popup.winfo_screenwidth() // 2) - (width // 2)
                y = (popup.winfo_screenheight() // 2) - (height // 2)
                popup.geometry(f"{width}x{height}+{x}+{y}")
        else:
            # 使用默认尺寸并居中
            popup.geometry(f"{default_width}x{default_height}")
            popup.update_idletasks()
            x = (popup.winfo_screenwidth() // 2) - (default_width // 2)
            y = (popup.winfo_screenheight() // 2) - (default_height // 2)
            popup.geometry(f"{default_width}x{default_height}+{x}+{y}")
        
        # 设置最小尺寸
        popup.minsize(default_width, default_height)
        
        # 绑定关闭事件，保存配置
        def on_popup_close():
            self.save_popup_config(window_name, popup)
            popup.destroy()
        
        popup.protocol("WM_DELETE_WINDOW", on_popup_close)
        
        return popup
    
    def on_closing(self):
        """窗口关闭时的处理"""
        # 保存窗口配置
        self.save_window_config()
        # 关闭窗口
        self.root.destroy()
    
    def load_history_limit(self):
        """加载历史记录数量限制"""
        try:
            self.history_limit = self.store.get('history_limit', 100)
            print(f"✓ 历史记录限制：{self.history_limit} 条")
        except Exception as e:
            print(f"⚠️ 加载历史记录限制失败: {e}")
            self.history_limit = 100
    
    def save_history_limit(self):
        """保存历史记录数量限制"""
        try:
            self.store.set('history_limit', self.history_limit)
            print(f"✓ 已保存历史记录限制：{self.history_limit} 条")
        except Exception as e:
            print(f"⚠️ 保存历史记录限制失败: {e}")
    
    def load_history(self):
        """加载历史记录"""
        try:
            self.history_data = self.store.get('history', [])
            print(f"✓ 已加载历史记录：{len(self.history_data)} 条")
        except Exception as e:
            print(f"⚠️ 加载历史记录失败: {e}")
            self.history_data = []
    
    def save_history(self):
        """保存历史记录"""
        try:
            self.store.set('history', self.history_data)
            print(f"✓ 已保存历史记录：{len(self.history_data)} 条")
        except Exception as e:
            print(f"⚠️ 保存历史记录失败: {e}")
    
    def add_to_history(self, ocr_type, results):
        """添加识别结果到历史记录"""
        try:
            print(f"📝 开始添加历史记录：{ocr_type}, 结果数量：{len(results)}")
            
            # 过滤掉跳过的结果
            valid_results = [r for r in results if r.get('count', 0) > 0 and not r.get('skipped', False)]
            
            if not valid_results:
                print("⚠️ 没有有效的识别结果，跳过保存历史记录")
                return
            
            history_item = {
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'type': ocr_type,
                'file_count': len(valid_results),
                'total_lines': sum(r['count'] for r in valid_results),
                'files': []
            }
            
            # 添加文件信息（保存所有内容）
            for result in valid_results:
                file_info = {
                    'name': result['file'],
                    'lines': result['count'],
                    'content': result['lines']  # 保存所有行
                }
                history_item['files'].append(file_info)
                print(f"  - {result['file']}: {result['count']} 行")
            
            # 添加到历史记录列表开头
            self.history_data.insert(0, history_item)
            
            # 限制历史记录数量
            if len(self.history_data) > self.history_limit:
                self.history_data = self.history_data[:self.history_limit]
            
            # 保存到文件
            self.save_history()
            print(f"✓ 历史记录添加成功：{history_item['file_count']} 个文件，{history_item['total_lines']} 行")
        except Exception as e:
            print(f"⚠️ 添加历史记录失败: {e}")
            import traceback
            traceback.print_exc()
    
    def load_stats(self):
        """加载统计数据"""
        try:
            self.stats = self.store.get('stats', {})
            print(f"✓ 已加载统计数据：{len(self.stats)} 天的记录")
        except Exception as e:
            print(f"⚠️ 加载统计数据失败: {e}")
            self.stats = {}
    
    def load_size_limits(self):
        """加载尺寸限制配置"""
        try:
            saved_limits = self.store.get('size_limits', {})
            if saved_limits:
                self.size_limits.update(saved_limits)
                print(f"✓ 已加载尺寸限制配置: {saved_limits}")
            # 如果界已经创建，立即更新显示
            if hasattr(self, 'size_hint_label'):
                self.update_size_hint_display()
        except Exception as e:
            print(f"⚠️ 加载尺寸限制配置失败: {e}")
    
    def save_size_limits(self):
        """保存尺寸限制配置"""
        try:
            self.store.set('size_limits', self.size_limits)
            print(f"✓ 尺寸限制配置已保存")
            # 保存后立即更新界面显示
            self.update_size_hint_display()
        except Exception as e:
            print(f"⚠️ 保存尺寸限制配置失败: {e}")
    
    def load_font_config(self):
        """加载字号配置"""
        try:
            config = self.store.get('font_config', {})
            if config:
                self.current_font_size = config.get('font_size', 11)
            else:
                self.current_font_size = 11
            print(f"✓ 已加载字号配置: {self.current_font_size}")
        except Exception as e:
            print(f"⚠️ 加载字号配置失败: {e}")
            self.current_font_size = 11
    
    def save_font_config(self):
        """保存字号配置"""
        try:
            config = {'font_size': self.current_font_size}
            self.store.set('font_config', config)
            print(f"✓ 字号配置已保存: {self.current_font_size}")
        except Exception as e:
            print(f"⚠️ 保存字号配置失败: {e}")
    
    def update_size_hint_display(self):
        """更新界面上的尺寸提示信息"""
        try:
            if hasattr(self, 'size_hint_label'):
                if self.size_limit_unlocked:
                    bas_range = f"{self.size_limits['basic_min_width']}~{self.size_limits['basic_max_width']}x{self.size_limits['basic_min_height']}~{self.size_limits['basic_max_height']}"
                    self.size_hint_label.config(text=f"💡 高精度(已解锁限制) | 快速({bas_range})")
                else:
                    acc_range = f"{self.size_limits['accurate_min_width']}~{self.size_limits['accurate_max_width']}x{self.size_limits['accurate_min_height']}~{self.size_limits['accurate_max_height']}"
                    bas_range = f"{self.size_limits['basic_min_width']}~{self.size_limits['basic_max_width']}x{self.size_limits['basic_min_height']}~{self.size_limits['basic_max_height']}"
                    self.size_hint_label.config(text=f"💡 高精度({acc_range}) | 快速({bas_range})")
        except Exception as e:
            print(f"⚠️ 更新界面提示信息失败: {e}")
    
    def show_size_settings(self):
        """显示尺寸设置窗口（需要解锁）"""
        # 检查是否已解锁
        if not self.size_limit_unlocked:
            messagebox.showwarning("需要解锁", 
                "尺寸设置需要先解锁！\n\n"
                "请点击「🔒 解锁限制」按钮并输入密码")
            return
        
        settings_window = self.create_popup_window(self.root, "图片尺寸限制设置", "size_limit_settings", 600, 700)
        
        tk.Label(settings_window, text="⚙️ 图片尺寸限制设置", 
                font=("Arial", 14, "bold")).pack(pady=15)
        
        tk.Label(settings_window, text="设置OCR识别的图片尺寸范围要求", 
                fg="gray").pack(pady=5)
        
        # 设置框架
        settings_frame = tk.Frame(settings_window)
        settings_frame.pack(pady=20, padx=30, fill=tk.BOTH, expand=True)
        
        # 高精度识别设置
        tk.Label(settings_frame, text="高精度识别范围（适合大图）：", 
                font=("Arial", 11, "bold")).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        tk.Label(settings_frame, text="最小宽度 (px):").grid(row=1, column=0, sticky=tk.W, pady=5)
        acc_min_width_var = tk.StringVar(value=str(self.size_limits['accurate_min_width']))
        acc_min_width_entry = tk.Entry(settings_frame, textvariable=acc_min_width_var, width=15)
        acc_min_width_entry.grid(row=1, column=1, sticky=tk.W, pady=5, padx=10)
        
        tk.Label(settings_frame, text="最大宽度 (px):").grid(row=2, column=0, sticky=tk.W, pady=5)
        acc_max_width_var = tk.StringVar(value=str(self.size_limits['accurate_max_width']))
        acc_max_width_entry = tk.Entry(settings_frame, textvariable=acc_max_width_var, width=15)
        acc_max_width_entry.grid(row=2, column=1, sticky=tk.W, pady=5, padx=10)
        
        tk.Label(settings_frame, text="最小高度 (px):").grid(row=3, column=0, sticky=tk.W, pady=5)
        acc_min_height_var = tk.StringVar(value=str(self.size_limits['accurate_min_height']))
        acc_min_height_entry = tk.Entry(settings_frame, textvariable=acc_min_height_var, width=15)
        acc_min_height_entry.grid(row=3, column=1, sticky=tk.W, pady=5, padx=10)
        
        tk.Label(settings_frame, text="最大高度 (px):").grid(row=4, column=0, sticky=tk.W, pady=5)
        acc_max_height_var = tk.StringVar(value=str(self.size_limits['accurate_max_height']))
        acc_max_height_entry = tk.Entry(settings_frame, textvariable=acc_max_height_var, width=15)
        acc_max_height_entry.grid(row=4, column=1, sticky=tk.W, pady=5, padx=10)
        
        # 快速识别设置
        tk.Label(settings_frame, text="快速识别范围（适合小图）：", 
                font=("Arial", 11, "bold")).grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        tk.Label(settings_frame, text="最小宽度 (px):").grid(row=6, column=0, sticky=tk.W, pady=5)
        bas_min_width_var = tk.StringVar(value=str(self.size_limits['basic_min_width']))
        bas_min_width_entry = tk.Entry(settings_frame, textvariable=bas_min_width_var, width=15)
        bas_min_width_entry.grid(row=6, column=1, sticky=tk.W, pady=5, padx=10)
        
        tk.Label(settings_frame, text="最大宽度 (px):").grid(row=7, column=0, sticky=tk.W, pady=5)
        bas_max_width_var = tk.StringVar(value=str(self.size_limits['basic_max_width']))
        bas_max_width_entry = tk.Entry(settings_frame, textvariable=bas_max_width_var, width=15)
        bas_max_width_entry.grid(row=7, column=1, sticky=tk.W, pady=5, padx=10)
        
        tk.Label(settings_frame, text="最小高度 (px):").grid(row=8, column=0, sticky=tk.W, pady=5)
        bas_min_height_var = tk.StringVar(value=str(self.size_limits['basic_min_height']))
        bas_min_height_entry = tk.Entry(settings_frame, textvariable=bas_min_height_var, width=15)
        bas_min_height_entry.grid(row=8, column=1, sticky=tk.W, pady=5, padx=10)
        
        tk.Label(settings_frame, text="最大高度 (px):").grid(row=9, column=0, sticky=tk.W, pady=5)
        bas_max_height_var = tk.StringVar(value=str(self.size_limits['basic_max_height']))
        bas_max_height_entry = tk.Entry(settings_frame, textvariable=bas_max_height_var, width=15)
        bas_max_height_entry.grid(row=9, column=1, sticky=tk.W, pady=5, padx=10)
        
        # 提示信息
        hint_text = "💡 提示：修改后将立即生效，并保存到配置文件\n范围格式：最小值 ≤ 图片尺寸 ≤ 最大值"
        tk.Label(settings_frame, text=hint_text, fg="blue", justify=tk.LEFT,
                font=("Arial", 9)).grid(row=10, column=0, columnspan=2, pady=15)
        
        def save_settings():
            try:
                # 验证输入
                acc_min_w = int(acc_min_width_var.get())
                acc_max_w = int(acc_max_width_var.get())
                acc_min_h = int(acc_min_height_var.get())
                acc_max_h = int(acc_max_height_var.get())
                bas_min_w = int(bas_min_width_var.get())
                bas_max_w = int(bas_max_width_var.get())
                bas_min_h = int(bas_min_height_var.get())
                bas_max_h = int(bas_max_height_var.get())
                
                # 验证范围合理性
                if acc_min_w < 0 or acc_max_w < 0 or acc_min_h < 0 or acc_max_h < 0:
                    messagebox.showerror("错误", "高精度识别尺寸不能为负数！")
                    return
                
                if bas_min_w < 0 or bas_max_w < 0 or bas_min_h < 0 or bas_max_h < 0:
                    messagebox.showerror("错误", "快速识别尺寸不能为负数！")
                    return
                
                if acc_min_w > acc_max_w or acc_min_h > acc_max_h:
                    messagebox.showerror("错误", "高精度识别：最小值不能大于最大值！")
                    return
                
                if bas_min_w > bas_max_w or bas_min_h > bas_max_h:
                    messagebox.showerror("错误", "快速识别：最小值不能大于最大值！")
                    return
                
                # 保存设置
                self.size_limits['accurate_min_width'] = acc_min_w
                self.size_limits['accurate_max_width'] = acc_max_w
                self.size_limits['accurate_min_height'] = acc_min_h
                self.size_limits['accurate_max_height'] = acc_max_h
                self.size_limits['basic_min_width'] = bas_min_w
                self.size_limits['basic_max_width'] = bas_max_w
                self.size_limits['basic_min_height'] = bas_min_h
                self.size_limits['basic_max_height'] = bas_max_h
                
                self.save_size_limits()
                
                # 更新提示信息
                if hasattr(self, 'size_hint_label'):
                    if self.size_limit_unlocked:
                        self.size_hint_label.config(text=f"💡 高精度(已解锁限制) | 快速({bas_min_w}~{bas_max_w}x{bas_min_h}~{bas_max_h})")
                    else:
                        self.size_hint_label.config(text=f"💡 高精度({acc_min_w}~{acc_max_w}x{acc_min_h}~{acc_max_h}) | 快速({bas_min_w}~{bas_max_w}x{bas_min_h}~{bas_max_h})")
                else:
                    # 兼容旧版本的更新方式
                    for widget in self.progress_frame.winfo_children():
                        if isinstance(widget, tk.Label) and "高精度" in widget.cget("text"):
                            if self.size_limit_unlocked:
                                widget.config(text=f"💡 高精度(已解锁限制) | 快速({bas_min_w}~{bas_max_w}x{bas_min_h}~{bas_max_h})")
                            else:
                                widget.config(text=f"💡 高精度({acc_min_w}~{acc_max_w}x{acc_min_h}~{acc_max_h}) | 快速({bas_min_w}~{bas_max_w}x{bas_min_h}~{bas_max_h})")
                
                # 保存窗口尺寸配置
                self.save_popup_config("size_limit_settings", settings_window)
                
                settings_window.destroy()
                messagebox.showinfo("成功", "尺寸限制设置已保存！")
                
                # 如果已选择文件，重新检查
                if self.image_paths:
                    if len(self.image_paths) == 1:
                        self.select_file_internal(self.image_paths[0])
                    else:
                        self.batch_select_files_internal(self.image_paths)
            
            except ValueError:
                messagebox.showerror("错误", "请输入有效的数字！")
        
        def reset_defaults():
            acc_min_width_var.set("3500")
            acc_max_width_var.set("15000")
            acc_min_height_var.set("4000")
            acc_max_height_var.set("15000")
            bas_min_width_var.set("0")
            bas_max_width_var.set("8100")
            bas_min_height_var.set("0")
            bas_max_height_var.set("3000")
        
        # 按钮
        btn_frame = tk.Frame(settings_window)
        btn_frame.pack(pady=15)
        
        tk.Button(btn_frame, text="保存", command=save_settings,
                 bg="#4CAF50", fg="white", padx=30, pady=8).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="恢复默认", command=reset_defaults,
                 bg="#FF9800", fg="white", padx=30, pady=8).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="取消", command=settings_window.destroy,
                 bg="#757575", fg="white", padx=30, pady=8).pack(side=tk.LEFT, padx=5)
    
    def save_stats(self):
        """保存统计数据"""
        try:
            self.store.set('stats', self.stats)
        except Exception as e:
            print(f"⚠️ 保存统计数据失败: {e}")
            messagebox.showerror("错误", f"统计数据保存失败：{e}")
    
    def record_ocr(self, ocr_type, success_count, failed_count, lines):
        """记录识别统计"""
        today = datetime.now().strftime("%Y-%m-%d")
        
        if today not in self.stats:
            self.stats[today] = {
                'accurate': {'count': 0, 'success': 0, 'failed': 0, 'skipped': 0, 'lines': 0},
                'basic': {'count': 0, 'success': 0, 'failed': 0, 'lines': 0},
                'general': {'count': 0, 'success': 0, 'failed': 0, 'lines': 0}
            }
        
        # 确保所有模式都存在
        if 'general' not in self.stats[today]:
            self.stats[today]['general'] = {'count': 0, 'success': 0, 'failed': 0, 'lines': 0}
        
        if 'accurate' not in self.stats[today]:
            self.stats[today]['accurate'] = {'count': 0, 'success': 0, 'failed': 0, 'skipped': 0, 'lines': 0}
        
        if 'basic' not in self.stats[today]:
            self.stats[today]['basic'] = {'count': 0, 'success': 0, 'failed': 0, 'lines': 0}
        
        self.stats[today][ocr_type]['count'] += 1
        self.stats[today][ocr_type]['success'] += success_count
        self.stats[today][ocr_type]['failed'] += failed_count
        self.stats[today][ocr_type]['lines'] += lines
        
        self.save_stats()

    
    def show_stats(self):
        """显示统计信息"""
        stats_window = self.create_popup_window(self.root, "识别统计", "stats_window", 1100, 850)
        
        tk.Label(stats_window, text="📊 OCR 识别统计", 
                font=("Arial", 16, "bold")).pack(pady=15)
        
        # 创建选项卡
        from tkinter import ttk
        notebook = ttk.Notebook(stats_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 总计选项卡
        total_tab = tk.Frame(notebook)
        notebook.add(total_tab, text="📈 总计统计")
        
        # 按日统计选项卡
        daily_tab = tk.Frame(notebook)
        notebook.add(daily_tab, text="📅 按日统计")
        
        # 按月统计选项卡
        monthly_tab = tk.Frame(notebook)
        notebook.add(monthly_tab, text="📊 按月统计")
        
        # === 总计统计 ===
        self._show_total_stats(total_tab)
        
        # === 按日统计 ===
        self._show_daily_stats(daily_tab)
        
        # === 按月统计 ===
        self._show_monthly_stats(monthly_tab)
        
        # 按钮
        btn_frame = tk.Frame(stats_window)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="关闭", command=stats_window.destroy,
                 bg="#757575", fg="white", padx=20, pady=8).pack()
    
    def show_history(self):
        """显示历史记录"""
        history_window = self.create_popup_window(self.root, "识别历史记录", "history_window", 1200, 800)
        
        tk.Label(history_window, text="📜 OCR 识别历史记录", 
                font=("Arial", 16, "bold")).pack(pady=15)
        
        # 创建表格框架
        from tkinter import ttk
        
        table_frame = tk.Frame(history_window)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 创建滚动条
        scrollbar = tk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建表格
        columns = ("时间", "类型", "文件数", "总行数", "操作")
        # 使用自定义样式 History.Treeview，避免影响全局 Treeview 样式
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", 
                            yscrollcommand=scrollbar.set, height=25, style="History.Treeview")
        
        # 设置列标题
        tree.heading("时间", text="识别时间")
        tree.heading("类型", text="识别类型")
        tree.heading("文件数", text="文件数")
        tree.heading("总行数", text="总行数")
        tree.heading("操作", text="操作")
        
        # 设置列宽度
        tree.column("时间", width=180, anchor=tk.CENTER)
        tree.column("类型", width=120, anchor=tk.CENTER)
        tree.column("文件数", width=100, anchor=tk.CENTER)
        tree.column("总行数", width=100, anchor=tk.CENTER)
        tree.column("操作", width=150, anchor=tk.CENTER)
        
        scrollbar.config(command=tree.yview)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 配置样式 (使用自定义样式名)
        style = ttk.Style()
        style.configure("History.Treeview", font=("Microsoft YaHei", 10), rowheight=30)
        style.configure("History.Treeview.Heading", font=("Microsoft YaHei", 11, "bold"))
        
        # 插入数据
        for idx, item in enumerate(self.history_data):
            tree.insert("", tk.END, 
                       values=(item['timestamp'], 
                              item['type'], 
                              item['file_count'], 
                              item['total_lines'],
                              "查看详情"),
                       tags=(f"item_{idx}",))
        
        # 设置行颜色
        for i in range(len(self.history_data)):
            if i % 2 == 0:
                tree.tag_configure(f"item_{i}", background="#F5F5F5")
        
        # 双击查看详情
        def on_double_click(event):
            item = tree.selection()
            if item:
                item_values = tree.item(item[0])['values']
                timestamp = item_values[0]
                
                # 查找对应的历史记录
                for history_item in self.history_data:
                    if history_item['timestamp'] == timestamp:
                        self.show_history_detail(history_item)
                        break
        
        tree.bind("<Double-1>", on_double_click)
        
        # 按钮框架
        btn_frame = tk.Frame(history_window)
        btn_frame.pack(pady=10)
        
        def clear_history():
            if messagebox.askyesno("确认", "确定要清空所有历史记录吗？\n此操作不可恢复！"):
                self.history_data = []
                self.save_history()
                history_window.destroy()
                messagebox.showinfo("成功", "历史记录已清空")
        
        def set_history_limit():
            """设置历史记录数量限制"""
            limit_window = self.create_popup_window(history_window, "历史记录数量设置", "history_limit_settings", 450, 300)
            
            tk.Label(limit_window, text="📝 历史记录数量设置", 
                    font=("Arial", 14, "bold")).pack(pady=20)
            
            tk.Label(limit_window, text=f"当前限制：{self.history_limit} 条", 
                    fg="blue", font=("Arial", 11)).pack(pady=10)
            
            tk.Label(limit_window, text="设置新的历史记录数量限制：", 
                    font=("Arial", 10)).pack(pady=10)
            
            # 输入框
            limit_var = tk.StringVar(value=str(self.history_limit))
            limit_entry = tk.Entry(limit_window, textvariable=limit_var, 
                                  font=("Arial", 12), width=15, justify=tk.CENTER)
            limit_entry.pack(pady=10)
            limit_entry.focus_set()
            
            # 快捷按钮
            quick_frame = tk.Frame(limit_window)
            quick_frame.pack(pady=10)
            
            tk.Label(quick_frame, text="快捷设置：", font=("Arial", 9)).pack(side=tk.LEFT, padx=5)
            
            for value in [50, 100, 200, 500, 1000]:
                tk.Button(quick_frame, text=str(value), 
                         command=lambda v=value: limit_var.set(str(v)),
                         bg="#2196F3", fg="white", padx=10, pady=3, 
                         font=("Arial", 9)).pack(side=tk.LEFT, padx=2)
            
            hint_text = "💡 提示：\n• 设置为0表示不限制\n• 超出限制的旧记录会被自动删除"
            tk.Label(limit_window, text=hint_text, fg="gray", 
                    font=("Arial", 9), justify=tk.LEFT).pack(pady=10)
            
            def save_limit():
                try:
                    new_limit = int(limit_var.get())
                    if new_limit < 0:
                        messagebox.showerror("错误", "数量不能为负数！")
                        return
                    
                    old_limit = self.history_limit
                    self.history_limit = new_limit
                    self.save_history_limit()
                    
                    # 如果新限制小于当前记录数，裁剪历史记录
                    if new_limit > 0 and len(self.history_data) > new_limit:
                        removed_count = len(self.history_data) - new_limit
                        self.history_data = self.history_data[:new_limit]
                        self.save_history()
                        messagebox.showinfo("成功", 
                            f"历史记录限制已更新！\n\n"
                            f"旧限制：{old_limit} 条\n"
                            f"新限制：{new_limit} 条\n"
                            f"已删除：{removed_count} 条旧记录")
                    else:
                        limit_text = "不限制" if new_limit == 0 else f"{new_limit} 条"
                        messagebox.showinfo("成功", 
                            f"历史记录限制已更新！\n\n"
                            f"旧限制：{old_limit} 条\n"
                            f"新限制：{limit_text}")
                    
                    limit_window.destroy()
                    history_window.destroy()
                    self.show_history()
                
                except ValueError:
                    messagebox.showerror("错误", "请输入有效的数字！")
            
            btn_frame2 = tk.Frame(limit_window)
            btn_frame2.pack(pady=15)
            
            tk.Button(btn_frame2, text="保存", command=save_limit,
                     bg="#4CAF50", fg="white", padx=25, pady=8).pack(side=tk.LEFT, padx=5)
            
            tk.Button(btn_frame2, text="取消", command=limit_window.destroy,
                     bg="#757575", fg="white", padx=25, pady=8).pack(side=tk.LEFT, padx=5)
            
            limit_entry.bind("<Return>", lambda e: save_limit())
        
        tk.Button(btn_frame, text="数量设置", command=set_history_limit,
                 bg="#2196F3", fg="white", padx=20, pady=8).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="清空历史", command=clear_history,
                 bg="#F44336", fg="white", padx=20, pady=8).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="关闭", command=history_window.destroy,
                 bg="#757575", fg="white", padx=20, pady=8).pack(side=tk.LEFT, padx=5)
        
        # 显示统计信息
        limit_text = "不限制" if self.history_limit == 0 else f"{self.history_limit} 条"
        info_text = f"共 {len(self.history_data)} 条历史记录 | 限制: {limit_text}"
        if self.history_data:
            total_files = sum(item['file_count'] for item in self.history_data)
            total_lines = sum(item['total_lines'] for item in self.history_data)
            info_text += f" | 总文件数: {total_files} | 总行数: {total_lines}"
        
        tk.Label(history_window, text=info_text, fg="gray", font=("Arial", 10)).pack(pady=5)
    
    def show_api_key_settings(self):
        """显示API密钥设置窗口"""
        settings_window = self.create_popup_window(self.root, "API密钥设置", "api_key_settings", 700, 700)
        
        tk.Label(settings_window, text="🔑 百度OCR API密钥设置", 
                font=("Arial", 14, "bold")).pack(pady=15)
        
        tk.Label(settings_window, text="修改后将自动保存到 .env 文件", 
                fg="gray", font=("Arial", 10)).pack(pady=5)
        
        # 设置框架
        settings_frame = tk.Frame(settings_window)
        settings_frame.pack(pady=20, padx=30, fill=tk.BOTH, expand=True)
        
        # 高精度识别密钥
        tk.Label(settings_frame, text="高精度识别密钥：", 
                font=("Arial", 11, "bold"), fg="#2196F3").grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        tk.Label(settings_frame, text="API Key:").grid(row=1, column=0, sticky=tk.W, pady=5)
        api_key_var = tk.StringVar(value=API_KEY)
        api_key_entry = tk.Entry(settings_frame, textvariable=api_key_var, width=50, font=("Arial", 10))
        api_key_entry.grid(row=1, column=1, sticky=tk.W, pady=5, padx=10)
        
        tk.Label(settings_frame, text="Secret Key:").grid(row=2, column=0, sticky=tk.W, pady=5)
        secret_key_var = tk.StringVar(value=SECRET_KEY)
        secret_key_entry = tk.Entry(settings_frame, textvariable=secret_key_var, width=50, font=("Arial", 10))
        secret_key_entry.grid(row=2, column=1, sticky=tk.W, pady=5, padx=10)
        
        # 分隔线
        tk.Frame(settings_frame, height=2, bg="gray").grid(row=3, column=0, columnspan=2, sticky=tk.EW, pady=15)
        
        # 快速识别密钥
        tk.Label(settings_frame, text="快速识别密钥（可选，留空则使用高精度密钥）：", 
                font=("Arial", 11, "bold"), fg="#00BCD4").grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        tk.Label(settings_frame, text="API Key:").grid(row=5, column=0, sticky=tk.W, pady=5)
        api_key_basic_var = tk.StringVar(value=API_KEY_BASIC if API_KEY_BASIC != API_KEY else "")
        api_key_basic_entry = tk.Entry(settings_frame, textvariable=api_key_basic_var, width=50, font=("Arial", 10))
        api_key_basic_entry.grid(row=5, column=1, sticky=tk.W, pady=5, padx=10)
        
        tk.Label(settings_frame, text="Secret Key:").grid(row=6, column=0, sticky=tk.W, pady=5)
        secret_key_basic_var = tk.StringVar(value=SECRET_KEY_BASIC if SECRET_KEY_BASIC != SECRET_KEY else "")
        secret_key_basic_entry = tk.Entry(settings_frame, textvariable=secret_key_basic_var, width=50, font=("Arial", 10))
        secret_key_basic_entry.grid(row=6, column=1, sticky=tk.W, pady=5, padx=10)
        
        # 分隔线
        tk.Frame(settings_frame, height=2, bg="gray").grid(row=7, column=0, columnspan=2, sticky=tk.EW, pady=15)
        
        # 通用识别密钥
        tk.Label(settings_frame, text="通用识别密钥（可选，留空则使用快速识别密钥）：", 
                font=("Arial", 11, "bold"), fg="#9C27B0").grid(row=8, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        tk.Label(settings_frame, text="API Key:").grid(row=9, column=0, sticky=tk.W, pady=5)
        api_key_general_var = tk.StringVar(value=API_KEY_GENERAL if API_KEY_GENERAL != API_KEY_BASIC else "")
        api_key_general_entry = tk.Entry(settings_frame, textvariable=api_key_general_var, width=50, font=("Arial", 10))
        api_key_general_entry.grid(row=9, column=1, sticky=tk.W, pady=5, padx=10)
        
        tk.Label(settings_frame, text="Secret Key:").grid(row=10, column=0, sticky=tk.W, pady=5)
        secret_key_general_var = tk.StringVar(value=SECRET_KEY_GENERAL if SECRET_KEY_GENERAL != SECRET_KEY_BASIC else "")
        secret_key_general_entry = tk.Entry(settings_frame, textvariable=secret_key_general_var, width=50, font=("Arial", 10))
        secret_key_general_entry.grid(row=10, column=1, sticky=tk.W, pady=5, padx=10)
        
        # 提示信息
        hint_text = "💡 提示：\n• 高精度识别密钥为必填项\n• 快速识别密钥可选，留空则使用高精度密钥\n• 通用识别密钥可选，留空则使用快速识别密钥\n• 修改后立即生效，无需重启程序"
        tk.Label(settings_frame, text=hint_text, fg="blue", justify=tk.LEFT,
                font=("Arial", 9)).grid(row=11, column=0, columnspan=2, pady=15, sticky=tk.W)
        
        def save_api_keys():
            try:
                new_api_key = api_key_var.get().strip()
                new_secret_key = secret_key_var.get().strip()
                new_api_key_basic = api_key_basic_var.get().strip()
                new_secret_key_basic = secret_key_basic_var.get().strip()
                new_api_key_general = api_key_general_var.get().strip()
                new_secret_key_general = secret_key_general_var.get().strip()
                
                # 验证必填项
                if not new_api_key or not new_secret_key:
                    messagebox.showerror("错误", "高精度识别的API Key和Secret Key不能为空！")
                    return
                
                # 更新全局变量
                global API_KEY, SECRET_KEY, API_KEY_BASIC, SECRET_KEY_BASIC, API_KEY_GENERAL, SECRET_KEY_GENERAL
                API_KEY = new_api_key
                SECRET_KEY = new_secret_key
                API_KEY_BASIC = new_api_key_basic if new_api_key_basic else new_api_key
                SECRET_KEY_BASIC = new_secret_key_basic if new_secret_key_basic else new_secret_key
                API_KEY_GENERAL = new_api_key_general if new_api_key_general else API_KEY_BASIC
                SECRET_KEY_GENERAL = new_secret_key_general if new_secret_key_general else SECRET_KEY_BASIC
                
                # 保存到.env文件
                env_path = Path(__file__).parent / '.env'
                env_lines = []
                
                # 读取现有的.env文件（如果存在）
                existing_keys = set()
                if env_path.exists():
                    with open(env_path, 'r', encoding='utf-8') as f:
                        for line in f:
                            line = line.strip()
                            if line and not line.startswith('#') and '=' in line:
                                key = line.split('=', 1)[0].strip()
                                if key not in ['BAIDU_API_KEY', 'BAIDU_SECRET_KEY', 
                                             'BAIDU_API_KEY_BASIC', 'BAIDU_SECRET_KEY_BASIC',
                                             'BAIDU_API_KEY_GENERAL', 'BAIDU_SECRET_KEY_GENERAL']:
                                    env_lines.append(line)
                
                # 添加新的密钥
                env_lines.append(f"BAIDU_API_KEY={new_api_key}")
                env_lines.append(f"BAIDU_SECRET_KEY={new_secret_key}")
                
                if new_api_key_basic:
                    env_lines.append(f"BAIDU_API_KEY_BASIC={new_api_key_basic}")
                if new_secret_key_basic:
                    env_lines.append(f"BAIDU_SECRET_KEY_BASIC={new_secret_key_basic}")
                
                if new_api_key_general:
                    env_lines.append(f"BAIDU_API_KEY_GENERAL={new_api_key_general}")
                if new_secret_key_general:
                    env_lines.append(f"BAIDU_SECRET_KEY_GENERAL={new_secret_key_general}")
                
                # 写入文件
                with open(env_path, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(env_lines))
                
                settings_window.destroy()
                messagebox.showinfo("成功", 
                    "API密钥已保存！\n\n"
                    "密钥已更新并保存到 .env 文件\n"
                    "立即生效，无需重启程序")
            
            except Exception as e:
                messagebox.showerror("错误", f"保存失败：{str(e)}")
        
        # 按钮
        btn_frame = tk.Frame(settings_window)
        btn_frame.pack(pady=15)
        
        tk.Button(btn_frame, text="保存", command=save_api_keys,
                 bg="#4CAF50", fg="white", padx=30, pady=8, font=("Arial", 11)).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="取消", command=settings_window.destroy,
                 bg="#757575", fg="white", padx=30, pady=8, font=("Arial", 11)).pack(side=tk.LEFT, padx=5)
    
    def show_history_detail(self, history_item):
        """显示历史记录详情"""
        detail_window = self.create_popup_window(self.root, "历史记录详情", "history_detail", 900, 700)
        
        # 标题
        title_text = f"📄 {history_item['type']} - {history_item['timestamp']}"
        tk.Label(detail_window, text=title_text, 
                font=("Arial", 14, "bold")).pack(pady=15)
        
        # 信息
        info_text = f"文件数: {history_item['file_count']} | 总行数: {history_item['total_lines']}"
        tk.Label(detail_window, text=info_text, fg="gray").pack(pady=5)
        
        # 创建文本框显示内容（ScrolledText自带滚动条）
        text_widget = scrolledtext.ScrolledText(detail_window, width=100, height=30,
                                                font=("Microsoft YaHei", 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 显示内容
        for file_info in history_item['files']:
            text_widget.insert(tk.END, f"\n{'='*80}\n")
            text_widget.insert(tk.END, f"文件: {file_info['name']}\n")
            text_widget.insert(tk.END, f"行数: {file_info['lines']}\n")
            text_widget.insert(tk.END, f"{'='*80}\n\n")
            
            for line in file_info['content']:
                text_widget.insert(tk.END, line + "\n")
            
            if file_info['lines'] > len(file_info['content']):
                text_widget.insert(tk.END, f"\n... (还有 {file_info['lines'] - len(file_info['content'])} 行未显示)\n")
        
        text_widget.config(state=tk.DISABLED)
        
        # 按钮
        btn_frame = tk.Frame(detail_window)
        btn_frame.pack(pady=10)
        
        def copy_content():
            all_text = text_widget.get(1.0, tk.END)
            self.root.clipboard_clear()
            self.root.clipboard_append(all_text)
            messagebox.showinfo("成功", "内容已复制到剪贴板")
        
        def export_history_item():
            """导出历史记录到文件"""
            filepath = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
                initialfile=f"历史记录_{history_item['timestamp'].replace(':', '-').replace(' ', '_')}.txt"
            )
            
            if filepath:
                try:
                    with open(filepath, 'w', encoding='utf-8') as f:
                        f.write(f"识别时间: {history_item['timestamp']}\n")
                        f.write(f"识别类型: {history_item['type']}\n")
                        f.write(f"文件数量: {history_item['file_count']}\n")
                        f.write(f"总行数: {history_item['total_lines']}\n")
                        f.write("="*80 + "\n\n")
                        
                        for file_info in history_item['files']:
                            f.write("="*80 + "\n")
                            f.write(f"文件: {file_info['name']}\n")
                            f.write(f"行数: {file_info['lines']}\n")
                            f.write("="*80 + "\n\n")
                            
                            for line in file_info['content']:
                                f.write(line + "\n")
                            f.write("\n")
                    
                    messagebox.showinfo("成功", f"已导出到：{os.path.basename(filepath)}")
                except Exception as e:
                    messagebox.showerror("错误", f"导出失败：{str(e)}")
        
        tk.Button(btn_frame, text="复制内容", command=copy_content,
                 bg="#2196F3", fg="white", padx=20, pady=8).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="导出文件", command=export_history_item,
                 bg="#4CAF50", fg="white", padx=20, pady=8).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="关闭", command=detail_window.destroy,
                 bg="#757575", fg="white", padx=20, pady=8).pack(side=tk.LEFT, padx=5)
    
    def _show_total_stats(self, parent):
        """显示总计统计"""
        # 计算总计
        total_acc_count = 0
        total_acc_success = 0
        total_bas_count = 0
        total_bas_success = 0
        total_gen_count = 0
        total_gen_success = 0
        
        for day_data in self.stats.values():
            if 'accurate' in day_data:
                total_acc_count += day_data['accurate'].get('count', 0)
                total_acc_success += day_data['accurate'].get('success', 0)
                total_bas_count += day_data['basic'].get('count', 0)
                total_bas_success += day_data['basic'].get('success', 0)
                total_gen_count += day_data.get('general', {}).get('count', 0)
                total_gen_success += day_data.get('general', {}).get('success', 0)
        
        total_days = len(self.stats)
        
        info_frame = tk.Frame(parent)
        info_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # 计算日平均
        avg_acc_count = total_acc_count / total_days if total_days > 0 else 0
        avg_acc_success = total_acc_success / total_days if total_days > 0 else 0
        avg_bas_count = total_bas_count / total_days if total_days > 0 else 0
        avg_bas_success = total_bas_success / total_days if total_days > 0 else 0
        avg_gen_count = total_gen_count / total_days if total_days > 0 else 0
        avg_gen_success = total_gen_success / total_days if total_days > 0 else 0
        
        total_all_count = total_acc_count + total_bas_count + total_gen_count
        total_all_success = total_acc_success + total_bas_success + total_gen_success
        
        total_info = f"""
使用天数: {total_days} 天

【高精度识别】
  识别次数: {total_acc_count} 次
  成功图片: {total_acc_success} 张
  日平均次数: {avg_acc_count:.1f} 次/天
  日平均成功: {avg_acc_success:.1f} 张/天

【快速识别】
  识别次数: {total_bas_count} 次
  成功图片: {total_bas_success} 张
  日平均次数: {avg_bas_count:.1f} 次/天
  日平均成功: {avg_bas_success:.1f} 张/天

【通用识别】
  识别次数: {total_gen_count} 次
  成功图片: {total_gen_success} 张
  日平均次数: {avg_gen_count:.1f} 次/天
  日平均成功: {avg_gen_success:.1f} 张/天

【总计】
  总识别次数: {total_all_count} 次
  总成功图片: {total_all_success} 张
  日平均识别: {total_all_count / total_days if total_days > 0 else 0:.1f} 次/天
  日平均成功: {total_all_success / total_days if total_days > 0 else 0:.1f} 张/天
        """
        tk.Label(info_frame, text=total_info, font=("Arial", 11), 
                justify=tk.LEFT, anchor=tk.W).pack(fill=tk.BOTH, expand=True)
    
    def _show_daily_stats(self, parent):
        """显示按日统计（表格形式）"""
        from tkinter import ttk
        
        # 创建表格框架
        table_frame = tk.Frame(parent)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 创建滚动条
        scrollbar = tk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建表格
        columns = ("日期", "类型", "次数", "成功")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", 
                           yscrollcommand=scrollbar.set, height=25)
        
        # 设置列标题
        tree.heading("日期", text="日期")
        tree.heading("类型", text="类型")
        tree.heading("次数", text="次数")
        tree.heading("成功", text="成功")
        
        # 设置列宽度和对齐方式
        tree.column("日期", width=150, anchor=tk.CENTER)
        tree.column("类型", width=120, anchor=tk.CENTER)
        tree.column("次数", width=100, anchor=tk.CENTER)
        tree.column("成功", width=100, anchor=tk.CENTER)
        
        # 配置滚动条
        scrollbar.config(command=tree.yview)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 配置表格样式
        style = ttk.Style()
        style.configure("Treeview", font=("Microsoft YaHei", 10), rowheight=25)
        style.configure("Treeview.Heading", font=("Microsoft YaHei", 11, "bold"))
        
        # 插入数据
        sorted_dates = sorted(self.stats.keys(), reverse=True)
        
        for date in sorted_dates:
            day_data = self.stats[date]
            
            if 'accurate' in day_data:
                acc = day_data['accurate']
                bas = day_data['basic']
                gen = day_data.get('general', {'count': 0, 'success': 0})
                
                # 插入高精度数据
                tree.insert("", tk.END, values=(date, "高精度", 
                                               acc.get('count', 0), 
                                               acc.get('success', 0)),
                           tags=("accurate",))
                
                # 插入快速识别数据
                tree.insert("", tk.END, values=("", "快速", 
                                               bas.get('count', 0), 
                                               bas.get('success', 0)),
                           tags=("basic",))
                
                # 插入通用识别数据
                tree.insert("", tk.END, values=("", "通用", 
                                               gen.get('count', 0), 
                                               gen.get('success', 0)),
                           tags=("general",))
                
                # 插入日合计
                day_total_count = acc.get('count', 0) + bas.get('count', 0) + gen.get('count', 0)
                day_total_success = acc.get('success', 0) + bas.get('success', 0) + gen.get('success', 0)
                tree.insert("", tk.END, values=("", "日合计", 
                                               day_total_count, 
                                               day_total_success),
                           tags=("total",))
        
        # 设置行颜色
        tree.tag_configure("accurate", background="#E3F2FD")
        tree.tag_configure("basic", background="#FFF3E0")
        tree.tag_configure("general", background="#F3E5F5")
        tree.tag_configure("total", background="#E8F5E9", font=("Microsoft YaHei", 10, "bold"))
    
    def _show_monthly_stats(self, parent):
        """显示按月统计"""
        # 按月汇总数据
        monthly_data = {}
        
        for date, day_data in self.stats.items():
            if 'accurate' in day_data:
                month = date[:7]  # YYYY-MM
                
                if month not in monthly_data:
                    monthly_data[month] = {
                        'accurate': {'count': 0, 'success': 0},
                        'basic': {'count': 0, 'success': 0},
                        'general': {'count': 0, 'success': 0},
                        'days': set()
                    }
                
                monthly_data[month]['days'].add(date)
                
                acc = day_data['accurate']
                bas = day_data['basic']
                gen = day_data.get('general', {'count': 0, 'success': 0})
                
                monthly_data[month]['accurate']['count'] += acc.get('count', 0)
                monthly_data[month]['accurate']['success'] += acc.get('success', 0)
                monthly_data[month]['basic']['count'] += bas.get('count', 0)
                monthly_data[month]['basic']['success'] += bas.get('success', 0)
                monthly_data[month]['general']['count'] += gen.get('count', 0)
                monthly_data[month]['general']['success'] += gen.get('success', 0)
        
        # 创建表格框架
        from tkinter import ttk
        
        table_frame = tk.Frame(parent)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 创建滚动条
        scrollbar = tk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 创建表格
        columns = ("月份", "天数", "类型", "次数", "成功", "日均")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", 
                           yscrollcommand=scrollbar.set, height=25)
        
        # 设置列标题
        tree.heading("月份", text="月份")
        tree.heading("天数", text="天数")
        tree.heading("类型", text="类型")
        tree.heading("次数", text="次数")
        tree.heading("成功", text="成功")
        tree.heading("日均", text="日均")
        
        # 设置列宽度和对齐方式
        tree.column("月份", width=120, anchor=tk.CENTER)
        tree.column("天数", width=80, anchor=tk.CENTER)
        tree.column("类型", width=100, anchor=tk.CENTER)
        tree.column("次数", width=80, anchor=tk.CENTER)
        tree.column("成功", width=80, anchor=tk.CENTER)
        tree.column("日均", width=100, anchor=tk.CENTER)
        
        # 配置滚动条
        scrollbar.config(command=tree.yview)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 配置表格样式
        style = ttk.Style()
        style.configure("Treeview", font=("Microsoft YaHei", 10), rowheight=25)
        style.configure("Treeview.Heading", font=("Microsoft YaHei", 11, "bold"))
        
        # 插入数据
        sorted_months = sorted(monthly_data.keys(), reverse=True)
        
        for month in sorted_months:
            data = monthly_data[month]
            acc = data['accurate']
            bas = data['basic']
            days = len(data['days'])
            
            # 计算日平均
            avg_acc = acc['count'] / days if days > 0 else 0
            avg_bas = bas['count'] / days if days > 0 else 0
            
            # 插入高精度数据
            tree.insert("", tk.END, values=(month, days, "高精度", 
                                           acc['count'], acc['success'], 
                                           f"{avg_acc:.1f}"),
                       tags=("accurate",))
            
            # 插入快速识别数据
            tree.insert("", tk.END, values=("", "", "快速", 
                                           bas['count'], bas['success'], 
                                           f"{avg_bas:.1f}"),
                       tags=("basic",))
            
            # 插入月合计
            month_total_count = acc['count'] + bas['count']
            month_total_success = acc['success'] + bas['success']
            avg_total = month_total_count / days if days > 0 else 0
            tree.insert("", tk.END, values=("", "", "月合计", 
                                           month_total_count, month_total_success, 
                                           f"{avg_total:.1f}"),
                       tags=("total",))
        
        # 设置行颜色
        tree.tag_configure("accurate", background="#E3F2FD")
        tree.tag_configure("basic", background="#FFF3E0")
        tree.tag_configure("total", background="#E8F5E9", font=("Microsoft YaHei", 10, "bold"))
    
    def export_results(self):
        """导出识别结果"""
        if not self.all_results:
            messagebox.showwarning("警告", "没有可导出的结果！")
            return
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        
        if not filepath:
            return
        
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                for result in self.all_results:
                    f.write("="*80 + "\n")
                    f.write(f"文件: {result['file']}\n")
                    f.write(f"识别行数: {result['count']}\n")
                    f.write("="*80 + "\n\n")
                    
                    if result['count'] > 0:
                        for line in result['lines']:
                            f.write(line + "\n")
                    else:
                        f.write("识别失败\n")
                    
                    f.write("\n\n")
            
            self.progress_label.config(text=f"✓ 已导出到：{os.path.basename(filepath)}")
        
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")
    
    def merge_images(self):
        """拼接图片功能"""
        file_paths = filedialog.askopenfilenames(
            title="选择要拼接的图片（按住Ctrl多选）",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp"), ("所有文件", "*.*")]
        )
        
        if not file_paths:
            return  # 用户取消选择
        
        if len(file_paths) < 2:
            messagebox.showwarning("警告", "请至少选择2张图片！\n\n提示：按住Ctrl键可以多选图片")
            return
        
        try:
            # 加载所有图片
            images = []
            for path in file_paths:
                img = Image.open(path)
                images.append(img)
            
            # 计算拼接后的尺寸
            total_width = sum(img.width for img in images)
            max_height = max(img.height for img in images)
            
            # 创建拼接图片
            merged_image = Image.new('RGB', (total_width, max_height), 'white')
            
            # 从右到左拼接（默认）
            x_offset = 0
            for img in reversed(images):
                y_offset = (max_height - img.height) // 2
                merged_image.paste(img, (x_offset, y_offset))
                x_offset += img.width
            
            # 询问是否保存
            save_choice = messagebox.askyesnocancel(
                "拼接完成",
                f"拼接完成！\n\n"
                f"图片数量: {len(images)}\n"
                f"拼接尺寸: {total_width}x{max_height}\n\n"
                f"是否保存拼接后的图片？\n\n"
                f"「是」= 保存图片并识别\n"
                f"「否」= 只识别不保存\n"
                f"「取消」= 取消操作"
            )
            
            if save_choice is None:  # 取消
                return
            
            # 保存到临时文件（用于识别）
            import tempfile
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, "merged_temp.jpg")
            merged_image.save(temp_path, format='JPEG', quality=90)
            
            # 如果选择保存
            if save_choice:
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".jpg",
                    filetypes=[("JPEG图片", "*.jpg"), ("PNG图片", "*.png"), ("所有文件", "*.*")],
                    initialfile=f"merged_{len(images)}images_{total_width}x{max_height}.jpg"
                )
                
                if save_path:
                    # 保存到用户指定位置
                    if save_path.lower().endswith('.png'):
                        merged_image.save(save_path, format='PNG')
                    else:
                        merged_image.save(save_path, format='JPEG', quality=95)
                    
                    self.progress_label.config(
                        text=f"✓ 拼接图片已保存到：{os.path.basename(save_path)}")
                    
                    # 使用保存的文件进行识别
                    temp_path = save_path
            
            # 继续识别流程
            result = messagebox.askyesno("开始识别", 
                f"是否立即识别拼接后的图片？\n\n"
                f"拼接尺寸: {total_width}x{max_height}")
            
            if result:
                self.image_paths = [temp_path]
                self.file_label.config(
                    text=f"已选择: 拼接图片 ({len(images)}张) - {total_width}x{max_height}", 
                    fg="blue")
                
                # 检查尺寸并启用相应按钮（宽度和高度都在范围内）
                width_in_accurate = self.size_limits["accurate_min_width"] <= total_width <= self.size_limits["accurate_max_width"]
                height_in_accurate = self.size_limits["accurate_min_height"] <= max_height <= self.size_limits["accurate_max_height"]
                meets_accurate = width_in_accurate and height_in_accurate
                
                width_in_basic = self.size_limits["basic_min_width"] <= total_width <= self.size_limits["basic_max_width"]
                height_in_basic = self.size_limits["basic_min_height"] <= max_height <= self.size_limits["basic_max_height"]
                meets_basic = width_in_basic and height_in_basic
                
                if meets_accurate:
                    self.ocr_btn.config(state=tk.NORMAL)
                else:
                    self.ocr_btn.config(state=tk.DISABLED)
                
                if meets_basic:
                    self.quick_ocr_btn.config(state=tk.NORMAL)
                else:
                    self.quick_ocr_btn.config(state=tk.DISABLED)
                
                self.progress_label.config(text="")
                
                # 选择识别方式
                if meets_accurate and meets_basic:
                    ocr_choice = messagebox.askyesno("选择识别方式",
                        f"是否使用高精度识别？\n\n"
                        f"「是」= 高精度识别\n"
                        f"「否」= 快速识别")
                    if ocr_choice:
                        self.root.after(500, self.perform_ocr)
                    else:
                        self.root.after(500, self.perform_quick_ocr)
                elif meets_accurate:
                    self.root.after(500, self.perform_ocr)
                elif meets_basic:
                    self.root.after(500, self.perform_quick_ocr)
                else:
                    messagebox.showwarning("警告", 
                        f"拼接后的图片尺寸不符合任何识别要求\n\n"
                        f"当前尺寸: {total_width}x{max_height}\n"
                        f"高精度要求: 宽≥{self.size_limits['accurate_min_width']} 且 高≥{self.size_limits['accurate_min_height']}\n"
                        f"快速识别要求: 宽<{self.size_limits['basic_max_width']} 且 高<{self.size_limits['basic_max_height']}")
        
        except Exception as e:
            messagebox.showerror("错误", f"拼接失败：{str(e)}")
    
    def crop_and_merge_direct(self):
        """直接从主界面调用裁剪并拼接功能"""
        file_paths = filedialog.askopenfilenames(
            title="选择要裁剪的图片（可多选）",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp"), ("所有文件", "*.*")]
        )
        
        if not file_paths:
            return
        
        self._open_crop_window(file_paths)
    
    def _open_crop_window(self, file_paths):
        """打开裁剪窗口"""
        crop_window = tk.Toplevel(self.root)
        crop_window.title("裁剪并拼接 - 框选区域")
        
        screen_width = crop_window.winfo_screenwidth()
        screen_height = crop_window.winfo_screenheight()
        
        window_width = int(screen_width * 0.9)
        window_height = int(screen_height * 0.9)
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        crop_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        crop_window.state('zoomed')
        
        try:
            images_data = []
            for path in file_paths:
                img = Image.open(path)
                images_data.append({
                    'path': path,
                    'name': os.path.basename(path),
                    'original': img,
                    'crop_areas': []
                })
            
            display_mode = ['dual' if len(images_data) >= 2 else 'single']
            current_image_index = [0]
            
            max_display_size = min(window_width - 100, window_height - 300)
            
            def get_display_image(img, is_dual_mode=False):
                max_width = (max_display_size // 2 - 20) if is_dual_mode else max_display_size
                max_height = max_display_size
                
                if img.width > max_width or img.height > max_height:
                    scale = min(max_width / img.width, max_height / img.height)
                    new_width = int(img.width * scale)
                    new_height = int(img.height * scale)
                    display_img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                    return display_img, scale
                return img.copy(), 1.0
            
            current_rect = None
            start_x = start_y = 0
            zoom_level = [1.0]
            is_panning = [False]
            
            title_frame = tk.Frame(crop_window, bg="#FF9800")
            title_frame.pack(fill=tk.X)
            
            tk.Label(title_frame, text="✂️ 裁剪并拼接", font=("Arial", 14, "bold"),
                    bg="#FF9800", fg="white", pady=8).pack(side=tk.LEFT, padx=20)
            
            tk.Label(title_frame, text="💡 左键框选 | 右键删除 | 滚轮缩放 | 中键拖动 | Ctrl+0适合屏幕", 
                    font=("Arial", 10), bg="#FF9800", fg="white", pady=8).pack(side=tk.RIGHT, padx=20)
            
            nav_frame = tk.Frame(crop_window)
            nav_frame.pack(fill=tk.X, padx=20, pady=8)
            
            image_label = tk.Label(nav_frame, text="", font=("Arial", 11, "bold"), fg="blue")
            image_label.pack(side=tk.LEFT)
            
            canvas_frame = tk.Frame(crop_window)
            canvas_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
            
            h_scrollbar = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL)
            h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            
            v_scrollbar = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL)
            v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            canvas = tk.Canvas(canvas_frame, bg="gray", cursor="cross",
                             xscrollcommand=h_scrollbar.set,
                             yscrollcommand=v_scrollbar.set)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            h_scrollbar.config(command=canvas.xview)
            v_scrollbar.config(command=canvas.yview)
            
            status_label = tk.Label(crop_window, text="", fg="blue", font=("Arial", 10))
            status_label.pack(pady=5)
            
            merge_info_frame = tk.Frame(crop_window, bg="#f0f0f0", relief=tk.RIDGE, bd=2)
            merge_info_frame.pack(fill=tk.X, padx=20, pady=5)
            
            merge_info_label = tk.Label(merge_info_frame, text="", bg="#f0f0f0", 
                                       font=("Arial", 10, "bold"), fg="#333")
            merge_info_label.pack(pady=8)
            
            def update_status():
                current_img = images_data[current_image_index[0]]
                total_areas = sum(len(img['crop_areas']) for img in images_data)
                
                total_width = 0
                max_height = 0
                
                for img_data in images_data:
                    for area in img_data['crop_areas']:
                        x1, y1, x2, y2 = area['coords']
                        width = x2 - x1
                        height = y2 - y1
                        total_width += width
                        max_height = max(max_height, height)
                
                status_text = f"当前图片已框选 {len(current_img['crop_areas'])} 个区域 | 总共 {total_areas} 个区域"
                status_label.config(text=status_text, fg="blue")
                
                if total_areas > 0:
                    remaining_width = 8100 - total_width
                    usage_percent = (total_width / 8100) * 100
                    
                    merge_text = f"📏 拼接尺寸: 宽 {total_width}px × 高 {max_height}px"
                    merge_text += f"  |  已用: {usage_percent:.1f}%"
                    
                    if total_width > self.size_limits["basic_max_width"]:
                        merge_text += f"  |  ❌ 超限 {total_width - 8100}px"
                        merge_info_label.config(text=merge_text, fg="red")
                        merge_info_frame.config(bg="#ffe0e0")
                        merge_info_label.config(bg="#ffe0e0")
                    elif total_width > 7000:
                        merge_text += f"  |  ⚠️ 剩余 {remaining_width}px"
                        merge_info_label.config(text=merge_text, fg="#ff6600")
                        merge_info_frame.config(bg="#fff3e0")
                        merge_info_label.config(bg="#fff3e0")
                    else:
                        merge_text += f"  |  ✓ 剩余 {remaining_width}px"
                        merge_info_label.config(text=merge_text, fg="green")
                        merge_info_frame.config(bg="#e8f5e9")
                        merge_info_label.config(bg="#e8f5e9")
                else:
                    merge_info_label.config(text="💡 请框选要拼接的区域（左键拖动框选，右键删除）", 
                                          fg="#666")
                    merge_info_frame.config(bg="#f0f0f0")
                    merge_info_label.config(bg="#f0f0f0")
            
            def display_current_image():
                """显示当前图片"""
                canvas.delete("all")
                from PIL import ImageTk
                
                if display_mode[0] == 'dual' and len(images_data) >= 2:
                    img1_data = images_data[0]
                    img2_data = images_data[1]
                    
                    base_img1, base_scale1 = get_display_image(img1_data['original'], is_dual_mode=True)
                    base_img2, base_scale2 = get_display_image(img2_data['original'], is_dual_mode=True)
                    
                    final_scale1 = base_scale1 * zoom_level[0]
                    final_scale2 = base_scale2 * zoom_level[0]
                    
                    final_width1 = int(img1_data['original'].width * final_scale1)
                    final_height1 = int(img1_data['original'].height * final_scale1)
                    final_width2 = int(img2_data['original'].width * final_scale2)
                    final_height2 = int(img2_data['original'].height * final_scale2)
                    
                    display_img1 = img1_data['original'].resize((final_width1, final_height1), Image.Resampling.LANCZOS)
                    display_img2 = img2_data['original'].resize((final_width2, final_height2), Image.Resampling.LANCZOS)
                    
                    gap = 20
                    total_width = final_width1 + gap + final_width2
                    total_height = max(final_height1, final_height2)
                    
                    canvas.config(scrollregion=(0, 0, total_width, total_height))
                    
                    photo1 = ImageTk.PhotoImage(display_img1)
                    canvas.photo1 = photo1
                    canvas.create_image(0, 0, anchor=tk.NW, image=photo1, tags="image1")
                    canvas.create_text(final_width1 // 2, 20, text=f"图1: {img1_data['name']}", 
                                     font=("Arial", 12, "bold"), fill="yellow", tags="label1")
                    
                    x_offset = final_width1 + gap
                    photo2 = ImageTk.PhotoImage(display_img2)
                    canvas.photo2 = photo2
                    canvas.create_image(x_offset, 0, anchor=tk.NW, image=photo2, tags="image2")
                    canvas.create_text(x_offset + final_width2 // 2, 20, text=f"图2: {img2_data['name']}", 
                                     font=("Arial", 12, "bold"), fill="yellow", tags="label2")
                    
                    canvas.image_info = [
                        {'x_offset': 0, 'scale': final_scale1, 'data': img1_data},
                        {'x_offset': x_offset, 'scale': final_scale2, 'data': img2_data}
                    ]
                    
                    area_counter = 1
                    for img_idx, img_info in enumerate(canvas.image_info):
                        img_data = img_info['data']
                        scale = img_info['scale']
                        x_off = img_info['x_offset']
                        
                        for area in img_data['crop_areas']:
                            orig_x1, orig_y1, orig_x2, orig_y2 = area['coords']
                            x1 = x_off + orig_x1 * scale
                            y1 = orig_y1 * scale
                            x2 = x_off + orig_x2 * scale
                            y2 = orig_y2 * scale
                            
                            rect_id = canvas.create_rectangle(x1, y1, x2, y2, outline="red", width=2, tags="rect")
                            text_id = canvas.create_text((x1+x2)/2, (y1+y2)/2, text=str(area_counter),
                                                        font=("Arial", 20, "bold"), fill="red", tags="text")
                            area['rect_id'] = rect_id
                            area['text_id'] = text_id
                            area['display_coords'] = (x1, y1, x2, y2)
                            area['image_index'] = img_idx
                            area_counter += 1
                    
                    zoom_percent = int(zoom_level[0] * 100)
                    image_label.config(text=f"双图模式: {img1_data['name']} + {img2_data['name']} | 缩放: {zoom_percent}%")
                
                else:
                    current_img = images_data[current_image_index[0]]
                    base_display_img, base_scale = get_display_image(current_img['original'], is_dual_mode=False)
                    
                    final_scale = base_scale * zoom_level[0]
                    final_width = int(current_img['original'].width * final_scale)
                    final_height = int(current_img['original'].height * final_scale)
                    
                    display_img = current_img['original'].resize((final_width, final_height), Image.Resampling.LANCZOS)
                    
                    photo = ImageTk.PhotoImage(display_img)
                    canvas.photo = photo
                    canvas.image_info = [{'x_offset': 0, 'scale': final_scale, 'data': current_img}]
                    
                    canvas.config(scrollregion=(0, 0, final_width, final_height))
                    canvas.create_image(0, 0, anchor=tk.NW, image=photo, tags="image")
                    
                    for i, area in enumerate(current_img['crop_areas']):
                        orig_x1, orig_y1, orig_x2, orig_y2 = area['coords']
                        x1 = orig_x1 * final_scale
                        y1 = orig_y1 * final_scale
                        x2 = orig_x2 * final_scale
                        y2 = orig_y2 * final_scale
                        
                        rect_id = canvas.create_rectangle(x1, y1, x2, y2, outline="red", width=2, tags="rect")
                        text_id = canvas.create_text((x1+x2)/2, (y1+y2)/2, text=str(i+1),
                                                    font=("Arial", 20, "bold"), fill="red", tags="text")
                        area['rect_id'] = rect_id
                        area['text_id'] = text_id
                        area['display_coords'] = (x1, y1, x2, y2)
                        area['image_index'] = 0
                    
                    zoom_percent = int(zoom_level[0] * 100)
                    image_label.config(text=f"图片 {current_image_index[0]+1}/{len(images_data)}: {current_img['name']} | 缩放: {zoom_percent}%")
                
                update_status()


            
            def on_mouse_down(event):
                nonlocal start_x, start_y, current_rect
                start_x = canvas.canvasx(event.x)
                start_y = canvas.canvasy(event.y)
                current_rect = canvas.create_rectangle(start_x, start_y, start_x, start_y,
                                                       outline="red", width=2)
            
            def on_mouse_move(event):
                if current_rect:
                    current_x = canvas.canvasx(event.x)
                    current_y = canvas.canvasy(event.y)
                    canvas.coords(current_rect, start_x, start_y, current_x, current_y)
            
            def on_mouse_up(event):
                nonlocal current_rect
                if current_rect:
                    x1, y1, x2, y2 = canvas.coords(current_rect)
                    
                    if abs(x2 - x1) > 10 and abs(y2 - y1) > 10:
                        center_x = (x1 + x2) / 2
                        target_img = None
                        target_img_info = None
                        
                        for img_info in canvas.image_info:
                            img_data = img_info['data']
                            x_off = img_info['x_offset']
                            scale = img_info['scale']
                            img_width = img_data['original'].width * scale
                            
                            if x_off <= center_x <= x_off + img_width:
                                target_img = img_data
                                target_img_info = img_info
                                break
                        
                        if target_img and target_img_info:
                            scale = target_img_info['scale']
                            x_off = target_img_info['x_offset']
                            
                            orig_x1 = int((min(x1, x2) - x_off) / scale)
                            orig_y1 = int(min(y1, y2) / scale)
                            orig_x2 = int((max(x1, x2) - x_off) / scale)
                            orig_y2 = int(max(y1, y2) / scale)
                            
                            orig_x1 = max(0, min(orig_x1, target_img['original'].width))
                            orig_y1 = max(0, min(orig_y1, target_img['original'].height))
                            orig_x2 = max(0, min(orig_x2, target_img['original'].width))
                            orig_y2 = max(0, min(orig_y2, target_img['original'].height))
                            
                            total_areas = sum(len(img['crop_areas']) for img in images_data)
                            
                            area = {
                                'rect_id': current_rect,
                                'coords': (orig_x1, orig_y1, orig_x2, orig_y2),
                                'display_coords': (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
                            }
                            target_img['crop_areas'].append(area)
                            
                            label_x = (x1 + x2) / 2
                            label_y = (y1 + y2) / 2
                            text_id = canvas.create_text(label_x, label_y, 
                                                         text=str(total_areas + 1),
                                                         font=("Arial", 20, "bold"), fill="red")
                            area['text_id'] = text_id
                            
                            update_status()
                        else:
                            canvas.delete(current_rect)
                    else:
                        canvas.delete(current_rect)
                    
                    current_rect = None
            
            def on_canvas_click(event):
                click_x = canvas.canvasx(event.x)
                click_y = canvas.canvasy(event.y)
                
                deleted = False
                
                if display_mode[0] == 'dual' and len(images_data) >= 2:
                    for img_data in images_data:
                        for i, area in enumerate(img_data['crop_areas']):
                            x1, y1, x2, y2 = area['display_coords']
                            if x1 <= click_x <= x2 and y1 <= click_y <= y2:
                                canvas.delete(area['rect_id'])
                                canvas.delete(area['text_id'])
                                img_data['crop_areas'].pop(i)
                                deleted = True
                                break
                        if deleted:
                            break
                else:
                    current_img = images_data[current_image_index[0]]
                    for i, area in enumerate(current_img['crop_areas']):
                        x1, y1, x2, y2 = area['display_coords']
                        if x1 <= click_x <= x2 and y1 <= click_y <= y2:
                            canvas.delete(area['rect_id'])
                            canvas.delete(area['text_id'])
                            current_img['crop_areas'].pop(i)
                            deleted = True
                            break
                
                if deleted:
                    display_current_image()
            
            def on_mouse_wheel(event):
                """鼠标滚轮缩放"""
                old_zoom = zoom_level[0]
                
                if event.delta > 0:
                    zoom_level[0] *= 1.15
                else:
                    zoom_level[0] /= 1.15
                
                zoom_level[0] = max(0.1, min(zoom_level[0], 10.0))
                
                display_current_image()
            
            def on_pan_start(event):
                """开始平移（中键拖动）"""
                canvas.config(cursor="fleur")
                canvas.scan_mark(event.x, event.y)
                is_panning[0] = True
            
            def on_pan_move(event):
                """平移中（中键拖动）"""
                if is_panning[0]:
                    canvas.scan_dragto(event.x, event.y, gain=1)
            
            def on_pan_end(event):
                """结束平移"""
                canvas.config(cursor="cross")
                is_panning[0] = False
            
            def prev_image():
                """上一张图片"""
                if current_image_index[0] > 0:
                    current_image_index[0] -= 1
                    zoom_level[0] = 1.0
                    display_current_image()
            
            def next_image():
                """下一张图片"""
                if current_image_index[0] < len(images_data) - 1:
                    current_image_index[0] += 1
                    zoom_level[0] = 1.0
                    display_current_image()
            
            def on_key_press(event):
                """键盘快捷键处理"""
                if event.keysym == 'r' or event.keysym == 'R':
                    zoom_level[0] = 1.0
                    display_current_image()
                elif event.keysym == 'Left':
                    prev_image()
                elif event.keysym == 'Right':
                    next_image()
                elif event.keysym == 'plus' or event.keysym == 'equal':
                    zoom_level[0] *= 1.2
                    zoom_level[0] = min(zoom_level[0], 10.0)
                    display_current_image()
                elif event.keysym == 'minus':
                    zoom_level[0] /= 1.2
                    zoom_level[0] = max(zoom_level[0], 0.1)
                    display_current_image()
                elif event.keysym == '0' and (event.state & 0x4):  # Ctrl+0
                    fit_screen()
            
            crop_window.bind("<Key>", on_key_press)
            canvas.focus_set()
            
            canvas.bind("<ButtonPress-1>", on_mouse_down)
            canvas.bind("<B1-Motion>", on_mouse_move)
            canvas.bind("<ButtonRelease-1>", on_mouse_up)
            canvas.bind("<Button-3>", on_canvas_click)
            canvas.bind("<MouseWheel>", on_mouse_wheel)
            canvas.bind("<ButtonPress-2>", on_pan_start)
            canvas.bind("<B2-Motion>", on_pan_move)
            canvas.bind("<ButtonRelease-2>", on_pan_end)
            
            display_current_image()
            
            def do_crop_and_merge():
                all_crop_areas = []
                for img_data in images_data:
                    if img_data['crop_areas']:
                        all_crop_areas.extend([
                            (img_data['original'], area['coords'], img_data['name']) 
                            for area in img_data['crop_areas']
                        ])
                
                if not all_crop_areas:
                    messagebox.showwarning("警告", "请至少框选一个区域！")
                    return
                
                try:
                    cropped_images = []
                    for i, (original_img, coords, img_name) in enumerate(all_crop_areas):
                        x1, y1, x2, y2 = coords
                        cropped = original_img.crop((x1, y1, x2, y2))
                        cropped_images.append(cropped)
                    
                    total_width = sum(img.width for img in cropped_images)
                    max_height = max(img.height for img in cropped_images)
                    
                    if total_width > self.size_limits["basic_max_width"]:
                        messagebox.showerror("图片尺寸超限",
                            f"拼接后的图片宽度超过限制！\n\n"
                            f"当前宽度: {total_width}px\n"
                            f"最大宽度: 8100px\n"
                            f"超出: {total_width - 8100}px")
                        return
                    
                    # 根据默认方向拼接图片（从右到左）
                    merged = Image.new('RGB', (total_width, max_height), 'white')
                    
                    # 从右到左拼接（默认）
                    x_offset = 0
                    for i, img in enumerate(reversed(cropped_images)):
                        y_offset = (max_height - img.height) // 2
                        merged.paste(img, (x_offset, y_offset))
                        x_offset += img.width
                    
                    import tempfile
                    temp_dir = tempfile.gettempdir()
                    temp_path = os.path.join(temp_dir, "cropped_merged_ocr.jpg")
                    merged.save(temp_path, format='JPEG', quality=90)
                    
                    crop_window.destroy()
                    
                    # 询问是否保存拼接图片
                    from tkinter import simpledialog
                    
                    save_dialog = tk.Toplevel(self.root)
                    save_dialog.title("保存选项")
                    save_dialog.geometry("500x300")
                    save_dialog.minsize(500, 300)  # 设置最小尺寸
                    save_dialog.transient(self.root)
                    save_dialog.grab_set()
                    
                    # 居中显示
                    save_dialog.update_idletasks()
                    x = (save_dialog.winfo_screenwidth() // 2) - (500 // 2)
                    y = (save_dialog.winfo_screenheight() // 2) - (300 // 2)
                    save_dialog.geometry(f"500x300+{x}+{y}")
                    
                    user_choice = [None]  # 用列表存储选择结果
                    
                    tk.Label(save_dialog, text="拼接完成！", 
                            font=("Arial", 14, "bold")).pack(pady=15)
                    
                    info_text = f"区域数量: {len(cropped_images)}\n"
                    info_text += f"拼接尺寸: 宽{total_width} x 高{max_height}"
                    tk.Label(save_dialog, text=info_text, 
                            font=("Arial", 11)).pack(pady=10)
                    
                    tk.Label(save_dialog, text="是否保存拼接后的图片？", 
                            font=("Arial", 11, "bold")).pack(pady=10)
                    
                    def on_yes():
                        user_choice[0] = 'save'
                        save_dialog.destroy()
                    
                    def on_no():
                        user_choice[0] = 'no_save'
                        save_dialog.destroy()
                    
                    def on_cancel():
                        user_choice[0] = 'cancel'
                        save_dialog.destroy()
                    
                    btn_frame = tk.Frame(save_dialog)
                    btn_frame.pack(pady=20)
                    
                    tk.Button(btn_frame, text="是 - 保存并识别", command=on_yes,
                             bg="#4CAF50", fg="white", font=("Arial", 11),
                             padx=20, pady=10).pack(side=tk.LEFT, padx=5)
                    
                    tk.Button(btn_frame, text="否 - 只识别不保存", command=on_no,
                             bg="#2196F3", fg="white", font=("Arial", 11),
                             padx=20, pady=10).pack(side=tk.LEFT, padx=5)
                    
                    tk.Button(btn_frame, text="取消", command=on_cancel,
                             bg="#757575", fg="white", font=("Arial", 11),
                             padx=20, pady=10).pack(side=tk.LEFT, padx=5)
                    
                    # 等待用户选择
                    self.root.wait_window(save_dialog)
                    
                    if user_choice[0] == 'cancel':
                        # 用户取消操作
                        return
                    
                    # 如果选择保存
                    if user_choice[0] == 'save':
                        save_path = filedialog.asksaveasfilename(
                            defaultextension=".jpg",
                            filetypes=[
                                ("JPEG图片", "*.jpg"),
                                ("PNG图片", "*.png"),
                                ("所有文件", "*.*")
                            ],
                            initialfile=f"merged_{len(cropped_images)}regions_w{total_width}xh{max_height}.jpg"
                        )
                        
                        if save_path:
                            # 保存图片
                            if save_path.lower().endswith('.png'):
                                merged.save(save_path, format='PNG')
                            else:
                                merged.save(save_path, format='JPEG', quality=95)
                            
                            self.progress_label.config(
                                text=f"✓ 拼接图片已保存到：{os.path.basename(save_path)}"
                            )
                        else:
                            # 用户取消了保存对话框，但仍然继续识别
                            pass
                    
                    # 继续识别流程
                    self.result_text.delete(1.0, tk.END)
                    self.result_text.insert(tk.END, f"✓ 已裁剪 {len(cropped_images)} 个区域并拼接\n")
                    self.result_text.insert(tk.END, f"✓ 拼接尺寸: 宽{total_width} x 高{max_height}\n")
                    if user_choice[0] == 'save':
                        self.result_text.insert(tk.END, "="*80 + "\n")
                        self.result_text.insert(tk.END, f"✓ 图片已保存\n")
                    self.result_text.insert(tk.END, "正在识别拼接后的图片，请稍候...\n\n")
                    
                    self.image_paths = [temp_path]
                    self.file_label.config(
                        text=f"裁剪拼接图片 ({len(cropped_images)}个区域) - 宽{total_width} x 高{max_height}",
                        fg="blue"
                    )
                    
                    # 检查尺寸（宽度和高度都在范围内）
                    width_in_accurate = self.size_limits["accurate_min_width"] <= total_width <= self.size_limits["accurate_max_width"]
                    height_in_accurate = self.size_limits["accurate_min_height"] <= max_height <= self.size_limits["accurate_max_height"]
                    meets_accurate = width_in_accurate and height_in_accurate
                    
                    width_in_basic = self.size_limits["basic_min_width"] <= total_width <= self.size_limits["basic_max_width"]
                    height_in_basic = self.size_limits["basic_min_height"] <= max_height <= self.size_limits["basic_max_height"]
                    meets_basic = width_in_basic and height_in_basic
                    
                    if meets_accurate:
                        self.ocr_btn.config(state=tk.NORMAL)
                    else:
                        self.ocr_btn.config(state=tk.DISABLED)
                    
                    if meets_basic:
                        self.quick_ocr_btn.config(state=tk.NORMAL)
                    else:
                        self.quick_ocr_btn.config(state=tk.DISABLED)
                    
                    self.progress_label.config(text="")
                    
                    if meets_accurate and meets_basic:
                        result = messagebox.askyesno("选择识别方式",
                            f"是否使用高精度识别？\n\n"
                            f"「是」= 高精度识别\n"
                            f"「否」= 快速识别")
                        if result:
                            self.root.after(500, self.perform_ocr)
                        else:
                            self.root.after(500, self.perform_quick_ocr)
                    elif meets_accurate:
                        self.root.after(500, self.perform_ocr)
                    elif meets_basic:
                        self.root.after(500, self.perform_quick_ocr)
                
                except Exception as e:
                    messagebox.showerror("错误", f"裁剪拼接失败：{str(e)}")
            
            btn_frame = tk.Frame(crop_window)
            btn_frame.pack(pady=15)
            
            def zoom_in():
                zoom_level[0] *= 1.2
                zoom_level[0] = min(zoom_level[0], 10.0)
                display_current_image()
            
            def zoom_out():
                zoom_level[0] /= 1.2
                zoom_level[0] = max(zoom_level[0], 0.1)
                display_current_image()
            
            def zoom_reset():
                zoom_level[0] = 1.0
                display_current_image()
            
            def fit_screen():
                """适合屏幕 - 自动调整缩放以填充可视区域"""
                try:
                    # 获取canvas的可视区域大小
                    canvas_width = canvas.winfo_width()
                    canvas_height = canvas.winfo_height()
                    
                    if canvas_width <= 1 or canvas_height <= 1:
                        # 如果canvas还没有渲染，使用默认值
                        canvas_width = max_display_size
                        canvas_height = max_display_size
                    
                    if display_mode[0] == 'dual' and len(images_data) >= 2:
                        # 双图模式：计算两张图片的总宽度
                        img1 = images_data[0]['original']
                        img2 = images_data[1]['original']
                        
                        # 获取基础缩放
                        _, base_scale1 = get_display_image(img1, is_dual_mode=True)
                        _, base_scale2 = get_display_image(img2, is_dual_mode=True)
                        
                        # 计算总宽度（包括间隔）
                        total_width = img1.width * base_scale1 + 20 + img2.width * base_scale2
                        max_height = max(img1.height * base_scale1, img2.height * base_scale2)
                        
                        # 计算适合屏幕的缩放比例
                        scale_x = canvas_width / total_width
                        scale_y = canvas_height / max_height
                        fit_scale = min(scale_x, scale_y) * 0.95  # 留5%边距
                        
                        zoom_level[0] = fit_scale
                    else:
                        # 单图模式
                        current_img = images_data[current_image_index[0]]['original']
                        _, base_scale = get_display_image(current_img, is_dual_mode=False)
                        
                        # 计算适合屏幕的缩放比例
                        img_width = current_img.width * base_scale
                        img_height = current_img.height * base_scale
                        
                        scale_x = canvas_width / img_width
                        scale_y = canvas_height / img_height
                        fit_scale = min(scale_x, scale_y) * 0.95  # 留5%边距
                        
                        zoom_level[0] = fit_scale
                    
                    # 限制缩放范围
                    zoom_level[0] = max(0.1, min(zoom_level[0], 10.0))
                    
                    display_current_image()
                    
                    # 居中显示
                    canvas.update_idletasks()
                    canvas.xview_moveto(0)
                    canvas.yview_moveto(0)
                
                except Exception as e:
                    print(f"适合屏幕失败: {e}")
                    zoom_level[0] = 1.0
                    display_current_image()
            
            tk.Button(btn_frame, text="🔍+", command=zoom_in,
                     bg="#009688", fg="white", font=("Arial", 11),
                     padx=15, pady=10).pack(side=tk.LEFT, padx=3)
            
            tk.Button(btn_frame, text="🔍-", command=zoom_out,
                     bg="#009688", fg="white", font=("Arial", 11),
                     padx=15, pady=10).pack(side=tk.LEFT, padx=3)
            
            tk.Button(btn_frame, text="重置", command=zoom_reset,
                     bg="#009688", fg="white", font=("Arial", 11),
                     padx=15, pady=10).pack(side=tk.LEFT, padx=3)
            
            tk.Button(btn_frame, text="📐 适合屏幕", command=fit_screen,
                     bg="#009688", fg="white", font=("Arial", 11),
                     padx=15, pady=10).pack(side=tk.LEFT, padx=3)
            
            tk.Frame(btn_frame, width=2, bg="gray").pack(side=tk.LEFT, padx=10, fill=tk.Y)
            
            if len(images_data) > 1:
                tk.Button(btn_frame, text="◀ 上一张", command=prev_image,
                         bg="#2196F3", fg="white", font=("Arial", 11),
                         padx=20, pady=10).pack(side=tk.LEFT, padx=5)
                
                tk.Button(btn_frame, text="下一张 ▶", command=next_image,
                         bg="#2196F3", fg="white", font=("Arial", 11),
                         padx=20, pady=10).pack(side=tk.LEFT, padx=5)
                
                tk.Frame(btn_frame, width=2, bg="gray").pack(side=tk.LEFT, padx=10, fill=tk.Y)
            
            tk.Button(btn_frame, text="✓ 确认拼接", command=do_crop_and_merge,
                     bg="#4CAF50", fg="white", font=("Arial", 12, "bold"),
                     padx=40, pady=12).pack(side=tk.LEFT, padx=10)
            
            tk.Button(btn_frame, text="✗ 取消", command=crop_window.destroy,
                     bg="#757575", fg="white", font=("Arial", 12),
                     padx=40, pady=12).pack(side=tk.LEFT, padx=10)
        
        except Exception as e:
            messagebox.showerror("错误", f"加载图片失败：{str(e)}")


if __name__ == '__main__':
    try:
        # 尝试使用TkinterDnD支持拖放
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        # 如果没有安装tkinterdnd2，使用普通Tk
        print("提示：安装 tkinterdnd2 可以启用拖放功能")
        print("安装命令：pip install tkinterdnd2")
        root = tk.Tk()
    
    app = OCRApp(root)
    root.mainloop()
