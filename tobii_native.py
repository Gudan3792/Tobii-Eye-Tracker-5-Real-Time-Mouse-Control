
import ctypes
import threading
import time
import os
import sys
import collections
import configparser
import math
import tkinter as tk
import subprocess
import winsound
from PIL import Image, ImageTk, ImageGrab

# --- 작업 디렉토리를 스크립트/EXE 위치로 설정 ---
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))
else:
    os.path.chdir(os.path.dirname(os.path.abspath(__file__)))

# --- 상수 및 열거형 ---
TOBII_ERROR_NO_ERROR = 0
TOBII_FIELD_OF_USE_INTERACTIVE = 1
TOBII_VALIDITY_VALID = 1

# --- DLL 로드 ---
def get_dll_path(filename):
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, filename)

DLL_PATH = get_dll_path("tobii_stream_engine.dll")

try:
    tobii = ctypes.CDLL(DLL_PATH)
except FileNotFoundError:
    try:
        tobii = ctypes.CDLL("tobii_stream_engine.dll")
    except Exception:
        print(f"Error: Could not find {DLL_PATH}")
        sys.exit(1)
except OSError as e:
    print(f"Error loading DLL: {e}")
    sys.exit(1)

# --- 구조체 및 타입 ---
c_void_p = ctypes.c_void_p
c_char_p = ctypes.c_char_p
c_int = ctypes.c_int
c_float = ctypes.c_float
c_uint64 = ctypes.c_uint64

class tobii_gaze_point_t(ctypes.Structure):
    _fields_ = [
        ("timestamp_us", ctypes.c_longlong),
        ("validity", c_int),
        ("position_xy", c_float * 2)
    ]

tobii.tobii_api_create.argtypes = [ctypes.POINTER(c_void_p), c_void_p, c_void_p]
tobii.tobii_api_create.restype = c_int
tobii.tobii_api_destroy.argtypes = [c_void_p]
tobii.tobii_api_destroy.restype = c_int
URL_RECEIVER_T = ctypes.CFUNCTYPE(None, c_char_p, c_void_p)
tobii.tobii_enumerate_local_device_urls.argtypes = [c_void_p, URL_RECEIVER_T, c_void_p]
tobii.tobii_enumerate_local_device_urls.restype = c_int
tobii.tobii_device_create.argtypes = [c_void_p, c_char_p, c_int, ctypes.POINTER(c_void_p)]
tobii.tobii_device_create.restype = c_int
tobii.tobii_device_destroy.argtypes = [c_void_p]
tobii.tobii_device_destroy.restype = c_int
GAZE_CALLBACK_T = ctypes.CFUNCTYPE(None, ctypes.POINTER(tobii_gaze_point_t), c_void_p)
tobii.tobii_gaze_point_subscribe.argtypes = [c_void_p, GAZE_CALLBACK_T, c_void_p]
tobii.tobii_gaze_point_subscribe.restype = c_int
tobii.tobii_gaze_point_unsubscribe.argtypes = [c_void_p]
tobii.tobii_gaze_point_unsubscribe.restype = c_int
tobii.tobii_wait_for_callbacks.argtypes = [c_int, ctypes.POINTER(c_void_p)]
tobii.tobii_wait_for_callbacks.restype = c_int
tobii.tobii_device_process_callbacks.argtypes = [c_void_p]
tobii.tobii_device_process_callbacks.restype = c_int

user32 = ctypes.windll.user32
SCREEN_W = user32.GetSystemMetrics(0)
SCREEN_H = user32.GetSystemMetrics(1)

# --- 정밀 줌 창 클래스 ---
# 안구 마우스 사용 시 타겟을 정밀하게 클릭하기 위해, 특정 영역을 캡처한 뒤 확대해서 보여주는 오버레이 창입니다.
class ZoomWindow(tk.Toplevel):
    """
    화면의 특정 부분을 캡처(비트맵)하여 2~3배 확대한 이미지를 화면 중앙에 띄웁니다.
    사용자는 이 창 안에서 눈을 움직여 중앙 가이드(빨간 원)를 타겟에 맞춘 후, 
    눈을 깜빡여 클릭을 수행하게 됩니다. (오조작을 방지하는 돋보기 역할)
    """
    def __init__(self, parent, capture_pos, size, scale):
        super().__init__(parent)
        self.capture_pos = capture_pos
        self.size = size
        self.scale = scale
        self.display_size = int(size * scale)
        
        self.overrideredirect(True)
        self.attributes('-topmost', True)
        
        x = (SCREEN_W - self.display_size) // 2
        y = (SCREEN_H - self.display_size) // 2
        self.geometry(f"{self.display_size}x{self.display_size}+{x}+{y}")
        self.zx, self.zy = x, y
        
        cx, cy = capture_pos
        left = max(0, int(cx - size // 2))
        top = max(0, int(cy - size // 2))
        if left + size > SCREEN_W: left = SCREEN_W - size
        if top + size > SCREEN_H: top = SCREEN_H - size
        
        self.capture_origin = (left, top)
        
        try:
            img = ImageGrab.grab(bbox=(left, top, left + size, top + size))
            img = img.resize((self.display_size, self.display_size), Image.Resampling.LANCZOS)
            self.photo = ImageTk.PhotoImage(img)
            
            self.canvas = tk.Canvas(self, width=self.display_size, height=self.display_size, 
                                     highlightthickness=0)
            self.canvas.create_image(0, 0, anchor='nw', image=self.photo)
            self.canvas.pack()
            
            mid = self.display_size // 2
            self.guide = self.canvas.create_oval(mid-15, mid-15, mid+15, mid+15, outline="red", width=3)
        except Exception as e:
            print(f"Zoom Error: {e}")
            self.destroy()

    def update_guide(self, screen_x, screen_y):
        rx = screen_x - self.zx
        ry = screen_y - self.zy
        if 0 <= rx <= self.display_size and 0 <= ry <= self.display_size:
            self.canvas.coords(self.guide, rx-15, ry-15, rx+15, ry+15)
            self.canvas.itemconfig(self.guide, state='normal')
        else:
            self.canvas.itemconfig(self.guide, state='hidden')

    def get_real_coords(self, screen_x, screen_y):
        rx = screen_x - self.zx
        ry = screen_y - self.zy
        ox = self.capture_origin[0] + (rx / self.scale)
        oy = self.capture_origin[1] + (ry / self.scale)
        return int(ox), int(oy)

# --- 확장 액션 메뉴(우상단 도구 모음) UI 클래스 ---
# 안구 마우스의 한계를 보완하기 위해 더블클릭, 우클릭, 드래그, 창닫기 기능을 보조하는 플라이아웃 형태의 메뉴입니다.
class ActionMenuWindow(tk.Toplevel):
    """
    평소에는 화면 우측 가장자리에 얇고 반투명한 막대로 대기하다가,
    사용자의 시선(또는 마우스 포인터)이 근처(약 80% 거리)에 접근하면 스르륵 패널이 전개됩니다.
    마우스 포인터가 영역 밖으로 벗어나면 다시 얇은 막대 형태로 눈에 띄지 않게 축소됩니다.
    """
    def __init__(self, parent, main_app):
        super().__init__(parent)
        self.main_app = main_app
        self.overrideredirect(True)
        self.attributes('-topmost', True, '-alpha', 0.4) # 최초 막대: 거의 보일듯 말듯 더 투명하게 (60% 투명)
        
        self.expanded_w = 180  
        self.collapsed_w = 15  
        self.h = 320           
        self.y_pos = int(SCREEN_H * 0.15) 
        
        self.geometry(f"{self.collapsed_w}x{self.h}+{SCREEN_W - self.collapsed_w}+{self.y_pos}")
        self.config(bg="#D5D8DC") # 은은하고 차분한 라이트 실버/그레이
        
        self.is_expanded = False
        self.animating = False
        self.enter_timer = None
        
        self.btn_frame = tk.Frame(self, bg="#F8F9F9") # 눈부시지 않은 부드러운 오프화이트
        
        btn_font = ("Arial", 16, "bold")
        fg_col = "#2C3E50" # 눈이 가장 편안한 짙은 슬레이트 폰트 색상
        
        # 눈부심을 없앤 소프트 톤 (탁하지도, 쨍하지도 않은 차분하고 정돈된 색상)
        tk.Button(self.btn_frame, text="더블클릭", command=self.cmd_double, font=btn_font, bg="#FDEBD0", activebackground="#FAD7A1", fg=fg_col).pack(fill=tk.BOTH, expand=True, pady=1)
        tk.Button(self.btn_frame, text="우클릭", command=self.cmd_right, font=btn_font, bg="#FADBD8", activebackground="#F5B7B1", fg=fg_col).pack(fill=tk.BOTH, expand=True, pady=1)
        tk.Button(self.btn_frame, text="드래그", command=self.cmd_drag, font=btn_font, bg="#D6EAF8", activebackground="#AED6F1", fg=fg_col).pack(fill=tk.BOTH, expand=True, pady=1)
        tk.Button(self.btn_frame, text="창 닫기", command=self.cmd_close, font=btn_font, bg="#E6B0AA", activebackground="#D98880", fg=fg_col).pack(fill=tk.BOTH, expand=True, pady=1)
        
        self.leave_timer = None
        self.after(50, self.check_proximity)

    def check_proximity(self):
        """
        0.05초(50ms)마다 현재 OS 마우스(안구 커서)의 절대 좌표를 주기적으로 폴링(Polling)합니다.
        가장자리 끝까지 시선을 정밀하게 옮기기 힘든 안구 마우스 사용자의 특성을 고려하여, 
        물리적인 패널 범위보다 조금 더 넓은 가상 영역(전체 폭의 80% 거리)에 진입하면 선제적으로 UI를 엽니다.
        """
        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
        pt = POINT()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
        cx, cy = pt.x, pt.y
        
        # 튀어나오는 전개 박스 폭의 약 80% 범위까지 접근하면 동작 (화면 끝까지 보지 않아도 됨)
        target_x = SCREEN_W - (self.expanded_w * 0.8)
        
        # 이미 패널이 펼쳐졌다면 시선이 약간 떨어져도 바로 닫히지 않게 조작 여유 범위(40px) 추가
        if self.is_expanded or self.animating:
            target_x = SCREEN_W - (self.expanded_w + 40)
            
        top = self.y_pos - 20
        bottom = self.y_pos + self.h + 20
        
        in_zone = (cx >= target_x) and (top <= cy <= bottom)
        
        if in_zone:
            if getattr(self, 'leave_timer', None) is not None:
                self.after_cancel(self.leave_timer)
                self.leave_timer = None
                
            if not self.is_expanded and not self.animating:
                if getattr(self, 'enter_timer', None) is None:
                    self.enter_timer = self.after(150, lambda: self.slide(True))
        else:
            if getattr(self, 'enter_timer', None) is not None:
                self.after_cancel(self.enter_timer)
                self.enter_timer = None
                
            if self.is_expanded and not self.animating:
                if getattr(self, 'leave_timer', None) is None:
                    self.leave_timer = self.after(300, self._do_collapse)
                    
        self.after(50, self.check_proximity)

    def _do_collapse(self):
        self.leave_timer = None
        if self.is_expanded and not self.animating:
            self.slide(False)

    def slide(self, expand):
        self.animating = True
        target_w = self.expanded_w if expand else self.collapsed_w
        
        if expand and not self.is_expanded:
            self.btn_frame.place(x=self.collapsed_w - self.expanded_w, y=0, width=self.expanded_w, height=self.h)

        steps = 10
        step_time = 15
        current_w = self.winfo_width()
        dw = (target_w - current_w) / steps
        
        def do_step(step):
            if step <= steps:
                new_w = int(current_w + dw * step)
                self.geometry(f"{new_w}x{self.h}+{SCREEN_W - new_w}+{self.y_pos}")
                self.btn_frame.place_configure(x=new_w - self.expanded_w)
                
                # 내부 UI는 불투명(1.0), 대기 막대일 때는 더 투명한(0.4) 상태로 자연스럽게 전환
                new_alpha = 0.4 + (0.6 * (step / steps)) if expand else 1.0 - (0.6 * (step / steps))
                self.attributes('-alpha', new_alpha)
                
                self.after(step_time, do_step, step+1)
            else:
                self.geometry(f"{target_w}x{self.h}+{SCREEN_W - target_w}+{self.y_pos}")
                self.btn_frame.place_configure(x=target_w - self.expanded_w)
                self.is_expanded = expand
                self.animating = False
                self.attributes('-alpha', 1.0 if expand else 0.4)
                if not expand:
                    self.btn_frame.place_forget()
                
        do_step(1)

    def _prepare_action(self):
        winsound.Beep(800, 100)
        self.slide(False)

    def cmd_double(self):
        self.main_app.next_click_action = "double"
        self._prepare_action()

    def cmd_right(self):
        self.main_app.next_click_action = "right"
        self._prepare_action()

    def cmd_drag(self):
        self.main_app.next_click_action = "drag"
        self._prepare_action()

    def cmd_close(self):
        user32 = ctypes.windll.user32
        import os
        pid = os.getpid()
        
        target_hwnd = None
        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        
        def enum_callback(hwnd, lParam):
            nonlocal target_hwnd
            if user32.IsWindowVisible(hwnd):
                window_pid = ctypes.c_ulong()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
                
                # 본인 프로그램(UI)이 아닐 때
                if window_pid.value != pid:
                    # 메인 창 조건 (Owner == 0)
                    owner = user32.GetWindow(hwnd, 4) # GW_OWNER
                    if owner == 0:
                        length = user32.GetWindowTextLengthW(hwnd)
                        if length > 0:
                            title_buf = ctypes.create_unicode_buffer(length + 1)
                            user32.GetWindowTextW(hwnd, title_buf, length + 1)
                            title = title_buf.value
                            if title not in ("Program Manager", "Windows 입력 환경", "설정"):
                                target_hwnd = hwnd
                                return False # 탐색 종료 (가장 위쪽의 유효 앱 창 발견)
            return True # 계속 탐색
            
        user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
            
        if target_hwnd:
            user32.PostMessageW(target_hwnd, 0x0010, 0, 0)
            
        winsound.Beep(800, 200)
        self.slide(False)

# --- 안구 마우스 메인 애플리케이션 클래스 ---
# Tobii 아이트래커와 연동하여 사용자의 시선을 추적하고, 눈 깜빡임(Blink)을 마우스 제어로 변환하는 핵심 비즈니스 로직입니다.
class TobiiMouseApp:
    """
    백그라운드 스레드에서 Tobii C API와 통신하며 시선 좌표를 화면 보정(Calibration) 모델에 맞게 변환합니다.
    사용자가 눈을 감는 체공 시간(Duration)에 따라 [일반 클릭], [줌 모드 진입], [수면 보호막 모드 전환] 등으로 분기 처리하며,
    ctypes의 user32 API를 통해 OS 시스템 레벨에 가상 물리 마우스 입력을 주입합니다.
    """
    def __init__(self):
        self.api = c_void_p(None)
        self.device = c_void_p(None)
        self.device_url = None
        self.running = False
        self.lock = threading.Lock()
        
        self.cur_x, self.cur_y = SCREEN_W / 2, SCREEN_H / 2
        self.offset_x, self.offset_y = 0, 0
        self.cal_grid = []
        self.scale_x, self.scale_y = 1.0, 1.0
        
        self.last_valid_time = time.time()
        self.last_valid_pos = (self.cur_x, self.cur_y)
        self.is_tracking = False
        self.zoom_mode = False
        self.zoom_win = None
        
        self.last_click_time = 0
        self.click_delay = 0.1
        self.stable_click_pos = (self.cur_x, self.cur_y)
        
        self.next_click_action = "left"
        self.is_dragging = False

        
        self.smooth = 0.1
        self.avg_samples = 10
        self.gaze_buffer = collections.deque(maxlen=self.avg_samples)
        self.tracking_enabled = True
        
        # 신규 노이즈 방지 설정값 (기본값)
        self.deep_sleep_threshold = 5.0  # 이 시간 이상 눈을 감으면 '장시간 수면'으로 판정
        self.bottom_block_threshold = 0.96  # 장기 수면 후 복귀 시 클릭을 차단할 하단 임계값 (0.0~1.0)
        
        # 스티키 수면 보호막 관련 변수
        self.sleep_shield_active = False # 수면 보호막 활성화 여부
        self.wake_confirm_duration = 2.0 # 보호막해제를 위해 필요한 연속 시선 유지 시간
        self.last_invalid_time = time.time() # 연속 시선 유지 체크용

        self.url_cb = URL_RECEIVER_T(self.url_receiver_func)
        self.gaze_cb = GAZE_CALLBACK_T(self.gaze_callback_func)
        
        self.init_ui()
        
        # 액션/오버레이 바 생성
        self.action_menu = ActionMenuWindow(self.root, self)
        
        self.load_calibration()
        self.check_settings_loop() # 설정 파일 실시간 감시 시작

    def init_ui(self):
        self.root = tk.Tk()
        self.root.title("TobiiMouse Controller")
        self.root.geometry("300x280")
        self.root.resizable(False, False)
        self.root.configure(bg="#2c3e50")
        
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"300x280+{(sw-300)//2}+{(sh-280)//2}")

        top_frame = tk.Frame(self.root, bg="#2c3e50")
        top_frame.pack(pady=20)

        self.toggle_btn = tk.Button(top_frame, text="ON", bg="#2ecc71", fg="white", 
                                    command=self.toggle_tracking, font=("Arial", 16, "bold"), 
                                    height=3, width=8, bd=4, relief="raised")
        self.toggle_btn.pack(side=tk.LEFT, padx=10)

        self.cal_btn = tk.Button(top_frame, text="CALI", bg="#3498db", fg="white", 
                                 command=self.run_calibration, font=("Arial", 16, "bold"), 
                                 height=3, width=8, bd=4, relief="raised")
        self.cal_btn.pack(side=tk.LEFT, padx=10)

        self.exit_btn = tk.Button(self.root, text="EXIT", bg="#e74c3c", fg="white", 
                                  command=self.on_exit, font=("Arial", 16, "bold"), 
                                  height=2, width=18, bd=4, relief="raised")
        self.exit_btn.pack(pady=10)

        self.status_label = tk.Label(self.root, text="", bg="#2c3e50", fg="#bdc3c7", font=("Arial", 12, "bold"))
        self.status_label.pack(pady=5)
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)

    def load_calibration(self):
        ini_path = "eye_setting.ini"
        if not os.path.exists(ini_path): return
        try:
            mtime = os.path.getmtime(ini_path)
            if hasattr(self, 'cal_mtime') and mtime <= self.cal_mtime: return
            
            config = configparser.ConfigParser()
            try:
                with open(ini_path, 'r', encoding='utf-8-sig') as f: config.read_file(f)
            except:
                with open(ini_path, 'r', encoding='cp949') as f: config.read_file(f)
                
            if 'Calibration' in config:
                self.smooth = float(config['Calibration'].get('smooth', '0.1'))
                new_avg = int(config['Calibration'].get('avg_samples', '10'))
                self.blink_min = float(config['Calibration'].get('blink_min', '0.3'))
                self.blink_max = float(config['Calibration'].get('blink_max', '1.0'))
                self.zoom_min = float(config['Calibration'].get('zoom_min', '1.0'))
                self.zoom_max = float(config['Calibration'].get('zoom_max', '3.0'))
                self.zoom_size = int(config['Calibration'].get('zoom_size', '300'))
                self.zoom_scale = float(config['Calibration'].get('zoom_scale', '3.0'))
                self.click_delay = float(config['Calibration'].get('click_delay', '0.1'))
                self.deep_sleep_threshold = float(config['Calibration'].get('deep_sleep_threshold', '5.0'))
                self.bottom_block_threshold = float(config['Calibration'].get('bottom_block_threshold', '0.96'))
                
                if new_avg != self.avg_samples:
                    self.avg_samples = max(1, min(100, new_avg))
                    with self.lock: self.gaze_buffer = collections.deque(maxlen=self.avg_samples)
            
            self.cal_grid = []
            if 'CalibrationGrid' in config:
                count_str = config['CalibrationGrid'].get('point_count', '0')
                try:
                    count = int(count_str)
                    for i in range(count):
                        self.cal_grid.append({
                            'raw_x': float(config['CalibrationGrid'][f'raw_x_{i}']),
                            'raw_y': float(config['CalibrationGrid'][f'raw_y_{i}']),
                            'target_x': float(config['CalibrationGrid'][f'target_x_{i}']),
                            'target_y': float(config['CalibrationGrid'][f'target_y_{i}'])
                        })
                except: pass
            
            self.root.title(f"TobiiMouse (S:{self.smooth:.2f} | A:{self.avg_samples})")
            self.status_label.config(text=f"Smooth: {self.smooth:.2f} | Avg: {self.avg_samples}")
            self.cal_mtime = mtime
        except Exception as e: print(f"Load Error: {e}")

    def check_settings_loop(self):
        """1초마다 설정 파일의 변경 사항을 체크하여 자동 로드"""
        if self.running:
            self.load_calibration()
        self.root.after(1000, self.check_settings_loop)

    def gaze_callback_func(self, gaze_point_ptr, user_data):
        """
        Tobii 센서로부터 시선 데이터가 발생할 때마다 매우 빠른 주기(약 30~90Hz)로 호출되는 콜백 함수입니다.
        데이터가 유효함(Validity==1)일 경우 버퍼에 좌표를 쌓아서 마우스 커서를 부드럽게 이동시킵니다.
        반대로 시선 데이터가 유실(눈을 감음)되었다가 다시 유효해지는 순간, 
        눈을 감고 있었던 '유실 시간(Duration)'을 측정하여 클릭 또는 줌 액션을 실행합니다.
        """
        data = gaze_point_ptr.contents
        if data.validity == TOBII_VALIDITY_VALID:
            mx, my = self.map_gaze(data.position_xy[0], data.position_xy[1])
            tx, ty = int(mx * SCREEN_W), int(my * SCREEN_H)
            
            with self.lock:
                now = time.time()
                
                # 2. 트래킹 유실 후 복구 시점 판정 (눈을 떴을 때)
                if self.tracking_enabled and not self.is_tracking:
                    duration = now - self.last_valid_time
                    
                    # 장시간 눈 감음 여부 판정 -> 보호막 활성화
                    if duration >= self.deep_sleep_threshold:
                        self.sleep_shield_active = True
                        print(f"[SHIELD] Sleep Shield ACTIVATED (Duration: {duration:.2f}s)")

                    # (1) 일반/확정 클릭 (0.3 ~ 1.0)
                    if self.blink_max < 999 and self.blink_min <= duration < self.blink_max:
                        # 보호막이 켜져 있는 동안 하단 노이즈는 무조건 차단
                        if self.sleep_shield_active and data.position_xy[1] >= self.bottom_block_threshold:
                            print(f"[BLOCK] Shielded noise (CLICK) at bottom: Y={data.position_xy[1]:.4f}")
                            self.last_valid_time, self.last_valid_pos = now, (tx, ty)
                            self.last_invalid_time = now # 보호막 해제 타이머 리셋
                            self.is_tracking = True
                            return

                        sx, sy = self.stable_click_pos
                        if self.zoom_mode and self.zoom_win:
                            rx, ry = self.zoom_win.get_real_coords(sx, sy)
                            self.cur_x, self.cur_y = rx, ry
                            self.last_click_time = now
                            self.exit_zoom_mode()
                            self.execute_click(rx, ry)
                        else:
                            self.cur_x, self.cur_y = sx, sy
                            self.last_click_time = now
                            self.execute_click(int(sx), int(sy))
                        
                        self.gaze_buffer.clear()
                        self.last_valid_time, self.last_valid_pos = now, (tx, ty)
                        self.last_invalid_time = now # 보호막 해제 타이머 리셋
                        self.is_tracking = True
                        return
                        
                    # (2) 줌 모드 진입 (1.0 ~ 3.0)
                    elif self.zoom_max < 999 and self.zoom_min <= duration < self.zoom_max:
                        if not self.zoom_mode:
                            # 보호막이 켜져 있는 동안 하단 노이즈는 무조건 차단
                            if self.sleep_shield_active and data.position_xy[1] >= self.bottom_block_threshold:
                                print(f"[BLOCK] Shielded noise (ZOOM) at bottom: Y={data.position_xy[1]:.4f}")
                                self.last_valid_time, self.last_valid_pos = now, (tx, ty)
                                self.last_invalid_time = now # 보호막 해제 타이머 리셋
                                self.is_tracking = True
                                return

                            sx, sy = self.stable_click_pos
                            self.enter_zoom_mode(int(sx), int(sy))
                            self.cur_x, self.cur_y = sx, sy
                            self.gaze_buffer.clear()
                            self.last_valid_time, self.last_valid_pos = now, (tx, ty)
                            self.last_invalid_time = now # 보호막 해제 타이머 리셋
                            self.is_tracking = True
                            return

                    # (3) 취소 혹은 종료 (> 줌 최대시간)
                    elif duration >= self.zoom_max and self.zoom_mode:
                        self.exit_zoom_mode()
                
                # 3. 보호막 해제 조건 체크 (연속 시선 유지)
                if self.sleep_shield_active:
                    # 하단 영역(4%)을 보고 있는 동안에는 '깨어남'으로 인정하지 않음 (실눈 노이즈 방지)
                    if data.position_xy[1] >= self.bottom_block_threshold:
                        self.last_invalid_time = now # 타이머 계속 초기화

                    if now - self.last_invalid_time >= self.wake_confirm_duration:
                        self.sleep_shield_active = False
                        print("[SHIELD] Sleep Shield DEACTIVATED - User is awake.")

                # 4. 버퍼 업데이트
                self.gaze_buffer.append((tx, ty))
                self.last_valid_time, self.last_valid_pos = now, (tx, ty)
                self.is_tracking = True

    def map_gaze(self, rx, ry):
        if not self.cal_grid: return rx, ry
        nume_x, nume_y, deno = 0.0, 0.0, 0.0
        for pt in self.cal_grid:
            dist = math.hypot(rx - pt['raw_x'], ry - pt['raw_y'])
            w = 1.0 / (max(dist, 1e-5) ** 2)
            nume_x += w * (pt['target_x'] - pt['raw_x'])
            nume_y += w * (pt['target_y'] - pt['raw_y'])
            deno += w
        return rx + nume_x / deno, ry + nume_y / deno

    def mouse_loop(self):
        while self.running:
            now = time.time()
            with self.lock:
                if self.is_tracking and (now - self.last_valid_time > 0.2):
                    self.is_tracking = False
                    self.last_invalid_time = now # 추적 유실 시점 기록 (보호막 해제 타이머 리셋용)
                    # 유실 순간의 안정된 커서 위치 저장 (클릭 타겟용)
                    self.stable_click_pos = (self.cur_x, self.cur_y)
                
                # 클릭 지연 시간 동안 마우스 이동 억제
                in_click_delay = (now - self.last_click_time < self.click_delay)
                
                if self.zoom_mode and self.zoom_win:
                    try: self.zoom_win.update_guide(self.cur_x, self.cur_y)
                    except: self.zoom_mode = False
                
                if self.tracking_enabled and len(self.gaze_buffer) > 0:
                    # 클릭 지연 시간 동안에는 좌표 계산 및 이동 완전 중단
                    if not in_click_delay:
                        tx = sum(p[0] for p in self.gaze_buffer) / len(self.gaze_buffer)
                        ty = sum(p[1] for p in self.gaze_buffer) / len(self.gaze_buffer)
                        alpha = max(0.01, min(1.0, getattr(self, 'smooth', 0.1)))
                        self.cur_x += (tx - self.cur_x) * alpha
                        self.cur_y += (ty - self.cur_y) * alpha
                        if not self.zoom_mode:
                            user32.SetCursorPos(max(0, min(SCREEN_W, int(self.cur_x))), 
                                               max(0, min(SCREEN_H, int(self.cur_y))))
            time.sleep(0.01)

    def execute_click(self, x, y):
        """
        최종 산출된 (x, y) 좌표로 OS의 마우스 포인터를 순간 이동시킨 후, 
        액션 바에 예약된 동작(일반클릭, 더블클릭, 우클릭, 드래그 온/오프)을 실행합니다.
        
        주의: UI 프레임워크(예: Tkinter)나 일부 웹 브라우저가 너무 빠른 가상 클릭을 무시하는 현상을 막기 위해,
        MOUSEEVENTF_LEFTDOWN(누름)과 LEFTUP(뗌) 사이에 0.05초 이상의 딜레이를 주입하여 
        정상적인 물리 마우스의 클릭 리듬을 모사합니다.
        """
        user32.SetCursorPos(x, y)
        # 위치 이동 후 OS가 좌표를 완전히 인지할 때까지 아주 잠깐 대기
        time.sleep(0.02)
        
        action = getattr(self, "next_click_action", "left")
        
        if action == "left":
            user32.mouse_event(2, 0, 0, 0, 0) # MOUSEEVENTF_LEFTDOWN
            time.sleep(0.05) # UI 버튼 등에서 빠른 클릭을 인지하도록 딜레이 추가
            user32.mouse_event(4, 0, 0, 0, 0) # MOUSEEVENTF_LEFTUP
        elif action == "double":
            user32.mouse_event(2, 0, 0, 0, 0)
            time.sleep(0.02)
            user32.mouse_event(4, 0, 0, 0, 0)
            time.sleep(0.05)
            user32.mouse_event(2, 0, 0, 0, 0)
            time.sleep(0.02)
            user32.mouse_event(4, 0, 0, 0, 0)
            self.next_click_action = "left" # 초기화
        elif action == "right":
            user32.mouse_event(8, 0, 0, 0, 0) # RIGHTDOWN
            time.sleep(0.05) 
            user32.mouse_event(0x10, 0, 0, 0, 0) # RIGHTUP
            self.next_click_action = "left" # 초기화
        elif action == "drag":
            if getattr(self, "is_dragging", False):
                # 두 번째 클릭: 드래그 종료 (Drop)
                user32.mouse_event(4, 0, 0, 0, 0) # LEFTUP
                self.is_dragging = False
                self.next_click_action = "left" # 초기화
            else:
                # 첫 번째 클릭: 드래그 시작 (Drag)
                user32.mouse_event(2, 0, 0, 0, 0) # LEFTDOWN
                self.is_dragging = True
        
        # 사운드 피드백
        if action == "drag" and self.is_dragging:
            winsound.Beep(500, 100) # 드래그 시작음
        elif action == "drag" and not self.is_dragging:
            winsound.Beep(400, 100) # 드래그 종료음
        else:
            winsound.Beep(440, 100) # 기본 클릭음

    def enter_zoom_mode(self, x, y):
        self.zoom_mode = True
        winsound.Beep(600, 100); winsound.Beep(600, 100)
        self.root.after(0, self._create_zoom_win, x, y)

    def _create_zoom_win(self, x, y):
        if self.zoom_win: self.zoom_win.destroy()
        self.zoom_win = ZoomWindow(self.root, (x, y), self.zoom_size, self.zoom_scale)

    def exit_zoom_mode(self):
        self.zoom_mode = False
        if self.zoom_win:
            # 즉각적으로 화면에서 숨김 (클릭 간섭 방지)
            self.root.after(0, self.zoom_win.withdraw)
            # 이후 안전하게 창 파괴
            self.root.after(0, self.zoom_win.destroy)
            self.zoom_win = None

    def toggle_tracking(self):
        self.tracking_enabled = not self.tracking_enabled
        self.toggle_btn.config(text="ON" if self.tracking_enabled else "OFF", 
                               bg="#2ecc71" if self.tracking_enabled else "#95a5a6")

    def run_calibration(self):
        if os.path.exists("calibration.exe"): subprocess.Popen(["calibration.exe"])
        else: subprocess.Popen([sys.executable, "calibration.py"])

    def on_exit(self):
        self.running = False
        self.root.quit(); self.cleanup(); sys.exit(0)

    def url_receiver_func(self, url, user_data):
        if not self.device_url: self.device_url = url

    def run(self):
        tobii.tobii_api_create(ctypes.byref(self.api), None, None)
        tobii.tobii_enumerate_local_device_urls(self.api, self.url_cb, None)
        
        self.running = True
        if self.device_url:
            tobii.tobii_device_create(self.api, self.device_url, TOBII_FIELD_OF_USE_INTERACTIVE, ctypes.byref(self.device))
            tobii.tobii_gaze_point_subscribe(self.device, self.gaze_cb, None)
            threading.Thread(target=self.tobii_process_loop, daemon=True).start()
        else:
            print("Tobii device not found. Running in Test Mode.")
            
        threading.Thread(target=self.mouse_loop, daemon=True).start()
        self.root.mainloop()

    def tobii_process_loop(self):
        devs = (c_void_p * 1)(self.device)
        while self.running:
            tobii.tobii_wait_for_callbacks(1, devs)
            tobii.tobii_device_process_callbacks(self.device)

    def cleanup(self):
        if self.device: tobii.tobii_gaze_point_unsubscribe(self.device); tobii.tobii_device_destroy(self.device)
        if self.api: tobii.tobii_api_destroy(self.api)

if __name__ == "__main__":
    TobiiMouseApp().run()
