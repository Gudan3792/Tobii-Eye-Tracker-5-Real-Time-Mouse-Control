
import tkinter as tk
import ctypes
import threading
import time
import os
import sys
import collections
import configparser
import math

# --- Ensure working directory is the script/EXE location ---
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))
else:
    os.path.chdir(os.path.dirname(os.path.abspath(__file__)))

# --- Constants & Enums (Copied from tobii_native.py for portability) ---
TOBII_ERROR_NO_ERROR = 0
TOBII_FIELD_OF_USE_INTERACTIVE = 1
TOBII_VALIDITY_VALID = 1

# --- Load DLL ---
def get_dll_path(filename):
    # If running as PyInstaller EXE, use sys._MEIPASS if bundled, 
    # but since we are not bundling it (onefile), it should be next to the EXE.
    # sys.executable gives the path to the EXE when frozen.
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, filename)

DLL_PATH = get_dll_path("tobii_stream_engine.dll")

try:
    tobii = ctypes.CDLL(DLL_PATH)
except Exception as e:
    # Final fallback attempt with just filename
    try:
        tobii = ctypes.CDLL("tobii_stream_engine.dll")
    except Exception:
        print(f"Error loading DLL ({DLL_PATH}): {e}")
        sys.exit(1)

# --- Structures & Types ---
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

# Function Prototypes
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


# --- 안구 조준점 보정(캘리브레이션) 앱 클래스 ---
# 사용자의 시선과 실제 모니터 좌표 간의 오차를 바로잡아주는 12포인트 캘리브레이션 로직입니다.
class CalibrationApp:
    """
    최초 실행 시 검은색 전체 화면에 12개의 점(4x3 배열)을 차례대로 띄웁니다.
    사용자가 한 점을 약 1.5초간 안정적으로 응시하면(Fixation), 
    해당 시점의 생체 시선 데이터(Raw Gaze)들을 모아 중앙값을 구하고 '실제로 화면에 표시된 점의 위치'와의 차이를 계산하여 구성 파일에 저장합니다.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Eye Tracker Calibration")
        self.root.attributes('-fullscreen', True)
        self.root.configure(bg='black')
        self.root.bind("<Escape>", self.exit_app)
        self.root.bind("<space>", self.skip_step_event) # Added skip in case of failure
        self.root.bind("<Return>", self.save_and_exit_event)

        self.width = self.root.winfo_screenwidth()
        self.height = self.root.winfo_screenheight()

        self.canvas = tk.Canvas(root, width=self.width, height=self.height, bg='black', highlightthickness=0)
        self.canvas.pack()

        # 12-Point Calibration Grid (4x3)
        self.points = [
            (0.1, 0.1), (0.36, 0.1), (0.63, 0.1), (0.9, 0.1), # Row 1
            (0.1, 0.5), (0.36, 0.5), (0.63, 0.5), (0.9, 0.5), # Row 2
            (0.1, 0.9), (0.36, 0.9), (0.63, 0.9), (0.9, 0.9)  # Row 3
        ]
        self.current_point_idx = 0
        self.calibration_results = {} # Index -> (avg_x, avg_y)

        # Auto-collection State
        self.focus_start_time = 0
        self.is_focusing = False
        self.collecting = False
        self.gaze_data = [] # List of (x, y) for current point
        self.recent_gaze_history = collections.deque(maxlen=30) # ~1 sec at 30Hz
        self.verifying = False

        # Tobii Setup
        self.api = c_void_p(None)
        self.device = c_void_p(None)
        self.device_url = None
        self.running = False
        self.lock = threading.Lock()
        
        self.url_cb = URL_RECEIVER_T(self.url_receiver_func)
        self.gaze_cb = GAZE_CALLBACK_T(self.gaze_callback_func)
        
        self.init_tobii()
        
        # Start UI
        self.show_point()

    def init_tobii(self):
        print("Initializing Tobii...")
        res = tobii.tobii_api_create(ctypes.byref(self.api), None, None)
        if res != TOBII_ERROR_NO_ERROR:
            print("Failed to create API")
            return

        tobii.tobii_enumerate_local_device_urls(self.api, self.url_cb, None)
        
        self.running = True
        if not self.device_url:
            print("No device found. Running Calibration in Test Mode.")
            return

        tobii.tobii_device_create(self.api, self.device_url, TOBII_FIELD_OF_USE_INTERACTIVE, ctypes.byref(self.device))
        tobii.tobii_gaze_point_subscribe(self.device, self.gaze_cb, None)
        
        self.thread = threading.Thread(target=self.tobii_loop, daemon=True)
        self.thread.start()

    def url_receiver_func(self, url, user_data):
        if not self.device_url:
            self.device_url = url
            print(f"Found device: {url.decode('utf-8')}")

    def gaze_callback_func(self, gaze_point_ptr, user_data):
        data = gaze_point_ptr.contents
        if data.validity == TOBII_VALIDITY_VALID:
            rx, ry = data.position_xy[0], data.position_xy[1]
            self.latest_gaze_raw = (rx, ry)
            
            with self.lock:
                self.recent_gaze_history.append((rx, ry))
                if self.collecting:
                    self.gaze_data.append((rx, ry))

    def tobii_loop(self):
        device_ptr_array = (c_void_p * 1)(self.device)
        p_devices = ctypes.cast(device_ptr_array, ctypes.POINTER(c_void_p))
        
        while self.running:
            tobii.tobii_wait_for_callbacks(1, p_devices)
            tobii.tobii_device_process_callbacks(self.device)

    def draw_circle(self, x, y, r=20, color="yellow"):
        x_px = int(x * self.width)
        y_px = int(y * self.height)
        self.canvas.create_oval(x_px-r, y_px-r, x_px+r, y_px+r, fill=color, outline="white", tags="point")

    def show_point(self):
        self.canvas.delete("all")
        self.is_focusing = False
        self.focus_start_time = 0
        self.recent_gaze_history.clear()
        
        if self.current_point_idx < len(self.points):
            pt = self.points[self.current_point_idx]
            self.draw_circle(pt[0], pt[1])
            self.canvas.create_text(self.width/2, self.height - 50, text="Look at the circle steadily to calibrate", fill="white", font=("Arial", 20))
            self.canvas.create_text(self.width/2, self.height - 20, text="(Press SPACE to skip this point manually)", fill="gray", font=("Arial", 12))
            
            # Start monitoring fixation and drawing live cursor
            self.check_fixation()
            self.draw_live_cursor()
        else:
            self.finish_calibration()

    def check_fixation(self):
        """
        초당 수십 번씩 사용자의 안구 시선(최근 15프레임 분량)을 확인하여, 시선이 바깥으로 크게 튀지 않고
        노란색 원 반경 영역 내에 좁게(안정적으로) 고정되어 있는지 지속적으로 판별합니다.
        시선이 안정적이라고 판단되면 파란색 게이지(Progress Arc)를 채우며, 1.5초를 채우면 데이터를 수집/완료합니다.
        """
        if self.current_point_idx >= len(self.points) or self.verifying:
            return

        with self.lock:
            history = list(self.recent_gaze_history)
            
        current_time = time.time()
        is_stable = False
        
        # Check if we have enough data (e.g. at least 15 points = ~0.5 sec)
        if len(history) > 15:
            xs = [p[0] for p in history]
            ys = [p[1] for p in history]
            
            width = max(xs) - min(xs)
            height = max(ys) - min(ys)
            
            # If the bounding box of recent gaze is small enough, consider it "stable" fixation
            # Adjust the threshold (0.05 ~ 5% of screen) depending on noise
            if width < 0.08 and height < 0.08:
                is_stable = True

        if is_stable:
            if not self.is_focusing:
                self.is_focusing = True
                self.focus_start_time = current_time
                print("Fixation started...")
            
            # Calculate duration
            focused_duration = current_time - self.focus_start_time
            required_duration = 1.5 # seconds
            
            # Draw progress arc
            progress = min(1.0, focused_duration / required_duration)
            self.draw_progress_arc(progress)
            
            if focused_duration >= required_duration:
                # Trigger collection
                print("Fixation complete! Collecting data...")
                self.trigger_collection()
                return # Stop checking for this point
        else:
            if self.is_focusing:
                print("Fixation lost.")
                self.is_focusing = False
                self.canvas.delete("progress_arc")

        # Repeat check every 50ms
        self.root.after(50, self.check_fixation)
        
    def draw_live_cursor(self):
        if hasattr(self, 'latest_gaze_raw') and self.latest_gaze_raw:
             rx, ry = self.latest_gaze_raw
             px, py = int(rx * self.width), int(ry * self.height)
             
             # Visualize raw (uncalibrated) cursor during collection
             self.canvas.delete("live_cursor")
             self.canvas.create_oval(px-8, py-8, px+8, py+8, fill="gray", outline="white", tag="live_cursor")
             
        # Only draw if we haven't reached verification and aren't done
        if not self.verifying and self.current_point_idx < len(self.points):
            self.root.after(30, self.draw_live_cursor)

    def draw_progress_arc(self, progress):
        self.canvas.delete("progress_arc")
        if progress <= 0: return
        
        pt = self.points[self.current_point_idx]
        x_px = int(pt[0] * self.width)
        y_px = int(pt[1] * self.height)
        r = 35 # slightly larger than the yellow circle

        extent = -(progress * 360) # Draw clockwise
        self.canvas.create_arc(x_px-r, y_px-r, x_px+r, y_px+r, start=90, extent=extent, outline="cyan", width=5, style=tk.ARC, tags="progress_arc")

    def trigger_collection(self):
        self.canvas.delete("progress_arc")
        self.canvas.create_text(int(self.points[self.current_point_idx][0] * self.width), int(self.points[self.current_point_idx][1] * self.height) - 50, text="OK!", fill="green", font=("Arial", 16), tags="ok_text")
        self.root.update()
        
        # We already have stable data in history! Let's just use it instead of waiting another second.
        with self.lock:
            valid_points = list(self.recent_gaze_history)
        
        if len(valid_points) > 0:
            avg_x = sum(p[0] for p in valid_points) / len(valid_points)
            avg_y = sum(p[1] for p in valid_points) / len(valid_points)
            self.calibration_results[self.current_point_idx] = (avg_x, avg_y)
            print(f"Point {self.current_point_idx}: Auto-Collected Avg Gaze ({avg_x:.4f}, {avg_y:.4f})")
        else:
             self.calibration_results[self.current_point_idx] = (0.5, 0.5)

        # Brief pause to show OK! text
        self.root.after(500, self.move_to_next)

    def move_to_next(self):
        self.current_point_idx += 1
        self.show_point()

    def skip_step_event(self, event):
        if self.current_point_idx >= len(self.points) or self.verifying:
            return
            
        print(f"Skipping point {self.current_point_idx} manually")
        self.calibration_results[self.current_point_idx] = (0.5, 0.5) # Default fallback
        self.current_point_idx += 1
        self.show_point()

    def finish_calibration(self):
        self.canvas.delete("all")
        self.canvas.create_text(self.width/2, self.height/2, text="Calculating Mapping Data...", fill="white", font=("Arial", 30))
        self.root.update()
        
        try:
            # Save raw to target mapping in temp storage
            self.temp_cal_grid = []
            
            for i, target_pt in enumerate(self.points):
                if i in self.calibration_results:
                    raw_pt = self.calibration_results[i]
                else:
                    # Fallback to perfect hit if not collected/skipped
                    raw_pt = target_pt
                    
                self.temp_cal_grid.append({
                    'raw_x': raw_pt[0], 'raw_y': raw_pt[1],
                    'target_x': target_pt[0], 'target_y': target_pt[1]
                })

            # Show collected points for debug
            self.draw_debug_points()
            
            self.canvas.create_text(self.width/2, self.height - 150, text="TEST MODE: Look around to verify.", fill="white", font=("Arial", 20))
            self.canvas.create_text(self.width/2, self.height - 100, text="Automatically SAVING and EXITING in 5 seconds...", fill="yellow", font=("Arial", 16))
            self.canvas.create_text(self.width/2, self.height - 60, text="(Press Enter to Save Now, Esc to Cancel)", fill="gray", font=("Arial", 12))
            
            # Start verification loop
            self.verifying = True
            self.verify_loop()
            
            # Auto-save timer
            self.root.after(5000, self.save_and_exit)
            
        except Exception as e:
            self.canvas.create_text(self.width/2, self.height/2, text=f"Error: {e}", fill="red", font=("Arial", 20))

    def draw_debug_points(self):
        # Draw targets (gray) and measured (red)
        targets = self.points
        for i, pt in enumerate(targets):
            tx, ty = pt[0] * self.width, pt[1] * self.height
            self.canvas.create_oval(tx-10, ty-10, tx+10, ty+10, outline="gray")
            
            if i in self.calibration_results:
                mx, my = self.calibration_results[i]
                mx_px, my_px = mx * self.width, my * self.height
                self.canvas.create_oval(mx_px-5, my_px-5, mx_px+5, my_px+5, fill="red", outline="red")
                self.canvas.create_line(tx, ty, mx_px, my_px, fill="gray", dash=(4, 4))
                
    def map_gaze(self, rx, ry):
        """
        [수학적 보정 핵심 로직: 역거리 가중치 보간법 (Inverse Distance Weighting, IDW)]
        수집 완료된 12개의 캘리브레이션 오차 스펙트럼 데이터를 기반으로 구동됩니다.
        현재 사용자가 쳐다보고 있는 생체 원시 좌표(rx, ry)를 기준 삼아, 가장 가까운 주변 점들의 왜곡 오차값을 추출한 뒤 
        역산(가까울수록 높은 가중치 부여)하여 최종적으로 가장 완벽한 화면 좌표(모니터 마우스 위치)로 맵핑(Mapping)시켜줍니다.
        """
        if not hasattr(self, 'temp_cal_grid') or not self.temp_cal_grid:
            return rx, ry

        nume_x, nume_y = 0.0, 0.0
        deno = 0.0
        
        p = 2.0  # Power param

        for pt in self.temp_cal_grid:
            dist = math.hypot(rx - pt['raw_x'], ry - pt['raw_y'])
            
            # If we are exactly on a mapping point, return it precisely
            if dist < 1e-5:
                # Add tiny offset to avoid div by zero
                dist = 1e-5
                
            w = 1.0 / (dist ** p)
            
            # We want to know how far the target is from raw for EACH point,
            # then interpolate this displacement vector
            dx = pt['target_x'] - pt['raw_x']
            dy = pt['target_y'] - pt['raw_y']
            
            nume_x += w * dx
            nume_y += w * dy
            deno += w
            
        # Global interpolated displacement
        disp_x = nume_x / deno
        disp_y = nume_y / deno
        
        # Apply displacement
        return rx + disp_x, ry + disp_y

    def verify_loop(self):
        """
        12개의 캘리브레이션 포인트 수집이 모두 끝난 직후, ini 파일(설정 파일)에 영구 저장하기 전에 
        방금 조정한 보간법(map_gaze)이 정상적으로 눈을 따라가는지 십자선 커서로 테스트 해보는 '시선 검증(Verification)' 전용 무한루프입니다.
        """
        if not self.verifying:
            return
            
        # Get latest gaze
        gaze = None
        with self.lock:
            if self.gaze_data: # Use recent buffer
                gaze = self.gaze_data[-1] 
                
        # Actually need real-time gaze, so we need to read 'latest' from callback
        # My callback appends to list. Let's just peek last item.
        # But list grows indefinitely? No, we cleared it.
        # Wait, gaze_callback appends if self.collecting is True.
        # verification mode needs collecting too.
        
        # Enable collection implicitly or handle it?
        # Let's change gaze_callback to always update 'latest_gaze'
        
        if hasattr(self, 'latest_gaze_raw') and self.latest_gaze_raw:
             rx, ry = self.latest_gaze_raw
             
             # Apply IDW Map
             mx, my = self.map_gaze(rx, ry)
             
             fx = mx * self.width
             fy = my * self.height
             
             # Visualize cursor
             self.canvas.delete("cursor")
             self.canvas.create_oval(fx-15, fy-15, fx+15, fy+15, fill="cyan", outline="white", tag="cursor")
             self.canvas.create_line(fx-20, fy, fx+20, fy, fill="cyan", tag="cursor")
             self.canvas.create_line(fx, fy-20, fx, fy+20, fill="cyan", tag="cursor")

        self.root.after(20, self.verify_loop)

    # Removed duplicate gaze_callback_func we accidentally left in before

    # Removed duplicate next_step_event

    def save_and_exit_event(self, event):
        if hasattr(self, 'verifying') and self.verifying:
            self.save_and_exit()

    def save_and_exit(self):
        self.save_ini()
        print("Saved IDW Grid data!")
        self.exit_app(None)

    def save_ini(self):
        config = configparser.ConfigParser()
        ini_path = 'eye_setting.ini'
        
        # 기존 설정 읽기 (다른 섹션 보존을 위해)
        if os.path.exists(ini_path):
            try:
                with open(ini_path, 'r', encoding='utf-8-sig') as f:
                    config.read_file(f)
            except Exception:
                try:
                    with open(ini_path, 'r', encoding='cp949') as f:
                        config.read_file(f)
                except Exception:
                    pass
            
        if 'CalibrationGrid' not in config:
            config['CalibrationGrid'] = {}
            
        config['CalibrationGrid']['point_count'] = str(len(self.temp_cal_grid))
        
        for i, pt in enumerate(self.temp_cal_grid):
            config['CalibrationGrid'][f'raw_x_{i}'] = f"{pt['raw_x']:.5f}"
            config['CalibrationGrid'][f'raw_y_{i}'] = f"{pt['raw_y']:.5f}"
            config['CalibrationGrid'][f'target_x_{i}'] = f"{pt['target_x']:.5f}"
            config['CalibrationGrid'][f'target_y_{i}'] = f"{pt['target_y']:.5f}"

        # Calibration 섹션 보존 및 기본값 설정
        if 'Calibration' not in config:
            config['Calibration'] = {}
        
        # 값이 없는 경우에만 기본값 설정 (기존 값 보존)
        if 'smooth' not in config['Calibration']:
             config['Calibration']['smooth'] = "0.1"
        if 'avg_samples' not in config['Calibration']:
             config['Calibration']['avg_samples'] = "10"
        if 'zoom_min' not in config['Calibration']:
             config['Calibration']['zoom_min'] = "1.0"
        if 'zoom_max' not in config['Calibration'] or config['Calibration']['zoom_max'] == "3.0":
             config['Calibration']['zoom_max'] = "3.0"
        if 'zoom_size' not in config['Calibration']:
             config['Calibration']['zoom_size'] = "300"
        if 'zoom_scale' not in config['Calibration']:
             config['Calibration']['zoom_scale'] = "3.0"
        if 'click_delay' not in config['Calibration']:
             config['Calibration']['click_delay'] = "0.1"
        if 'blink_min' not in config['Calibration']:
             config['Calibration']['blink_min'] = "0.3"
        if 'blink_max' not in config['Calibration']:
             config['Calibration']['blink_max'] = "1.0"
        if 'deep_sleep_threshold' not in config['Calibration']:
             config['Calibration']['deep_sleep_threshold'] = "5.0"
        if 'bottom_block_threshold' not in config['Calibration']:
             config['Calibration']['bottom_block_threshold'] = "0.96"

        # 파일 저장 (UTF-8 with BOM으로 저장하여 윈도우 호환성 극대화)
        with open(ini_path, 'w', encoding='utf-8-sig') as f:
            # Grid Data 섹션 작성
            f.write("[CalibrationGrid]\n")
            f.write(f"point_count = {config['CalibrationGrid']['point_count']}\n")
            for i in range(int(config['CalibrationGrid']['point_count'])):
                f.write(f"raw_x_{i} = {config['CalibrationGrid'][f'raw_x_{i}']}\n")
                f.write(f"raw_y_{i} = {config['CalibrationGrid'][f'raw_y_{i}']}\n")
                f.write(f"target_x_{i} = {config['CalibrationGrid'][f'target_x_{i}']}\n")
                f.write(f"target_y_{i} = {config['CalibrationGrid'][f'target_y_{i}']}\n")
            
            f.write("\n[Calibration]\n")
            f.write("; smooth: 마우스 움직임의 부드러움 정도 (0.01 ~ 1.0)\n")
            f.write("; 값이 작을수록 더 부드럽게 움직이지만 약간의 지연이 생길 수 있습니다 (추천: 0.1 ~ 0.2)\n")
            f.write(f"smooth = {config['Calibration']['smooth']}\n\n")

            f.write("; avg_samples: 몇 개의 시선 데이터를 평균 내어 사용할지 결정 (1 ~ 50)\n")
            f.write("; 값이 클수록 떨림이 적어지지만 반응 속도가 느려질 수 있습니다 (추천: 10 ~ 20)\n")
            f.write(f"avg_samples = {config['Calibration']['avg_samples']}\n\n")
            
            f.write("; blink_min: 눈을 최소한 이 시간(초)만큼 감아야 클릭으로 인식합니다 (기본: 0.3)\n")
            f.write(f"blink_min = {config['Calibration']['blink_min']}\n\n")
            f.write("; blink_max: 이 시간(초) 이상 감으면 줌 모드로 넘어갑니다. (999로 설정 시 일반 클릭 비활성화)\n")
            f.write(f"blink_max = {config['Calibration']['blink_max']}\n\n")
            
            f.write("; zoom_min: 눈을 이 시간(초) 이상 감으면 줌 모드에 진입합니다 (기본: 1.0)\n")
            f.write(f"zoom_min = {config['Calibration']['zoom_min']}\n\n")
            f.write("; zoom_max: 이 시간(초) 이상 감으면 취소됩니다. (999로 설정 시 줌 모드 비활성화 / 기본: 3.0)\n")
            f.write(f"zoom_max = {config['Calibration']['zoom_max']}\n\n")
            
            f.write("; zoom_size: 줌 모드에서 캡처할 원본 영역의 크기 (기본: 300)\n")
            f.write(f"zoom_size = {config['Calibration']['zoom_size']}\n\n")
            f.write("; zoom_scale: 줌 모드의 확대 배율 (기본: 3.0)\n")
            f.write(f"zoom_scale = {config['Calibration']['zoom_scale']}\n\n")
            
            f.write("; 클릭 시 마우스가 튀는 현상을 방지합니다.\n")
            f.write(f"click_delay = {config['Calibration']['click_delay']}\n\n")
            
            f.write("; deep_sleep_threshold: 이 시간(초) 이상 눈을 감으면 '장시간 수면' 상태로 간주합니다. (기본: 5.0)\n")
            f.write(f"deep_sleep_threshold = {config['Calibration']['deep_sleep_threshold']}\n\n")

            f.write("; bottom_block_threshold: 장기 수면 후 복귀 시 클릭을 차단할 하단 영역의 임계값 (기본: 0.96)\n")
            f.write("; 윈도우 하단 바의 절반보다 살짝 작은 영역(약 하단 4%)에서 발생하는 노이즈 클릭을 차단합니다.\n")
            f.write(f"bottom_block_threshold = {config['Calibration']['bottom_block_threshold']}\n")

    def exit_app(self, event):
        self.running = False
        self.root.destroy()
        if self.device:
            tobii.tobii_gaze_point_unsubscribe(self.device)
            tobii.tobii_device_destroy(self.device)
        if self.api:
            tobii.tobii_api_destroy(self.api)
        sys.exit()

if __name__ == "__main__":
    root = tk.Tk()
    app = CalibrationApp(root)
    root.mainloop()
