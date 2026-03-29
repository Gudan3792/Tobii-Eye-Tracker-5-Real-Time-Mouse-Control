
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
    # 파일 이름만으로 마지막 폴백 시도
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
c_uint64 = ctypes.c_uint64 # 타임스탬프는 int64_t

class tobii_gaze_point_t(ctypes.Structure):
    _fields_ = [
        ("timestamp_us", ctypes.c_longlong),
        ("validity", c_int),
        ("position_xy", c_float * 2)
    ]

# 함수 프로토타입
# tobii_error_t tobii_api_create( tobii_api_t** api, void* alloc, void* log );
tobii.tobii_api_create.argtypes = [ctypes.POINTER(c_void_p), c_void_p, c_void_p]
tobii.tobii_api_create.restype = c_int

# tobii_error_t tobii_api_destroy( tobii_api_t* api );
tobii.tobii_api_destroy.argtypes = [c_void_p]
tobii.tobii_api_destroy.restype = c_int

# tobii_error_t tobii_enumerate_local_device_urls( tobii_api_t* api, receiver, void* user_data );
URL_RECEIVER_T = ctypes.CFUNCTYPE(None, c_char_p, c_void_p)
tobii.tobii_enumerate_local_device_urls.argtypes = [c_void_p, URL_RECEIVER_T, c_void_p]
tobii.tobii_enumerate_local_device_urls.restype = c_int

# tobii_error_t tobii_device_create( tobii_api_t* api, char const* url, int field_of_use, tobii_device_t** device );
tobii.tobii_device_create.argtypes = [c_void_p, c_char_p, c_int, ctypes.POINTER(c_void_p)]
tobii.tobii_device_create.restype = c_int

# tobii_error_t tobii_device_destroy( tobii_device_t* device );
tobii.tobii_device_destroy.argtypes = [c_void_p]
tobii.tobii_device_destroy.restype = c_int

# tobii_error_t tobii_gaze_point_subscribe( tobii_device_t* device, callback, void* user_data );
GAZE_CALLBACK_T = ctypes.CFUNCTYPE(None, ctypes.POINTER(tobii_gaze_point_t), c_void_p)
tobii.tobii_gaze_point_subscribe.argtypes = [c_void_p, GAZE_CALLBACK_T, c_void_p]
tobii.tobii_gaze_point_subscribe.restype = c_int

# tobii_error_t tobii_gaze_point_unsubscribe( tobii_device_t* device );
tobii.tobii_gaze_point_unsubscribe.argtypes = [c_void_p]
tobii.tobii_gaze_point_unsubscribe.restype = c_int

# tobii_error_t tobii_wait_for_callbacks( int device_count, tobii_device_t* const* devices );
tobii.tobii_wait_for_callbacks.argtypes = [c_int, ctypes.POINTER(c_void_p)]
tobii.tobii_wait_for_callbacks.restype = c_int

# tobii_error_t tobii_device_process_callbacks( tobii_device_t* device );
tobii.tobii_device_process_callbacks.argtypes = [c_void_p]
tobii.tobii_device_process_callbacks.restype = c_int


# --- 마우스 애플리케이션 ---

SMOOTHING_FACTOR = 0.5
user32 = ctypes.windll.user32
SCREEN_W = user32.GetSystemMetrics(0)
SCREEN_H = user32.GetSystemMetrics(1)

class TobiiMouseApp:
    def __init__(self):
        self.api = c_void_p(None)
        self.device = c_void_p(None)
        self.device_url = None
        self.running = False
        
        self.lock = threading.Lock()
        self.latest_gaze = None
        
        self.cur_x = SCREEN_W / 2
        self.cur_y = SCREEN_H / 2
        
        # 보정 오프셋
        self.offset_x = 0
        self.offset_y = 0
        self.cal_grid = []
        self.scale_x = 1.0
        self.scale_y = 1.0
        self.load_calibration()
        
        # 이동 평균을 위한 버퍼
        # 반응성 향상을 위해 15로 줄임
        self.gaze_buffer = collections.deque(maxlen=15)

        # 가비지 컬렉션을 방지하기 위해 콜백 참조 유지
        self.url_cb = URL_RECEIVER_T(self.url_receiver_func)
        self.gaze_cb = GAZE_CALLBACK_T(self.gaze_callback_func)
        
        # UI 상태
        self.tracking_enabled = True
        
        self.init_ui()

    def init_ui(self):
        self.root = tk.Tk()
        self.root.title("TobiiMouse Controller")
        self.root.geometry("300x280")
        self.root.resizable(False, False)
        self.root.configure(bg="#2c3e50")
        
        # 창을 화면 중앙에 배치
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - 300) // 2
        y = (sh - 280) // 2
        self.root.geometry(f"300x280+{x}+{y}")

        # 버튼 스타일
        btn_style = {
            "font": ("Arial", 16, "bold"),
            "height": 2,
            "width": 15,
            "bd": 4,
            "relief": "raised"
        }

        # 버튼들을 가로로 나열하기 위한 상단 프레임
        top_frame = tk.Frame(self.root, bg="#2c3e50")
        top_frame.pack(pady=20)

        self.toggle_btn = tk.Button(
            top_frame, 
            text="ON", 
            bg="#2ecc71", 
            fg="white",
            command=self.toggle_tracking,
            font=("Arial", 16, "bold"),
            height=3,
            width=8,
            bd=4,
            relief="raised"
        )
        self.toggle_btn.pack(side=tk.LEFT, padx=10)

        self.cal_btn = tk.Button(
            top_frame, 
            text="CALI", 
            bg="#3498db", 
            fg="white",
            command=self.run_calibration,
            font=("Arial", 16, "bold"),
            height=3,
            width=8,
            bd=4,
            relief="raised"
        )
        self.cal_btn.pack(side=tk.LEFT, padx=10)

        self.exit_btn = tk.Button(
            self.root, 
            text="EXIT", 
            bg="#e74c3c", 
            fg="white",
            command=self.on_exit,
            font=("Arial", 16, "bold"),
            height=2,
            width=18,
            bd=4,
            relief="raised"
        )
        self.exit_btn.pack(pady=10)

        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)

    def toggle_tracking(self):
        self.tracking_enabled = not self.tracking_enabled
        if self.tracking_enabled:
            self.toggle_btn.config(text="ON", bg="#2ecc71")
            print("Tracking Resumed")
        else:
            self.toggle_btn.config(text="OFF", bg="#95a5a6")
            print("Tracking Paused")

    def run_calibration(self):
        print("Launching Calibration...")
        try:
            # EXE 파일이 존재하면 우선 사용
            if os.path.exists("calibration.exe"):
                subprocess.Popen(["calibration.exe"])
            elif os.path.exists("run_calibration.bat"):
                subprocess.Popen(["run_calibration.bat"], shell=True)
            else:
                # 파이썬으로 폴백
                subprocess.Popen([sys.executable, "calibration.py"])
        except Exception as e:
            print(f"Error launching calibration: {e}")

    def on_exit(self):
        print("Exiting...")
        self.running = False
        self.root.quit()
        self.root.destroy()
        self.cleanup()
        sys.exit(0)

    def load_calibration(self):
        ini_path = "eye_setting.ini"
        if not os.path.exists(ini_path):
             print(f"{ini_path} not found. Reverting to default.")
             self.cal_grid = []
             return

        try:
            mtime = os.path.getmtime(ini_path)
            # 파일이 최신인 경우에만 다시 로드
            if hasattr(self, 'cal_mtime') and mtime <= self.cal_mtime:
                return

            config = configparser.ConfigParser()
            config.read(ini_path, encoding='utf-8')
            
            # 전역 설정 로드
            if 'Calibration' in config:
                self.smooth = float(config['Calibration'].get('smooth', 0.15))
                
                # 동적 평균 샘플 업데이트
                new_avg = int(config['Calibration'].get('avg_samples', 15))
                if not hasattr(self, 'avg_samples') or new_avg != self.avg_samples:
                    self.avg_samples = new_avg
                    with self.lock:
                        self.gaze_buffer = collections.deque(maxlen=self.avg_samples)
                        print(f"Update: Buffer Size set to {self.avg_samples}")
            else:
                self.smooth = 0.15
                self.avg_samples = 15
                
            # 그리드 데이터 로드
            self.cal_grid = []
            if 'CalibrationGrid' in config:
                count = int(config['CalibrationGrid'].get('point_count', '0'))
                for i in range(count):
                    try:
                        pt = {
                            'raw_x': float(config['CalibrationGrid'][f'raw_x_{i}']),
                            'raw_y': float(config['CalibrationGrid'][f'raw_y_{i}']),
                            'target_x': float(config['CalibrationGrid'][f'target_x_{i}']),
                            'target_y': float(config['CalibrationGrid'][f'target_y_{i}'])
                        }
                        self.cal_grid.append(pt)
                    except KeyError:
                        pass
                print(f"Reloaded {len(self.cal_grid)}-Point Calibration Grid | Smooth: {self.smooth}")
            else:
                print("No [CalibrationGrid] found, falling back to basic if present...")
                # 이전 버전 호환성을 위해 구형 방식으로 폴백
                if 'Calibration' in config:
                    left = int(config['Calibration'].get('left', 0))
                    right = int(config['Calibration'].get('right', 0))
                    up = int(config['Calibration'].get('up', 0))
                    down = int(config['Calibration'].get('down', 0))
                    self.scale_x = float(config['Calibration'].get('scale_x', 1.0))
                    self.scale_y = float(config['Calibration'].get('scale_y', 1.0))
                    self.offset_x = right - left
                    self.offset_y = down - up
                    
            self.cal_mtime = mtime
            
        except Exception as e:
            print(f"Error reading {ini_path}: {e}")

    def url_receiver_func(self, url, user_data):
        if not self.device_url:
            self.device_url = url
            print(f"Found device: {url.decode('utf-8')}")

    def map_gaze(self, rx, ry):
        if not self.cal_grid:
            # 이전 동작으로 폴백
            sx = (rx - 0.5) * self.scale_x + 0.5
            sy = (ry - 0.5) * self.scale_y + 0.5
            return sx, sy

        nume_x, nume_y = 0.0, 0.0
        deno = 0.0
        p = 2.0  # 거듭제곱 매개변수

        for pt in self.cal_grid:
            dist = math.hypot(rx - pt['raw_x'], ry - pt['raw_y'])
            if dist < 1e-5:
                dist = 1e-5
                
            w = 1.0 / (dist ** p)
            dx = pt['target_x'] - pt['raw_x']
            dy = pt['target_y'] - pt['raw_y']
            
            nume_x += w * dx
            nume_y += w * dy
            deno += w
            
        disp_x = nume_x / deno
        disp_y = nume_y / deno
        
        return rx + disp_x, ry + disp_y

    def gaze_callback_func(self, gaze_point_ptr, user_data):
        data = gaze_point_ptr.contents
        if data.validity == TOBII_VALIDITY_VALID:
            # 정규화된 원본 좌표
            rx = data.position_xy[0]
            ry = data.position_xy[1]
            
            # IDW 매핑 적용
            mx, my = self.map_gaze(rx, ry)
            
            # 픽셀로 변환
            tx = int(mx * SCREEN_W)
            ty = int(my * SCREEN_H)
            
            with self.lock:
                # 버퍼에 좌표 추가
                self.gaze_buffer.append((tx, ty))

    def check_error(self, res, msg):
        if res != TOBII_ERROR_NO_ERROR:
            raise RuntimeError(f"{msg} failed with error code: {res}")

    def mouse_loop(self):
        print("Mouse thread started.")
        loop_count = 0
        while self.running:
            # 약 1초마다 보정 설정 확인 (100 * 10ms)
            loop_count += 1
            if loop_count % 100 == 0:
                self.load_calibration()
                
            target = None
            with self.lock:
                if len(self.gaze_buffer) > 0:
                    sum_x = sum(p[0] for p in self.gaze_buffer)
                    sum_y = sum(p[1] for p in self.gaze_buffer)
                    count = len(self.gaze_buffer)
                    target = (sum_x / count, sum_y / count)
            
            if target and self.tracking_enabled:
                tx, ty = target
                
                # 부드러움 처리
                # 로드된 부드러움 계수 사용 (기본값 0.15)
                # 부드러움 계수가 적정 범위 내에 있는지 확인
                alpha = max(0.01, min(1.0, getattr(self, 'smooth', 0.15)))
                
                self.cur_x += (tx - self.cur_x) * alpha
                self.cur_y += (ty - self.cur_y) * alpha
                
                # 기본 보정으로 폴백하는 경우 레거시 오프셋 적용
                # (IDW 그리드 사용 시 보통 0)
                final_x = self.cur_x
                final_y = self.cur_y
                if not self.cal_grid:
                    final_x += self.offset_x
                    final_y += self.offset_y
                
                ix = max(0, min(SCREEN_W, int(final_x)))
                iy = max(0, min(SCREEN_H, int(final_y)))
                
                user32.SetCursorPos(ix, iy)
            
            time.sleep(0.01) # 100Hz 업데이트

    def run(self):
        print("Initializing Tobii Stream Engine...")
        self.check_error(tobii.tobii_api_create(ctypes.byref(self.api), None, None), "tobii_api_create")
        
        print("Enumerating devices...")
        self.check_error(tobii.tobii_enumerate_local_device_urls(self.api, self.url_cb, None), "enumerate_urls")
        
        if not self.device_url:
            print("No devices found.")
            return

        print("Creating device...")
        self.check_error(tobii.tobii_device_create(self.api, self.device_url, TOBII_FIELD_OF_USE_INTERACTIVE, ctypes.byref(self.device)), "device_create")

        print("Subscribing to gaze...")
        self.check_error(tobii.tobii_gaze_point_subscribe(self.device, self.gaze_cb, None), "subscribe")

        self.running = True
        
        # 스레드 1: 마우스 이동 및 부드러움 처리
        t_mouse = threading.Thread(target=self.mouse_loop, daemon=True)
        t_mouse.start()
        
        # 스레드 2: Tobii 콜백 처리 (블로킹)
        t_tobii = threading.Thread(target=self.tobii_process_loop, daemon=True)
        t_tobii.start()
        
        print("Running! Exit via UI button.")
        # UI 루프 시작 (메인 스레드)
        self.root.mainloop()

    def tobii_process_loop(self):
        device_ptr_array = (c_void_p * 1)(self.device)
        p_devices = ctypes.cast(device_ptr_array, ctypes.POINTER(c_void_p))
        
        while self.running:
            # 콜백 대기 (블로킹)
            res = tobii.tobii_wait_for_callbacks(1, p_devices)
            if res != TOBII_ERROR_NO_ERROR:
                pass 
            
            # 콜백 처리
            res = tobii.tobii_device_process_callbacks(self.device)
            if res != TOBII_ERROR_NO_ERROR:
                print(f"Process callback error: {res}")

    def cleanup(self):
        self.running = False
        if self.device:
            tobii.tobii_gaze_point_unsubscribe(self.device)
            tobii.tobii_device_destroy(self.device)
        if self.api:
            tobii.tobii_api_destroy(self.api)
        print("Cleanup done.")

if __name__ == "__main__":
    app = TobiiMouseApp()
    app.run()
