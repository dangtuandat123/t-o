import tkinter as tk
from tkinter import colorchooser, messagebox, ttk
import pystray
from PIL import Image, ImageDraw, ImageGrab
import threading
import ctypes
from ctypes import wintypes
import sys
import json
import os
import time
import datetime
import wave
import base64

try:
    import pyaudio
except ImportError:
    pyaudio = None

try:
    import requests
except ImportError:
    requests = None

CONFIG_FILE = "config.json"
DEFAULT_AI_MODEL = "xiaomi/mimo-v2-omni"
AVAILABLE_MODELS = [
    "xiaomi/mimo-v2-omni",
    "google/gemini-3-flash-preview",
    "google/gemini-3.1-flash-lite-preview",
    "google/gemini-3.1-pro-preview"
]

if sys.platform == 'win32':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

    WH_MOUSE_LL = 14
    WM_XBUTTONDOWN = 0x020B
    WM_XBUTTONUP = 0x020C
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    class MSLLHOOKSTRUCT(ctypes.Structure):
        _fields_ = [("pt", wintypes.POINT),
                    ("mouseData", wintypes.DWORD),
                    ("flags", wintypes.DWORD),
                    ("time", wintypes.DWORD),
                    ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]
                    
    HOOKPROC = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)
    
    user32.SetWindowsHookExW.argtypes = [ctypes.c_int, HOOKPROC, wintypes.HINSTANCE, wintypes.DWORD]
    user32.SetWindowsHookExW.restype = wintypes.HHOOK

    user32.CallNextHookEx.argtypes = [wintypes.HHOOK, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM]
    user32.CallNextHookEx.restype = ctypes.c_long

    user32.UnhookWindowsHookEx.argtypes = [wintypes.HHOOK]
    user32.UnhookWindowsHookEx.restype = wintypes.BOOL

    kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
    kernel32.GetModuleHandleW.restype = wintypes.HINSTANCE

class OverlayApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.trans_color = '#010203'
        self.root.config(bg=self.trans_color)
        self.root.wm_attributes('-transparentcolor', self.trans_color)
        self.root.wm_attributes('-topmost', True)
        self.root.title("Text Overlay")
        
        self.text_sequence = []
        self.text_index = -1
        self.text_str = "..."
        
        # Default Text Settings
        self.text_color = "red"
        self.text_size = 120
        self.pos_x = None
        self.pos_y = None
        self.show_border = True 
        
        # Default Rect Settings
        self.rect_enabled = True
        self.rect_color = "yellow"
        self.rect_alpha = 0.5
        self.rect_x = 100
        self.rect_y = 100
        self.rect_w = 300
        self.rect_h = 200
        
        # AI Config
        self.api_key1 = "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"
        self.api_key2 = ""
        self.api_key3 = ""
        self.ai_model = DEFAULT_AI_MODEL
        
        self.current_key_idx = 0
        self.load_config()
        
        self.mouse5_down_time = 0
        self.mouse5_action_done = False
        
        self.mouse4_down_time = 0
        self.mouse4_action_done = False
        
        self.is_recording = False
        self.audio_frames = []
        self.p_audio = None
        self.audio_stream = None
        
        self.canvas = tk.Canvas(self.root, bg=self.trans_color, highlightthickness=0)
        self.canvas.pack()
        self.update_text_render()
        
        self.rect_win = tk.Toplevel(self.root)
        self.rect_win.overrideredirect(True)
        self.rect_win.wm_attributes('-topmost', True)
        self.update_rect_render()
        
        if self.pos_x is None or self.pos_y is None:
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            self.pos_x = (sw - self.canvas.winfo_reqwidth()) // 2
            self.pos_y = (sh - self.canvas.winfo_reqheight()) // 2
        
        self.make_clickthrough(self.root)
        self.make_clickthrough(self.rect_win)
        
        self.update_geometry()
        self.root.after(100, self.update_geometry)
        
        self.setup_tray()
        self.keep_on_top()
        
        if sys.platform == 'win32':
            self.start_mouse_hook()

    def start_mouse_hook(self):
        self._hook_thread = threading.Thread(target=self._run_hook, daemon=True)
        self._hook_thread.start()

    def _run_hook(self):
        @HOOKPROC
        def low_level_mouse_handler(nCode, wParam, lParam):
            if nCode >= 0:
                struct = ctypes.cast(lParam, ctypes.POINTER(MSLLHOOKSTRUCT))[0]
                mouseData = struct.mouseData >> 16
                
                if wParam == WM_XBUTTONDOWN:
                    if mouseData == 1: 
                        self.mouse4_down_time = time.time()
                        self.mouse4_action_done = False
                        threading.Thread(target=self.check_mouse4_hold, args=(self.mouse4_down_time,), daemon=True).start()
                        return 1 
                    elif mouseData == 2: 
                        self.mouse5_down_time = time.time()
                        self.mouse5_action_done = False
                        threading.Thread(target=self.check_mouse5_hold, args=(self.mouse5_down_time,), daemon=True).start()
                        return 1
                elif wParam == WM_XBUTTONUP:
                    if mouseData == 1:
                        if not getattr(self, 'mouse4_action_done', False):
                            self.mouse4_action_done = True
                            self.root.after(0, self.prev_text)
                        elif getattr(self, 'is_recording', False):
                            self.root.after(0, self.stop_recording)
                        return 1
                    elif mouseData == 2:
                        if not getattr(self, 'mouse5_action_done', False):
                            self.mouse5_action_done = True
                            self.root.after(0, self.next_text)
                        return 1
            return user32.CallNextHookEx(self.hook, nCode, wParam, lParam)

        self.hook_id = low_level_mouse_handler 
        kernel32.GetModuleHandleW.restype = wintypes.HMODULE
        hMod = kernel32.GetModuleHandleW(None)
        self.hook = user32.SetWindowsHookExW(WH_MOUSE_LL, self.hook_id, hMod, 0)
        
        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

    def check_mouse5_hold(self, start_time):
        time.sleep(1.0)
        if getattr(self, 'mouse5_down_time', 0) == start_time:
            if not getattr(self, 'mouse5_action_done', False):
                self.mouse5_action_done = True
                self.root.after(0, self.take_screenshot)

    def check_mouse4_hold(self, start_time):
        time.sleep(1.0)
        if getattr(self, 'mouse4_down_time', 0) == start_time:
            if not getattr(self, 'mouse4_action_done', False):
                self.mouse4_action_done = True
                self.root.after(0, self.start_recording)

    def start_recording(self):
        if pyaudio is None:
            print("Cần cài đặt thư viện 'pyaudio'.")
            return
            
        if self.is_recording:
            return
            
        self.is_recording = True
        self.audio_frames = []
        
        self.old_text_record = self.text_str
        self.old_color_record = self.text_color
        self.text_str = "🎙"
        self.update_text_render()
        
        try:
            self.p_audio = pyaudio.PyAudio()
            self.audio_stream = self.p_audio.open(format=pyaudio.paInt16,
                                                  channels=1,
                                                  rate=44100,
                                                  input=True,
                                                  frames_per_buffer=1024)
            def record_thread():
                import audioop
                while getattr(self, 'is_recording', False):
                    try:
                        data = self.audio_stream.read(1024)
                        # Giảm hệ số khuếch đại từ 4.0 (dễ rè/vỡ tiếng) xuống 2.5 để tiếng trong vắt hơn
                        amplified_data = audioop.mul(data, 2, 2.5)
                        self.audio_frames.append(amplified_data)
                    except Exception:
                        pass
                        
                # --- SAU KHI UI DỪNG THU ÂM ---
                if hasattr(self, 'audio_stream') and self.audio_stream:
                    try:
                        self.audio_stream.stop_stream()
                        self.audio_stream.close()
                    except Exception: pass
                if hasattr(self, 'p_audio') and self.p_audio:
                    try:
                        self.p_audio.terminate()
                    except Exception: pass
                    
                if len(self.audio_frames) > 0:
                    os.makedirs("recordings", exist_ok=True)
                    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = os.path.join(os.getcwd(), "recordings", f"record_{timestamp}.wav")
                    try:
                        wf = wave.open(filename, 'wb')
                        wf.setnchannels(1)
                        wf.setsampwidth(2)
                        wf.setframerate(44100)
                        wf.writeframes(b''.join(self.audio_frames))
                        wf.close()
                        print(f"Đã lưu ghi âm: {filename}")
                        
                        # Set UI chờ
                        old_color = getattr(self, 'old_color_record', "red")
                        def ui_wait():
                            self.text_str = "⏳"
                            self.update_text_render()
                        self.root.after(0, ui_wait)
                        
                        self.process_ai(
                            image_path=getattr(self, 'pending_audio_image_path', None),
                            audio_path=filename, 
                            restore_color=old_color
                        )
                    except Exception as e:
                        print("Lỗi lưu file:", e)
                        # Rollback UI khi lưu WAV thất bại
                        def ui_rollback_save():
                            self.text_str = getattr(self, 'old_text_record', "...")
                            self.text_color = getattr(self, 'old_color_record', "red")
                            self.update_text_render()
                        self.root.after(0, ui_rollback_save)
                else:
                    # Thu âm quá ngắn (0 frame) → Phục hồi UI về trạng thái cũ
                    print("[WARN] Thu âm rỗng, bỏ qua.")
                    def ui_rollback_empty():
                        self.text_str = getattr(self, 'old_text_record', "...")
                        self.text_color = getattr(self, 'old_color_record', "red")
                        self.update_text_render()
                    self.root.after(0, ui_rollback_empty)
                        
            threading.Thread(target=record_thread, daemon=True).start()
        except Exception as e:
            print("Lỗi mở Mic:", e)
            self.is_recording = False
            # Rollback UI nếu Mic lỗi (Vì Thread chưa được tạo nên không ai tự dọn dẹp)
            if hasattr(self, 'old_text_record'):
                self.text_str = self.old_text_record
                self.text_color = getattr(self, 'old_color_record', "red")
                self.update_text_render()
            if hasattr(self, 'p_audio') and self.p_audio:
                try: self.p_audio.terminate()
                except Exception: pass

    def stop_recording(self):
        if not getattr(self, 'is_recording', False):
            return
        
        # QUAN TRỌNG: Chụp ảnh TRƯỚC rồi mới hạ cờ is_recording
        # Nếu hạ cờ trước, Recording Thread sẽ thoát vòng lặp và đọc pending_audio_image_path
        # trước khi Main Thread kịp chụp ảnh xong → mất ảnh (Race Condition)
        self.pending_audio_image_path = None
        
        if self.rect_w > 0 and self.rect_h > 0:
            hidden_rect = False
            try:
                if getattr(self, 'rect_enabled', False):
                    self.rect_win.withdraw()
                    hidden_rect = True
                self.root.update()
                
                x1, y1 = self.rect_x, self.rect_y
                x2, y2 = x1 + self.rect_w, y1 + self.rect_h
                try:
                    from PIL import ImageGrab
                    img = ImageGrab.grab(bbox=(x1, y1, x2, y2), all_screens=True)
                except TypeError:
                    img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
                    
                os.makedirs("screenshots", exist_ok=True)
                t_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                self.pending_audio_image_path = os.path.join(os.getcwd(), "screenshots", f"scr_{t_str}.png")
                img.save(self.pending_audio_image_path)
            except Exception as e:
                print("Lỗi chụp màn hình kèm Audio:", e)
            finally:
                if hidden_rect and getattr(self, 'rect_enabled', False):
                    self.rect_win.deiconify()
        
        # HẠ CỜ SAU CÙNG - Recording Thread giờ mới được phép thoát vòng lặp
        self.is_recording = False

    def take_screenshot(self):
        if self.rect_w <= 0 or self.rect_h <= 0:
            return
            
        if self.rect_enabled:
            self.rect_win.withdraw()
            
        self.root.update()
        
        def do_capture():
            try:
                x1, y1 = self.rect_x, self.rect_y
                x2, y2 = x1 + self.rect_w, y1 + self.rect_h
                
                try:
                    img = ImageGrab.grab(bbox=(x1, y1, x2, y2), all_screens=True)
                except TypeError:
                    img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
                
                os.makedirs("screenshots", exist_ok=True)
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                filepath = os.path.join(os.getcwd(), "screenshots", f"capture_{timestamp}.png")
                img.save(filepath)
                
                old_text = self.text_str
                old_color = self.text_color
                self.text_str = "⏳" # Báo hiệu đang xử lý AI
                self.update_text_render()
                
                self.process_ai(image_path=filepath, restore_color=old_color)
            except Exception as e:
                print("Lỗi chụp ảnh:", e)
            finally:
                if self.rect_enabled:
                    self.rect_win.deiconify()
                
        self.root.after(100, do_capture)

    def call_openrouter(self, image_path=None, audio_path=None):
        if not requests:
            return "LỖI: Chưa cài thư viện requests"
            
        available_keys = [k.strip() for k in [self.api_key1, self.api_key2, self.api_key3] if k.strip()]
        if not available_keys:
            return "LỖI: Bạn chưa nhập OpenRouter API Key nào!"

        content = []
        if image_path:
            try:
                with open(image_path, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode('utf-8')
                content.append({
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/png;base64,{b64}"
                    }
                })
            except Exception: pass

        if audio_path:
            try:
                with open(audio_path, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode('utf-8')
                content.append({
                    "type": "input_audio",
                    "input_audio": {
                        "data": b64,
                        "format": "wav"
                    }
                })
            except Exception: pass
            
        if not content:
            return "! Lỗi: Không có file Ảnh/Âm thanh"
            
        text_prompt = "Identify ALL complete English multiple-choice questions (grammar, vocabulary, reading, TOEIC listening, error identification) in the media and solve them. Output ONLY a raw JSON object containing an 'answers' array. Example: {\"answers\": [{\"cau_hoi\": \"12\", \"dap_an\": \"A\"}, {\"cau_hoi\": \"13\", \"dap_an\": \"B\"}]}"
        if audio_path:
            text_prompt = "Listen to the TOEIC audio carefully. An image is also attached but it MIGHT be irrelevant or blank. ONLY use the image if it contains explicit TOEIC visuals (e.g., Part 1 photos or Part 3/4 graphics). Otherwise, rely purely on the audio. First, write a brief 'transcript'. Then, identify ALL correct TOEIC Listening answers. Output ONLY a JSON object containing an 'answers' array: {\"answers\": [{\"transcript\": \"...\", \"cau_hoi\": \"12\", \"dap_an\": \"A\"}]}"

        content.append({
            "type": "text",
            "text": text_prompt
        })
        
        system_role = "You are an expert English Language teacher and advanced AI data-extraction engine dedicated EXCLUSIVELY to solving English multiple-choice exercises and TOEIC Listening tests.\n\nSTRICT RULES:\n1. If the input contains multiple complete questions, you MUST solve ALL of them. Ignore ONLY the incomplete/cropped questions.\n2. Carefully analyze grammar, vocabulary context, TOEIC listening comprehension patterns, or error pinpointing before selecting the answers. Note that attached images during audio tasks might be irrelevant; prioritize audio if the image lacks context.\n3. NO conversational text or markdown formatting. Do not explain your reasoning.\n4. Output ONLY a single raw JSON object with an 'answers' array exactly like this:\n{\"answers\": [{\"transcript\": \"<Optional audio transcript>\", \"cau_hoi\": \"<Question ID>\", \"dap_an\": \"<A, B, C, or D>\"}, ...]}\n\nViolating these rules will cause a system failure. Proceed."
        
        payload = {
            "model": self.ai_model,
            "temperature": 0.0,
            "messages": [
                {
                    "role": "system",
                    "content": system_role
                },
                {
                    "role": "user",
                    "content": content
                }
            ],
            "response_format": { "type": "json_object" },
            "plugins": [
                {"id": "response-healing"}
            ]
        }

        if not hasattr(self, 'current_key_idx'):
            self.current_key_idx = 0

        for attempt in range(3):
            if self.current_key_idx >= len(available_keys):
                self.current_key_idx = 0
                
            active_key = available_keys[self.current_key_idx]
            headers = {
                "Authorization": f"Bearer {active_key}",
                "Content-Type": "application/json",
                "HTTP-Referer": "http://localhost",
                "X-OpenRouter-Title": "OverlayApp",
            }
            
            try:
                print(f"[RETRY {attempt+1}] Gửi lệnh với Key #{self.current_key_idx+1} ({active_key[:8]}...)")
                response = requests.post("https://openrouter.ai/api/v1/chat/completions", json=payload, headers=headers, timeout=45)
                data = response.json()
                print(f"\n[AI RESPONSE DUMP - THỬ LẦN {attempt+1}]", json.dumps(data, indent=2, ensure_ascii=False))
                
                if "choices" in data and len(data["choices"]) > 0:
                    msg_content = data["choices"][0]["message"].get("content", "")
                    
                    try:
                        if msg_content is None: msg_content = ""
                        
                        import re
                        blocks = re.findall(r'```(?:json)?\s*([\s\S]*?)\s*```', msg_content)
                        if blocks:
                            clean = blocks[-1].strip()
                        else:
                            last_obj = ""
                            stack = 0
                            start = -1
                            for i, char in enumerate(msg_content):
                                if char in '{[':
                                    if stack == 0: start = i
                                    stack += 1
                                elif char in '}]':
                                    stack -= 1
                                    if stack == 0 and start != -1:
                                        last_obj = msg_content[start:i+1]
                            clean = last_obj if last_obj else msg_content.strip()
                        
                        js = json.loads(clean)
                        
                        answers_list = []
                        if isinstance(js, dict):
                            if "answers" in js and isinstance(js["answers"], list):
                                answers_list = js["answers"]
                            else:
                                answers_list = [js]
                        elif isinstance(js, list):
                            answers_list = js
                            
                        results = []
                        for item in answers_list:
                            if not isinstance(item, dict): continue
                            if "transcript" in item and item["transcript"] and str(item["transcript"]).strip().lower() != "null":
                                print(f"[AI TRANSCRIPT] 👉 {item['transcript']}\n")
                                
                            so, da = None, None
                            for k, v in item.items():
                                kl = str(k).lower()
                                if 'hoi' in kl or 'number' in kl or 'socau' in kl or 'so_cau' in kl or 'cau' in kl or kl == 'id':
                                    so = str(v)
                                # Tránh dính chữ question_text, options, ...
                                if 'answer' in kl or 'dap_an' in kl or 'dapan' in kl or 'correct' in kl:
                                    if isinstance(v, dict): continue
                                    import re
                                    val_str = str(v).strip().upper()
                                    match_abcd = re.search(r'\b(A|B|C|D)\b', val_str)
                                    if match_abcd: da = match_abcd.group(1)
                                    else: da = val_str[:5]
                                    
                            if so or da:
                                results.append(f"{so or '?'} {da or '?'}".strip())
                                
                        if not results: raise ValueError("Keys missing")
                        return results if len(results) > 1 else results[0]
                    except Exception:
                        # Nếu JSON vỡ, bắt Regex chặt chẽ hơn (chỉ bắt A B C D đứng sau các từ khóa chỉ định)
                        import re
                        m_ans = re.search(r'(?i)(?:answer|đáp án|correct|dap an|chọn)[^A-D]*([A-D])\b', msg_content)
                        if m_ans:
                            ans_char = m_ans.group(1).upper()
                            m_so = re.search(r'(?i)(?:câu|question)[^\d]*(\d+)', msg_content)
                            return f"{m_so.group(1) if m_so else '?'} {ans_char}"
                            
                        # Lưới lọc cuối (Regex tìm trơ trọi)
                        m = re.search(r'(?i)\b(A|B|C|D)\b', msg_content[-50:]) # Ưu tiên phần kết luận
                        if m: return f"? {m.group(1).upper()}"
                        
                        # Không tìm ra = Lỗi JSON Model ngáo chữ -> Đổi Key + Retry
                        if attempt < 2:
                            self.current_key_idx = (self.current_key_idx + 1) % len(available_keys)
                            time.sleep(2)
                            continue
                        short_msg = str(msg_content).replace('\n', ' ')
                        return f"! LỖI JSON: {short_msg[:40].strip()}"
                else:
                    err_obj = data.get('error', {})
                    err_str = str(err_obj)
                    
                    if attempt < 2:
                        print(f"Lỗi API 401/429/500, tự động Switch qua Key tiếp theo...")
                        self.current_key_idx = (self.current_key_idx + 1) % len(available_keys)
                        time.sleep(2)
                        continue
                        
                    if 'User not found' in err_str or '401' in err_str:
                        return "! LỖI API: Cả 3 API Key đều đã chết/sai mã!"
                    return f"! LỖI API: {str(err_obj.get('message', err_obj))[:30]}"
            except Exception as e:
                print(f"[NETWORK ERROR - LẦN {attempt+1}]", e)
                if attempt < 2:
                    self.current_key_idx = (self.current_key_idx + 1) % len(available_keys)
                    time.sleep(2)
                    continue
                return "! Lỗi Mạng (Thất bại 3 lần + 3 Keys)"

    def process_ai(self, image_path=None, audio_path=None, restore_color=None):
        def worker():
            ans = self.call_openrouter(image_path, audio_path)
            if ans:
                def update_ui():
                    if restore_color and not getattr(self, 'is_recording', False):
                        self.text_color = restore_color
                    
                    if isinstance(ans, list):
                        start_idx = len(self.text_sequence)
                        for a in ans:
                            self.text_sequence.append(a)
                        
                        # Không giật quyền vẽ màn hình nếu user đang bấm Mic thu âm câu tiếp theo
                        if not getattr(self, 'is_recording', False):
                            self.text_index = start_idx
                            self.text_str = self.text_sequence[self.text_index]
                            self.update_style()
                        
                    elif isinstance(ans, str):
                        self.text_sequence.append(ans)
                        if not getattr(self, 'is_recording', False):
                            self.text_index = len(self.text_sequence) - 1
                            self.text_str = self.text_sequence[self.text_index]
                            self.update_style()
                self.root.after(0, update_ui)
        threading.Thread(target=worker, daemon=True).start()

    def update_text_render(self):
        self.canvas.delete("all")
        font_spec = ('Arial', self.text_size, 'bold')
        
        border_thick = 4 if self.show_border else 0
        pad = border_thick * 2
        
        text_id = self.canvas.create_text(border_thick, border_thick, text=self.text_str, font=font_spec, anchor='nw')
        bbox = self.canvas.bbox(text_id)
        
        if bbox:
            w = bbox[2] - bbox[0] + pad
            h = bbox[3] - bbox[1] + pad
            self.canvas.config(width=w, height=h)
        
        self.canvas.delete("all")
        x, y = border_thick, border_thick
        
        if self.show_border:
            for dx in [-border_thick, 0, border_thick]:
                for dy in [-border_thick, 0, border_thick]:
                    if dx != 0 or dy != 0:
                        self.canvas.create_text(x + dx, y + dy, text=self.text_str, font=font_spec, fill='black', anchor='nw')
                        
        self.canvas.create_text(x, y, text=self.text_str, font=font_spec, fill=self.text_color, anchor='nw')
        self.root.update_idletasks()

    def update_rect_render(self):
        if self.rect_enabled:
            self.rect_win.deiconify() 
            self.rect_win.config(bg=self.rect_color)
            self.rect_win.attributes('-alpha', self.rect_alpha)
            
            x_str = f"+{int(self.rect_x)}" if int(self.rect_x) >= 0 else str(int(self.rect_x))
            y_str = f"+{int(self.rect_y)}" if int(self.rect_y) >= 0 else str(int(self.rect_y))
            self.rect_win.geometry(f"{int(self.rect_w)}x{int(self.rect_h)}{x_str}{y_str}")
            
            self.rect_win.update_idletasks()
            self.make_clickthrough(self.rect_win)
        else:
            self.rect_win.withdraw()

    def next_text(self):
        if not self.text_sequence:
            return
        self.text_index = (self.text_index + 1) % len(self.text_sequence)
        self.text_str = self.text_sequence[self.text_index]
        self.update_style()

    def prev_text(self):
        if not self.text_sequence:
            return
        self.text_index = (self.text_index - 1) % len(self.text_sequence)
        self.text_str = self.text_sequence[self.text_index]
        self.update_style()

    def keep_on_top(self):
        self.root.attributes('-topmost', True)
        self.root.lift()
        self.rect_win.attributes('-topmost', True)
        self.rect_win.lift()
        self.root.after(1000, self.keep_on_top)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.text_color = data.get("color", "red")
                    self.text_size = data.get("size", 120)
                    self.pos_x = data.get("x", None)
                    self.pos_y = data.get("y", None)
                    self.show_border = data.get("show_border", True)
                    
                    self.rect_enabled = data.get("rect_enabled", True)
                    self.rect_color = data.get("rect_color", "yellow")
                    self.rect_alpha = float(data.get("rect_alpha", 0.5))
                    self.rect_x = data.get("rect_x", 100)
                    self.rect_y = data.get("rect_y", 100)
                    self.rect_w = data.get("rect_w", 300)
                    self.rect_h = data.get("rect_h", 200)
                    
                    # Cứu key cũ nếu bản save json phiên bản trước chỉ có 1 biến api_key
                    old_key = data.get("api_key", "")
                    self.api_key1 = data.get("api_key1", old_key if old_key else "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")
                    self.api_key2 = data.get("api_key2", "")
                    self.api_key3 = data.get("api_key3", "")
                    
                    self.ai_model = data.get("ai_model", DEFAULT_AI_MODEL)
            except Exception:
                pass

    def save_config(self):
        data = {
            "color": self.text_color,
            "size": self.text_size,
            "x": self.pos_x,
            "y": self.pos_y,
            "show_border": self.show_border,
            "rect_enabled": self.rect_enabled,
            "rect_color": self.rect_color,
            "rect_alpha": self.rect_alpha,
            "rect_x": self.rect_x,
            "rect_y": self.rect_y,
            "rect_w": self.rect_w,
            "rect_h": self.rect_h,
            "api_key1": self.api_key1,
            "api_key2": self.api_key2,
            "api_key3": self.api_key3,
            "ai_model": self.ai_model
        }
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
        except Exception:
            pass

    def make_clickthrough(self, win):
        try:
            if sys.platform == 'win32':
                hwnd = ctypes.windll.user32.GetParent(win.winfo_id())
                ex_style = ctypes.windll.user32.GetWindowLongW(hwnd, -20)
                ctypes.windll.user32.SetWindowLongW(hwnd, -20, ex_style | 0x00080000 | 0x00000020)
        except Exception:
            pass

    def update_geometry(self):
        self.root.update_idletasks()
        width = self.canvas.winfo_reqwidth()
        height = self.canvas.winfo_reqheight()
        
        x_str = f"+{int(self.pos_x)}" if self.pos_x is not None and int(self.pos_x) >= 0 else str(int(self.pos_x) if self.pos_x is not None else 0)
        y_str = f"+{int(self.pos_y)}" if self.pos_y is not None and int(self.pos_y) >= 0 else str(int(self.pos_y) if self.pos_y is not None else 0)
        
        self.root.geometry(f'{width}x{height}{x_str}{y_str}')
        self.root.update()

    def update_style(self):
        self.update_text_render()
        self.update_geometry()
        self.update_rect_render()

    def create_tray_image(self):
        image = Image.new('RGB', (64, 64), color=(255, 255, 255))
        draw = ImageDraw.Draw(image)
        draw.rectangle((16, 16, 48, 48), fill=(255, 0, 0))
        return image

    def on_quit(self, icon, item):
        if hasattr(self, 'hook'):
            user32.UnhookWindowsHookEx(self.hook)
        icon.stop()
        self.root.quit()

    def on_setting(self, icon, item):
        self.root.after(0, self.open_settings)

    def on_reset_text(self, icon, item):
        def do_reset():
            self.text_sequence = []
            self.text_index = -1
            self.text_str = "..."
            self.update_style()
            print("[RESET] 🗑️ Đã xóa toàn bộ lịch sử text!")
        self.root.after(0, do_reset)

    def create_model_menu(self):
        def set_model(model_name):
            def inner(icon, item):
                self.ai_model = model_name
                self.save_config()
                print(f"\n[HOT SWAP] 🚀 AI Model changed to: {model_name}\n")
            return inner
            
        def is_checked(model_name):
            def inner(item):
                return getattr(self, 'ai_model', DEFAULT_AI_MODEL) == model_name
            return inner

        model_items = []
        for m in AVAILABLE_MODELS:
            model_items.append(pystray.MenuItem(m, set_model(m), checked=is_checked(m), radio=True))
            
        return pystray.Menu(*model_items)

    def setup_tray(self):
        image = self.create_tray_image()
        menu = pystray.Menu(
            pystray.MenuItem('⚡ Chuyển Đổi Model Nhanh', self.create_model_menu()),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('🗑️ Reset Text', self.on_reset_text),
            pystray.MenuItem('Cài Đặt', self.on_setting),
            pystray.MenuItem('Thoát', self.on_quit)
        )
        self.tray_icon = pystray.Icon("Overlay", image, "Ứng dụng AI Overlay", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def open_settings(self):
        if hasattr(self, 'settings_win') and self.settings_win.winfo_exists():
            self.settings_win.focus()
            return

        self.settings_win = tk.Toplevel(self.root)
        self.settings_win.title("Cài đặt Overlay")
        self.settings_win.geometry("420x420")
        self.settings_win.resizable(False, False)
        self.settings_win.attributes('-topmost', True)
        
        notebook = ttk.Notebook(self.settings_win)
        notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        tab_text = ttk.Frame(notebook)
        notebook.add(tab_text, text="Cài Đặt Chữ")
        
        tab_rect = ttk.Frame(notebook)
        notebook.add(tab_rect, text="Cài Đặt Khung")

        tab_ai = ttk.Frame(notebook)
        notebook.add(tab_ai, text="Cài Đặt AI API")

        # -- TAB TEXT --
        tk.Label(tab_text, text="Tọa độ X chữ:").grid(row=0, column=0, padx=15, pady=10, sticky='w')
        x_var = tk.StringVar(value=str(self.pos_x))
        tk.Entry(tab_text, textvariable=x_var).grid(row=0, column=1)

        tk.Label(tab_text, text="Tọa độ Y chữ:").grid(row=1, column=0, padx=15, pady=10, sticky='w')
        y_var = tk.StringVar(value=str(self.pos_y))
        tk.Entry(tab_text, textvariable=y_var).grid(row=1, column=1)

        tk.Label(tab_text, text="Kích cỡ chữ:").grid(row=2, column=0, padx=15, pady=10, sticky='w')
        size_var = tk.StringVar(value=str(self.text_size))
        tk.Entry(tab_text, textvariable=size_var).grid(row=2, column=1)

        tk.Label(tab_text, text="Màu chữ:").grid(row=3, column=0, padx=15, pady=10, sticky='w')
        color_btn_t = tk.Button(tab_text, text="[ Chọn màu ]", bg=self.text_color, width=15)
        color_btn_t.grid(row=3, column=1)

        def choose_color_text():
            c = colorchooser.askcolor(parent=self.settings_win, title="Chọn màu chữ", initialcolor=self.text_color)[1]
            if c:
                self.text_color = c
                color_btn_t.config(bg=c)
        color_btn_t.config(command=choose_color_text)

        border_var = tk.BooleanVar(value=self.show_border)
        tk.Checkbutton(tab_text, text="Bật viền chữ đen", variable=border_var).grid(row=4, column=0, columnspan=2, pady=5, sticky='w', padx=10)

        # -- TAB RECT --
        rect_en_var = tk.BooleanVar(value=self.rect_enabled)
        tk.Checkbutton(tab_rect, text="BẬT HIỂN THỊ KHUNG HÌNH CHỮ NHẬT", variable=rect_en_var, font=('', 9, 'bold')).grid(row=0, column=0, columnspan=2, pady=5, sticky='w', padx=10)

        tk.Label(tab_rect, text="Tọa độ X khung:").grid(row=1, column=0, padx=15, pady=5, sticky='w')
        rx_var = tk.StringVar(value=str(self.rect_x))
        tk.Entry(tab_rect, width=12, textvariable=rx_var).grid(row=1, column=1, sticky='w')

        tk.Label(tab_rect, text="Tọa độ Y khung:").grid(row=2, column=0, padx=15, pady=5, sticky='w')
        ry_var = tk.StringVar(value=str(self.rect_y))
        tk.Entry(tab_rect, width=12, textvariable=ry_var).grid(row=2, column=1, sticky='w')

        tk.Label(tab_rect, text="Chiều dài (W):").grid(row=3, column=0, padx=15, pady=5, sticky='w')
        rw_var = tk.StringVar(value=str(self.rect_w))
        tk.Entry(tab_rect, width=12, textvariable=rw_var).grid(row=3, column=1, sticky='w')

        tk.Label(tab_rect, text="Chiều cao (H):").grid(row=4, column=0, padx=15, pady=5, sticky='w')
        rh_var = tk.StringVar(value=str(self.rect_h))
        tk.Entry(tab_rect, width=12, textvariable=rh_var).grid(row=4, column=1, sticky='w')
        
        tk.Label(tab_rect, text="Độ mờ (Alpha 0-1):").grid(row=5, column=0, padx=15, pady=5, sticky='w')
        ralpha_var = tk.StringVar(value=str(self.rect_alpha))
        tk.Entry(tab_rect, width=12, textvariable=ralpha_var).grid(row=5, column=1, sticky='w')

        tk.Label(tab_rect, text="Màu khung:").grid(row=6, column=0, padx=15, pady=5, sticky='w')
        color_btn_r = tk.Button(tab_rect, text="[ Chọn màu ]", bg=self.rect_color, width=15)
        color_btn_r.grid(row=6, column=1, sticky='w')

        def choose_color_rect():
            c = colorchooser.askcolor(parent=self.settings_win, title="Chọn màu khung", initialcolor=self.rect_color)[1]
            if c:
                self.rect_color = c
                color_btn_r.config(bg=c)
        color_btn_r.config(command=choose_color_rect)

        # -- TAB AI --
        tk.Label(tab_ai, text="OpenRouter API Key 1:").grid(row=0, column=0, padx=15, pady=5, sticky='w')
        api_var1 = tk.StringVar(value=self.api_key1)
        tk.Entry(tab_ai, textvariable=api_var1, width=22, show='*').grid(row=0, column=1, sticky='w')
        
        tk.Label(tab_ai, text="OpenRouter API Key 2:").grid(row=1, column=0, padx=15, pady=5, sticky='w')
        api_var2 = tk.StringVar(value=self.api_key2)
        tk.Entry(tab_ai, textvariable=api_var2, width=22, show='*').grid(row=1, column=1, sticky='w')
        
        tk.Label(tab_ai, text="OpenRouter API Key 3:").grid(row=2, column=0, padx=15, pady=5, sticky='w')
        api_var3 = tk.StringVar(value=self.api_key3)
        tk.Entry(tab_ai, textvariable=api_var3, width=22, show='*').grid(row=2, column=1, sticky='w')
        
        tk.Label(tab_ai, text="AI Model (OpenRouter):").grid(row=3, column=0, padx=15, pady=10, sticky='w')
        model_var = tk.StringVar(value=self.ai_model)
        model_cb = ttk.Combobox(tab_ai, textvariable=model_var, values=AVAILABLE_MODELS, width=21)
        model_cb.grid(row=3, column=1, sticky='w')

        def apply_settings():
            try:
                self.ai_model = model_var.get().strip()
                self.pos_x = int(x_var.get())
                self.pos_y = int(y_var.get())
                self.text_size = int(size_var.get())
                self.show_border = border_var.get()
                
                self.rect_enabled = rect_en_var.get()
                self.rect_x = int(rx_var.get())
                self.rect_y = int(ry_var.get())
                self.rect_w = int(rw_var.get())
                self.rect_h = int(rh_var.get())
                
                self.api_key1 = api_var1.get().strip()
                self.api_key2 = api_var2.get().strip()
                self.api_key3 = api_var3.get().strip()
                
                alpha_val = float(ralpha_var.get())
                if alpha_val < 0.0: alpha_val = 0.0
                if alpha_val > 1.0: alpha_val = 1.0
                self.rect_alpha = alpha_val
                
                self.update_style() 
                self.save_config()
                messagebox.showinfo("Thành công", "Đã lưu cài đặt!", parent=self.settings_win)
            except Exception:
                messagebox.showerror("Lỗi", "Vui lòng kiểm tra lại số liệu nhập vào!", parent=self.settings_win)

        tk.Button(self.settings_win, text="LƯU CÀI ĐẶT", command=apply_settings, bg='green', fg='white', width=30, font=('', 10, 'bold')).pack(pady=10)

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = OverlayApp()
    app.run()
