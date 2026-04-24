import tkinter as tk
from tkinter import ttk, messagebox
import pyautogui
import pygetwindow as gw
import pandas as pd
import pyperclip
import uuid6
import time
import re
import os
import threading
import traceback
import winocr
from PIL import Image
import ctypes
import queue
from datetime import datetime

# 실행 환경 설정
import sys
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 인증 및 설정 파일 경로
CREDENTIALS_FILE = os.path.join(BASE_DIR, 'credentials.json')
TOKEN_FILE = os.path.join(BASE_DIR, 'token.json')
CALENDAR_ID_FILE = os.path.join(BASE_DIR, 'calendar_id.txt')

# DPI 인식 설정
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except: pass

# Windows Native OCR 사용 (winocr) — 한국어 인식률 최고

class KakaoMacroGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("카카오톡 일정 수집기 v1.0")
        self.root.geometry("650x600")
        
        self.running = False
        self.selected_window = None
        self.msg_queue = queue.Queue()
        self.processed_titles = set() # 제목 기준 중복 방지 추가

        # Windows Native OCR (winocr) — 별도 초기화 불필요
        # 한국어 언어팩이 Windows에 설치되어 있어야 합니다.

        # UI 구성 (동일)
        frame_top = tk.Frame(root, pady=10)
        frame_top.pack(fill="x")
        tk.Label(frame_top, text="대상 창:").pack(side="left", padx=10)
        self.combo_windows = ttk.Combobox(frame_top, width=40, state="readonly")
        self.combo_windows.pack(side="left", padx=5)
        tk.Button(frame_top, text="새로고침", command=self.refresh_windows).pack(side="left", padx=5)
        tk.Button(frame_top, text="창 고정", command=self.fix_and_resize, bg="#e1f5fe").pack(side="left", padx=5)

        frame_mid = tk.Frame(root, pady=10)
        frame_mid.pack(fill="x")
        tk.Label(frame_mid, text="저장 이름:").pack(side="left", padx=10)
        self.entry_name = tk.Entry(frame_mid, width=35)
        self.entry_name.pack(side="left", padx=5)

        # 수집 설정 (개수, 날짜 범위)
        frame_config = tk.Frame(root, pady=5)
        frame_config.pack(fill="x")
        
        tk.Label(frame_config, text="수집 개수:").pack(side="left", padx=10)
        self.entry_limit = tk.Entry(frame_config, width=5)
        self.entry_limit.pack(side="left", padx=5)
        
        tk.Label(frame_config, text="시작일(YYYYMMDD):").pack(side="left", padx=10)
        self.entry_start_date = tk.Entry(frame_config, width=12)
        self.entry_start_date.pack(side="left", padx=5)

        tk.Label(frame_config, text="종료일(YYYYMMDD):").pack(side="left", padx=10)
        self.entry_end_date = tk.Entry(frame_config, width=12)
        self.entry_end_date.pack(side="left", padx=5)

        frame_btns = tk.Frame(root, pady=10)
        frame_btns.pack(fill="x", padx=20)
        self.btn_run = tk.Button(frame_btns, text="매크로 실행 시작", command=self.start_thread, bg="#4caf50", fg="white", font=("Arial", 11, "bold"), height=2, width=25)
        self.btn_run.pack(side="left", expand=True, padx=5)
        self.btn_stop = tk.Button(frame_btns, text="실행 중지", command=self.stop_macro, bg="#f44336", fg="white", font=("Arial", 11, "bold"), height=2, width=25, state="disabled")
        self.btn_stop.pack(side="left", expand=True, padx=5)

        frame_test = tk.Frame(root, pady=5)
        frame_test.pack(fill="x", padx=20)
        self.btn_capture = tk.Button(frame_test, text="📸 화면 캡쳐 테스트", command=self.capture_test, bg="#2196f3", fg="white", font=("Arial", 10, "bold"), height=1, width=25)
        self.btn_capture.pack(side="left", expand=True, padx=5)
        self.btn_gcal = tk.Button(frame_test, text="📅 구글 캘린더 등록", command=self.upload_to_gcal, bg="#ff9800", fg="white", font=("Arial", 10, "bold"), height=1, width=25)
        self.btn_gcal.pack(side="left", expand=True, padx=5)

        # 로그 검색창 추가
        frame_search = tk.Frame(root, padx=15)
        frame_search.pack(fill="x")
        tk.Label(frame_search, text="로그 검색:").pack(side="left")
        self.entry_search = tk.Entry(frame_search, width=30)
        self.entry_search.pack(side="left", padx=5)
        self.entry_search.bind("<Return>", lambda e: self.search_log())
        tk.Button(frame_search, text="검색", command=self.search_log).pack(side="left", padx=5)
        tk.Button(frame_search, text="지우기", command=self.clear_search).pack(side="left")

        # 로그창 + 스크롤바
        frame_log = tk.Frame(root, padx=15, pady=10)
        frame_log.pack(fill="both", expand=True)
        
        self.log_text = tk.Text(frame_log, height=15, state="disabled", bg="#f5f5f5", wrap="none")
        scrollbar_y = tk.Scrollbar(frame_log, orient="vertical", command=self.log_text.yview)
        scrollbar_x = tk.Scrollbar(frame_log, orient="horizontal", command=self.log_text.xview)
        self.log_text.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        self.log_text.pack(side="left", fill="both", expand=True)
        
        # 검색 태그 설정
        self.log_text.tag_configure("match", background="yellow", foreground="black")
        
        self.root.after(100, self.process_queue)
        self.refresh_windows()

    def search_log(self):
        query = self.entry_search.get().strip()
        self.log_text.tag_remove("match", "1.0", tk.END)
        if not query: return
        
        start_pos = "1.0"
        while True:
            start_pos = self.log_text.search(query, start_pos, stopindex=tk.END, nocase=True)
            if not start_pos: break
            end_pos = f"{start_pos}+{len(query)}c"
            self.log_text.tag_add("match", start_pos, end_pos)
            start_pos = end_pos
        
        # 첫 번째 검색 결과로 스크롤
        first_match = self.log_text.tag_ranges("match")
        if first_match:
            self.log_text.see(first_match[0])

    def clear_search(self):
        self.entry_search.delete(0, tk.END)
        self.log_text.tag_remove("match", "1.0", tk.END)

    def log(self, message):
        self.msg_queue.put(("LOG", message))

    def process_queue(self):
        try:
            while True:
                task, val = self.msg_queue.get_nowait()
                if task == "LOG":
                    self.log_text.config(state="normal")
                    self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {val}\n")
                    self.log_text.see(tk.END)
                    self.log_text.config(state="disabled")
                elif task == "STATE":
                    self.btn_run.config(state=val[0])
                    self.btn_stop.config(state=val[1])
                elif task == "CAPTURE_DONE":
                    self.btn_capture.config(state="normal", text="📸 화면 캡쳐 테스트")
                elif task == "GCAL_DONE":
                    self.btn_gcal.config(state="normal", text="📅 구글 캘린더 등록")
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.process_queue)

    def refresh_windows(self):
        titles = [w.title for w in gw.getAllWindows() if w.title.strip()]
        self.combo_windows['values'] = titles

    def fix_and_resize(self):
        title = self.combo_windows.get()
        if not title: return
        try:
            self.selected_window = gw.getWindowsWithTitle(title)[0]
            self.selected_window.activate()
            self.selected_window.resizeTo(800, 1000)
            self.selected_window.moveTo(0, 0)
            clean_name = re.sub(r'[\/:*?"<>|]', '', title).strip()
            self.entry_name.delete(0, tk.END)
            self.entry_name.insert(0, clean_name)
            self.log(f"창 고정 완료: {title} (800x1000)")
        except Exception as e:
            messagebox.showerror("오류", f"창 조절 실패: {e}")

    def capture_test(self):
        """현재 화면 캡쳐 후 OCR 테스트 (별도 스레드)"""
        if not self.selected_window:
            messagebox.showwarning("경고", "먼저 대상 창을 선택하고 '창 고정'을 눌러주세요.")
            return
        self.btn_capture.config(state="disabled", text="📸 캡쳐 중...")
        threading.Thread(target=self._do_capture_test, daemon=True).start()

    def _do_capture_test(self):
        try:
            self.log("=" * 50)
            self.log("📸 화면 캡쳐 테스트 시작")
            self.log("=" * 50)

            # 매크로 실행 없이 현재 화면만 OCR
            self.running = True  # run_ocr_and_parse 내부의 check_running 우회
            info = self.run_ocr_and_parse()
            self.running = False

            if info is None:
                self.log("⚠️ OCR 결과 없음 (상세 페이지가 아니거나 로딩 중)")
            else:
                self.log("-" * 40)
                self.log(f"  📌 제목:   {info['title']}")
                self.log(f"  👤 작성자: {info['author']}")
                self.log(f"  📅 시작:   {info['start_date']} {info['start_time']}")
                self.log(f"  🏁 종료:   {info['end_date']} {info['end_time']}")
                self.log(f"  🔑 UUID:   {info['id']}")
                self.log("-" * 40)
                self.log("✅ 캡쳐 테스트 완료!")

        except Exception as e:
            self.log(f"❌ 캡쳐 테스트 오류: {e}")
            import traceback
            self.log(traceback.format_exc())
        finally:
            self.running = False
            self.msg_queue.put(("CAPTURE_DONE", None))

    def upload_to_gcal(self):
        """CSV 파일을 선택하여 Google Calendar에 등록"""
        from tkinter import filedialog, simpledialog
        csv_path = filedialog.askopenfilename(
            title="등록할 CSV 파일 선택",
            filetypes=[("CSV 파일", "*.csv"), ("모든 파일", "*.*")],
            initialdir=BASE_DIR
        )
        if not csv_path:
            return
            
        # 이전에 사용한 캘린더 ID 불러오기
        default_cal_id = "primary"
        if os.path.exists(CALENDAR_ID_FILE):
            try:
                with open(CALENDAR_ID_FILE, 'r', encoding='utf-8') as f:
                    default_cal_id = f.read().strip() or "primary"
            except: pass

        # 그룹 캘린더 등 특정 캘린더 ID 입력받기
        cal_id_input = simpledialog.askstring(
            "캘린더 ID 입력",
            "등록할 구글 캘린더 ID를 입력하세요.\n(기본 내 캘린더에 등록하려면 그냥 '확인'을 누르세요.)",
            initialvalue=default_cal_id
        )
        if cal_id_input is None:
            return # 취소
        calendar_id = cal_id_input.strip() or "primary"

        # 사용한 캘린더 ID 저장
        try:
            with open(CALENDAR_ID_FILE, 'w', encoding='utf-8') as f:
                f.write(calendar_id)
        except: pass

        self.btn_gcal.config(state="disabled", text="📅 등록 중...")
        threading.Thread(target=self._do_gcal_upload, args=(csv_path, calendar_id), daemon=True).start()

    def _do_gcal_upload(self, csv_path, calendar_id):
        try:
            from google_calendar_upload import upload_csv_to_calendar
            self.log("=" * 50)
            self.log(f"📅 Google Calendar 등록 시작 (대상: {calendar_id})")
            self.log("=" * 50)
            success, failed, skipped = upload_csv_to_calendar(csv_path, calendar_id=calendar_id, log_func=self.log)
            self.log("=" * 50)
        except BaseException as e:
            self.log(f"❌ Google Calendar 등록 오류: {e}")
            try:
                import traceback
                self.log(traceback.format_exc())
            except:
                pass
        self.msg_queue.put(("GCAL_DONE", None))

    def start_thread(self):
        if not self.selected_window: return
        self.running = True
        self.processed_titles.clear() # 시작 시 초기화
        self.msg_queue.put(("STATE", ("disabled", "normal")))
        threading.Thread(target=self.run_macro, daemon=True).start()

    def stop_macro(self):
        self.running = False
        self.log("🛑 중지 예약됨.")

    def check_running(self):
        if not self.running: raise StopIteration

    def interruptible_sleep(self, seconds):
        """중단 가능한 sleep: 100ms 단위로 running 플래그 확인"""
        end_time = time.time() + seconds
        while time.time() < end_time:
            if not self.running:
                raise StopIteration
            time.sleep(0.1)

    def run_macro(self):
        save_name = self.entry_name.get().strip() or "kakao_events"
        limit = self.entry_limit.get().strip()
        start_date = self.entry_start_date.get().strip()
        end_date = self.entry_end_date.get().strip()
        
        limit_int = int(limit) if limit.isdigit() else 0
        
        filename = f"{save_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        filepath = os.path.join(BASE_DIR, filename)
        try:
            results = self.execute_logic(save_name, limit_int, start_date, end_date)
            if results:
                pd.DataFrame(results).to_csv(filepath, index=False, encoding='utf-8-sig')
                self.log(f"✅ 저장 완료: {filepath}")
        except StopIteration:
            self.log("🛑 작업을 중단했습니다.")
        except Exception as e:
            self.log(f"❌ 오류 발생: {e}")
        finally:
            self.running = False
            self.msg_queue.put(("STATE", ("normal", "disabled")))

    # --- 뒤로가기 클릭 (좌표 기반 절대 실패 없도록 강화) ---
    def click_back_button_with_retry(self):
        win = self.selected_window
        # 800x1000 기준 뒤로가기 "← Details" 버튼
        # 타이틀바(약 30px) 아래에 위치. 다양한 오프셋을 두어 클릭 실패율을 최소화합니다.
        back_positions = [
            (win.left + 40, win.top + 72),   # 중심
            (win.left + 35, win.top + 65),   # 약간 위
            (win.left + 40, win.top + 80),   # 약간 아래
            (win.left + 25, win.top + 72),   # 약간 왼쪽
            (win.left + 55, win.top + 72),   # 약간 오른쪽
        ]
        for bx, by in back_positions:
            self.check_running()
            pyautogui.click(bx, by)
            time.sleep(0.5)  # 화면 전환 대기 시간 증가
            # 뒤로가기가 성공했는지 확인: 상세보기 텍스트가 사라졌는지 체크
            try:
                check_shot = pyautogui.screenshot(region=(win.left, win.top + 50, 200, 50))
                result = winocr.recognize_pil_sync(check_shot, 'ko')
                check_text = result.get('text', '')
                if "상세보기" not in check_text and "details" not in check_text.lower():
                    return  # 성공적으로 뒤로감
            except:
                pass
        # 마지막 수단: 보조 좌표 재시도
        self.log("  [경고] 뒤로가기 버튼 클릭 실패")

    def refresh_event_tab(self):
        """일정 탭 무한로딩 시 공지 → 일정 탭 전환으로 복구"""
        win = self.selected_window
        self.log("🔄 일정 탭 로딩 실패 → 공지/일정 탭 전환 시도")
        # 공지 탭 클릭
        pyautogui.click(win.left + 45, win.top + 107)
        self.interruptible_sleep(2.0)
        # 일정 탭 클릭
        pyautogui.click(win.left + 110, win.top + 107)
        self.log("⏳ 일정 목록 로딩 대기 중 (5초)...")
        self.interruptible_sleep(5.0)

    def execute_logic(self, save_name="kakao_events", limit=0, start_date_str="", end_date_str=""):
        self.selected_window.activate()
        scraped_data = []
        event_index = 0
        consecutive_failures = 0
        consecutive_skips = 0  # 연속 건너뜀 횟수
        collected_count = 0
        win = self.selected_window

        # 날짜 객체 변환
        start_date_obj = None
        end_date_obj = None
        
        def parse_date(d_str):
            if not d_str: return None
            clean_str = d_str.replace('-', '').replace('.', '').replace('/', '').strip()
            return datetime.strptime(clean_str, "%Y%m%d")

        try:
            start_date_obj = parse_date(start_date_str)
            end_date_obj = parse_date(end_date_str)
        except:
            self.log("⚠️ 날짜 형식이 잘못되었습니다. 날짜 제한 없이 진행합니다.")
            start_date_obj = None
            end_date_obj = None

        self.log("🚀 키보드 탐색 모드로 시작 (Home → Down → Enter)")

        while self.running:
            self.check_running()

            # 1. Home → Down × (event_index) → Enter 로 이벤트 진입
            pyautogui.click(win.left + 400, win.top + 500)  # 창 포커스
            time.sleep(0.2)
            pyautogui.press('home')
            time.sleep(0.3)
            for _ in range(event_index):
                pyautogui.press('down')
                time.sleep(0.1)
            pyautogui.press('enter')
            # 고정 대기 제거 (process_detail_page 내부의 스마트 대기 로직 활용)

            # 2. 상세 페이지 확인
            info = self.process_detail_page()

            if info == "RETRY" or info is None:
                consecutive_failures += 1
                self.log(f"  [실패] 이벤트 #{event_index} 진입 실패 ({consecutive_failures}회 연속)")
                # 1단계: 뒤로가기 (상세 페이지에서 목록으로 복귀)
                self.click_back_button_with_retry()
                # 2단계: Notices→Events 탭 전환 (목록 로딩 복구)
                self.refresh_event_tab()
                if consecutive_failures >= 30:
                    self.log("⚠️ 30회 연속 로딩 실패. 수집을 강제 종료합니다.")
                    break
                # event_index 유지 → 같은 이벤트 재시도
                continue

            consecutive_failures = 0

            # 3. 정보 처리
            # 날짜 필터링
            try:
                # info['start_date'] 형식: "2026/02/09 월요일"
                cur_date_str = info['start_date'].split(' ')[0]
                cur_date_obj = datetime.strptime(cur_date_str, "%Y/%m/%d")
                
                if start_date_obj and cur_date_obj < start_date_obj:
                    self.log(f"  [건너뜀] 시작일({start_date_str}) 이전 일정: {info['title']}")
                    self.click_back_button_with_retry()
                    self.log("⏳ 일정 목록 로딩 대기 중 (5초)...")
                    self.interruptible_sleep(5.0)
                    event_index += 1
                    continue
                
                if end_date_obj and cur_date_obj > end_date_obj:
                    self.log(f"  [수집 종료] 종료일({end_date_str}) 이후 일정 도달: {info['title']}")
                    break
            except Exception as e:
                pass # 날짜 파싱 실패 시 그냥 진행

            unique_key = f"{info['title']}_{info['start_date']}"
            if unique_key not in self.processed_titles:
                # ID가 없는 경우에만 새 UUID 등록 (대시 제거)
                if info['id'].startswith('ocr_'):
                    new_uuid = str(uuid6.uuid6()).replace('-', '')
                    self.post_uuid_comment(new_uuid)
                    info['id'] = new_uuid
                
                scraped_data.append(info)
                self.processed_titles.add(unique_key)
                collected_count += 1
                self.log(f"✅ [수집 #{collected_count}] {info['title']}")
                consecutive_skips = 0
                
                # 개수 제한 확인
                if limit > 0 and collected_count >= limit:
                    self.log(f"🏁 설정한 수집 개수({limit}개)를 달성했습니다.")
                    break
            else:
                consecutive_skips += 1
                self.log(f"  [중복] 이미 처리된 일정: {info['title']} ({consecutive_skips}회 연속)")
                if consecutive_skips >= 3:
                    self.log("더 이상 새로운 항목이 없습니다. 수집을 종료합니다.")
                    break

            # 4. 뒤로가기 → 다음 이벤트
            self.click_back_button_with_retry()
            self.log("⏳ 일정 목록 로딩 대기 중 (5초)...")
            self.interruptible_sleep(5.0)
            event_index += 1

        return scraped_data

    def process_detail_page(self):
        try:
            self.check_running()
            start_wait = time.time()
            info = None
            title_noise = ["minute", "hour", "before", "betore", "thour", "ago", "reminder",
                           "종료", "진행", "습니다"]

            null_count = 0
            while time.time() - start_wait < 20: # 최대 20초 대기
                self.check_running()
                info = self.run_ocr_and_parse()
                if info == "LIST_PAGE":
                    # 목록 페이지에 머물러 있음 → 상세 진입 실패로 간주하고 즉시 반환
                    return "RETRY"
                if info is None:
                    null_count += 1
                    if null_count >= 10: # 10초 이상 로딩 안 되면 포기
                        # 10회 연속 None → 로딩 중 또는 목록 페이지, 즉시 탈출
                        return "RETRY"
                    self.interruptible_sleep(1.0)
                    continue
                null_count = 0
                self.log(f"  [디버그] 추출된 제목: {info['title']}")
                if info['title'] != "제목 미검출" and not any(n in info['title'].lower() for n in title_noise):
                    break
                self.interruptible_sleep(1.0)

            if not info or info['title'] == "제목 미검출":
                return "RETRY"

            # --- UUID 추가 확인 (댓글이 많아 밀린 경우 대비) ---
            if info['id'].startswith('ocr_'):
                self.log("  [UUID 탐색] 첫 화면에 ID가 없어 하단으로 스크롤하여 재확인합니다.")
                pyautogui.press('end')
                self.interruptible_sleep(1.2)
                # 하단 스크롤 후 다시 OCR (ID만 업데이트)
                info_bottom = self.run_ocr_and_parse()
                if info_bottom and not info_bottom['id'].startswith('ocr_'):
                    info['id'] = info_bottom['id']
                    self.log(f"  [UUID 발견] 하단 스크롤 후 ID 확인 완료: {info['id']}")

            return info
        except Exception as e:
            self.log(f"  [오류] process_detail_page: {e}")
            self.log(traceback.format_exc())
            return None

    def post_uuid_comment(self, new_uuid):
        """새 UUID를 댓글로 등록"""
        win = self.selected_window
        pyautogui.press('end')
        time.sleep(0.3)
        pyautogui.click(win.left + 150, win.top + 955)
        time.sleep(0.2)
        pyperclip.copy(f"id: {new_uuid}")
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.3)
        pyautogui.click(win.left + 740, win.top + 955) # 엔터 안 먹힐 대비 클릭 (둘 다 해도 무방)
        self.interruptible_sleep(2.0) # 댓글 달리는 시간 확실히 대기

    def _merge_korean_ascii(self, winocr_word, tess_word):
        """winocr 한글 + Tesseract 숫자/기호 병합.
        예: winocr='테스틔', tess='테스트1' → '테스트1'
        예: winocr='[젬민이병신1', tess='[PA WO SAl' → '[젬민이병신]'
        """
        # Tesseract에 숫자가 있고 winocr에 없는 경우:
        # winocr에서 한글 부분, Tesseract에서 뒷부분(숫자) 가져오기
        wk_chars = list(winocr_word)
        ts_chars = list(tess_word)

        # 한글 부분 끝 위치 찾기 (winocr)
        wk_last_korean = -1
        for i, c in enumerate(wk_chars):
            if '\uac00' <= c <= '\ud7a3':
                wk_last_korean = i

        # Tesseract에서 한글 끝난 이후 부분 찾기
        ts_last_korean = -1
        for i, c in enumerate(ts_chars):
            if '\uac00' <= c <= '\ud7a3':
                ts_last_korean = i

        if ts_last_korean >= 0 and ts_last_korean < len(ts_chars) - 1:
            # Tesseract 한글 뒤의 숫자/기호 부분
            ts_suffix = tess_word[ts_last_korean + 1:]
            # winocr의 한글 앞부분 (한글 직전까지의 prefix + 한글들)
            # Tesseract의 한글이 더 정확할 수도 있으므로 Tesseract 한글 부분 사용
            ts_prefix = tess_word[:ts_last_korean + 1]

            # winocr 한글이 더 많으면 winocr 한글 사용
            wk_korean_count = sum(1 for c in winocr_word if '\uac00' <= c <= '\ud7a3')
            ts_korean_count = sum(1 for c in ts_prefix if '\uac00' <= c <= '\ud7a3')

            if wk_korean_count >= ts_korean_count:
                # winocr 한글 + Tesseract suffix(숫자)
                # winocr에서 마지막 한글까지 잘라내기
                wk_prefix = winocr_word[:wk_last_korean + 1]
                return wk_prefix + ts_suffix
            else:
                return tess_word
        else:
            # Tesseract에 한글 뒤 숫자가 없으면 winocr 그대로
            return winocr_word

    def run_ocr_and_parse(self):
        win = self.selected_window

        # --- 단일 캡처 (800x1000 창 전체가 한 화면에 보임) ---
        screenshot = pyautogui.screenshot(region=(win.left, win.top + 30, 800, 940))

        # Windows Native OCR 로 텍스트 추출 (한국어+영어 이중 인식)
        result_ko = winocr.recognize_pil_sync(screenshot, 'ko')
        result_en = winocr.recognize_pil_sync(screenshot, 'en')

        # 두 결과를 합침: 한국어 결과를 기본으로 하되 영어 결과로 보완
        # winocr 결과를 기존 코드 호환 형태로 변환: [(bbox, text, conf), ...]
        results_detail = []
        results_text_only = []

        # 한국어 결과 파싱
        for line in result_ko.get('lines', []):
            for word in line.get('words', []):
                text = word.get('text', '').strip()
                rect = word.get('bounding_rect', {})
                if text:
                    # bbox를 EasyOCR 호환 형태로 변환 [[x1,y1],[x2,y1],[x2,y2],[x1,y2]]
                    x = rect.get('x', 0)
                    y = rect.get('y', 0)
                    w = rect.get('width', 0)
                    h = rect.get('height', 0)
                    bbox = [[x, y], [x+w, y], [x+w, y+h], [x, y+h]]
                    results_detail.append((bbox, text, 1.0))
                    results_text_only.append(text)

        # 영어 결과에서 UUID/날짜/시간 등 라틴 문자 보강
        en_texts = []
        en_results_detail = []
        for line in result_en.get('lines', []):
            for word in line.get('words', []):
                text = word.get('text', '').strip()
                rect = word.get('bounding_rect', {})
                if text:
                    en_texts.append(text)
                    x = rect.get('x', 0)
                    y = rect.get('y', 0)
                    w = rect.get('width', 0)
                    h = rect.get('height', 0)
                    bbox = [[x, y], [x+w, y], [x+w, y+h], [x, y+h]]
                    en_results_detail.append((bbox, text, 1.0))

        full_text_all = " ".join(results_text_only)
        full_text_en = " ".join(en_texts)

        self.log(f"  [OCR 한국어] {full_text_all[:200]}")
        self.log(f"  [OCR 영어] {full_text_en[:200]}")

        if len(results_text_only) < 3:
            return None

        # 전처리: OCR이 ':'를 '•'이나 '·'로 읽는 문제 보정
        full_text_all = full_text_all.replace('•', ':').replace('·', ':').replace('．', '.')
        for i, t in enumerate(results_text_only):
            results_text_only[i] = t.replace('•', ':').replace('·', ':').replace('．', '.')
        # results_detail의 텍스트도 보정
        results_detail = [(bbox, text.replace('•', ':').replace('·', ':'), conf) for bbox, text, conf in results_detail]

        # 영어 인식 결과도 함께 사용 (UUID, 날짜 등은 영어 결과가 더 정확)
        clean_text = re.sub(r'\s+', ' ', full_text_all)
        clean_text_en = re.sub(r'\s+', ' ', full_text_en)

        # --- 목록 페이지 감지 (상세 페이지가 아닌 경우 즉시 반환) ---
        list_page_markers_en = ["notices", "events", "polls", "quiz"]
        list_page_markers_ko = ["공지", "일정", "투표", "퀴즈"]
        lower_text = clean_text.lower()
        en_count = sum(1 for m in list_page_markers_en if m in lower_text)
        ko_count = sum(1 for m in list_page_markers_ko if m in clean_text)
        if en_count >= 3 or ko_count >= 3:
            # self.log("  [감지] 목록 페이지 (상세 페이지 아님)") # 로그 너무 많아질 수 있어 주석처리
            return "LIST_PAGE"

        # --- 상세 페이지 로딩 중 감지 ---
        if ("detail" in lower_text or "상세보기" in clean_text) and len(results_text_only) < 8:
            self.log("  [감지] 상세 페이지 로딩 중 (콘텐츠 없음) → 즉시 반환")
            return None

        # --- 1. UUID 추출 (영어 OCR 결과 우선 사용) ---
        item_id = None
        
        # UUID 추출 전처리 강화: 'id:' 뒤의 텍스트를 우선 추출하고 typo 보정
        for search_text in [clean_text_en, clean_text]:
            m_prefix = re.search(r'id[:\s]+([a-zA-Z0-9]+)', search_text, re.IGNORECASE)
            if m_prefix:
                raw_id = m_prefix.group(1)
                # 흔한 OCR 오타 보정: I/i/l -> 1, O/o -> 0
                raw_id = raw_id.replace('I', '1').replace('i', '1').replace('l', '1').replace('o', '0').replace('O', '0')
                if len(raw_id) >= 25:
                    item_id = raw_id.lower()
                    self.log(f"  [UUID 감지] id: 접두어 기반 발견: {item_id}")
                    break
        
        if not item_id:
            # 기존 정규식 방식 폴백
            for search_text in [clean_text_en, clean_text]:
                clean_text_for_id = search_text.replace('-', '').replace(' ', '')
                uuid_pattern = re.compile(r'([A-Fa-f0-9]{25,32})')
                u_match = uuid_pattern.search(clean_text_for_id)
                if u_match:
                    item_id = u_match.group(1).lower()
                    self.log(f"  [UUID 감지] 정규식 기반 발견: {item_id}")
                    break
        
        if not item_id:
            # 폴백: 더 관대한 패턴으로 재시도
            for search_text in [clean_text_en, clean_text]:
                clean_text_for_id = search_text.replace('-', '').replace(' ', '')
                uuid_pattern_loose = re.compile(r'([A-Za-z0-9]{25,32})')
                u_match = uuid_pattern_loose.search(clean_text_for_id)
                if u_match:
                    raw_id = u_match.group(1)
                    raw_id = raw_id.replace('O', '0').replace('o', '0').replace('I', '1').replace('l', '1')
                    item_id = raw_id.lower()
                    self.log(f"  [UUID 감지] 느슨한 정규식 발견: {item_id}")
                    break
        
        if not item_id:
            item_id = f"ocr_{hash(clean_text[:30])}"

        # --- 2. 날짜 추출 (한국어 + 영어 둘 다 지원) ---
        date_str = "날짜 미상"
        d_match = None
        d_match_ko = None

        # 한국어 날짜 형식: "2026년 4월 8일 (수)" 또는 "2026년 4월 8일"
        d_match_ko = re.search(
            r'(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일(?:\s*\(([가-힣])\))?',
            clean_text
        )
        if d_match_ko:
            year_num = int(d_match_ko.group(1))
            month_num = int(d_match_ko.group(2))
            day_num = int(d_match_ko.group(3))
            day_name = d_match_ko.group(4) or ""
            day_name_kr = {"월": "월요일", "화": "화요일", "수": "수요일",
                           "목": "목요일", "금": "금요일", "토": "토요일", "일": "일요일"}.get(day_name, "")
            date_str = f"{year_num}/{month_num:02d}/{day_num:02d} {day_name_kr}"
        else:
            # 영어 날짜 형식 폴백: "Wed, Apr 8, 2026"
            for search_text in [clean_text_en, clean_text]:
                d_match = re.search(
                    r'([A-Za-z]{3,4}[,.]?\s*(?:[A-Za-z]{3,4}\s*)?\d{1,2}[,.]?\s*\d{4})',
                    search_text
                )
                if d_match:
                    date_str = d_match.group(1).strip()
                    break

        # --- 3. 시간 추출 (한국어 + 영어 둘 다 지원) 및 종료 날짜 추출 ---
        start_date = date_str
        end_date = date_str
        start_time = "하루 종일"
        end_time = "하루 종일"

        # 전처리: OCR이 콜론을 누락하는 경우 보정 (오후 730 → 오후 7:30, 오후 1030 → 오후 10:30)
        clean_text_time = re.sub(r'(오[전후]\s*)(\d{1,2})(\d{2})', r'\1\2:\3', clean_text)

        # 한국어 다중일 형식: "오전 12:00 ~ 2026년 2월 7일 (토) 오전 6:00"
        t_match_multi = re.search(
            r'(오[전후]\s*\d{1,2}:?\d{2})\s*[~\-–]\s*'
            r'(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일(?:\s*\(([가-힣])\))?\s*'
            r'(오[전후]\s*\d{1,2}:?\d{2})',
            clean_text_time
        )
        if t_match_multi:
            start_time = t_match_multi.group(1).strip()
            
            end_y, end_m, end_d = int(t_match_multi.group(2)), int(t_match_multi.group(3)), int(t_match_multi.group(4))
            end_day_kr = {"월": "월요일", "화": "화요일", "수": "수요일", "목": "목요일", "금": "금요일", "토": "토요일", "일": "일요일"}.get(t_match_multi.group(5), "")
            end_date = f"{end_y}/{end_m:02d}/{end_d:02d} {end_day_kr}".strip()
            
            end_time = t_match_multi.group(6).strip()
        else:
            # 한국어 단일일 형식: "오전 1:00 ~ 오전 2:00"
            t_match_ko = re.search(
                r'(오[전후]\s*\d{1,2}:?\d{2})\s*[~\-–]\s*(오[전후]\s*\d{1,2}:?\d{2})',
                clean_text_time
            )
            if t_match_ko:
                start_time = t_match_ko.group(1).strip()
                end_time = t_match_ko.group(2).strip()
            else:
                # 영어 시간 형식 폴백: "1:00 AM ~ 2:00 AM" (단일일 가정)
                for search_text in [clean_text_en, clean_text]:
                    t_match = re.search(
                        r'(\d{1,2}[:.]?\d{2}\s*[APap][Mm])\s*[~\-–]\s*(\d{1,2}[:.]?\d{2}\s*[APap][Mm])',
                        search_text
                    )
                    if t_match:
                        raw_s = t_match.group(1).strip()
                        raw_s = re.sub(r'(\d{1,2})[:.]?(\d{2})\s*([APap][Mm])', r'\1:\2 \3', raw_s)
                        start_time = raw_s
                        
                        raw_e = t_match.group(2).strip()
                        raw_e = re.sub(r'(\d{1,2})[:.]?(\d{2})\s*([APap][Mm])', r'\1:\2 \3', raw_e)
                        end_time = raw_e
                        break

        # --- 4. 제목/작성자 추출 (위치 기반) ---
        title = "제목 미검출"
        author = "작성자 미상"

        # 한국어 OCR 결과 (제목/작성자 추출용)
        items_with_y = []
        for bbox, text, conf in results_detail:
            y_center = (bbox[0][1] + bbox[2][1]) / 2
            x_center = (bbox[0][0] + bbox[1][0]) / 2
            items_with_y.append((y_center, x_center, text, conf))
        items_with_y.sort(key=lambda x: x[0])

        # 영어 OCR 결과 (status/date 앵커 검색용)
        en_items_with_y = []
        for bbox, text, conf in en_results_detail:
            y_center = (bbox[0][1] + bbox[2][1]) / 2
            x_center = (bbox[0][0] + bbox[1][0]) / 2
            en_items_with_y.append((y_center, x_center, text, conf))
        en_items_with_y.sort(key=lambda x: x[0])

        # --- 작성자: 연속된 한글 1글자를 합쳐서 이름 생성 ---
        korean_char_pattern = re.compile(r'^[가-힣]$')
        korean_name_pattern = re.compile(r'^[가-힣]{2,4}$')
        # UI 키워드 (작성자로 오인식 방지)
        author_skip_words = {"상세보기", "수", "일정", "공지", "투표", "퀴즈",
                             "종료", "진행", "알림", "오전", "오후"}

        # 방법1: 인접한 한글 1글자 합치기 (김 찬 진 → 김찬진)
        i = 0
        while i < len(items_with_y):
            y, x, text, conf = items_with_y[i]
            t_strip = text.strip().replace(" ", "")
            if t_strip in author_skip_words:
                i += 1
                continue
            if korean_char_pattern.match(t_strip):
                # 같은 y좌표(±15px) 연속 한글 1글자 수집
                name_chars_with_x = [(x, t_strip)]
                j = i + 1
                while j < len(items_with_y):
                    y2, x2, t2, c2 = items_with_y[j]
                    if abs(y2 - y) < 15:
                        t2_strip = t2.strip().replace(" ", "")
                        if korean_char_pattern.match(t2_strip):
                            name_chars_with_x.append((x2, t2_strip))
                        j += 1
                    else:
                        break
                name_chars_with_x.sort(key=lambda item: item[0])
                combined = "".join(item[1] for item in name_chars_with_x)
                if 2 <= len(combined) <= 4 and combined not in author_skip_words:
                    author = combined
                    break
            elif korean_name_pattern.match(t_strip) and t_strip not in author_skip_words:
                author = t_strip
                break
            i += 1

        # --- 날짜 변환 (영어 형식인 경우만): "Wed, Apr 8, 2026" → "2026/04/08 수요일" ---
        day_map = {
            'mon': '월요일', 'tue': '화요일', 'wed': '수요일',
            'thu': '목요일', 'fri': '금요일', 'sat': '토요일', 'sun': '일요일'
        }
        month_map = {
            'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
            'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
        }
        if d_match and not d_match_ko:
            raw_date = d_match.group(1)
            parts = re.findall(r'[A-Za-z]+|\d+', raw_date)
            day_name_kr = ""
            month_num = 0
            day_num = 0
            year_num = 0
            for p in parts:
                p_lower = p.lower()[:3]
                if p_lower in day_map:
                    day_name_kr = day_map[p_lower]
                elif p_lower in month_map:
                    month_num = month_map[p_lower]
                elif p.isdigit():
                    n = int(p)
                    if n > 31:
                        year_num = n
                    elif day_num == 0:
                        day_num = n
            if month_num == 0 and year_num and day_num:
                month_num = 4
            if year_num and day_num:
                date_str = f"{year_num}/{month_num:02d}/{day_num:02d} {day_name_kr}"

        # --- 상태 라인 y좌표 (종료/진행/알림 등) ---
        status_y = None
        date_y = None

        # 한국어 상태 키워드
        status_keywords_ko = ["종료", "진행", "알림", "남았습니다", "일정이"]
        status_keywords_en = ["ended", "progress", "reminder"]

        # 1차: 한국어 OCR에서 상태/날짜 앵커 검색
        for y, x, text, conf in items_with_y:
            t_text = text.strip()
            t_lower = t_text.lower()
            if status_y is None:
                if any(kw in t_text for kw in status_keywords_ko) or any(kw in t_lower for kw in status_keywords_en):
                    status_y = y
            if date_y is None:
                # 한국어 날짜: "2026년" 또는 "4월" 등
                if d_match_ko and re.search(r'\d{4}년', t_text):
                    date_y = y
                elif d_match:
                    d_str = d_match.group(0)
                    if len(t_text) >= 3 and t_text in d_str:
                        date_y = y

        # 2차 폴백: 영어 OCR에서도 검색
        if status_y is None:
            for y, x, text, conf in en_items_with_y:
                t_lower = text.lower()
                if any(kw in t_lower for kw in status_keywords_en):
                    status_y = y
                    break
        if date_y is None and d_match:
            for y, x, text, conf in en_items_with_y:
                d_str = d_match.group(0)
                t_strip = text.strip()
                if len(t_strip) >= 3 and t_strip in d_str:
                    date_y = y
                    break

        self.log(f"  [디버그] status_y: {status_y}, date_y: {date_y}")

        # --- 제목 추출: 같은 줄의 모든 단어를 합침 ---
        title_noise_words = {
            # 영어
            "day", "days", "day(s)", "until", "event", "reminder", "reminder:",
            "read", "the", "has", "this", "is", "in", "be", "to",
            "attending", "not", "like", "first", "enter", "comment",
            "before", "ago", "minute", "hour", "minutes", "hours",
            # 한국어
            "일정이", "종료되었습니다.", "진행", "중입니다.", "알림:",
            "남았습니다.", "가장", "먼저", "좋아요를", "남겨보세요.",
            "댓글을", "등록",
        }
        def is_title_noise(text):
            t = text.strip()
            t_lower = t.lower()
            if t_lower in title_noise_words or t in title_noise_words:
                return True
            # 한국어 노이즈 패턴 (부분 매칭)
            if any(kw in t for kw in ["종료", "진행", "남았습니다", "좋아요", "남겨보세요", "댓글"]):
                return True
            return False

        def collect_line_text(target_y, items):
            """같은 y좌표(±15px) 줄의 모든 텍스트를 x좌표순으로 합침"""
            line_items = []
            for y, text, conf in items:
                if abs(y - target_y) < 15:
                    # x좌표도 필요 → results_detail에서 가져옴
                    line_items.append(text.strip())
            return " ".join(line_items)

        # 제목 줄 찾기: status_y와 date_y 사이
        title_y = None
        if status_y is not None and date_y is not None:
            for y, x, text, conf in items_with_y:
                if y > status_y + 8 and y < date_y - 8:
                    t = text.strip()
                    if not is_title_noise(t) and len(t) >= 1:
                        title_y = y
                        break
        elif status_y is not None:
            for y, x, text, conf in items_with_y:
                if y > status_y + 8:
                    t = text.strip()
                    if not is_title_noise(t) and len(t) >= 1:
                        title_y = y
                        break

        if title_y is not None:
            self.log(f"  [디버그] title_y: {title_y}")

            # 제목 줄의 모든 텍스트를 합침 (전체 페이지 winocr 결과)
            title_parts = []
            for y, x, text, conf in items_with_y:
                if abs(y - title_y) < 15:
                    title_parts.append((x, text.strip()))
            title_parts.sort(key=lambda item: item[0])
            combined_title = " ".join(item[1] for item in title_parts)

            if combined_title and len(combined_title) >= 2:
                # --- Tesseract 한글 보정 (winocr가 '협'→'리' 등 오인식하는 문제) ---
                try:
                    import pytesseract
                    from PIL import ImageOps, ImageEnhance
                    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

                    crop_top = max(0, int(title_y - 25))
                    crop_bottom = min(screenshot.height, int(title_y + 25))
                    title_crop = screenshot.crop((20, crop_top, 770, crop_bottom))
                    
                    # --- 이미지 전처리 강화 ---
                    # 1. 2배 확대 (인식률 향상에 가장 효과적)
                    title_crop = title_crop.resize((title_crop.width * 2, title_crop.height * 2), Image.Resampling.LANCZOS)
                    # 2. 대비 향상 (글자를 더 뚜렷하게)
                    enhancer = ImageEnhance.Contrast(title_crop)
                    title_crop = enhancer.enhance(2.0)
                    # 3. 여백 추가
                    title_crop = ImageOps.expand(title_crop, border=20, fill='white')

                    tess_title = pytesseract.image_to_string(
                        title_crop, lang='kor', config='--psm 7 --dpi 300'
                    ).strip()
                    self.log(f"  [Tesseract 보정] {tess_title}")

                    # 단어별 교체: WinOCR 결과가 부실한 경우에만 Tesseract로 보완
                    if tess_title:
                        wk_words = combined_title.split()
                        ts_words = tess_title.split()
                        merged = []
                        ts_idx = 0
                        for wk in wk_words:
                            # 1. 한글이 포함된 경우: WinOCR이 이미 잘 읽었다면 우선 (예: 초심자교육)
                            has_kr = any('\uac00' <= c <= '\ud7a3' for c in wk)
                            has_en_num = any(c.isalnum() for c in wk)
                            
                            if has_kr:
                                # WinOCR이 이미 한글을 잘 읽었다면(2글자 이상 등) 그대로 유지
                                # 단, 숫자나 기호가 섞여있어 부자연스러운 경우에만 Tesseract 참고
                                if len(wk) >= 2 and not re.search(r'[0-9!@#$%^&*()]', wk):
                                    merged.append(wk)
                                    if ts_idx < len(ts_words): ts_idx += 1
                                    continue
                            
                            # 2. 영문/숫자인 경우: WinOCR이 이미 잘 읽었다면(File 등) 그대로 유지
                            if not has_kr and has_en_num:
                                if len(wk) >= 3: # 3글자 이상 영문/숫자면 신뢰 (File, 2026 등)
                                    merged.append(wk)
                                    if ts_idx < len(ts_words): ts_idx += 1
                                    continue

                            # 3. 그 외 부실한 결과인 경우 Tesseract와 비교 보완
                            if ts_idx < len(ts_words):
                                for ti in range(ts_idx, min(len(ts_words), ts_idx + 3)):
                                    # Tesseract에 의미 있는 한글/영문이 있으면 채택
                                    if any(c.isalnum() for c in ts_words[ti]):
                                        merged.append(ts_words[ti])
                                        ts_idx = ti + 1
                                        break
                                else:
                                    merged.append(wk)
                            else:
                                merged.append(wk)
                        combined_title = " ".join(merged)
                except Exception as e:
                    self.log(f"  [Tesseract 보정 실패] {e}")

                # 후처리 보정
                # 0. 중국어/한자 OCR 쓰레기 제거 (winocr가 한글을 CJK로 오인식)
                combined_title = re.sub(r'[\u4e00-\u9fff\u3400-\u4dbf]+', '', combined_title)
                combined_title = re.sub(r'\s+', ' ', combined_title).strip()
                # 1. [한글1 → [한글] (닫는 괄호 ']'가 '1'로 오인식)
                combined_title = re.sub(r'\[([가-힣A-Za-z\s]+)1(\s|$)', r'[\1]\2', combined_title)
                # 2. testll → test11, testl → test1 (소문자 l → 숫자 1)
                combined_title = re.sub(r'(test)([0-9]*)ll\b', lambda m: m.group(1) + m.group(2) + '11', combined_title)
                combined_title = re.sub(r'(test)([0-9]*)l\b', lambda m: m.group(1) + m.group(2) + '1', combined_title)
                # 3. tO → to
                combined_title = combined_title.replace(' tO ', ' to ')
                # 4. [한글 → [한글] (닫는 괄호 누락 보정)
                if '[' in combined_title and ']' not in combined_title:
                    combined_title = re.sub(r'\[([가-힣A-Za-z]+)(\s)', r'[\1]\2', combined_title)

                title = combined_title


        return {"author": author, "id": item_id, "title": title, "start_date": start_date, "start_time": start_time, "end_date": end_date, "end_time": end_time}

if __name__ == "__main__":
    root = tk.Tk()
    app = KakaoMacroGUI(root)
    root.mainloop()