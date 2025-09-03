import os
from Reflash_list_generator import get_reflash_list
from _ast import Pass
from collections import defaultdict
from datetime import datetime, timedelta
from FCT_NG_Table_Embedded_경로파일명수정본 import create_excel_file_for_table, display_excel_embedded
from tkinter import *  # Imports all general tkinter modules
from tkinter import messagebox, filedialog
from tkinter.ttk import Combobox
from tkcalendar import DateEntry
import locale
from reportlab.pdfgen import canvas
from PIL import ImageGrab
import tempfile


# ✅ 병렬 분석 실행 함수
import concurrent.futures
from collections import defaultdict

def run_analysis_scripts_parallel(input_date, shift):
    def safe_execute(fn, name):
        try:
            fn(input_date, shift)
            print(f"✅ {name} 완료")
        except Exception as e:
            print(f"❌ {name} 실패: {e}")

    from FCT_NG_Backend_debug_verbose import run_fct_ng_analysis
    from Ford_A_C_LED_NG_Backend import run_led_ng_analysis
    from FORD_A_C_FCT_Percentage_Backend import run_fct_passrate_analysis

    scripts = [
        (run_fct_ng_analysis, "FCT NG 분석"),
        (run_led_ng_analysis, "LED NG 분석"),
        (run_fct_passrate_analysis, "FCT PASS율 분석")
    ]

    with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
        futures = [executor.submit(safe_execute, fn, name) for fn, name in scripts]
        for future in concurrent.futures.as_completed(futures):
            pass

# ✅ 생산량, OEE 계산 병렬 처리 함수
def calculate_production_metrics(time_slot_counts, cycle_times):
    def calc_part_summary():
        part_summary = defaultdict(lambda: {"양품": 0, "불량": 0})
        for slot_data in time_slot_counts.values():
            for part, counts in slot_data.items():
                part_summary[part]["양품"] += counts["양품"]
                part_summary[part]["불량"] += counts["불량"]
        return part_summary

    def calc_total_ok_ng():
        ok, ng = 0, 0
        for slot_data in time_slot_counts.values():
            for counts in slot_data.values():
                ok += counts["양품"]
                ng += counts["불량"]
        return ok, ng

    def calc_standard_times(part_summary):
        std_times = {}
        total_std_time = 0.0
        for part, counts in part_summary.items():
            ct = cycle_times.get(part, 0)
            std_time = (counts["양품"] * ct) / 3600
            std_times[part] = std_time
            total_std_time += std_time
        return std_times, total_std_time

    with concurrent.futures.ThreadPoolExecutor() as executor:
        future_summary = executor.submit(calc_part_summary)
        future_ok_ng = executor.submit(calc_total_ok_ng)

        part_summary = future_summary.result()
        total_ok, total_ng = future_ok_ng.result()

        std_times, total_std_time = calc_standard_times(part_summary)

    return part_summary, total_ok, total_ng, std_times, total_std_time


# PDF로 저장하는 함수
def save_as_pdf(result_window):
    import datetime  # 날짜 데이터를 처리하기 위해 필요
    from tkinter import messagebox  # 메시지 박스를 띄우기 위한 임포트
    from tkinter import filedialog  # 파일 저장 대화 상자를 관리하기 위한 임포트
    from reportlab.pdfgen import canvas  # PDF 생성 라이브러리
    from PIL import ImageGrab  # 화면 캡처를 위해 사용
    import tempfile  # 임시 파일 저장에 사용
    import os  # 파일 경로 처리를 위해 사용

    # --- 파일명 생성 로직 시작 ---
    # 날짜 가져오기
    selected_date = date_entry.get()  # 사용자 입력 (yyyy-mm-dd 형태라고 가정)
    if not selected_date:  # 선택하지 않은 경우 오늘 날짜 사용
        selected_date = datetime.datetime.now().strftime('%Y-%m-%d')

    # yyyy.mm.dd 형태로 변환
    formatted_date = selected_date.replace("-", ".")

    # 주간/야간 선택
    try:
        selected_shift = shift_combobox.get()  # '주간' 또는 '야간'
    except Exception:
        selected_shift = "주간"  # 기본값

    # 주간 또는 야간의 기본값 설정
    if selected_shift not in ['주간', '야간']:
        selected_shift = "주간"

    # 파일명 생성
    generated_file_name = f"{formatted_date} {selected_shift} Production film.pdf"
    # --- 파일명 생성 로직 끝 ---

    # GUI 창의 이미지를 캡처
    try:
        x = result_window.winfo_rootx()
        y = result_window.winfo_rooty()
        width = result_window.winfo_width()
        height = result_window.winfo_height()

        # 화면 캡처 조정 옵션
        offset_x = 1  # 좌우 여백
        offset_y = 10  # 상하 여백
        scale_factor = 1

        # 조정된 크기 계산
        adjusted_width = int(width * scale_factor)
        adjusted_height = int(height * scale_factor)

        # 캡처 영역 설정
        screenshot = ImageGrab.grab(bbox=(
            x - offset_x,  # 좌측 조정
            y - offset_y,  # 상측 조정
            x - offset_x + adjusted_width,  # 우측 조정
            y - offset_y + adjusted_height  # 하측 조정
        ))

        # 임시 파일 저장 경로 설정
        temp_image_path = tempfile.NamedTemporaryFile(suffix=".png", delete=False).name
        screenshot.save(temp_image_path, "PNG")
    except Exception as e:
        messagebox.showerror("오류", f"이미지 캡처 중 오류가 발생했습니다: {str(e)}")
        return

    # PDF 저장 경로 선택
    base_path = os.path.expanduser("~/Documents")  # 기본 저장 위치를 사용자 문서 폴더로 설정
    default_file_path = os.path.join(base_path, generated_file_name)

    # PDF 저장 대화상자
    pdf_file_path = filedialog.asksaveasfilename(
        initialfile=generated_file_name,
        defaultextension=".pdf",
        filetypes=[("PDF 파일", "*.pdf")],
        title="PDF 파일로 저장"
    )
    if not pdf_file_path:  # 사용자가 경로를 선택하지 않은 경우
        return

    # PDF 파일 생성
    try:
        pdf = canvas.Canvas(pdf_file_path)

        # PDF 크기를 이미지 크기와 일치하도록 설정
        image_width, image_height = screenshot.size
        pdf.setPageSize((image_width, image_height))

        # 이미지 삽입
        pdf.drawImage(temp_image_path, 0, 0, width=image_width, height=image_height)

        # PDF 저장
        pdf.save()

        # PDF 저장 완료 메시지 출력
        messagebox.showinfo("저장 성공", f"PDF 파일이 성공적으로 저장되었습니다: {pdf_file_path}")
    except Exception as e:
        messagebox.showerror("오류", f"PDF 저장 중 오류가 발생했습니다: {str(e)}")
    finally:
        # 임시 이미지 파일 삭제
        if os.path.exists(temp_image_path):
            os.remove(temp_image_path)

# **Locale 설정 (달력 표기를 한국어로 설정)**
try:
    locale.setlocale(locale.LC_TIME, 'ko_KR.UTF-8')
except locale.Error:
    try:
        # Windows용 한국어 로케일 설정
        locale.setlocale(locale.LC_TIME, 'Korean')
    except locale.Error:
        locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')  # 기본값이 안 되면 영어로 설정

import tkinter as tk
from tkinter import simpledialog, messagebox
import json
import os

# **기본 경로 및 초기 데이터**
BASE_PATH = r"C:\Users\user\Desktop\FORD A+C VISION 로그파일"
TXT_FILE_PATH = r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C_PN,CT\PN,CT.txt"

# 초기 설정 파일에서 데이터 읽기
def load_default_values_from_file():
    global DEFAULT_MAPPING, DEFAULT_CYCLE_TIMES
    try:
        if os.path.exists(TXT_FILE_PATH):
            with open(TXT_FILE_PATH, "r", encoding="utf-8") as txt_file:
                lines = txt_file.readlines()
                if len(lines) >= 2:
                    DEFAULT_MAPPING = eval(lines[0].strip())  # 첫 번째 줄의 데이터를 eval을 통해 딕셔너리로 변환
                    DEFAULT_CYCLE_TIMES = eval(lines[1].strip())  # 두 번째 줄의 데이터를 eval을 통해 딕셔너리로 변환
                else:
                    raise ValueError("PN,CT.txt 파일의 형식이 잘못되었습니다!")
        else:
            raise FileNotFoundError(f"{TXT_FILE_PATH} 파일이 존재하지 않습니다!")
    except Exception as e:
        messagebox.showerror("오류", f"기본값을 로드하는 데 실패했습니다: {e}")
        # 기본값 사용할 경우
        DEFAULT_MAPPING = {'C': '35643009', 'P': '35643010', '1': '35654264', 'N': '35749091', 'J': '35915729',
                           'S': '35915730'}
        DEFAULT_CYCLE_TIMES = {'35643009': 8.2, '35643010': 8.2, '35654264': 8.2, '35749091': 8.2, '35915729': 9.25,
                               '35915730': 9.25}


load_default_values_from_file()  # 파일에서 초기값 로드

# 설정 파일 경로 및 초기값
SETTINGS_FILE = "settings.json"
SETTINGS_PASSWORD = "leejangwoo1!"

# 전역 변수
MAPPING = DEFAULT_MAPPING

# ==== 매핑 정보 전역 변수 ====
NORMAL_MAPPING = {}
NORMAL_CT = {}
REFLASH_MAPPING = {}
REFLASH_CT = {}
REFLASH_FILE_NAMES = set()

def load_normal_mapping():
    pn_ct_path = r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C_PN,CT\PN,CT.txt"
    try:
        with open(pn_ct_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
            global NORMAL_MAPPING, NORMAL_CT
            NORMAL_MAPPING = eval(lines[0].strip())
            NORMAL_CT = eval(lines[1].strip())
        print("[INFO] NORMAL 매핑 로드 완료:", NORMAL_MAPPING)
    except Exception as e:
        print(f"[ERROR] NORMAL 매핑 로드 실패: {e}")

def load_reflash_mapping_and_list(input_date, shift):
    reflash_list_path = fr"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C_Reflash\{input_date}_{shift}_Reflash_list.txt"
    pn_ct_reflash_path = r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C_PN,CT\PN,CT_Reflash.txt"
    try:
        with open(pn_ct_reflash_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
            global REFLASH_MAPPING, REFLASH_CT
            REFLASH_MAPPING = eval(lines[0].strip())
            REFLASH_CT = eval(lines[1].strip())
        with open(reflash_list_path, "r", encoding="utf-8") as f:
            global REFLASH_FILE_NAMES
            REFLASH_FILE_NAMES = set(line.strip() for line in f if line.strip())
        print("[INFO] REFLASH 매핑 로드 완료:", REFLASH_MAPPING)
        print("[INFO] REFLASH 파일 개수:", len(REFLASH_FILE_NAMES))
    except Exception as e:
        print(f"[ERROR] REFLASH 매핑 로드 실패: {e}")

def get_part_info(file_name):
    """파일명에 따라 적절한 매핑/CT 반환"""
    if file_name in REFLASH_FILE_NAMES:
        key_char = file_name[17] if len(file_name) > 17 else None
        return REFLASH_MAPPING.get(key_char), REFLASH_CT
    else:
        key_char = file_name[17] if len(file_name) > 17 else None
        return NORMAL_MAPPING.get(key_char), NORMAL_CT

CYCLE_TIMES = DEFAULT_CYCLE_TIMES


# **설정 파일에서 데이터 로드**
def load_settings():
    global MAPPING, CYCLE_TIMES
    if os.path.exists(SETTINGS_FILE):  # JSON 파일이 존재하는 경우
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                MAPPING = data.get("MAPPING", DEFAULT_MAPPING)
                CYCLE_TIMES = data.get("CYCLE_TIMES", DEFAULT_CYCLE_TIMES)
        except Exception as e:
            messagebox.showerror("오류", f"설정을 로드하지 못했습니다: {e}")
    else:
        # 파일이 없으면 기본값으로 초기화 후 저장
        save_settings_to_file()


# **설정 파일에 데이터 저장**
def save_settings_to_file():
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump({"MAPPING": MAPPING, "CYCLE_TIMES": CYCLE_TIMES}, f, indent=4, ensure_ascii=False)

        # 텍스트 파일에도 저장
        with open(TXT_FILE_PATH, "w", encoding="utf-8") as txt_file:
            txt_file.write(str(MAPPING) + "\n")
            txt_file.write(str(CYCLE_TIMES) + "\n")
    except Exception as e:
        messagebox.showerror("오류", f"설정을 저장하지 못했습니다: {e}")

# **설정 창 열기**
def open_settings():
    # 비밀번호 입력받기
    password = simpledialog.askstring("비밀번호", "설정에 접근하려면 비밀번호를 입력하세요:", show="*")

    if password == SETTINGS_PASSWORD:  # 비밀번호 확인
        settings_window = tk.Toplevel(root)
        settings_window.title("설정")
        settings_window.geometry("400x400")

        def save_settings():
            try:
                # MAPPING 업데이트
                mappings_input = mapping_entry.get("1.0", tk.END).strip()
                updated_mappings = eval(mappings_input)  # 문자열을 딕셔너리로 변환
                if not isinstance(updated_mappings, dict):
                    raise ValueError("MAPPING의 형식이 올바르지 않습니다!")

                # CYCLE_TIMES 업데이트
                cycle_times_input = cycle_times_entry.get("1.0", tk.END).strip()
                updated_cycle_times = eval(cycle_times_input)  # 문자열을 딕셔너리로 변환
                if not isinstance(updated_cycle_times, dict):
                    raise ValueError("CYCLE_TIMES의 형식이 올바르지 않습니다!")

                # 전역 변수에 변경사항 적용
                global MAPPING, CYCLE_TIMES
                MAPPING = updated_mappings
                CYCLE_TIMES = updated_cycle_times

                # 변경사항 저장
                save_settings_to_file()

                # 성공 메시지
                messagebox.showinfo("완료", "설정이 성공적으로 저장되었습니다!")
                settings_window.destroy()
            except Exception as e:
                messagebox.showerror("오류", f"설정을 저장하지 못했습니다: {e}")

        # MAPPING 현재 상태 표시
        tk.Label(settings_window, text="MAPPING").pack()
        mapping_entry = tk.Text(settings_window, height=8, width=50)
        mapping_entry.insert(tk.END, str(MAPPING))
        mapping_entry.pack()

        # CYCLE_TIMES 현재 상태 표시
        tk.Label(settings_window, text="CYCLE_TIMES").pack()
        cycle_times_entry = tk.Text(settings_window, height=8, width=50)
        cycle_times_entry.insert(tk.END, str(CYCLE_TIMES))
        cycle_times_entry.pack()

        # 저장 버튼
        tk.Button(settings_window, text="저장", command=save_settings).pack(pady=10)

    else:
        messagebox.showerror("접근 불가", "비밀번호가 올바르지 않습니다!")

# **시간 조건**
DAY_START = datetime.strptime("08:30:00", "%H:%M:%S").time()
DAY_END = datetime.strptime("20:29:59", "%H:%M:%S").time()
NIGHT_START = datetime.strptime("20:30:00", "%H:%M:%S").time()
NIGHT_END = datetime.strptime("08:29:59", "%H:%M:%S").time()

EXCLUDE_START = datetime.strptime("00:00:00", "%H:%M:%S").time()
EXCLUDE_END = datetime.strptime("08:29:59", "%H:%M:%S").time()

# **파일명에서 년월일 추출**
def extract_file_date(file_name):
    if len(file_name) < 52:  # 파일명이 충분한 길이를 가지지 않으면 None 반환
        return None
    try:
        # 파일명 18번째 문자가 C, J, 1일 경우
        if file_name[17] in ['C', 'J', '1']:
            year = file_name[31:35]  # 32~35번째 문자
            month = file_name[35:37]  # 36, 37번째 문자
            day = file_name[37:39]  # 38, 39번째 문자
        # 파일명 18번째 문자가 P, N, S일 경우
        elif file_name[17] in ['P', 'N', 'S']:
            year = file_name[32:36]  # 33~36번째 문자
            month = file_name[36:38]  # 37, 38번째 문자
            day = file_name[38:40]  # 39, 40번째 문자
        else:
            return None
        # yyyymmdd 형식의 날짜 반환
        return f"{year}{month}{day}"
    except Exception:
        return None

# **파일명에서 시간 추출**
def extract_file_time(file_name):
    if len(file_name) < 52:
        return None
    try:
        if file_name[17] in ['P', 'N', 'S']:
            hh = file_name[40:42]
            mm = file_name[42:44]
            ss = file_name[44:46]
        elif file_name[17] in ['C', '1', 'J']:
            hh = file_name[39:41]
            mm = file_name[41:43]
            ss = file_name[43:45]
        else:
            return None
        return datetime.strptime(f"{hh}:{mm}:{ss}", "%H:%M:%S").time()
    except ValueError:
        return None

# **시간 범위 확인 함수**
def is_file_in_slot(file_time, shift, folder_type, file_date, base_date):
    # 기준 날짜와 +1일 계산
    today_date = base_date.strftime("%Y%m%d")  # 기준 날짜
    tomorrow_date = (base_date + timedelta(days=1)).strftime("%Y%m%d")  # 기준 날짜 +1일

    # 파일 날짜가 기준 날짜와 동일한 경우
    if file_date == today_date:
        # 시간 예외 조건 적용 (00:00:00 ~ 08:29:59 제외)
        if folder_type == "today" and EXCLUDE_START <= file_time <= EXCLUDE_END:
            return False
        # 주간/야간 시간 조건 확인
        if shift == "주간":
            return DAY_START <= file_time <= DAY_END
        elif shift == "야간":
            return NIGHT_START <= file_time or file_time <= NIGHT_END
        return False

    # 파일 날짜가 기준 날짜의 +1일인 경우 (내일 날짜)
    if file_date == tomorrow_date:
        # 시간 조건에 관계없이 포함
        return True

    # 그 외 날짜는 제외
    return False

# **시간 차이 계산**
def calculate_time_difference(start_str, end_str):
    try:
        start_time = datetime.strptime(start_str.strip(), "%H:%M")
        end_time = datetime.strptime(end_str.strip(), "%H:%M")
        if end_time < start_time:
            end_time += timedelta(days=1)
        return (end_time - start_time).total_seconds() / 3600
    except ValueError:
        return 0

# **시간 선택 콤보박스 생성**
def create_time_picker(parent):
    frame = Frame(parent)
    frame.pack(side=LEFT, padx=2)
    hours = [f"{i:02}" for i in range(24)]
    minutes = ["00", "05", "10", "15", "20", "25", "30", "35", "40", "45", "50", "55"]
    hour_combo = Combobox(frame, values=hours, width=3, state="readonly")
    minute_combo = Combobox(frame, values=minutes, width=3, state="readonly")
    hour_combo.set(hours[0])
    minute_combo.set(minutes[0])
    hour_combo.pack(side=LEFT)
    Label(frame, text=":").pack(side=LEFT)
    minute_combo.pack(side=LEFT)
    return hour_combo, minute_combo


# **시간 행 추가**
def add_time_row(parent, rows, initial=False):
    row_frame = Frame(parent)
    row_frame.pack(pady=2, anchor="w")
    start_hour, start_minute = create_time_picker(row_frame)
    Label(row_frame, text=" ~ ").pack(side=LEFT)
    end_hour, end_minute = create_time_picker(row_frame)

    # 텍스트 입력 필드 추가 (계획 정지/비가동 시간 옆에)
    # "내용" 텍스트 추가 (왼쪽에 배치)
    Label(row_frame, text="내용").pack(side=LEFT, padx=5)
    # 텍스트 입력 필드 추가 (계획 정지/비가동 시간 옆에)
    annotation_entry = Entry(row_frame, width=20)  # 텍스트 입력 필드(넓이 조절 가능)
    annotation_entry.insert(0, "")  # 기본 값은 빈 문자열
    annotation_entry.pack(side=LEFT, padx=5)

    if initial:
        Button(row_frame, text="+", command=lambda: add_time_row(parent, rows)).pack(side=LEFT, padx=5)
    else:
        Button(row_frame, text="-", command=lambda: delete_time_row(row_frame, rows)).pack(side=LEFT, padx=5)

    rows.append((start_hour, start_minute, end_hour, end_minute, row_frame, annotation_entry))  # 추가 필드 포함

# **시간 행 삭제**
def delete_time_row(row_frame, rows):
    for row in rows:
        if row[4] == row_frame:
            rows.remove(row)
            row_frame.destroy()
            break

# **시간 범위 포함 여부 확인**
def is_time_in_range(start_time, end_time, compare_time):
    start = datetime.strptime(start_time, "%H:%M").time()
    end = datetime.strptime(end_time, "%H:%M").time()
    compare = datetime.strptime(compare_time, "%H:%M").time()

    if start > end:  # 밤 사이 시간대
        return compare >= start or compare <= end
    return start <= compare <= end


# **분석 실행 함수**
from FCT_Graph_Backend_embed_final_SAFE import generate_graphs_embedded  # 📊 그래프 분석 모듈
from sparepart_graph_backend import create_sparepart_graph_embedded



def run_analysis():

    # === Step 0: Generate Reflash list before any analysis ===
    try:
        input_date = date_entry.get_date().strftime("%Y%m%d")
        shift = shift_combobox.get()
        print("[INFO] Reflash list generation started...")
        get_reflash_list(input_date, shift)
        print("[INFO] Reflash list generation completed.")
    except Exception as e:
        print(f"[ERROR] Failed to generate Reflash list: {e}")
    # === 매핑 데이터 로드 ===
    input_date = date_entry.get_date().strftime("%Y%m%d")
    shift = shift_combobox.get()
    load_normal_mapping()
    load_reflash_mapping_and_list(input_date, shift)

    try:
        print(f"[GUI] 날짜={input_date}, Shift={shift}")

        # ✅ FCT NG 분석
        from FCT_NG_Backend_debug_verbose import run_fct_ng_analysis
        run_fct_ng_analysis(input_date, shift)

        # ✅ LED NG 분석
        from Ford_A_C_LED_NG_Backend import run_led_ng_analysis
        run_led_ng_analysis(input_date, shift)

        # ✅ FCT PASS율 분석
        from FORD_A_C_FCT_Percentage_Backend import run_fct_passrate_analysis
        run_fct_passrate_analysis(input_date, shift)

    except Exception as e:
        print(f"[GUI] run_analysis 오류: {e}")

        # ✅ FCT NG 분석
        from FCT_NG_Backend_debug_verbose import run_fct_ng_analysis
        run_fct_ng_analysis(input_date, shift)

        # ✅ LED NG 분석
        from Ford_A_C_LED_NG_Backend import run_led_ng_analysis
        run_led_ng_analysis(input_date, shift)

    except Exception as e:
        print(f"[GUI] run_analysis 오류: {e}")


        # ✅ FCT 분석 자동 실행
        from FCT_NG_Backend_debug_verbose import run_fct_ng_analysis
        run_fct_ng_analysis(input_date, shift)
    except Exception as e:
        print(f"[GUI] run_analysis 내부 오류: {e}")


    # **주간/야간 선택 여부 확인**
    shift = shift_combobox.get()
    if not shift:
        messagebox.showerror("경고", "Shift를 선택해주세요!")
        return  # 함수 종료

    # **작업자 이름 가져오기**
    worker_name = worker_name_entry.get() if worker_name_entry.get() else ""

    # **Order No. 가져오기 (Text 위젯은 다르게 처리)**
    try:
        order_no = order_no_entry.get("1.0", tk.END).strip()  # 공백 제거
    except Exception:
        messagebox.showerror("오류", "Order No.를 입력해주세요!")
        return

    # **Box 현황 가져오기**
    box_status = box_status_entry.get().strip()

    # **마스터 샘플 상태 확인**
    master_sample_status = master_sample_entry.get().strip()

    # **Sparepart 사용량 확인 및 유효성 검사**
    try:
        spare_parts_usage = {}
        for label, entry in spare_parts_entries.items():
            value = entry.get().strip()
            if not value:  # 입력값이 없을 경우
                messagebox.showerror("경고", f"SPAREPART 사용량을 입력하세요. 사용량이 없으면 0이라고 입력하세요.")
                return  # 함수 종료
            try:
                spare_parts_usage[label] = int(value)  # 숫자로 변환
            except ValueError:
                messagebox.showerror("오류", f"'{label}' 사용량은 숫자로 입력해야 합니다.")
                return  # 함수 종료
    except Exception as e:
        messagebox.showerror("오류", f"Sparepart 입력 확인 중 오류 발생: {e}")
        return

    # **주요 불량 내용 가져오기 (Text 위젯 사용)**
    try:
        defect_details = defect_entry.get("1.0", tk.END).strip()
    except Exception:
        defect_details = ""

    # **생산 건의 내용 가져오기**
    try:
        suggestions = suggestion_entry.get("1.0", tk.END).strip()
    except Exception:
        suggestions = ""

    # **계획 정지 시간과 비가동 시간 계산**
    planned_downtime, total_downtime = 0.0, 0.0

    # **계획 정지 시간 계산**
    for row in planned_downtime_rows:
        try:
            start = f"{row[0].get()}:{row[1].get()}"
            end = f"{row[2].get()}:{row[3].get()}"
            annotation = row[5].get()
            planned_downtime += calculate_time_difference(start, end)
        except Exception as e:
            messagebox.showwarning("경고", f"계획 정지 시간 입력 오류: {e}")
            continue

    # **비가동 시간 계산**
    for row in downtime_rows:
        try:
            start = f"{row[0].get()}:{row[1].get()}"
            end = f"{row[2].get()}:{row[3].get()}"
            annotation = row[5].get()
            total_downtime += calculate_time_difference(start, end)
        except Exception as e:
            messagebox.showwarning("경고", f"비가동 시간 입력 오류: {e}")
            continue

    # **날짜 형식 확인**
    try:
        base_date = datetime.strptime(input_date, "%Y%m%d")
    except ValueError:
        messagebox.showerror("에러", "유효한 날짜를 입력해주세요.")
        return

    # **폴더 경로 생성**
    today_folder = os.path.join(BASE_PATH, input_date, "GoodFile")
    tomorrow_folder = os.path.join(BASE_PATH, (base_date + timedelta(days=1)).strftime("%Y%m%d"), "GoodFile")

    # **데이터 분석을 위한 변수 초기화**
    time_slot_counts = defaultdict(lambda: defaultdict(lambda: {"양품": 0, "불량": 0}))

    time_slots = {
        "주간": {
            "A": ("08:30:00", "10:29:59"),
            "B": ("10:30:00", "12:29:59"),
            "C": ("12:30:00", "14:29:59"),
            "D": ("14:30:00", "16:29:59"),
            "E": ("16:30:00", "18:29:59"),
            "F": ("18:30:00", "20:29:59"),
        },
        "야간": {
            "A'": ("20:30:00", "22:29:59"),
            "B'": ("22:30:00", "00:29:59"),
            "C'": ("00:30:00", "02:29:59"),
            "D'": ("02:30:00", "04:29:59"),
            "E'": ("04:30:00", "06:29:59"),
            "F'": ("06:30:00", "08:29:59"),
        },
    }

    # **폴더 목록 순회 및 분석**
    folders = [(today_folder, "today"), (tomorrow_folder, "tomorrow")]
    for folder_path, folder_type in folders:
        # 주간일 경우 tomorrow 폴더 생략
        if shift == "주간" and folder_type == "tomorrow":
            continue

        if not os.path.exists(folder_path):
            messagebox.showinfo("경로 없음", f"경로가 존재하지 않습니다: {folder_path}")
            continue

        for file_name in os.listdir(folder_path):
            file_time = extract_file_time(file_name)  # 파일 시간 추출
            file_date = extract_file_date(file_name)  # 파일 날짜 추출
            if not file_time or not file_date or not is_file_in_slot(file_time, shift, folder_type, file_date,
                                                                     base_date):
                continue

            part_number, ct_map = get_part_info(file_name)
            if not part_number:
                continue
            # CYCLE_TIMES는 ct_map을 사용
            CYCLE_TIMES.update(ct_map)
            if not part_number:
                continue

            is_ok = file_name[50] == 'P' or file_name[51] == 'P'

            for slot, (start, end) in time_slots[shift].items():
                start_time = datetime.strptime(start, "%H:%M:%S").time()
                end_time = datetime.strptime(end, "%H:%M:%S").time()

                if start_time <= file_time <= end_time or (
                        start_time > end_time and (file_time >= start_time or file_time <= end_time)):
                    time_slot_counts[slot][part_number]["양품" if is_ok else "불량"] += 1
                    break

    # **총계 및 표준 생산 시간 계산**
    total_ok, total_ng = 0, 0
    part_summary = defaultdict(lambda: {"양품": 0, "불량": 0})
    standard_production_time_summary = 0.0
    part_standard_times = {}

    for slot, parts in time_slot_counts.items():
        for part, counts in parts.items():
            total_ok += counts["양품"]
            total_ng += counts["불량"]
            part_summary[part]["양품"] += counts["양품"]
            part_summary[part]["불량"] += counts["불량"]

    for part, counts in part_summary.items():
        ct = CYCLE_TIMES.get(part, 0)
        standard_time = (counts["양품"] * ct) / 3600
        part_standard_times[part] = standard_time
        standard_production_time_summary += standard_time

    total_work_time = 12.0
    working_time = total_work_time - planned_downtime
    loss_time = total_work_time - standard_production_time_summary - total_downtime
    oee = (standard_production_time_summary / working_time) * 100 if working_time > 0 else 0

    # PASS율(양품률) 초기값 설정
    pass_rate = 0.0  # 초기화하여 참조 오류 방지

    # PASS율(양품률) 계산
    if (total_ok + total_ng) > 0:  # 0으로 나누는 경우 방지
        pass_rate = (total_ok / (total_ok + total_ng)) * 100

    # **결과 저장**
    try:
        save_path = "C:\\Ford A+C Test center_생산 분석 프로그램_Rev02\\Data\\FORD A+C_Data\\FORD A+C_OEE"  # 경로
        os.makedirs(save_path, exist_ok=True)
        file_name = f"{base_date.strftime('%y.%m.%d')}_{shift}.txt"
        with open(os.path.join(save_path, file_name), "w", encoding="utf-8") as f:
            f.write(f"날짜: {base_date.strftime('%Y-%m-%d')}\n")
            f.write(f"Shift: {shift}\n")
            f.write(f"작업자명: {worker_name}\n")
            f.write(f"\nOEE: {oee:.2f}%\n")
            f.write(f"FCT > LED PASS율: {pass_rate:.2f}%\n")  # PASS율 추가
            f.write("\n=== Spareparts 사용량 ===\n")

            # === Spareparts 사용량 ===
            f.write("\n=== Spareparts 사용량 ===\n")
            for part, quantity in spare_parts_usage.items():
                f.write(f"{part}: {quantity}\n")

            # === 계획 정지 시간 ===
            f.write("\n=== 계획 정지 시간 ===\n")
            f.write(f"총 계획 정지 시간: {planned_downtime:.2f} 시간\n")
            for row in planned_downtime_rows:
                start = f"{row[0].get()}:{row[1].get()}"
                end = f"{row[2].get()}:{row[3].get()}"
                annotation = row[5].get()
                if start != end:
                    f.write(f"- {start} ~ {end} / 내용: {annotation}\n")

            # === 비가동 시간 ===
            f.write("\n=== 비가동 시간 ===\n")
            f.write(f"총 비가동 시간: {total_downtime:.2f} 시간\n")
            for row in downtime_rows:
                start = f"{row[0].get()}:{row[1].get()}"
                end = f"{row[2].get()}:{row[3].get()}"
                annotation = row[5].get()
                if start != end:
                    f.write(f"- {start} ~ {end} / 내용: {annotation}\n")

        # 저장 완료 시 메시지 표시 없음
    except Exception as e:
        # 저장 실패 시에만 메시지 표시
        messagebox.showerror("저장 실패", f"결과 저장 중 오류가 발생했습니다: {e}")

    # **결과 출력**
    result_window = Toplevel(root)
    result_window.title("생산 일보")

    # 출력창 크기와 위치 설정
    width, height = 1800, 900
    screen_width = result_window.winfo_screenwidth()
    screen_height = result_window.winfo_screenheight()

    # 화면 중앙에서 창 위치 계산
    x = int((screen_width - width) / 2)
    y = int((screen_height - height) / 7)

    # 창의 크기 및 설정
    result_window.geometry(f"{width}x{height}+{x}+{y}")

    # 출력창 내용
    result_text = Text(result_window, wrap=WORD)

    # **결과 출력 프레임 생성**
    main_frame = Frame(result_window)
    graph_frame = Frame(main_frame)
    graph_frame.pack(side=RIGHT, anchor='n', padx=10, pady=10)
    table_frame = Frame(graph_frame)
    table_frame.pack(side=BOTTOM, fill=BOTH, padx=5, pady=10)
    main_frame.pack(expand=1, fill=BOTH, padx=10, pady=28)

    # 표 셀(Cell) 높이를 동적으로 변경하기 위한 변수 및 함수 설정
    cell_height = DoubleVar(value=0.01)  # 기본 높이 값 (0.1)
    cell_height.set(0.01)  # 초기 높이 설정

    def update_table_height(event=None):
        """
        Scale 값 변경 시 테이블의 셀 높이를 동적으로 업데이트하는 함수
        """
        for widget in table_frame.winfo_children():
            widget.config(height=int(cell_height.get() * 1))  # 높이를 동적으로 설정 (0.01 단위에도 맞춤)

    # Canvas와 사각형(Rectangle)을 이용해 얇은 테이블 생성
    rows = 148
    columns = 5
    output_window_width = 400  # 출력 창의 너비 (픽셀 기준)
    output_window_height = 820  # 출력 창의 높이 (픽셀 기준)

    # 각 셀의 높이를 0.5픽셀처럼 보이도록 계산
    cell_height = output_window_height / rows

    # Canvas 생성
    canvas = Canvas(main_frame, width=output_window_width, height=output_window_height, bg="white")
    canvas.pack(side=LEFT, fill=BOTH, expand=True, padx=10, pady=10)

    # Shift(주/야간 교대) 및 관련 텍스트 설정
    shift_text = "08:30" if shift == "주간" else "20:30"  # 주간 선택 시 "08:30", 야간 선택 시 "20:30"
    production_text = "생산"  # '생산' 텍스트

    # 병합 셀 정의 (R1C1-R4C1과 R1C2-R4C2)
    cell_width = output_window_width / columns  # 각 셀의 너비
    merged_height = cell_height * 4  # 병합된 셀(R1~R4)의 높이

    # R1C1-R4C1 (Shift 시간) 병합된 셀 생성 및 텍스트 추가
    canvas.create_rectangle(0, 0, cell_width, merged_height, outline="black", fill="white")  # 사각형 생성
    canvas.create_text(
        cell_width / 2, merged_height - 1,  # 텍스트 위치: 셀 하단 정렬
        text=shift_text,  # Shift에 따라 "08:30" 또는 "20:30"
        font=("Arial", 9),  # 글꼴 및 크기
        anchor="s"  # 텍스트 하단 정렬(south)
    )

    # R1C2-R4C2 (생산 텍스트) 병합된 셀 생성 및 텍스트 추가
    start_x = cell_width  # 열 2 시작 X 좌표
    end_x = 2 * cell_width  # 열 2 끝 X 좌표
    canvas.create_rectangle(start_x, 0, end_x, merged_height, outline="black", fill="white")  # 사각형 생성
    canvas.create_text(
        (start_x + end_x) / 2, merged_height / 2,  # 텍스트 위치: 셀의 중앙
        text=production_text,  # '생산' 텍스트
        font=("Arial", 10, "bold"),  # 글꼴과 굵기 설정
        anchor="center"  # 텍스트 중앙 정렬
    )

    # R5C1 ~ R148C1: 6칸씩 병합된 셀 생성
    # 각 Shift에 맞는 시간 설정
    time_intervals_day = [
        "09:00", "09:30", "10:00", "10:30", "11:00", "11:30",
        "12:00", "12:30", "13:00", "13:30", "14:00", "14:30",
        "15:00", "15:30", "16:00", "16:30", "17:00", "17:30",
        "18:00", "18:30", "19:00", "19:30", "20:00", "20:30"
    ]

    time_intervals_night = [
        "21:00", "21:30", "22:00", "22:30", "23:00", "23:30",
        "00:00", "00:30", "01:00", "01:30", "02:00", "02:30",
        "03:00", "03:30", "04:00", "04:30", "05:00", "05:30",
        "06:00", "06:30", "07:00", "07:30", "08:00", "08:30"
    ]

    # Shift에 따른 시간 리스트 선택
    time_intervals = time_intervals_day if shift == "주간" else time_intervals_night

    # 병합된 셀을 순차적으로 생성 (R5C1 ~ R148C1에 해당)
    for i, time_text in enumerate(time_intervals):
        start_row = 5 + (i * 6)  # 병합된 섹션의 시작 행 (각 섹션 6칸씩)
        end_row = start_row + 6  # 병합된 영역의 끝 행

        start_y = (start_row - 1) * cell_height  # 시작 Y 좌표
        end_y = (end_row - 1) * cell_height  # 끝 Y 좌표

        # 병합된 셀 그리기
        canvas.create_rectangle(
            0, start_y, cell_width, end_y,  # X 좌표(C1 열 고정), Y 좌표
            outline="black", fill="white"  # 흰 배경과 검정 테두리
        )

        # 병합된 셀 안에 Shift에 따른 시간 텍스트 삽입
        canvas.create_text(
            cell_width / 2, end_y - 1,  # 셀의 중앙(가로)과 하단(세로 - 1px 여백)
            text=time_text,  # 시간 텍스트 리스트에서 가져옴
            font=("Arial", 8),  # 글꼴과 크기
            anchor="s"  # 텍스트 하단 정렬 (south)
        )

    # R5C2 이후: 개별 셀 생성 (초록색으로 음영 기본 처리)
    for row in range(4, rows):  # R5부터 마지막 행까지
        start_y = row * cell_height  # 셀 시작 Y 좌표
        end_y = start_y + cell_height  # 셀 끝 Y 좌표
        canvas.create_rectangle(
            cell_width, start_y, 2 * cell_width, end_y,
            outline="black"
        )

    # 나머지 셀 (C3-C5) 생성
    for col in range(2, columns):  # 3번째 열부터 마지막 열까지
        for row in range(rows):  # 모든 행에 대해 반복
            # 병합된 R1C2-R4C2와 겹치는 경우 건너뜀
            if 0 <= row < 4 and col == 1:  # R1C2-R4C2 제외
                continue
            start_x = col * cell_width  # 열의 시작 X 좌표
            end_x = start_x + cell_width  # 열의 끝 X 좌표
            start_y = row * cell_height  # 행의 시작 Y 좌표
            end_y = start_y + cell_height  # 행의 끝 Y 좌표

            # 일반 셀 생성
            canvas.create_rectangle(
                start_x, start_y, end_x, end_y,  # 셀의 위치
                outline="black", fill="white"  # 테두리와 배경색
            )

            # 병합된 R1C3 ~ R4C5 생성 및 텍스트 추가
            start_x = 2 * cell_width  # 열 3 시작 X 좌표
            end_x = 5 * cell_width  # 열 5 끝 X 좌표
            start_y = 0  # 첫 행 시작 Y 좌표
            end_y = merged_height  # 네 번째 행 끝 Y 좌표

            # 병합된 셀 생성
            canvas.create_rectangle(
                start_x, start_y, end_x, end_y,
                outline="black", fill="white"  # 테두리와 배경색
            )

            # 병합된 셀 안에 텍스트 "내용" 삽입 (가운데 정렬)
            canvas.create_text(
                (start_x + end_x) / 2, (start_y + end_y) / 2,  # 텍스트 위치: 병합된 셀 중앙
                text="내용",
                font=("Arial", 10, "bold"),  # 글꼴과 텍스트 스타일
                anchor="center"  # 텍스트 중앙 정렬
            )
            # 나머지 C3, C4, C5 열 병합 (R5 이후부터 끝까지)
            start_x = 2 * cell_width  # C3 시작 X 좌표
            end_x = 5 * cell_width  # C5 끝 X 좌표
            start_y = merged_height  # R5 시작 Y 좌표
            end_y = rows * cell_height  # 마지막 행 끝 Y 좌표

            # 병합된 C3 ~ C5 셀 생성 (R5 이후)
            canvas.create_rectangle(
                start_x, start_y, end_x, end_y,
                outline="black", fill="white"  # 배경색만 흰색, 텍스트 없음
            )

            # 범위 내 시간을 확인하는 함수 (동일 시간 처리 추가)
            def is_time_in_range(start, end, check_time):
                # 시간을 datetime 객체로 변환
                start_time = datetime.strptime(start, "%H:%M")
                end_time = datetime.strptime(end, "%H:%M")
                check_time_obj = datetime.strptime(check_time, "%H:%M")

                # 시작 시간과 종료 시간이 동일할 경우 범위로 취급하지 않음
                if start_time == end_time:
                    return False

                # 끝 시간이 다음 날로 넘어가는 경우 처리
                if end_time < start_time:
                    end_time += timedelta(days=1)  # 다음 날로 끝 시간을 이동

                    # check_time도 다음 날로 넘어가는 경우를 처리
                    if check_time_obj < start_time:
                        check_time_obj += timedelta(days=1)

                # check_time이 범위 내에 있는지 확인
                return start_time <= check_time_obj <= end_time

            # Shift에 따라 동적으로 additional_time_intervals 생성
            def generate_intervals(start_time, count, interval_minutes):
                intervals = []  # 시간을 저장할 리스트
                current_time = start_time
                for _ in range(count):  # 총 count 번 반복 (예: 144개)
                    intervals.append(current_time.strftime("%H:%M"))  # 현재 시간을 "HH:MM" 형식으로 추가
                    current_time += timedelta(minutes=interval_minutes)  # interval_minutes(5분)만큼 추가
                return intervals

            # Shift에 따라 추가적인 시간대(intervals) 생성
            if shift == "주간":
                start_time = datetime.strptime("08:35", "%H:%M")  # 주간 시작 시간
                additional_time_intervals = generate_intervals(start_time, 144, 5)  # 08:35 ~ 20:30, 5분 간격으로 144개
            elif shift == "야간":
                start_time = datetime.strptime("20:35", "%H:%M")  # 야간 시작 시간
                additional_time_intervals = generate_intervals(start_time, 144, 5)  # 20:35 ~ 08:30, 5분 간격으로 144개
            else:
                additional_time_intervals = []  # 예외 처리: 유효하지 않은 Shift일 경우 비워둠

            # 텍스트 출력 추적을 위한 변수
            processed_time_ranges = []  # 출력된 시간 범위를 저장하는 리스트 [(start_time, end_time)]

            # 추가적인 시간 구간에 따라 Canvas에 사각형 및 텍스트 렌더링
            for i, time_text in enumerate(additional_time_intervals):
                # 현재 시간 슬롯 정의
                time_start = time_text  # 현재 시간 슬롯의 시작
                time_end = additional_time_intervals[i + 1] if i + 1 < len(additional_time_intervals) else \
                    additional_time_intervals[0]  # 마지막 슬롯의 종료시간 설정

                # 시간 범위에 따라 색상과 텍스트 결정
                fill_color = "lightgreen"  # 기본 색상: 연한 녹색(lightgreen)
                display_text = None  # 출력할 텍스트 초기화
                annotation_text = None  # 내용 부분 (새줄로 출력될 텍스트)
                matching_range = None  # 현재 시간 슬롯에 매칭된 범위를 저장

                # 계획 정지 시간 확인
                for row in planned_downtime_rows:
                    planned_start = f"{row[0].get()}:{row[1].get()}"  # 계획 정지 시작 시간
                    planned_end = f"{row[2].get()}:{row[3].get()}"  # 계획 정지 종료 시간
                    annotation = row[5].get()  # 계획 정지 설명

                    # 시작과 종료 시간이 같은 경우 무시
                    if planned_start == planned_end:
                        continue

                    if is_time_in_range(planned_start, planned_end, time_start):  # 수정된 함수 사용
                        fill_color = "yellow"  # 노란색 설정
                        display_text = f"계획 정지 시간({planned_start} ~ {planned_end})"  # 첫 번째 줄 텍스트
                        annotation_text = f"내용: {annotation}"  # 두 번째 줄 텍스트
                        matching_range = (planned_start, planned_end)
                        break

                # 비가동 시간 확인
                for row in downtime_rows:
                    downtime_start = f"{row[0].get()}:{row[1].get()}"  # 비가동 시작 시간
                    downtime_end = f"{row[2].get()}:{row[3].get()}"  # 비가동 종료 시간
                    annotation = row[5].get()  # 비가동 설명

                    # 시작과 종료 시간이 같은 경우 무시
                    if downtime_start == downtime_end:
                        continue

                    if is_time_in_range(downtime_start, downtime_end, time_start):  # 수정된 함수 사용
                        fill_color = "red"  # 빨간색 설정
                        display_text = f"비가동 시간({downtime_start} ~ {downtime_end})"  # 첫 번째 줄 텍스트
                        annotation_text = f"내용: {annotation}"  # 두 번째 줄 텍스트
                        matching_range = (downtime_start, downtime_end)
                        break

                # **사각형(시간 슬롯) 생성**
                # 칸의 좌표와 크기 설정
                x1 = 160  # 사각형의 좌측 상단 X 좌표
                y1 = i * 5.54 + 22  # 사각형의 좌측 상단 Y 좌표 (i에 따라 아래로 배치됨)
                x2 = 80  # 사각형의 우측 하단 X 좌표
                y2 = y1 + 5.760  # 사각형의 우측 하단 Y 좌표 (사각형 높이: 약 5.760)

                # 사각형을 Canvas에 그리기
                canvas.create_rectangle(x1, y1, x2, y2, fill=fill_color, outline="black")

                # **텍스트 출력 처리**
                if matching_range and matching_range not in processed_time_ranges:
                    # 해당 범위가 아직 출력되지 않았다면 텍스트 출력

                    # 첫 번째 줄 텍스트 출력
                    canvas.create_text(
                        x2 + 88,  # 텍스트는 사각형 오른쪽 5픽셀 간격
                        (y1 + y2) / 2 - 0,  # 텍스트는 사각형 중앙에서 8픽셀 위
                        text=display_text,  # 출력할 첫 번째 줄 텍스트
                        font=("Arial", 8),  # 글꼴: Arial, 크기: 8
                        fill="black",  # 글자 색상: 검정색
                        anchor="w",  # 텍스트의 위치 기준점: 왼쪽 정렬(west)
                    )

                    # 두 번째 줄 텍스트 출력
                    canvas.create_text(
                        x2 + 88,  # 텍스트는 사각형 오른쪽 5픽셀 간격
                        (y1 + y2) / 2 + 11,  # 텍스트는 사각형 중앙에서 8픽셀 아래
                        text=annotation_text,  # 출력할 두 번째 줄 텍스트
                        font=("Arial", 8),  # 글꼴: Arial, 크기: 8
                        fill="black",  # 글자 색상: 검정색
                        anchor="w",  # 텍스트의 위치 기준점: 왼쪽 정렬(west)
                    )

                    # 출력된 범위를 저장
                    processed_time_ranges.append(matching_range)

    # **출력 텍스트 생성 (오른쪽)**
    result_text = Text(main_frame, wrap=WORD, width=78, height=28, bg="lightgray", state="disabled")
    result_text.pack(side=LEFT, fill=BOTH, expand=True, padx=0, pady=10)
    right_panel = Frame(main_frame)
    right_panel.pack(side=RIGHT, fill=Y, expand=True)  # 부모 높이에 맞게 확장, 너비는 고정
    graph_frame = Frame(right_panel)
    graph_frame.pack(side=TOP, fill=BOTH, padx=5, pady=(0, 0))
    table_frame = Frame(right_panel)
    table_frame.pack(side=TOP, fill=BOTH, padx=5, pady=(0, 0))
    # x 위치(padx): 왼쪽 410, 오른쪽 600
    # y 위치(pady): 위쪽 10, 아래쪽 12

    # 그리드 확장 설정
    main_frame.rowconfigure(0, weight=1)
    main_frame.columnconfigure(0, weight=1)

    # 태그 정의
    result_text.tag_configure("title", font=("Arial", 8, "bold"))
    result_text.tag_configure("header", font=("Arial", 7, "bold"), foreground="blue")
    result_text.tag_configure("normal", font=("Arial", 7, "bold"))
    result_text.tag_configure("highlight", foreground="red", font=("Arial", 7, "bold"))

    # 제목 출력
    # 📊 그래프 분석 실행
    generate_graphs_embedded(input_date, shift, graph_frame)
    create_sparepart_graph_embedded(graph_frame)

    # 테이블을 위한 엑셀 파일 생성
    table_excel_path = create_excel_file_for_table(
        f"table_{input_date}_{shift}.xlsx",  # 생성될 엑셀 파일 이름
        shift,  # 주간/야간 정보
        f"C:/Ford A+C Test center_생산 분석 프로그램_Rev02/Data/FORD A+C_Data/FORD A+C FCT NG List/{input_date}_{shift}_FCT NG List.txt",
        # FCT NG 파일 경로
        f"C:/Ford A+C Test center_생산 분석 프로그램_Rev02/Data/FORD A+C_Data/FORD A+C LED NG List/{input_date}_{shift}_LED NG List.txt"
        # LED NG 파일 경로
    )

    # 엑셀 파일을 GUI에 임베드된 상태로 출력
    display_excel_embedded(
        table_excel_path,  # 생성된 엑셀 파일 경로
        table_frame  # 출력할 테이블의 프레임
    )
    # 제목 출력
    
# 🔓 텍스트 삽입을 위해 상태를 normal로 설정
    result_text.config(state="normal")
    result_text.insert(END, f"=== {base_date.strftime('%Y.%m.%d')} {shift} 생산 일보(FCT 양품 > LED) ===\n", "title")
    result_text.insert(END, f"작업자: {worker_name} / 생산 품명: FORD A+C\n", "header")
    result_text.insert(END, f"Order No.: {order_no}\n", "normal")
    result_text.insert(END, f"Box 현황: {box_status} / 마스터 샘플 테스트(O/X): {master_sample_status}\n\n", "normal")

    # 생산 시간 정리
    result_text.insert(END, "== 생산 시간 정리 ==\n", "header")
    result_text.insert(END, f"총 계획 정지 시간 : {planned_downtime:.2f} 시간", "normal")
    result_text.insert(END, f" / 총 비가동 시간 : {total_downtime:.2f} 시간\n", "normal")
    result_text.insert(END, f"총 작업 시간 : {total_work_time:.2f} 시간", "normal")
    result_text.insert(END, f" / 실작업 시간 : {working_time:.2f} 시간\n", "normal")
    result_text.insert(END, f"유실 시간 : {loss_time:.2f} 시간\n", "highlight")

    # 품번별 생산 실적
    result_text.insert(END, "\n== 품번별 생산 실적 ==\n", "header")
    ordered_slots = list(time_slots[shift].keys())

    for slot in ordered_slots:
        if slot not in time_slot_counts:
            continue
        start, end = time_slots[shift][slot]
        result_text.insert(END, f"{slot}시간대({start} ~ {end})\n", "header")
        parts = time_slot_counts[slot]

        for part, counts in parts.items():
            result_text.insert(END, f"  품번 {part} / 양품 개수: {counts['양품']} / 불량 개수: {counts['불량']}\n", "normal")
        result_text.insert(END, "\n", "normal")

    # 품번별 요약
    for part, counts in part_summary.items():
        result_text.insert(
            END,
            f"품번 {part} = FCT OK > LED / 양품 개수: {counts['양품']} & 불량 개수: {counts['불량']}\n",
            "header"
        )

    # PASS율(양품률) 계산
    if (total_ok + total_ng) > 0:  # 0으로 나누는 경우를 방지
        pass_rate = (total_ok / (total_ok + total_ng)) * 100
        result_text.insert(
            END,
            f"전체 FCT OK > LED 양품 개수: {total_ok} / FCT OK > LED 불량 개수: {total_ng} / PASS율(양품률): {pass_rate:.2f}%\n",
            "highlight"
        )
    else:
        result_text.insert(
            END,
            f"전체 FCT OK > LED 양품 개수: {total_ok} \nFCT OK > LED 불량 개수: {total_ng}\nPASS율(양품률): 계산 불가 (0으로 나눌 수 없음)\n",
            "highlight"
        )

    # CYCLE_TIME 출력
    result_text.insert(END, "\n== 표준 생산 시간 계산 ==\n", "header")
    for part in part_summary.keys():
        if part in CYCLE_TIMES:
            result_text.insert(END, f"품번 {part} 의 CYCLE_TIME: {CYCLE_TIMES[part]} 초\n", "normal")

    result_text.insert(END, f" 표준 생산시간 합계: {standard_production_time_summary:.2f} 시간\n\n", "highlight")
    result_text.insert(END, f"주요 불량 내용(검사 항목 제외): {defect_details}\n\n", "normal")
    result_text.insert(END, f"생산 건의 내용: {suggestions}\n\n", "normal")

    # OEE 텍스트를 빨간색과 큰 글씨로 출력
    result_text.tag_configure("highlight", foreground="red", font=("Arial", 11, "bold"))  # 태그 정의
    result_text.insert(END, f"OEE: {oee:.2f}%\n", "highlight")  # 태그 적용

    # OEE 값이 95 이상인 경우 추가 메시지 출력
    if oee >= 95:
        result_text.insert(END, "시간 기입 관련 오류가 있습니다, 확인바랍니다.\n", "highlight")  # 동일 태그로 적용
# ✅ Text 위젯 생성 시 Read-only 상태 설정
# 🔒 텍스트 삽입 후 다시 비활성화하여 수정 불가하게 설정
    
    # 🔓 텍스트 삽입 위해 다시 활성화
    result_text.config(state="normal")

    # ✅ FCT 2회 NG 결과 출력 (출력창 하단에)
    from Ford_A_C_FCT_2회_NG_List_backend import run_fct_2nd_ng_analysis

    # 결과를 출력하며 'normal' 태그를 적용
    run_fct_2nd_ng_analysis(
        input_date,
        shift,
        output_callback=lambda msg: result_text.insert("end", msg + "\n", "normal")
    )

    # 출력이 완료되면 Text 위젯 수정 불가로 설정
    
    # ✅ Vision NG 분석 (BA1WJ + 17번째 문자 & 라인 끝 판별, Normal 폰트 적용)
    vision_file_path = fr"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C LED NG List\{input_date}_{shift}_LED NG List.txt"

    vision_ng_1 = {}
    vision_ng_2 = {}

    if os.path.exists(vision_file_path):
        with open(vision_file_path, "r", encoding="utf-8") as vf:
            for line in vf:
                # 시간대별 요약 이전까지만 읽기
                if "======== 시간대별 & LED별 조건별 요약 ========" in line:
                    break

                # BA1WJ 위치 찾기
                pos = line.find("BA1WJ")
                if pos == -1:
                    continue

                # BA1WJ + YYJJJSSSSSS + 아무 문자 1개 → 그 다음 문자
                target_index = pos + 17
                if target_index >= len(line):
                    continue

                key_char = line[target_index]

                mapping = {
                    "C": "35643009",
                    "J": "35915729",
                    "1": "35654264",
                    "P": "35643010",
                    "N": "35749091",
                    "S": "35915730"
                }
                part_no = mapping.get(key_char)
                if not part_no:
                    continue

                # Vision NG 유형 판별 (라인 끝 확인)
                if line.strip().endswith("Vision 2회 발생"):
                    vision_ng_2[part_no] = vision_ng_2.get(part_no, 0) + 1
                elif line.strip().endswith("Vision NG"):
                    vision_ng_1[part_no] = vision_ng_1.get(part_no, 0) + 1

    # 출력 (Normal 폰트 적용)
    result_text.insert(END, "\n[Vision 1회 NG]:\n", "normal")
    for part_no, count in sorted(vision_ng_1.items()):
        result_text.insert(END, f" - {part_no} / {count}개\n", "normal")

    result_text.insert(END, "\n[Vision 2회 NG]:\n", "normal")
    for part_no, count in sorted(vision_ng_2.items()):
        result_text.insert(END, f" - {part_no} / {count}개\n", "normal")

    result_text.config(state="disabled")

    # 저장 버튼 변경
    save_button = Button(result_window, text="저  장(Save)", command=lambda: save_as_pdf(result_window))
    save_button.place(x=650, y=865)  # x=수평 위치, y=수직 위치


# 비밀번호 상수 설정
SPAREPARTS_PASSWORD = "test1234"


def open_spareparts_settings():
    # 비밀번호 확인
    password = simpledialog.askstring("비밀번호 입력", "비밀번호를 입력하세요:", show="*")
    if password != SPAREPARTS_PASSWORD:
        messagebox.showerror("오류", "비밀번호가 올바르지 않습니다.")
        return

    # 새로운 창 생성
    spareparts_window = tk.Toplevel(root)
    spareparts_window.title("Spareparts 설정")
    spareparts_window.geometry("400x400")

    entries = {}
    parts = ["Mini B", "USB-C", "USB-A", "Power"]

    # 각 부품의 현재 재고 및 안전 수량 입력 필드 추가
    row = 0
    for part in parts:
        tk.Label(spareparts_window, text=f"{part} 현재 재고:").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        entries[f"{part}_current"] = tk.Entry(spareparts_window)
        entries[f"{part}_current"].grid(row=row, column=1, padx=10, pady=5)
        row += 1

        tk.Label(spareparts_window, text=f"{part} 안전 수량:").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        entries[f"{part}_safe"] = tk.Entry(spareparts_window)
        entries[f"{part}_safe"].grid(row=row, column=1, padx=10, pady=5)
        row += 1

    def save_spareparts():
        # 데이터 저장 경로 정의
        save_dir = "C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C_Spareparts\\"
        os.makedirs(save_dir, exist_ok=True)  # 경로가 없다면 생성

        # 현재 날짜로 파일 이름 생성
        filename = datetime.now().strftime("%Y.%m.%d_sparepart list.txt")
        file_path = os.path.join(save_dir, filename)

        # 데이터 저장
        try:
            with open(file_path, "w", encoding="utf-8") as file:
                file.write("Spareparts 설정 정보\n")
                file.write(f"저장 날짜: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                file.write("-" * 30 + "\n")
                for part in parts:
                    current = entries[f"{part}_current"].get()
                    safe = entries[f"{part}_safe"].get()
                    file.write(f"{part} 현재 재고: {current}\n")
                    file.write(f"{part} 안전 수량: {safe}\n")
                    file.write("-" * 30 + "\n")
            messagebox.showinfo("저장 완료", f"파일이 저장되었습니다: {file_path}")
        except Exception as e:
            messagebox.showerror("오류", f"파일 저장 중 오류가 발생했습니다:\n{str(e)}")

    # 저장 버튼 추가
    save_button = tk.Button(spareparts_window, text="저장", command=save_spareparts)
    save_button.grid(row=row, column=0, columnspan=2, pady=10)


# **GUI 구성**
root = Tk()
root.title("생산 분석 프로그램      Copyright 2025. JW All rights reserved.")
root.geometry("500x900")  # 창 크기를 고정 (너비x높이)

# 설정 메뉴 접근 버튼
settings_button = tk.Button(root, text="설정(비번 필요)", command=open_settings)
settings_button.place(x=350, y=10)  # 버튼을 (X=320, Y=10) 위치로 배치

# Spareparts 버튼 추가
spareparts_button = tk.Button(root, text="Spareparts 설정(비번 필요)", command=open_spareparts_settings)
spareparts_button.pack(padx=0, pady=10)


Label(root, text="기준 날짜 선택").pack()
date_entry = DateEntry(root, background='darkblue', foreground='white', borderwidth=2, locale='ko_KR')
date_entry.pack()

Label(root, text="Shift 선택").pack()
shift_combobox = Combobox(root, values=["주간", "야간"], state="readonly", width=10)
shift_combobox.pack()

Label(root, text="작업자 이름").pack()
worker_name_entry = Entry(root)
worker_name_entry.pack()

# Order No. 필드 조정
Label(root, text="Order No.( / 키 이용할 것)").pack()
order_no_entry = Text(root, width=40, height=1)  # 너비를 65으로 설정
order_no_entry.pack()

from tkinter import Tk, Label, Entry, Text

# 한 줄로 수평 배치
row = Frame(root)
row.pack(fill=X, pady=10)  # 한 줄 위젯을 담을 프레임

# Box 현황
Label(row, text="Box 현황").pack(side=LEFT, padx=(100, 10))  # Box 현황 레이블 (오른쪽 간격 10 추가)
box_status_entry = Entry(row, width=5)  # 고정된 입력 필드 폭
box_status_entry.pack(side=LEFT, padx=(0, 10))  # 여백 추가 (오른쪽 20)

# 마스터 샘플 테스트
Label(row, text="마스터 샘플 테스트(O/X)").pack(side=LEFT, padx=(5, 10))  # 마스터 샘플 테스트 레이블
master_sample_entry = Entry(row, width=5)
master_sample_entry.pack(side=LEFT, padx=5)  # 기본 간격

# 수평 배치용 Frame 생성
Label(root, text="- Sparepart 사용량 -").pack()
spare_parts_frame = Frame(root)
spare_parts_frame.pack(pady=5)

# 스페어파트 이름 리스트
spare_parts_labels = [
    "FCT1 Mini B", "FCT1 USB-C", "FCT1 USB-A", "FCT1 Power",
    "FCT2 Mini B", "FCT2 USB-C", "FCT2 USB-A", "FCT2 Power",
    "FCT3 Mini B", "FCT3 USB-C", "FCT3 USB-A", "FCT3 Power",
    "FCT4 Mini B", "FCT4 USB-C", "FCT4 USB-A", "FCT4 Power"
]
spare_parts_entries = {}

# 라벨과 Entry들을 행(row) 단위로 배치
row_frame = None
for idx, part in enumerate(spare_parts_labels):
    if idx % 4 == 0:  # 4개의 Label과 Entry를 한 행에 배치
        row_frame = Frame(spare_parts_frame)
        row_frame.pack(pady=1)  # 행 사이 간격 설정

    label = Label(row_frame, text=part, width=10, anchor='w')  # 라벨 생성
    label.pack(side="left", padx=5)  # 수평 배치

    entry = Entry(row_frame, width=2)  # Entry 생성
    entry.pack(side="left", padx=1)  # 수평 배치

    spare_parts_entries[part] = entry


# 주요 불량 내용 필드 추가 및 조정
Label(root, text="특이사항(Enter키 대신 '/' 키 이용할 것)").pack()
defect_entry = Text(root, width=65, height=2)  # 너비를 65, 높이를 5 줄로 설정
defect_entry.pack()

Label(root, text="생산 건의 내용( Enter키 대신 '/' 키 이용할 것)").pack()
suggestion_entry = Text(root, width=65, height=2)  # 너비를 65, 높이를 5 줄로 설정
suggestion_entry.pack()

Label(root, text="계획 정지 시간 (Film내 노란색)").pack()
planned_downtime_frame = Frame(root)
planned_downtime_frame.pack()
planned_downtime_rows = []
add_time_row(planned_downtime_frame, planned_downtime_rows, initial=True)

Label(root, text="비가동 시간 (Film내 빨간색)").pack()
downtime_frame = Frame(root)
downtime_frame.pack()
downtime_rows = []
add_time_row(downtime_frame, downtime_rows, initial=True)

Button(root, text="분석 실행", command=run_analysis).pack()
root.mainloop()