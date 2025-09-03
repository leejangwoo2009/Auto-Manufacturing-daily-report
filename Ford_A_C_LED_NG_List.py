import os
from datetime import datetime, timedelta
from tkinter import *
from tkinter.ttk import Combobox
from tkcalendar import DateEntry
import locale

# Locale 설정
try:
    locale.setlocale(locale.LC_TIME, 'ko_KR.UTF-8')
except locale.Error:
    locale.setlocale(locale.LC_TIME, 'ko_KR')

# 데이터 경로 설정
BASE_PATH = r"C:\Users\user\Desktop\FORD A+C VISION 로그파일"
OUTPUT_DIRS = [r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C FCT NG List"]

# 시간대 정의
TIME_SLOTS = {
    "주간": [
        ("A 시간대", timedelta(hours=8, minutes=30), timedelta(hours=10, minutes=29, seconds=59)),
        ("B 시간대", timedelta(hours=10, minutes=30), timedelta(hours=12, minutes=29, seconds=59)),
        ("C 시간대", timedelta(hours=12, minutes=30), timedelta(hours=14, minutes=29, seconds=59)),
        ("D 시간대", timedelta(hours=14, minutes=30), timedelta(hours=16, minutes=29, seconds=59)),
        ("E 시간대", timedelta(hours=16, minutes=30), timedelta(hours=18, minutes=29, seconds=59)),
        ("F 시간대", timedelta(hours=18, minutes=30), timedelta(hours=20, minutes=29, seconds=59)),
    ],
    "야간": [
        ("A' 시간대", timedelta(hours=20, minutes=30), timedelta(hours=22, minutes=29, seconds=59)),
        ("B' 시간대", timedelta(hours=22, minutes=30), timedelta(days=1, seconds=-1)),
        ("C' 시간대", timedelta(days=1), timedelta(days=1, hours=2, minutes=29, seconds=59)),
        ("D' 시간대", timedelta(days=1, hours=2, minutes=30), timedelta(days=1, hours=4, minutes=29, seconds=59)),
        ("E' 시간대", timedelta(days=1, hours=4, minutes=30), timedelta(days=1, hours=6, minutes=29, seconds=59)),
        ("F' 시간대", timedelta(days=1, hours=6, minutes=30), timedelta(days=1, hours=8, minutes=29, seconds=59)),
    ],
}

# 파일에서 시간 추출
def parse_file_time(file_name):
    """파일 이름에서 시간을 추출"""
    if len(file_name) < 46:
        return None
    try:
        identifier = file_name[17]
        date_time_str = file_name[31:45] if identifier in ['C', 'J', '1'] else file_name[32:46]
        return datetime.strptime(date_time_str, '%Y%m%d%H%M%S')
    except ValueError:
        return None

# 파일 시간대 분류
def classify_time_period(file_datetime, 기준날짜, shift):
    """시간대를 기준으로 파일을 분류"""
    for label, start, end in TIME_SLOTS[shift]:
        if 기준날짜 + start <= file_datetime <= 기준날짜 + end:
            return label, (기준날짜 + start).strftime("%H:%M:%S"), (기준날짜 + end).strftime("%H:%M:%S")
    return None, None, None

# NG 데이터 분석
def analyze_ng_files(input_date, shift):
    """선택된 날짜와 교대를 기준으로 NG 데이터를 분석"""
    기준날짜 = datetime.strptime(input_date, '%Y%m%d')
    results = {}
    repeated_ng_files = set()
    file_registry = {}

    # 시간대 저장소 초기화
    for label, _, _ in TIME_SLOTS[shift]:
        results[label] = []

    for base_path in [BASE_PATH]:
        day_folder = os.path.join(base_path, input_date, "GoodFile")
        next_day_folder = os.path.join(base_path, (기준날짜 + timedelta(days=1)).strftime('%Y%m%d'), "GoodFile")
        folders_to_check = [day_folder] if shift == "주간" else [day_folder, next_day_folder]

        for folder in folders_to_check:
            if not os.path.isdir(folder):
                continue
            for file_name in os.listdir(folder):
                if not file_name.endswith("F.txt"):
                    continue

                # 유일 ID 추출
                unique_id = file_name[:31]
                file_datetime = parse_file_time(file_name)
                if not file_datetime:
                    continue

                # 2회 발생 여부 처리
                if unique_id in file_registry:
                    previous_file_name = file_registry[unique_id]
                    if file_datetime > parse_file_time(previous_file_name):
                        repeated_ng_files.add(unique_id)
                file_registry[unique_id] = file_name

                # 시간대 확인
                time_label, _, _ = classify_time_period(file_datetime, 기준날짜, shift)
                if time_label:
                    # 파일 형태 처리
                    display_name = f"{time_label}_{file_name}_"
                    display_name += "Vision 2회 발생" if unique_id in repeated_ng_files else "Vision NG"
                    results[time_label].append(display_name)

    return results

# 결과 파일로 저장
def save_results_to_file(input_date, shift, results):
    """분석 결과를 텍스트 파일로 저장"""
    output_file_name = f"{input_date}_{shift}_LED NG List.txt"
    output_file_path = os.path.join(OUTPUT_DIR, output_file_name)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(output_file_path, "w", encoding="utf-8") as file:
        # 각 시간대별 상세 저장소
        condition_counts = {label: {"Vision NG": 0, "Vision 2회 NG": 0} for label in results.keys()}

        for time_label in results.keys():
            for file_name in results[time_label]:
                file.write(f"{file_name}\n")
                if "Vision 2회 발생" in file_name:
                    condition_counts[time_label]["Vision 2회 NG"] += 1
                else:
                    condition_counts[time_label]["Vision NG"] += 1

        # 요약 값 출력
        file.write("\n======== 시간대별 & LED별 조건별 요약 ========\n")
        for time_label, counts in condition_counts.items():
            if counts["Vision NG"] > 0:
                file.write(f"{time_label} & Vision NG : {counts['Vision NG']}개\n")
            if counts["Vision 2회 NG"] > 0:
                file.write(f"{time_label} & Vision 2회 NG : {counts['Vision 2회 NG']}개\n")

    print(f"결과 파일 저장 완료: {output_file_path}")

# GUI를 통해 분석 실행
def run_analysis():
    """NG 분석 실행 및 결과 GUI에 표시"""
    input_date = date_entry.get_date().strftime('%Y%m%d')
    shift = shift_combobox.get()
    if shift not in ["주간", "야간"]:
        return

    results = analyze_ng_files(input_date, shift)
    save_results_to_file(input_date, shift, results)

# GUI 설정
root = Tk()
root.title("LED NG 분석 도구")

Label(root, text="기준 날짜 선택:").pack()
date_entry = DateEntry(root, width=12, year=datetime.now().year, month=datetime.now().month, day=datetime.now().day)
date_entry.pack()

Label(root, text="Shift 선택:").pack()
shift_combobox = Combobox(root, values=["주간", "야간"], state="readonly", width=10)
shift_combobox.pack()

Button(root, text="분석 실행", command=run_analysis).pack()

root.mainloop()
