
from datetime import datetime, timedelta
import os

# 데이터 경로 설정
BASE_PATH = r"C:\Users\user\Desktop\FORD A+C VISION 로그파일"
OUTPUT_DIRS = [r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C LED NG List"]

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

def parse_file_time(file_name):
    if len(file_name) < 46:
        return None
    try:
        identifier = file_name[17]
        date_time_str = file_name[31:45] if identifier in ['C', 'J', '1'] else file_name[32:46]
        return datetime.strptime(date_time_str, '%Y%m%d%H%M%S')
    except ValueError:
        return None

def classify_time_period(file_datetime, 기준날짜, shift):
    for label, start, end in TIME_SLOTS[shift]:
        if 기준날짜 + start <= file_datetime <= 기준날짜 + end:
            return label, (기준날짜 + start).strftime("%H:%M:%S"), (기준날짜 + end).strftime("%H:%M:%S")
    return None, None, None

def analyze_led_ng_files(input_date, shift):
    기준날짜 = datetime.strptime(input_date, '%Y%m%d')
    results = {}
    repeated_ng_files = set()
    file_registry = {}

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

                unique_id = file_name[:31]
                file_datetime = parse_file_time(file_name)
                if not file_datetime:
                    continue

                if unique_id in file_registry:
                    previous_file_name = file_registry[unique_id]
                    if file_datetime > parse_file_time(previous_file_name):
                        repeated_ng_files.add(unique_id)
                file_registry[unique_id] = file_name

                time_label, _, _ = classify_time_period(file_datetime, 기준날짜, shift)
                if time_label:
                    display_name = f"{time_label}_{file_name}_"
                    display_name += "Vision 2회 발생" if unique_id in repeated_ng_files else "Vision NG"
                    results[time_label].append(display_name)

    return results

# 분석 결과 저장 함수 수정
def save_led_ng_results(input_date, shift, results):
    output_file_name = f"{input_date}_{shift}_LED NG List.txt"

    for output_dir in OUTPUT_DIRS:
        output_file_path = os.path.join(output_dir, output_file_name)

        os.makedirs(output_dir, exist_ok=True)

        with open(output_file_path, "w", encoding="utf-8") as file:
            condition_counts = {label: {"Vision NG": 0, "Vision 2회 NG": 0} for label in results.keys()}

            for time_label in results.keys():
                for file_name in results[time_label]:
                    file.write(f"{file_name}\n")
                    if "Vision 2회 발생" in file_name:
                        condition_counts[time_label]["Vision 2회 NG"] += 1
                    else:
                        condition_counts[time_label]["Vision NG"] += 1

            file.write("\n======== 시간대별 & LED별 조건별 요약 ========\n")
            for time_label, counts in condition_counts.items():
                if counts["Vision NG"] > 0:
                    file.write(f"{time_label} & Vision NG : {counts['Vision NG']}개\n")
                if counts["Vision 2회 NG"] > 0:
                    file.write(f"{time_label} & Vision 2회 NG : {counts['Vision 2회 NG']}개\n")

        print(f"[LED_NG_Backend] 결과 파일 저장 완료: {output_file_path}")

def run_led_ng_analysis(input_date, shift):
    print(f"[LED_NG_Backend] 분석 시작: {input_date}, {shift}")
    results = analyze_led_ng_files(input_date, shift)
    save_led_ng_results(input_date, shift, results)
    print("[LED_NG_Backend] 분석 및 저장 완료")

if __name__ == "__main__":
    print("이 모듈은 메인 프로그램에서 import 하여 사용하는 LED NG 백엔드 전용 모듈입니다.")
