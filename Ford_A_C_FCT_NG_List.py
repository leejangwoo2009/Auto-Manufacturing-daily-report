import os
from datetime import datetime, timedelta
from tkinter import *
from tkinter.ttk import Combobox
from tkcalendar import DateEntry
import locale

# **Locale 설정 (달력 표기를 한국어로 표시)**
try:
    locale.setlocale(locale.LC_TIME, 'ko_KR.UTF-8')
except locale.Error:
    try:
        # Windows용 한국어 로케일
        locale.setlocale(locale.LC_TIME, 'Korean')
    except locale.Error:
        locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')

# 경로 설정
BASE_PATHS = [
    r"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC6",  # FCT1
    r"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC7",  # FCT2
    r"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC8",  # FCT3
    r"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC9"   # FCT4
]
OUTPUT_DIRS = [r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C FCT NG List"]

# 파일명에서 시간 정보를 파싱
def parse_file_time(file_name):
    """
    파일명에서 시간을 파싱합니다. 조건에 따라 날짜와 시간을 식별합니다.
    """
    if len(file_name) < 46:  # 파일명이 최소 길이를 만족해야 유효
        return None
    try:
        # 17번째 인덱스 식별자 가져오기
        identifier = file_name[17]

        # 식별자 조건에 따라 날짜+시간 추출 범위 결정
        if identifier in ["C", "J", "1"]:  # C, J, 1의 경우
            date_time_str = file_name[31:45]  # 인덱스 31~44 (YYYYMMDDHHMMSS)
        elif identifier in ["P", "N", "S"]:  # P, N, S의 경우
            date_time_str = file_name[32:46]  # 인덱스 32~45 (YYYYMMDDHHMMSS)
        else:
            return None  # 조건에 맞지 않는 경우 None 반환

        # 추출된 문자열을 datetime 객체로 변환
        return datetime.strptime(date_time_str, "%Y%m%d%H%M%S")
    except ValueError:
        return None  # 형식 오류가 발생하면 None 반환


# 시간대 분류 함수 (정확한 A~F, A'~F' 시간대 적용)
def classify_time_period(file_datetime, 기준날짜, shift):
    time_slots = {
        "A 시간대": (timedelta(hours=8, minutes=30), timedelta(hours=10, minutes=29, seconds=59)),
        "B 시간대": (timedelta(hours=10, minutes=30), timedelta(hours=12, minutes=29, seconds=59)),
        "C 시간대": (timedelta(hours=12, minutes=30), timedelta(hours=14, minutes=29, seconds=59)),
        "D 시간대": (timedelta(hours=14, minutes=30), timedelta(hours=16, minutes=29, seconds=59)),
        "E 시간대": (timedelta(hours=16, minutes=30), timedelta(hours=18, minutes=29, seconds=59)),
        "F 시간대": (timedelta(hours=18, minutes=30), timedelta(hours=20, minutes=29, seconds=59)),
        "A' 시간대": (timedelta(hours=20, minutes=30), timedelta(hours=22, minutes=29, seconds=59)),
        "B' 시간대": (timedelta(hours=22, minutes=30), timedelta(days=1, hours=0, minutes=29, seconds=59)),
        "C' 시간대": (timedelta(days=1, hours=0, minutes=30), timedelta(days=1, hours=2, minutes=29, seconds=59)),
        "D' 시간대": (timedelta(days=1, hours=2, minutes=30), timedelta(days=1, hours=4, minutes=29, seconds=59)),
        "E' 시간대": (timedelta(days=1, hours=4, minutes=30), timedelta(days=1, hours=6, minutes=29, seconds=59)),
        "F' 시간대": (timedelta(days=1, hours=6, minutes=30), timedelta(days=1, hours=8, minutes=29, seconds=59)),
    }

    for label, (start_delta, end_delta) in time_slots.items():
        if ("'" in label and shift == "야간") or ("'" not in label and shift == "주간"):
            start = 기준날짜 + start_delta
            end = 기준날짜 + end_delta
            if start <= file_datetime <= end:
                return label
    return None

# NG 파일 분석 - 조건 2, 3 메시지 포함
def process_ng_file_content(file_path, file_name, repeated_ng_latest_files):
    try:
        messages = []
        if os.path.exists(file_path):
            with open(file_path, "r", encoding="utf-8") as f:
                lines = f.readlines()

            # 조건 1: NG 2회 발생 확인
            if file_name in repeated_ng_latest_files:
                messages.append("FCT 2회 NG")

            # 조건 2와 조건 3: 식별자로 구분
            identifier = file_name[17]
            diagnostics = {}
            if identifier in ["C", "1", "P", "N"]:
                diagnostics.update({
                    1.00: "제품 Mini B 문제",
                    1.01: "제품 USB-A 문제",
                    1.02: "제품 USB-C 문제",
                    1.03: "제품 Power Pin 문제",
                    1.04: "제품 Power Pin 문제",
                    1.05: "제품 Power Pin 문제",
                    1.06: "테스터기 Power 관련 문제",
                    1.07: "제품 Q소자 문제 가능성 많음",
                    1.08: "SW 설치 Or Tester기 parts 문제",
                    1.09: "SW 설치 Or Tester기 parts 문제",
                    1.10: "USB-A 관련 문제",
                    1.11: "USB-A 관련 문제",
                    1.12: "USB-C 관련 문제",
                    1.13: "USB-C 관련 문제",
                    1.14: "USB-C 관련 문제",
                    1.15: "USB-A 관련 문제",
                    1.16: "USB-C 관련 문제",
                    1.17: "USB-C 관련 문제",
                    1.18: "USB-A 관련 문제",
                    1.19: "USB-A 관련 문제",
                    1.20: "USB-C 관련 문제",
                    1.21: "USB-C 관련 문제",
                    1.22: "Mini B 관련 문제 or Carplay",
                    1.23: "Mini B 관련 문제 or Carplay",
                    1.24: "USB-C 관련 문제",
                    1.25: "USB-C 관련 문제",
                    1.26: "USB-C 관련 문제",
                    1.27: "USB-C 관련 문제",
                    1.28: "USB-A 관련 문제",
                    1.29: "USB-A 관련 문제",
                    1.30: "USB-A 관련 문제",
                    1.31: "USB-A 관련 문제",
                    1.32: "회로 문제(암전류)"
                })

                # 조건 3-식별자 ['J', 'S'] (다른 로직 추가)
            elif identifier in ['J', 'S']:
                diagnostics.update({
                    1.00: "Reflash NG",
                    1.01: "테스터기 Power 관련 문제",
                    1.02: "USB-C 관련 문제",
                    1.03: "USB-C 관련 문제",
                    1.04: "USB-C 관련 문제",
                    1.05: "USB-C 관련 문제",
                    1.06: "USB-A 관련 문제",
                    1.07: "USB-A 관련 문제",
                    1.08: "USB-A 관련 문제",
                    1.09: "USB-A 관련 문제",
                    1.10: "SW 설치 Or Tester기 parts 문제",
                    1.11: "SW 설치 Or Tester기 parts 문제",
                    1.12: "Mini B 관련 문제 or Carplay",
                    1.13: "Mini B 관련 문제 or Carplay",
                    1.14: "제품 충전 프로파일 문제",
                    1.15: "제품 Power Pin 문제",
                    1.16: "제품 Power Pin 문제",
                    1.17: "제품 Power Pin 문제",
                    1.18: "테스터기 Power 관련 문제",
                    1.19: "제품 Q소자 문제 가능성 많음",
                    1.20: "USB-C 관련 문제",
                    1.21: "USB-A 관련 문제",
                    1.22: "USB-A 관련 문제",
                    1.23: "USB-A 관련 문제",
                    1.24: "USB-C 관련 문제",
                    1.25: "USB-C 관련 문제",
                    1.26: "USB-C 관련 문제",
                    1.27: "USB-A 관련 문제",
                    1.28: "USB-C 관련 문제",
                    1.29: "USB-C 관련 문제",
                    1.30: "USB-A 관련 문제",
                    1.31: "USB-A 관련 문제",
                    1.32: "USB-C 관련 문제",
                    1.33: "USB-C 관련 문제",
                    1.34: "USB-C 관련 문제",
                    1.35: "USB-C 관련 문제",
                    1.36: "회로 문제(암전류)"
                })

            # FAIL 키워드가 있는 데이터를 검사
            check_lines = lines[18:]
            for line in check_lines:
                if "FAIL" in line:
                    try:
                        value = float(line[:4].strip("_ "))
                        if value in diagnostics:
                            messages.append(diagnostics[value])
                        else:
                            messages.append(f"미확인 데이터: {value}")
                    except ValueError:
                        messages.append(f"잘못된 데이터 형식: {line.strip()}")

        else:
            messages.append("파일 없음")
        return messages
    except Exception as e:
        return [f"파일 분석 중 오류: {e}"]


# 분석 결과 출력 파일 저장 (조건별 세부 개수 그대로 추가)
def save_results_to_file(input_date, shift, results):
    try:
        # 시간대 및 FCT, 조건별 데이터를 저장할 딕셔너리 초기화
        time_order = [
            "A 시간대", "B 시간대", "C 시간대", "D 시간대", "E 시간대", "F 시간대",
            "A' 시간대", "B' 시간대", "C' 시간대", "D' 시간대", "E' 시간대", "F' 시간대"
        ]
        fcts = ["FCT1", "FCT2", "FCT3", "FCT4"]

        condition_counts = {
            f"{time}_{fct}_{condition}": 0
            for time in time_order for fct in fcts
            for condition in [
                "FCT 2회 NG", "USB-C 관련 문제", "USB-A 관련 문제", "제품 USB-A 문제", "제품 Q소자 문제 가능성 많음", "테스터기 Power 관련 문제",
                "제품 Mini B 문제", "제품 USB-C 문제", "제품 Power Pin 문제", "Mini B 관련 문제 or Carplay",
                "제품 충전 프로파일 문제", "SW 설치 Or Tester기 parts 문제", "회로 문제(암전류)", "Reflash NG"
            ]
        }

        # 모든 경로에 대해 저장 작업 수행
        for OUTPUT_DIR in OUTPUT_DIRS:
            try:
                # 경로가 없으면 생성
                os.makedirs(OUTPUT_DIR, exist_ok=True)
                output_file_path = os.path.join(OUTPUT_DIR, f"{input_date}_{shift}_FCT NG List.txt")

                with open(output_file_path, "w", encoding="utf-8") as output_file:
                    # 각 시간대와 FCT별로 파일 처리
                    for time_label in time_order:
                        for fct_result in results:
                            fct_label = fct_result["fct"]

                            if time_label in fct_result["time_buckets"]:
                                files = fct_result["time_buckets"][time_label]["files"]
                                for file_name in files:
                                    # 파일 세부 내용 기록
                                    output_file.write(f"{file_name}\n")

                                    # 조건별 카운트 증가 (FCT 2회 NG 포함 시 다른 조건 제외)
                                    if "FCT 2회 NG" in file_name:
                                        key = f"{time_label}_{fct_label}_FCT 2회 NG"
                                        condition_counts[key] += 1
                                    else:
                                        if "USB-C 관련 문제" in file_name:
                                            key = f"{time_label}_{fct_label}_USB-C 관련 문제"
                                            condition_counts[key] += 1
                                        if "USB-A 관련 문제" in file_name:
                                            key = f"{time_label}_{fct_label}_USB-A 관련 문제"
                                            condition_counts[key] += 1
                                        if "제품 USB-A 문제" in file_name:
                                            key = f"{time_label}_{fct_label}_제품 USB-A 문제"
                                            condition_counts[key] += 1
                                        if "제품 Q소자 문제 가능성 많음" in file_name:
                                            key = f"{time_label}_{fct_label}_제품 Q소자 문제 가능성 많음"
                                            condition_counts[key] += 1
                                        if "테스터기 Power 관련 문제" in file_name:
                                            key = f"{time_label}_{fct_label}_테스터기 Power 관련 문제"
                                            condition_counts[key] += 1
                                        if "제품 Mini B 문제" in file_name:
                                            key = f"{time_label}_{fct_label}_제품 Mini B 문제"
                                            condition_counts[key] += 1
                                        if "제품 USB-C 문제" in file_name:
                                            key = f"{time_label}_{fct_label}_제품 USB-C 문제"
                                            condition_counts[key] += 1
                                        if "제품 Power Pin 문제" in file_name:
                                            key = f"{time_label}_{fct_label}_제품 Power Pin 문제"
                                            condition_counts[key] += 1
                                        if "Mini B 관련 문제 or Carplay" in file_name:
                                            key = f"{time_label}_{fct_label}_Mini B 관련 문제 or Carplay"
                                            condition_counts[key] += 1
                                        if "제품 충전 프로파일 문제" in file_name:
                                            key = f"{time_label}_{fct_label}_제품 충전 프로파일 문제"
                                            condition_counts[key] += 1
                                        if "SW 설치 Or Tester기 parts 문제" in file_name:
                                            key = f"{time_label}_{fct_label}_SW 설치 Or Tester기 parts 문제"
                                            condition_counts[key] += 1
                                        if "회로 문제(암전류)" in file_name:
                                            key = f"{time_label}_{fct_label}_회로 문제(암전류)"
                                            condition_counts[key] += 1
                                        if "Reflash NG" in file_name:
                                            key = f"{time_label}_{fct_label}_Reflash NG"
                                            condition_counts[key] += 1

                    # 요약 보고서 작성
                    output_file.write("\n======== 시간대별 & FCT별 조건별 요약 ========\n")
                    for time_label in time_order:
                        for fct_label in fcts:
                            for condition in [
                                "FCT 2회 NG", "USB-C 관련 문제", "USB-A 관련 문제", "제품 USB-A 문제", "제품 Q소자 문제 가능성 많음",
                                "테스터기 Power 관련 문제", "제품 Mini B 문제", "제품 USB-C 문제", "제품 Power Pin 문제", "Mini B 관련 문제 or Carplay",
                                "제품 충전 프로파일 문제", "SW 설치 Or Tester기 parts 문제", "회로 문제(암전류)", "Reflash NG"
                            ]:
                                key = f"{time_label}_{fct_label}_{condition}"
                                count = condition_counts[key]
                                if count > 0:
                                    output_file.write(f"{time_label} & {fct_label} & {condition} : {count}개\n")

                print(f"결과 파일 저장 완료: {output_file_path}")
            except Exception as e:
                print(f"경로 {OUTPUT_DIR}에 파일 저장 중 오류: {e}")
    except Exception as e:
        print(f"결과 파일 저장 중 오류 발생: {e}")


# NG 파일 분석 (조건별 카운트 수집 로직 추가)
def analyze_ng_files(input_date, shift):
    기준날짜 = datetime.strptime(input_date, "%Y%m%d")
    results = []

    # 🔁 모든 FCT 폴더에서 제품별 NG 이력 수집
    from collections import defaultdict
    file_registry_all = defaultdict(list)  # key: 제품번호(1~18자리), value: (파일명, datetime)

    for base_path in BASE_PATHS:
        for day in [기준날짜, 기준날짜 + timedelta(days=1)]:
            folder = os.path.join(base_path, day.strftime("%Y%m%d"), "GoodFile")
            if not os.path.isdir(folder):
                continue
            for file_name in os.listdir(folder):
                if not file_name.endswith("F.txt"):
                    continue
                product_id = file_name[:18]  # ✅ 1~18자리만 사용
                file_dt = parse_file_time(file_name)
                if file_dt:
                    file_registry_all[product_id].append((file_name, file_dt))

    # ✅ 2개 이상의 NG 발생 제품 중 최신 파일만 수집
    repeated_ng_latest_files = set()
    for product_id, file_list in file_registry_all.items():
        if len(file_list) >= 2:
            latest_file = max(file_list, key=lambda x: x[1])[0]
            repeated_ng_latest_files.add(latest_file)

    # 각 FCT별로 분석
    for idx, base_path in enumerate(BASE_PATHS, 1):
        fct_label = f"FCT{idx}"
        day_folder = os.path.join(base_path, input_date, "GoodFile")
        next_day_folder = os.path.join(base_path, (기준날짜 + timedelta(days=1)).strftime("%Y%m%d"), "GoodFile")
        folders_to_check = [day_folder] if shift == "주간" else [day_folder, next_day_folder]

        time_buckets = {}

        for folder in folders_to_check:
            if not os.path.isdir(folder):
                continue

            for file_name in os.listdir(folder):
                if not file_name.endswith("F.txt"):
                    continue

                file_path = os.path.join(folder, file_name)
                file_datetime = parse_file_time(file_name)
                if not file_datetime:
                    continue

                time_label = classify_time_period(file_datetime, 기준날짜, shift)
                if time_label:
                    if time_label not in time_buckets:
                        time_buckets[time_label] = {"files": []}
                    messages = process_ng_file_content(file_path, file_name, repeated_ng_latest_files)

                    # 파일 명세 생성
                    serial_number = file_name[31:]
                    file_description = f"{time_label}_{fct_label}_{file_name}_{serial_number}_{', '.join(messages)}"
                    time_buckets[time_label]["files"].append(file_description)

        results.append({"fct": fct_label, "time_buckets": time_buckets})

    save_results_to_file(input_date, shift, results)
    return results


    # ✅ 2개 이상 NG 발생한 제품 중 가장 최신 파일만 수집
    repeated_ng_latest_files = set()
    for product_id, file_list in file_registry_all.items():
        if len(file_list) >= 2:
            latest_file = max(file_list, key=lambda x: x[1])[0]
            repeated_ng_latest_files.add(latest_file)

    # 각 FCT별로 분석
    for idx, base_path in enumerate(BASE_PATHS, 1):
        fct_label = f"FCT{idx}"
        day_folder = os.path.join(base_path, input_date, "GoodFile")
        next_day_folder = os.path.join(base_path, (기준날짜 + timedelta(days=1)).strftime("%Y%m%d"), "GoodFile")
        folders_to_check = [day_folder] if shift == "주간" else [day_folder, next_day_folder]

        time_buckets = {}

        for folder in folders_to_check:
            if not os.path.isdir(folder):
                continue

            for file_name in os.listdir(folder):
                if not file_name.endswith("F.txt"):
                    continue

                file_path = os.path.join(folder, file_name)
                file_datetime = parse_file_time(file_name)
                if not file_datetime:
                    continue

                time_label = classify_time_period(file_datetime, 기준날짜, shift)
                if time_label:
                    if time_label not in time_buckets:
                        time_buckets[time_label] = {"files": []}
                    messages = process_ng_file_content(file_path, file_name, repeated_ng_latest_files)

                    # 파일명 생성 규칙
                    serial_number = file_name[31:]
                    file_description = f"{time_label}_{fct_label}_{file_name}_{serial_number}_{', '.join(messages)}"
                    time_buckets[time_label]["files"].append(file_description)

        results.append({"fct": fct_label, "time_buckets": time_buckets})

    save_results_to_file(input_date, shift, results)

# GUI 실행 및 이벤트
def run_analysis():
    input_date = date_entry.get_date().strftime("%Y%m%d")
    shift = shift_combobox.get()
    if shift not in ["주간", "야간"]:
        return

    try:
        analyze_ng_files(input_date, shift)
    except Exception as e:
        print(f"분석 중 오류: {e}")