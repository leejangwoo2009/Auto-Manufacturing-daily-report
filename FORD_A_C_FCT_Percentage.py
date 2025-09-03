import os
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar

# 기본 폴더 경로 및 결과 저장 경로
base_paths = [
    r"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC6",  # FCT1
    r"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC7",  # FCT2
    r"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC8",  # FCT3
    r"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC9"   # FCT4
]
output_base_path = r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C_FCT Percentage"


def parse_file_name(file_name):
    """파일명에서 yyyymmddHHmmss와 기타 정보를 추출."""
    if len(file_name) < 46:
        return None
    identifier = file_name[17]
    if identifier in ['C', 'J', '1']:
        date_time = file_name[31:45]
    elif identifier in ['P', 'N', 'S']:
        date_time = file_name[32:46]
    else:
        return None

    try:
        return datetime.strptime(date_time, '%Y%m%d%H%M%S')
    except ValueError:
        return None


def count_files(path, 기준날짜, shift, check_ng=False):
    """
    특정 폴더에서 shift 조건에 따른 모든 파일 개수 및 NG 파일 확인.
    """
    if not os.path.exists(path) or not os.path.isdir(path):
        return {"file_count": 0, "files": []}

    file_count = 0
    filtered_files = []

    for file_name in os.listdir(path):
        if check_ng and not file_name.endswith('F.txt'):
            continue

        file_datetime = parse_file_name(file_name)
        if not file_datetime:
            continue

        if shift == '주간':
            lower_bound = 기준날짜 + timedelta(hours=8, minutes=30)
            upper_bound = 기준날짜 + timedelta(hours=20, minutes=29, seconds=59)
            if lower_bound <= file_datetime <= upper_bound:
                file_count += 1
                filtered_files.append(file_name)
        elif shift == '야간':
            night_start = 기준날짜 + timedelta(hours=20, minutes=30)
            night_end = 기준날짜.replace(hour=23, minute=59, second=59)
            next_day = 기준날짜 + timedelta(days=1)
            early_morning_start = next_day.replace(hour=0, minute=0, second=0)
            early_morning_end = next_day.replace(hour=8, minute=29, second=59)

            if night_start <= file_datetime <= night_end or early_morning_start <= file_datetime <= early_morning_end:
                file_count += 1
                filtered_files.append(file_name)

    return {"file_count": file_count, "files": filtered_files}


def calculate_pass_rate_by_fct(입력날짜, shift):
    """
    입력받은 날짜와 shift 조건(주간/야간)에 따라 각 FCT 호기별 PASS율 계산.
    """
    try:
        기준날짜 = datetime.strptime(입력날짜, '%Y%m%d')
        fct_results = []

        for idx, base_path in enumerate(base_paths, start=1):
            경로_1 = os.path.join(base_path, 입력날짜, "GoodFile")
            경로_2 = os.path.join(base_path, (기준날짜 + timedelta(days=1)).strftime('%Y%m%d'), "GoodFile")

            paths_to_check = [경로_1]
            if shift == '야간':
                paths_to_check.append(경로_2)

            total_files = 0
            total_ng_files = 0

            for path in paths_to_check:
                if os.path.exists(path) and os.path.isdir(path):
                    total_result = count_files(path, 기준날짜, shift, check_ng=False)
                    ng_result = count_files(path, 기준날짜, shift, check_ng=True)
                    total_files += total_result["file_count"]
                    total_ng_files += ng_result["file_count"]

            pass_rate = 100 - (total_ng_files / total_files) * 100 if total_files > 0 else None

            fct_results.append({
                "fct": f"FCT{idx}",
                "total_files": total_files,
                "total_ng_files": total_ng_files,
                "pass_rate": pass_rate
            })

        return fct_results

    except ValueError:
        return "유효하지 않은 날짜 형식입니다. 'yyyymmdd' 형식으로 입력해주세요."


def save_results_to_file(입력날짜, 입력시간대, 결과목록, 전체총파일, 전체총NG, 전체PASS율):
    """
    결과를 파일로 저장.
    """
    try:
        output_dir = os.path.join(output_base_path)
        os.makedirs(output_dir, exist_ok=True)
        output_file_path = os.path.join(output_dir, f"{입력날짜}_{입력시간대}_FCT.txt")

        with open(output_file_path, "w", encoding="utf-8") as f:
            f.write(f"Ford A+C_{입력날짜}_{입력시간대} 결과:\n")
            for 결과 in 결과목록:
                fct = 결과["fct"]
                total_files = 결과["total_files"]
                total_ng_files = 결과["total_ng_files"]
                pass_rate = 결과["pass_rate"]

                f.write(f"{fct} 결과:\n")
                f.write(f"  - 총 파일 개수: {total_files}\n")
                f.write(f"  - NG 파일 개수: {total_ng_files}\n")
                if pass_rate is not None:
                    f.write(f"  - PASS율: {pass_rate:.2f}%\n")
                else:
                    f.write("  - PASS율: 데이터 없음\n")

            f.write("\n전체 통합 PASS율:\n")
            f.write(f"  - 총 파일 개수: {전체총파일}\n")
            f.write(f"  - NG 파일 개수: {전체총NG}\n")
            if 전체PASS율 is not None:
                f.write(f"  - PASS율: {전체PASS율:.2f}%\n")
            else:
                f.write("  - PASS율: 데이터 없음\n")

    except Exception as e:
        print(f"결과 저장 중 오류가 발생했습니다: {e}")


# GUI 코드
def run_gui():
    def calculate():
        입력날짜 = cal.get_date()
        입력날짜 = 입력날짜.replace('-', '')  # 날짜를 yyyymmdd 형식으로 변환
        입력시간대 = time_var.get()

        if 입력시간대 not in ['주간', '야간']:
            print("시간대는 '주간' 또는 '야간'만 선택 가능합니다.")
            return

        결과목록 = calculate_pass_rate_by_fct(입력날짜, 입력시간대)

        if not isinstance(결과목록, str):
            전체총파일 = sum(결과["total_files"] for 결과 in 결과목록)
            전체총NG = sum(결과["total_ng_files"] for 결과 in 결과목록)
            전체PASS율 = 100 - (전체총NG / 전체총파일) * 100 if 전체총파일 > 0 else None
            save_results_to_file(입력날짜, 입력시간대, 결과목록, 전체총파일, 전체총NG, 전체PASS율)

    root = tk.Tk()
    root.title("Ford A+C 분석 프로그램")

    tk.Label(root, text="기준 날짜 선택").pack(pady=5)
    cal = Calendar(root, selectmode='day', date_pattern='yyyy-mm-dd')
    cal.pack(pady=10)

    tk.Label(root, text="Shift 선택").pack(pady=5)
    time_var = tk.StringVar(value='주간')
    ttk.Combobox(root, textvariable=time_var, values=['주간', '야간'], state="readonly").pack(pady=10)

    tk.Button(root, text="분석 실행", command=calculate).pack(pady=20)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
