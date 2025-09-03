
import os
from datetime import datetime, timedelta

# 기본 경로
base_paths = [
    r"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC6",  # FCT1
    r"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC7",  # FCT2
    r"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC8",  # FCT3
    r"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC9"   # FCT4
]
output_base_path = r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C_FCT Percentage"

def parse_file_name(file_name):
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

def save_results_to_file(입력날짜, 입력시간대, 결과목록, 전체총파일, 전체총NG, 전체PASS율):
    try:
        output_dir = os.path.join(output_base_path)
        os.makedirs(output_dir, exist_ok=True)
        output_file_path = os.path.join(output_dir, f"{입력날짜}_{입력시간대}_FCT.txt")

        with open(output_file_path, "w", encoding="utf-8") as f:
            f.write(f"Ford A+C_{입력날짜}_{입력시간대} 결과:\n")
            for 결과 in 결과목록:
                f.write(f"{결과['fct']} 결과:\n")
                f.write(f"  - 총 파일 개수: {결과['total_files']}\n")
                f.write(f"  - NG 파일 개수: {결과['total_ng_files']}\n")
                if 결과['pass_rate'] is not None:
                    f.write(f"  - PASS율: {결과['pass_rate']:.2f}%\n")
                else:
                    f.write("  - PASS율: 데이터 없음\n")

            f.write("\n전체 통합 PASS율:\n")
            f.write(f"  - 총 파일 개수: {전체총파일}\n")
            f.write(f"  - NG 파일 개수: {전체총NG}\n")
            if 전체PASS율 is not None:
                f.write(f"  - PASS율: {전체PASS율:.2f}%\n")
            else:
                f.write("  - PASS율: 데이터 없음\n")
        print(f"[FCT_PERCENTAGE] 결과 저장 완료: {output_file_path}")
    except Exception as e:
        print(f"[FCT_PERCENTAGE] 결과 저장 오류: {e}")

def run_fct_passrate_analysis(입력날짜, 입력시간대):
    print(f"[FCT_PERCENTAGE] 분석 시작: {입력날짜}, {입력시간대}")
    결과목록 = calculate_pass_rate_by_fct(입력날짜, 입력시간대)
    전체총파일 = sum(결과["total_files"] for 결과 in 결과목록)
    전체총NG = sum(결과["total_ng_files"] for 결과 in 결과목록)
    전체PASS율 = 100 - (전체총NG / 전체총파일) * 100 if 전체총파일 > 0 else None
    save_results_to_file(입력날짜, 입력시간대, 결과목록, 전체총파일, 전체총NG, 전체PASS율)
    print("[FCT_PERCENTAGE] 분석 및 저장 완료")

if __name__ == "__main__":
    print("이 모듈은 메인 프로그램에서 import 하여 사용하는 FCT PASS율 백엔드입니다.")
