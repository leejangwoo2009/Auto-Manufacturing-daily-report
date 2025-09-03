import os
import re
from collections import Counter

def find_file_by_date_and_shift(directory, input_date, shift_type):
    expected_file_name = f"{input_date}_{shift_type}_FCT NG List.txt"
    expected_file_path = os.path.join(directory, expected_file_name)
    if os.path.exists(expected_file_path):
        return expected_file_path
    return None

# 품번 매핑
품번_mapping = {
    'C': '35643009', 'J': '35915729', '1': '35654264',
    'P': '35643010', 'N': '35749091', 'S': '35915730'
}

# 품번 추출 함수
def extract_part_number(line):
    # BA1WJ + YYJJJ + SSSSSS + (아무 문자 1개) + 품번코드
    match_품번 = re.search(r'BA1WJ\d{11}.{1}([CJ1PNS])', line)
    if match_품번:
        품번_code = match_품번.group(1)
        return 품번_mapping.get(품번_code, "Unknown")
    return "Unknown"

def extract_fct_1st_ng_base_only(file_path):
    categories = [
        '제품 USB-A 문제', '제품 USB-C 문제', '테스터기 Power 관련 문제',
        'USB-C 관련 문제', 'USB-A 관련 문제', '제품 Power Pin 문제',
        '제품 Q소자 문제 가능성 많음', 'SW 설치 Or Tester기 parts 문제',
        'Mini B 관련 문제 or Carplay', '제품 충전 프로파일 문제', '회로 문제(암전류)', 'Reflash NG', '제품 Mini B 문제'
    ]
    category_pattern = '|'.join(re.escape(cat) for cat in categories)
    results_counter = Counter()

    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    base_lines = []
    for line in lines:
        if '======== 시간대별 & FCT별 조건별 요약 ========' in line:
            break
        base_lines.append(line.strip())

    for line in base_lines:
        if "FCT 2회 NG" in line:
            continue
        match = re.search(rf"({category_pattern})", line)
        if match:
            category = match.group(1)
            품번 = extract_part_number(line)
            key = f"{품번} / {category}"
            results_counter[key] += 1
    return results_counter

def extract_fct_2nd_ng_all(file_path):
    categories = [
        '제품 USB-A 문제', '제품 USB-C 문제', '테스터기 Power 관련 문제',
        'USB-C 관련 문제', 'USB-A 관련 문제', '제품 Power Pin 문제',
        '제품 Q소자 문제 가능성 많음', 'SW 설치 Or Tester기 parts 문제',
        'Mini B 관련 문제 or Carplay', '제품 충전 프로파일 문제', '회로 문제(암전류)', 'Reflash NG', '제품 Mini B 문제'
    ]
    category_pattern = '|'.join(re.escape(cat) for cat in categories)
    pattern = rf"FCT 2회 NG.*({category_pattern})"
    results_counter = Counter()

    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    for line in lines:
        match = re.search(pattern, line)
        if match:
            category = match.group(1)
            품번 = extract_part_number(line)
            key = f"{품번} / {category}"
            results_counter[key] += 1
    return results_counter

if __name__ == "__main__":
    base_dir = r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C FCT NG List"
    input_date = input("날짜를 입력하세요 (형식: yyyymmdd): ").strip()
    shift_type = input("근무 시간대를 입력하세요 (주간/야간): ").strip()

    if not re.match(r"\d{8}", input_date):
        print("❌ 날짜 형식이 잘못되었습니다. yyyymmdd 형식으로 입력해주세요.")
    elif shift_type not in ["주간", "야간"]:
        print("❌ 근무 시간대 입력이 잘못되었습니다. '주간' 또는 '야간'만 입력해주세요.")
    else:
        file_path = find_file_by_date_and_shift(base_dir, input_date, shift_type)
        if not file_path:
            print(f"❌ {input_date} {shift_type}에 해당하는 파일을 찾을 수 없습니다.")
        else:
            print("\n✅ FCT 1회 NG (요약 전 기준, '2회 NG' 제외):")
            result_1st = extract_fct_1st_ng_base_only(file_path)
            if result_1st:
                for key, count in result_1st.items():
                    print(f"{key} : {count}개")
            else:
                print("⚠️ FCT 1회 NG 관련 결과 없음.")

            print("\n✅ FCT 2회 NG (전체 기준):")
            result_2nd = extract_fct_2nd_ng_all(file_path)
            if result_2nd:
                for key, count in result_2nd.items():
                    print(f"{key} : {count}개")
            else:
                print("⚠️ FCT 2회 NG 관련 결과 없음.")
