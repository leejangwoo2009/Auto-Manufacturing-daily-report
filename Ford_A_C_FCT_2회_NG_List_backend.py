
import os
import re
from collections import Counter

def run_fct_2nd_ng_analysis(input_date: str, shift_type: str, output_callback=print):
    """
    FCT 1회 및 2회 NG 분석 결과를 출력 콜백으로 전달합니다. (품번/카테고리 기준)
    """
    dir_path = r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C FCT NG List"
    expected_file_name = f"{input_date}_{shift_type}_FCT NG List.txt"
    file_path = os.path.join(dir_path, expected_file_name)

    if not os.path.exists(file_path):
        output_callback(f"파일을 찾을 수 없습니다: {file_path}")
        return

    categories = [
        '제품 USB-A 문제', '제품 USB-C 문제', '테스터기 Power 관련 문제',
        'USB-C 관련 문제', 'USB-A 관련 문제', '제품 Power Pin 문제',
        '제품 Q소자 문제 가능성 많음', 'SW 설치 Or Tester기 parts 문제',
        'Mini B 관련 문제 or Carplay', '제품 충전 프로파일 문제', '회로 문제(암전류)', 'Reflash NG', '제품 Mini B 문제'
    ]

    품번_mapping = {
        'C': '35643009', 'J': '35915729', '1': '35654264',
        'P': '35643010', 'N': '35749091', 'S': '35915730'
    }

    # 품번 추출 함수 (정규식 기반)
    def extract_part_number(line: str) -> str:
        match_품번 = re.search(r'BA1WJ\d{11}.{1}([CJ1PNS])', line)
        if match_품번:
            품번_code = match_품번.group(1)
            return 품번_mapping.get(품번_code, "Unknown")
        return "Unknown"

    category_pattern = '|'.join(re.escape(cat) for cat in categories)
    fct2_pattern = rf"FCT 2회 NG.*?({category_pattern})"
    fct1_pattern = rf"^(?!.*FCT 2회 NG).*?({category_pattern})"

    # 파일 로딩
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    # 요약 이전 줄만 추출
    base_lines = []
    for line in lines:
        if '======== 시간대별 & FCT별 조건별 요약 ========' in line:
            break
        base_lines.append(line.strip())

    # FCT 1회 NG 분석
    fct1_counter = Counter()
    for line in base_lines:
        match = re.search(fct1_pattern, line)
        if match:
            category = match.group(1)
            품번 = extract_part_number(line)
            fct1_counter[f"{품번} / {category}"] += 1

    # FCT 2회 NG 분석
    fct2_counter = Counter()
    for line in lines:
        match = re.search(fct2_pattern, line)
        if match:
            category = match.group(1)
            품번 = extract_part_number(line)
            fct2_counter[f"{품번} / {category}"] += 1

    # 출력
    output_callback("\n[FCT 1회 NG] 분석 결과:")
    if fct1_counter:
        for key, cnt in fct1_counter.items():
            output_callback(f" - {key}: {cnt}개")
    else:
        output_callback(" - 관련 항목을 찾을 수 없습니다.")

    output_callback("\n[FCT 2회 NG] 분석 결과:")
    if fct2_counter:
        for key, cnt in fct2_counter.items():
            output_callback(f" - {key}: {cnt}개")
    else:
        output_callback(" - 관련 항목을 찾을 수 없습니다.")
