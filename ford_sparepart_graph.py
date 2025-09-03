import os
import re
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm

# 1. 'Malgun Gothic' 폰트가 유효한지 확인
if 'Malgun Gothic' in [f.name for f in fm.fontManager.ttflist]:
    plt.rcParams['font.family'] = 'Malgun Gothic'  # Windows: Malgun Gothic
else:
    print("Warning: 'Malgun Gothic' font not found. Using default font.")
    plt.rcParams['font.family'] = 'sans-serif'  # 기본 폰트 설정

# '-' 표시 깨짐 방지
plt.rcParams['axes.unicode_minus'] = False

# 2. 가장 최신의 sparepart list 파일 찾기
spareparts_dir = r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C_Spareparts\\"
sparepart_files = [f for f in os.listdir(spareparts_dir) if re.match(r'\d{4}\.\d{2}\.\d{2}_sparepart list\.txt', f)]
latest_sparepart_file = max(sparepart_files)  # 가장 최신 파일

# 최신 sparepart 파일 날짜 추출
latest_sparepart_date = latest_sparepart_file[:10].replace('.', '')  # yyyymmdd 형식으로 변환
latest_sparepart_datetime = datetime.strptime(latest_sparepart_date, "%Y%m%d") + timedelta(days=1)  # 다음 날짜 계산
latest_sparepart_next_date = latest_sparepart_datetime.strftime("%y.%m.%d")  # OEE 파일 형식에 맞춤

# 3. 최신 sparepart 파일 데이터 추출
sparepart_file_path = os.path.join(spareparts_dir, latest_sparepart_file)
with open(sparepart_file_path, 'r', encoding='utf-8') as file:
    lines = file.readlines()

    # 데이터 정제 함수 정의 (줄 끝 숫자 추출)
    def extract_number(text):
        try:
            return int(text.strip().split(":")[-1].strip())  # 줄 끝의 숫자 추출
        except ValueError:
            return 0  # 숫자가 없으면 0 반환

    # 각 줄에서 데이터 추출
    mini_b_stock = extract_number(lines[3])  # 4번째 줄: Mini B 현재 재고
    usb_c_stock = extract_number(lines[6])  # 7번째 줄: USB-C 현재 재고
    usb_a_stock = extract_number(lines[9])  # 10번째 줄: USB-A 현재 재고
    power_stock = extract_number(lines[12])  # 13번째 줄: Power 현재 재고

    # 안전 수량 추출
    mini_b_safe = extract_number(lines[4])  # 5번째 줄: Mini B 안전 수량
    usb_c_safe = extract_number(lines[7])  # 8번째 줄: USB-C 안전 수량
    usb_a_safe = extract_number(lines[10])  # 11번째 줄: USB-A 안전 수량
    power_safe = extract_number(lines[13])  # 14번째 줄: Power 안전 수량

# 4. 주간 및 야간 OEE 파일 필터링
oee_dir = r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C_OEE\\"
oee_files = [
    f for f in os.listdir(oee_dir)
    if (re.match(r'^\d{2}\.\d{2}\.\d{2}_주간\.txt$', f) or re.match(r'^\d{2}\.\d{2}\.\d{2}_야간\.txt$', f))
    and f[:8] >= latest_sparepart_next_date  # 최신 sparepart 파일의 다음 날짜부터 포함
]

# 소비량 계산
mini_b_usage = 0
usb_c_usage = 0
usb_a_usage = 0
power_usage = 0

for oee_file in oee_files:
    oee_file_path = os.path.join(oee_dir, oee_file)
    with open(oee_file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

        # 각 줄에서 소비량 추출
        for line in lines:
            if "Mini B" in line:
                mini_b_usage += extract_number(line)
            elif "USB-C" in line:
                usb_c_usage += extract_number(line)
            elif "USB-A" in line:
                usb_a_usage += extract_number(line)
            elif "Power" in line:
                power_usage += extract_number(line)

# 5. 현재 재고에서 소비량 반영
mini_b_remaining = mini_b_stock - mini_b_usage
usb_c_remaining = usb_c_stock - usb_c_usage
usb_a_remaining = usb_a_stock - usb_a_usage
power_remaining = power_stock - power_usage

# 데이터 준비
labels_group1 = ['Mini B', 'USB-C']
labels_group2 = ['USB-A', 'Power']

used_values_group1 = [mini_b_usage, usb_c_usage]
used_values_group2 = [usb_a_usage, power_usage]

remaining_values_group1 = [mini_b_remaining, usb_c_remaining]
remaining_values_group2 = [usb_a_remaining, power_remaining]

safe_values_group1 = [mini_b_safe, usb_c_safe]
safe_values_group2 = [usb_a_safe, power_safe]

colors_group1 = ['red' if remaining < safe else 'skyblue' for remaining, safe in zip(remaining_values_group1, safe_values_group1)]
colors_group2 = ['red' if remaining < safe else 'skyblue' for remaining, safe in zip(remaining_values_group2, safe_values_group2)]

# 6. 좌우 서브플롯 그래프 생성
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(7.5, 1.1))  # 좌우로 나란히 배치

# 그래프 1: Mini B, USB-C
x_positions1 = range(len(labels_group1))
remaining_bars1 = ax1.barh(x_positions1, remaining_values_group1, color=colors_group1, label='현재 재고')
ax1.barh(x_positions1, used_values_group1, left=remaining_values_group1, color='black', label='현 사용량')
for i, bar in enumerate(remaining_bars1):
    width = bar.get_width()
    height = bar.get_y() + bar.get_height() / 2
    safe_threshold = safe_values_group1[i]
    comment = ">안전 재고 이하 구매 신청할 것" if width < safe_threshold else ""
    ax1.text(width + 1, height, f'{width} {comment}', va='center', ha='left', fontsize=7, color='red', fontweight='bold')

ax1.set_yticks(x_positions1)
ax1.set_yticklabels(labels_group1)
ax1.set_title('Mini B & USB-C 재고 수량', fontsize=10, fontweight='bold')

# 그래프 2: USB-A, Power
x_positions2 = range(len(labels_group2))
remaining_bars2 = ax2.barh(x_positions2, remaining_values_group2, color=colors_group2, label='현재 재고')
ax2.barh(x_positions2, used_values_group2, left=remaining_values_group2, color='black', label='현 사용량')
for i, bar in enumerate(remaining_bars2):
    width = bar.get_width()
    height = bar.get_y() + bar.get_height() / 2
    safe_threshold = safe_values_group2[i]
    comment = ">안전 재고 이하 구매 신청할 것" if width < safe_threshold else ""
    ax2.text(width + 1, height, f'{width} {comment}', va='center', ha='left', fontsize=7, color='red', fontweight='bold')

ax2.set_yticks(x_positions2)
ax2.set_yticklabels(labels_group2)
ax2.set_title('USB-A & Power 재고 수량', fontsize=10, fontweight='bold')

# 공통 구성
plt.tight_layout()
plt.show()
