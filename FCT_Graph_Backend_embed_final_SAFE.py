
import os
from collections import defaultdict
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib import font_manager, rc

# 한글 폰트 설정
font_path = "C:/Windows/Fonts/malgun.ttf"
if os.path.exists(font_path):
    font_name = font_manager.FontProperties(fname=font_path).get_name()
    rc('font', family=font_name)

def generate_graphs_embedded(selected_date, selected_shift, parent_frame):
    file_directory = r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C FCT NG List"
    file_name = f"{selected_date}_{selected_shift}_FCT NG List.txt"
    file_path = os.path.join(file_directory, file_name)

    if not os.path.exists(file_path):
        print(f"[GRAPH] ❌ 파일이 존재하지 않습니다: {file_path}")
        return

    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    fct_counts = defaultdict(int)
    fault_counts = defaultdict(int)
    total_fct_count = 0

    is_within_summary = False
    for line in lines:
        line = line.strip()
        if "======== 시간대별 & FCT별 조건별 요약 ========" in line:
            is_within_summary = True
            continue
        if is_within_summary and line:
            parts = line.split('&')
            if len(parts) == 3:
                _, fct_part, fault_part = parts
                fct = fct_part.strip()
                fct_counts[fct] += 1
                total_fct_count += 1
                fault_value = fault_part.split(':')[0].strip()
                fault_counts[fault_value] += 1

    filtered_fcts = ['FCT1', 'FCT2', 'FCT3', 'FCT4']
    fct_percentages = {fct: (fct_counts[fct] / total_fct_count) * 100 if fct in fct_counts else 0 for fct in filtered_fcts}
    total_fault_count = sum(fault_counts.values())
    fault_percentages = {fault: (count / total_fault_count) * 100 for fault, count in fault_counts.items()}
    fault_percentages = dict(sorted(fault_percentages.items(), key=lambda x: x[1], reverse=True))

    max_fct = max(fct_percentages, key=fct_percentages.get)
    if not fault_percentages:
        print("[GRAPH] ❌ NG 없음 - FCT 그래프 생략")
        return
    max_fault = max(fault_percentages, key=fault_percentages.get)

    fig, axes = plt.subplots(1, 2, figsize=(7.5, 1.5))
    fig.subplots_adjust(bottom=0.4)

    axes[0].bar(fct_percentages.keys(), fct_percentages.values(), color='salmon')
    axes[0].set_title('FCT호기별 NG 퍼센티지', fontsize=8, fontweight='bold')
    axes[0].tick_params(axis='x', labelsize=6)
    axes[0].tick_params(axis='y', labelsize=6)
    for i, (fct, pct) in enumerate(fct_percentages.items()):
        axes[0].text(i, pct - 2.5, f"{pct:.2f}%", ha='center', fontsize=8)

    wrapped_faults = {fault: '\n'.join([fault[i:i + 10] for i in range(0, len(fault), 10)]) for fault in fault_percentages.keys()}
    axes[1].bar(wrapped_faults.values(), fault_percentages.values(), color='salmon')
    axes[1].set_title('NG 항목별 퍼센티지', fontsize=8, fontweight='bold')
    axes[1].tick_params(axis='x', labelsize=5)
    axes[1].tick_params(axis='y', labelsize=6)
    for i, (fault, pct) in enumerate(fault_percentages.items()):
        axes[1].text(i, pct - 4.5, f"{pct:.2f}%", ha='center', fontsize=8)

    fig.text(0.5, 0.10, f"'{max_fct}'의 '{max_fault}'에 대한 검토 필요", ha='center', fontsize=8, color='red', fontweight='bold')

    canvas = FigureCanvasTkAgg(fig, master=parent_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(side='top', padx=5, pady=5)
