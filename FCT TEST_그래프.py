import os
import matplotlib.pyplot as plt
from matplotlib import font_manager, rc
from collections import defaultdict
from tkinter import Tk, Label, Button, StringVar, Toplevel
from tkinter.ttk import Combobox
from tkcalendar import DateEntry

# 한글 깨짐 문제 해결
font_path = "C:/Windows/Fonts/malgun.ttf"  # Windows 환경의 기본 한글 폰트 경로
if os.path.exists(font_path):
    font_name = font_manager.FontProperties(fname=font_path).get_name()
    rc('font', family=font_name)


# 데이터 계산 및 그래프 생성 함수
def generate_graphs(selected_date, selected_shift,
                    graph_width=14, graph_height=8,  # 그래프 크기 설정
                    title_size=16, axis_label_size=14,  # 제목 및 축 레이블 글자 크기
                    tick_label_size=12, text_size=12):  # 눈금 글자 크기 및 텍스트 표시 크기

    file_directory = r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C FCT NG List"
    file_name = f"{selected_date}_{selected_shift}_FCT NG List.txt"
    file_path = os.path.join(file_directory, file_name)

    if not os.path.exists(file_path):
        error_window = Toplevel()
        Label(error_window, text=f"파일이 존재하지 않습니다: {file_path}").pack()
        return

    # 파일 읽기
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    # 데이터 처리
    fct_counts = defaultdict(int)
    fault_counts = defaultdict(int)
    total_fct_count = 0

    is_within_summary = False
    for line in lines:
        line = line.strip()
        if "======== 시간대별 & FCT별 조건별 요약 ========" in line:
            is_within_summary = True
            continue
        if is_within_summary:
            if line:
                parts = line.split('&')
                if len(parts) == 3:
                    _, fct_part, fault_part = parts
                    # FCT 데이터 추출
                    fct = fct_part.strip()
                    fct_counts[fct] += 1
                    total_fct_count += 1

                    # Fault Value 데이터 추출
                    fault_value = fault_part.split(':')[0].strip()
                    fault_counts[fault_value] += 1

    # FCT 퍼센티지 계산 (FCT1 ~ FCT4 고정하여 표현)
    filtered_fcts = ['FCT1', 'FCT2', 'FCT3', 'FCT4']
    fct_percentages = {fct: (fct_counts[fct] / total_fct_count) * 100 if fct in fct_counts else 0 for fct in filtered_fcts}

    # Fault Value 퍼센티지 계산
    total_fault_count = sum(fault_counts.values())
    fault_percentages = {fault: (count / total_fault_count) * 100 for fault, count in fault_counts.items()}
    fault_percentages = dict(sorted(fault_percentages.items(), key=lambda x: x[1], reverse=True))

    # 최대 퍼센티지를 찾기 위한 데이터
    max_fct = max(fct_percentages, key=fct_percentages.get)
    max_fault = max(fault_percentages, key=fault_percentages.get)

    # 그래프 생성
    fig, axes = plt.subplots(1, 2, figsize=(graph_width, graph_height))
    plt.subplots_adjust(bottom=0.4)  # 하단 여백 설정

    # 첫 번째 그래프: FCT 퍼센티지 (FCT1 ~ FCT4만 표시)
    axes[0].bar(fct_percentages.keys(), fct_percentages.values(), color='lightcoral')
    axes[0].set_title('FCT호기별 NG 퍼센티지', fontsize=title_size, fontweight='bold')
    axes[0].set_xlabel('', fontsize=axis_label_size)  # X축 레이블 글자 크기
    axes[0].tick_params(axis='x', labelsize=tick_label_size)  # X축 눈금 글자 크기
    axes[0].tick_params(axis='y', labelsize=tick_label_size)  # Y축 눈금 글자 크기
    for i, (fct, pct) in enumerate(fct_percentages.items()):
        axes[0].text(i, pct - 2.5, f"{pct:.2f}%", ha='center', fontsize=text_size)

    # 두 번째 그래프: Fault Value 퍼센티지 (긴 X축 레이블 줄바꿈)
    wrapped_faults = {fault: '\n'.join([fault[i:i + 10] for i in range(0, len(fault), 10)]) for fault in fault_percentages.keys()}
    axes[1].bar(wrapped_faults.values(), fault_percentages.values(), color='lightcoral')
    axes[1].set_title('NG 항목별 퍼센티지', fontsize=title_size, fontweight='bold')
    axes[1].set_xlabel('', fontsize=axis_label_size)  # X축 레이블 글자 크기
    axes[1].tick_params(axis='x', labelsize=tick_label_size)  # X축 눈금 글자 크기
    axes[1].tick_params(axis='y', labelsize=tick_label_size)  # Y축 눈금 글자 크기
    for i, (fault, pct) in enumerate(fault_percentages.items()):
        axes[1].text(i, pct -4.5, f"{pct:.2f}%", ha='center', fontsize=text_size)

    # 그래프 하단 메시지
    fig.text(0.5, 0.12,  # 위치를 약간 아래로 이동
             f"'{max_fct}'의 '{max_fault}'에 대한 검토 필요",
             ha='center', fontsize=text_size + 2,  # 글자 크기를 키움
             color='red', fontweight='bold')  # 글씨 굵게 설정
    plt.show()


# GUI 생성
def main():
    root = Tk()
    root.title("Ford A+C FCT NG List 분석 도구")

    Label(root, text="기준 날짜 선택 (yyyymmdd):").pack(pady=5)
    date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
    date_entry.pack(pady=5)

    Label(root, text="Shift 선택").pack(pady=5)
    shift_var = StringVar()
    shift_combobox = Combobox(root, textvariable=shift_var, values=["주간", "야간"], state="readonly", width=10)
    shift_combobox.pack(pady=5)

    def on_submit():
        selected_date = date_entry.get_date().strftime("%Y%m%d")
        selected_shift = shift_combobox.get()
        if not selected_shift:
            error_window = Toplevel()
            Label(error_window, text="주간 또는 야간을 선택해주세요!").pack()
            return

        # 그래프 생성 함수 호출
        generate_graphs(
            selected_date, selected_shift,
            graph_width=7,  # 그래프 너비
            graph_height=2,  # 그래프 높이
            title_size=8,  # 그래프 제목 글자 크기
            axis_label_size=10,  # X축, Y축 라벨 글자 크기
            tick_label_size=5,  # X축, Y축 눈금 레이블 크기
            text_size=7  # 그래프 수치 텍스트 크기
        )

    Button(root, text="분석 실행", command=on_submit).pack(pady=20)
    root.mainloop()


if __name__ == "__main__":
    main()
