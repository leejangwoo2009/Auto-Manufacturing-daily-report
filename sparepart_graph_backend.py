
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import os
import re
from datetime import datetime, timedelta

def create_sparepart_graph_embedded(target_frame):
    from matplotlib import font_manager as fm
    if 'Malgun Gothic' in [f.name for f in fm.fontManager.ttflist]:
        plt.rcParams['font.family'] = 'Malgun Gothic'
    else:
        plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['axes.unicode_minus'] = False

    spareparts_dir = r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C_Spareparts\\"
    oee_dir = r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C_OEE\\"

    sparepart_files = [f for f in os.listdir(spareparts_dir) if re.match(r'\d{4}\.\d{2}\.\d{2}_sparepart list\.txt', f)]
    if not sparepart_files:
        return

    latest_sparepart_file = max(sparepart_files)
    latest_sparepart_date = latest_sparepart_file[:10].replace('.', '')
    latest_sparepart_datetime = datetime.strptime(latest_sparepart_date, "%Y%m%d") + timedelta(days=1)
    latest_sparepart_next_date = latest_sparepart_datetime.strftime("%y.%m.%d")

    with open(os.path.join(spareparts_dir, latest_sparepart_file), 'r', encoding='utf-8') as file:
        lines = file.readlines()

    def extract_number(text):
        try:
            return int(text.strip().split(":")[-1].strip())
        except:
            return 0

    mini_b_stock = extract_number(lines[3])
    usb_c_stock = extract_number(lines[6])
    usb_a_stock = extract_number(lines[9])
    power_stock = extract_number(lines[12])
    mini_b_safe = extract_number(lines[4])
    usb_c_safe = extract_number(lines[7])
    usb_a_safe = extract_number(lines[10])
    power_safe = extract_number(lines[13])

    oee_files = [
        f for f in os.listdir(oee_dir)
        if re.match(r'^\d{2}\.\d{2}\.\d{2}_(주간|야간)\.txt$', f)
        and f[:8] >= latest_sparepart_next_date
    ]

    mini_b_usage = usb_c_usage = usb_a_usage = power_usage = 0
    for fname in oee_files:
        with open(os.path.join(oee_dir, fname), encoding='utf-8') as f:
            for line in f:
                if "Mini B" in line: mini_b_usage += extract_number(line)
                elif "USB-C" in line: usb_c_usage += extract_number(line)
                elif "USB-A" in line: usb_a_usage += extract_number(line)
                elif "Power" in line: power_usage += extract_number(line)

    mini_b_remaining = mini_b_stock - mini_b_usage
    usb_c_remaining = usb_c_stock - usb_c_usage
    usb_a_remaining = usb_a_stock - usb_a_usage
    power_remaining = power_stock - power_usage

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(7.5, 1.1))
    labels1, labels2 = ['Mini B', 'USB-C'], ['USB-A', 'Power']
    used1, used2 = [mini_b_usage, usb_c_usage], [usb_a_usage, power_usage]
    remain1, remain2 = [mini_b_remaining, usb_c_remaining], [usb_a_remaining, power_remaining]
    safe1, safe2 = [mini_b_safe, usb_c_safe], [usb_a_safe, power_safe]
    colors1 = ['red' if r < s else 'skyblue' for r, s in zip(remain1, safe1)]
    colors2 = ['red' if r < s else 'skyblue' for r, s in zip(remain2, safe2)]

    x1 = range(len(labels1))
    x2 = range(len(labels2))
    bars1 = ax1.barh(x1, remain1, color=colors1, label='재고')
    ax1.barh(x1, used1, left=remain1, color='black', label='사용')
    for i, bar in enumerate(bars1):
        width = bar.get_width()
        height = bar.get_y() + bar.get_height() / 2
        comment = ">안전 재고 이하 구매 신청할 것" if width < safe1[i] else ""
        ax1.text(width + 1, height, f"{width} {comment}", va='center', ha='left',
                 fontsize=7, color='red', fontweight='bold')

    ax1.set_yticks(x1)
    ax1.set_yticklabels(labels1)
    ax1.set_title("Mini B & USB-C 재고 수량", fontsize=10, fontweight='bold')

    bars2 = ax2.barh(x2, remain2, color=colors2, label='재고')
    ax2.barh(x2, used2, left=remain2, color='black', label='사용')
    for i, bar in enumerate(bars2):
        width = bar.get_width()
        height = bar.get_y() + bar.get_height() / 2
        comment = ">안전 재고 이하 구매 신청할 것" if width < safe2[i] else ""
        ax2.text(width + 1, height, f"{width} {comment}", va='center', ha='left',
                 fontsize=7, color='red', fontweight='bold')

    ax2.set_yticks(x2)
    ax2.set_yticklabels(labels2)
    ax2.set_title("USB-A & Power 재고 수량", fontsize=10, fontweight='bold')

    fig.tight_layout()
    canvas = FigureCanvasTkAgg(fig, master=target_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill='both', expand=True)
