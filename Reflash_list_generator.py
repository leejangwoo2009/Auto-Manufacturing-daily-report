import os
from datetime import datetime, timedelta

def get_reflash_list(date_str, shift):
    # FCT GoodFile 경로
    fct_dirs = [
        fr"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC6\{date_str}\GoodFile",  # FCT1 경로
        fr"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC7\{date_str}\GoodFile",  # FCT2 경로
        fr"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC8\{date_str}\GoodFile",  # FCT3 경로
        fr"C:\Users\user\Desktop\FORD A+C_FCT LOG\TC9\{date_str}\GoodFile"  # FCT4 경로
    ]
    # Vision GoodFile 경로
    vision_dir = fr"C:\Users\user\Desktop\FORD A+C VISION 로그파일\{date_str}\GoodFile"

    cond_prefixes = set()

    # 1, 2. FCT 조건 판정
    for dir_path in fct_dirs:
        if not os.path.exists(dir_path):
            continue
        for fname in os.listdir(dir_path):
            if not fname.lower().endswith(".txt"):
                continue
            if len(fname) < 48:
                continue
            if fname[17] == 'J' and fname[46] == 'R':
                cond_prefixes.add(fname[:17])
            elif fname[17] == 'S' and fname[47] == 'R':
                cond_prefixes.add(fname[:17])

    # 3. Vision prefix
    vision_prefixes = set()
    vision_files = []
    if os.path.exists(vision_dir):
        for fname in os.listdir(vision_dir):
            if fname.lower().endswith(".txt") and len(fname) >= 17:
                vision_prefixes.add(fname[:17])
                vision_files.append(fname)

    # 4. 교집합
    match_prefixes = cond_prefixes & vision_prefixes

    # 5. 조건 prefix로 Vision 전체 파일명 추출
    matched_files = [f for f in vision_files if f[:17] in match_prefixes]

    # === 시간 필터 ===
    filtered_files = []
    base_date = datetime.strptime(date_str, "%Y%m%d")

    for fname in matched_files:
        try:
            timestamp_str = fname.split("_")[1][:14]  # yyyymmddhhmmss
            file_dt = datetime.strptime(timestamp_str, "%Y%m%d%H%M%S")
        except Exception:
            continue

        if shift == "주간":
            start = datetime.combine(base_date.date(), datetime.strptime("08:30:00", "%H:%M:%S").time())
            end = datetime.combine(base_date.date(), datetime.strptime("20:29:59", "%H:%M:%S").time())
            if start <= file_dt <= end:
                filtered_files.append(fname)

        elif shift == "야간":
            # 첫째날 20:30~23:59:59
            night_start = datetime.combine(base_date.date(), datetime.strptime("20:30:00", "%H:%M:%S").time())
            night_end = datetime.combine(base_date.date(), datetime.strptime("23:59:59", "%H:%M:%S").time())
            # 둘째날 00:00~08:29:59
            next_day = base_date + timedelta(days=1)
            morning_start = datetime.combine(next_day.date(), datetime.strptime("00:00:00", "%H:%M:%S").time())
            morning_end = datetime.combine(next_day.date(), datetime.strptime("08:29:59", "%H:%M:%S").time())

            if (night_start <= file_dt <= night_end) or (morning_start <= file_dt <= morning_end):
                filtered_files.append(fname)

    # 6. 결과 저장 (없으면 저장 안 함)
    if not filtered_files:
        print("[안내] 조건에 맞는 파일이 없습니다. 저장하지 않습니다.")
        return

    save_dir = r"C:\Ford A+C Test center_생산 분석 프로그램_Rev02\Data\FORD A+C_Data\FORD A+C_Reflash"
    os.makedirs(save_dir, exist_ok=True)
    save_path = os.path.join(save_dir, f"{date_str}_{shift}_Reflash_list.txt")

    with open(save_path, "w", encoding="utf-8") as f:
        for fname in filtered_files:
            f.write(fname + "\n")

    print(f"[완료] {len(filtered_files)}개 파일 저장됨 → {save_path}")

if __name__ == "__main__":
    date_str = input("날짜(yyyymmdd): ").strip()
    shift = input("주간 or 야간: ").strip()
    get_reflash_list(date_str, shift)
