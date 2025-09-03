
import os
from datetime import datetime, timedelta

def run_fct_ng_analysis(input_date: str, shift: str):
    try:
        print(f"[FCT_NG_Backend] ✅ 분석 시작됨: input_date={input_date}, shift={shift}")

        # 🔍 모듈 임포트 시도
        print("[FCT_NG_Backend] 📦 Ford_A_C_FCT_NG_List 모듈 import 시도")
        from Ford_A_C_FCT_NG_List import analyze_ng_files
        print("[FCT_NG_Backend] ✅ analyze_ng_files 함수 불러오기 성공")

        # 🔍 분석 실행
        print("[FCT_NG_Backend] 🚀 analyze_ng_files 실행")
        result = analyze_ng_files(input_date, shift)
        print("[FCT_NG_Backend] ✅ analyze_ng_files 실행 완료")

        return result

    except Exception as e:
        print(f"[FCT_NG_Backend] ❌ 오류 발생: {e}")
        raise

# 단독 실행 방지용
if __name__ == "__main__":
    print("이 모듈은 메인 프로그램에서 import 하여 사용하는 백엔드 전용 모듈입니다.")
