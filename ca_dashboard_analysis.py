"""
CA Dashboard — 매니저별 성과 분석 스크립트
===========================================
PRD.md의 Block A~H 구조를 기반으로:
  1) 구글 시트에서 블록별 데이터 로드 (탐색 레이블 기반)
  2) 모든 블록을 manager_name + weekly/monthly 키로 병합
  3) '종결 상태값이 아닌 리드 비율' 계산
  4) 콜 시도 횟수 ↔ 결제 금액 상관관계 분석
  5) 매니저별 리드 진척 / 콜 활동 / 최종 성과 리포트 출력

인증 방식:
  gspread.oauth() — 개인 구글 계정의 OAuth 2.0 인증
  (시트를 공개하지 않아도 됨, 본인 계정에 읽기 권한만 있으면 OK)

사전 설정:
  1) Google Cloud Console에서 OAuth 2.0 클라이언트 ID 생성 (완료)
  2) credentials.json을 CA_AX 폴더에 저장 (완료)
  3) 최초 실행 시 브라우저에서 구글 로그인 → 토큰 자동 저장

실행 방법:
  pip install gspread pandas numpy
  python ca_dashboard_analysis.py
"""

import pandas as pd
import numpy as np
import math
import os
import sys
import subprocess
from pathlib import Path

# ── gspread 자동 설치 시도 ──
GSPREAD_AVAILABLE = False
try:
    import gspread
    GSPREAD_AVAILABLE = True
except ImportError:
    try:
        print("  gspread 미설치 → 자동 설치 중...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "gspread", "-q"],
            stderr=subprocess.DEVNULL,
        )
        import gspread
        GSPREAD_AVAILABLE = True
    except Exception:
        print("  [!] gspread 자동 설치 실패. 로컬에서 pip install gspread 후 재실행하세요.")
        print("      (샘플 데이터로 로직 검증을 진행합니다)")

pd.set_option("display.float_format", lambda x: f"{x:.1f}")
pd.set_option("display.max_columns", 30)
pd.set_option("display.width", 200)

# ============================================================
# 0. 구글 시트 설정
# ============================================================
SHEET_ID = "1RZdwVEOjwUZZKJZbPjTYWyzS7ebNZmwSmBFLb-Z0DqA"
SHEET_GID = "1334115015"  # 추출 탭의 gid (수정됨)
EXTRACT_TAB_NAME = "추출"  # 탭 이름 직접 지정

# ── credentials.json 경로 탐색 (우선순위) ──
# 1순위: 스크립트와 같은 폴더의 CA_AX/credentials.json
# 2순위: 상위 CA dashboard 프로젝트의 CA_AX/credentials.json
# 3순위: 사용자 홈의 Documents 경로 (해나님 실제 경로)
# 4순위: gspread 기본 경로 (~/.config/gspread/)
SCRIPT_DIR = Path(__file__).resolve().parent
CANDIDATE_PATHS = [
    SCRIPT_DIR / "CA_AX" / "credentials.json",
    SCRIPT_DIR.parent / "CA_AX" / "credentials.json",
    Path.home() / "Documents" / "Claude" / "Projects" / "CA dashboard" / "CA_AX" / "credentials.json",
    Path.home() / ".config" / "gspread" / "credentials.json",
]

def find_credentials() -> Path:
    """우선순위에 따라 credentials.json 경로를 탐색"""
    for p in CANDIDATE_PATHS:
        if p.exists():
            return p
    return CANDIDATE_PATHS[-1]  # 없으면 기본 경로 반환 (에러 메시지용)

CREDENTIALS_FILE = find_credentials()
# authorized_user.json은 credentials.json과 같은 폴더에 저장
AUTHORIZED_USER_FILE = CREDENTIALS_FILE.parent / "authorized_user.json"

# 매니저 이름 목록 (PRD 기준 7인 + Others)
VALID_MANAGERS = ["Tommy", "Jane", "Owen", "Mia", "Jinny", "UBASE", "Bella"]

# 블록별 탐색 레이블 → 컬럼 매핑
BLOCK_LABELS = {
    "A": "lead&call >>",
    "B": "Booking >>",
    "C": "New Stdt >>",
    "D": "up&cross >>",
    "E": "call history >>",
    "F": "블록Fpush >>",
    "G": "Paid Cost (M)",
    "H": "Paid Cost (W)",
    "I": "REFUND",
}

def normalize_manager(name) -> str:
    """매니저 이름을 7인 + Others 로 통일"""
    if pd.isna(name):
        return "Others"
    name = str(name).strip()
    # 대소문자 무시 매칭
    for valid in VALID_MANAGERS:
        if name.lower() == valid.lower():
            return valid
    return "Others"


# ============================================================
# 1. gspread OAuth 2.0 인증 + 구글 시트 로드
# ============================================================
def authenticate_gspread():
    """
    gspread.oauth()를 사용하여 OAuth 2.0 인증 수행.

    동작 방식:
    1) authorized_user.json이 있으면 → 토큰 재사용 (브라우저 팝업 없음)
    2) 없으면 → credentials.json 기반으로 브라우저 팝업 → 구글 로그인 → 토큰 자동 저장

    Returns: gspread.Client 객체
    """
    print(f"  credentials.json 경로: {CREDENTIALS_FILE}")

    if not CREDENTIALS_FILE.exists():
        print(f"\n  [오류] credentials.json을 찾을 수 없습니다!")
        print(f"         확인한 경로들:")
        for p in CANDIDATE_PATHS:
            status = "O" if p.exists() else "X"
            print(f"           [{status}] {p}")
        return None

    try:
        gc = gspread.oauth(
            credentials_filename=str(CREDENTIALS_FILE),
            authorized_user_filename=str(AUTHORIZED_USER_FILE),
        )
        print(f"  [OK] OAuth 인증 성공")
        if AUTHORIZED_USER_FILE.exists():
            print(f"       (토큰 재사용: {AUTHORIZED_USER_FILE.name})")
        else:
            print(f"       (최초 인증 완료 — 토큰이 저장되었습니다)")
        return gc

    except Exception as e:
        print(f"  [오류] OAuth 인증 실패: {e}")
        print(f"\n  확인 사항:")
        print(f"    - Google Cloud Console에서 Google Sheets API + Drive API가 활성화되어 있는지")
        print(f"    - OAuth 클라이언트 유형이 '데스크톱 앱'인지")
        print(f"    - 토큰 만료 시: {AUTHORIZED_USER_FILE} 삭제 후 재실행")
        return None


def load_sheet_via_oauth() -> pd.DataFrame:
    """
    gspread.oauth()로 비공개 구글 시트에 접속하여
    추출 탭의 전체 데이터를 DataFrame으로 반환.

    탭 탐색 우선순위:
    1) EXTRACT_TAB_NAME이 설정되어 있으면 해당 이름으로 검색
    2) gid(SHEET_GID)로 워크시트 매칭
    3) 각 탐색 레이블이 포함된 탭 자동 탐색
    """
    gc = authenticate_gspread()
    if gc is None:
        return None

    try:
        # 스프레드시트 열기
        spreadsheet = gc.open_by_key(SHEET_ID)
        print(f"  [OK] 스프레드시트 열기 성공: '{spreadsheet.title}'")

        # --- 탭 찾기 ---
        worksheet = None

        # 방법 1: 탭 이름으로 검색
        if EXTRACT_TAB_NAME:
            try:
                worksheet = spreadsheet.worksheet(EXTRACT_TAB_NAME)
                print(f"  [OK] 탭 '{EXTRACT_TAB_NAME}' 발견")
            except Exception:
                print(f"  [!] 탭 '{EXTRACT_TAB_NAME}'을 찾지 못함, gid로 재시도...")

        # 방법 2: gid로 검색
        if worksheet is None and SHEET_GID:
            for ws in spreadsheet.worksheets():
                if str(ws.id) == str(SHEET_GID):
                    worksheet = ws
                    print(f"  [OK] gid={SHEET_GID} 매칭 탭 발견: '{ws.title}'")
                    break

        # 방법 3: 탐색 레이블이 포함된 탭 자동 탐색
        if worksheet is None:
            first_label = list(BLOCK_LABELS.values())[0]  # "lead&call >>"
            print(f"  [!] gid 매칭 실패, 탐색 레이블 '{first_label}'로 탭 검색 중...")
            for ws in spreadsheet.worksheets():
                try:
                    cell = ws.find(first_label)
                    if cell:
                        worksheet = ws
                        print(f"  [OK] 레이블 '{first_label}' 발견: 탭 '{ws.title}'")
                        break
                except Exception:
                    continue

        if worksheet is None:
            print(f"  [오류] 추출 탭을 찾을 수 없습니다.")
            print(f"         SHEET_GID({SHEET_GID}) 또는 탐색 레이블을 확인하세요.")
            # 사용 가능한 탭 목록 출력
            print(f"\n  현재 시트의 탭 목록:")
            for ws in spreadsheet.worksheets():
                print(f"    - '{ws.title}' (gid: {ws.id})")
            return None

        # --- 전체 데이터 로드 ---
        print(f"  데이터 로드 중... (탭: '{worksheet.title}')")
        all_values = worksheet.get_all_values()

        if not all_values:
            print(f"  [오류] 탭 '{worksheet.title}'에 데이터가 없습니다.")
            return None

        df = pd.DataFrame(all_values)
        # 빈 문자열을 NaN으로 변환 (FutureWarning 방지)
        df = df.replace("", np.nan)
        try:
            df = df.infer_objects(copy=False)
        except Exception:
            pass

        print(f"  [OK] 구글 시트 로드 성공: {df.shape[0]}행 x {df.shape[1]}열")
        return df

    except Exception as e:
        print(f"  [오류] 구글 시트 접속 실패: {e}")
        return None


def find_block_in_sheet(raw_df: pd.DataFrame, label: str) -> pd.DataFrame:
    """
    raw DataFrame에서 탐색 레이블 셀을 찾고,
    해당 셀의 우측 영역을 블록 데이터로 추출.

    PRD 정의: "[탐색 레이블]이 적힌 셀을 찾고, 해당 셀의 우측 영역을 데이터셋으로 인식"
    """
    if raw_df is None:
        return None

    # 모든 블록 레이블 목록 (블록 경계 인식용)
    all_labels = list(BLOCK_LABELS.values())

    def _is_block_label(cell_val, exclude_label):
        """셀 값이 다른 블록의 레이블인지 판별 ('>>' 포함 + 현재 블록 제외)"""
        s = str(cell_val).strip().lower()
        if ">>" not in s:
            return False
        # 현재 블록 자신의 레이블은 제외
        if exclude_label.lower() in s:
            return False
        # 다른 블록 레이블 중 하나와 매칭되는지 확인
        for lbl in all_labels:
            if lbl.lower() in s:
                return True
        # ">>" 를 포함하고 "블록" 을 포함하면 블록 경계로 판단
        if "블록" in s:
            return True
        return False

    # 레이블 셀 위치 찾기
    for row_idx in range(raw_df.shape[0]):
        for col_idx in range(raw_df.shape[1]):
            cell = str(raw_df.iloc[row_idx, col_idx]).strip()
            if label.lower() in cell.lower():
                # 레이블 셀의 우측부터가 데이터 영역
                # 첫 번째 행 = 헤더, 이후 = 데이터
                data_start_col = col_idx + 1

                # 헤더 행 추출 — 다음 블록 레이블을 만나면 중단
                headers = []
                for hcol in range(data_start_col, raw_df.shape[1]):
                    hval = raw_df.iloc[row_idx, hcol]
                    # 다음 블록 레이블이면 여기서 중단
                    if _is_block_label(hval, label):
                        break
                    if pd.notna(hval) and str(hval).strip():
                        headers.append(str(hval).strip())
                    else:
                        # 빈 셀(NaN)이면 연속 데이터 영역 끝으로 판단
                        break

                if not headers:
                    return None

                num_cols = len(headers)

                # 데이터 행 추출 (빈 행을 만날 때까지)
                data_rows = []
                for data_row in range(row_idx + 1, raw_df.shape[0]):
                    row_data = raw_df.iloc[data_row, data_start_col:data_start_col + num_cols].tolist()
                    if all(pd.isna(v) for v in row_data):
                        break
                    data_rows.append(row_data)

                if data_rows:
                    block_df = pd.DataFrame(data_rows, columns=headers)
                    # ── 중복 컬럼명 제거 (시트에 같은 헤더가 반복될 수 있음) ──
                    dup_mask = block_df.columns.duplicated()
                    if dup_mask.any():
                        block_df = block_df.loc[:, ~dup_mask]
                        print(f"    블록 [{label}] 중복 컬럼 제거 → {block_df.shape[1]}열")
                    print(f"    블록 [{label}] 추출 완료: {block_df.shape[0]}행 x {block_df.shape[1]}열")
                    print(f"      컬럼: {list(block_df.columns)}")
                    return block_df

    print(f"    블록 [{label}] 레이블을 찾지 못했습니다.")
    return None


# ============================================================
# 2. 샘플 데이터 (구글 시트 로드 실패 시 폴백)
# ============================================================
from itertools import product as iter_product

WEEKS_SAMPLE = pd.to_datetime([
    "2026-03-02", "2026-03-09", "2026-03-16", "2026-03-23", "2026-03-30"
])

def load_sample_block_a() -> pd.DataFrame:
    """Block A: 리드 & 상담 현황 (weekly × manager_name)"""
    np.random.seed(42)
    rows = []
    for w, mgr in iter_product(WEEKS_SAMPLE, VALID_MANAGERS):
        lead = np.random.randint(30, 120)
        clean = int(lead * np.random.uniform(0.75, 0.95))
        consulted = int(clean * np.random.uniform(0.40, 0.85))
        success = int(consulted * np.random.uniform(0.15, 0.50))
        rows.append({
            "monthly": w.strftime("%Y-%m"),
            "weekly": w,
            "manager_name": mgr,
            "lead_cnt": lead,
            "clean_lead_cnt": clean,
            "consulted_cnt": consulted,
            "success_cnt": success,
            "total_call_attempts": np.random.randint(50, 400),
            "mobile_call_attempts": np.random.randint(30, 300),
            "phone_connected_cnt": np.random.randint(10, 100),
            "total_paid_cost": np.random.randint(500_000, 8_000_000),
            "coupon_usage": np.random.randint(0, 20),
        })
    return pd.DataFrame(rows)


def load_sample_block_b() -> pd.DataFrame:
    rows = []
    for w in WEEKS_SAMPLE:
        rows.append({
            "monthly": w.strftime("%Y-%m"),
            "weekly": w,
            "net_booking_in_krw": np.random.randint(10_000_000, 80_000_000),
        })
    return pd.DataFrame(rows)


def load_sample_block_c() -> pd.DataFrame:
    rows = []
    for w in WEEKS_SAMPLE:
        total = np.random.randint(20, 100)
        rows.append({
            "monthly": w.strftime("%Y-%m"),
            "weekly": w,
            "newstdt_total": total,
            "newstdt_ac": int(total * np.random.uniform(0.3, 0.7)),
        })
    return pd.DataFrame(rows)


def load_sample_block_d() -> pd.DataFrame:
    rows = []
    for w in WEEKS_SAMPLE:
        rows.append({
            "monthly": w.strftime("%Y-%m"),
            "weekly": w,
            "up_cross_selling_count": np.random.randint(5, 40),
        })
    return pd.DataFrame(rows)


def load_sample_block_e() -> pd.DataFrame:
    np.random.seed(99)
    rows = []
    for w, mgr in iter_product(WEEKS_SAMPLE, VALID_MANAGERS):
        total_calls = np.random.randint(80, 500)
        mobile = int(total_calls * np.random.uniform(0.5, 0.9))
        success = int(total_calls * np.random.uniform(0.2, 0.6))
        fail = total_calls - success
        distinct_num = np.random.randint(20, 150)
        rows.append({
            "monthly": w.strftime("%Y-%m"),
            "weekly": w,
            "manager_name": mgr,
            "total_calls": total_calls,
            "mobile_010_calls": mobile,
            "non_mobile_calls": total_calls - mobile,
            "success_calls": success,
            "fail_calls": fail,
            "distinct_numbers": distinct_num,
            "attempts_per_number": round(total_calls / max(distinct_num, 1), 2),
            "total_duration_sum": np.random.randint(5000, 30000),
            "success_real_duration_sum": np.random.randint(3000, 20000),
            "effort_seconds_est": np.random.randint(8000, 40000),
            "mobile_010_success_calls": int(success * np.random.uniform(0.4, 0.8)),
        })
    return pd.DataFrame(rows)


def load_sample_block_g() -> pd.DataFrame:
    rows = []
    for m, mgr in iter_product(["2026-03"], VALID_MANAGERS):
        rows.append({
            "monthly": m,
            "manager_name": mgr,
            "paid_cost_m": np.random.randint(1_000_000, 15_000_000),
            "order_cnt_m": np.random.randint(3, 30),
            "coupon_cnt_m": np.random.randint(0, 10),
        })
    return pd.DataFrame(rows)


def load_sample_block_h() -> pd.DataFrame:
    np.random.seed(7)
    rows = []
    for w, mgr in iter_product(WEEKS_SAMPLE, VALID_MANAGERS):
        rows.append({
            "weekly": w,
            "manager_name": mgr,
            "paid_cost_w": np.random.randint(200_000, 5_000_000),
            "order_cnt_w": np.random.randint(1, 15),
            "coupon_cnt_w": np.random.randint(0, 5),
        })
    return pd.DataFrame(rows)


def load_sample_block_i() -> pd.DataFrame:
    """Block I: 환불 현황 (weekly × manager_name)"""
    np.random.seed(77)
    rows = []
    for w, mgr in iter_product(WEEKS_SAMPLE, VALID_MANAGERS):
        # 환불은 매주 0~3건 정도 발생하는 것으로 시뮬레이션
        refund_count = np.random.randint(0, 4)
        rows.append({
            "month_kst": w.strftime("%Y-%m"),
            "week_monday_kst": w,
            "manager_name": mgr,
            "total_refund_amount": np.random.randint(0, 2_000_000) if refund_count > 0 else 0,
            "refund_count": refund_count,
            "refund_order_ids": "",
        })
    return pd.DataFrame(rows)


def load_sample_block_f() -> pd.DataFrame:
    """Block F: Push 관리 현황 (weekly × manager_name)"""
    np.random.seed(55)
    rows = []
    for w, mgr in iter_product(WEEKS_SAMPLE, VALID_MANAGERS):
        b_total = np.random.randint(5, 40)
        b_target = np.random.randint(0, max(b_total, 1))
        b_done = np.random.randint(0, max(b_target, 1))
        c_total = np.random.randint(3, 25)
        c_target = np.random.randint(0, max(c_total, 1))
        c_done = np.random.randint(0, max(c_target, 1))
        rows.append({
            "lead_month": w.strftime("%Y-%m"),
            "lead_week": w,
            "manager_name": mgr,
            "A_Link_Created": np.random.randint(10, 60),
            "B_Total": b_total,
            "B_Push_Target": b_target,
            "B_Push_Done": b_done,
            "B_Push_Fail": b_target - b_done,
            "C_Total": c_total,
            "C_Push_Target": c_target,
            "C_Push_Done": c_done,
            "C_Push_Fail": c_target - c_done,
            "D_Paid_Success": np.random.randint(1, 15),
            "B_Push_Skip": np.random.randint(0, 5),
            "C_Push_Skip": np.random.randint(0, 3),
        })
    return pd.DataFrame(rows)


# ============================================================
# 3. 데이터 로드 통합 함수
# ============================================================
def load_all_blocks():
    """
    gspread OAuth 2.0으로 비공개 구글 시트에서 데이터를 로드하고,
    실패 시 샘플 데이터로 폴백.

    인증 흐름:
    1) gspread 설치 확인
    2) credentials.json 존재 확인
    3) OAuth 인증 → 시트 접속 → 블록별 데이터 추출
    """
    print("\n[1단계] 데이터 로드 (gspread OAuth 2.0)")
    print("-" * 40)

    if not GSPREAD_AVAILABLE:
        print("  [!] gspread 라이브러리가 없습니다. pip install gspread 후 재실행하세요.")
        return _fallback_sample_data()

    if not CREDENTIALS_FILE.exists():
        print(f"  [!] credentials.json을 찾을 수 없습니다.")
        print(f"      CA_AX 폴더에 credentials.json이 있는지 확인해주세요.")
        return _fallback_sample_data()

    # --- gspread OAuth 인증 → 시트 로드 ---
    raw = load_sheet_via_oauth()

    if raw is not None:
        # 구글 시트에서 블록별 추출 시도
        blocks = {}
        for block_name, label in BLOCK_LABELS.items():
            block_df = find_block_in_sheet(raw, label)
            if block_df is not None:
                blocks[block_name] = block_df

        if len(blocks) >= 4:  # 주요 블록이 4개 이상 추출되면 성공
            print(f"\n  [OK] 구글 시트에서 {len(blocks)}개 블록 추출 완료")
            return blocks, "google_sheet"
        else:
            print(f"\n  [!] {len(blocks)}개 블록만 추출됨 (최소 4개 필요)")
            print(f"      탐색 레이블이 시트 내에 정확히 존재하는지 확인하세요.")

    # --- 폴백: 샘플 데이터 ---
    return _fallback_sample_data()


def _fallback_sample_data():
    """샘플 데이터로 분석 로직을 검증"""
    print("\n  → 샘플 데이터를 사용하여 분석 로직을 검증합니다.")
    print("    (실제 데이터 사용 시: OAUTH_SETUP_GUIDE.md 참고)")
    return {
        "A": load_sample_block_a(),
        "B": load_sample_block_b(),
        "C": load_sample_block_c(),
        "D": load_sample_block_d(),
        "E": load_sample_block_e(),
        "F": load_sample_block_f(),
        "G": load_sample_block_g(),
        "H": load_sample_block_h(),
        "I": load_sample_block_i(),
    }, "sample"


# ============================================================
# 4. 블록 전처리 — Pre-Aggregation (SUMIFS 재현)
# ============================================================
# 원칙:
#   1) 중복 컬럼명 제거 (시트 헤더 중복 대비)
#   2) 키 컬럼 정규화 (이름·타입 통일)
#   3) 값 컬럼 전부 to_numeric
#   4) groupby(keys).sum() → 유니크 키 보장
#   5) 값 컬럼에 블록 접두사 부착 → 이름 절대 안 겹침
# ============================================================

def _sanitize_block(raw_df, key_cols, rename_map=None, prefix=""):
    """
    하나의 블록을 받아 깨끗한 (키 + 접두사 값) DataFrame 반환.

    Parameters
    ----------
    raw_df      : find_block_in_sheet 에서 추출한 원본
    key_cols    : 조인 키가 될 컬럼 목록 (정규화 후 이름)
    rename_map  : 원본 컬럼 → 정규화 이름 매핑 (optional)
    prefix      : 값 컬럼에 붙일 접두사, 예: "A_"
    """
    df = raw_df.copy()

    # ① 중복 컬럼 제거 (첫 번째만 유지)
    dup = df.columns.duplicated()
    if dup.any():
        df = df.loc[:, ~dup]

    # ② 컬럼 리네임 (예: months→monthly, weeks→weekly)
    if rename_map:
        # 충돌 방지: 대상 이름이 이미 있으면 먼저 제거
        for src, dst in rename_map.items():
            if src in df.columns and dst in df.columns and src != dst:
                df = df.drop(columns=[dst])
        df = df.rename(columns=rename_map)

    # ② -bis 다시 한번 중복 컬럼 제거
    dup = df.columns.duplicated()
    if dup.any():
        df = df.loc[:, ~dup]

    # ③ 키 컬럼 중 존재하지 않는 것은 NaN으로 생성
    for k in key_cols:
        if k not in df.columns:
            df[k] = np.nan

    # ④ 값 컬럼 식별 + 숫자 강제 변환
    val_cols = [c for c in df.columns if c not in key_cols]
    for c in val_cols:
        # 통화 기호(₩, $)와 쉼표 제거 후 숫자 변환
        if df[c].dtype == object:
            df[c] = (df[c].astype(str)
                     .str.replace("₩", "", regex=False)
                     .str.replace(",", "", regex=False)
                     .str.replace("$", "", regex=False)
                     .str.strip())
        df[c] = pd.to_numeric(df[c], errors="coerce")
    # 숫자 변환 실패 컬럼 제거 (모두 NaN)
    num_val_cols = [c for c in val_cols
                    if df[c].dtype.kind in ("i", "f") and df[c].notna().any()]
    df = df[key_cols + num_val_cols].copy()

    # ⑤ 날짜 변환 (weekly)
    if "weekly" in key_cols and "weekly" in df.columns:
        df["weekly"] = pd.to_datetime(df["weekly"], errors="coerce")
    # monthly → "YYYY-MM" 형식 통일 ("2026. 4. 1" / "2026-04" 모두 처리)
    if "monthly" in key_cols and "monthly" in df.columns:
        def _normalize_monthly(val):
            s = str(val).strip()
            # "2026-04" 형식이면 그대로
            if len(s) == 7 and s[4] == "-":
                return s
            # "2026. 4. 1" 또는 "2026.4.1" 형식 → "2026-04"
            cleaned = s.replace(" ", "")
            if "." in cleaned:
                parts = cleaned.split(".")
                if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
                    return f"{parts[0]}-{int(parts[1]):02d}"
            # pd.to_datetime 시도
            try:
                dt = pd.to_datetime(s)
                return dt.strftime("%Y-%m")
            except Exception:
                return s
        df["monthly"] = df["monthly"].apply(_normalize_monthly)

    # ⑥ 매니저명 정규화
    if "manager_name" in key_cols and "manager_name" in df.columns:
        df["manager_name"] = df["manager_name"].apply(normalize_manager)

    # ⑦ NaT / NaN 키 행 제거 (집계 불가)
    df = df.dropna(subset=key_cols)

    # ⑧ groupby().sum() — SUMIFS 재현
    if num_val_cols:
        df = df.groupby(key_cols, as_index=False)[num_val_cols].sum()
    else:
        df = df.drop_duplicates(subset=key_cols)

    # ⑨ 접두사 부착 (키 컬럼은 제외)
    if prefix:
        df = df.rename(columns={c: f"{prefix}{c}" for c in df.columns if c not in key_cols})

    print(f"    [{prefix.rstrip('_') or '?'}] 정제 후 {df.shape[0]}행 × {df.shape[1]}열  "
          f"컬럼={list(df.columns)}")
    return df


# ============================================================
# 5. 블록 병합 (Merge) — Unique Naming 방식
# ============================================================
def merge_all_blocks(blocks: dict) -> tuple:
    """
    Returns
    -------
    weekly_mgr  : 주별 × 매니저 병합 테이블
    monthly_mgr : 월별 × 매니저 병합 테이블
    """
    print("\n  [Pre-Aggregation] 각 블록 정제 중...")

    # ── 필수 블록 검증 ──
    required = ["A", "E"]
    missing = [k for k in required if k not in blocks]
    if missing:
        loaded = list(blocks.keys())
        raise ValueError(
            f"필수 블록 {missing}이(가) 시트에서 로드되지 않았습니다.\n"
            f"  로드된 블록: {loaded}\n"
            f"  [추출] 탭 1행의 블록 레이블(예: 'lead&call >>')이 변경되었는지 확인하세요."
        )

    # ── 블록별 키·리네임·접두사 정의 ──
    WMM = ["weekly", "monthly", "manager_name"]   # weekly+monthly+manager
    WM  = ["weekly", "monthly"]                    # weekly+monthly (집계 블록)
    MM  = ["monthly", "manager_name"]              # monthly+manager

    a = _sanitize_block(blocks["A"], key_cols=WMM, prefix="A_")
    e = _sanitize_block(blocks["E"], key_cols=WMM, prefix="E_")
    h = _sanitize_block(blocks["H"], key_cols=["weekly", "manager_name"], prefix="H_") if "H" in blocks else None

    # Block I: 원본 컬럼명이 month_kst / week_monday_kst 일 수 있음
    i = _sanitize_block(blocks["I"], key_cols=WMM,
                        rename_map={"month_kst": "monthly", "week_monday_kst": "weekly"},
                        prefix="I_") if "I" in blocks else None

    b = _sanitize_block(blocks["B"], key_cols=WM, prefix="B_") if "B" in blocks else None

    # Block C: 원본 컬럼명이 months / weeks 일 수 있음
    c = _sanitize_block(blocks["C"], key_cols=WM,
                        rename_map={"months": "monthly", "weeks": "weekly"},
                        prefix="C_") if "C" in blocks else None

    d = _sanitize_block(blocks["D"], key_cols=WM, prefix="D_") if "D" in blocks else None

    # Block F: lead_month→monthly, lead_week→weekly
    f = _sanitize_block(blocks["F"], key_cols=WMM,
                        rename_map={"lead_month": "monthly", "lead_week": "weekly"},
                        prefix="F_") if "F" in blocks else None

    g = _sanitize_block(blocks["G"], key_cols=MM, prefix="G_") if "G" in blocks else None

    # ══════════════════════════════════════════════════════════
    # 주별 × 매니저 병합
    # 접두사가 모두 달라서 컬럼 충돌 0%
    # ══════════════════════════════════════════════════════════
    print("\n  [Merge] 주별 × 매니저 병합 시작...")
    weekly_mgr = a  # 기준 테이블

    weekly_mgr = weekly_mgr.merge(e, on=WMM, how="left")
    if h is not None:
        weekly_mgr = weekly_mgr.merge(h, on=["weekly", "manager_name"], how="left")
    if i is not None:
        weekly_mgr = weekly_mgr.merge(i, on=WMM, how="left")
    if b is not None:
        weekly_mgr = weekly_mgr.merge(b, on=WM, how="left")
    if c is not None:
        weekly_mgr = weekly_mgr.merge(c, on=WM, how="left")
    if d is not None:
        weekly_mgr = weekly_mgr.merge(d, on=WM, how="left")
    if f is not None:
        weekly_mgr = weekly_mgr.merge(f, on=WMM, how="left")

    print(f"  [Merge] weekly_mgr 완료: {weekly_mgr.shape}")

    # ══════════════════════════════════════════════════════════
    # 월별 × 매니저 병합
    # ══════════════════════════════════════════════════════════
    # Block A 월별 합산
    a_num = [c for c in a.columns if c.startswith("A_")]
    monthly_mgr = a.groupby(MM, as_index=False)[a_num].sum()

    # Block G 조인
    if g is not None:
        monthly_mgr = monthly_mgr.merge(g, on=MM, how="left")

    # Block I 월별 합산 후 조인
    if i is not None:
        i_num = [c for c in i.columns if c.startswith("I_")]
        if i_num:
            i_monthly = i.groupby(MM, as_index=False)[i_num].sum()
            # 월별용 접두사 변경 (I_ → IM_)
            i_monthly = i_monthly.rename(columns={c: c.replace("I_", "IM_") for c in i_num})
            monthly_mgr = monthly_mgr.merge(i_monthly, on=MM, how="left")

    print(f"  [Merge] monthly_mgr 완료: {monthly_mgr.shape}")

    # ══════════════════════════════════════════════════════════
    # 편의 앨리어스: 접두사 → 기존 KPI 함수가 기대하는 이름
    # (dashboard.py / compute_kpis 와 호환 유지)
    # ══════════════════════════════════════════════════════════
    alias = {
        # Block A
        "A_lead_cnt":            "lead_cnt",
        "A_clean_lead_cnt":      "clean_lead_cnt",
        "A_consulted_cnt":       "consulted_cnt",
        "A_success_cnt":         "success_cnt",
        "A_total_call_attempts": "total_call_attempts",
        "A_mobile_call_attempts":"mobile_call_attempts",
        "A_phone_connected_cnt": "phone_connected_cnt",
        "A_total_paid_cost":     "total_paid_cost",
        "A_coupon_usage":        "coupon_usage",
        # Block E
        "E_total_calls":              "total_calls",
        "E_mobile_010_calls":         "mobile_010_calls",
        "E_non_mobile_calls":         "non_mobile_calls",
        "E_success_calls":            "success_calls",
        "E_fail_calls":               "fail_calls",
        "E_distinct_numbers":         "distinct_numbers",
        "E_attempts_per_number":      "attempts_per_number",
        "E_total_duration_sum":       "total_duration_sum",
        "E_success_real_duration_sum":"success_real_duration_sum",
        "E_effort_seconds_est":       "effort_seconds_est",
        "E_mobile_010_success_calls": "mobile_010_success_calls",
        # Block H (시트 컬럼: paid_cost, order_cnt, coupon_cnt — _w 접미사 없음)
        "H_paid_cost":    "paid_cost_w",
        "H_order_cnt":    "order_cnt_w",
        "H_coupon_cnt":   "coupon_cnt_w",
        # Block I
        "I_total_refund_amount": "total_refund_amount",
        "I_refund_count":        "refund_count",
        # Block B
        "B_net_booking_in_krw": "net_booking_in_krw",
        # Block C
        "C_newstdt_total": "newstdt_total",
        "C_newstdt_ac":    "newstdt_ac",
        # Block D
        "D_up_cross_selling_count": "up_cross_selling_count",
        # Block F (Push 관리)
        "F_A_Link_Created":  "f_link_created",
        "F_B_Total":         "f_b_total",
        "F_B_Push_Target":   "f_b_push_target",
        "F_B_Push_Done":     "f_b_push_done",
        "F_B_Push_Fail":     "f_b_push_fail",
        "F_C_Total":         "f_c_total",
        "F_C_Push_Target":   "f_c_push_target",
        "F_C_Push_Done":     "f_c_push_done",
        "F_C_Push_Fail":     "f_c_push_fail",
        "F_D_Paid_Success":  "f_paid_success",
        "F_B_Push_Skip":     "f_b_push_skip",
        "F_C_Push_Skip":     "f_c_push_skip",
        # Block G (시트 컬럼: paid_cost, order_cnt, coupon_cnt — _m 접미사 없음)
        "G_paid_cost":    "paid_cost_m",
        "G_order_cnt":    "order_cnt_m",
        "G_coupon_cnt":   "coupon_cnt_m",
    }
    # weekly_mgr 앨리어스
    for src, dst in alias.items():
        if src in weekly_mgr.columns:
            weekly_mgr[dst] = weekly_mgr[src]
    # monthly_mgr 앨리어스
    for src, dst in alias.items():
        if src in monthly_mgr.columns:
            monthly_mgr[dst] = monthly_mgr[src]
    # IM_ 접두사도 앨리어스
    for c in monthly_mgr.columns:
        if c.startswith("IM_"):
            monthly_mgr[c.replace("IM_", "refund_m_")] = monthly_mgr[c]

    # NaN → 0 (환불 등 없는 주차)
    for c in ["total_refund_amount", "refund_count", "paid_cost_w",
              "order_cnt_w", "coupon_cnt_w"]:
        if c in weekly_mgr.columns:
            weekly_mgr[c] = weekly_mgr[c].fillna(0)

    return weekly_mgr, monthly_mgr


# ============================================================
# 6. KPI 계산
# ============================================================
def compute_kpis(df: pd.DataFrame) -> pd.DataFrame:
    """주별 × 매니저 DataFrame에 파생 KPI 컬럼 추가.
    컬럼이 없으면 건너뛰어 KeyError 방지."""
    d = df.copy()

    def _safe(col):
        """컬럼이 있으면 Series 반환, 없으면 0으로 채운 Series"""
        if col in d.columns:
            return pd.to_numeric(d[col], errors="coerce").fillna(0)
        return pd.Series(0, index=d.index, dtype=float)

    def _div(num, den):
        """안전한 나눗셈: den이 0이면 NaN (numpy 배열 연산으로 ZeroDivision 방지)"""
        n = _safe(num) if isinstance(num, str) else num
        de = _safe(den) if isinstance(den, str) else den
        n = np.asarray(n, dtype=float)
        de = np.asarray(de, dtype=float)
        return np.where(de > 0, n / np.where(de == 0, 1, de), np.nan)

    # --- A. 리드 유입 & 진척 ---
    d["non_closed_lead_ratio"] = np.round(
        _div(_safe("clean_lead_cnt") - _safe("consulted_cnt"), "clean_lead_cnt") * 100, 1
    )
    d["invalid_lead_ratio"] = np.round(
        _div(_safe("lead_cnt") - _safe("clean_lead_cnt"), "lead_cnt") * 100, 1
    )

    # --- B. 콜 활동 & 효율 ---
    d["calls_per_lead"] = np.round(_div("total_calls", "clean_lead_cnt"), 1)
    d["call_attempt_ratio"] = np.round(_div("total_call_attempts", "lead_cnt"), 1)
    if "effort_seconds_est" in d.columns:
        d["effort_hours"] = np.where(
            d["effort_seconds_est"].notna(),
            (d["effort_seconds_est"] / 3600).round(1), np.nan,
        )

    # --- C. 전환 & 매출 ---
    d["cvr_total"] = np.round(_div("success_cnt", "lead_cnt") * 100, 1)
    d["cvr_clean"] = np.round(_div("success_cnt", "clean_lead_cnt") * 100, 1)
    d["cvr_consulted"] = np.round(_div("success_cnt", "consulted_cnt") * 100, 1)
    d["coupon_ratio_w"] = np.round(_div("coupon_cnt_w", "order_cnt_w") * 100, 1)

    # --- D. 환불 지표 ---
    d["refund_ratio_w"] = np.round(
        _div("total_refund_amount", "paid_cost_w") * 100, 1
    )

    # --- E. Block F 파생: 리드 관리 비율 ---
    # (B) 결제 링크 접속 후 미관리된 리드 비율 = B_Push_Target / B_Total
    d["b_unmanaged_ratio"] = np.round(
        _div("f_b_push_target", "f_b_total") * 100, 1
    )
    # (C) 결제 진행중 미관리된 리드 비율 = C_Push_Target / C_Total
    d["c_unmanaged_ratio"] = np.round(
        _div("f_c_push_target", "f_c_total") * 100, 1
    )

    # --- F. 시도 비율 (통화시도/리드) ---
    d["call_attempt_per_total_lead"] = np.round(
        _div("total_calls", "lead_cnt") * 100, 1
    )
    d["call_attempt_per_clean_lead"] = np.round(
        _div("total_calls", "clean_lead_cnt") * 100, 1
    )

    return d


# ============================================================
# 6. 피어슨 상관계수 (scipy 없이 numpy로 계산)
# ============================================================
def pearson_r(x, y):
    """numpy 기반 피어슨 상관계수 + t-test p-value 계산"""
    n = len(x)
    if n < 3:
        return np.nan, np.nan

    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)

    mx, my = x.mean(), y.mean()
    dx, dy = x - mx, y - my

    denom = np.sqrt(np.sum(dx**2) * np.sum(dy**2))
    if denom == 0:
        return 0.0, 1.0

    r = np.sum(dx * dy) / denom

    # t-test for significance
    if abs(r) >= 1.0:
        p = 0.0
    else:
        t_stat = r * np.sqrt((n - 2) / (1 - r**2))
        # 간이 p-value 근사 (|t| 기준)
        # 자유도 n-2에서 양측 검정
        df = n - 2
        # Student's t CDF 근사
        p = 2 * (1 - t_cdf_approx(abs(t_stat), df))

    return round(r, 1), round(p, 4)


def t_cdf_approx(t, df):
    """Student's t CDF 간이 근사 (Abramowitz & Stegun)"""
    x = df / (df + t**2)
    if df <= 0:
        return 0.5
    # 정규분포 근사 (df가 작아도 방향성 판단에는 충분)
    z = t * (1 - 1/(4*df)) / math.sqrt(1 + t**2/(2*df))
    return 0.5 * (1 + math.erf(z / math.sqrt(2)))


# ============================================================
# 7. 상관관계 분석 — 콜 시도 횟수 vs 결제 금액
# ============================================================
def correlation_analysis(df: pd.DataFrame):
    """매니저별 콜 시도 ↔ 결제 금액 피어슨 상관계수"""
    print("\n" + "=" * 70)
    print("  [분석 1] 매니저별 콜 시도 횟수 vs 결제 금액 상관관계")
    print("=" * 70)

    # 전체 상관
    valid = df.dropna(subset=["total_calls", "paid_cost_w"])
    if len(valid) >= 3:
        r, p = pearson_r(valid["total_calls"], valid["paid_cost_w"])
        print(f"\n  전체 피어슨 r = {r},  p-value = {p}")
        if p < 0.05:
            print("    -> 통계적으로 유의미한 상관관계가 존재합니다.")
        else:
            print("    -> 유의미한 상관관계가 발견되지 않았습니다 (p >= 0.05).")

    # 매니저별 상관
    print(f"\n  {'매니저':<10} {'r':>8} {'p-value':>10} {'주차수':>6}  해석")
    print("  " + "-" * 55)
    for mgr, grp in valid.groupby("manager_name"):
        if len(grp) >= 3:
            r, p = pearson_r(grp["total_calls"], grp["paid_cost_w"])
            sig = "유의" if p < 0.05 else "비유의"
            direction = "양(+)" if r > 0 else "음(-)"
            print(f"  {mgr:<10} {r:>8} {p:>10} {len(grp):>6}  {direction} / {sig}")
        else:
            print(f"  {mgr:<10} {'N/A':>8} {'N/A':>10} {len(grp):>6}  데이터 부족")


# ============================================================
# 8. 종결 상태값이 아닌 리드 비율 — 이상치 탐지 및 원인 추론
# ============================================================
def anomaly_detection(df: pd.DataFrame):
    """non_closed_lead_ratio가 유독 높은 매니저/주차 식별 + 원인 추론"""
    print("\n" + "=" * 70)
    print("  [분석 2] 종결 상태값이 아닌 리드 비율 — 이상치 & 원인 추론")
    print("=" * 70)

    valid = df.dropna(subset=["non_closed_lead_ratio"]).copy()
    overall_mean = float(valid["non_closed_lead_ratio"].mean())
    overall_std = float(valid["non_closed_lead_ratio"].std())
    threshold = overall_mean + 1.0 * overall_std  # 1 시그마 초과 = '높음'

    print(f"\n  전체 평균: {overall_mean:.1f}%  |  표준편차: {overall_std:.1f}%")
    print(f"  이상치 기준(평균 + 1 시그마): {threshold:.1f}% 초과\n")

    anomalies = valid[valid["non_closed_lead_ratio"] > threshold].sort_values(
        "non_closed_lead_ratio", ascending=False
    )

    if anomalies.empty:
        print("  -> 기준을 초과하는 이상치가 없습니다.")
        return

    print(f"  {'매니저':<10} {'주차':>12} {'비율(%)':>8} {'유효리드':>8} {'상담완료':>8} "
          f"{'콜시도':>8} {'리드당콜':>8}  추론")
    print("  " + "-" * 90)

    for _, row in anomalies.iterrows():
        calls_per = float(row.get("calls_per_lead", 0) or 0)
        ratio = float(row["non_closed_lead_ratio"])

        if calls_per < 1.5:
            reason = ">> 콜 시도 부족 -> 리드 누락 가능성 높음"
        elif calls_per >= 1.5 and ratio > overall_mean + 1.5 * overall_std:
            reason = ">> 콜 충분하나 비율 매우 높음 -> 상담 지연 의심"
        else:
            reason = ">> 콜 시도 있음, 일반적 상담 지연 범위"

        weekly_str = str(row['weekly'])[:10] if hasattr(row['weekly'], 'date') else str(row['weekly'])[:10]
        print(
            f"  {row['manager_name']:<10} "
            f"{weekly_str:>12} "
            f"{ratio:>8.1f} "
            f"{int(row['clean_lead_cnt']):>8} "
            f"{int(row['consulted_cnt']):>8} "
            f"{int(row['total_calls']):>8} "
            f"{calls_per:>8.1f}  "
            f"{reason}"
        )


# ============================================================
# 9. 매니저별 리포트 출력
# ============================================================
def manager_report(df: pd.DataFrame):
    """매니저별 주차 성과를 3개 섹션(리드 진척 / 콜 활동 / 최종 성과)으로 출력"""
    for mgr in sorted(df["manager_name"].unique()):
        sub = df[df["manager_name"] == mgr].sort_values("weekly")
        if sub.empty:
            continue

        print(f"\n{'=' * 80}")
        print(f"  [매니저: {mgr}]")
        print(f"{'=' * 80}")

        # -- A. 리드 진척 --
        print("\n  [A] 리드 진척 (Funnel)")
        headers_a = ["주차", "전체리드", "유효리드", "상담완료", "미종결비율(%)", "허수비율(%)"]
        print(f"  {'  '.join(f'{h:>12}' for h in headers_a)}")
        for _, r in sub.iterrows():
            w = str(r['weekly'])[:10]
            print(f"  {w:>12}  "
                  f"{int(r['lead_cnt']):>12}  {int(r['clean_lead_cnt']):>12}  "
                  f"{int(r['consulted_cnt']):>12}  {float(r['non_closed_lead_ratio']):>12.1f}  "
                  f"{float(r['invalid_lead_ratio']):>12.1f}")

        # -- B. 콜 활동 --
        print("\n  [B] 콜 활동 & 공수 (Activity)")
        headers_b = ["주차", "전체콜", "정상번호콜", "연결성공", "리드당콜", "노동시간(h)"]
        print(f"  {'  '.join(f'{h:>12}' for h in headers_b)}")
        for _, r in sub.iterrows():
            w = str(r['weekly'])[:10]
            print(f"  {w:>12}  "
                  f"{int(r['total_calls']):>12}  {int(r['mobile_010_calls']):>12}  "
                  f"{int(r['success_calls']):>12}  {float(r['calls_per_lead']):>12.1f}  "
                  f"{float(r['effort_hours']):>12.1f}")

        # -- C. 최종 성과 --
        print("\n  [C] 최종 성과 & 전환 (Conversion & Sales)")
        headers_c = ["주차", "결제건수", "CVR전체(%)", "CVR유효(%)", "CVR상담(%)",
                     "결제금액(W)", "환불금액", "환불비율(%)", "쿠폰비율(%)"]
        print(f"  {'  '.join(f'{h:>12}' for h in headers_c)}")
        for _, r in sub.iterrows():
            w = str(r['weekly'])[:10]
            paid = float(r.get("paid_cost_w", 0) or 0)
            refund_amt = float(r.get("total_refund_amount", 0) or 0)
            refund_r = float(r.get("refund_ratio_w", 0) or 0)
            coupon_r = float(r.get("coupon_ratio_w", 0) or 0)
            print(f"  {w:>12}  "
                  f"{int(r['success_cnt']):>12}  {float(r['cvr_total']):>12.1f}  "
                  f"{float(r['cvr_clean']):>12.1f}  {float(r['cvr_consulted']):>12.1f}  "
                  f"{int(paid):>12,}  {int(refund_amt):>12,}  "
                  f"{refund_r:>12.1f}  {coupon_r:>12.1f}")


# ============================================================
# 10. 데이터 갭 검토 — 계산 불가 지표 보고
# ============================================================
def data_gap_report():
    """PRD에 정의되었으나 현재 추출 탭 데이터로 계산 불가능한 지표 목록"""
    print("\n" + "=" * 70)
    print("  [검토] 데이터 갭 현황")
    print("=" * 70)

    print("\n  [해결됨] 환불 금액 및 비율")
    print("           -> Block I (REFUND) 추가로 계산 가능해짐")
    print("           -> refund_ratio_w = total_refund_amount / paid_cost_w * 100")

    print("\n  [제외됨] B_Push_Skip, C_Push_Skip (미관리 리드)")
    print("           -> Block F 미정의 + 팀 결정에 따라 분석 대상에서 제외")

    print("\n  -> 현재 모든 핵심 지표가 계산 가능합니다.")


# ============================================================
# 11. 메인 실행
# ============================================================
def main():
    print("=" * 70)
    print("  CA Dashboard — 매니저별 종합 성과 분석 리포트")
    print("  분석 일자: 2026-04-06  |  소수점 첫째 자리 반올림 적용")
    print("=" * 70)
    print(f"\n  데이터 소스: 비공개 구글 시트 (gspread OAuth 2.0)")
    print(f"  Sheet ID: {SHEET_ID}")
    print(f"  GID: {SHEET_GID}")
    print(f"  인증 파일: {CREDENTIALS_FILE}")

    # 1) 데이터 로드
    blocks, source = load_all_blocks()

    # 2) 블록 병합
    print("\n[2단계] 블록 병합 (Merge)")
    print("-" * 40)
    weekly_mgr, monthly_mgr = merge_all_blocks(blocks)
    print(f"  주별 병합 테이블: {weekly_mgr.shape[0]}행 × {weekly_mgr.shape[1]}열")
    print(f"  월별 병합 테이블: {monthly_mgr.shape[0]}행 × {monthly_mgr.shape[1]}열")

    # 3) KPI 계산
    print("\n[3단계] KPI 계산")
    print("-" * 40)
    weekly_kpi = compute_kpis(weekly_mgr)
    print("  -> 파생 KPI 컬럼 추가 완료")

    # 4) 상관관계 분석
    correlation_analysis(weekly_kpi)

    # 5) 이상치 탐지
    anomaly_detection(weekly_kpi)

    # 6) 매니저별 리포트
    print("\n" + "=" * 70)
    print("  [분석 3] 매니저별 주간 성과 상세 리포트")
    print("=" * 70)
    manager_report(weekly_kpi)

    # 7) 데이터 갭 보고
    data_gap_report()

    print("\n" + "=" * 70)
    print("  분석 완료")
    print("=" * 70)


if __name__ == "__main__":
    main()
