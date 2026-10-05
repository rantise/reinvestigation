import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from html import escape


# =========================================================
# 1. Strategy Pattern 정의 (분석 전략)
# =========================================================

class AreaAnalysisStrategy:
    """면적 계산 및 증감 분석을 담당하는 전략 클래스"""

    @staticmethod
    def calculate(df):
        df = df.copy()

        # 면적 숫자 변환
        df["종전_면적"] = pd.to_numeric(
            df["종전_면적"],
            errors="coerce"
        ).fillna(0)

        df["확정_면적"] = pd.to_numeric(
            df["확정_면적"],
            errors="coerce"
        ).fillna(0)

        # 종전 ID 생성
        df["종전_ID"] = (
            df["종전_읍면"].fillna("").astype(str).str.strip()
            + "_"
            + df["종전_동리"].fillna("").astype(str).str.strip()
            + "_"
            + df["종전_지목"].fillna("").astype(str).str.strip()
            + "_"
            + df["종전_지번"].fillna("").astype(str).str.strip()
        )

        # 확정 ID 생성
        df["확정_ID"] = (
            df["확정_읍면"].fillna("").astype(str).str.strip()
            + "_"
            + df["확정_동리"].fillna("").astype(str).str.strip()
            + "_"
            + df["확정_지목"].fillna("").astype(str).str.strip()
            + "_"
            + df["확정_지번"].fillna("").astype(str).str.strip()
        )

        # 종전 면적 집계
        prev_area_map = (
            df.groupby("종전_ID")["종전_면적"]
            .sum()
            .to_dict()
        )

        # 확정 ID 기준으로 종전 총면적 매칭
        df["종전_총면적"] = (
            df["확정_ID"]
            .map(prev_area_map)
            .fillna(0)
        )

        # 증가 면적
        df["증감면적"] = np.where(
            df["확정_면적"] > df["종전_총면적"],
            df["확정_면적"] - df["종전_총면적"],
            0
        )

        # 감소 면적
        df["감소면적"] = np.where(
            df["종전_총면적"] > df["확정_면적"],
            df["종전_총면적"] - df["확정_면적"],
            0
        )

        # 업로드 엑셀의 면적 정밀도(소수 첫째 자리)에 맞춰
        # 계산 결과도 소수 첫째 자리까지 반올림합니다.
        df["종전_면적"] = df["종전_면적"].round(1)
        df["확정_면적"] = df["확정_면적"].round(1)
        df["종전_총면적"] = df["종전_총면적"].round(1)
        df["증감면적"] = df["증감면적"].round(1)
        df["감소면적"] = df["감소면적"].round(1)

        # 증감 여부
        df["증감여부"] = np.where(
            df["증감면적"] > 0,
            "증가",
            np.where(
                df["감소면적"] > 0,
                "감소",
                "변동없음"
            )
        )

        # 처리 상태
        df["처리상태"] = np.where(
            df["확정_면적"] == 0,
            "말소",
            "유지"
        )

        return df


# =========================================================
# 2. 소유권 검증 전략
# =========================================================

class OwnershipValidationStrategy:
    """소유권 유형별 검증 및 지분 합계 체크"""

    @staticmethod
    def validate_share(owner_data, ownership_type):
        """공유/상속 지분 합계를 검증하고 요청된 문구를 반환합니다."""
        if ownership_type not in ["공유", "상속"]:
            return None

        total_share = sum(
            item.get("지분", 0.0)
            for item in owner_data
        )

        if np.isclose(total_share, 1.0, atol=1e-5):
            return None

        if total_share > 1.0:
            return "지분의 합이 1이 아닙니다.(1 초과)"

        return "지분의 합이 1이 아닙니다.(1 미만)"


# =========================================================
# 3. 엑셀 관련 함수
# =========================================================

def convert_excel_columns(df):
    """
    엑셀의 2단 헤더를 프로그램 내부 컬럼명으로 변환
    """

    new_columns = []

    for col in df.columns:

        if isinstance(col, tuple):
            first = str(col[0]).strip()
            second = str(col[1]).strip()
        else:
            first = str(col).strip()
            second = ""

        if first == "종전":
            new_columns.append(
                f"종전_{second}"
            )

        elif first == "확정":
            new_columns.append(
                f"확정_{second}"
            )

        elif first == "기준년도":
            new_columns.append(
                "기준년도"
            )

        elif first == "사업지구명":
            new_columns.append(
                "사업지구명"
            )

        elif first == "소유사항":
            new_columns.append(
                "소유사항"
            )

        elif first == "공유 또는 상속인 수":
            new_columns.append(
                "공유_상속인수"
            )

        else:
            new_columns.append(first)

    df = df.copy()
    df.columns = new_columns

    return df


def create_excel_download(df):
    """계산 결과를 엑셀 파일 데이터로 변환"""

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="계산결과"
        )

    return output.getvalue()


# =========================================================
# 4. Streamlit 기본 설정
# =========================================================

st.set_page_config(
    page_title="토지 면적 증감 및 소유 분석 프로그램",
    layout="wide"
)

st.title(
    "🏡 토지 면적 증감 및 소유 분석 프로그램"
)


# =========================================================
# 5. 소유구분 맵핑
# =========================================================

OWNERSHIP_OPTIONS = [
    "1. 개인",
    "2. 법인",
    "3. 국유지",
    "4. 시유지",
    "5. 시군구유지"
]

# 예:
# "1. 개인" -> "개인"
OWNERSHIP_MAP = {
    option: option.split(". ", 1)[1]
    for option in OWNERSHIP_OPTIONS
}

# 예:
# "1" -> "개인"
OWNERSHIP_CODE_MAP = {
    option.split(". ", 1)[0]:
    option.split(". ", 1)[1]
    for option in OWNERSHIP_OPTIONS
}


# =========================================================
# 6. Session State
# =========================================================

if "collapsed_rows" not in st.session_state:
    st.session_state["collapsed_rows"] = {}


# =========================================================
# 7. 탭 생성
# =========================================================

tab1, tab2 = st.tabs(
    [
        "📥 엑셀 파일 업로드",
        "📝 직접 입력 모드"
    ]
)


# =========================================================
# TAB 1
# 엑셀 파일 업로드
# =========================================================

with tab1:

    uploaded_file = st.file_uploader(
        "엑셀 파일을 선택하세요.",
        type=["xlsx"]
    )

    if uploaded_file:

        try:
            df_input = pd.read_excel(
                uploaded_file,
                header=[0, 1]
            )

            df_input = convert_excel_columns(df_input)

            st.subheader("📄 업로드 데이터 미리보기")
            st.dataframe(
                df_input.head(),
                use_container_width=True
            )

            if st.button(
                "엑셀 데이터로 계산 시작",
                key="calculate_excel"
            ):

                required_columns = [
                    "종전_읍면", "종전_동리", "종전_지목", "종전_지번",
                    "종전_면적", "확정_읍면", "확정_동리", "확정_지목",
                    "확정_지번", "확정_면적"
                ]

                missing_columns = [
                    col for col in required_columns
                    if col not in df_input.columns
                ]

                if missing_columns:
                    st.error(
                        "필수 컬럼이 누락되었습니다.\n\n"
                        + ", ".join(missing_columns)
                    )
                else:
                    calculated_df = AreaAnalysisStrategy.calculate(df_input)

                    # 내부 계산용 컬럼은 화면/다운로드에서 숨김
                    result_df = calculated_df.drop(
                        columns=[
                            c for c in [
                                "종전_ID",
                                "확정_ID",
                                "종전_총면적"
                            ]
                            if c in calculated_df.columns
                        ],
                        errors="ignore"
                    ).copy()

                    st.session_state["excel_result_df"] = result_df
                    st.session_state["excel_result_source_df"] = calculated_df
                    st.session_state["excel_owner_data"] = {}
                    st.session_state["excel_owner_collapsed"] = {}

                    st.success("✅ 계산이 완료되었습니다.")

            if "excel_result_df" in st.session_state:

                result_df = st.session_state["excel_result_df"].copy()
                st.subheader("📊 계산 결과 및 소유자 입력")

                # =====================================================
                # test3.xlsx 스타일 결과표
                # =====================================================
                # 중요: 계산 결과(result_df)의 원본 값을 그대로 사용합니다.
                # 이전 버전에서 HTML 셀만 따로 렌더링하면서 값이 공란처럼
                # 보이는 문제가 있었으므로, 결과값은 Streamlit 셀에 직접
                # 출력하고 헤더와 동일한 열 비율을 사용합니다.

                def parse_share(value):
                    if value is None or pd.isna(value):
                        return 0.0
                    text = str(value).strip()
                    if not text:
                        return 0.0
                    try:
                        if "/" in text:
                            a, b = text.split("/", 1)
                            denominator = float(b.strip())
                            if denominator == 0:
                                return 0.0
                            return float(a.strip()) / denominator
                        return float(text)
                    except (ValueError, TypeError, ZeroDivisionError):
                        return 0.0

                def fmt_value(value):
                    if value is None or pd.isna(value):
                        return ""
                    if isinstance(value, (float, np.floating)):
                        # 144.10000000002 같은 부동소수점 표현을 방지
                        return f"{float(value):.1f}"
                    if isinstance(value, (int, np.integer)):
                        return str(int(value))
                    return str(value)

                def normalize_ownership(value):
                    text = "" if value is None or pd.isna(value) else str(value).strip()
                    mapping = {"1": "1. 단독", "2": "2. 공유", "3": "3. 상속"}
                    if text in mapping:
                        return mapping[text]
                    if text in ["단독", "공유", "상속"]:
                        return {"단독": "1. 단독", "공유": "2. 공유", "상속": "3. 상속"}[text]
                    if text in ["1. 단독", "2. 공유", "3. 상속"]:
                        return text
                    return "1. 단독"

                OWNERSHIP_TYPE_OPTIONS = ["1. 단독", "2. 공유", "3. 상속"]

                if "excel_owner_data" not in st.session_state:
                    st.session_state["excel_owner_data"] = {}
                if "excel_owner_collapsed" not in st.session_state:
                    st.session_state["excel_owner_collapsed"] = {}
                if "share_dialog_signature" not in st.session_state:
                    st.session_state["share_dialog_signature"] = None
                if "pending_share_dialog" not in st.session_state:
                    st.session_state["pending_share_dialog"] = None

                if hasattr(st, "dialog"):
                    @st.dialog("지분 합계 확인")
                    def share_error_dialog(message):
                        st.error(message)
                        st.write("입력한 지분을 확인해 주세요.")

                # 열 폭은 헤더와 데이터 행에서 완전히 동일하게 사용합니다.
                # 전체 표 폭은 v5보다 줄이고, 지번은 1324-567 정도가 들어가도록
                # 확보하며, 지분 열은 기존(v5)보다 2배 넓게 잡습니다.
                col_widths = [
                    0.80, 1.25,
                    0.70, 1.15, 0.80, 1.40, 0.80,
                    0.70, 1.15, 0.80, 1.40, 0.80,
                    0.90, 0.90, 0.85, 0.85,
                    0.95, 1.15,
                    1.10, 1.20, 1.80
                ]
                total_col_width = sum(col_widths)
                result_table_width = 1900
                header_colgroup = "".join(
                    f'<col style="width:{w / total_col_width * 100:.4f}%">'
                    for w in col_widths
                )

                static_keys = [
                    "기준년도", "사업지구명",
                    "종전_읍면", "종전_동리", "종전_지목", "종전_지번", "종전_면적",
                    "확정_읍면", "확정_동리", "확정_지목", "확정_지번", "확정_면적",
                    "증감면적", "감소면적", "증감여부", "처리상태"
                ]

                def render_value_cell(container, value, extra_class=""):
                    """값이 실제로 보이는 일반 표 셀."""
                    safe_value = escape(fmt_value(value))
                    with container:
                        st.markdown(
                            f'<div class="land-value-cell {extra_class}" style="color:#222 !important;">{safe_value}</div>',
                            unsafe_allow_html=True,
                        )

                # 화면 폭을 넓혀 헤더와 데이터의 세로선을 정확히 맞춥니다.
                st.markdown(
                    """
                    <style>
                    /* 결과표 전체: 가로 스크롤은 한 곳에서만 */
                    div.st-key-result_table {
                        overflow-x: auto !important;
                        overflow-y: visible !important;
                        width: 100% !important;
                        max-width: 100% !important;
                        padding: 0 0 8px 0 !important;
                    }
                    div.st-key-result_table > div {
                        min-width: 1900px !important;
                    }
                    div.st-key-result_table .stHorizontalBlock {
                        min-width: 1900px !important;
                        gap: 0 !important;
                        margin-top: 0 !important;
                        margin-bottom: 0 !important;
                        padding-top: 0 !important;
                        padding-bottom: 0 !important;
                        align-items: stretch !important;
                    }
                    div.st-key-result_table .stHorizontalBlock + .stHorizontalBlock {
                        margin-top: -1px !important;
                    }
                    /* 행과 행 사이에 Streamlit이 남기는 기본 여백을 제거 */
                    div.st-key-result_table .stHorizontalBlock {
                        margin-bottom: -7px !important;
                    }
                    div.st-key-result_table .stHorizontalBlock:last-child {
                        margin-bottom: 0 !important;
                    }
                    /* Streamlit의 행별 wrapper가 만드는 세로 간격까지 제거하여
                       업로드 미리보기(st.dataframe)처럼 행을 서로 맞닿게 합니다. */
                    div.st-key-result_table [data-testid="stVerticalBlock"],
                    div.st-key-result_table [data-testid="stVerticalBlockBorderWrapper"],
                    div.st-key-result_table [data-testid="stHorizontalBlock"],
                    div.st-key-result_table [data-testid="column"],
                    div.st-key-result_table [data-testid="column"] > div,
                    div.st-key-result_table [data-testid="stVerticalBlock"] > div,
                    div.st-key-result_table [data-testid="stHorizontalBlock"] > div {
                        margin-top: 0 !important;
                        margin-bottom: 0 !important;
                        padding-top: 0 !important;
                        padding-bottom: 0 !important;
                    }
                    div.st-key-result_table [data-testid="stVerticalBlock"],
                    div.st-key-result_table [data-testid="stVerticalBlock"] > div,
                    div.st-key-result_table [data-testid="stHorizontalBlock"] {
                        gap: 0 !important;
                        row-gap: 0 !important;
                    }
                    /* 각 결과행의 Streamlit 기본 블록 간격을 완전히 상쇄합니다. */
                    div.st-key-result_table .land-row-block {
                        margin: 0 !important;
                        padding: 0 !important;
                        line-height: 0 !important;
                        height: 40px !important;
                        min-height: 40px !important;
                        overflow: visible !important;
                    }
                    div.st-key-result_table [data-testid="stElementContainer"] {
                        margin: 0 !important;
                        padding: 0 !important;
                    }
                    div.st-key-result_table [data-testid="stMarkdownContainer"] {
                        margin: 0 !important;
                        padding: 0 !important;
                    }
                    div.st-key-result_table [data-testid="column"] {
                        padding-left: 0 !important;
                        padding-right: 0 !important;
                        min-width: 0 !important;
                    }
                    div.st-key-result_table [data-testid="stVerticalBlock"] {
                        gap: 0 !important;
                    }

                    .land-header-table {
                        width: 1900px;
                        min-width: 1900px;
                        border-collapse: collapse;
                        table-layout: fixed;
                        margin: 0;
                        font-size: 13px;
                    }
                    .land-header-table th {
                        box-sizing: border-box;
                        border: 1px solid #b7b7b7;
                        background: #eeeeee;
                        text-align: center;
                        vertical-align: middle;
                        height: 34px;
                        padding: 4px 2px;
                        font-weight: 700;
                        white-space: normal;
                        overflow-wrap: anywhere;
                        color: #202020;
                    }
                    .land-header-table th.group {
                        background: #e2e2e2;
                    }

                    /* 실제 데이터 셀 */
                    .land-value-cell {
                        box-sizing: border-box;
                        width: 100%;
                        min-height: 40px;
                        height: 40px;
                        margin: 0 !important;
                        border: 1px solid #c8c8c8;
                        background: #ffffff;
                        color: #222222 !important;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        padding: 4px 3px;
                        font-size: 13px;
                        line-height: 1.2;
                        white-space: nowrap;
                        overflow: hidden;
                        text-overflow: ellipsis;
                    }
                    .land-value-cell.left { justify-content: flex-start; }
                    .land-value-cell.right { justify-content: flex-end; }
                    .land-value-cell.share-error {
                        color: #d00000 !important;
                        background: #ffe4e4;
                        border: 2px solid #ff0000;
                        font-weight: 700;
                    }

                    /* 입력 위젯도 일반 셀과 동일한 높이 */
                    div.st-key-result_table input,
                    div.st-key-result_table [data-baseweb="select"] > div,
                    div.st-key-result_table [data-testid="stNumberInput"] input {
                        min-height: 40px !important;
                        height: 40px !important;
                        border-radius: 0 !important;
                        box-sizing: border-box !important;
                    }
                    div.st-key-result_table input {
                        padding-left: 1.5px !important;
                        padding-right: 1.5px !important;
                    }
                    div.st-key-result_table [data-testid="stTextInputRootElement"],
                    div.st-key-result_table [data-testid="stSelectbox"],
                    div.st-key-result_table [data-testid="stNumberInput"] {
                        min-height: 40px !important;
                        height: 40px !important;
                        margin: 0 !important;
                        padding: 0 !important;
                    }
                    div.st-key-result_table label,
                    div.st-key-result_table [data-testid="stWidgetLabel"] {
                        display: none !important;
                    }

                    /* ┗ 버튼: 공유/상속 추가행 첫 번째 소유구분 바로 왼쪽 */
                    div.st-key-result_table [class*="st-key-fold_cell_"] {
                        width: 100% !important;
                        min-height: 40px !important;
                        height: 40px !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: flex-end !important;
                        box-sizing: border-box !important;
                    }
                    div.st-key-result_table [class*="st-key-fold_cell_"] button {
                        display: block !important;
                        margin-left: auto !important;
                        min-width: 34px !important;
                        width: 34px !important;
                        height: 30px !important;
                        margin: 3px 2px 3px auto !important;
                        padding: 0 !important;
                        border: 1px solid #8c8c8c !important;
                        border-radius: 6px !important;
                        background: linear-gradient(#ffffff, #e2e2e2) !important;
                        color: #202020 !important;
                        font-size: 18px !important;
                        font-weight: 800 !important;
                        box-shadow: 0 2px 0 #888, 0 3px 5px rgba(0,0,0,.16) !important;
                        transform: translateY(-1px);
                    }
                    div.st-key-result_table [class*="st-key-fold_cell_"] button:hover {
                        background: linear-gradient(#f8f8f8, #d8d8d8) !important;
                    }
                    div.st-key-result_table [class*="st-key-fold_cell_"] button:active {
                        transform: translateY(1px) !important;
                        box-shadow: 0 1px 0 #888 !important;
                    }

                    /* 공유/상속 하위행의 좌측 병합 영역 */
                    .land-fold-cell {
                        min-height: 40px;
                        height: 40px;
                        width: 100%;
                        box-sizing: border-box;
                        border: 1px solid #c8c8c8;
                        background: #ffffff;
                        display: flex;
                        align-items: center;
                        justify-content: flex-end;
                        padding: 0 3px 0 0;
                    }
                    .land-fold-merged {
                        min-height: 40px;
                        height: 40px;
                        width: 100%;
                        box-sizing: border-box;
                        border: 1px solid #c8c8c8;
                        background: #ffffff;
                        display: flex;
                        align-items: center;
                        justify-content: flex-end;
                        padding: 0 3px 0 0;
                    }
                    /* 공유/상속인 수 -/+ 버튼 */
                    div.st-key-result_table [class*="st-key-count_controls_"] button {
                        min-width: 24px !important;
                        width: 24px !important;
                        height: 30px !important;
                        padding: 0 !important;
                        margin: 0 !important;
                        border-radius: 5px !important;
                        font-weight: 800 !important;
                        font-size: 16px !important;
                    }
                    .land-count-number {
                        min-width: 28px;
                        text-align: center;
                        font-size: 14px;
                        line-height: 30px;
                    }

                    .land-owner-input {
                        width: 100%;
                        min-height: 40px;
                        height: 40px;
                        box-sizing: border-box;
                    }
                    </style>
                    """,
                    unsafe_allow_html=True,
                )

                # 결과표를 하나의 컨테이너로 묶습니다.
                table = st.container(key="result_table", border=True)
                owner_changed = False

                with table:
                    # -------------------------------------------------
                    # 2단 헤더: test3.xlsx와 같은 병합 구조
                    # -------------------------------------------------
                    st.markdown(
                        f"""
                        <table class="land-header-table">
                          <colgroup>{header_colgroup}</colgroup>
                          <tr>
                            <th rowspan="2">기준년도</th>
                            <th rowspan="2">사업지구명</th>
                            <th colspan="5" class="group">종전</th>
                            <th colspan="5" class="group">확정</th>
                            <th rowspan="2">증감면적</th>
                            <th rowspan="2">감소면적</th>
                            <th rowspan="2">증감여부</th>
                            <th rowspan="2">처리상태</th>
                            <th rowspan="2">소유사항</th>
                            <th rowspan="2">공유 또는 상속인 수</th>
                            <th rowspan="2">소유구분</th>
                            <th rowspan="2">토지소유자</th>
                            <th rowspan="2">지분</th>
                          </tr>
                          <tr>
                            <th>읍면</th><th>동리</th><th>지목</th><th>지번</th><th>면적</th>
                            <th>읍면</th><th>동리</th><th>지목</th><th>지번</th><th>면적</th>
                          </tr>
                        </table>
                        """,
                        unsafe_allow_html=True,
                    )

                    # -------------------------------------------------
                    # 결과 데이터 행
                    # -------------------------------------------------
                    # 각 행은 동일 높이의 Streamlit 블록으로 렌더링하되 CSS에서
                    # 블록 간 기본 여백을 제거하여 업로드 미리보기처럼 붙입니다.
                    for idx, row in result_df.iterrows():
                        raw_type = row.get("소유사항", "")
                        ownership_type = normalize_ownership(raw_type)

                        raw_count = row.get("공유_상속인수", np.nan)
                        if raw_count is None or pd.isna(raw_count):
                            count = 0
                        else:
                            try:
                                count = max(0, int(float(raw_count)))
                            except (TypeError, ValueError):
                                count = 0

                        if ownership_type == "1. 단독":
                            count = 0

                        owner_key = f"excel_owner_{idx}"
                        required = max(1, count)
                        if owner_key not in st.session_state["excel_owner_data"]:
                            st.session_state["excel_owner_data"][owner_key] = [
                                {"소유구분": "1. 개인", "토지소유자": "", "지분": ""}
                                for _ in range(required)
                            ]
                        else:
                            owners = st.session_state["excel_owner_data"][owner_key]
                            if len(owners) < required:
                                owners.extend(
                                    {"소유구분": "1. 개인", "토지소유자": "", "지분": ""}
                                    for _ in range(required - len(owners))
                                )
                            elif len(owners) > required:
                                st.session_state["excel_owner_data"][owner_key] = owners[:required]

                        owners = st.session_state["excel_owner_data"][owner_key]

                        # 지분 합계는 공유/상속 기본행에 표시합니다.
                        # 단, 모든 공유/상속 소유자의 지분이 입력되기 전에는
                        # 1 초과/미만 오류로 판정하지 않습니다.
                        def is_valid_share_input(value):
                            if value is None or (isinstance(value, float) and pd.isna(value)):
                                return False
                            text = str(value).strip()
                            if not text:
                                return False
                            try:
                                if "/" in text:
                                    a, b = text.split("/", 1)
                                    denominator = float(b.strip())
                                    if denominator == 0:
                                        return False
                                    float(a.strip())
                                else:
                                    float(text)
                                return True
                            except (ValueError, TypeError, ZeroDivisionError):
                                return False

                        share_complete = (
                            count > 0
                            and len(owners) >= count
                            and all(is_valid_share_input(o.get("지분", "")) for o in owners[:count])
                        )
                        total_share = (
                            sum(parse_share(o.get("지분", "")) for o in owners[:count])
                            if count > 0 else 1.0
                        )
                        invalid_share = share_complete and not np.isclose(total_share, 1.0, atol=1e-5)

                        # -------------------------
                        # 기본행: 21개 열
                        # -------------------------
                        cells = st.columns(col_widths, gap=None)
                        values = [
                            row.get("기준년도", ""),
                            row.get("사업지구명", ""),
                            row.get("종전_읍면", ""),
                            row.get("종전_동리", ""),
                            row.get("종전_지목", ""),
                            row.get("종전_지번", ""),
                            row.get("종전_면적", ""),
                            row.get("확정_읍면", ""),
                            row.get("확정_동리", ""),
                            row.get("확정_지목", ""),
                            row.get("확정_지번", ""),
                            row.get("확정_면적", ""),
                            row.get("증감면적", ""),
                            row.get("감소면적", ""),
                            row.get("증감여부", ""),
                            row.get("처리상태", ""),
                        ]

                        for pos, value in enumerate(values):
                            render_value_cell(cells[pos], value)

                        # 소유사항: 업로드 값이 있어도 수정 가능
                        with cells[16]:
                            selected_type = st.selectbox(
                                "소유사항",
                                OWNERSHIP_TYPE_OPTIONS,
                                index=OWNERSHIP_TYPE_OPTIONS.index(ownership_type),
                                key=f"excel_ownership_type_{idx}",
                                label_visibility="collapsed",
                            )

                        selected_plain = selected_type.split(". ", 1)[1]
                        old_plain = ownership_type.split(". ", 1)[1]

                        if selected_plain != old_plain:
                            # 소유사항 변경 시 기존 결과 데이터도 즉시 갱신
                            st.session_state["excel_result_df"].loc[idx, "소유사항"] = selected_plain
                            if "excel_result_source_df" in st.session_state:
                                st.session_state["excel_result_source_df"].loc[idx, "소유사항"] = selected_plain
                            owner_changed = True

                            if selected_plain == "단독":
                                st.session_state["excel_result_df"].loc[idx, "공유_상속인수"] = 0
                                st.session_state["excel_owner_data"][owner_key] = [owners[0] if owners else {"소유구분":"1. 개인","토지소유자":"","지분":""}]
                                count = 0
                            else:
                                count = max(1, count)
                                st.session_state["excel_result_df"].loc[idx, "공유_상속인수"] = count
                                current = st.session_state["excel_owner_data"][owner_key]
                                while len(current) < count:
                                    current.append({"소유구분":"1. 개인","토지소유자":"","지분":""})
                                owners = current

                        # 공유 또는 상속인 수
                        with cells[17]:
                            if selected_plain == "단독":
                                st.number_input(
                                    "공유 또는 상속인 수",
                                    min_value=0, max_value=100, value=0, step=1,
                                    disabled=True,
                                    key=f"excel_count_{idx}_single",
                                    label_visibility="collapsed",
                                )
                                count_value = 0
                            else:
                                # v6 방식: Streamlit 기본 숫자 입력칸으로 직접 인원 수를 입력/증감
                                count_value = st.number_input(
                                    "공유 또는 상속인 수",
                                    min_value=1, max_value=100,
                                    value=max(1, count), step=1,
                                    key=f"excel_count_{idx}_{selected_plain}",
                                    label_visibility="collapsed",
                                )
                                if int(count_value) != int(count):
                                    count = int(count_value)
                                    st.session_state["excel_result_df"].loc[idx, "공유_상속인수"] = count
                                    current = st.session_state["excel_owner_data"][owner_key]
                                    if len(current) < count:
                                        current.extend(
                                            {"소유구분":"1. 개인","토지소유자":"","지분":""}
                                            for _ in range(count - len(current))
                                        )
                                    elif len(current) > count:
                                        st.session_state["excel_owner_data"][owner_key] = current[:count]
                                    owners = st.session_state["excel_owner_data"][owner_key]
                                    owner_changed = True

                        # 단독은 기본행에 소유자 정보를 입력
                        if selected_plain == "단독":
                            owner = owners[0]
                            with cells[18]:
                                owner_type = st.selectbox(
                                    "소유구분", OWNERSHIP_OPTIONS,
                                    index=OWNERSHIP_OPTIONS.index(owner.get("소유구분", "1. 개인"))
                                    if owner.get("소유구분", "1. 개인") in OWNERSHIP_OPTIONS else 0,
                                    key=f"excel_owner_type_{idx}_0",
                                    label_visibility="collapsed",
                                )
                                if owner.get("소유구분") != owner_type:
                                    owner_changed = True
                                owner["소유구분"] = owner_type

                            with cells[19]:
                                new_name = st.text_input(
                                    "토지소유자",
                                    value=owner.get("토지소유자", ""),
                                    key=f"excel_owner_name_{idx}_0",
                                    label_visibility="collapsed",
                                )
                                if owner.get("토지소유자", "") != new_name:
                                    owner_changed = True
                                owner["토지소유자"] = new_name

                            with cells[20]:
                                st.text_input(
                                    "지분", value="1", disabled=True,
                                    key=f"excel_owner_share_{idx}_0",
                                    label_visibility="collapsed",
                                )
                                owner["지분"] = "1"

                        else:
                            # 공유/상속 기본행: 소유구분/소유자/지분은 입력행에서 관리
                            render_value_cell(cells[18], "")
                            render_value_cell(cells[19], "")
                            share_class = "share-error" if invalid_share else ""
                            render_value_cell(cells[20], total_share, share_class)

                        # -------------------------------------------------
                        # 공유/상속 추가행
                        # 좌측 18개 열을 하나의 병합 셀로 만들고
                        # ┗ 버튼을 우측(소유구분 바로 왼쪽)에 배치합니다.
                        # -------------------------------------------------
                        if selected_plain in ["공유", "상속"]:
                            if owner_key not in st.session_state["excel_owner_collapsed"]:
                                st.session_state["excel_owner_collapsed"][owner_key] = False

                            collapsed = st.session_state["excel_owner_collapsed"][owner_key]

                            # 접힌 상태에서도 ┗ 버튼이 남아 있어야 다시 펼칠 수 있습니다.
                            # 펼친 상태에서는 소유자 수만큼 실제 입력행을 표시합니다.
                            render_owner_indices = [0] if collapsed else list(range(count_value))

                            for owner_idx in render_owner_indices:
                                owner = st.session_state["excel_owner_data"][owner_key][owner_idx]

                                # 공유/상속 추가행의 좌측 전체(기준년도~공유/상속인 수)를
                                # 하나의 셀처럼 병합하고, 그 오른쪽 끝에 ┗ 버튼을 둡니다.
                                # 바로 다음 칸이 첫 번째 소유구분이므로 ┗가 소유구분 바로 옆에 옵니다.
                                child_widths = [sum(col_widths[:18]), col_widths[18], col_widths[19], col_widths[20]]
                                child = st.columns(child_widths, gap=None)

                                with child[0]:
                                    if owner_idx == 0:
                                        fold_box = st.container(key=f"fold_cell_{idx}")
                                        with fold_box:
                                            if st.button(
                                                "┗",
                                                key=f"fold_excel_{idx}",
                                                help="공유/상속 소유자 행 접기/펼치기",
                                                use_container_width=False,
                                            ):
                                                st.session_state["excel_owner_collapsed"][owner_key] = not collapsed
                                                st.rerun()
                                    else:
                                        st.markdown('<div class="land-fold-merged"></div>', unsafe_allow_html=True)

                                # 접힌 상태에서는 소유자 입력 3칸을 숨깁니다.
                                if collapsed:
                                    continue

                                with child[1]:
                                    owner_type = st.selectbox(
                                        "소유구분",
                                        OWNERSHIP_OPTIONS,
                                        index=OWNERSHIP_OPTIONS.index(owner.get("소유구분", "1. 개인"))
                                        if owner.get("소유구분", "1. 개인") in OWNERSHIP_OPTIONS else 0,
                                        key=f"excel_child_owner_type_{idx}_{owner_idx}",
                                        label_visibility="collapsed",
                                    )
                                    if owner.get("소유구분") != owner_type:
                                        owner_changed = True
                                    owner["소유구분"] = owner_type

                                with child[2]:
                                    new_name = st.text_input(
                                        "토지소유자",
                                        value=owner.get("토지소유자", ""),
                                        key=f"excel_child_owner_name_{idx}_{owner_idx}",
                                        label_visibility="collapsed",
                                    )
                                    if owner.get("토지소유자", "") != new_name:
                                        owner_changed = True
                                    owner["토지소유자"] = new_name

                                with child[3]:
                                    new_share = st.text_input(
                                        "지분",
                                        value=owner.get("지분", ""),
                                        placeholder="0.5 또는 1/2",
                                        key=f"excel_child_owner_share_{idx}_{owner_idx}",
                                        label_visibility="collapsed",
                                    )
                                    if owner.get("지분", "") != new_share:
                                        owner_changed = True
                                    owner["지분"] = new_share

                            # 현재 필지의 모든 소유자 지분이 입력된 경우에만
                            # 합계를 계산하여 1 초과/미만 여부를 판단합니다.
                            owners = st.session_state["excel_owner_data"][owner_key]
                            share_complete = (
                                count_value > 0
                                and len(owners) >= count_value
                                and all(is_valid_share_input(o.get("지분", "")) for o in owners[:count_value])
                            )
                            total_share = sum(parse_share(o.get("지분", "")) for o in owners[:count_value])
                            invalid_share = share_complete and not np.isclose(total_share, 1.0, atol=1e-5)

                            if share_complete and invalid_share and owner_changed and hasattr(st, "dialog"):
                                warning_msg = (
                                    "지분의 합이 1이 아닙니다.(1 초과)"
                                    if total_share > 1.0
                                    else "지분의 합이 1이 아닙니다.(1 미만)"
                                )
                                # 같은 필지에서 같은 합계로 반복 팝업되는 것을 방지합니다.
                                # 합계가 정상(1)으로 돌아오면 다음 오류를 다시 알릴 수 있도록
                                # 이전 서명을 초기화합니다.
                                signature = f"{idx}:{selected_plain}:{count_value}:{total_share:.10f}"
                                if signature != st.session_state.get("share_dialog_signature"):
                                    st.session_state["share_dialog_signature"] = signature
                                    st.session_state["pending_share_dialog"] = warning_msg
                            elif share_complete and not invalid_share:
                                st.session_state["share_dialog_signature"] = None

                            # 접힌 상태에서도 합계 셀은 기본행에 표시되므로
                            # 사용자가 현재 합계를 바로 확인할 수 있습니다.

                    # 소유자 입력이 변경되면 다시 렌더링하여 기본행 합계를 갱신
                    if owner_changed:
                        st.rerun()

                # 이전에 예약된 자동 팝업은 위젯 이벤트로 인한 rerun 후 표시
                if st.session_state.get("pending_share_dialog") and hasattr(st, "dialog"):
                    message = st.session_state.pop("pending_share_dialog")
                    share_error_dialog(message)

                st.subheader("📥 결과 엑셀 다운로드")
                download_df = result_df.copy()
                owner_summary = []
                for idx in result_df.index:
                    owners = st.session_state["excel_owner_data"].get(f"excel_owner_{idx}", [])
                    owner_summary.append(
                        "; ".join(
                            f"{o.get('소유구분','')} / {o.get('토지소유자','')} / {o.get('지분','')}"
                            for o in owners
                        )
                    )
                download_df["소유자 입력내역"] = owner_summary
                excel_data = create_excel_download(download_df)
                st.download_button(
                    "📥 결과 엑셀 다운로드",
                    data=excel_data,
                    file_name="result_calc.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel_result",
                )

        except Exception as e:

            st.error(
                f"파일 처리 중 오류 발생: {e}"
            )


# =========================================================
# TAB 2
# 직접 입력 모드
# =========================================================

with tab2:

    st.subheader(
        "📍 토지 기본 정보"
    )

    # -----------------------------------------------------
    # 토지 기본 정보
    # -----------------------------------------------------

    col1, col2, col3 = st.columns(3)

    district = col1.text_input(
        "읍면동리"
    )

    land_type = col2.text_input(
        "지목"
    )

    lot_num = col3.text_input(
        "지번"
    )

    col1, col2 = st.columns(2)

    prev_area = col1.number_input(
        "종전 면적",
        min_value=0.0,
        value=0.0
    )

    conf_area = col2.number_input(
        "확정 면적",
        min_value=0.0,
        value=0.0
    )

    # -----------------------------------------------------
    # 소유 정보
    # -----------------------------------------------------

    st.divider()

    st.subheader(
        "👥 소유 정보"
    )

    ownership_type = st.selectbox(
        "소유 사항",
        [
            "단독",
            "공유",
            "상속"
        ],
        key="ownership_type"
    )

    is_multi = (
        ownership_type in
        ["공유", "상속"]
    )

    num_persons = st.number_input(
        "공유 또는 상속인 수",
        min_value=1,
        max_value=10,
        value=1,
        step=1,
        disabled=not is_multi,
        key="num_persons"
    )

    owner_data = []


    # =====================================================
    # 소유자 입력
    # =====================================================

    for i in range(
        int(num_persons)
    ):

        # -------------------------------------------------
        # 공유일 때 첫 번째 행은 기존 로직대로 비활성화
        # -------------------------------------------------

        if (
            i == 0
            and ownership_type == "공유"
        ):

            c1, c2, c3 = st.columns(
                [3, 3, 2]
            )

            c1.text_input(
                "소유 구분",
                value="",
                disabled=True,
                key=f"dummy_type_{i}"
            )

            c2.text_input(
                "토지소유자",
                value="",
                disabled=True,
                key=f"dummy_name_{i}"
            )

            c3.text_input(
                "지분",
                value="",
                disabled=True,
                key=f"dummy_share_{i}"
            )

            continue


        # -------------------------------------------------
        # 공유 / 상속 접기 기능
        # -------------------------------------------------

        if is_multi:

            if (
                i not in
                st.session_state[
                    "collapsed_rows"
                ]
            ):
                st.session_state[
                    "collapsed_rows"
                ][i] = False


            if st.button(
                "┗",
                key=f"fold_{i}"
            ):

                st.session_state[
                    "collapsed_rows"
                ][i] = not st.session_state[
                    "collapsed_rows"
                ][i]


            if st.session_state[
                "collapsed_rows"
            ][i]:

                st.write(
                    f"소유자 {i + 1} "
                    "입력 행이 접혀 있습니다."
                )

                continue


            st.write(
                f"소유자 {i + 1} 상세 정보"
            )


        # -------------------------------------------------
        # 실제 소유자 입력
        # -------------------------------------------------

        c1, c2, c3 = st.columns(
            [3, 3, 2]
        )

        val1 = c1.selectbox(
            "소유 구분",
            OWNERSHIP_OPTIONS,
            key=f"owner_type_{i}"
        )

        val2 = c2.text_input(
            "토지소유자",
            key=f"owner_name_{i}"
        )

        val3 = c3.text_input(
            "지분",
            key=f"owner_share_{i}",
            placeholder="예: 0.5"
        )


        # -------------------------------------------------
        # 지분 숫자 변환
        # -------------------------------------------------

        try:

            share_text = val3.strip()

            if not share_text:
                share_val = 0.0

            elif "/" in share_text:
                numerator, denominator = share_text.split("/", 1)
                denominator_value = float(denominator.strip())

                if denominator_value == 0:
                    raise ValueError

                share_val = (
                    float(numerator.strip())
                    / denominator_value
                )

            else:
                share_val = float(share_text)


            # 음수 방지
            if share_val < 0:

                st.warning(
                    f"소유자 {i + 1}: "
                    "지분은 음수가 될 수 없습니다."
                )

                share_val = 0.0


            owner_data.append(
                {
                    "소유구분": val1,
                    "토지소유자": val2,
                    "지분": share_val
                }
            )

        except ValueError:

            st.warning(
                f"소유자 {i + 1}: "
                "지분은 숫자로 입력해주세요."
            )


    # =====================================================
    # 지분 검증
    # =====================================================

    warning_msg = (
        OwnershipValidationStrategy
        .validate_share(
            owner_data,
            ownership_type
        )
    )

    if warning_msg:

        st.warning(
            warning_msg
        )


    # =====================================================
    # 직접 입력 데이터 계산
    # =====================================================

    if st.button(
        "입력 내용으로 계산하기",
        key="calculate_manual"
    ):

        new_data = pd.DataFrame(
            [
                {
                    "종전_읍면": district,
                    "종전_동리": district,
                    "종전_지목": land_type,
                    "종전_지번": lot_num,
                    "종전_면적": prev_area,

                    "확정_읍면": district,
                    "확정_동리": district,
                    "확정_지목": land_type,
                    "확정_지번": lot_num,
                    "확정_면적": conf_area,

                    "소유사항": ownership_type,
                    "공유_상속인수": num_persons
                }
            ]
        )

        try:

            result_df = (
                AreaAnalysisStrategy.calculate(
                    new_data
                )
            )

            st.success(
                "✅ 계산이 완료되었습니다."
            )

            st.subheader(
                "📊 계산 결과"
            )

            manual_display_df = result_df.drop(
                columns=[
                    c for c in [
                        "종전_ID",
                        "확정_ID",
                        "종전_총면적"
                    ]
                    if c in result_df.columns
                ],
                errors="ignore"
            )

            st.dataframe(
                manual_display_df,
                use_container_width=True
            )

            # 엑셀 다운로드
            excel_data = (
                create_excel_download(
                    result_df
                )
            )

            st.download_button(
                "📥 결과 엑셀 다운로드",
                data=excel_data,
                file_name="result_calc.xlsx",
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                key="download_manual_result"
            )

        except Exception as e:

            st.error(
                f"계산 중 오류 발생: {e}"
            )