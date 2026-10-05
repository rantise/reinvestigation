import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO


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

                # -------------------------------------------------
                # 엑셀의 2행 헤더와 동일한 그룹 구조
                # -------------------------------------------------
                st.markdown(
                    """
                    <style>
                    .result-scroll {
                        width: 100%;
                        overflow-x: auto;
                        overflow-y: visible;
                        border: 1px solid #d9d9d9;
                        border-radius: 6px;
                        padding: 0;
                        margin-bottom: 10px;
                    }

                    .result-header-table {
                        border-collapse: collapse;
                        min-width: 2140px;
                        width: max-content;
                        table-layout: fixed;
                        font-size: 13px;
                    }

                    .result-header-table th {
                        border: 1px solid #b8b8b8;
                        padding: 7px 8px;
                        text-align: center;
                        background: #f2f2f2;
                        white-space: nowrap;
                        height: 34px;
                    }

                    .result-header-table .group {
                        background: #e5e5e5;
                        font-weight: 700;
                    }

                    .result-header-table .owner-group {
                        background: #dfe8f5;
                    }

                    .result-header-table .owner-sub {
                        background: #edf3fa;
                    }

                    .owner-error input {
                        color: red !important;
                        border: 2px solid red !important;
                        background-color: #fff0f0 !important;
                    }

                    /* 결과 행도 헤더와 동일한 칸 폭/테두리를 사용 */
                    div[data-testid="stHorizontalBlock"]:has(.result-row-marker) {
                        min-width: 2140px !important;
                        width: 2140px !important;
                        max-width: none !important;
                        gap: 0 !important;
                        margin: 0 !important;
                    }
                    div[data-testid="stHorizontalBlock"]:has(.result-row-marker) > div[data-testid="stColumn"] {
                        overflow: visible !important;
                    }
                    div[data-testid="stHorizontalBlock"]:has(.result-row-marker) > div[data-testid="stColumn"] {
                        border-right: 1px solid #b8b8b8;
                        border-bottom: 1px solid #b8b8b8;
                        box-sizing: border-box;
                        min-height: 56px;
                        padding: 0 !important;
                    }
                    .result-cell {
                        min-height: 56px;
                        height: 100%;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        padding: 6px 7px;
                        box-sizing: border-box;
                        font-size: 13px;
                        line-height: 1.25;
                        word-break: break-word;
                        overflow-wrap: anywhere;
                    }
                    .result-cell-left {
                        justify-content: flex-start;
                    }
                    .result-cell-input {
                        min-height: 56px;
                        display: flex;
                        align-items: center;
                        padding: 4px 6px;
                        box-sizing: border-box;
                    }

                    .owner-row-label {
                        font-weight: 600;
                        margin-top: 3px;
                        margin-bottom: 3px;
                    }
                    </style>

                    <div class="result-scroll">
                    <table class="result-header-table">
                      <colgroup>
                        <col style="width:90px">
                        <col style="width:150px">
                        <col style="width:75px">
                        <col style="width:95px">
                        <col style="width:75px">
                        <col style="width:105px">
                        <col style="width:100px">
                        <col style="width:75px">
                        <col style="width:95px">
                        <col style="width:75px">
                        <col style="width:105px">
                        <col style="width:100px">
                        <col style="width:105px">
                        <col style="width:125px">
                        <col style="width:105px">
                        <col style="width:85px">
                        <col style="width:105px">
                        <col style="width:125px">
                        <col style="width:100px">
                        <col style="width:150px">
                        <col style="width:100px">
                      </colgroup>
                      <tr>
                        <th rowspan="2">기준년도</th>
                        <th rowspan="2">사업지구명</th>
                        <th class="group" colspan="5">종전</th>
                        <th class="group" colspan="5">확정</th>
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
                        <th>읍면</th>
                        <th>동리</th>
                        <th>지목</th>
                        <th>지번</th>
                        <th>면적</th>
                        <th>읍면</th>
                        <th>동리</th>
                        <th>지목</th>
                        <th>지번</th>
                        <th>면적</th>
                      </tr>
                    </table>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                def parse_share(value):
                    if value is None:
                        return 0.0

                    value = str(value).strip()

                    if not value:
                        return 0.0

                    try:
                        if "/" in value:
                            numerator, denominator = value.split("/", 1)
                            denominator = float(denominator.strip())

                            if denominator == 0:
                                return 0.0

                            return (
                                float(numerator.strip())
                                / denominator
                            )

                        return float(value)

                    except (ValueError, TypeError, ZeroDivisionError):
                        return 0.0

                if "share_error_dialog" not in st.session_state:
                    st.session_state["share_error_dialog"] = None

                # Streamlit 팝업
                if hasattr(st, "dialog"):

                    @st.dialog("지분 합계 확인")
                    def share_error_dialog(message):
                        st.error(message)
                        st.write("입력한 지분을 확인해 주세요.")

                # -------------------------------------------------
                # 결과 행 표시용 셀 함수
                # -------------------------------------------------
                def result_cell(column, value, left=False):
                    cls = "result-cell result-cell-left" if left else "result-cell"
                    if pd.isna(value):
                        value = ""
                    column.markdown(
                        f'<div class="{cls}">{str(value)}</div>',
                        unsafe_allow_html=True
                    )

                # -------------------------------------------------
                # 결과 행
                # -------------------------------------------------
                for idx, row in result_df.iterrows():

                    ownership_type = str(
                        row.get("소유사항", "")
                    ).strip()

                    raw_count = row.get(
                        "공유_상속인수",
                        np.nan
                    )

                    # 빈 셀 = 0
                    if pd.isna(raw_count):
                        owner_count_value = 0
                    else:
                        try:
                            owner_count_value = max(
                                0,
                                int(float(raw_count))
                            )
                        except (TypeError, ValueError):
                            owner_count_value = 0

                    # '단독'이면 공유/상속인 수를 비활성 상태로 취급
                    if ownership_type == "단독":
                        owner_count_value = 0

                    # -------------------------------------------------
                    # 0명/빈값/단독:
                    # 기본 결과 행에 소유구분/소유자/지분을 바로 배치
                    # 추가 행은 만들지 않음
                    # -------------------------------------------------
                    if owner_count_value == 0:

                        owner_key = f"excel_owner_{idx}"

                        if owner_key not in st.session_state["excel_owner_data"]:
                            st.session_state["excel_owner_data"][owner_key] = [
                                {
                                    "소유구분": "1. 개인",
                                    "토지소유자": "",
                                    "지분": ""
                                }
                            ]

                        owner = st.session_state["excel_owner_data"][owner_key][0]

                        # 가로 스크롤 가능한 넓은 결과 행
                        row_box = st.container()

                        with row_box:

                            cols = st.columns(
                                [90, 150, 75, 95, 75, 105, 100,
                                 75, 95, 75, 105, 100, 105, 125, 105, 85,
                                 105, 125, 100, 150, 100],
                                gap=None
                            )

                            # 기본 정보
                            cols[0].markdown('<span class="result-row-marker"></span>', unsafe_allow_html=True)
                            result_cell(cols[0], row.get("기준년도", ""))
                            result_cell(cols[1], row.get("사업지구명", ""), left=True)

                            result_cell(cols[2], row.get("종전_읍면", ""))
                            result_cell(cols[3], row.get("종전_동리", ""))
                            result_cell(cols[4], row.get("종전_지목", ""))
                            result_cell(cols[5], row.get("종전_지번", ""))
                            result_cell(cols[6], row.get("종전_면적", ""))

                            result_cell(cols[7], row.get("확정_읍면", ""))
                            result_cell(cols[8], row.get("확정_동리", ""))
                            result_cell(cols[9], row.get("확정_지목", ""))
                            result_cell(cols[10], row.get("확정_지번", ""))
                            result_cell(cols[11], row.get("확정_면적", ""))

                            result_cell(cols[12], row.get("증감면적", ""))
                            result_cell(cols[13], row.get("감소면적", ""))
                            result_cell(cols[14], row.get("증감여부", ""))
                            result_cell(cols[15], row.get("처리상태", ""))
                            result_cell(cols[16], row.get("소유사항", ""))

                            # 공유_상속인수:
                            # 단독이면 비활성화된 0,
                            # 빈 값도 0으로 표시
                            cols[17].number_input(
                                "공유 또는 상속인 수",
                                min_value=0,
                                value=0,
                                step=1,
                                disabled=True,
                                key=f"owner_count_disabled_{idx}",
                                label_visibility="collapsed"
                            )

                            selected = cols[18].selectbox(
                                "소유구분",
                                OWNERSHIP_OPTIONS,
                                index=(
                                    OWNERSHIP_OPTIONS.index(
                                        owner["소유구분"]
                                    )
                                    if owner["소유구분"] in OWNERSHIP_OPTIONS
                                    else 0
                                ),
                                key=f"same_row_owner_type_{idx}",
                                label_visibility="collapsed"
                            )
                            st.session_state["excel_owner_data"][owner_key][0]["소유구분"] = selected

                            owner_name = cols[19].text_input(
                                "토지소유자",
                                value=str(owner.get("토지소유자", "")),
                                key=f"same_row_owner_name_{idx}",
                                label_visibility="collapsed"
                            )

                            st.session_state["excel_owner_data"][owner_key][0]["토지소유자"] = owner_name

                            share_text = cols[20].text_input(
                                "지분",
                                value=str(owner.get("지분", "")),
                                placeholder="예: 0.5 또는 1/2",
                                key=f"same_row_owner_share_{idx}",
                                label_visibility="collapsed"
                            )

                            st.session_state["excel_owner_data"][owner_key][0]["지분"] = share_text

                        st.divider()

                        # 단독/0명 행은 여기서 끝
                        continue

                    # -------------------------------------------------
                    # 공유/상속이고 숫자가 있는 경우:
                    # 기존 방식대로 숫자만큼 소유자 추가 행 생성
                    # -------------------------------------------------
                    owner_key = f"excel_owner_{idx}"

                    if owner_key not in st.session_state["excel_owner_data"]:
                        st.session_state["excel_owner_data"][owner_key] = [
                            {
                                "소유구분": "1. 개인",
                                "토지소유자": "",
                                "지분": ""
                            }
                            for _ in range(owner_count_value)
                        ]
                    else:
                        current = st.session_state["excel_owner_data"][owner_key]

                        if len(current) < owner_count_value:
                            current.extend(
                                [
                                    {
                                        "소유구분": "1. 개인",
                                        "토지소유자": "",
                                        "지분": ""
                                    }
                                    for _ in range(
                                        owner_count_value - len(current)
                                    )
                                ]
                            )
                        elif len(current) > owner_count_value:
                            st.session_state["excel_owner_data"][owner_key] = (
                                current[:owner_count_value]
                            )

                    collapse_key = f"excel_collapse_{idx}"

                    if collapse_key not in st.session_state["excel_owner_collapsed"]:
                        st.session_state["excel_owner_collapsed"][collapse_key] = False

                    # 기본 결과 행은 읽기 전용 정보로 표시
                    base_cols = st.columns(
                        [90, 150, 75, 95, 75, 105, 100,
                         75, 95, 75, 105, 100, 105, 125, 105, 85,
                         105, 125, 100],
                        gap=None
                    )

                    base_values = [
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
                        row.get("소유사항", ""),
                    ]

                    base_cols[0].markdown('<span class="result-row-marker"></span>', unsafe_allow_html=True)
                    for col, value in zip(base_cols[:17], base_values):
                        result_cell(col, value)

                    base_cols[17].number_input(
                        "공유 또는 상속인 수",
                        min_value=0,
                        value=owner_count_value,
                        step=1,
                        disabled=True,
                        key=f"owner_count_{idx}",
                        label_visibility="collapsed"
                    )

                    if base_cols[18].button(
                        "┗",
                        key=f"excel_fold_{idx}"
                    ):
                        st.session_state["excel_owner_collapsed"][collapse_key] = (
                            not st.session_state["excel_owner_collapsed"][collapse_key]
                        )
                        st.rerun()

                    if st.session_state["excel_owner_collapsed"][collapse_key]:
                        st.info(
                            f"소유자 입력 {owner_count_value}개 행이 접혀 있습니다."
                        )
                        st.divider()
                        continue

                    owner_rows_for_validation = []

                    # 추가 소유자 행
                    for owner_idx in range(owner_count_value):

                        owner = st.session_state["excel_owner_data"][owner_key][owner_idx]

                        c1, c2, c3 = st.columns([3, 4, 2], gap="small")

                        selected = c1.selectbox(
                            "소유구분",
                            OWNERSHIP_OPTIONS,
                            index=(
                                OWNERSHIP_OPTIONS.index(
                                    owner["소유구분"]
                                )
                                if owner["소유구분"] in OWNERSHIP_OPTIONS
                                else 0
                            ),
                            key=f"extra_owner_type_{idx}_{owner_idx}",
                            label_visibility="collapsed"
                        )
                        st.session_state["excel_owner_data"][owner_key][owner_idx]["소유구분"] = selected

                        owner_name = c2.text_input(
                            "토지소유자",
                            value=str(owner.get("토지소유자", "")),
                            key=f"extra_owner_name_{idx}_{owner_idx}",
                            label_visibility="collapsed"
                        )

                        st.session_state["excel_owner_data"][owner_key][owner_idx]["토지소유자"] = owner_name

                        share_text = c3.text_input(
                            "지분",
                            value=str(owner.get("지분", "")),
                            placeholder="예: 0.5 또는 1/2",
                            key=f"extra_owner_share_{idx}_{owner_idx}",
                            label_visibility="collapsed"
                        )

                        st.session_state["excel_owner_data"][owner_key][owner_idx]["지분"] = share_text

                        owner_rows_for_validation.append(
                            {
                                "소유구분": selected,
                                "토지소유자": owner_name,
                                "지분": parse_share(share_text)
                            }
                        )

                    warning_msg = OwnershipValidationStrategy.validate_share(
                        owner_rows_for_validation,
                        ownership_type
                    )

                    if warning_msg:

                        st.markdown(
                            f"""
                            <div style="
                                color:red;
                                border:2px solid red;
                                background:#fff0f0;
                                padding:10px;
                                border-radius:6px;
                                font-weight:bold;
                                margin:8px 0;
                            ">
                            {warning_msg}
                            </div>
                            """,
                            unsafe_allow_html=True
                        )

                        if hasattr(st, "dialog"):
                            if st.button(
                                "⚠️ 지분 오류 팝업 열기",
                                key=f"share_error_{idx}"
                            ):
                                share_error_dialog(warning_msg)

                    st.divider()

                # -------------------------------------------------
                # 결과 엑셀 다운로드
                # -------------------------------------------------
                st.subheader("📥 결과 엑셀 다운로드")

                download_df = result_df.copy()

                owner_summary = []

                for idx in result_df.index:

                    owner_key = f"excel_owner_{idx}"

                    owners = st.session_state["excel_owner_data"].get(
                        owner_key,
                        []
                    )

                    owner_summary.append(
                        "; ".join(
                            [
                                (
                                    f"{o['소유구분']} / "
                                    f"{o['토지소유자']} / "
                                    f"{o['지분']}"
                                )
                                for o in owners
                            ]
                        )
                    )

                download_df["소유자 입력내역"] = owner_summary

                excel_data = create_excel_download(
                    download_df
                )

                st.download_button(
                    "📥 결과 엑셀 다운로드",
                    data=excel_data,
                    file_name="result_calc.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    key="download_excel_result"
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