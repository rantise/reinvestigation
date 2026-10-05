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

                def parse_share(value):
                    if value is None or pd.isna(value):
                        return 0.0
                    text = str(value).strip()
                    if not text:
                        return 0.0
                    try:
                        if "/" in text:
                            a, b = text.split("/", 1)
                            b = float(b.strip())
                            if b == 0:
                                return 0.0
                            return float(a.strip()) / b
                        return float(text)
                    except (ValueError, TypeError, ZeroDivisionError):
                        return 0.0

                def fmt_value(value):
                    if value is None or pd.isna(value):
                        return ""
                    if isinstance(value, (float, np.floating)):
                        return f"{float(value):.1f}"
                    return str(value)

                def normalize_ownership(value):
                    text = "" if value is None or pd.isna(value) else str(value).strip()
                    mapping = {"1": "1. 단독", "2": "2. 공유", "3": "3. 상속"}
                    if text in mapping:
                        return mapping[text]
                    if text in ["단독", "공유", "상속"]:
                        return {"단독":"1. 단독", "공유":"2. 공유", "상속":"3. 상속"}[text]
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

                if hasattr(st, "dialog"):
                    @st.dialog("지분 합계 확인")
                    def share_error_dialog(message):
                        st.error(message)
                        st.write("입력한 지분을 확인해 주세요.")

                # 지분 입력 직후 다음 렌더링에서 자동 팝업을 표시합니다.
                if st.session_state.get("pending_share_dialog") and hasattr(st, "dialog"):
                    _pending_message = st.session_state.pop("pending_share_dialog")
                    share_error_dialog(_pending_message)

                # test3.xlsx와 같은 2단 헤더.
                # 하나의 바깥 컨테이너 안에 헤더와 모든 행을 넣어
                # 가로 스크롤이 하나만 생기도록 구성합니다.
                st.markdown(
                    """
                    <style>
                    /* 결과표 전체를 하나의 가로 스크롤 영역으로 사용 */
                    div.st-key-result_table {
                        overflow-x: auto !important;
                        overflow-y: visible !important;
                        width: 100% !important;
                        max-width: 100% !important;
                        padding-bottom: 8px !important;
                    }
                    div.st-key-result_table > div {
                        min-width: 2100px !important;
                    }
                    div.st-key-result_table .stHorizontalBlock {
                        min-width: 2100px !important;
                    }
                    div.st-key-result_table [data-testid="stVerticalBlock"] {
                        min-width: 2100px !important;
                    }
                    .land-result-header {
                        width: 2100px;
                        border-collapse: collapse;
                        table-layout: fixed;
                        margin: 0 0 4px 0;
                        font-size: 13px;
                    }
                    .land-result-header th {
                        border: 1px solid #b7b7b7;
                        background: #eeeeee;
                        text-align: center;
                        vertical-align: middle;
                        height: 32px;
                        padding: 4px 3px;
                        font-weight: 700;
                        white-space: nowrap;
                    }
                    .land-result-header th.group {
                        background: #e2e2e2;
                    }
                    .land-result-cell {
                        box-sizing: border-box;
                        width: 100%;
                        min-height: 38px;
                        border: 1px solid #c5c5c5;
                        background: white;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        padding: 5px 4px;
                        overflow-wrap: anywhere;
                        font-size: 13px;
                    }
                    .land-result-cell.muted {
                        color: #777;
                        background: #f7f7f7;
                    }
                    .land-result-cell.share-error {
                        color: red;
                        background: #ffe5e5;
                        border: 2px solid red;
                        font-weight: 700;
                    }
                    .land-result-special {
                        font-weight: 700;
                        font-size: 18px;
                    }
                    /* widget label 공간 제거 */
                    div.st-key-result_table label {
                        display: none !important;
                    }
                    div.st-key-result_table .stTextInput,
                    div.st-key-result_table .stSelectbox,
                    div.st-key-result_table .stNumberInput {
                        margin: 0 !important;
                        padding: 0 !important;
                    }
                    div.st-key-result_table [data-testid="stWidgetLabel"] {
                        display: none !important;
                    }
                    div.st-key-result_table [data-testid="stTextInputRootElement"],
                    div.st-key-result_table [data-testid="stSelectbox"],
                    div.st-key-result_table [data-testid="stNumberInput"] {
                        min-height: 38px !important;
                    }
                    </style>
                    """,
                    unsafe_allow_html=True
                )

                # 표 열 폭: test3.xlsx의 열 비율을 최대한 유지하면서
                # 계산 결과 열을 추가한 구조입니다.
                col_widths = [1.0, 1.55, 0.90, 1.55, 1.05, 1.55, 1.00,
                              0.90, 1.55, 1.05, 1.55, 1.15,
                              1.15, 1.15, 1.10, 1.10, 1.00, 1.30,
                              1.45, 1.30, 1.25]

                def cell_html(value="", extra_class=""):
                    return (
                        f'<div class="land-result-cell {extra_class}">'
                        f'{value if value is not None else ""}</div>'
                    )

                def read_row_value(row, key):
                    value = row.get(key, "")
                    return "" if pd.isna(value) else value

                # 결과표의 기본 데이터를 소유자 데이터와 함께 준비합니다.
                row_infos = []
                for idx, row in result_df.iterrows():
                    raw_type = read_row_value(row, "소유사항")
                    ownership_type = normalize_ownership(raw_type) if raw_type else "1. 단독"
                    raw_count = row.get("공유_상속인수", np.nan)
                    if pd.isna(raw_count):
                        count = 0
                    else:
                        try:
                            count = max(0, int(float(raw_count)))
                        except (TypeError, ValueError):
                            count = 0
                    if ownership_type == "1. 단독":
                        count = 0

                    owner_key = f"excel_owner_{idx}"
                    needed = count if count > 0 else 1
                    if owner_key not in st.session_state["excel_owner_data"]:
                        st.session_state["excel_owner_data"][owner_key] = [
                            {"소유구분": "1. 개인", "토지소유자": "", "지분": ""}
                            for _ in range(needed)
                        ]
                    else:
                        owners = st.session_state["excel_owner_data"][owner_key]
                        if len(owners) < needed:
                            owners.extend(
                                {"소유구분": "1. 개인", "토지소유자": "", "지분": ""}
                                for _ in range(needed - len(owners))
                            )
                        elif len(owners) > needed:
                            st.session_state["excel_owner_data"][owner_key] = owners[:needed]

                    row_infos.append({
                        "idx": idx,
                        "row": row,
                        "ownership_type": ownership_type,
                        "count": count,
                        "owner_key": owner_key,
                    })

                # 헤더: test3.xlsx의 병합 구조를 그대로 반영
                owner_changed = False
                table = st.container(key="result_table", border=True)
                with table:
                    st.markdown(
                        """
                        <table class="land-result-header">
                          <colgroup>
                            <col style="width:80px"><col style="width:124px">
                            <col style="width:72px"><col style="width:124px"><col style="width:84px"><col style="width:124px"><col style="width:80px">
                            <col style="width:72px"><col style="width:124px"><col style="width:84px"><col style="width:124px"><col style="width:92px">
                            <col style="width:92px"><col style="width:92px"><col style="width:88px"><col style="width:88px">
                            <col style="width:80px"><col style="width:104px"><col style="width:116px"><col style="width:104px"><col style="width:100px">
                          </colgroup>
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

                    # 행별로 하나의 가로 블록을 만들고, 바깥 result_table
                    # 컨테이너가 모든 블록을 함께 가로 스크롤합니다.
                    for info in row_infos:
                        idx = info["idx"]
                        row = info["row"]
                        ownership_type = info["ownership_type"]
                        count = info["count"]
                        owner_key = info["owner_key"]

                        owners = st.session_state["excel_owner_data"][owner_key]
                        total_share = sum(parse_share(o.get("지분", "")) for o in owners) if count > 0 else 1.0
                        invalid_share = count > 0 and not np.isclose(total_share, 1.0, atol=1e-5)

                        # 기본행
                        cells = st.columns(col_widths, gap="small")
                        values = [
                            fmt_value(read_row_value(row, "기준년도")),
                            fmt_value(read_row_value(row, "사업지구명")),
                            fmt_value(read_row_value(row, "종전_읍면")),
                            fmt_value(read_row_value(row, "종전_동리")),
                            fmt_value(read_row_value(row, "종전_지목")),
                            fmt_value(read_row_value(row, "종전_지번")),
                            fmt_value(read_row_value(row, "종전_면적")),
                            fmt_value(read_row_value(row, "확정_읍면")),
                            fmt_value(read_row_value(row, "확정_동리")),
                            fmt_value(read_row_value(row, "확정_지목")),
                            fmt_value(read_row_value(row, "확정_지번")),
                            fmt_value(read_row_value(row, "확정_면적")),
                            fmt_value(read_row_value(row, "증감면적")),
                            fmt_value(read_row_value(row, "감소면적")),
                            fmt_value(read_row_value(row, "증감여부")),
                            fmt_value(read_row_value(row, "처리상태")),
                        ]
                        for c, value in zip(cells[:16], values):
                            c.markdown(cell_html(value), unsafe_allow_html=True)

                        # 소유사항: 업로드 값이 없거나 있어도 수정 가능
                        with cells[16]:
                            type_value = st.selectbox(
                                "소유사항",
                                OWNERSHIP_TYPE_OPTIONS,
                                index=OWNERSHIP_TYPE_OPTIONS.index(ownership_type),
                                key=f"excel_ownership_type_{idx}",
                                label_visibility="collapsed",
                            )
                        selected_type = type_value
                        selected_plain = selected_type.split(". ", 1)[1]

                        # 소유사항 변경을 다음 렌더링부터 적용
                        if selected_type != ownership_type:
                            st.session_state["excel_owner_data"].setdefault(owner_key, [{"소유구분":"1. 개인","토지소유자":"","지분":""}])
                            if selected_plain == "단독":
                                st.session_state["excel_owner_data"][owner_key] = [st.session_state["excel_owner_data"][owner_key][0]]
                            else:
                                target = count if count > 0 else 1
                                current = st.session_state["excel_owner_data"][owner_key]
                                if len(current) < target:
                                    current.extend({"소유구분":"1. 개인","토지소유자":"","지분":""} for _ in range(target-len(current)))
                            st.session_state["excel_result_source_df"].loc[idx, "소유사항"] = selected_plain
                            st.session_state["excel_result_df"].loc[idx, "소유사항"] = selected_plain
                            ownership_type = selected_type
                            count = 0 if selected_plain == "단독" else max(count, 1)

                        # 공유/상속인 수: 단독이면 비활성화, 나머지는 수정 가능
                        with cells[17]:
                            if selected_plain == "단독":
                                st.number_input(
                                    "공유 또는 상속인 수", min_value=0, max_value=100,
                                    value=0, step=1, disabled=True,
                                    key=f"excel_count_{idx}_{selected_plain}", label_visibility="collapsed"
                                )
                                count_value = 0
                            else:
                                count_value = st.number_input(
                                    "공유 또는 상속인 수", min_value=1, max_value=100,
                                    value=max(1, int(count)), step=1,
                                    key=f"excel_count_{idx}_{selected_plain}", label_visibility="collapsed"
                                )
                                if int(count_value) != int(count):
                                    current = st.session_state["excel_owner_data"][owner_key]
                                    if len(current) < int(count_value):
                                        current.extend({"소유구분":"1. 개인","토지소유자":"","지분":""} for _ in range(int(count_value)-len(current)))
                                    elif len(current) > int(count_value):
                                        st.session_state["excel_owner_data"][owner_key] = current[:int(count_value)]
                                    st.session_state["excel_result_df"].loc[idx, "공유_상속인수"] = int(count_value)
                                    count = int(count_value)

                        # 소유구분/토지소유자/지분은 단독이면 기본행에 표시
                        # 공유/상속이면 아래 추가행에서 입력합니다.
                        if selected_plain == "단독":
                            owner = st.session_state["excel_owner_data"][owner_key][0]
                            with cells[18]:
                                owner_type = st.selectbox(
                                    "소유구분", OWNERSHIP_OPTIONS,
                                    index=OWNERSHIP_OPTIONS.index(owner.get("소유구분", "1. 개인")) if owner.get("소유구분", "1. 개인") in OWNERSHIP_OPTIONS else 0,
                                    key=f"excel_owner_type_{idx}_0",
                                    label_visibility="collapsed",
                                )
                                if owner.get("소유구분") != owner_type:
                                    owner_changed = True
                                owner["소유구분"] = owner_type
                            with cells[19]:
                                _new_name = st.text_input(
                                    "토지소유자", value=owner.get("토지소유자", ""),
                                    key=f"excel_owner_name_{idx}_0", label_visibility="collapsed"
                                )
                                if owner.get("토지소유자", "") != _new_name:
                                    owner_changed = True
                                owner["토지소유자"] = _new_name
                            with cells[20]:
                                st.text_input(
                                    "지분", value="1", disabled=True,
                                    key=f"excel_owner_share_{idx}_0", label_visibility="collapsed"
                                )
                                owner["지분"] = "1"
                        else:
                            with cells[18]:
                                st.markdown(cell_html(""), unsafe_allow_html=True)
                            with cells[19]:
                                st.markdown(cell_html(""), unsafe_allow_html=True)
                            with cells[20]:
                                share_cls = "land-result-cell share-error" if invalid_share else "land-result-cell"
                                total_text = f"{total_share:.6g}"
                                st.markdown(cell_html(total_text, share_cls), unsafe_allow_html=True)

                        # 공유/상속 하위행: test3처럼 기준년도 칸에 ┗ 배치
                        if selected_plain in ["공유", "상속"]:
                            if owner_key not in st.session_state["excel_owner_collapsed"]:
                                st.session_state["excel_owner_collapsed"][owner_key] = False

                            # test3의 ┗ 기호는 첫 번째 공유/상속 하위행의
                            # 기준년도 칸에 배치하고, 그 기호 자체를 버튼으로
                            # 사용해 하위 행을 접고 펼칩니다.
                            if not st.session_state["excel_owner_collapsed"][owner_key]:
                                for owner_idx in range(int(count_value)):
                                    owner = st.session_state["excel_owner_data"][owner_key][owner_idx]
                                    child = st.columns(col_widths, gap="small")
                                    # test3의 특수기호는 기준년도 열에 위치합니다.
                                    # 첫 번째 ┗ 자체를 클릭 버튼으로 사용합니다.
                                    with child[0]:
                                        if owner_idx == 0:
                                            if st.button("┗", key=f"fold_excel_{idx}", use_container_width=True):
                                                st.session_state["excel_owner_collapsed"][owner_key] = not st.session_state["excel_owner_collapsed"][owner_key]
                                                st.rerun()
                                        else:
                                            st.markdown(cell_html(""), unsafe_allow_html=True)
                                    for c in child[1:18]:
                                        c.markdown(cell_html(""), unsafe_allow_html=True)
                                    with child[18]:
                                        owner_type = st.selectbox(
                                            "소유구분", OWNERSHIP_OPTIONS,
                                            index=OWNERSHIP_OPTIONS.index(owner.get("소유구분", "1. 개인")) if owner.get("소유구분", "1. 개인") in OWNERSHIP_OPTIONS else 0,
                                            key=f"excel_child_owner_type_{idx}_{owner_idx}",
                                            label_visibility="collapsed",
                                        )
                                        if owner.get("소유구분") != owner_type:
                                            owner_changed = True
                                        owner["소유구분"] = owner_type
                                    with child[19]:
                                        _new_name = st.text_input(
                                            "토지소유자", value=owner.get("토지소유자", ""),
                                            key=f"excel_child_owner_name_{idx}_{owner_idx}", label_visibility="collapsed"
                                        )
                                        if owner.get("토지소유자", "") != _new_name:
                                            owner_changed = True
                                        owner["토지소유자"] = _new_name
                                    with child[20]:
                                        _new_share = st.text_input(
                                            "지분", value=owner.get("지분", ""), placeholder="예: 0.5 또는 1/2",
                                            key=f"excel_child_owner_share_{idx}_{owner_idx}", label_visibility="collapsed"
                                        )
                                        if owner.get("지분", "") != _new_share:
                                            owner_changed = True
                                        owner["지분"] = _new_share

                            # 현재 입력값을 다시 계산하여 자동 팝업
                            total_share = sum(parse_share(o.get("지분", "")) for o in st.session_state["excel_owner_data"][owner_key])
                            invalid_share = not np.isclose(total_share, 1.0, atol=1e-5)
                            if invalid_share:
                                warning_msg = (
                                    "지분의 합이 1이 아닙니다.(1 초과)"
                                    if total_share > 1.0
                                    else "지분의 합이 1이 아닙니다.(1 미만)"
                                )
                                signature = f"{idx}:{selected_plain}:{total_share:.10f}"
                                if signature != st.session_state.get("share_dialog_signature"):
                                    st.session_state["share_dialog_signature"] = signature
                                    if owner_changed:
                                        st.session_state["pending_share_dialog"] = warning_msg

                # 입력이 변경된 경우 한 번 다시 렌더링하여
                # 기본행의 지분 합계 셀도 즉시 새 값으로 갱신합니다.
                if owner_changed:
                    st.rerun()

                    # 결과 행이 너무 많아졌을 때도 동일한 표 안에서
                    # 세로로 자연스럽게 이어지도록 구성합니다.

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