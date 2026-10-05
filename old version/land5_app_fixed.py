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
        df["증가면적"] = np.where(
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

        # 증감 여부
        df["증감여부"] = np.where(
            df["증가면적"] > 0,
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
            # 엑셀 읽기
            df_input = pd.read_excel(
                uploaded_file,
                header=[0, 1]
            )

            # 컬럼명 변환
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
                    "종전_읍면",
                    "종전_동리",
                    "종전_지목",
                    "종전_지번",
                    "종전_면적",
                    "확정_읍면",
                    "확정_동리",
                    "확정_지목",
                    "확정_지번",
                    "확정_면적"
                ]

                missing_columns = [
                    col
                    for col in required_columns
                    if col not in df_input.columns
                ]

                if missing_columns:
                    st.error(
                        "필수 컬럼이 누락되었습니다.\n\n"
                        + ", ".join(missing_columns)
                    )
                else:
                    result_df = AreaAnalysisStrategy.calculate(df_input)

                    # 계산용 내부 컬럼은 결과 화면/다운로드에서 제외
                    hidden_columns = [
                        "종전_ID",
                        "확정_ID",
                        "종전_총면적"
                    ]

                    display_df = result_df.drop(
                        columns=[
                            c for c in hidden_columns
                            if c in result_df.columns
                        ]
                    ).copy()

                    # 업로드 결과를 세션에 저장하여 위젯 상호작용 시에도 유지
                    st.session_state["excel_result_df"] = display_df
                    st.session_state["excel_result_source_df"] = result_df
                    st.session_state["excel_owner_data"] = {}
                    st.session_state["excel_owner_collapsed"] = {}

                    st.success("✅ 계산이 완료되었습니다.")

            # 계산 결과가 세션에 있으면 계속 표시
            if "excel_result_df" in st.session_state:
                result_df = st.session_state["excel_result_df"].copy()

                st.subheader("📊 계산 결과 및 소유자 입력")

                # 화면에 보여줄 기본 결과 컬럼 순서
                preferred_columns = [
                    "기준년도",
                    "사업지구명",
                    "종전_읍면",
                    "종전_동리",
                    "종전_지목",
                    "종전_지번",
                    "종전_면적",
                    "확정_읍면",
                    "확정_동리",
                    "확정_지목",
                    "확정_지번",
                    "확정_면적",
                    "증가면적",
                    "감소면적",
                    "증감여부",
                    "처리상태",
                    "소유사항",
                    "공유_상속인수"
                ]

                visible_columns = [
                    c for c in preferred_columns
                    if c in result_df.columns
                ]

                # preferred_columns에 없는 사용자 컬럼도 표시
                visible_columns += [
                    c for c in result_df.columns
                    if c not in visible_columns
                    and c not in ["종전_ID", "확정_ID", "종전_총면적"]
                ]

                # 결과 표에는 요청하신 순서대로 처리상태 뒤에 소유사항/공유_상속인수를 배치
                if "처리상태" in visible_columns:
                    visible_columns = [
                        c for c in visible_columns
                        if c not in ["소유사항", "공유_상속인수"]
                    ]
                    pos = visible_columns.index("처리상태") + 1
                    visible_columns[pos:pos] = [
                        c for c in ["소유사항", "공유_상속인수"]
                        if c in result_df.columns
                    ]

                # 계산 결과 기본 표
                st.dataframe(
                    result_df[visible_columns],
                    use_container_width=True,
                    hide_index=True
                )

                # -------------------------------------------------
                # 소유자 입력용 CSS
                # -------------------------------------------------
                st.markdown(
                    """
                    <style>
                    .owner-title {
                        font-weight: 700;
                        font-size: 16px;
                        margin-top: 8px;
                        margin-bottom: 4px;
                    }
                    .share-error input {
                        color: red !important;
                        border: 2px solid red !important;
                        background-color: #fff0f0 !important;
                    }
                    .owner-row {
                        border: 1px solid #dddddd;
                        border-radius: 6px;
                        padding: 6px 8px 2px 8px;
                        margin-bottom: 4px;
                    }
                    </style>
                    """,
                    unsafe_allow_html=True
                )

                def parse_share(value):
                    """지분: 소수(0.5)와 분수(1/2)를 모두 지원."""
                    if value is None:
                        return 0.0

                    value = str(value).strip()

                    if not value:
                        return 0.0

                    try:
                        if "/" in value:
                            parts = value.split("/")
                            if len(parts) != 2:
                                return 0.0

                            numerator = float(parts[0].strip())
                            denominator = float(parts[1].strip())

                            if denominator == 0:
                                return 0.0

                            return numerator / denominator

                        return float(value)

                    except (ValueError, TypeError, ZeroDivisionError):
                        return 0.0

                def ownership_from_code(value):
                    """1~5 숫자 입력을 드롭다운의 동일 항목으로 변환."""
                    value = str(value).strip()

                    if value in OWNERSHIP_CODE_MAP:
                        code = value
                        return f"{code}. {OWNERSHIP_CODE_MAP[code]}"

                    # '1. 개인'처럼 전체 항목을 입력해도 허용
                    for option in OWNERSHIP_OPTIONS:
                        if value == option:
                            return option

                    # '개인'처럼 명칭만 입력해도 허용
                    for option in OWNERSHIP_OPTIONS:
                        if value == option.split(". ", 1)[1]:
                            return option

                    return None

                def show_share_dialog(message):
                    """지분 오류를 팝업 대화상자로 표시."""
                    try:
                        st.session_state["share_error_dialog_message"] = message
                    except Exception:
                        pass

                # Streamlit의 dialog 기능을 사용합니다.
                if hasattr(st, "dialog"):

                    @st.dialog("지분 합계 확인")
                    def share_error_dialog(message):
                        st.error(message)
                        st.write("입력한 지분을 확인해 주세요.")

                # -------------------------------------------------
                # 각 토지 결과 행 + 그 아래 소유자 행
                # -------------------------------------------------
                for idx, row in result_df.iterrows():

                    # 공유/상속인 수 안전 처리
                    raw_owner_count = row.get(
                        "공유_상속인수",
                        1
                    )

                    try:
                        if pd.isna(raw_owner_count):
                            owner_count = 1
                        else:
                            owner_count = int(float(raw_owner_count))
                            owner_count = max(1, min(100, owner_count))
                    except (TypeError, ValueError):
                        owner_count = 1

                    ownership_type = str(
                        row.get("소유사항", "")
                    ).strip()

                    # 단독은 1개 입력 행, 공유/상속은 엑셀 값만큼 입력 행 생성
                    if ownership_type not in ["공유", "상속"]:
                        owner_count = 1

                    st.markdown(
                        f'<div class="owner-title">📍 {idx + 1}번 토지 소유자 입력</div>',
                        unsafe_allow_html=True
                    )

                    # 토지의 핵심 기본 정보는 읽기 전용으로 표시
                    basic_cols = st.columns(6)

                    basic_cols[0].write(
                        f"**확정 지번**  \n{row.get('확정_지번', '')}"
                    )
                    basic_cols[1].write(
                        f"**처리 상태**  \n{row.get('처리상태', '')}"
                    )
                    basic_cols[2].write(
                        f"**소유사항**  \n{row.get('소유사항', '')}"
                    )
                    basic_cols[3].write(
                        f"**공유·상속인 수**  \n{owner_count}"
                    )
                    basic_cols[4].write(
                        f"**종전 면적**  \n{row.get('종전_면적', 0)}"
                    )
                    basic_cols[5].write(
                        f"**확정 면적**  \n{row.get('확정_면적', 0)}"
                    )

                    # 이 토지의 소유자 데이터 초기화
                    owner_key = f"excel_owner_{idx}"

                    if owner_key not in st.session_state["excel_owner_data"]:
                        st.session_state["excel_owner_data"][owner_key] = [
                            {
                                "소유구분": "1. 개인",
                                "토지소유자": "",
                                "지분": ""
                            }
                            for _ in range(owner_count)
                        ]
                    else:
                        current = st.session_state["excel_owner_data"][owner_key]

                        # 인원수가 바뀐 경우 행 수를 엑셀 값에 맞춤
                        if len(current) < owner_count:
                            current.extend(
                                [
                                    {
                                        "소유구분": "1. 개인",
                                        "토지소유자": "",
                                        "지분": ""
                                    }
                                    for _ in range(
                                        owner_count - len(current)
                                    )
                                ]
                            )
                        elif len(current) > owner_count:
                            st.session_state["excel_owner_data"][owner_key] = current[:owner_count]

                    # 접기 상태
                    collapse_key = f"excel_collapse_{idx}"

                    if collapse_key not in st.session_state["excel_owner_collapsed"]:
                        st.session_state["excel_owner_collapsed"][collapse_key] = False

                    # 첫 번째 소유자 행 앞의 ┗ 버튼
                    if owner_count > 1:
                        button_col, title_col = st.columns([0.6, 9.4])

                        with button_col:
                            if st.button(
                                "┗",
                                key=f"excel_fold_{idx}"
                            ):
                                st.session_state["excel_owner_collapsed"][collapse_key] = (
                                    not st.session_state["excel_owner_collapsed"][collapse_key]
                                )
                                st.rerun()

                        with title_col:
                            st.caption(
                                f"소유자 입력 ({owner_count}명)"
                                + (
                                    " — 접혀 있음"
                                    if st.session_state["excel_owner_collapsed"][collapse_key]
                                    else ""
                                )
                            )

                    if (
                        owner_count > 1
                        and st.session_state["excel_owner_collapsed"][collapse_key]
                    ):
                        st.info("┗ 소유자 입력 행이 접혀 있습니다. 다시 누르면 펼쳐집니다.")
                        continue

                    # 지분 입력값
                    owner_rows_for_validation = []

                    for owner_idx in range(owner_count):

                        owner = st.session_state["excel_owner_data"][owner_key][owner_idx]

                        st.markdown(
                            '<div class="owner-row">',
                            unsafe_allow_html=True
                        )

                        # 한 행에 소유구분 / 토지소유자 / 지분만 배치
                        c1, c2, c3 = st.columns([3, 4, 2])

                        # -----------------------------------------
                        # 소유구분: 드롭다운
                        # -----------------------------------------
                        selected_option = c1.selectbox(
                            "소유구분",
                            OWNERSHIP_OPTIONS,
                            index=(
                                OWNERSHIP_OPTIONS.index(owner["소유구분"])
                                if owner["소유구분"] in OWNERSHIP_OPTIONS
                                else 0
                            ),
                            key=f"excel_owner_select_{idx}_{owner_idx}"
                        )

                        # 숫자 입력으로 동일한 소유구분 선택 가능
                        code_input = c1.text_input(
                            "번호 입력(선택)",
                            value="",
                            placeholder="예: 1",
                            key=f"excel_owner_code_{idx}_{owner_idx}"
                        )

                        if code_input.strip():
                            converted = ownership_from_code(code_input)

                            if converted:
                                selected_option = converted
                                st.session_state["excel_owner_data"][owner_key][owner_idx]["소유구분"] = converted
                            else:
                                c1.warning("1~5 중 하나를 입력하세요.")

                        else:
                            st.session_state["excel_owner_data"][owner_key][owner_idx]["소유구분"] = selected_option

                        # -----------------------------------------
                        # 토지소유자
                        # -----------------------------------------
                        owner_name = c2.text_input(
                            "토지소유자",
                            value=str(owner.get("토지소유자", "")),
                            key=f"excel_owner_name_{idx}_{owner_idx}"
                        )

                        st.session_state["excel_owner_data"][owner_key][owner_idx]["토지소유자"] = owner_name

                        # -----------------------------------------
                        # 지분
                        # -----------------------------------------
                        share_text = c3.text_input(
                            "지분",
                            value=str(owner.get("지분", "")),
                            placeholder="예: 0.5 또는 1/2",
                            key=f"excel_owner_share_{idx}_{owner_idx}"
                        )

                        st.session_state["excel_owner_data"][owner_key][owner_idx]["지분"] = share_text

                        share_value = parse_share(share_text)

                        owner_rows_for_validation.append(
                            {
                                "소유구분": selected_option,
                                "토지소유자": owner_name,
                                "지분": share_value
                            }
                        )

                        st.markdown(
                            "</div>",
                            unsafe_allow_html=True
                        )

                    # ---------------------------------------------
                    # 공유/상속 지분 합계 검사
                    # ---------------------------------------------
                    if ownership_type in ["공유", "상속"]:

                        warning_msg = (
                            OwnershipValidationStrategy.validate_share(
                                owner_rows_for_validation,
                                ownership_type
                            )
                        )

                        if warning_msg:
                            # 빨간색 안내
                            st.markdown(
                                f"""
                                <div style="
                                    color:red;
                                    border:2px solid red;
                                    background:#fff0f0;
                                    padding:10px;
                                    border-radius:6px;
                                    font-weight:bold;
                                ">
                                {warning_msg}
                                </div>
                                """,
                                unsafe_allow_html=True
                            )

                            # 실제 팝업 대화상자
                            if hasattr(st, "dialog"):
                                if st.button(
                                    "⚠️ 지분 오류 팝업 열기",
                                    key=f"share_error_{idx}"
                                ):
                                    share_error_dialog(warning_msg)

                    st.divider()

                # -------------------------------------------------
                # 소유자 입력 결과 저장 / 다운로드
                # -------------------------------------------------
                st.subheader("📥 결과 엑셀 다운로드")

                # 원본 계산 결과에서 내부 ID 3개 제거
                download_df = result_df.drop(
                    columns=[
                        c for c in [
                            "종전_ID",
                            "확정_ID",
                            "종전_총면적"
                        ]
                        if c in result_df.columns
                    ],
                    errors="ignore"
                ).copy()

                # 소유자 입력 내용을 별도 컬럼으로 저장
                owner_summary = []

                for idx in result_df.index:
                    key = f"excel_owner_{idx}"
                    owners = st.session_state["excel_owner_data"].get(
                        key,
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

                excel_data = create_excel_download(download_df)

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