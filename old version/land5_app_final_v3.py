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
                # 결과표는 헤더와 데이터가 분리되지 않는 하나의
                # st.data_editor 표로 표시합니다.
                # - 하나의 표 안에서 가로 스크롤
                # - 헤더와 결과값의 열 위치 동일
                # - 공유/상속이면 실제 엑셀 행처럼 하위 행 추가
                # - 하위 행의 첫 번째 소유사항 칸에 ┗ 표시
                # -------------------------------------------------
                st.markdown(
                    """
                    <style>
                    /* 결과표 자체가 하나의 스크롤 영역이 되도록 합니다. */
                    div[data-testid="stDataEditor"] {
                        width: 100% !important;
                    }

                    /* 편집 셀의 줄바꿈을 허용합니다. */
                    div[data-testid="stDataEditor"] [role="gridcell"] {
                        white-space: normal !important;
                    }
                    </style>
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

                if hasattr(st, "dialog"):

                    @st.dialog("지분 합계 확인")
                    def share_error_dialog(message):
                        st.error(message)
                        st.write("입력한 지분을 확인해 주세요.")

                # -------------------------------------------------
                # 소유자 데이터 초기화 및 결과표용 행 생성
                # -------------------------------------------------
                display_rows = []
                row_meta = []

                for idx, row in result_df.iterrows():

                    ownership_type = str(
                        row.get("소유사항", "")
                    ).strip()

                    raw_count = row.get(
                        "공유_상속인수",
                        np.nan
                    )

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

                    # 단독은 공유/상속인 수를 0으로 고정
                    if ownership_type == "단독":
                        owner_count_value = 0

                    owner_key = f"excel_owner_{idx}"

                    # 단독/0명은 기본행에 소유자 정보를 입력합니다.
                    required_owner_rows = (
                        1 if owner_count_value == 0
                        else owner_count_value
                    )

                    if owner_key not in st.session_state["excel_owner_data"]:
                        st.session_state["excel_owner_data"][owner_key] = [
                            {
                                "소유구분": "1. 개인",
                                "토지소유자": "",
                                "지분": ""
                            }
                            for _ in range(required_owner_rows)
                        ]
                    else:
                        current = st.session_state["excel_owner_data"][owner_key]

                        if len(current) < required_owner_rows:
                            current.extend(
                                [
                                    {
                                        "소유구분": "1. 개인",
                                        "토지소유자": "",
                                        "지분": ""
                                    }
                                    for _ in range(
                                        required_owner_rows - len(current)
                                    )
                                ]
                            )
                        elif len(current) > required_owner_rows:
                            st.session_state["excel_owner_data"][owner_key] = (
                                current[:required_owner_rows]
                            )

                    owners = st.session_state["excel_owner_data"][owner_key]

                    base = {
                        "기준년도": "" if pd.isna(row.get("기준년도", "")) else row.get("기준년도", ""),
                        "사업지구명": "" if pd.isna(row.get("사업지구명", "")) else row.get("사업지구명", ""),
                        "종전\n읍면": "" if pd.isna(row.get("종전_읍면", "")) else row.get("종전_읍면", ""),
                        "종전\n동리": "" if pd.isna(row.get("종전_동리", "")) else row.get("종전_동리", ""),
                        "종전\n지목": "" if pd.isna(row.get("종전_지목", "")) else row.get("종전_지목", ""),
                        "종전\n지번": "" if pd.isna(row.get("종전_지번", "")) else row.get("종전_지번", ""),
                        "종전\n면적": "" if pd.isna(row.get("종전_면적", "")) else row.get("종전_면적", ""),
                        "확정\n읍면": "" if pd.isna(row.get("확정_읍면", "")) else row.get("확정_읍면", ""),
                        "확정\n동리": "" if pd.isna(row.get("확정_동리", "")) else row.get("확정_동리", ""),
                        "확정\n지목": "" if pd.isna(row.get("확정_지목", "")) else row.get("확정_지목", ""),
                        "확정\n지번": "" if pd.isna(row.get("확정_지번", "")) else row.get("확정_지번", ""),
                        "확정\n면적": "" if pd.isna(row.get("확정_면적", "")) else row.get("확정_면적", ""),
                        "증감면적": "" if pd.isna(row.get("증감면적", "")) else row.get("증감면적", ""),
                        "감소면적": "" if pd.isna(row.get("감소면적", "")) else row.get("감소면적", ""),
                        "증감여부": "" if pd.isna(row.get("증감여부", "")) else row.get("증감여부", ""),
                        "처리상태": "" if pd.isna(row.get("처리상태", "")) else row.get("처리상태", ""),
                        "소유사항": ownership_type,
                        "공유 또는 상속인 수": owner_count_value,
                        "소유구분": "",
                        "토지소유자": "",
                        "지분": "",
                    }

                    # -------------------------------------------------
                    # 단독/0명:
                    # 기존 결과행 바로 옆에 소유구분/소유자/지분
                    # -------------------------------------------------
                    if owner_count_value == 0:

                        owner = owners[0]

                        base["소유구분"] = owner.get(
                            "소유구분",
                            "1. 개인"
                        )
                        base["토지소유자"] = owner.get(
                            "토지소유자",
                            ""
                        )
                        base["지분"] = owner.get(
                            "지분",
                            ""
                        )

                        display_rows.append(base)
                        row_meta.append(
                            {
                                "source_idx": idx,
                                "owner_idx": 0,
                                "is_base": True,
                                "ownership_type": ownership_type,
                                "owner_count": 0,
                            }
                        )

                    # -------------------------------------------------
                    # 공유/상속:
                    # 기본행 1개 + 소유자 수만큼 하위행을 같은 표에
                    # 바로 붙입니다.
                    #
                    # test2.xlsx 형태:
                    # 공유 | 3
                    # ┗  |   | 소유구분 | 토지소유자 | 지분
                    # -------------------------------------------------
                    else:

                        display_rows.append(base)
                        row_meta.append(
                            {
                                "source_idx": idx,
                                "owner_idx": None,
                                "is_base": True,
                                "ownership_type": ownership_type,
                                "owner_count": owner_count_value,
                            }
                        )

                        for owner_idx in range(owner_count_value):

                            owner = owners[owner_idx]

                            child = {
                                key: ""
                                for key in base.keys()
                            }

                            # 특수기호는 소유구분 바로 왼쪽의
                            # '소유사항' 칸에 배치합니다.
                            child["소유사항"] = (
                                "┗" if owner_idx == 0 else ""
                            )

                            child["소유구분"] = owner.get(
                                "소유구분",
                                "1. 개인"
                            )
                            child["토지소유자"] = owner.get(
                                "토지소유자",
                                ""
                            )
                            child["지분"] = owner.get(
                                "지분",
                                ""
                            )

                            display_rows.append(child)
                            row_meta.append(
                                {
                                    "source_idx": idx,
                                    "owner_idx": owner_idx,
                                    "is_base": False,
                                    "ownership_type": ownership_type,
                                    "owner_count": owner_count_value,
                                }
                            )

                # -------------------------------------------------
                # 실제 표시 컬럼 순서
                # -------------------------------------------------
                display_columns = [
                    "기준년도",
                    "사업지구명",
                    "종전\n읍면",
                    "종전\n동리",
                    "종전\n지목",
                    "종전\n지번",
                    "종전\n면적",
                    "확정\n읍면",
                    "확정\n동리",
                    "확정\n지목",
                    "확정\n지번",
                    "확정\n면적",
                    "증감면적",
                    "감소면적",
                    "증감여부",
                    "처리상태",
                    "소유사항",
                    "공유 또는 상속인 수",
                    "소유구분",
                    "토지소유자",
                    "지분",
                ]

                display_df = pd.DataFrame(
                    display_rows,
                    columns=display_columns
                )

                # -------------------------------------------------
                # 하나의 data_editor에서 헤더 + 값 + 입력을 모두
                # 처리합니다. 따라서 별도의 헤더 스크롤이 없습니다.
                # -------------------------------------------------
                disabled_columns = [
                    col for col in display_columns
                    if col not in [
                        "소유구분",
                        "토지소유자",
                        "지분",
                    ]
                ]

                edited_df = st.data_editor(
                    display_df,
                    key="excel_result_editor",
                    hide_index=True,
                    use_container_width=True,
                    num_rows="fixed",
                    height=min(
                        720,
                        max(220, 44 * len(display_df) + 90)
                    ),
                    disabled=disabled_columns,
                    column_config={
                        "기준년도": st.column_config.TextColumn(
                            "기준년도",
                            width="small"
                        ),
                        "사업지구명": st.column_config.TextColumn(
                            "사업지구명",
                            width="medium"
                        ),
                        "종전\n읍면": st.column_config.TextColumn(
                            "종전\n읍면",
                            width="small"
                        ),
                        "종전\n동리": st.column_config.TextColumn(
                            "종전\n동리",
                            width="small"
                        ),
                        "종전\n지목": st.column_config.TextColumn(
                            "종전\n지목",
                            width="small"
                        ),
                        "종전\n지번": st.column_config.TextColumn(
                            "종전\n지번",
                            width="small"
                        ),
                        "종전\n면적": st.column_config.NumberColumn(
                            "종전\n면적",
                            format="%.1f",
                            width="small"
                        ),
                        "확정\n읍면": st.column_config.TextColumn(
                            "확정\n읍면",
                            width="small"
                        ),
                        "확정\n동리": st.column_config.TextColumn(
                            "확정\n동리",
                            width="small"
                        ),
                        "확정\n지목": st.column_config.TextColumn(
                            "확정\n지목",
                            width="small"
                        ),
                        "확정\n지번": st.column_config.TextColumn(
                            "확정\n지번",
                            width="small"
                        ),
                        "확정\n면적": st.column_config.NumberColumn(
                            "확정\n면적",
                            format="%.1f",
                            width="small"
                        ),
                        "증감면적": st.column_config.NumberColumn(
                            "증감면적",
                            format="%.1f",
                            width="small"
                        ),
                        "감소면적": st.column_config.NumberColumn(
                            "감소면적",
                            format="%.1f",
                            width="small"
                        ),
                        "증감여부": st.column_config.TextColumn(
                            "증감여부",
                            width="small"
                        ),
                        "처리상태": st.column_config.TextColumn(
                            "처리상태",
                            width="small"
                        ),
                        "소유사항": st.column_config.TextColumn(
                            "소유사항",
                            width="small"
                        ),
                        "공유 또는 상속인 수": st.column_config.NumberColumn(
                            "공유 또는 상속인 수",
                            format="%d",
                            width="medium"
                        ),
                        "소유구분": st.column_config.SelectboxColumn(
                            "소유구분",
                            options=OWNERSHIP_OPTIONS,
                            width="medium",
                            required=False
                        ),
                        "토지소유자": st.column_config.TextColumn(
                            "토지소유자",
                            width="medium"
                        ),
                        "지분": st.column_config.TextColumn(
                            "지분",
                            width="small",
                            help="예: 0.5 또는 1/2"
                        ),
                    },
                )

                # -------------------------------------------------
                # 편집된 소유자 정보 저장
                # -------------------------------------------------
                for display_idx, meta in enumerate(row_meta):

                    source_idx = meta["source_idx"]
                    owner_idx = meta["owner_idx"]

                    if owner_idx is None:
                        continue

                    if display_idx >= len(edited_df):
                        continue

                    edited_row = edited_df.iloc[display_idx]

                    owner = st.session_state["excel_owner_data"][
                        f"excel_owner_{source_idx}"
                    ][owner_idx]

                    selected_type = edited_row.get(
                        "소유구분",
                        owner.get("소유구분", "1. 개인")
                    )

                    if (
                        selected_type not in OWNERSHIP_OPTIONS
                        and str(selected_type).strip() in OWNERSHIP_CODE_MAP
                    ):
                        selected_type = (
                            f"{str(selected_type).strip()}. "
                            f"{OWNERSHIP_CODE_MAP[str(selected_type).strip()]}"
                        )

                    if selected_type not in OWNERSHIP_OPTIONS:
                        selected_type = "1. 개인"

                    owner["소유구분"] = selected_type
                    owner["토지소유자"] = (
                        ""
                        if pd.isna(edited_row.get("토지소유자", ""))
                        else str(edited_row.get("토지소유자", ""))
                    )
                    owner["지분"] = (
                        ""
                        if pd.isna(edited_row.get("지분", ""))
                        else str(edited_row.get("지분", ""))
                    )

                # -------------------------------------------------
                # 단독/0명 행의 소유자 정보도 저장
                # -------------------------------------------------
                for display_idx, meta in enumerate(row_meta):

                    if not meta["is_base"]:
                        continue

                    source_idx = meta["source_idx"]
                    owner_count = meta["owner_count"]

                    # 공유/상속 기본행은 소유자 입력 대상이 아닙니다.
                    if owner_count > 0:
                        continue

                    if display_idx >= len(edited_df):
                        continue

                    edited_row = edited_df.iloc[display_idx]

                    owner = st.session_state["excel_owner_data"][
                        f"excel_owner_{source_idx}"
                    ][0]

                    selected_type = edited_row.get(
                        "소유구분",
                        owner.get("소유구분", "1. 개인")
                    )

                    if (
                        selected_type not in OWNERSHIP_OPTIONS
                        and str(selected_type).strip() in OWNERSHIP_CODE_MAP
                    ):
                        selected_type = (
                            f"{str(selected_type).strip()}. "
                            f"{OWNERSHIP_CODE_MAP[str(selected_type).strip()]}"
                        )

                    if selected_type not in OWNERSHIP_OPTIONS:
                        selected_type = "1. 개인"

                    owner["소유구분"] = selected_type
                    owner["토지소유자"] = (
                        ""
                        if pd.isna(edited_row.get("토지소유자", ""))
                        else str(edited_row.get("토지소유자", ""))
                    )
                    owner["지분"] = (
                        ""
                        if pd.isna(edited_row.get("지분", ""))
                        else str(edited_row.get("지분", ""))
                    )

                # -------------------------------------------------
                # 공유/상속 지분 합계 검증
                # -------------------------------------------------
                for idx, row in result_df.iterrows():

                    ownership_type = str(
                        row.get("소유사항", "")
                    ).strip()

                    if ownership_type not in ["공유", "상속"]:
                        continue

                    owners = st.session_state["excel_owner_data"].get(
                        f"excel_owner_{idx}",
                        []
                    )

                    owner_rows_for_validation = [
                        {
                            "소유구분": owner.get("소유구분", ""),
                            "토지소유자": owner.get("토지소유자", ""),
                            "지분": parse_share(owner.get("지분", ""))
                        }
                        for owner in owners
                    ]

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