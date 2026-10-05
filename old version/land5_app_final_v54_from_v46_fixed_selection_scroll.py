import streamlit as st
import pandas as pd
import numpy as np
import hashlib
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

        # ---------------------------------------------------------
        # 완전히 빈 행 제거
        #
        # Excel에서 내용 없이 서식만 남아 있는 행도 pandas에서는
        # 실제 데이터 행처럼 읽힐 수 있습니다. 이 행을 계산 전에
        # 제거하지 않으면 빈 행이 여러 개의 결과 행으로 생성됩니다.
        # 첫 열부터 마지막 열까지 값이 하나라도 있는 행만 유지합니다.
        # ---------------------------------------------------------
        if not df.empty:
            non_empty_mask = df.apply(
                lambda row: any(
                    not (
                        pd.isna(value)
                        or (isinstance(value, str) and value.strip() == "")
                    )
                    for value in row.tolist()
                ),
                axis=1,
            )
            df = df.loc[non_empty_mask].copy().reset_index(drop=True)

        # 기준년도는 숫자로 변환하지 않고 텍스트로 유지합니다.
        if "기준년도" in df.columns:
            df["기준년도"] = df["기준년도"].apply(
                lambda value: "" if pd.isna(value) else str(value).strip()
            )

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

        # ---------------------------------------------------------
        # 면적 증감 계산
        #
        # 최초 프로그램의 기준대로 각 행의
        #   종전 면적 - 확정 면적
        # 을 계산합니다.
        #
        # 결과가 +이면 면적이 감소한 것이므로 감소면적에 표시하고,
        # 결과가 -이면 면적이 증가한 것이므로 증가면적에 표시합니다.
        # 두 항목 모두 화면에는 절대값으로 표시합니다.
        #
        # 이전 버전에서는 확정_ID로 종전 면적을 다시 매칭하면서
        # 매칭되지 않는 경우 종전 면적이 0으로 처리되어
        # 확정 면적 전체가 증가면적으로 표시되는 문제가 있었습니다.
        # 따라서 여기서는 해당 행의 종전_면적과 확정_면적을 직접 비교합니다.
        # ---------------------------------------------------------
        area_difference = df["종전_면적"] - df["확정_면적"]

        # + 값 → 감소면적, - 값 → 증가면적
        df["증가면적"] = np.where(
            area_difference < 0,
            np.abs(area_difference),
            0
        )

        df["감소면적"] = np.where(
            area_difference > 0,
            np.abs(area_difference),
            0
        )

        # 업로드 엑셀의 면적 정밀도(소수 첫째 자리)에 맞춰
        # 계산 결과도 소수 첫째 자리까지 반올림합니다.
        df["종전_면적"] = df["종전_면적"].round(1)
        df["확정_면적"] = df["확정_면적"].round(1)
        df["증가면적"] = df["증가면적"].round(1).abs()
        df["감소면적"] = df["감소면적"].round(1).abs()

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


def render_excel_preview(df, max_rows=5):
    """업로드 엑셀 미리보기.

    병합 셀로 인해 pandas가 만든 'Unnamed: ..._level_1' 헤더는 숨기고,
    실제 Excel의 병합 구조처럼 2단 헤더를 HTML table의 rowspan/colspan으로 표시합니다.
    추가된 열도 그대로 모두 표시합니다.
    """
    preview = df.head(max_rows).copy()

    def clean_header(value):
        text = "" if pd.isna(value) else str(value).strip()
        if text.startswith("Unnamed:"):
            return ""
        return text

    # 2단 헤더(MultiIndex)
    if isinstance(preview.columns, pd.MultiIndex) and preview.columns.nlevels >= 2:
        levels = preview.columns.tolist()
        ncols = len(levels)

        html = [
            '<div style="overflow-x:auto; width:100%;">',
            '<table class="excel-preview-table">',
            '<thead><tr>'
        ]

        # 1행: 동일한 상위 헤더는 colspan, 하위 헤더가 Unnamed이면 rowspan=2
        i = 0
        while i < ncols:
            first = clean_header(levels[i][0])
            second = clean_header(levels[i][1])

            # 첫 번째 헤더가 비어 있으면 열 단위로 표시
            if not first:
                label = second
                html.append(
                    f'<th rowspan="2">{escape(label)}</th>'
                )
                i += 1
                continue

            # 같은 상위 헤더가 연속된 범위 찾기
            j = i + 1
            while j < ncols and clean_header(levels[j][0]) == first:
                j += 1

            span = j - i
            all_second_empty = all(
                not clean_header(levels[k][1]) for k in range(i, j)
            )

            if span == 1 and all_second_empty:
                html.append(
                    f'<th rowspan="2">{escape(first)}</th>'
                )
            else:
                html.append(
                    f'<th colspan="{span}">{escape(first)}</th>'
                )
            i = j

        html.append('</tr><tr>')

        # 2행: 하위 헤더가 실제로 존재하는 열만 출력
        i = 0
        while i < ncols:
            first = clean_header(levels[i][0])
            second = clean_header(levels[i][1])

            if not first:
                i += 1
                continue

            j = i + 1
            while j < ncols and clean_header(levels[j][0]) == first:
                j += 1

            span = j - i
            if not (span == 1 and not second):
                for k in range(i, j):
                    sub = clean_header(levels[k][1])
                    html.append(f'<th>{escape(sub)}</th>')
            i = j

        html.append('</tr></thead><tbody>')

    else:
        # 1단 헤더
        html = [
            '<div style="overflow-x:auto; width:100%;">',
            '<table class="excel-preview-table">',
            '<thead><tr>'
        ]
        for col in preview.columns:
            label = clean_header(col)
            html.append(f'<th>{escape(label)}</th>')
        html.append('</tr></thead><tbody>')

    # 데이터
    for _, row in preview.iterrows():
        html.append('<tr>')
        for value in row.tolist():
            if pd.isna(value):
                text = ""
            else:
                text = str(value)
            html.append(f'<td>{escape(text)}</td>')
        html.append('</tr>')

    html.extend([
        '</tbody></table></div>',
        '<style>',
        '.excel-preview-table { border-collapse: collapse; width: 100%; min-width: max-content; font-size: 13px; }',
        '.excel-preview-table th, .excel-preview-table td { border: 1px solid #d9d9d9; padding: 6px 8px; text-align: center; white-space: nowrap; }',
        '.excel-preview-table th { background: #f3f4f6; font-weight: 600; }',
        '.excel-preview-table td { background: #ffffff; }',
        '</style>'
    ])

    return "".join(html)


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


# 소유사항 값을 업로드 Excel의 여러 표현에서 표준 형식으로 변환합니다.
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


def normalize_owner_category(value):
    """업로드 Excel의 소유구분 값을 결과표용 표준값으로 변환합니다."""
    text = "" if value is None or pd.isna(value) else str(value).strip()
    if text in OWNERSHIP_OPTIONS:
        return text
    if text in OWNERSHIP_CODE_MAP:
        return f"{text}. {OWNERSHIP_CODE_MAP[text]}"
    if text in OWNERSHIP_MAP.values():
        for option, plain in OWNERSHIP_MAP.items():
            if plain == text:
                return option
    return "1. 개인"


def clean_owner_value(value):
    """업로드 Excel의 소유자/지분 값을 문자열로 안전하게 정리합니다."""
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


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

        # 업로드 파일이 바뀌면 이전 파일에서 사용하던 숫자/소유자 상태를
        # 새 파일에 재사용하지 않도록 파일 내용 기준으로 상태를 초기화합니다.
        try:
            _uploaded_bytes = uploaded_file.getvalue()
            _uploaded_signature = hashlib.sha256(_uploaded_bytes).hexdigest()
        except Exception:
            _uploaded_signature = f"{uploaded_file.name}_{getattr(uploaded_file, 'size', '')}"

        if st.session_state.get("excel_uploaded_signature") != _uploaded_signature:
            # 새 파일이면 이전 결과와 공유/상속인 수 위젯 상태를 모두 제거합니다.
            for _key in list(st.session_state.keys()):
                if (
                    _key.startswith("excel_count_")
                    or _key.startswith("excel_owner_")
                    or _key.startswith("excel_fold_")
                    or _key.startswith("excel_child_")
                    or _key.startswith("excel_result_edit_")
                    or _key.startswith("excel_ownership_type_")
                ):
                    del st.session_state[_key]

            for _key in [
                "excel_result_df",
                "excel_result_source_df",
                "excel_owner_data",
                "excel_owner_collapsed",
                "share_dialog_signature",
                "pending_share_dialog",
            ]:
                st.session_state.pop(_key, None)

            st.session_state["excel_uploaded_signature"] = _uploaded_signature

        try:
            # ---------------------------------------------------------
            # 업로드 파일은 앞으로 기존 항목 외에 새로운 열/항목이
            # 계속 추가될 수 있으므로, 미리보기 단계에서는 특정
            # 컬럼 목록에 의존하지 않습니다.
            #
            # 1) 우선 기존 test3 형태인 2줄 헤더로 읽습니다.
            # 2) 2줄 헤더가 없는 파일도 업로드 자체는 가능하도록
            #    1줄 헤더로 다시 읽습니다.
            # 3) 변환할 수 없는 특이한 헤더가 있더라도 원본 열을
            #    그대로 유지하여 미리보기는 항상 표시합니다.
            # ---------------------------------------------------------
            uploaded_file.seek(0)
            try:
                df_raw = pd.read_excel(
                    uploaded_file,
                    header=[0, 1],
                    dtype=str,
                )
                preview_header_mode = "2줄 헤더"
            except Exception:
                uploaded_file.seek(0)
                df_raw = pd.read_excel(
                    uploaded_file,
                    header=0,
                    dtype=str,
                )
                preview_header_mode = "1줄 헤더"

            # 미리보기용 데이터는 업로드된 열을 임의로 삭제하지 않습니다.
            # 기존 항목 + 앞으로 추가될 항목까지 모두 표시합니다.
            df_input = df_raw.copy()

            try:
                converted_df = convert_excel_columns(df_raw.copy())
            except Exception:
                # 알 수 없는 헤더 구조라도 업로드 미리보기는 계속 표시합니다.
                converted_df = df_raw.copy()

            # 계산 버튼에서 사용할 내부 데이터
            df_input = converted_df

            st.subheader("📄 업로드 데이터 미리보기")
            st.caption(
                f"업로드된 모든 항목을 표시합니다. ({preview_header_mode}) "
                "기존 항목에 새로운 열이 추가되어도 미리보기는 그대로 표시됩니다."
            )
            st.markdown(
                render_excel_preview(df_raw),
                unsafe_allow_html=True
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
                    # ---------------------------------------------------------
                    # 업로드 Excel의 공유/상속 소유자 하위행 분리
                    #
                    # 업로드 파일이 아래처럼 되어 있는 경우를 지원합니다.
                    #
                    # 공유 | 3 |            | ...
                    #      |   | 1. 개인 | 홍길동 | 1/3
                    #      |   | 1. 개인 | 김철수 | 1/3
                    #      |   | 1. 개인 | 이영희 | 1/3
                    #
                    # 위 3개 하위행은 새로운 토지 행이 아니라 소유자 추가행입니다.
                    # 따라서 면적 계산 대상에서 제외하고, 공유/상속인 수에 맞춰
                    # 결과표의 추가 소유자 행에만 넣습니다.
                    # ---------------------------------------------------------
                    def _clean_excel_value(value):
                        if value is None or pd.isna(value):
                            return ""
                        return str(value).strip()

                    land_identity_columns = [
                        "기준년도", "사업지구명",
                        "종전_읍면", "종전_동리", "종전_지목", "종전_지번", "종전_면적",
                        "확정_읍면", "확정_동리", "확정_지목", "확정_지번", "확정_면적",
                    ]

                    def _is_owner_continuation_row(row):
                        # ┗ 기호가 어느 열에 들어가더라도 하위 소유자 행으로 인식합니다.
                        has_fold_marker = any(
                            _clean_excel_value(value) == "┗"
                            for value in row.tolist()
                        )

                        owner_has_value = any(
                            _clean_excel_value(row.get(col, ""))
                            for col in ["소유구분", "토지소유자", "지분"]
                        )

                        land_has_value = any(
                            _clean_excel_value(row.get(col, ""))
                            for col in land_identity_columns
                            if col in row.index
                        )

                        # 토지 기본정보가 전혀 없고 소유자 정보만 있는 행은
                        # 실제 토지 행이 아니라 공유/상속 소유자 추가행입니다.
                        return (has_fold_marker or owner_has_value) and not land_has_value

                    calculation_rows = []
                    uploaded_owner_seeds = []
                    source_rows = df_input.reset_index(drop=True)
                    source_idx = 0

                    while source_idx < len(source_rows):
                        current_row = source_rows.iloc[source_idx]

                        if _is_owner_continuation_row(current_row):
                            # 선행 기본행이 없는 고아 소유자 행은 결과 토지행으로 만들지 않습니다.
                            source_idx += 1
                            continue

                        base_row = current_row.copy()
                        calculation_rows.append(base_row)

                        ownership_type_seed = normalize_ownership(
                            base_row.get("소유사항", "")
                        )
                        raw_seed_count = base_row.get("공유_상속인수", 0)
                        try:
                            seed_count = max(0, int(float(raw_seed_count)))
                        except (TypeError, ValueError):
                            seed_count = 0

                        if ownership_type_seed == "1. 단독":
                            seed_count = 0

                        owner_seed_rows = []
                        next_idx = source_idx + 1

                        # 공유/상속 기본행 바로 아래에 붙어 있는 소유자 행을
                        # 실제 인원수만큼만 읽습니다.
                        if ownership_type_seed in ["2. 공유", "3. 상속"] and seed_count > 0:
                            while next_idx < len(source_rows) and len(owner_seed_rows) < seed_count:
                                candidate = source_rows.iloc[next_idx]

                                if not _is_owner_continuation_row(candidate):
                                    break

                                owner_seed_rows.append({
                                    "소유구분": normalize_owner_category(
                                        candidate.get("소유구분", "")
                                    ),
                                    "토지소유자": _clean_excel_value(
                                        candidate.get("토지소유자", "")
                                    ),
                                    "지분": _clean_excel_value(
                                        candidate.get("지분", "")
                                    ),
                                })
                                next_idx += 1

                        # 기본행 자체에 소유자 정보가 있는 경우에는
                        # 공유/상속의 첫 번째 소유자로 사용합니다.
                        base_owner_has_value = any(
                            _clean_excel_value(base_row.get(col, ""))
                            for col in ["소유구분", "토지소유자", "지분"]
                        )

                        if base_owner_has_value:
                            base_owner = {
                                "소유구분": normalize_owner_category(
                                    base_row.get("소유구분", "")
                                ),
                                "토지소유자": _clean_excel_value(
                                    base_row.get("토지소유자", "")
                                ),
                                "지분": _clean_excel_value(
                                    base_row.get("지분", "")
                                ),
                            }
                            if seed_count == 0:
                                owner_seed_rows = [base_owner]
                            elif not owner_seed_rows:
                                owner_seed_rows.insert(0, base_owner)
                            elif len(owner_seed_rows) < seed_count:
                                owner_seed_rows.insert(0, base_owner)

                        uploaded_owner_seeds.append(owner_seed_rows[:seed_count] if seed_count > 0 else owner_seed_rows[:1])

                        # 소비한 하위 소유자 행은 계산 대상에서 다시 읽지 않습니다.
                        source_idx = next_idx if next_idx > source_idx + 1 else source_idx + 1

                    calculation_input = pd.DataFrame(calculation_rows, columns=df_input.columns)
                    calculated_df = AreaAnalysisStrategy.calculate(calculation_input)

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

                    # 공유/상속인 수는 결과표에서 숫자 위젯으로 자유롭게 수정할 수 있어야
                    # 하므로, 업로드 Excel에서 pandas가 StringDtype/float64 등으로
                    # 추론한 dtype에 직접 정수값을 대입하지 않도록 object 타입으로 고정합니다.
                    if "공유_상속인수" in result_df.columns:
                        result_df["공유_상속인수"] = result_df["공유_상속인수"].astype(object)
                    if "공유_상속인수" in calculated_df.columns:
                        calculated_df["공유_상속인수"] = calculated_df["공유_상속인수"].astype(object)

                    st.session_state["excel_result_df"] = result_df
                    st.session_state["excel_result_source_df"] = calculated_df
                    st.session_state["excel_uploaded_owner_seeds"] = uploaded_owner_seeds

                    # 업로드 Excel의 소유자 정보를 결과표 초기값으로 반영합니다.
                    initial_owner_data = {}

                    for result_pos, seed_row in result_df.reset_index(drop=True).iterrows():
                        seed_type = normalize_ownership(seed_row.get("소유사항", ""))
                        raw_count = seed_row.get("공유_상속인수", 0)
                        try:
                            seed_count = max(0, int(float(raw_count)))
                        except (TypeError, ValueError):
                            seed_count = 0
                        if seed_type == "1. 단독":
                            seed_count = 0

                        seeded = []
                        if result_pos < len(uploaded_owner_seeds):
                            seeded = uploaded_owner_seeds[result_pos]

                        required_seed = max(1, seed_count)
                        owners_seed = list(seeded[:seed_count]) if seed_count > 0 else list(seeded[:1])

                        # 단독 행에 소유자 정보가 기본행에 있는 경우
                        # seed에서 빠졌더라도 기본행 값을 직접 반영합니다.
                        if seed_count == 0 and not owners_seed:
                            owners_seed = [{
                                "소유구분": normalize_owner_category(seed_row.get("소유구분", "")),
                                "토지소유자": _clean_excel_value(seed_row.get("토지소유자", "")),
                                "지분": _clean_excel_value(seed_row.get("지분", "")),
                            }]

                        while len(owners_seed) < required_seed:
                            owners_seed.append({
                                "소유구분": "1. 개인",
                                "토지소유자": "",
                                "지분": "",
                            })

                        initial_owner_data[f"excel_owner_{result_pos}"] = owners_seed[:required_seed]

                    st.session_state["excel_owner_data"] = initial_owner_data
                    st.session_state["excel_owner_collapsed"] = {}
                    
                    st.success("✅ 계산이 완료되었습니다.")

            if "excel_result_df" in st.session_state:

                result_df = st.session_state["excel_result_df"].copy()

                if "excel_selected_rows" not in st.session_state:
                    st.session_state["excel_selected_rows"] = set()

                def _clear_result_widget_states():
                    prefixes = ("excel_result_edit_", "excel_ownership_type_", "excel_count_", "excel_owner_type_", "excel_owner_name_", "excel_owner_share_", "excel_child_owner_type_", "excel_child_owner_name_", "excel_child_owner_share_", "excel_row_check_", "excel_owner_row_check_")
                    for key in list(st.session_state.keys()):
                        if key.startswith(prefixes):
                            del st.session_state[key]

                def _make_empty_result_row():
                    return {"기준년도":"", "사업지구명":"", "종전_읍면":"", "종전_동리":"", "종전_지목":"", "종전_지번":"", "종전_면적":0.0, "확정_읍면":"", "확정_동리":"", "확정_지목":"", "확정_지번":"", "확정_면적":0.0, "증가면적":0.0, "감소면적":0.0, "증감여부":"변동없음", "처리상태":"유지", "소유사항":"단독", "공유_상속인수":0, "소유구분":"1. 개인", "토지소유자":"", "지분":"1"}

                title_col, action_col = st.columns([0.78, 0.22], gap="small")
                with title_col:
                    st.subheader("📊 계산 결과 및 소유자 입력")
                with action_col:
                    add_col, delete_col = st.columns(2, gap="small")
                    with add_col:
                        add_row_clicked = st.button("행 추가", key="excel_add_row", use_container_width=True)
                    with delete_col:
                        delete_row_clicked = st.button("행 삭제", key="excel_delete_row", use_container_width=True)

                if add_row_clicked:
                    result_columns = list(result_df.columns)
                    new_row = _make_empty_result_row()
                    new_row = {c: new_row.get(c, "") for c in result_columns}
                    st.session_state["excel_result_df"] = pd.concat([st.session_state["excel_result_df"], pd.DataFrame([new_row], columns=result_columns)], ignore_index=True)
                    if "excel_result_source_df" in st.session_state:
                        src = st.session_state["excel_result_source_df"]
                        st.session_state["excel_result_source_df"] = pd.concat([src, pd.DataFrame([{c:new_row.get(c, "") for c in src.columns}], columns=src.columns)], ignore_index=True)
                    new_idx = len(st.session_state["excel_result_df"]) - 1
                    st.session_state.setdefault("excel_owner_data", {})[f"excel_owner_{new_idx}"] = [{"소유구분":"1. 개인", "토지소유자":"", "지분":"1"}]
                    st.session_state["excel_selected_rows"] = set()
                    _clear_result_widget_states()
                    st.rerun()

                if delete_row_clicked:
                    selected_items = st.session_state.get("excel_selected_rows", set())

                    # 구버전에서 저장된 정수형 선택값도 기본행 선택으로 호환합니다.
                    normalized_selected = set()
                    for item in selected_items:
                        if isinstance(item, tuple):
                            normalized_selected.add(item)
                        elif isinstance(item, int):
                            normalized_selected.add(("base", item))
                        elif isinstance(item, str) and item.isdigit():
                            normalized_selected.add(("base", int(item)))

                    selected_base = sorted({
                        item[1] for item in normalized_selected
                        if len(item) == 2 and item[0] == "base" and isinstance(item[1], int)
                        and 0 <= item[1] < len(result_df)
                    })

                    selected_owner = {}
                    for item in normalized_selected:
                        if (
                            isinstance(item, tuple) and len(item) == 3
                            and item[0] == "owner"
                            and isinstance(item[1], int) and isinstance(item[2], int)
                            and 0 <= item[1] < len(result_df)
                        ):
                            selected_owner.setdefault(item[1], set()).add(item[2])

                    changed = False

                    # 기본행이 선택되면 해당 필지 전체를 삭제합니다.
                    if selected_base:
                        selected_set = set(selected_base)
                        st.session_state["excel_result_df"] = (
                            st.session_state["excel_result_df"]
                            .iloc[[i not in selected_set for i in range(len(result_df))]]
                            .reset_index(drop=True)
                        )
                        if "excel_result_source_df" in st.session_state:
                            src = st.session_state["excel_result_source_df"]
                            st.session_state["excel_result_source_df"] = (
                                src.iloc[[i not in selected_set for i in range(len(src))]]
                                .reset_index(drop=True)
                            )

                        old = st.session_state.get("excel_owner_data", {})
                        new = {}
                        for new_i, old_i in enumerate(i for i in range(len(result_df)) if i not in selected_set):
                            if f"excel_owner_{old_i}" in old:
                                new[f"excel_owner_{new_i}"] = old[f"excel_owner_{old_i}"]
                        st.session_state["excel_owner_data"] = new
                        changed = True

                        # 기본행 삭제로 인해 남은 owner 선택값은 모두 초기화합니다.
                        selected_owner = {}

                    # 공유/상속 추가행이 선택되면 해당 소유자 행만 삭제합니다.
                    # 마지막 1명까지 삭제하면 공유/상속의 인원수가 0명이 되므로,
                    # 최소 1명은 유지합니다.
                    if selected_owner and not selected_base:
                        result_state = st.session_state["excel_result_df"]
                        for row_idx, owner_indices in sorted(selected_owner.items()):
                            if not (0 <= row_idx < len(result_state)):
                                continue

                            ownership_plain = normalize_ownership(
                                result_state.loc[row_idx, "소유사항"]
                            )
                            if ownership_plain not in ["2. 공유", "3. 상속"]:
                                continue

                            owner_key = f"excel_owner_{row_idx}"
                            owners = st.session_state.get("excel_owner_data", {}).get(owner_key, [])
                            if len(owners) <= 1:
                                continue

                            valid_remove = sorted(
                                {i for i in owner_indices if 0 <= i < len(owners)},
                                reverse=True,
                            )
                            # 최소 1명의 소유자는 남깁니다.
                            max_remove = max(0, len(owners) - 1)
                            valid_remove = valid_remove[:max_remove]
                            if not valid_remove:
                                continue

                            for owner_idx in valid_remove:
                                del owners[owner_idx]

                            new_count = len(owners)
                            result_state["공유_상속인수"] = result_state["공유_상속인수"].astype(object)
                            result_state.loc[row_idx, "공유_상속인수"] = new_count

                            if "excel_result_source_df" in st.session_state:
                                src = st.session_state["excel_result_source_df"]
                                if "공유_상속인수" in src.columns and row_idx < len(src):
                                    src["공유_상속인수"] = src["공유_상속인수"].astype(object)
                                    src.loc[row_idx, "공유_상속인수"] = new_count

                            changed = True

                    if changed:
                        st.session_state["excel_selected_rows"] = set()
                        _clear_result_widget_states()
                        st.rerun()

                result_df = st.session_state["excel_result_df"].copy()

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

                    # 면적 계열은 소수 첫째 자리까지 계산하되,
                    # 소수 부분이 0이면 .0을 표시하지 않습니다.
                    # 예: 100.0 -> 100, 100.5 -> 100.5
                    if isinstance(value, (float, np.floating)):
                        number = float(value)
                        if np.isclose(number, round(number), atol=1e-9):
                            return str(int(round(number)))
                        return f"{number:.1f}"

                    if isinstance(value, (int, np.integer)):
                        return str(int(value))

                    # Excel을 dtype=str로 읽었더라도 계산 결과에 숫자
                    # 문자열이 남아 있는 경우를 안전하게 정리합니다.
                    text = str(value).strip()
                    try:
                        number = float(text)
                        if np.isclose(number, round(number), atol=1e-9):
                            return str(int(round(number)))
                        # 면적 표시 정밀도는 소수 첫째 자리까지
                        return f"{number:.1f}"
                    except (ValueError, TypeError):
                        return text

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
                # 첫 번째 열은 행 선택 전용 열입니다.
                # 기존 21개 데이터 열은 그대로 유지하고 앞에 1개 열을 추가합니다.
                col_widths = [
                    0.45,
                    0.80, 1.25,
                    0.70, 1.15, 0.80, 1.40, 0.80,
                    0.70, 1.15, 0.80, 1.40, 0.80,
                    0.90, 0.90, 0.85, 0.85,
                    0.95, 1.15,
                    1.10, 1.20, 1.80
                ]
                total_col_width = sum(col_widths)
                result_table_width = 1900
                # v39: 선택 열도 실제 표의 첫 번째 열로 포함합니다.
                # 헤더와 모든 데이터 행이 동일한 22개 열 폭을 사용합니다.
                header_colgroup = "".join(
                    f'<col style="width:{w / total_col_width * 100:.4f}%">'
                    for w in col_widths
                )

                static_keys = [
                    "기준년도", "사업지구명",
                    "종전_읍면", "종전_동리", "종전_지목", "종전_지번", "종전_면적",
                    "확정_읍면", "확정_동리", "확정_지목", "확정_지번", "확정_면적",
                    "증가면적", "감소면적", "증감여부", "처리상태"
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
                    /* 결과표 바깥에는 가로 스크롤을 만들지 않습니다.
                       가로 스크롤은 오른쪽 데이터 영역(result_scroll) 하나만 담당합니다. */
                    div.st-key-result_table {
                        overflow: visible !important;
                        width: 100% !important;
                        max-width: 100% !important;
                        padding: 0 0 8px 0 !important;
                    }
                    div.st-key-result_table > div {
                        min-width: 0 !important;
                        width: 100% !important;
                        max-width: 100% !important;
                    }
                    div.st-key-result_table .stHorizontalBlock {
                        min-width: 0 !important;
                        width: 100% !important;
                        max-width: 100% !important;
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
                    /* 헤더 셀은 데이터 셀과 달리 모서리를 둥글게 하지 않습니다. */
                    .land-header-table th,
                    .land-select-header {
                        border-radius: 0 !important;
                    }
                    .land-select-header {
                        box-sizing: border-box !important;
                        display: flex !important;
                        width: 100% !important;
                        min-width: 100% !important;
                        max-width: 100% !important;
                        height: 68px !important;
                        min-height: 68px !important;
                        max-height: 68px !important;
                        margin: 0 !important;
                        padding: 4px 2px !important;
                        border: 1px solid #b7b7b7 !important;
                        border-radius: 0 !important;
                        background: #eeeeee !important;
                        color: #202020 !important;
                        align-items: center !important;
                        justify-content: center !important;
                        font-size: 13px !important;
                        font-weight: 700 !important;
                        line-height: 1.2 !important;
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
                        border-radius: 6px !important;
                        border: 1px solid #c8c8c8 !important;
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
                    /* 입력 위젯의 라벨만 숨깁니다. 체크박스 라벨/표시는 유지합니다. */
                    div.st-key-result_table [data-testid="stTextInput"] label,
                    div.st-key-result_table [data-testid="stSelectbox"] label,
                    div.st-key-result_table [data-testid="stNumberInput"] label {
                        display: none !important;
                    }
                    /* 선택 체크박스: 셀 전체를 사용하고 실제 체크박스도 중앙에 배치합니다. */
                    div.st-key-result_table [data-testid="stCheckbox"] {
                        min-height: 40px !important;
                        height: 40px !important;
                        width: 100% !important;
                        min-width: 100% !important;
                        max-width: 100% !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: center !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        box-sizing: border-box !important;
                    }
                    div.st-key-result_table [data-testid="stCheckbox"] > label {
                        display: flex !important;
                        align-items: center !important;
                        justify-content: center !important;
                        width: 100% !important;
                        height: 40px !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        box-sizing: border-box !important;
                    }
                    div.st-key-result_table [data-testid="stCheckbox"] > label > div {
                        margin: 0 auto !important;
                        padding: 0 !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: center !important;
                    }
                    div.st-key-result_table [data-testid="stCheckbox"] input {
                        margin: 0 !important;
                    }

                    /*
                       [중요] 선택 열 고정은 checkbox 위젯이 아니라
                       '실제 스크롤 영역(result_table)의 바로 아래에서 움직이는
                       선택 셀의 내부 래퍼'까지 sticky로 잡아야 합니다.

                       기존 v31은 바깥 column만 sticky로 지정했습니다.
                       Streamlit의 st.columns는 내부에 여러 wrapper를 만들기 때문에
                       column의 배경은 고정되어도 실제 checkbox 위젯은 가로 스크롤을
                       따라 움직일 수 있습니다.

                       아래에서는 column -> vertical block -> element container ->
                       checkbox까지 모두 같은 sticky 좌표계(left:0)에 둡니다.
                       숫자/+/- 등의 nested columns에는 :has() 조건으로 적용하지 않습니다.
                    */
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] > div [data-testid="stCheckbox"]) > [data-testid="column"]:first-child,
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] > div [data-testid="stCheckbox"]) > [data-testid="column"]:first-child > div,
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] > div [data-testid="stCheckbox"]) > [data-testid="column"]:first-child [data-testid="stVerticalBlock"],
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] > div [data-testid="stCheckbox"]) > [data-testid="column"]:first-child [data-testid="stElementContainer"] {
                        position: sticky !important;
                        left: 0 !important;
                        z-index: 60 !important;
                        background: #ffffff !important;
                        align-self: stretch !important;
                        box-shadow: 2px 0 0 rgba(0,0,0,.08) !important;
                    }

                    /* checkbox도 같은 sticky 좌표계에 둡니다.
                       width/height는 기존 셀 크기를 유지합니다. */
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] > div [data-testid="stCheckbox"]) > [data-testid="column"]:first-child [data-testid="stCheckbox"],
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] > div [data-testid="stCheckbox"]) > [data-testid="column"]:first-child [data-testid="stCheckbox"] > label {
                        position: sticky !important;
                        left: 0 !important;
                        z-index: 70 !important;
                        background: #ffffff !important;
                    }
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] > div [data-testid="stCheckbox"]) > [data-testid="column"]:first-child [data-testid="stCheckbox"] > label {
                        width: 100% !important;
                        min-width: 100% !important;
                        height: 40px !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: center !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        box-sizing: border-box !important;
                    }
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] > div [data-testid="stCheckbox"]) > [data-testid="column"]:first-child [data-testid="stCheckbox"] input {
                        width: 16px !important;
                        height: 16px !important;
                        margin: 0 !important;
                    }
                    div.st-key-result_table .land-header-table th.select-col {
                        position: sticky !important;
                        left: 0 !important;
                        z-index: 80 !important;
                        background: #eeeeee !important;
                    }
                    div.st-key-result_table .land-header-table th.select-col {
                        position: sticky !important;
                        left: 0 !important;
                        z-index: 30 !important;
                        background: #eeeeee !important;
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
                        min-width: 23px !important;
                        width: 23px !important;
                        height: 10px !important;
                        min-height: 10px !important;
                        max-height: 10px !important;
                        margin: 4px 1px 4px auto !important;
                        padding: 0 !important;
                        border: 1px solid #8c8c8c !important;
                        border-radius: 6px !important;
                        background: linear-gradient(#ffffff, #e2e2e2) !important;
                        color: #202020 !important;
                        font-size: 13px !important;
                        font-weight: 800 !important;
                        box-shadow: 0 2px 0 #888, 0 3px 5px rgba(0,0,0,.16) !important;
                        transform: none !important;
                    }
                    div.st-key-result_table [class*="st-key-fold_cell_"] button:hover {
                        background: linear-gradient(#f8f8f8, #d8d8d8) !important;
                    }
                    div.st-key-result_table [class*="st-key-fold_cell_"] button:active {
                        transform: translateY(1px) !important;
                        box-shadow: 0 1px 0 #888 !important;
                    }

                    /* 두 번째 이후 공유/상속 추가행은 테두리를 표시하지 않습니다.
                       첫 번째 추가행(child_first)은 기존 테두리를 유지합니다. */
                    div.st-key-result_table [class*="st-key-child_no_border_"] .land-fold-merged,
                    div.st-key-result_table [class*="st-key-child_no_border_"] input,
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-baseweb="select"] > div,
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="stTextInputRootElement"],
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="stSelectbox"] {
                        border: none !important;
                        box-shadow: none !important;
                    }
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-baseweb="select"] {
                        border: none !important;
                        box-shadow: none !important;
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
                        border: none;
                        background: #ffffff;
                        display: flex;
                        align-items: center;
                        justify-content: flex-end;
                        padding: 0 3px 0 0;
                    }
                    /* 공유/상속인 수 컨트롤은 결과표 전체 폭 규칙의 영향을 받지 않게 합니다. */
                    div.st-key-result_table [class*="st-key-count_controls_"] {
                        width: 100% !important;
                        min-width: 0 !important;
                        max-width: 100% !important;
                        padding: 0 !important;
                        margin: 0 !important;
                    }
                    div.st-key-result_table [class*="st-key-count_controls_"] .stHorizontalBlock {
                        width: 100% !important;
                        min-width: 0 !important;
                        max-width: 100% !important;
                        gap: 4px !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        align-items: center !important;
                    }
                    div.st-key-result_table [class*="st-key-count_controls_"] [data-testid="column"] {
                        min-width: 0 !important;
                        width: auto !important;
                        padding: 0 !important;
                        margin: 0 !important;
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

                    /* v33: 선택 열과 데이터 영역을 실제로 분리합니다.
                       왼쪽 선택 패널은 가로 스크롤 영역 밖에 두고,
                       오른쪽 데이터 패널만 별도의 가로 스크롤을 사용합니다. */
                    div.st-key-result_table > div:has(div.st-key-result_fixed),
                    div.st-key-result_table > div:has(div.st-key-result_scroll) {
                        min-width: 0 !important;
                    }
                    div.st-key-result_table [class*="st-key-result_fixed"] {
                        width: 100% !important;
                        min-width: 0 !important;
                        max-width: 100% !important;
                        overflow: visible !important;
                        position: relative !important;
                        z-index: 100 !important;
                        background: #ffffff !important;
                    }
                    div.st-key-result_table [class*="st-key-result_scroll"] {
                        width: 100% !important;
                        min-width: 0 !important;
                        max-width: 100% !important;
                        overflow-x: auto !important;
                        overflow-y: visible !important;
                        background: #ffffff !important;
                        box-sizing: border-box !important;
                    }
                    /* 실제 1900px 폭은 스크롤 영역 내부의 첫 콘텐츠 블록에만 적용합니다. */
                    div.st-key-result_table [class*="st-key-result_scroll"] > div > div {
                        min-width: 1900px !important;
                        width: 1900px !important;
                        max-width: none !important;
                    }
                    /* v35: 모든 셀/비활성 입력도 모서리를 둥글게 처리합니다. */
                    div.st-key-result_table .land-value-cell:not(.land-select-header),
                    div.st-key-result_table .land-fold-cell,
                    div.st-key-result_table .land-fold-merged,
                    div.st-key-result_table input,
                    div.st-key-result_table [data-testid="stTextInputRootElement"],
                    div.st-key-result_table [data-testid="stSelectbox"] [data-baseweb="select"] > div,
                    div.st-key-result_table [data-baseweb="select"] > div,
                    div.st-key-result_table [data-testid="stNumberInput"] input {
                        border: 1px solid #c8c8c8 !important;
                        border-radius: 6px !important;
                        box-shadow: none !important;
                    }
                    /* 비활성화된 입력/선택 위젯도 동일하게 둥근 셀 모양을 유지합니다. */
                    div.st-key-result_table input:disabled,
                    div.st-key-result_table [aria-disabled="true"],
                    div.st-key-result_table [data-disabled="true"],
                    div.st-key-result_table [data-baseweb="select"][aria-disabled="true"] > div {
                        border-radius: 6px !important;
                    }
                    /* v36: 선택 체크박스는 각 데이터 행과 동일한 40px 셀 안에서
                       가로/세로 모두 정확히 중앙에 배치합니다.
                       임의의 translateY는 사용하지 않습니다. */
                    /* v38: 선택 열의 모든 데이터 셀은 오른쪽 데이터 셀과 동일한 높이를
                       가지면서 각각 독립적인 테두리를 표시합니다. */
                    div.st-key-result_table [class*="st-key-base_select_cell_"],
                    div.st-key-result_table [class*="st-key-owner_select_cell_"] {
                        width: 100% !important;
                        min-width: 0 !important;
                        max-width: 100% !important;
                        height: 40px !important;
                        min-height: 40px !important;
                        max-height: 40px !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: center !important;
                        box-sizing: border-box !important;
                        border: 1px solid #c8c8c8 !important;
                        border-radius: 6px !important;
                        background: #ffffff !important;
                    }
                    div.st-key-result_table [class*="st-key-base_select_cell_"] [data-testid="stCheckbox"],
                    div.st-key-result_table [class*="st-key-owner_select_cell_"] [data-testid="stCheckbox"] {
                        width: 100% !important;
                        min-width: 100% !important;
                        max-width: 100% !important;
                        height: 40px !important;
                        min-height: 40px !important;
                        max-height: 40px !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: center !important;
                        transform: none !important;
                        box-sizing: border-box !important;
                    }
                    div.st-key-result_table [class*="st-key-base_select_cell_"] [data-testid="stCheckbox"] > label,
                    div.st-key-result_table [class*="st-key-owner_select_cell_"] [data-testid="stCheckbox"] > label {
                        width: 100% !important;
                        height: 40px !important;
                        min-height: 40px !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: center !important;
                        box-sizing: border-box !important;
                    }
                    div.st-key-result_table [class*="st-key-base_select_cell_"] [data-testid="stCheckbox"] > label > div,
                    div.st-key-result_table [class*="st-key-owner_select_cell_"] [data-testid="stCheckbox"] > label > div {
                        margin: 0 !important;
                        padding: 0 !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: center !important;
                    }
                    div.st-key-result_table [class*="st-key-base_select_cell_"] [data-testid="stCheckbox"] input,
                    div.st-key-result_table [class*="st-key-owner_select_cell_"] [data-testid="stCheckbox"] input {
                        width: 16px !important;
                        height: 16px !important;
                        margin: 0 !important;
                    }

                    /* v36: 선택 열은 별도 고정 패널 자체로 분리되어 있으므로
                       checkbox/헤더에는 sticky를 적용하지 않습니다. */
                    div.st-key-result_table [class*="st-key-result_fixed"] [data-testid="stCheckbox"],
                    div.st-key-result_table [class*="st-key-result_fixed"] [data-testid="stCheckbox"] > label,
                    div.st-key-result_table .land-header-table th.select-col {
                        position: relative !important;
                        left: auto !important;
                        transform: none !important;
                    }
                    div.st-key-result_table .land-header-table th.select-col {
                        border-radius: 0 !important;
                    }

                    /* v33에서는 checkbox 자체에 sticky를 주지 않습니다.
                       체크박스는 왼쪽 고정 패널 안에 있기 때문에 자연스럽게 고정됩니다. */
                    div.st-key-result_table [class*="st-key-result_fixed"] [data-testid="stCheckbox"],
                    div.st-key-result_table [class*="st-key-result_fixed"] [data-testid="stCheckbox"] > label {
                        position: relative !important;
                        left: auto !important;
                        z-index: 101 !important;
                    }
                    /* v39: 선택 열은 별도 패널이 아니라 전체 표의 첫 번째 열입니다.
                       position: sticky로 엑셀의 틀 고정과 같은 방식으로 가로 스크롤 시 고정합니다. */
                    div.st-key-result_table [class*="st-key-result_scroll"] {
                        width: 100% !important;
                        max-width: 100% !important;
                        min-width: 0 !important;
                        overflow-x: auto !important;
                        overflow-y: visible !important;
                        box-sizing: border-box !important;
                    }
                    div.st-key-result_table [class*="st-key-result_scroll"] > div {
                        width: 1900px !important;
                        min-width: 1900px !important;
                        max-width: none !important;
                    }
                    /* 헤더의 선택 셀도 동일한 첫 번째 열에 고정 */
                    div.st-key-result_table .land-header-table {
                        width: 1900px !important;
                        min-width: 1900px !important;
                        max-width: none !important;
                    }
                    div.st-key-result_table .land-header-table th.select-col {
                        position: sticky !important;
                        left: 0 !important;
                        z-index: 50 !important;
                        border-radius: 0 !important;
                        box-sizing: border-box !important;
                    }
                    /* 체크박스가 들어 있는 실제 첫 번째 Streamlit 열을 sticky 처리 */
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] [class*="st-key-base_select_cell_"]) > [data-testid="column"]:first-child,
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] [class*="st-key-owner_select_cell_"]) > [data-testid="column"]:first-child {
                        position: sticky !important;
                        left: 0 !important;
                        z-index: 40 !important;
                        background: #ffffff !important;
                        align-self: stretch !important;
                    }
                    div.st-key-result_table [class*="st-key-base_select_cell_"],
                    div.st-key-result_table [class*="st-key-owner_select_cell_"] {
                        width: 100% !important;
                        height: 40px !important;
                        min-height: 40px !important;
                        max-height: 40px !important;
                        box-sizing: border-box !important;
                    }
                    /* 선택 열은 각 행마다 독립된 셀을 가지며 체크박스는 그 셀 중앙 */
                    div.st-key-result_table [class*="st-key-base_select_cell_"] [data-testid="stCheckbox"],
                    div.st-key-result_table [class*="st-key-owner_select_cell_"] [data-testid="stCheckbox"] {
                        width: 100% !important;
                        height: 40px !important;
                        min-height: 40px !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: center !important;
                    }
                    /* 공유/상속 추가행의 입력 불가 영역(기준년도~공유/상속인 수)은
                       모든 추가행에서 테두리를 제거합니다. ┗ 버튼 자체의 테두리는 유지합니다. */
                    div.st-key-result_table [class*="st-key-child_first_"] .land-fold-merged,
                    div.st-key-result_table [class*="st-key-child_first_"] [class*="st-key-fold_cell_"],
                    div.st-key-result_table [class*="st-key-child_no_border_"] .land-fold-merged,
                    div.st-key-result_table [class*="st-key-child_no_border_"] [class*="st-key-fold_cell_"] {
                        border: none !important;
                        box-shadow: none !important;
                    }
                    /* 위 영역의 1~18번 실제 셀에 들어간 입력/표시 요소도 무테두리 */
                    div.st-key-result_table [class*="st-key-child_first_"] .land-fold-merged,
                    div.st-key-result_table [class*="st-key-child_no_border_"] .land-fold-merged {
                        border-radius: 0 !important;
                    }
                    /* ┗ 버튼은 2/3 크기를 유지 */
                    div.st-key-result_table [class*="st-key-fold_cell_"] button {
                        min-width: 23px !important;
                        width: 23px !important;
                        height: 20px !important;
                    }
                    /* v40 final: 선택 열 전체를 엑셀 틀 고정처럼 고정하고,
                       실제 가로 스크롤은 기준년도 열부터 시작합니다. */
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] > div [data-testid="stCheckbox"]) > [data-testid="column"]:first-child {
                        position: sticky !important;
                        left: 0 !important;
                        z-index: 120 !important;
                        background: #ffffff !important;
                        align-self: stretch !important;
                    }
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] > div [data-testid="stCheckbox"]) > [data-testid="column"]:first-child [data-testid="stCheckbox"],
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] > div [data-testid="stCheckbox"]) > [data-testid="column"]:first-child [data-testid="stCheckbox"] > label {
                        position: relative !important;
                        left: auto !important;
                        width: 100% !important;
                        min-width: 100% !important;
                        max-width: 100% !important;
                        height: 40px !important;
                        min-height: 40px !important;
                        max-height: 40px !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: center !important;
                        box-sizing: border-box !important;
                    }
                    /* 추가행의 기준년도~공유/상속인 수 영역: 테두리 색까지 완전 제거 */
                    div.st-key-result_table [class*="st-key-child_first_"] .land-fold-merged,
                    div.st-key-result_table [class*="st-key-child_no_border_"] .land-fold-merged,
                    div.st-key-result_table [class*="st-key-child_first_"] input,
                    div.st-key-result_table [class*="st-key-child_no_border_"] input,
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="stTextInputRootElement"],
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="stTextInputRootElement"],
                    div.st-key-result_table [class*="st-key-child_first_"] [data-baseweb="select"] > div,
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-baseweb="select"] > div {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-color: transparent !important;
                        box-shadow: none !important;
                        outline: none !important;
                    }
                    /* v42: ┗ 버튼 높이 15px */
                    div.st-key-result_table [class*="st-key-fold_cell_"] button {
                        height: 18px !important;
                        min-height: 18px !important;
                        max-height: 18px !important;
                        line-height: 16px !important;
                        padding-top: 0 !important;
                        padding-bottom: 0 !important;
                    }
                    /* v42: 선택 열의 실제 셀/체크박스까지 같은 sticky 좌표계로 고정 */
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] [class*="st-key-base_select_cell_"]) > [data-testid="column"]:first-child,
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] [class*="st-key-owner_select_cell_"]) > [data-testid="column"]:first-child {
                        position: sticky !important;
                        left: 0 !important;
                        z-index: 130 !important;
                        background: #ffffff !important;
                        align-self: stretch !important;
                    }
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] [class*="st-key-base_select_cell_"]) > [data-testid="column"]:first-child [class*="st-key-base_select_cell_"],
                    div.st-key-result_table .stHorizontalBlock:has(> [data-testid="column"] [class*="st-key-owner_select_cell_"]) > [data-testid="column"]:first-child [class*="st-key-owner_select_cell_"] {
                        position: relative !important;
                        left: auto !important;
                        z-index: 131 !important;
                    }

                    /* v42: 공유/상속 추가행의 기준년도~공유 또는 상속인 수까지
                       모든 셀 모양을 완전히 숨깁니다.
                       (선택 열과 실제 ┗/소유구분/토지소유자/지분 입력 영역은 제외) */
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19),
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-style: none !important;
                        border-color: transparent !important;
                        outline: 0 !important;
                        box-shadow: none !important;
                        background: transparent !important;
                    }
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) *,
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) * {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-style: none !important;
                        border-color: transparent !important;
                        outline: 0 !important;
                        box-shadow: none !important;
                    }
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+18) > div,
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+18) > div,
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+18) [data-testid="stVerticalBlock"],
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+18) [data-testid="stVerticalBlock"],
                    div.st-key-result_table [class*="st-key-child_first_"] [class*="st-key-land-fold-merged"],
                    div.st-key-result_table [class*="st-key-child_no_border_"] [class*="st-key-land-fold-merged"] {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-color: transparent !important;
                        outline: 0 !important;
                        box-shadow: none !important;
                        background: transparent !important;
                    }

                    /* v42: 기본행의 공유 또는 상속인 수 셀은 하단 테두리까지
                       확실하게 보이도록 바깥 컨테이너와 NumberInput 양쪽에 적용합니다. */
                    div.st-key-result_table [class*="st-key-count_controls_"] {
                        border: 1px solid #c8c8c8 !important;
                        border-radius: 6px !important;
                        box-sizing: border-box !important;
                        background: #ffffff !important;
                        min-height: 40px !important;
                        height: 40px !important;
                        border-bottom: 1px solid #c8c8c8 !important;
                    }
                    div.st-key-result_table [data-testid="stNumberInput"] {
                        border: 1px solid #c8c8c8 !important;
                        border-radius: 6px !important;
                        box-shadow: none !important;
                        border-bottom: 1px solid #c8c8c8 !important;
                    }
                    div.st-key-result_table [data-testid="stNumberInput"] input {
                        border: 0 !important;
                        border-radius: 0 !important;
                        box-shadow: none !important;
                    }

                    /* v42: 기본행의 소유사항/소유구분 값 입력 영역은 셀 외곽 테두리를
                       보이지 않게 하여 하나의 연결된 입력 영역처럼 표시합니다. */
                    div.st-key-result_table .stHorizontalBlock > [data-testid="column"]:nth-child(18) .land-value-cell,
                    div.st-key-result_table .stHorizontalBlock > [data-testid="column"]:nth-child(20) [data-testid="stSelectbox"],
                    div.st-key-result_table .stHorizontalBlock > [data-testid="column"]:nth-child(20) [data-baseweb="select"] > div {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-color: transparent !important;
                        box-shadow: none !important;
                    }

                    /* v42: 소유사항/소유구분 드롭다운은 바깥 셀 하나만 테두리를 갖게 합니다. */
                    div.st-key-result_table [data-testid="stSelectbox"] {
                        border: 1px solid #c8c8c8 !important;
                        border-radius: 6px !important;
                        box-sizing: border-box !important;
                        overflow: hidden !important;
                        background: #ffffff !important;
                    }
                    div.st-key-result_table [data-testid="stSelectbox"] [data-baseweb="select"],
                    div.st-key-result_table [data-testid="stSelectbox"] [data-baseweb="select"] > div {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-radius: 0 !important;
                        box-shadow: none !important;
                        outline: 0 !important;
                    }
                    div.st-key-result_table [data-testid="stSelectbox"] [data-baseweb="select"] > div {
                        min-height: 38px !important;
                        height: 38px !important;
                        background: #ffffff !important;
                    }
                    /* v42: 소유구분 드롭다운의 외곽 테두리도 숨깁니다. */
                    div.st-key-result_table .stHorizontalBlock > [data-testid="column"]:nth-child(20) [data-testid="stSelectbox"] {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-color: transparent !important;
                        box-shadow: none !important;
                        outline: 0 !important;
                    }

                    /* v43: 공유/상속 추가행의 기준년도~공유/상속인 수 영역은
                       '셀 모양' 자체가 보이지 않도록 내부 요소까지 숨깁니다.
                       레이아웃 공간은 유지하되 테두리/배경/그림자는 모두 제거합니다.
                       ┗ 버튼만 예외적으로 다시 표시합니다. */
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19),
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-style: none !important;
                        border-color: transparent !important;
                        background: transparent !important;
                        box-shadow: none !important;
                        outline: 0 !important;
                        border-radius: 0 !important;
                    }
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) > div,
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) > div,
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) [data-testid="stVerticalBlock"],
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) [data-testid="stVerticalBlock"] {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-style: none !important;
                        border-color: transparent !important;
                        background: transparent !important;
                        box-shadow: none !important;
                        outline: 0 !important;
                        border-radius: 0 !important;
                    }
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) * ,
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) * {
                        border-color: transparent !important;
                        box-shadow: none !important;
                        outline-color: transparent !important;
                    }
                    /* 실제로 표시되는 콘텐츠가 없는 영역은 완전히 투명하게 유지합니다. */
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) .land-fold-merged,
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) .land-fold-merged {
                        visibility: hidden !important;
                        border: 0 !important;
                        background: transparent !important;
                    }
                    /* 공유/상속인 수 열의 ┗ 버튼만 보이게 복원 */
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(19) [class*="st-key-fold_cell_"],
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(19) [class*="st-key-fold_cell_"] *,
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(19) [class*="st-key-fold_cell_"],
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(19) [class*="st-key-fold_cell_"] * {
                        visibility: visible !important;
                    }

                    /* v43: ┗ 버튼은 18px 높이이며 해당 40px 행 안에서 위/아래 중앙 정렬 */
                    div.st-key-result_table [class*="st-key-fold_cell_"] {
                        height: 40px !important;
                        min-height: 40px !important;
                        max-height: 40px !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: flex-end !important;
                        padding: 0 1px 0 0 !important;
                        margin: 0 !important;
                        box-sizing: border-box !important;
                        border: 0 !important;
                        background: transparent !important;
                    }
                    div.st-key-result_table [class*="st-key-fold_cell_"] button {
                        width: 23px !important;
                        min-width: 23px !important;
                        max-width: 23px !important;
                        height: 18px !important;
                        min-height: 18px !important;
                        max-height: 18px !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        line-height: 16px !important;
                        transform: none !important;
                        box-sizing: border-box !important;
                    }

                    /* v43: 소유구분은 '셀 + 드롭다운' 이중 테두리가 생기지 않도록
                       바깥 셀의 테두리는 없애고, 실제 드롭다운 하나만 테두리를 표시합니다. */
                    div.st-key-result_table .stHorizontalBlock > [data-testid="column"]:nth-child(20) {
                        border: 0 !important;
                        box-shadow: none !important;
                        background: transparent !important;
                    }
                    div.st-key-result_table .stHorizontalBlock > [data-testid="column"]:nth-child(20) [data-testid="stSelectbox"] {
                        border: 1px solid #c8c8c8 !important;
                        border-radius: 6px !important;
                        box-shadow: none !important;
                        background: #ffffff !important;
                        overflow: hidden !important;
                    }
                    div.st-key-result_table .stHorizontalBlock > [data-testid="column"]:nth-child(20) [data-testid="stSelectbox"] [data-baseweb="select"],
                    div.st-key-result_table .stHorizontalBlock > [data-testid="column"]:nth-child(20) [data-testid="stSelectbox"] [data-baseweb="select"] > div {
                        border: 0 !important;
                        box-shadow: none !important;
                        outline: 0 !important;
                        border-radius: 0 !important;
                    }

                    /* v44: 추가행의 빈 셀은 실제 셀처럼 보이지 않도록 column 자체까지 완전히 투명화합니다.
                       기존 전역 .land-fold-merged 테두리 규칙보다 더 강하게 적용합니다. */
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19),
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19),
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) > div,
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) > div,
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) [data-testid="stVerticalBlock"],
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) [data-testid="stVerticalBlock"] {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-style: none !important;
                        border-color: transparent !important;
                        border-radius: 0 !important;
                        box-shadow: none !important;
                        outline: 0 !important;
                        background: transparent !important;
                    }
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) .land-fold-merged,
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) .land-fold-merged {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-style: none !important;
                        border-color: transparent !important;
                        border-radius: 0 !important;
                        box-shadow: none !important;
                        outline: 0 !important;
                        background: transparent !important;
                        visibility: hidden !important;
                    }

                    /* v45: 소유사항/소유구분은 드롭다운 전체에 테두리 하나만 표시합니다.
                       Streamlit의 stSelectbox wrapper에는 테두리를 만들지 않고,
                       실제 BaseWeb select 하나에만 테두리를 표시합니다. */
                    div.st-key-result_table [data-testid="stSelectbox"],
                    div.st-key-result_table [data-testid="stSelectbox"] > div,
                    div.st-key-result_table [data-testid="stSelectbox"] [data-testid="stSelectboxLabel"] {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-style: none !important;
                        border-color: transparent !important;
                        box-shadow: none !important;
                        outline: 0 !important;
                        background: transparent !important;
                    }
                    div.st-key-result_table [data-testid="stSelectbox"] [data-baseweb="select"] {
                        border: 1px solid #c8c8c8 !important;
                        border-radius: 6px !important;
                        box-shadow: none !important;
                        outline: 0 !important;
                        background: #ffffff !important;
                        overflow: hidden !important;
                        box-sizing: border-box !important;
                    }
                    div.st-key-result_table [data-testid="stSelectbox"] [data-baseweb="select"] > div,
                    div.st-key-result_table [data-testid="stSelectbox"] [data-baseweb="select"] > div > div,
                    div.st-key-result_table [data-testid="stSelectbox"] [data-baseweb="select"] input {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-style: none !important;
                        border-color: transparent !important;
                        border-radius: 0 !important;
                        box-shadow: none !important;
                        outline: 0 !important;
                        background: transparent !important;
                    }
                    div.st-key-result_table [data-testid="column"]:has([data-testid="stSelectbox"]) {
                        border: 0 !important;
                        box-shadow: none !important;
                        outline: 0 !important;
                        background: transparent !important;
                    }

                    /* v44: ┗ 주변의 부모 셀/컨테이너 테두리는 제거하고 버튼만 표시합니다. */
                    div.st-key-result_table [class*="st-key-fold_cell_"] {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-style: none !important;
                        border-color: transparent !important;
                        box-shadow: none !important;
                        outline: 0 !important;
                        background: transparent !important;
                        height: 40px !important;
                        min-height: 40px !important;
                        max-height: 40px !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: flex-end !important;
                        margin: 0 !important;
                        padding: 0 1px 0 0 !important;
                    }
                    div.st-key-result_table [class*="st-key-fold_cell_"] button {
                        height: 18px !important;
                        min-height: 18px !important;
                        max-height: 18px !important;
                        width: 23px !important;
                        min-width: 23px !important;
                        max-width: 23px !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        line-height: 16px !important;
                        box-sizing: border-box !important;
                    }
                    /* ┗가 들어 있는 공유/상속인 수 열의 부모 column도 무테두리입니다. */
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(19),
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(19) {
                        border: 0 !important;
                        box-shadow: none !important;
                        outline: 0 !important;
                        background: transparent !important;
                    }

                    /* v45 최종: 추가행의 입력 불가 영역에는 HTML placeholder 자체를 만들지 않습니다.
                       혹시 남아 있는 land-fold-merged가 있더라도 화면에 표시하지 않습니다. */
                    div.st-key-result_table [class*="st-key-child_first_"] .land-fold-merged,
                    div.st-key-result_table [class*="st-key-child_no_border_"] .land-fold-merged {
                        display: none !important;
                        visibility: hidden !important;
                        width: 0 !important;
                        height: 0 !important;
                        min-width: 0 !important;
                        min-height: 0 !important;
                        border: 0 !important;
                        padding: 0 !important;
                        margin: 0 !important;
                        background: transparent !important;
                        box-shadow: none !important;
                    }
                    /* 추가행의 기준년도~공유/상속인 수 실제 column에도 테두리를 만들지 않습니다. */
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19),
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19),
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) > div,
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+19) > div {
                        border: 0 !important;
                        border-width: 0 !important;
                        border-style: none !important;
                        border-color: transparent !important;
                        background: transparent !important;
                        box-shadow: none !important;
                        outline: 0 !important;
                    }
                    /* v46: 추가 소유자 행의 입력 불가 영역은 실제 column 자체를 보이지 않게 합니다.
                       빈 column에 남아 있던 Streamlit/상위 래퍼의 시각적 박스까지 제거합니다.
                       선택 열은 유지하고, 기준년도~소유사항 영역은 완전 비표시합니다. */
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+18),
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(n+2):nth-child(-n+18) {
                        visibility: hidden !important;
                        border: 0 !important;
                        background: transparent !important;
                        box-shadow: none !important;
                        outline: none !important;
                    }
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(19) {
                        visibility: hidden !important;
                        border: 0 !important;
                        background: transparent !important;
                        box-shadow: none !important;
                        outline: none !important;
                    }
                    /* 첫 번째 추가행의 19번째 열은 ┗ 버튼만 보입니다. */
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(19) {
                        visibility: visible !important;
                        border: 0 !important;
                        background: transparent !important;
                        box-shadow: none !important;
                        outline: none !important;
                    }
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(19) > div,
                    div.st-key-result_table [class*="st-key-child_first_"] [data-testid="column"]:nth-child(19) [data-testid="stVerticalBlock"],
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(19) > div,
                    div.st-key-result_table [class*="st-key-child_no_border_"] [data-testid="column"]:nth-child(19) [data-testid="stVerticalBlock"] {
                        border: 0 !important;
                        background: transparent !important;
                        box-shadow: none !important;
                        outline: none !important;
                    }

                    /* v46: 소유구분 드롭다운은 BaseWeb 실제 select 하나에만 테두리를 표시합니다.
                       stSelectbox 바깥 wrapper의 테두리는 제거하여 이중 박스를 없앱니다. */
                    div.st-key-result_table [data-testid="stSelectbox"] {
                        border: 0 !important;
                        box-shadow: none !important;
                        outline: none !important;
                        background: transparent !important;
                        border-radius: 0 !important;
                        overflow: visible !important;
                    }
                    div.st-key-result_table [data-testid="stSelectbox"] [data-baseweb="select"] {
                        border: 1px solid #c8c8c8 !important;
                        border-radius: 6px !important;
                        box-shadow: none !important;
                        outline: none !important;
                        background: #ffffff !important;
                        overflow: hidden !important;
                        box-sizing: border-box !important;
                    }
                    div.st-key-result_table [data-testid="stSelectbox"] [data-baseweb="select"] > div,
                    div.st-key-result_table [data-testid="stSelectbox"] [data-baseweb="select"] > div > div {
                        border: 0 !important;
                        border-radius: 0 !important;
                        box-shadow: none !important;
                        outline: none !important;
                        background: transparent !important;
                    }

                    /* v46: ┗ 버튼은 18px 높이로 유지하고 행 중앙에 배치합니다. */
                    div.st-key-result_table [class*="st-key-fold_cell_"] {
                        height: 40px !important;
                        min-height: 40px !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: flex-end !important;
                        border: 0 !important;
                        background: transparent !important;
                        box-shadow: none !important;
                    }
                    div.st-key-result_table [class*="st-key-fold_cell_"] button {
                        height: 18px !important;
                        min-height: 18px !important;
                        max-height: 18px !important;
                        width: 23px !important;
                        min-width: 23px !important;
                        max-width: 23px !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        box-sizing: border-box !important;
                    }

                    /* v54: 선택 열 전체는 고정 패널, 기준년도부터 오른쪽은 단일 스크롤 패널. */
                    div.st-key-result_table [class*="st-key-result_fixed"] {
                        overflow: visible !important;
                        position: relative !important;
                        z-index: 100 !important;
                        background: #fff !important;
                    }
                    div.st-key-result_table [class*="st-key-result_scroll"] {
                        overflow-x: auto !important;
                        overflow-y: visible !important;
                        min-width: 0 !important;
                        width: 100% !important;
                        max-width: 100% !important;
                        background: #fff !important;
                    }
                    div.st-key-result_table [class*="st-key-result_scroll"] > div {
                        min-width: 1900px !important;
                        width: 1900px !important;
                        max-width: none !important;
                    }
                    .land-fixed-header-cell {
                        box-sizing: border-box !important;
                        width: 42px !important;
                        min-width: 42px !important;
                        height: 68px !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: center !important;
                        border: 1px solid #b7b7b7 !important;
                        background: #eeeeee !important;
                        color: #202020 !important;
                        font-size: 13px !important;
                        font-weight: 700 !important;
                        border-radius: 0 !important;
                    }
                    .land-data-header-table {
                        width: 1858px !important;
                        min-width: 1858px !important;
                        max-width: none !important;
                    }
                    div.st-key-result_table [class*="st-key-base_select_cell_"],
                    div.st-key-result_table [class*="st-key-owner_select_cell_"] {
                        width: 42px !important;
                        min-width: 42px !important;
                        max-width: 42px !important;
                        height: 40px !important;
                        min-height: 40px !important;
                        display: flex !important;
                        align-items: center !important;
                        justify-content: center !important;
                        margin: 0 !important;
                        padding: 0 !important;
                        background: #fff !important;
                        border: 1px solid #c8c8c8 !important;
                        border-radius: 0 !important;
                        box-sizing: border-box !important;
                    }
                    div.st-key-result_table [class*="st-key-base_select_cell_"] [data-testid="stCheckbox"],
                    div.st-key-result_table [class*="st-key-owner_select_cell_"] [data-testid="stCheckbox"],
                    div.st-key-result_table [class*="st-key-base_select_cell_"] [data-testid="stCheckbox"] > label,
                    div.st-key-result_table [class*="st-key-owner_select_cell_"] [data-testid="stCheckbox"] > label {
                        position: relative !important;
                        left: auto !important;
                        right: auto !important;
                        top: auto !important;
                    }
                    /* v46에 남아 있는 sticky 규칙을 무효화합니다. */
                    div.st-key-result_table .land-header-table th.select-col,
                    div.st-key-result_table .stHorizontalBlock:has([class*="st-key-base_select_cell_"]) > [data-testid="column"]:first-child,
                    div.st-key-result_table .stHorizontalBlock:has([class*="st-key-owner_select_cell_"]) > [data-testid="column"]:first-child,
                    div.st-key-result_table [class*="st-key-base_select_cell_"] [data-testid="stCheckbox"],
                    div.st-key-result_table [class*="st-key-owner_select_cell_"] [data-testid="stCheckbox"] {
                        position: relative !important;
                        left: auto !important;
                        right: auto !important;
                    }
                    </style>
                    """,
                    unsafe_allow_html=True,
                )

                # 결과표를 하나의 컨테이너로 묶습니다.
                table = st.container(key="result_table", border=True)
                owner_changed = False

                with table:
                    # v54: 선택 열과 데이터 영역을 DOM 레벨에서 완전히 분리합니다.
                    # 왼쪽 선택 패널에는 가로 스크롤을 주지 않고, 오른쪽 데이터 패널만
                    # 가로 스크롤을 갖게 하여 엑셀의 틀 고정처럼 선택 열 전체가 움직이지 않게 합니다.
                    fixed_panel, scroll_panel = st.columns([0.45, 21.0], gap=None)

                    with fixed_panel:
                        st.markdown(
                            '<div class="land-fixed-header-cell">선택</div>',
                            unsafe_allow_html=True,
                        )

                    data_colgroup = "".join(
                        f'<col style="width:{w / sum(col_widths[1:]) * 100:.4f}%">'
                        for w in col_widths[1:]
                    )
                    with scroll_panel:
                        st.markdown(
                            f"""
                            <div class="land-scroll-header-wrap">
                              <table class="land-header-table land-data-header-table">
                                <colgroup>{data_colgroup}</colgroup>
                                <tr>
                                  <th rowspan="2">기준년도</th>
                                  <th rowspan="2">사업지구명</th>
                                  <th colspan="5" class="group">종전</th>
                                  <th colspan="5" class="group">확정</th>
                                  <th rowspan="2">증가면적</th>
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
                            </div>
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
                        # 기본행: 22개 열(첫 번째 선택 열 + 기존 21개 데이터 열)
                        # 기준년도 ~ 확정 면적(12개 항목)은 결과표에서
                        # 직접 수정/입력할 수 있도록 편집 위젯으로 표시합니다.
                        # -------------------------
                        # 22개 열을 동일한 폭 비율로 한 행에 배치합니다.
                        # 첫 번째 열은 CSS sticky로 고정되어 엑셀의 틀 고정처럼 동작합니다.
                        with fixed_panel:
                            with st.container(key=f"base_select_cell_{idx}"):
                                checked = st.checkbox(
                                    "",
                                    key=f"excel_row_check_{idx}",
                                    label_visibility="collapsed",
                                )
                                selected_rows = st.session_state.setdefault("excel_selected_rows", set())
                                selection_key = ("base", idx)
                                selected_rows.discard(idx)
                                if checked:
                                    selected_rows.add(selection_key)
                                else:
                                    selected_rows.discard(selection_key)

                        # 기준년도부터 지분까지의 21개 열만 스크롤 영역에 배치합니다.
                        with scroll_panel:
                            cells = [None] + list(st.columns(col_widths[1:], gap=None))

                        editable_text_fields = [
                            (1, "기준년도"),
                            (2, "사업지구명"),
                            (3, "종전_읍면"),
                            (4, "종전_동리"),
                            (5, "종전_지목"),
                            (6, "종전_지번"),
                            (8, "확정_읍면"),
                            (9, "확정_동리"),
                            (10, "확정_지목"),
                            (11, "확정_지번"),
                        ]

                        editable_area_fields = [
                            (7, "종전_면적"),
                            (12, "확정_면적"),
                        ]

                        # 텍스트 항목
                        for pos, column_name in editable_text_fields:
                            widget_key = f"excel_result_edit_{idx}_{column_name}"
                            initial_value = row.get(column_name, "")
                            if pd.isna(initial_value):
                                initial_value = ""
                            initial_value = str(initial_value).strip()
                            if widget_key not in st.session_state:
                                st.session_state[widget_key] = initial_value

                            with cells[pos]:
                                edited_value = st.text_input(column_name, key=widget_key, label_visibility="collapsed")
                            st.session_state["excel_result_df"].loc[idx, column_name] = edited_value
                            if "excel_result_source_df" in st.session_state and column_name in st.session_state["excel_result_source_df"].columns:
                                st.session_state["excel_result_source_df"].loc[idx, column_name] = edited_value

                        # 면적 항목도 직접 입력/수정할 수 있게 합니다.
                        # 화면에는 100.0 대신 100처럼 표시합니다.
                        # 면적 입력값을 DataFrame에 저장하기 위한 숫자 변환 함수입니다.
                        # 화면의 text_input 값은 문자열이므로 float64 컬럼에 그대로 대입하지 않습니다.
                        def parse_area_input(value):
                            try:
                                text = "" if value is None else str(value).strip().replace(",", "")
                                if not text:
                                    return 0.0
                                return round(float(text), 1)
                            except (TypeError, ValueError):
                                return 0.0

                        for pos, column_name in editable_area_fields:
                            widget_key = f"excel_result_edit_{idx}_{column_name}"
                            initial_value = row.get(column_name, "")
                            if pd.isna(initial_value):
                                initial_text = ""
                            else:
                                initial_text = fmt_value(initial_value)
                            if widget_key not in st.session_state:
                                st.session_state[widget_key] = initial_text

                            with cells[pos]:
                                edited_text = st.text_input(
                                    column_name,
                                    key=widget_key,
                                    label_visibility="collapsed",
                                )

                            # text_input은 문자열을 반환하지만 DataFrame의 면적 컬럼은 float64입니다.
                            # 반드시 숫자로 변환하여 저장합니다.
                            edited_area_value = parse_area_input(edited_text)
                            st.session_state["excel_result_df"].loc[idx, column_name] = edited_area_value
                            if "excel_result_source_df" in st.session_state and column_name in st.session_state["excel_result_source_df"].columns:
                                st.session_state["excel_result_source_df"].loc[idx, column_name] = edited_area_value

                        edited_prev_area = parse_area_input(st.session_state.get(f"excel_result_edit_{idx}_종전_면적", row.get("종전_면적", 0)))
                        edited_conf_area = parse_area_input(st.session_state.get(f"excel_result_edit_{idx}_확정_면적", row.get("확정_면적", 0)))
                        area_difference = edited_prev_area - edited_conf_area
                        new_increase = abs(area_difference) if area_difference < 0 else 0.0
                        new_decrease = abs(area_difference) if area_difference > 0 else 0.0
                        st.session_state["excel_result_df"].loc[idx, "증가면적"] = round(new_increase, 1)
                        st.session_state["excel_result_df"].loc[idx, "감소면적"] = round(new_decrease, 1)
                        st.session_state["excel_result_df"].loc[idx, "증감여부"] = (
                            "증가" if new_increase > 0 else "감소" if new_decrease > 0 else "변동없음"
                        )
                        st.session_state["excel_result_df"].loc[idx, "처리상태"] = (
                            "말소" if edited_conf_area == 0 else "유지"
                        )

                        # 계산 결과 항목은 읽기 전용으로 표시합니다.
                        render_value_cell(cells[13], new_increase)
                        render_value_cell(cells[14], new_decrease)
                        render_value_cell(cells[15], st.session_state["excel_result_df"].loc[idx, "증감여부"])
                        render_value_cell(cells[16], st.session_state["excel_result_df"].loc[idx, "처리상태"])

                        # 소유사항: 업로드 값이 있어도 수정 가능
                        with cells[17]:
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
                            st.session_state["excel_result_df"].loc[idx, "소유사항"] = selected_plain
                            if "excel_result_source_df" in st.session_state:
                                st.session_state["excel_result_source_df"].loc[idx, "소유사항"] = selected_plain
                            owner_changed = True

                            if selected_plain == "단독":
                                st.session_state["excel_result_df"]["공유_상속인수"] = \
                                    st.session_state["excel_result_df"]["공유_상속인수"].astype(object)
                                st.session_state["excel_result_df"].loc[idx, "공유_상속인수"] = 0
                                st.session_state["excel_owner_data"][owner_key] = [owners[0] if owners else {"소유구분":"1. 개인","토지소유자":"","지분":""}]
                                count = 0
                            else:
                                count = max(1, count)
                                st.session_state["excel_result_df"]["공유_상속인수"] = \
                                    st.session_state["excel_result_df"]["공유_상속인수"].astype(object)
                                st.session_state["excel_result_df"].loc[idx, "공유_상속인수"] = count
                                current = st.session_state["excel_owner_data"][owner_key]
                                while len(current) < count:
                                    current.append({"소유구분":"1. 개인","토지소유자":"","지분":""})
                                owners = current

                        # 공유 또는 상속인 수
                        with cells[18]:
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
                                count_key = f"excel_count_{idx}_{selected_plain}"
                                current_count = max(1, min(100, int(count)))
                                if count_key not in st.session_state:
                                    st.session_state[count_key] = current_count
                                else:
                                    try:
                                        st.session_state[count_key] = max(1, min(100, int(st.session_state[count_key])))
                                    except (TypeError, ValueError):
                                        st.session_state[count_key] = current_count

                                def _change_excel_count(widget_key, delta):
                                    try:
                                        current_value = int(st.session_state.get(widget_key, current_count))
                                    except (TypeError, ValueError):
                                        current_value = current_count
                                    st.session_state[widget_key] = max(1, min(100, current_value + delta))

                                # 요청하신 순서: 숫자 → + → -
                                with st.container(key=f"count_controls_{idx}_{selected_plain}"):
                                    count_input, btn_plus, btn_minus = st.columns([0.50, 0.25, 0.25], gap="small")
                                    with count_input:
                                        count_value = st.number_input(
                                            "공유 또는 상속인 수",
                                            min_value=1, max_value=100, step=1,
                                            key=count_key, label_visibility="collapsed",
                                        )
                                    with btn_plus:
                                        st.button(
                                            "+", key=f"excel_count_plus_{idx}_{selected_plain}",
                                            help="공유 또는 상속인 수 1명 증가", use_container_width=True,
                                            on_click=_change_excel_count, args=(count_key, 1),
                                        )
                                    with btn_minus:
                                        st.button(
                                            "-", key=f"excel_count_minus_{idx}_{selected_plain}",
                                            help="공유 또는 상속인 수 1명 감소", use_container_width=True,
                                            on_click=_change_excel_count, args=(count_key, -1),
                                        )

                                if int(count_value) != int(count):
                                    count = int(count_value)
                                    # 업로드 당시의 값은 초기값으로만 사용합니다.
                                    # 이후 숫자/+/- 조작은 현재 화면 상태만 변경합니다.
                                    # dtype에 관계없이 자유롭게 수정할 수 있도록 object 컬럼에
                                    # 정수값을 저장합니다.
                                    st.session_state["excel_result_df"]["공유_상속인수"] = \
                                        st.session_state["excel_result_df"]["공유_상속인수"].astype(object)
                                    st.session_state["excel_result_df"].loc[idx, "공유_상속인수"] = count
                                    if "excel_result_source_df" in st.session_state:
                                        st.session_state["excel_result_source_df"]["공유_상속인수"] = \
                                            st.session_state["excel_result_source_df"]["공유_상속인수"].astype(object)
                                        st.session_state["excel_result_source_df"].loc[idx, "공유_상속인수"] = count
                                    current = st.session_state["excel_owner_data"][owner_key]
                                    if len(current) < count:
                                        current.extend({"소유구분":"1. 개인","토지소유자":"","지분":""} for _ in range(count - len(current)))
                                    elif len(current) > count:
                                        st.session_state["excel_owner_data"][owner_key] = current[:count]
                                    owners = st.session_state["excel_owner_data"][owner_key]
                                    owner_changed = True

                        # 단독은 기본행에 소유자 정보를 입력
                        if selected_plain == "단독":
                            owner = owners[0]
                            with cells[19]:
                                owner_type = st.selectbox(
                                    "소유구분", OWNERSHIP_OPTIONS,
                                    index=OWNERSHIP_OPTIONS.index(owner.get("소유구분", "1. 개인"))
                                    if owner.get("소유구분", "1. 개인") in OWNERSHIP_OPTIONS else 0,
                                    key=f"excel_owner_type_{idx}_0", label_visibility="collapsed",
                                )
                                if owner.get("소유구분") != owner_type:
                                    owner_changed = True
                                owner["소유구분"] = owner_type

                            with cells[20]:
                                new_name = st.text_input(
                                    "토지소유자", value=owner.get("토지소유자", ""),
                                    key=f"excel_owner_name_{idx}_0", label_visibility="collapsed",
                                )
                                if owner.get("토지소유자", "") != new_name:
                                    owner_changed = True
                                owner["토지소유자"] = new_name

                            with cells[21]:
                                st.text_input(
                                    "지분", value="1", disabled=True,
                                    key=f"excel_owner_share_{idx}_0", label_visibility="collapsed",
                                )
                                owner["지분"] = "1"

                        elif selected_plain == "상속":
                            # 상속인 경우 기본행의 소유구분/토지소유자도 직접 입력·수정할 수 있도록 합니다.
                            # 기본행은 첫 번째 상속인의 정보를 사용하며, 아래 추가행과 동일한
                            # excel_owner_data를 공유하므로 어느 위치에서 수정해도 값이 유지됩니다.
                            owner = owners[0] if owners else {"소유구분":"1. 개인", "토지소유자":"", "지분":""}
                            if not owners:
                                st.session_state["excel_owner_data"][owner_key] = [owner]
                                owners = st.session_state["excel_owner_data"][owner_key]

                            with cells[19]:
                                owner_type = st.selectbox(
                                    "소유구분", OWNERSHIP_OPTIONS,
                                    index=OWNERSHIP_OPTIONS.index(owner.get("소유구분", "1. 개인"))
                                    if owner.get("소유구분", "1. 개인") in OWNERSHIP_OPTIONS else 0,
                                    key=f"excel_owner_type_{idx}_0", label_visibility="collapsed",
                                )
                                if owner.get("소유구분") != owner_type:
                                    owner_changed = True
                                owner["소유구분"] = owner_type

                            with cells[20]:
                                new_name = st.text_input(
                                    "토지소유자", value=owner.get("토지소유자", ""),
                                    key=f"excel_owner_name_{idx}_0", label_visibility="collapsed",
                                )
                                if owner.get("토지소유자", "") != new_name:
                                    owner_changed = True
                                owner["토지소유자"] = new_name

                            share_class = "share-error" if invalid_share else ""
                            render_value_cell(cells[21], total_share, share_class)
                        else:
                            # 공유는 기존처럼 기본행에는 소유자 상세를 표시하지 않고
                            # 아래 추가 소유자 행에서 입력합니다.
                            render_value_cell(cells[19], "")
                            render_value_cell(cells[20], "")
                            share_class = "share-error" if invalid_share else ""
                            render_value_cell(cells[21], total_share, share_class)

                        # -------------------------------------------------
                        # 공유/상속 추가행
                        # ┗ 버튼은 공유 또는 상속인 수의 숫자/+/- 컨트롤 바로 아래에
                        # 오도록 병합영역의 오른쪽 끝(인원수 열 아래)에 배치합니다.
                        # -------------------------------------------------
                        if selected_plain in ["공유", "상속"]:
                            if owner_key not in st.session_state["excel_owner_collapsed"]:
                                st.session_state["excel_owner_collapsed"][owner_key] = False

                            collapsed = st.session_state["excel_owner_collapsed"][owner_key]
                            render_owner_indices = [0] if collapsed else list(range(count_value))

                            for owner_idx in render_owner_indices:
                                owner = st.session_state["excel_owner_data"][owner_key][owner_idx]

                                # 첫 번째 추가행은 기존 테두리를 유지하고,
                                # 두 번째 이후 추가행은 테두리를 제거합니다.
                                if owner_idx == 0:
                                    row_wrapper = st.container(key=f"child_first_{idx}")
                                else:
                                    row_wrapper = st.container(key=f"child_no_border_{idx}_{owner_idx}")

                                with row_wrapper:
                                    # 기본행과 정확히 같은 22개 열을 사용합니다. 첫 번째 열도
                                    # 실제 행의 한 셀로 존재하며 CSS sticky로만 고정됩니다.
                                    # 기존 v23은 앞의 18개 열을 하나로 합친 뒤 그 안에 ┗ 버튼을
                                    # 넣었기 때문에 Streamlit의 nested column 폭 계산에 따라
                                    # 버튼이 빈 영역의 상단/가운데처럼 보일 수 있었습니다.
                                    # 이제 ┗ 버튼을 실제 '공유 또는 상속인 수' 열(index 17)에
                                    # 직접 배치하여 숫자 / + / - 바로 아래에 고정합니다.
                                    # 추가 소유자 행의 선택 셀도 고정 패널에만 배치합니다.
                                    with fixed_panel:
                                        with st.container(key=f"owner_select_cell_{idx}_{owner_idx}"):
                                            owner_checked = st.checkbox(
                                                "",
                                                key=f"excel_owner_row_check_{idx}_{owner_idx}",
                                                label_visibility="collapsed",
                                            )
                                            selected_rows = st.session_state.setdefault("excel_selected_rows", set())
                                            owner_selection_key = ("owner", idx, owner_idx)
                                            if owner_checked:
                                                selected_rows.add(owner_selection_key)
                                            else:
                                                selected_rows.discard(owner_selection_key)

                                    with scroll_panel:
                                        child = [None] + list(st.columns(col_widths[1:], gap=None))

                                    # 추가 행의 기준년도~공유/상속인 수 영역은
                                    # "빈 셀"을 렌더링하지 않습니다.
                                    # Streamlit column만 존재시키고 내부 요소를 만들지 않아
                                    # 빈 셀의 사각형/둥근 테두리가 화면에 나타나지 않게 합니다.
                                    # 실제 행 높이는 같은 행의 소유구분/토지소유자/지분 입력칸이
                                    # 결정하므로 행 높이도 유지됩니다.

                                    with child[18]:
                                        if owner_idx == 0:
                                            fold_box = st.container(key=f"fold_cell_{idx}")
                                            with fold_box:
                                                if st.button(
                                                    "┗",
                                                    key=f"fold_excel_{idx}",
                                                    help="공유/상속 소유자 행 접기/펼치기",
                                                    use_container_width=True,
                                                ):
                                                    st.session_state["excel_owner_collapsed"][owner_key] = not collapsed
                                                    st.rerun()
                                        else:
                                            # 첫 번째 소유자 행이 아닌 경우에는
                                            # ┗ 버튼도 표시하지 않고 빈 column만 유지합니다.
                                            # 별도의 빈 HTML 셀을 만들지 않아 테두리가 나타나지 않습니다.
                                            pass

                                    # 접힌 상태에서는 소유자 입력 3칸을 숨깁니다.
                                    if collapsed:
                                        continue

                                    with child[19]:
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

                                    with child[20]:
                                        new_name = st.text_input(
                                            "토지소유자",
                                            value=owner.get("토지소유자", ""),
                                            key=f"excel_child_owner_name_{idx}_{owner_idx}",
                                            label_visibility="collapsed",
                                        )
                                        if owner.get("토지소유자", "") != new_name:
                                            owner_changed = True
                                        owner["토지소유자"] = new_name

                                    with child[21]:
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


# 업로드 파일을 삭제한 경우에도 이전 업로드 파일의 숫자 위젯 상태가
# 다음 파일에 남지 않도록 초기화합니다.
if not uploaded_file:
    if st.session_state.get("excel_uploaded_signature") is not None:
        for _key in list(st.session_state.keys()):
            if (
                _key.startswith("excel_count_")
                or _key.startswith("excel_owner_")
                or _key.startswith("excel_fold_")
                or _key.startswith("excel_child_")
                or _key.startswith("excel_result_edit_")
                or _key.startswith("excel_ownership_type_")
            ):
                del st.session_state[_key]

        for _key in [
            "excel_result_df",
            "excel_result_source_df",
            "excel_owner_data",
            "excel_owner_collapsed",
            "share_dialog_signature",
            "pending_share_dialog",
        ]:
            st.session_state.pop(_key, None)

        st.session_state.pop("excel_uploaded_signature", None)


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