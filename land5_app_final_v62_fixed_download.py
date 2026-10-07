import streamlit as st
import pandas as pd
import numpy as np
import hashlib
import json
import re
from datetime import datetime
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
        df["토지변경"] = np.where(
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

def _normalize_header_text(value):
    """헤더 비교용 정규화: 줄바꿈/탭/공백 차이를 무시합니다."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return "".join(str(value).replace("\r", "").replace("\n", "").replace("\t", "").split())


def _normalize_appraisal_header(first, second):
    """감정평가 업체명/번호의 표현이 달라도 업체 A/B로 인식합니다."""
    f = _normalize_header_text(first)
    sec = _normalize_header_text(second)
    # 예: 감정평가 업체1 / 감정평가업체 1 / 감정평가 1 / 감정평가1 / 업체1
    company = None
    if ("감정평가" in f or "업체" in f) and ("1" in f or "１" in f):
        company = "A"
    elif ("감정평가" in f or "업체" in f) and ("2" in f or "２" in f):
        company = "B"
    if not company:
        return None
    if any(x in sec for x in ("㎡당가격", "m²당가격", "m2당가격", "제곱미터당가격")):
        return f"㎡당가격(원) ({company})"
    if "감정평가액" in sec:
        return f"감정평가액(원) ({company})"
    return None


def _prepare_two_row_headers(df):
    """2행 헤더의 병합 셀/Unnamed 표기를 정리하되 원본 데이터는 변경하지 않습니다."""
    if not isinstance(df.columns, pd.MultiIndex) or df.columns.nlevels < 2:
        return df.copy()
    cols = []
    last_first = ""
    for col in df.columns:
        first = "" if col[0] is None else str(col[0]).strip()
        second = "" if len(col) < 2 or col[1] is None else str(col[1]).strip()
        if first.lower().startswith("unnamed:"):
            first = ""
        if second.lower().startswith("unnamed:"):
            second = ""
        if first:
            last_first = first
        elif last_first:
            first = last_first
        cols.append((first, second))
    out = df.copy()
    out.columns = pd.MultiIndex.from_tuples(cols)
    return out


def convert_excel_columns(df):
    """엑셀 2단 헤더를 내부 컬럼으로 변환하고 중복명을 안전하게 고유화합니다."""
    new_columns = []
    used = {}
    for col in df.columns:
        if isinstance(col, tuple):
            first = "" if col[0] is None else str(col[0]).strip()
            second = "" if len(col) < 2 or col[1] is None else str(col[1]).strip()
        else:
            first, second = str(col).strip(), ""
        appraisal = _normalize_appraisal_header(first, second)
        if appraisal:
            name = appraisal
        else:
            nf = _normalize_header_text(first)
            if nf == _normalize_header_text("종전"):
                name = "종전_" + second.replace("\r", "").replace("\n", "").strip()
            elif nf == _normalize_header_text("확정"):
                name = "확정_" + second.replace("\r", "").replace("\n", "").strip()
            elif nf == _normalize_header_text("기준년도"):
                name = "기준년도"
            elif nf == _normalize_header_text("사업지구명"):
                name = "사업지구명"
            elif nf == _normalize_header_text("소유사항"):
                name = "소유사항"
            elif nf == _normalize_header_text("공유 또는 상속인 수"):
                name = "공유_상속인수"
            else:
                # 알 수 있는 2행 헤더도 첫 행/둘째 행의 정보를 모두 보존합니다.
                # 어느 한쪽만 값이 있는 경우에는 그 값을, 둘 다 있으면 두 값을
                # 결합하여 내부 컬럼명으로 사용합니다. 이후 헤더 비교 시에는
                # 각 구성요소를 다시 분리하여 결과 헤더와 비교할 수 있습니다.
                first_clean = _normalize_header_text(first)
                second_clean = _normalize_header_text(second)
                if first_clean and second_clean:
                    name = f"{first_clean}__{second_clean}"
                else:
                    name = first_clean or second_clean or first or second
        used[name] = used.get(name, 0) + 1
        new_columns.append(name if used[name] == 1 else f"{name}__dup{used[name]}")
    result = df.copy()
    result.columns = new_columns
    return result


def _copy_matching_upload_values(result_df, source_df):
    """업로드 헤더와 결과 헤더가 같으면 업로드 값을 결과에 복사합니다.

    2행 헤더에서는 1행/2행 중 어느 한쪽에 결과 헤더가 있거나,
    두 행을 합친 이름이 결과 헤더와 같으면 매칭합니다. 줄바꿈/공백은 무시합니다.
    중복 헤더는 물리적인 열 위치를 기준으로 처리합니다.
    """
    result = result_df.copy().astype(object)
    source = source_df.copy()
    if result.empty or source.empty:
        return result

    def norm(v):
        return _normalize_header_text(v).replace("__dup2", "").replace("__dup3", "")

    def source_header_candidates(col):
        vals = []
        if isinstance(col, tuple):
            parts = ["" if x is None else str(x) for x in col]
            cleaned = [norm(x) for x in parts if norm(x)]
            vals.extend(cleaned)
            if len(cleaned) >= 2:
                vals.append(norm("__".join(cleaned)))
        else:
            text = str(col)
            vals.append(norm(text))
            if "__" in text:
                vals.extend([norm(x) for x in text.split("__") if norm(x)])
        return list(dict.fromkeys(x for x in vals if x))

    source_map = {}
    for pos, col in enumerate(source.columns):
        for key in source_header_candidates(col):
            source_map.setdefault(key, []).append(pos)

    for rpos, rcol in enumerate(result.columns):
        key = norm(rcol)
        positions = source_map.get(key, [])
        if not positions:
            continue
        for row_pos in range(min(len(result), len(source))):
            chosen = None
            for spos in positions:
                value = source.iloc[row_pos, spos]
                if value is not None and not pd.isna(value) and str(value).strip() != "":
                    chosen = value
                    break
            if chosen is not None:
                result.iloc[row_pos, rpos] = chosen
    return result


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


def _safe_filename_part(value):
    text = "" if value is None or pd.isna(value) else str(value).strip()
    text = re.sub(r'[\\/:*?"<>|]+', "_", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or "미입력"


def _make_result_save_basename(result_df):
    year = ""
    district = ""
    if isinstance(result_df, pd.DataFrame) and not result_df.empty:
        first_row = result_df.iloc[0]
        for col in result_df.columns:
            n = _normalize_header_text(col)
            if n == _normalize_header_text("기준년도"):
                year = first_row.get(col, "")
            elif n == _normalize_header_text("사업지구명"):
                district = first_row.get(col, "")
    return f"{_safe_filename_part(year)}_{_safe_filename_part(district)}_({datetime.now().strftime('%Y%m%d')})"


def _json_default(value):
    if value is None or value is pd.NA:
        return None
    if isinstance(value, np.integer):
        return int(value)
    if isinstance(value, np.floating):
        return None if not np.isfinite(value) else float(value)
    if isinstance(value, np.bool_):
        return bool(value)
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.isoformat()
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    return str(value)


def create_result_json(result_df, owner_data=None, header_meta=None):
    clean_df = result_df.copy().where(pd.notna(result_df), None)
    records = [{str(k): _json_default(v) for k, v in row.items()} for row in clean_df.to_dict(orient="records")]
    owners_clean = {}
    for key, value in (owner_data or {}).items():
        if isinstance(value, list):
            owners_clean[str(key)] = [
                {str(k): _json_default(v) for k, v in owner.items()}
                for owner in value if isinstance(owner, dict)
            ]
    payload = {
        "format": "land5_result_save",
        "version": 1,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "result_columns": [str(c) for c in result_df.columns],
        "result_records": records,
        "owner_data": owners_clean,
        "header_meta": header_meta if isinstance(header_meta, list) else [],
    }
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")


def load_result_json(uploaded_file):
    uploaded_file.seek(0)
    payload = json.loads(uploaded_file.getvalue().decode("utf-8-sig"))
    if payload.get("format") != "land5_result_save":
        raise ValueError("이 프로그램에서 저장한 결과 JSON 파일이 아닙니다.")
    df = pd.DataFrame(payload.get("result_records", []))
    columns = payload.get("result_columns", [])
    if columns:
        df = df.reindex(columns=columns)
    return payload, df.astype(object)


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


def create_adjustment_error_excel(errors):
    """조정금/감정평가 오류를 오류별 개별 Sheet로 만들고 오류값은 빨간 글자로 표시합니다."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for n, error in enumerate(errors, start=1):
            row = {
                "종전 지번": error.get("before_land_address", ""),
                "확정 지번": error.get("after_land_address", ""),
                "문제 내용": error.get("issue", ""),
                "증감여부": error.get("change_type", ""),
                "증가 또는 감소 항목": error.get("area_item", ""),
                "증가 또는 감소 값": error.get("area_value", ""),
                "증가면적": error.get("increase_area", ""),
                "감소면적": error.get("decrease_area", ""),
                "㎡당가격 (A+B) / 2": error.get("avg_unit_price", ""),
                "면적 × ㎡당가격 (A+B) / 2 계산값": error.get("formula1_value", ""),
                "2개의 감정평가액의 평균 값": error.get("formula2_value", ""),
                "조정금 - 납부": error.get("pay_value", ""),
                "조정금 - 지급": error.get("give_value", ""),
                "감정평가 업체1 ㎡당가격(원) (A)": error.get("a_unit_price", ""),
                "감정평가 업체2 ㎡당가격(원) (B)": error.get("b_unit_price", ""),
                "A·B 단가 평균 계산값": error.get("ab_avg_calculated", ""),
                "감정평가액(원) (A)": error.get("a_amount", ""),
                "감정평가액(원) (B)": error.get("b_amount", ""),
            }
            df = pd.DataFrame([row])
            df.to_excel(writer, index=False, sheet_name=f"오류{n}")
            ws = writer.book[f"오류{n}"]
            # 헤더 가독성
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
            # 오류와 직접 관련된 값은 빨간 글자로 표시
            red_cols = set(error.get("red_export_columns", []))
            for col_idx, col_name in enumerate(df.columns, start=1):
                if col_name in red_cols or col_name == "문제 내용":
                    ws.cell(row=2, column=col_idx).font = ws.cell(row=2, column=col_idx).font.copy(color="FF0000")
            for column_cells in ws.columns:
                max_len = max(len(str(c.value or "")) for c in column_cells)
                ws.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 12), 60)
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
        type=["xlsx", "json"]
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
            is_saved_json = str(getattr(uploaded_file, "name", "")).lower().endswith(".json")
            json_restore_payload = None
            json_restore_df = None
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
            if is_saved_json:
                json_restore_payload, json_restore_df = load_result_json(uploaded_file)
                df_raw = json_restore_df.copy()
                converted_df = json_restore_df.copy()
                preview_header_mode = "저장 JSON"
                st.session_state["excel_upload_header_meta"] = json_restore_payload.get("header_meta", [])
            else:
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

            if not is_saved_json:
                try:
                    converted_source_df = _prepare_two_row_headers(df_raw.copy())
                    converted_df = convert_excel_columns(converted_source_df)

                    # 업로드 Excel의 물리적인 헤더 위치/표기를 그대로 기억합니다.
                    # 이후 결과표 헤더를 만들 때 이 정보를 사용하므로,
                    # Excel에 새로운 열이 추가되어도 결과표에서 같은 순서/위치로 표시됩니다.
                    upload_header_meta = []
                    for _hpos, _hcol in enumerate(converted_source_df.columns):
                        if isinstance(_hcol, tuple):
                            _hfirst = "" if _hcol[0] is None else str(_hcol[0]).strip()
                            _hsecond = "" if len(_hcol) < 2 or _hcol[1] is None else str(_hcol[1]).strip()
                        else:
                            _hfirst = str(_hcol).strip()
                            _hsecond = ""
                        if _hfirst.lower().startswith("unnamed:"):
                            _hfirst = ""
                        if _hsecond.lower().startswith("unnamed:"):
                            _hsecond = ""
                        upload_header_meta.append({
                            "position": _hpos,
                            "internal": str(converted_df.columns[_hpos]),
                            "first": _hfirst,
                            "second": _hsecond,
                        })
                    st.session_state["excel_upload_header_meta"] = upload_header_meta
                except Exception:
                    # 알 수 없는 헤더 구조라도 업로드 미리보기는 계속 표시합니다.
                    converted_df = df_raw.copy()
                    st.session_state["excel_upload_header_meta"] = []

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
                "저장 JSON 결과 불러오기" if is_saved_json else "엑셀 데이터로 계산 시작",
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

                    # 결과창의 감정평가 업체1/2 하위 4개 항목을 항상 별도 컬럼으로 확보합니다.
                    # 업로드 파일에 해당 값이 있으면 아래 헤더 매칭 단계에서 자동으로 채워집니다.
                    appraisal_result_columns = []
                    for _app_col in appraisal_result_columns:
                        if _app_col not in result_df.columns:
                            result_df[_app_col] = ""

                    # 결과표에 표시할 컬럼은 업로드 파일에 실제 존재했던 컬럼만 사용합니다.
                    # 계산 과정에서 내부적으로 생성되는 컬럼은 결과 DataFrame에 남겨두되
                    # 화면에는 절대로 추가하지 않습니다. 따라서 업로드 파일에 없는
                    # 처리상태/토지변경/감정평가 고정 컬럼 등이 임의로 나타나지 않습니다.
                    _upload_column_order = [c for c in calculation_input.columns if c in result_df.columns and not str(c).startswith('__')]
                    result_df = result_df.reindex(columns=[c for c in _upload_column_order if c in result_df.columns])

                    # 중요: 저장 JSON의 값은 문자열/숫자/목록 등 원래 저장된 형태를 그대로
                    # 복원할 수 있으므로, JSON 값을 덮어쓰기 전에 반드시 object dtype으로
                    # 변환합니다. 기존 코드처럼 float64 컬럼에 리스트를 대입하면
                    # "Invalid value [...] for dtype 'float64'" 오류가 발생할 수 있습니다.
                    result_df = result_df.astype(object)
                    calculated_df = calculated_df.astype(object)

                    if is_saved_json and isinstance(json_restore_df, pd.DataFrame):
                        _saved_df = json_restore_df.reindex(columns=result_df.columns).copy().astype(object)
                        if len(_saved_df) == len(result_df):
                            _saved_df.index = result_df.index
                            for _col in result_df.columns:
                                # 열 전체를 list로 한 번에 대입하지 않고 object 컬럼에
                                # 개별 값을 넣어 복원하여 pandas dtype 충돌을 방지합니다.
                                for _row_idx in result_df.index:
                                    result_df.at[_row_idx, _col] = _saved_df.at[_row_idx, _col]

                    st.session_state["excel_result_df"] = result_df
                    st.session_state["excel_result_source_df"] = calculated_df
                    if is_saved_json:
                        saved_owner_data = json_restore_payload.get("owner_data", {}) if isinstance(json_restore_payload, dict) else {}
                        uploaded_owner_seeds = [
                            list(saved_owner_data.get(f"excel_owner_{_ridx}", []))
                            for _ridx in result_df.index
                        ]
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
                    prefixes = ("excel_result_edit_", "excel_result_appraisal_", "excel_ownership_type_", "excel_count_", "excel_owner_type_", "excel_owner_name_", "excel_owner_share_", "excel_child_owner_type_", "excel_child_owner_name_", "excel_child_owner_share_", "excel_row_check_", "excel_owner_row_check_")
                    for key in list(st.session_state.keys()):
                        if key.startswith(prefixes):
                            del st.session_state[key]

                def _make_empty_result_row():
                    return {"기준년도":"", "사업지구명":"", "종전_읍면":"", "종전_동리":"", "종전_지목":"", "종전_지번":"", "종전_면적":0.0, "확정_읍면":"", "확정_동리":"", "확정_지목":"", "확정_지번":"", "확정_면적":0.0, "증가면적":0.0, "감소면적":0.0, "증감여부":"변동없음", "토지변경":"유지", "소유사항":"단독", "공유_상속인수":0, "소유구분":"1. 개인", "토지소유자":"", "지분":"1", "㎡당가격(원) (A)":"", "감정평가액(원) (A)":"", "㎡당가격(원) (B)":"", "감정평가액(원) (B)":""}

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
                if "adjustment_dialog_signature" not in st.session_state:
                    st.session_state["adjustment_dialog_signature"] = None
                if "pending_adjustment_dialog" not in st.session_state:
                    st.session_state["pending_adjustment_dialog"] = None
                if "adjustment_error_appraisal_cells" not in st.session_state:
                    st.session_state["adjustment_error_appraisal_cells"] = {}

                if hasattr(st, "dialog"):
                    @st.dialog("지분 합계 확인")
                    def share_error_dialog(message):
                        st.error(message)
                        st.write("입력한 지분을 확인해 주세요.")

                    @st.dialog("조정금 오류 확인", width="large")
                    def adjustment_error_dialog(errors):
                        if not errors:
                            return

                        current = int(st.session_state.get("adjustment_error_page", 0))
                        current = max(0, min(current, len(errors) - 1))
                        payload = errors[current]

                        st.markdown(
                            f'<div style="color:#ff0000;font-size:24px;font-weight:800;text-align:center;">'
                            f'({len(errors)}개) 오류가 있습니다.</div>',
                            unsafe_allow_html=True,
                        )

                        nav_left, page_col, nav_right = st.columns([1, 8, 1])
                        with nav_left:
                            if st.button("<", key="adjust_err_prev", use_container_width=True):
                                st.session_state["adjustment_error_page"] = (current - 1) % len(errors)
                                st.rerun()
                        with page_col:
                            st.markdown(
                                f'<div style="text-align:center;font-weight:700;padding-top:8px;">'
                                f'오류 {current + 1} / {len(errors)}</div>',
                                unsafe_allow_html=True,
                            )
                        with nav_right:
                            if st.button(">", key="adjust_err_next", use_container_width=True):
                                st.session_state["adjustment_error_page"] = (current + 1) % len(errors)
                                st.rerun()

                        st.markdown(
                            f'**종전 지번 - {escape(str(payload.get("before_land_address", "")))}**,  '
                            f'**확정 지번 - {escape(str(payload.get("after_land_address", "")))}**'
                        )
                        st.error(escape(str(payload.get("issue", "해당 행의 조정금에 문제가 있습니다."))))
                        st.markdown(f'**증감여부:** {escape(str(payload.get("change_type", "")))}')
                        st.markdown(f'**증가 또는 감소 항목 - {escape(str(payload.get("area_item", "")))}: {escape(str(payload.get("area_value", "")))}㎡**')
                        st.markdown(f'**㎡당가격 (A+B) / 2: {escape(str(payload.get("avg_unit_price", "")))}원/㎡**')
                        st.markdown(f'**증가면적 또는 감소면적 × ㎡당 감정평가 금액의 합 (A+B) / 2:** {escape(str(payload.get("formula1", "")))} = {escape(str(payload.get("formula1_value", "")))}원')
                        st.markdown(f'**2개의 감정평가액의 평균 값:** {escape(str(payload.get("formula2", "")))} = {escape(str(payload.get("formula2_value", "")))}원')
                        st.markdown(f'**감정평가 업체1 ㎡당가격(원) (A):** {escape(str(payload.get("a_unit_price", "")))}')
                        st.markdown(f'**감정평가 업체2 ㎡당가격(원) (B):** {escape(str(payload.get("b_unit_price", "")))}')
                        st.markdown(f'**(A+B) / 2 계산값:** {escape(str(payload.get("ab_avg_calculated", "")))}원/㎡')
                        st.markdown(f'**조정금 - 납부:** {escape(str(payload.get("pay_value", "")))}')
                        st.markdown(f'**조정금 - 지급:** {escape(str(payload.get("give_value", "")))}')

                        error_excel = create_adjustment_error_excel(errors)
                        st.download_button(
                            "📥 오류 내역 Excel 다운로드",
                            data=error_excel,
                            file_name=f"조정금_오류내역_{datetime.now().strftime("%Y%m%d")}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_adjustment_errors",
                            use_container_width=True,
                        )


                # 열 폭은 헤더와 데이터 행에서 완전히 동일하게 사용합니다.
                # 전체 표 폭은 v5보다 줄이고, 지번은 1324-567 정도가 들어가도록
                # 확보하며, 지분 열은 기존(v5)보다 2배 넓게 잡습니다.
                # 첫 번째 열은 행 선택 전용 열입니다.
                # 기존 21개 데이터 열은 그대로 유지하고 앞에 1개 열을 추가합니다.
                # 기본 26개 열 + 업로드 파일에서 자동으로 발견한 추가 열
                # 결과표에 실제로 표시되는 순서. 선택 열은 화면에서만 앞에 추가합니다.
                # 화면에는 업로드 파일에 실제 존재하는 열만 표시합니다.
                # display_columns의 순서 = Excel의 실제 열 순서입니다.
                display_columns = list(result_df.columns)
                result_col_pos = {name: pos + 1 for pos, name in enumerate(display_columns)}

                # 업로드 파일에서 들어온 모든 추가 열을 자동으로 편집 대상으로 잡습니다.
                # 값이 하나도 없는 열이라도 result_df에는 열 자체가 존재하므로,
                # 해당 열 전체가 단순한 빈 공간으로 남지 않고 각 행마다 입력칸이 표시됩니다.
                # 앞으로 Excel에 새로운 항목을 추가해도 여기의 컬럼명을 수정할 필요가 없습니다.
                _fixed_result_columns = {
                    "기준년도", "사업지구명",
                    "종전_읍면", "종전_동리", "종전_지목", "종전_지번", "종전_면적",
                    "확정_읍면", "확정_동리", "확정_지목", "확정_지번", "확정_면적",
                    "증가면적", "감소면적", "증감여부", "토지변경",
                    "소유사항", "공유_상속인수", "소유구분", "토지소유자", "지분",
                    "㎡당가격(원) (A)", "감정평가액(원) (A)",
                    "㎡당가격(원) (B)", "감정평가액(원) (B)",
                    "조정금 - 납부", "조정금-납부", "조정금__납부", "조정금 - 지급", "조정금-지급", "조정금__지급",
                }
                def _is_special_appraisal_column(_c):
                    n = _normalize_header_text(_c)
                    return n in {
                        _normalize_header_text("㎡당가격(원) (A)"),
                        _normalize_header_text("감정평가액(원) (A)"),
                        _normalize_header_text("㎡당가격(원) (B)"),
                        _normalize_header_text("감정평가액(원) (B)"),
                        _normalize_header_text("㎡당가격 (A+B) / 2"),
                        _normalize_header_text("㎡당가격(원) (A+B) / 2"),
                    }
                extra_result_columns = [
                    c for c in display_columns
                    if c not in _fixed_result_columns and not _is_special_appraisal_column(c)
                ]
                _width_map = {
                    "기준년도":0.80,"사업지구명":1.25,"종전_읍면":0.70,"종전_동리":1.15,"종전_지목":0.80,"종전_지번":1.40,"종전_면적":0.80,
                    "확정_읍면":0.70,"확정_동리":1.15,"확정_지목":0.80,"확정_지번":1.40,"확정_면적":0.80,
                    "증가면적":0.90,"감소면적":0.90,"증감여부":0.85,"토지변경":0.85,"소유사항":0.95,"공유_상속인수":1.15,"소유구분":1.10,"토지소유자":1.20,"지분":1.80,
                    "㎡당가격(원) (A)":1.15,"감정평가액(원) (A)":1.35,"㎡당가격(원) (B)":1.15,"감정평가액(원) (B)":1.35,
                }
                col_widths = [0.45] + [_width_map.get(c, 1.25) for c in display_columns]
                total_col_width = sum(col_widths)
                result_table_width = max(2360, int(2360 * total_col_width / 26.95))
                # v39: 선택 열도 실제 표의 첫 번째 열로 포함합니다.
                # 헤더와 모든 데이터 행이 동일한 22개 열 폭을 사용합니다.
                header_colgroup = "".join(
                    f'<col style="width:{w / total_col_width * 100:.4f}%">'
                    for w in col_widths
                )

                def _display_result_header(column_name):
                    """업로드 헤더의 실제 1행/2행 표기를 우선 사용합니다.

                    내부 컬럼명은 계산을 위해 정규화되어 있지만 화면에서는
                    업로드 Excel의 헤더를 그대로 보여줘야 하므로, 업로드 당시
                    물리적 열 위치에 저장해 둔 원본 헤더를 찾아 사용합니다.
                    """
                    target = str(column_name)
                    target_base = target.replace("__dup2", "").replace("__dup3", "")
                    for _meta in st.session_state.get("excel_upload_header_meta", []):
                        if str(_meta.get("internal", "")) in (target, target_base):
                            _first = str(_meta.get("first", "") or "").strip()
                            _second = str(_meta.get("second", "") or "").strip()
                            if _first and _second:
                                return f"{escape(_first)}<br>{escape(_second)}"
                            if _first:
                                return escape(_first)
                            if _second:
                                return escape(_second)
                    text = target_base
                    if "__" in text:
                        a, b = text.split("__", 1)
                        return f"{escape(a)}<br>{escape(b)}"
                    return escape(text)

                def _build_dynamic_result_header_html(columns):
                    """업로드 Excel의 1행/2행 헤더 구조를 그대로 결과표 헤더에 재현합니다."""
                    meta_by_internal = {str(m.get("internal")): m for m in st.session_state.get("excel_upload_header_meta", [])}
                    cells = []
                    for c in columns:
                        m = meta_by_internal.get(str(c), {})
                        first = str(m.get("first", "") or "").strip()
                        second = str(m.get("second", "") or "").strip()
                        cells.append((first, second))

                    row1 = ['<tr>', '<th rowspan="2" class="select-col">선택</th>']
                    row2 = ['<tr>']
                    i = 0
                    while i < len(columns):
                        first, second = cells[i]
                        if first:
                            j = i + 1
                            while j < len(columns) and cells[j][0] == first and cells[j][0] != "":
                                j += 1
                            span = j - i
                            # 같은 1행 헤더가 여러 열을 묶으면 colspan으로 표시합니다.
                            if span > 1 or second:
                                if span > 1:
                                    row1.append(f'<th colspan="{span}" class="group">{escape(first)}</th>')
                                    for k in range(i, j):
                                        sub = cells[k][1]
                                        if sub:
                                            row2.append(f'<th>{escape(sub)}</th>')
                                        else:
                                            row2.append('<th></th>')
                                else:
                                    if second:
                                        row1.append(f'<th>{escape(first)}</th>')
                                        row2.append(f'<th>{escape(second)}</th>')
                                    else:
                                        row1.append(f'<th rowspan="2">{escape(first)}</th>')
                                i = j
                                continue
                            row1.append(f'<th rowspan="2">{escape(first)}</th>')
                        else:
                            label = second or ""
                            row1.append(f'<th rowspan="2">{escape(label)}</th>')
                        i += 1
                    row1.append('</tr>')
                    row2.append('</tr>')
                    return ''.join(row1 + row2)

                static_keys = [
                    "기준년도", "사업지구명",
                    "종전_읍면", "종전_동리", "종전_지목", "종전_지번", "종전_면적",
                    "확정_읍면", "확정_동리", "확정_지목", "확정_지번", "확정_면적",
                    "증가면적", "감소면적", "증감여부", "토지변경"
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
                result_table_css = """
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
                        width: {result_table_width}px;
                        min-width: {result_table_width}px;
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
                    /* 실제 {result_table_width}px 폭은 스크롤 영역 내부의 첫 콘텐츠 블록에만 적용합니다. */
                    div.st-key-result_table [class*="st-key-result_scroll"] > div > div {
                        min-width: {result_table_width}px !important;
                        width: {result_table_width}px !important;
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
                        width: {result_table_width}px !important;
                        min-width: {result_table_width}px !important;
                        max-width: none !important;
                    }
                    /* 헤더의 선택 셀도 동일한 첫 번째 열에 고정 */
                    div.st-key-result_table .land-header-table {
                        width: {result_table_width}px !important;
                        min-width: {result_table_width}px !important;
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
                    /* 조정금 검증 오류: 해당 납부/지급 입력 셀 전체에 빨간 테두리를 표시합니다. */
                    div.st-key-result_table [class*="st-key-adjustment_error_"] [data-testid="stTextInput"] {
                        border: 2px solid #ff0000 !important;
                        border-radius: 6px !important;
                        box-sizing: border-box !important;
                    }
                    div.st-key-result_table [class*="st-key-adjustment_error_"] [data-baseweb="input"] {
                        border: 2px solid #ff0000 !important;
                        box-shadow: none !important;
                    }
                    div.st-key-result_table [class*="st-key-adjustment_error_"] [data-baseweb="input"] > div {
                        border: 0 !important;
                    }
                    div.st-key-result_table [class*="st-key-appraisal_error_"] [data-testid="stTextInput"] {
                        border: 2px solid #ff0000 !important;
                        border-radius: 6px !important;
                        box-sizing: border-box !important;
                    }
                    div.st-key-result_table [class*="st-key-appraisal_error_"] [data-baseweb="input"] {
                        border: 2px solid #ff0000 !important;
                        box-shadow: none !important;
                    }
                    div.st-key-result_table [class*="st-key-appraisal_error_"] [data-baseweb="input"] > div {
                        border: 0 !important;
                    }
                    /* 오류 팝업 좌우 페이지 이동 버튼 */
                    div.st-key-adjust_err_prev button,
                    div.st-key-adjust_err_next button {
                        background: #ff0000 !important;
                        color: #ffffff !important;
                        border: 1px solid #ff0000 !important;
                        font-weight: 800 !important;
                        font-size: 22px !important;
                    }
                    </style>
                    """
                st.markdown(result_table_css.replace("{result_table_width}", str(result_table_width)), unsafe_allow_html=True)

                def _find_result_column(*aliases):
                    normalized = {_normalize_header_text(a) for a in aliases}
                    for _c in display_columns:
                        if _normalize_header_text(_c) in normalized:
                            return _c
                    return None

                def _to_number(value):
                    if value is None or (isinstance(value, float) and pd.isna(value)):
                        return 0.0
                    text = str(value).strip().replace(',', '')
                    if not text:
                        return 0.0
                    try:
                        return float(text)
                    except (ValueError, TypeError):
                        return 0.0

                def _format_calc_number(value):
                    value = float(value)
                    if not np.isfinite(value):
                        return ""
                    if np.isclose(value, round(value), atol=1e-9):
                        return str(int(round(value)))
                    return f"{value:.1f}"

                _app_a_col = _find_result_column("㎡당가격(원) (A)", "㎡당가격(원)(A)")
                _app_a_amount_col = _find_result_column("감정평가액(원) (A)", "감정평가액(원)(A)")
                _app_b_col = _find_result_column("㎡당가격(원) (B)", "㎡당가격(원)(B)")
                _app_b_amount_col = _find_result_column("감정평가액(원) (B)", "감정평가액(원)(B)")
                _app_avg_col = _find_result_column(
                    "㎡당가격 (A+B) / 2",
                    "㎡당가격(A+B)/2",
                    "㎡당가격 (A + B) / 2",
                    "㎡당가격(A + B)/2",
                    "㎡당가격(원) (A+B) / 2",
                    "㎡당가격(원)(A+B)/2",
                    "㎡당가격(원) (A + B) / 2",
                )
                _adjust_pay_col = _find_result_column("조정금 - 납부", "조정금-납부", "조정금__납부", "조정금 납부")
                _adjust_give_col = _find_result_column("조정금 - 지급", "조정금-지급", "조정금__지급", "조정금 지급")

                # 결과표를 하나의 컨테이너로 묶습니다.
                table = st.container(key="result_table", border=True)
                owner_changed = False

                with table:
                    # v39: 선택 열을 별도 패널로 분리하지 않고, 전체 결과표의 첫 번째 열로
                    # 유지합니다. CSS의 position: sticky로 엑셀의 "틀 고정"과 같은 동작을
                    # 구현합니다. 따라서 선택 열과 나머지 열은 같은 행/같은 높이의 셀입니다.
                    scroll = st.container(key="result_scroll")
                    with scroll:
                        st.markdown(
                            f"""
                            <table class="land-header-table">
                              <colgroup>{header_colgroup}</colgroup>
                              { _build_dynamic_result_header_html(display_columns) }
                            </table>
                            """,
                            unsafe_allow_html=True,
                        )

                    # 모든 데이터 행도 같은 result_scroll 안에 놓고, 첫 번째 열만 sticky로
                    # 고정합니다. 별도의 고정 패널을 만들지 않으므로 셀 높이와 행 위치가
                    # 오른쪽 데이터 열과 정확히 동일합니다.

                    # -------------------------------------------------
                    # 결과 데이터 행
                    # -------------------------------------------------
                    # 각 행은 동일 높이의 Streamlit 블록으로 렌더링하되 CSS에서
                    # 블록 간 기본 여백을 제거하여 업로드 미리보기처럼 붙입니다.
                    adjustment_validation_errors = []
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
                        with scroll:
                            cells = list(st.columns(col_widths, gap=None))

                        with cells[0]:
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

                        editable_text_fields = [
                            (result_col_pos["기준년도"], "기준년도"), (result_col_pos["사업지구명"], "사업지구명"),
                            (result_col_pos["종전_읍면"], "종전_읍면"), (result_col_pos["종전_동리"], "종전_동리"),
                            (result_col_pos["종전_지목"], "종전_지목"), (result_col_pos["종전_지번"], "종전_지번"),
                            (result_col_pos["확정_읍면"], "확정_읍면"), (result_col_pos["확정_동리"], "확정_동리"),
                            (result_col_pos["확정_지목"], "확정_지목"), (result_col_pos["확정_지번"], "확정_지번"),
                        ]
                        editable_area_fields = [
                            (result_col_pos["종전_면적"], "종전_면적"),
                            (result_col_pos["확정_면적"], "확정_면적"),
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
                        st.session_state["excel_result_df"].loc[idx, "토지변경"] = (
                            "말소" if edited_conf_area == 0 else "유지"
                        )

                        # 계산 결과 항목은 읽기 전용으로 표시합니다.
                        render_value_cell(cells[result_col_pos["증가면적"]], new_increase)
                        render_value_cell(cells[result_col_pos["감소면적"]], new_decrease)
                        render_value_cell(cells[result_col_pos["증감여부"]], st.session_state["excel_result_df"].loc[idx, "증감여부"])
                        render_value_cell(cells[result_col_pos["토지변경"]], st.session_state["excel_result_df"].loc[idx, "토지변경"])

                        # 소유사항: 업로드 값이 있어도 수정 가능
                        with cells[result_col_pos["소유사항"]]:
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
                        with cells[result_col_pos["공유_상속인수"]]:
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
                            with cells[result_col_pos["소유구분"]]:
                                owner_type = st.selectbox(
                                    "소유구분", OWNERSHIP_OPTIONS,
                                    index=OWNERSHIP_OPTIONS.index(owner.get("소유구분", "1. 개인"))
                                    if owner.get("소유구분", "1. 개인") in OWNERSHIP_OPTIONS else 0,
                                    key=f"excel_owner_type_{idx}_0", label_visibility="collapsed",
                                )
                                if owner.get("소유구분") != owner_type:
                                    owner_changed = True
                                owner["소유구분"] = owner_type

                            with cells[result_col_pos["토지소유자"]]:
                                new_name = st.text_input(
                                    "토지소유자", value=owner.get("토지소유자", ""),
                                    key=f"excel_owner_name_{idx}_0", label_visibility="collapsed",
                                )
                                if owner.get("토지소유자", "") != new_name:
                                    owner_changed = True
                                owner["토지소유자"] = new_name

                            with cells[result_col_pos["지분"]]:
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

                            with cells[result_col_pos["소유구분"]]:
                                owner_type = st.selectbox(
                                    "소유구분", OWNERSHIP_OPTIONS,
                                    index=OWNERSHIP_OPTIONS.index(owner.get("소유구분", "1. 개인"))
                                    if owner.get("소유구분", "1. 개인") in OWNERSHIP_OPTIONS else 0,
                                    key=f"excel_owner_type_{idx}_0", label_visibility="collapsed",
                                )
                                if owner.get("소유구분") != owner_type:
                                    owner_changed = True
                                owner["소유구분"] = owner_type

                            with cells[result_col_pos["토지소유자"]]:
                                new_name = st.text_input(
                                    "토지소유자", value=owner.get("토지소유자", ""),
                                    key=f"excel_owner_name_{idx}_0", label_visibility="collapsed",
                                )
                                if owner.get("토지소유자", "") != new_name:
                                    owner_changed = True
                                owner["토지소유자"] = new_name

                            share_class = "share-error" if invalid_share else ""
                            render_value_cell(cells[result_col_pos["지분"]], total_share, share_class)
                        else:
                            # 공유는 기존처럼 기본행에는 소유자 상세를 표시하지 않고
                            # 아래 추가 소유자 행에서 입력합니다.
                            render_value_cell(cells[result_col_pos["소유구분"]], "")
                            render_value_cell(cells[result_col_pos["토지소유자"]], "")
                            share_class = "share-error" if invalid_share else ""
                            render_value_cell(cells[result_col_pos["지분"]], total_share, share_class)

                        # 감정평가 단가/평가액은 업로드 값 또는 직접 입력값을 사용합니다.
                        # 평가액은 최초/자동 상태에서만 면적×단가로 계산하고, 사용자가 직접
                        # 수정한 값은 유지합니다.
                        _increase_for_app = _to_number(row.get("증가면적", 0))
                        _decrease_for_app = _to_number(row.get("감소면적", 0))
                        _appraisal_area = abs(_increase_for_app) + abs(_decrease_for_app)

                        # 현재 행의 감정평가 관련 오류 셀을 렌더링할 때 사용합니다.
                        # 검증은 아래 입력값을 모두 읽은 뒤 수행되므로, 렌더링 시점에는
                        # 아직 오류 여부가 결정되지 않은 상태입니다. 따라서 먼저 빈 set으로
                        # 초기화하고, 이전 검증에서 저장된 오류 셀은 별도로 불러옵니다.
                        _error_appraisal = set()

                        def _render_editable_appraisal(_col, _label, _auto_value=None, _signature=None, _error=False):
                            if not _col or _col not in result_col_pos:
                                return ""
                            widget_key = f"excel_result_appraisal_{idx}_{_col}"
                            initial_value = row.get(_col, "")
                            if pd.isna(initial_value):
                                initial_value = ""
                            initial_value = str(initial_value).strip()
                            auto_key = f"excel_result_auto_{idx}_{_col}"
                            sig_key = f"excel_result_auto_sig_{idx}_{_col}"
                            previous_auto = st.session_state.get(auto_key, None)
                            previous_sig = st.session_state.get(sig_key, None)

                            if widget_key not in st.session_state:
                                st.session_state[widget_key] = initial_value

                            if _auto_value is not None:
                                calc_text = _format_calc_number(_auto_value)
                                current = str(st.session_state.get(widget_key, "")).strip()
                                if previous_sig is None and calc_text:
                                    # 업로드 파일에 기존 값이 있어도 A/B 단가를 기준으로
                                    # 최초 표시값은 요구된 계산식으로 산정합니다.
                                    # 이후 사용자가 직접 수정한 값은 자동 계산값과 구분하여 유지합니다.
                                    st.session_state[widget_key] = calc_text
                                    st.session_state[auto_key] = calc_text
                                elif previous_sig != _signature and current == str(previous_auto or "").strip():
                                    st.session_state[widget_key] = calc_text
                                    st.session_state[auto_key] = calc_text
                                st.session_state[sig_key] = _signature

                            with cells[result_col_pos[_col]]:
                                _wrapper_key = f"appraisal_error_{idx}_{_col}" if _error else f"appraisal_ok_{idx}_{_col}"
                                with st.container(key=_wrapper_key):
                                    edited_value = st.text_input(_label, key=widget_key, label_visibility="collapsed")
                            st.session_state["excel_result_df"].loc[idx, _col] = edited_value
                            if "excel_result_source_df" in st.session_state and _col in st.session_state["excel_result_source_df"].columns:
                                st.session_state["excel_result_source_df"].loc[idx, _col] = edited_value
                            return edited_value

                        # 직전 검증에서 오류로 판정된 감정평가 셀은 즉시 빨간 테두리로 유지합니다.
                        _stored_appraisal_error_cells = set(st.session_state.get("adjustment_error_appraisal_cells", {}).get(idx, []))

                        # 단가 A/B를 먼저 표시하고 입력값을 최신 상태로 읽습니다.
                        if _app_a_col:
                            _render_editable_appraisal(_app_a_col, "㎡당가격(원) (A)", _error=(_app_a_col in _stored_appraisal_error_cells or _app_a_col in _error_appraisal))
                        if _app_b_col:
                            _render_editable_appraisal(_app_b_col, "㎡당가격(원) (B)", _error=(_app_b_col in _stored_appraisal_error_cells or _app_b_col in _error_appraisal))

                        _a_num = _to_number(st.session_state.get(f"excel_result_appraisal_{idx}_{_app_a_col}", "")) if _app_a_col else 0.0
                        _b_num = _to_number(st.session_state.get(f"excel_result_appraisal_{idx}_{_app_b_col}", "")) if _app_b_col else 0.0
                        _area_sig = f"{_appraisal_area:.10f}"

                        if _app_a_amount_col:
                            _render_editable_appraisal(
                                _app_a_amount_col, "감정평가액(원) (A)",
                                _appraisal_area * _a_num, f"{_area_sig}|A={_a_num:.10f}", _error=(_app_a_amount_col in _stored_appraisal_error_cells or _app_a_amount_col in _error_appraisal)
                            )
                        if _app_b_amount_col:
                            _render_editable_appraisal(
                                _app_b_amount_col, "감정평가액(원) (B)",
                                _appraisal_area * _b_num, f"{_area_sig}|B={_b_num:.10f}", _error=(_app_b_amount_col in _stored_appraisal_error_cells or _app_b_amount_col in _error_appraisal)
                            )

                        # B가 공란/0이면 A, B가 0 이외이면 (A+B)/2를 자동 표시합니다.
                        # 이 셀도 직접 입력/변경할 수 있으며, 사용자가 변경하면 그 값을 유지합니다.
                        if _app_avg_col:
                            _avg_value = _a_num if np.isclose(_b_num, 0.0, atol=1e-12) else (_a_num + _b_num) / 2.0
                            _render_editable_appraisal(
                                _app_avg_col, "㎡당가격 (A+B) / 2", _avg_value,
                                f"A={_a_num:.10f}|B={_b_num:.10f}", _error=(_app_avg_col in _stored_appraisal_error_cells or _app_avg_col in _error_appraisal)
                            )

                        # -------------------------------------------------
                        # 조정금 - 납부 / 조정금 - 지급
                        # -------------------------------------------------
                        def _has_real_value(_value):
                            if _value is None or pd.isna(_value):
                                return False
                            return str(_value).strip() != ""

                        def _render_adjustment_cell(_col, _label, _auto_value=None, _error=False):
                            if not _col or _col not in result_col_pos:
                                return ""
                            widget_key = f"excel_result_adjustment_{idx}_{_col}"
                            initial_value = row.get(_col, "")
                            if pd.isna(initial_value):
                                initial_value = ""
                            initial_value = str(initial_value).strip()
                            if widget_key not in st.session_state:
                                st.session_state[widget_key] = (
                                    _format_calc_number(_auto_value)
                                    if (not initial_value and _auto_value is not None)
                                    else initial_value
                                )
                            _kind = "pay" if _col == _adjust_pay_col else "give"
                            wrapper_key = f"adjustment_error_{idx}_{_kind}" if _error else f"adjustment_ok_{idx}_{_kind}"
                            with cells[result_col_pos[_col]]:
                                with st.container(key=wrapper_key):
                                    edited_value = st.text_input(_label, key=widget_key, label_visibility="collapsed")
                            st.session_state["excel_result_df"].loc[idx, _col] = edited_value
                            if "excel_result_source_df" in st.session_state and _col in st.session_state["excel_result_source_df"].columns:
                                st.session_state["excel_result_source_df"].loc[idx, _col] = edited_value
                            return edited_value

                        _change_text = str(row.get("증감여부", "")).strip() if "증감여부" in row.index else ""
                        _increase_num = _to_number(row.get("증가면적", 0))
                        _decrease_num = _to_number(row.get("감소면적", 0))
                        _avg_num = _to_number(st.session_state.get(f"excel_result_appraisal_{idx}_{_app_avg_col}", "")) if _app_avg_col else 0.0
                        _a_amount_num = _to_number(st.session_state.get(f"excel_result_appraisal_{idx}_{_app_a_amount_col}", "")) if _app_a_amount_col else 0.0
                        _b_amount_num = _to_number(st.session_state.get(f"excel_result_appraisal_{idx}_{_app_b_amount_col}", "")) if _app_b_amount_col else 0.0
                        _is_increase = "증가" in _change_text
                        _is_decrease = "감소" in _change_text
                        _adjust_col = _adjust_pay_col if _is_increase else (_adjust_give_col if _is_decrease else None)
                        _other_adjust_col = _adjust_give_col if _is_increase else (_adjust_pay_col if _is_decrease else None)

                        def _current_adjustment_value(_col):
                            if not _col:
                                return ""
                            existing = st.session_state.get(f"excel_result_adjustment_{idx}_{_col}", None)
                            if existing is None:
                                existing = row.get(_col, "")
                            return "" if existing is None or pd.isna(existing) else str(existing).strip()

                        _pay_value_before = _current_adjustment_value(_adjust_pay_col)
                        _give_value_before = _current_adjustment_value(_adjust_give_col)
                        _pay_has_value = _has_real_value(_pay_value_before)
                        _give_has_value = _has_real_value(_give_value_before)
                        _any_adjustment_column = bool(_adjust_pay_col or _adjust_give_col)
                        _any_adjustment_value = _pay_has_value or _give_has_value

                        # 감정평가 단가 3개와 감정평가액 2개의 값이 있으면 조정금 열의 존재와 관계없이 검증합니다.
                        _a_unit_has = _has_real_value(st.session_state.get(f"excel_result_appraisal_{idx}_{_app_a_col}", "")) if _app_a_col else False
                        _b_unit_has = _has_real_value(st.session_state.get(f"excel_result_appraisal_{idx}_{_app_b_col}", "")) if _app_b_col else False
                        _avg_unit_has = _has_real_value(st.session_state.get(f"excel_result_appraisal_{idx}_{_app_avg_col}", "")) if _app_avg_col else False
                        _a_amount_has = _has_real_value(st.session_state.get(f"excel_result_appraisal_{idx}_{_app_a_amount_col}", "")) if _app_a_amount_col else False
                        _b_amount_has = _has_real_value(st.session_state.get(f"excel_result_appraisal_{idx}_{_app_b_amount_col}", "")) if _app_b_amount_col else False
                        _appraisal_values_present = _a_unit_has or _b_unit_has or _avg_unit_has or _a_amount_has or _b_amount_has
                        _required_appraisal_values_ready = all([_a_unit_has, _b_unit_has, _avg_unit_has, _a_amount_has, _b_amount_has])
                        _formula_area = _increase_num if _is_increase else (_decrease_num if _is_decrease else 0.0)
                        _formula1_value = _formula_area * _avg_num
                        _formula2_value = (_a_amount_num + _b_amount_num) / 2.0
                        _ab_avg_calculated = (_a_num + _b_num) / 2.0
                        _unit_avg_mismatch = (_a_unit_has and _b_unit_has and _avg_unit_has and not np.isclose(_ab_avg_calculated, _avg_num, atol=0.01, rtol=1e-9))
                        _formula_values_match = (_required_appraisal_values_ready and _formula_area > 0 and np.isclose(_formula1_value, _formula2_value, atol=0.01, rtol=1e-9))

                        _validation_requested = _any_adjustment_column or _any_adjustment_value or _appraisal_values_present
                        _adjust_error = False
                        _error_pay = False
                        _error_give = False
                        # _error_appraisal은 위에서 렌더링 전에 초기화했습니다.
                        # 여기서 다시 초기화하면 이후 검증에는 문제가 없지만, 코드 흐름상
                        # 렌더링 시 참조되는 변수가 미정의되는 회귀가 발생할 수 있으므로
                        # 검증 단계에서는 명시적으로 비우지 않습니다.
                        _error_issue_parts = []
                        _adjust_auto_value = None

                        if _validation_requested:
                            if not _is_increase and not _is_decrease:
                                if _any_adjustment_value:
                                    _adjust_error = True
                                    _error_pay = _pay_has_value
                                    _error_give = _give_has_value
                                    _error_issue_parts.append("증감여부가 증가/감소로 구분되지 않은 상태에서 조정금이 입력되어 있습니다.")
                            else:
                                if _is_increase and _give_has_value:
                                    _adjust_error = True
                                    _error_give = True
                                    _error_issue_parts.append("증가 행인데 조정금 - 지급에 금액이 입력되어 있습니다. 증가 행은 조정금 - 납부만 사용해야 합니다.")
                                if _is_decrease and _pay_has_value:
                                    _adjust_error = True
                                    _error_pay = True
                                    _error_issue_parts.append("감소 행인데 조정금 - 납부에 금액이 입력되어 있습니다. 감소 행은 조정금 - 지급만 사용해야 합니다.")
                                if _pay_has_value and _give_has_value:
                                    _adjust_error = True
                                    _error_pay = True
                                    _error_give = True
                                    _error_issue_parts.append("조정금 - 납부와 조정금 - 지급에 동시에 금액이 입력되어 있습니다. 두 항목 중 하나만 입력해야 합니다.")

                            # A/B 단가 평균과 사용자가 입력한 ㎡당가격 (A+B) / 2가 다른 경우
                            if _unit_avg_mismatch:
                                _adjust_error = True
                                _error_appraisal.update({_app_a_col, _app_b_col, _app_avg_col})
                                _error_issue_parts.append("감정평가 업체1·2의 ㎡당가격 평균과 ㎡당가격 (A+B) / 2의 값이 서로 다릅니다.")

                            # 단가 평균과 감정평가액 평균이 다른 경우
                            if _required_appraisal_values_ready and _formula_area > 0 and not _formula_values_match:
                                _adjust_error = True
                                _error_appraisal.update({_app_avg_col, _app_a_amount_col, _app_b_amount_col})
                                _error_issue_parts.append("증가면적 또는 감소면적 × ㎡당가격 (A+B) / 2의 계산값과 2개의 감정평가액 평균 값이 서로 다릅니다.")

                            if _is_increase or _is_decrease:
                                expected_col = _adjust_col
                                expected_has_value = _pay_has_value if _is_increase else _give_has_value
                                expected_num = _to_number(_pay_value_before if _is_increase else _give_value_before)
                                if _required_appraisal_values_ready and _formula_area > 0 and expected_has_value and not np.isclose(expected_num, _formula1_value, atol=0.01, rtol=1e-9):
                                    _adjust_error = True
                                    if expected_col == _adjust_pay_col:
                                        _error_pay = True
                                    else:
                                        _error_give = True
                                    _error_issue_parts.append(f"{('조정금 - 납부' if _is_increase else '조정금 - 지급')} 금액이 산정된 금액과 다릅니다.")
                                if (not expected_has_value) and _formula_values_match and expected_col:
                                    _adjust_auto_value = _formula1_value

                        if _adjust_error:

                            _before_dongri = row.get("종전_동리", row.get("종전-동리", ""))
                            _before_jibun = row.get("종전_지번", row.get("종전-지번", ""))
                            _after_dongri = row.get("확정_동리", row.get("확정-동리", ""))
                            _after_jibun = row.get("확정_지번", row.get("확정-지번", ""))
                            _pay_display = _pay_value_before
                            _give_display = _give_value_before
                            _formula1_text = (
                                f"{('증가면적' if _is_increase else '감소면적')} {_format_calc_number(_formula_area)}㎡ × "
                                f"{_format_calc_number(_avg_num)}원/㎡ = { _format_calc_number(_formula1_value) }원"
                            )
                            _formula2_text = (
                                f"({_format_calc_number(_a_amount_num)}원 + "
                                f"{_format_calc_number(_b_amount_num)}원) ÷ 2 = { _format_calc_number(_formula2_value) }원"
                            )
                            _red_export_columns = []
                            if _error_pay: _red_export_columns.append("조정금 - 납부")
                            if _error_give: _red_export_columns.append("조정금 - 지급")
                            if _app_avg_col in _error_appraisal: _red_export_columns.append("㎡당가격 (A+B) / 2")
                            if _app_a_col in _error_appraisal: _red_export_columns.append("감정평가 업체1 ㎡당가격(원) (A)")
                            if _app_b_col in _error_appraisal: _red_export_columns.append("감정평가 업체2 ㎡당가격(원) (B)")
                            if _app_a_amount_col in _error_appraisal: _red_export_columns.append("감정평가액(원) (A)")
                            if _app_b_amount_col in _error_appraisal: _red_export_columns.append("감정평가액(원) (B)")
                            # 현재 행에서 오류로 판정된 감정평가 셀을 다음 렌더링에도 유지합니다.
                            _current_error_cells = st.session_state.setdefault("adjustment_error_appraisal_cells", {})
                            _current_error_cells[idx] = [c for c in _error_appraisal if c]
                            adjustment_validation_errors.append({
                                "row_index": idx,
                                "before_dongri": _before_dongri,
                                "before_jibun": _before_jibun,
                                "after_dongri": _after_dongri,
                                "after_jibun": _after_jibun,
                                "before_land_address": f"{_before_dongri} {_before_jibun}".strip(),
                                "after_land_address": f"{_after_dongri} {_after_jibun}".strip(),
                                "change_type": _change_text,
                                "issue": " ".join(_error_issue_parts),
                                "formula1": _formula1_text,
                                "formula1_value": _format_calc_number(_formula1_value),
                                "formula2": _formula2_text,
                                "formula2_value": _format_calc_number(_formula2_value),
                                "increase_area": _format_calc_number(_increase_num),
                                "decrease_area": _format_calc_number(_decrease_num),
                                "area_item": "증가면적" if _is_increase else ("감소면적" if _is_decrease else ""),
                                "area_value": _format_calc_number(_formula_area),
                                "avg_unit_price": _format_calc_number(_avg_num),
                                "a_unit_price": _format_calc_number(_a_num),
                                "b_unit_price": _format_calc_number(_b_num),
                                "ab_avg_calculated": _format_calc_number(_ab_avg_calculated),
                                "a_amount": _format_calc_number(_a_amount_num),
                                "b_amount": _format_calc_number(_b_amount_num),
                                "pay_value": _pay_display,
                                "give_value": _give_display,
                                "red_export_columns": _red_export_columns,
                            })

                        if _adjust_col:
                            _payment_initial = _pay_value_before if _is_increase else _give_value_before
                            _payment_has_value_now = _has_real_value(_payment_initial)
                            _auto_for_expected = _adjust_auto_value if not _payment_has_value_now else None
                            _payment_value = _render_adjustment_cell(
                                _adjust_col,
                                "조정금 - 납부" if _adjust_col == _adjust_pay_col else "조정금 - 지급",
                                _auto_for_expected, _adjust_error and (_error_pay if _adjust_col == _adjust_pay_col else _error_give)
                            )
                        else:
                            _payment_value = ""

                        if _other_adjust_col:
                            _other_value = _render_adjustment_cell(
                                _other_adjust_col,
                                "조정금 - 지급" if _other_adjust_col == _adjust_give_col else "조정금 - 납부",
                                None,
                                _adjust_error and (_error_give if _other_adjust_col == _adjust_give_col else _error_pay)
                            )

                        # 현재 위젯에서 최종 입력값을 다시 읽어 두 열의 값도 검증 결과에 반영합니다.
                        _final_pay = _current_adjustment_value(_adjust_pay_col)
                        _final_give = _current_adjustment_value(_adjust_give_col)
                        if _adjust_error and adjustment_validation_errors:
                            adjustment_validation_errors[-1]["pay_value"] = _final_pay
                            adjustment_validation_errors[-1]["give_value"] = _final_give

                        # 업로드 파일에 새 열이 추가되면 해당 열도 자동으로 결과표에
                        # 생성하고, 업로드 값을 초기값으로 표시한 뒤 직접 수정할 수 있게 합니다.
                        # 이 부분 때문에 새 항목마다 Python 코드에 컬럼명을 추가할 필요가 없습니다.
                        for _extra_col in extra_result_columns:
                            _extra_pos = result_col_pos[_extra_col]
                            widget_key = f"excel_result_extra_{idx}_{_extra_pos}"
                            initial_value = row.get(_extra_col, "")
                            if pd.isna(initial_value):
                                initial_value = ""
                            if widget_key not in st.session_state:
                                st.session_state[widget_key] = str(initial_value).strip()
                            with cells[_extra_pos]:
                                edited_extra = st.text_input(
                                    str(_extra_col),
                                    key=widget_key,
                                    label_visibility="collapsed",
                                )
                            st.session_state["excel_result_df"].loc[idx, _extra_col] = edited_extra
                            if "excel_result_source_df" in st.session_state and _extra_col in st.session_state["excel_result_source_df"].columns:
                                st.session_state["excel_result_source_df"].loc[idx, _extra_col] = edited_extra

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
                                    with scroll:
                                        child = list(st.columns(col_widths, gap=None))

                                    # 추가 소유자 행의 선택 셀도 동일한 첫 번째 열에 배치합니다.
                                    with child[0]:
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

                                    # 추가 행의 기준년도~공유/상속인 수 영역은
                                    # "빈 셀"을 렌더링하지 않습니다.
                                    # Streamlit column만 존재시키고 내부 요소를 만들지 않아
                                    # 빈 셀의 사각형/둥근 테두리가 화면에 나타나지 않게 합니다.
                                    # 실제 행 높이는 같은 행의 소유구분/토지소유자/지분 입력칸이
                                    # 결정하므로 행 높이도 유지됩니다.

                                    with child[result_col_pos["공유_상속인수"]]:
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

                                    with child[result_col_pos["소유구분"]]:
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

                                    with child[result_col_pos["토지소유자"]]:
                                        new_name = st.text_input(
                                            "토지소유자",
                                            value=owner.get("토지소유자", ""),
                                            key=f"excel_child_owner_name_{idx}_{owner_idx}",
                                            label_visibility="collapsed",
                                        )
                                        if owner.get("토지소유자", "") != new_name:
                                            owner_changed = True
                                        owner["토지소유자"] = new_name

                                    with child[result_col_pos["지분"]]:
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

                                    # 추가 소유자 행에서도 업로드 파일의 추가 열은
                                    # 각 셀을 직접 입력할 수 있도록 합니다.
                                    # 감정평가 4개 항목은 기본행에서 입력하므로 여기서는
                                    # 업로드 파일에 새로 존재하는 동적 추가 열만 처리합니다.
                                    for _extra_col in extra_result_columns:
                                        _extra_pos = result_col_pos[_extra_col]
                                        # 산정방법은 공유/상속으로 추가된 행에서는 반드시 공란으로
                                        # 유지하고 입력 위젯도 만들지 않습니다.
                                        if _normalize_header_text(_extra_col) == _normalize_header_text("산정방법"):
                                            owner[_extra_col] = ""
                                            continue

                                        _extra_key = f"excel_child_extra_{idx}_{owner_idx}_{_extra_pos}"
                                        _extra_initial = owner.get(_extra_col, "")
                                        if pd.isna(_extra_initial):
                                            _extra_initial = ""
                                        if _extra_key not in st.session_state:
                                            st.session_state[_extra_key] = str(_extra_initial).strip()
                                        with child[_extra_pos]:
                                            _extra_value = st.text_input(
                                                str(_extra_col),
                                                key=_extra_key,
                                                label_visibility="collapsed",
                                            )
                                        owner[_extra_col] = _extra_value

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

                # 이번 검증에서 오류가 없었던 행의 감정평가 오류 테두리는 제거합니다.
                _live_error_rows = {e.get("row_index") for e in adjustment_validation_errors}
                _stored_error_rows = set(st.session_state.get("adjustment_error_appraisal_cells", {}).keys())
                for _old_row in list(_stored_error_rows - _live_error_rows):
                    st.session_state["adjustment_error_appraisal_cells"].pop(_old_row, None)

                # 조정금 오류가 하나라도 있으면 오류를 모아서 한 개의 팝업에서 페이지별로 확인합니다.
                if hasattr(st, "dialog"):
                    if adjustment_validation_errors:
                        if st.session_state.get("adjustment_error_page", 0) >= len(adjustment_validation_errors):
                            st.session_state["adjustment_error_page"] = 0
                        adjustment_error_dialog(adjustment_validation_errors)
                    else:
                        st.session_state["adjustment_error_page"] = 0

                # 기존 지분 오류 팝업은 기존 동작을 유지합니다.
                if st.session_state.get("pending_share_dialog") and hasattr(st, "dialog"):
                    message = st.session_state.pop("pending_share_dialog")
                    share_error_dialog(message)

                st.subheader("💾 저장 및 다운로드")
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
                col_dl, col_save = st.columns(2)
                with col_dl:
                    st.download_button(
                        "💾 다운로드",
                        data=excel_data,
                        file_name=_make_result_save_basename(download_df) + ".xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_excel_result",
                        use_container_width=True,
                    )
                json_data = create_result_json(
                    result_df,
                    st.session_state.get("excel_owner_data", {}),
                    st.session_state.get("excel_upload_header_meta", []),
                )
                with col_save:
                    st.download_button(
                        "💾 저장",
                        data=json_data,
                        file_name=_make_result_save_basename(result_df) + ".json",
                        mime="application/json",
                        key="save_json_result",
                        use_container_width=True,
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
            "adjustment_dialog_signature",
            "pending_adjustment_dialog",
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
                "💾 저장 및 다운로드",
                data=excel_data,
                file_name=_make_result_save_basename(result_df) + ".xlsx",
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                key="download_manual_result"
            )

            json_data = create_result_json(result_df, {}, st.session_state.get("excel_upload_header_meta", []))
            with col_save:
                st.download_button(
                    "💾 저장",
                    data=json_data,
                    file_name=_make_result_save_basename(result_df) + ".json",
                    mime="application/json",
                    key="save_json_manual_result",
                    use_container_width=True,
                )

        except Exception as e:

            st.error(
                f"계산 중 오류 발생: {e}"
            )