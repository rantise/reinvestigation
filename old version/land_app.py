import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO


# =========================================================
# 페이지 설정
# =========================================================

st.set_page_config(
    page_title="토지 면적 증감 및 소유 분석 프로그램",
    layout="wide"
)

st.title("🏡 토지 면적 증감 및 소유 분석 프로그램")

st.markdown("""
이 프로그램은 엑셀 파일을 업로드하거나 직접 입력하여
**종전/확정 면적을 비교**하고,
**소유사항 및 공유·상속 정보를 확인**합니다.
""")


# =========================================================
# 공통 함수
# =========================================================

def make_id(df, prefix):
    """
    종전 또는 확정 데이터를 이용하여 비교용 ID 생성
    """

    return (
        df[f'{prefix}_읍면'].fillna('').astype(str).str.strip()
        + "_"
        + df[f'{prefix}_동리'].fillna('').astype(str).str.strip()
        + "_"
        + df[f'{prefix}_지목'].fillna('').astype(str).str.strip()
        + "_"
        + df[f'{prefix}_지번'].fillna('').astype(str).str.strip()
    )


def calculate_results(df):
    """
    종전/확정 면적을 비교하여 결과 생성
    """

    df = df.copy()

    # -----------------------------------------------------
    # 필요한 컬럼
    # -----------------------------------------------------

    required_columns = [
        '종전_읍면',
        '종전_동리',
        '종전_지목',
        '종전_지번',
        '종전_면적',
        '확정_읍면',
        '확정_동리',
        '확정_지목',
        '확정_지번',
        '확정_면적'
    ]

    missing = [
        col for col in required_columns
        if col not in df.columns
    ]

    if missing:
        raise ValueError(
            "필수 컬럼이 없습니다.\n\n"
            f"누락된 컬럼: {missing}\n\n"
            f"현재 컬럼: {df.columns.tolist()}"
        )

    # -----------------------------------------------------
    # 면적 숫자 변환
    # -----------------------------------------------------

    df['종전_면적'] = pd.to_numeric(
        df['종전_면적'],
        errors='coerce'
    ).fillna(0)

    df['확정_면적'] = pd.to_numeric(
        df['확정_면적'],
        errors='coerce'
    ).fillna(0)

    # -----------------------------------------------------
    # 종전 / 확정 ID 생성
    # -----------------------------------------------------

    df['종전_ID'] = make_id(df, '종전')
    df['확정_ID'] = make_id(df, '확정')

    # -----------------------------------------------------
    # 종전 면적 집계
    # -----------------------------------------------------

    prev_area_map = (
        df.groupby('종전_ID')['종전_면적']
        .sum()
        .to_dict()
    )

    # -----------------------------------------------------
    # 확정 ID와 종전 ID가 같은 경우 매칭
    # -----------------------------------------------------

    df['종전_총면적'] = (
        df['확정_ID']
        .map(prev_area_map)
        .fillna(0)
    )

    # -----------------------------------------------------
    # 면적 증감
    # -----------------------------------------------------

    df['증가면적'] = np.where(
        df['확정_면적'] > df['종전_총면적'],
        df['확정_면적'] - df['종전_총면적'],
        0
    )

    df['감소면적'] = np.where(
        df['종전_총면적'] > df['확정_면적'],
        df['종전_총면적'] - df['확정_면적'],
        0
    )

    # -----------------------------------------------------
    # 증감 여부
    # -----------------------------------------------------

    df['증감여부'] = np.where(
        df['증가면적'] > 0,
        '증가',
        np.where(
            df['감소면적'] > 0,
            '감소',
            '변동없음'
        )
    )

    # -----------------------------------------------------
    # 처리 상태
    # -----------------------------------------------------

    df['처리상태'] = np.where(
        df['확정_면적'] == 0,
        '말소',
        '유지'
    )

    return df


# =========================================================
# 엑셀 헤더 변환 함수
# =========================================================

def convert_excel_columns(df):
    """
    2단 헤더 형태의 엑셀 컬럼을
    프로그램에서 사용할 컬럼명으로 변환

    예:
    ('종전', '읍면') -> 종전_읍면
    ('확정', '면적') -> 확정_면적
    """

    new_columns = []

    for col in df.columns:

        first = str(col[0]).strip()
        second = str(col[1]).strip()

        # 종전
        if first == '종전':
            new_columns.append(
                f'종전_{second}'
            )

        # 확정
        elif first == '확정':
            new_columns.append(
                f'확정_{second}'
            )

        # 기본 정보
        elif first == '기준년도':
            new_columns.append('기준년도')

        elif first == '사업지구명':
            new_columns.append('사업지구명')

        # 소유 정보
        elif first == '소유사항':
            new_columns.append('소유사항')

        elif first == '공유 또는 상속인 수':
            new_columns.append('공유_상속인수')

        else:
            new_columns.append(first)

    df.columns = new_columns

    return df


# =========================================================
# 엑셀 다운로드 함수
# =========================================================

def create_excel_download(df):
    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine='openpyxl'
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name='계산결과'
        )

    return output.getvalue()


# =========================================================
# UI
# =========================================================

tab1, tab2 = st.tabs(
    [
        "📥 엑셀 파일 업로드",
        "📝 직접 입력 모드"
    ]
)


# =========================================================
# TAB 1 : 엑셀 업로드
# =========================================================

with tab1:

    uploaded_file = st.file_uploader(
        "엑셀 파일을 선택하세요.",
        type=["xlsx"]
    )

    if uploaded_file:

        try:

            # ---------------------------------------------
            # 2단 헤더로 엑셀 읽기
            # ---------------------------------------------

            df_input = pd.read_excel(
                uploaded_file,
                header=[0, 1]
            )

            # ---------------------------------------------
            # 컬럼명 변환
            # ---------------------------------------------

            df_input = convert_excel_columns(
                df_input
            )

            # ---------------------------------------------
            # 인식된 컬럼 확인
            # ---------------------------------------------

            st.subheader("📋 인식된 엑셀 컬럼")

            st.write(
                df_input.columns.tolist()
            )

            # ---------------------------------------------
            # 데이터 미리보기
            # ---------------------------------------------

            st.subheader("📄 업로드 데이터")

            st.dataframe(
                df_input,
                use_container_width=True
            )

            # ---------------------------------------------
            # 계산 버튼
            # ---------------------------------------------

            if st.button(
                "엑셀 데이터로 계산 시작",
                key="excel_calculate"
            ):

                try:

                    result_df = calculate_results(
                        df_input
                    )

                    st.success(
                        "✅ 계산이 완료되었습니다."
                    )

                    # -------------------------------------
                    # 결과 표시
                    # -------------------------------------

                    st.subheader("📊 계산 결과")

                    st.dataframe(
                        result_df,
                        use_container_width=True
                    )

                    # -------------------------------------
                    # 다운로드
                    # -------------------------------------

                    excel_data = create_excel_download(
                        result_df
                    )

                    st.download_button(
                        label="📥 결과 엑셀 다운로드",
                        data=excel_data,
                        file_name="result_calc.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        key="excel_download"
                    )

                except Exception as e:

                    st.error(
                        f"계산 중 오류가 발생했습니다.\n\n{e}"
                    )

        except Exception as e:

            st.error(
                f"엑셀 파일을 읽는 중 오류가 발생했습니다.\n\n{e}"
            )


# =========================================================
# TAB 2 : 직접 입력
# =========================================================

with tab2:

    st.subheader("📍 토지 기본 정보")

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

    st.subheader("👥 소유 정보")

    ownership_type = st.selectbox(
        "소유 사항",
        [
            "단독",
            "공유",
            "상속"
        ]
    )

    is_multi = ownership_type in [
        "공유",
        "상속"
    ]

    num_persons = st.number_input(
        "공유 또는 상속인 수",
        min_value=1,
        max_value=10,
        value=1,
        step=1,
        disabled=not is_multi
    )

    # -----------------------------------------------------
    # 소유자 데이터
    # -----------------------------------------------------

    owner_data = []

    for i in range(int(num_persons)):

        st.divider()

        c1, c2, c3 = st.columns(
            [3, 3, 2]
        )

        # 공유 첫 번째 줄은 빈칸
        if (
            i == 0
            and ownership_type == "공유"
        ):

            val1 = c1.text_input(
                "소유 구분",
                value="",
                disabled=True,
                key=f"owner_type_{i}"
            )

            val2 = c2.text_input(
                "토지소유자",
                value="",
                disabled=True,
                key=f"owner_name_{i}"
            )

            val3 = c3.text_input(
                "지분",
                value="",
                disabled=True,
                key=f"owner_share_{i}"
            )

        else:

            val1 = c1.selectbox(
                "소유 구분",
                [
                    "개인",
                    "법인",
                    "국유지",
                    "시유지",
                    "시군구유지"
                ],
                key=f"owner_type_{i}"
            )

            val2 = c2.text_input(
                "토지소유자",
                key=f"owner_name_{i}"
            )

            val3 = c3.text_input(
                "지분",
                key=f"owner_share_{i}"
            )

            try:

                share_float = (
                    float(val3)
                    if val3.strip()
                    else 0.0
                )

            except ValueError:

                share_float = 0.0

                st.warning(
                    f"{i + 1}번째 지분은 숫자로 입력해주세요."
                )

            owner_data.append(
                {
                    "소유구분": val1,
                    "소유자": val2,
                    "지분": share_float
                }
            )

    # -----------------------------------------------------
    # 지분 검증
    # -----------------------------------------------------

    if ownership_type in [
        "공유",
        "상속"
    ]:

        total_share = sum(
            item["지분"]
            for item in owner_data
        )

        st.write(
            f"현재 지분 합계: **{total_share:.6f}**"
        )

        if np.isclose(
            total_share,
            1.0,
            atol=1e-9
        ):

            st.success(
                "✅ 지분 합계가 1입니다."
            )

        elif total_share > 1.0:

            st.error(
                "⚠️ 지분 합계가 1을 초과합니다."
            )

        else:

            st.error(
                "⚠️ 지분 합계가 1보다 작습니다."
            )

    # -----------------------------------------------------
    # 직접 입력 계산
    # -----------------------------------------------------

    if st.button(
        "입력 내용으로 계산하기",
        key="manual_calculate"
    ):

        # ---------------------------------------------
        # 직접 입력 데이터를 실제 엑셀 구조와 동일하게 생성
        # ---------------------------------------------

        new_data = {

            "기준년도": [""],

            "사업지구명": [""],

            "종전_읍면": [district],

            "종전_동리": [district],

            "종전_지목": [land_type],

            "종전_지번": [lot_num],

            "종전_면적": [prev_area],

            "확정_읍면": [district],

            "확정_동리": [district],

            "확정_지목": [land_type],

            "확정_지번": [lot_num],

            "확정_면적": [conf_area],

            "소유사항": [ownership_type],

            "공유_상속인수": [
                int(num_persons)
                if is_multi
                else 1
            ]
        }

        df_temp = pd.DataFrame(
            new_data
        )

        # ---------------------------------------------
        # 계산
        # ---------------------------------------------

        try:

            df_final = calculate_results(
                df_temp
            )

            # -----------------------------------------
            # 소유자 정보 추가
            # -----------------------------------------

            if owner_data:

                owner_df = pd.DataFrame(
                    owner_data
                )

                st.subheader(
                    "👥 소유자 정보"
                )

                st.dataframe(
                    owner_df,
                    use_container_width=True
                )

            # -----------------------------------------
            # 계산 결과
            # -----------------------------------------

            st.success(
                "✅ 계산이 완료되었습니다."
            )

            st.subheader(
                "📊 계산 결과"
            )

            st.dataframe(
                df_final,
                use_container_width=True
            )

            # -----------------------------------------
            # 결과 다운로드
            # -----------------------------------------

            excel_data = create_excel_download(
                df_final
            )

            st.download_button(
                label="📥 결과 엑셀 다운로드",
                data=excel_data,
                file_name="result_calc.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                key="manual_download"
            )

        except Exception as e:

            st.error(
                f"계산 중 오류가 발생했습니다.\n\n{e}"
            )
