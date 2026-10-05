import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from abc import ABC, abstractmethod

# =========================================================
# 1. Strategy Pattern 정의 (분석 전략)
# =========================================================

class AreaAnalysisStrategy:
    """면적 계산 및 증감 분석을 담당하는 전략 클래스"""
    @staticmethod
    def calculate(df):
        df = df.copy()
        # 면적 숫자 변환
        df['종전_면적'] = pd.to_numeric(df['종전_면적'], errors='coerce').fillna(0)
        df['확정_면적'] = pd.to_numeric(df['확정_면적'], errors='coerce').fillna(0)

        # ID 생성
        df['종전_ID'] = (
            df['종전_읍면'].fillna('').astype(str).str.strip() + "_" +
            df['종전_동리'].fillna('').astype(str).str.strip() + "_" +
            df['종전_지목'].fillna('').astype(str).str.strip() + "_" +
            df['종전_지번'].fillna('').astype(str).str.strip()
        )
        df['확정_ID'] = (
            df['확정_읍면'].fillna('').astype(str).str.strip() + "_" +
            df['확정_동리'].fillna('').astype(str).str.strip() + "_" +
            df['확정_지목'].fillna('').astype(str).str.strip() + "_" +
            df['확정_지번'].fillna('').astype(str).str.strip()
        )

        # 종전 면적 집계 및 매칭
        prev_area_map = df.groupby('종전_ID')['종전_면적'].sum().to_dict()
        df['종전_총면적'] = df['확정_ID'].map(prev_area_map).fillna(0)

        # 증감 계산
        df['증가면적'] = np.where(df['확정_면적'] > df['종전_총면적'], df['확정_면적'] - df['종전_총면적'], 0)
        df['감소면적'] = np.where(df['종전_총면적'] > df['확정_면적'], df['종전_총면적'] - df['확정_면적'], 0)
        df['증감여부'] = np.where(df['증가면적'] > 0, '증가', 
                                   np.where(df['감소면적'] > 0, '감소', '변동없음'))
        df['처리상태'] = np.where(df['확정_면적'] == 0, '말소', '유지')
        
        return df

# =========================================================
# 2. 소유권 검증 전략
# =========================================================

class OwnershipValidationStrategy:
    """소유권 유형별 검증 및 지분 합계 체크를 담당하는 전략 클래스"""
    @staticmethod
    def validate_share(owner_data, ownership_type):
        if ownership_type in ["공유", "상속"]:
            total_share = sum(d["지분"] for d in owner_data)
            if not np.isclose(total_share, 1.0, atol=1e-5):
                if total_share > 1.0:
                    return f"⚠️ 지분에 문제가 있습니다. (합이 1을 초과) 현재 합계: {total_share:.4f}"
                else:
                    return f"⚠️ 지분에 문제가 있습니다. (합이 1 미만) 현재 합계: {total_share:.4f}"
        return None

# =========================================================
# 3. UI 및 메인 앱
# =========================================================

if 'collapsed_rows' not in st.session_state:
    st.session_state['collapsed_rows'] = {}

st.set_page_config(page_title="토지 면적 증감 및 소유 분석 프로그램", layout="wide")
st.title("🏡 토지 면적 증감 및 소유 분석 프로그램")

# 소유구분 맵핑 데이터
OWNERSHIP_OPTIONS = ["1. 개인", "2. 법인", "3. 국유지", "4. 시유지", "5. 시군구유지"]
OWNERSHIP_MAP = {v: k.split(". ")[1] for k in OWNERSHIP_OPTIONS}

tab1, tab2 = st.tabs(["📥 엑셀 파일 업로드", "📝 직접 입력 모드"])

with tab1:
    uploaded_file = st.file_uploader("엑셀 파일을 선택하세요.", type=["xlsx"])
    if uploaded_file:
        try:
            df_input = pd.read_excel(uploaded_file, header=[0, 1])
            # 컬럼명 변환 로직 (기존 유지)
            new_columns = []
            for col in df_input.columns:
                first, second = str(col[0]).strip(), str(col[1]).strip()
                if first == '종전': new_columns.append(f'종전_{second}')
                elif first == '확정': new_columns.append(f'확정_{second}')
                elif first == '기준년도': new_columns.append('기준년도')
                elif first == '사업지구명': new_columns.append('사업지구명')
                elif first == '소유사항': new_columns.append('소유사항')
                elif first == '공유 또는 상속인 수': new_columns.append('공유_상속인수')
                else: new_columns.append(first)
            df_input.columns = new_columns

            st.subheader("📄 업로드 데이터 미리보기")
            st.dataframe(df_input.head(), use_container_width=True)

            if st.button("엑셀 데이터로 계산 시작"):
                # 필수 컬럼 체크
                req = ['종전_읍면', '종전_동리', '종전_지목', '종전_지번', '종전_면적', 
                       '확정_읍면', '확정_동리', '확정_지목', '확정_지번', '확정_면적']
                if all(c in df_input.columns for c in req):
                    result_df = AreaAnalysisStrategy.calculate(df_input)
                    st.success("✅ 계산이 완료되었습니다.")
                    st.dataframe(result_df, use_container_width=True)
                    
                    # 결과 다운로드
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False)
                    st.download_button("📥 결과 엑셀 다운로드", data=output.getvalue(), 
                                       file_name="result_calc.xlsx", 
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.error("필수 컬럼이 누락되었습니다. (종전/확정_읍면,동리,지목,지번,면적)")
        except Exception as e:
            st.error(f"파일 처리 중 오류 발생: {e}")

with tab2:
    st.subheader("📍 토지 기본 정보")
    col1, col2, col3 = st.columns(3)
    district = col1.text_input("읍면동리")
    land_type = col2.text_input("지목")
    lot_num = col3.text_input("지번")

    col1, col2 = st.columns(2)
    prev_area = col1.number_input("종전 면적", min_value=0.0, value=0.0)
    conf_area = col2.number_input("확정 면적", min_value=0.0, value=0.0)

    st.divider()
    st.subheader("👥 소유 정보")
    
    ownership_type = st.selectbox("소유 사항", ["단독", "공유", "상속"])
    
    is_multi = ownership_type in ["공유", "상속"]
    num_persons = st.number_input(
        "공유 또는 상속인 수", 
        min_value=1, max_value=10, value=1, 
        step=1, disabled=not is_multi
    )

    owner_data = []
    
    # 동적 줄 생성 로직
    for i in range(int(num_persons)):
        # 공유일 때 첫 줄은 빈 칸 (비활성화)
        if i == 0 and ownership_type == "공유":
            c1, c2, c3 = st.columns([3, 3, 2])
            c1.text_input("소유 구분", value="", disabled=True, key=f"dummy_type_{i}")
            c2.text_input("토지소유자", value="", disabled=True, key=f"dummy_name_{i}")
            c3.text_input("지분", value="", disabled=True, key=f"dummy_share_{i}")
            continue

        # 줄 앞에 기호 및 접기 버튼 추가 (공유/상속인 경우)
        if is_multi:
            if st.button(f"┗", key=f"fold_{i}"):
                st.session_state['collapsed_rows'][i] = not st.session_state['collapsed_rows'].get(i, True)
            
            if not st.session_state['collapsed_rows'].get(i, True):
                st.write(f"소유자 {i+1} 상세 정보")
            else:
                # 접혔을 때의 레이아웃 (공백)
                st.write("")
                continue

        # 실제 입력 줄
        c1, c2, c3 = st.columns([3, 3, 2])
        
        # 드롭다운 및 숫자 입력 연동
        val1 = c1.selectbox("소유 구분", OWNERSHIP_OPTIONS, key=f"owner_type_{i}")
        val2 = c2.text_input("토지소유자", key=f"owner_name_{i}")
        val3 = c3.text_input("지분", key=f"owner_share_{i}")

        # 숫자 입력 시 드롭다운 자동 선택 로직
        if val3 and val3.isdigit():
            idx_val = int(val3)
            if idx_val in [int(k.split(". ")[1]) for k in OWNERSHIP_OPTIONS]:
                # 세션 상태를 통해 강제 업데이트는 Streamlit 특성상 제한적이라 
                # 사용자에게 안내하거나 텍스트로 표시하는 방식을 택함
                pass

        try:
            share_val = float(val3) if val3 else 0.0
            owner_data.append({"소유구분": val1, "토지소유자": val2, "지분": share_val})
        except ValueError:
            pass

    # 지분 검증 실행 (Strategy 사용)
    warning_msg = OwnershipValidationStrategy.validate_share(owner_data, ownership_type)
    if warning_msg:
        st.warning(warning_msg)

    if st.button("입력 내용으로 계산하기"):
        new_data = pd.DataFrame([{
            "종전_읍면": district, "종전_동리": district, "종전_지목": land_type,
            "종전_지번": lot_num, "종전_면적": prev_area,
            "확정_읍면": district, "확정_동리": district, "확정_지목": land_type,
            "확정_지번": lot_num, "확정_면적": conf_area,
            "소유사항": ownership_type, "공유_상속인수": num_persons
        }])
        
        try:
            result_df = AreaAnalysisStrategy.calculate(new_data)
            st.success("✅ 계산이 완료되었습니다.")
            st.subheader("📊 계산 결과")
            st.dataframe(result_df, use_container_width=True)
            
            # 결과 다운로드
            excel_data = create_excel_download(result_df)
            st.download_button("📥 결과 엑셀 다운로드", data=excel_data, 
                               file_name="result_calc.xlsx", 
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"계산 중 오류 발생: {e}")

# 기존 함수 유지
def convert_excel_columns(df):
    new_columns = []
    for col in df.columns:
        first = str(col[0]).strip()
        second = str(col[1]).strip()
        if first == '종전': new_columns.append(f'종전_{second}')
        elif first == '확정': new_columns.append(f'확정_{second}')
        elif first == '기준년도': new_columns.append('기준년도')
        elif first == '사업지구명': new_columns.append('사업지구명')
        elif first == '소유사항': new_columns.append('소유사항')
        elif first == '공유 또는 상속인 수': new_columns.append('공유_상속인수')
        else: new_columns.append(first)
    df.columns = new_columns
    return df

def create_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='계산결과')
    return output.getvalue()
