import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# =========================================================
# 페이지 설정
# =========================================================
st.set_page_config(page_title="토지 면적 증감 및 소유 분석 프로그램", layout="wide")

# 세션 상태 초기화 (접기/펴기 상태 저장용)
if 'collapsed_rows' not in st.session_state:
    st.session_state['collapsed_rows'] = {}

# =========================================================
# 공통 함수
# =========================================================

def make_id(df, prefix):
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
    df = df.copy()
    
    # 필수 컬럼 확인
    required_columns = [
        '종전_읍면', '종전_동리', '종전_지목', '종전_지번', '종전_면적',
        '확정_읍면', '확정_동리', '확정_지목', '확정_지번', '확정_면적'
    ]
    
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"필수 컬럼이 누락되었습니다: {col}")

    # 면적 숫자 변환
    df['종전_면적'] = pd.to_numeric(df['종전_면적'], errors='coerce').fillna(0)
    df['확정_면적'] = pd.to_numeric(df['확정_면적'], errors='coerce').fillna(0)

    # ID 생성 및 매칭
    df['종전_ID'] = make_id(df, '종전')
    df['확정_ID'] = make_id(df, '확정')
    prev_area_map = df.groupby('종전_ID')['종전_면적'].sum().to_dict()
    df['종전_총면적'] = df['확정_ID'].map(prev_area_map).fillna(0)

    # 증감 계산
    df['증가면적'] = np.where(df['확정_면적'] > df['종전_총면적'], df['확정_면적'] - df['종전_총면적'], 0)
    df['감소면적'] = np.where(df['종전_총면적'] > df['확정_면적'], df['종전_총면적'] - df['확정_면적'], 0)
    df['증감여부'] = np.where(df['증가면적'] > 0, '증가', 
                                np.where(df['감소면적'] > 0, '감소', '변동없음'))
    df['처리상태'] = np.where(df['확정_면적'] == 0, '말소', '유지')

    # 컬럼 순서 재배치
    # 요청사항: 소유사항, 공유_상속인수 -> 증감여부 왼쪽
    cols_to_keep = ['종전_읍면', '종전_동리', '종전_지목', '종전_지번', '종전_면적', 
                     '확정_읍면', '확정_동리', '확정_지목', '확정_지번', '확정_면적',
                     '소유사항', '공유_상속인수', '증가면적', '감소면적', '증감여부', '처리상태']
    
    # 존재하는 컬럼만 필터링
    final_cols = [c for c in cols_to_keep if c in df.columns]
    return df[final_cols]

# =========================================================
# UI 영역
# =========================================================

st.title("🏡 토지 면적 증감 및 소유 분석 프로그램")

tab1, tab2 = st.tabs(["📥 엑셀 파일 업로드", "📝 직접 입력 모드"])

# 소유구분 맵핑 데이터
ownership_map = {
    "1. 개인": "개인",
    "2. 법인": "법인",
    "3. 국유지": "국유지",
    "4. 시유지": "시유지",
    "5. 시군구유지": "시군구유지"
}

with tab1:
    uploaded_file = st.file_uploader("엑셀 파일을 선택하세요.", type=["xlsx"])
    
    if uploaded_file:
        try:
            df_input = pd.read_excel(uploaded_file, header=[0, 1])
            
            # 컬럼명 변환 (기존 로직 유지)
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

            # --- 추가 입력 섹션 ---
            st.divider()
            st.subheader("📝 소유 정보 추가 입력")
            
            # 예시를 위해 첫 번째 행에 대해 입력을 받는 구조 (실제 구현시 루프나 data_editor 활용 가능)
            with st.container():
                st.info("파일을 업로드한 후, 하단에서 소유 사항을 상세 입력해주세요.")
                
                col_info1, col_info2, col_info3 = st.columns(3)
                row_district = col_info1.text_input("읍면동리(수정)")
                row_land_type = col_info2.text_input("지목(수정)")
                row_lot_num = col_info3.text_input("지번(수정)")

                # 소유 사항 입력
                own_type = st.selectbox("소유 사항", list(ownership_map.keys()))
                is_multi = own_type in ["공유", "상속"] # 실제 값은 "공유", "상속"이 포함된 키
                
                # 공유/상속인 수 입력 (비활성화 처리)
                num_persons = st.number_input(
                    "공유 또는 상속인 수", 
                    min_value=1, max_value=10, value=1, 
                    disabled=not (own_type == "공유" or own_type == "상속")
                )

                owner_list = []
                # 줄 추가 로직
                for i in range(int(num_persons)):
                    # 첫 줄 접기/펴기 (공유일 때만)
                    is_first = (i == 0)
                    if is_first and "공유" in own_type:
                        if st.button("┗", key=f"fold_{i}"):
                            st.session_state['collapsed_rows'][i] = not st.session_state['collapsed_rows'].get(i, True)
                        
                        if not st.session_state['collapsed_rows'].get(i, True):
                            st.write("자세한 소유자 정보는 아래에 있습니다.")
                        else:
                            # 접혔을 때 빈 칸 처리 (요청사항: 첫줄은 공란 및 비활성화)
                            c1, c2, c3 = st.columns(3)
                            c1.text_input("소유 구분", value="", disabled=True, key=f"o1_{i}")
                            c2.text_input("토지소유자", value="", disabled=True, key=f"o2_{i}")
                            c3.text_input("지분", value="", disabled=True, key=f"o3_{i}")
                            continue

                    # 일반 줄 입력
                    c1, c2, c3 = st.columns([3, 3, 2])
                    val1 = c1.selectbox("소유 구분", list(ownership_map.keys()), key=f"o1_{i}")
                    val2 = c2.text_input("토지소유자", key=f"o2_{i}")
                    val3 = c3.text_input("지분", key=f"o3_{i}")

                    # 숫자 매핑 기능 (입력창에 숫자를 넣으면 드롭다운 자동 선택)
                    if val3 != "" and val3.isdigit():
                        idx = int(val3)
                        if idx in ownership_map:
                            st.session_state[f"o1_{i}"] = ownership_map[idx]

                    # 데이터 저장
                    try:
                        share_val = float(val3) if val3 else 0.0
                        owner_list.append({"소유구분": val1, "토지소유자": val2, "지분": share_val})
                    except:
                        pass

                # 지분 검증
                if "공유" in own_type or "상속" in own_type:
                    total_share = sum(d["지분"] for d in owner_list)
                    if not np.isclose(total_share, 1.0, atol=1e-5):
                        if total_share > 1.0:
                            st.error(f"⚠️ 지분에 문제가 있습니다. (합이 1을 초과) 현재 합계: {total_share}")
                        else:
                            st.error(f"⚠️ 지분에 문제가 있습니다. (합이 1 미만) 현재 합계: {total_share}")
                    else:
                        st.success("✅ 지분 합계가 1입니다.")

                if st.button("결과 계산하기"):
                    # 업로드 데이터와 사용자 입력 데이터 결합 로직 수행
                    # (이 부분은 실제 구현 시 df_input의 인덱스와 매칭하여 합산 처리)
                    st.info("입력된 데이터 기반 계산을 진행합니다...")
                    # 계산 로직 호출...

        except Exception as e:
            st.error(f"오류 발생: {e}")

else:
    st.info("엑셀 파일을 업로드하면 분석 도구가 활성화됩니다.")

# [참고] 직접 입력 모드는 기존과 동일하게 유지하되, 
# 위와 동일한 동적 UI 로직을 적용하여 구성할 수 있습니다.
