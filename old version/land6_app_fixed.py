import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re

# =========================================================
# 1. 전략 패턴 정의 (분석 전략)
# =========================================================

class AreaAnalysisStrategy:
    """면적 계산 및 증감 분석을 담당하는 전략 클래스"""

    @staticmethod
    def calculate(df):
        df = df.copy()

        # 면적 숫자 변환
        df["종전_면적"] = pd.to_numeric(df["종전_면적"], errors="coerce").fillna(0)
        df["확정_면적"] = pd.to_numeric(df["확정_면적"], errors="coerce").fillna(0)

        # 종전 ID 생성
        df["종전_ID"] = (
            df["종전_읍면"].fillna("").astype(str).str.strip()
            + "_" + df["종전_동리"].fillna("").astype(str).str.strip()
            + "_" + df["종전_지목"].fillna("").astype(str).str.strip()
            + "_" + df["종전_지번"].fillna("").astype(str).str.strip()
        )

        # 확정 ID 생성
        df["확정_ID"] = (
            df["확정_읍면"].fillna("").astype(str).str.strip()
            + "_" + df["확정_동리"].fillna("").astype(str).str.strip()
            + "_" + df["확정_지목"].fillna("").astype(str).str.strip()
            + "_" + df["확정_지번"].fillna("").astype(str).str.strip()
        )

        # 종전 면적 집계
        prev_area_map = df.groupby("종전_ID")["종전_면적"].sum().to_dict()

        # 확정 ID 기준으로 종전 총면적 매칭
        df["종전_총면적"] = df["확정_ID"].map(prev_area_map).fillna(0)

        # 증감 계산
        df["증가면적"] = np.where(df["확정_면적"] > df["종전_총면적"], df["확정_면적"] - df["종전_총면적"], 0)
        df["감소면적"] = np.where(df["종전_총면적"] > df["확정_면적"], df["종전_총면적"] - df["확정_면적"], 0)
        df["증감여부"] = np.where(df["증가면적"] > 0, "증가", np.where(df["감소면적"] > 0, "감소", "변동없음"))
        df["처리상태"] = np.where(df["확정_면적"] == 0, "말소", "유지")

        return df


# =========================================================
# 2. 소유권 검증 및 유틸리티
# =========================================================

def parse_share(value):
    """분수(1/2) 또는 소수(0.5)를 float로 변환"""
    if not value: return 0.0
    try:
        if "/" in str(value):
            num, den = str(value).split("/")
            return float(num) / float(den)
        return float(str(value))
    except:
        return 0.0

class OwnershipValidationStrategy:
    @staticmethod
    def validate_share(owner_data):
        if not owner_data: return None
        total_share = sum(item["지분"] for item in owner_data)
        
        if not np.isclose(total_share, 1.0, atol=1e-5):
            if total_share > 1.0:
                return f"⚠️ 지분의 합이 1이 아닙니다.(1 초과) 현재 합계: {total_share:.4f}"
            else:
                return f"⚠️ 지분의 합이 1이 아닙니다.(1 미만) 현재 합계: {total_share:.4f}"
        return None


# =========================================================
# 3. 엑셀 관련 함수
# =========================================================

def convert_excel_columns(df):
    new_columns = []
    for col in df.columns:
        if isinstance(col, tuple):
            first = str(col[0]).strip()
            second = str(col[1]).strip()
        else:
            first = str(col).strip()
            second = ""

        if first == "종전": new_columns.append(f"종전_{second}")
        elif first == "확정": new_columns.append(f"확정_{second}")
        elif first == "기준년도": new_columns.append("기준년도")
        elif first == "사업지구명": new_columns.append("사업지구명")
        elif first == "소유사항": new_columns.append("소유사항")
        elif first == "공유 또는 상속인 수": new_columns.append("공유_상속인수")
        else: new_columns.append(first)

    df = df.copy()
    df.columns = new_columns
    return df

def create_excel_download(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="계산결과")
    return output.getvalue()


# =========================================================
# 4. Streamlit 설정 및 기본 데이터
# =========================================================

st.set_page_config(page_title="토지 분석 프로그램", layout="wide")
st.title("🏡 토지 면적 증감 및 소유 분석 프로그램")

OWNERSHIP_OPTIONS = ["1. 개인", "2. 법인", "3. 국유지", "4. 시유지", "5. 시군구유지"]
OWNERSHIP_MAP = {option: option.split(". ", 1)[1] for option in OWNERSHIP_OPTIONS}

if "collapsed_rows" not in st.session_state:
    st.session_state["collapsed_rows"] = {}

tab1, tab2 = st.tabs(["📥 엑셀 파일 업로드", "📝 직접 입력 모드"])


# =========================================================
# TAB 1: 엑셀 파일 업로드
# =========================================================
with tab1:
    uploaded_file = st.file_uploader("엑셀 파일을 선택하세요.", type=["xlsx"])

    if uploaded_file:
        try:
            df_input = pd.read_excel(uploaded_file, header=[0, 1])
            df_input = convert_excel_columns(df_input)

            st.subheader("📄 업로드 데이터 미리보기")
            st.dataframe(df_input.head(), use_container_width=True)

            if st.button("엑셀 데이터로 계산 시작", key="calculate_excel"):
                required_columns = [
                    "종전_읍면", "종전_동리", "종전_지목", "종전_지번", "종전_면적",
                    "확정_읍면", "확정_동리", "확정_지목", "확정_지번", "확정_면적"
                ]

                missing_columns = [col for col in required_columns if col not in df_input.columns]
                if missing_columns:
                    st.error(f"필수 컬럼이 누락되었습니다: {', '.join(missing_columns)}")
                else:
                    result_df = AreaAnalysisStrategy.calculate(df_input)
                    
                    # 요청 사항: 특정 컬럼 제외
                    cols_to_drop = ["종전_ID", "확정_ID", "종전_총면적"]
                    result_df = result_df.drop(columns=[c for c in cols_to_drop if c in result_df.columns])

                    st.success("✅ 계산이 완료되었습니다.")

                    # 결과 표시 및 동적 행 생성 UI
                    st.subheader("📊 계산 결과 및 소유자 입력")
                    
                    # 결과 데이터 보관용 리스트 (상호작용을 위해)
                    display_data = result_df.to_dict('records')
                    
                    # 결과 순회 및 UI 렌더링
                    for idx, row in enumerate(display_data):
                        with st.expander(f"Row {idx+1}: {row['확정_지번']} ({row['처리상태']})", expanded=True):
                            # 기본 정보 표시
                            c1, c2, c3, c4, c5, c6 = st.columns(6)
                            c1.write(f"**종전**: {row['종전_읍면']}{row['종전_동리']} {row['종전_지목']} {row['종전_지번']}")
                            c2.write(f"**확정**: {row['확정_읍면']}{row['확정_동리']} {row['확정_지목']} {row['확정_지번']}")
                            c3.write(f"**면적**: {row['종전_면적']} $\rightarrow$ {row['확정_면적']}")
                            c4.write(f"**증감**: {row['증감여부']} ({row['증가면적']}+{row['감소면적']})")
                            c5.write(f"**소유사항**: {row['소유사항']}")
                            c6.write(f"**공유인수**: {row['공유_상속인수']}")

                            # 소유 정보 입력 영역
                            st.divider()
                            owner_rows = []

                            # 엑셀의 빈 셀은 NaN이 될 수 있으므로 안전하게 정수로 변환
                            raw_num_owners = row.get("공유_상속인수", 1)
                            try:
                                if pd.isna(raw_num_owners):
                                    num_owners = 1
                                else:
                                    num_owners = int(float(raw_num_owners))
                                    if num_owners < 1:
                                        num_owners = 1
                            except (TypeError, ValueError):
                                num_owners = 1
                            
                            # 요청 사항: 공유_상속인수에 따른 동적 행 생성
                            for i in range(num_owners):
                                is_first = (i == 0)
                                
                                # 접기 기능 (첫번째 행만)
                                if is_first and num_owners > 1:
                                    if "collapsed_" not in st.session_state: st.session_state["collapsed_rows"] = {}
                                    if "fold_" + str(idx) not in st.session_state: st.session_state["fold_" + str(idx)] = False
                                    
                                    if st.button("┗", key=f"fold_{idx}"):
                                        st.session_state["fold_" + str(idx)] = not st.session_state["fold_" + str(idx)]
                                    
                                    if st.session_state["fold_" + str(idx)]:
                                        st.info("소유자 상세 정보가 접혀 있습니다.")
                                        continue

                                # 입력 필드
                                row_col1, row_col2, row_col3 = st.columns([3, 3, 2])
                                
                                # 소유구분 드롭다운 (직접 입력 모드와 동일하게 번호 포함)
                                val_type = row_col1.selectbox("소유구분", OWNERSHIP_OPTIONS, 
                                                                 key=f"res_type_{idx}_{i}",
                                                                 index=0 if i == 0 else 0)
                                
                                val_name = row_col2.text_input("토지소유자", key=f"res_name_{idx}_{i}")
                                
                                val_share_str = row_col3.text_input("지분", key=f"res_share_{idx}_{i}", placeholder="예: 0.5 또는 1/2")
                                
                                # 지분 계산 및 합계 검증
                                share_val = parse_share(val_share_str)
                                owner_rows.append({"지분": share_val})

                                # 합계 검증 및 스타일링
                                if i == num_owners - 1 and num_owners > 1:
                                    warning = OwnershipValidationStrategy.validate_share(owner_rows)
                                    if warning:
                                        st.error(warning) # 팝업 대신 st.error 사용
                                        # 빨간색 강조 표시 (Markdown 활용)
                                        st.markdown(f"⚠️ <span style='color:red; font-weight:bold;'>지분 합계 오류: {val_share_str}</span>", unsafe_allow_html=True)

                            # 엑셀 다운로드 버튼 (전체 결과용)
                            # 실제 다운로드를 위해 결과 df를 업데이트하는 로직은 추가 구현 필요
                            # 여기서는 계산된 결과물 자체를 다운로드 가능하게 함
                            st.info("※ 결과 데이터는 계산 완료 후 위에서 확인 가능하며, 다운로드 버튼을 통해 저장하세요.")

            # 결과 다운로드 버튼
            excel_data = create_excel_download(result_df)
            st.download_button("📥 결과 엑셀 다운로드", data=excel_data, 
                                file_name="result_calc.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        except Exception as e:
            st.error(f"파일 처리 중 오류 발생: {e}")


# =========================================================
# TAB 2: 직접 입력 모드 (기존 유지하되 수정사항 반영 가능)
# =========================================================
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
    ownership_type = st.selectbox("소유 사항", ["단독", "공유", "상속"], key="ownership_type")
    
    is_multi = ownership_type in ["공유", "상속"]
    num_persons = st.number_input("공유 또는 상속인 수", min_value=1, max_value=10, value=1, step=1, 
                                    disabled=not is_multi, key="num_persons")

    owner_data = []
    try:
        safe_num_persons = int(float(num_persons))
    except (TypeError, ValueError):
        safe_num_persons = 1
    safe_num_persons = max(1, min(10, safe_num_persons))

    for i in range(safe_num_persons):
        if i == 0 and ownership_type == "공유":
            c1, c2, c3 = st.columns([3, 3, 2])
            c1.text_input("소유 구분", value="", disabled=True, key=f"dummy_type_{i}")
            c2.text_input("토지소유자", value="", disabled=True, key=f"dummy_name_{i}")
            c3.text_input("지분", value="", disabled=True, key=f"dummy_share_{i}")
            continue

        if is_multi:
            if i not in st.session_state.get("collapsed_rows", {}):
                st.session_state["collapsed_rows"][i] = False
            if st.button("┗", key=f"fold_{i}"):
                st.session_state["collapsed_rows"][i] = not st.session_state["collapsed_rows"][i]
            if st.session_state["collapsed_rows"][i]:
                st.write(f"소유자 {i + 1} 입력 행이 접혀 있습니다.")
                continue
            st.write(f"소유자 {i + 1} 상세 정보")

        c1, c2, c3 = st.columns([3, 3, 2])
        val1 = c1.selectbox("소유 구분", OWNERSHIP_OPTIONS, key=f"owner_type_{i}")
        val2 = c2.text_input("토지소유자", key=f"owner_name_{i}")
        val3 = c3.text_input("지분", key=f"owner_share_{i}", placeholder="예: 0.5")

        try:
            share_val = parse_share(val3)
            if share_val < 0:
                st.warning(f"소유자 {i + 1}: 지분은 음수가 될 수 없습니다.")
                share_val = 0.0
            owner_data.append({"소유구분": val1, "토지소유자": val2, "지분": share_val})
        except ValueError:
            st.warning(f"소유자 {i + 1}: 지분은 숫자로 입력해주세요.")

    warning_msg = OwnershipValidationStrategy.validate_share(owner_data)
    if warning_msg:
        st.warning(warning_msg)

    if st.button("입력 내용으로 계산하기", key="calculate_manual"):
        new_data = pd.DataFrame([{
            "종전_읍면": district, "종전_동리": district, "종전_지목": land_type, "종전_지번": lot_num, "종전_면적": prev_area,
            "확정_읍면": district, "확정_동리": district, "확정_지목": land_type, "확정_지번": lot_num, "확정_면적": conf_area,
            "소유사항": ownership_type, "공유_상속인수": num_persons
        }])
        try:
            result_df = AreaAnalysisStrategy.calculate(new_data)
            st.success("✅ 계산이 완료되었습니다.")
            st.dataframe(result_df, use_container_width=True)
            excel_data = create_excel_download(result_df)
            st.download_button("📥 결과 엑셀 다운로드", data=excel_data, 
                                file_name="result_calc.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="download_manual_result")
        except Exception as e:
            st.error(f"계산 중 오류 발생: {e}")
