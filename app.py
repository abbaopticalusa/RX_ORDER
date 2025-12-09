import streamlit as st
import pandas as pd
import openpyxl
import base64
from io import BytesIO
import os # 🚨 추가
import sys # 🚨 추가

# ============================================================================
# 1. 환경 설정 및 변수 정의
# ============================================================================
BASE64_FILE_NAME = "excel_template.txt"
BASE64_EXCEL = ""

# 🚨 EXE 환경에서 파일을 읽기 위한 안정적인 경로 설정 로직 🚨
def find_data_file(filename):
    if getattr(sys, "frozen", False):
        # PyInstaller로 빌드된 EXE 환경
        base_path = sys._MEIPASS
    else:
        # 일반 Python 환경
        base_path = os.path.dirname(__file__)

    return os.path.join(base_path, filename)

# Base64 코드를 파일에서 읽어옴 (IndexError 및 경로 오류 방지)
try:
    data_file_path = find_data_file(BASE64_FILE_NAME)
    with open(data_file_path, "r", encoding='utf-8') as f:
        BASE64_EXCEL = f.read()
except FileNotFoundError:
    st.error(f"Error: The required file '{BASE64_FILE_NAME}' was not found at {data_file_path}. Please check the spec file and file location.")
    st.stop()
except Exception as e:
    st.error(f"An unexpected error occurred while reading the Base64 file: {e}")
    st.stop()


# A열부터 R열까지 모든 열을 정의 (총 18개 열)
ALL_COLUMNS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']
DROPDOWN_COLS = ['J', 'K', 'L', 'M'] # 드롭다운이 필요한 열 

# A열부터 R열까지의 사용자 정의 헤더 (VBA 원본 기반)
USER_HEADERS = [
    'Order #', 'Eye', 'Sph', 'Cyl', 'Axis', 'Prism 1', 'Add', 'PD', 'HT', 
    'MATERIAL', 'Products', 'tint', 'Coating', 'A', 'B', 'DBL', 'ED', 'Qty'
]
HEADER_MAPPING = {f'Col_{ALL_COLUMNS[i]}': USER_HEADERS[i] for i in range(len(ALL_COLUMNS))}


# ============================================================================
# 2. 기능 구현
# ============================================================================

# Base64 문자열을 엑셀 객체로 복원하는 함수
def get_workbook_from_code():
    decoded_data = base64.b64decode(BASE64_EXCEL)
    return BytesIO(decoded_data)

# 'DATA' 시트에서 드롭다운 목록 읽어오기
@st.cache_data
def load_options():
    try:
        # DATA 시트를 읽기 위해 포맷을 유지한 채 로드합니다.
        wb = openpyxl.load_workbook(get_workbook_from_code(), data_only=True, keep_vba=True)
        # 🚨 시트 이름 확인: 드롭다운 목록이 있는 시트 이름으로 변경하세요.
        ws = wb['DATA'] 
        
        options = {'J':[], 'K':[], 'L':[], 'M':[]}
        col_map = {1:'J', 2:'K', 3:'L', 4:'M'} 
        
        for col_idx, key in col_map.items():
            for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx, values_only=True):
                if row[0] is not None:
                    options[key].append(str(row[0]))
        return options
    except Exception as e:
        st.error(f"데이터 로드 중 오류 발생: {e}")
        return {}

# 엑셀 생성 및 데이터 주입 함수
def create_order_file(user_df):
    # 1. 메모리 상에서 원본 형식을 가진 엑셀 로드
    input_stream = get_workbook_from_code()
    # 파일 형식 오류 해결: VBA를 로드하지 않음 (keep_vba=False)
    wb = openpyxl.load_workbook(input_stream, keep_vba=False) 
    
    # 🚨 시트 이름 확인: 주문 데이터를 입력할 시트 이름으로 변경하세요.
    # ws = wb['ORDER'] # 이전 코드에서 확인된 시트 이름
    # ws = wb['Sheet1'] # Sheet1으로 가정하고 진행합니다. 만약 'ORDER'가 확실하다면 'ORDER'로 유지해주세요.
    ws = wb['ORDER'] 
    
    # helper 함수: 값 추출 
    def extract_value(data):
        if isinstance(data, list) and len(data) > 0:
            return data[0]
        return data

    # 2. 사용자 데이터 입력 (Row 3 ~ 33, A열~R열)
    start_row = 3
    for i, row in user_df.iterrows():
        current_row = start_row + i
        if current_row > 33: break
        
        # A(1) ~ R(18) 열에 값 입력
        for col_index, col_name in enumerate(ALL_COLUMNS):
            df_col_key = f'Col_{col_name}'
            excel_col_index = col_index + 1 
            
            ws.cell(row=current_row, column=excel_col_index).value = extract_value(row[df_col_key])
    
    
    # 2.5. 열 너비 수정 (Column Width Adjustment)
    COLUMN_WIDTHS = {
        'A': 45, 'B': 5, 'C': 5, 'D': 5, 'E': 5, 'F': 8, 'G': 5, 'H': 5, 'I': 5,
        'J': 20, 'K': 50, 'L': 20, 'M': 15, 'N': 5, 'O': 5, 'P': 5, 'Q': 5, 'R': 5 
    }
    
    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width
    
    # 3. 저장
    output = BytesIO()
    wb.save(output) 
    output.seek(0)
    return output

# ============================================================================
# 3. 화면 UI (Streamlit)
# ============================================================================
st.set_page_config(page_title="Plazma Order System", layout="wide")

st.title("ABBA Optical Champion Order App – from the last legacy of SANG.")
st.caption("Fill out the forms, save your work, and export to Excel when done.")

# 옵션 로드
opts = load_options()
if not opts:
    st.warning("옵션을 불러오지 못했습니다. Base64 코드를 확인해주세요.")
    st.stop()

# 데이터 편집기 초기화 (A열부터 R열까지 31줄 초기화)
if 'df_input' not in st.session_state:
    initial_data = {f'Col_{col}': [None] * 31 for col in ALL_COLUMNS}
    st.session_state.df_input = pd.DataFrame(initial_data)

# 컬럼 설정 (사용자 정의 헤더 이름 적용)
col_conf = {}
for col in ALL_COLUMNS:
    col_key = f'Col_{col}'
    header_name = HEADER_MAPPING.get(col_key, f"{col}열")
    
    if col in DROPDOWN_COLS:
        # J, K, L, M은 드롭다운
        col_conf[col_key] = st.column_config.SelectboxColumn(
            header_name, options=opts[col], required=False
        )
    else:
        # 나머지 열은 일반 텍스트 입력 필드
        col_conf[col_key] = st.column_config.TextColumn(
            header_name, required=False
        )


# 그리드 표시
st.write("### 주문 내역 입력 (A3:R33)")
edited_data = st.data_editor(
    st.session_state.df_input,
    column_config=col_conf,
    num_rows="fixed",
    hide_index=True,
    use_container_width=True,
    height=600
)

# 다운로드 버튼
st.write("---")
if st.button("🚀 엑셀 파일 생성 및 다운로드", type="primary"):
    excel_file = create_order_file(edited_data)
    
    st.download_button(
        label="📥 DOWNLOAD / 결과물 다운로드 (.xlsx)",
        data=excel_file,
        file_name="Plazma_Order_Result.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.success("FINISHED. CHECK THE EXCEL FILE THAT WAS DOWNLOADED / 완료! 다운로드된 엑셀 파일을 확인해주세요.")