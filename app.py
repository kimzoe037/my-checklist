import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import re
import os

# 1. 페이지 설정 및 디자인
st.set_page_config(page_title="체크 마스터", layout="wide")

def apply_custom_style():
    st.markdown("""
        <style>
        .main { background-color: #F5F5DC; }
        html, body, [class*="st-"] { font-size: 18px; }
        h1 { font-family: 'Pretendard', sans-serif; font-weight: 800; color: #2c3e50; font-size: 2.5rem !important; }
        .step1-container { text-align: center; max-width: 600px; margin: 0 auto; padding-top: 20px; }
        .nav-bar { background-color: rgba(93, 64, 55, 0.1); padding: 15px; border-radius: 10px; margin-bottom: 20px; display: flex; flex-wrap: wrap; gap: 10px; border: 1px solid #d7ccc8; }
        .nav-item { padding: 5px 15px; background-color: #5d4037; color: white !important; border-radius: 20px; text-decoration: none; font-size: 16px; font-weight: bold; }
        .theme-header { margin-top: 40px; margin-bottom: 25px; padding-bottom: 12px; border-bottom: 3px solid #5d4037; color: #3e2723; }
        div[data-testid="stCheckbox"] { padding: 15px 20px; margin-bottom: 12px; background-color: rgba(255, 255, 255, 0.7); border-radius: 8px; border: 1px solid #d7ccc8; }
        div[data-testid="stCheckbox"] label p { font-size: 20px !important; font-weight: 500; }
        .alert-box { background-color: #fff3cd; border-left: 5px solid #ffca28; padding: 20px; border-radius: 5px; margin-bottom: 20px; }
        .preview-paper { background-color: white; width: 100%; max-width: 800px; padding: 40px; margin: 0 auto; box-shadow: 0 0 10px rgba(0,0,0,0.1); color: black; border: 1px solid #ddd; }
        .preview-theme { font-size: 20px; font-weight: bold; border-bottom: 2px solid #333; margin-top: 30px; padding-bottom: 5px; }
        .preview-item { font-size: 16px; margin: 12px 0; border-bottom: 1px solid #eee; padding-bottom: 5px; }
        </style>
    """, unsafe_allow_html=True)

apply_custom_style()

# 2. 유틸리티 함수
def get_color(importance):
    imp = str(importance).strip()
    if imp == '상': return '🔴'
    if imp == '중': return '🟡'
    if imp == '하': return '🟢'
    return '⚪'

def sanitize_id(text):
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', text)

# [핵심 변경] 입력된 카테고리 텍스트에 포함된 단어로 시트를 정확하게 찾아내는 함수
def load_excel_template(keyword):
    file_name = 'Checkmaster.xlsx'
    if not os.path.exists(file_name):
        return []
    try:
        excel_file = pd.ExcelFile(file_name)
        target_sheet = None
        
        # 유저가 입력한 대분류 이름 (예: '전공시험준비', '상하이 여행준비')
        user_input = keyword.strip()
        
        # 엑셀 시트 이름들 중에서 유저가 입력한 글자에 포함되는 단어가 있는지 확인
        # 예: 시트 이름이 '시험' 일 때, '전공시험준비' 안에 '시험'이 들어있으므로 매칭 성공!
        for sheet in excel_file.sheet_names:
            sheet_name_clean = sheet.strip()
            if sheet_name_clean in user_input:
                target_sheet = sheet
                break
                
        # 만약 글자가 뒤집혀서 매칭될 수도 있으니 반대 방향도 한 번 더 검사 (예: 입력이 '여행', 시트명이 '가족여행')
        if not target_sheet:
            for sheet in excel_file.sheet_names:
                sheet_name_clean = sheet.strip()
                if user_input in sheet_name_clean:
                    target_sheet = sheet
                    break
                    
        # 그래도 매칭되는 시트가 전혀 없으면 첫 번째 시트를 안전장치로 가져옵니다
        if not target_sheet and len(excel_file.sheet_names) > 0:
            target_sheet = excel_file.sheet_names[0]
            
        if not target_sheet: 
            return [] 
        
        # 매칭된 시트 데이터 읽기
        df = pd.read_excel(file_name, sheet_name=target_sheet, header=None)
        category_items = []
        for i, row in df.iterrows():
            if len(row) < 4: continue
            content = str(row[2]).strip()
            if not content or content in ['체크 항목', '내용', 'nan', 'None']: continue
            category_items.append({
                'id': f"{sanitize_id(target_sheet)}_{i}_{datetime.now().timestamp()}", 
                'theme': str(row[1]) if pd.notna(row[1]) else "기본",
                'content': content,
                'importance': str(row[3]) if pd.notna(row[3]) else '중'
            })
        return category_items
    except Exception as e:
        return []

# 3. 세션 상태 초기화
if 'step' not in st.session_state: st.session_state.step = 1  
if 'items_data' not in st.session_state: st.session_state.items_data = {}  
if 'cat_dates' not in st.session_state: st.session_state.cat_dates = {}  

# STEP 1
if st.session_state.step == 1:
    st.markdown("<div class='step1-container'>", unsafe_allow_html=True)
    image_path = "character.png" 
    if os.path.exists(image_path):
        st.image(image_path, width=200)
    else:
        st.markdown("<h1 style='font-size: 60px; margin: 0;'>📋</h1>", unsafe_allow_html=True)
    st.title("Checklist")
    st.subheader("지금 어떤 프로젝트를 준비하고 계신가요?")
    st.write("나만의 맞춤형 체크리스트 대분류(카테고리)를 입력해 주세요.")
    st.markdown("</div>", unsafe_allow_html=True)
    
    with st.form(key="step1_form"):
        categories_input = st.text_input("카테고리 대분류 입력 (쉼표 `,` 로 구분하여 여러 개 등록 가능)", placeholder="예: 상하이 여행준비, 기말시험, 거실 시공")
        submit_btn = st.form_submit_button("나만의 체크리스트 생성하기 🚀")
        if submit_btn:
            if categories_input.strip():
                input_list = [c.strip() for c in categories_input.split(",") if c.strip()]
                for cat in input_list:
                    if cat not in st.session_state.items_data:
                        st.session_state.cat_dates[cat] = date.today() + timedelta(days=10)
                        # 입력창 텍스트 통째로 넘겨 똑똑해진 시트 검색 기능 실행
                        st.session_state.items_data[cat] = load_excel_template(cat)
                st.session_state.step = 2
                st.rerun()

# STEP 2
elif st.session_state.step == 2:
    sheet_names = list(st.session_state.items_data.keys())
    if not sheet_names:
        st.session_state.step = 1
        st.rerun()
    if st.button("⬅️ 다른 프로젝트 새로 만들기 (초기화)"):
        st.session_state.clear()
        st.session_state.step = 1
        st.rerun()
    tabs = st.tabs(sheet_names)
    for idx, tab in enumerate(tabs):
        with tab:
            cat = sheet_names[idx]
            current_items = st.session_state.items_data.get(cat, [])
            themes = sorted(list(
