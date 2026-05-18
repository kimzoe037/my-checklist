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

# 유틸리티 및 템플릿 매핑
TEMPLATE_MAP = {'여행': '여행준비', '시험': '시험준비', '시공': '건축시공'}

def get_color(importance):
    imp = str(importance).strip()
    if imp == '상': return '🔴'
    if imp == '중': return '🟡'
    if imp == '하': return '🟢'
    return '⚪'

def sanitize_id(text):
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', text)

def load_excel_template(keyword):
    file_name = 'Checkmaster.xlsx'
    if not os.path.exists(file_name): return []
    try:
        excel_file = pd.ExcelFile(file_name)
        target_sheet = next((s for s in excel_file.sheet_names if keyword in s), None)
        if not target_sheet: return []

        df = pd.read_excel(file_name, sheet_name=target_sheet, header=None)
        items = []
        for i, row in df.iterrows():
            if len(row) < 4 or str(row[2]) in ['체크 항목', '내용', 'nan', 'None']: continue
            items.append({
                'id': f"{sanitize_id(target_sheet)}_{i}_{datetime.now().timestamp()}", 
                'theme': str(row[1]) if pd.notna(row[1]) else "기본",
                'content': str(row[2]).strip(),
                'importance': str(row[3]) if pd.notna(row[3]) else '중'
            })
        return items
    except Exception: return []

# 세션 초기화
if 'step' not in st.session_state: st.session_state.step = 1
if 'items_data' not in st.session_state: st.session_state.items_data = {}
if 'cat_dates' not in st.session_state: st.session_state.cat_dates = {}

# STEP 1: 화면
if st.session_state.step == 1:
    st.markdown("<div class='step1-container'><h1>🔍 Checklist</h1><h3>어떤 프로젝트를 준비하시나요?</h3></div>", unsafe_allow_html=True)
    with st.form(key="step1_form"):
        categories_input = st.text_input("카테고리 입력 (쉼표로 구분)", placeholder="예: 상하이 여행준비, 기말시험, 거실 시공")
        if st.form_submit_button("생성하기 🚀"):
            for cat in [c.strip() for c in categories_input.split(",") if c.strip()]:
                if cat not in st.session_state.items_data:
                    st.session_state.cat_dates[cat] = date.today() + timedelta(days=10)
                    # 키워드 매칭 로직 적용
                    matched_key = next((k for k in TEMPLATE_MAP if k in cat), None)
                    st.session_state.items_data[cat] = load_excel_template(TEMPLATE_MAP[matched_key]) if matched_key else []
            st.session_state.step = 2
            st.rerun()

# STEP 2: 체크리스트 메인
elif st.session_state.step == 2:
    sheet_names = list(st.session_state.items_data.keys())
    if st.button("⬅️ 초기화"):
        st.session_state.clear(); st.session_state.step = 1; st.rerun()

    tabs = st.tabs(sheet_names)
    for idx, tab in enumerate(tabs):
        with tab:
            cat = sheet_names[idx]
            current_items = st.session_state.items_data[cat]
            themes = sorted(list(set(item['theme'] for item in current_items)))
            
            # 레이아웃 및 기능 구현
            st.title(f"{cat.upper()} LIST")
            edit_mode = st.toggle("🛠️ 편집 모드", key=f"ed_{cat}")
            
            # 항목 추가
            with st.expander("➕ 항목 추가"):
                with st.form(key=f"form_{cat}"):
                    t = st.text_input("테마", "기본")
                    c = st.text_input("내용")
                    i = st.select_slider("중요도", ['하','중','상'], '중')
                    if st.form_submit_button("추가") and c:
                        st.session_state.items_data[cat].append({'id': f"new_{datetime.now().timestamp()}", 'theme': t, 'content': c, 'importance': i})
                        st.rerun()

            # 리스트 렌더링
            for theme in themes:
                st.markdown(f"### 📍 {theme}")
                t_items = [i for i in current_items if i['theme'] == theme]
                for item in t_items:
                    c1, c2 = st.columns([9, 1])
                    with c1: st.checkbox(f"{get_color(item['importance'])} {item['content']}", key=item['id'])
                    with c2:
                        if edit_mode and st.button("🗑️", key=f"del_{item['id']}"):
                            st.session_state.items_data[cat].remove(item); st.rerun()
