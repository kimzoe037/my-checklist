import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import re
import os

# 1. 페이지 설정
st.set_page_config(page_title="체크 마스터", layout="wide")

# 2. 스타일 정의
st.markdown("""
    <style>
    .main { background-color: #F5F5DC; }
    .theme-header { margin-top: 30px; border-bottom: 2px solid #5d4037; }
    .alert-box { background-color: #fff3cd; border-left: 5px solid #ffca28; padding: 20px; border-radius: 5px; margin-bottom: 20px; }
    .preview-paper { background-color: white; padding: 40px; box-shadow: 0 0 10px rgba(0,0,0,0.1); color: black; }
    </style>
""", unsafe_allow_html=True)

# 3. 유틸리티 함수
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
        # 키워드가 포함된 시트 찾기
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
    except: return []

# 4. 세션 상태 초기화
if 'items_data' not in st.session_state: st.session_state.items_data = {}
if 'cat_dates' not in st.session_state: st.session_state.cat_dates = {}

# 5. 메인 로직
st.title("📋 Checkmaster")
input_text = st.text_input("프로젝트 카테고리 입력 (쉼표로 구분)", placeholder="건축시공, 여행준비, 시험준비")

if input_text:
    categories = [c.strip() for c in input_text.split(",") if c.strip()]
    for cat in categories:
        if cat not in st.session_state.items_data:
            st.session_state.cat_dates[cat] = date.today() + timedelta(days=10)
            matched = next((k for k in ['건축시공', '여행', '시험'] if k in cat), None)
            st.session_state.items_data[cat] = load_excel_template(matched) if matched else []

    tabs = st.tabs(categories)
    for idx, tab in enumerate(tabs):
        cat = categories[idx]
        with tab:
            current_items = st.session_state.items_data[cat]
            
            # 진행률 및 D-day 계산
            done = sum(1 for i in current_items if st.session_state.get(i['id'], False))
            total = len(current_items)
            
            col1, col2 = st.columns([2, 1])
            with col1:
                st.write(f"🏗️ 진행률: {done}/{total} ({(done/total*100) if total>0 else 0:.1f}%)")
                st.progress((done/total) if total > 0 else 0)
            with col2:
                st.session_state.cat_dates[cat] = st.date_input(f"📅 마감일", st.session_state.cat_dates[cat])
            
            # 편집 및 추가
            edit_mode = st.toggle("🛠️ 편집 모드", key=f"ed_{cat}")
            with st.expander("➕ 항목 추가"):
                with st.form(f"form_{cat}"):
                    t = st.text_input("테마", "기본")
                    c = st.text_input("내용")
                    i = st.select_slider("중요도", ['하', '중', '상'], '중')
                    if st.form_submit_button("추가"):
                        current_items.append({'id': f"new_{cat}_{datetime.now().timestamp()}", 'theme': t, 'content': c, 'importance': i})
                        st.rerun()

            # 리스트 렌더링
            themes = sorted(list(set(i['theme'] for i in current_items)))
            for theme in themes:
                st.markdown(f"<div class='theme-header'><h3>📍 {theme}</h3></div>", unsafe_allow_html=True)
                for item in [i for i in current_items if i['theme'] == theme]:
                    c1, c2 = st.columns([9, 1])
                    with c1:
                        st.checkbox(f"{get_color(item['importance'])} {item['content']}", key=item['id'])
                    with c2:
                        if edit_mode and st.button("🗑️", key=f"del_{item['id']}"):
                            current_items.remove(item); st.rerun()
