import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import os
import re

# 1. 페이지 설정
st.set_page_config(page_title="Checkmaster", layout="wide")

# 2. 스타일 정의 (스크린샷 UI 구현)
st.markdown("""
    <style>
    .nav-bar { background-color: #f0f0f0; padding: 15px; border-radius: 10px; margin-bottom: 20px; display: flex; flex-wrap: wrap; gap: 10px; border: 1px solid #ddd; }
    .nav-item { padding: 5px 15px; background-color: #5d4037; color: white !important; border-radius: 20px; font-weight: bold; }
    .theme-header { margin-top: 30px; border-bottom: 2px solid #5d4037; color: #3e2723; }
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
        target_sheet = next((s for s in excel_file.sheet_names if keyword in s), None)
        if not target_sheet: return []
        df = pd.read_excel(file_name, sheet_name=target_sheet, header=None)
        items = []
        for i, row in df.iterrows():
            if len(row) < 3: continue
            items.append({
                'id': f"{sanitize_id(str(row[2]))}_{i}_{datetime.now().timestamp()}", 
                'theme': str(row[1]) if pd.notna(row[1]) else "기본",
                'content': str(row[2]).strip(),
                'importance': str(row[3]) if pd.notna(row[3]) else '중'
            })
        return items
    except: return []

# 4. 세션 초기화
if 'items_data' not in st.session_state: st.session_state.items_data = {}
if 'cat_dates' not in st.session_state: st.session_state.cat_dates = {}

# 5. 메인 로직
st.title("📋 Checkmaster")
input_text = st.text_input("프로젝트 입력 (예: 건축시공, 여행준비, 시험준비)", placeholder="건축시공, 여행준비, 시험준비")

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
            
            # [상단 레이아웃]
            col_h1, col_h2 = st.columns([3, 1])
            with col_h1:
                st.header(f"{cat.upper()} LIST")
            with col_h2:
                show_preview = st.toggle("🔍 인쇄 미리보기 모드")
            
            # [태그 바]
            themes = sorted(list(set(i['theme'] for i in current_items)))
            st.markdown('<div class="nav-bar">' + "".join([f'<span class="nav-item"># {t}</span>' for t in themes]) + '</div>', unsafe_allow_html=True)
            
            # [진행률 및 마감일]
            c1, c2 = st.columns([2, 1])
            done = sum(1 for i in current_items if st.session_state.get(i['id'], False))
            with c1:
                st.write(f"🏗️ 전체 진행률: {done}/{len(current_items)} ({(done/len(current_items)*100) if current_items else 0:.1f}%)")
                st.progress((done/len(current_items)) if current_items else 0)
            with c2:
                st.session_state.cat_dates[cat] = st.date_input(f"📅 {cat} 마감일", st.session_state.cat_dates[cat])

            # [편집 및 추가]
            edit_mode = st.toggle("🛠️ 편집 모드 (수정/삭제)", key=f"ed_{cat}")
            with st.expander("➕ 새로운 항목 추가하기"):
                with st.form(f"form_{cat}", clear_on_submit=True):
                    t = st.text_input("테마", "기본")
                    c = st.text_input("내용")
                    i = st.select_slider("중요도", ['하', '중', '상'], '중')
                    if st.form_submit_button("추가"):
                        current_items.append({'id': f"new_{cat}_{datetime.now().timestamp()}", 'theme': t, 'content': c, 'importance': i})
                        st.rerun()

            # [리스트 출력]
            if show_preview:
                st.info("💡 Ctrl+P를 눌러 인쇄하세요.")
            
            for theme in themes:
                st.markdown(f"<div class='theme-header'><h3>📍 {theme}</h3></div>", unsafe_allow_html=True)
                theme_items = [i for i in current_items if i['theme'] == theme]
                for item in theme_items:
                    c_box, c_del = st.columns([9, 1])
                    with c_box:
                        st.checkbox(f"{get_color(item['importance'])} {item['content']}", key=item['id'])
                    with c_del:
                        if edit_mode and st.button("🗑️", key=f"del_{item['id']}"):
                            current_items.remove(item); st.rerun()