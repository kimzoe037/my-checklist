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
        .alert-box { background-color: #fff3cd; border-left: 5px solid #ffca28; padding: 20px; border-radius: 5px; margin-bottom: 20px; }
        </style>
    """, unsafe_allow_html=True)

apply_custom_style()

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
        excel_file = pd.ExcelFile(file_name, engine='openpyxl')
        target_sheet = None
        user_input = keyword.strip()
        
        # 시트 매칭 로직
        for sheet in excel_file.sheet_names:
            sheet_name_clean = sheet.strip()
            if sheet_name_clean in user_input or user_input in sheet_name_clean:
                target_sheet = sheet
                break
        
        if not target_sheet:
            for k in ["시공", "여행", "시험"]:
                if k in user_input:
                    for sheet in excel_file.sheet_names:
                        if k in sheet: target_sheet = sheet; break
        
        if not target_sheet and excel_file.sheet_names: target_sheet = excel_file.sheet_names[0]
        
        df = pd.read_excel(file_name, sheet_name=target_sheet, header=None, engine='openpyxl')
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
    except Exception: return []

if 'step' not in st.session_state: st.session_state.step = 1  
if 'items_data' not in st.session_state: st.session_state.items_data = {}  
if 'cat_dates' not in st.session_state: st.session_state.cat_dates = {}  

if st.session_state.step == 1:
    st.markdown("<div class='step1-container'>", unsafe_allow_html=True)
    if os.path.exists("character.png"): st.image("character.png", width=200)
    st.title("Checklist")
    with st.form(key="step1_form"):
        categories_input = st.text_input("카테고리 입력 (예: 상하이 여행준비, 전공시험준비)")
        if st.form_submit_button("시작하기"):
            for cat in [c.strip() for c in categories_input.split(",") if c.strip()]:
                st.session_state.cat_dates[cat] = date.today() + timedelta(days=10)
                st.session_state.items_data[cat] = load_excel_template(cat)
            st.session_state.step = 2
            st.rerun()

elif st.session_state.step == 2:
    if st.button("초기화"): st.session_state.clear(); st.session_state.step = 1; st.rerun()
    tabs = st.tabs(list(st.session_state.items_data.keys()))
    for idx, tab in enumerate(tabs):
        with tab:
            cat = list(st.session_state.items_data.keys())[idx]
            items = st.session_state.items_data.get(cat, [])
            total = len(items)
            done = sum(1 for item in items if st.session_state.get(item['id'], False))
            
            # 안전한 진행률 계산
            prog = (done / total) if total > 0 else 0
            st.write(f"🏗️ **전체 진행률: {prog*100:.1f}%** ({done}/{total})")
            st.progress(prog)
            
            d_day = (st.session_state.cat_dates.get(cat, date.today()) - date.today()).days
            if 0 <= d_day <= 3:
                st.markdown(f'<div class="alert-box">⚠️ <b>D-{d_day if d_day > 0 else "Day"}!</b> 마감이 얼마 남지 않았습니다!</div>', unsafe_allow_html=True)
            
            edit_mode = st.toggle("🛠️ 편집 모드", key=f"ed_{cat}")
            themes = sorted(list(set(item['theme'] for item in items)))
            for theme in themes:
                st.markdown(f"### 📍 {theme}")
                for item in [i for i in items if i['theme'] == theme]:
                    c1, c2 = st.columns([10, 1])
                    with c1: st.checkbox(f"{get_color(item['importance'])} {item['content']}", key=item['id'])
                    with c2:
                        if edit_mode and st.button("🗑️", key=f"del_{item['id']}"):
                            st.session_state.items_data[cat].remove(item); st.rerun()
