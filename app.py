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

# 입력한 문장 안에 시트 키워드가 들어있는지 똑똑하게 매칭하는 로더
def load_excel_template(keyword):
    file_name = 'Checkmaster.xlsx'
    if not os.path.exists(file_name):
        return []
    try:
        excel_file = pd.ExcelFile(file_name, engine='openpyxl')
        target_sheet = None
        user_input = keyword.strip()
        
        # 1차 검사: 엑셀 시트명(예: 시험)이 유저 입력 글자(예: 전공시험준비)에 들어있는지 확인
        for sheet in excel_file.sheet_names:
            sheet_name_clean = sheet.strip()
            if sheet_name_clean in user_input:
                target_sheet = sheet
                break
                
        # 2차 검사: 반대 방향 검사
        if not target_sheet:
            for sheet in excel_file.sheet_names:
                sheet_name_clean = sheet.strip()
                if user_input in sheet_name_clean:
                    target_sheet = sheet
                    break
                    
        # 아무것도 매칭 안 되면 첫 번째 시트를 디폴트로 안전하게 로드
        if not target_sheet and len(excel_file.sheet_names) > 0:
            target_sheet = excel_file.sheet_names[0]
            
        if not target_sheet: 
            return [] 
        
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
                        # 초기 기본 마감일은 10일 뒤로 지정하되, 알림창은 3일 전부터만 뜨게 로직 제어
                        st.session_state.cat_dates[cat] = date.today() + timedelta(days=10)
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
            themes = sorted(list(set(item['theme'] for item in current_items)))
            today = date.today()
            
            if cat in st.session_state.cat_dates:
                target_date = st.session_state.cat_dates[cat]
                d_day = (target_date - today).days 
                
                # [수정 요구사항 반영] 마감 임박 알림은 '3일 전(D-3)'부터 '당일(D-Day)'까지만 출력하도록 엄격히 제어
                if 0 <= d_day <= 3:
                    d_day_text = f"<b>{d_day}일</b> 남았습니다. (D-{d_day})" if d_day > 0 else "<b>오늘이 마감일입니다! (D-Day)</b>"
                    done_count = sum(1 for i in current_items if st.session_state.get(i['id'], False))
                    total_count = len(current_items)
                    prog_val = (done_count / total_count * 100) if total_count > 0 else 0
                    st.markdown(f'<div class="alert-box"><h3 style="margin-top:0;">⚠️ D-day 임박 알림!</h3><strong>{cat}</strong> 리스트의 마감이 {d_day_text} (목표일: {target_date})<br>현재 전체 진행률은 <b>{prog_val:.1f}%</b> 입니다.</div>', unsafe_allow_html=True)
            
            col_t1, col_t2 = st.columns([3, 1])
            with col_t1: st.title(f"{cat.upper()} LIST")
            with col_t2: show_preview = st.toggle("🔍 인쇄 미리보기 모드", key=f"prev_{cat}")
            
            if show_preview:
                preview_html = f"<div class='preview-paper'><h1>📋 {cat.upper()} 체크리스트</h1>"
                for theme in themes:
                    preview_html += f"<div class='preview-theme'>📍 {theme}</div>"
                    for item in [i for i in current_items if i['theme'] == theme]:
                        preview_html += f"<div class='preview-item'>□ {get_color(item['importance'])} {item['content']}</div>"
                preview_html += "</div>"
                st.markdown(preview_html, unsafe_allow_html=True)
            else:
                if themes:
                    nav_html = '<div class="nav-bar">'
                    for theme in themes: nav_html += f'<a href="#{sanitize_id(theme)}" class="nav-item"># {theme}</a>'
                    nav_html += '</div>'
                    st.markdown(nav_html, unsafe_allow_html=True)
                    
                done = sum(1 for item in current_items if st.session_state.get(item['id'], False))
                total = len(current_items)
                col_p1, col_p2 = st.columns([2, 1])
                with col_p1:
                    if total > 0:
                        st.write(f"🏗️ **전체 진행률: {done/total*100:.1f}%** ({done}/{total})")
                        st.progress(done/total)
                with col_p2:
                    saved_date = st.session_state.cat_dates.get(cat, date.today())
                    st.session_state.cat_dates[cat] = st.date_input(f"📅 {cat} 총 마감일", value=saved_date, key=f"cat_date_{cat}")
                    
                st.divider()
                c_opt, c_add = st.columns([1, 2])
                with c_opt: edit_mode = st.toggle("🛠️ 편집 모드 (수정/삭제 가능)", key=f"ed_{cat}")
                with c_add:
                    with st.expander("➕ 새로운 항목 추가하기"):
                        with st.form(key=f"form_{cat}", clear_on_submit=True):
                            add_theme = st.text_input("소분류 테마 입력", placeholder="기본")
                            new_content = st.text_input("체크리스트 세부 내용")
                            new_imp = st.select_slider("중요도", options=['하', '중', '상'], value='중')
                            if st.form_submit_button("추가") and new_content:
                                st.session_state.items_data[cat].append({'id': f"new_{sanitize_id(cat)}_{datetime.now().timestamp()}", 'theme': add_theme.strip() if add_theme.strip() else "기본", 'content': new_content, 'importance': new_imp})
                                st.rerun()
                                
                for theme in themes:
                    st.markdown(f'<div id="{sanitize_id(theme)}"></div>', unsafe_allow_html=True)
                    st.markdown(f"<div class='theme-header'><h2>📍 {theme}</h2></div>", unsafe_allow_html=True)
                    cols = st.columns(2)
                    for i, item in enumerate([i for i in current_items if i['theme'] == theme]):
                        with cols[i % 2]:
                            c_box, c_del = st.columns([7, 1])
                            with c_box: st.checkbox(f"{get_color(item['importance'])} {item['content']}", key=item['id'])
                            with c_del:
                                if edit_mode and st.button("🗑️", key=f"del_{item['id']}"):
                                    st.session_state.items_data[cat] = [x for x in st.session_state.items_data[cat] if x['id'] != item['id']]
                                    st.rerun()
