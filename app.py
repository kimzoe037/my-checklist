import streamlit as st
import pandas as pd
from datetime import datetime, date
import re

# 1. 페이지 설정 및 디자인
st.set_page_config(page_title="민채의 체크 마스터", layout="wide")

def apply_custom_style():
    st.markdown("""
        <style>
        .main { background-color: #F5F5DC; }
        html, body, [class*="st-"] { font-size: 18px; }
        h1 { font-family: 'Pretendard', sans-serif; font-weight: 800; color: #2c3e50; font-size: 2.5rem !important; }
        
        /* 네비게이션 바 */
        .nav-bar { background-color: rgba(93, 64, 55, 0.1); padding: 15px; border-radius: 10px; margin-bottom: 20px; display: flex; flex-wrap: wrap; gap: 10px; border: 1px solid #d7ccc8; }
        .nav-item { padding: 5px 15px; background-color: #5d4037; color: white !important; border-radius: 20px; text-decoration: none; font-size: 16px; font-weight: bold; }

        .theme-header { margin-top: 60px; margin-bottom: 25px; padding-bottom: 12px; border-bottom: 3px solid #5d4037; color: #3e2723; }
        
        div[data-testid="stCheckbox"] { padding: 15px 20px; margin-bottom: 12px; background-color: rgba(255, 255, 255, 0.7); border-radius: 8px; border: 1px solid #d7ccc8; }
        div[data-testid="stCheckbox"] label p { font-size: 20px !important; font-weight: 500; }
        
        /* 알림창 스타일 */
        .alert-box { background-color: #fff3cd; border-left: 5px solid #ffca28; padding: 20px; border-radius: 5px; margin-bottom: 20px; }
        
        /* 인쇄 미리보기 전용 스타일 */
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

# 3. 메인 앱 로직
try:
    file_name = 'Checkmaster.xlsx'
    excel_file = pd.ExcelFile(file_name)
    sheet_names = excel_file.sheet_names

    # 세션 데이터 초기화
    if 'items_data' not in st.session_state:
        st.session_state.items_data = {}
        st.session_state.cat_dates = {} # { "카테고리명": date_obj }
        for sheet in sheet_names:
            df = pd.read_excel(file_name, sheet_name=sheet, header=None)
            category_items = []
            for i, row in df.iterrows():
                content = str(row[2]).strip()
                if not content or content in ['체크 항목', '내용', 'nan', 'None']: continue
                category_items.append({
                    'id': f"{sheet}_{i}_{datetime.now().timestamp()}",
                    'theme': row[1] if pd.notna(row[1]) else "기본",
                    'content': content,
                    'importance': row[3] if pd.notna(row[3]) else '중'
                })
            st.session_state.items_data[sheet] = category_items

    tabs = st.tabs(sheet_names)
    
    for idx, tab in enumerate(tabs):
        with tab:
            cat = sheet_names[idx]
            current_items = st.session_state.items_data.get(cat, [])
            themes = sorted(list(set(item['theme'] for item in current_items)))
            
            # --- [기능] 마감 2일 전 알림창 ---
            today = date.today()
            if cat in st.session_state.cat_dates:
                target_date = st.session_state.cat_dates[cat]
                d_day = (target_date - today).days
                if 0 <= d_day <= 2:
                    done_count = sum(1 for i in current_items if st.session_state.get(i['id'], False))
                    total_count = len(current_items)
                    prog_val = (done_count / total_count * 100) if total_count > 0 else 0
                    
                    st.markdown(f"""
                        <div class="alert-box">
                            <h3 style="margin-top:0;">⚠️ 마감 임박 알림!</h3>
                            <strong>{cat}</strong> 리스트의 마감이 <b>{d_day}일</b> 남았습니다.<br>
                            현재 전체 진행률은 <b>{prog_val:.1f}%</b> 입니다. 조금만 더 힘내세요!
                        </div>
                    """, unsafe_allow_html=True)

            # 상단 제목 및 미리보기 토글
            col_t1, col_t2 = st.columns([3, 1])
            with col_t1:
                st.title(f" {cat.upper()} LIST")
            with col_t2:
                show_preview = st.toggle("🔍 인쇄 미리보기 모드", key=f"prev_{cat}")

            if show_preview:
                st.info("💡 Ctrl+P를 눌러 인쇄하세요.")
                preview_html = f"<div class='preview-paper'><h1>📋 {cat.upper()} 체크리스트</h1>"
                if cat in st.session_state.cat_dates:
                    preview_html += f"<p><b>총 마감 기한: {st.session_state.cat_dates[cat]}</b></p>"
                for theme in themes:
                    preview_html += f"<div class='preview-theme'>📍 {theme}</div>"
                    t_items = [i for i in current_items if i['theme'] == theme]
                    for item in t_items:
                        preview_html += f"<div class='preview-item'>□ {get_color(item['importance'])} {item['content']}</div>"
                preview_html += "</div>"
                st.markdown(preview_html, unsafe_allow_html=True)
                st.stop()

            # 퀵 네비게이션
            if themes:
                nav_html = '<div class="nav-bar">'
                for theme in themes:
                    nav_html += f'<a href="#{sanitize_id(theme)}" class="nav-item"># {theme}</a>'
                nav_html += '</div>'
                st.markdown(nav_html, unsafe_allow_html=True)

            # --- [수정] 전체 공정률 라인에 마감일 설정 추가 ---
            done = sum(1 for item in current_items if st.session_state.get(item['id'], False))
            total = len(current_items)
            
            col_p1, col_p2 = st.columns([2, 1])
            with col_p1:
                if total > 0:
                    prog = done / total
                    st.write(f"🏗️ **전체 공정률: {prog*100:.1f}%** ({done}/{total})")
                    st.progress(prog)
            with col_p2:
                # 카테고리 전체에 대한 단 하나의 마감 기한 설정
                saved_cat_date = st.session_state.cat_dates.get(cat, date.today())
                st.session_state.cat_dates[cat] = st.date_input(f"📅 {cat} 총 마감일", value=saved_cat_date, key=f"cat_date_{cat}")

            st.divider()

            # 편집 모드 및 항목 추가
            c_opt, c_add = st.columns([1, 2])
            with c_opt:
                edit_mode = st.toggle("🛠️ 편집 모드 (수정/삭제 가능)", key=f"ed_{cat}")
            with c_add:
                with st.expander("➕ 새로운 항목 추가하기"):
                    with st.form(key=f"form_{cat}", clear_on_submit=True):
                        add_theme = st.selectbox("테마 선택", themes if themes else ["기본"], key=f"at_{cat}")
                        new_content = st.text_input("내용")
                        new_imp = st.select_slider("중요도", options=['하', '중', '상'], value='중')
                        if st.form_submit_button("추가") and new_content:
                            st.session_state.items_data[cat].append({
                                'id': f"new_{datetime.now().timestamp()}",
                                'theme': add_theme, 'content': new_content, 'importance': new_imp
                            })
                            st.rerun()

            # 메인 리스트
            for theme in themes:
                st.markdown(f'<div id="{sanitize_id(theme)}"></div>', unsafe_allow_html=True)
                st.markdown(f"<div class='theme-header'><h2>📍 {theme}</h2></div>", unsafe_allow_html=True)
                
                t_items = [i for i in current_items if i['theme'] == theme]
                t_items.sort(key=lambda x: st.session_state.get(x['id'], False))
                
                cols = st.columns(2)
                for i, item in enumerate(t_items):
                    with cols[i % 2]:
                        c_box, c_del = st.columns([7, 1])
                        with c_box:
                            st.checkbox(f"{get_color(item['importance'])} {item['content']}", key=item['id'])
                        with c_del:
                            if edit_mode:
                                if st.button("🗑️", key=f"del_{item['id']}"):
                                    st.session_state.items_data[cat] = [x for x in st.session_state.items_data[cat] if x['id'] != item['id']]
                                    st.rerun()

except Exception as e:
    st.error(f"오류가 발생했습니다: {e}")
