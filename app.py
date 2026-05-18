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
        
        /* 첫 페이지 중앙 정렬 컨테이너 (테두리 및 배경 박스 완전 제거) */
        .step1-container { text-align: center; max-width: 600px; margin: 0 auto; padding-top: 20px; }
        
        /* 네비게이션 바 */
        .nav-bar { background-color: rgba(93, 64, 55, 0.1); padding: 15px; border-radius: 10px; margin-bottom: 20px; display: flex; flex-wrap: wrap; gap: 10px; border: 1px solid #d7ccc8; }
        .nav-item { padding: 5px 15px; background-color: #5d4037; color: white !important; border-radius: 20px; text-decoration: none; font-size: 16px; font-weight: bold; }

        .theme-header { margin-top: 40px; margin-bottom: 25px; padding-bottom: 12px; border-bottom: 3px solid #5d4037; color: #3e2723; }
        
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

# 엑셀에서 특정 키워드에 맞는 데이터를 찾아와 기본값 리스트를 만들어주는 함수
def load_excel_template(keyword):
    file_name = 'Checkmaster.xlsx'
    try:
        excel_file = pd.ExcelFile(file_name)
        target_sheet = None
        for sheet in excel_file.sheet_names:
            if keyword in sheet:
                target_sheet = sheet
                break
        
        if not target_sheet:
            if keyword in excel_file.sheet_names:
                target_sheet = keyword
            else:
                return [] 

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
        st.error(f"엑셀 템플릿을 불러오는 중 오류 발생: {e}")
        return []

# 3. 세션 상태 및 화면 전환 변수 초기화
if 'step' not in st.session_state:
    st.session_state.step = 1  
if 'items_data' not in st.session_state:
    st.session_state.items_data = {}  
if 'cat_dates' not in st.session_state:
    st.session_state.cat_dates = {}  

# ==========================================
# STEP 1: 첫 진입 화면 (무엇을 만들겠습니까?)
# ==========================================
if st.session_state.step == 1:
    # 중앙 정렬을 위한 컨테이너 시작
    st.markdown("<div class='step1-container'>", unsafe_allow_html=True)
    
    # [수정 완료] 흰색 원형 테두리를 완전히 없애고, 보내주신 곰돌이 이미지만 아담하게(width=200) 중앙 배치
    # 실행 환경에 있는 이미지 파일명에 맞게 "character.png"를 수정해서 쓰세요. (예: bear.png)
    image_path = "character.png" 
    if os.path.exists(image_path):
        # 대형 박스나 원형 테두리 없이 이미지만 깔끔하게 표출
        st.image(image_path, width=200)
    else:
        st.markdown("<h1 style='font-size: 60px; margin: 0;'>📋</h1>", unsafe_allow_html=True)
        
    st.title("Checklist")
    st.subheader("지금 어떤 프로젝트를 준비하고 계신가요?")
    st.write("나만의 맞춤형 체크리스트 대분류(카테고리)를 입력해 주세요.")
    st.markdown("</div>", unsafe_allow_html=True)
    
    with st.form(key="step1_form"):
        categories_input = st.text_input(
            "카테고리 대분류 입력 (쉼표 `,` 로 구분하여 여러 개 등록 가능)", 
            placeholder="예: 상하이 여행준비, 기말시험, 거실 시공, 자취방 이사"
        )
        submit_btn = st.form_submit_button("나만의 체크리스트 생성하기 🚀")
        
        if submit_btn:
            if categories_input.strip():
                input_list = [c.strip() for c in categories_input.split(",") if c.strip()]
                
                for cat in input_list:
                    if cat not in st.session_state.items_data:
                        # 초기 생성 시 기본 목표 날짜를 오늘로부터 10일 뒤로 세팅 (예: 오늘 18일 -> 기본값 28일)
                        st.session_state.cat_dates[cat] = date.today() + timedelta(days=10)
                        
                        if '여행' in cat:
                            st.session_state.items_data[cat] = load_excel_template('여행')
                        elif '시험' in cat:
                            st.session_state.items_data[cat] = load_excel_template('시험')
                        elif '시공' in cat:
                            st.session_state.items_data[cat] = load_excel_template('시공')
                        else:
                            st.session_state.items_data[cat] = []
                
                st.session_state.step = 2
                st.rerun()
            else:
                st.warning("⚠️ 최소 한 개 이상의 카테고리명을 입력해 주세요!")

# ==========================================
# STEP 2: 메인 체크리스트 화면
# ==========================================
elif st.session_state.step == 2:
    sheet_names = list(st.session_state.items_data.keys())
    
    if not sheet_names:
        st.warning("생성된 카테고리가 없습니다. 첫 화면으로 돌아갑니다.")
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
            
            # --- [수정 완료] 사용자가 설정한 목표 날짜에 맞춰 실시간으로 D-Day 계산 ---
            today = date.today()
            if cat in st.session_state.cat_dates:
                target_date = st.session_state.cat_dates[cat]
                d_day = (target_date - today).days # 설정일(예: 28일) - 오늘(예: 18일) = 10일
                
                # 남은 일수에 맞춰 D-day 텍스트 동적 변경 (D-2, D-10 완벽 적용)
                if d_day > 0:
                    d_day_text = f"<b>{d_day}일</b> 남았습니다. (D-{d_day})"
                elif d_day == 0:
                    d_day_text = "<b>오늘이 마감일입니다! (D-Day)</b>"
                else:
                    d_day_text = f"<b>{abs(d_day)}일 지났습니다. (D+{abs(d_day)})</b>"
                
                done_count = sum(1 for i in current_items if st.session_state.get(i['id'], False))
                total_count = len(current_items)
                prog_val = (done_count / total_count * 100) if total_count > 0 else 0
                
                st.markdown(f"""
                    <div class="alert-box">
                        <h3 style="margin-top:0;">⚠️ D-day 임박 알림!</h3>
                        <strong>{cat}</strong> 리스트의 마감이 {d_day_text} (목표일: {target_date})<br>
                        현재 전체 진행률은 <b>{prog_val:.1f}%</b> 입니다. 조금만 더 힘내세요!
                    </div>
                """, unsafe_allow_html=True)

            # 상단 제목 및 미리보기 토글
            col_t1, col_t2 = st.columns([3, 1])
            with col_t1:
                st.title(f" {cat.upper()} LIST")
            with col_t2:
                show_preview = st.toggle("🔍 인쇄 미리보기 모드", key=f"prev_{cat}")

            # 인쇄 미리보기 모드
            if show_preview:
                st.info("💡 Ctrl+P를 눌러 인쇄하세요. 미리보기를 끄려면 우측 상단 토글을 다시 누르세요.")
                preview_html = f"<div class='preview-paper'><h1>📋 {cat.upper()} 체크리스트</h1>"
                if cat in st.session_state.cat_dates:
                    preview_html += f"<p><b>총 마감 기한: {st.session_state.cat_dates[cat]}</b></p>"
                if not themes:
                    preview_html += "<p style='color:gray;'>등록된 항목이 없습니다. 먼저 항목을 추가해 주세요.</p>"
                for theme in themes:
                    preview_html += f"<div class='preview-theme'>📍 {theme}</div>"
                    t_items = [i for i in current_items if i['theme'] == theme]
                    for item in t_items:
                        preview_html += f"<div class='preview-item'>□ {get_color(item['importance'])} {item['content']}</div>"
                preview_html += "</div>"
                st.markdown(preview_html, unsafe_allow_html=True)
                
            else:
                # 일반 대시보드 모드
                if themes:
                    nav_html = '<div class="nav-bar">'
                    for theme in themes:
                        nav_html += f'<a href="#{sanitize_id(theme)}" class="nav-item"># {theme}</a>'
                    nav_html += '</div>'
                    st.markdown(nav_html, unsafe_allow_html=True)

                # 전체 공정률 및 마감일 설정
                done = sum(1 for item in current_items if st.session_state.get(item['id'], False))
                total = len(current_items)
                
                col_p1, col_p2 = st.columns([2, 1])
                with col_p1:
                    if total > 0:
                        prog = done / total
                        st.write(f"🏗️ **전체 진행률: {prog*100:.1f}%** ({done}/{total})")
                        st.progress(prog)
                    else:
                        st.info("💡 아래 ➕ 버튼을 눌러 소분류(테마)와 체크리스트 세부 내용을 추가해 보세요!")
                with col_p2:
                    # 설정값을 바꾸면 위의 D-day 문구가 실시간 연동되어 즉각 바뀝니다.
                    saved_date = st.session_state.cat_dates.get(cat, date.today())
                    st.session_state.cat_dates[cat] = st.date_input(f"📅 {cat} 총 마감일", value=saved_date, key=f"cat_date_{cat}")

                st.divider()

                # 편집 모드 및 항목 추가
                c_opt, c_add = st.columns([1, 2])
                with c_opt:
                    edit_mode = st.toggle("🛠️ 편집 모드 (수정/삭제 가능)", key=f"ed_{cat}")
                with c_add:
                    with st.expander("➕ 새로운 항목 추가하기", expanded=True if total==0 else False):
                        with st.form(key=f"form_{cat}", clear_on_submit=True):
                            add_theme = st.text_input("소분류 테마 입력 (예: 필수 준비물, 1교시 과목 등)", placeholder="기본")
                            if not add_theme.strip():
                                add_theme = "기본"
                            new_content = st.text_input("체크리스트 세부 내용")
                            new_imp = st.select_slider("중요도", options=['하', '중', '상'], value='중')
                            
                            if st.form_submit_button("추가") and new_content:
                                st.session_state.items_data[cat].append({
                                    'id': f"new_{sanitize_id(cat)}_{datetime.now().timestamp()}",
                                    'theme': add_theme.strip(), 'content': new_content, 'importance': new_imp
                                })
                                st.rerun()

                # 메인 리스트 렌더링
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