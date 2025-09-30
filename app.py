import streamlit as st
from datetime import datetime, timedelta
import pandas as pd
import os
import glob
import base64

# 페이지 설정
st.set_page_config(
    page_title="통합 식단 관리 시스템",
    page_icon="🍽️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 커스텀 CSS
st.markdown("""
<style>
/* 메인 배경 */
.main {
    padding: 2rem 3rem;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
}
.card {
    background: white;
    padding: 2rem;
    border-radius: 15px;
    box-shadow: 0 10px 30px rgba(0,0,0,0.1);
    margin-bottom: 2rem;
    border: 1px solid rgba(255,255,255,0.2);
}
.login-card {
    max-width: 400px;
    margin: 5rem auto;
    background: white;
    padding: 3rem;
    border-radius: 20px;
    box-shadow: 0 15px 35px rgba(0,0,0,0.1);
    text-align: center;
}
.success-banner {
    background: linear-gradient(90deg, #56ab2f, #a8e6cf);
    color: white;
    padding: 1rem;
    border-radius: 10px;
    text-align: center;
    font-weight: 600;
    margin-bottom: 2rem;
}
.admin-header {
    background: linear-gradient(90deg, #667eea, #764ba2);
    color: white;
    padding: 1.5rem;
    border-radius: 10px;
    text-align: center;
    margin-bottom: 2rem;
}
.user-header {
    background: linear-gradient(90deg, #4facfe, #00f2fe);
    color: white;
    padding: 1.5rem;
    border-radius: 10px;
    text-align: center;
    margin-bottom: 2rem;
}
.stButton > button {
    background: linear-gradient(45deg, #667eea, #764ba2);
    color: white;
    border: none;
    padding: 0.75rem 2rem;
    border-radius: 25px;
    font-weight: 600;
    transition: all 0.3s ease;
    box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
}
.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(102, 126, 234, 0.4);
}
.danger-button button {
    background: linear-gradient(45deg, #ff416c, #ff4b2b) !important;
    box-shadow: 0 4px 15px rgba(255, 65, 108, 0.3) !important;
}
.start-button button {
    background: linear-gradient(45deg, #56ab2f, #a8e6cf) !important;
    box-shadow: 0 4px 15px rgba(86, 171, 47, 0.3) !important;
    font-size: 1.1rem !important;
    padding: 1rem 2.5rem !important;
}
.stat-card {
    background: white;
    padding: 1.5rem;
    border-radius: 10px;
    text-align: center;
    box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    margin: 0.5rem;
}
.stat-number {
    font-size: 2rem;
    font-weight: 700;
    color: #667eea;
}
.stat-label {
    color: #7f8c8d;
    font-size: 0.9rem;
    margin-top: 0.5rem;
}
/* 사이드바 스타일 */
.css-1d391kg {
    background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
}
section[data-testid="stSidebar"] > div {
    background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
}
</style>
""", unsafe_allow_html=True)

LOG_FILE = "log.csv"
UPLOAD_FOLDER = "uploads"
HTML_FILE = "index.html"
MENU_XLSX = "menu.xlsx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 사용자 설정
user_dict = {
    "SR01": "test01", "SR02": "test02", "SR03": "test03", "SR04": "test04",
    "SR05": "test05", "SR06": "test06", "SR07": "test07", "SR08": "test08",
    "SR09": "test09", "SR10": "test10", "SR11": "test11", "SR12": "test12",
    "SR13": "test13", "admin": "admin"
}

def get_kst_now():
    return datetime.utcnow() + timedelta(hours=9)

# 초기 상태
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.start_time = None
if "meal_type" not in st.session_state:
    st.session_state.meal_type = "식단표A"

# HTML 파일을 읽어서 iframe으로 표시하는 함수
def render_index_html_with_injected_xlsx():
    html_path = "index.html"  # 지금 쓰시는 경로
    if not os.path.exists(html_path):
        st.error("⚠️ index.html 파일을 찾을 수 없습니다. 파일을 같은 디렉토리에 배치해주세요.")
        return

    # 1) HTML 읽기
    try:
        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
    except:
        with open(html_path, 'r', encoding='cp949', errors='ignore') as f:
            html_content = f.read()

    # 2) Excel 후보 경로 중 존재하는 것 선택
    xlsx_candidates = [
        "menu.xlsx",
        "/mnt/data/menu.xlsx",
        "/mnt/data/정선_음식 데이터_간식제외.xlsx"
    ]
    xlsx_path = next((p for p in xlsx_candidates if os.path.exists(p)), None)

    # 3) base64로 주입 스크립트 생성
    inject_script = ""
    if xlsx_path:
        with open(xlsx_path, "rb") as xf:
            b64 = base64.b64encode(xf.read()).decode()
        inject_script = f"<script>window.__XLSX_BASE64__='{b64}';</script>"
    else:
        # 주입이 없으면 null로 선언 (HTML이 fetch()로 폴백)
        inject_script = "<script>window.__XLSX_BASE64__=null;</script>"

    # 4) 주입 스크립트 + 기존 HTML을 합쳐서 iframe 렌더
    final_html = inject_script + html_content
    st.components.v1.html(final_html, height=900, scrolling=True)

# 사이드바 - 탭 선택
with st.sidebar:
    st.markdown("""
    <div style='text-align: center; padding: 2rem 0;'>
        <h1 style='color: white; margin-bottom: 0.5rem;'>🍽️</h1>
        <h2 style='color: white; font-size: 1.5rem;'>통합 식단 시스템</h2>
    </div>
    """, unsafe_allow_html=True)
    
    if st.session_state.logged_in:
        st.markdown(f"""
        <div style='background: rgba(255,255,255,0.2); padding: 1rem; border-radius: 10px; text-align: center; color: white; margin-bottom: 2rem;'>
            <strong>👤 {st.session_state.username}</strong>
        </div>
        """, unsafe_allow_html=True)
    
    selected_tab = st.radio(
        "메뉴 선택",
        ["📝 식단 제출", "🔍 메뉴 관리"],
        label_visibility="collapsed"
    )
    
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    if st.session_state.logged_in:
        if st.button("🚪 로그아웃", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.username = ""
            st.session_state.start_time = None
            st.session_state.meal_type = "식단표A"
            st.rerun()

# 로그인 화면
if not st.session_state.logged_in:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
    <div class="login-card">
        <h1 style="color: #667eea; margin-bottom: 0.5rem;">🍽️</h1>
        <h2 style="color: #2c3e50; margin-bottom: 0.5rem;">식단 설계 시스템</h2>
        <p style="color: #7f8c8d; margin-bottom: 2rem;">로그인하여 시작하세요</p>
    </div>
    """, unsafe_allow_html=True)
    
    with st.container():
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            username = st.text_input("👤 아이디", placeholder="아이디를 입력하세요")
            password = st.text_input("🔒 비밀번호", type="password", placeholder="비밀번호를 입력하세요")
            
            if st.button("🚀 로그인", use_container_width=True):
                if username in user_dict and user_dict[username] == password:
                    st.session_state.logged_in = True
                    st.session_state.username = username
                    st.rerun()
                else:
                    st.error("❌ 아이디 또는 비밀번호가 올바르지 않습니다.")

# 로그인 후 화면
else:
    st.markdown(f"""
    <div class="success-banner">
        🎉 {st.session_state.username}님 환영합니다!
    </div>
    """, unsafe_allow_html=True)
    
    # 탭 1: 식단 제출
    if selected_tab == "📝 식단 제출":
        # 관리자 페이지
        if st.session_state.username == "admin":
            st.markdown("""
            <div class="admin-header">
                <h1>🔧 관리자 페이지</h1>
                <p>시스템 관리 및 제출 기록 확인</p>
            </div>
            """, unsafe_allow_html=True)
            
            # 통계 카드들
            if os.path.exists(LOG_FILE):
                df = pd.read_csv(LOG_FILE)
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.markdown(f"""<div class="stat-card"><div class="stat-number">{len(df)}</div><div class="stat-label">총 제출 수</div></div>""", unsafe_allow_html=True)
                with col2:
                    st.markdown(f"""<div class="stat-card"><div class="stat-number">{df['사용자'].nunique()}</div><div class="stat-label">참여 사용자</div></div>""", unsafe_allow_html=True)
                with col3:
                    avg_time = int(df['소요시간(초)'].mean()) if '소요시간(초)' in df.columns else 0
                    st.markdown(f"""<div class="stat-card"><div class="stat-number">{avg_time}초</div><div class="stat-label">평균 소요시간</div></div>""", unsafe_allow_html=True)
                with col4:
                    today_str = datetime.now().strftime('%Y-%m-%d')
                    today_count = len(df[df['제출시간'].astype(str).str.contains(today_str)])
                    st.markdown(f"""<div class="stat-card"><div class="stat-number">{today_count}</div><div class="stat-label">오늘 제출</div></div>""", unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # 관리 버튼들
                col1, col2, col3 = st.columns([1, 1, 2])
                with col1:
                    st.markdown('<div class="danger-button">', unsafe_allow_html=True)
                    if st.button("🗑️ 기록 전체 삭제", use_container_width=True):
                        if os.path.exists(LOG_FILE):
                            os.remove(LOG_FILE)
                            st.success("✅ 로그 파일이 삭제되었습니다.")
                        else:
                            st.warning("⚠️ 삭제할 로그 파일이 없습니다.")
                    st.markdown('</div>', unsafe_allow_html=True)
                
                # 제출 기록 테이블
                st.markdown("""<div class="card"><h3>📊 제출 기록</h3></div>""", unsafe_allow_html=True)
                st.dataframe(df, use_container_width=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                col1, col2 = st.columns(2)
                
                with col1:
                    user_list = df["사용자"].unique().tolist()
                    selected_user = st.selectbox("👤 사용자 선택", user_list)
                
                with col2:
                    pattern = os.path.join(UPLOAD_FOLDER, f"{selected_user}_식단표*.xlsx")
                    files = sorted(glob.glob(pattern))
                    if files:
                        for path in files:
                            base = os.path.basename(path)
                            label = f"📥 {os.path.splitext(base)[0]} 다운로드"
                            with open(path, "rb") as f:
                                st.download_button(
                                    label=label,
                                    data=f,
                                    file_name=base,
                                    use_container_width=True
                                )
                    else:
                        st.warning(f"⚠️ {selected_user}님의 제출 파일이 존재하지 않습니다.")
            else:
                st.info("📝 제출 기록이 아직 없습니다.")
        
        # 사용자 페이지
        else:
            # 식단표 선택
            st.markdown("""
            <div class="card">
                <h3>🧾 식단표 선택</h3>
                <p>작업하실 식단표를 먼저 선택해주세요.</p>
            </div>
            """, unsafe_allow_html=True)
            
            st.session_state.meal_type = st.radio(
                "식단표 유형",
                options=["식단표A", "식단표B"],
                index=0 if st.session_state.meal_type == "식단표A" else 1,
                horizontal=True
            )
            
            # 시작 버튼 섹션
            st.markdown("""
            <div class="card">
                <h3>🚀 작업 시작</h3>
                <p>아래 버튼을 클릭하여 식단 개선 작업을 시작하세요.</p>
            </div>
            """, unsafe_allow_html=True)
            
            if st.session_state.start_time is None:
                col1, col2 = st.columns([1, 1])
                with col1:
                    st.markdown('<div class="start-button">', unsafe_allow_html=True)
                    if st.button("🍽️ 식단 설계 시작", use_container_width=True):
                        st.session_state.start_time = get_kst_now()
                        st.success(f"⏰ 시작 시간: {st.session_state.start_time.strftime('%H:%M:%S')}")
                        st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)
            else:
                # 진행 중 상태 표시
                current_time = get_kst_now()
                elapsed = current_time - st.session_state.start_time
                elapsed_seconds = int(elapsed.total_seconds())
                
                st.markdown(f"""
                <div style="background: linear-gradient(90deg, #56ab2f, #a8e6cf); color: white; padding: 1rem; border-radius: 10px; text-align: center; margin-bottom: 2rem;">
                    ⏱️ 작업 진행 중... | 시작 시간: {st.session_state.start_time.strftime('%H:%M:%S')} | 경과 시간: {elapsed_seconds}초 | 선택: {st.session_state.meal_type}
                </div>
                """, unsafe_allow_html=True)
                
                # 파일 업로드 섹션
                st.markdown("""
                <div class="card">
                    <h3>📁 파일 업로드</h3>
                    <p>완성된 식단 설계 엑셀 파일을 업로드해주세요.</p>
                </div>
                """, unsafe_allow_html=True)
                
                uploaded_file = st.file_uploader(
                    "📊 엑셀 파일 선택",
                    type=["xlsx", "xls"],
                    help="xlsx 또는 xls 파일만 업로드 가능합니다."
                )
                
                if uploaded_file:
                    st.success(f"✅ 파일 선택됨: {uploaded_file.name}")
                    
                    col1, col2, col3 = st.columns([1, 2, 1])
                    with col2:
                        if st.button("📤 제출하기", use_container_width=True):
                            submit_time = get_kst_now()
                            duration = (submit_time - st.session_state.start_time).total_seconds()
                            
                            safe_meal = st.session_state.meal_type
                            save_name = f"{st.session_state.username}_{safe_meal}.xlsx"
                            file_path = os.path.join(UPLOAD_FOLDER, save_name)
                            
                            with open(file_path, "wb") as f:
                                f.write(uploaded_file.getbuffer())
                            
                            # 로그 데이터 구성
                            log_row = {
                                "사용자": st.session_state.username,
                                "시작시간": st.session_state.start_time.strftime('%Y-%m-%d %H:%M:%S'),
                                "제출시간": submit_time.strftime('%Y-%m-%d %H:%M:%S'),
                                "소요시간(초)": int(duration),
                                "식단표종류": safe_meal,
                                "파일경로": file_path
                            }
                            
                            if os.path.exists(LOG_FILE):
                                existing = pd.read_csv(LOG_FILE)
                                for col in ["파일경로", "식단표종류"]:
                                    if col not in existing.columns:
                                        existing[col] = None
                                log_df = pd.concat([existing, pd.DataFrame([log_row])], ignore_index=True)
                            else:
                                log_df = pd.DataFrame([log_row])
                            
                            log_df.to_csv(LOG_FILE, index=False)
                            
                            st.success("🎉 제출이 완료되었습니다!")
                            st.markdown(f"""
                            <div style="background: #e8f5e8; padding: 1.5rem; border-radius: 10px; margin: 1rem 0;">
                                <h4>📋 제출 완료 요약</h4>
                                <p><strong>👤 사용자:</strong> {st.session_state.username}</p>
                                <p><strong>🧾 식단표:</strong> {safe_meal}</p>
                                <p><strong>⏰ 소요 시간:</strong> {int(duration)}초</p>
                                <p><strong>📅 제출 시간:</strong> {submit_time.strftime('%Y-%m-%d %H:%M:%S')}</p>
                                <p><strong>💾 저장 파일명:</strong> {save_name}</p>
                            </div>
                            """, unsafe_allow_html=True)
                            
                            st.session_state.start_time = None
    
    # 탭 2: 메뉴 관리
    elif selected_tab == "🔍 메뉴 관리":
        st.markdown("""
        <div class="user-header">
            <h1>🔍 메뉴 관리</h1>
            <p>메뉴 데이터베이스 조회 및 검색</p>
        </div>
        """, unsafe_allow_html=True)
        
        render_index_html_with_injected_xlsx()
