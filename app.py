import streamlit as st
from datetime import datetime, timedelta
import pandas as pd
import os
import glob
import io

# ⬇️ HTML 임베드용
import streamlit.components.v1 as components

# =========================
# 페이지/스타일 설정
# =========================
st.set_page_config(
    page_title=" 기존 수기 방식 식단 개선 시스템",
    page_icon="🍽️",
    layout="wide",
    initial_sidebar_state="expanded"  # ← 사이드바 기본 펼침
)

st.markdown("""
<style>
    /* 메인 배경 */
    .main {
        padding: 2rem 3rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        min-height: 100vh;
    }
    .card { background: white; padding: 2rem; border-radius: 15px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); margin-bottom: 2rem; border: 1px solid rgba(255,255,255,0.2); }
    .login-card { max-width: 400px; margin: 5rem auto; background: white; padding: 3rem; border-radius: 20px; box-shadow: 0 15px 35px rgba(0,0,0,0.1); text-align: center; }
    .title { color: #2c3e50; font-size: 2.5rem; font-weight: 700; margin-bottom: 1rem; text-align: center; }
    .subtitle { color: #7f8c8d; font-size: 1.1rem; margin-bottom: 2rem; text-align: center; }
    .success-banner { background: linear-gradient(90deg, #56ab2f, #a8e6cf); color: white; padding: 1rem; border-radius: 10px; text-align: center; font-weight: 600; margin-bottom: 2rem; }
    .admin-header { background: linear-gradient(90deg, #667eea, #764ba2); color: white; padding: 1.5rem; border-radius: 10px; text-align: center; margin-bottom: 2rem; }
    .user-header { background: linear-gradient(90deg, #4facfe, #00f2fe); color: white; padding: 1.5rem; border-radius: 10px; text-align: center; margin-bottom: 2rem; }
    .stButton > button { background: linear-gradient(45deg, #667eea, #764ba2); color: white; border: none; padding: 0.75rem 2rem; border-radius: 25px; font-weight: 600; transition: all 0.3s ease; box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3); }
    .stButton > button:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(102, 126, 234, 0.4); }
    .danger-button button { background: linear-gradient(45deg, #ff416c, #ff4b2b) !important; box-shadow: 0 4px 15px rgba(255, 65, 108, 0.3) !important; }
    .danger-button button:hover { box-shadow: 0 6px 20px rgba(255, 65, 108, 0.4) !important; }
    .start-button button { background: linear-gradient(45deg, #56ab2f, #a8e6cf) !important; box-shadow: 0 4px 15px rgba(86, 171, 47, 0.3) !important; font-size: 1.1rem !important; padding: 1rem 2.5rem !important; }
    .stFileUploader { background: #f8f9fa; border: 2px dashed #667eea; border-radius: 10px; padding: 2rem; text-align: center; }
    .dataframe { border-radius: 10px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
    .stat-card { background: white; padding: 1.5rem; border-radius: 10px; text-align: center; box-shadow: 0 4px 15px rgba(0,0,0,0.1); margin: 0.5rem; }
    .stat-number { font-size: 2rem; font-weight: 700; color: #667eea; }
    .stat-label { color: #7f8c8d; font-size: 0.9rem; margin-top: 0.5rem; }
    .stTextInput > div > div > input { border-radius: 10px; border: 2px solid #e1e8ed; padding: 0.75rem; transition: all 0.3s ease; }
    .stTextInput > div > div > input:focus { border-color: #667eea; box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1); }
    .stSelectbox > div > div > select { border-radius: 10px; border: 2px solid #e1e8ed; }
</style>
""", unsafe_allow_html=True)

# =========================
# 상수/폴더
# =========================
LOG_FILE = "log.csv"
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

DEFAULT_HTML_PATH = "/mnt/data/index.html"
DEFAULT_XLSX_CANDIDATES = [
    "/mnt/data/menu.xlsx",
    "/mnt/data/정선_음식 데이터_간식제외.xlsx"
]

# =========================
# 사용자
# =========================
user_dict = {
    "SR01": "test01",
    "SR02": "test02",
    "SR03": "test03",
    "SR04": "test04",
    "SR05": "test05",
    "SR06": "test06",
    "SR07": "test07",
    "SR08": "test08",
    "SR09": "test09",
    "SR10": "test10",
    "SR11": "test11",
    "SR12": "test12",
    "SR13": "test13",
    "admin": "admin"
}

def get_kst_now():
    return datetime.utcnow() + timedelta(hours=9)

# =========================
# 상태 초기화
# =========================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.start_time = None
if "meal_type" not in st.session_state:
    st.session_state.meal_type = "식단표A"

# =========================
# 사이드바: 네비게이션
# =========================
with st.sidebar:
    st.markdown("### 📚 메뉴")
    nav = st.radio(
        "이동",
        options=[
            "🚀 작업/제출",
            "📊 업로드 데이터 미리보기",
            "🧩 메뉴 시각화(HTML)"
        ],
        index=0
    )
    st.markdown("---")
    st.caption("※ HTML/엑셀은 /mnt/data 아래 기본 경로를 우선 사용합니다.")

# =========================
# 로그인 화면
# =========================
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

else:
    # =========================
    # 공통 상단 배너
    # =========================
    st.markdown(f"""
    <div class="success-banner">
        🎉 {st.session_state.username}님 환영합니다!
    </div>
    """, unsafe_allow_html=True)

    # =========================
    # 관리자 페이지
    # =========================
    if st.session_state.username == "admin":
        st.markdown("""
        <div class="admin-header">
            <h1>🔧 관리자 페이지</h1>
            <p>시스템 관리 및 제출 기록 확인</p>
        </div>
        """, unsafe_allow_html=True)

        # 통계 카드
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

        # 관리 버튼
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

        if os.path.exists(LOG_FILE):
            df = pd.read_csv(LOG_FILE)
            st.dataframe(df, use_container_width=True)

            st.markdown("<br>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)

            with col1:
                user_list = df["사용자"].unique().tolist()
                selected_user = st.selectbox("👤 사용자 선택", user_list)

            with col2:
                # 사용자 제출 파일(A/B) 다운로드 버튼 자동 생성
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

        # 관리자도 사이드바 전환 메뉴 사용 가능 (HTML/엑셀 미리보기)
        st.markdown("---")

    # =========================
    # 사용자/관리자 공통: 사이드바 섹션별 화면
    # =========================
    if nav == "🧩 메뉴 시각화(HTML)":
        st.markdown("""<div class="card"><h3>🧩 메뉴 시각화(HTML 임베드)</h3><p>/mnt/data/index.html 파일을 임베드합니다.</p></div>""", unsafe_allow_html=True)

        # 1) /mnt/data/index.html 우선
        html_path = DEFAULT_HTML_PATH
        # 2) 업로더로 대체 업로드 옵션 제공
        uploaded_html = st.file_uploader("또는 HTML 파일 업로드(옵션)", type=["html", "htm"], key="html_uploader")

        html_to_render = None
        if uploaded_html is not None:
            try:
                html_bytes = uploaded_html.getvalue()
                html_to_render = html_bytes.decode("utf-8", errors="ignore")
            except Exception as e:
                st.error(f"업로드한 HTML을 읽는 중 오류: {e}")
        elif os.path.exists(html_path):
            try:
                with open(html_path, "r", encoding="utf-8") as f:
                    html_to_render = f.read()
            except Exception:
                # 인코딩 재시도
                with open(html_path, "r", encoding="cp949", errors="ignore") as f:
                    html_to_render = f.read()

        if html_to_render:
            components.html(html_to_render, height=900, scrolling=True)
            st.success("✅ HTML 렌더 완료")
        else:
            st.warning("⚠️ 렌더할 HTML이 없습니다. /mnt/data/index.html을 배치하거나 파일을 업로드하세요.")

    elif nav == "📊 업로드 데이터 미리보기":
        st.markdown("""<div class="card"><h3>📊 엑셀 데이터 미리보기</h3><p>/mnt/data의 기본 파일을 자동 탐색하거나 직접 업로드할 수 있습니다.</p></div>""", unsafe_allow_html=True)

        # 기본 후보 경로에서 첫 번째로 존재하는 파일을 사용
        default_found = None
        for p in DEFAULT_XLSX_CANDIDATES:
            if os.path.exists(p):
                default_found = p
                break

        st.write("🔍 기본 경로 탐색 결과:", default_found if default_found else "없음")

        tab1, tab2 = st.tabs(["🔎 기본 파일 열기", "⬆️ 직접 업로드"])

        with tab1:
            if default_found:
                try:
                    df0 = pd.read_excel(default_found)
                    st.info(f"기본 파일 열기: `{os.path.basename(default_found)}`")
                    st.dataframe(df0, use_container_width=True)
                except Exception as e:
                    st.error(f"엑셀 읽기 오류: {e}")
            else:
                st.warning("기본 경로에 엑셀이 없습니다. 아래 탭에서 업로드하거나 /mnt/data에 파일을 두세요.")

        with tab2:
            up = st.file_uploader("엑셀 업로드", type=["xlsx", "xls"], key="xlsx_preview")
            if up:
                try:
                    df_up = pd.read_excel(up)
                    st.success(f"✅ 업로드 파일: {up.name}")
                    st.dataframe(df_up, use_container_width=True)
                except Exception as e:
                    st.error(f"엑셀 읽기 오류: {e}")

    else:  # "🚀 작업/제출"
        # 사용자 페이지 (기존 플로우)
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

        st.markdown("""
        <div class="card">
            <h3>🚀 작업 시작</h3>
            <p>아래 버튼을 클릭하여 기존 방식으로 식단 개선 작업을 시작하세요.</p>
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
            with col2:
                st.markdown('<div class="start-button">', unsafe_allow_html=True)
                if st.button("🚪 로그아웃", use_container_width=True):
                    st.session_state.logged_in = False
                    st.session_state.username = ""
                    st.session_state.start_time = None
                    st.session_state.meal_type = "식단표A"
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

        else:
            current_time = get_kst_now()
            elapsed = current_time - st.session_state.start_time
            elapsed_seconds = int(elapsed.total_seconds())

            st.markdown(f"""
            <div style="background: linear-gradient(90deg, #56ab2f, #a8e6cf); color: white; 
                        padding: 1rem; border-radius: 10px; text-align: center; margin-bottom: 2rem;">
                ⏱️ 작업 진행 중... | 시작 시간: {st.session_state.start_time.strftime('%H:%M:%S')} 
                | 경과 시간: {elapsed_seconds}초 | 선택: {st.session_state.meal_type}
            </div>
            """, unsafe_allow_html=True)

        # 파일 업로드/제출
        if st.session_state.start_time:
            st.markdown("""
            <div class="card">
                <h3>📁 파일 업로드</h3>
                <p>완성된 식단 설계 엑셀 파일을 업로드해주세요.</p>
            </div>
            """, unsafe_allow_html=True)

            uploaded_file = st.file_uploader(
                "📊 엑셀 파일 선택",
                type=["xlsx", "xls"],
                help="xlsx 또는 xls 파일만 업로드 가능합니다.",
                key="xlsx_submit"
            )

            if uploaded_file:
                st.success(f"✅ 파일 선택됨: {uploaded_file.name}")

                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("📤 제출하기", use_container_width=True):
                        submit_time = get_kst_now()
                        duration = (submit_time - st.session_state.start_time).total_seconds()

                        safe_meal = st.session_state.meal_type  # "식단표A"/"식단표B"
                        save_name = f"{st.session_state.username}_{safe_meal}.xlsx"
                        file_path = os.path.join(UPLOAD_FOLDER, save_name)
                        with open(file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())

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
