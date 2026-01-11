import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import io
import sqlite3
import datetime
import re
import smtplib
from email.mime.text import MIMEText
from scipy.stats import gmean
from PIL import Image
import streamlit.components.v1 as components

# =============================================================================
# 0. 시스템 설정 및 유틸리티 (DB, 메일, 유효성 검사)
# =============================================================================

# DB 초기화 함수
def init_db():
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    # users 테이블: id, pw, role(temp/official/admin), signup_date, expiry_date
    c.execute('''CREATE TABLE IF NOT EXISTS users
                 (id TEXT PRIMARY KEY, pw TEXT, role TEXT, signup_date TEXT, expiry_date TEXT)''')
    
    # 관리자 계정 자동 생성 (없을 경우)
    # 관리자: shjeon / @jsh2143033
    try:
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?, ?)", 
                  ('shjeon', '@jsh2143033', 'admin', 
                   str(datetime.date.today()), '9999-12-31'))
        conn.commit()
    except sqlite3.IntegrityError:
        pass # 이미 존재하면 패스
    conn.close()

# 유효성 검사 함수
def validate_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email)

def validate_password(password):
    # 문자, 특수문자 포함 여부 확인 (간단한 정규식 예시)
    if len(password) < 4: return False
    has_char = re.search(r'[a-zA-Z]', password)
    has_special = re.search(r'[!@#$%^&*(),.?":{}|<>]', password)
    return has_char and has_special

# 메일 발송 함수 (SMTP 설정 필요)
def send_application_email(user_email):
    # [주의] 실제 메일 발송을 위해서는 아래 SMTP 서버 설정에 본인의 계정 정보를 입력해야 합니다.
    # 현재는 기능 동작을 위한 틀만 제공하며, 설정이 없으면 콘솔에 로그만 출력합니다.
    sender_email = "your_email@gmail.com"  # 보내는 사람 이메일 (설정 필요)
    sender_password = "your_app_password"  # 보내는 사람 앱 비밀번호 (설정 필요)
    recipient_email = "jeon080423@gmail.com"
    
    subject = f"[AHP 앱] 정식 사용자 가입 신청: {user_email}"
    body = f"새로운 정식 사용자가 가입했습니다.\n\n아이디: {user_email}\n가입일: {datetime.date.today()}"
    
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email

    try:
        # SMTP 서버 예시 (Gmail) - 실제 사용시 주석 해제 및 설정 필요
        # with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        #     server.login(sender_email, sender_password)
        #     server.sendmail(sender_email, recipient_email, msg.as_string())
        print(f"메일 발송 성공 (Simulated): {user_email} -> {recipient_email}")
    except Exception as e:
        print(f"메일 발송 실패: {e}")

# DB 관련 함수
def add_user(user_id, pw, role):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    signup_date = datetime.date.today()
    
    if role == 'official':
        # 정식 사용자는 30일 후 만료
        expiry_date = signup_date + datetime.timedelta(days=30)
    else:
        # 임시/관리자는 만료일 없음 (또는 먼 미래)
        expiry_date = datetime.date(9999, 12, 31)
        
    try:
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?, ?)", 
                  (user_id, pw, role, str(signup_date), str(expiry_date)))
        conn.commit()
        success = True
    except sqlite3.IntegrityError:
        success = False
    finally:
        conn.close()
    return success

def check_login(user_id, pw):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("SELECT role, expiry_date FROM users WHERE id=? AND pw=?", (user_id, pw))
    result = c.fetchone()
    conn.close()
    return result # (role, expiry_date) or None

def get_all_users():
    conn = sqlite3.connect('users.db')
    df = pd.read_sql_query("SELECT * FROM users", conn)
    conn.close()
    return df

def update_user_info(user_id, new_role, new_expiry):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("UPDATE users SET role=?, expiry_date=? WHERE id=?", (new_role, new_expiry, user_id))
    conn.commit()
    conn.close()

# -----------------------------------------------------------------------------
# 1. AHP Utility Functions (기존 코드)
# -----------------------------------------------------------------------------

def get_ri(n):
    ri_dict = {
        1: 0.00, 2: 0.00, 3: 0.58, 4: 0.90, 5: 1.12, 
        6: 1.24, 7: 1.32, 8: 1.41, 9: 1.45, 10: 1.49,
        11: 1.51, 12: 1.48, 13: 1.56, 14: 1.57, 15: 1.59
    }
    return ri_dict.get(n, 1.59)

def calculate_weights_gm(matrix):
    geom_means = gmean(matrix, axis=1)
    weights = geom_means / geom_means.sum()
    return weights

def calculate_consistency(matrix):
    n = matrix.shape[0]
    if n <= 2:
        return 0.0, 0.0, n
    
    weights = calculate_weights_gm(matrix)
    weighted_sum = matrix.dot(weights)
    lambda_values = weighted_sum / weights
    lambda_max = lambda_values.mean()
    
    ci = (lambda_max - n) / (n - 1)
    ri = get_ri(n)
    cr = ci / ri if ri > 0 else 0.0
    
    return cr, ci, lambda_max

def improve_consistency(matrix, threshold, max_iter=500, learning_rate=0.2):
    current_matrix = matrix.copy()
    n = current_matrix.shape[0]
    cr, ci, _ = calculate_consistency(current_matrix)
    
    iterations = 0
    if cr <= threshold:
        return current_matrix, cr, iterations, False

    for it in range(max_iter):
        if cr <= threshold:
            break
        
        w = calculate_weights_gm(current_matrix)
        consistent_matrix = np.outer(w, 1/w)
        new_matrix = (current_matrix * (1 - learning_rate)) + (consistent_matrix * learning_rate)
        
        for i in range(n):
            new_matrix[i, i] = 1.0
            for j in range(i + 1, n):
                val = new_matrix[i, j]
                new_matrix[j, i] = 1.0 / val
                
        current_matrix = new_matrix
        cr, ci, _ = calculate_consistency(current_matrix)
        iterations += 1
        
    was_corrected = iterations > 0
    return current_matrix, cr, iterations, was_corrected

def parse_input_value(val):
    if val == 0: return 1.0
    elif val < 0: return abs(val)
    elif val == 1: return 1.0
    else: return 1.0 / val

def infer_factors_from_columns(cols):
    m = len(cols)
    delta = 1 + 8 * m
    n = int((1 + np.sqrt(delta)) / 2)
    factors = [f"Factor{i+1}" for i in range(n)]
    return factors, n

def process_ahp_data(df, cr_threshold, max_iter):
    meta_cols = df.columns[:2]
    comp_cols = df.columns[2:]
    
    factors, n = infer_factors_from_columns(comp_cols)
    results_list = []
    
    for idx, row in df.iterrows():
        respondent_id = row.iloc[0]
        respondent_type = row.iloc[1]
        
        matrix = np.eye(n)
        col_idx = 0
        
        for i in range(n):
            for j in range(i + 1, n):
                if col_idx < len(comp_cols):
                    raw_val = row[comp_cols[col_idx]]
                    ahp_val = parse_input_value(raw_val)
                    matrix[i, j] = ahp_val
                    matrix[j, i] = 1.0 / ahp_val
                    col_idx += 1
        
        orig_cr, orig_ci, _ = calculate_consistency(matrix)
        final_matrix = matrix.copy()
        final_cr = orig_cr
        iterations = 0
        corrected_flag = False
        
        if orig_cr > cr_threshold:
            final_matrix, final_cr, iterations, corrected_flag = improve_consistency(
                matrix, cr_threshold, max_iter=max_iter
            )
        
        final_weights = calculate_weights_gm(final_matrix)
        
        res = {
            "ID": respondent_id,
            "Type": respondent_type,
            "Original_CR": orig_cr,
            "Final_CR": final_cr,
            "Iterations": iterations,
            "Corrected": corrected_flag,
            "Matrix_Object": final_matrix
        }
        
        for f_idx, f_name in enumerate(factors):
            res[f"Weight_{f_name}"] = final_weights[f_idx]
            
        results_list.append(res)
        
    results_df = pd.DataFrame(results_list)
    return results_df, factors

# -----------------------------------------------------------------------------
# 2. Main Setup & UI
# -----------------------------------------------------------------------------

# 앱 초기화 및 DB 설정
init_db()

# 페이지 설정
try:
    icon_img = Image.open("image_4.png")
    st.set_page_config(page_title="AHP Analysis Tool", layout="wide", page_icon=icon_img)
except FileNotFoundError:
    st.set_page_config(page_title="AHP Analysis Tool", layout="wide", page_icon="📊")

# 세션 상태 초기화
if 'user_id' not in st.session_state:
    st.session_state.user_id = None
if 'user_role' not in st.session_state:
    st.session_state.user_role = None # 'temp', 'official', 'admin'
if 'expiry_date' not in st.session_state:
    st.session_state.expiry_date = None

# 스타일링
st.markdown("""
<style>
    .stDataFrame {font-size: 0.9rem;}
    div[data-testid="stMetricValue"] {font-size: 1.2rem;}
</style>
""", unsafe_allow_html=True)

# 메인 헤더
col_h1, col_h2 = st.columns([1, 15])
with col_h1:
    try:
        st.image("image_4.png", width=80) 
    except Exception:
        st.header("📊")
with col_h2:
    st.title("AHP 분석 자동화 시스템")

st.markdown("""
**Analytic Hierarchy Process (AHP)** 분석 및 일관성 자동 보정 도구입니다.
엑셀 파일을 업로드하면 개인별 가중치 산출, 일관성 보정(CR), 그룹별 집계 결과를 제공합니다.
""")

# -----------------------------------------------------------------------------
# 3. 사이드바: 로그인 / 회원가입 / 관리자 메뉴
# -----------------------------------------------------------------------------

# 페이팔 코드
paypal_html = """
<div id="paypal-container-Y5JKCC6YSVDRC"></div>
<script>
  paypal.HostedButtons({
    hostedButtonId: "Y5JKCC6YSVDRC",
  }).render("#paypal-container-Y5JKCC6YSVDRC")
</script>
<script src="https://www.paypal.com/sdk/js?client-id=BAA&components=hosted-buttons&disable-funding=venmo&currency=USD"></script>
"""

with st.sidebar:
    if st.session_state.user_id is None:
        # 비로그인 상태: 로그인/회원가입 탭
        tab_login, tab_signup = st.tabs(["로그인", "회원가입"])
        
        with tab_login:
            st.header("로그인")
            l_id = st.text_input("아이디 (ID)", key="l_id")
            l_pw = st.text_input("비밀번호 (PW)", type="password", key="l_pw")
            if st.button("로그인 실행"):
                result = check_login(l_id, l_pw)
                if result:
                    st.session_state.user_id = l_id
                    st.session_state.user_role = result[0]
                    st.session_state.expiry_date = result[1]
                    st.success(f"환영합니다, {l_id}님!")
                    st.rerun()
                else:
                    st.error("아이디 또는 비밀번호가 일치하지 않습니다.")

        with tab_signup:
            st.header("회원가입")
            s_id = st.text_input("아이디 (이메일)", key="s_id", help="이메일 형식으로 입력해주세요.")
            s_pw = st.text_input("비밀번호", type="password", key="s_pw", help="문자와 특수문자를 혼합하여 사용해주세요.")
            s_role = st.radio("이용 권한 선택", 
                              ("임시 사용자 (5 Sample)", "정식 사용자 (1개월 무제한)"), 
                              index=0)
            
            # 정식 사용자 선택 시 페이팔 노출
            if "정식" in s_role:
                st.markdown("---")
                st.write("**정식 사용자 결제 (PayPal)**")
                components.html(paypal_html, height=150)
                st.markdown("---")

            if st.button("가입하기"):
                # 유효성 검사
                if not validate_email(s_id):
                    st.error("아이디는 올바른 이메일 형식이어야 합니다.")
                elif not validate_password(s_pw):
                    st.error("비밀번호는 문자와 특수문자를 혼합해야 합니다.")
                else:
                    role_code = 'official' if "정식" in s_role else 'temp'
                    if add_user(s_id, s_pw, role_code):
                        st.success("가입이 완료되었습니다! 로그인 탭에서 로그인해주세요.")
                        # 정식 사용자일 경우 신청 메일 발송
                        if role_code == 'official':
                            send_application_email(s_id)
                    else:
                        st.error("이미 존재하는 아이디입니다.")

    else:
        # 로그인 상태
        st.success(f"**{st.session_state.user_id}** 님 접속 중")
        
        # 역할 표시
        role_disp = "관리자" if st.session_state.user_role == 'admin' else \
                    ("정식 사용자" if st.session_state.user_role == 'official' else "임시 사용자")
        st.info(f"권한: {role_disp}")
        
        if st.button("로그아웃"):
            st.session_state.user_id = None
            st.session_state.user_role = None
            st.session_state.expiry_date = None
            st.rerun()

        st.markdown("---")
        
        # 관리자 전용 메뉴
        if st.session_state.user_role == 'admin':
            st.subheader("⚙️ 관리자 메뉴")
            if st.checkbox("회원 관리 모드 켜기"):
                st.session_state['admin_mode'] = True
            else:
                st.session_state['admin_mode'] = False

        # 분석 설정 (로그인 상태일 때만 보임)
        st.header("분석 설정")
        cr_threshold = st.selectbox("일관성 비율(CR) 임계값", [0.1, 0.2], index=0)
        max_iter = st.number_input("최대 보정 반복 횟수", min_value=10, max_value=2000, value=500, step=50)

    st.markdown("---")
    st.markdown("**제작: 전상현**")
    st.markdown("jeon080423@gmail.com")

# -----------------------------------------------------------------------------
# 4. 메인 컨텐츠 (관리자 모드 or 분석 도구)
# -----------------------------------------------------------------------------

# 관리자 모드가 켜져있으면 회원 관리 화면 표시
if st.session_state.get('admin_mode', False) and st.session_state.user_role == 'admin':
    st.subheader("👥 가입자 현황 및 관리")
    users_df = get_all_users()
    st.dataframe(users_df)

    with st.expander("회원 권한/기간 수정"):
        edit_id = st.selectbox("수정할 회원 ID", users_df['id'].unique())
        selected_user = users_df[users_df['id'] == edit_id].iloc[0]
        
        new_role = st.selectbox("권한 변경", ['temp', 'official', 'admin'], 
                                index=['temp', 'official', 'admin'].index(selected_user['role']))
        new_expiry = st.text_input("만료일 변경 (YYYY-MM-DD)", value=selected_user['expiry_date'])
        
        if st.button("정보 수정 적용"):
            update_user_info(edit_id, new_role, new_expiry)
            st.success(f"{edit_id} 회원의 정보가 수정되었습니다.")
            st.rerun()
    
    st.divider()

# 분석 도구 화면 (항상 보이지만, 기능 실행 시 권한 체크)
uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        excel_obj = pd.ExcelFile(uploaded_file)
        target_sheet = excel_obj.sheet_names[0]
        df_input = pd.read_excel(uploaded_file, sheet_name=target_sheet)
        
        st.info(f"파일 로드됨: {target_sheet} (데이터 행 수: {len(df_input)}개)")

        # -----------------------------------------------------
        # [권한 체크 및 제한 로직]
        # -----------------------------------------------------
        permission_granted = False
        message = ""

        if st.session_state.user_id is None:
            message = "⚠️ 분석 기능을 이용하시려면 **로그인**이 필요합니다."
        
        else:
            role = st.session_state.user_role
            today = datetime.date.today()
            expiry_str = st.session_state.expiry_date
            expiry = datetime.datetime.strptime(expiry_str, "%Y-%m-%d").date()

            if role == 'admin':
                permission_granted = True
            
            elif role == 'temp':
                # 임시 사용자: 5 Sample 제한
                if len(df_input) <= 5:
                    permission_granted = True
                else:
                    message = f"⛔ **임시 사용자**는 최대 5개 표본까지만 분석 가능합니다. (현재: {len(df_input)}개)\n\n정식 사용자로 전환하여 무제한 분석을 이용하세요."
            
            elif role == 'official':
                # 정식 사용자: 기간 제한 확인
                if today <= expiry:
                    permission_granted = True
                else:
                    message = f"⛔ **이용 기간이 만료**되었습니다. (만료일: {expiry_str})\n관리자에게 문의하여 기간을 연장하세요."

        # -----------------------------------------------------
        # 분석 실행
        # -----------------------------------------------------
        if permission_granted:
            with st.spinner("행렬 구성 및 일관성 보정 작업 중..."):
                results_df, factors = process_ahp_data(df_input, cr_threshold, max_iter)
            
            # 순위 산출
            sorted_weight_cols = [f"Weight_{f}" for f in factors]
            rank_df = results_df[sorted_weight_cols].rank(axis=1, ascending=False, method='min')
            sorted_rank_cols = [f"Rank_{f}" for f in factors]
            rank_df.columns = sorted_rank_cols
            results_df = pd.concat([results_df, rank_df], axis=1)

            # 컬럼 정리
            meta_cols = ["ID", "Type", "Original_CR", "Final_CR", "Iterations", "Corrected"]
            final_col_order = meta_cols + sorted_weight_cols + sorted_rank_cols + ["Matrix_Object"]
            results_df = results_df[final_col_order]

            # 집계
            agg_df = results_df.groupby("Type")[sorted_weight_cols].mean().reset_index()
            overall_weights = results_df[sorted_weight_cols].mean()
            overall_weights.index = [c.replace("Weight_", "") for c in overall_weights.index]

            # 종합 행렬
            all_matrices = np.stack(results_df["Matrix_Object"].values)
            overall_matrix = gmean(all_matrices, axis=0)
            overall_matrix_df = pd.DataFrame(overall_matrix, index=factors, columns=factors)
            overall_combined_df = overall_matrix_df.copy()
            overall_combined_df['Weight'] = overall_weights
            overall_combined_df['Rank'] = overall_combined_df['Weight'].rank(ascending=False).astype(int)
            final_cols = factors + ['Weight', 'Rank']
            overall_combined_df = overall_combined_df[final_cols]

            # 그룹별 분석
            group_results = {}
            unique_types = results_df["Type"].unique()
            for g_type in unique_types:
                sub_df = results_df[results_df["Type"] == g_type]
                if len(sub_df) > 0:
                    g_matrices = np.stack(sub_df["Matrix_Object"].values)
                    g_matrix = gmean(g_matrices, axis=0)
                    g_weights = sub_df[sorted_weight_cols].mean()
                    g_weights.index = [c.replace("Weight_", "") for c in g_weights.index]
                    g_df = pd.DataFrame(g_matrix, index=factors, columns=factors)
                    g_df['Weight'] = g_weights
                    g_df['Rank'] = g_df['Weight'].rank(ascending=False).astype(int)
                    g_df = g_df[final_cols]
                    group_results[g_type] = g_df

            # 일관성 통계
            group_consistency = results_df.groupby("Type").agg({
                'Original_CR': 'mean',
                'Final_CR': 'mean',
                'Corrected': lambda x: (x.sum() / len(x)) * 100
            }).reset_index()
            group_consistency.columns = ['Group', 'Avg Orig CR', 'Avg Final CR', 'Corrected(%)']

            # ---------------------------
            # UI 표시
            # ---------------------------
            tab1, tab2, tab3 = st.tabs(["분석 결과 요약", "시각화", "데이터 다운로드"])
            
            with tab1:
                st.subheader("🌐 종합 분석 결과")
                format_dict = {col: "{:.3f}" for col in overall_combined_df.columns if col != 'Rank'}
                format_dict['Rank'] = "{:.0f}"

                st.markdown("#### 1️⃣ 전체 (Overall)")
                st.dataframe(overall_combined_df.style.format(format_dict), use_container_width=True)
                
                if len(group_results) > 0:
                    st.divider()
                    st.markdown("#### 2️⃣ 그룹별 상세")
                    for g_type, g_df in group_results.items():
                        st.markdown(f"**📌 그룹: {g_type}**")
                        st.dataframe(g_df.style.format(format_dict), use_container_width=True)
                
                st.divider()
                col1, col2 = st.columns([1, 1])
                with col1:
                    st.subheader("📊 일관성 현황")
                    avg_cr_orig = results_df["Original_CR"].mean()
                    avg_cr_final = results_df["Final_CR"].mean()
                    pct_corrected = (results_df["Corrected"].sum() / len(results_df)) * 100
                    st.metric("전체 평균 보정 전 CR", f"{avg_cr_orig:.3f}")
                    st.metric("전체 평균 보정 후 CR", f"{avg_cr_final:.3f}")
                    st.metric("전체 보정된 응답자 비율", f"{pct_corrected:.1f}%")
                with col2:
                    st.markdown("#### [그룹별 일관성 요약]")
                    st.dataframe(group_consistency.style.format({'Avg Orig CR': '{:.3f}', 'Avg Final CR': '{:.3f}', 'Corrected(%)': '{:.1f}%'}), use_container_width=True)

                st.divider()
                st.subheader("개인별 상세 분석 결과")
                preview_cols = ["ID", "Type", "Original_CR", "Final_CR"] + sorted_weight_cols + sorted_rank_cols
                st.dataframe(results_df[preview_cols].head(10).style.format({"Original_CR": "{:.3f}", "Final_CR": "{:.3f}", **{c: "{:.3f}" for c in sorted_weight_cols}, **{c: "{:.0f}" for c in sorted_rank_cols}}))

            with tab2:
                col_v1, col_v2 = st.columns(2)
                with col_v1:
                    st.subheader("전체 요인 가중치")
                    plot_df = overall_combined_df.reset_index().rename(columns={'index': 'Factor'})
                    fig, ax = plt.subplots(figsize=(6, 4))
                    sns.barplot(x="Weight", y="Factor", data=plot_df, palette="viridis", ax=ax)
                    ax.set_xlim(0, max(plot_df["Weight"])*1.2)
                    for i, v in enumerate(plot_df["Weight"]): ax.text(v, i, f" {v:.3f}", va='center')
                    st.pyplot(fig)
                with col_v2:
                    st.subheader("CR 보정 전후 분포")
                    fig2, ax2 = plt.subplots(figsize=(6, 4))
                    sns.scatterplot(data=results_df, x="Original_CR", y="Final_CR", hue="Type", ax=ax2)
                    ax2.axhline(y=cr_threshold, color='r', linestyle='--', label=f'Threshold ({cr_threshold})')
                    ax2.legend()
                    st.pyplot(fig2)
                
                st.divider()
                st.subheader("그룹별 가중치 분포")
                long_df = results_df.melt(id_vars=["ID", "Type"], value_vars=sorted_weight_cols, var_name="Factor", value_name="Weight")
                long_df["Factor"] = long_df["Factor"].str.replace("Weight_", "")
                fig3, ax3 = plt.subplots(figsize=(10, 5))
                try: sns.boxplot(data=long_df, x="Factor", y="Weight", hue="Type", ax=ax3)
                except: sns.stripplot(data=long_df, x="Factor", y="Weight", hue="Type", dodge=True, ax=ax3)
                st.pyplot(fig3)

            with tab3:
                st.subheader("📥 엑셀 다운로드")
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    worksheet = workbook.add_worksheet('Summary_Analysis')
                    writer.sheets['Summary_Analysis'] = worksheet
                    start_row = 0
                    
                    worksheet.write_string(start_row, 0, "[1] Overall Analysis")
                    start_row += 1
                    overall_combined_df.round(3).to_excel(writer, sheet_name='Summary_Analysis', startrow=start_row)
                    start_row += len(overall_combined_df) + 3
                    
                    if len(group_results) > 0:
                        for g_type, g_df in group_results.items():
                            worksheet.write_string(start_row, 0, f"[2] Group Analysis: {g_type}")
                            start_row += 1
                            g_df.round(3).to_excel(writer, sheet_name='Summary_Analysis', startrow=start_row)
                            start_row += len(g_df) + 3
                    
                    worksheet.write_string(start_row, 0, "[3] Consistency Statistics")
                    start_row += 1
                    group_consistency.round(3).to_excel(writer