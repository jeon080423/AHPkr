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
import json
import platform
import os
import matplotlib.font_manager as fm
from matplotlib import rc
from email.mime.text import MIMEText
from scipy.stats import gmean, ttest_rel, f_oneway
from PIL import Image
import itertools
from math import pi
from dateutil.relativedelta import relativedelta

# [필수] plotly 라이브러리 (requirements.txt에 plotly 추가 필요)
import plotly.express as px
import plotly.graph_objects as go
from scipy import stats
import gspread
from google.oauth2.service_account import Credentials

# ANOVA 및 사후검정을 위한 라이브러리 (없을 경우 예외처리)
try:
    from statsmodels.stats.multicomp import pairwise_tukeyhsd
    STATSMODELS_AVAILABLE = True
except ImportError:
    STATSMODELS_AVAILABLE = False

# =============================================================================
# 0. 시스템 설정 및 유틸리티
# =============================================================================

# [폰트 설정]
def set_font_config():
    system_name = platform.system()
    try:
        if system_name == 'Windows':
            font_path = "c:/Windows/Fonts/malgun.ttf"
            if os.path.exists(font_path):
                font_name = fm.FontProperties(fname=font_path).get_name()
                rc('font', family=font_name)
        elif system_name == 'Darwin': # Mac
            rc('font', family='AppleGothic')
        else: # Linux
            font_path = "NanumGothic.ttf"
            if not os.path.exists(font_path):
                import urllib.request
                url = "https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Regular.ttf"
                urllib.request.urlretrieve(url, font_path)
            fm.fontManager.addfont(font_path)
            font_prop = fm.FontProperties(fname=font_path)
            rc('font', family=font_prop.get_name())
    except Exception as e:
        pass
    plt.rcParams['axes.unicode_minus'] = False 

set_font_config()

# DB 초기화
def init_db():
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users
                  (id TEXT PRIMARY KEY, pw TEXT, role TEXT, signup_date TEXT, expiry_date TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS saved_analyses
                  (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id TEXT, filename TEXT, save_date TEXT, file_data BLOB)''')
    c.execute('''CREATE TABLE IF NOT EXISTS user_models
                  (user_id TEXT PRIMARY KEY, model_data TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS visit_logs
                  (ip_address TEXT, visit_date TEXT, PRIMARY KEY (ip_address, visit_date))''')
    try:
        c.execute("INSERT OR IGNORE INTO users VALUES (?, ?, ?, ?, ?)", 
                  ('shjeon', '@jsh2143033', 'admin', str(datetime.date.today()), '9999-12-31'))
        conn.commit()
    except sqlite3.IntegrityError:
        pass 
    conn.close()

# 방문자 추적
def track_visitor():
    try:
        if hasattr(st, "context") and hasattr(st.context, "headers"):
            ip = st.context.headers.get("X-Forwarded-For", "unknown_ip")
        else:
            ip = "localhost"
        today = str(datetime.date.today())
        conn = sqlite3.connect('users.db')
        c = conn.cursor()
        c.execute("INSERT OR IGNORE INTO visit_logs (ip_address, visit_date) VALUES (?, ?)", (ip, today))
        conn.commit()
        conn.close()
    except Exception as e:
        pass

track_visitor()

# 유효성 검사
def validate_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email)

def validate_password(password):
    if len(password) < 4: return False
    has_char = re.search(r'[a-zA-Z]', password)
    has_special = re.search(r'[!@#$%^&*(),.?":{}|<>]', password)
    return has_char and has_special

# 메일 발송 함수들
def send_application_email(user_email):
    sender_email = "jeon080423@gmail.com"
    password = "csuh xxru wqdy mttt" 
    recipient_email = "jeon080423@gmail.com"
    subject = f"[AHP 앱] 정식 사용자 승인 요청: {user_email}"
    body = f"사용자가 정식 권한 신청.\nID: {user_email}\n신청일: {datetime.date.today()}"
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
    except: pass

def send_approval_email(user_email):
    sender_email = "jeon080423@gmail.com"
    password = "csuh xxru wqdy mttt" 
    recipient_email = user_email
    subject = "[AHP 분석 시스템] 정식 사용자 승인 완료"
    body = f"{user_email}님, 정식 사용자로 승인되었습니다. 오늘부터 2개월간 모든 기능을 무제한으로 사용하실 수 있습니다."
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
        return True
    except: return False

def send_password_recovery_email(user_email, user_pw):
    sender_email = "jeon080423@gmail.com"
    password = "csuh xxru wqdy mttt"
    recipient_email = user_email
    subject = "[AHP 분석 시스템] 비밀번호 안내"
    body = f"""안녕하세요. 요청하신 계정 정보를 안내해 드립니다.

ID: {user_email}
PW: {user_pw}

로그인 후 비밀번호를 변경하시기를 권장합니다.
감사합니다.
"""
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
        return True
    except Exception as e:
        return False

# --- DB CRUD ---

def log_to_sheets(user_id, role, signup_date):
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
        client = gspread.authorize(creds)
        sheet = client.open('AHPkr_Users').sheet1
        sheet.append_row([user_id, role, str(signup_date)])
    except:
        pass
def add_user(user_id, pw, role):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    signup_date = datetime.date.today()
    expiry_date = datetime.date(9999, 12, 31)
    try:
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?, ?)", 
                  (user_id, pw, role, str(signup_date), str(expiry_date)))
        conn.commit()
        log_to_sheets(user_id, role, signup_date)
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
    return result

def get_user_password(user_id):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("SELECT pw FROM users WHERE id=?", (user_id,))
    result = c.fetchone()
    conn.close()
    return result[0] if result else None

def change_user_password(user_id, new_pw):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("UPDATE users SET pw=? WHERE id=?", (new_pw, user_id))
    conn.commit()
    conn.close()
    return True

def get_all_users():
    conn = sqlite3.connect('users.db')
    df = pd.read_sql_query("SELECT * FROM users", conn)
    conn.close()
    return df

def update_user_full_info(user_id, new_pw, new_role, new_expiry):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    if new_pw and new_pw.strip() != "":
        c.execute("UPDATE users SET pw=?, role=?, expiry_date=? WHERE id=?", (new_pw, new_role, new_expiry, user_id))
    else:
        c.execute("UPDATE users SET role=?, expiry_date=? WHERE id=?", (new_role, new_expiry, user_id))
    conn.commit()
    conn.close()

def delete_user(user_id):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("DELETE FROM users WHERE id=?", (user_id,))
    c.execute("DELETE FROM saved_analyses WHERE user_id=?", (user_id,))
    c.execute("DELETE FROM user_models WHERE user_id=?", (user_id,))
    conn.commit()
    conn.close()

def save_analysis_to_db(user_id, filename, file_data):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    save_date = str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    c.execute("INSERT INTO saved_analyses (user_id, filename, save_date, file_data) VALUES (?, ?, ?, ?)",
              (user_id, filename, save_date, file_data))
    conn.commit()
    conn.close()

def get_user_analyses(user_id):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("SELECT id, filename, save_date FROM saved_analyses WHERE user_id=? ORDER BY save_date DESC", (user_id,))
    rows = c.fetchall()
    conn.close()
    return rows

def get_analysis_file(analysis_id):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("SELECT filename, file_data FROM saved_analyses WHERE id=?", (analysis_id,))
    result = c.fetchone()
    conn.close()
    return result

def delete_analysis(analysis_id):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("DELETE FROM saved_analyses WHERE id=?", (analysis_id,))
    conn.commit()
    conn.close()

def save_user_model(user_id, model_dict):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    model_json = json.dumps(model_dict, ensure_ascii=False)
    c.execute("INSERT OR REPLACE INTO user_models (user_id, model_data) VALUES (?, ?)", (user_id, model_json))
    conn.commit()
    conn.close()

def load_user_model(user_id):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("SELECT model_data FROM user_models WHERE user_id=?", (user_id,))
    result = c.fetchone()
    conn.close()
    if result:
        return json.loads(result[0])
    return None

# -----------------------------------------------------------------------------
# 1. AHP Functions
# -----------------------------------------------------------------------------
def get_ri(n):
    ri_dict = {1: 0.00, 2: 0.00, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24, 7: 1.32, 8: 1.41, 9: 1.45, 10: 1.49}
    return ri_dict.get(n, 1.49)

def calculate_weights(matrix, method='geometric'):
    if method == 'arithmetic':
        col_sum = matrix.sum(axis=0)
        col_sum[col_sum == 0] = 1
        normalized_matrix = matrix / col_sum
        weights = normalized_matrix.mean(axis=1)
    else:
        geom_means = gmean(matrix, axis=1)
        weights = geom_means / geom_means.sum()
    return weights

def calculate_consistency(matrix, method='geometric'):
    n = matrix.shape[0]
    if n <= 2: return 0.0, 0.0, n
    weights = calculate_weights(matrix, method)
    weighted_sum = matrix.dot(weights)
    weights_safe = weights.copy()
    weights_safe[weights_safe == 0] = 1e-10
    lambda_values = weighted_sum / weights_safe
    lambda_max = lambda_values.mean()
    ci = (lambda_max - n) / (n - 1)
    ri = get_ri(n)
    cr = ci / ri if ri > 0 else 0.0
    return cr, ci, lambda_max

def improve_consistency(matrix, threshold, max_iter=500, learning_rate=0.2, method='geometric'):
    current_matrix = matrix.copy()
    n = current_matrix.shape[0]
    cr, ci, _ = calculate_consistency(current_matrix, method)
    iterations = 0
    if cr <= threshold: return current_matrix, cr, iterations, False
    for it in range(max_iter):
        if cr <= threshold: break
        w = calculate_weights(current_matrix, method)
        consistent_matrix = np.outer(w, 1/w)
        new_matrix = (current_matrix * (1 - learning_rate)) + (consistent_matrix * learning_rate)
        for i in range(n):
            new_matrix[i, i] = 1.0
            for j in range(i + 1, n):
                val = new_matrix[i, j]
                new_matrix[j, i] = 1.0 / val
        current_matrix = new_matrix
        cr, ci, _ = calculate_consistency(current_matrix, method)
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
    extracted_factors = []
    seen = set()
    for c in cols:
        parts = str(c).split('_')
        for p in parts:
            p_str = p.strip()
            if p_str not in seen:
                seen.add(p_str)
                extracted_factors.append(p_str)
    if len(extracted_factors) == n:
        factors = extracted_factors 
    else:
        factors = [f"F{i+1}" for i in range(n)]
    return factors, n

def calculate_pairwise_ttest(df, factors):
    n = len(factors)
    p_values = pd.DataFrame(index=factors, columns=factors)
    weight_cols = [f"Weight_{f}" for f in factors]
    for i in range(n):
        for j in range(n):
            if i == j:
                p_values.iloc[i, j] = 1.0
            else:
                col1 = weight_cols[i]
                col2 = weight_cols[j]
                if col1 in df.columns and col2 in df.columns and len(df) > 1:
                    try:
                        _, p = ttest_rel(df[col1], df[col2], nan_policy='omit')
                        p_values.iloc[i, j] = p
                    except:
                        p_values.iloc[i, j] = np.nan
                else:
                    p_values.iloc[i, j] = np.nan
    return p_values

def process_single_sheet(df, cr_threshold, max_iter, method='geometric'):
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
        orig_cr, orig_ci, _ = calculate_consistency(matrix, method)
        final_matrix = matrix.copy()
        final_cr = orig_cr
        iterations = 0
        corrected_flag = False
        if orig_cr > cr_threshold:
            final_matrix, final_cr, iterations, corrected_flag = improve_consistency(
                matrix, cr_threshold, max_iter=max_iter, method=method
            )
        _, final_ci, _ = calculate_consistency(final_matrix, method)
        final_weights = calculate_weights(final_matrix, method)
        res = {
            "ID": respondent_id,
            "Type": respondent_type,
            "Original_CR": orig_cr,
            "Final_CR": final_cr,
            "Original_CI": orig_ci,
            "Final_CI": final_ci,
            "Iterations": iterations,
            "Corrected": corrected_flag,
            "Matrix_Object": final_matrix 
        }
        for f_idx, f_name in enumerate(factors):
            res[f"Weight_{f_name}"] = final_weights[f_idx]
        results_list.append(res)
    results_df = pd.DataFrame(results_list)
    return results_df, factors

def create_sample_excel():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        main_cols = ["ID", "Type", "거버넌스_계획타당성", "거버넌스_실현가능성", "거버넌스_사업효과", 
                      "계획타당성_실현가능성", "계획타당성_사업효과", "실현가능성_사업효과"]
        main_data = [
            [1, "전문가", 5, -5, 5, 5, -5, 5],      
            [2, "전문가", 7, 7, -7, -7, 2, -2],   
            [3, "일반", -5, 5, 5, -5, 5, 5],
            [4, "일반", 3, -3, 3, -3, 3, -3],
            [5, "공무원", 9, -9, 9, -9, 9, -9]
        ]
        df_main = pd.DataFrame(main_data, columns=main_cols)
        df_main.to_excel(writer, sheet_name="Main_Criteria", index=False)
        
        inconsistent_pattern = [
            [1, "전문가", 5, -5, 5],
            [2, "전문가", 7, -7, 7],
            [3, "일반", 3, -3, 3],
            [4, "일반", 9, -9, 9],
            [5, "공무원", 4, -4, 4]
        ]
        sub1_cols = ["ID", "Type", "행정지원_지역공동체", "행정지원_총괄사업관리자", "지역공동체_총괄사업관리자"]
        pd.DataFrame(inconsistent_pattern, columns=sub1_cols).to_excel(writer, sheet_name="거버넌스", index=False)
        sub2_cols = ["ID", "Type", "현안적정성_대안적정성", "현안적정성_목표구체성", "대안적정성_목표구체성"]
        pd.DataFrame(inconsistent_pattern, columns=sub2_cols).to_excel(writer, sheet_name="계획타당성", index=False)
        sub3_cols = ["ID", "Type", "부지확보_사업구체화", "부지확보_사업비적정성", "사업구체화_사업비적정성"]
        pd.DataFrame(inconsistent_pattern, columns=sub3_cols).to_excel(writer, sheet_name="실현가능성", index=False)
        sub4_cols = ["ID", "Type", "경제적효과_사회적효과", "경제적효과_성과관리", "사회적효과_성과관리"]
        pd.DataFrame(inconsistent_pattern, columns=sub4_cols).to_excel(writer, sheet_name="사업효과", index=False)
    output.seek(0)
    return output

def calculate_anova_and_posthoc(full_data):
    results = []
    unique_factors = full_data['Factor'].unique()
    
    for factor in unique_factors:
        subset = full_data[full_data['Factor'] == factor]
        groups = [group['Global_Weight'].values for name, group in subset.groupby('Type')]
        
        if len(groups) < 2:
            continue
            
        f_stat, p_val = f_oneway(*groups)
        
        row = {
            "요인": factor,
            "F-Statistic": f_stat,
            "P-Value": p_val,
            "유의성": "유의함" if p_val < 0.05 else "유의하지 않음",
            "사후검정(Tukey HSD)": ""
        }
        
        if p_val < 0.05 and STATSMODELS_AVAILABLE:
            try:
                tukey = pairwise_tukeyhsd(endog=subset['Global_Weight'], groups=subset['Type'], alpha=0.05)
                tukey_df = pd.DataFrame(data=tukey.summary().data[1:], columns=tukey.summary().data[0])
                sig_pairs = tukey_df[tukey_df['reject'] == True]
                if not sig_pairs.empty:
                    pairs_str = []
                    for _, r in sig_pairs.iterrows():
                        pairs_str.append(f"{r['group1']} vs {r['group2']}")
                    row["사후검정(Tukey HSD)"] = ", ".join(pairs_str) + " 차이 있음"
                else:
                    row["사후검정(Tukey HSD)"] = "집단 간 구체적 차이 발견 못함"
            except Exception as e:
                row["사후검정(Tukey HSD)"] = "계산 오류"
        
        results.append(row)
        
    return pd.DataFrame(results)

# -----------------------------------------------------------------------------
# 2. Setup & Layout
# -----------------------------------------------------------------------------

init_db()

try:
    icon_img = Image.open("image_4.png")
    st.set_page_config(page_title="AHP Analysis Tool", layout="wide", page_icon=icon_img)
except FileNotFoundError:
    st.set_page_config(page_title="AHP Analysis Tool", layout="wide", page_icon="📊")

# CSS 최적화
st.markdown("""
<style>
    .stDataFrame {font-size: 0.9rem;} 
    div[data-testid="stMetricValue"] {font-size: 1.2rem;}
</style>
""", unsafe_allow_html=True)

if 'user_id' not in st.session_state: st.session_state.user_id = None
if 'user_role' not in st.session_state: st.session_state.user_role = None
if 'expiry_date' not in st.session_state: st.session_state.expiry_date = None
if 'admin_mode' not in st.session_state: st.session_state.admin_mode = False
if 'model_structure' not in st.session_state: st.session_state.model_structure = {}

col_h1, col_h2 = st.columns([1, 15])
with col_h1:
    try: 
        st.image("image_4.png", width=80) 
    except: 
        st.header("📊")
with col_h2:
    st.title("AHP 분석 자동화 시스템")

st.markdown("Analytic Hierarchy Process (AHP) 분석 및 일관성 자동 보정 도구입니다. 엑셀 파일을 업로드하면 개인별 가중치 산출, 일관성 보정(CR), 그룹별 집계 결과를 제공합니다.")

# =============================================================================
# 3. Sidebar (Auth & Settings)
# =============================================================================

# 이용 요금 공통 안내 텍스트 정의
fee_info_text = """
---
### 💰 서비스 이용 금액 안내
- **학위논문 분석**: 40만원
- **일반 연구 분석**: 50만원

**결제 정보**
- **계좌번호**: 카카오뱅크 333-26-7331429
- **예금주**: 전상현(프레쉬푸드)
- **주의**: 송금자명에 **가입한 이메일 주소**를 기입해주세요.
"""

with st.sidebar:
    with st.expander("💡 사용자 권한 안내", expanded=False):
        st.info("**비로그인(Guest)**: 샘플 파일 분석만 가능 (5행 제한)")
        st.info("**임시 사용자**: 나만의 모델 생성 가능, 분석 5행 제한")
        st.info("**정식 사용자**: 모든 기능 무제한 (2개월)")
        st.info("**관리자**: 모든 기능 무제한 + 관리자 도구")

    if st.session_state.user_id is None:
        tab_login, tab_signup, tab_find_pw = st.tabs(["로그인", "회원가입", "비밀번호 찾기"])
        
        with tab_login:
            st.header("🔐 로그인")
            l_id = st.text_input("아이디 (이메일 주소)", key="l_id")
            l_pw = st.text_input("비밀번호 (PW)", type="password", key="l_pw")
            if st.button("로그인 실행"):
                result = check_login(l_id.strip(), l_pw.strip())
                if result:
                    # 만료 체크 로직
                    today = datetime.date.today()
                    expiry_date = datetime.datetime.strptime(result[1], "%Y-%m-%d").date()
                    if today > expiry_date:
                        st.error(f"❌ 이용 기간이 만료되었습니다. (만료일: {result[1]})")
                    else:
                        st.session_state.user_id = l_id.strip()
                        st.session_state.user_role = result[0]
                        st.session_state.expiry_date = result[1]
                        st.success(f"환영합니다, {l_id}님!")
                        st.rerun()
                else:
                    st.error("아이디 또는 비밀번호가 일치하지 않습니다.")
            
            # [추가] 로그인 탭 내 서비스 이용 요금 안내
            st.markdown(fee_info_text)

        with tab_signup:
            st.header("📝 회원가입")
            s_id = st.text_input("아이디 (이메일 주소)", key="s_id")
            s_pw = st.text_input("비밀번호", type="password", key="s_pw")
            s_role_selection = st.radio("이용 권한 선택", ("임시 사용자 (5 Sample)", "정식 사용자 (2개월 무제한)"), index=0)
            
            if "정식" in s_role_selection:
                st.warning("⚠️ 정식 사용자 가입 안내")
                st.info("정식 사용자 신청 시 즉시 **임시 사용자** 권한이 부여됩니다.")
                st.info("관리자가 입금 확인 후 **정식 사용자**로 권한을 승인하며(승인 시점부터 2개월로 제한), 승인 완료 시 이메일로 알림을 보내 드립니다.")
            
            if st.button("가입신청"):
                if not validate_email(s_id): st.error("올바른 이메일 형식이 아닙니다.")
                elif not validate_password(s_pw): st.error("비밀번호는 문자+특수문자여야 합니다.")
                else:
                    initial_role = 'temp' 
                    if add_user(s_id.strip(), s_pw.strip(), initial_role):
                        if "정식" in s_role_selection:
                            send_application_email(s_id)
                            st.success("가입 신청 접수됨 (입금 확인 전까지 임시 권한 부여)")
                        else:
                            st.success("가입 완료!")
                    else:
                        st.error("이미 존재하는 아이디입니다.")
            
            # 회원가입 탭 내 서비스 이용 요금 안내
            st.markdown(fee_info_text)

        with tab_find_pw:
            st.header("🔑 비밀번호 찾기")
            st.write("가입 시 사용한 이메일 주소를 입력해주세요.")
            f_id = st.text_input("가입한 아이디 (이메일)", key="f_id")
            if st.button("비밀번호 이메일 전송"):
                if not f_id:
                    st.warning("이메일 주소를 입력해주세요.")
                else:
                    found_pw = get_user_password(f_id.strip())
                    if found_pw:
                        if send_password_recovery_email(f_id.strip(), found_pw):
                            st.success(f"'{f_id}'로 비밀번호를 전송했습니다.\n이메일을 확인해주세요.")
                        else:
                            st.error("이메일 전송 중 오류가 발생했습니다.")
                    else:
                        st.error("등록되지 않은 아이디입니다.")

    else:
        st.success(f"**{st.session_state.user_id}** 님")
        role_disp = "관리자" if st.session_state.user_role == 'admin' else ("정식 사용자" if st.session_state.user_role == 'official' else "임시 사용자")
        st.info(f"권한: {role_disp}")
        
        if st.session_state.expiry_date:
            st.warning(f"📅 사용 만료일: {st.session_state.expiry_date}")
        
        if st.session_state.user_role == 'admin':
            if st.button("🔧 관리자 화면 접속"):
                st.session_state.admin_mode = not st.session_state.admin_mode
                st.rerun()

        with st.expander("🔐 비밀번호 변경"):
            cur_pw = st.text_input("현재 비밀번호", type="password", key="chg_cur")
            new_pw = st.text_input("새 비밀번호", type="password", key="chg_new")
            confirm_pw = st.text_input("새 비밀번호 확인", type="password", key="chg_conf")
            
            if st.button("비밀번호 변경"):
                if new_pw != confirm_pw:
                    st.error("새 비밀번호가 일치하지 않습니다.")
                elif not validate_password(new_pw):
                    st.error("비밀번호는 4자 이상, 영문+특수문자를 포함해야 합니다.")
                else:
                    chk_res = check_login(st.session_state.user_id, cur_pw)
                    if chk_res:
                        change_user_password(st.session_state.user_id, new_pw)
                        st.success("비밀번호가 변경되었습니다.")
                    else:
                        st.error("현재 비밀번호가 올바르지 않습니다.")

        if st.button("로그아웃"):
            st.session_state.user_id = None
            st.session_state.user_role = None
            st.session_state.expiry_date = None
            st.session_state.admin_mode = False
            st.rerun()

    st.markdown("---")
    st.header("분석 설정")
    mean_method_label = st.radio("평균 산출 방식", ('기하평균 (Geometric)', '산술평균 (Arithmetic)'), index=0)
    mean_method = 'geometric' if '기하' in mean_method_label else 'arithmetic'
    cr_threshold = st.selectbox("일관성 비율(CR) 임계값", [0.1, 0.2], index=0)
    max_iter = st.number_input("최대 보정 반복 횟수", min_value=10, max_value=2000, value=500, step=50)

    st.markdown("---")
    with st.expander("ℹ️ 일관성 보정 안내", expanded=False):
        st.markdown("""
        **보정 방법: 반복 수렴 조정법(Iterative Adjustment)**
        판단 행렬이 비일관적(CR > 임계값)인 경우, 수학적으로 일관된 행렬과 원본 행렬을 일정 비율로 혼합하여 반복적으로 가중치를 미세 조정한 결과를 제시합니다.
        
        **현재 방법의 특징:**
        1. **최소 판단 왜곡**: 원본 설문 응답의 경향성을 최대한 보존하면서 수학적 일관성만을 확보합니다.
        2. **자동 수렴**: 설정된 반복 횟수 내에서 CR 값을 임계값 이하로 자동 개선합니다.
        """)

# =============================================================================
# 4. Main Content Logic
# =============================================================================

# CASE: Admin Mode
if st.session_state.get('admin_mode', False) and st.session_state.user_role == 'admin':
    st.subheader("👥 가입자 현황 및 관리")
    try:
        conn = sqlite3.connect('users.db')
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM visit_logs")
        total_visits = c.fetchone()[0]
        daily_df = pd.read_sql_query("SELECT visit_date, COUNT(*) as count FROM visit_logs GROUP BY visit_date ORDER BY visit_date ASC", conn)
        conn.close()
        st.write(f"**총 누적 방문자 수:** {total_visits:,}명")
        st.write("#### 📅 일별 방문자 현황")
        if not daily_df.empty:
            st.bar_chart(daily_df.set_index("visit_date"))
        else:
            st.info("방문 기록이 없습니다.")
    except Exception as e:
        st.error(f"통계 오류: {e}")
    st.divider()
    
    users_df = get_all_users()
    st.dataframe(users_df)

    with st.expander("회원 정보 수정 (비밀번호 초기화 포함)"):
        edit_id = st.selectbox("수정할 회원 ID", users_df['id'].unique())
        selected_user = users_df[users_df['id'] == edit_id].iloc[0]
        new_role = st.selectbox("권한 변경", ['temp', 'official', 'admin'], 
                                index=['temp', 'official', 'admin'].index(selected_user['role']))
        
        # 관리자가 'official'로 변경할 때 2개월 기한 제안
        if new_role == 'official' and selected_user['role'] != 'official':
            suggested_date = datetime.date.today() + relativedelta(months=2)
            new_expiry = st.text_input("만료일 설정 (YYYY-MM-DD) - 2개월 기한 자동 제안됨", value=str(suggested_date))
        else:
            new_expiry = st.text_input("만료일 변경 (YYYY-MM-DD)", value=selected_user['expiry_date'])
            
        new_pw = st.text_input("새 비밀번호 (입력 시 변경됨)", type="password", placeholder="변경하지 않으려면 비워두세요")
        
        if st.button("정보 수정 적용"):
            update_user_full_info(edit_id, new_pw, new_role, new_expiry)
            if new_role == 'official' and selected_user['role'] != 'official':
                send_approval_email(edit_id)
            st.success(f"{edit_id} 회원의 정보가 수정되었습니다.")
            st.rerun()
    
    with st.expander("회원 삭제"):
        del_id = st.selectbox("삭제할 회원 ID 선택", users_df['id'].unique(), key='del_user_select')
        if st.button("선택한 회원 삭제"):
            if del_id == st.session_state.user_id:
                st.error("본인은 삭제할 수 없습니다.")
            else:
                delete_user(del_id)
                st.success("삭제 완료")
                st.rerun()
    st.divider()

# CASE: Analysis View (Everyone)
st.subheader("1. AHP 분석 모델 설정 및 입력 템플릿 다운로드")

if st.session_state.user_id is None:
    st.info("🔒 **로그인 후** '나만의 분석 모델'을 만들 수 있습니다. (비로그인 상태에서는 기본 모델 및 샘플 데이터만 제공)")
else:
    saved_model = load_user_model(st.session_state.user_id)
    default_main = "거버넌스, 계획타당성, 실현가능성, 사업효과"
    default_subs = {
        "거버넌스": "행정지원, 지역공동체, 총괄사업관리자",
        "계획타당성": "현안적정성, 대안적정성, 목표구체성",
        "실현가능성": "부지확보, 사업구체화, 사업비적정성",
        "사업효과": "경제적효과, 사회적효과, 성과관리"
    }
    
    if saved_model:
        default_main = saved_model.get('main', default_main)
        default_subs = saved_model.get('subs', default_subs)

    with st.expander("📌 나의 분석 모델 만들기", expanded=False):
        st.info("대항목과 세부항목을 입력하여 나만의 입력 엑셀 템플릿을 생성하세요.\n\n현재 입력되어 있는 내용은 샘플 모델입니다. 삭제하시고 이용자님의 AHP 모델을 입력하세요.")
        main_criteria_input = st.text_input("대항목 (Main Criteria, 콤마 구분)", value=default_main)
        main_criteria_list = [x.strip() for x in main_criteria_input.split(',') if x.strip()]
        
        model_structure = {}
        if main_criteria_list:
            for mc in main_criteria_list:
                d_val = default_subs.get(mc, "")
                if isinstance(d_val, list): d_val = ", ".join(d_val)
                sub_input = st.text_input(f"'{mc}'의 세부항목", value=d_val, key=f"sub_{mc}")
                sub_list = [x.strip() for x in sub_input.split(',') if x.strip()]
                model_structure[mc] = sub_list
        
        if st.button("설정한 모델로 입력 엑셀 템플릿 생성"):
            if not main_criteria_list:
                st.error("대항목 입력 필요")
            else:
                current_model = {'main': main_criteria_input, 'subs': model_structure}
                save_user_model(st.session_state.user_id, current_model)
                st.toast("모델 저장 완료")
                
                output_template = io.BytesIO()
                with pd.ExcelWriter(output_template, engine='xlsxwriter') as writer:
                    main_pairs = list(itertools.combinations(main_criteria_list, 2))
                    main_cols = ["ID", "Type"] + [f"{a}_{b}" for a, b in main_pairs]
                    df_template_main = pd.DataFrame(columns=main_cols)
                    df_template_main.loc[0] = [1, ""] + [0]*len(main_pairs)
                    df_template_main.to_excel(writer, sheet_name="Main_Criteria", index=False)
                    
                    for mc, subs in model_structure.items():
                        if len(subs) < 2:
                            df_sub = pd.DataFrame(columns=["ID", "Type"])
                        else:
                            sub_pairs = list(itertools.combinations(subs, 2))
                            sub_cols = ["ID", "Type"] + [f"{a}_{b}" for a, b in sub_pairs]
                            df_sub = pd.DataFrame(columns=sub_cols)
                            df_sub.loc[0] = [1, ""] + [0]*len(sub_pairs)
                        safe_sheet_name = mc[:31]
                        df_sub.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                output_template.seek(0)
                st.download_button(
                    label="📥 엑셀 템플릿 다운로드",
                    data=output_template,
                    file_name="AHP_Custom_Template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.markdown("""
                ---
                ### 📝 데이터 입력 가이드
                1. **엑셀 파일 열기**: 위 버튼을 눌러 다운로드한 엑셀 파일을 실행합니다.
                2. **쌍대비교 데이터 입력**:
                    - **왼쪽** 항목이 더 중요하면: **음수** 입력 (예: -3)
                    - **오른쪽** 항목이 더 중요하면: **양수** 입력 (예: 3)
                    - **동등**하면: `1` 입력
                3. **필수 정보 입력**: A열(ID), **B열(Type)에 그룹명 입력 (예: 전문가, 주민 등)**
                """)
                if os.path.exists("ahp_input_guide.png"):
                    st.image("ahp_input_guide.png", caption="[참고] 설문 응답을 엑셀에 입력하는 방법")

st.markdown("---")

if st.session_state.user_role == 'official':
    with st.expander("📂 나의 분석 보관함"):
        my_analyses = get_user_analyses(st.session_state.user_id)
        if not my_analyses: st.info("저장된 분석 없음")
        else:
            for item in my_analyses:
                a_id, filename, save_date = item
                col_List1, col_List2, col_List3, col_List4 = st.columns([3, 2, 1, 1])
                with col_List1: st.text(f"{filename}")
                with col_List2: st.caption(f"{save_date}")
                with col_List3:
                    file_info = get_analysis_file(a_id)
                    if file_info:
                        fname, fdata = file_info
                        st.download_button("⬇️", fdata, fname, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_{a_id}")
                with col_List4:
                    if st.button("🗑️", key=f"del_{a_id}"):
                        delete_analysis(a_id)
                        st.rerun()

with st.container(border=True):
    st.markdown("#### ⚡ 빠른 시작 (도시재생 뉴딜사업 모델)")
    st.info("아래 버튼을 누르면 테스트용 샘플 엑셀 파일이 다운로드 됩니다.\n\n"
            "다운받은 테스트 샘플 엑셀 파일을 아래 2. 데이터 업로드 및 분석에 드롭다운 하거나 파일을 찾아 업로드 하세요.")
    
    sample_excel = create_sample_excel()
    st.download_button(
        label="📂 테스트용 샘플 엑셀 다운로드 (CR > 0.3, 5 Rows)",
        data=sample_excel,
        file_name="AHP_UrbanRegeneration_Sample.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("---")

# 엑셀 병합 출력 및 병합 서식 함수
def write_custom_ahp_table(writer, sheet_name, df, title_text, start_row, formats):
    workbook = writer.book
    if sheet_name in writer.sheets: worksheet = writer.sheets[sheet_name]
    else:
        worksheet = workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = worksheet
    
    header_fmt = formats['header']
    merge_fmt = formats['merge']
    body_fmt = formats['body']
    num_fmt = formats['num']
    sum_row_fmt = formats['sum_row']
    
    worksheet.merge_range(start_row, 0, start_row, 6, title_text, workbook.add_format({'bold': True, 'font_size': 12}))
    start_row += 1
    
    headers = ["대분류", "가중치(a)", "중분류", "가중치(b)", "종합 가중치(a x b)", "종합 순위", "비고"]
    for col, h in enumerate(headers):
        worksheet.write(start_row, col, h, header_fmt)
    start_row += 1
    
    main_criteria = df['대분류'].unique()
    current_row = start_row
    
    for main_c in main_criteria:
        sub_df = df[df['대분류'] == main_c]
        n_subs = len(sub_df)
        main_w = sub_df.iloc[0]['대분류 가중치']
        sub_cr = sub_df.iloc[0]['CR(중분류)']
        sum_sub_w = sub_df['중분류 가중치'].sum()
        
        merge_span = n_subs + 2 
        if merge_span > 1:
            worksheet.merge_range(current_row, 0, current_row + merge_span - 1, 0, main_c, merge_fmt)
            worksheet.merge_range(current_row, 1, current_row + merge_span - 1, 1, main_w, num_fmt)
        else:
            worksheet.write(current_row, 0, main_c, merge_fmt)
            worksheet.write(current_row, 1, main_w, num_fmt)
            
        for idx, row in sub_df.iterrows():
            worksheet.write(current_row, 2, row['중분류'], body_fmt)
            worksheet.write(current_row, 3, row['중분류 가중치'], num_fmt)
            worksheet.write(current_row, 4, row['Global Weight'], num_fmt)
            worksheet.write(current_row, 5, row['Global Rank'], body_fmt)
            worksheet.write(current_row, 6, "", body_fmt)
            current_row += 1
        
        worksheet.write(current_row, 2, "합계", sum_row_fmt)
        worksheet.write(current_row, 3, sum_sub_w, formats['sum_val'])
        worksheet.write_blank(current_row, 4, "", sum_row_fmt)
        worksheet.write_blank(current_row, 5, "", sum_row_fmt)
        worksheet.write_blank(current_row, 6, "", sum_row_fmt)
        current_row += 1
        
        worksheet.write(current_row, 2, "일관성 비율(CR)", sum_row_fmt)
        worksheet.write(current_row, 3, sub_cr, formats['num_sum'])
        worksheet.write_blank(current_row, 4, "", sum_row_fmt)
        worksheet.write_blank(current_row, 5, "", sum_row_fmt)
        worksheet.write_blank(current_row, 6, "", sum_row_fmt)
        current_row += 1

    worksheet.write(current_row, 0, "합계", sum_row_fmt)
    worksheet.write(current_row, 1, 1, formats['sum_val'])
    worksheet.write(current_row, 2, "합계", sum_row_fmt)
    worksheet.write_blank(current_row, 3, "", sum_row_fmt)
    worksheet.write(current_row, 4, 1, formats['sum_val'])
    worksheet.write_blank(current_row, 5, "", sum_row_fmt)
    worksheet.write_blank(current_row, 6, "", sum_row_fmt)
    
    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 12)
    worksheet.set_column('C:C', 25)
    worksheet.set_column('D:F', 12)
    return current_row + 2

def add_borders_to_data(worksheet, start_row, start_col, df, border_fmt, has_header=True, has_index=False):
    rows = len(df) + (1 if has_header else 0)
    cols = len(df.columns) + (1 if has_index else 0)
    worksheet.conditional_format(start_row, start_col, start_row+rows-1, start_col+cols-1,
                                   {'type': 'formula', 'criteria': '=TRUE', 'format': border_fmt})

st.subheader("2. 데이터 업로드 및 분석")
uploaded_file = st.file_uploader("작성된 엑셀 파일 업로드 (.xlsx)", type=['xlsx', 'xls'])

if uploaded_file:
    try:
        excel_obj = pd.ExcelFile(uploaded_file)
        sheet_names = excel_obj.sheet_names
        df_main = pd.read_excel(uploaded_file, sheet_name=sheet_names[0])
        main_cols = df_main.columns[2:]
        main_factors, n_main = infer_factors_from_columns(main_cols)

        permission_granted = False
        message = ""
        role = st.session_state.user_role
        user_id = st.session_state.user_id

        if role == 'admin' or role == 'official':
            permission_granted = True
            if role == 'official':
                today = datetime.date.today()
                expiry = datetime.datetime.strptime(st.session_state.expiry_date, "%Y-%m-%d").date()
                if today > expiry:
                    permission_granted = False
                    message = "⛔ 이용 기간이 만료되었습니다."
        else: 
            rows_ok = True
            for sn in sheet_names:
                if len(pd.read_excel(uploaded_file, sheet_name=sn)) > 5:
                    rows_ok = False
                    break
            if rows_ok: permission_granted = True
            else: message = f"⛔ **임시 사용자**는 시트당 최대 5개 표본까지만 분석 가능합니다."

        if permission_granted:
            with st.spinner("계층 분석 수행 중..."):
                main_results_df, main_factors = process_single_sheet(df_main, cr_threshold, max_iter, mean_method)
                main_sig_df = calculate_pairwise_ttest(main_results_df, main_factors)
                main_weight_cols = [f"Weight_{f}" for f in main_factors]
                
                if mean_method == 'arithmetic':
                    group_main_weights = main_results_df[main_weight_cols].mean(axis=0)
                else:
                    group_main_weights = gmean(main_results_df[main_weight_cols], axis=0)
                group_main_weights = group_main_weights / group_main_weights.sum()
                main_cr_final_avg = main_results_df['Final_CR'].mean()
                
                main_matrices = np.stack(main_results_df['Matrix_Object'].values)
                main_group_matrix = np.mean(main_matrices, axis=0) if mean_method == 'arithmetic' else gmean(main_matrices, axis=0)
                main_grp_cr, main_grp_ci, _ = calculate_consistency(main_group_matrix, mean_method)
                
                indiv_global_data = []
                all_ids = main_results_df['ID'].unique()
                
                sub_results_storage = {} 
                for i, sub_sheet_name in enumerate(sheet_names[1:]):
                    parent_factor = main_factors[i]
                    df_sub = pd.read_excel(uploaded_file, sheet_name=sub_sheet_name)
                    sub_res_df, sub_facts = process_single_sheet(df_sub, cr_threshold, max_iter, mean_method)
                    sub_sig_df = calculate_pairwise_ttest(sub_res_df, sub_facts)
                    sub_w_cols = [f"Weight_{f}" for f in sub_facts]
                    group_sub_w = sub_res_df[sub_w_cols].mean(axis=0) if mean_method == 'arithmetic' else gmean(sub_res_df[sub_w_cols], axis=0)
                    group_sub_w = group_sub_w / group_sub_w.sum()
                    sub_cr_final_avg = sub_res_df['Final_CR'].mean()
                    sub_matrices = np.stack(sub_res_df['Matrix_Object'].values)
                    sub_group_matrix = np.mean(sub_matrices, axis=0) if mean_method == 'arithmetic' else gmean(sub_matrices, axis=0)
                    sub_grp_cr, _, _ = calculate_consistency(sub_group_matrix, mean_method)
                    sub_results_storage[parent_factor] = {
                        'weights': group_sub_w, 'factors': sub_facts, 'cr': sub_cr_final_avg,
                        'df': sub_res_df, 'group_matrix': sub_group_matrix, 'group_cr': sub_grp_cr, 'sig_df': sub_sig_df
                    }

                for uid in all_ids:
                    u_main = main_results_df[main_results_df['ID'] == uid]
                    if u_main.empty: continue
                    u_type = u_main['Type'].values[0]
                    for mf in main_factors:
                        m_w = u_main[f"Weight_{mf}"].values[0]
                        s_df = sub_results_storage[mf]['df']
                        u_sub = s_df[s_df['ID'] == uid]
                        if u_sub.empty: continue
                        for sf in sub_results_storage[mf]['factors']:
                            s_w = u_sub[f"Weight_{sf}"].values[0]
                            indiv_global_data.append({
                                "ID": uid, "Type": str(u_type), "Factor": sf, "Global_Weight": m_w * s_w
                            })
                indiv_df = pd.DataFrame(indiv_global_data)
                
                anova_df = pd.DataFrame()
                if not indiv_df.empty and len(indiv_df['Type'].unique()) >= 2:
                    anova_df = calculate_anova_and_posthoc(indiv_df)

                summary_rows = []
                for idx, main_f in enumerate(main_factors):
                    m_weight = group_main_weights[idx]
                    sub_info = sub_results_storage[main_f]
                    for s_idx, sub_f in enumerate(sub_info['factors']):
                        s_weight = sub_info['weights'][s_idx]
                        global_w = m_weight * s_weight
                        summary_rows.append({
                            "대분류": main_f, "대분류 가중치": m_weight, "중분류": sub_f, "중분류 가중치": s_weight,
                            "Global Weight": global_w, "CR(대분류)": main_cr_final_avg, "CR(중분류)": sub_info['cr']
                        })
                
                final_df = pd.DataFrame(summary_rows)
                final_df['Global Rank'] = final_df['Global Weight'].rank(ascending=False, method='min').astype(int)
                cols_order = ["대분류", "대분류 가중치", "중분류", "중분류 가중치", "Global Weight", "Global Rank", "CR(대분류)", "CR(중분류)"]
                final_df = final_df[cols_order]

                unique_groups = sorted(main_results_df['Type'].astype(str).unique())
                group_analysis_results = {}
                group_full_dfs = {} 
                
                for grp in unique_groups:
                    grp_main_df = main_results_df[main_results_df['Type'].astype(str) == grp]
                    if grp_main_df.empty: continue
                    g_main_w = grp_main_df[main_weight_cols].mean(axis=0) if mean_method == 'arithmetic' else gmean(grp_main_df[main_weight_cols], axis=0)
                    g_main_w = g_main_w / g_main_w.sum()
                    g_main_mats = np.stack(grp_main_df['Matrix_Object'].values)
                    g_main_mat_obj = np.mean(g_main_mats, axis=0) if mean_method == 'arithmetic' else gmean(g_main_mats, axis=0)
                    g_main_cr, _, _ = calculate_consistency(g_main_mat_obj, mean_method)
                    
                    grp_rows = []
                    for idx, main_f in enumerate(main_factors):
                        m_w = g_main_w[idx]
                        full_sub_df = sub_results_storage[main_f]['df']
                        grp_sub_df = full_sub_df[full_sub_df['Type'].astype(str) == grp]
                        sub_facts = sub_results_storage[main_f]['factors']
                        if grp_sub_df.empty: continue
                        s_w_cols = [f"Weight_{f}" for f in sub_facts]
                        g_sub_w = grp_sub_df[s_w_cols].mean(axis=0) if mean_method == 'arithmetic' else gmean(grp_sub_df[s_w_cols], axis=0)
                        g_sub_w = g_sub_w / g_sub_w.sum()
                        g_sub_mats = np.stack(grp_sub_df['Matrix_Object'].values)
                        g_sub_mat_obj = np.mean(g_sub_mats, axis=0) if mean_method == 'arithmetic' else gmean(g_sub_mats, axis=0)
                        g_sub_cr, _, _ = calculate_consistency(g_sub_mat_obj, mean_method)
                        for s_idx, sf in enumerate(sub_facts):
                            grp_rows.append({
                                "대분류": main_f, "대분류 가중치": m_w, "중분류": sf, "중분류 가중치": g_sub_w[s_idx],
                                "Global Weight": m_w * g_sub_w[s_idx], "CR(대분류)": g_main_cr, "CR(중분류)": g_sub_cr
                            })
                    g_df = pd.DataFrame(grp_rows)
                    if not g_df.empty:
                        g_df['Global Rank'] = g_df['Global Weight'].rank(ascending=False, method='min').astype(int)
                        group_full_dfs[grp] = g_df[cols_order]
                        group_analysis_results[grp] = group_full_dfs[grp][['중분류', 'Global Weight']]

                comparison_df = final_df[['중분류', 'Global Weight']].copy()
                comparison_df.rename(columns={'Global Weight': 'Overall'}, inplace=True)
                for grp, df_res in group_analysis_results.items():
                    temp_df = df_res.rename(columns={'Global Weight': grp})
                    comparison_df = comparison_df.merge(temp_df, on='중분류', how='left')

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    formats = {
                        'header': workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#000000', 'font_color': '#FFFFFF', 'border': 1}),
                        'merge': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}),
                        'body': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}),
                        'num': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.000'}),
                        'sum_row': workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'align': 'center', 'valign': 'vcenter', 'border': 1}),
                        'sum_val': workbook.add_format({'num_format': '0', 'bg_color': '#D3D3D3', 'border': 1, 'align':'center'}),
                        'num_sum': workbook.add_format({'num_format': '0.000', 'bg_color': '#D3D3D3', 'border': 1, 'align':'center'})
                    }
                    border_fmt = workbook.add_format({'border': 1})
                    fmt_float_no_border = workbook.add_format({'num_format': '0.000', 'align': 'center', 'valign': 'vcenter'})
                    fmt_plain_no_border = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
                    fmt_diagonal = workbook.add_format({'num_format': '0', 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E7E6E6'})

                    current_row = write_custom_ahp_table(writer, '종합분석', final_df, "1) 전체_종합결과", 1, formats)
                    for grp in unique_groups:
                        if grp in group_full_dfs:
                            current_row = write_custom_ahp_table(writer, '종합분석', group_full_dfs[grp], f"▶ [그룹: {grp}] 분석 결과", current_row, formats)

                    comparison_df.to_excel(writer, sheet_name='Group_Comparison', index=False)
                    ws_comp = writer.sheets['Group_Comparison']
                    add_borders_to_data(ws_comp, 0, 0, comparison_df, border_fmt)

                    if not anova_df.empty:
                        anova_df.to_excel(writer, sheet_name='Statistical_Test', index=False)
                        ws_anova = writer.sheets['Statistical_Test']
                        add_borders_to_data(ws_anova, 0, 0, anova_df, border_fmt)

                    # 그룹별 종합 행렬표 엑셀 출력 기능
                    def write_detailed_sheet(sheet_name, matrix_df, detail_df, matrix_title, row_labels, group_matrices=None):
                        ws = workbook.add_worksheet(sheet_name)
                        writer.sheets[sheet_name] = ws
                        s_row = 0
                        ws.write_string(s_row, 0, matrix_title)
                        s_row += 1
                        
                        pd.DataFrame(matrix_df, index=row_labels, columns=row_labels).to_excel(writer, sheet_name=sheet_name, startrow=s_row)
                        for r in range(len(matrix_df)):
                            for c in range(len(matrix_df)):
                                val = 1 if r==c else matrix_df[r][c]
                                ws.write(s_row+r+1, c+1, val, border_fmt if r!=c else fmt_diagonal)
                                if r!=c: ws.write(s_row+r+1, c+1, val, fmt_float_no_border)
                        s_row += len(matrix_df) + 3

                        if group_matrices:
                            for g_name, g_mat in group_matrices.items():
                                ws.write_string(s_row, 0, f"] 그룹 종합 행렬: {g_name}")
                                s_row += 1
                                pd.DataFrame(g_mat, index=row_labels, columns=row_labels).to_excel(writer, sheet_name=sheet_name, startrow=s_row)
                                for r in range(len(g_mat)):
                                    for c in range(len(g_mat)):
                                        val = 1 if r==c else g_mat[r][c]
                                        ws.write(s_row+r+1, c+1, val, border_fmt if r!=c else fmt_diagonal)
                                        if r!=c: ws.write(s_row+r+1, c+1, val, fmt_float_no_border)
                                s_row += len(g_mat) + 3

                        detail_df.to_excel(writer, sheet_name=sheet_name, startrow=s_row, index=False)
                        add_borders_to_data(ws, s_row, 0, detail_df, border_fmt)

                    main_group_mats = {}
                    for grp in unique_groups:
                        g_df = main_results_df[main_results_df['Type'].astype(str) == grp]
                        if not g_df.empty:
                            mats = np.stack(g_df['Matrix_Object'].values)
                            main_group_mats[grp] = np.mean(mats, axis=0) if mean_method == 'arithmetic' else gmean(mats, axis=0)

                    out_main = main_results_df.drop(columns=['Matrix_Object'], errors='ignore')
                    write_detailed_sheet('Result_Main', main_group_matrix, out_main, f"[1] 전체 종합 행렬", main_factors, group_matrices=main_group_mats)
                    
                    for mf, info in sub_results_storage.items():
                        safe_name = f"Result_{mf}"[:31]
                        sub_grp_mats = {}
                        for grp in unique_groups:
                            g_sub_df = info['df'][info['df']['Type'].astype(str) == grp]
                            if not g_sub_df.empty:
                                mats = np.stack(g_sub_df['Matrix_Object'].values)
                                sub_grp_mats[grp] = np.mean(mats, axis=0) if mean_method == 'arithmetic' else gmean(mats, axis=0)
                        out_sub = info['df'].drop(columns=['Matrix_Object'], errors='ignore')
                        write_detailed_sheet(safe_name, info['group_matrix'], out_sub, f"[1] 전체 종합 행렬", info['factors'], group_matrices=sub_grp_mats)

                st.success("분석이 완료되었습니다!")
                if st.session_state.user_role == 'official':
                    save_analysis_to_db(st.session_state.user_id, f"{uploaded_file.name.split('.')[0]}_Result.xlsx", output.getvalue())

                tab1, tab2, tab3, tab4, tab5 = st.tabs(["🌐 종합 분석 (Global)", "👨‍👩‍👧‍👦 그룹별 분석", "🧪 통계 검정 (ANOVA)", "📊 시각화 센터", "📑 상세 데이터"])
                with tab1:
                    st.subheader("🌐 종합 중요도 및 순위")
                    st.dataframe(final_df.style.format(precision=3), use_container_width=True)
                with tab2:
                    st.markdown("#### 그룹별 가중치 상세 비교")
                    st.dataframe(comparison_df.style.format(precision=4), use_container_width=True)
                with tab3:
                    st.markdown("#### 집단 간 유의성 분석")
                    if not anova_df.empty: st.dataframe(anova_df.style.format(precision=5), use_container_width=True)
                    else: st.info("통계 검정을 위해 2개 이상의 그룹 데이터가 필요합니다.")
                with tab4:
                    st.markdown("#### 📊 시각화 센터")
                    col_chart1, col_chart2 = st.columns(2)
                    with col_chart1:
                        st.write("**종합 중요도 (Bar)**")
                        fig_bar = px.bar(final_df.sort_values('Global Weight'), y='중분류', x='Global Weight', orientation='h', text_auto='.3f')
                        st.plotly_chart(fig_bar, use_container_width=True)
                    with col_chart2:
                        st.write("**그룹별 중요도 패턴 (Radar)**")
                        radar_plot_df = indiv_global_data = []
                        all_ids = main_results_df['ID'].unique()
                        for rid in all_ids:
                            m_row = main_results_df[main_results_df['ID'] == rid].iloc[0]
                            rtype = m_row['Type']
                            for m_f in main_factors:
                                mw_indiv = m_row[f"Weight_{m_f}"]
                                s_row = sub_results_storage[m_f]['df'][sub_results_storage[m_f]['df']['ID'] == rid].iloc[0]
                                for s_f in sub_results_storage[m_f]['factors']:
                                    indiv_global_data.append({"Type": rtype, "Factor": s_f, "Global_Weight": mw_indiv * s_row[f"Weight_{s_f}"]})
                        radar_indiv_df = pd.DataFrame(indiv_global_data)
                        radar_plot_df = radar_indiv_df.groupby(['Type', 'Factor'])['Global_Weight'].mean().reset_index()
                        fig_radar = go.Figure()
                        for t in radar_plot_df['Type'].unique():
                            t_data = radar_plot_df[radar_plot_df['Type'] == t]
                            fig_radar.add_trace(go.Scatterpolar(r=t_data['Global_Weight'], theta=t_data['Factor'], fill='toself', name=t))
                        st.plotly_chart(fig_radar, use_container_width=True)

                    st.markdown("---")
                    st.write("**3. 일관성 비율(CR) 분포도 (Violin/Box Plot)**")
                    cr_dist_data = main_results_df[['ID', 'Type', 'Final_CR']].copy()
                    cr_dist_data['Level'] = '대분류'
                    for m_f in main_factors:
                        temp_cr = sub_results_storage[m_f]['df'][['ID', 'Type', 'Final_CR']].copy()
                        temp_cr['Level'] = f'중분류({m_f})'
                        cr_dist_data = pd.concat([cr_dist_data, temp_cr])
                    fig_cr_dist = px.violin(cr_dist_data, y="Final_CR", x="Level", color="Level", box=True, points="all", title="응답자별 일관성 지수 분포")
                    st.plotly_chart(fig_cr_dist, use_container_width=True)

                    st.markdown("---")
                    st.write("**4. 항목별 우선순위 산점도 (중요도 vs. 합의도)**")
                    scatter_df = radar_indiv_df.groupby('Factor')['Global_Weight'].agg(['mean', 'std']).reset_index()
                    scatter_df.columns = ['Factor', 'Weight_Mean', 'Weight_SD']
                    fig_scatter = px.scatter(scatter_df, x="Weight_Mean", y="Weight_SD", text="Factor", size="Weight_Mean", color="Weight_Mean",
                                            labels={'Weight_Mean': '중요도(평균)', 'Weight_SD': '의견차이(표준편차)'},
                                            title="중요도-합의도 분석 (우측 하단일수록 중요하고 합의된 항목)")
                    fig_scatter.update_traces(textposition='top center')
                    st.plotly_chart(fig_scatter, use_container_width=True)

                with tab5:
                    st.download_button("📥 결과 파일 다운로드 (Excel)", data=output.getvalue(), file_name="AHP_Result.xlsx")
                    st.dataframe(radar_indiv_df, use_container_width=True)
        else:
            st.warning(message)
    except Exception as e:
        st.error(f"오류 발생: {e}")

st.markdown("---")
st.caption("© 2026 AHP Analysis System. All rights reserved.")


