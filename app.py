import streamlit as st
# Force rebuild 2026-01-29 v5 (Sync & UI Polish)
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
from signup_agreement import show_agreement_ui, save_agreement_to_sheets, validate_all_agreements

# 1. 추가해야 할 라이브러리 (기존 Credentials 바로 아래 추가)
from streamlit_javascript import st_javascript
import base64

# IP 위치 추적 및 공인 IP 추출을 위한 라이브러리 추가
import requests

# [최적화 추가] 비동기 처리를 위한 스레딩 라이브러리
import threading

# ANOVA 및 사후검정을 위한 라이브러리 (없을 경우 예외처리)
try:
    from statsmodels.stats.multicomp import pairwise_tukeyhsd
    STATSMODELS_AVAILABLE = True
except ImportError:
    STATSMODELS_AVAILABLE = False

# =============================================================================
# 0. 시스템 설정 및 유틸리티
# =============================================================================

# [수정] Base64 문자열의 패딩 및 정제를 위한 유틸리티 함수 강화
def fix_base64_padding(data):
    """
    Base64 문자열의 패딩(Incorrect padding) 오류를 수정하는 함수
    """
    if isinstance(data, str):
        # 1. 모든 공백 및 줄바꿈 문자 제거 (가장 중요한 수정)
        data = re.sub(r'\s+', '', data)
        
        # 2. 패딩(=) 계산 및 추가
        missing_padding = len(data) % 4
        if missing_padding:
            data += '=' * (4 - missing_padding)
    return data

# [수정 반영] 1) SEO 태그 삽입, 2) 서비스 명 변경(AHP 마스터), 4) 파비콘 설정
try:
    logo_path = "ahp_master_logo.png"
    if os.path.exists(logo_path):
        logo_img = Image.open(logo_path)
    else:
        logo_img = "📊"
    
    st.set_page_config(
        page_title="AHP 마스터", 
        layout="wide", 
        page_icon=logo_img,
        menu_items={
            'Get Help': None,
            'Report a bug': None,
            'About': "AHP 마스터 - 스마트 의사결정 분석 시스템"
        }
    )
except Exception:
    st.set_page_config(page_title="AHP 마스터", layout="wide", page_icon="📊")

# [수정 반영] 메타 코드가 화면에 노출되지 않도록 display:none 스타일을 추가한 SEO 태그
seo_tags = """
    <div style="display:none;">
        <head>
            <meta name="description" content="AHP 마스터 - 학위 논문 및 정책 연구를 위한 최적의 AHP 분석 자동화 솔루션. 일관성 비율(CR) 자동 보정 및 통계 검정 제공.">
            <meta name="keywords" content="AHP, 무료, 프로그램, AHP분석, 계층분석과정, 일관성보정, CR보정, 학위논문통계, AHP계산기, AHP 마스터">
            <meta name="author" content="AHP Master">
            <meta property="og:title" content="AHP 마스터: 분석 자동화 시스템">
            <meta property="og:description" content="수학적 일관성 보정과 고도화된 통계 분석을 지원하는 최신 AHP 전문 도구">
        </head>
    </div>
"""
st.markdown(seo_tags, unsafe_allow_html=True)

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
    except Exception:
        pass
    plt.rcParams['axes.unicode_minus'] = False 

set_font_config()

# [중요 수정] 구글 시트 연결 헬퍼 함수 - 인증 정보 로드 로직 전면 재검토 및 수정
# TOML(Dict), JSON String, Base64 Encoded String 등 다양한 포맷에 대응하도록 강화
def get_gspread_client():
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    
    # st.secrets에서 값 가져오기 (없을 경우 에러 처리)
    if "gcp_service_account" not in st.secrets:
        st.error("Secrets에 'gcp_service_account' 설정이 없습니다.")
        return None

    raw_auth = st.secrets["gcp_service_account"]
    auth_info = {}

    # Case 1: 이미 딕셔너리 형태인 경우 (TOML 포맷) - 가장 일반적인 경우
    if isinstance(raw_auth, dict) or hasattr(raw_auth, "keys"): 
        auth_info = dict(raw_auth) # AttrDict 등을 dict로 변환
    
    # Case 2: 문자열 형태인 경우 (JSON 문자열 혹은 Base64 인코딩 문자열)
    elif isinstance(raw_auth, str):
        # 앞뒤 공백 및 따옴표 제거
        auth_str = raw_auth.strip().strip('"').strip("'")
        
        try:
            # 2-1. 순수 JSON 문자열로 파싱 시도
            auth_info = json.loads(auth_str)
        except json.JSONDecodeError:
            # 2-2. JSON 파싱 실패 -> Base64 인코딩된 값으로 가정하고 디코딩 시도
            try:
                # 1단계: 문자열 정제 (모든 공백 제거)
                clean_b64 = re.sub(r'\s+', '', auth_str)
                
                # 2단계: 패딩(=) 보정
                missing_padding = len(clean_b64) % 4
                if missing_padding:
                    clean_b64 += '=' * (4 - missing_padding)
                
                # 3단계: Base64 디코딩 (Standard 및 URL-Safe 방식 모두 시도)
                try:
                    decoded_bytes = base64.b64decode(clean_b64)
                except Exception:
                    # Standard 실패 시 URL-Safe 방식 시도 (-와 _ 문자 처리)
                    decoded_bytes = base64.urlsafe_b64decode(clean_b64)
                    
                decoded_info = decoded_bytes.decode('utf-8')
                auth_info = json.loads(decoded_info)
            except Exception as e:
                st.error(f"서비스 계정 키 디코딩 실패 (Base64/JSON 오류): {e}")
                return None
    else:
        st.error("gcp_service_account 형식을 인식할 수 없습니다.")
        return None

    # [중요] Private Key 내의 줄바꿈 문자(\n) 처리
    # TOML 등에서 문자열로 읽어올 때 \\n으로 이스케이프된 경우 실제 줄바꿈으로 변경 필요
    if auth_info and "private_key" in auth_info:
        auth_info["private_key"] = auth_info["private_key"].replace("\\n", "\n")

    # 필수 필드 확인 (Missing fields 에러 방지)
    required_fields = ["private_key", "client_email", "token_uri"]
    missing = [f for f in required_fields if f not in auth_info]
    if missing:
        st.error(f"서비스 계정 정보에 필수 필드가 누락되었습니다: {', '.join(missing)}")
        return None

    creds = Credentials.from_service_account_info(auth_info, scopes=scope)
    return gspread.authorize(creds)

# DB 초기화 및 구글 시트로부터 데이터(회원+방문로그) 복구 로직
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
    
    # [커뮤니티 기능 테이블 수정/추가]
    # views: 조회수 카운트 컬럼 추가
    c.execute('''CREATE TABLE IF NOT EXISTS community_posts
                  (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id TEXT, title TEXT, content TEXT, reg_date TEXT, is_secret INTEGER, is_notice INTEGER, likes INTEGER DEFAULT 0, non_user_pw TEXT, views INTEGER DEFAULT 0)''')
    
    try:
        c.execute("ALTER TABLE community_posts ADD COLUMN non_user_pw TEXT")
    except sqlite3.OperationalError:
        pass
    
    # [요청사항 3] 조회수 컬럼 추가 로직
    try:
        c.execute("ALTER TABLE community_posts ADD COLUMN views INTEGER DEFAULT 0")
    except sqlite3.OperationalError:
        pass

    c.execute('''CREATE TABLE IF NOT EXISTS community_comments
                (id INTEGER PRIMARY KEY AUTOINCREMENT, post_id INTEGER, user_id TEXT, 
                content TEXT, reg_date TEXT, is_secret INTEGER)''')
    
    # [최적화 추가] DB 검색 성능 향상을 위한 인덱스 생성
    c.execute("CREATE INDEX IF NOT EXISTS idx_post_id ON community_comments(post_id)")
    
    # 관리자 계정 생성
    try:
        kst_now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))
        signup_date_str = kst_now.strftime("%Y-%m-%d")
        c.execute("INSERT OR IGNORE INTO users VALUES (?, ?, ?, ?, ?)", 
                  ('shjeon', '@jsh2143033', 'admin', signup_date_str, '9999-12-31'))
        conn.commit()
    except sqlite3.IntegrityError:
        pass 

    # [복구 로직 1] 회원 정보 복구
    c.execute("SELECT COUNT(*) FROM users")
    if c.fetchone()[0] <= 1:
        try:
            client = get_gspread_client()
            if client:
                spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
                sheet = spreadsheet.sheet1 
                all_values = sheet.get_all_values()
                if len(all_values) > 1:
                    for row in all_values[1:]:
                        if row[0] == 'shjeon': continue
                        c.execute("INSERT OR IGNORE INTO users (id, role, signup_date, pw, expiry_date) VALUES (?, ?, ?, ?, ?)", 
                                  (row[0], row[1], row[2], row[3], '9999-12-31'))
                    conn.commit()
        except Exception:
            pass

    # [복구 로직 2] 방문 로그 복구
    c.execute("SELECT COUNT(*) FROM visit_logs")
    if c.fetchone()[0] == 0:
        try:
            client = get_gspread_client()
            if client:
                spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
                try:
                    visit_sheet = spreadsheet.worksheet("Visit_Logs")
                    records = visit_sheet.get_all_records()
                    for row in records:
                        c.execute("INSERT OR IGNORE INTO visit_logs (ip_address, visit_date) VALUES (?, ?)", 
                                  (row['IP'], row['Date']))
                    conn.commit()
                except gspread.exceptions.WorksheetNotFound:
                    pass
        except Exception:
            pass

    conn.close()
    
    # [요청사항 4] 어플 재부팅 시 구글 시트 내용(회원, 게시글, 댓글) 불러오기
    sync_db_from_sheets()

# [신규 기능 1 & 요청사항 4] 구글 시트의 내용을 강제로 DB에 동기화하는 함수
def sync_db_from_sheets():
    """구글 시트의 데이터를 읽어와 DB에 없는 데이터를 강제로 추가합니다. (회원, 게시글, 댓글)"""
    try:
        client = get_gspread_client()
        if not client: return -1
        
        spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
        conn = sqlite3.connect('users.db')
        c = conn.cursor()
        
        # 1. 회원 정보 동기화
        try:
            sheet = spreadsheet.sheet1
            all_values = sheet.get_all_values()
            if len(all_values) > 1:
                for row in all_values[1:]:
                    if len(row) >= 4:
                        user_id = row[0]
                        role = row[1]
                        signup_date = row[2]
                        pw = row[3]
                        expiry_date = '9999-12-31' 
                        c.execute("INSERT OR IGNORE INTO users (id, pw, role, signup_date, expiry_date) VALUES (?, ?, ?, ?, ?)", 
                                  (user_id, pw, role, signup_date, expiry_date))
        except: pass

        # 2. 게시글 동기화 (Community_Posts)
        try:
            post_sheet = spreadsheet.worksheet("Community_Posts")
            posts = post_sheet.get_all_values()
            # Header: ID, UserID, Title, Content, RegDate, IsSecret, IsNotice, Likes, NonUserPW, Views
            if len(posts) > 1:
                for row in posts[1:]:
                    if len(row) >= 8: # 최소 필드 확보
                        p_id = row[0]
                        u_id = row[1]
                        ttl = row[2]
                        cnt = row[3]
                        reg = row[4]
                        sec = int(row[5]) if row[5] else 0
                        notc = int(row[6]) if row[6] else 0
                        lks = int(row[7]) if row[7] else 0
                        npw = row[8] if len(row) > 8 else None
                        vws = int(row[9]) if len(row) > 9 and row[9] else 0
                        
                        c.execute("INSERT OR IGNORE INTO community_posts (id, user_id, title, content, reg_date, is_secret, is_notice, likes, non_user_pw, views) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                  (p_id, u_id, ttl, cnt, reg, sec, notc, lks, npw, vws))
        except gspread.exceptions.WorksheetNotFound:
            pass
            
        # 3. 댓글 동기화 (Community_Comments)
        try:
            com_sheet = spreadsheet.worksheet("Community_Comments")
            comments = com_sheet.get_all_values()
            # Header: ID, PostID, UserID, Content, RegDate, IsSecret
            if len(comments) > 1:
                for row in comments[1:]:
                    if len(row) >= 6:
                        c_id = row[0]
                        p_id = row[1]
                        u_id = row[2]
                        cnt = row[3]
                        reg = row[4]
                        sec = int(row[5]) if row[5] else 0
                        c.execute("INSERT OR IGNORE INTO community_comments (id, post_id, user_id, content, reg_date, is_secret) VALUES (?, ?, ?, ?, ?, ?)",
                                  (c_id, p_id, u_id, cnt, reg, sec))
        except gspread.exceptions.WorksheetNotFound:
            pass

        conn.commit()
        conn.close()
        # [최적화 추가] 게시판 데이터 변경 시 캐시 초기화
        st.cache_data.clear()
        return 1
    except Exception:
        return -1

# 방문자 추적 및 구글 시트 실시간 저장
def track_visitor():
    js_ip_script = 'await fetch("https://api.ipify.org?format=json").then(r => r.json()).then(d => d.ip)'
    client_ip = st_javascript(js_ip_script)
    if not client_ip:
        return 

    ip = str(client_ip).strip()
    
    if st.session_state.get('visited'):
        return

    try:
        now_ts = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M:%S")
        
        country, region, city, lat, lon = "", "", "", "", ""
        if ip not in ["localhost", "unknown_ip", "127.0.0.1"] and not ip.startswith("192.168."):
            try:
                response = requests.get(f"http://ip-api.com/json/{ip}", timeout=5)
                if response.status_code == 200:
                    data = response.json()
                    if data.get("status") == "success":
                        country = data.get("country", "")
                        region = data.get("regionName", "")
                        city = data.get("city", "")
                        lat = data.get("lat", "")
                        lon = data.get("lon", "")
            except:
                pass

        conn = sqlite3.connect('users.db')
        c = conn.cursor()
        c.execute("INSERT OR IGNORE INTO visit_logs (ip_address, visit_date) VALUES (?, ?)", (ip, now_ts))
        conn.commit()
        conn.close()

        try:
            client = get_gspread_client()
            if client:
                spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
                try:
                    visit_sheet = spreadsheet.worksheet("Visit_Logs")
                except gspread.exceptions.WorksheetNotFound:
                    visit_sheet = spreadsheet.add_worksheet(title="Visit_Logs", rows="1000", cols="10")
                    visit_sheet.append_row(["IP", "Date", "Country", "Region", "City", "Latitude", "Longitude"])
                
                existing_logs = visit_sheet.get_all_values()
                if [ip, now_ts] not in [row[:2] for row in existing_logs]:
                    # [최적화 추가] 방문 로그는 스레드 처리하여 응답성 향상
                    threading.Thread(target=visit_sheet.append_row, args=([ip, now_ts, country, region, city, lat, lon],)).start()
                
                st.session_state.visited = True
            
        except Exception:
            pass
    except Exception:
        pass

# 방문자 추적 실행부
if 'visited' not in st.session_state:
    st.session_state.visited = False
track_visitor()

def validate_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email)

def validate_password(password):
    if len(password) < 4: return False
    has_char = re.search(r'[a-zA-Z]', password)
    has_special = re.search(r'[!@#$%^&*(),.?":{}|<>]', password)
    return has_char and has_special

def send_application_email(user_email):
    sender_email = "jeon080423@gmail.com"
    password = "csuh xxru wqdy mttt"
    recipient_email = "jeon080423@gmail.com"
    subject = f"[AHP 마스터] 정식 사용자 승인 요청: {user_email}"
    kst_today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).date()
    body = f"사용자가 정식 권한 신청.\nID: {user_email}\n신청일: {kst_today}"
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
    except: pass

def send_conversion_request_email(user_email):
    sender_email = "jeon080423@gmail.com"
    password = "csuh xxru wqdy mttt"
    recipient_email = "jeon080423@gmail.com"
    subject = f"[AHP 마스터] 정식사용자 전환 요청: {user_email}"
    body = f"임시 사용자가 정식사용자로 전환 요청 했습니다\nID: {user_email}"
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

def send_approval_email(user_email):
    sender_email = "jeon080423@gmail.com"
    password = "csuh xxru wqdy mttt"
    recipient_email = user_email
    subject = "[AHP 마스터] 정식 사용자 승인 완료"
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
    subject = "[AHP 마스터] 비밀번호 안내"
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
    except Exception:
        return False

# --- DB CRUD ---

def log_to_sheets(user_id, role, signup_date, pw, agree_info="미동의"):
    try:
        client = get_gspread_client()
        if client:
            spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
            sheet = spreadsheet.sheet1
            
            headers = sheet.row_values(1)
            if "agree_info" not in headers:
                sheet.update_cell(1, 5, "agree_info")
            if "expiry_date" not in headers:
                sheet.update_cell(1, 6, "expiry_date")

            sheet.append_row([user_id, role, str(signup_date), pw, agree_info, "9999-12-31"])
    except Exception as e:
        st.error(f"Google Sheets 로깅 오류: {e}")

def add_user(user_id, pw, role, agree_info="미동의"):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    signup_date = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime("%Y-%m-%d")
    expiry_date = "9999-12-31"
    try:
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?, ?)", 
                  (user_id, pw, role, signup_date, expiry_date))
        conn.commit()
        log_to_sheets(user_id, role, signup_date, pw, agree_info)
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

    try:
        client = get_gspread_client()
        if client:
            spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
            sheet = spreadsheet.sheet1
            cell = sheet.find(user_id)
            if cell:
                sheet.update_cell(cell.row, 4, new_pw)
    except Exception:
        pass
    return True

def get_all_users():
    conn = sqlite3.connect('users.db')
    df = pd.read_sql_query("SELECT * FROM users", conn)
    conn.close()
    return df

def update_user_full_info(user_id, new_pw, new_role, new_expiry):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    if new_pw is not None and new_pw != "":
        c.execute("UPDATE users SET pw=?, role=?, expiry_date=? WHERE id=?", (new_pw, new_role, new_expiry, user_id))
    else:
        c.execute("UPDATE users SET role=?, expiry_date=? WHERE id=?", (new_role, new_expiry, user_id))
    conn.commit()
    conn.close()
    
    try:
        client = get_gspread_client()
        if client:
            spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
            sheet = spreadsheet.sheet1
            
            headers = sheet.row_values(1)
            if "expiry_date" not in headers:
                sheet.update_cell(1, 6, "expiry_date")

            cell = sheet.find(user_id)
            if cell:
                row_num = cell.row
                sheet.update_cell(row_num, 2, new_role)
                sheet.update_cell(row_num, 6, new_expiry)
                if new_pw is not None and new_pw != "":
                    sheet.update_cell(row_num, 4, new_pw)
    except Exception:
        pass 

def delete_user(user_id):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("DELETE FROM users WHERE id=?", (user_id,))
    c.execute("DELETE FROM saved_analyses WHERE user_id=?", (user_id,))
    c.execute("DELETE FROM user_models WHERE user_id=?", (user_id,))
    conn.commit()
    conn.close()

    try:
        client = get_gspread_client()
        if client:
            spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
            sheet = spreadsheet.sheet1
            
            try:
                del_sheet = spreadsheet.worksheet("Deleted_Users")
            except gspread.exceptions.WorksheetNotFound:
                del_sheet = spreadsheet.add_worksheet(title="Deleted_Users", rows="1000", cols="10")
                del_sheet.append_row(["ID", "Role", "SignupDate", "PW", "DeletedDate"])

            all_values = sheet.get_all_values()
            target_row_index = -1
            row_data = []
            for i, row in enumerate(all_values):
                if row[0] == user_id:
                    target_row_index = i + 1
                    row_data = row
                    break
            
            if target_row_index != -1:
                kst_now_ts = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M:%S")
                row_data.append(str(kst_now_ts))
                del_sheet.append_row(row_data)
                sheet.delete_rows(target_row_index)
    except Exception:
        pass

def restore_from_deleted_sheet(user_id):
    try:
        client = get_gspread_client()
        if client:
            spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
            try:
                del_sheet = spreadsheet.worksheet("Deleted_Users")
                cell = del_sheet.find(user_id)
                if cell:
                    del_sheet.delete_rows(cell.row)
            except (gspread.exceptions.WorksheetNotFound, gspread.exceptions.CellNotFound):
                pass
    except Exception:
        pass

def save_analysis_to_db(user_id, filename, file_data):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    save_date = str(datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M:%S"))
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
# Saaty(1980) AHP Functions
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

def improve_consistency(matrix, threshold, min_val, max_val, max_iter=500, learning_rate=0.4, method='geometric', allow_even=False):
    current_matrix = matrix.copy()
    n = current_matrix.shape[0]
    cr, ci, _ = calculate_consistency(current_matrix, method)
    iterations = 0
    if cr <= threshold: return current_matrix, cr, iterations, False
    
    triu_indices = np.triu_indices(n, k=1)
    
    for it in range(max_iter):
        if cr <= threshold: break
        
        w = calculate_weights(current_matrix, method)
        consistent_matrix = np.outer(w, 1/w)
        
        new_matrix = (current_matrix * (1 - learning_rate)) + (consistent_matrix * learning_rate)
        np.fill_diagonal(new_matrix, 1.0)
        
        vals = new_matrix[triu_indices]
        
        temp_raw = np.where(vals == 1.0, 1.0, 
                   np.where(vals > 1.0, -np.round(vals), 
                   np.round(1.0/vals)))
        
        temp_raw = np.clip(temp_raw, min_val, max_val)
        
        abs_raw = np.abs(temp_raw)
        signs = np.sign(temp_raw)
        
        if not allow_even:
            abs_raw = np.where((abs_raw % 2 == 0) & (abs_raw != 0), np.maximum(1, abs_raw - 1), abs_raw)
            
        temp_raw = np.where(temp_raw == 0, 1, (signs * abs_raw)).astype(int)
        
        final_vals = np.where(temp_raw == 0, 1.0,
                     np.where(temp_raw < 0, np.abs(temp_raw).astype(float),
                     np.where(temp_raw == 1, 1.0, 1.0 / temp_raw)))
        
        new_matrix[triu_indices] = final_vals
        new_matrix.T[triu_indices] = 1.0 / final_vals
        
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
    
    all_comp_values = df[comp_cols].values.flatten()
    sheet_min = int(np.min(all_comp_values))
    sheet_max = int(np.max(all_comp_values))
    
    has_even = np.any((np.abs(all_comp_values) % 2 == 0) & (np.abs(all_comp_values) > 1))
    
    results_list = []
    excluded_list = []
    excluded_count = 0
    for idx, row in df.iterrows():
        respondent_id = row.iloc[0]
        respondent_type = row.iloc[1]
        matrix = np.eye(n)
        
        raw_values = []
        col_idx = 0
        for i in range(n):
            for j in range(i + 1, n):
                if col_idx < len(comp_cols):
                    raw_val = row[comp_cols[col_idx]]
                    raw_values.append(raw_val)
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
                matrix, cr_threshold, sheet_min, sheet_max, max_iter=max_iter, method=method, allow_even=has_even
            )
        
        if final_cr > cr_threshold:
            excluded_count += 1
            ex_res = {"ID": respondent_id, "Type": respondent_type}
            for k, col_name in enumerate(comp_cols):
                ex_res[col_name] = raw_values[k]
            ex_res["CR"] = final_cr
            excluded_list.append(ex_res)
            continue

        final_raw_values = []
        for i in range(n):
            for j in range(i + 1, n):
                val = final_matrix[i, j]
                if val == 1.0: final_raw_val = 1
                elif val > 1.0: final_raw_val = -int(round(val)) 
                else: final_raw_val = int(round(1.0/val)) 
                final_raw_values.append(final_raw_val)

        _, final_ci, _ = calculate_consistency(final_matrix, method)
        final_weights = calculate_weights(final_matrix, method)
        
        res = {
            "ID": respondent_id,
            "Type": respondent_type
        }
        
        for k, col_name in enumerate(comp_cols):
            res[f"Raw_Orig_{col_name}"] = raw_values[k]
        
        res["Original_CI"] = orig_ci
        res["Original_CR"] = orig_cr
        
        for k, col_name in enumerate(comp_cols):
            res[f"Raw_Final_{col_name}"] = final_raw_values[k]
            
        res["Final_CI"] = final_ci
        res["Final_CR"] = final_cr
        
        res["Iterations"] = iterations
        res["Corrected"] = corrected_flag
        res["Matrix_Object"] = final_matrix 
        
        for f_idx, f_name in enumerate(factors):
            res[f"Weight_{f_name}"] = final_weights[f_idx]
            
        results_list.append(res)
        
    results_df = pd.DataFrame(results_list)
    excluded_df = pd.DataFrame(excluded_list)
    return results_df, factors, excluded_count, excluded_df

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
            "F-값": f_stat,
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
            except Exception:
                row["사후검정(Tukey HSD)"] = "계산 오류"
        
        results.append(row)
        
    return pd.DataFrame(results)

# -----------------------------------------------------------------------------
# [신규] 커뮤니티 데이터베이스 및 UI 로직 (구글 시트 연동 포함)
# -----------------------------------------------------------------------------

def log_community_to_sheets(sheet_name, data):
    try:
        client = get_gspread_client()
        if client:
            spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
            try:
                work_sheet = spreadsheet.worksheet(sheet_name)
            except gspread.exceptions.WorksheetNotFound:
                if sheet_name == "Community_Posts":
                    work_sheet = spreadsheet.add_worksheet(title="Community_Posts", rows="1000", cols="10")
                    work_sheet.append_row(["ID", "UserID", "Title", "Content", "RegDate", "IsSecret", "IsNotice", "Likes", "NonUserPW", "Views"])
                else:
                    work_sheet = spreadsheet.add_worksheet(title="Community_Comments", rows="1000", cols="10")
                    work_sheet.append_row(["ID", "PostID", "UserID", "Content", "RegDate", "IsSecret"])
            
            # [요청사항 1] 헤더 유무 확인 및 강제 기록 (데이터 꼬임 방지)
            if sheet_name == "Community_Posts":
                if not work_sheet.row_values(1):
                    work_sheet.append_row(["ID", "UserID", "Title", "Content", "RegDate", "IsSecret", "IsNotice", "Likes", "NonUserPW", "Views"])
            
            if isinstance(data, list): 
                # [최적화 추가] 시트 저장을 비동기 처리하여 게시판 응답 속도 향상
                threading.Thread(target=work_sheet.append_row, args=(data,)).start()
    except Exception:
        pass

# [최적화 추가] 게시판 목록 읽기 로직 캐싱 적용
@st.cache_data(show_spinner=False)
def get_posts():
    conn = sqlite3.connect('users.db')
    df = pd.read_sql_query("SELECT * FROM community_posts ORDER BY is_notice DESC, id DESC", conn)
    conn.close()
    return df

# [요청사항 3] 조회수 증가 및 구글 시트 기록 함수
def increment_views(pid):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("UPDATE community_posts SET views = views + 1 WHERE id=?", (pid,))
    conn.commit()
    
    # 갱신된 조회수 가져오기
    c.execute("SELECT views FROM community_posts WHERE id=?", (pid,))
    new_views = c.fetchone()[0]
    conn.close()
    
    # [요청사항 1 & 3] 구글 시트 업데이트 (기존 데이터에 덧쓰지 않고 행 업데이트)
    try:
        client = get_gspread_client()
        if client:
            spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
            sheet = spreadsheet.worksheet("Community_Posts")
            cell = sheet.find(str(pid))
            if cell:
                # Views는 J열(10번째 열)에 위치한다고 가정 (헤더 순서 기반)
                # [최적화 추가] 비동기 업데이트
                threading.Thread(target=sheet.update_cell, args=(cell.row, 10, new_views)).start()
    except: pass
    
    # [최적화 추가] 조회수 변경 시 캐시 무효화
    st.cache_data.clear()

def add_post(uid, title, content, is_secret, is_notice, non_user_pw=None):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M:%S")
    c.execute("INSERT INTO community_posts (user_id, title, content, reg_date, is_secret, is_notice, non_user_pw, views) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
              (uid, title, content, now, 1 if is_secret else 0, 1 if is_notice else 0, non_user_pw, 0))
    post_id = c.lastrowid
    conn.commit()
    conn.close()
    
    # [요청사항 1] 구글 시트에 새 행으로 추가 (append_row 사용으로 기존 데이터 보존, 정확한 컬럼 매핑)
    log_community_to_sheets("Community_Posts", [post_id, uid, title, content, now, 1 if is_secret else 0, 1 if is_notice else 0, 0, non_user_pw, 0])
    
    # [최적화 추가] 데이터 추가 시 캐시 무효화
    st.cache_data.clear()

def update_post(pid, title, content, is_secret, is_notice):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("UPDATE community_posts SET title=?, content=?, is_secret=?, is_notice=? WHERE id=?", 
              (title, content, 1 if is_secret else 0, 1 if is_notice else 0, pid))
    conn.commit()
    conn.close()
    
    try:
        client = get_gspread_client()
        if client:
            spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
            sheet = spreadsheet.worksheet("Community_Posts")
            cell = sheet.find(str(pid))
            if cell:
                # 덧쓰지 않고 특정 열 범위만 업데이트 (C:Title ~ G:IsNotice)
                # [최적화 추가] 비동기 처리
                threading.Thread(target=sheet.update, kwargs={"range_name": f'C{cell.row}:G{cell.row}', "values": [[title, content, "", 1 if is_secret else 0, 1 if is_notice else 0]]}).start()
    except: pass
    
    # [최적화 추가] 캐시 무효화
    st.cache_data.clear()

def delete_post(pid):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("DELETE FROM community_posts WHERE id=?", (pid,))
    c.execute("DELETE FROM community_comments WHERE post_id=?", (pid,))
    conn.commit()
    conn.close()
    
    try:
        client = get_gspread_client()
        if client:
            spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
            sheet = spreadsheet.worksheet("Community_Posts")
            cell = sheet.find(str(pid))
            if cell: 
                # [최적화 추가] 비동기 삭제
                threading.Thread(target=sheet.delete_rows, args=(cell.row,)).start()
    except: pass
    
    # [최적화 추가] 캐시 무효화
    st.cache_data.clear()

# [최적화 추가] 댓글 읽기 캐싱 적용
@st.cache_data(show_spinner=False)
def get_comments(pid):
    conn = sqlite3.connect('users.db')
    df = pd.read_sql_query(f"SELECT * FROM community_comments WHERE post_id={pid} ORDER BY id ASC", conn)
    conn.close()
    return df

def add_comment(pid, uid, content, is_secret):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).strftime("%Y-%m-%d %H:%M:%S")
    c.execute("INSERT INTO community_comments (post_id, user_id, content, reg_date, is_secret) VALUES (?, ?, ?, ?, ?)",
              (pid, uid, content, now, 1 if is_secret else 0))
    com_id = c.lastrowid
    conn.commit()
    conn.close()
    log_community_to_sheets("Community_Comments", [com_id, pid, uid, content, now, 1 if is_secret else 0])
    
    # [최적화 추가] 캐시 무효화
    st.cache_data.clear()

def like_post(pid):
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("UPDATE community_posts SET likes = likes + 1 WHERE id=?", (pid,))
    conn.commit()
    conn.close()
    
    try:
        client = get_gspread_client()
        if client:
            spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
            sheet = spreadsheet.worksheet("Community_Posts")
            cell = sheet.find(str(pid))
            if cell:
                curr_likes = sheet.cell(cell.row, 8).value
                # [최적화 추가] 비동기 처리
                threading.Thread(target=sheet.update_cell, args=(cell.row, 8, int(curr_likes or 0) + 1)).start()
    except: pass
    
    # [최적화 추가] 캐시 무효화
    st.cache_data.clear()

# [요청사항 2] 커뮤니티 게시판 UI 수정 (이미지 형태의 목록형 -> 클릭 시 내용 표시)
def show_community_board():
    st.header("💬 커뮤니티 게시판")
    if st.button("⬅️ 분석 시스템으로 돌아가기"):
        st.session_state.page = "main"
        st.rerun()

    posts = get_posts()
    tab_list, tab_write = st.tabs(["글 목록", "🖋️ 글 쓰기"])

    with tab_list:
        if posts.empty:
            st.info("등록된 게시글이 없습니다.")
        else:
            # [요청사항 2] 표 형식 헤더 (번호, 제목, 이름, 날짜, 조회, 추천, 댓글) - 댓글 추가
            # 레이아웃 비율 설정
            h_col1, h_col2, h_col3, h_col4, h_col5, h_col6, h_col7 = st.columns([0.8, 5, 1.5, 1.5, 0.8, 0.8, 0.8])
            h_col1.markdown("**번호**")
            h_col2.markdown("**제목**")
            h_col3.markdown("**이름**")
            h_col4.markdown("**날짜**")
            h_col5.markdown("**조회**")
            h_col6.markdown("**추천**")
            h_col7.markdown("**댓글**")
            st.divider()

            # [요청사항 2] 게시글 목록 출력
            for _, row in posts.iterrows():
                c1, c2, c3, c4, c5, c6, c7 = st.columns([0.8, 5, 1.5, 1.5, 0.8, 0.8, 0.8])
                
                # 1. 번호 (공지사항은 '공지'로 표시, 일반글은 ID 표시)
                # [요청사항 2] 공지사항 핑크색 박스 및 중앙 정렬
                if row['is_notice']:
                    c1.markdown('<div style="background-color:#FFEBEE; color:#D32F2F; padding:2px 5px; border-radius:5px; font-weight:bold; text-align:center; font-size:0.8rem;">공지</div>', unsafe_allow_html=True)
                else:
                    c1.markdown(f'<div style="text-align:center; font-size:0.9rem;">{row["id"]}</div>', unsafe_allow_html=True)
                
                # 2. 제목 (클릭 가능한 버튼으로 구현하여 내용 토글)
                # [요청사항 3] 회색 박스(버튼) 너비 일정하게 고정 -> use_container_width=True
                title_text = f"{'🔒 ' if row['is_secret'] else ''}{row['title']}"
                # 클릭 시 해당 글의 ID를 active_post_id 세션에 저장
                if c2.button(title_text, key=f"btn_title_{row['id']}", use_container_width=True):
                    if st.session_state.get('active_post_id') == row['id']:
                        st.session_state.active_post_id = None # 이미 열려있으면 닫기
                    else:
                        st.session_state.active_post_id = row['id']
                        increment_views(row['id']) # 조회수 증가
                    st.rerun()
                
                # 3. 이름 (마스킹 처리)
                author_id = row['user_id']
                display_author = author_id[:3] + "*" * (len(author_id.split('@')[0]) - 3) if '@' in author_id and len(author_id.split('@')[0]) > 3 else author_id[:3] + "***"
                c3.markdown(f'<div style="text-align:center; font-size:0.9rem;">{display_author}</div>', unsafe_allow_html=True)
                
                # 4. 날짜 (YYYY-MM-DD 형식만 표시)
                c4.markdown(f'<div style="text-align:center; font-size:0.9rem;">{row["reg_date"][:10]}</div>', unsafe_allow_html=True)
                
                # 5. 조회수
                c5.markdown(f'<div style="text-align:center; font-size:0.9rem;">{row.get("views", 0)}</div>', unsafe_allow_html=True)
                
                # 6. 추천수
                c6.markdown(f'<div style="text-align:center; font-size:0.9rem;">{row["likes"]}</div>', unsafe_allow_html=True)
                
                # 7. [요청사항 1] 댓글 수 표시
                comment_count = len(get_comments(row['id']))
                c7.markdown(f'<div style="text-align:center; font-size:0.9rem;">{comment_count}</div>', unsafe_allow_html=True)

                # [내용 표시 영역] 선택된 글인 경우 하단에 내용 표시
                if st.session_state.get('active_post_id') == row['id']:
                    with st.container(border=True):
                        # 권한 체크
                        can_view = True
                        if row['is_secret']:
                            if not st.session_state.user_id:
                                can_view = False
                            elif st.session_state.user_role != 'admin' and st.session_state.user_id != row['user_id']:
                                can_view = False
                        
                        if not can_view:
                            st.warning("🔒 비밀글입니다. 작성자와 관리자만 볼 수 있습니다.")
                        else:
                            # [요청사항 3] 줄바꿈 및 띄어쓰기 보존 처리
                            content_display = row['content'].replace("\n", "  \n")
                            st.markdown(content_display)
                            st.divider()
                            
                            # 기능 버튼 (좋아요, 수정, 삭제)
                            ac1, ac2, ac3 = st.columns([1,1,1])
                            with ac1:
                                if st.button(f"👍 좋아요 ({row['likes']})", key=f"like_inner_{row['id']}"):
                                    if st.session_state.user_id:
                                        like_post(row['id'])
                                        st.rerun()
                                    else: st.warning("회원 가입 후 가능합니다.")
                            
                            is_author = st.session_state.user_id and (st.session_state.user_id == row['user_id'])
                            is_admin = st.session_state.user_role == 'admin'
                            
                            if is_author or is_admin:
                                with ac2:
                                    if st.button("📝 수정", key=f"edit_btn_{row['id']}"):
                                        st.session_state.edit_pid = row['id']
                                        st.session_state.edit_mode = True
                                        st.rerun()
                                with ac3:
                                    if st.button("🗑️ 삭제", key=f"del_btn_{row['id']}"):
                                        delete_post(row['id'])
                                        st.session_state.active_post_id = None
                                        st.rerun()
                            elif row['non_user_pw'] and not st.session_state.user_id:
                                with st.popover("비회원 글 관리"):
                                    check_pw = st.text_input("비밀번호", type="password", key=f"check_pw_{row['id']}")
                                    if st.button("확인 및 삭제", key=f"del_non_{row['id']}"):
                                        if check_pw == row['non_user_pw']:
                                            delete_post(row['id'])
                                            st.session_state.active_post_id = None
                                            st.success("삭제되었습니다.")
                                            st.rerun()
                                        else:
                                            st.error("비밀번호가 일치하지 않습니다.")
                            
                            # 댓글 영역
                            st.write("---")
                            st.markdown("**💬 댓글**")
                            comments = get_comments(row['id'])
                            for _, com in comments.iterrows():
                                com_view = True
                                if com['is_secret']:
                                    if not st.session_state.user_id: com_view = False
                                    elif st.session_state.user_role != 'admin' and st.session_state.user_id != com['user_id'] and st.session_state.user_id != row['user_id']:
                                        com_view = False
                                
                                if com_view:
                                    com_author = com['user_id']
                                    display_com_author = com_author[:3] + "***"
                                    st.markdown(f"- **{display_com_author}**: {com['content']} {'🔒' if com['is_secret'] else ''}")
                                else:
                                    st.markdown(f"- 🔒 비밀 댓글입니다.")
                            
                            if st.session_state.user_id:
                                with st.form(f"com_f_{row['id']}", clear_on_submit=True):
                                    ct = st.text_input("댓글 입력")
                                    cs = st.checkbox("비밀댓글")
                                    if st.form_submit_button("작성"):
                                        if ct: add_comment(row['id'], st.session_state.user_id, ct, cs); st.rerun()
                            else:
                                st.info("로그인 후 댓글 작성이 가능합니다.")

        if st.session_state.get('edit_mode'):
            st.divider()
            edit_post_data = posts[posts['id'] == st.session_state.edit_pid]
            if not edit_post_data.empty:
                curr = edit_post_data.iloc[0]
                with st.form("edit_post_f"):
                    st.write("### 글 수정")
                    et = st.text_input("제목", value=curr['title'])
                    # [요청사항 3] 수정 에디터 높이 2배 증가 (height=500)
                    ec = st.text_area("내용", value=curr['content'], height=500)
                    es = st.checkbox("비밀글", value=bool(curr['is_secret']))
                    en = st.checkbox("공지사항", value=bool(curr['is_notice']), disabled=(st.session_state.user_role != 'admin'))
                    if st.form_submit_button("수정 완료"):
                        update_post(st.session_state.edit_pid, et, ec, es, en)
                        st.session_state.edit_mode = False
                        st.rerun()
                    if st.form_submit_button("취소"):
                        st.session_state.edit_mode = False
                        st.rerun()

    with tab_write:
        if st.session_state.user_id:
            with st.form("new_post_f", clear_on_submit=True):
                wt = st.text_input("제목")
                # [요청사항 3] 글쓰기 에디터 높이 조정 (기본 200 -> 500)
                wc = st.text_area("내용", height=500)
                ws = st.checkbox("비밀 글쓰기 (관리자만 읽을 수 있음)")
                wn = st.checkbox("공지사항 등록 (관리자 전용)", disabled=(st.session_state.user_role != 'admin'))
                if st.form_submit_button("등록"):
                    if wt and wc:
                        add_post(st.session_state.user_id, wt, wc, ws, wn)
                        st.success("등록되었습니다.")
                        st.rerun()
                    else: st.error("내용을 입력하세요.")
        else:
            st.info("💡 비회원으로 글을 작성하실 수 있습니다. 작성 후 회원가입을 고려해 주세요.")
            with st.form("non_user_post_f", clear_on_submit=True):
                non_name = st.text_input("작성자명")
                non_pw = st.text_input("비밀번호 (수정/삭제용)", type="password")
                wt = st.text_input("제목")
                wc = st.text_area("내용", height=500)
                if st.form_submit_button("등록"):
                    if non_name and non_pw and wt and wc:
                        add_post(non_name, wt, wc, False, False, non_user_pw=non_pw)
                        st.success("게시글이 성공적으로 등록되었습니다.")
                        st.info("🙏 안녕하세요! 게시글을 작성해 주셔서 감사합니다. 회원으로 가입하시면 작성하신 글을 더 체계적으로 관리하고, 분석 보관함 등 AHP 마스터의 모든 전문 기능을 자유롭게 이용하실 수 있습니다. 지금 바로 가입해 보시는 건 어떨까요?")
                    else:
                        st.error("모든 항목(작성자명, 비밀번호, 제목, 내용)을 입력해야 합니다.")

# -----------------------------------------------------------------------------
# 2. Setup & Layout
# -----------------------------------------------------------------------------

init_db()

st.markdown("""
<style>
    .stDataFrame {font-size: 0.9rem;} 
    div[data-testid="stMetricValue"] {font-size: 1.2rem;}
    .stDownloadButton > button {
        background-color: #d32f2f;
        color: white;
        border-radius: 5px;
        border: none;
        padding: 0.6rem 1.2rem;
        font-weight: bold;
        transition: background-color 0.3s ease;
    }
    .stDownloadButton > button:hover {
        background-color: #b71c1c;
    }
    div.stButton > button:first-child[kind="primary"] {
        background-color: #90EE90 !important; 
        color: black !important;
        border: none !important;
    }
</style>
""", unsafe_allow_html=True)

if 'user_id' not in st.session_state: st.session_state.user_id = None
if 'user_role' not in st.session_state: st.session_state.user_role = None
if 'expiry_date' not in st.session_state: st.session_state.expiry_date = None
if 'admin_mode' not in st.session_state: st.session_state.admin_mode = False
if 'model_structure' not in st.session_state: st.session_state.model_structure = {}
if 'page' not in st.session_state: st.session_state.page = "main"

# =============================================================================
# 3. Sidebar (Auth & Settings)
# =============================================================================

fee_info_text = """
---
### 💰 서비스 이용료
- **무료사용자**: 무료 (5표본 제한 외 기능제한 없음)
- **학위논문**: 40만원
- **일반연구**: 50만원

**결제 정보**
- **계좌번호**: 카카오뱅크 3333-23-8667708
- **예금주**: 전상현
- **주의**: 송금자명에 **가입한 이메일 주소**를 기입해주세요.
"""

with st.sidebar:
    try:
        st.image("ahp_master_logo.png", use_container_width=True)
    except:
        st.subheader("📊 AHP 마스터")
    
    with st.expander("ℹ️ 일관성 보정 기준", expanded=False):
        st.markdown("""
        **보정 방법: 반복 수렴 조정법(Iterative Adjustment)**
        가중치 산출 알고리즘(Saaty)에 의해 판단 행렬이 비일관적(CR > 임계값)인 경우, 수학적으로 일관된 행렬과 원본 행렬을 일정 비율로 혼합하여 반복적으로 가중치를 미세 조정한 결과를 제시합니다.
        
        **현재 방법의 특징:**
        1. **최소 판단 왜곡**: 원본 설문 응답의 경향성을 90%보존하면서 수학적 일관성만을 확보합니다.
        2. **자동 수렴**: 설정된 반복 횟수 내에서 CR 값을 임계값 이하로 자동 개선합니다. ($New = Old^{0.9} \\times Ideal^{0.1}$)
        
        """)        
    
    if st.button("🌐 커뮤니티 게시판", use_container_width=True, type="primary"):
        st.session_state.page = "community"
        st.rerun()
    
    if st.session_state.user_id is None:
        tab_login, tab_signup, tab_find_pw = st.tabs(["로그인", "회원가입", "비밀번호 찾기"])
        
        with tab_login:
            st.header("🔐 로그인")
            l_id = st.text_input("아이디 (이메일 주소)", key="l_id")
            l_pw = st.text_input("비밀번호 (PW)", type="password", key="l_pw")
            if st.button("로그인 실행"):
                result = check_login(l_id.strip(), l_pw)
                if result:
                    today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).date()
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
            
            st.markdown(fee_info_text)

        with tab_signup:
            st.header("📝 회원가입")
            agreements = show_agreement_ui()
            s_id = st.text_input("아이디 (이메일 주소)", key="s_id")
            s_pw = st.text_input("비밀번호", type="password", key="s_pw")
            s_role_selection = st.radio("이용 권한 선택", ("무료사용자", "정식 사용자 (2개월, 기능 무제한)"), index=0)
            
            if "정식" in s_role_selection:
                st.warning("⚠️ 정식 사용자 가입 안내")
                st.info("정식 사용자는 입급 전까지 **무료사용자** 권한이 부여됩니다.")
                st.info("관리자가 입금 확인 후 **정식 사용자**로 권한이 변경됩니다, 승인 완료 시 이메일로 안내해 드립니다. (사용 기간은 2개월 입니다)")
            
            if st.button("가입신청"):
                if not agreements.get("agree_personal_info"):
                    st.error("개인정보 수집·이용에 동의해야 가입신청할 수 있습니다.")
                elif not validate_email(s_id):
                    st.error("올바른 이메일 형식이 아닙니다.")
                elif not validate_password(s_pw):
                    st.error("비밀번호는 문자+특수문자여야 합니다.")
                else:
                    restore_from_deleted_sheet(s_id.strip())
                    initial_role = 'temp'
                    actual_requested_role = 'official' if "정식" in s_role_selection else 'temp'
                    agree_text = "동의" if agreements.get("agree_personal_info") else "미동의"
                    if add_user(s_id.strip(), s_pw, initial_role, agree_info=agree_text):
                        if actual_requested_role == 'official':
                            send_application_email(s_id)
                        st.success("무료사용자로 가입 완료 되었습니다")
                    else:
                        st.error("이미 존재하는 아이디입니다.")
            
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
        role_disp = "관리자" if st.session_state.user_role == 'admin' else ("정식 사용자" if st.session_state.user_role == 'official' else "무료사용자")
        st.info(f"권한: {role_disp}")
        
        if st.session_state.user_role == 'temp':
            if st.button("정식 사용자 전환 요청"):
                if send_conversion_request_email(st.session_state.user_id):
                    st.success("정식 사용자 전환요청이 완료 되었습니다. 입금 확인 후 정식사용자로 전환해 드립니다")
                else:
                    st.error("요청 전송 실패. 관리자에게 문의바랍니다.")

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
    max_iter = st.number_input("최대 보정 반복 횟수", min_value=10, max_value=500, value=500, step=50)

    st.markdown("---")
    with st.expander("💡 사용자 권한 안내", expanded=False):
        st.info("**비로그인(Guest)**: 샘플 파일 분석만 가능")
        st.info("**무료사용자**: 나만의 모델 생성, 분석 가능 (무료 5표본 제한)")
        st.info("**정식 사용자**: 모든 기능 무제한 (2개월/필요시 1개월 연장)")
    
    st.markdown("### 📞 문의처")
    st.markdown("- **이메일**: jeon080423@gmail.com")
    st.markdown("- **카톡ID**: AHPkr")
    st.markdown("- **전화**: 010-2142-2610")
    st.markdown("- **[사용설명서](https://morison.tistory.com/97)**")

# =============================================================================
# 4. Main Content Logic
# =============================================================================

if st.session_state.page == "community":
    show_community_board()
else:
    st.title("AHP 마스터: 분석 자동화 시스템")

    st.markdown("Saaty(1980)의 Analytic Hierarchy Process (AHP) 분석 및 일관성 자동 보정 도구입니다. 엑셀 파일을 업로드하면 개인별 가중치 산출, 일관성 보정(CR), 그룹별 집계 결과를 제공합니다.\n\n**■ 코딩 프로그램**: Python\n\n**■ 제작/관리**: 제온 https://blog.naver.com/morison00")
            

    if st.session_state.get('admin_mode', False) and st.session_state.user_role == 'admin':
        st.subheader("👥 가입자 현황 및 관리")
        
        col_sync1, col_sync2 = st.columns([2, 8])
        with col_sync1:
            if st.button("🔄 구글 시트와 동기화"):
                with st.spinner("구글 시트 데이터 불러오는 중..."):
                    added_count = sync_db_from_sheets()
                if added_count >= 0:
                    st.success(f"동기화 완료! (복구된 회원 수: {added_count}명)")
                    st.rerun()
                else:
                    st.error("동기화 중 오류가 발생했습니다.")
        
        try:
            client = get_gspread_client()
            if client:
                spreadsheet = client.open_by_key(st.secrets["SPREADSHEET_ID"])
                try:
                    visit_sheet = spreadsheet.worksheet("Visit_Logs")
                    visit_data_gs = visit_sheet.get_all_records()
                    daily_df_logs = pd.DataFrame(visit_data_gs)
                    if not daily_df_logs.empty:
                        daily_df_logs['Date_Only'] = daily_df_logs['Date'].astype(str).str[:10]
                        daily_df_counts = daily_df_logs.groupby('Date_Only').size().reset_index(name='count')
                        total_visits = len(daily_df_logs)

                        st.write("#### 🗺️ 접속자 실시간 위치 분포")
                        if 'Latitude' in daily_df_logs.columns and 'Longitude' in daily_df_logs.columns:
                            map_data = daily_df_logs[daily_df_logs['Latitude'].astype(str).str.strip() != ""].copy()
                            if not map_data.empty:
                                map_data['lat'] = pd.to_numeric(map_data['Latitude'], errors='coerce')
                                map_data['lon'] = pd.to_numeric(map_data['Longitude'], errors='coerce')
                                map_data = map_data.dropna(subset=['lat', 'lon'])
                                if not map_data.empty:
                                    map_display = map_data.groupby(['lat', 'lon']).size().reset_index(name='visit_count')
                                    map_display['size'] = map_display['visit_count'] * 20
                                    st.map(map_display, latitude='lat', longitude='lon', size='size')
                                else:
                                    st.info("유효한 좌표 데이터가 없습니다.")
                            else:
                                st.info("지도에 표시할 위치 정보 데이터가 아직 수집되지 않았습니다.")
                        else:
                            st.info("위치 정보 컬럼이 존재하지 않습니다.")
                    else:
                        total_visits = 0
                        daily_df_counts = pd.DataFrame()
                except gspread.exceptions.WorksheetNotFound:
                    total_visits = 0
                    daily_df_counts = pd.DataFrame()

                st.write(f"**총 누적 방문자 수 (시간 기반):** {total_visits:,}회")
                st.write("#### 📅 일별 방문자 현황 (날짜별 합산)")
                if not daily_df_counts.empty:
                    fig_visit = px.bar(daily_df_counts, x='Date_Only', y='count', text='count',
                                        labels={'Date_Only': '날짜', 'count': '방문자 수'})
                    fig_visit.update_traces(textposition='outside')
                    fig_visit.update_layout(xaxis_title="날짜", yaxis_title="방문자 수", showlegend=False, xaxis={'type': 'category'})
                    st.plotly_chart(fig_visit, use_container_width=True)
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
            
            if new_role == 'official' and selected_user['role'] != 'official':
                suggested_date = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).date() + relativedelta(months=2)
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

    st.subheader("1. AHP 분석 모델 설정 및 입력 템플릿 다운로드")

    if st.session_state.user_id is None:
        st.info("🔒 **로그인 후** '나만의 분석 모델'을 만들 수 있습니다. (비로그인 상태에서도 샘플 데이터로 최종 분석 결과를 미리볼 수 있습니다)")
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
                        file_name="AHP_Master_Template.xlsx",
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
        with st.expander("📂 나의 분석 보관함 (!중요) 반드시 컴퓨터에 백업해 주세요"):
            my_analyses = get_user_analyses(st.session_state.user_id)
            if not my_analyses: st.info("저장된 분석 없음")
            else:
                for item in my_analyses:
                    a_id, filename, save_date = item
                    col_List1, col_List2, col_List3, col_List4 = st.columns([3, 2, 1, 1])
                    with col_List1: st.text(f"{filename}")
                    with col_List2: st.caption(f"{save_date}")
                    with col_List3:
                        file_info = get_analysis_file(analysis_id=a_id)
                        if file_info:
                            fname, fdata = file_info
                            st.download_button("⬇️", fdata, fname, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_{a_id}")
                    with col_List4:
                        if st.button("🗑️", key=f"del_{a_id}"):
                            delete_analysis(a_id)
                            st.rerun()

    with st.container(border=True):
        st.markdown("#### ⚡ 빠른 시작 (도시재생 사업 모델)")
        st.info("아래 버튼을 누르면 테스트용 샘플 엑셀 파일이 다운로드 됩니다.\n\n"
                "다운받은 테스트 샘플 엑셀 파일을 아래 '데이터 업로드 및 분석'에 업로드 하세요.")
        
        sample_excel = create_sample_excel()
        st.download_button(
            label="📂 테스트용 샘플 데이터 다운로드",
            data=sample_excel,
            file_name="AHP_UrbanRegeneration_Sample.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.markdown("---")

    def write_custom_ahp_table(writer, sheet_name, df, title_text, start_row, formats, excluded_df=None):
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
        
        if excluded_df is not None:
            worksheet.write(start_row, 0, f"※ 분석 제외 사례수: {len(excluded_df)}건", workbook.add_format({'bold': True, 'font_color': 'red'}))
            start_row += 1
            if not excluded_df.empty:
                worksheet.write(start_row, 0, "▶ 제외된 응답 데이터 (보정 실패)", workbook.add_format({'bold': True}))
                start_row += 1
                excluded_df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
                start_row += len(excluded_df) + 2

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
            main_cols_names = df_main.columns[2:]
            main_factors, n_main = infer_factors_from_columns(main_cols_names)

            permission_granted = False
            message = ""
            role = st.session_state.user_role
            user_id = st.session_state.user_id

            if role == 'admin' or role == 'official':
                permission_granted = True
                if role == 'official':
                    today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9))).date()
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
                else: message = f"⛔ **무료사용자**는 시트당 최대 5개 표본까지만 분석 가능합니다."

            if permission_granted:
                with st.spinner("계층 분석 수행 중..."):
                    main_results_df, main_factors, main_excluded, main_excluded_df = process_single_sheet(df_main, cr_threshold, max_iter, mean_method)
                    
                    total_excluded = main_excluded
                    st.markdown(f"**분석 제외: {total_excluded}건**")

                    main_sig_df = calculate_pairwise_ttest(main_results_df, main_factors)
                    main_weight_cols = [f"Weight_{f}" for f in main_factors]
                    
                    if mean_method == 'arithmetic':
                        group_main_weights = main_results_df[main_weight_cols].mean(axis=0)
                    else:
                        group_main_weights = gmean(main_results_df[main_weight_cols].values, axis=0)
                    group_main_weights = group_main_weights / group_main_weights.sum()
                    main_cr_final_avg = main_results_df['Final_CR'].mean()
                    
                    main_matrices = np.stack(main_results_df['Matrix_Object'].values)
                    main_group_matrix = np.mean(main_matrices, axis=0) if mean_method == 'arithmetic' else gmean(main_matrices, axis=0)
                    main_grp_cr, main_grp_ci, _ = calculate_consistency(main_group_matrix, mean_method)
                    
                    indiv_global_data = []
                    all_ids = main_results_df['ID'].unique()
                    
                    sub_results_storage = {} 
                    total_excl_df_list = [main_excluded_df]
                    for i, sub_sheet_name in enumerate(sheet_names[1:]):
                        parent_factor = main_factors[i]
                        df_sub = pd.read_excel(uploaded_file, sheet_name=sub_sheet_name)
                        sub_res_df, sub_facts, sub_excl, sub_excl_df = process_single_sheet(df_sub, cr_threshold, max_iter, mean_method)
                        sub_sig_df = calculate_pairwise_ttest(sub_res_df, sub_facts)
                        sub_w_cols = [f"Weight_{f}" for f in sub_facts]
                        group_sub_w = sub_res_df[sub_w_cols].mean(axis=0) if mean_method == 'arithmetic' else gmean(sub_res_df[sub_w_cols].values, axis=0)
                        group_sub_w = group_sub_w / group_sub_w.sum()
                        sub_cr_final_avg = sub_res_df['Final_CR'].mean()
                        sub_matrices = np.stack(sub_res_df['Matrix_Object'].values)
                        sub_group_matrix = np.mean(sub_matrices, axis=0) if mean_method == 'arithmetic' else gmean(sub_matrices, axis=0)
                        sub_grp_cr, _, _ = calculate_consistency(sub_group_matrix, method=mean_method)
                        sub_results_storage[parent_factor] = {
                            'weights': group_sub_w, 'factors': sub_facts, 'cr': sub_cr_final_avg,
                            'df': sub_res_df, 'group_matrix': sub_group_matrix, 'group_cr': sub_grp_cr, 'sig_df': sub_sig_df
                        }
                        if not sub_excl_df.empty:
                            sub_excl_df['Sheet'] = sub_sheet_name
                            total_excl_df_list.append(sub_excl_df)

                    for uid in all_ids:
                        u_main = main_results_df[main_results_df['ID'] == uid]
                        if u_main.empty: continue
                        u_type = u_main['Type'].values[0]
                        for mf in main_factors:
                            m_w = u_main[f"Weight_{mf}"].values[0]
                            s_row_df = sub_results_storage[mf]['df']
                            u_sub = s_row_df[s_row_df['ID'] == uid]
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
                        g_main_w = grp_main_df[main_weight_cols].mean(axis=0) if mean_method == 'arithmetic' else gmean(grp_main_df[main_weight_cols].values, axis=0)
                        g_main_w = g_main_w / g_main_w.sum()
                        g_main_mats = np.stack(grp_main_df['Matrix_Object'].values)
                        g_main_mat_obj = np.mean(g_main_mats, axis=0) if mean_method == 'arithmetic' else gmean(g_main_mats, axis=0)
                        g_main_cr, _, _ = calculate_consistency(g_main_mat_obj, method=mean_method)
                        
                        grp_rows = []
                        for idx, main_f in enumerate(main_factors):
                            m_w = g_main_w[idx]
                            full_sub_df = sub_results_storage[main_f]['df']
                            grp_sub_df = full_sub_df[full_sub_df['Type'].astype(str) == grp]
                            sub_facts_list = sub_results_storage[main_f]['factors']
                            if grp_sub_df.empty: continue
                            s_w_cols = [f"Weight_{f}" for f in sub_facts_list]
                            g_sub_w = grp_sub_df[s_w_cols].mean(axis=0) if mean_method == 'arithmetic' else gmean(grp_sub_df[s_w_cols].values, axis=0)
                            g_sub_w = g_sub_w / g_sub_w.sum()
                            g_sub_mats = np.stack(grp_sub_df['Matrix_Object'].values)
                            g_sub_mat_obj = np.mean(g_sub_mats, axis=0) if mean_method == 'arithmetic' else gmean(g_sub_mats, axis=0)
                            g_sub_cr, _, _ = calculate_consistency(g_sub_mat_obj, method=mean_method)
                            for s_idx, sf in enumerate(sub_facts_list):
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
                            'num_sum': workbook.add_format({'num_format': '0.000', 'bg_color': '#D3D3D3', 'border': 1, 'align':'center'}),
                            'yellow': workbook.add_format({'bg_color': 'yellow', 'border': 1, 'align': 'center', 'num_format': '0.000'})
                        }
                        border_fmt = workbook.add_format({'border': 1})
                        fmt_float_no_border = workbook.add_format({'num_format': '0.000', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                        fmt_diagonal = workbook.add_format({'num_format': '0', 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E7E6E6', 'border': 1})

                        total_excluded_df = pd.concat(total_excl_df_list, ignore_index=True)
                        current_row = write_custom_ahp_table(writer, '종합분석', final_df, "1) 전체_종합결과", 1, formats, excluded_df=total_excluded_df)
                        for grp in unique_groups:
                            if grp in group_full_dfs:
                                current_row = write_custom_ahp_table(writer, '종합분석', group_full_dfs[grp], f"▶ [그룹: {grp}] 분석 결과", current_row, formats)

                        if len(unique_groups) >= 1:
                            ws_comp = workbook.add_worksheet('Group_Comparison')
                            writer.sheets['Group_Comparison'] = ws_comp
                            s_row = 1
                            ws_comp.write_string(s_row, 0, "그룹 간 비교(일원배치 분산분석: ANOVA)", workbook.add_format({'bold': True, 'font_size': 12}))
                            s_row += 1
                            
                            if not anova_df.empty:
                                anova_for_merge = anova_df.rename(columns={'요인': '중분류'})
                                integrated_df = comparison_df.merge(anova_for_merge, on='중분류', how='left')
                            else:
                                integrated_df = comparison_df
                            
                            integrated_df.to_excel(writer, sheet_name='Group_Comparison', startrow=s_row, index=False)
                            add_borders_to_data(ws_comp, s_row, 0, integrated_df, border_fmt)
                            
                            num_format_3 = workbook.add_format({'num_format': '0.000', 'border': 1, 'align': 'center'})
                            for r in range(len(integrated_df)):
                                for c in range(1, len(integrated_df.columns)):
                                    val = integrated_df.iloc[r, c]
                                    if pd.notnull(val) and isinstance(val, (int, float)):
                                        ws_comp.write_number(s_row + 1 + r, c, val, num_format_3)
                                    elif pd.notnull(val):
                                        ws_comp.write(s_row + 1 + r, c, val, border_fmt)

                            guide_start_row = s_row + len(integrated_df) + 3
                            bold_fmt = workbook.add_format({'bold': True, 'font_size': 11, 'valign': 'vcenter', 'align': 'left', 'bg_color': '#F2F2F2', 'border': 1})
                            text_fmt = workbook.add_format({'font_size': 10, 'text_wrap': True, 'valign': 'top', 'align': 'left', 'border': 1})
                            ws_comp.set_column('A:G', 20) 
                            ws_comp.merge_range(guide_start_row, 0, guide_start_row, 6, "※ 그룹 간 중요도의 차이가 있지만 통계적으로 유의하지 않게 나타나는 이유", bold_fmt)

                            guide_content = [
                                ("1. 그룹 내 편차(분산)가 너무 큰 경우", "ANOVA는 '그룹 간의 차이'와 '그룹 내의 차이'를 비교합니다.\n\n■ 원리: 그룹 간 평균 차이가 크더라도, 각 그룹 내부 데이터들이 서로 들쭉날쭉(분산이 큼)하다면 통계적으로는 '이 차이가 우연히 발생했을 가능성이 높다'고 판단합니다.\n■ 분석: 현재 데이터에서 평균값의 절대적인 차이는 커 보일 수 있지만, 각 그룹(A~D)에 속한 개별 응답자들의 값들이 평균에서 멀리 떨어져 있다면 F-값이 낮아지고 P-Value는 올라가게 됩니다."),
                                ("2. 표본 크기(Sample Size)의 부족", "통계적 유의성은 표본의 수에 매우 민감합니다.\n\n■ 현상: 각 그룹의 데이터 개수(표본수)가 너무 적다면(예: 그룹당 3~5개 미만) 아무리 평균 차이가 커도 통계적 힘(Power)이 부족하여 유의미한 차이를 찾아내지 못합니다.\n■ 확인 사항: 현재 분석에 사용된 각 그룹의 n수(표본수)가 충분한지 검토가 필요합니다."),
                                ("3. 데이터의 단위(Scale)와 변동성", "표에 나타난 수치들이 대부분 0.1 미만 혹은 0.2 수준의 매우 작은 소수점 단위입니다.\n\n■ 분석: 수치 자체가 작기 때문에 시각적으로는 0.05와 0.15가 3배 차이로 커 보일 수 있지만, 실제 계산 과정에서 발생하는 표준오차(Standard Error) 범위 안에 해당 수치들이 포함되어 있다면 통계적으로는 '측정 오차 범위 내의 흔들림'으로 간주됩니다.")
                            ]

                            current_row_comp = guide_start_row + 1
                            for title, body in guide_content:
                                ws_comp.set_row(current_row_comp, 25)
                                ws_comp.merge_range(current_row_comp, 0, current_row_comp, 6, title, bold_fmt)
                                ws_comp.set_row(current_row_comp + 1, 120)
                                ws_comp.merge_range(current_row_comp + 1, 0, current_row_comp + 1, 6, body, text_fmt)
                                current_row_comp += 2

                        def write_detailed_sheet(sheet_name, matrix_data, detail_data_df, matrix_title, row_labels, group_matrices=None, sheet_excl_count=0):
                            ws = workbook.add_worksheet(sheet_name)
                            writer.sheets[sheet_name] = ws
                            s_row_det = 0
                            
                            ws.write(s_row_det, 0, f"분석 제외 사례수: {sheet_excl_count}건", workbook.add_format({'bold': True, 'font_color': 'red'}))
                            s_row_det += 1
                            
                            ws.write_string(s_row_det, 0, matrix_title)
                            s_row_det += 1
                            m_df_obj = pd.DataFrame(matrix_data, index=row_labels, columns=row_labels)
                            m_df_obj.to_excel(writer, sheet_name=sheet_name, startrow=s_row_det)
                            add_borders_to_data(ws, s_row_det, 0, m_df_obj, border_fmt, has_header=True, has_index=True)
                            for r in range(len(matrix_data)):
                                for c in range(len(matrix_data)):
                                    val = 1 if r==c else matrix_data[r][c]
                                    ws.write(s_row_det+r+1, c+1, val, border_fmt if r!=c else fmt_diagonal)
                                    if r!=c: ws.write(s_row_det+r+1, c+1, val, fmt_float_no_border)
                            
                            s_row_det += len(matrix_data) + 3
                            
                            if group_matrices:
                                for g_name, g_mat in group_matrices.items():
                                    ws.write_string(s_row_det, 0, f"] 그룹 종합 행렬: {g_name}")
                                    s_row_det += 1
                                    gm_df_obj = pd.DataFrame(g_mat, index=row_labels, columns=row_labels)
                                    gm_df_obj.to_excel(writer, sheet_name=sheet_name, startrow=s_row_det)
                                    add_borders_to_data(ws, s_row_det, 0, gm_df_obj, border_fmt, has_header=True, has_index=True)
                                    for r in range(len(g_mat)):
                                        for c in range(len(g_mat)):
                                            val = 1 if r==c else g_mat[r][c]
                                            ws.write(s_row_det+r+1, c+1, val, border_fmt if r!=c else fmt_diagonal)
                                            if r!=c: ws.write(s_row_det+r+1, c+1, val, fmt_float_no_border)
                                    s_row_det += len(g_mat) + 3
                            
                            detail_data_df.to_excel(writer, sheet_name=sheet_name, startrow=s_row_det, index=False)
                            
                            for c_idx, col_val in enumerate(detail_data_df.columns):
                                ws.write(s_row_det, c_idx, col_val, formats['header'])
                            
                            for r_idx in range(len(detail_data_df)):
                                orig_cr_val = detail_data_df.iloc[r_idx]['Original_CR']
                                final_cr_val = detail_data_df.iloc[r_idx]['Final_CR']
                                row_pos = s_row_det + 1 + r_idx
                                
                                for c_idx, col_name in enumerate(detail_data_df.columns):
                                    val = detail_data_df.iloc[r_idx, c_idx]
                                    current_fmt = border_fmt
                                    
                                    if col_name == 'Original_CR' and orig_cr_val > 0.1:
                                        current_fmt = formats['yellow']
                                    elif col_name == 'Final_CR' and final_cr_val > 0.1:
                                        current_fmt = formats['yellow']
                                    elif isinstance(val, (float, np.float64)):
                                        current_fmt = formats['num']
                                    else:
                                        current_fmt = formats['body']
                                        
                                    if pd.isnull(val):
                                        ws.write_blank(row_pos, c_idx, "", current_fmt)
                                    else:
                                        ws.write(row_pos, c_idx, val, current_fmt)

                        main_group_mats = {}
                        for grp in unique_groups:
                            g_df_m = main_results_df[main_results_df['Type'].astype(str) == grp]
                            if not g_df_m.empty:
                                mats_stack = np.stack(g_df_m['Matrix_Object'].values)
                                main_group_mats[grp] = np.mean(mats_stack, axis=0) if mean_method == 'arithmetic' else gmean(mats_stack, axis=0)

                        out_main = main_results_df.drop(columns=['Matrix_Object'], errors='ignore')
                        write_detailed_sheet('Result_Main', main_group_matrix, out_main, f"[1] 전체 종합 행렬", main_factors, group_matrices=main_group_mats, sheet_excl_count=main_excluded)
                        for mf, info in sub_results_storage.items():
                            safe_name = f"Result_{mf}"[:31]
                            sub_grp_mats = {}
                            for grp in unique_groups:
                                g_sub_df = info['df'][info['df']['Type'].astype(str) == grp]
                                if not g_sub_df.empty:
                                    mats_stack = np.stack(g_sub_df['Matrix_Object'].values)
                                    sub_grp_mats[grp] = np.mean(mats_stack, axis=0) if mean_method == 'arithmetic' else gmean(mats_stack, axis=0)
                            out_sub = info['df'].drop(columns=['Matrix_Object'], errors='ignore')
                            
                            sub_excl_val = 0
                            for df_ex in total_excl_df_list:
                                if 'Sheet' in df_ex.columns and not df_ex.empty:
                                     if df_ex['Sheet'].iloc[0] == mf or (mf in df_ex['Sheet'].unique()):
                                          sub_excl_val = len(df_ex[df_ex['Sheet'] == mf])
                                          
                            write_detailed_sheet(safe_name, info['group_matrix'], out_sub, f"[1] 전체 종합 행렬", info['factors'], group_matrices=sub_grp_mats, sheet_excl_count=sub_excl_val)

                        theory_ws = workbook.add_worksheet("Consistency_Theory")
                        theory_title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_name': 'NanumGothic'})
                        theory_body_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top', 'font_name': 'NanumGothic'})
                        theory_text = [
                            ["의사결정론적 관점에서의 AHP 일관성 보정 원리 및 학술적 근거"],
                            [""],
                            ["1. 서론: 계층분석과정(AHP)의 일관성 문제"],
                            ["Saaty(1980)에 의해 제안된 계층분석과정(Analytic Hierarchy Process, AHP)은 인간의 주관적 판단을 정량화하는 강력한 다기준 의사결정 도구이다. 그러나 의사결정자의 인지적 한계로 인해 쌍대비교 행렬에서 이행성(Transitivity)이 결여된 비일관적 판단이 발생할 수 있다. 본 시스템은 이러한 비일관성을 수학적으로 교정하여 분석의 신뢰성을 확보한다."],
                            [""],
                            ["2. 보정 알고리즘: 반복 수렴 조정법(Iterative Adjustment Method)"],
                            ["본 시스템에 적용된 보정 로직은 '반복적 선형 결합 수렴법'에 근거한다. 비일관적 행렬 A가 주어졌을 때, 일관성 비율(Consistency Ratio, CR)이 임계값(0.1 또는 0.2)을 초과할 경우 다음과 같은 프로세스를 수행한다."],
                            ["    가. 고유벡터법(Eigenvector Method) 또는 기하평균법을 통해 현재 행렬의 가중치 벡터 w를 도출한다."],
                            ["    나. 가중치 벡터 w를 기반으로 완벽한 일관성을 가진 행렬 W = [wi/wj]를 생성한다. 이를 '이상적 일관 행렬'이라 정의한다."],
                            ["    다. 원본 행렬 A와 이상적 행렬 W를 특정 학습률(Learning Rate, α=0.4)에 따라 선형 결합(Linear Combination)한다: A_new = (1-α)A + αW."],
                            ["    라. 교정된 행렬 A_new의 역수성(Reciprocity)을 재설정하고, CR이 임계값 이하로 수렴할 때까지 위 과정을 최대 500회 반복한다."],
                            [""],
                            ["3. 학술적 근거 및 효과"],
                            ["첫째, 최소 판단 왜곡의 원리(Principle of Minimal Distortion): Cao et al.(2008)에 따르면, 원본 행렬과 일관 행렬의 가중 평균을 이용한 조정은 의사결정자의 원래 선호 경향성을 최대한 보존하면서 수학적 일관성만을 선택적으로 향상시키는 효과가 입증되었다."],
                            ["둘째, 수렴 안정성: 반복적 조정 프로세스는 행렬의 최대 고유값(λmax)을 차원 수 n에 수렴하게 함으로써 일관성 지수(CI)를 통계적으로 유의미한 수준으로 감소시킨다."],
                            ["셋째, 실무적 유용성: 설문 응답자에게 재설문을 요구하기 어려운 연구 환경에서, 본 보정법은 데이터의 대푯값을 훼손하지 않는 범위 내에서 분석의 논리적 타당성을 부여하는 학술적 대안으로 활용된다."],
                            [""],
                            ["본 시스템의 분석 결과는 위와 같은 엄밀한 수치적 보정을 거쳐 산출되었으므로, 학술 연구 및 정책 의사결정의 기초 자료로 활용하기에 적합한 신뢰도를 보유함을 확인한다."]
                        ]
                        theory_ws.set_column('A:A', 100)
                        for r_idx, row_content in enumerate(theory_text):
                            fmt = theory_title_fmt if r_idx == 0 else theory_body_fmt
                            theory_ws.write(r_idx, 0, row_content[0], fmt)

                    st.success("분석이 완료되었습니다.")
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
                            indiv_global_radar = []
                            all_ids_r = main_results_df['ID'].unique()
                            for rid in all_ids_r:
                                m_row_r = main_results_df[main_results_df['ID'] == rid].iloc[0]
                                rtype_r = m_row_r['Type']
                                for m_f in main_factors:
                                    mw_indiv = m_row_r[f"Weight_{m_f}"]
                                    s_row_df_r = sub_results_storage[m_f]['df']
                                    s_row_r = s_row_df_r[s_row_df_r['ID'] == rid].iloc[0]
                                    for s_f in sub_results_storage[m_f]['factors']:
                                        indiv_global_radar.append({"Type": rtype_r, "Factor": s_f, "Global_Weight": mw_indiv * s_row_r[f"Weight_{s_f}"]})
                            radar_indiv_df = pd.DataFrame(indiv_global_radar)
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
st.caption("© 2026 AHP Master. All rights reserved.")
