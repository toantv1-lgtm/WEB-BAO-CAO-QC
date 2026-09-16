from datetime import date, datetime
import io
import os
import re
import sqlite3
import numpy as np
import openpyxl
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ================= 1. KẾT NỐI CSDL AN TOÀN & TỰ ĐỘNG CẬP NHẬT CẤU TRÚC =================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "Report_Database.db")
IMG_DIR = os.path.join(BASE_DIR, "Anh_kiem_tra_dau_vao")

os.makedirs(IMG_DIR, exist_ok=True)


try:
  USE_TURSO = "TURSO_DATABASE_URL" in st.secrets
except Exception:
  USE_TURSO = False


# ---------- Lớp tương thích libsql_experimental <-> sqlite3.Row ----------
# libsql_experimental hiện CHƯA hỗ trợ row_factory, nên các lớp dưới đây mô
# phỏng lại hành vi row["ten_cot"] / row[index] của sqlite3.Row để toàn bộ
# logic truy vấn cũ (kể cả pd.read_sql_query) chạy nguyên vẹn.
class _LibsqlRow:
  __slots__ = ("_cols", "_vals")

  def __init__(self, cols, vals):
    self._cols = cols
    self._vals = vals

  def __getitem__(self, key):
    if isinstance(key, str):
      return self._vals[self._cols.index(key)]
    return self._vals[key]

  def get(self, key, default=None):
    try:
      return self[key]
    except (ValueError, IndexError):
      return default

  def keys(self):
    return list(self._cols)

  def __iter__(self):
    return iter(self._vals)

  def __len__(self):
    return len(self._vals)

  def __repr__(self):
    return repr(dict(zip(self._cols, self._vals)))


class _LibsqlCursorWrapper:
  def __init__(self, cursor):
    self._cursor = cursor

  def _cols(self):
    return [d[0] for d in (self._cursor.description or [])]

  def execute(self, sql, params=()):
    self._cursor.execute(sql, params if params is not None else ())
    return self

  def executemany(self, sql, seq_of_params):
    self._cursor.executemany(sql, seq_of_params)
    return self

  def fetchone(self):
    row = self._cursor.fetchone()
    return _LibsqlRow(self._cols(), row) if row is not None else None

  def fetchall(self):
    cols = self._cols()
    return [_LibsqlRow(cols, r) for r in self._cursor.fetchall()]

  @property
  def description(self):
    return self._cursor.description

  @property
  def lastrowid(self):
    return getattr(self._cursor, "lastrowid", None)

  @property
  def rowcount(self):
    return getattr(self._cursor, "rowcount", -1)


class _LibsqlConnWrapper:
  """Bọc connection libsql_experimental để tương thích code viết cho sqlite3:
  conn.execute(), conn.cursor(), row['col'], pd.read_sql_query(sql, conn, ...)."""

  def __init__(self, raw_conn):
    self._conn = raw_conn
    self.row_factory = None  # giữ để tương thích, không thực sự dùng

  def cursor(self):
    return _LibsqlCursorWrapper(self._conn.cursor())

  def execute(self, sql, params=()):
    return _LibsqlCursorWrapper(self._conn.cursor()).execute(sql, params)

  def executemany(self, sql, seq_of_params):
    return _LibsqlCursorWrapper(self._conn.cursor()).executemany(sql, seq_of_params)

  def commit(self):
    try:
      self._conn.commit()
    except Exception:
      pass

  def rollback(self):
    try:
      self._conn.rollback()
    except Exception:
      pass

  def close(self):
    try:
      self._conn.close()
    except Exception:
      pass


def get_db_connection():
  """Ưu tiên kết nối Turso (đọc từ st.secrets, giống WEB_BAOCA_2.py để đảm bảo
  Dashboard và Web nhập liệu luôn đọc/ghi cùng một CSDL trên cloud). Nếu không
  có secrets Turso (vd. chạy thử trên máy local) thì rơi về SQLite file."""
  if USE_TURSO:
    import libsql_experimental as libsql

    raw = libsql.connect(
        database=st.secrets["TURSO_DATABASE_URL"],
        auth_token=st.secrets.get("TURSO_AUTH_TOKEN", ""),
    )
    return _LibsqlConnWrapper(raw)
  else:
    conn = sqlite3.connect(DB_PATH, timeout=30.0)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
  try:
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
            CREATE TABLE IF NOT EXISTS tb_qc_dau_vao (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                loai_qc TEXT DEFAULT 'DAU_VAO',
                so_lot TEXT, ma_vt TEXT, ten_vt TEXT, ncc TEXT, ngay_ve TEXT,
                tong_sl_ve REAL, sl_kiem REAL, sl_khong_dat REAL, sl_dat REAL,
                ket_luan TEXT, nguoi_kiem TEXT, ngay_kiem DATETIME, ghi_chu TEXT,
                kieu_loi TEXT DEFAULT '', cong_viec_con TEXT DEFAULT '', img1 TEXT, img2 TEXT
            )
        """)

    cursor.execute("PRAGMA table_info(tb_qc_dau_vao)")
    cols = [col[1] for col in cursor.fetchall()]
    if "loai_qc" not in cols:
      cursor.execute(
          "ALTER TABLE tb_qc_dau_vao ADD COLUMN loai_qc TEXT DEFAULT 'DAU_VAO'"
      )
    if "kieu_loi" not in cols:
      cursor.execute(
          "ALTER TABLE tb_qc_dau_vao ADD COLUMN kieu_loi TEXT DEFAULT ''"
      )
    if "cong_viec_con" not in cols:
      cursor.execute(
          "ALTER TABLE tb_qc_dau_vao ADD COLUMN cong_viec_con TEXT DEFAULT ''"
      )
    if "sl_huy" not in cols:
      cursor.execute(
          "ALTER TABLE tb_qc_dau_vao ADD COLUMN sl_huy REAL DEFAULT 0"
      )

    cursor.execute("""
            CREATE TABLE IF NOT EXISTS tb_dm_loai_loi (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                phan_he TEXT,
                ten_loi TEXT
            )
        """)

    cursor.execute("""
            CREATE TABLE IF NOT EXISTS tb_dm_cong_viec (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                phan_he TEXT,
                ten_cong_viec TEXT
            )
        """)
    
    # Bảng lưu trữ mục tiêu chất lượng
    cursor.execute("""
            CREATE TABLE IF NOT EXISTS tb_dm_muc_tieu (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                phan_he TEXT,
                mat_prefix TEXT,
                muc_tieu REAL
            )
        """)

    conn.commit()
    conn.close()
  except Exception:
    pass


init_db()

# ================= 2. CẤU HÌNH DASHBOARD & HỆ THỐNG THIẾT KẾ =================
st.set_page_config(
    page_title="EMIC - Dashboard Tổng Hợp",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800;900&display=swap');
        #MainMenu { visibility: hidden; }
        footer { visibility: hidden; }
        header[data-testid="stHeader"] { background: transparent !important; box-shadow: none !important; }
        div[data-baseweb="popover"] { z-index: 999999 !important; }
        div[data-baseweb="calendar"] { z-index: 999999 !important; }

        :root {
            --bg: #F5F6F9;
            --bg-accent: #EEF1FA;
            --surface: #FFFFFF;
            --border: #E6E9F0;
            --border-strong: #D6DAE5;
            --text: #0F1222;
            --text-muted: #6B7280;
            --text-faint: #9CA3AF;
            --primary: #4F46E5;
            --primary-dark: #3730A3;
            --primary-soft: #EEF0FF;
            --success: #0EA968;
            --success-soft: #ECFDF5;
            --danger: #E23D4D;
            --danger-soft: #FEF2F3;
            --warning: #EA8A0A;
            --warning-soft: #FFF8EB;
            --purple: #9333EA;
            --radius-lg: 18px;
            --radius: 14px;
            --radius-sm: 10px;
            --shadow-xs: 0 1px 2px rgba(15,18,34,0.05);
            --shadow: 0 1px 3px rgba(15,18,34,0.05), 0 6px 16px -6px rgba(15,18,34,0.08);
            --shadow-hover: 0 8px 24px -6px rgba(15,18,34,0.14);
        }

        html, body, [class*="css"], .stApp {
            font-family: 'Plus Jakarta Sans', -apple-system, 'Segoe UI', sans-serif !important;
            color: var(--text);
        }
        .stApp {
            background: linear-gradient(135deg, #EEF2FF 0%, #E0E7FF 50%, #F3E8FF 100%) !important;
            background-attachment: fixed !important;
        }
        h1, h2, h3, h4, h5, h6, p, span, div, label { font-family: 'Plus Jakarta Sans', sans-serif; }

        .main .block-container, div[data-testid="stAppViewBlockContainer"] {
            padding-top: 1rem !important;
            padding-bottom: 2.5rem !important;
            padding-left: 1.2rem !important;
            padding-right: 1.2rem !important;
            max-width: 100% !important;
            margin: 0 auto !important;
        }

        .filter-banner {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 16px 22px;
            box-shadow: var(--shadow);
            margin-bottom: 22px;
        }
        .filter-banner .filter-title {
            font-size: 13px; font-weight: 700; color: var(--text-muted);
            text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 12px;
            display: flex; align-items: center; gap: 6px;
        }

        .page-header {
            display: flex; align-items: center; justify-content: space-between;
            background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 50%, #1E3A8A 100%);
            padding: 20px 26px; border-radius: 18px;
            box-shadow: 0 10px 25px -5px rgba(37, 99, 235, 0.35);
            margin-bottom: 20px; border: 1px solid rgba(255,255,255,0.15);
        }
        .page-header .ph-logo {
            font-size: 11px; font-weight: 900; letter-spacing: 0.15em; color: #FDE047;
            text-transform: uppercase; margin: 0 0 4px 0;
        }
        .page-header .ph-title {
            font-size: 20px; font-weight: 900; color: #FFFFFF; margin: 0;
            letter-spacing: -0.01em; text-shadow: 0 2px 4px rgba(0,0,0,0.15);
        }
        .page-header .ph-subtitle { font-size: 12.5px; color: rgba(255,255,255,0.85); margin: 4px 0 0 0; font-weight: 500; }
        .page-header .ph-meta {
            font-size: 12px; font-weight: 700; color: #1E3A8A;
            background: #FFFFFF; border: 1px solid rgba(255,255,255,0.6);
            padding: 7px 15px; border-radius: 999px; white-space: nowrap;
        }

        .stTabs [data-baseweb="tab-list"] {
            justify-content: flex-start !important;
            gap: 24px !important;
            background-color: transparent !important;
            padding: 0px !important;
            border-bottom: 1px solid var(--border) !important;
            margin-bottom: 22px !important;
        }
        .stTabs [data-baseweb="tab"] {
            background-color: transparent !important;
            color: var(--text-muted) !important;
            border-radius: 0px !important;
            padding: 4px 2px 13px 2px !important;
            border: none !important;
            border-bottom: 2.5px solid transparent !important;
            font-weight: 600 !important;
            font-size: 14.5px !important;
            font-family: 'Plus Jakarta Sans', sans-serif !important;
            transition: color 0.15s ease, border-color 0.15s ease;
        }
        .stTabs [data-baseweb="tab"]:hover { color: var(--text) !important; }
        .stTabs [aria-selected="true"] {
            color: var(--primary) !important;
            border-bottom: 2.5px solid var(--primary) !important;
        }
        .stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] { display: none !important; }

        .kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 20px; }
        .kpi-card {
            background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius);
            padding: 18px 20px; box-shadow: var(--shadow); position: relative; overflow: hidden;
            transition: box-shadow 0.18s ease, transform 0.18s ease, border-color 0.18s ease;
        }
        .kpi-card:hover { box-shadow: var(--shadow-hover); transform: translateY(-2px); border-color: var(--border-strong); }
        .kpi-card::before {
            content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 4px;
            background: var(--accent, var(--primary));
        }
        .kpi-card .kpi-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 14px; }
        .kpi-card .kpi-label {
            font-size: 11.5px; font-weight: 700; color: var(--text-muted);
            text-transform: uppercase; letter-spacing: 0.05em;
        }
        .kpi-card .kpi-icon {
            width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center;
            justify-content: center; font-size: 14px; background: var(--accent-soft, var(--primary-soft));
        }
        .kpi-card .kpi-value { font-size: 27px; font-weight: 800; color: var(--text); line-height: 1.15; letter-spacing: -0.02em; }
        .kpi-card .kpi-sub { font-size: 11.5px; color: var(--text-muted); margin-top: 6px; font-weight: 500; }

        .chart-card {
            background-color: var(--surface); border-radius: var(--radius); border: 1px solid var(--border);
            padding: 18px 22px 10px 22px; box-shadow: 0 2px 4px rgba(15,18,34,0.06), 0 12px 28px -8px rgba(15,18,34,0.14);
            margin-bottom: 20px; transition: box-shadow 0.18s ease, transform 0.18s ease;
        }
        .chart-card:hover { box-shadow: 0 4px 8px rgba(15,18,34,0.08), 0 18px 36px -8px rgba(15,18,34,0.18); transform: translateY(-1px); }
        .chart-card-header {
            display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px;
            padding-bottom: 10px; border-bottom: 1px solid var(--border);
        }
        .chart-card-title {
            font-size: 15.5px; font-weight: 800; color: var(--text); letter-spacing: -0.01em;
            display: flex; align-items: center; gap: 8px;
        }
        .chart-card-title::before {
            content: ""; width: 4px; height: 15px; border-radius: 2px;
            background: var(--primary); display: inline-block;
        }
        .chart-card-caption {
            font-size: 11px; color: var(--text-muted); font-weight: 600;
            background: var(--bg); padding: 3px 9px; border-radius: 999px;
        }

        .section-heading { font-size: 16.5px; font-weight: 800; color: var(--text); margin: 6px 0 14px 2px; letter-spacing: -0.01em; }

        div[data-testid="stDataFrame"] {
            border: 1px solid var(--border) !important; border-radius: var(--radius) !important;
            box-shadow: 0 2px 4px rgba(15,18,34,0.05), 0 8px 20px -6px rgba(15,18,34,0.10);
            overflow: hidden;
        }
        div[data-testid="stDataFrame"] [data-testid="stDataFrameResizable"] { font-family: 'Plus Jakarta Sans', sans-serif !important; }

        [data-testid="stWidgetLabel"] p {
            font-size: 12px !important; font-weight: 700 !important; color: var(--text-muted) !important;
            text-transform: uppercase; letter-spacing: 0.03em;
        }
        div[data-baseweb="select"] > div {
            border-radius: var(--radius-sm) !important; border-color: var(--border) !important;
            font-family: 'Plus Jakarta Sans', sans-serif !important; background-color: var(--surface) !important;
        }
        .stDateInput input, .stTextInput input {
            border-radius: var(--radius-sm) !important; font-family: 'Plus Jakarta Sans', sans-serif !important;
            border-color: var(--border) !important; background-color: var(--surface) !important;
            font-weight: 600 !important;
        }

        .stButton button, .stDownloadButton button {
            border-radius: var(--radius-sm) !important; font-weight: 600 !important;
            font-family: 'Plus Jakarta Sans', sans-serif !important; border: 1px solid var(--border) !important;
            transition: transform 0.12s ease, box-shadow 0.12s ease, background-color 0.12s ease;
        }
        .stButton button:hover, .stDownloadButton button:hover { transform: translateY(-1px); }
        .stButton button[kind="primary"], .stDownloadButton button[kind="primary"] {
            background-color: var(--primary) !important; border: none !important;
            box-shadow: 0 3px 10px rgba(79, 70, 229, 0.32) !important;
        }

        ::-webkit-scrollbar { width: 9px; height: 9px; }
        ::-webkit-scrollbar-thumb { background: #C9CEDA; border-radius: 8px; }
        ::-webkit-scrollbar-track { background: transparent; }

        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #FFFFFF 0%, #F7F8FD 100%) !important;
            border-right: 1px solid var(--border) !important;
            box-shadow: 2px 0 12px rgba(15,18,34,0.04);
            min-width: 300px !important;
            max-width: 320px !important;
        }
        section[data-testid="stSidebar"] > div { padding-top: 1.2rem !important; }
        section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: 0.4rem !important; }
        /* Mọi khối con (cả tiêu đề lẫn nút) trong sidebar dùng chung 1 mốc lề trái duy nhất */
        section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div,
        section[data-testid="stSidebar"] [data-testid="stElementContainer"] {
            margin-left: 0 !important;
            padding-left: 0 !important;
        }
        .sidebar-heading {
            font-size: 11.5px; font-weight: 800; color: var(--text-muted);
            text-transform: uppercase; letter-spacing: 0.08em;
            margin: 0 !important; padding: 0 0 10px 12px;
            display: flex; align-items: center; gap: 6px;
        }
        .nav-group-label {
            font-size: 11px; font-weight: 800; color: var(--text-faint);
            text-transform: uppercase; letter-spacing: 0.06em;
            margin: 0 !important; padding: 16px 0 4px 12px;
        }

        /* Menu điều hướng dạng cây ở sidebar: nút phẳng, thẳng hàng lề trái, không icon */
        section[data-testid="stSidebar"] div.stButton,
        section[data-testid="stSidebar"] div[data-testid="stButton"] {
            width: 100% !important; margin: 0 !important; padding: 0 !important;
        }
        section[data-testid="stSidebar"] div.stButton > button,
        section[data-testid="stSidebar"] div[data-testid="stButton"] > button {
            background-color: transparent !important;
            border: none !important;
            box-shadow: none !important;
            display: flex !important;
            justify-content: flex-start !important;
            align-items: center !important;
            text-align: left !important;
            font-family: 'Plus Jakarta Sans', sans-serif !important;
            font-weight: 600 !important;
            font-size: 13.5px !important;
            color: var(--text) !important;
            padding: 9px 12px !important;
            min-height: 38px !important;
            width: 100% !important;
            border-radius: 10px !important;
            margin: 0 0 2px 0 !important;
            transition: background-color 0.12s ease, color 0.12s ease;
        }
        /* Bỏ mọi căn giữa / margin ẩn bên trong nút (thẻ p, div, span con) để chữ luôn bám sát lề trái */
        section[data-testid="stSidebar"] div.stButton > button *,
        section[data-testid="stSidebar"] div[data-testid="stButton"] > button * {
            text-align: left !important;
            justify-content: flex-start !important;
            margin: 0 !important;
            padding: 0 !important;
            width: auto !important;
        }
        section[data-testid="stSidebar"] div.stButton > button:hover {
            background-color: var(--primary-soft) !important;
            color: var(--primary-dark) !important;
            transform: none !important;
        }
        section[data-testid="stSidebar"] div.stButton > button[kind="primary"] {
            background-color: var(--primary) !important;
            color: #FFFFFF !important;
            box-shadow: none !important;
            font-weight: 700 !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)


def _html(raw):
  return "".join(line.strip() for line in raw.strip().splitlines())


def render_page_header(logo_text, title, meta_text):
  st.markdown(
      _html(f"""
      <div class="page-header">
          <div>
              <p class="ph-logo">{logo_text}</p>
              <p class="ph-title">{title}</p>
          </div>
          <div class="ph-meta">📅 {meta_text}</div>
      </div>
      """),
      unsafe_allow_html=True,
  )


def render_section_heading(text):
  st.markdown(
      f'<div class="section-heading">{text}</div>', unsafe_allow_html=True
  )


def render_kpi_cards(items):
  cards_html = ""
  for it in items:
    color = it.get("color", "#4F46E5")
    color_soft = it.get("color_soft", "#EEF0FF")
    sub = it.get("sub", "")
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    cards_html += _html(f"""
        <div class="kpi-card" style="--accent: {color}; --accent-soft: {color_soft};">
            <div class="kpi-head">
                <span class="kpi-label">{it['label']}</span>
                <span class="kpi-icon">{it.get('icon', '📌')}</span>
            </div>
            <div class="kpi-value">{it['value']}</div>
            {sub_html}
        </div>
    """)
  st.markdown(f'<div class="kpi-row">{cards_html}</div>', unsafe_allow_html=True)


def chart_card_open(title, caption=""):
  cap_html = f'<span class="chart-card-caption">{caption}</span>' if caption else ""
  st.markdown(
      _html(f"""<div class="chart-card">
          <div class="chart-card-header">
              <span class="chart-card-title">{title}</span>
              {cap_html}
          </div>"""),
      unsafe_allow_html=True,
  )


def chart_card_close():
  st.markdown("</div>", unsafe_allow_html=True)


PHAN_HE_OPTIONS = ["DAU_VAO", "CO_KHI", "TU_TI", "CONG_TO", "TTTB_CNC"]


def render_catalog_manager(table_name, name_col, item_label):
  """Trang quản trị dùng chung cho danh mục Loại Lỗi / Công Việc Con:
  thêm nhanh 1 dòng, nhập hàng loạt từ Excel, và sửa/xoá trực tiếp trên bảng.
  """
  st.markdown(f"##### ➕ Thêm Nhanh {item_label} Mới")
  col_add1, col_add2, col_add3 = st.columns([1, 1.6, 1])
  with col_add1:
    add_ph = st.selectbox(
        "Chọn Xưởng / Phân hệ:", PHAN_HE_OPTIONS, key=f"add_ph_{table_name}"
    )
  with col_add2:
    add_name = st.text_input(
        f"Tên {item_label.lower()} mới:",
        placeholder=f"Gõ tên {item_label.lower()}...",
        key=f"add_name_{table_name}",
    )
  with col_add3:
    st.markdown("<div style='height:25px;'></div>", unsafe_allow_html=True)
    if st.button(
        f"➕ Thêm {item_label}",
        type="primary",
        use_container_width=True,
        key=f"btn_add_{table_name}",
    ):
      if add_name.strip():
        try:
          conn = get_db_connection()
          cursor = conn.cursor()
          cursor.execute(
              f"INSERT INTO {table_name} (phan_he, {name_col}) VALUES (?, ?)",
              (add_ph, add_name.strip()),
          )
          conn.commit()
          conn.close()
          st.success(f"✅ Đã thêm: {add_name.strip()}")
          st.rerun()
        except Exception as ex:
          st.error(f"Lỗi thêm mới: {ex}")
      else:
        st.warning("⚠️ Vui lòng nhập tên!")

  st.markdown(
      "<hr style='margin:16px 0; border-color:#E4E8F0;'>",
      unsafe_allow_html=True,
  )

  # ---------- NHẬP HÀNG LOẠT TỪ EXCEL ----------
  st.markdown("##### 📥 Nhập Danh Sách Hàng Loạt Từ Excel")
  col_up1, col_up2 = st.columns([2, 1])
  with col_up1:
    up_file = st.file_uploader(
        f"Chọn file Excel (.xlsx) — cần đúng 2 cột 'Phân hệ' và '{item_label}'",
        type=["xlsx"],
        key=f"upload_{table_name}",
    )
  with col_up2:
    template_df = pd.DataFrame({
        "Phân hệ": ["CO_KHI"],
        item_label: [f"Ví dụ {item_label.lower()}..."],
    })
    tpl_buf = io.BytesIO()
    with pd.ExcelWriter(tpl_buf, engine="openpyxl") as writer:
      template_df.to_excel(writer, index=False, sheet_name="Mau")
    st.markdown("<div style='height:29px;'></div>", unsafe_allow_html=True)
    st.download_button(
        "📄 Tải File Mẫu",
        data=tpl_buf.getvalue(),
        file_name=f"Mau_{item_label.replace(' ', '_')}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument"
            ".spreadsheetml.sheet"
        ),
        use_container_width=True,
        key=f"tpl_{table_name}",
    )

  if up_file is not None:
    try:
      df_up = pd.read_excel(up_file)
      df_up.columns = [str(c).strip() for c in df_up.columns]
      if not {"Phân hệ", item_label}.issubset(set(df_up.columns)):
        st.error(
            "❌ File Excel cần có đúng 2 cột tiêu đề: 'Phân hệ' và"
            f" '{item_label}' (tải file mẫu ở trên để đúng định dạng)."
        )
      else:
        df_up = df_up[["Phân hệ", item_label]].dropna(how="all")
        df_up["Phân hệ"] = df_up["Phân hệ"].astype(str).str.strip().str.upper()
        df_up[item_label] = df_up[item_label].astype(str).str.strip()
        df_up = df_up[(df_up["Phân hệ"] != "") & (df_up[item_label] != "")]
        df_valid = df_up[df_up["Phân hệ"].isin(PHAN_HE_OPTIONS)]
        n_invalid = len(df_up) - len(df_valid)

        st.markdown(
            f"**Xem trước: {len(df_valid):,} dòng hợp lệ / {len(df_up):,}"
            " dòng trong file**"
        )
        if not df_valid.empty:
          st.dataframe(df_valid, use_container_width=True, hide_index=True)
        if n_invalid > 0:
          st.warning(
              f"⚠️ Bỏ qua {n_invalid} dòng có 'Phân hệ' không hợp lệ (phải là"
              f" một trong: {', '.join(PHAN_HE_OPTIONS)})."
          )

        if not df_valid.empty and st.button(
            f"✅ Xác Nhận Nhập {len(df_valid):,} Dòng Vào Danh Mục",
            type="primary",
            key=f"confirm_up_{table_name}",
        ):
          try:
            conn = get_db_connection()
            cursor = conn.cursor()
            for _, row in df_valid.iterrows():
              cursor.execute(
                  f"INSERT INTO {table_name} (phan_he, {name_col}) VALUES"
                  " (?, ?)",
                  (row["Phân hệ"], row[item_label]),
              )
            conn.commit()
            conn.close()
            st.success(f"🎉 Đã nhập thành công {len(df_valid):,} dòng!")
            st.rerun()
          except Exception as ex:
            st.error(f"Lỗi nhập dữ liệu: {ex}")
    except Exception as ex:
      st.error(f"Lỗi đọc file Excel: {ex}")

  st.markdown(
      "<hr style='margin:16px 0; border-color:#E4E8F0;'>",
      unsafe_allow_html=True,
  )

  # ---------- BẢNG DANH MỤC HIỆN CÓ: SỬA / XOÁ TRỰC TIẾP ----------
  st.markdown(
      f"##### 📜 Danh Mục {item_label} Hiện Có — sửa hoặc xoá dòng rồi bấm"
      " Lưu"
  )
  try:
    conn = get_db_connection()
    df_list = pd.read_sql_query(
        f"SELECT id, phan_he, {name_col} FROM {table_name} ORDER BY phan_he"
        f" ASC, {name_col} ASC",
        conn,
    )
    conn.close()
  except Exception as ex:
    df_list = pd.DataFrame(columns=["id", "phan_he", name_col])
    st.error(f"Lỗi tải danh mục: {ex}")

  if df_list.empty:
    st.info(
        f"💡 Chưa có {item_label.lower()} nào trong danh mục. Thêm mới ở"
        " trên hoặc nhập từ Excel."
    )
    return

  df_list_display = df_list.rename(
      columns={"id": "ID", "phan_he": "Phân Hệ / Xưởng", name_col: item_label}
  )

  edited = st.data_editor(
      df_list_display,
      use_container_width=True,
      hide_index=True,
      num_rows="dynamic",
      column_config={
          "ID": st.column_config.NumberColumn("ID", disabled=True),
          "Phân Hệ / Xưởng": st.column_config.SelectboxColumn(
              "Phân Hệ / Xưởng", options=PHAN_HE_OPTIONS
          ),
      },
      key=f"editor_{table_name}",
  )

  if st.button(
      "💾 Lưu Thay Đổi Danh Mục", type="primary", key=f"save_{table_name}"
  ):
    try:
      conn = get_db_connection()
      cursor = conn.cursor()
      original_ids = set(df_list_display["ID"].dropna().astype(int))
      edited_ids = set(edited["ID"].dropna().astype(int))
      for did in original_ids - edited_ids:
        cursor.execute(f"DELETE FROM {table_name} WHERE id = ?", (int(did),))
      for _, row in edited.iterrows():
        ph = str(row["Phân Hệ / Xưởng"]).strip()
        nm = str(row[item_label]).strip()
        if not nm or not ph:
          continue
        if pd.isna(row["ID"]):
          cursor.execute(
              f"INSERT INTO {table_name} (phan_he, {name_col}) VALUES (?,"
              " ?)",
              (ph, nm),
          )
        else:
          cursor.execute(
              f"UPDATE {table_name} SET phan_he = ?, {name_col} = ? WHERE"
              " id = ?",
              (ph, nm, int(row["ID"])),
          )
      conn.commit()
      conn.close()
      st.success("✅ Đã lưu thay đổi danh mục!")
      st.rerun()
    except Exception as ex:
      st.error(f"Lỗi lưu thay đổi: {ex}")


COLOR_SUCCESS = "#0EA968"
COLOR_PRIMARY = "#4F46E5"
COLOR_DANGER = "#E23D4D"
COLOR_WARNING = "#EA8A0A"
COLOR_PURPLE = "#9333EA"
COLOR_TEXT = "#0F1222"

DISTINCT_COLORS = [
    "#4F46E5",
    "#0EA968",
    "#EA8A0A",
    "#E23D4D",
    "#9333EA",
    "#0891B2",
]

PLOTLY_FONT = "Plus Jakarta Sans, -apple-system, Segoe UI, sans-serif"
PLOTLY_GRID = "#F1F2F6"
PLOTLY_AXIS_TEXT = "#6B7280"


def clean_emoji(text):
  return re.sub(r"[^\w\s\(\)\-\/\.\,\:]", "", str(text)).strip()


LENH_PREFIX_TO_PHAN_HE = {
    "3011": "TU_TI",
    "3012": "CO_KHI",
    "3013": "CONG_TO",
    "3014": "CONG_TO",
    "3016": "TTTB_CNC",
}


def derive_phan_he(lenh_sx):
  """Xác định đúng xưởng theo 4 số đầu lệnh sản xuất:
  3011=TU_TI, 3012=CO_KHI, 3013/3014=CONG_TO, 3016=TTTB_CNC,
  các đầu lệnh khác = KHAC (lệnh bảo hành / cải tạo)."""
  prefix = str(lenh_sx).strip()[:4]
  return LENH_PREFIX_TO_PHAN_HE.get(prefix, "KHAC")


def _normalize_ten(ten_tp):
  return re.sub(r"[^A-Z0-9]", "", str(ten_tp).upper())


# ================= 3B. PHÂN LOẠI DÒNG SẢN PHẨM THEO BẢNG 1 / BẢNG 2 =================
def classify_cong_to(ma_tp, ten_tp):
  """Dòng sản phẩm xưởng Công Tơ (3013, 3014) theo Bảng 1 + carve-out AMI (Bảng 2)."""
  ma_tp = str(ma_tp).strip()
  ten_upper = str(ten_tp).strip().upper()
  ten_norm = _normalize_ten(ten_tp)
  is_dau5 = ma_tp.startswith("5")
  is_dau4 = ma_tp.startswith("4")

  # AMI: các mã ME41A / ME42A / CE14A tách riêng khỏi ME / CE (Bảng 2)
  if "ME41A" in ten_norm or "ME42A" in ten_norm or "CE14A" in ten_norm:
    return "AMI"

  prefix2 = ten_upper[:2]
  if prefix2 == "CE":
    if is_dau5:
      return "CE"
    if is_dau4 and (ten_upper.endswith("KSS") or ten_upper.endswith("CXX")):
      return "CE"
  if prefix2 == "ME":
    if is_dau5:
      return "ME"
    if is_dau4 and (ten_upper.endswith("KSS") or ten_upper.endswith("CXX")):
      return "ME"
  if prefix2 == "VA":
    return "VA"
  if prefix2 == "EW":
    if is_dau5:
      return "EW"
    if is_dau4 and ten_upper.endswith("CKC"):
      return "EW"
  if prefix2 in ("DC", "MD", "PM", "MO", "HU"):
    if prefix2 == "MD" and "WM" in ten_upper:
      return "TBTT ĐHN"
    return "TBTT Công tơ"

  # Bán thành phẩm mô tả chung (BP, CỤ, BỘ, NẮ...): tìm mã con Bảng 2 trong tên
  for code in ["CE18", "CE38", "CE14", "CE28", "CE58"]:
    if code in ten_norm:
      return "CE"
  for code in ["ME40", "ME41", "ME42", "ME43"]:
    if code in ten_norm:
      return "ME"
  return "Khác"


def classify_tuti(ma_tp, ten_tp):
  """Dòng sản phẩm xưởng TU/TI (3011) theo Bảng 1."""
  ten_upper = str(ten_tp).strip().upper()
  
  # 1. KIỂM TRA 2 KÝ TỰ ĐẦU TIÊN TRƯỚC (Ưu tiên Thành phẩm)
  prefix2 = ten_upper[:2]
  if prefix2 in ["HT", "HB", "VT"]:
    return prefix2
  if prefix2 == "CT":
    return "CT"
  if prefix2 == "PT":
    return "PT"

  # 2. SAU ĐÓ MỚI KIỂM TRA 4 KÝ TỰ (LPVT, LPCT)
  prefix4 = ten_upper[:4]
  if prefix4 == "LPVT":
    return "LPVT"
  if prefix4 == "LPCT":
    return "LPCT"

  # 3. CUỐI CÙNG MỚI QUÉT TÌM BÁN THÀNH PHẨM (kiểm tra bằng cách cắt đúng 3 ký tự đầu)
  prefix3 = ten_upper[:3]
  if prefix3 in ["CT1", "CT2", "CT4", "CT5", "CT7", "CT8"]:
    return "CT"
  if prefix3 in ["CT3", "CT6", "CT9"]:
    return "TI"
  if prefix3 in ["PT1", "PT2", "PT4", "PT5", "PT7", "PT8"]:
    return "PT"
  if prefix3 in ["PT3", "PT6", "PT9"]:
    return "TU"
  if prefix3 in ["HT1", "HT2", "HT3"]:
    return "HT"

  return prefix2 if len(prefix2) > 0 else "N/A"


def classify_tttb_cnc(ma_tp, ten_tp):
  """Dòng sản phẩm xưởng TTTB CNC (3016) theo Bảng 1."""
  ten_upper = str(ten_tp).strip().upper()
  if ten_upper[:2] == "FE":
    return "TBTT báo cháy"
  for pfx in ["TH", "HỆ", "BỘ", "TB"]:
    if ten_upper.startswith(pfx.upper()):
      return "Bộ báo cháy liên gia"
  return "Khác"


def derive_mat_prefix(phan_he, ma_tp, ten_tp):
  if phan_he == "CONG_TO":
    return classify_cong_to(ma_tp, ten_tp)
  if phan_he == "TU_TI":
    return classify_tuti(ma_tp, ten_tp)
  if phan_he == "TTTB_CNC":
    return classify_tttb_cnc(ma_tp, ten_tp)
  return None  # Giữ nguyên mat_prefix gốc cho các xưởng khác (VD: CO_KHI)


# Bảng 2 — mã con chi tiết trong từng nhóm CE / ME / AMI / EW (dùng cho biểu
# đồ "Tổng Sản Lượng Cả Năm" khi người dùng chọn đúng 1 trong 4 nhóm này)
BANG2_SUB_CODES = {
    "CE": ["CE18", "CE38", "CE14", "CE28", "CE58"],
    "ME": ["ME40", "ME41", "ME42", "ME43"],
    "AMI": ["ME41A", "ME42A", "CE14A"],
}


def classify_sub_code_bang2(dong_chinh, ten_tp):
  """Trả về mã con Bảng 2 (VD: CE18, ME41A...) cho 1 dòng sản phẩm CE/ME/AMI.
  Với EW: tự gộp nhóm theo 8 ký tự đầu tiên của tên (đúng như Bảng 2 mô tả)."""
  ten_norm = _normalize_ten(ten_tp)
  if dong_chinh == "EW":
    ten_clean = str(ten_tp).strip().upper()
    return ten_clean[:8] if len(ten_clean) >= 8 else ten_clean
  for code in BANG2_SUB_CODES.get(dong_chinh, []):
    if code in ten_norm:
      return code
  return "Khác"


@st.cache_data(ttl=15)
def load_data(tu_date, den_date):
  if not USE_TURSO and not os.path.exists(DB_PATH):
    return pd.DataFrame(), pd.DataFrame()

  tu_iso = tu_date.strftime("%Y-%m-%d 00:00:00")
  den_iso = den_date.strftime("%Y-%m-%d 23:59:59")
  try:
    conn = get_db_connection()
    df_qa32 = pd.read_sql_query(
        "SELECT * FROM tb_sap_qa32 WHERE ngay_ve_dt >= ? AND ngay_ve_dt <= ?",
        conn,
        params=(tu_iso, den_iso),
    )
    df_coois = pd.read_sql_query(
        "SELECT * FROM tb_sap_coois WHERE ngay_lenh_dt >= ? AND ngay_lenh_dt"
        " <= ?",
        conn,
        params=(tu_iso, den_iso),
    )
    conn.close()
    if not df_coois.empty and "lenh_sx" in df_coois.columns:
      df_coois["phan_he"] = df_coois["lenh_sx"].apply(derive_phan_he)
      new_mat_prefix = df_coois.apply(
          lambda r: derive_mat_prefix(
              r["phan_he"], r.get("ma_tp", ""), r.get("ten_tp", "")
          ),
          axis=1,
      )
      df_coois["mat_prefix"] = new_mat_prefix.where(
          new_mat_prefix.notna(), df_coois.get("mat_prefix")
      )
    return df_qa32, df_coois
  except Exception:
    return pd.DataFrame(), pd.DataFrame()


# ================= 4. BỘ LỌC THỜI GIAN CỐ ĐỊNH =================
st.markdown('<div class="filter-banner">', unsafe_allow_html=True)
st.markdown(
    '<div class="filter-title">📅 BỘ LỌC THỜI GIAN BÁO CÁO TOÀN HỆ THỐNG</div>',
    unsafe_allow_html=True,
)

today = date.today()
first_day_of_year = date(today.year, 1, 1)

col_f1, col_f2, col_f3, col_f4 = st.columns([1.3, 1.3, 1, 1.3])
with col_f1:
  tu_date = st.date_input(
      "Từ ngày:", first_day_of_year, format="DD/MM/YYYY"
  )
with col_f2:
  den_date = st.date_input("Đến ngày:", today, format="DD/MM/YYYY")
with col_f3:
  st.markdown("<div style='height: 25px;'></div>", unsafe_allow_html=True)
  if st.button("🔄 CẬP NHẬT BÁO CÁO", use_container_width=True, type="primary"):
    st.cache_data.clear()
    st.rerun()
with col_f4:
  st.markdown("<div style='height: 25px;'></div>", unsafe_allow_html=True)
  if st.button("📊 XUẤT EXCEL TOÀN BỘ", use_container_width=True):
    st.session_state["trigger_export_all"] = True

if st.session_state.get("trigger_export_all"):
  try:
    conn = get_db_connection()
    tu_iso_exp = tu_date.strftime("%Y-%m-%d 00:00:00")
    den_iso_exp = den_date.strftime("%Y-%m-%d 23:59:59")

    df_exp_qa32 = pd.read_sql_query(
        "SELECT * FROM tb_sap_qa32 WHERE ngay_ve_dt >= ? AND ngay_ve_dt <= ?",
        conn,
        params=(tu_iso_exp, den_iso_exp),
    )
    df_exp_coois = pd.read_sql_query(
        "SELECT * FROM tb_sap_coois WHERE ngay_lenh_dt >= ? AND"
        " ngay_lenh_dt <= ?",
        conn,
        params=(tu_iso_exp, den_iso_exp),
    )
    df_exp_qc = pd.read_sql_query(
        "SELECT * FROM tb_qc_dau_vao WHERE ngay_kiem >= ? AND ngay_kiem <= ?"
        " ORDER BY ngay_kiem DESC",
        conn,
        params=(tu_iso_exp, den_iso_exp),
    )
    df_exp_loai_loi = pd.read_sql_query(
        "SELECT * FROM tb_dm_loai_loi ORDER BY phan_he, ten_loi", conn
    )
    df_exp_cong_viec = pd.read_sql_query(
        "SELECT * FROM tb_dm_cong_viec ORDER BY phan_he, ten_cong_viec", conn
    )
    conn.close()

    # Đảm bảo cột sl_huy có sẵn (nếu CSDL cũ chưa có)
    if "sl_huy" not in df_exp_qc.columns:
        df_exp_qc["sl_huy"] = 0.0

    # Bảng tổng hợp theo tháng (sản lượng COOIS)
    if not df_exp_coois.empty:
      df_exp_coois_m = df_exp_coois.copy()
      df_exp_coois_m["Tháng"] = pd.to_datetime(
          df_exp_coois_m["ngay_lenh_dt"], errors="coerce"
      ).dt.strftime("%m/%Y")
      df_sum_thang = (
          df_exp_coois_m.groupby(["Tháng", "phan_he"])
          .agg(
              So_Luong_Lenh=("lenh_sx", "count"),
              Tong_SL=("sl_tong", "sum"),
              SL_Hoan_Thanh=("sl_ht", "sum"),
          )
          .reset_index()
      )
    else:
      df_sum_thang = pd.DataFrame()

    # Bảng tổng hợp sai hỏng theo kiểu lỗi
    if not df_exp_qc.empty:
      df_sum_loi = (
          df_exp_qc[df_exp_qc["kieu_loi"].astype(str).str.strip() != ""]
          .groupby(["kieu_loi"])
          .agg(
              So_Luong_Loi=("sl_khong_dat", "sum"),
              So_Luong_Huy=("sl_huy", "sum"),
              So_Ca_Bao_Loi=("id", "count"),
          )
          .reset_index()
          .sort_values("So_Luong_Loi", ascending=False)
      )
      df_sum_ns = (
          df_exp_qc.groupby(["nguoi_kiem"])
          .agg(
              So_Luot_Kiem=("id", "count"),
              Tong_SL_Kiem=("sl_kiem", "sum"),
              Tong_SL_Loi=("sl_khong_dat", "sum"),
              Tong_SL_Huy=("sl_huy", "sum"),
          )
          .reset_index()
          .sort_values("So_Luot_Kiem", ascending=False)
      )
    else:
      df_sum_loi = pd.DataFrame()
      df_sum_ns = pd.DataFrame()

    buf_all = io.BytesIO()
    with pd.ExcelWriter(buf_all, engine="openpyxl") as writer:
      df_exp_qa32.to_excel(writer, sheet_name="QA32_ThoDuLieu", index=False)
      df_exp_coois.to_excel(writer, sheet_name="COOIS_ThoDuLieu", index=False)
      df_exp_qc.to_excel(writer, sheet_name="QC_BaoCao_ThoDuLieu", index=False)
      df_sum_thang.to_excel(
          writer, sheet_name="TongHop_SanLuongTheoThang", index=False
      )
      df_sum_loi.to_excel(
          writer, sheet_name="TongHop_SaiHongTheoLoi", index=False
      )
      df_sum_ns.to_excel(
          writer, sheet_name="TongHop_NangSuatNhanSu", index=False
      )
      df_exp_loai_loi.to_excel(
          writer, sheet_name="DanhMuc_LoaiLoi", index=False
      )
      df_exp_cong_viec.to_excel(
          writer, sheet_name="DanhMuc_CongViecCon", index=False
      )

    st.download_button(
        label=(
            "📥 Tải file Excel tổng hợp toàn bộ"
            f" ({tu_date.strftime('%d/%m/%Y')} → {den_date.strftime('%d/%m/%Y')})"
        ),
        data=buf_all.getvalue(),
        file_name=(
            f"EMIC_XuatToanBo_{tu_date.strftime('%Y%m%d')}_"
            f"{den_date.strftime('%Y%m%d')}.xlsx"
        ),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key="dl_export_all",
    )
  except Exception as ex:
    st.error(f"Lỗi xuất Excel tổng hợp: {ex}")

st.markdown("</div>", unsafe_allow_html=True)

if "nav_selected" not in st.session_state:
  st.session_state.nav_selected = "1. Báo Cáo Vật Tư"


def _nav_button(label, key):
  is_active = st.session_state.nav_selected == label
  if st.button(
      label,
      key=key,
      use_container_width=True,
      type="primary" if is_active else "secondary",
  ):
    st.session_state.nav_selected = label
    st.rerun()


with st.sidebar:
  st.markdown(
      '<div class="sidebar-heading">DANH MỤC BÁO CÁO</div>',
      unsafe_allow_html=True,
  )
  _nav_button("0. Báo Cáo Chung", "nav_0")
  _nav_button("1. Báo Cáo Vật Tư", "nav_1")
  _nav_button("2. Báo Cáo Xưởng Cơ Khí", "nav_2")
  _nav_button("3. Báo Cáo Xưởng Công Tơ", "nav_3")
  _nav_button("4. Báo Cáo Xưởng TU/TI", "nav_4")
  _nav_button("5. Báo Cáo Xưởng TTTB CNC", "nav_5")

  st.markdown(
      '<div class="nav-group-label">6. Báo Cáo Tổng Hợp</div>',
      unsafe_allow_html=True,
  )
  _nav_button("6.1 Danh Sách Vật Tư", "nav_51")
  _nav_button("6.2 Danh Sách Lệnh Sản Xuất", "nav_52")
  _nav_button("6.3 Sai Hỏng", "nav_53")
  _nav_button("6.4 Năng Suất", "nav_54")

  st.markdown(
      '<div class="nav-group-label">7. Cài Đặt</div>',
      unsafe_allow_html=True,
  )
  _nav_button("7.1 Lỗi Sai Hỏng", "nav_61")
  _nav_button("7.2 Công Việc Con", "nav_62")
  _nav_button("7.3 Mục tiêu chất lượng", "nav_63")

  nav = st.session_state.nav_selected

  st.markdown(
      '<div class="sidebar-heading" style="margin-top:18px;">CẤU HÌNH DỰ'
      " PHÒNG</div>",
      unsafe_allow_html=True,
  )
  st.info("💡 Bạn cũng có thể chọn ngày ở đây:")
  sb_tu = st.date_input(
      "Từ ngày (Sidebar)", tu_date, key="sb_tu", format="DD/MM/YYYY"
  )
  sb_den = st.date_input(
      "Đến ngày (Sidebar)", den_date, key="sb_den", format="DD/MM/YYYY"
  )
  if sb_tu != tu_date or sb_den != den_date:
    tu_date, den_date = sb_tu, sb_den

DANH_SACH_LABELS = (
    "6.1 Danh Sách Vật Tư",
    "6.2 Danh Sách Lệnh Sản Xuất",
    "6.3 Sai Hỏng",
    "6.4 Năng Suất",
)
CAI_DAT_LABELS = ("7.1 Lỗi Sai Hỏng", "7.2 Công Việc Con", "7.3 Mục tiêu chất lượng")
PROTECTED_LABELS = DANH_SACH_LABELS + CAI_DAT_LABELS

is_authenticated = True
st.session_state["admin_authenticated"] = True

df_qa32, df_coois = load_data(tu_date, den_date)

render_page_header(
    "EMIC - Phòng Quản Lý Chất Lượng",
    "Phần Mềm Báo Cáo Chất Lượng",
    f"{tu_date.strftime('%d/%m/%Y')} → {den_date.strftime('%d/%m/%Y')}",
)


# ================= 5. NAVIGATION: MENU DẠNG CÂY Ở SIDEBAR =================

# ================= 6. TAB 1: BÁO CÁO VẬT TƯ (NGUYÊN BẢN CÓ ĐỦ BIỂU ĐỒ) =================
# ================= 5B. TAB 0: BÁO CÁO CHUNG (CHỜ BỔ SUNG NỘI DUNG) =================
if nav == "0. Báo Cáo Chung":
  render_section_heading("📌 BÁO CÁO CHUNG")
  st.info(
      "💡 Trang này đang để trống, sẵn sàng bổ sung nội dung theo yêu cầu"
      " tiếp theo của bạn."
  )

if nav == "1. Báo Cáo Vật Tư":
  if df_qa32.empty:
    st.info("💡 Chưa có dữ liệu QA32 trong khoảng thời gian đã chọn.")
  else:
    months_labels = [f"T{i}" for i in range(1, 13)]
    ud01_m, ud02_m, ud03_m, uninspected_m = (
        [0] * 12,
        [0] * 12,
        [0] * 12,
        [0] * 12,
    )
    ft_qty_m, by_inspected_m, by_uninspected_m = (
        [0.0] * 12,
        [0.0] * 12,
        [0.0] * 12,
    )
    total_ca_block, total_ft_all = 0.0, 0.0
    top_block_dict = {}

    for _, r in df_qa32.iterrows():
      try:
        m_idx = (
            datetime.strptime(
                str(r["ngay_ve_dt"]).split()[0], "%Y-%m-%d"
            ).month
            - 1
        )
      except Exception:
        m_idx = 0
      if not (0 <= m_idx < 12):
        m_idx = 0

      st_clean = (
          str(r["xac_nhan_sap"]).strip().upper().replace(" ", "")
          if "xac_nhan_sap" in r and pd.notna(r["xac_nhan_sap"])
          else ""
      )
      ft_val = (
          float(r["ft_qty"])
          if ("ft_qty" in r and pd.notna(r["ft_qty"]))
          else 0.0
      )
      by_val = (
          float(r["by_sample"])
          if ("by_sample" in r and pd.notna(r["by_sample"]))
          else 0.0
      )
      ca_val = (
          float(r["ca_qty"])
          if ("ca_qty" in r and pd.notna(r["ca_qty"]))
          else 0.0
      )

      ft_qty_m[m_idx] += ft_val
      total_ft_all += ft_val
      total_ca_block += ca_val

      is_uninspected = (
          "CHƯA" in st_clean or not st_clean or st_clean in ["NAN", "NONE", "❌CHƯAXN"]
      )
      is_ud02 = any(
          k in st_clean for k in ["02", "UD2", "ĐẶCNHƯỢNG", "DACNHUONG"]
      )
      is_ud03 = any(
          k in st_clean
          for k in ["03", "UD3", "TRẢLẠI", "TRALAI", "TỪCHỐI", "TUCHOI", "KHÔNG", "KHONG"]
      )
      is_ud01 = any(k in st_clean for k in ["01", "UD1", "ĐẠT", "DAT"]) and not (
          is_ud02 or is_ud03
      )

      if is_uninspected:
        uninspected_m[m_idx] += 1
        by_uninspected_m[m_idx] += by_val
      else:
        by_inspected_m[m_idx] += by_val
        if is_ud02:
          ud02_m[m_idx] += 1
        elif is_ud03:
          ud03_m[m_idx] += 1
        else:
          ud01_m[m_idx] += 1

      ma_vt_str = (
          str(r["ma_vt"]).strip()
          if "ma_vt" in r and pd.notna(r["ma_vt"])
          else ""
      )
      ten_vt_str = (
          str(r["ten_vt"]).strip()
          if "ten_vt" in r and pd.notna(r["ten_vt"])
          else ""
      )
      ncc_str = (
          str(r["ncc"]).strip() if "ncc" in r and pd.notna(r["ncc"]) else ""
      )

      if (
          is_ud02
          or is_ud03
          or ca_val > 0
          or (st_clean and not is_ud01 and not is_uninspected)
      ):
        if "VIHA" in ncc_str.upper():
          key = ("Mặt số công tơ", ncc_str if ncc_str else "Cty TNHH CN VIHA")
          ma_display, ten_display = "Mặt số công tơ", "Mặt số công tơ"
        else:
          key = (ma_vt_str, ncc_str)
          ma_display, ten_display = ma_vt_str, ten_vt_str

        if key not in top_block_dict:
          top_block_dict[key] = {
              "ma_vt": ma_display,
              "ten_vt": ten_display,
              "ncc": key[1],
              "ud02": 0,
              "ud03": 0,
              "ca_block": 0.0,
              "ft_total": 0.0,
          }
        if is_ud02:
          top_block_dict[key]["ud02"] += 1
        if is_ud03:
          top_block_dict[key]["ud03"] += 1
        top_block_dict[key]["ca_block"] += ca_val
        top_block_dict[key]["ft_total"] += ft_val

    total_ud01 = sum(ud01_m)
    total_ud02 = sum(ud02_m)
    total_ud03 = sum(ud03_m)
    total_uninspected = sum(uninspected_m)
    total_lots = total_ud01 + total_ud02 + total_ud03 + total_uninspected
    pct_ud01 = (total_ud01 / total_lots * 100) if total_lots > 0 else 0.0
    total_by_all = sum(by_inspected_m) + sum(by_uninspected_m)

    render_kpi_cards([
        {
            "label": "Tổng Vật Tư Về (FT)",
            "value": f"{int(total_ft_all):,}",
            "icon": "📦",
            "color": COLOR_PRIMARY,
            "color_soft": "#EEF0FF",
            "sub": f"{int(total_by_all):,} mẫu đã kiểm (BY)",
        },
        {
            "label": "Tỷ Lệ Đạt (UD 01)",
            "value": f"{pct_ud01:.1f}%",
            "icon": "✅",
            "color": COLOR_SUCCESS,
            "color_soft": "#ECFDF5",
            "sub": f"{int(total_ud01):,} / {int(total_lots):,} lô",
        },
        {
            "label": "Đặc Nhượng / Trả Lại",
            "value": f"{int(total_ud02 + total_ud03):,}",
            "icon": "⚠️",
            "color": COLOR_WARNING,
            "color_soft": "#FFF8EB",
            "sub": f"UD02: {int(total_ud02):,} · UD03: {int(total_ud03):,}",
        },
        {
            "label": "SL Bị Block (CA)",
            "value": f"{int(total_ca_block):,}",
            "icon": "🚫",
            "color": COLOR_DANGER,
            "color_soft": "#FEF2F3",
            "sub": f"{len(top_block_dict):,} mã vật tư liên quan",
        },
    ])

    PLOT_HEIGHT = 380

    col1, col2 = st.columns([2.1, 1.0])
    with col1:
      chart_card_open(
          "Số Lượng Lệnh Kiểm & Tổng Vật Tư Về / Số Mẫu Kiểm", "Theo tháng"
      )
      fig1 = make_subplots(specs=[[{"secondary_y": True}]])
      fig1.add_trace(
          go.Bar(
              x=months_labels,
              y=ud01_m,
              name="UD 01 (Đạt)",
              marker_color=COLOR_SUCCESS,
          ),
          secondary_y=False,
      )
      fig1.add_trace(
          go.Bar(
              x=months_labels,
              y=ud02_m,
              name="UD 02 (Đặc nhượng)",
              marker_color=COLOR_WARNING,
          ),
          secondary_y=False,
      )
      fig1.add_trace(
          go.Bar(
              x=months_labels,
              y=ud03_m,
              name="UD 03 (Trả lại)",
              marker_color=COLOR_DANGER,
          ),
          secondary_y=False,
      )

      total_by = [by_inspected_m[i] + by_uninspected_m[i] for i in range(12)]
      fig1.add_trace(
          go.Scatter(
              x=months_labels,
              y=total_by,
              name="Số mẫu phải kiểm (BY)",
              mode="lines+markers+text",
              line=dict(color=COLOR_PURPLE, width=2, dash="dash"),
              text=[f"{int(v):,}" if v > 0 else "" for v in total_by],
              textposition="top center",
              textfont=dict(size=9.5, family=PLOTLY_FONT, color=COLOR_PURPLE),
          ),
          secondary_y=True,
      )
      fig1.add_trace(
          go.Scatter(
              x=months_labels,
              y=ft_qty_m,
              name="Tổng số hàng về (FT)",
              mode="lines+markers+text",
              line=dict(color=COLOR_PRIMARY, width=2),
              text=[f"{int(v):,}" if v > 0 else "" for v in ft_qty_m],
              textposition="bottom center",
              textfont=dict(size=9.5, family=PLOTLY_FONT, color=COLOR_PRIMARY),
          ),
          secondary_y=True,
      )

      total_lots_m = [ud01_m[i] + ud02_m[i] + ud03_m[i] for i in range(12)]
      for i in range(12):
        if total_lots_m[i] > 0:
          fig1.add_annotation(
              x=months_labels[i],
              y=total_lots_m[i],
              text=f"<b>{int(total_lots_m[i]):,}</b>",
              showarrow=False,
              yshift=12,
              font=dict(size=10.5, family=PLOTLY_FONT, color=COLOR_TEXT),
              yref="y1",
          )

      fig1.update_layout(
          font=dict(family=PLOTLY_FONT, color=PLOTLY_AXIS_TEXT),
          barmode="stack",
          margin=dict(l=30, r=20, t=8, b=55),
          height=PLOT_HEIGHT,
          paper_bgcolor="#FFFFFF",
          plot_bgcolor="#FFFFFF",
          legend=dict(
              orientation="h",
              yanchor="top",
              y=-0.22,
              xanchor="center",
              x=0.5,
              font=dict(size=11, family=PLOTLY_FONT, color="#6B7280"),
          ),
      )
      fig1.update_xaxes(
          showgrid=False, tickfont=dict(size=11, family=PLOTLY_FONT, color="#6B7280")
      )
      fig1.update_yaxes(
          title_text="← Số Lượng Lệnh",
          title_font=dict(size=12, color=COLOR_PRIMARY),
          tickformat=",d",
          secondary_y=False,
          showgrid=True,
          gridcolor=PLOTLY_GRID,
          zeroline=False,
      )
      fig1.update_yaxes(
          title_text="Vật Tư / Mẫu (Log) →",
          title_font=dict(size=12, color=COLOR_PURPLE),
          type="log",
          dtick=1,
          tickformat="~s",
          secondary_y=True,
          showgrid=False,
      )

      st.plotly_chart(
          fig1,
          use_container_width=True,
          config={"displayModeBar": False},
          key="vt_chart_fig1",
      )
      chart_card_close()

    with col2:
      chart_card_open("Tỷ Lệ Vật Tư Đạt vs Bị Block Lỗi")
      ok_cnt = max(0.0, total_ft_all - total_ca_block)
      pct_ok = (ok_cnt / total_ft_all * 100) if total_ft_all > 0 else 0
      pct_block = (
          (total_ca_block / total_ft_all * 100) if total_ft_all > 0 else 0
      )

      fig2 = go.Figure(
          data=[
              go.Pie(
                  labels=["Vật tư Đạt", "Bị Block (Lỗi)"],
                  values=[ok_cnt, total_ca_block],
                  hole=0.6,
                  marker=dict(
                      colors=[COLOR_SUCCESS, COLOR_DANGER],
                      line=dict(color="#FFFFFF", width=3),
                  ),
                  text=[
                      f"<b>Vật tư Đạt</b><br>{pct_ok:.1f}%<br>({ok_cnt:,.0f})",
                      f"<b>Bị Block (Lỗi)</b><br>{pct_block:.1f}%<br>({total_ca_block:,.0f})",
                  ],
                  textinfo="text",
                  textposition="outside",
                  textfont=dict(
                      size=12,
                      family=PLOTLY_FONT,
                      color=[COLOR_SUCCESS, COLOR_DANGER],
                  ),
                  direction="clockwise",
                  sort=False,
              )
          ]
      )
      fig2.update_layout(
          font=dict(family=PLOTLY_FONT, color=PLOTLY_AXIS_TEXT),
          margin=dict(l=95, r=95, t=30, b=30),
          height=PLOT_HEIGHT,
          paper_bgcolor="#FFFFFF",
          showlegend=False,
          annotations=[
              dict(
                  text=(
                      "<b>TỔNG VẬT TƯ"
                      " VỀ</b><br><span"
                      f" style='font-size:16px'>{int(total_ft_all):,}</span>"
                  ),
                  x=0.5,
                  y=0.5,
                  font_size=12,
                  font_family=PLOTLY_FONT,
                  showarrow=False,
              )
          ],
      )
      st.plotly_chart(
          fig2,
          use_container_width=True,
          config={"displayModeBar": False},
          key="vt_chart_fig2",
      )
      chart_card_close()

    sorted_blocks = sorted(
        top_block_dict.values(),
        key=lambda x: (x["ud03"] + x["ud02"], x["ca_block"], x["ft_total"]),
        reverse=True,
    )
    if sorted_blocks:
      col_rank, col_table = st.columns([1.0, 1.6])

      supplier_agg = {}
      for item in sorted_blocks:
        ncc_name = item["ncc"].strip() if item["ncc"] else "Không rõ NCC"
        supplier_agg[ncc_name] = (
            supplier_agg.get(ncc_name, 0.0) + item["ca_block"]
        )
      top_suppliers = sorted(
          supplier_agg.items(), key=lambda x: x[1], reverse=True
      )[:8]

      with col_rank:
        if top_suppliers:
          sup_names = [clean_emoji(s[0])[:28] for s in top_suppliers][::-1]
          sup_vals = [s[1] for s in top_suppliers][::-1]
          chart_card_open(
              "Top Nhà Cung Cấp Bị Block Nhiều Nhất", "Theo tổng SL (CA)"
          )
          fig_sup = go.Figure(
              go.Bar(
                  x=sup_vals,
                  y=sup_names,
                  orientation="h",
                  marker=dict(color=COLOR_DANGER, cornerradius=6),
                  text=[f"{v:,.0f}" for v in sup_vals],
                  textposition="outside",
                  textfont=dict(
                      size=11, family=PLOTLY_FONT, color=COLOR_DANGER
                  ),
              )
          )
          fig_sup.update_layout(
              font=dict(family=PLOTLY_FONT, color=PLOTLY_AXIS_TEXT),
              margin=dict(l=10, r=55, t=10, b=10),
              height=max(230, 32 * len(sup_names)),
              paper_bgcolor="#FFFFFF",
              plot_bgcolor="#FFFFFF",
              showlegend=False,
          )
          fig_sup.update_xaxes(
              showgrid=True,
              gridcolor=PLOTLY_GRID,
              tickfont=dict(
                  size=10.5, family=PLOTLY_FONT, color=PLOTLY_AXIS_TEXT
              ),
          )
          fig_sup.update_yaxes(
              showgrid=False,
              automargin=True,
              tickfont=dict(size=11, family=PLOTLY_FONT, color=COLOR_TEXT),
          )
          st.plotly_chart(
              fig_sup,
              use_container_width=True,
              config={"displayModeBar": False},
              key="vt_chart_fig_sup",
          )
          chart_card_close()

      with col_table:
        chart_card_open(
            "🚨 Danh Sách Vật Tư Bị Block & UD02, UD03",
            f"{len(sorted_blocks)} mã vật tư",
        )
        df_block = pd.DataFrame(sorted_blocks)
        df_block["Tổng SL Block (CA) / SL Về"] = df_block.apply(
            lambda r: f"{r['ca_block']:,.0f} / {r['ft_total']:,.0f}", axis=1
        )
        df_block = df_block[[
            "ma_vt",
            "ten_vt",
            "ncc",
            "ud02",
            "ud03",
            "Tổng SL Block (CA) / SL Về",
        ]]
        df_block.columns = [
            "Mã Vật Tư",
            "Tên Vật Tư",
            "Nhà Cung Cấp",
            "Số Lượt UD 02",
            "Số Lượt UD 03",
            "Tổng SL Block (CA) / SL Về",
        ]
        styled_block = (
            df_block.style.background_gradient(
                subset=["Số Lượt UD 02"], cmap="Oranges", vmin=0
            )
            .background_gradient(
                subset=["Số Lượt UD 03"], cmap="Reds", vmin=0
            )
            .format({"Số Lượt UD 02": "{:.0f}", "Số Lượt UD 03": "{:.0f}"})
        )
        st.dataframe(
            styled_block, use_container_width=True, hide_index=True, height=360
        )
        chart_card_close()
    else:
      st.success("🎉 Không có vật tư nào bị Block hoặc UD 02, 03")


# ================= 7. HÀM COOIS TÍCH HỢP TỰ ĐỘNG SỐ LIỆU SAI HỎNG =================
def render_coois_tab_layout(phan_he_code, title_text):
  df_sub = (
      df_coois[df_coois["phan_he"] == phan_he_code]
      if not df_coois.empty
      else pd.DataFrame()
  )
  if df_sub.empty:
    st.info(f"💡 Chưa có dữ liệu sản xuất cho phân hệ {title_text}.")
    return

  # Truy xuất mục tiêu chất lượng
  try:
    conn = get_db_connection()
    df_targets = pd.read_sql_query("SELECT mat_prefix, muc_tieu FROM tb_dm_muc_tieu WHERE phan_he = ?", conn, params=(phan_he_code,))
    conn.close()
    target_dict = dict(zip(df_targets['mat_prefix'], df_targets['muc_tieu']))
  except Exception:
    target_dict = {}

  # Truy vấn số liệu QC Sai Hỏng từ CSDL Mobile
  try:
    conn = get_db_connection()
    df_qc_sub = pd.read_sql_query(
        "SELECT so_lot, ma_vt, ten_vt, sl_kiem, sl_khong_dat, kieu_loi,"
        " cong_viec_con, nguoi_kiem, ngay_kiem FROM tb_qc_dau_vao WHERE loai_qc"
        " = 'SAN_XUAT' AND (sl_khong_dat > 0 OR (kieu_loi IS NOT NULL AND"
        " kieu_loi != '')) ORDER BY ngay_kiem DESC",
        conn,
    )
    conn.close()
  except Exception:
    df_qc_sub = pd.DataFrame()

  # --- CHỈ LỌC THÀNH PHẨM CHO BẢNG QC_SUB (TUTI CHỈ LẤY ĐẦU 5) ---
  if not df_qc_sub.empty:
      ma_vt_qc = df_qc_sub["ma_vt"].astype(str).str.lstrip("0")
      ten_vt_qc = df_qc_sub["ten_vt"].astype(str).str.strip().str.upper()
      
      cond_qc_5 = ma_vt_qc.str.startswith("5")
      if phan_he_code == "TU_TI":
          df_qc_sub = df_qc_sub[cond_qc_5].copy()
      else:
          cond_qc_4_special = ma_vt_qc.str.startswith("4") & (
              ten_vt_qc.str.endswith("KSS") | 
              ten_vt_qc.str.endswith("CXX") | 
              ten_vt_qc.str.endswith("CKC")
          )
          df_qc_sub = df_qc_sub[cond_qc_5 | cond_qc_4_special].copy()

  # Nhận diện an toàn cột chứa Số lệnh sản xuất (so_lenh hoặc lenh_sx)
  col_order = (
      "so_lenh"
      if "so_lenh" in df_sub.columns
      else ("lenh_sx" if "lenh_sx" in df_sub.columns else "")
  )

  # Ghép nối Lô/Lệnh COOIS với Dòng sản phẩm (mat_prefix)
  if not df_qc_sub.empty and not df_sub.empty:
    if col_order and "mat_prefix" in df_sub.columns:
      order_map = dict(
          zip(df_sub[col_order].astype(str), df_sub["mat_prefix"].astype(str))
      )
      allowed_orders = set(df_sub[col_order].astype(str).unique())
    else:
      order_map = {}
      allowed_orders = set()

    if "ma_tp" in df_sub.columns and "mat_prefix" in df_sub.columns:
      code_map = dict(
          zip(df_sub["ma_tp"].astype(str), df_sub["mat_prefix"].astype(str))
      )
      allowed_codes = set(df_sub["ma_tp"].astype(str).unique())
    else:
      code_map = {}
      allowed_codes = set()

    df_qc_sub["mat_prefix"] = df_qc_sub["so_lot"].astype(str).map(order_map)
    df_qc_sub["mat_prefix"] = (
        df_qc_sub["mat_prefix"]
        .fillna(df_qc_sub["ma_vt"].astype(str).map(code_map))
        .fillna("Khác")
    )

    df_qc_sub = df_qc_sub[
        df_qc_sub["so_lot"].astype(str).isin(allowed_orders)
        | df_qc_sub["ma_vt"].astype(str).isin(allowed_codes)
    ].copy()
  else:
    df_qc_sub = pd.DataFrame()

  months_labels = [f"T{i}" for i in range(1, 13)]
  m_comp_qty, m_uncomp_qty = [0.0] * 12, [0.0] * 12
  m_tot_orders, m_uncomp_orders = [0] * 12, [0] * 12
  tot_qty_all, deliv_qty_all = 0.0, 0.0

  for _, r in df_sub.iterrows():
    try:
      m_idx = (
          datetime.strptime(
              str(r["ngay_lenh_dt"]).split()[0], "%Y-%m-%d"
          ).month
          - 1
      )
    except Exception:
      m_idx = 0
    if not (0 <= m_idx < 12):
      m_idx = 0
    sl_t, sl_h = float(r["sl_tong"]), float(r["sl_ht"])
    uncomp_q = max(0.0, sl_t - sl_h)
    tot_qty_all += sl_t
    deliv_qty_all += sl_h
    m_comp_qty[m_idx] += sl_h
    m_uncomp_qty[m_idx] += uncomp_q
    m_tot_orders[m_idx] += 1
    if sl_h < sl_t:
      m_uncomp_orders[m_idx] += 1

  m_comp_orders = [m_tot_orders[i] - m_uncomp_orders[i] for i in range(12)]
  title_clean = clean_emoji(title_text)

  rem_qty_kpi = max(0.0, tot_qty_all - deliv_qty_all)
  pct_deliv_kpi = (
      (deliv_qty_all / tot_qty_all * 100) if tot_qty_all > 0 else 0.0
  )
  total_orders_kpi = sum(m_tot_orders)
  tot_defect_qty = (
      df_qc_sub["sl_khong_dat"].sum() if not df_qc_sub.empty else 0.0
  )
  pct_defect_kpi = (
      (tot_defect_qty / deliv_qty_all * 100.0) if deliv_qty_all > 0 else 0.0
  )

  render_kpi_cards([
      {
          "label": "Tổng Kế Hoạch",
          "value": f"{int(tot_qty_all):,}",
          "icon": "🎯",
          "color": COLOR_PRIMARY,
          "color_soft": "#EEF0FF",
          "sub": f"{int(total_orders_kpi):,} lệnh sản xuất",
      },
      {
          "label": "Đã Hoàn Thành",
          "value": f"{int(deliv_qty_all):,}",
          "icon": "✅",
          "color": COLOR_SUCCESS,
          "color_soft": "#ECFDF5",
          "sub": f"Tỷ lệ SL: {pct_deliv_kpi:.1f}%",
      },
      {
          "label": "TỔNG SỐ LƯỢNG LỖI",
          "value": f"{int(tot_defect_qty):,}",
          "icon": "⚠️",
          "color": COLOR_DANGER,
          "color_soft": "#FEF2F3",
          "sub": f"Tỷ lệ sai hỏng: {pct_defect_kpi:.2f}%",
      },
      {
          "label": "Còn Lại Chưa Xong",
          "value": f"{int(rem_qty_kpi):,}",
          "icon": "⏳",
          "color": COLOR_WARNING,
          "color_soft": "#FFF8EB",
          "sub": f"{int(sum(m_uncomp_orders)):,} lệnh chưa xong",
      },
  ])

  PLOT_HEIGHT = 380

  col1, col2 = st.columns([2.1, 1.0])
  with col1:
    chart_card_open(
        f"Sản Lượng & Tỷ Lệ Hoàn Thành — {title_clean}", "Theo tháng"
    )
    fig1 = make_subplots(specs=[[{"secondary_y": True}]])

    fig1.add_trace(
        go.Bar(
            x=months_labels,
            y=m_comp_qty,
            name="SL Hoàn Thành",
            marker_color=COLOR_SUCCESS,
            text=[f"{int(v):,}" if v > 0 else "" for v in m_comp_qty],
            textposition="inside",
            textfont=dict(size=9.5, family=PLOTLY_FONT, color="#FFFFFF"),
        ),
        secondary_y=False,
    )
    fig1.add_trace(
        go.Bar(
            x=months_labels,
            y=m_uncomp_qty,
            name="SL Chưa Xong",
            marker_color=COLOR_WARNING,
            base=m_comp_qty,
            text=[f"{int(v):,}" if v > 0 else "" for v in m_uncomp_qty],
            textposition="inside",
            textfont=dict(size=9.5, family=PLOTLY_FONT, color="#FFFFFF"),
        ),
        secondary_y=False,
    )
    pct_hoanthanh_m = [
        (m_comp_qty[i] / (m_comp_qty[i] + m_uncomp_qty[i]) * 100.0)
        if (m_comp_qty[i] + m_uncomp_qty[i]) > 0
        else None
        for i in range(12)
    ]
    fig1.add_trace(
        go.Scatter(
            x=months_labels,
            y=pct_hoanthanh_m,
            name="% Hoàn Thành",
            mode="lines+markers+text",
            line=dict(color=COLOR_PRIMARY, width=2.5),
            marker=dict(size=6, color=COLOR_PRIMARY),
            text=[f"{v:.0f}%" if v is not None else "" for v in pct_hoanthanh_m],
            textposition="top center",
            textfont=dict(size=10, family=PLOTLY_FONT, color=COLOR_PRIMARY),
            connectgaps=False,
        ),
        secondary_y=True,
    )

    fig1.update_layout(
        font=dict(family=PLOTLY_FONT, color=PLOTLY_AXIS_TEXT),
        barmode="stack",
        margin=dict(l=30, r=20, t=8, b=55),
        height=PLOT_HEIGHT,
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.22,
            xanchor="center",
            x=0.5,
            font=dict(size=11, family=PLOTLY_FONT, color="#6B7280"),
        ),
    )
    fig1.update_xaxes(
        showgrid=False, tickfont=dict(size=11, family=PLOTLY_FONT, color="#6B7280")
    )
    fig1.update_yaxes(
        title_text="← Sản Lượng",
        title_font=dict(size=12, color=COLOR_SUCCESS),
        tickformat="~s",
        tickfont=dict(size=11, family=PLOTLY_FONT, color=PLOTLY_AXIS_TEXT),
        secondary_y=False,
        showgrid=True,
        gridcolor=PLOTLY_GRID,
        zeroline=False,
    )
    fig1.update_yaxes(
        title_text="% Hoàn Thành →",
        title_font=dict(size=12, color=COLOR_PRIMARY),
        tickformat=".0f",
        ticksuffix="%",
        range=[0, 105],
        tickfont=dict(size=11, family=PLOTLY_FONT, color=PLOTLY_AXIS_TEXT),
        secondary_y=True,
        showgrid=False,
    )

    st.plotly_chart(
        fig1,
        use_container_width=True,
        config={"displayModeBar": False},
        key=f"coois_fig1_{phan_he_code}",
    )
    chart_card_close()

  with col2:
    chart_card_open("Tỷ Lệ Hoàn Thành Tổng Quan")
    pct_deliv = (deliv_qty_all / tot_qty_all * 100) if tot_qty_all > 0 else 0
    pct_rem = 100.0 - pct_deliv if tot_qty_all > 0 else 0.0

    text_labels = [
        f"<b>Hoàn thành</b><br>{pct_deliv:.1f}%<br>({int(deliv_qty_all):,})",
        f"<b>Chưa xong</b><br>{pct_rem:.1f}%<br>({int(rem_qty_kpi):,})",
    ]
    fig2 = go.Figure(
        data=[
            go.Pie(
                labels=["Hoàn thành", "Chưa xong"],
                values=[deliv_qty_all, rem_qty_kpi],
                hole=0.6,
                marker=dict(
                    colors=[COLOR_SUCCESS, COLOR_WARNING],
                    line=dict(color="#FFFFFF", width=3),
                ),
                text=text_labels,
                textinfo="text",
                textposition="outside",
                textfont=dict(
                    size=12,
                    family=PLOTLY_FONT,
                    color=[COLOR_SUCCESS, COLOR_WARNING],
                ),
                direction="clockwise",
                sort=False,
            )
        ]
    )
    fig2.update_layout(
        font=dict(family=PLOTLY_FONT, color=PLOTLY_AXIS_TEXT),
        margin=dict(l=95, r=95, t=30, b=30),
        height=PLOT_HEIGHT,
        paper_bgcolor="#FFFFFF",
        showlegend=False,
        annotations=[
            dict(
                text=(
                    "<b>TỔNG KẾ HOẠCH</b><br><span"
                    f" style='font-size:16px'>{int(tot_qty_all):,}</span>"
                ),
                x=0.5,
                y=0.5,
                font_size=12,
                font_family=PLOTLY_FONT,
                showarrow=False,
            )
        ],
    )
    st.plotly_chart(
        fig2,
        use_container_width=True,
        config={"displayModeBar": False},
        key=f"coois_fig2_{phan_he_code}",
    )
    chart_card_close()

  try:
    conn = get_db_connection()
    df_qc_rate = pd.read_sql_query(
        "SELECT so_lot, ma_vt, ten_vt, sl_kiem, sl_khong_dat, ngay_kiem FROM"
        " tb_qc_dau_vao WHERE loai_qc = 'SAN_XUAT'",
        conn,
    )
    conn.close()
  except Exception:
    df_qc_rate = pd.DataFrame()

  # --- CHỈ LỌC THÀNH PHẨM CHO BẢNG QC_RATE ---
  if not df_qc_rate.empty:
      ma_vt_rate = df_qc_rate["ma_vt"].astype(str).str.lstrip("0")
      ten_vt_rate = df_qc_rate["ten_vt"].astype(str).str.strip().str.upper()
      
      cond_rate_5 = ma_vt_rate.str.startswith("5")
      if phan_he_code == "TU_TI":
          df_qc_rate = df_qc_rate[cond_rate_5].copy()
      else:
          cond_rate_4_special = ma_vt_rate.str.startswith("4") & (
              ten_vt_rate.str.endswith("KSS") | 
              ten_vt_rate.str.endswith("CXX") | 
              ten_vt_rate.str.endswith("CKC")
          )
          df_qc_rate = df_qc_rate[cond_rate_5 | cond_rate_4_special].copy()

  if not df_qc_rate.empty and not df_sub.empty and col_order:
    order_map_rate = (
        dict(zip(df_sub[col_order].astype(str), df_sub["mat_prefix"].astype(str)))
        if "mat_prefix" in df_sub.columns
        else {}
    )
    code_map_rate = (
        dict(zip(df_sub["ma_tp"].astype(str), df_sub["mat_prefix"].astype(str)))
        if "ma_tp" in df_sub.columns and "mat_prefix" in df_sub.columns
        else {}
    )
    allowed_orders_rate = set(df_sub[col_order].astype(str).unique())
    allowed_codes_rate = (
        set(df_sub["ma_tp"].astype(str).unique())
        if "ma_tp" in df_sub.columns
        else set()
    )
    df_qc_rate = df_qc_rate[
        df_qc_rate["so_lot"].astype(str).isin(allowed_orders_rate)
        | df_qc_rate["ma_vt"].astype(str).isin(allowed_codes_rate)
    ].copy()
    df_qc_rate["mat_prefix"] = df_qc_rate["so_lot"].astype(str).map(order_map_rate)
    df_qc_rate["mat_prefix"] = (
        df_qc_rate["mat_prefix"]
        .fillna(df_qc_rate["ma_vt"].astype(str).map(code_map_rate))
        .fillna("Khác")
    )
    df_qc_rate["month"] = pd.to_datetime(
        df_qc_rate["ngay_kiem"], errors="coerce"
    ).dt.month
  else:
    df_qc_rate = pd.DataFrame()

  # --- BỘ LỌC DÒNG SP ---
  if not df_sub.empty:
      ma_tp_prefix = df_sub["ma_tp"].astype(str).str.split(".").str[0].str.lstrip("0")
      ten_tp_upper = df_sub["ten_tp"].astype(str).str.strip().str.upper()
      
      cond_5 = ma_tp_prefix.str.startswith("5")
      if phan_he_code == "TU_TI":
          sub_5 = df_sub[cond_5].copy()
      else:
          cond_4_special = ma_tp_prefix.str.startswith("4") & (
              ten_tp_upper.str.endswith("KSS") | 
              ten_tp_upper.str.endswith("CXX") | 
              ten_tp_upper.str.endswith("CKC")
          )
          sub_5 = df_sub[cond_5 | cond_4_special].copy()
  else:
      sub_5 = pd.DataFrame()

  raw_fams = (
      [
          str(x).strip()
          for x in sub_5["mat_prefix"].unique()
          if pd.notna(x)
          and str(x).strip()
          and str(x).strip().lower() not in ["none", "nan"]
      ]
      if not sub_5.empty
      else []
  )
  available_fams = ["Tất cả dòng sản phẩm"] + sorted(list(set(raw_fams)))
  sel_fam = st.selectbox(
      "🎯 Chọn Dòng SP (Đầu 5):", available_fams, key=f"cb_{phan_he_code}"
  )

  col3, col4 = st.columns([1, 1])

  with col3:
    chart_card_open(f"Sản Lượng — Dòng: {clean_emoji(sel_fam)}")
    m3_qty = [0.0] * 12
    if not sub_5.empty:
      sub_5_df = sub_5.copy()
      sub_5_df["month"] = pd.to_datetime(
          sub_5_df["ngay_lenh_dt"], errors="coerce"
      ).dt.month
      sub_filtered = (
          sub_5_df[sub_5_df["mat_prefix"] == sel_fam]
          if sel_fam != "Tất cả dòng sản phẩm"
          else sub_5_df
      )
      for _, r in sub_filtered.iterrows():
        m_val = r["month"]
        if pd.notna(m_val) and 1 <= int(m_val) <= 12:
          m3_qty[int(m_val) - 1] += float(r["sl_ht"])

    # Tỷ lệ sai hỏng theo tháng CỦA ĐÚNG DÒNG SP ĐANG CHỌN (trục phải)
    m_loi_fam = [0.0] * 12
    if not df_qc_rate.empty:
      df_qc_fam_rate = (
          df_qc_rate[df_qc_rate["mat_prefix"] == sel_fam]
          if sel_fam != "Tất cả dòng sản phẩm"
          else df_qc_rate
      )
      for _, r in df_qc_fam_rate.iterrows():
        m_val = r["month"]
        if pd.notna(m_val) and 1 <= int(m_val) <= 12:
          idx = int(m_val) - 1
          m_loi_fam[idx] += float(r["sl_khong_dat"] or 0.0)
    
    # --- TÍNH TỶ LỆ SAI HỎNG = LỖI / SẢN LƯỢNG (sl_ht) ---
    pct_loi_fam_m = [
        (m_loi_fam[i] / m3_qty[i] * 100.0) if m3_qty[i] > 0 else None
        for i in range(12)
    ]

    fig3 = make_subplots(specs=[[{"secondary_y": True}]])
    fig3.add_trace(
        go.Bar(
            x=months_labels,
            y=m3_qty,
            name="SL Sản Xuất",
            marker=dict(color=COLOR_PRIMARY, cornerradius=6),
            text=[f"{int(v):,}" if v > 0 else "" for v in m3_qty],
            textposition="outside",
            textfont=dict(color=COLOR_PRIMARY, size=11, family=PLOTLY_FONT),
        ),
        secondary_y=False,
    )
    fig3.add_trace(
        go.Scatter(
            x=months_labels,
            y=pct_loi_fam_m,
            name="Tỷ Lệ Sai Hỏng (%)",
            mode="lines+markers",
            line=dict(color=COLOR_DANGER, width=2.5),
            marker=dict(size=6, color=COLOR_DANGER),
            connectgaps=False,
        ),
        secondary_y=True,
    )
    
    # --- VẼ ĐƯỜNG MỤC TIÊU ---
    target_val = target_dict.get(sel_fam, None)
    if target_val is not None:
        fig3.add_trace(
            go.Scatter(
                x=months_labels,
                y=[target_val] * 12,
                name="Mục Tiêu Lỗi (%)",
                mode="lines",
                line=dict(color="#0EA968", width=2, dash="dash"),
                hoverinfo="y+name"
            ),
            secondary_y=True,
        )

    fig3.update_layout(
        font=dict(family=PLOTLY_FONT, color=PLOTLY_AXIS_TEXT),
        margin=dict(l=30, r=30, t=8, b=36),
        height=PLOT_HEIGHT - 40,
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        showlegend=False,
        bargap=0.3,
    )
    fig3.update_xaxes(
        showgrid=False, tickfont=dict(size=11, family=PLOTLY_FONT, color="#6B7280")
    )
    fig3.update_yaxes(
        title_text="SL Hoàn Thành",
        title_font=dict(size=12, color=COLOR_PRIMARY),
        tickformat="~s",
        tickfont=dict(size=11, family=PLOTLY_FONT, color=PLOTLY_AXIS_TEXT),
        showgrid=True,
        gridcolor=PLOTLY_GRID,
        zeroline=False,
        rangemode="tozero",
        secondary_y=False,
    )
    fig3.update_yaxes(
        title_text="% Sai Hỏng",
        title_font=dict(size=12, color=COLOR_DANGER),
        ticksuffix="%",
        tickfont=dict(size=11, family=PLOTLY_FONT, color=COLOR_DANGER),
        showgrid=False,
        secondary_y=True,
    )

    st.plotly_chart(
        fig3,
        use_container_width=True,
        config={"displayModeBar": False},
        key=f"coois_fig3_{phan_he_code}",
    )
    chart_card_close()

  with col4:
    use_bang2_breakdown = (
        phan_he_code == "CONG_TO" and sel_fam in ("CE", "ME", "AMI", "EW")
    )
    chart_card_open(
        f"Tổng Sản Lượng Cả Năm — Chi Tiết Mã Con Dòng {sel_fam}"
        if use_bang2_breakdown
        else "Tổng Sản Lượng Cả Năm Các Mã Đầu 5"
    )
    if use_bang2_breakdown:
      sub_fam_5 = (
          sub_5[sub_5["mat_prefix"] == sel_fam].copy()
          if not sub_5.empty
          else pd.DataFrame()
      )
      if not sub_fam_5.empty:
        sub_fam_5["sub_code"] = sub_fam_5["ten_tp"].apply(
            lambda t: classify_sub_code_bang2(sel_fam, t)
        )
        summary_fams = (
            sub_fam_5.groupby("sub_code")[["sl_ht"]].sum().reset_index()
        )
        fams_x = [
            str(val)
            for val in summary_fams["sub_code"].tolist()
            if pd.notna(val) and str(val).strip()
        ]
        if not fams_x:
          fams_x, deliv_fams = ["Trống"], [0.0]
        else:
          summary_fams = summary_fams[summary_fams["sub_code"].isin(fams_x)]
          fams_x, deliv_fams = (
              summary_fams["sub_code"].tolist(),
              summary_fams["sl_ht"].values,
          )
      else:
        fams_x, deliv_fams = ["Không có SP"], [0.0]
    elif not sub_5.empty:
      summary_fams = (
          sub_5.groupby("mat_prefix")[["sl_ht"]].sum().reset_index()
      )
      fams_x = [
          str(val)
          for val in summary_fams["mat_prefix"].tolist()
          if pd.notna(val)
          and str(val).strip()
          and str(val).strip().lower() not in ["none", "nan"]
      ]
      if not fams_x:
        fams_x, deliv_fams = ["Trống"], [0.0]
      else:
        summary_fams = summary_fams[summary_fams["mat_prefix"].isin(fams_x)]
        fams_x, deliv_fams = (
            summary_fams["mat_prefix"].tolist(),
            summary_fams["sl_ht"].values,
        )
    else:
      fams_x, deliv_fams = ["Không có SP"], [0.0]

    if len(fams_x) > 1:
      pairs = sorted(zip(fams_x, deliv_fams), key=lambda p: p[1], reverse=True)
      fams_x = [p[0] for p in pairs]
      deliv_fams = [p[1] for p in pairs]

    bar_colors = [
        DISTINCT_COLORS[i % len(DISTINCT_COLORS)] for i in range(len(fams_x))
    ]

    # --- TÍNH TỶ LỆ THEO TỪNG MÃ SP (SỬ DỤNG MẪU SỐ LÀ deliv_fams) ---
    pct_loi_by_fam = []
    if use_bang2_breakdown:
      code_to_ten = (
          dict(zip(sub_5["ma_tp"].astype(str), sub_5["ten_tp"].astype(str)))
          if not sub_5.empty
          else {}
      )
      if not df_qc_rate.empty:
        df_qc_rate_fam = df_qc_rate[
            df_qc_rate["mat_prefix"] == sel_fam
        ].copy()
        df_qc_rate_fam["ten_lookup"] = df_qc_rate_fam["ma_vt"].astype(
            str
        ).map(code_to_ten)
        df_qc_rate_fam["sub_code"] = df_qc_rate_fam["ten_lookup"].apply(
            lambda t: classify_sub_code_bang2(sel_fam, t) if pd.notna(t) else None
        )
        for i, fam in enumerate(fams_x):
          sub_code_rate = df_qc_rate_fam[df_qc_rate_fam["sub_code"] == fam]
          loi_sum = float(sub_code_rate["sl_khong_dat"].sum()) if not sub_code_rate.empty else 0.0
          prod_qty = float(deliv_fams[i])
          pct_loi_by_fam.append((loi_sum / prod_qty * 100.0) if prod_qty > 0 else None)
      else:
        pct_loi_by_fam = [None] * len(fams_x)
    elif not df_qc_rate.empty:
      for i, fam in enumerate(fams_x):
        sub_fam_rate = df_qc_rate[df_qc_rate["mat_prefix"] == fam]
        loi_sum = float(sub_fam_rate["sl_khong_dat"].sum()) if not sub_fam_rate.empty else 0.0
        prod_qty = float(deliv_fams[i])
        pct_loi_by_fam.append((loi_sum / prod_qty * 100.0) if prod_qty > 0 else None)
    else:
      pct_loi_by_fam = [None] * len(fams_x)

    fig4 = make_subplots(specs=[[{"secondary_y": True}]])
    fig4.add_trace(
        go.Bar(
            x=fams_x,
            y=deliv_fams,
            name="SL Hoàn Thành",
            marker=dict(color=bar_colors, cornerradius=6),
            text=[f"{int(v):,}" if v > 0 else "" for v in deliv_fams],
            textposition="outside",
            textfont=dict(color=bar_colors, size=11, family=PLOTLY_FONT),
        ),
        secondary_y=False,
    )
    fig4.add_trace(
        go.Scatter(
            x=fams_x,
            y=pct_loi_by_fam,
            name="Tỷ Lệ Sai Hỏng (%)",
            mode="lines+markers",
            line=dict(color=COLOR_DANGER, width=2.5),
            marker=dict(size=6, color=COLOR_DANGER),
            connectgaps=False,
        ),
        secondary_y=True,
    )
    
    # --- VẼ CÁC ĐIỂM/ĐƯỜNG MỤC TIÊU ---
    target_vals_by_fam = [target_dict.get(fam, None) for fam in fams_x]
    if any(v is not None for v in target_vals_by_fam):
        fig4.add_trace(
            go.Scatter(
                x=fams_x,
                y=target_vals_by_fam,
                name="Mục Tiêu Lỗi (%)",
                mode="lines+markers",
                line=dict(color="#0EA968", width=2, dash="dash"),
                marker=dict(size=8, symbol="diamond", color="#0EA968"),
                connectgaps=False
            ),
            secondary_y=True,
        )

    fig4.update_layout(
        font=dict(family=PLOTLY_FONT, color=PLOTLY_AXIS_TEXT),
        margin=dict(l=30, r=30, t=8, b=36),
        height=PLOT_HEIGHT - 40,
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        showlegend=False,
        bargap=0.3,
    )
    fig4.update_xaxes(
        showgrid=False, tickfont=dict(size=11, family=PLOTLY_FONT, color="#6B7280")
    )
    fig4.update_yaxes(
        title_text="Số Lượng SP (Log)",
        title_font=dict(size=12, color=COLOR_SUCCESS),
        type="log",
        dtick=1,
        tickformat="~s",
        tickfont=dict(size=11, family=PLOTLY_FONT, color=PLOTLY_AXIS_TEXT),
        showgrid=True,
        gridcolor=PLOTLY_GRID,
        zeroline=False,
        secondary_y=False,
    )
    fig4.update_yaxes(
        title_text="% Sai Hỏng",
        title_font=dict(size=12, color=COLOR_DANGER),
        ticksuffix="%",
        tickfont=dict(size=11, family=PLOTLY_FONT, color=COLOR_DANGER),
        showgrid=False,
        secondary_y=True,
    )

    st.plotly_chart(
        fig4,
        use_container_width=True,
        config={"displayModeBar": False},
        key=f"coois_fig4_{phan_he_code}",
    )
    chart_card_close()

# ================= 8. RENDER CÁC TAB COOIS MÀN HÌNH =================
if nav == "2. Báo Cáo Xưởng Cơ Khí":
  render_coois_tab_layout("CO_KHI", "⚙️ BÁO CÁO CƠ KHÍ (LỆNH 3012)")
if nav == "4. Báo Cáo Xưởng TU/TI":
  render_coois_tab_layout("TU_TI", "🔌 BÁO CÁO TUTI (LỆNH 3011)")
if nav == "3. Báo Cáo Xưởng Công Tơ":
  render_coois_tab_layout("CONG_TO", "⚡ BÁO CÁO CÔNG TƠ (LỆNH 3013, 3014)")
if nav == "5. Báo Cáo Xưởng TTTB CNC":
  render_coois_tab_layout("TTTB_CNC", "🛠️ BÁO CÁO TTTB CNC (LỆNH 3016)")


# ================= 9. TAB 5: DANH SÁCH CHI TIẾT, NĂNG SUẤT & QUẢN LÝ CÔNG VIỆC CON =================
if nav in DANH_SACH_LABELS:
  render_section_heading(
      "🔍 QUẢN LÝ DANH SÁCH CHI TIẾT VẬT TƯ, LỆNH SẢN XUẤT, NĂNG SUẤT & CÔNG"
      " VIỆC"
  )

  # SUB-TAB 1: QA32
  if nav == "6.1 Danh Sách Vật Tư":
    chart_card_open("⚙️ Bộ Lọc Dữ Liệu QA32")
    col_flt1, col_flt2, col_flt3, col_flt4 = st.columns([1, 1, 1, 1.5])
    with col_flt1:
      filter_status_vt = st.selectbox(
          "Trạng thái kiểm:",
          ["Tất cả", "Đã kiểm (Đã UD)", "Chưa kiểm (Chưa UD)"],
          key="ds_filter_status_vt",
      )
    with col_flt3:
      filter_loai_sp_vt = st.selectbox(
          "Loại sản phẩm:",
          ["Tất cả", "Bán thành phẩm (Đầu 5)", "Thành phẩm (Khác Đầu 5)"],
          key="ds_filter_loai_sp_vt",
      )
    with col_flt4:
      search_keyword_vt = st.text_input(
          "🔎 Tìm kiếm nhanh (Mã/Tên/NCC):", "", key="ds_search_keyword_vt"
      )
    chart_card_close()

    if not df_qa32.empty:
      df_qa32_view = df_qa32.copy()
      if "ngay_ve_dt" in df_qa32_view.columns:
        df_qa32_view["ngay_ve_format"] = pd.to_datetime(
            df_qa32_view["ngay_ve_dt"], errors="coerce"
        ).dt.strftime("%d/%m/%Y")
      else:
        df_qa32_view["ngay_ve_format"] = "-"

      if "xac_nhan_sap" in df_qa32_view.columns:
        if filter_status_vt == "Đã kiểm (Đã UD)":
          df_qa32_view = df_qa32_view[
              df_qa32_view["xac_nhan_sap"].notna()
              & (~df_qa32_view["xac_nhan_sap"]
                  .astype(str)
                  .str.contains("CHƯA|NAN|NONE", case=False, na=False))
          ]
        elif filter_status_vt == "Chưa kiểm (Chưa UD)":
          df_qa32_view = df_qa32_view[
              df_qa32_view["xac_nhan_sap"].isna()
              | df_qa32_view["xac_nhan_sap"]
              .astype(str)
              .str.contains("CHƯA|NAN|NONE", case=False, na=False)
          ]

      if "ma_vt" in df_qa32_view.columns:
        if filter_loai_sp_vt == "Bán thành phẩm (Đầu 5)":
          df_qa32_view = df_qa32_view[
              df_qa32_view["ma_vt"]
              .astype(str)
              .str.lstrip("0")
              .str.startswith("5")
          ]
        elif filter_loai_sp_vt == "Thành phẩm (Khác Đầu 5)":
          df_qa32_view = df_qa32_view[
              ~df_qa32_view["ma_vt"]
              .astype(str)
              .str.lstrip("0")
              .str.startswith("5")
          ]

      if search_keyword_vt.strip():
        kw = search_keyword_vt.strip().lower()
        m1 = (
            df_qa32_view["ma_vt"]
            .astype(str)
            .str.lower()
            .str.contains(kw, na=False)
            if "ma_vt" in df_qa32_view.columns
            else False
        )
        m2 = (
            df_qa32_view["ten_vt"]
            .astype(str)
            .str.lower()
            .str.contains(kw, na=False)
            if "ten_vt" in df_qa32_view.columns
            else False
        )
        m3 = (
            df_qa32_view["ncc"]
            .astype(str)
            .str.lower()
            .str.contains(kw, na=False)
            if "ncc" in df_qa32_view.columns
            else False
        )
        df_qa32_view = df_qa32_view[m1 | m2 | m3]

      if not df_qa32_view.empty:
        df_vt_display = pd.DataFrame()
        df_vt_display["STT"] = np.arange(1, len(df_qa32_view) + 1)
        df_vt_display["Ngày tháng năm"] = df_qa32_view["ngay_ve_format"].values
        df_vt_display["Lot"] = (
            df_qa32_view["lot"].values
            if "lot" in df_qa32_view.columns
            else (
                df_qa32_view["so_lot"].values
                if "so_lot" in df_qa32_view.columns
                else df_qa32_view.index + 1
            )
        )
        df_vt_display["Mã vật tư"] = (
            df_qa32_view["ma_vt"].values
            if "ma_vt" in df_qa32_view.columns
            else ""
        )
        df_vt_display["Tên vật tư"] = (
            df_qa32_view["ten_vt"].values
            if "ten_vt" in df_qa32_view.columns
            else ""
        )
        df_vt_display["Nhà cung cấp"] = (
            df_qa32_view["ncc"].values if "ncc" in df_qa32_view.columns else ""
        )

        raw_ud = (
            df_qa32_view["xac_nhan_sap"].fillna("Chưa kiểm (Chưa UD)").values
            if "xac_nhan_sap" in df_qa32_view.columns
            else ["Chưa kiểm (Chưa UD)"] * len(df_qa32_view)
        )
        clean_ud = [
            "Chưa kiểm (Chưa UD)"
            if (
                "CHƯA" in str(u).upper()
                or str(u).strip() in ["nan", "None", ""]
            )
            else str(u)
            for u in raw_ud
        ]
        df_vt_display["Giá trị kiểm"] = clean_ud

        st.markdown(f"##### 📋 Danh Sách Vật Tư ({len(df_vt_display):,} bản ghi)")
        st.dataframe(
            df_vt_display,
            column_config={
                "STT": st.column_config.NumberColumn("STT", width="small"),
                "Giá trị kiểm": st.column_config.TextColumn(
                    "Giá trị kiểm", width="medium"
                ),
            },
            use_container_width=True,
            hide_index=True,
            height=480,
        )

        buf_vt = io.BytesIO()
        with pd.ExcelWriter(buf_vt, engine="openpyxl") as writer:
          df_vt_display.to_excel(
              writer, sheet_name="DanhSach_VatTu_QA32", index=False
          )
        st.download_button(
            label="📥 Xuất Bảng Vật Tư (Excel)",
            data=buf_vt.getvalue(),
            file_name=(
                "DanhSach_VatTu_QA32_"
                f"{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
      else:
        st.warning("⚠️ Không tìm thấy bản ghi vật tư nào phù hợp với bộ lọc.")
    else:
      st.info("💡 Chưa có dữ liệu vật tư.")

  # SUB-TAB 2: COOIS
  if nav == "6.2 Danh Sách Lệnh Sản Xuất":
    chart_card_open("⚙️ Bộ Lọc Dữ Luợu COOIS")
    col_f1, col_f2, col_f3, col_f4 = st.columns([1, 1, 1, 1.5])
    with col_f1:
      filter_status_l = st.selectbox(
          "Trạng thái:",
          ["Tất cả", "Đã xong", "Chưa xong"],
          key="ds_filter_status_l",
      )
    with col_f2:
      filter_phan_he_l = st.selectbox(
          "Xưởng / Phân hệ:",
          [
              "Tất cả",
              "Cơ khí (CO_KHI)",
              "TU/TI (TU_TI)",
              "Công tơ (CONG_TO)",
              "TTTB CNC (TTTB_CNC)",
          ],
          key="ds_filter_phan_he_l",
      )
    with col_f3:
      filter_loai_sp_l = st.selectbox(
          "Loại sản phẩm:",
          ["Tất cả", "Bán thành phẩm (Đầu 4)", "Sản phẩm (Đầu 5)"],
          key="ds_filter_loai_sp_l",
      )
    with col_f4:
      search_keyword_l = st.text_input(
          "🔎 Tìm kiếm nhanh (Số lệnh/Mã/Tên):", "", key="ds_search_keyword_l"
      )
    chart_card_close()

    phan_he_code_map = {
        "Cơ khí (CO_KHI)": "CO_KHI",
        "TU/TI (TU_TI)": "TU_TI",
        "Công tơ (CONG_TO)": "CONG_TO",
        "TTTB CNC (TTTB_CNC)": "TTTB_CNC",
    }

    if not df_coois.empty:
      df_coois_view = df_coois.copy()
      if "ngay_lenh_dt" in df_coois_view.columns:
        df_coois_view["ngay_lenh_format"] = pd.to_datetime(
            df_coois_view["ngay_lenh_dt"], errors="coerce"
        ).dt.strftime("%d/%m/%Y")
      else:
        df_coois_view["ngay_lenh_format"] = "-"

      if filter_phan_he_l != "Tất cả" and "phan_he" in df_coois_view.columns:
        target_ph = phan_he_code_map.get(filter_phan_he_l)
        df_coois_view = df_coois_view[df_coois_view["phan_he"] == target_ph]

      if (
          "sl_tong" in df_coois_view.columns
          and "sl_ht" in df_coois_view.columns
      ):
        if filter_status_l == "Đã xong":
          df_coois_view = df_coois_view[
              df_coois_view["sl_ht"] >= df_coois_view["sl_tong"]
          ]
        elif filter_status_l == "Chưa xong":
          df_coois_view = df_coois_view[
              df_coois_view["sl_ht"] < df_coois_view["sl_tong"]
          ]

      if "ma_tp" in df_coois_view.columns:
        if filter_loai_sp_l == "Bán thành phẩm (Đầu 4)":
          df_coois_view = df_coois_view[
              df_coois_view["ma_tp"]
              .astype(str)
              .str.split(".")
              .str[0]
              .str.lstrip("0")
              .str.startswith("4")
          ]
        elif filter_loai_sp_l == "Sản phẩm (Đầu 5)":
          df_coois_view = df_coois_view[
              df_coois_view["ma_tp"]
              .astype(str)
              .str.split(".")
              .str[0]
              .str.lstrip("0")
              .str.startswith("5")
          ]

      if search_keyword_l.strip():
        kw = search_keyword_l.strip().lower()
        col_order_check = (
            "so_lenh"
            if "so_lenh" in df_coois_view.columns
            else ("lenh_sx" if "lenh_sx" in df_coois_view.columns else "")
        )

        m1 = (
            df_coois_view[col_order_check]
            .astype(str)
            .str.lower()
            .str.contains(kw, na=False)
            if col_order_check
            else False
        )
        m2 = (
            df_coois_view["ma_tp"]
            .astype(str)
            .str.lower()
            .str.contains(kw, na=False)
            if "ma_tp" in df_coois_view.columns
            else False
        )
        m3 = (
            df_coois_view["ten_tp"]
            .astype(str)
            .str.lower()
            .str.contains(kw, na=False)
            if "ten_tp" in df_coois_view.columns
            else False
        )
        df_coois_view = df_coois_view[m1 | m2 | m3]

      if not df_coois_view.empty:
        col_order_disp = (
            "so_lenh"
            if "so_lenh" in df_coois_view.columns
            else ("lenh_sx" if "lenh_sx" in df_coois_view.columns else "")
        )
        df_lenh_display = pd.DataFrame()
        df_lenh_display["STT"] = np.arange(1, len(df_coois_view) + 1)
        df_lenh_display["Ngày tháng năm"] = df_coois_view[
            "ngay_lenh_format"
        ].values
        df_lenh_display["Lệnh"] = (
            df_coois_view[col_order_disp].values if col_order_disp else ""
        )
        df_lenh_display["Mã sản phẩm"] = (
            df_coois_view["ma_tp"].values
            if "ma_tp" in df_coois_view.columns
            else ""
        )
        df_lenh_display["Tên sản phẩm"] = (
            df_coois_view["ten_tp"].values
            if "ten_tp" in df_coois_view.columns
            else ""
        )
        df_lenh_display["Tổng số lượng"] = (
            df_coois_view["sl_tong"].values
            if "sl_tong" in df_coois_view.columns
            else 0
        )
        df_lenh_display["Tổng số đã giao"] = (
            df_coois_view["sl_ht"].values
            if "sl_ht" in df_coois_view.columns
            else 0
        )
        df_lenh_display["Text ghi chú"] = (
            df_coois_view["ghi_chu"].values
            if "ghi_chu" in df_coois_view.columns
            else (
                df_coois_view["phan_he"].values
                if "phan_he" in df_coois_view.columns
                else ""
            )
        )

        st.markdown(
            "##### ⚙️ Danh Sách Lệnh Kiểm Tra / Sản Xuất"
            f" ({len(df_lenh_display):,} bản ghi)"
        )
        st.dataframe(
            df_lenh_display,
            column_config={
                "STT": st.column_config.NumberColumn("STT", width="small"),
                "Tổng số lượng": st.column_config.NumberColumn(
                    "Tổng số lượng", format="%d"
                ),
                "Tổng số đã giao": st.column_config.NumberColumn(
                    "Tổng số đã giao", format="%d"
                ),
            },
            use_container_width=True,
            hide_index=True,
            height=480,
        )

        buf_lenh = io.BytesIO()
        with pd.ExcelWriter(buf_lenh, engine="openpyxl") as writer:
          df_lenh_display.to_excel(
              writer, sheet_name="DanhSach_Lenh_COOIS", index=False
          )
        st.download_button(
            label="📥 Xuất Bảng Lệnh Kiểm Tra (Excel)",
            data=buf_lenh.getvalue(),
            file_name=(
                "DanhSach_Lenh_COOIS_"
                f"{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
      else:
        st.warning("⚠️ Không tìm thấy bản ghi lệnh nào phù hợp với bộ lọc.")
    else:
      st.info("💡 Chưa có dữ liệu lệnh sản xuất.")

  # SUB-TAB 3: BÁO CÁO NĂNG SUẤT CÁ NHÂN VÀ TÍNH NĂNG QUẢN LÝ CÔNG VIỆC CON
  if nav == "6.4 Năng Suất":
    try:
      conn = get_db_connection()
      df_qc_logs = pd.read_sql_query(
          "SELECT id, loai_qc, so_lot, ma_vt, ten_vt, ncc, sl_kiem,"
          " sl_khong_dat, sl_dat, ket_luan, nguoi_kiem, ngay_kiem, ghi_chu,"
          " kieu_loi, cong_viec_con, sl_huy FROM tb_qc_dau_vao ORDER BY ngay_kiem"
          " DESC",
          conn,
      )
      conn.close()

      if not df_qc_logs.empty:
        df_qc_logs["ngay_kiem_dt"] = pd.to_datetime(
            df_qc_logs["ngay_kiem"], errors="coerce"
        )
        df_qc_logs["Ngay_Format"] = df_qc_logs["ngay_kiem_dt"].dt.strftime(
            "%d/%m/%Y"
        )
        df_qc_logs["Ngay_Nhap_Format"] = df_qc_logs["ngay_kiem_dt"].dt.strftime(
            "%d/%m/%Y %H:%M:%S"
        )

        chart_card_open("👨‍💼 Bộ Lọc Tính Năng Suất Làm Việc QC")
        col_flt_person, col_flt_type, col_flt_tu, col_flt_den = st.columns(
            [1.3, 1, 1, 1]
        )

        list_inspectors = ["Tất cả nhân sự"] + sorted([
            str(x).strip()
            for x in df_qc_logs["nguoi_kiem"].dropna().unique()
            if str(x).strip()
        ])

        with col_flt_person:
          selected_inspector = st.selectbox(
              "👤 Chọn Nhân sự QC:", list_inspectors, key="ns_inspector"
          )

        with col_flt_type:
          selected_loai_qc = st.selectbox(
              "🎯 Phân hệ kiểm:",
              ["Tất cả", "QC Đầu Vào (DAU_VAO)", "QC Sản Xuất (SAN_XUAT)", "QC Phát Sinh (PHAT_SINH)"],
              key="ns_loai_qc",
          )

        ns_min_date = df_qc_logs["ngay_kiem_dt"].min()
        ns_max_date = df_qc_logs["ngay_kiem_dt"].max()
        ns_default_tu = (
            ns_min_date.date() if pd.notna(ns_min_date) else date.today()
        )
        ns_default_den = (
            ns_max_date.date() if pd.notna(ns_max_date) else date.today()
        )
        with col_flt_tu:
          ns_tu_ngay = st.date_input(
              "📅 Từ ngày:",
              ns_default_tu,
              key="ns_tu_ngay",
              format="DD/MM/YYYY",
          )
        with col_flt_den:
          ns_den_ngay = st.date_input(
              "📅 Đến ngày:",
              ns_default_den,
              key="ns_den_ngay",
              format="DD/MM/YYYY",
          )
        chart_card_close()

        df_filtered = df_qc_logs.copy()

        if selected_inspector != "Tất cả nhân sự":
          df_filtered = df_filtered[
              df_filtered["nguoi_kiem"] == selected_inspector
          ]

        if selected_loai_qc == "QC Đầu Vào (DAU_VAO)":
          df_filtered = df_filtered[df_filtered["loai_qc"] == "DAU_VAO"]
        elif selected_loai_qc == "QC Sản Xuất (SAN_XUAT)":
          df_filtered = df_filtered[df_filtered["loai_qc"] == "SAN_XUAT"]
        elif selected_loai_qc == "QC Phát Sinh (PHAT_SINH)":
          df_filtered = df_filtered[df_filtered["loai_qc"] == "PHAT_SINH"]

        df_filtered = df_filtered[
            (df_filtered["ngay_kiem_dt"].dt.date >= ns_tu_ngay)
            & (df_filtered["ngay_kiem_dt"].dt.date <= ns_den_ngay)
        ]

        tot_luot = len(df_filtered)
        tot_sl_kiem = (
            df_filtered["sl_kiem"].sum() if "sl_kiem" in df_filtered else 0
        )
        tot_sl_loi = (
            df_filtered["sl_khong_dat"].sum()
            if "sl_khong_dat" in df_filtered
            else 0
        )
        ty_le_loi = (
            (tot_sl_loi / tot_sl_kiem * 100) if tot_sl_kiem > 0 else 0.0
        )

        render_kpi_cards([
            {
                "label": "TỔNG LƯỢT KIỂM TRẢ",
                "value": f"{tot_luot:,} lượt",
                "icon": "📝",
                "color": COLOR_PRIMARY,
            },
            {
                "label": "TỔNG SỐ LƯỢNG ĐÃ KIỂM",
                "value": f"{int(tot_sl_kiem):,}",
                "icon": "🔍",
                "color": COLOR_SUCCESS,
            },
            {
                "label": "TỔNG SỐ LƯỢNG LỖI",
                "value": f"{int(tot_sl_loi):,}",
                "icon": "⚠️",
                "color": COLOR_DANGER,
            },
            {
                "label": "TỶ LỆ LỖI PHÁT HIỆN",
                "value": f"{ty_le_loi:.1f}%",
                "icon": "📈",
                "color": COLOR_WARNING,
            },
        ])

        df_display = pd.DataFrame()
        df_display["ID"] = df_filtered["id"].values
        df_display["STT"] = np.arange(1, len(df_filtered) + 1)
        df_display["Ngày kiểm"] = df_filtered["ngay_kiem_dt"].dt.date.values
        df_display["Ngày nhập"] = df_filtered["Ngay_Nhap_Format"].values
        df_display["Người kiểm tra"] = df_filtered["nguoi_kiem"].values
        df_display["Loại QC"] = df_filtered["loai_qc"].map(
            {"DAU_VAO": "QC Đầu Vào", "SAN_XUAT": "QC Sản Xuất", "PHAT_SINH": "QC Phát Sinh"}
        )
        df_display["Lô / Lệnh SX"] = df_filtered["so_lot"].values
        df_display["Mã mặt hàng"] = df_filtered["ma_vt"].values
        df_display["Tên mặt hàng"] = df_filtered["ten_vt"].values
        df_display["Đơn vị / NCC"] = df_filtered["ncc"].values
        df_display["Công việc con"] = df_filtered["cong_viec_con"].values
        df_display["SL Kiểm"] = df_filtered["sl_kiem"].values
        df_display["SL Lỗi"] = df_filtered["sl_khong_dat"].values
        df_display["SL Hủy"] = df_filtered["sl_huy"].values
        df_display["SL Đạt"] = df_filtered["sl_dat"].values
        df_display["Kết luận"] = df_filtered["ket_luan"].values
        df_display["Ghi chú"] = df_filtered["ghi_chu"].values

        st.markdown(
            f"##### 📋 BẢNG NHẬT KÝ KIỂM TRA CHI TIẾT ({len(df_display):,} bản"
            " ghi) — sửa bất kỳ ô nào hoặc xoá cả dòng rồi bấm Lưu"
        )
        edited_ns = st.data_editor(
            df_display,
            column_config={
                "ID": None,
                "STT": st.column_config.NumberColumn(
                    "STT", width="small", disabled=True
                ),
                "Ngày kiểm": st.column_config.DateColumn(
                    "Ngày kiểm", format="DD/MM/YYYY"
                ),
                "Ngày nhập": st.column_config.TextColumn(
                    "Ngày nhập", disabled=True
                ),
                "Loại QC": st.column_config.SelectboxColumn(
                    "Loại QC", options=["QC Đầu Vào", "QC Sản Xuất", "QC Phát Sinh"]
                ),
                "SL Kiểm": st.column_config.NumberColumn(
                    "SL Kiểm", format="%d", min_value=0
                ),
                "SL Lỗi": st.column_config.NumberColumn(
                    "SL Lỗi", format="%d", min_value=0
                ),
                "SL Hủy": st.column_config.NumberColumn(
                    "SL Hủy", format="%d", min_value=0
                ),
                "SL Đạt": st.column_config.NumberColumn(
                    "SL Đạt", format="%d", disabled=True
                ),
                "Kết luận": st.column_config.SelectboxColumn(
                    "Kết luận",
                    options=["Đạt", "Không đạt", "Chấp nhận", "Đạt tiêu chuẩn (UD 01)", "Không đạt - Trả lại (UD 03)"],
                ),
            },
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            height=450,
            key="editor_ns_logs",
        )

        if st.button(
            "💾 Lưu Thay Đổi Nhật Ký Năng Suất",
            type="primary",
            key="save_ns_logs",
        ):
          try:
            conn = get_db_connection()
            cursor = conn.cursor()
            original_ids_ns = set(df_display["ID"].dropna().astype(int))
            edited_ids_ns = set(edited_ns["ID"].dropna().astype(int))
            n_deleted_ns = 0
            for did in original_ids_ns - edited_ids_ns:
              cursor.execute(
                  "DELETE FROM tb_qc_dau_vao WHERE id = ?", (int(did),)
              )
              n_deleted_ns += 1
            loai_qc_rev = {
                "QC Đầu Vào": "DAU_VAO",
                "QC Sản Xuất": "SAN_XUAT",
                "QC Phát Sinh": "PHAT_SINH",
            }
            n_saved_ns = 0
            for _, row in edited_ns.iterrows():
              if pd.isna(row["ID"]):
                continue
              rid = int(row["ID"])
              sl_kiem_v = (
                  float(row["SL Kiểm"]) if pd.notna(row["SL Kiểm"]) else 0.0
              )
              sl_loi_v = (
                  float(row["SL Lỗi"]) if pd.notna(row["SL Lỗi"]) else 0.0
              )
              sl_huy_v = (
                  float(row["SL Hủy"]) if "SL Hủy" in row and pd.notna(row["SL Hủy"]) else 0.0
              )
              sl_dat_v = max(0.0, sl_kiem_v - sl_loi_v)
              
              ngay_kiem_v = row["Ngày kiểm"]
              if pd.isna(ngay_kiem_v):
                ngay_kiem_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
              else:
                # Giữ nguyên Giờ:Phút:Giây từ cột 'Ngày nhập' cũ để không bị ghi đè
                old_time_str = str(row.get("Ngày nhập", ""))[-8:] 
                try:
                    time_part = datetime.strptime(old_time_str, "%H:%M:%S").time()
                except:
                    time_part = datetime.now().time() # Dự phòng nếu lỗi
                
                ngay_kiem_str = datetime.combine(
                    ngay_kiem_v
                    if not isinstance(ngay_kiem_v, str)
                    else pd.to_datetime(ngay_kiem_v).date(),
                    time_part,
                ).strftime("%Y-%m-%d %H:%M:%S")
              
              cursor.execute(
                  "UPDATE tb_qc_dau_vao SET nguoi_kiem = ?, loai_qc = ?,"
                  " so_lot = ?, ma_vt = ?, ten_vt = ?, ncc = ?,"
                  " cong_viec_con = ?, sl_kiem = ?, sl_khong_dat = ?,"
                  " sl_dat = ?, ket_luan = ?, ghi_chu = ?, ngay_kiem = ?, sl_huy = ?"
                  " WHERE id = ?",
                  (
                      str(row["Người kiểm tra"] or "").strip(),
                      loai_qc_rev.get(row["Loại QC"], "SAN_XUAT"),
                      str(row["Lô / Lệnh SX"] or "").strip(),
                      str(row["Mã mặt hàng"] or "").strip(),
                      str(row["Tên mặt hàng"] or "").strip(),
                      str(row["Đơn vị / NCC"] or "").strip(),
                      str(row["Công việc con"] or "").strip(),
                      sl_kiem_v,
                      sl_loi_v,
                      sl_dat_v,
                      str(row["Kết luận"] or "").strip(),
                      str(row["Ghi chú"] or "").strip(),
                      ngay_kiem_str,
                      sl_huy_v,
                      rid,
                  ),
              )
              n_saved_ns += 1
            conn.commit()
            conn.close()
            st.success(
                f"✅ Đã lưu {n_saved_ns} bản ghi, xoá {n_deleted_ns} bản ghi!"
            )
            st.cache_data.clear()
            st.rerun()
          except Exception as ex:
            st.error(f"Lỗi lưu thay đổi: {ex}")

        buf_ns = io.BytesIO()
        with pd.ExcelWriter(buf_ns, engine="openpyxl") as writer:
          df_display.to_excel(
              writer, sheet_name="NangSuat_QC_ChiTiet", index=False
          )

        st.download_button(
            label="📥 XUẤT BÁO CÁO NĂNG SUẤT QC (EXCEL)",
            data=buf_ns.getvalue(),
            file_name=(
                "BaoCao_NangSuat_QC_"
                f"{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
      else:
        st.info("💡 Chưa có nhật ký báo cáo QC nào trong Cơ sở dữ liệu.")
    except Exception as e:
      st.error(f"⚠️ Lỗi khi tải nhật ký QC: {e}")

  # SUB-TAB 4: BÁO CÁO SAI HỎNG CHI TIẾT & BẢNG QUẢN LÝ DANH MỤC LỖI
  if nav == "6.3 Sai Hỏng":
    st.markdown("#### 🚨 BÁO CÁO PHÂN TÍCH SAI HỎNG")

    try:
      conn = get_db_connection()
      df_defects = pd.read_sql_query(
          "SELECT * FROM tb_qc_dau_vao WHERE sl_khong_dat > 0 OR (kieu_loi"
          " IS NOT NULL AND kieu_loi != '') ORDER BY ngay_kiem DESC",
          conn,
      )
      conn.close()

      if not df_defects.empty:
        # Đảm bảo cột sl_huy có sẵn (nếu CSDL cũ)
        if "sl_huy" not in df_defects.columns:
            df_defects["sl_huy"] = 0.0

        # Ghép "Dòng sản phẩm" (mat_prefix) từ COOIS để lọc đúng dòng SP
        order_map_sh, code_map_sh = {}, {}
        if not df_coois.empty and "mat_prefix" in df_coois.columns:
          if "lenh_sx" in df_coois.columns:
            order_map_sh = dict(
                zip(
                    df_coois["lenh_sx"].astype(str),
                    df_coois["mat_prefix"].astype(str),
                )
            )
          if "ma_tp" in df_coois.columns:
            code_map_sh = dict(
                zip(
                    df_coois["ma_tp"].astype(str),
                    df_coois["mat_prefix"].astype(str),
                )
            )
        df_defects = df_defects.copy()
        df_defects["mat_prefix"] = df_defects["so_lot"].astype(str).map(
            order_map_sh
        )
        df_defects["mat_prefix"] = (
            df_defects["mat_prefix"]
            .fillna(df_defects["ma_vt"].astype(str).map(code_map_sh))
            .fillna("Khác")
        )

        col_sh_f1, col_sh_f2 = st.columns(2)
        with col_sh_f1:
          sh_filter_loai = st.selectbox(
              "Lọc Phân Hệ:",
              ["Tất cả", "DAU_VAO", "CO_KHI", "TU_TI", "CONG_TO", "TTTB_CNC", "PHAT_SINH"],
              key="sh_flt_ph",
          )
        with col_sh_f2:
          fam_options_sh = ["Tất cả"] + sorted(
              df_defects["mat_prefix"].dropna().unique().tolist()
          )
          sh_filter_fam = st.selectbox(
              "Lọc Dòng Sản Phẩm:", fam_options_sh, key="sh_flt_fam"
          )

        df_sh_view = df_defects.copy()
        if sh_filter_loai != "Tất cả":
          df_sh_view = df_sh_view[
              df_sh_view["loai_qc"].str.contains(sh_filter_loai, na=False)
              | df_sh_view["ncc"].str.contains(sh_filter_loai, na=False)
          ]
        if sh_filter_fam != "Tất cả":
          df_sh_view = df_sh_view[df_sh_view["mat_prefix"] == sh_filter_fam]

        if not df_sh_view.empty and "kieu_loi" in df_sh_view.columns:
          defect_counts = (
              df_sh_view.groupby("kieu_loi")["sl_khong_dat"]
              .sum()
              .reset_index()
          )
          defect_counts = defect_counts[defect_counts["kieu_loi"] != ""]
          defect_counts = defect_counts.sort_values(
              by="sl_khong_dat", ascending=False
          ).reset_index(drop=True)

          if not defect_counts.empty:
            defect_counts["cum_pct"] = (
                defect_counts["sl_khong_dat"].cumsum()
                / defect_counts["sl_khong_dat"].sum()
                * 100.0
            )
            bar_colors_pareto = [
                DISTINCT_COLORS[i % len(DISTINCT_COLORS)]
                for i in range(len(defect_counts))
            ]

            fig_err = make_subplots(specs=[[{"secondary_y": True}]])
            fig_err.add_trace(
                go.Bar(
                    x=defect_counts["kieu_loi"],
                    y=defect_counts["sl_khong_dat"],
                    marker=dict(color=bar_colors_pareto),
                    width=0.35,
                    text=[
                        f"{v:,.0f}" for v in defect_counts["sl_khong_dat"]
                    ],
                    textposition="outside",
                    textfont=dict(size=11, family=PLOTLY_FONT),
                    name="Số Lượng Lỗi",
                ),
                secondary_y=False,
            )
            fig_err.add_trace(
                go.Scatter(
                    x=defect_counts["kieu_loi"],
                    y=defect_counts["cum_pct"],
                    mode="lines+markers+text",
                    line=dict(color=COLOR_WARNING, width=2.5),
                    marker=dict(size=6, color=COLOR_WARNING),
                    text=[f"{v:.0f}%" for v in defect_counts["cum_pct"]],
                    textposition="top center",
                    textfont=dict(
                        size=10, family=PLOTLY_FONT, color=COLOR_WARNING
                    ),
                    name="% Lũy Kế",
                ),
                secondary_y=True,
            )
            fig_err.update_layout(
                font=dict(family=PLOTLY_FONT, color=PLOTLY_AXIS_TEXT),
                title="<b>PARETO — CÁC KIỂU SAI HỎNG PHÁT HIỆN NHIỀU NHẤT</b>",
                margin=dict(l=10, r=40, t=40, b=90),
                height=380,
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#FFFFFF",
                showlegend=False,
            )
            fig_err.update_xaxes(
                tickangle=-30, tickfont=dict(size=10, family=PLOTLY_FONT)
            )
            fig_err.update_yaxes(
                title_text="Số Lượng Lỗi",
                tickfont=dict(size=11, family=PLOTLY_FONT),
                showgrid=True,
                gridcolor=PLOTLY_GRID,
                secondary_y=False,
            )
            fig_err.update_yaxes(
                title_text="% Lũy Kế",
                range=[0, 105],
                ticksuffix="%",
                tickfont=dict(size=11, family=PLOTLY_FONT, color=COLOR_WARNING),
                showgrid=False,
                secondary_y=True,
            )
            st.plotly_chart(
                fig_err,
                use_container_width=True,
                config={"displayModeBar": False},
                key="sh_chart_fig_err",
            )

        st.markdown(
            "##### 📋 Danh Sách Ca Báo Lỗi Chi Tiết — sửa Công Việc Con /"
            " Kiểu Sai Hỏng / SL Lỗi / SL Hủy / Ghi Chú, hoặc xoá cả dòng, rồi bấm Lưu"
        )
        df_sh_view = df_sh_view.copy()
        df_sh_view["ngay_kiem_fmt"] = pd.to_datetime(
            df_sh_view["ngay_kiem"], errors="coerce"
        ).dt.strftime("%d/%m/%Y %H:%M:%S")
        df_sh_edit = df_sh_view[[
            "id",
            "ngay_kiem_fmt",
            "nguoi_kiem",
            "so_lot",
            "ma_vt",
            "ten_vt",
            "ncc",
            "cong_viec_con",
            "kieu_loi",
            "sl_kiem",
            "sl_khong_dat",
            "sl_huy",
            "ghi_chu",
        ]].copy()
        df_sh_edit.columns = [
            "ID",
            "Ngày nhập",
            "Người Kiểm",
            "Lô/Lệnh",
            "Mã Hàng",
            "Tên Mặt Hàng",
            "Xưởng/NCC",
            "Công Việc Con",
            "Kiểu Sai Hỏng",
            "SL Kiểm",
            "SL Lỗi",
            "SL Hủy",
            "Ghi Chú Chi Tiết",
        ]

        edited_sh = st.data_editor(
            df_sh_edit,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            disabled=[
                "ID",
                "Ngày nhập",
                "Người Kiểm",
                "Lô/Lệnh",
                "Mã Hàng",
                "Tên Mặt Hàng",
                "Xưởng/NCC",
                "SL Kiểm",
            ],
            column_config={
                "SL Lỗi": st.column_config.NumberColumn("SL Lỗi", min_value=0),
                "SL Hủy": st.column_config.NumberColumn("SL Hủy", min_value=0),
            },
            key="editor_sh_defects",
        )

        if st.button(
            "💾 Lưu Thay Đổi Số Liệu Sai Hỏng",
            type="primary",
            key="save_sh_defects",
        ):
          try:
            conn = get_db_connection()
            cursor = conn.cursor()
            n_saved = 0
            original_ids_sh = set(df_sh_edit["ID"].dropna().astype(int))
            edited_ids_sh = set(edited_sh["ID"].dropna().astype(int))
            n_deleted = 0
            for did in original_ids_sh - edited_ids_sh:
              cursor.execute(
                  "DELETE FROM tb_qc_dau_vao WHERE id = ?", (int(did),)
              )
              n_deleted += 1
            for _, row in edited_sh.iterrows():
              if pd.isna(row["ID"]):
                continue
              rid = int(row["ID"])
              cvc = str(row["Công Việc Con"] or "").strip()
              kl = str(row["Kiểu Sai Hỏng"] or "").strip()
              sl_loi = (
                  float(row["SL Lỗi"]) if pd.notna(row["SL Lỗi"]) else 0.0
              )
              sl_huy = (
                  float(row["SL Hủy"]) if "SL Hủy" in row and pd.notna(row["SL Hủy"]) else 0.0
              )
              sl_kiem_val = (
                  float(row["SL Kiểm"]) if pd.notna(row["SL Kiểm"]) else 0.0
              )
              sl_dat_val = max(0.0, sl_kiem_val - sl_loi)
              ghi_chu_val = str(row["Ghi Chú Chi Tiết"] or "").strip()
              ket_luan_val = (
                  "Đạt tiêu chuẩn (UD 01)"
                  if sl_loi == 0
                  else "Không đạt - Trả lại (UD 03)"
              )
              cursor.execute(
                  "UPDATE tb_qc_dau_vao SET cong_viec_con = ?, kieu_loi = ?,"
                  " sl_khong_dat = ?, sl_dat = ?, ket_luan = ?, ghi_chu = ?, sl_huy = ?"
                  " WHERE id = ?",
                  (cvc, kl, sl_loi, sl_dat_val, ket_luan_val, ghi_chu_val, sl_huy, rid),
              )
              n_saved += 1
            conn.commit()
            conn.close()
            st.success(
                f"✅ Đã lưu {n_saved} bản ghi sửa đổi, xoá {n_deleted} bản"
                " ghi!"
            )
            st.cache_data.clear()
            st.rerun()
          except Exception as ex:
            st.error(f"Lỗi lưu thay đổi: {ex}")

        st.markdown(
            "<hr style='margin: 16px 0; border-color: #E4E8F0;'>",
            unsafe_allow_html=True,
        )
        st.markdown("##### 🖼️ Xem Ảnh Chụp Kiểm Tra Đã Lưu")
        img_options = {
            f"{r['so_lot']} - {r['ma_vt']} - {r['nguoi_kiem']} ({r['ngay_kiem_fmt']})": r[
                "id"
            ]
            for _, r in df_sh_view.iterrows()
        }
        sel_img_label = st.selectbox(
            "Chọn bản ghi cần xem ảnh:",
            list(img_options.keys()),
            key="sh_img_record_select",
        )
        sel_img_id = img_options[sel_img_label]
        row_img = df_sh_view[df_sh_view["id"] == sel_img_id].iloc[0]

        col_img1, col_img2 = st.columns(2)
        with col_img1:
          path1 = str(row_img.get("img1", "") or "").strip()
          if path1.startswith("http"):
            st.link_button("🖼️ Mở Ảnh 1 (Google Drive)", path1, use_container_width=True)
          elif path1:
            # Đường dẫn ảnh cũ lưu trên ổ cứng local (trước khi chuyển sang Drive)
            if os.path.exists(path1):
              st.image(path1, use_container_width=True)
            else:
              st.warning(f"⚠️ Không tìm thấy file ảnh 1 trên máy chủ ({path1}).")
          else:
            st.caption("Chưa có Ảnh 1 cho bản ghi này.")
        with col_img2:
          path2 = str(row_img.get("img2", "") or "").strip()
          if path2.startswith("http"):
            st.link_button("🖼️ Mở Ảnh 2 (Google Drive)", path2, use_container_width=True)
          elif path2:
            if os.path.exists(path2):
              st.image(path2, use_container_width=True)
            else:
              st.warning(f"⚠️ Không tìm thấy file ảnh 2 trên máy chủ ({path2}).")
          else:
            st.caption("Chưa có Ảnh 2 cho bản ghi này.")

      else:
        st.success("🎉 Chưa ghi nhận ca phát sinh sai hỏng nào!")
    except Exception as e:
      st.error(f"Lỗi tải báo cáo sai hỏng: {e}")



# ================= 10. TAB 6: CÀI ĐẶT DANH MỤC (LỖI SAI HỎNG / CÔNG VIỆC CON) =================
if nav == "7.1 Lỗi Sai Hỏng":
  render_section_heading("🐞 CÀI ĐẶT DANH MỤC: LỖI SAI HỎNG")
  render_catalog_manager("tb_dm_loai_loi", "ten_loi", "Kiểu Sai Hỏng")

if nav == "7.2 Công Việc Con":
  render_section_heading("📋 CÀI ĐẶT DANH MỤC: CÔNG VIỆC CON")
  render_catalog_manager("tb_dm_cong_viec", "ten_cong_viec", "Công Việc Con")

# ================= 11. TAB 7: CÀI ĐẶT MỤC TIÊU CHẤT LƯỢNG =================
if nav == "7.3 Mục tiêu chất lượng":
  render_section_heading("🎯 CÀI ĐẶT DANH MỤC: MỤC TIÊU CHẤT LƯỢNG")
  st.markdown("Nhập tỷ lệ lỗi mục tiêu (%) cho từng phân hệ và dòng sản phẩm. Hệ thống sẽ sử dụng dữ liệu này để vẽ đường ranh giới mục tiêu trên biểu đồ.")
  
  try:
    conn = get_db_connection()
    df_targets = pd.read_sql_query("SELECT id, phan_he, mat_prefix, muc_tieu FROM tb_dm_muc_tieu ORDER BY phan_he ASC, mat_prefix ASC", conn)
    conn.close()
  except Exception as ex:
    df_targets = pd.DataFrame(columns=["id", "phan_he", "mat_prefix", "muc_tieu"])
    st.error(f"Lỗi tải danh mục mục tiêu: {ex}")

  df_t_disp = df_targets.rename(
      columns={"id": "ID", "phan_he": "Phân Hệ / Xưởng", "mat_prefix": "Dòng Sản Phẩm / Mã Con", "muc_tieu": "Mục Tiêu Lỗi (%)"}
  )

  edited_t = st.data_editor(
      df_t_disp,
      use_container_width=True,
      hide_index=True,
      num_rows="dynamic",
      column_config={
          "ID": st.column_config.NumberColumn("ID", disabled=True),
          "Phân Hệ / Xưởng": st.column_config.SelectboxColumn(
              "Phân Hệ / Xưởng", options=PHAN_HE_OPTIONS
          ),
          "Mục Tiêu Lỗi (%)": st.column_config.NumberColumn(
              "Mục Tiêu Lỗi (%)", min_value=0.0, max_value=100.0, format="%.2f"
          )
      },
      key="editor_muc_tieu"
  )

  if st.button("💾 Lưu Mục Tiêu Chất Lượng", type="primary", key="save_muc_tieu"):
    try:
      conn = get_db_connection()
      cursor = conn.cursor()
      original_ids = set(df_t_disp["ID"].dropna().astype(int))
      edited_ids = set(edited_t["ID"].dropna().astype(int))
      for did in original_ids - edited_ids:
        cursor.execute("DELETE FROM tb_dm_muc_tieu WHERE id = ?", (int(did),))
      
      for _, row in edited_t.iterrows():
        ph = str(row["Phân Hệ / Xưởng"]).strip()
        mp = str(row["Dòng Sản Phẩm / Mã Con"]).strip()
        mt = float(row["Mục Tiêu Lỗi (%)"]) if pd.notna(row["Mục Tiêu Lỗi (%)"]) else 0.0
        
        if not ph or not mp or str(ph) == "nan" or str(mp) == "nan":
          continue
          
        if pd.isna(row["ID"]):
          cursor.execute(
              "INSERT INTO tb_dm_muc_tieu (phan_he, mat_prefix, muc_tieu) VALUES (?, ?, ?)",
              (ph, mp, mt),
          )
        else:
          cursor.execute(
              "UPDATE tb_dm_muc_tieu SET phan_he = ?, mat_prefix = ?, muc_tieu = ? WHERE id = ?",
              (ph, mp, mt, int(row["ID"])),
          )
      conn.commit()
      conn.close()
      st.success("✅ Đã lưu mục tiêu chất lượng!")
      st.rerun()
    except Exception as ex:
      st.error(f"Lỗi lưu thay đổi: {ex}")