import streamlit as st
import pandas as pd
import os
import sys
import re
from datetime import datetime
import calendar
import glob
import subprocess
import io
import uuid
from streamlit_gsheets import GSheetsConnection
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ページ設定
st.set_page_config(
    page_title="売上月報管理システム",
    page_icon="📊",
    layout="wide"
)

# スプレッドシートとの接続
conn = st.connection("gsheets", type=GSheetsConnection)

# Airレジからの取得データ(CSV)一時保存フォルダ
OUTPUT_DIR = os.path.abspath("./output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Airレジ専用スプレッドシートのURL取得（Secretsから）
AIRREGI_SPREADSHEET_URL = None
if "AIRREGI_SPREADSHEET" in st.secrets:
    AIRREGI_SPREADSHEET_URL = str(st.secrets["AIRREGI_SPREADSHEET"])
elif "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
    if "AIRREGI_SPREADSHEET" in st.secrets["connections"]["gsheets"]:
        AIRREGI_SPREADSHEET_URL = str(st.secrets["connections"]["gsheets"]["AIRREGI_SPREADSHEET"])

# 店舗定義
STORE_CREDENTIALS = {
    "三田おにぎり 本店": {"id": "001", "pass": "0000", "short_label": "三田店"},
    "おにぎり維心 恵比寿店": {"id": "002", "pass": "0000", "short_label": "恵比寿店"},
    "おにぎり維心 五反田店": {"id": "003", "pass": "0000", "short_label": "五反田店"},
    "おにぎり維心 中目黒店": {"id": "004", "pass": "0000", "short_label": "中目黒店"}
}

STORES = list(STORE_CREDENTIALS.keys())
TIME_SLOTS = ["早番", "中番", "遅番"]

def normalize_time_slot(val):
    s = str(val)
    if "早" in s: return "早番"
    elif "中" in s: return "中番"
    elif "遅" in s: return "遅番"
    return s

def parse_int(val):
    if pd.isna(val): return 0
    clean_str = re.sub(r"[^\d-]", "", str(val))
    try: return int(clean_str)
    except ValueError: return 0

def get_short_store_key(store_name):
    for key in ["三田", "恵比寿", "五反田", "中目黒"]:
        if key in store_name: return key
    return store_name

def match_slot_row(store_df, slot_name):
    short_slot = slot_name[0] if slot_name else ""
    return store_df[store_df[1].astype(str).str.contains(short_slot, na=False)]

# --- キャッシュ付きデータロード関数群 ---
@st.cache_data(ttl=600)
def load_budget_prev_data():
    try:
        df = conn.read(worksheet="budget_prev", ttl=0).dropna(how="all")
        if not df.empty:
            df["time_slot"] = df["time_slot"].apply(normalize_time_slot)
            df["date"] = df["date"].astype(str).str.strip()
        return df
    except Exception:
        return pd.DataFrame(columns=["date", "store", "time_slot", "budget_sales", "prev_sales", "prev_cust", "budget_cust"])

@st.cache_data(ttl=600)
def load_delivery_data():
    try:
        df = conn.read(worksheet="delivery_sales", ttl=0).dropna(how="all")
        if not df.empty:
            df["time_slot"] = df["time_slot"].apply(normalize_time_slot)
            df["date"] = df["date"].astype(str).str.strip()
            if "rn" in df.columns and "rocket_now" not in df.columns:
                df = df.rename(columns={"rn": "rocket_now"})
            return df
        return pd.DataFrame(columns=["date", "store", "time_slot", "uber", "rocket_now", "demaekan", "other"])
    except Exception:
        return pd.DataFrame(columns=["date", "store", "time_slot", "uber", "rocket_now", "demaekan", "other"])

@st.cache_data(ttl=600)
def load_reservations():
    try:
        df = conn.read(worksheet="reservations", ttl=0).dropna(how="all")
        if not df.empty:
            df["id"] = df["id"].astype(str)
            df["date"] = df["date"].astype(str).str.strip()
            df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0).astype(int)
            return df
        return pd.DataFrame(columns=["id", "store", "date", "time_slot", "management_name", "amount", "customer_name", "phone", "notes"])
    except Exception:
        return pd.DataFrame(columns=["id", "store", "date", "time_slot", "management_name", "amount", "customer_name", "phone", "notes"])

@st.cache_data(ttl=600)
def load_reservation_sales():
    try:
        df = conn.read(worksheet="reservation_sales", ttl=0).dropna(how="all")
        if not df.empty:
            df["id"] = df["id"].astype(str)
            df["date"] = df["date"].astype(str).str.strip()
            df["reservation_id"] = df["reservation_id"].astype(str)
            df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0).astype(int)
            return df
        return pd.DataFrame(columns=["id", "reservation_id", "store", "date", "time_slot", "amount"])
    except Exception:
        return pd.DataFrame(columns=["id", "reservation_id", "store", "date", "time_slot", "amount"])

@st.cache_data(ttl=600)
def get_airregi_sheet_for_year(year_str):
    if not AIRREGI_SPREADSHEET_URL:
        return pd.DataFrame()
    try:
        df = conn.read(spreadsheet=AIRREGI_SPREADSHEET_URL, worksheet=str(year_str), ttl=0).dropna(how="all")
        return df if not df.empty else pd.DataFrame()
    except Exception:
        return pd.DataFrame()

def save_airregi_csv_to_sheets(target_date_str=None):
    if not AIRREGI_SPREADSHEET_URL:
        st.warning("⚠️ AIRREGI_SPREADSHEET の設定が見つかりません。Secretsを確認してください。")
        return 0

    try:
        pattern = os.path.join(OUTPUT_DIR, f"*{target_date_str}*.csv") if target_date_str else os.path.join(OUTPUT_DIR, "*.csv")
        csv_files = glob.glob(pattern)
        if not csv_files:
            return 0

        files_by_year = {}
        for f in csv_files:
            m = re.search(r"(\d{4})-(\d{2}-\d{2})", os.path.basename(f))
            if m:
                year_str = m.group(1)
                date_str = f"{year_str}-{m.group(2)}"
                files_by_year.setdefault(year_str, []).append((date_str, f))

        total_synced_dates = 0
        for y_str, f_list in files_by_year.items():
            new_records = []
            for date_str, f_path in f_list:
                try:
                    raw = pd.read_csv(f_path, header=None)
                    if raw.empty: continue
                    if raw.shape[1] > 1:
                        raw = raw[~raw[1].astype(str).str.strip().isin(["時間帯", "現金"])]
                    if raw.empty: continue
                    raw.insert(0, "date", date_str)
                    raw.columns = ["date"] + [f"c{i}" for i in range(raw.shape[1] - 1)]
                    new_records.append(raw)
                except Exception:
                    continue

            if not new_records:
                continue

            year_new_df = pd.concat(new_records, ignore_index=True)
            try:
                current_sheet_df = conn.read(spreadsheet=AIRREGI_SPREADSHEET_URL, worksheet=str(y_str), ttl=0).dropna(how="all")
            except Exception:
                current_sheet_df = pd.DataFrame()

            if not current_sheet_df.empty and "date" in current_sheet_df.columns:
                if "c1" in current_sheet_df.columns:
                    current_sheet_df = current_sheet_df[~current_sheet_df["c1"].astype(str).str.strip().isin(["時間帯", "現金"])]
                dates_to_replace = year_new_df["date"].unique()
                kept_df = current_sheet_df[~current_sheet_df["date"].isin(dates_to_replace)]
                final_df = pd.concat([kept_df, year_new_df], ignore_index=True)
            else:
                final_df = year_new_df

            if "date" in final_df.columns:
                sort_cols = ["date"]
                if "c0" in final_df.columns: sort_cols.append("c0")
                final_df = final_df.sort_values(by=sort_cols).reset_index(drop=True)

            conn.update(spreadsheet=AIRREGI_SPREADSHEET_URL, worksheet=str(y_str), data=final_df)
            total_synced_dates += len(year_new_df["date"].unique())

        st.cache_data.clear()
        return total_synced_dates
    except Exception as e:
        st.error(f"スプレッドシートへの保存中にエラーが発生しました: {e}")
        return 0

def find_airregi_data_in_memory(date_str, df_year):
    if df_year is not None and not df_year.empty and "date" in df_year.columns:
        sub = df_year[df_year["date"].astype(str).str.strip() == date_str.strip()]
        if not sub.empty:
            sub_data = sub.drop(columns=["date"]).reset_index(drop=True)
            sub_data.columns = range(sub_data.shape[1])
            return sub_data

    d_dash = date_str
    d_nodash = date_str.replace("-", "")
    d_under = date_str.replace("-", "_")
    all_files = glob.glob(os.path.join(OUTPUT_DIR, "*.csv"))
    matched = [f for f in all_files if d_dash in os.path.basename(f) or d_nodash in os.path.basename(f) or d_under in os.path.basename(f)]
    if matched:
        matched.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        try: return pd.read_csv(matched[0], header=None)
        except Exception: pass
    return None

def get_available_months():
    months = set()
    display_year = st.session_state.get("display_year", str(datetime.now().year))
    df_cur_year = get_airregi_sheet_for_year(display_year)
    if not df_cur_year.empty and "date" in df_cur_year.columns:
        for d in df_cur_year["date"].dropna().astype(str):
            m = re.search(r"(\d{4})-(\d{2})", d)
            if m: months.add(f"{m.group(1)}-{m.group(2)}")

    all_files = glob.glob(os.path.join(OUTPUT_DIR, "*.csv"))
    for f in all_files:
        base = os.path.basename(f)
        m = re.search(r"(\d{4})-(\d{2})-\d{2}", base)
        if m: months.add(f"{m.group(1)}-{m.group(2)}")

    if not months:
        return [f"{display_year}-10", f"{display_year}-01"]
    return sorted(list(months), reverse=True)

# --- 予約・デリバリー更新関数 ---
def save_delivery_data(date_str, store, time_slot, uber, rocket_now, demaekan, other):
    df = load_delivery_data()
    df = df[~((df["date"] == date_str) & (df["store"] == store) & (df["time_slot"] == time_slot))]
    new_row = pd.DataFrame([{"date": date_str, "store": store, "time_slot": time_slot, "uber": uber, "rocket_now": rocket_now, "demaekan": demaekan, "other": other}])
    df = pd.concat([df, new_row], ignore_index=True)
    conn.update(worksheet="delivery_sales", data=df)
    load_delivery_data.clear()

def save_reservation(store, date_str, time_slot, mng_name, amount, cust_name, phone, notes):
    df = load_reservations()
    res_id = str(uuid.uuid4())
    new_row = pd.DataFrame([{"id": res_id, "store": store, "date": date_str, "time_slot": time_slot, "management_name": mng_name, "amount": amount, "customer_name": cust_name, "phone": phone, "notes": notes}])
    df = pd.concat([df, new_row], ignore_index=True)
    conn.update(worksheet="reservations", data=df)
    load_reservations.clear()

def update_reservation(res_id, date_str, time_slot, mng_name, amount, cust_name, phone, notes):
    df =
