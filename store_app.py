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
        return df
    except Exception:
        return pd.DataFrame(columns=["date", "store", "time_slot", "budget_sales", "prev_sales", "prev_cust", "budget_cust"])

@st.cache_data(ttl=600)
def load_delivery_data():
    try:
        df = conn.read(worksheet="delivery_sales", ttl=0).dropna(how="all")
        if not df.empty:
            df["time_slot"] = df["time_slot"].apply(normalize_time_slot)
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
    df = load_reservations()
    idx = df.index[df["id"] == res_id].tolist()
    if idx:
        df.at[idx[0], "date"] = date_str
        df.at[idx[0], "time_slot"] = time_slot
        df.at[idx[0], "management_name"] = str(mng_name)
        df.at[idx[0], "amount"] = amount
        df.at[idx[0], "customer_name"] = str(cust_name) if cust_name else ""
        df.at[idx[0], "phone"] = str(phone) if phone else ""
        df.at[idx[0], "notes"] = str(notes) if notes else ""
        conn.update(worksheet="reservations", data=df)
        load_reservations.clear()

def delete_reservation(res_id):
    df = load_reservations()
    df = df[df["id"] != res_id]
    conn.update(worksheet="reservations", data=df)
    load_reservations.clear()
    
    sales_df = load_reservation_sales()
    if not sales_df.empty:
        sales_df = sales_df[sales_df["reservation_id"] != res_id]
        conn.update(worksheet="reservation_sales", data=sales_df)
        load_reservation_sales.clear()

def save_reservation_sale(res_id, store, date_str, time_slot, amount):
    df = load_reservation_sales()
    sale_id = str(uuid.uuid4())
    new_row = pd.DataFrame([{"id": sale_id, "reservation_id": res_id, "store": store, "date": date_str, "time_slot": time_slot, "amount": amount}])
    df = pd.concat([df, new_row], ignore_index=True)
    conn.update(worksheet="reservation_sales", data=df)
    load_reservation_sales.clear()

def update_reservation_sale(sale_id, date_str, time_slot, amount):
    df = load_reservation_sales()
    idx = df.index[df["id"] == sale_id].tolist()
    if idx:
        df.at[idx[0], "date"] = date_str
        df.at[idx[0], "time_slot"] = time_slot
        df.at[idx[0], "amount"] = amount
        conn.update(worksheet="reservation_sales", data=df)
        load_reservation_sales.clear()

def delete_reservation_sale(sale_id):
    df = load_reservation_sales()
    df = df[df["id"] != sale_id]
    conn.update(worksheet="reservation_sales", data=df)
    load_reservation_sales.clear()

# --- 表生成関数（異常値を防ぎ、レイアウトを整えた完成版） ---
def generate_month_df(selected_store, year, month, target_slot_keyword, bp_df, deliv_df, res_df, res_sales_df, airregi_df):
    _, max_days = calendar.monthrange(year, month)

    # 1. 予算・前年データの事前辞書化
    bp_dict = {}
    if not bp_df.empty:
        bp_sub = bp_df[bp_df["store"] == selected_store]
        # 「合計」タブ(None)の場合はフィルタせず、日付ごとの全シフトを合算する。
        # 特定シフトの場合はそのシフトだけでフィルタする。
        if target_slot_keyword:
            bp_sub = bp_sub[bp_sub["time_slot"] == target_slot_keyword]
            
        if not bp_sub.empty:
            b_cust_col = None
            for c in ["budget_cust", "予算客数", "b_cust", "budget_customers"]:
                if c in bp_sub.columns:
                    b_cust_col = c
                    break

            # 数値型に変換してからSUMをとる（異常な文字列結合を防ぐ）
            bp_sub = bp_sub.copy()
            bp_sub["prev_sales_n"] = pd.to_numeric(bp_sub["prev_sales"], errors="coerce").fillna(0)
            bp_sub["budget_sales_n"] = pd.to_numeric(bp_sub["budget_sales"], errors="coerce").fillna(0)
            bp_sub["prev_cust_n"] = pd.to_numeric(bp_sub["prev_cust"], errors="coerce").fillna(0)
            if b_cust_col:
                bp_sub["budget_cust_n"] = pd.to_numeric(bp_sub[b_cust_col], errors="coerce").fillna(0)
            else:
                bp_sub["budget_cust_n"] = 0

            for d, grp in bp_sub.groupby("date"):
                bp_dict[str(d).strip()] = {
                    "prev_sales": int(grp["prev_sales_n"].sum()),
                    "budget_sales": int(grp["budget_sales_n"].sum()),
                    "prev_cust": int(grp["prev_cust_n"].sum()),
                    "budget_cust": int(grp["budget_cust_n"].sum())
                }

    # 2. デリバリー売上の事前辞書化
    deliv_dict = {}
    if not deliv_df.empty:
        d_sub = deliv_df[deliv_df["store"] == selected_store]
        if target_slot_keyword:
            d_sub = d_sub[d_sub["time_slot"] == target_slot_keyword]
        if not d_sub.empty:
            d_sub = d_sub.copy()
            d_sub["uber_n"] = pd.to_numeric(d_sub["uber"], errors="coerce").fillna(0)
            d_sub["rn_n"] = pd.to_numeric(d_sub["rocket_now"], errors="coerce").fillna(0)
            d_sub["demae_n"] = pd.to_numeric(d_sub["demaekan"], errors="coerce").fillna(0)
            d_sub["other_n"] = pd.to_numeric(d_sub["other"], errors="coerce").fillna(0)
            d_sub["total_deliv"] = d_sub["uber_n"] + d_sub["rn_n"] + d_sub["demae_n"]
            for d, grp in d_sub.groupby("date"):
                deliv_dict[str(d).strip()] = {
                    "deliv": int(grp["total_deliv"].sum()),
                    "other": int(grp["other_n"].sum())
                }

    # 3. 予約金の事前辞書化
    res_dict = {}
    if not res_df.empty:
        r_sub = res_df[res_df["store"] == selected_store]
        if target_slot_keyword:
            r_sub = r_sub[(r_sub["time_slot"] == target_slot_keyword) | (r_sub["time_slot"] == "時間帯未指定")]
        if not r_sub.empty:
            r_sub = r_sub.copy()
            r_sub["amount_n"] = pd.to_numeric(r_sub["amount"], errors="coerce").fillna(0)
            for d, grp in r_sub.groupby("date"):
                res_dict[str(d).strip()] = -int(grp["amount_n"].sum())

    # 4. 売上計上の事前辞書化
    res_sales_dict = {}
    if not res_sales_df.empty:
        rs_sub = res_sales_df[res_sales_df["store"] == selected_store]
        if target_slot_keyword:
            rs_sub = rs_sub[rs_sub["time_slot"] == target_slot_keyword]
        if not rs_sub.empty:
            rs_sub = rs_sub.copy()
            rs_sub["amount_n"] = pd.to_numeric(rs_sub["amount"], errors="coerce").fillna(0)
            for d, grp in rs_sub.groupby("date"):
                res_sales_dict[str(d).strip()] = int(grp["amount_n"].sum())

    # 5. Airレジデータの事前辞書化
    air_dict = {}
    if airregi_df is not None and not airregi_df.empty and "date" in airregi_df.columns:
        short_key = get_short_store_key(selected_store)
        c_store = "c0" if "c0" in airregi_df.columns else airregi_df.columns[1]
        c_slot = "c1" if "c1" in airregi_df.columns else airregi_df.columns[2]
        c_sales = "c2" if "c2" in airregi_df.columns else airregi_df.columns[3]
        c_cust = "c5" if "c5" in airregi_df.columns else airregi_df.columns[6]
        c_other = "c11" if "c11" in airregi_df.columns else (airregi_df.columns[12] if len(airregi_df.columns) > 12 else None)

        store_air = airregi_df[airregi_df[c_store].astype(str).str.contains(short_key, na=False)]
        if target_slot_keyword is None:
            matched_air = store_air[store_air[c_slot].astype(str).str.strip() == "合計"]
        else:
            matched_air = store_air[store_air[c_slot].astype(str).str.contains(target_slot_keyword[0], na=False)]

        for _, row in matched_air.iterrows():
            d_key = str(row["date"]).strip()
            a_sales = parse_int(row[c_sales])
            a_cust = parse_int(row[c_cust])
            a_oth = parse_int(row[c_other]) if c_other and pd.notna(row.get(c_other)) else 0
            air_dict[d_key] = (a_sales, a_cust, a_oth)

    weekday_kanji_list = ["月", "火", "水", "木", "金", "土", "日"]
    rows = []

    for day in range(1, max_days + 1):
        date_str = f"{year:04d}-{month:02d}-{day:02d}"

        bp_item = bp_dict.get(date_str, {})
        p_sales = bp_item.get("prev_sales", 0)
        b_sales = bp_item.get("budget_sales", 0)
        p_cust = bp_item.get("prev_cust", 0)
        b_cust = bp_item.get("budget_cust", 0)

        air_sales, air_cust, air_other_sales = 0, 0, 0
        if date_str in air_dict:
            air_sales, air_cust, air_other_sales = air_dict[date_str]
        else:
            raw_df = find_airregi_data_in_memory(date_str, airregi_df)
            if raw_df is not None:
                try:
                    short_key = get_short_store_key(selected_store)
                    store_df = raw_df[raw_df[0].astype(str).str.contains(short_key, na=False)]
                    matched = store_df[store_df[1].astype(str).str.strip() == "合計"] if target_slot_keyword is None else match_slot_row(store_df, target_slot_keyword)
                    if not matched.empty:
                        air_sales = parse_int(matched.iloc[0, 2])
                        air_cust = parse_int(matched.iloc[0, 5])
                        if matched.shape[1] > 11: air_other_sales = parse_int(matched.iloc[0, 11])
                except Exception:
                    pass

        deliv_item = deliv_dict.get(date_str, {})
        day_delivery_sales = deliv_item.get("deliv", 0)
        deliv_other_sales = deliv_item.get("other", 0)

        day_res_deposit = res_dict.get(date_str, 0)
        day_res_sales = res_sales_dict.get(date_str, 0)

        adjusted_air_sales = air_sales + day_res_deposit + day_res_sales
        total_other_sales = air_other_sales + deliv_other_sales
        actual_sales = adjusted_air_sales + day_delivery_sales + total_other_sales
        actual_cust = air_cust

        # 比率・単価計算
        prev_ratio = f"{(actual_sales / p_sales * 100):.1f}%" if p_sales > 0 and actual_sales != 0 else ("-" if actual_sales == 0 else "0.0%")
        budget_ratio = f"{(actual_sales / b_sales * 100):.1f}%" if b_sales > 0 and actual_sales != 0 else ("-" if actual_sales == 0 else "0.0%")
        
        p_price = int(p_sales / p_cust) if p_cust > 0 else 0
        b_price = int(b_sales / b_cust) if b_cust > 0 else 0
        a_price = int(actual_sales / actual_cust) if actual_cust > 0 else 0
        
        cust_prev_ratio = f"{(actual_cust / p_cust * 100):.1f}%" if p_cust > 0 and actual_cust > 0 else "-"
        cust_budget_ratio = f"{(actual_cust / b_cust * 100):.1f}%" if b_cust > 0 and actual_cust > 0 else "-"
        
        price_prev_ratio = f"{(a_price / p_price * 100):.1f}%" if p_price > 0 and a_price != 0 else "-"
        price_budget_ratio = f"{(a_price / b_price * 100):.1f}%" if b_price > 0 and a_price != 0 else "-"
        
        weekday_kanji = weekday_kanji_list[pd.to_datetime(date_str).weekday()]

        rows.append({
            "日": day, "曜日": weekday_kanji, 
            "前年売上": p_sales, "予算売上": b_sales, "実績売上": actual_sales, "前年比": prev_ratio, "予算比": budget_ratio, 
            "店頭売上": adjusted_air_sales, "デリバリー売上": day_delivery_sales, "その他売上": total_other_sales, "予約金": day_res_deposit, "売上計上": day_res_sales,
            "前年客数": p_cust, "予算客数": b_cust, "実績客数": actual_cust, "客数前年比": cust_prev_ratio, "客数予算比": cust_budget_ratio, 
            "前年単価": p_price, "予算単価": b_price, "実績単価": a_price, "単価前年比": price_prev_ratio, "単価予算比": price_budget_ratio
        })

    return pd.DataFrame(rows)

def calculate_summary_rows(df_slot):
