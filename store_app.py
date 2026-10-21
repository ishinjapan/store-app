import calendar
from datetime import datetime
import glob
import io
import os
import re
import subprocess
import sys
import time
import uuid
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st
from streamlit_gsheets import GSheetsConnection

APP_VARIANT = os.environ.get("STORE_APP_VARIANT", "combined")
APP_MODE_BY_VARIANT = {
    "admin": "📊 本部用（月報一覧表示）",
    "store": "📱 店舗用（実績入力）",
}
if APP_VARIANT not in {*APP_MODE_BY_VARIANT, "combined"}:
  raise ValueError(f"未対応のアプリ種別です: {APP_VARIANT}")

# ページ設定
st.set_page_config(
    page_title=(
        "本部用 月報一覧"
        if APP_VARIANT == "admin"
        else (
            "店舗用 実績・予約入力"
            if APP_VARIANT == "store"
            else "売上月報管理システム"
        )
    ),
    page_icon="📱" if APP_VARIANT == "store" else "📊",
    layout="wide",
)
if APP_VARIANT == "store":
  st.markdown(
      """
      <style>
      section[data-testid="stSidebar"], [data-testid="collapsedControl"] {
          display: none !important;
      }
      </style>
      """,
      unsafe_allow_html=True,
  )
elif APP_VARIANT == "admin":
  st.markdown(
      """
      <style>
      @media (max-width: 640px) {
          .st-key-admin_year_navigation [data-testid="stHorizontalBlock"] {
              display: grid !important;
              grid-template-columns: minmax(44px, 1fr) minmax(60px, 2fr) minmax(44px, 1fr) !important;
              align-items: center !important;
          }
          .st-key-admin_year_navigation [data-testid="column"] {
              min-width: 0 !important;
              width: auto !important;
              max-width: none !important;
              flex: none !important;
          }
      }
      [data-testid="manage-app-button"],
      [data-testid="stAppDeployButton"],
      .stAppDeployButton,
      button[aria-label="Manage app"] {
          display: none !important;
      }
      </style>
      """,
      unsafe_allow_html=True,
  )
st.session_state["performance_timings"] = []
st.session_state.setdefault("monthly_report_cache", {})
APP_RUN_STARTED_AT = time.perf_counter()

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
    AIRREGI_SPREADSHEET_URL = str(
        st.secrets["connections"]["gsheets"]["AIRREGI_SPREADSHEET"]
    )

# 店舗定義
STORE_CREDENTIALS = {
    "三田おにぎり 本店": {
        "id": "001",
        "pass": "0000",
        "short_label": "三田店",
    },
    "おにぎり維心 恵比寿店": {
        "id": "002",
        "pass": "0000",
        "short_label": "恵比寿店",
    },
    "おにぎり維心 五反田店": {
        "id": "003",
        "pass": "0000",
        "short_label": "五反田店",
    },
    "おにぎり維心 中目黒店": {
        "id": "004",
        "pass": "0000",
        "short_label": "中目黒店",
    },
}

STORES = list(STORE_CREDENTIALS.keys())
TIME_SLOTS = ["早番", "中番", "遅番"]
DELIVERY_UNIT_PRICE = 1800
AIRREGI_REPORT_START_DATE = "2026-10-01"


def normalize_time_slot(val):
  s = str(val)
  if "早" in s:
    return "早番"
  elif "中" in s:
    return "中番"
  elif "遅" in s:
    return "遅番"
  return s


def parse_int(val):
  if pd.isna(val):
    return 0
  clean_str = re.sub(r"[^\d.-]", "", str(val))
  try:
    return int(float(clean_str))
  except (ValueError, OverflowError):
    return 0


def normalize_date_key(val):
  if pd.isna(val):
    return ""
  parsed = pd.to_datetime(val, errors="coerce")
  return parsed.strftime("%Y-%m-%d") if pd.notna(parsed) else ""


def log_performance_time(label, started_at):
  elapsed_ms = (time.perf_counter() - started_at) * 1000
  print(f"[PERF] {label}: {elapsed_ms:.0f} ms")
  st.session_state.setdefault("performance_timings", []).append({
      "処理": label,
      "時間（秒）": round(elapsed_ms / 1000, 3),
  })


def clear_monthly_report_cache():
  st.session_state["monthly_report_cache"] = {}


def refresh_monthly_report_data():
  load_budget_prev_data.clear()
  load_delivery_data.clear()
  load_reservations.clear()
  load_reservation_sales.clear()
  get_airregi_sheet_for_year.clear()
  clear_monthly_report_cache()


def timed_call(label, function, *args):
  started_at = time.perf_counter()
  try:
    return function(*args)
  finally:
    log_performance_time(label, started_at)


def estimate_delivery_count(delivery_sales):
  sales = max(0, parse_int(delivery_sales))
  return (sales + DELIVERY_UNIT_PRICE // 2) // DELIVERY_UNIT_PRICE


def normalize_management_name(value):
  if pd.isna(value):
    return ""
  if isinstance(value, (int, float)) and not isinstance(value, bool):
    numeric_value = float(value)
    if numeric_value.is_integer():
      return str(int(numeric_value))
  return str(value)


def get_short_store_key(store_name):
  for key in ["三田", "恵比寿", "五反田", "中目黒"]:
    if key in store_name:
      return key
  return store_name


def match_slot_row(store_df, slot_name):
  short_slot = slot_name[0] if slot_name else ""
  return store_df[store_df[1].astype(str).str.contains(short_slot, na=False)]


# --- キャッシュ付きデータロード関数群 ---
def deduplicate_budget_prev_data(df):
  if df.empty:
    return df

  df = df.copy()
  df["time_slot"] = df["time_slot"].apply(normalize_time_slot)
  date_keys = df["date"].apply(normalize_date_key)
  df["_normalized_date"] = date_keys
  valid_date = date_keys != ""
  duplicate_indices = df.loc[valid_date].index[
      pd.DataFrame({
          "date": date_keys.loc[valid_date],
          "store": df.loc[valid_date, "store"].astype(str).str.strip(),
          "time_slot": df.loc[valid_date, "time_slot"],
      }).duplicated(keep="first")
  ]
  return df.drop(index=duplicate_indices)


def prepare_actual_customer_csv(uploaded_file):
  df = pd.read_csv(uploaded_file)
  required_columns = ["date", "store", "time_slot", "actual_cust"]
  missing_columns = [column for column in required_columns if column not in df.columns]
  if missing_columns:
    raise ValueError(
        "CSVに必要な列がありません: " + ", ".join(missing_columns)
    )

  df = df[required_columns].copy()
  df["date"] = df["date"].apply(normalize_date_key)
  df["store"] = df["store"].astype(str).str.strip()
  df["time_slot"] = df["time_slot"].apply(normalize_time_slot)
  if (df["date"] == "").any():
    raise ValueError("日付を読み取れない行があります。")
  if not df["store"].isin(STORES).all():
    raise ValueError("登録されていない店舗名がCSVに含まれています。")
  if not df["time_slot"].isin(TIME_SLOTS).all():
    raise ValueError("時間帯は早番・中番・遅番のいずれかを指定してください。")

  numeric_counts = pd.to_numeric(df["actual_cust"], errors="coerce")
  invalid_counts = df["actual_cust"].notna() & numeric_counts.isna()
  if invalid_counts.any():
    raise ValueError("実績客数に数値以外の値が含まれています。")
  if (numeric_counts.dropna() < 0).any():
    raise ValueError("実績客数に負数は指定できません。")
  if ((numeric_counts.dropna() % 1) != 0).any():
    raise ValueError("実績客数は整数で指定してください。")

  df["actual_cust"] = numeric_counts.astype("Int64")
  key_columns = ["date", "store", "time_slot"]
  if df.duplicated(key_columns).any():
    raise ValueError("日付・店舗・時間帯が重複する行があります。")
  return df


def update_actual_customer_column(uploaded_file):
  csv_df = prepare_actual_customer_csv(uploaded_file)
  sheet_df = conn.read(worksheet="budget_prev", ttl=0).dropna(how="all")
  key_columns = ["date", "store", "time_slot"]
  missing_columns = [column for column in key_columns if column not in sheet_df]
  if missing_columns:
    raise ValueError(
        "budget_prevシートに必要な列がありません: "
        + ", ".join(missing_columns)
    )

  sheet_df = sheet_df.copy()
  sheet_df["_date_key"] = sheet_df["date"].apply(normalize_date_key)
  sheet_df["_store_key"] = sheet_df["store"].astype(str).str.strip()
  sheet_df["_slot_key"] = sheet_df["time_slot"].apply(normalize_time_slot)
  sheet_key_columns = ["_date_key", "_store_key", "_slot_key"]
  sheet_keys = sheet_df[sheet_key_columns].rename(
      columns=dict(zip(sheet_key_columns, key_columns))
  )
  if sheet_keys.duplicated(key_columns).any():
    raise ValueError(
        "budget_prevシートに日付・店舗・時間帯の重複があります。"
        "データを変更せず中止しました。"
    )

  csv_keys = set(map(tuple, csv_df[key_columns].itertuples(index=False, name=None)))
  sheet_key_set = set(map(tuple, sheet_keys.itertuples(index=False, name=None)))
  if csv_keys != sheet_key_set:
    missing_in_csv = len(sheet_key_set - csv_keys)
    missing_in_sheet = len(csv_keys - sheet_key_set)
    raise ValueError(
        "CSVとbudget_prevシートの行が一致しないため、"
        "データを変更せず中止しました。"
        f"（CSVにないシート行: {missing_in_csv}、シートにないCSV行: {missing_in_sheet}）"
    )

  counts_by_key = csv_df.set_index(key_columns)["actual_cust"]
  csv_index = pd.MultiIndex.from_frame(sheet_keys[key_columns])
  actual_counts = counts_by_key.reindex(csv_index)
  sheet_df = sheet_df.drop(columns=sheet_key_columns)
  delivery_sales = (
      pd.to_numeric(sheet_df["delivery_sales"], errors="coerce").fillna(0)
      if "delivery_sales" in sheet_df.columns
      else pd.Series(0, index=sheet_df.index)
  )
  estimated_delivery_counts = delivery_sales.apply(estimate_delivery_count)
  delivery_counts = (
      pd.to_numeric(sheet_df["delivery_count"], errors="coerce")
      if "delivery_count" in sheet_df.columns
      else pd.Series(pd.NA, index=sheet_df.index, dtype="Float64")
  ).fillna(estimated_delivery_counts).astype("Int64")
  delivery_unit_prices = (
      pd.to_numeric(sheet_df["delivery_unit_price"], errors="coerce")
      if "delivery_unit_price" in sheet_df.columns
      else pd.Series(pd.NA, index=sheet_df.index, dtype="Float64")
  ).fillna(DELIVERY_UNIT_PRICE).astype("Int64")

  column_numbers = {}
  next_column = len(sheet_df.columns) + 1
  for column in ["actual_cust", "delivery_count", "delivery_unit_price"]:
    if column in sheet_df.columns:
      column_numbers[column] = sheet_df.columns.get_loc(column) + 1
    else:
      column_numbers[column] = next_column
      next_column += 1
  last_row = len(sheet_df) + 1
  worksheet = conn.client._select_worksheet(worksheet="budget_prev")
  worksheet.batch_update(
      [
          {
              "range": (
                  f"{get_column_letter(column_numbers[column])}1:"
                  f"{get_column_letter(column_numbers[column])}{last_row}"
              ),
              "values": [[column]]
              + [
                  ["" if pd.isna(value) else int(value)]
                  for value in column_values
              ],
          }
          for column, column_values in [
              ("actual_cust", actual_counts),
              ("delivery_count", delivery_counts),
              ("delivery_unit_price", delivery_unit_prices),
          ]
      ],
      value_input_option="USER_ENTERED",
  )
  load_budget_prev_data.clear()
  clear_monthly_report_cache()
  return (
      len(csv_df),
      int(csv_df["actual_cust"].isna().sum()),
      int(delivery_counts.sum()),
  )


@st.cache_data(ttl=60)
def load_budget_prev_data():
  try:
    df = conn.read(worksheet="budget_prev", ttl=0).dropna(how="all")
    return deduplicate_budget_prev_data(df)
  except Exception as error:
    raise RuntimeError(
        "budget_prevシートを読み込めません。Google Sheets接続先と"
        "サービスアカウントの共有権限を確認してください。"
    ) from error


@st.cache_data(ttl=60)
def load_delivery_data():
  try:
    df = conn.read(worksheet="delivery_sales", ttl=0).dropna(how="all")
    if not df.empty:
      df["time_slot"] = df["time_slot"].apply(normalize_time_slot)
      if "rn" in df.columns and "rocket_now" not in df.columns:
        df = df.rename(columns={"rn": "rocket_now"})
      return df
    return pd.DataFrame(
        columns=[
            "date",
            "store",
            "time_slot",
            "uber",
            "rocket_now",
            "demaekan",
            "other",
            "delivery_count",
            "delivery_unit_price",
        ]
    )
  except Exception:
    return pd.DataFrame(
        columns=[
            "date",
            "store",
            "time_slot",
            "uber",
            "rocket_now",
            "demaekan",
            "other",
            "delivery_count",
            "delivery_unit_price",
        ]
    )


@st.cache_data(ttl=60)
def load_reservations():
  try:
    df = conn.read(worksheet="reservations", ttl=0).dropna(how="all")
    if not df.empty:
      df["id"] = df["id"].astype(str)
      df["management_name"] = df["management_name"].apply(
          normalize_management_name
      )
      df["amount"] = (
          pd.to_numeric(df["amount"], errors="coerce").fillna(0).astype(int)
      )
      return df
    return pd.DataFrame(
        columns=[
            "id",
            "store",
            "date",
            "time_slot",
            "management_name",
            "amount",
            "customer_name",
            "phone",
            "notes",
        ]
    )
  except Exception:
    return pd.DataFrame(
        columns=[
            "id",
            "store",
            "date",
            "time_slot",
            "management_name",
            "amount",
            "customer_name",
            "phone",
            "notes",
        ]
    )


@st.cache_data(ttl=60)
def load_reservation_sales():
  try:
    df = conn.read(worksheet="reservation_sales", ttl=0).dropna(how="all")
    if not df.empty:
      df["id"] = df["id"].astype(str)
      df["reservation_id"] = df["reservation_id"].astype(str)
      df["amount"] = (
          pd.to_numeric(df["amount"], errors="coerce").fillna(0).astype(int)
      )
      return df
    return pd.DataFrame(
        columns=["id", "reservation_id", "store", "date", "time_slot", "amount"]
    )
  except Exception:
    return pd.DataFrame(
        columns=["id", "reservation_id", "store", "date", "time_slot", "amount"]
    )


@st.cache_data(ttl=60)
def get_airregi_sheet_for_year(year_str):
  if not AIRREGI_SPREADSHEET_URL:
    raise RuntimeError(
        "AIRREGI_SPREADSHEET が設定されていません。"
        "本部アプリのSecretsで、Airレジ年別シートの接続先を確認してください。"
    )
  try:
    df = conn.read(
        spreadsheet=AIRREGI_SPREADSHEET_URL, worksheet=str(year_str), ttl=0
    ).dropna(how="all")
    return df if not df.empty else pd.DataFrame()
  except Exception as error:
    raise RuntimeError(
        f"Airレジの{year_str}年シートを読み込めません。"
        "本部アプリのAIRREGI_SPREADSHEET設定、シート名、"
        "サービスアカウントの共有権限を確認してください。"
    ) from error


def save_airregi_csv_to_sheets(
    target_date_str=None, start_date_str=None, end_date_str=None
):
  if not AIRREGI_SPREADSHEET_URL:
    raise RuntimeError(
        "AIRREGI_SPREADSHEET の設定が見つかりません。Secretsを確認してください。"
    )
    return 0

  try:
    pattern = (
        os.path.join(OUTPUT_DIR, f"*{target_date_str}*.csv")
        if target_date_str
        else os.path.join(OUTPUT_DIR, "*.csv")
    )
    csv_files = glob.glob(pattern)
    if not csv_files:
      return 0

    files_by_year = {}
    for f in csv_files:
      m = re.search(r"(\d{4})-(\d{2}-\d{2})", os.path.basename(f))
      if m:
        year_str = m.group(1)
        date_str = f"{year_str}-{m.group(2)}"
        if start_date_str and date_str < start_date_str:
          continue
        if end_date_str and date_str > end_date_str:
          continue
        files_by_year.setdefault(year_str, []).append((date_str, f))

    total_synced_dates = 0
    for y_str, f_list in files_by_year.items():
      new_records = []
      for date_str, f_path in f_list:
        try:
          raw = pd.read_csv(f_path, header=None)
          if raw.empty:
            continue
          if raw.shape[1] > 1:
            raw = raw[
                ~raw[1].astype(str).str.strip().isin(["時間帯", "現金"])
            ]
          if raw.empty:
            continue
          raw.insert(0, "date", date_str)
          raw.columns = ["date"] + [f"c{i}" for i in range(raw.shape[1] - 1)]
          new_records.append(raw)
        except Exception:
          continue

      if not new_records:
        continue

      year_new_df = pd.concat(new_records, ignore_index=True)
      try:
        current_sheet_df = conn.read(
            spreadsheet=AIRREGI_SPREADSHEET_URL, worksheet=str(y_str), ttl=0
        ).dropna(how="all")
      except Exception as e:
        raise RuntimeError(
            f"{y_str}年シートの既存データを読み込めないため、"
            "既存データ保護のため同期を中止しました。"
        ) from e

      if not current_sheet_df.empty and "date" in current_sheet_df.columns:
        if "c1" in current_sheet_df.columns:
          current_sheet_df = current_sheet_df[
              ~current_sheet_df["c1"]
              .astype(str)
              .str.strip()
              .isin(["時間帯", "現金"])
          ]
        dates_to_replace = year_new_df["date"].unique()
        kept_df = current_sheet_df[
            ~current_sheet_df["date"].isin(dates_to_replace)
        ]
        final_df = pd.concat([kept_df, year_new_df], ignore_index=True)
      else:
        final_df = year_new_df

      if "date" in final_df.columns:
        sort_cols = ["date"]
        if "c0" in final_df.columns:
          sort_cols.append("c0")
        final_df = final_df.sort_values(by=sort_cols).reset_index(drop=True)

      conn.update(
          spreadsheet=AIRREGI_SPREADSHEET_URL,
          worksheet=str(y_str),
          data=final_df,
      )
      total_synced_dates += len(year_new_df["date"].unique())

    st.cache_data.clear()
    clear_monthly_report_cache()
    return total_synced_dates
  except Exception as e:
    raise RuntimeError(
        f"スプレッドシートへの保存中にエラーが発生しました: {e}"
    ) from e


def find_airregi_data_in_memory(date_str, df_year):
  if df_year is not None and not df_year.empty and "date" in df_year.columns:
    date_keys = df_year["date"].map(normalize_date_key)
    sub = df_year[date_keys == normalize_date_key(date_str)]
    if not sub.empty:
      sub_data = sub.drop(columns=["date"]).reset_index(drop=True)
      sub_data.columns = range(sub_data.shape[1])
      return sub_data

  d_dash = date_str
  d_nodash = date_str.replace("-", "")
  d_under = date_str.replace("-", "_")
  all_files = glob.glob(os.path.join(OUTPUT_DIR, "*.csv"))
  matched = [
      f
      for f in all_files
      if d_dash in os.path.basename(f)
      or d_nodash in os.path.basename(f)
      or d_under in os.path.basename(f)
  ]
  if matched:
    matched.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    try:
      return pd.read_csv(matched[0], header=None)
    except Exception:
      pass
  return None


def get_available_months():
  months = set()
  display_year = st.session_state.get("display_year", str(datetime.now().year))

  # 1. budget_prev から過去月を自動取得（2026年9月以前に対応）
  try:
    bp_df_all = load_budget_prev_data()
  except RuntimeError:
    bp_df_all = pd.DataFrame()
  if not bp_df_all.empty and "date" in bp_df_all.columns:
    for d in bp_df_all["date"].dropna().astype(str):
      m = re.search(r"(\d{4})-(\d{2})", d)
      if m:
        months.add(f"{m.group(1)}-{m.group(2)}")

  # 2. Airレジシートから取得
  try:
    df_cur_year = get_airregi_sheet_for_year(display_year)
  except RuntimeError:
    df_cur_year = pd.DataFrame()
  if not df_cur_year.empty and "date" in df_cur_year.columns:
    for d in df_cur_year["date"].dropna().astype(str):
      m = re.search(r"(\d{4})-(\d{2})", d)
      if m:
        months.add(f"{m.group(1)}-{m.group(2)}")

  # 3. CSVファイルから取得
  all_files = glob.glob(os.path.join(OUTPUT_DIR, "*.csv"))
  for f in all_files:
    base = os.path.basename(f)
    m = re.search(r"(\d{4})-(\d{2})-\d{2}", base)
    if m:
      months.add(f"{m.group(1)}-{m.group(2)}")

  if not months:
    return [f"{display_year}-10", f"{display_year}-01"]
  return sorted(list(months), reverse=True)


# --- 予約・デリバリー更新関数 ---
def save_delivery_data(
    date_str, store, time_slot, uber, rocket_now, demaekan, other
):
  df = load_delivery_data()
  existing = df[
      (df["date"] == date_str)
      & (df["store"] == store)
      & (df["time_slot"] == time_slot)
  ]
  delivery_count = estimate_delivery_count(uber + rocket_now + demaekan)
  delivery_unit_price = DELIVERY_UNIT_PRICE
  if not existing.empty:
    if "delivery_count" in existing.columns and pd.notna(
        existing["delivery_count"].iloc[0]
    ):
      delivery_count = parse_int(existing["delivery_count"].iloc[0])
    if "delivery_unit_price" in existing.columns and pd.notna(
        existing["delivery_unit_price"].iloc[0]
    ):
      delivery_unit_price = parse_int(
          existing["delivery_unit_price"].iloc[0]
      )
  df = df[
      ~(
          (df["date"] == date_str)
          & (df["store"] == store)
          & (df["time_slot"] == time_slot)
      )
  ]
  new_row = pd.DataFrame([{
      "date": date_str,
      "store": store,
      "time_slot": time_slot,
      "uber": uber,
      "rocket_now": rocket_now,
      "demaekan": demaekan,
      "other": other,
      "delivery_count": delivery_count,
      "delivery_unit_price": delivery_unit_price,
  }])
  df = pd.concat([df, new_row], ignore_index=True)
  conn.update(worksheet="delivery_sales", data=df)
  load_delivery_data.clear()
  clear_monthly_report_cache()


def save_reservation(
    store, date_str, time_slot, mng_name, amount, cust_name, phone, notes
):
  df = load_reservations()
  res_id = str(uuid.uuid4())
  new_row = pd.DataFrame([{
      "id": res_id,
      "store": store,
      "date": date_str,
      "time_slot": time_slot,
      "management_name": str(mng_name),
      "amount": amount,
      "customer_name": cust_name,
      "phone": phone,
      "notes": notes,
  }])
  df = pd.concat([df, new_row], ignore_index=True)
  conn.update(worksheet="reservations", data=df)
  load_reservations.clear()
  clear_monthly_report_cache()


def update_reservation(
    res_id, date_str, time_slot, mng_name, amount, cust_name, phone, notes
):
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
    clear_monthly_report_cache()


def delete_reservation(res_id):
  df = load_reservations()
  df = df[df["id"] != res_id]
  conn.update(worksheet="reservations", data=df)
  load_reservations.clear()
  clear_monthly_report_cache()

  sales_df = load_reservation_sales()
  if not sales_df.empty:
    sales_df = sales_df[sales_df["reservation_id"] != res_id]
    conn.update(worksheet="reservation_sales", data=sales_df)
    load_reservation_sales.clear()
    clear_monthly_report_cache()


def save_reservation_sale(res_id, store, date_str, time_slot, amount):
  df = load_reservation_sales()
  sale_id = str(uuid.uuid4())
  new_row = pd.DataFrame([{
      "id": sale_id,
      "reservation_id": res_id,
      "store": store,
      "date": date_str,
      "time_slot": time_slot,
      "amount": amount,
  }])
  df = pd.concat([df, new_row], ignore_index=True)
  conn.update(worksheet="reservation_sales", data=df)
  load_reservation_sales.clear()
  clear_monthly_report_cache()


def update_reservation_sale(sale_id, date_str, time_slot, amount):
  df = load_reservation_sales()
  idx = df.index[df["id"] == sale_id].tolist()
  if idx:
    df.at[idx[0], "date"] = date_str
    df.at[idx[0], "time_slot"] = time_slot
    df.at[idx[0], "amount"] = amount
    conn.update(worksheet="reservation_sales", data=df)
    load_reservation_sales.clear()
    clear_monthly_report_cache()


def delete_reservation_sale(sale_id):
  df = load_reservation_sales()
  df = df[df["id"] != sale_id]
  conn.update(worksheet="reservation_sales", data=df)
  load_reservation_sales.clear()
  clear_monthly_report_cache()


# --- 表生成関数（過去月報の店頭売上・デリバリー売上反映に対応） ---
def filter_rows_to_month(df, year, month):
  if df.empty or "date" not in df.columns:
    return df
  month_prefix = f"{year:04d}-{month:02d}"
  date_keys = df["date"].map(normalize_date_key)
  return df.loc[date_keys.str.startswith(month_prefix)]


def build_airregi_match_diagnostics(df, selected_store, year, month):
  columns = ["日付", "店舗行数", "合計行数", "合計行の店頭売上"]
  if df.empty or "date" not in df.columns or len(df.columns) < 3:
    return pd.DataFrame(columns=columns)

  c_store = "c0" if "c0" in df.columns else df.columns[1]
  c_slot = "c1" if "c1" in df.columns else df.columns[2]
  c_sales = "c2" if "c2" in df.columns else df.columns[3]
  short_key = get_short_store_key(selected_store)
  date_keys = df["date"].map(normalize_date_key)
  month_prefix = f"{year:04d}-{month:02d}"
  month_rows = df.loc[date_keys.str.startswith(month_prefix)].copy()
  month_rows["_date_key"] = date_keys.loc[month_rows.index]
  store_rows = month_rows[
      month_rows[c_store].astype(str).str.contains(short_key, na=False)
  ]

  diagnostics = []
  _, days_in_month = calendar.monthrange(year, month)
  for day in range(1, days_in_month + 1):
    date_key = f"{year:04d}-{month:02d}-{day:02d}"
    day_rows = store_rows[store_rows["_date_key"] == date_key]
    total_rows = day_rows[
        day_rows[c_slot].astype(str).str.strip() == "合計"
    ]
    sales_values = [
        f"{parse_int(value):,}" for value in total_rows[c_sales].tolist()
    ]
    diagnostics.append({
        "日付": date_key,
        "店舗行数": len(day_rows),
        "合計行数": len(total_rows),
        "合計行の店頭売上": ", ".join(sales_values) if sales_values else "-",
    })
  return pd.DataFrame(diagnostics, columns=columns)


def generate_month_df(
    selected_store,
    year,
    month,
    target_slot_keyword,
    bp_df,
    deliv_df,
    res_df,
    res_sales_df,
    airregi_df,
):
  started_at = time.perf_counter()
  cache = st.session_state["monthly_report_cache"]
  now = time.monotonic()
  cache_key = (selected_store, year, month, target_slot_keyword)
  cached_result = cache.get(cache_key)
  if cached_result:
    cached_at, result = cached_result
    if now - cached_at < 45:
      log_performance_time("月報データ生成（セッションキャッシュ）", started_at)
      return result
    del cache[cache_key]

  result = _build_month_df(
      selected_store, year, month, target_slot_keyword,
      bp_df, deliv_df, res_df, res_sales_df, airregi_df,
  )
  cache[cache_key] = (now, result)
  while len(cache) > 12:
    oldest_key = min(cache, key=lambda key: cache[key][0])
    del cache[oldest_key]
  log_performance_time("月報データ生成（再計算）", started_at)
  return result


def _build_month_df(
    selected_store,
    year,
    month,
    target_slot_keyword,
    bp_df,
    deliv_df,
    res_df,
    res_sales_df,
    airregi_df,
):
  _, max_days = calendar.monthrange(year, month)

  # 1. 予算・前年・過去実績データの事前辞書化
  bp_dict = {}
  if not bp_df.empty:
    bp_sub = bp_df[bp_df["store"] == selected_store]
    if not bp_sub.empty:
      bp_sub = bp_sub.copy()
      bp_sub["_date_key"] = (
          bp_sub["_normalized_date"]
          if "_normalized_date" in bp_sub.columns
          else bp_sub["date"].apply(normalize_date_key)
      )
      bp_sub = bp_sub[bp_sub["_date_key"].str.startswith(
          f"{year:04d}-{month:02d}"
      )]
    if target_slot_keyword and not bp_sub.empty:
      bp_sub = bp_sub[bp_sub["time_slot"] == target_slot_keyword]
    if not bp_sub.empty:
      bp_sub = bp_sub[bp_sub["_date_key"] != ""]

      b_cust_col = None
      for c in ["budget_cust", "予算客数", "b_cust", "budget_customers"]:
        if c in bp_sub.columns:
          b_cust_col = c
          break

      actual_cust_col = "actual_cust" if "actual_cust" in bp_sub.columns else None
      delivery_count_col = (
          "delivery_count" if "delivery_count" in bp_sub.columns else None
      )
      delivery_unit_price_col = (
          "delivery_unit_price"
          if "delivery_unit_price" in bp_sub.columns
          else None
      )

      s_sales_col = None
      for c in [
          "store_sales",
          "店頭売上",
          "店舗売上",
          "本店売上",
          "store_sale",
      ]:
        if c in bp_sub.columns:
          s_sales_col = c
          break

      d_sales_col = None
      for c in [
          "delivery_sales",
          "D売上",
          "デリバリー売上",
          "uber_sales",
          "uber",
      ]:
        if c in bp_sub.columns:
          d_sales_col = c
          break

      for d, grp in bp_sub.groupby("_date_key"):
        b_cust_val = parse_int(grp[b_cust_col].sum()) if b_cust_col else 0
        s_sales_val = parse_int(grp[s_sales_col].sum()) if s_sales_col else 0
        d_sales_val = parse_int(grp[d_sales_col].sum()) if d_sales_col else 0
        actual_cust_values = (
            pd.to_numeric(grp[actual_cust_col], errors="coerce").dropna()
            if actual_cust_col
            else pd.Series(dtype="float64")
        )
        actual_cust_val = (
            int(actual_cust_values.sum())
            if actual_cust_col and len(actual_cust_values) == len(grp)
            else None
        )
        delivery_count_values = (
            pd.to_numeric(grp[delivery_count_col], errors="coerce").dropna()
            if delivery_count_col
            else pd.Series(dtype="float64")
        )
        delivery_count_val = (
            int(delivery_count_values.sum())
            if delivery_count_col and len(delivery_count_values) == len(grp)
            else None
        )
        delivery_unit_price_values = (
            pd.to_numeric(grp[delivery_unit_price_col], errors="coerce").dropna()
            if delivery_unit_price_col
            else pd.Series(dtype="float64")
        )
        delivery_unit_price_val = (
            int(delivery_unit_price_values.iloc[0])
            if delivery_unit_price_col
            and len(delivery_unit_price_values) == len(grp)
            else None
        )
        bp_dict[d] = {
            "prev_sales": parse_int(grp["prev_sales"].sum()),
            "budget_sales": parse_int(grp["budget_sales"].sum()),
            "prev_cust": parse_int(grp["prev_cust"].sum()),
            "budget_cust": b_cust_val,
            "actual_cust": actual_cust_val,
            "delivery_count": delivery_count_val,
            "delivery_unit_price": delivery_unit_price_val,
            "store_sales": s_sales_val,
            "delivery_sales": d_sales_val,
        }

  # 2. デリバリー売上の事前辞書化
  deliv_dict = {}
  if not deliv_df.empty:
    d_sub = filter_rows_to_month(
        deliv_df[deliv_df["store"] == selected_store], year, month
    )
    if target_slot_keyword and not d_sub.empty:
      d_sub = d_sub[d_sub["time_slot"] == target_slot_keyword]
    if not d_sub.empty:
      d_sub = d_sub.copy()
      d_sub["uber_n"] = pd.to_numeric(d_sub["uber"], errors="coerce").fillna(0)
      d_sub["rn_n"] = pd.to_numeric(
          d_sub["rocket_now"], errors="coerce"
      ).fillna(0)
      d_sub["demae_n"] = pd.to_numeric(
          d_sub["demaekan"], errors="coerce"
      ).fillna(0)
      d_sub["other_n"] = pd.to_numeric(d_sub["other"], errors="coerce").fillna(
          0
      )
      if "delivery_count" in d_sub.columns:
        d_sub["delivery_count_n"] = pd.to_numeric(
            d_sub["delivery_count"], errors="coerce"
        )
      else:
        d_sub["delivery_count_n"] = pd.NA
      if "delivery_unit_price" in d_sub.columns:
        d_sub["delivery_unit_price_n"] = pd.to_numeric(
            d_sub["delivery_unit_price"], errors="coerce"
        )
      else:
        d_sub["delivery_unit_price_n"] = pd.NA
      d_sub["total_deliv"] = (
          d_sub["uber_n"] + d_sub["rn_n"] + d_sub["demae_n"]
      )
      for d, grp in d_sub.groupby("date"):
        explicit_counts = grp["delivery_count_n"].dropna()
        explicit_prices = grp["delivery_unit_price_n"].dropna()
        deliv_dict[str(d).strip()] = {
            "deliv": int(grp["total_deliv"].sum()),
            "other": int(grp["other_n"].sum()),
            "delivery_count": (
                int(explicit_counts.sum())
                if len(explicit_counts) == len(grp)
                else None
            ),
            "delivery_unit_price": (
                int(explicit_prices.iloc[0])
                if len(explicit_prices) == len(grp)
                else None
            ),
        }

  # 3. 予約金の事前辞書化
  res_dict = {}
  if not res_df.empty:
    r_sub = filter_rows_to_month(
        res_df[res_df["store"] == selected_store], year, month
    )
    if target_slot_keyword and not r_sub.empty:
      r_sub = r_sub[
          (r_sub["time_slot"] == target_slot_keyword)
          | (r_sub["time_slot"] == "時間帯未指定")
      ]
    if not r_sub.empty:
      r_sub = r_sub.copy()
      r_sub["amount_n"] = pd.to_numeric(r_sub["amount"], errors="coerce").fillna(
          0
      )
      for d, grp in r_sub.groupby("date"):
        res_dict[str(d).strip()] = -int(grp["amount_n"].sum())

  # 4. 売上計上の事前辞書化
  res_sales_dict = {}
  if not res_sales_df.empty:
    rs_sub = filter_rows_to_month(
        res_sales_df[res_sales_df["store"] == selected_store], year, month
    )
    if target_slot_keyword and not rs_sub.empty:
      rs_sub = rs_sub[rs_sub["time_slot"] == target_slot_keyword]
    if not rs_sub.empty:
      rs_sub = rs_sub.copy()
      rs_sub["amount_n"] = pd.to_numeric(
          rs_sub["amount"], errors="coerce"
      ).fillna(0)
      for d, grp in rs_sub.groupby("date"):
        res_sales_dict[str(d).strip()] = int(grp["amount_n"].sum())

  # 5. Airレジデータの事前辞書化
  air_dict = {}
  if (
      airregi_df is not None
      and not airregi_df.empty
      and "date" in airregi_df.columns
  ):
    airregi_month_df = filter_rows_to_month(airregi_df, year, month)
    short_key = get_short_store_key(selected_store)
    c_store = (
        "c0" if "c0" in airregi_month_df.columns else airregi_month_df.columns[1]
    )
    c_slot = (
        "c1" if "c1" in airregi_month_df.columns else airregi_month_df.columns[2]
    )
    c_sales = (
        "c2" if "c2" in airregi_month_df.columns else airregi_month_df.columns[3]
    )
    c_cust = (
        "c5" if "c5" in airregi_month_df.columns else airregi_month_df.columns[6]
    )
    c_other = (
        "c11"
        if "c11" in airregi_month_df.columns
        else (
            airregi_month_df.columns[12]
            if len(airregi_month_df.columns) > 12
            else None
        )
    )

    store_air = airregi_month_df[
        airregi_month_df[c_store].astype(str).str.contains(short_key, na=False)
    ]
    if target_slot_keyword is None:
      matched_air = store_air[
          store_air[c_slot].astype(str).str.strip() == "合計"
      ]
    else:
      matched_air = store_air[
          store_air[c_slot]
          .astype(str)
          .str.contains(target_slot_keyword[0], na=False)
      ]

    for _, row in matched_air.iterrows():
      d_key = normalize_date_key(row["date"])
      if not d_key:
        continue
      a_sales = parse_int(row[c_sales])
      a_cust = parse_int(row[c_cust])
      a_oth = (
          parse_int(row[c_other])
          if c_other and pd.notna(row.get(c_other))
          else 0
      )
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

    # 過去データ判定（2026年9月以前は過去月報の数値を優先反映）
    is_past_archive = date_str < AIRREGI_REPORT_START_DATE

    air_sales, air_cust, air_other_sales = 0, 0, 0
    if date_str in air_dict:
      air_sales, air_cust, air_other_sales = air_dict[date_str]
    else:
      raw_df = find_airregi_data_in_memory(date_str, airregi_df)
      if raw_df is not None:
        try:
          short_key = get_short_store_key(selected_store)
          store_df = raw_df[
              raw_df[0].astype(str).str.contains(short_key, na=False)
          ]
          matched = (
              store_df[store_df[1].astype(str).str.strip() == "合計"]
              if target_slot_keyword is None
              else match_slot_row(store_df, target_slot_keyword)
          )
          if not matched.empty:
            air_sales = parse_int(matched.iloc[0, 2])
            air_cust = parse_int(matched.iloc[0, 5])
            if matched.shape[1] > 11:
              air_other_sales = parse_int(matched.iloc[0, 11])
        except Exception:
          pass

    deliv_item = deliv_dict.get(date_str, {})
    day_res_deposit = res_dict.get(date_str, 0)
    day_res_sales = res_sales_dict.get(date_str, 0)

    # 店頭売上とデリバリー売上の決定
    if is_past_archive and (
        "store_sales" in bp_item or "delivery_sales" in bp_item
    ):
      adjusted_air_sales = bp_item.get("store_sales", 0)
      day_delivery_sales = bp_item.get("delivery_sales", 0)
      total_other_sales = 0
    else:
      adjusted_air_sales = air_sales + day_res_deposit + day_res_sales
      day_delivery_sales = deliv_item.get("deliv", 0)
      total_other_sales = air_other_sales + deliv_item.get("other", 0)

    actual_sales = adjusted_air_sales + day_delivery_sales + total_other_sales
    store_cust = bp_item.get("actual_cust")
    if store_cust is None:
      store_cust = air_cust
    delivery_count = (
        bp_item.get("delivery_count")
        if is_past_archive
        else deliv_item.get("delivery_count")
    )
    if delivery_count is None:
      delivery_count = estimate_delivery_count(day_delivery_sales)
    delivery_unit_price = (
        bp_item.get("delivery_unit_price")
        if is_past_archive
        else deliv_item.get("delivery_unit_price")
    )
    if delivery_unit_price is None:
      delivery_unit_price = DELIVERY_UNIT_PRICE
    actual_cust = store_cust + delivery_count

    # 売上比率
    prev_ratio = (
        f"{(actual_sales / p_sales * 100):.1f}%"
        if p_sales > 0 and actual_sales != 0
        else ("-" if actual_sales == 0 else "0.0%")
    )
    budget_ratio = (
        f"{(actual_sales / b_sales * 100):.1f}%"
        if b_sales > 0 and actual_sales != 0
        else ("-" if actual_sales == 0 else "0.0%")
    )

    # 単価計算
    p_price = int(p_sales / p_cust) if p_cust > 0 else 0
    b_price = int(b_sales / b_cust) if b_cust > 0 else 0
    a_price = int(actual_sales / actual_cust) if actual_cust > 0 else 0

    # 稼働（客数）比率
    cust_prev_ratio = (
        f"{(actual_cust / p_cust * 100):.1f}%"
        if p_cust > 0 and actual_cust > 0
        else "-"
    )
    cust_budget_ratio = (
        f"{(actual_cust / b_cust * 100):.1f}%"
        if b_cust > 0 and actual_cust > 0
        else "-"
    )

    # 客単価比率
    price_prev_ratio = (
        f"{(a_price / p_price * 100):.1f}%"
        if p_price > 0 and a_price != 0
        else "-"
    )
    price_budget_ratio = (
        f"{(a_price / b_price * 100):.1f}%"
        if b_price > 0 and a_price != 0
        else "-"
    )

    weekday_kanji = weekday_kanji_list[pd.to_datetime(date_str).weekday()]

    rows.append({
        "日": day,
        "曜日": weekday_kanji,
        "前年売上": p_sales,
        "予算売上": b_sales,
        "実績売上": actual_sales,
        "前年比": prev_ratio,
        "予算比": budget_ratio,
        "店頭売上": adjusted_air_sales,
        "D売上": day_delivery_sales,
        "その他売上": total_other_sales,
        "予約金": day_res_deposit,
        "売上計上": day_res_sales,
        "前年客数": p_cust,
        "予算客数": b_cust,
        "店頭客数": store_cust,
        "実績客数": actual_cust,
        "D件数": delivery_count,
        "客数前年比": cust_prev_ratio,
        "客数予算比": cust_budget_ratio,
        "前年単価": p_price,
        "予算単価": b_price,
        "実績単価": a_price,
        "D単価": delivery_unit_price,
        "単価前年比": price_prev_ratio,
        "単価予算比": price_budget_ratio,
    })

  return pd.DataFrame(rows)


def calculate_summary_rows(df_slot):
  df_act = df_slot[df_slot["実績売上"] != 0]
  if df_act.empty:
    df_act = df_slot.iloc[:1]
  num_cols = [
      "前年売上",
      "予算売上",
      "実績売上",
      "店頭売上",
      "D売上",
      "D件数",
      "その他売上",
      "予約金",
      "売上計上",
      "前年客数",
      "予算客数",
      "店頭客数",
      "実績客数",
  ]

  # 合計
  sum_vals = {
      col: int(pd.to_numeric(df_act[col], errors="coerce").fillna(0).sum())
      for col in num_cols
  }
  sum_delivery_unit_price = int(
      pd.to_numeric(df_act.loc[df_act["D売上"] != 0, "D単価"], errors="coerce")
      .fillna(0)
      .mean()
  ) if (df_act["D売上"] != 0).any() else 0
  s_p_ratio = (
      f"{(sum_vals['実績売上'] / sum_vals['前年売上'] * 100):.1f}%"
      if sum_vals["前年売上"] > 0
      else "-"
  )
  s_b_ratio = (
      f"{(sum_vals['実績売上'] / sum_vals['予算売上'] * 100):.1f}%"
      if sum_vals["予算売上"] > 0
      else "-"
  )
  s_c_ratio = (
      f"{(sum_vals['実績客数'] / sum_vals['前年客数'] * 100):.1f}%"
      if sum_vals["前年客数"] > 0
      else "-"
  )
  s_cb_ratio = (
      f"{(sum_vals['実績客数'] / sum_vals['予算客数'] * 100):.1f}%"
      if sum_vals["予算客数"] > 0
      else "-"
  )
  s_p_price = (
      int(sum_vals["前年売上"] / sum_vals["前年客数"])
      if sum_vals["前年客数"] > 0
      else 0
  )
  s_b_price = (
      int(sum_vals["予算売上"] / sum_vals["予算客数"])
      if sum_vals["予算客数"] > 0
      else 0
  )
  s_a_price = (
      int(sum_vals["実績売上"] / sum_vals["実績客数"])
      if sum_vals["実績客数"] > 0
      else 0
  )
  s_pr_ratio = (
      f"{(s_a_price / s_p_price * 100):.1f}%" if s_p_price > 0 else "-"
  )
  s_prb_ratio = (
      f"{(s_a_price / s_b_price * 100):.1f}%" if s_b_price > 0 else "-"
  )
  row_sum = {
      "日": "合計",
      "曜日": "",
      **sum_vals,
      "D単価": sum_delivery_unit_price,
      "前年比": s_p_ratio,
      "予算比": s_b_ratio,
      "客数前年比": s_c_ratio,
      "客数予算比": s_cb_ratio,
      "前年単価": s_p_price,
      "予算単価": s_b_price,
      "実績単価": s_a_price,
      "単価前年比": s_pr_ratio,
      "単価予算比": s_prb_ratio,
  }

  # 平均
  ave_vals = {
      col: int(pd.to_numeric(df_act[col], errors="coerce").fillna(0).mean())
      for col in num_cols
  }
  ave_delivery_unit_price = sum_delivery_unit_price
  a_p_ratio = (
      f"{(ave_vals['実績売上'] / ave_vals['前年売上'] * 100):.1f}%"
      if ave_vals["前年売上"] > 0
      else "-"
  )
  a_b_ratio = (
      f"{(ave_vals['実績売上'] / ave_vals['予算売上'] * 100):.1f}%"
      if ave_vals["予算売上"] > 0
      else "-"
  )
  a_c_ratio = (
      f"{(ave_vals['実績客数'] / ave_vals['前年客数'] * 100):.1f}%"
      if ave_vals["前年客数"] > 0
      else "-"
  )
  a_cb_ratio = (
      f"{(ave_vals['実績客数'] / ave_vals['予算客数'] * 100):.1f}%"
      if ave_vals["予算客数"] > 0
      else "-"
  )
  a_p_price = (
      int(ave_vals["前年売上"] / ave_vals["前年客数"])
      if ave_vals["前年客数"] > 0
      else 0
  )
  a_b_price = (
      int(ave_vals["予算売上"] / ave_vals["予算客数"])
      if ave_vals["予算客数"] > 0
      else 0
  )
  a_a_price = (
      int(ave_vals["実績売上"] / ave_vals["実績客数"])
      if ave_vals["実績客数"] > 0
      else 0
  )
  a_pr_ratio = (
      f"{(a_a_price / a_p_price * 100):.1f}%" if a_p_price > 0 else "-"
  )
  a_prb_ratio = (
      f"{(a_a_price / a_b_price * 100):.1f}%" if a_b_price > 0 else "-"
  )
  row_ave = {
      "日": "AVE",
      "曜日": "",
      **ave_vals,
      "D単価": ave_delivery_unit_price,
      "前年比": a_p_ratio,
      "予算比": a_b_ratio,
      "客数前年比": a_c_ratio,
      "客数予算比": a_cb_ratio,
      "前年単価": a_p_price,
      "予算単価": a_b_price,
      "実績単価": a_a_price,
      "単価前年比": a_pr_ratio,
      "単価予算比": a_prb_ratio,
  }

  # 最小
  min_vals = {
      col: int(pd.to_numeric(df_act[col], errors="coerce").fillna(0).min())
      for col in num_cols
  }
  m_p_ratio = (
      f"{(min_vals['実績売上'] / min_vals['前年売上'] * 100):.1f}%"
      if min_vals["前年売上"] > 0
      else "-"
  )
  m_b_ratio = (
      f"{(min_vals['実績売上'] / min_vals['予算売上'] * 100):.1f}%"
      if min_vals["予算売上"] > 0
      else "-"
  )
  m_c_ratio = (
      f"{(min_vals['実績客数'] / min_vals['前年客数'] * 100):.1f}%"
      if min_vals["前年客数"] > 0
      else "-"
  )
  m_cb_ratio = (
      f"{(min_vals['実績客数'] / min_vals['予算客数'] * 100):.1f}%"
      if min_vals["予算客数"] > 0
      else "-"
  )
  m_p_price = int(
      pd.to_numeric(df_act["前年単価"], errors="coerce").fillna(0).min()
  )
  m_b_price = int(
      pd.to_numeric(df_act["予算単価"], errors="coerce").fillna(0).min()
  )
  m_a_price = int(
      pd.to_numeric(df_act["実績単価"], errors="coerce").fillna(0).min()
  )
  m_pr_ratio = (
      f"{(m_a_price / m_p_price * 100):.1f}%" if m_p_price > 0 else "-"
  )
  m_prb_ratio = (
      f"{(m_a_price / m_b_price * 100):.1f}%" if m_b_price > 0 else "-"
  )
  row_min = {
      "日": "MIN",
      "曜日": "",
      **min_vals,
      "D単価": int(pd.to_numeric(df_act["D単価"], errors="coerce").fillna(0).min()),
      "前年比": m_p_ratio,
      "予算比": m_b_ratio,
      "客数前年比": m_c_ratio,
      "客数予算比": m_cb_ratio,
      "前年単価": m_p_price,
      "予算単価": m_b_price,
      "実績単価": m_a_price,
      "単価前年比": m_pr_ratio,
      "単価予算比": m_prb_ratio,
  }

  # 最大
  max_vals = {
      col: int(pd.to_numeric(df_act[col], errors="coerce").fillna(0).max())
      for col in num_cols
  }
  mx_p_ratio = (
      f"{(max_vals['実績売上'] / max_vals['前年売上'] * 100):.1f}%"
      if max_vals["前年売上"] > 0
      else "-"
  )
  mx_b_ratio = (
      f"{(max_vals['実績売上'] / max_vals['予算売上'] * 100):.1f}%"
      if max_vals["予算売上"] > 0
      else "-"
  )
  mx_c_ratio = (
      f"{(max_vals['実績客数'] / max_vals['前年客数'] * 100):.1f}%"
      if max_vals["前年客数"] > 0
      else "-"
  )
  mx_cb_ratio = (
      f"{(max_vals['実績客数'] / max_vals['予算客数'] * 100):.1f}%"
      if max_vals["予算客数"] > 0
      else "-"
  )
  mx_p_price = int(
      pd.to_numeric(df_act["前年単価"], errors="coerce").fillna(0).max()
  )
  mx_b_price = int(
      pd.to_numeric(df_act["予算単価"], errors="coerce").fillna(0).max()
  )
  mx_a_price = int(
      pd.to_numeric(df_act["実績単価"], errors="coerce").fillna(0).max()
  )
  mx_pr_ratio = (
      f"{(mx_a_price / mx_p_price * 100):.1f}%" if mx_p_price > 0 else "-"
  )
  mx_prb_ratio = (
      f"{(mx_a_price / mx_b_price * 100):.1f}%" if mx_b_price > 0 else "-"
  )
  row_max = {
      "日": "MAX",
      "曜日": "",
      **max_vals,
      "D単価": int(pd.to_numeric(df_act["D単価"], errors="coerce").fillna(0).max()),
      "前年比": mx_p_ratio,
      "予算比": mx_b_ratio,
      "客数前年比": mx_c_ratio,
      "客数予算比": mx_cb_ratio,
      "前年単価": mx_p_price,
      "予算単価": mx_b_price,
      "実績単価": mx_a_price,
      "単価前年比": mx_pr_ratio,
      "単価予算比": mx_prb_ratio,
  }

  return [row_sum, row_ave, row_min, row_max]


def create_multi_sheet_excel(
    selected_store,
    year,
    month,
    bp_df,
    deliv_df,
    res_df,
    res_sales_df,
    airregi_df,
):
  output = io.BytesIO()
  slots = [("合計", None), ("早番", "早番"), ("中番", "中番"), ("遅番", "遅番")]

  header_fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  header_font = Font(name="游ゴシック", size=11, bold=True, color="FFFFFF")
  total_row_fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  total_row_font = Font(name="游ゴシック", size=11, bold=True, color="FFFFFF")
  even_row_fill = PatternFill(
      start_color="F2F4F7", end_color="F2F4F7", fill_type="solid"
  )
  odd_row_fill = PatternFill(
      start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
  )
  data_font = Font(name="游ゴシック", size=10)
  summary_label_font = Font(name="游ゴシック", size=10, bold=True)
  center_alignment = Alignment(horizontal="center", vertical="center")
  thin_border = Border(
      left=Side(style="thin", color="D9D9D9"),
      right=Side(style="thin", color="D9D9D9"),
      top=Side(style="thin", color="D9D9D9"),
      bottom=Side(style="thin", color="D9D9D9"),
  )

  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    workbook = writer.book
    for sheet_name, slot_kw in slots:
      df_slot = generate_month_df(
          selected_store,
          year,
          month,
          slot_kw,
          bp_df,
          deliv_df,
          res_df,
          res_sales_df,
          airregi_df,
      )
      cols = list(df_slot.columns)
      ws = workbook.create_sheet(title=sheet_name)

      ws.cell(
          row=1,
          column=1,
          value=f"{selected_store} {year}年{month}月度（{sheet_name}）",
      ).font = Font(
          name="游ゴシック", size=13, bold=True, color="1F4E78"
      )

      for col_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=3, column=col_idx)
        cell.value = (
            "合計" if col_idx == 1 else ("" if col_idx == 2 else col_name)
        )
        cell.fill, cell.font, cell.alignment, cell.border = (
            header_fill,
            header_font,
            center_alignment,
            thin_border,
        )

      for s_idx, s_row in enumerate(calculate_summary_rows(df_slot), start=4):
        is_total_row = s_row.get("日") == "合計"
        fill = (
            total_row_fill
            if is_total_row
            else (even_row_fill if s_idx % 2 == 0 else odd_row_fill)
        )
        font = total_row_font if is_total_row else data_font

        for col_idx, col_name in enumerate(cols, start=1):
          cell = ws.cell(row=s_idx, column=col_idx)
          val = s_row.get(col_name, "")
          if col_idx == 1:
            cell.value = "" if is_total_row else val
            cell.font = total_row_font if is_total_row else summary_label_font
          elif col_idx == 2:
            cell.value = ""
          else:
            cell.value = val
            cell.font = font
          cell.fill, cell.alignment, cell.border = (
              fill,
              center_alignment,
              thin_border,
          )

      ws.merge_cells("A3:B4")
      ws.merge_cells("A5:B5")
      ws.merge_cells("A6:B6")
      ws.merge_cells("A7:B7")

      for col_idx, col_name in enumerate(cols, start=1):
        cell = ws.cell(row=9, column=col_idx, value=col_name)
        cell.fill, cell.font, cell.alignment, cell.border = (
            header_fill,
            header_font,
            center_alignment,
            thin_border,
        )

      for r_idx, row_data in enumerate(
          df_slot.to_dict(orient="records"), start=10
      ):
        fill = even_row_fill if r_idx % 2 == 0 else odd_row_fill
        for col_idx, col_name in enumerate(cols, start=1):
          cell = ws.cell(row=r_idx, column=col_idx, value=row_data[col_name])
          cell.fill, cell.font, cell.alignment, cell.border = (
              fill,
              data_font,
              center_alignment,
              thin_border,
          )

      ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 8, 8
      for col in list(ws.columns)[2:]:
        max_len = max(len(str(cell.value or "")) for cell in list(col)[2:])
        ws.column_dimensions[col[0].column_letter].width = max(max_len + 4, 12)

    if "Sheet" in workbook.sheetnames:
      workbook.remove(workbook["Sheet"])
  return output.getvalue()


current_year_num = datetime.now().year
available_years = [
    str(y) for y in range(current_year_num, current_year_num - 11, -1)
]

if "selected_store" not in st.session_state:
  st.session_state["selected_store"] = STORES[0]
if "display_year" not in st.session_state:
  st.session_state["display_year"] = str(current_year_num)
if "target_date_state" not in st.session_state:
  st.session_state["target_date_state"] = datetime.now().date()

if "form_d_key" not in st.session_state:
  st.session_state["form_d_key"] = 0
if "form_r_key" not in st.session_state:
  st.session_state["form_r_key"] = 0
if "form_s_key" not in st.session_state:
  st.session_state["form_s_key"] = 0


def prev_year_action():
  idx = (
      available_years.index(st.session_state["display_year"])
      if st.session_state["display_year"] in available_years
      else 0
  )
  if idx < len(available_years) - 1:
    st.session_state["display_year"] = available_years[idx + 1]
    m = get_available_months()
    if m:
      st.session_state["selected_year_month"] = m[0]


def next_year_action():
  idx = (
      available_years.index(st.session_state["display_year"])
      if st.session_state["display_year"] in available_years
      else 0
  )
  if idx > 0:
    st.session_state["display_year"] = available_years[idx - 1]
    m = get_available_months()
    if m:
      st.session_state["selected_year_month"] = m[0]


def run_airregi_extraction(start_date_str, end_date_str):
  subprocess.run(
      [sys.executable, "-m", "playwright", "install", "chromium"],
      capture_output=True,
      check=True,
  )

  custom_env = os.environ.copy()
  air_id = st.secrets.get("AIRREGI_ID")
  air_pass = st.secrets.get("AIRREGI_PASS")

  if (
      not air_id
      and "connections" in st.secrets
      and "gsheets" in st.secrets["connections"]
  ):
    gsheets_sec = st.secrets["connections"]["gsheets"]
    air_id = gsheets_sec.get("AIRREGI_ID")
    air_pass = gsheets_sec.get("AIRREGI_PASS")

  if air_id:
    custom_env["AIRREGI_ID"] = str(air_id)
  if air_pass:
    custom_env["AIRREGI_PASS"] = str(air_pass)

  command = [
      sys.executable,
      "-u",
      "get_monthly_data.py",
      start_date_str,
      end_date_str,
  ]
  start_dt = datetime.strptime(start_date_str, "%Y-%m-%d").date()
  end_dt = datetime.strptime(end_date_str, "%Y-%m-%d").date()
  total_days = (end_dt - start_dt).days + 1
  completed_days = 0
  output_lines = []
  progress = st.progress(0, text=f"取得準備中（0/{total_days}日）")

  with subprocess.Popen(
      command,
      stdout=subprocess.PIPE,
      stderr=subprocess.STDOUT,
      text=True,
      encoding="utf-8",
      errors="replace",
      bufsize=1,
      env=custom_env,
  ) as process:
    if process.stdout is None:
      raise RuntimeError("データ取得プロセスの出力を読み取れません。")
    for line in process.stdout:
      output_lines.append(line)
      match = re.search(r"処理日】\s*(\d{4}-\d{2}-\d{2})", line)
      if match:
        completed_days += 1
        progress.progress(
            min(completed_days / total_days, 1.0),
            text=f"取得中: {match.group(1)}（{completed_days}/{total_days}日）",
        )
    return_code = process.wait()

  output = "".join(output_lines)
  if return_code:
    progress.empty()
    raise subprocess.CalledProcessError(
        return_code,
        command,
        output=output,
    )

  progress.progress(1.0, text=f"取得完了（{total_days}/{total_days}日）")
  return subprocess.CompletedProcess(command, return_code, stdout=output)


def show_airregi_extraction_error(error):
  st.error("⚠️ データ取得処理でエラーが発生しました。")
  err_detail = ""
  if error.stderr:
    err_detail += f"【標準エラー】:\n{error.stderr}\n\n"
  if error.stdout:
    err_detail += f"【標準出力】:\n{error.stdout}"
  st.code(
      err_detail if err_detail else "エラーメッセージはありませんでした。",
      language="text",
  )
  if os.path.exists("error_screenshot.png"):
    st.image("error_screenshot.png", caption="エラー発生時のブラウザ画面")


if APP_VARIANT in {"admin", "combined"}:
  st.sidebar.title("メニュー")
  st.sidebar.markdown("---")
  st.sidebar.subheader("店舗・期間設定")
  st.sidebar.selectbox("閲覧店舗", STORES, key="selected_store")

if APP_VARIANT in {"admin", "combined"}:
  display_year = st.session_state["display_year"]
  available_months = get_available_months()
  months_in_year = [
      m for m in available_months if m.startswith(display_year)
  ] or [f"{display_year}-10", f"{display_year}-01"]
  st.sidebar.markdown("対象年")
  if APP_VARIANT == "admin":
    with st.sidebar.container(key="admin_year_navigation"):
      year_columns = st.columns([1, 2, 1], gap="small")
  else:
    year_columns = st.sidebar.columns([1, 2, 1])
  col_prev, col_year, col_next = year_columns
  with col_prev:
    st.button("◀", key="prev_year", on_click=prev_year_action)
  with col_year:
    st.markdown(
        f"<div style='text-align: center; padding-top: 5px; font-weight: bold;"
        f" font-size: 18px;'>{display_year}年</div>",
        unsafe_allow_html=True,
    )
  with col_next:
    st.button("▶", key="next_year", on_click=next_year_action)
  if st.session_state.get("selected_year_month") not in months_in_year:
    st.session_state["selected_year_month"] = months_in_year[0]
  st.sidebar.selectbox("対象月", months_in_year, key="selected_year_month")

  if APP_VARIANT == "combined":
    if st.sidebar.button(
        "🔄 最新データを再読み込み",
        key="refresh_monthly_report_data",
        help="店舗アプリで登録したデータを含め、スプレッドシートから最新情報を読み込みます。",
    ):
      refresh_monthly_report_data()
      st.rerun()

  selected_year_month = st.session_state["selected_year_month"]
else:
  selected_year_month = datetime.now().strftime("%Y-%m")

if APP_VARIANT == "combined":
  with st.sidebar.expander("📥 売上データを取得する"):
    target_date = st.date_input(
        "取得日",
        value=datetime.now().date(),
        key="fetch_date_input_sidebar",
    )
    if st.button("🚀 取得実行", type="primary", key="fetch_single_date"):
      target_str = target_date.strftime("%Y-%m-%d")
      with st.spinner(f"{target_str} 取得中... (完了までお待ちください)"):
        try:
          run_airregi_extraction(target_str, target_str)
          count = save_airregi_csv_to_sheets(target_str)
          st.success(
              f"🎉 {target_str} のデータを取得し、"
              f"スプレッドシートへ {count} 日分を保存しました！"
          )
          st.rerun()
        except subprocess.CalledProcessError as error:
          show_airregi_extraction_error(error)
        except Exception as error:
          st.error(
              "データの取得またはスプレッドシートへの保存に失敗しました: "
              f"{error}"
          )

    st.markdown("#### 期間を指定して取得")
    date_col_start, date_col_end = st.columns(2)
    with date_col_start:
      start_date = st.date_input(
          "開始日",
          value=datetime(2025, 10, 1).date(),
          key="fetch_start_date_sidebar",
      )
    with date_col_end:
      end_date = st.date_input(
          "終了日",
          value=datetime(2026, 9, 30).date(),
          key="fetch_end_date_sidebar",
      )
    st.caption(
        "期間取得は時間がかかります。処理中は画面を閉じずにお待ちください。"
    )
    if st.button("📅 期間取得を実行", key="fetch_date_range_button"):
      if start_date > end_date:
        st.error("開始日は終了日以前の日付を指定してください。")
      else:
        start_str = start_date.strftime("%Y-%m-%d")
        end_str = end_date.strftime("%Y-%m-%d")
        date_count = (end_date - start_date).days + 1
        with st.spinner(
            f"{start_str} 〜 {end_str} の {date_count} 日分を取得中です。"
            "完了まで画面を閉じずにお待ちください..."
        ):
          try:
            run_airregi_extraction(start_str, end_str)
            count = save_airregi_csv_to_sheets(
                start_date_str=start_str,
                end_date_str=end_str,
            )
            st.success(
                "🎉 期間取得とスプレッドシートへの保存が完了しました。"
                f"（対象 {date_count} 日、保存 {count} 日）"
            )
            st.rerun()
          except subprocess.CalledProcessError as error:
            st.error(
                f"⚠️ 期間取得に失敗しました。取得済みCSVは"
                f" {start_str} 〜 {end_str} で同期します。"
            )
            try:
              count = save_airregi_csv_to_sheets(
                  start_date_str=start_str,
                  end_date_str=end_str,
              )
              st.info(f"失敗前に取得できた {count} 日分を同期しました。")
            except Exception as sync_error:
              st.error(f"取得済みCSVの同期にも失敗しました: {sync_error}")
            show_airregi_extraction_error(error)
          except Exception as error:
            st.error(
                "期間データのスプレッドシートへの保存に失敗しました。"
                f"取得済みCSVはローカルに残っています: {error}"
            )

    st.markdown("#### 店頭客数CSVを反映")
    customer_csv = st.file_uploader(
        "店頭客数CSV（date, store, time_slot, actual_cust）",
        type=["csv"],
        key="actual_customer_csv_upload",
        help=(
            "日付・店舗・時間帯がbudget_prevシートの全行と一致するCSVのみ反映します。"
            "売上や予算など他の列は変更しません。"
        ),
    )
    if customer_csv and st.button(
        "店頭客数・D件数・D単価をスプレッドシートへ反映",
        key="update_actual_customer_csv_button",
    ):
      try:
        updated_count, blank_count, estimated_d_count = (
            update_actual_customer_column(customer_csv)
        )
        st.success(
            f"店頭客数を {updated_count - blank_count} 行反映しました。"
            f"空欄 {blank_count} 行は空欄のまま、"
            f"D件数 {estimated_d_count:,} 件を反映しました。"
        )
        st.rerun()
      except Exception as error:
        st.error(f"店頭客数CSVを反映できませんでした: {error}")

    if st.button(
        "☁️ 取得済みCSVを年別シートに全同期",
        key="sync_all_airregi_csv",
        help="サーバー上にある全CSVを年ごとのシートに自動で振り分けて保存します。",
    ):
      with st.spinner("スプレッドシートへ同期中..."):
        try:
          count = save_airregi_csv_to_sheets()
          if count > 0:
            st.success(
                f"🎉 サーバー上の {count} 日分の売上データを"
                "スプレッドシートに同期しました！"
            )
            st.rerun()
          else:
            st.info("同期対象のCSVファイルがありませんでした。")
        except Exception as error:
          st.error(f"スプレッドシートへの同期に失敗しました: {error}")

if APP_VARIANT == "combined":
  app_mode = st.sidebar.radio(
      "機能を選択してください",
      [
          APP_MODE_BY_VARIANT["admin"],
          APP_MODE_BY_VARIANT["store"],
          "📋 エクセルコピペ用データ生成",
          "📄 報告書作成",
      ],
  )
else:
  app_mode = APP_MODE_BY_VARIANT[APP_VARIANT]

selected_store = st.session_state["selected_store"]
year, month = map(int, selected_year_month.split("-"))

# 共通データの取得
monthly_data_setup_started_at = time.perf_counter()
try:
  bp_df = timed_call("予算・前年データ読み込み", load_budget_prev_data)
except RuntimeError as error:
  if APP_VARIANT in {"admin", "combined"}:
    st.error(f"月報の予算・前年データを読み込めません: {error}")
    st.stop()
  st.warning(f"予算・前年データは利用できません: {error}")
  bp_df = pd.DataFrame()

if APP_VARIANT in {"admin", "combined"}:
  if bp_df.empty:
    st.warning(
        "budget_prevシートにデータがありません。"
        "本部用アプリのSecretsに、予算・前年データがあるGoogleスプレッドシートの"
        "接続設定が登録されているか確認してください。"
    )
  elif not {
      "date",
      "store",
      "time_slot",
      "budget_sales",
      "prev_sales",
  }.issubset(bp_df.columns):
    st.warning(
        "budget_prevシートの列名が想定と異なります。"
        "date、store、time_slot、budget_sales、prev_sales列を確認してください。"
    )
  elif {"date", "store"}.issubset(bp_df.columns):
    matching_budget_rows = bp_df[
        (bp_df["store"].astype(str).str.strip() == selected_store)
        & bp_df["_normalized_date"].str.startswith(
            f"{year:04d}-{month:02d}"
        )
    ]
    if matching_budget_rows.empty:
      st.warning(
          f"budget_prevシートに{selected_store}の{year}年{month}月の"
          "行が見つかりません。接続先のシートと日付・店舗名を確認してください。"
      )

deliv_df = timed_call("デリバリーデータ読み込み", load_delivery_data)
res_df = timed_call("予約データ読み込み", load_reservations)
res_sales_df = timed_call("予約売上データ読み込み", load_reservation_sales)
needs_monthly_report_data = (
    APP_VARIANT == "admin"
    or app_mode == APP_MODE_BY_VARIANT["admin"]
)
uses_airregi_monthly_data = (
    f"{year:04d}-{month:02d}-01" >= AIRREGI_REPORT_START_DATE
)
airregi_year_df = pd.DataFrame()
if APP_VARIANT in {"admin", "combined"} and uses_airregi_monthly_data:
  try:
    airregi_year_df = timed_call(
        "Airレジデータ読み込み", get_airregi_sheet_for_year, str(year)
    )
  except RuntimeError as error:
    if needs_monthly_report_data and uses_airregi_monthly_data:
      st.error(f"本部月報のAirレジデータを読み込めません: {error}")
      st.stop()
    st.warning(f"Airレジデータを読み込めません: {error}")

if needs_monthly_report_data and uses_airregi_monthly_data:
  if airregi_year_df.empty:
    st.warning(
        f"Airレジの{year}年シートにデータがありません。"
        "本部アプリのSecretsに設定したAIRREGI_SPREADSHEETが、"
        "店舗アプリと同じスプレッドシートを指しているか確認してください。"
    )
  elif {"date", "c0"}.issubset(airregi_year_df.columns):
    month_prefix = f"{year:04d}-{month:02d}"
    month_date_keys = airregi_year_df["date"].map(normalize_date_key)
    month_rows = airregi_year_df[month_date_keys.str.startswith(month_prefix)]
    short_store_key = get_short_store_key(selected_store)
    store_month_rows = month_rows[
        month_rows["c0"].astype(str).str.contains(short_store_key, na=False)
    ]
    if store_month_rows.empty:
      st.warning(
          f"Airレジの{year}年シートに{selected_store}の{month}月データがありません。"
          "Airレジ年別シートの同期状況と、本部アプリの"
          "AIRREGI_SPREADSHEET設定を確認してください。"
      )
log_performance_time("月報データ準備（読み込み・検証を含む）", monthly_data_setup_started_at)

if app_mode == "📱 店舗用（実績入力）":
  st.markdown(
      """
        <style>
        .stTabs [data-baseweb="tab-list"] { gap: 10px; }
        .stTabs [data-baseweb="tab"] {
            background-color: #f8f9fa !important;
            border: 1px solid #ddd !important;
            border-radius: 5px !important;
            padding: 10px 20px !important;
            margin: 0 !important;
        }
        .stTabs [data-baseweb="tab"][aria-selected="true"] {
            background-color: #ef4444 !important;
            border-color: #ef4444 !important;
        }
        .stTabs [data-baseweb="tab"] * { color: #333 !important; font-weight: bold; }
        .stTabs [data-baseweb="tab"][aria-selected="true"] * { color: #fff !important; }
        .stTabs [data-baseweb="tab-highlight"] { display: none; }
        </style>
    """,
      unsafe_allow_html=True,
  )

  st.title("📱 実績・予約管理")
  st.selectbox("入力店舗", STORES, key="selected_store")
  selected_store = st.session_state["selected_store"]

  tab_input, tab_manage = st.tabs(
      ["📝 日次入力 (デリバリー売上 / 予約)", "⚙️ 予約データの確認・修正・削除"]
  )

  with tab_input:
    st.markdown(
        """
            <div style="background-color: #e3f2fd; padding: 10px 20px; border-radius: 5px; border-left: 5px solid #1976d2; margin-bottom: 15px;">
                <h3 style="margin: 0; color: #0d47a1;">🎯 入力対象の日付と時間帯</h3>
            </div>
        """,
        unsafe_allow_html=True,
    )
    col_t1, col_t2 = st.columns(2)
    with col_t1:
      selected_date_val = st.date_input(
          "対象日付",
          value=st.session_state["target_date_state"],
          key="daily_date_input",
      )
    with col_t2:
      selected_slot = st.selectbox(
          "時間帯を選択", TIME_SLOTS, key="daily_slot"
      )
    if selected_date_val != st.session_state["target_date_state"]:
      st.session_state["target_date_state"] = selected_date_val
    selected_date = selected_date_val.strftime("%Y-%m-%d")
    st.markdown("---")

    st.subheader(
        f"🛵 1. デリバリー売上 【対象: {selected_date} / {selected_slot}】"
    )
    just_saved_deliv = False
    if "msg_deliv" in st.session_state:
      st.success(st.session_state["msg_deliv"])
      del st.session_state["msg_deliv"]
      just_saved_deliv = True

    existing_row = deliv_df[
        (deliv_df["date"] == selected_date)
        & (deliv_df["store"] == selected_store)
        & (deliv_df["time_slot"] == selected_slot)
    ]

    d_uber, d_rn, d_demae, d_other = 0, 0, 0, 0
    if not existing_row.empty:
      d_uber = (
          parse_int(existing_row["uber"].values[0])
          if "uber" in existing_row.columns
          else 0
      )
      d_rn = (
          parse_int(existing_row["rocket_now"].values[0])
          if "rocket_now" in existing_row.columns
          else 0
      )
      d_demae = (
          parse_int(existing_row["demaekan"].values[0])
          if "demaekan" in existing_row.columns
          else 0
      )
      d_other = (
          parse_int(existing_row["other"].values[0])
          if "other" in existing_row.columns
          else 0
      )
      if not just_saved_deliv:
        st.info(
            f"ℹ️ {selected_date} / {selected_slot}"
            " のデータは登録済みです。上書き保存されます。"
        )

    css_input_rules = []
    if not existing_row.empty:
      c_uber = "#c0392b" if d_uber < 0 else "#000000"
      c_rn = "#c0392b" if d_rn < 0 else "#000000"
      c_demae = "#c0392b" if d_demae < 0 else "#000000"
      c_other = "#c0392b" if d_other < 0 else "#000000"
      css_input_rules.append(
          'input[aria-label="Uber Eats 売上 (円)"] { color:'
          f" {c_uber} !important; font-weight: bold !important; font-size:"
          " 16px !important; }"
      )
      css_input_rules.append(
          'input[aria-label="Rocket now 売上 (円)"] { color:'
          f" {c_rn} !important; font-weight: bold !important; font-size: 16px"
          " !important; }"
      )
      css_input_rules.append(
          'input[aria-label="出前館 売上 (円)"] { color:'
          f" {c_demae} !important; font-weight: bold !important; font-size:"
          " 16px !important; }"
      )
      css_input_rules.append(
          'input[aria-label="その他 売上 (円)"] { color:'
          f" {c_other} !important; font-weight: bold !important; font-size:"
          " 16px !important; }"
      )
    if css_input_rules:
      st.markdown(
          "<style>" + "\n".join(css_input_rules) + "</style>",
          unsafe_allow_html=True,
      )

    k_d_uber = f"d_uber_{selected_date}_{selected_slot}_{st.session_state['form_d_key']}"
    k_d_rn = (
        f"d_rn_{selected_date}_{selected_slot}_{st.session_state['form_d_key']}"
    )
    k_d_dem = f"d_dem_{selected_date}_{selected_slot}_{st.session_state['form_d_key']}"
    k_d_oth = f"d_oth_{selected_date}_{selected_slot}_{st.session_state['form_d_key']}"

    c1, c2 = st.columns(2)
    with c1:
      uber_val = st.number_input(
          "Uber Eats 売上 (円)", value=d_uber, step=500, key=k_d_uber
      )
      rn_val = st.number_input(
          "Rocket now 売上 (円)", value=d_rn, step=500, key=k_d_rn
      )
    with c2:
      demaekan_val = st.number_input(
          "出前館 売上 (円)", value=d_demae, step=500, key=k_d_dem
      )
      other_val = st.number_input(
          "その他 売上 (円)", value=d_other, step=500, key=k_d_oth
      )

    st.markdown("---")
    total_delivery = uber_val + rn_val + demaekan_val + other_val
    if total_delivery < 0:
      st.markdown(
          f"💡 この時間帯の売上合計: <span style='color: #c0392b; font-weight:"
          f" bold; font-size: 1.15em;'>¥{total_delivery:,}</span>",
          unsafe_allow_html=True,
      )
    else:
      st.markdown(f"💡 この時間帯の売上合計: **¥{total_delivery:,}**")

    if st.button("💾 デリバリー売上を保存", type="primary"):
      save_delivery_data(
          selected_date,
          selected_store,
          selected_slot,
          uber_val,
          rn_val,
          demaekan_val,
          other_val,
      )
      st.session_state["msg_deliv"] = (
          f"🎉 【{selected_date} /"
          f" {selected_slot}】のデリバリー売上を保存しました！"
      )
      st.session_state["form_d_key"] += 1
      st.rerun()

    st.markdown("---")

    st.subheader(
        "📈 2. 予約売上計上（商品引渡）【計上日:"
        f" {selected_date} / {selected_slot}】"
    )
    if "msg_res_sales" in st.session_state:
      st.success(st.session_state["msg_res_sales"])
      del st.session_state["msg_res_sales"]

    st.info(
        "過去に預かった予約金から、本日分の売上を計上します。ここで登録した金額は月報の店頭売上にプラス処理されます。"
    )
    store_reservations = res_df[res_df["store"] == selected_store].copy()

    active_reservations = []
    for _, r in store_reservations.iterrows():
      t_sales = res_sales_df[res_sales_df["reservation_id"] == r["id"]][
          "amount"
      ].sum()
      if (r["amount"] - t_sales) > 0:
        active_reservations.append(r)
    active_store_reservations = pd.DataFrame(active_reservations)

    if active_store_reservations.empty:
      st.warning(
          "現在、この店舗に計上可能な予約入金（未完了の予約）はありません。"
      )
    else:
      mng_list = active_store_reservations["management_name"].tolist()
      k_s_mng = f"s_mng_{selected_date}_{selected_slot}_{st.session_state['form_s_key']}"
      k_s_amt = f"s_amt_{selected_date}_{selected_slot}_{st.session_state['form_s_key']}"

      selected_mng = st.selectbox("対象の予約を選択", mng_list, key=k_s_mng)
      sel_res = active_store_reservations[
          active_store_reservations["management_name"] == selected_mng
      ].iloc[0]

      total_deposit = sel_res["amount"]
      total_sales = res_sales_df[
          res_sales_df["reservation_id"] == sel_res["id"]
      ]["amount"].sum()
      balance = total_deposit - total_sales

      st.markdown(
          f"**預かり金額**: ¥{total_deposit:,} | **計上済**: ¥{total_sales:,} |"
          f" **残額**: <span style='color:red;"
          f" font-size:1.2em;'>¥{balance:,}</span>",
          unsafe_allow_html=True,
      )
      s_amount = st.number_input(
          "本日の計上金額 (円)",
          min_value=0,
          max_value=int(balance),
          step=500,
          value=0,
          key=k_s_amt,
      )

      if st.button("📈 売上を計上する", type="primary"):
        if s_amount > 0:
          save_reservation_sale(
              sel_res["id"],
              selected_store,
              selected_date,
              selected_slot,
              s_amount,
          )
          st.session_state["msg_res_sales"] = (
              f"🎉 【{selected_date} / {selected_slot}】に「{selected_mng}」の売上"
              f" (¥{s_amount:,}) を計上しました！"
          )
          st.session_state["form_s_key"] += 1
          st.rerun()
        else:
          st.error("計上金額を正しく入力してください。")

    st.markdown("---")

    st.subheader(f"📥 3. 予約入金（事前預かり）【入金日: {selected_date}】")
    if "msg_res_dep" in st.session_state:
      st.success(st.session_state["msg_res_dep"])
      del st.session_state["msg_res_dep"]

    st.info(
        "Airレジに入金として打たれた予約預かり金を登録します。ここで登録した金額は月報でマイナス処理（相殺）されます。"
    )
    k_mng = (
        f"r_mng_{selected_date}_{selected_slot}_{st.session_state['form_r_key']}"
    )
    k_amt = (
        f"r_amt_{selected_date}_{selected_slot}_{st.session_state['form_r_key']}"
    )
    k_cust = (
        f"r_cust_{selected_date}_{selected_slot}_{st.session_state['form_r_key']}"
    )
    k_phone = f"r_phone_{selected_date}_{selected_slot}_{st.session_state['form_r_key']}"
    k_notes = f"r_notes_{selected_date}_{selected_slot}_{st.session_state['form_r_key']}"
    k_r_slot = f"r_slot_{selected_date}_{selected_slot}_{st.session_state['form_r_key']}"

    r_slot = st.selectbox(
        "Airレジ入力時間帯",
        ["時間帯未指定"] + TIME_SLOTS,
        index=(
            TIME_SLOTS.index(selected_slot) + 1
            if selected_slot in TIME_SLOTS
            else 0
        ),
        key=k_r_slot,
        help="未指定の場合は月報の「合計」のみで相殺されます",
    )
    c3, c4 = st.columns(2)
    with c3:
      r_mng_name = st.text_input("管理名 (例: 高山様 10月平日分)", key=k_mng)
    with c4:
      r_amount = st.number_input(
          "預かり金額 (円)", min_value=0, step=1000, value=0, key=k_amt
      )
    c5, c6 = st.columns(2)
    with c5:
      r_cust_name = st.text_input("顧客名", key=k_cust)
    with c6:
      r_phone = st.text_input("電話番号", key=k_phone)
    r_notes = st.text_area("備考", key=k_notes)

    if st.button("📥 入金を登録する", type="primary"):
      if r_mng_name and r_amount > 0:
        save_reservation(
            selected_store,
            selected_date,
            r_slot,
            r_mng_name,
            r_amount,
            r_cust_name,
            r_phone,
            r_notes,
        )
        st.session_state["msg_res_dep"] = (
            f"🎉 【{selected_date}】に「{r_mng_name}」の入金 (¥{r_amount:,})"
            " を登録しました！"
        )
        st.session_state["form_r_key"] += 1
        st.rerun()
      else:
        st.error("管理名と預かり金額を入力してください。")

  with tab_manage:
    if "flash_msg_manage" in st.session_state:
      st.success(st.session_state["flash_msg_manage"])
      del st.session_state["flash_msg_manage"]

    manage_mode = st.radio(
        "操作を選択",
        ["履歴一覧", "入金データの修正・削除", "売上計上データの修正・削除"],
        horizontal=True,
    )
    store_reservations = res_df[res_df["store"] == selected_store].copy()

    if manage_mode == "履歴一覧":
      if store_reservations.empty:
        st.write("データがありません。")
      else:
        filter_opt = st.radio(
            "表示フィルター",
            ["すべて表示", "未完了（残額あり）のみ"],
            horizontal=True,
        )

        events = []
        for _, r in store_reservations.iterrows():
          res_id = r["id"]
          mng_name = r["management_name"]

          events.append({
              "res_id": res_id,
              "管理名": mng_name,
              "実行日": r["date"],
              "時間帯": (
                  r["time_slot"] if r["time_slot"] != "時間帯未指定" else ""
              ),
              "預かり金額": r["amount"],
              "計上売上": None,
              "顧客名": (
                  r["customer_name"] if pd.notna(r["customer_name"]) else ""
              ),
              "電話番号": r["phone"] if pd.notna(r["phone"]) else "",
              "備考": r["notes"] if pd.notna(r["notes"]) else "",
              "type": "deposit",
          })

          sales = res_sales_df[res_sales_df["reservation_id"] == res_id]
          for _, s in sales.iterrows():
            events.append({
                "res_id": res_id,
                "管理名": mng_name,
                "実行日": s["date"],
                "時間帯": s["time_slot"],
                "預かり金額": None,
                "計上売上": s["amount"],
                "顧客名": "",
                "電話番号": "",
                "備考": "",
                "type": "sale",
            })

        grouped_events = {}
        for e in events:
          grouped_events.setdefault(e["res_id"], []).append(e)

        final_rows = []
        for res_id, evs in grouped_events.items():
          evs.sort(
              key=lambda x: (x["実行日"], 0 if x["type"] == "deposit" else 1)
          )
          balance = 0
          for e in evs:
            if e["type"] == "deposit":
              balance += e["預かり金額"]
            else:
              balance -= e["計上売上"]
            e["残額"] = balance

          if filter_opt == "未完了（残額あり）のみ" and balance <= 0:
            continue

          final_rows.extend(evs)

        final_rows.sort(
            key=lambda x: (x["実行日"], 0 if x["type"] == "deposit" else 1)
        )

        display_data = []
        for r in final_rows:
          d_parts = r["実行日"].replace("-", "/").split("/")
          d_str = (
              f"{d_parts[0]}/{int(d_parts[1])}/{int(d_parts[2])}"
              if len(d_parts) == 3
              else r["実行日"]
          )

          display_data.append({
              "管理名": r["管理名"],
              "実行日": d_str,
              "時間帯": r["時間帯"],
              "預かり金額": (
                  f"¥{int(r['預かり金額']):,}"
                  if r["預かり金額"] is not None
                  else ""
              ),
              "計上売上": (
                  f"¥{int(r['計上売上']):,}"
                  if r["計上売上"] is not None
                  else ""
              ),
              "残額": f"¥{int(r['残額']):,}",
              "顧客名": r["顧客名"],
              "電話番号": r["電話番号"],
              "備考": r["備考"],
          })

        if not display_data:
          st.write("該当するデータがありません。")
        else:
          html = (
              '<div style="max-height: 60vh; overflow-y: auto; overflow-x:'
              " auto; margin-bottom: 20px;\"><table style=\"width: 100%;"
              " border-collapse: collapse; font-family: sans-serif; font-size:"
              ' 14px; text-align: center; white-space: nowrap;"><thead><tr'
              ' style="text-align: center;">'
          )
          for col in [
              "管理名",
              "実行日",
              "時間帯",
              "預かり金額",
              "計上売上",
              "残額",
              "顧客名",
              "電話番号",
              "備考",
          ]:
            html += (
                '<th style="position: sticky; top: 0; background-color:'
                " #f39c12; color: #ffffff; border: 1px solid #e0e0e0; padding:"
                f' 8px; font-weight: bold; z-index: 1;">{col}</th>'
            )
          html += "</tr></thead><tbody>"
          for idx, row in enumerate(display_data):
            bg_color = "#fffbf0" if idx % 2 == 0 else "#ffffff"
            html += (
                f'<tr style="background-color: {bg_color}; text-align: center;">'
            )
            for col in [
                "管理名",
                "実行日",
                "時間帯",
                "預かり金額",
                "計上売上",
                "残額",
                "顧客名",
                "電話番号",
                "備考",
            ]:
              val = str(row[col])
              style = (
                  "border: 1px solid #e0e0e0; padding: 8px; white-space:"
                  " nowrap;"
              )
              if col in ["預かり金額", "計上売上", "残額"] and val != "":
                style += " font-weight: bold;"
              html += f'<td style="{style}">{val}</td>'
            html += "</tr>"
          html += "</tbody></table></div>"
          st.markdown(html, unsafe_allow_html=True)

    elif manage_mode == "入金データの修正・削除":
      if store_reservations.empty:
        st.write("修正できる入金データがありません。")
      else:
        edit_res_list = [
            f"{r['date']} | {r['management_name']} (¥{r['amount']:,})"
            for _, r in store_reservations.iterrows()
        ]
        selected_edit_str = st.selectbox(
            "修正・削除する予約を選択", edit_res_list
        )
        sel_idx = edit_res_list.index(selected_edit_str)
        sel_res = store_reservations.iloc[sel_idx]

        st.markdown(f"#### 【{sel_res['management_name']}】のデータ編集")

        k_e_date = f"e_date_{sel_res['id']}"
        k_e_slot = f"e_slot_{sel_res['id']}"
        k_e_mng = f"e_mng_{sel_res['id']}"
        k_e_amt = f"e_amt_{sel_res['id']}"
        k_e_cust = f"e_cust_{sel_res['id']}"
        k_e_phone = f"e_phone_{sel_res['id']}"
        k_e_notes = f"e_notes_{sel_res['id']}"

        c1, c2 = st.columns(2)
        with c1:
          e_date = st.date_input(
              "入金日",
              value=datetime.strptime(sel_res["date"], "%Y-%m-%d").date(),
              key=k_e_date,
          ).strftime("%Y-%m-%d")
        with c2:
          idx_slot = (
              (["時間帯未指定"] + TIME_SLOTS).index(sel_res["time_slot"])
              if sel_res["time_slot"] in ["時間帯未指定"] + TIME_SLOTS
              else 0
          )
          e_slot = st.selectbox(
              "時間帯",
              ["時間帯未指定"] + TIME_SLOTS,
              index=idx_slot,
              key=k_e_slot,
          )
        c3, c4 = st.columns(2)
        with c3:
          e_mng = st.text_input(
              "管理名", value=sel_res["management_name"], key=k_e_mng
          )
        with c4:
          e_amt = st.number_input(
              "預かり金額 (円)",
              value=int(sel_res["amount"]),
              min_value=0,
              step=1000,
              key=k_e_amt,
          )
        c5, c6 = st.columns(2)
        with c5:
          e_cust = st.text_input(
              "顧客名",
              value=(
                  str(sel_res["customer_name"])
                  if pd.notna(sel_res["customer_name"])
                  else ""
              ),
              key=k_e_cust,
          )
        with c6:
          e_phone = st.text_input(
              "電話番号",
              value=(
                  str(sel_res["phone"]) if pd.notna(sel_res["phone"]) else ""
              ),
              key=k_e_phone,
          )
        e_notes = st.text_area(
            "備考",
            value=str(sel_res["notes"]) if pd.notna(sel_res["notes"]) else "",
            key=k_e_notes,
        )

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
          if st.button("🔄 この内容で更新する", type="primary"):
            total_sales = res_sales_df[
                res_sales_df["reservation_id"] == sel_res["id"]
            ]["amount"].sum()
            original_amt = int(sel_res["amount"])

            if original_amt <= total_sales and e_amt != original_amt:
              st.error(
                  "完了済みのため、金額の変更はできません。売上計上の修正を先に行ってください。"
              )
            elif e_amt < total_sales:
              st.error(
                  f"計上済みの売上（¥{total_sales:,}）を下回る金額には変更できません。"
              )
            else:
              update_reservation(
                  sel_res["id"],
                  e_date,
                  e_slot,
                  e_mng,
                  e_amt,
                  e_cust,
                  e_phone,
                  e_notes,
              )
              st.session_state["flash_msg_manage"] = (
                  "🔄 予約データを更新しました。"
              )
              st.rerun()
        with col_btn2:
          if st.button("🗑️ この予約を削除する", type="secondary"):
            delete_reservation(sel_res["id"])
            st.session_state["flash_msg_manage"] = (
                "🗑️ 予約データおよび関連する売上計上データを削除しました。"
            )
            st.rerun()

    elif manage_mode == "売上計上データの修正・削除":
      store_sales = res_sales_df[
          res_sales_df["store"] == selected_store
      ].copy()
      if store_sales.empty:
        st.write("修正できる売上計上データがありません。")
      else:
        sale_options = []
        for _, s in store_sales.iterrows():
          parent = res_df[res_df["id"] == s["reservation_id"]]
          mng = (
              parent["management_name"].values[0]
              if not parent.empty
              else "不明な予約"
          )
          sale_options.append(
              f"{s['date']} | {mng} | {s['time_slot']} (¥{s['amount']:,})"
          )

        selected_sale_str = st.selectbox(
            "修正・削除する売上計上データを選択", sale_options
        )
        sel_sale_idx = sale_options.index(selected_sale_str)
        sel_sale = store_sales.iloc[sel_sale_idx]

        st.markdown("#### 売上計上データの編集")

        k_e_s_date = f"e_s_date_{sel_sale['id']}"
        k_e_s_slot = f"e_s_slot_{sel_sale['id']}"
        k_e_s_amt = f"e_s_amt_{sel_sale['id']}"

        c1, c2, c3 = st.columns(3)
        with c1:
          e_s_date = st.date_input(
              "計上日",
              value=datetime.strptime(sel_sale["date"], "%Y-%m-%d").date(),
              key=k_e_s_date,
          ).strftime("%Y-%m-%d")
        with c2:
          e_s_slot = st.selectbox(
              "計上 時間帯",
              TIME_SLOTS,
              index=(
                  TIME_SLOTS.index(sel_sale["time_slot"])
                  if sel_sale["time_slot"] in TIME_SLOTS
                  else 0
              ),
              key=k_e_s_slot,
          )
        with c3:
          e_s_amt = st.number_input(
              "計上金額 (円)",
              value=int(sel_sale["amount"]),
              min_value=0,
              step=500,
              key=k_e_s_amt,
          )

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
          if st.button("🔄 この計上データを更新する", type="primary"):
            update_reservation_sale(sel_sale["id"], e_s_date, e_s_slot, e_s_amt)
            st.session_state["flash_msg_manage"] = (
                "🔄 売上計上データを更新しました。"
            )
            st.rerun()
        with col_btn2:
          if st.button("🗑️ この計上データを削除する", type="secondary"):
            delete_reservation_sale(sel_sale["id"])
            st.session_state["flash_msg_manage"] = (
                "🗑️ 売上計上データを削除しました。"
            )
            st.rerun()

elif app_mode == "📋 エクセルコピペ用データ生成":
  st.title("📋 エクセル実績入力用データ生成")
  st.markdown(f"現在選択中の店舗: **{selected_store}**")
  target_date_val = st.date_input(
      "対象日付",
      value=st.session_state["target_date_state"],
      key="copy_date_input",
  )
  if target_date_val != st.session_state["target_date_state"]:
    st.session_state["target_date_state"] = target_date_val
  target_date_str = target_date_val.strftime("%Y-%m-%d")

  raw_df = find_airregi_data_in_memory(target_date_str, airregi_year_df)
  if raw_df is None:
    st.warning(f"⚠️ **{target_date_str}** のAirレジデータが見つかりません。")

  def get_slot_metrics(store_name, slot_name, df_data):
    air_sales, air_cust = 0, 0
    if df_data is not None:
      try:
        short_key = get_short_store_key(store_name)
        store_df = df_data[
            df_data[0].astype(str).str.contains(short_key, na=False)
        ]
        matched = match_slot_row(store_df, slot_name)
        if not matched.empty:
          air_sales = parse_int(matched.iloc[0, 2])
          air_cust = parse_int(matched.iloc[0, 5])
      except Exception:
        pass
    return air_sales, air_cust

  def get_slot_delivery_total(store_name, date_str, slot_name):
    sub = deliv_df[
        (deliv_df["date"].astype(str).str.strip() == date_str.strip())
        & (deliv_df["store"].astype(str).str.strip() == store_name.strip())
        & (deliv_df["time_slot"].astype(str).str.strip() == slot_name.strip())
    ]
    return (
        int(
            pd.to_numeric(sub["uber"], errors="coerce").fillna(0).sum()
            + pd.to_numeric(sub["rocket_now"], errors="coerce").fillna(0).sum()
            + pd.to_numeric(sub["demaekan"], errors="coerce").fillna(0).sum()
        )
        if not sub.empty
        else 0
    )

  early_d = get_slot_delivery_total(selected_store, target_date_str, "早番")
  mid_d = get_slot_delivery_total(selected_store, target_date_str, "中番")
  late_d = get_slot_delivery_total(selected_store, target_date_str, "遅番")
  early_sales, early_cust = get_slot_metrics(selected_store, "早番", raw_df)
  mid_sales, mid_cust = get_slot_metrics(selected_store, "中番", raw_df)
  late_sales, late_cust = get_slot_metrics(selected_store, "遅番", raw_df)

  df_month_all = generate_month_df(
      selected_store,
      target_date_val.year,
      target_date_val.month,
      None,
      bp_df,
      deliv_df,
      res_df,
      res_sales_df,
      airregi_year_df,
  )
  day_row = df_month_all[df_month_all["日"] == target_date_val.day]
  target_other_sales = (
      int(day_row["その他売上"].values[0]) if not day_row.empty else 0
  )

  parts = [
      str(early_cust),
      str(early_sales),
      str(early_d),
      str(mid_cust),
      str(mid_sales),
      str(mid_d),
      str(late_cust),
      str(late_sales),
      str(late_d),
      str(target_other_sales),
  ]
  tabs_list = [
      "\t" * 4,
      "\t" * 1,
      "\t" * 3,
      "\t" * 4,
      "\t" * 1,
      "\t" * 3,
      "\t" * 4,
      "\t" * 1,
      "\t" * 1,
  ]
  tsv_text = "".join([
      part + tabs_list[i] if i < len(tabs_list) else part
      for i, part in enumerate(parts)
  ])
  st.success(f"🎉 {target_date_str} のデータを自動生成しました！")
  st.code(tsv_text, language="text")

elif app_mode == "📄 報告書作成":
  st.title("📄 LINE報告書作成")
  report_date = st.date_input(
      "集計基準日",
      value=st.session_state["target_date_state"],
      key="report_date_input",
  )
  if report_date != st.session_state["target_date_state"]:
    st.session_state["target_date_state"] = report_date

  if st.button("📝 報告用テキストを生成する", type="primary"):
    _, max_days = calendar.monthrange(year, month)
    target_end_day = min(report_date.day, max_days)
    report_lines = [
        "お疲れ様です。",
        f"{report_date.month}月{report_date.day}日時点の中間報告です。",
        "",
    ]

    for store_name in STORES:
      short_label = STORE_CREDENTIALS[store_name]["short_label"]
      df_rep = generate_month_df(
          store_name,
          year,
          month,
          None,
          bp_df,
          deliv_df,
          res_df,
          res_sales_df,
          airregi_year_df,
      )
      df_rep_act = df_rep[df_rep["日"] <= target_end_day]

      total_p = df_rep_act["前年売上"].sum()
      total_b = df_rep_act["予算売上"].sum()
      total_act = df_rep_act["実績売上"].sum()

      prev_ratio = f"{(total_act / total_p * 100):.1f}%" if total_p > 0 else "-"
      budget_ratio = (
          f"{(total_act / total_b * 100):.1f}%" if total_b > 0 else "-"
      )
      diff_budget = total_act - total_b
      diff_str = (
          f"+{diff_budget:,}円"
          if diff_budget > 0
          else f"▲{abs(diff_budget):,}円"
          if diff_budget < 0
          else "0円"
      )

      report_lines.extend([
          f"【{short_label}】",
          f"前年比 : {prev_ratio}",
          f"予算比 : {budget_ratio}",
          f"累売上 : {total_act:,}円",
          f"予乖離 : {diff_str}",
          "",
      ])

    st.success("🎉 報告用テキストを生成しました！")
    st.code("\n".join(report_lines).strip(), language="text")

else:
  st.title("📊 売上月報一覧")
  selected_year_month = st.session_state["selected_year_month"]
  year, month = map(int, selected_year_month.split("-"))
  st.subheader(f"🏬 {selected_store} {year}年{month}月度")
  st.markdown("---")

  def format_currency(val):
    if pd.isna(val):
      return "-"
    try:
      num = int(val)
      return (
          f"-¥{abs(num):,}" if num < 0 else f"¥{num:,}" if num > 0 else "¥0"
      )
    except:
      return str(val)

  # 全体サマリー用
  df_month_total = generate_month_df(
      selected_store,
      year,
      month,
      None,
      bp_df,
      deliv_df,
      res_df,
      res_sales_df,
      airregi_year_df,
  )
  summary_calculation_started_at = time.perf_counter()
  df_achieved = df_month_total[df_month_total["実績売上"] != 0]
  if df_achieved.empty:
    max_day, df_achieved = 1, df_month_total.iloc[:1]
  else:
    max_day = int(df_achieved["日"].max())

  t_p, t_b, t_act = (
      df_achieved["前年売上"].sum(),
      df_achieved["予算売上"].sum(),
      df_achieved["実績売上"].sum(),
  )
  t_air, t_del, t_oth = (
      df_achieved["店頭売上"].sum(),
      df_achieved["D売上"].sum(),
      df_achieved["その他売上"].sum(),
  )
  t_dcount = df_achieved["D件数"].sum()
  t_dunit = (
      int(df_achieved.loc[df_achieved["D売上"] != 0, "D単価"].mean())
      if (df_achieved["D売上"] != 0).any()
      else 0
  )
  t_res_dep, t_res_sales = (
      df_achieved["予約金"].sum(),
      df_achieved["売上計上"].sum(),
  )
  t_pc, t_bc, t_ac = (
      df_achieved["前年客数"].sum(),
      df_achieved["予算客数"].sum(),
      df_achieved["実績客数"].sum(),
  )
  t_store_cust = df_achieved["店頭客数"].sum()

  tot_pprice = int(t_p / t_pc) if t_pc > 0 else 0
  tot_bprice = int(t_b / t_bc) if t_bc > 0 else 0
  tot_aprice = int(t_act / t_ac) if t_ac > 0 else 0

  p_ratio_str = f"{(t_act / t_p * 100):.1f}%" if t_p > 0 else "-"
  b_ratio_str = f"{(t_act / t_b * 100):.1f}%" if t_b > 0 else "-"
  cust_ratio_str = f"{(t_ac / t_pc * 100):.1f}%" if t_pc > 0 else "-"
  cust_budget_ratio_str = f"{(t_ac / t_bc * 100):.1f}%" if t_bc > 0 else "-"
  price_ratio_str = f"{(tot_aprice / tot_pprice * 100):.1f}%" if tot_pprice > 0 else "-"
  price_budget_ratio_str = (
      f"{(tot_aprice / tot_bprice * 100):.1f}%" if tot_bprice > 0 else "-"
  )

  summary_data = [{
      "日": f"{max_day}日時点",
      "曜日": "合計",
      "前年比": p_ratio_str,
      "予算比": b_ratio_str,
      "前年売上": format_currency(t_p),
      "予算売上": format_currency(t_b),
      "実績売上": format_currency(t_act),
      "店頭売上": format_currency(t_air),
      "D売上": format_currency(t_del),
      "その他売上": format_currency(t_oth),
      "予約金": format_currency(t_res_dep),
      "売上計上": format_currency(t_res_sales),
      "前年客数": int(t_pc),
      "予算客数": int(t_bc),
      "店頭客数": int(t_store_cust),
      "実績客数": int(t_ac),
      "D件数": int(t_dcount),
      "客数前年比": cust_ratio_str,
      "客数予算比": cust_budget_ratio_str,
      "前年単価": format_currency(tot_pprice),
      "予算単価": format_currency(tot_bprice),
      "実績単価": format_currency(tot_aprice),
      "D単価": format_currency(t_dunit),
      "単価前年比": price_ratio_str,
      "単価予算比": price_budget_ratio_str,
  }]
  log_performance_time("月報サマリー集計", summary_calculation_started_at)

  def render_styled_html_table(df, is_summary=False):
    header_bg, header_color = ("#f39c12", "#ffffff")
    wrap_style = (
        "overflow-x: auto; margin-bottom: 20px;"
        if is_summary
        else (
            "max-height: 60vh; overflow-y: auto; overflow-x: auto;"
            " margin-bottom: 20px;"
        )
    )
    cols = list(df.columns)

    html = (
        f'<div style="{wrap_style}"><table style="width: 100%;'
        " border-collapse: collapse; font-family: sans-serif; font-size: 14px;"
        ' white-space: nowrap;"><thead><tr style="text-align: center;">'
    )
    for col in cols:
      html += (
          f'<th style="position: sticky; top: 0; background-color: {header_bg};'
          f" color: {header_color}; border: 1px solid #e0e0e0; padding: 8px;"
          f' font-weight: bold; z-index: 1;">{col}</th>'
      )
    html += "</tr></thead><tbody>"

    for idx, row in enumerate(df.to_dict(orient="records")):
      bg_color = "#fffbf0" if idx % 2 == 0 else "#ffffff"
      html += (
          f'<tr style="background-color: {bg_color}; text-align: center;">'
      )
      for col in cols:
        val = str(row[col])
        style = (
            "border: 1px solid #e0e0e0; padding: 8px; white-space: nowrap;"
        )
        if (
            col
            in [
                "前年比",
                "予算比",
                "客数前年比",
                "客数予算比",
                "単価前年比",
                "単価予算比",
            ]
            or "%" in val
        ):
          if "%" in val:
            try:
              num = float(val.replace("%", "").strip())
              style += (
                  " color: #27ae60; font-weight: bold;"
                  if num >= 100.0
                  else " color: #c0392b; font-weight: bold;"
              )
            except Exception:
              style += " font-weight: bold;"
          else:
            style += " font-weight: bold;"
        elif "-" in val and any(char.isdigit() for char in val):
          style += " color: #c0392b; font-weight: bold;"
        elif (
            col
            in [
                "前年売上",
                "予算売上",
                "実績売上",
                "店頭売上",
                "D売上",
                "D単価",
                "その他売上",
                "予約金",
                "売上計上",
                "前年単価",
                "予算単価",
                "実績単価",
            ]
            or is_summary
        ):
          style += " font-weight: bold;"
        html += f'<td style="{style}">{val}</td>'
      html += "</tr>"
    return html + "</tbody></table></div>"

  st.markdown(f"### 【{max_day}日時点 累計サマリー】")
  render_started_at = time.perf_counter()
  summary_html = render_styled_html_table(
      pd.DataFrame(summary_data), is_summary=True
  )
  log_performance_time("月報サマリーHTML生成", render_started_at)
  st.markdown(
      summary_html,
      unsafe_allow_html=True,
  )

  st.markdown("---")
  st.markdown("### 🔍 表示する時間帯を選択してください")

  if "active_slot_tab" not in st.session_state:
    st.session_state["active_slot_tab"] = "合計"
  selected_tab = st.radio(
      "時間帯タブ",
      ["合計", "早番", "中番", "遅番"],
      horizontal=True,
      label_visibility="collapsed",
  )
  st.session_state["active_slot_tab"] = selected_tab
  current_kw = {"合計": None, "早番": "早番", "中番": "中番", "遅番": "遅番"}[
      selected_tab
  ]

  df_current_slot = generate_month_df(
      selected_store,
      year,
      month,
      current_kw,
      bp_df,
      deliv_df,
      res_df,
      res_sales_df,
      airregi_year_df,
  )
  formatting_started_at = time.perf_counter()
  df_current_slot_formatted = df_current_slot.copy()
  for col in [
      "前年売上",
      "予算売上",
      "実績売上",
      "店頭売上",
      "D売上",
      "D単価",
      "その他売上",
      "予約金",
      "売上計上",
      "前年単価",
      "予算単価",
      "実績単価",
  ]:
    df_current_slot_formatted[col] = df_current_slot_formatted[col].apply(
        format_currency
    )
  log_performance_time("明細データ整形", formatting_started_at)

  st.markdown("---")
  col_title, col_btn = st.columns([8, 2])
  with col_title:
    st.markdown("### 📅 日別実績一覧")
  with col_btn:
    st.markdown("<div style='padding-top: 10px;'></div>", unsafe_allow_html=True)
    st.download_button(
        label="📥 Excelダウンロード",
        data=lambda: create_multi_sheet_excel(
            selected_store,
            year,
            month,
            bp_df,
            deliv_df,
            res_df,
            res_sales_df,
            airregi_year_df,
        ),
        file_name=f"{selected_store}_{year}年{month}月度_売上月報.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True,
    )

  render_started_at = time.perf_counter()
  monthly_html = render_styled_html_table(df_current_slot_formatted)
  log_performance_time("月報明細HTML生成", render_started_at)
  st.markdown(monthly_html, unsafe_allow_html=True)
  log_performance_time("月報画面全体", APP_RUN_STARTED_AT)
  if APP_VARIANT != "admin":
    with st.expander("⏱ 月報の処理時間（今回の表示）"):
      st.dataframe(
          pd.DataFrame(st.session_state["performance_timings"]),
          hide_index=True,
          use_container_width=True,
      )
  if needs_monthly_report_data and APP_VARIANT != "admin":
    with st.expander("🔎 Airレジ年別シートの照合結果"):
      st.caption(
          "本部アプリが読み込んだAirレジ年別シートから、"
          "選択店舗・月の店舗行と「合計」行を確認します。"
      )
      st.dataframe(
          build_airregi_match_diagnostics(
              airregi_year_df, selected_store, year, month
          ),
          hide_index=True,
          use_container_width=True,
      )
