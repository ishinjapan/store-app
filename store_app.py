import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ページ設定
st.set_page_config(
    page_title="店舗用デリバリー売上入力システム",
    page_icon="📱",
    layout="centered"
)

DATA_DIR = os.path.abspath("./data")
os.makedirs(DATA_DIR, exist_ok=True)
DELIVERY_DATA_PATH = os.path.join(DATA_DIR, "delivery_sales.csv")

# 店舗定義（IDとパスワードの紐付け）
STORE_CREDENTIALS = {
    "三田おにぎり 本店": {"id": "001", "pass": "0000"},
    "おにぎり維心 恵比寿店": {"id": "002", "pass": "0000"},
    "おにぎり維心 五反田店": {"id": "003", "pass": "0000"},
    "おにぎり維心 中目黒店": {"id": "004", "pass": "0000"}
}

# 上書き用管理者パスワード
OVERWRITE_ADMIN_PASS = "1111"

TIME_SLOTS = ["早番", "中番", "遅番"]

def normalize_time_slot(val):
    s = str(val)
    if "早番" in s:
        return "早番"
    elif "中番" in s:
        return "中番"
    elif "遅番" in s:
        return "遅番"
    return s

def init_delivery_file():
    if not os.path.exists(DELIVERY_DATA_PATH):
        pd.DataFrame(columns=["date", "store", "time_slot", "uber", "rocket_now", "demaekan", "other"]).to_csv(DELIVERY_DATA_PATH, index=False, encoding="utf-8-sig")

init_delivery_file()

def load_delivery_data():
    df = pd.read_csv(DELIVERY_DATA_PATH)
    df["time_slot"] = df["time_slot"].apply(normalize_time_slot)
    if "rn" in df.columns and "rocket_now" not in df.columns:
        df = df.rename(columns={"rn": "rocket_now"})
    return df

def save_delivery_data(date_str, store, time_slot, uber, rocket_now, demaekan, other):
    df = load_delivery_data()
    df = df[~((df["date"] == date_str) & (df["store"] == store) & (df["time_slot"] == time_slot))]
    
    new_row = pd.DataFrame([{
        "date": date_str,
        "store": store,
        "time_slot": time_slot,
        "uber": uber,
        "rocket_now": rocket_now,
        "demaekan": demaekan,
        "other": other
    }])
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_csv(DELIVERY_DATA_PATH, index=False, encoding="utf-8-sig")

# セッション状態の初期化
if "logged_in_store" not in st.session_state:
    st.session_state["logged_in_store"] = None

if "show_overwrite_auth" not in st.session_state:
    st.session_state["show_overwrite_auth"] = False

# --- 画面描画 ---
if st.session_state["logged_in_store"] is None:
    st.title("🔐 店舗ログイン画面")
    st.write("デリバリー売上を入力する店舗の **店舗ID** と **パスワード** を入力してください。")
    
    with st.form("login_form"):
        input_id = st.text_input("店舗ID (例: 001)")
        input_pass = st.text_input("パスワード", type="password")
        login_btn = st.form_submit_button("ログイン", type="primary")
        
        if login_btn:
            matched_store = None
            for s_name, creds in STORE_CREDENTIALS.items():
                if creds["id"] == input_id.strip() and creds["pass"] == input_pass.strip():
                    matched_store = s_name
                    break
            
            if matched_store:
                st.session_state["logged_in_store"] = matched_store
                st.success(f"🎉 {matched_store} にログインしました！")
                st.rerun()
            else:
                st.error("❌ 店舗IDまたはパスワードが間違っています。正しい情報を入力してください。")

else:
    current_store = st.session_state["logged_in_store"]
    
    col_top1, col_top2 = st.columns([4, 1])
    with col_top1:
        st.title(f"📱 デリバリー売上 入力フォーム")
        st.markdown(f"現在ログイン中の店舗: **{current_store}**")
    with col_top2:
        st.write("###")
        if st.button("ログアウト", type="secondary"):
            st.session_state["logged_in_store"] = None
            st.session_state["show_overwrite_auth"] = False
            st.rerun()

    st.write("時間帯ごとのデリバリー売上（4サービス）を入力します。")
    
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        selected_date = st.date_input("対象日付", value=datetime.now().date()).strftime("%Y-%m-%d")
    with col_s2:
        selected_slot = st.selectbox("時間帯を選択", TIME_SLOTS)
    
    st.markdown("---")
    
    deliv_df = load_delivery_data()
    existing_row = deliv_df[(deliv_df["date"] == selected_date) & (deliv_df["store"] == current_store) & (deliv_df["time_slot"] == selected_slot)]
    
    is_already_registered = not existing_row.empty

    default_uber = 0
    default_rn = 0
    default_demae = 0
    default_other = 0

    if is_already_registered:
        st.warning(f"⚠️ この日時（{selected_date} / {selected_slot}）のデータはすでに登録されています。修正・上書き保存には管理者権限パスワードが必要です。")
        if "uber" in existing_row.columns:
            default_uber = int(existing_row["uber"].values[0])
        if "rocket_now" in existing_row.columns:
            default_rn = int(existing_row["rocket_now"].values[0])
        if "demaekan" in existing_row.columns:
            default_demae = int(existing_row["demaekan"].values[0])
        if "other" in existing_row.columns:
            default_other = int(existing_row["other"].values[0])

    col_d1, col_d2 = st.columns(2)
    with col_d1:
        uber_val = st.number_input("Uber Eats 売上 (円)", min_value=0, value=default_uber, step=500)
        rn_val = st.number_input("Rocket now 売上 (円)", min_value=0, value=default_rn, step=500)
    with col_d2:
        demaekan_val = st.number_input("出前館 売上 (円)", min_value=0, value=default_demae, step=500)
        other_val = st.number_input("その他 売上 (円)", min_value=0, value=default_other, step=500)

    st.markdown("---")
    total_delivery = uber_val + rn_val + demaekan_val + other_val
    st.info(f"💡 この時間帯のデリバリー売上合計: **¥{total_delivery:,}**")

    if not is_already_registered:
        if st.button("💾 デリバリー売上を保存する", type="primary"):
            save_delivery_data(selected_date, current_store, selected_slot, uber_val, rn_val, demaekan_val, other_val)
            st.success(f"🎉 {current_store} ({selected_date} / {selected_slot}) のデリバリー売上を保存しました！")
            st.rerun()
    else:
        if not st.session_state["show_overwrite_auth"]:
            if st.button("🔒 修正して上書き保存する（要パスワード）", type="primary"):
                st.session_state["show_overwrite_auth"] = True
                st.rerun()
        else:
            st.markdown("#### 🔐 管理者パスワード認証")
            admin_pass_input = st.text_input("上書き用パスワードを入力してください", type="password", key="overwrite_pass_input")
            
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("✅ 認証して上書き実行", type="primary"):
                    if admin_pass_input.strip() == OVERWRITE_ADMIN_PASS:
                        save_delivery_data(selected_date, current_store, selected_slot, uber_val, rn_val, demaekan_val, other_val)
                        st.session_state["show_overwrite_auth"] = False
                        st.success(f"🎉 パスワード認証成功！ {current_store} のデータを上書き保存しました！")
                        st.rerun()
                    else:
                        st.error("❌ パスワードが違います。")
            with col_btn2:
                if st.button("キャンセル"):
                    st.session_state["show_overwrite_auth"] = False
                    st.rerun()