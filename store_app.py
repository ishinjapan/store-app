from datetime import date
import os
import re
import pandas as pd

# ==========================================
# 設定
# ==========================================
TARGET_DIR = r"C:\Users\YOSHIDA\airregi-automation\過去月報"
OUTPUT_CSV = "app_sales_data.csv"

STORE_MAP = {
    "三田": "三田おにぎり 本店",
    "恵比寿": "おにぎり維心 恵比寿店",
    "五反田": "おにぎり維心 五反田店",
    "中目黒": "おにぎり維心 中目黒店",
}


def clean_text(val):
    if pd.isna(val):
        return ""
    return re.sub(r"[\s\u3000\xa0\ufeff]+", "", str(val))


def extract_year_month(file_path):
    file_name = os.path.basename(file_path)
    parent_dir = os.path.basename(os.path.dirname(file_path))

    for text in [file_name, parent_dir]:
        m = re.search(r"(202\d)[\.\-_年](\d{1,2})", text)
        if m:
            y, mo = int(m.group(1)), int(m.group(2))
            if 1 <= mo <= 12:
                return y, mo

    matches = re.findall(r"(202\d)[\.\-_年](\d{1,2})", file_path)
    for y_str, mo_str in reversed(matches):
        y, mo = int(y_str), int(mo_str)
        if 1 <= mo <= 12:
            return y, mo
    return None, None


def detect_columns_by_order(df, start_row):
    header_df = df.iloc[:start_row]
    num_cols = len(df.columns)

    prev_cols = []
    budget_cols = []
    store_col = None
    delivery_col = None

    for c in range(num_cols):
        col_cells = [
            clean_text(header_df.iat[r, c]) for r in range(len(header_df))
        ]
        col_text = "_".join(col_cells)
        col_text_lower = col_text.lower()

        # 除外キーワード
        is_ruikei = "累計" in col_text or "累前年" in col_text
        is_jisseki = "実績" in col_text
        is_hittatsu = "必達" in col_text or "店長" in col_text
        is_rate = "達成率" in col_text
        is_tanka = "単価" in col_text

        # 1. 店頭売上（「店舗」または「本店」、ただし単価列は除外）
        if (
            ("店舗" in col_text or "本店" in col_text)
            and not is_tanka
            and store_col is None
        ):
            store_col = c

        # 2. デリバリー売上（「uber」または「ubar」、ただし単価列は除外）
        if (
            ("uber" in col_text_lower or "ubar" in col_text_lower)
            and not is_tanka
            and delivery_col is None
        ):
            delivery_col = c

        # 3. 前年・予算列の判定（構成・単価等は除外）
        if is_ruikei or is_jisseki or is_hittatsu or is_rate or is_tanka:
            continue

        is_nichibetsu = "日別" in col_text

        if "前年" in col_text and is_nichibetsu:
            prev_cols.append(c)

        if (
            "予算" in col_text
            and is_nichibetsu
            and "日予算" not in col_text
            and "累予算" not in col_text
        ):
            budget_cols.append(c)

    col_map = {
        "prev_sales": prev_cols[0] if len(prev_cols) >= 1 else None,
        "prev_cust": prev_cols[1] if len(prev_cols) >= 2 else None,
        "budget_sales": budget_cols[0] if len(budget_cols) >= 1 else None,
        "budget_cust": budget_cols[1] if len(budget_cols) >= 2 else None,
        "store_sales": store_col,
        "delivery_sales": delivery_col,
    }
    return col_map


# ==========================================
# メイン処理
# ==========================================
records = []
print("=" * 65)
print(f"走査フォルダ: {TARGET_DIR}")

excel_files = []
for root, _, files in os.walk(TARGET_DIR):
    for f in files:
        if f.startswith("~$"):
            continue
        ext = os.path.splitext(f)[1].lower()
        if ext in [".xlsx", ".xlsb", ".xlsm"]:
            excel_files.append(os.path.join(root, f))

print(f"対象ファイル数: {len(excel_files)} 件")
print("データ抽出を開始します...\n")

logged_sample = False

for file_path in excel_files:
    file_name = os.path.basename(file_path)

    matched_store = None
    for kw, formal in STORE_MAP.items():
        if kw in file_name:
            matched_store = formal
            break
    if not matched_store:
        continue

    year, month = extract_year_month(file_path)
    if not year or not month:
        continue

    ext = os.path.splitext(file_name)[1].lower()
    engine = "pyxlsb" if ext == ".xlsb" else "openpyxl"

    try:
        xl = pd.ExcelFile(file_path, engine=engine)
        sheet_names = xl.sheet_names
    except Exception:
        continue

    for sheet in sheet_names:
        clean_sheet = str(sheet).strip()
        time_slot = None
        if "早" in clean_sheet:
            time_slot = "早番"
        elif "中" in clean_sheet:
            time_slot = "中番"
        elif "遅" in clean_sheet:
            time_slot = "遅番"

        if not time_slot:
            continue

        try:
            df = pd.read_excel(
                file_path, sheet_name=sheet, header=None, engine=engine
            )
        except Exception:
            continue

        # 営業日列（1, 2...）とデータ開始行の検出
        day_col = None
        start_row = None
        for r in range(min(10, len(df))):
            for c in range(min(5, len(df.columns))):
                v = df.iat[r, c]
                try:
                    if int(float(v)) == 1:
                        if r + 1 < len(df) and int(float(df.iat[r + 1, c])) == 2:
                            day_col = c
                            start_row = r
                            break
                except (ValueError, TypeError):
                    pass
            if start_row is not None:
                break

        if start_row is None:
            continue

        col_map = detect_columns_by_order(df, start_row)

        if not logged_sample:
            print(f"【列判定確認】ファイル: {file_name}（{sheet}）")
            print(
                f"  -> 売上(前年): 列{col_map['prev_sales']}, 売上(予算): 列{col_map['budget_sales']}"
            )
            print(
                f"  -> 客数(前年): 列{col_map['prev_cust']}, 客数(予算): 列{col_map['budget_cust']}"
            )
            print(
                f"  -> 店頭売上(店舗/本店): 列{col_map['store_sales']}, デリバリー(Uber/Ubar): 列{col_map['delivery_sales']}"
            )
            print("-" * 65)
            logged_sample = True

        for r in range(start_row, len(df)):
            day_val = df.iat[r, day_col]
            try:
                if pd.isna(day_val):
                    continue
                d = int(float(day_val))
                if not (1 <= d <= 31):
                    continue
                cur_date = date(year, month, d).strftime("%Y-%m-%d")
            except (ValueError, TypeError):
                continue

            def get_val(item_key):
                target_c = col_map.get(item_key)
                if target_c is None or target_c >= len(df.columns):
                    return ""
                try:
                    val = df.iat[r, target_c]
                    if pd.isna(val) or str(val).strip() in [
                        "",
                        "-",
                        "#REF!",
                        "#####",
                    ]:
                        return ""
                    return int(round(float(val)))
                except (ValueError, TypeError, IndexError):
                    return ""

            records.append(
                {
                    "date": cur_date,
                    "store": matched_store,
                    "time_slot": time_slot,
                    "budget_sales": get_val("budget_sales"),
                    "prev_sales": get_val("prev_sales"),
                    "prev_cust": get_val("prev_cust"),
                    "budget_cust": get_val("budget_cust"),
                    "store_sales": get_val("store_sales"),
                    "delivery_sales": get_val("delivery_sales"),
                }
            )

# 出力
result_df = pd.DataFrame(records)

print("=" * 65)
if not result_df.empty:
    slot_order = {"早番": 0, "中番": 1, "遅番": 2}
    result_df["slot_rank"] = result_df["time_slot"].map(slot_order)
    
    # ★ここが重要：完全に重複を排除（1日・1店舗・1シフトにつき1行に絞る）
    result_df = result_df.drop_duplicates(subset=["date", "store", "time_slot"], keep="first")
    
    result_df.sort_values(by=["date", "store", "slot_rank"], inplace=True)
    result_df.drop(columns=["slot_rank"], inplace=True)

    save_path = OUTPUT_CSV
    try:
        result_df.to_csv(save_path, index=False, encoding="utf-8-sig")
    except PermissionError:
        save_path = "app_sales_data_new.csv"
        result_df.to_csv(save_path, index=False, encoding="utf-8-sig")
        print("※ app_sales_data.csv が開かれていたため、別名で保存しました。")

    print(f"【抽出成功！】 合計 {len(result_df)} 件のデータを抽出しました。")
    print(f"保存先: {os.path.abspath(save_path)}")
else:
    print("データが抽出できませんでした。")
print("=" * 65)
