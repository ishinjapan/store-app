import os
import sys
import time
from datetime import datetime, timedelta
import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv()
AIRREGI_ID = os.getenv("AIRREGI_ID")
AIRREGI_PASS = os.getenv("AIRREGI_PASS")

OUTPUT_DIR = os.path.abspath("./output")
os.makedirs(OUTPUT_DIR, exist_ok=True)
USER_DATA_DIR = "./user_data"

TARGET_STORES = [
    "三田おにぎり 本店",
    "おにぎり維心 恵比寿店",
    "おにぎり維心 五反田店",
    "おにぎり維心 中目黒店"
]

def auto_login(page):
    page.goto("https://airregi.jp/CLP/view/top/")
    page.wait_for_timeout(2000)
    
    if "login" in page.url or page.locator("input[name='account']").is_visible() or page.get_by_role("button", name="ログイン").is_visible():
        print("・ログイン処理を自動実行中...")
        if page.get_by_role("button", name="ログイン").is_visible() and "login" not in page.url:
            page.get_by_role("button", name="ログイン").click()
            page.wait_for_timeout(2000)

        id_input = page.locator("input[name='account'], input[name='loginId'], input[type='text']").first
        if id_input.is_visible():
            id_input.fill(AIRREGI_ID)
        
        pass_input = page.locator("input[name='password'], input[type='password']").first
        if pass_input.is_visible():
            pass_input.fill(AIRREGI_PASS)

        login_btn = page.locator("button[type='submit'], input[type='submit'], button:has-text('ログイン')").first
        login_btn.click()
        page.wait_for_timeout(3000)

def select_store(page, store_name, is_first_store=False):
    print(f"\n--- 【店舗切り替え】「{store_name}」へ移動 ---")
    if page.locator("text='利用する店舗を選択'").is_visible() or page.locator("text='店舗を選択'").is_visible():
        page.wait_for_selector(f"text='{store_name}'", timeout=15000)
        page.locator(f"text='{store_name}'").first.click()
        page.wait_for_timeout(4000)
        return

    if is_first_store:
        try:
            page.wait_for_selector(f"text='{store_name}'", timeout=5000)
            return
        except Exception:
            pass

    store_trigger = page.locator("header button, header a, div[class*='shop'], div[class*='Header']").filter(has_text="店").first
    if not store_trigger.is_visible():
        store_trigger = page.get_by_text("三田おにぎり").or_(page.get_by_text("おにぎり維心")).first
    
    store_trigger.click()
    page.wait_for_timeout(1500)
    page.locator("text='店舗切り替え'").first.click()
    page.wait_for_timeout(2000)
    page.locator(f"text='{store_name}'").first.click()
    page.wait_for_timeout(3500)

def navigate_to_sales_page(page):
    sales_menu = page.locator("text='売上・分析'").first
    page.wait_for_selector("text='売上・分析'", timeout=15000)
    if not page.locator("text='日別売上'").first.is_visible():
        sales_menu.click()
        page.wait_for_timeout(1500)
    page.wait_for_selector("text='日別売上'", timeout=15000)
    page.locator("text='日別売上'").first.click()
    page.wait_for_timeout(2500)

def extract_sales_for_date(page, target_date):
    """指定日付の日付プルダウンを確実に操作してデータを取得する"""
    dt = datetime.strptime(target_date, "%Y-%m-%d")
    year_str, month_str, day_str = f"{dt.year}年", f"{dt.month}月", f"{dt.day}日"
    day_num_str = f"{dt.day}日"

    print(f"・ {target_date} ({year_str} {month_str} {day_num_str}) のデータを指定中...")

    # 集計対象が「時間別」になっているか確認・変更
    try:
        page.evaluate("""() => {
            const selects = Array.from(document.querySelectorAll('select'));
            for (let sel of selects) {
                for (let opt of sel.options) {
                    if (opt.text.includes('時間別')) {
                        sel.value = opt.value;
                        sel.dispatchEvent(new Event('change', { bubbles: true }));
                        break;
                    }
                }
            }
        }""")
    except Exception:
        pass

    # 日付のプルダウン操作（標準selectおよびカスタム要素へのフォールバック）
    success = page.evaluate("""({ yearStr, monthStr, dayNumStr }) => {
        const selects = Array.from(document.querySelectorAll('select'));
        let changed = false;
        for (let sel of selects) {
            for (let opt of sel.options) {
                let t = opt.text.trim();
                if (t === yearStr || t === monthStr || t === dayNumStr || t === String(parseInt(dayNumStr))) {
                    sel.value = opt.value;
                    sel.dispatchEvent(new Event('change', { bubbles: true }));
                    changed = true;
                }
            }
        }
        return changed;
    }""", {"yearStr": year_str, "monthStr": month_str, "dayNumStr": day_num_str})

    page.wait_for_timeout(1000)

    # 「表示する」ボタンをクリック
    btn = page.locator("button:has-text('表示する'), button:has-text('表示'), input[value*='表示']").first
    if btn.is_visible():
        btn.click(force=True)
        page.wait_for_timeout(3500)

    # テーブル行の抽出
    rows = page.locator("table tr").all()
    store_rows = []
    for row in rows:
        cells = row.locator("th, td").all()
        row_text = [cell.inner_text().strip().replace("\n", " ") for cell in cells]
        if any(row_text):
            store_rows.append(row_text)
    return store_rows

def main():
    if len(sys.argv) > 2:
        start_date_str = sys.argv[1]
        end_date_str = sys.argv[2]
    elif len(sys.argv) > 1:
        start_date_str = sys.argv[1]
        end_date_str = sys.argv[1]
    else:
        d = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        start_date_str, end_date_str = d, d

    start_dt = datetime.strptime(start_date_str, "%Y-%m-%d")
    end_dt = datetime.strptime(end_date_str, "%Y-%m-%d")
    
    date_list = []
    curr = start_dt
    while curr <= end_dt:
        date_list.append(curr.strftime("%Y-%m-%d"))
        curr += timedelta(days=1)

    print(f"=== 対象期間: {start_date_str} 〜 {end_date_str} ({len(date_list)}日間) ===")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=False,
            args=["--start-maximized"]
        )
        page = context.new_page()

        try:
            auto_login(page)

            for target_date in date_list:
                print(f"\n--- 【処理日】 {target_date} ---")
                all_stores_data = []

                for index, store_name in enumerate(TARGET_STORES):
                    is_first = (index == 0)
                    select_store(page, store_name, is_first_store=is_first)
                    
                    if index == 0 and target_date == date_list[0]:
                        navigate_to_sales_page(page)
                    else:
                        try:
                            page.locator("text='日別売上'").first.click()
                            page.wait_for_timeout(1500)
                        except Exception:
                            navigate_to_sales_page(page)

                    store_rows = extract_sales_for_date(page, target_date)
                    for r in store_rows:
                        all_stores_data.append([store_name] + r)

                if all_stores_data:
                    df = pd.DataFrame(all_stores_data)
                    csv_path = os.path.join(OUTPUT_DIR, f"4店舗_時間別売上_{target_date}.csv")
                    df.to_csv(csv_path, index=False, header=False, encoding="utf-8-sig")
                    print(f"[保存完了] {csv_path}")

        except Exception as e:
            print(f"[ERROR] 処理中にエラーが発生しました: {e}")
        finally:
            context.close()

if __name__ == "__main__":
    main()