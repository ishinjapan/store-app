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

def close_all_popups(page):
    """AirMATEや各種お知らせポップアップを閉じる。"""
    overlay_selector = "div[class*='overlay__Overlay']"

    def has_visible_overlay():
        return page.locator(overlay_selector).evaluate_all("""elements =>
            elements.some(element => {
                const rect = element.getBoundingClientRect();
                const style = getComputedStyle(element);
                return rect.width > 0 && rect.height > 0 &&
                    style.display !== 'none' && style.visibility !== 'hidden';
            })
        """)

    dismiss_selectors = [
        "button[aria-label='閉じる']",
        "button[aria-label='close' i]",
        "button[title='閉じる']",
        "button:has-text('閉じる')",
        "a:has-text('閉じる')",
    ]

    for _ in range(4):
        try:
            page.wait_for_timeout(500)
            clicked = False
            for selector in dismiss_selectors:
                buttons = page.locator(selector)
                for index in range(buttons.count()):
                    button = buttons.nth(index)
                    if button.is_visible():
                        button.click(timeout=1500)
                        page.wait_for_timeout(500)
                        clicked = True
                        break
                if clicked:
                    break

            page.keyboard.press("Escape")
            page.wait_for_timeout(500)
            if not has_visible_overlay():
                return

            # AirMATEの閉じるボタンがラベルのないアイコンの場合に備える。
            dismissed = page.evaluate("""() => {
                const visible = el => {
                    const r = el.getBoundingClientRect();
                    const s = getComputedStyle(el);
                    return r.width > 0 && r.height > 0 &&
                        s.display !== 'none' && s.visibility !== 'hidden';
                };
                const overlays = [...document.querySelectorAll(
                    'div[class*="overlay__Overlay"]'
                )].filter(visible);
                if (!overlays.length) return true;

                for (const overlay of overlays) {
                    const root = overlay.closest('[class*="gateway-local"]');
                    if (!root) continue;
                    const closeButton = [...root.querySelectorAll('button')]
                        .filter(visible)
                        .find(button => {
                            const r = button.getBoundingClientRect();
                            const label = [
                                button.innerText,
                                button.getAttribute('aria-label'),
                                button.getAttribute('title')
                            ].filter(Boolean).join(' ');
                            return /閉じる|close|×|✕/i.test(label) ||
                                (r.width < 80 && r.height < 80 &&
                                 button.querySelector('svg'));
                        });
                    if (closeButton) {
                        closeButton.click();
                        return true;
                    }
                }
                return false;
            }""")
            if dismissed:
                page.wait_for_timeout(700)
                if not has_visible_overlay():
                    print("・AirMATE / お知らせポップアップを閉じました。")
                    return
        except Exception as error:
            print(f"・ポップアップを閉じる処理を再試行します: {error}")

    if has_visible_overlay():
        raise RuntimeError(
            "AirMATEなどのポップアップが残っていて画面を操作できません。"
            "エラー画面のポップアップを閉じてから、もう一度実行してください。"
        )

def auto_login(page):
    if not AIRREGI_ID or not AIRREGI_PASS:
        print("[WARNING] AIRREGI_ID または AIRREGI_PASS が未設定です。")
        return

    print("・Airレジ トップページへアクセス中...")
    page.goto("https://airregi.jp/CLP/view/top/", wait_until="domcontentloaded")
    page.wait_for_timeout(3000)

    # ログインボタンのみがあるLP画面の場合はクリック
    login_btn_initial = page.locator("a:has-text('ログイン'), button:has-text('ログイン')").first
    pass_input_check = page.locator("input#password, input[name='password']").first
    if login_btn_initial.is_visible() and not pass_input_check.is_visible():
        print("・ログイン画面へ遷移します...")
        login_btn_initial.click()
        page.wait_for_timeout(3000)

    # 本物のID入力欄とパスワード入力欄を特定
    id_input = page.locator("input#account, input[name='account'], input#loginId, input[name='loginId'], input[name='command.common.loginId']").first
    pass_input = page.locator("input#password, input[name='password'], input[name='command.common.password']").first

    if id_input.is_visible() and pass_input.is_visible():
        print("・本物のID/パスワード欄に入力中...")
        id_input.click()
        id_input.fill("")
        id_input.fill(AIRREGI_ID)
        page.wait_for_timeout(500)

        pass_input.click()
        pass_input.fill("")
        pass_input.fill(AIRREGI_PASS)
        page.wait_for_timeout(500)

        submit_btn = page.locator("input[type='submit'], button[type='submit'], button:has-text('ログイン')").first
        print("・ログインボタンを押下...")
        submit_btn.click()
        
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            page.wait_for_timeout(5000)

    # ログイン直後のポップアップ消去
    close_all_popups(page)

def select_store(page, store_name, is_first_store=False):
    print(f"\n--- 【店舗切り替え】「{store_name}」へ移動 ---")
    page.wait_for_timeout(1500)
    close_all_popups(page)

    # 利用店舗選択画面
    if page.locator("text='利用する店舗を選択'").first.is_visible() or page.locator("text='店舗を選択'").first.is_visible():
        page.wait_for_selector(f"text='{store_name}'", timeout=15000)
        page.locator(f"text='{store_name}'").first.click()
        page.wait_for_timeout(4000)
        close_all_popups(page)
        return

    if is_first_store:
        try:
            page.wait_for_selector(f"text='{store_name}'", timeout=5000)
            close_all_popups(page)
            return
        except Exception:
            pass

    # ヘッダーの店舗切り替えボタン
    store_trigger = page.locator("header button, header a, div[class*='shop'], div[class*='Header']").filter(has_text="店").first
    if not store_trigger.is_visible():
        store_trigger = page.get_by_text("三田おにぎり").or_(page.get_by_text("おにぎり維心")).first

    store_trigger.click()
    page.wait_for_timeout(1500)
    page.locator("text='店舗切り替え'").first.click()
    page.wait_for_timeout(2000)
    page.locator(f"text='{store_name}'").first.click()
    page.wait_for_timeout(3500)
    close_all_popups(page)

def navigate_to_sales_page(page):
    close_all_popups(page)
    current_page_heading = page.locator("h1.title").filter(has_text="日別売上").first
    if current_page_heading.is_visible():
        return

    sales_menu = page.get_by_text("売上・分析", exact=True).first
    sales_menu.wait_for(state="visible", timeout=15000)
    sales_menu.click()
    page.wait_for_timeout(1000)
    close_all_popups(page)

    daily_sales_link = page.locator(
        "a, button, [role='menuitem']"
    ).filter(has_text="日別売上").first
    daily_sales_link.wait_for(state="visible", timeout=15000)
    daily_sales_link.click()
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(1000)
    close_all_popups(page)

def extract_sales_for_date(page, target_date):
    dt = datetime.strptime(target_date, "%Y-%m-%d")
    year_str, month_str, day_str = f"{dt.year}年", f"{dt.month}月", f"{dt.day}日"
    day_num_str = f"{dt.day}日"

    print(f"・ {target_date} ({year_str} {month_str} {day_num_str}) のデータを指定中...")
    close_all_popups(page)

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

    page.evaluate("""({ yearStr, monthStr, dayNumStr }) => {
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

    btn = page.locator("button:has-text('表示する'), button:has-text('表示'), input[value*='表示']").first
    if btn.is_visible():
        btn.click(force=True)
        page.wait_for_timeout(3500)

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
    if start_dt > end_dt:
        raise ValueError("開始日は終了日以前の日付を指定してください。")

    date_list = []
    curr = start_dt
    while curr <= end_dt:
        date_list.append(curr.strftime("%Y-%m-%d"))
        curr += timedelta(days=1)

    print(f"=== 対象期間: {start_date_str} 〜 {end_date_str} ({len(date_list)}日間) ===")

    screenshot_path = "error_screenshot.png"
    if os.path.exists(screenshot_path):
        try:
            os.remove(screenshot_path)
        except Exception:
            pass

    with sync_playwright() as p:
        user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        context = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=True,
            user_agent=user_agent,
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--disable-blink-features=AutomationControlled"
            ]
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
            try:
                page.screenshot(path=screenshot_path)
                print(f"エラー発生時の画面を保存しました: {screenshot_path}")
            except Exception:
                pass
            raise e
        finally:
            context.close()

if __name__ == "__main__":
    main()
