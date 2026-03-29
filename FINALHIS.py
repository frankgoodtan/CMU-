import tkinter as tk
from tkinter import messagebox, scrolledtext
import tkinter.ttk as ttk
import threading
import time
import re
import datetime
import os
import sys  # 新增這行
import ctypes
import random
import keyboard  
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

# 嘗試載入圖片處理套件
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# 全域變數：用來控制是否手動停止或暫停
stop_event = threading.Event()
pause_event = threading.Event()
pause_event.set() 

saved_initial_state = []

global_report_state = {
    "expected_map": {}, "completed_map": {}, "skipped_map": {},
    "new_consult_map": {}, "missing_chart_map": {}, "exist_record_map": {},
    "forced_draft_map": {}, "ghost_record_map": {}, "missing_today_map": {}
}
status_window = None  




def generate_current_report():
    expected_map = global_report_state["expected_map"]
    completed_map = global_report_state["completed_map"]
    new_consult_map = global_report_state["new_consult_map"]
    missing_chart_map = global_report_state["missing_chart_map"]
    missing_today_map = global_report_state["missing_today_map"]
    ghost_record_map = global_report_state["ghost_record_map"]
    exist_record_map = global_report_state["exist_record_map"]
    forced_draft_map = global_report_state["forced_draft_map"]

    expected_set = set(expected_map.keys())
    completed_set = set(completed_map.keys())
    new_consult_set = set(new_consult_map.keys())
    missing_chart_set = set(missing_chart_map.keys())
    missing_today_set = set(missing_today_map.keys())

    missing = expected_set - completed_set - new_consult_set - missing_chart_set - missing_today_set
    extra = completed_set - expected_set

    name_lookup = {c: info["name"] for c, info in expected_map.items()}
    name_lookup.update(completed_map)
    name_lookup.update(new_consult_map)
    name_lookup.update(missing_chart_map)
    name_lookup.update(ghost_record_map)
    name_lookup.update(missing_today_map)

    def fmt_list(chart_set):
        return [f"  • {c}（{name_lookup.get(c, '?')}）" for c in sorted(chart_set)]

    warn_lines = []
    if missing:
        warn_lines.append(f"【尚未處理 / 少寫 {len(missing)} 位】：")
        warn_lines += fmt_list(missing)
    if extra:
        warn_lines.append(f"\n【多寫 {len(extra)} 位】（完成了、但不在清單中）：")
        warn_lines += fmt_list(extra)
    if missing_today_map:
        warn_lines.append(f"\n【疑似未上傳門診 {len(missing_today_map)} 位】（跳過）：")
        warn_lines += fmt_list(missing_today_set)
    if new_consult_map:
        warn_lines.append(f"\n【疑似新會病歷 {len(new_consult_map)} 位】（無範圍內可選的ditto病歷）：")
        warn_lines += fmt_list(new_consult_set)
    if ghost_record_map:
        warn_lines.append(f"\n【幽靈病歷 {len(ghost_record_map)} 位】（無門診紀錄，隨機S並強制暫存）：")
        warn_lines += fmt_list(set(ghost_record_map.keys()))
    if missing_chart_map:
        warn_lines.append(f"\n【找不到的病歷號 {len(missing_chart_map)} 位】（清單中沒有/查無資料）：")
        warn_lines += fmt_list(missing_chart_set)
    if exist_record_map:
        warn_lines.append(f"\n【已有今日病歷 {len(exist_record_map)} 位】（跳過，視為完成）：")
        warn_lines += fmt_list(set(exist_record_map.keys()))
    if forced_draft_map:
        warn_lines.append(f"\n【醫師不符強制暫存 {len(forced_draft_map)} 位】（避免送錯醫師）：")
        warn_lines += fmt_list(set(forced_draft_map.keys()))

    warn_lines.append(f"\n【已完成清單（共 {len(completed_map)} 位）】：")
    if completed_map:
        warn_lines += [f"  • {c}（{n}）" for c, n in completed_map.items()]
    else:
        warn_lines.append("  (目前尚無完成資料)")

    has_warnings = bool(missing or extra or new_consult_map or missing_chart_map or forced_draft_map or ghost_record_map or missing_today_map)
    return "\n".join(warn_lines), has_warnings


def check_stop():
    if stop_event.is_set():
        raise Exception("【使用者手動停止】")

def safe_checkpoint(checkpoint_name=""):
    check_stop()
    if not pause_event.is_set():
        msg = f"    ⏸️ 系統已在安全段落 [{checkpoint_name}] 暫停，等待您按下繼續..." if checkpoint_name else "    ⏸️ 系統已暫停，等待繼續..."
        print(msg)
        pause_event.wait()
        print("    ▶️ 系統恢復執行...")
        check_stop() 


def update_ui_patient_processed(txt_names, txt_charts, txt_proc_charts, txt_proc_status, name, chart_no, status_msg):
    def _update():
        display_label = chart_no if chart_no else "[空白行]"
        txt_proc_charts.insert(tk.END, f"{display_label}\n")
        txt_proc_charts.see(tk.END)
        
        txt_proc_status.insert(tk.END, f"{status_msg}\n")
        txt_proc_status.see(tk.END)

        charts_lines = txt_charts.get("1.0", "end-1c").split('\n')
        names_lines = txt_names.get("1.0", "end-1c").split('\n')
        
        target_idx = -1
        tgt_val = chart_no.strip().lstrip("0")
        for i, c in enumerate(charts_lines):
            c_val = c.strip().lstrip("0")
            if c_val == tgt_val:
                target_idx = i
                break
        
        if target_idx != -1:
            charts_lines.pop(target_idx)
            txt_charts.delete("1.0", tk.END)
            if charts_lines:
                txt_charts.insert(tk.END, '\n'.join(charts_lines))
            
            if target_idx < len(names_lines):
                names_lines.pop(target_idx)
                txt_names.delete("1.0", tk.END)
                if names_lines:
                    txt_names.insert(tk.END, '\n'.join(names_lines))
                    
    root.after(0, _update)


def countdown_popup(seconds):
    proceed_event = threading.Event()
    try:
        top = tk.Toplevel()
        top.title("手動操作提示")
        window_width = 350
        window_height = 160
        screen_width = top.winfo_screenwidth()
        screen_height = top.winfo_screenheight()
        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int(screen_height - window_height - 150) 
        top.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")
        top.attributes("-topmost", True)  
        
        msg_label = tk.Label(top, text="請手動點選「允許網站自動執行」！", font=("Arial", 12, "bold"))
        msg_label.pack(pady=(15, 5))
        time_label = tk.Label(top, text=f"倒數 {seconds} 秒...", font=("Arial", 14, "bold"), fg="red")
        time_label.pack(pady=(0, 10))

        def on_confirm(event=None):
            proceed_event.set()
            try: top.destroy()
            except: pass

        confirm_btn = tk.Button(top, text="我已確認，繼續執行 (Enter)", font=("Arial", 11, "bold"), bg="#4CAF50", fg="white", command=on_confirm)
        confirm_btn.pack()
        top.bind("<Return>", on_confirm)
        top.focus_force()
        
        for i in range(seconds, 0, -1):
            if proceed_event.is_set() or stop_event.is_set(): break
            if not pause_event.is_set():
                try:
                    if top.winfo_exists():
                        time_label.config(text=f"⏸️ 暫停中... (剩餘 {i} 秒)")
                        top.update()
                except: pass
                pause_event.wait()
            try:
                if top.winfo_exists():
                    time_label.config(text=f"倒數 {i} 秒...")
                    top.update()
            except: break 
            proceed_event.wait(1)
            
        if not proceed_event.is_set():
            try:
                if top.winfo_exists(): top.destroy()
            except: pass
        check_stop()
    except Exception as e:
        if "使用者手動停止" in str(e): raise e
        time.sleep(seconds)


def auto_add_missing_patients(driver, wait, missing_charts, expected_map):
    print("\n    ▶ 進入 [自動新增病歷號] 流程...")
    
    try:
        print("      → 切換至病歷號搜尋模式")
        try:
            chart_radio_label = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//label[contains(normalize-space(), '病歷號')]")
            ))
            driver.execute_script("arguments[0].click();", chart_radio_label)
        except Exception as e:
            print(f"        ⚠️ 無法點擊病歷號 Radio 按鈕，嘗試備用方案：{e}")
            try:
                radio_inputs = driver.find_elements(By.XPATH, "//input[@type='radio' and @name='admPtType']")
                if len(radio_inputs) >= 2:
                    driver.execute_script("arguments[0].click();", radio_inputs[1])
            except: pass
                
        time.sleep(1) 
        
        chart_input = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//input[@maxlength='10' and @type='text']")
        ))

        for c in missing_charts:
            check_stop()
            name = expected_map.get(c, {}).get("name", "未知")
            print(f"\n      - 開始搜尋病歷號：{c} ({name})")
            
            try: chart_input.click()
            except: driver.execute_script("arguments[0].click();", chart_input)
                
            time.sleep(0.2)
            chart_input.send_keys(Keys.CONTROL + 'a')
            time.sleep(0.2)
            chart_input.send_keys(Keys.BACKSPACE)
            time.sleep(0.2)
            chart_input.send_keys(c)
            time.sleep(0.5)
            
            try:
                search_btn = driver.find_element(By.XPATH, "//button[.//span[contains(., '查詢')]]")
                driver.execute_script("arguments[0].click();", search_btn)
            except Exception as e:
                print(f"        ⚠️ 找不到查詢按鈕：{e}")
                continue
                
            time.sleep(2) 
            
            try:
                rows = driver.find_elements(By.CSS_SELECTOR, "tbody.p-datatable-tbody tr")
            except:
                rows = []
                
            valid_rows = []
            
            for row in rows:
                try:
                    tds = row.find_elements(By.TAG_NAME, "td")
                    if len(tds) < 11: continue
                        
                    row_chart = tds[3].text.strip().lstrip("0")
                    if row_chart != c.lstrip("0"): continue
                        
                    category = tds[10].text.strip()
                    date_str = tds[9].text.strip()
                    
                    if category == "住院":
                        try:
                            dt = datetime.datetime.strptime(date_str, "%Y-%m-%d %H:%M")
                        except:
                            dt = datetime.datetime.min 
                            
                        checkbox = tds[0].find_element(By.CSS_SELECTOR, "div.p-checkbox-box")
                        valid_rows.append({
                            "element": row,
                            "date": dt,
                            "checkbox": checkbox
                        })
                except Exception:
                    continue
            
            if not valid_rows:
                print(f"        ❌ 找不到符合「住院」條件的資料 (可能出院、格式錯誤或查無此人)")
                continue
                
            valid_rows.sort(key=lambda x: x["date"], reverse=True)
            target_record = valid_rows[0]
            
            checkbox_el = target_record["checkbox"]
            if "p-highlight" not in checkbox_el.get_attribute("class"):
                driver.execute_script("arguments[0].click();", checkbox_el)
                print(f"        ✅ 成功勾選 (判定為住院，最新就醫日期: {target_record['date'].strftime('%Y-%m-%d %H:%M')})")
            else:
                print(f"        ✅ 發現該病患已被勾選，略過操作。")
                
            time.sleep(1)

        print("\n    → 所有自動新增動作完畢，準備點擊「我的清單」返回...")
        try:
            my_list_label = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//label[contains(normalize-space(), '我的清單')]")
            ))
            driver.execute_script("arguments[0].click();", my_list_label)
        except Exception as inner_e:
            print(f"      ⚠️ 無法點擊我的清單：{inner_e}")
                
        time.sleep(2)
        print("    ✅ 已切換回我的清單！")

    except Exception as e:
        if "使用者手動停止" in str(e): raise e
        print(f"    ⚠️ 自動新增流程發生非預期錯誤：{e}")


def prompt_missing_patients(missing_charts, expected_map):
    proceed_event = threading.Event()
    action_result = {"status": "skip"}
    
    missing_info = [f"• {c} ({expected_map[c]['name']})" for c in missing_charts]
    missing_text = "\n".join(missing_info)

    def show_dialog():
        top = tk.Toplevel(root)
        top.title("⚠️ 找不到部分病歷號")
        top.attributes("-topmost", True)
        
        window_width = 520
        window_height = 370
        screen_width = top.winfo_screenwidth()
        screen_height = top.winfo_screenheight()
        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int(screen_height - window_height - 150)
        top.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")
        
        tk.Label(top, text="以下欲處理之病歷號未出現在「已勾選清單」中：", font=("Arial", 11, "bold"), fg="#c0392b").pack(pady=(10, 5))
        
        txt = scrolledtext.ScrolledText(top, width=50, height=8, font=("Arial", 10))
        txt.pack(padx=10, pady=5)
        txt.insert(tk.END, missing_text)
        txt.config(state="disabled")
        
        tk.Label(top, text="建議先手動新增，或點選「自動新增」由系統代勞。\n若未新增，執行時將自動跳過這些病歷號。\n\n請選擇下方操作：", font=("Arial", 10), justify="center").pack(pady=5)
        
        btn_frame = tk.Frame(top)
        btn_frame.pack(pady=5)
        
        is_paused = [False]
        
        def on_pause_resume():
            if not is_paused[0]:
                is_paused[0] = True
                btn_pause.config(text="我已手動新增完成，繼續", bg="#4CAF50")
                btn_auto.config(state="disabled")
                btn_skip.config(state="disabled", text="跳過")
                lbl_status.config(text="⏸️ 腳本已暫停，請在網頁上手動新增病人後點擊上方按鈕", fg="#2196F3", font=("Arial", 10, "bold"))
            else:
                action_result["status"] = "recheck"
                proceed_event.set()
                try: top.destroy()
                except: pass
                
        def on_auto_add():
            action_result["status"] = "auto_add"
            proceed_event.set()
            try: top.destroy()
            except: pass
        
        def on_skip():
            action_result["status"] = "skip"
            proceed_event.set()
            try: top.destroy()
            except: pass

        btn_pause = tk.Button(btn_frame, text="手動新增 (暫停進程)", bg="#FF9800", fg="white", font=("Arial", 10, "bold"), command=on_pause_resume)
        btn_pause.pack(side="left", padx=5)
        
        btn_auto = tk.Button(btn_frame, text="自動新增", bg="#9C27B0", fg="white", font=("Arial", 10, "bold"), command=on_auto_add)
        btn_auto.pack(side="left", padx=5)
        
        btn_skip = tk.Button(btn_frame, text="跳過 (剩餘 10 秒)", font=("Arial", 10), command=on_skip)
        btn_skip.pack(side="left", padx=5)
        
        lbl_status = tk.Label(top, text="", font=("Arial", 9))
        lbl_status.pack(pady=(0, 5))
        
        def countdown(left):
            if proceed_event.is_set(): return
            if is_paused[0]: return 
            
            if left > 0:
                try:
                    btn_skip.config(text=f"跳過 (剩餘 {left} 秒)")
                    top.after(1000, countdown, left - 1)
                except: pass
            else:
                on_skip()
                
        countdown(10)

    root.after(0, show_dialog)
    proceed_event.wait()
    return action_result["status"]


def prompt_all_patients_found():
    proceed_event = threading.Event()
    def show_dialog():
        top = tk.Toplevel(root)
        top.title("✅ 檢查完畢")
        top.attributes("-topmost", True)
        window_width = 350
        window_height = 120
        screen_width = top.winfo_screenwidth()
        screen_height = top.winfo_screenheight()
        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int(screen_height - window_height - 150)
        top.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")
        tk.Label(top, text="欲處理之病歷號皆已在病人清單中！", font=("Arial", 12, "bold"), fg="#4CAF50").pack(pady=(20, 10))
        lbl_time = tk.Label(top, text="將於 2 秒後自動繼續...", font=("Arial", 10))
        lbl_time.pack()
        
        def countdown(left):
            if stop_event.is_set():
                proceed_event.set()
                try: top.destroy()
                except: pass
                return
            if left > 0:
                try:
                    lbl_time.config(text=f"將於 {left} 秒後自動繼續...")
                    top.after(1000, countdown, left - 1)
                except: pass
            else:
                proceed_event.set()
                try: top.destroy()
                except: pass
        countdown(2)
    root.after(0, show_dialog)
    proceed_event.wait()


def parse_patient_list(name_text, chart_text):
    names  = [l.strip() for l in name_text.splitlines()]
    charts = [l.strip() for l in chart_text.splitlines()]
    max_len = max(len(names), len(charts))
    names  += [""] * (max_len - len(names))
    charts += [""] * (max_len - len(charts))
    patients = []
    for name, chart_no in zip(names, charts):
        if chart_no:
            normalized_chart = chart_no.lstrip("0")
            patients.append({"name": name, "chart_no": normalized_chart})
        else:
            patients.append({"name": name, "chart_no": ""}) 
    return patients


def scroll_to_load_all_rows(driver):
    prev_count = 0
    retries = 0
    max_retries = 3
    while True:
        check_stop()
        rows = driver.find_elements(By.CSS_SELECTOR, "tbody.p-datatable-tbody tr")
        current_count = len(rows)
        if current_count == prev_count:
            retries += 1
            if retries >= max_retries: break 
            time.sleep(1.0) 
            continue
        retries = 0
        prev_count = current_count
        if rows:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", rows[-1])
        time.sleep(0.5)
    print(f"    📋 表格共載入 {prev_count} 列")
    return prev_count


def find_row_by_chart_no(driver, chart_no):
    rows = driver.find_elements(By.CSS_SELECTOR, "tbody.p-datatable-tbody tr")
    for row in rows:
        check_stop()
        try:
            tds = row.find_elements(By.TAG_NAME, "td")
            if len(tds) < 4: continue
            row_chart = tds[3].text.strip().lstrip("0")
            if row_chart == chart_no:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", row)
                time.sleep(0.3)
                return row
        except Exception: continue
    return None


def keep_system_awake():
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000002 | 0x00000001)
        print("    🛡️ 已啟動防休眠機制，執行期間螢幕與系統將保持喚醒。")
    except Exception as e:
        print(f"    ⚠️ 防休眠機制啟動失敗（可能非 Windows 系統）：{e}")


def return_to_patient_list(driver, wait):
    print("    → 執行「回到病人清單函數」...")
    try:
        hamburger_menu = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@class, 'collapse-menu-link') and contains(@class, 'glyphicon-menu-hamburger')]")
        ))
        driver.execute_script("arguments[0].click();", hamburger_menu)
        time.sleep(1)
        
        patient_select_link = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@class, 'al-sidebar-list-link') and contains(., '病人選取')]")
        ))
        driver.execute_script("arguments[0].click();", patient_select_link)
        time.sleep(2)
        
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody.p-datatable-tbody tr")))
        print("    ✅ 已成功回到病人清單頁面！")
    except Exception as e:
        print(f"    ⚠️ 回到病人清單失敗，將嘗試使用備用方法：{e}")
        try:
            driver.back()
        except:
            pass


def final_countdown_and_close(driver, report_msg):
    def show_final_ui():
        top = tk.Toplevel(root)
        top.title("自動化執行完畢 / 系統提示")
        window_width = 480
        window_height = 500 
        screen_width = top.winfo_screenwidth()
        screen_height = top.winfo_screenheight()
        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int(screen_height - window_height - 150)
        top.geometry(f"{window_width}x{window_height}+{x_cordinate}+{y_cordinate}")
        top.attributes("-topmost", True)
        
        tk.Label(top, text="✅ 執行結果與提示：", font=("Arial", 12, "bold"), fg="#4CAF50").pack(pady=(10, 5))
        txt = scrolledtext.ScrolledText(top, width=60, height=18, font=("Arial", 10))
        txt.pack(padx=10, pady=5)
        txt.insert(tk.END, report_msg)
        txt.config(state="disabled")
        
        lbl_time = tk.Label(top, text="將於 300 秒後自動關閉網頁與程式...", font=("Arial", 12, "bold"), fg="red")
        lbl_time.pack(pady=5)
        is_cancelled = False
        
        def force_close():
            if is_cancelled: return 
            try: 
                if driver: driver.quit()  
            except: pass
            os._exit(0)         
            
        def continue_execution():
            nonlocal is_cancelled
            is_cancelled = True 
            try: top.destroy()   
            except: pass
            print("    ▶ 已取消自動關閉，將繼續保持網頁與程式開啟。")
            
        top.protocol("WM_DELETE_WINDOW", force_close)
        btn_frame = tk.Frame(top)
        btn_frame.pack(pady=(5, 10))
        tk.Button(btn_frame, text="繼續執行 (不關閉)", font=("Arial", 10, "bold"), bg="#2196F3", fg="white", command=continue_execution).pack(side="left", padx=10)
        tk.Button(btn_frame, text="立刻關閉", font=("Arial", 10), command=force_close).pack(side="left", padx=10)
        
        def countdown(left):
            if is_cancelled: return 
            if left > 0:
                try:
                    lbl_time.config(text=f"將於 {left} 秒後自動關閉網頁與程式...")
                    top.after(1000, countdown, left-1)
                except: pass
            else: force_close()
                
        countdown(300) 
    root.after(0, show_final_ui)


def step_6_submit_or_draft(driver, wait, action_mode, opd_dr_name, physician_code):
    safe_checkpoint("準備執行最終送件/暫存")
    print(f"\n    ▶ 進入第六步驟：準備執行【{action_mode}】...")
    try:
        btn_xpath = f"//button[contains(@class, 'btn') and contains(normalize-space(), '{action_mode}')]"
        action_btn = wait.until(EC.element_to_be_clickable((By.XPATH, btn_xpath)))
        driver.execute_script("arguments[0].click();", action_btn)
        print(f"    ✅ 成功點擊「{action_mode}」首要按鈕！")
        
        if action_mode == "送件":
            print("    → 等待第一個檢核視窗 (縮寫)...")
            confirm1_xpath = "//button[contains(@class, 'p-confirm-dialog-accept') and .//span[text()='確定']]"
            confirm1_btn = wait.until(EC.element_to_be_clickable((By.XPATH, confirm1_xpath)))
            driver.execute_script("arguments[0].click();", confirm1_btn)
            
            print("    → 等待第二個檢核視窗...")
            confirm2_xpath = "//button[contains(@class, 'p-button-success') and .//span[text()='確定']]"
            confirm2_btn = wait.until(EC.element_to_be_clickable((By.XPATH, confirm2_xpath)))
            driver.execute_script("arguments[0].click();", confirm2_btn)
            
            print("    → 視窗確認完畢，等待 2 秒準備帶回病歷...")
            time.sleep(2)
            check_stop()
            bring_back_xpath = "//button[contains(@class, 'p-button-info') and .//span[text()='病歷帶回']]"
            bring_back_btn = wait.until(EC.element_to_be_clickable((By.XPATH, bring_back_xpath)))
            driver.execute_script("arguments[0].click();", bring_back_btn)
            
            print("    → 進入主治醫師確認頁面，準備核對資料...")
            time.sleep(2) 
            check_stop()
            
            vs_dropdown_input_xpath = "//span[contains(@class, 'drOption')]//p-dropdown//input[contains(@class, 'p-dropdown-label')]"
            vs_input = wait.until(EC.presence_of_element_located((By.XPATH, vs_dropdown_input_xpath)))
            current_vs_val = vs_input.get_attribute("value")
            
            if opd_dr_name in current_vs_val and physician_code in current_vs_val:
                print("    ✅ 主治醫師姓名與代碼核對無誤！準備最終送件...")
                final_submit_xpath = "//button[contains(@class, 'p-button-success') and .//span[text()='送件']]"
                final_submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, final_submit_xpath)))
                driver.execute_script("arguments[0].click();", final_submit_btn)
                time.sleep(2)
            else:
                print(f"    ⚠️ 主治醫師資料不符！(門診紀錄醫師: {opd_dr_name}, 目標代碼: {physician_code})")
                print("    → 準備清除並重新輸入主治醫師代碼...")
                vs_input.click()
                time.sleep(0.5)
                vs_input.send_keys(Keys.CONTROL + 'a')
                time.sleep(0.2)
                vs_input.send_keys(Keys.BACKSPACE)
                time.sleep(0.5)
                vs_input.send_keys(physician_code)
                time.sleep(0.5)
                vs_input.send_keys(Keys.ENTER)
                time.sleep(1.5)
                check_stop()
                final_submit_xpath = "//button[contains(@class, 'p-button-success') and .//span[text()='送件']]"
                final_submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, final_submit_xpath)))
                driver.execute_script("arguments[0].click();", final_submit_btn)
                time.sleep(2)
                
            print("    → 等待最終的「確定」按鈕...")
            final_ok_xpath = "//button[contains(@class, 'p-confirm-dialog-accept') and .//span[text()='確定']]"
            final_ok_btn = wait.until(EC.element_to_be_clickable((By.XPATH, final_ok_xpath)))
            driver.execute_script("arguments[0].click();", final_ok_btn)
            print("    ✅ 已點選最終「確定」，完成所有送件流程！")
            time.sleep(2)
            
        else:
            print("    → 等待暫存確認視窗...")
            draft_confirm_xpath = "//button[contains(@class, 'p-confirm-dialog-accept') and .//span[text()='確定']]"
            draft_confirm_btn = wait.until(EC.element_to_be_clickable((By.XPATH, draft_confirm_xpath)))
            driver.execute_script("arguments[0].click();", draft_confirm_btn)
            time.sleep(2) 
            
    except Exception as e:
        if "使用者手動停止" in str(e): raise e
        print(f"    ⚠️ 第六步驟 ({action_mode}) 發生錯誤：{e}")
        raise e


def step_5_add_new_record(driver, wait, chart_no, patient_name, old_subjective, old_objective, ditto_o, ditto_a, old_plan, record_date, action_mode, opd_dr_name, physician_code):
    safe_checkpoint("準備自動填寫病歷資料")
    print(f"\n    ▶ 進入第五步驟：開始新增病歷 — {chart_no}（{patient_name}）")
    
    try:
        print("    → 尋找問題列表，準備點擊第一個診斷的淡藍色方塊...")
        first_diagnosis_td = wait.until(EC.presence_of_element_located(
            (By.XPATH, "(//tr[contains(@class, 'p-selectable-row')]//td[contains(@class, 'dt-status')])[1]")
        ))
        driver.execute_script("arguments[0].click();", first_diagnosis_td)
        time.sleep(1)
        
        formatted_subjective = old_subjective
        if old_subjective:
            y, m, d = record_date.split("-")
            roc_year = int(y) - 1911
            target_date_str = f"({roc_year}/{m}/{d})"
            target_date_str_full = f"（{roc_year}/{m}/{d}）"
            if target_date_str in old_subjective:
                formatted_subjective = old_subjective.split(target_date_str)[-1].strip()
            elif target_date_str_full in old_subjective:
                formatted_subjective = old_subjective.split(target_date_str_full)[-1].strip()

        formatted_plan = ""
        if old_plan:
            old_plan = re.sub(r'\(\d{2,3}/\d{2}/\d{2}\)', '', old_plan)
            old_plan = re.sub(r'(患者符合中醫健保.*?計畫\s*\(.*?\)).*?(?=#)', r'\1\n', old_plan, flags=re.DOTALL)
            plan_text = old_plan.replace("#", "\n#").replace("Timeout", "\nTimeout")
            plan_lines = [line.strip() for line in plan_text.splitlines() if line.strip()]
            formatted_plan = "\n".join(plan_lines)

        s_box = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='B1']/following-sibling::textarea[1]")))
        o_box = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='B2']/following-sibling::textarea[1]")))
        a_box = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='B3']/following-sibling::textarea[1]")))
        p_box = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='B4']/following-sibling::textarea[1]")))
        
        if formatted_subjective:
            s_box.clear()
            s_box.send_keys(formatted_subjective)
            
        if ditto_a:
            a_box.clear()
            a_box.send_keys(ditto_a)
            
        if formatted_plan:
            p_box.clear()
            p_box.send_keys(formatted_plan)
            
        try:
            calendar_input = wait.until(EC.presence_of_element_located((By.XPATH, "//p-calendar//input[contains(@class, 'p-inputtext')]")))
            calendar_input.click()
            time.sleep(0.5)
            current_datetime = calendar_input.get_attribute("value")
            
            if current_datetime:
                parts = current_datetime.split(" ")
                current_date_val = parts[0]
                current_time_val = parts[1] if len(parts) > 1 else ""

                if current_date_val != record_date:
                    new_datetime = f"{record_date} {current_time_val}".strip()
                    calendar_input.send_keys(Keys.CONTROL + 'a')
                    time.sleep(0.2)
                    calendar_input.send_keys(Keys.BACKSPACE)
                    time.sleep(0.5) 
                    calendar_input.send_keys(new_datetime)
                    time.sleep(0.5)
                    calendar_input.send_keys(Keys.ENTER) 
                    time.sleep(1)
                else:
                    calendar_input.send_keys(Keys.ESCAPE)
        except Exception as e:
            print(f"    ⚠️ 校正病歷完成時間發生錯誤 (繼續執行)：{e}")

        target_o = ditto_o if (ditto_o and ditto_o.strip()) else old_objective
        needs_new_reports = False
        
        if target_o:
            lines = target_o.splitlines()
            new_o_lines = []
            skipping_vitals = False
            for line in lines:
                line_lower = line.lower()
                if any(k in line for k in ["望診", "聞診", "舌診", "切診"]):
                    skipping_vitals = False
                if "vital signs" in line_lower:
                    skipping_vitals = True
                    continue
                if skipping_vitals:
                    if any(k in line_lower for k in ["spo2", "rr", "fio2", "peep", "pcv"]): continue
                    if re.search(r'\d+\.?\d*\s*/\s*\d+\.?\d*\s*=', line): continue
                    skipping_vitals = False
                if not skipping_vitals:
                    new_o_lines.append(line)
            
            target_o = "\n".join(new_o_lines).strip()
            
            if any(keyword in target_o for keyword in ["報告時間", "檢驗單", "影像部檢查"]):
                needs_new_reports = True
                match = re.search(r'(望診：.*?切診：[^\n]*)', target_o, flags=re.DOTALL)
                if match:
                    target_o = match.group(1)
                else:
                    extracted = [l for l in target_o.splitlines() if any(k in l for k in ["望診：", "聞診：", "舌診：", "切診："])]
                    target_o = "\n".join(extracted)

        if target_o is not None:
            o_box.clear()
            o_box.send_keys(target_o)
            o_box.send_keys("\n") 
            o_box.click() 
            time.sleep(0.5)
            
            try:
                o_box.click()
                time.sleep(0.5)
                def click_sidebar_menu(target_name):
                    menu_items = driver.find_elements(By.CSS_SELECTOR, "div.groupDetailSelect")
                    if not menu_items:
                        progress_img = driver.find_element(By.CSS_SELECTOR, "img[title='病程記錄']")
                        driver.execute_script("arguments[0].click();", progress_img)
                        time.sleep(1)
                        menu_items = driver.find_elements(By.CSS_SELECTOR, "div.groupDetailSelect")
                    for item in menu_items:
                        item_text = item.get_attribute("innerText") or item.text
                        if target_name in item_text:
                            driver.execute_script("arguments[0].click();", item)
                            return True
                    return False

                click_sidebar_menu("生命徵象")
                time.sleep(1.5) 
                
                if needs_new_reports:
                    o_box.click() 
                    if click_sidebar_menu("檢驗報告"):
                        time.sleep(1)
                        confirm_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'p-button-success') and .//span[text()='確定']]")))
                        driver.execute_script("arguments[0].click();", confirm_btn)
                        time.sleep(1.5)
                    if click_sidebar_menu("檢查報告"):
                        time.sleep(1)
                        confirm_btn2 = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'p-button-success') and .//span[text()='確定']]")))
                        driver.execute_script("arguments[0].click();", confirm_btn2)
                        time.sleep(1.5)

            except Exception as e:
                print(f"    ⚠️ 自動點擊帶入生命徵象/報告時發生狀況：{e}")

        print("    ✅ 成功完成所有 SOAP 資料填寫與帶入！")
        step_6_submit_or_draft(driver, wait, action_mode, opd_dr_name, physician_code)
        
    except Exception as e:
        if "使用者手動停止" in str(e): raise e
        raise e


def step_4_write_record(driver, wait, chart_no, patient_name, physician, physician_code, record_date, action_mode):
    label = f"{chart_no}（{patient_name}）"
    print(f"    ⏳ 開始執行病歷操作 — {label}")
    
    js_extract_text = """
        var text = '';
        for (var i = 0; i < arguments[0].childNodes.length; i++) {
            if (arguments[0].childNodes[i].nodeType === Node.TEXT_NODE) {
                text += arguments[0].childNodes[i].textContent + '\\n';
            }
        }
        return text;
    """
    
    mismatch_detected = False
    is_ghost_record = False 
    old_subjective = ""
    old_objective = ""
    old_plan = ""
    
    try:
        time.sleep(2)
        safe_checkpoint("準備抓取門診紀錄")
        
        long_wait = WebDriverWait(driver, 20)
        med_record_menu = long_wait.until(EC.presence_of_element_located((By.XPATH, "//img[@title='醫藥囑相關紀錄']")))
        driver.execute_script("arguments[0].click();", med_record_menu)
        time.sleep(1) 
        
        short_wait = WebDriverWait(driver, 5) 
        try:
            opd_record_link = short_wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'groupDetailSelect') and contains(., '門診-就醫病歷')]")))
            driver.execute_script("arguments[0].click();", opd_record_link)
        except Exception: pass 
        time.sleep(2) 
        
        try:
            short_wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tr.p-selectable-row")))
            opd_rows = driver.find_elements(By.CSS_SELECTOR, "tr.p-selectable-row")
        except Exception: opd_rows = []
        
        target_record_row = None
        
        if not opd_rows:
            is_ghost_record = True
        else:
            today_xpath = f"//tr[contains(@class, 'p-selectable-row')][.//td[1]//p[contains(., '{record_date}')]]"
            today_records = driver.find_elements(By.XPATH, today_xpath)
            
            if today_records:
                target_record_row = today_records[0]
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target_record_row)
                time.sleep(0.5)
                try: target_record_row.find_element(By.XPATH, "./td[1]").click()
                except Exception: driver.execute_script("arguments[0].click();", target_record_row)
                time.sleep(2)
                check_stop()
                
                opd_dr_val = ""
                try: opd_dr_val = target_record_row.find_element(By.XPATH, "./td[3]").text.strip()
                except Exception: pass

                if opd_dr_val and physician not in opd_dr_val:
                    mismatch_detected = True
                    action_mode = "暫存" 
                    pause_event.clear()
                    
                    def show_mismatch_warning():
                        top = tk.Toplevel(root)
                        top.title("醫師不符提示")
                        top.attributes("-topmost", True)
                        tk.Label(top, text=f"注意！門診紀錄醫師 ({opd_dr_val}) 與輸入醫師 ({physician}) 不符！\n\n系統已暫停，將於 5 秒後自動以「暫存」模式繼續執行。\n若要立刻繼續，請點擊下方按鈕。", padx=20, pady=20, font=("Arial", 10)).pack()
                        def force_resume():
                            pause_event.set()
                            try: top.destroy()
                            except: pass
                        tk.Button(top, text="立刻繼續執行 (強制暫存)", command=force_resume, bg="#2196F3", fg="white", font=("Arial", 10, "bold")).pack(pady=10)
                        def auto_close(left):
                            if pause_event.is_set():
                                try: top.destroy()
                                except: pass
                                return
                            if left <= 0:
                                pause_event.set()
                                try: top.destroy()
                                except: pass
                            else: top.after(1000, auto_close, left-1)
                        top.after(1000, auto_close, 5)

                    root.after(0, show_mismatch_warning)
                    root.after(0, lambda: btn_pause.config(text="繼續執行 (Alt+S)", bg="#2196F3"))
                    pause_event.wait(5.0) 
                    pause_event.set()     
                    root.after(0, lambda: btn_pause.config(text="暫停執行 (Alt+S)", bg="#FF9800"))
                    check_stop()

                driver.execute_script("arguments[0].click();", target_record_row)
                time.sleep(2) 
                
                try:
                    subjective_xpath = "//fieldset[.//legend[contains(., '主觀') or contains(., 'Subjective') or contains(., '會診目的') or contains(., '主訴')]]/div/div[contains(@class, 'p-fieldset-content')]"
                    subjective_el = driver.find_element(By.XPATH, subjective_xpath)
                    old_subjective = driver.execute_script(js_extract_text, subjective_el)
                    old_subjective = '\n'.join([line.strip() for line in old_subjective.splitlines() if line.strip()])
                except Exception: pass

                try:
                    objective_xpath = "//fieldset[.//legend[contains(., '理學檢查(Objective)')]]/div/div[contains(@class, 'p-fieldset-content')]"
                    objective_el = driver.find_element(By.XPATH, objective_xpath)
                    old_objective = driver.execute_script(js_extract_text, objective_el)
                    old_objective = '\n'.join([line.strip() for line in old_subjective.splitlines() if line.strip()])
                except Exception: pass

                try:
                    plan_xpath = "//fieldset[.//legend[contains(., '治療計畫(Plan)')]]/div/div[contains(@class, 'p-fieldset-content')]"
                    plan_el = driver.find_element(By.XPATH, plan_xpath)
                    old_plan = plan_el.text.strip()
                except Exception: pass
            else:
                raise Exception("疑似今日未上傳門診紀錄")
        
        try:
            close_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'pi-times') and contains(@class, 'p-dialog-header-close-icon')]")))
            driver.execute_script("arguments[0].click();", close_btn)
            time.sleep(2)
        except: pass
        
        safe_checkpoint("準備切換至 DITTO 操作")
        tcm_record_link = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(normalize-space(), '中醫病程記錄')]")))
        driver.execute_script("arguments[0].click();", tcm_record_link)
        time.sleep(2)
        
        ditto_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(normalize-space(), 'Ditto')]")))
        driver.execute_script("arguments[0].click();", ditto_btn)
        time.sleep(2)
        
        wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'p-datatable-scrollable-body')]//tbody/tr")))
        ditto_rows = driver.find_elements(By.XPATH, "//div[contains(@class, 'p-datatable-scrollable-body')]//tbody/tr")
        
        target_ditto_row = None
        current_date_obj = datetime.datetime.strptime(record_date, "%Y-%m-%d").date()
        past_7_days = [(current_date_obj - datetime.timedelta(days=i)).strftime("%Y-%m-%d") for i in range(1, 8)]
        
        for row in ditto_rows:
            try:
                date_text = row.find_element(By.XPATH, "./td[1]").text.strip()
                if date_text and date_text.split(" ")[0] == record_date:
                    try:
                        ditto_close_btn = driver.find_element(By.XPATH, "//span[contains(@class, 'pi-times') and contains(@class, 'p-dialog-header-close-icon')]")
                        driver.execute_script("arguments[0].click();", ditto_close_btn)
                        time.sleep(1)
                    except: pass
                    raise Exception("今日病歷已存在")
            except Exception as e:
                if str(e) == "今日病歷已存在" or "使用者手動停止" in str(e): raise e
        
        for target_date in past_7_days:
            for row in ditto_rows:
                try:
                    date_text = row.find_element(By.XPATH, "./td[1]").text.strip()
                    if date_text and date_text.split(" ")[0] == target_date:
                        target_ditto_row = row
                        break 
                except Exception: pass
            if target_ditto_row: break 

        if not target_ditto_row: raise Exception("疑似新會病歷")

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target_ditto_row)
        time.sleep(0.5)
        target_td = target_ditto_row.find_element(By.XPATH, "./td[1]")
        try: target_td.click()
        except Exception: driver.execute_script("arguments[0].click();", target_td)
        
        try:
            time.sleep(1)
            WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.TAG_NAME, "cmuh-article-list")))
            def extract_ditto_section(driver, section_title):
                try:
                    xpath = f"//h4[contains(@class, 'title') and contains(normalize-space(.), '{section_title}')]/parent::header/following-sibling::div[contains(@class, 'div-for-copy')][1]//pre"
                    return driver.find_element(By.XPATH, xpath).text
                except Exception: return ""

            ditto_o = extract_ditto_section(driver, "O")
            ditto_a = extract_ditto_section(driver, "A")
            ditto_p = extract_ditto_section(driver, "P")
        except Exception:
            ditto_o, ditto_a, ditto_p = "", "", ""

        ditto_close_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@class, 'pi-times') and contains(@class, 'p-dialog-header-close-icon')]")))
        driver.execute_script("arguments[0].click();", ditto_close_btn)
        time.sleep(2)
        
        if is_ghost_record:
            s_options = [
                ["納差", "胃口稍改善"],
                ["腹脹", "腹脹稍改善"],
                ["精神意識尚可", "精神意識稍差"],
                ["大便溏", "大便溏稍改善"],
                ["覺四肢稍無力", "四肢稍有力"]
            ]
            chosen_pairs = random.sample(s_options, 3)
            ghost_s_parts = [random.choice(pair) for pair in chosen_pairs]
            old_subjective = "，".join(ghost_s_parts) + "。"
            old_plan = ditto_p
            action_mode = "暫存"
            
        step_5_add_new_record(
            driver, wait, chart_no, patient_name,
            old_subjective, old_objective, ditto_o, ditto_a, old_plan, record_date, action_mode, 
            physician, physician_code
        )
        
        if is_ghost_record: return "ghost_record"
        return "forced_draft" if mismatch_detected else "success"
        
    except Exception as e:
        if "使用者手動停止" not in str(e): print(f"    ❌ 執行病歷操作發生狀況：{e}")
        raise e


# ==========================================
# 第三步驟：掃描已勾選病人，依群組逐一處理
# ==========================================
def step_3_process_patients(driver, wait, groups, action_mode, auto_add_flag=False, auto_uncheck_flag=False):
    print("\n--- 進入第三步驟：處理已勾選病人 ---")

    try:
        check_stop()
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody.p-datatable-tbody tr")))
        print("✅ 表格初始載入完成")
    except Exception as e:
        if "使用者手動停止" in str(e): raise e
        final_countdown_and_close(driver, f"找不到病人列表表格：\n{e}")
        return

    scroll_to_load_all_rows(driver)

    global global_report_state
    for k in global_report_state: global_report_state[k].clear()

    expected_map      = global_report_state["expected_map"]
    completed_map     = global_report_state["completed_map"]
    skipped_map       = global_report_state["skipped_map"]
    new_consult_map   = global_report_state["new_consult_map"]
    missing_chart_map = global_report_state["missing_chart_map"]
    exist_record_map  = global_report_state["exist_record_map"]
    forced_draft_map  = global_report_state["forced_draft_map"]
    ghost_record_map  = global_report_state["ghost_record_map"]
    missing_today_map = global_report_state["missing_today_map"]

    for g in groups:
        for p in g["patients"]:
            expected_map[p["chart_no"]] = {
                "name":           p["name"],
                "physician":      g["physician"],
                "physician_code": g["code"],
            }

    print(f"📋 期望處理病人（共 {len(expected_map)} 位，跨 {len(groups)} 個群組）：")

    def get_checked_chart_nos():
        rows = driver.find_elements(By.CSS_SELECTOR, "tbody.p-datatable-tbody tr")
        checked = []
        for row in rows:
            try:
                checkbox_box = row.find_element(By.CSS_SELECTOR, ".p-checkbox-box")
                if "p-highlight" not in checkbox_box.get_attribute("class"): continue
                tds = row.find_elements(By.TAG_NAME, "td")
                if len(tds) < 4: continue
                chart_no = tds[3].text.strip().lstrip("0")
                name     = tds[2].text.strip()
                if chart_no: checked.append((name, chart_no))
            except Exception: continue
        return checked

    checked_patients = get_checked_chart_nos()
    print(f"\n☑️  表格中已勾選病人（共 {len(checked_patients)} 位）")

    if not checked_patients:
        messagebox.showwarning("注意", "表格中找不到任何已勾選的病人！")
        return

    checked_set = {chart_no for _, chart_no in checked_patients}
    missing_from_checked = set(expected_map.keys()) - checked_set
    missing_from_checked = {c for c in missing_from_checked if c and c.strip() != ""}

    if missing_from_checked:
        print(f"\n    ⚠️ 發現 {len(missing_from_checked)} 位欲處理的病歷號不在勾選清單中...")
        
        if auto_add_flag:
            print("    ⚙️ 已勾選「自動新增未在清單中的病歷號」，準備自動執行...")
            auto_add_missing_patients(driver, wait, missing_from_checked, expected_map)
            
            print("    🔄 自動新增完畢，準備重新載入表格...")
            scroll_to_load_all_rows(driver)
            checked_patients = get_checked_chart_nos()
            checked_set = {chart_no for _, chart_no in checked_patients}
            print(f"    ☑️ 重新整理後，表格中已勾選病人共 {len(checked_patients)} 位")
        else:
            print("    💬 準備彈出提示視窗等待使用者選擇...")
            status = prompt_missing_patients(missing_from_checked, expected_map)
            
            if status == "auto_add":
                print("    ⚙️ 使用者點擊「自動新增」，準備由系統自動執行...")
                auto_add_missing_patients(driver, wait, missing_from_checked, expected_map)
                
                print("    🔄 自動新增完畢，準備重新載入表格...")
                scroll_to_load_all_rows(driver)
                checked_patients = get_checked_chart_nos()
                checked_set = {chart_no for _, chart_no in checked_patients}
                print(f"    ☑️ 重新整理後，表格中已勾選病人共 {len(checked_patients)} 位")
                
            elif status == "recheck":
                print("    🔄 使用者已手動新增，準備重新載入表格...")
                scroll_to_load_all_rows(driver)
                checked_patients = get_checked_chart_nos()
                checked_set = {chart_no for _, chart_no in checked_patients}
                print(f"    ☑️ 重新整理後，表格中已勾選病人共 {len(checked_patients)} 位")
            else:
                print("    ⏩ 使用者選擇跳過或已倒數結束，將略過這些病人...")
    else:
        print(f"\n    ✅ 欲處理之病歷號皆已在勾選清單中，彈出提示並於 2 秒後繼續...")
        prompt_all_patients_found()


    for g_idx, g in enumerate(groups):
        physician       = g["physician"]
        physician_code  = g["code"]
        record_date     = g["date"].get().strip()
        txt_names       = g["txt_names"]
        txt_charts      = g["txt_charts"]
        txt_proc_charts = g["txt_proc_charts"]
        txt_proc_status = g["txt_proc_status"]
        
        print(f"\n{'='*52}")
        print(f"▶▶ 群組 {g_idx+1}｜主治醫師：{physician}（{physician_code}）｜目標日期：{record_date}")
        print(f"{'='*52}")

        for p in g["patients"]:
            safe_checkpoint(f"準備處理病人：{p.get('name', '未知名稱')}")
            
            chart_no     = p["chart_no"]
            patient_name = p["name"]
            label        = f"{chart_no}（{patient_name}）" if patient_name else (chart_no if chart_no else "[空白病歷號]")

            if not chart_no or chart_no.strip() == "":
                missing_chart_map["[空白病歷號]"] = patient_name
                update_ui_patient_processed(txt_names, txt_charts, txt_proc_charts, txt_proc_status, patient_name, chart_no, "❌ 格式空白")
                continue

            if chart_no not in checked_set:
                skipped_map[chart_no] = patient_name
                update_ui_patient_processed(txt_names, txt_charts, txt_proc_charts, txt_proc_status, patient_name, chart_no, "⚠️ 未勾選跳過")
                continue

            print(f"\n  ▶ 處理：{label}")
            try:
                scroll_to_load_all_rows(driver)
                target_row = find_row_by_chart_no(driver, chart_no)

                if target_row is None:
                    missing_chart_map[chart_no] = patient_name
                    update_ui_patient_processed(txt_names, txt_charts, txt_proc_charts, txt_proc_status, patient_name, chart_no, "❌ 找不到病歷")
                    continue

                current_url = driver.current_url
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target_row)
                time.sleep(0.5)
                actions = ActionChains(driver)
                actions.double_click(target_row).perform()
                time.sleep(2) 
                
                try: wait.until(EC.url_changes(current_url))
                except Exception: wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.al-sidebar-list-link")))

                status = step_4_write_record(driver, wait, chart_no, patient_name, physician, physician_code, record_date, action_mode)

                if status == "ghost_record":
                    ghost_record_map[chart_no] = patient_name
                    completed_map[chart_no] = patient_name
                    update_ui_patient_processed(txt_names, txt_charts, txt_proc_charts, txt_proc_status, patient_name, chart_no, "✅ 幽靈強制暫存")
                elif status == "forced_draft":
                    forced_draft_map[chart_no] = patient_name
                    completed_map[chart_no] = patient_name 
                    update_ui_patient_processed(txt_names, txt_charts, txt_proc_charts, txt_proc_status, patient_name, chart_no, "✅ 醫師不符暫存")
                else:
                    completed_map[chart_no] = patient_name
                    update_ui_patient_processed(txt_names, txt_charts, txt_proc_charts, txt_proc_status, patient_name, chart_no, "✅ 完成")
                
                return_to_patient_list(driver, wait)

            except Exception as e:
                error_msg = str(e)
                if "使用者手動停止" in error_msg: raise e
                elif "疑似新會病歷" in error_msg:
                    new_consult_map[chart_no] = patient_name
                    update_ui_patient_processed(txt_names, txt_charts, txt_proc_charts, txt_proc_status, patient_name, chart_no, "🌟 疑似新會病歷")
                elif "疑似今日未上傳門診紀錄" in error_msg:
                    missing_today_map[chart_no] = patient_name
                    update_ui_patient_processed(txt_names, txt_charts, txt_proc_charts, txt_proc_status, patient_name, chart_no, "⚠️ 未上傳門診")
                elif "今日病歷已存在" in error_msg:
                    exist_record_map[chart_no] = patient_name 
                    completed_map[chart_no] = patient_name    
                    update_ui_patient_processed(txt_names, txt_charts, txt_proc_charts, txt_proc_status, patient_name, chart_no, "✅ 今日病歷已存")
                else:
                    skipped_map[chart_no] = patient_name
                    update_ui_patient_processed(txt_names, txt_charts, txt_proc_charts, txt_proc_status, patient_name, chart_no, "❌ 處理發生錯誤")

                return_to_patient_list(driver, wait)

    # ==========================================
    # 新增：病歷完成後自動反勾選處理
    # ==========================================
    if auto_uncheck_flag and completed_map:
        print("\n    → 準備執行「病歷完成後自動反勾選」...")
        try:
            print("      → 執行「病人選取」以整理並重新載入清單...")
            # 這裡呼叫一次 return_to_patient_list 確保畫面回到最乾淨的「病人選取」表格
            return_to_patient_list(driver, wait)
            time.sleep(2)
            scroll_to_load_all_rows(driver)
            
            rows = driver.find_elements(By.CSS_SELECTOR, "tbody.p-datatable-tbody tr")
            uncheck_count = 0
            
            for row in rows:
                check_stop()
                try:
                    tds = row.find_elements(By.TAG_NAME, "td")
                    if len(tds) < 4: continue
                    row_chart = tds[3].text.strip().lstrip("0")
                    
                    if row_chart in completed_map:
                        checkbox_box = row.find_element(By.CSS_SELECTOR, ".p-checkbox-box")
                        if "p-highlight" in checkbox_box.get_attribute("class"):
                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", row)
                            time.sleep(0.3)
                            driver.execute_script("arguments[0].click();", checkbox_box)
                            uncheck_count += 1
                            time.sleep(0.2)
                except Exception as e:
                    continue
                    
            print(f"    ✅ 自動反勾選完成，共取消勾選 {uncheck_count} 位已處理之病人！")
        except Exception as e:
            print(f"    ⚠️ 自動反勾選發生錯誤：{e}")


    msg, has_warnings = generate_current_report()
    if has_warnings: final_msg = "⚠️ 執行完畢 (含例外狀況)\n\n" + msg
    else: final_msg = "✅ 所有病人處理完畢，清單核對一致！\n\n" + msg
        
    print("\n=== 第三步驟完成，準備自動關閉 ===")
    root.after(0, lambda: btn_start.config(state="normal", text="開始執行"))
    root.after(0, lambda: btn_pause.config(state="disabled", text="暫停執行 (Alt+S)", bg="#FF9800"))
    root.after(0, lambda: btn_stop.config(state="disabled", text="停止並重置 (Alt+A)"))
    final_countdown_and_close(driver, final_msg)


def step_2_next_actions(driver, wait, groups, action_mode, auto_add_flag, auto_uncheck_flag):
    print("\n--- 進入第二步驟：開啟住院醫囑系統 ---")
    try:
        check_stop()
        programs_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '程式集')]")))
        driver.execute_script("arguments[0].click();", programs_btn)
        time.sleep(1)

        inpatient_sys_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '住院醫囑系統')]")))
        driver.execute_script("arguments[0].click();", inpatient_sys_btn)
        
        countdown_popup(10)
        
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])

        step_3_process_patients(driver, wait, groups, action_mode, auto_add_flag, auto_uncheck_flag)

    except Exception as e:
        if "使用者手動停止" in str(e): raise e
        final_countdown_and_close(driver, f"第二步驟發生錯誤：\n{e}")


def step_1_login(emp_id, emp_pwd, groups, action_mode, auto_add_flag, auto_uncheck_flag):
    print("\n--- 進入第一步驟：系統登入 ---")
    keep_system_awake()
    options = webdriver.EdgeOptions()
    prefs = {
        "protocol_handler.excluded_schemes.runpcallmainp": False,
        "custom_handlers.enabled": True
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Edge(options=options)
    wait = WebDriverWait(driver, 10)

    try:
        check_stop()
        driver.get("https://his.cmuh.org.tw/webapp/login/")

        auth_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '(原)帳號密碼')]")))
        auth_btn.click()

        userid_input = wait.until(EC.presence_of_element_located((By.ID, "userid")))
        userid_input.clear()
        userid_input.send_keys(emp_id)
        password_input = driver.find_element(By.ID, "password")
        password_input.clear()
        password_input.send_keys(emp_pwd)

        login_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(text(), '登 入')]")))
        login_btn.click()

        step_2_next_actions(driver, wait, groups, action_mode, auto_add_flag, auto_uncheck_flag)

    except Exception as e:
        if "使用者手動停止" in str(e):
            if 'driver' in locals() and driver:
                try: driver.quit()
                except: pass
            return
            
        driver_to_pass = driver if 'driver' in locals() else None
        final_countdown_and_close(driver_to_pass, f"系統發生錯誤：\n{e}")
        root.after(0, lambda: btn_start.config(state="normal", text="開始執行"))
        root.after(0, lambda: btn_pause.config(state="disabled", text="暫停執行 (Alt+S)", bg="#FF9800"))
        root.after(0, lambda: btn_stop.config(state="disabled", text="停止並重置 (Alt+A)"))


# ==========================================
# 介面功能與啟動
# ==========================================

def show_disclaimer(): # 移除 is_first_run 參數
    top = tk.Toplevel(root)
    top.title("免責與使用說明")
    top.attributes("-topmost", True)
    top.minsize(700, 750)
    
    # 圖片區塊
    if HAS_PIL:
        try:
            def resource_path(relative_path):
                try:
                    base_path = sys._MEIPASS
                except Exception:
                    base_path = os.path.dirname(os.path.abspath(__file__))
                return os.path.join(base_path, relative_path)
            
            img_path = resource_path("CMU黑奴.png")

            if os.path.exists(img_path):
                img = Image.open(img_path)
                img.thumbnail((650, 450))  
                photo = ImageTk.PhotoImage(img)
                lbl_img = tk.Label(top, image=photo)
                lbl_img.image = photo 
                lbl_img.pack(pady=(15, 5))
            else:
                tk.Label(top, text=f"[圖片檔案找不到]\n路徑: {img_path}", fg="red", font=("Arial", 10)).pack(pady=10)
        except Exception as e:
            tk.Label(top, text=f"[圖片載入失敗: {e}]", fg="red").pack(pady=10)
    else:
        tk.Label(top, text="[未安裝 Pillow 套件]", fg="orange").pack(pady=10)

    tk.Label(top, text="Copyright and Developed by PBCM-38 譚皓宇", font=("Arial", 10), fg="gray").pack(pady=(0, 10))

    desc_text = (
        "1. 本系統本為複製並撰寫針灸科病程使用，不具主動外傳病人資料或令裝置感染病毒之代碼功能。\n\n"
        "2. 本系統撰寫之內容與流程(SOAP)基於譚哥於2026/03/29時之前的經驗後實作，如後續有變動，再請自行修正。\n\n"
        "3. 本系統執行時為使用爬蟲分析病歷系統之資料結構並自動化，使用過程中盡可能保持畫面穩定，若有位移、醫院網站系統更新等干擾網頁結構資料之載入與分析的狀況則可能會有BUG，再請自行修正。\n\n"
        "4. 本系統只是輔助撰寫病歷，大家還是要為自己負責的病歷做確認喔!\n\n"
        "5. 以上如果操作過程有問題可以再聯繫 BY 譚哥"
    )
    
    lbl_desc = tk.Label(top, text=desc_text, justify="left", wraplength=650, font=("Arial", 12, "bold"))
    lbl_desc.pack(padx=20, pady=10)

    # 統一為關閉按鈕
    btn_close = tk.Button(top, text="我已了解並關閉", font=("Arial", 14, "bold"), bg="#4CAF50", fg="white", command=top.destroy)
    btn_close.pack(pady=15)
        
#def check_first_run():
#    file_path = os.path.join(get_app_dir(), "his_first_run_done.txt")
#    if not os.path.exists(file_path):
#        root.after(500, lambda: show_disclaimer(is_first_run=True))


def update_mandatory_stars():
    try:
        mode = action_var.get()
        for grp in group_frames:
            if mode == "送件":
                grp["star_vs"].config(text="*", fg="#c0392b", font=("Arial", 11, "bold"))
                grp["star_code"].config(text="*", fg="#c0392b", font=("Arial", 11, "bold"))
            else:
                grp["star_vs"].config(text="(暫存可空白)", fg="gray", font=("Arial", 9, "normal"))
                grp["star_code"].config(text="(暫存可空白)", fg="gray", font=("Arial", 9, "normal"))
    except: pass

def delete_group_by_frame(target_frame):
    global group_frames
    for g in group_frames:
        if g["frame"] == target_frame:
            g["frame"].destroy()
            group_frames.remove(g)
            break
            
    if len(group_frames) < 5:
        btn_add.config(state="normal", text="＋ 新增群組")
        
    groups_container.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))

def add_group():
    idx = len(group_frames)
    if idx >= 5: return

    frame = tk.LabelFrame(groups_container, text=f"群組 {idx + 1}", padx=8, pady=6)
    frame.pack(fill="x", pady=4, padx=4)
    
    if idx > 0:
        btn_del = tk.Button(frame, text="🗑️ 刪除此群組", bg="#ff4d4d", fg="white", font=("Arial", 9),
                            command=lambda f=frame: delete_group_by_frame(f))
        btn_del.pack(anchor="ne", pady=(0, 5))

    row_vs = tk.Frame(frame)
    row_vs.pack(fill="x", pady=2)
    tk.Label(row_vs, text="主治醫師：", width=10, anchor="e").pack(side="left")
    ent_physician = tk.Entry(row_vs, width=20)
    ent_physician.pack(side="left", padx=4)
    lbl_star_vs = tk.Label(row_vs, text="*", fg="#c0392b", font=("Arial", 11, "bold"))
    lbl_star_vs.pack(side="left")

    row_code = tk.Frame(frame)
    row_code.pack(fill="x", pady=2)
    tk.Label(row_code, text="醫師代碼：", width=10, anchor="e").pack(side="left")
    ent_code = tk.Entry(row_code, width=20)
    ent_code.pack(side="left", padx=4)
    lbl_star_code = tk.Label(row_code, text="*", fg="#c0392b", font=("Arial", 11, "bold"))
    lbl_star_code.pack(side="left")

    row_date = tk.Frame(frame)
    row_date.pack(fill="x", pady=2)
    
    tk.Label(row_date, text="病歷年份：", width=10, anchor="e").pack(side="left")
    ent_year = tk.Entry(row_date, width=6)
    ent_year.insert(0, datetime.date.today().strftime("%Y"))
    ent_year.pack(side="left", padx=(4, 2))
    
    tk.Label(row_date, text="日期(MM-DD)：").pack(side="left")
    ent_date = tk.Entry(row_date, width=8)
    ent_date.insert(0, datetime.date.today().strftime("%m-%d"))
    ent_date.pack(side="left", padx=(2, 4))
    tk.Label(row_date, text="*", fg="#c0392b", font=("Arial", 11, "bold")).pack(side="left")

    tk.Label(frame, text="病患清單（每行一位，姓名與病歷號行數需對應）", fg="gray", font=("Arial", 8)).pack(anchor="w", pady=(4, 1))

    row_pt = tk.Frame(frame)
    row_pt.pack(fill="x", pady=2)

    col_name = tk.Frame(row_pt)
    col_name.pack(side="left", padx=(0, 4))
    tk.Label(col_name, text="姓名", font=("Arial", 9, "bold")).pack(anchor="w")
    txt_names = scrolledtext.ScrolledText(col_name, width=10, height=6)
    txt_names.pack()

    col_chart = tk.Frame(row_pt)
    col_chart.pack(side="left", padx=(0, 4))
    tk.Label(col_chart, text="待處理患者病歷號，必填 *", font=("Arial", 9, "bold"), fg="#c0392b").pack(anchor="w")
    txt_charts = scrolledtext.ScrolledText(col_chart, width=20, height=6)
    txt_charts.pack()

    col_proc_chart = tk.Frame(row_pt)
    col_proc_chart.pack(side="left", padx=(4, 4))
    tk.Label(col_proc_chart, text="已處理病歷號", font=("Arial", 9, "bold"), fg="#4CAF50").pack(anchor="w")
    txt_proc_charts = scrolledtext.ScrolledText(col_proc_chart, width=14, height=6)
    txt_proc_charts.pack()

    col_proc_status = tk.Frame(row_pt)
    col_proc_status.pack(side="left")
    tk.Label(col_proc_status, text="處理狀態", font=("Arial", 9, "bold"), fg="#4CAF50").pack(anchor="w")
    txt_proc_status = scrolledtext.ScrolledText(col_proc_status, width=16, height=6)
    txt_proc_status.pack()
    
    # 同步捲動
    def sync_scroll_charts(f, l):
        txt_proc_charts.vbar.set(f, l)
        txt_proc_status.yview_moveto(f)

    def sync_scroll_status(f, l):
        txt_proc_status.vbar.set(f, l)
        txt_proc_charts.yview_moveto(f)

    txt_proc_charts.config(yscrollcommand=sync_scroll_charts)
    txt_proc_status.config(yscrollcommand=sync_scroll_status)

    group_frames.append({
        "frame":         frame,
        "physician":     ent_physician,
        "code":          ent_code,
        "year":          ent_year,   
        "date":          ent_date,
        "names":         txt_names,
        "charts":        txt_charts,
        "proc_charts":   txt_proc_charts, 
        "proc_status":   txt_proc_status, 
        "star_vs":       lbl_star_vs,    
        "star_code":     lbl_star_code,  
    })

    if len(group_frames) >= 5:
        btn_add.config(state="disabled", text="已達五群上限")

    groups_container.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))
    update_mandatory_stars()

def start_automation():
    emp_id  = entry_id.get().strip()
    emp_pwd = entry_pwd.get().strip()
    action_mode = action_var.get()
    
    auto_add_flag = auto_add_var.get()
    auto_uncheck_flag = auto_uncheck_var.get()

    if not emp_id or not emp_pwd:
        messagebox.showwarning("格式錯誤", "員工代號與密碼為必填欄位！")
        return

    global saved_initial_state
    saved_initial_state.clear()
    for grp in group_frames:
        saved_initial_state.append({
            "grp": grp,
            "names": grp["names"].get("1.0", tk.END).strip(),
            "charts": grp["charts"].get("1.0", tk.END).strip()
        })

    groups = []
    for i, grp in enumerate(group_frames):
        physician   = grp["physician"].get().strip()
        code        = grp["code"].get().strip()
        year_str    = grp["year"].get().strip()
        date_raw    = grp["date"].get().strip().replace("/", "-")
        
        try:
            m, d = date_raw.split("-")
            date_str = f"{int(m):02d}-{int(d):02d}"
        except: date_str = date_raw 
            
        record_date = f"{year_str}-{date_str}" 
        name_text   = grp["names"].get("1.0", tk.END).strip()
        chart_text  = grp["charts"].get("1.0", tk.END).strip()

        grp["proc_charts"].delete("1.0", tk.END)
        grp["proc_status"].delete("1.0", tk.END)

        if chart_text or physician or code:
            if action_mode == "送件":
                if not physician or not code:
                    messagebox.showwarning("格式錯誤", f"群組 {i+1}：選擇「送件」時，主治醫師與醫師代碼不得為空！")
                    return
            if not record_date:
                messagebox.showwarning("格式錯誤", f"群組 {i+1}：病歷日期不得為空！")
                return
            if not chart_text:
                messagebox.showwarning("格式錯誤", f"群組 {i+1}：病患清單（病歷號）不得為空！")
                return

            patients = parse_patient_list(name_text, chart_text)
            groups.append({
                "physician": physician, 
                "code": code, 
                "patients": patients,
                "date": tk.StringVar(value=record_date),
                "txt_names": grp["names"],
                "txt_charts": grp["charts"],
                "txt_proc_charts": grp["proc_charts"],
                "txt_proc_status": grp["proc_status"]
            })

    if not groups:
        messagebox.showwarning("格式錯誤", "請至少填寫一組資料！")
        return

    stop_event.clear()
    pause_event.set()
    btn_start.config(state="disabled", text="執行中...")
    btn_pause.config(state="normal", text="暫停執行 (Alt+S)", bg="#FF9800")
    btn_stop.config(state="normal", text="停止並重置 (Alt+A)")
    root.update_idletasks() 

    threading.Thread(target=step_1_login, args=(emp_id, emp_pwd, groups, action_mode, auto_add_flag, auto_uncheck_flag), daemon=True).start()

def toggle_pause():
    global status_window
    if pause_event.is_set():
        pause_event.clear()
        btn_pause.config(text="繼續執行 (Alt+S)", bg="#2196F3")
        root.update_idletasks() 
        print("\n    ⏳ 已收到暫停指令，將在到達下一個安全檢查點時暫停...")
        
        if status_window and status_window.winfo_exists():
            status_window.lift() 
        else:
            status_window = tk.Toplevel(root)
            status_window.title("⏸️ 暫停中 - 目前病人處理進度")
            status_window.geometry("450x550")
            status_window.attributes("-topmost", True)
            
            tk.Label(status_window, text="目前進度統計", font=("Arial", 12, "bold")).pack(pady=10)
            txt = scrolledtext.ScrolledText(status_window, width=55, height=25, font=("Arial", 10))
            txt.pack(padx=10, pady=5, fill="both", expand=True)
            
            report_msg, _ = generate_current_report()
            if not report_msg.strip() or "共 0 位" in report_msg:
                report_msg = "目前尚未開始處理病人資料..."
                
            txt.insert(tk.END, report_msg)
            txt.config(state="disabled")
            
            tk.Button(status_window, text="關閉此視窗", command=status_window.destroy, bg="#eee").pack(pady=10)
    else:
        pause_event.set()
        btn_pause.config(text="暫停執行 (Alt+S)", bg="#FF9800")
        root.update_idletasks() 
        
        if status_window and status_window.winfo_exists():
            status_window.destroy()

def stop_automation():
    print("\n🛑 已收到停止指令，正在安全中止程序並重置視窗...")
    stop_event.set()
    pause_event.set() 
    btn_pause.config(state="disabled")
    btn_stop.config(state="disabled", text="停止中...")
    root.update_idletasks() 
    
    def reset_ui():
        if saved_initial_state:
            for state in saved_initial_state:
                grp = state["grp"]
                grp["names"].delete("1.0", tk.END)
                if state["names"]: grp["names"].insert(tk.END, state["names"])
                
                grp["charts"].delete("1.0", tk.END)
                if state["charts"]: grp["charts"].insert(tk.END, state["charts"])
                    
                grp["proc_charts"].delete("1.0", tk.END)
                grp["proc_status"].delete("1.0", tk.END)
        
        btn_start.config(state="normal", text="開始執行")
        btn_pause.config(state="disabled", text="暫停執行 (Alt+S)", bg="#FF9800")
        btn_stop.config(state="disabled", text="停止並重置 (Alt+A)")
        print("    ✅ 已成功重置為最一開始執行前的參數狀態。")
        
    root.after(1500, reset_ui)


# ==========================================
# 建立主視窗
# ==========================================
root = tk.Tk()
root.title("HIS 病歷自動化系統")
root.geometry("800x850") 
root.configure(padx=20, pady=20)

# 修改頂部區域加入「使用說明按鈕」
top_header_frame = tk.Frame(root)
top_header_frame.pack(fill="x", pady=(0, 15))
tk.Label(top_header_frame, text="系統登入參數", font=("Arial", 14, "bold")).pack(side="left")
btn_info = tk.Button(top_header_frame, text="📖 使用說明", font=("Arial", 9), command=show_disclaimer)
btn_info.pack(side="right")

frame_form = tk.Frame(root)
frame_form.pack(fill="x")

tk.Label(frame_form, text="員工代號:").grid(row=0, column=0, sticky="e", pady=5)
entry_id = tk.Entry(frame_form, width=20)
entry_id.grid(row=0, column=1, padx=10)

tk.Label(frame_form, text="系統密碼:").grid(row=1, column=0, sticky="e", pady=5)
entry_pwd = tk.Entry(frame_form, width=20, show="*")
entry_pwd.grid(row=1, column=1, padx=10)

ttk.Separator(root, orient="horizontal").pack(fill="x", pady=12)
tk.Label(root, text="病患群組設定", font=("Arial", 14, "bold")).pack(anchor="w", pady=(0, 6))

outer_frame = tk.Frame(root)
outer_frame.pack(fill="both", expand=True)

canvas = tk.Canvas(outer_frame, highlightthickness=0)
scrollbar = ttk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True)

groups_container = tk.Frame(canvas)
canvas_window = canvas.create_window((0, 0), window=groups_container, anchor="nw")

def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

def on_canvas_configure(event):
    canvas.itemconfig(canvas_window, width=event.width)

groups_container.bind("<Configure>", on_frame_configure)
canvas.bind("<Configure>", on_canvas_configure)
canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

group_frames = []
add_group()

btn_add = tk.Button(root, text="＋ 新增群組", font=("Arial", 10), bg="#2196F3", fg="white", command=add_group)
btn_add.pack(fill="x", pady=(8, 4))

frame_action = tk.Frame(root)
frame_action.pack(fill="x", pady=(10, 5))
tk.Label(frame_action, text="最後執行動作：", font=("Arial", 10, "bold")).pack(side="left")

action_var = tk.StringVar(value="送件") 
tk.Radiobutton(frame_action, text="送件 (預設)", variable=action_var, value="送件", command=update_mandatory_stars).pack(side="left")
tk.Radiobutton(frame_action, text="暫存", variable=action_var, value="暫存", command=update_mandatory_stars).pack(side="left")

# 新增：自動新增未在清單中的病歷號 選項
auto_add_var = tk.BooleanVar(value=False)
tk.Checkbutton(frame_action, text="自動新增未在清單中的病歷號", variable=auto_add_var, font=("Arial", 10)).pack(side="left", padx=(15, 0))

# 新增：病歷完成後自動反勾選 選項
auto_uncheck_var = tk.BooleanVar(value=False)
tk.Checkbutton(frame_action, text="病歷完成後自動反勾選", variable=auto_uncheck_var, font=("Arial", 10)).pack(side="left", padx=(15, 0))

btn_frame = tk.Frame(root)
btn_frame.pack(fill="x", pady=(0, 5))

btn_start = tk.Button(btn_frame, text="開始執行", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", command=start_automation)
btn_start.pack(side="left", fill="x", expand=True, padx=(0, 5))

btn_pause = tk.Button(btn_frame, text="暫停執行 (Alt+S)", font=("Arial", 12, "bold"), bg="#FF9800", fg="white", state="disabled", command=toggle_pause)
btn_pause.pack(side="left", fill="x", expand=True, padx=5)

btn_stop = tk.Button(btn_frame, text="停止並重置 (Alt+A)", font=("Arial", 12, "bold"), bg="#f44336", fg="white", state="disabled", command=stop_automation)
btn_stop.pack(side="left", fill="x", expand=True, padx=(5, 0))


def safe_hk_start():
    if btn_start['state'] == 'normal': start_automation()

def safe_hk_pause():
    if btn_pause['state'] == 'normal': toggle_pause()

def safe_hk_stop():
    if btn_stop['state'] == 'normal': stop_automation()

try:
    keyboard.add_hotkey('alt+s', lambda: root.after(0, safe_hk_pause))
    keyboard.add_hotkey('alt+a', lambda: root.after(0, safe_hk_stop))
    print("⌨️ 全域快捷鍵已註冊成功！(Alt+S: 暫停/繼續, Alt+A: 停止並重置)")
except Exception as e:
    print(f"⚠️ 快捷鍵註冊失敗，請確認是否已安裝 keyboard 套件: {e}")

# 在主介面最下方加入版權宣告小字
lbl_copyright = tk.Label(root, text="Copyright and Developed by PBCM-38 譚皓宇", font=("Arial", 8), fg="gray")
lbl_copyright.pack(side="bottom", pady=5)


root.mainloop()
