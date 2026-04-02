# Author: 譚皓宇
import tkinter as tk
from tkinter import messagebox, scrolledtext
import tkinter.ttk as ttk
import threading
import time
import re
import datetime
import os
import sys
import ctypes
import random
import urllib.request
import json
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
    "forced_draft_map": {}, "ghost_record_map": {}, "missing_today_map": {},
    "chinese_herb_map": {}, "discharged_map": {}
}
status_window = None
discord_webhook_url = ""   # 進階設定：Discord Webhook URL
notify_per_patient = False  # 進階設定：每完成一位病人就發 Discord 通知

# ==========================================
# 確保瀏覽器視窗在前景的輔助函式
# ==========================================
def ensure_window_focus(driver):
    try:
        handles = driver.window_handles
        if handles:
            current = driver.current_window_handle
            if current not in handles:
                driver.switch_to.window(handles[-1])
        driver.execute_script("""
            window.focus();
            document.dispatchEvent(new MouseEvent('mousemove', {bubbles: true}));
        """)
        if driver.get_window_rect().get('height', 600) < 100:
            driver.maximize_window()
            time.sleep(0.5)
    except Exception as e:
        print(f"    ⚠️ ensure_window_focus 發生非預期狀況（略過）：{e}")

def send_discord_notification(report_msg):
    """執行完畢後透過 Discord Webhook 發送結算報告"""
    global discord_webhook_url
    url = discord_webhook_url.strip() if discord_webhook_url else ""
    if not url:
        return
    try:
        content = f"✅ **HIS 病歷自動化完成！**\n`\n{report_msg[:1850]}\n`"
        payload = json.dumps({"content": content}).encode("utf-8")
        
        req = urllib.request.Request(
            url,
            data=payload,
            headers={
                "Content-Type": "application/json",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            },
            method="POST"
        )
        urllib.request.urlopen(req, timeout=10)
        print("    ✅ Discord 通知已發送！")
    except Exception as e:
        print(f"    ⚠️ Discord 通知發送失敗：{e}")

def send_discord_progress(title, msg=""):
    """發送即時進度通知（暫停/群組完成/全部完成）"""
    global discord_webhook_url
    url = discord_webhook_url.strip() if discord_webhook_url else ""
    if not url:
        return
    try:
        report_snippet, _ = generate_current_report()
        combined = f"**{title}**\n"
        if msg:
            combined += f"{msg}\n"
        combined += f"```\n{report_snippet[:1500]}\n```"
        payload = json.dumps({"content": combined}).encode("utf-8")
        req = urllib.request.Request(
            url,
            data=payload,
            headers={
                "Content-Type": "application/json",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            },
            method="POST"
        )
        urllib.request.urlopen(req, timeout=10)
        print(f"    📨 Discord 進度通知已發送：{title}")
    except Exception as e:
        print(f"    ⚠️ Discord 進度通知失敗：{e}")

# ✅ 更新：加入 final_soap 參數，若有值則附加於 Discord 訊息中
def send_discord_per_patient(chart_no, patient_name, status_msg, final_soap=None):
    """每完成一位病人時發送 Discord 通知（需在進階設定中啟用）"""
    global discord_webhook_url, notify_per_patient
    if not notify_per_patient:
        return
    url = discord_webhook_url.strip() if discord_webhook_url else ""
    if not url:
        return
    try:
        label = f"{chart_no}（{patient_name}）" if patient_name else chart_no
        content = f"🔔 **病人處理完成**\n病歷號：`{label}`\n狀態：{status_msg}"
        
        # 將最終的 SOAP 加入到 Discord 訊息中
        if final_soap:
            soap_text = f"\n\n**[最終 SOAP 預覽]**\n"
            soap_text += f"**【S】**\n{final_soap.get('S', '')}\n\n"
            soap_text += f"**【O】**\n{final_soap.get('O', '')}\n\n"
            soap_text += f"**【A】**\n{final_soap.get('A', '')}\n\n"
            soap_text += f"**【P】**\n{final_soap.get('P', '')}"
            
            # 確保不會超過 Discord 單則訊息 2000 字的限制
            if len(content) + len(soap_text) > 1900:
                content += soap_text[:1900 - len(content)] + "\n\n...(字數過長已截斷)"
            else:
                content += soap_text

        payload = json.dumps({"content": content}).encode("utf-8")
        req = urllib.request.Request(
            url,
            data=payload,
            headers={
                "Content-Type": "application/json",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            },
            method="POST"
        )
        urllib.request.urlopen(req, timeout=10)
        print(f"    📨 Discord 個人通知已發送：{label} → {status_msg}")
    except Exception as e:
        print(f"    ⚠️ Discord 個人通知失敗：{e}")

def generate_current_report():
    expected_map      = global_report_state["expected_map"]
    completed_map     = global_report_state["completed_map"]
    skipped_map       = global_report_state["skipped_map"]        
    new_consult_map   = global_report_state["new_consult_map"]
    missing_chart_map = global_report_state["missing_chart_map"]
    missing_today_map = global_report_state["missing_today_map"]
    ghost_record_map  = global_report_state["ghost_record_map"]
    exist_record_map  = global_report_state["exist_record_map"]
    forced_draft_map  = global_report_state["forced_draft_map"]
    chinese_herb_map  = global_report_state["chinese_herb_map"]
    discharged_map    = global_report_state["discharged_map"]

    expected_set      = set(expected_map.keys())
    completed_set     = set(completed_map.keys())
    skipped_set       = set(skipped_map.keys())                   
    new_consult_set   = set(new_consult_map.keys())
    missing_chart_set = set(missing_chart_map.keys())
    missing_today_set = set(missing_today_map.keys())

    # 計算真正的 missing 
    missing = expected_set - completed_set - new_consult_set - missing_chart_set - missing_today_set - skipped_set
    extra   = completed_set - expected_set

    name_lookup = {c: info["name"] for c, info in expected_map.items()}
    name_lookup.update(completed_map)
    name_lookup.update(skipped_map)                               
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
        
    if skipped_map:
        warn_lines.append(f"\n【略過 / 處理發生錯誤 {len(skipped_map)} 位】（清單未勾選或過程異常）：")
        warn_lines += fmt_list(skipped_set)
        
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
    if chinese_herb_map:
        warn_lines.append(f"\n【含中藥治療改暫存 {len(chinese_herb_map)} 位】（P欄含中藥/中藥水治療）：")
        warn_lines += fmt_list(set(chinese_herb_map.keys()))
    if discharged_map:
        warn_lines.append(f"\n【已出院病人 {len(discharged_map)} 位】（S已補充出院提示，改暫存）：")
        warn_lines += fmt_list(set(discharged_map.keys()))

    warn_lines.append(f"\n【已完成清單（共 {len(completed_map)} 位）】：")
    if completed_map:
        warn_lines += [f"  • {c}（{n}）" for c, n in completed_map.items()]
    else:
        warn_lines.append("  (目前尚無完成資料)")

    has_warnings = bool(
        missing or extra or new_consult_map or missing_chart_map
        or forced_draft_map or ghost_record_map or missing_today_map
        or chinese_herb_map or skipped_map or discharged_map
    )
    return "\n".join(warn_lines), has_warnings

def check_stop():
    if stop_event.is_set():
        raise Exception("【使用者手動停止】")

def safe_checkpoint(checkpoint_name=""):
    check_stop()
    if not pause_event.is_set():
        msg = f"    ⏸️ 系統已在安全段落 [{checkpoint_name}] 暫停，等待您按下繼續…" if checkpoint_name else "    ⏸️ 系統已暫停，等待繼續…"
        print(msg)
        pause_event.wait()
        print("    ▶️ 系統恢復執行…")
        check_stop()

def fetch_and_format_chinese_medicine(driver, wait, record_date, physician):
    """
    開啟藥囑紀錄，爬取指定日期的中藥內容，計算日劑量並回傳格式化字串。
    完成後自動切回「病程記錄」頁籤。
    """
    try:
        med_record_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@title='醫藥囑相關紀錄']")))
        driver.execute_script("arguments[0].click();", med_record_menu)
        time.sleep(1.5)

        med_order_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'groupDetailSelect') and contains(text(), '藥囑紀錄')]")))
        driver.execute_script("arguments[0].click();", med_order_btn)
        time.sleep(2)

        target_date_obj = datetime.datetime.strptime(record_date, "%Y-%m-%d").date()
        all_rows_xpath = "//div[contains(@class, 'p-dialog')]//tbody/tr[td]"
        all_rows = driver.find_elements(By.XPATH, all_rows_xpath)

        best_row_primary = None
        min_days_diff_primary = 9999

        best_row_fallback = None
        min_days_diff_fallback = 9999

        # 掃描所有紀錄，根據條件分級鎖定
        for row in all_rows:
            try:
                tds = row.find_elements(By.TAG_NAME, "td")
                if len(tds) < 3: continue
                row_date_str = tds[0].text.strip()
                row_type = tds[1].text.strip()
                row_dr = tds[2].text.strip()

                row_date_obj = datetime.datetime.strptime(row_date_str.split(" ")[0], "%Y-%m-%d").date()
                if row_date_obj <= target_date_obj:
                    diff = (target_date_obj - row_date_obj).days
                    
                    # 條件 1 (首選)：醫師姓名相符且日期最近
                    if physician and physician in row_dr:
                        if diff < min_days_diff_primary:
                            min_days_diff_primary = diff
                            best_row_primary = row
                            
                    # 條件 2 (備案)：住院且 30 天內最近
                    if "住院" in row_type and diff <= 30:
                        if diff < min_days_diff_fallback:
                            min_days_diff_fallback = diff
                            best_row_fallback = row
            except Exception:
                continue

        best_row = best_row_primary if best_row_primary else best_row_fallback

        if best_row:
            matched_type = "醫師相符" if best_row_primary else "近30天內住院"
            print(f"        👉 已鎖定最佳紀錄 ({matched_type})，準備點擊...")
            driver.execute_script("arguments[0].click();", best_row)
            time.sleep(2)
            
            # ✅ 新增：點擊「住院用藥」頁籤
            try:
                inpatient_tab_xpath = "//a[contains(@class, 'p-tabview-nav-link')]//span[contains(text(), '住院用藥')]"
                inpatient_tab = driver.find_element(By.XPATH, inpatient_tab_xpath)
                driver.execute_script("arguments[0].click();", inpatient_tab)
                time.sleep(1.5)
                print("        👉 已成功切換至「住院用藥」頁籤")
            except Exception:
                print("        ⚠️ 找不到「住院用藥」頁籤，將直接抓取當前畫面")
                
        else:
            print("        ⚠️ 找不到符合條件 (醫師相符 或 近30天內住院) 的紀錄")
            raise Exception("No suitable record found")

        med_rows_xpath = "//div[contains(@class, 'p-dialog')]//div[contains(@class, 'p-datatable-scrollable-body')]//tbody/tr"
        med_rows = driver.find_elements(By.XPATH, med_rows_xpath)
        
        powders = []
        decoctions = []
        current_med_name = ""
        current_freq = ""

        for row in med_rows:
            classes = row.get_attribute("class") or ""
            advise_tds = row.find_elements(By.CSS_SELECTOR, "td.advise")

            # 處理水藥 (advise 囑咐欄位)
            if "advise" in classes or advise_tds:
                if current_med_name == "中藥":
                    advise_td = advise_tds[0] if advise_tds else row.find_element(By.TAG_NAME, "td")
                    advise_text = advise_td.text.strip()
                    if advise_text.startswith("囑咐:"):
                        advise_text = advise_text[3:].strip()

                    parts = advise_text.split(',')
                    formatted_herbs = []
                    for part in parts:
                        # 嚴格鎖定單位為「錢」
                        match = re.search(r'([^\d]+?)(?:\(自\))?\s*(\d+\.\d+)\s*錢', part)
                        if match:
                            raw_name = match.group(1)
                            dose = match.group(2)
                            
                            clean_name = raw_name.strip('() ')
                            clean_name = clean_name.replace('(生)', '生').replace('(法)', '法')
                            # 完美過濾雜訊，並保留 (炒)、(苦)
                            clean_name = re.sub(r'＊|☆|△|-包煎|-後下|\(自\)', '', clean_name)
                            clean_name = re.sub(r'^\(+|\)+$', '', clean_name).strip()

                            # 全形空白對齊四個字元
                            padded_name = clean_name
                            if len(clean_name) < 4:
                                padded_name = clean_name + chr(12288) * (4 - len(clean_name))

                            formatted_herbs.append(f"{padded_name} {dose}錢")

                    if formatted_herbs:
                        decoctions.append(f"#中藥水治療 {current_freq}")
                        for h in formatted_herbs:
                            decoctions.append(h)
                continue

            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) < 10: continue

            med_name = ""
            dose_str = ""
            freq_str = ""
            if len(cols) == 14:
                med_name = cols[3].text.strip()
                dose_str = cols[4].text.strip()
                freq_str = cols[7].text.strip()
            elif len(cols) == 13:
                med_name = cols[2].text.strip()
                dose_str = cols[3].text.strip()
                freq_str = cols[6].text.strip()
            else:
                continue

            if not med_name: continue

            current_med_name = med_name
            current_freq = freq_str

            # 1. 過濾純英文西藥 (藥名必須含有中文)
            has_chinese = any('\u4e00' <= char <= '\u9fff' for char in med_name)
            if not has_chinese: continue
            
            # 2. 過濾帶有中文的西藥 (例如 "Nalbuphine (芯奔)")
            has_english = re.search(r'[a-zA-Z]', med_name)
            if has_english and med_name != "中藥": 
                continue

            # 3. 過濾行政費用
            if "調劑費" in med_name or "藥費加成" in med_name: continue

            # 若為科學中藥 (非水藥的標籤 "中藥")
            if med_name != "中藥":
                try:
                    dose_val = float(re.search(r'\d+(\.\d+)?', dose_str).group())
                except (ValueError, AttributeError):
                    continue

                freq_upper = freq_str.upper()
                if "QD" in freq_upper: freq_factor = 1
                elif "BID" in freq_upper: freq_factor = 2
                elif "TID" in freq_upper: freq_factor = 3
                elif "QID" in freq_upper: freq_factor = 4
                else: freq_factor = 1
                
                match = re.search(r'(QD|BID|TID|QID)', freq_upper)
                display_freq = match.group(1) if match else freq_str

                daily_dose = dose_val * freq_factor
                rounded_dose = round(daily_dose * 2) / 2
                powders.append(f"{med_name} {rounded_dose:.2f} GM {display_freq}")

        try:
            close_btn = driver.find_element(By.XPATH, "//div[contains(@class, 'p-dialog')]//span[contains(@class, 'p-dialog-header-close-icon')]")
            driver.execute_script("arguments[0].click();", close_btn)
            time.sleep(1)
        except Exception as e:
            print(f"        ⚠️ 關閉藥囑視窗發生狀況：{e}")

        try:
            progress_note_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@title='病程記錄']")))
            driver.execute_script("arguments[0].click();", progress_note_menu)
            time.sleep(1.5)
            print("        👉 已成功切回「病程記錄」標籤...")
        except Exception as e:
            print(f"        ⚠️ 切回病程記錄發生狀況：{e}")

        result = ""
        if decoctions:
            result += "\n".join(decoctions)
        if powders:
            if result: result += "\n\n"
            result += "#科學中藥\n" + "\n".join(powders)
        
        return result

    except Exception as e:
        print(f"        ⚠️ 抓取中藥紀錄失敗 (直接略過)：{e}")
        try:
            close_btn = driver.find_element(By.XPATH, "//div[contains(@class, 'p-dialog')]//span[contains(@class, 'p-dialog-header-close-icon')]")
            driver.execute_script("arguments[0].click();", close_btn)
            time.sleep(0.5)
            progress_note_menu = driver.find_element(By.XPATH, "//img[@title='病程記錄']")
            driver.execute_script("arguments[0].click();", progress_note_menu)
            time.sleep(1)
        except: pass
        return ""

# ✅ 更新：加入已處理姓名床位的同步更新邏輯，並接收 final_soap 往下傳遞
def update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status, name, chart_no, status_msg, final_soap=None):
    def _update():
        # 空白病歷號：只從待處理欄位中移除對應行，不寫入已處理三欄
        if not chart_no or chart_no.strip() == "":
            charts_lines = txt_charts.get("1.0", "end-1c").split('\n')
            names_lines  = txt_names.get("1.0", "end-1c").split('\n')
            # 找到第一個空白病歷號行並移除（連同對應的姓名行）
            target_idx = -1
            for i, c in enumerate(charts_lines):
                if c.strip() == "":
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
            return  # 空白行無視，不寫入已處理欄位

        # 正常病歷號：三欄同步寫入（確保順序一致）
        display_name = name if name else ""
        txt_proc_names.insert(tk.END, f"{display_name}\n")
        txt_proc_names.see(tk.END)

        txt_proc_charts.insert(tk.END, f"{chart_no}\n")
        txt_proc_charts.see(tk.END)

        txt_proc_status.insert(tk.END, f"{status_msg}\n")
        txt_proc_status.see(tk.END)

        # 每完成一位病人通知（若有啟用），傳遞 final_soap
        threading.Thread(
            target=send_discord_per_patient,
            args=(chart_no, name, status_msg, final_soap),
            daemon=True
        ).start()
        
        charts_lines = txt_charts.get("1.0", "end-1c").split('\n')
        names_lines  = txt_names.get("1.0", "end-1c").split('\n')
        
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

    try:
        driver.execute_script("document.querySelector('tbody.p-datatable-tbody').closest('.p-datatable-scrollable-body, .p-datatable-wrapper, .p-scroller, [class*=\"scroll\"]').scrollTop = 0;")
    except Exception:
        try:
            driver.execute_script("window.scrollTo(0, 0);")
        except Exception:
            pass
    time.sleep(0.3)

    js_count_checked = """
    var rows = document.querySelectorAll("tbody.p-datatable-tbody tr");
    var count = 0;
    for (var i = 0; i < rows.length; i++) {
        var cb = rows[i].querySelector(".p-checkbox-box");
        if (cb && cb.className.includes("p-highlight")) count++;
    }
    return count;
    """
    stable_count = -1
    stable_retries = 0
    while stable_retries < 4:
        check_stop()
        try:
            current_checked = driver.execute_script(js_count_checked)
        except Exception:
            current_checked = 0
        if current_checked == stable_count:
            stable_retries += 1
        else:
            stable_count = current_checked
            stable_retries = 0
        if stable_retries >= 2:
            break
        time.sleep(0.8)

    print(f"    📋 表格共載入 {prev_count} 列，已勾選 {stable_count} 筆（checkbox 狀態已穩定）")
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
    print("    → 執行「回到病人清單函數」…")
    try:
        ensure_window_focus(driver)
        
        try:
            close_btns = driver.find_elements(By.XPATH, "//button[contains(@class, 'p-dialog-header-close')]|//span[contains(@class, 'p-dialog-header-close-icon') and contains(@class, 'pi-times')]")
            for btn in close_btns:
                if btn.is_displayed():
                    driver.execute_script("arguments[0].click();", btn)
                    print("        🧹 已清除殘留的對話框視窗")
                    time.sleep(0.5)
        except Exception:
            pass

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
        try: driver.back()
        except: pass

def final_countdown_and_close(driver, report_msg):
    threading.Thread(target=send_discord_notification, args=(report_msg,), daemon=True).start()

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
                    top.after(1000, countdown, left - 1)
                except: pass
            else: force_close()
                
        countdown(300)
    root.after(0, show_final_ui)

def step_6_submit_or_draft(driver, wait, action_mode, opd_dr_name, physician_code):
    safe_checkpoint("準備執行最終送件/暫存")
    print(f"\n    ▶ 進入第六步驟：準備執行【{action_mode}】…")
    try:
        ensure_window_focus(driver)
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

# ✅ 最終修正版：無重複，包含完美健保頭擷取與中藥爬蟲呼叫
def step_5_add_new_record(driver, wait, chart_no, patient_name,
                          old_subjective, old_objective, ditto_o, ditto_a,
                          opd_plan, plan_for_herb_check,
                          record_date, action_mode, opd_dr_name, physician_code,
                          draft_on_herb=False, append_discharge_note=False, keep_jianbao_flag=True):
    safe_checkpoint("準備自動填寫病歷資料")
    print(f"\n    ▶ 進入第五步驟：開始新增病歷 — {chart_no}（{patient_name}）")

    # 1. 定義中藥關鍵字
    herb_keywords = ["中藥", "科學中藥", "中藥粉", "中藥水", "水藥", "水煎藥"]
    
    # 2. 取得要檢查的文字範圍（只看門診P，ditto P 為舊病歷不代表今天）
    full_text_to_scan = (opd_plan or "")
    
    # 判斷是否需要強制改「暫存」
    herb_detected_for_draft = draft_on_herb and any(kw in full_text_to_scan for kw in herb_keywords)
    if herb_detected_for_draft:
        print(f"    🌿 偵測到 P 欄或計畫中含中藥相關內容，將改為【暫存】模式！")
        action_mode = "暫存"

    try:
        ensure_window_focus(driver)
        # 點擊診斷列表進入編輯介面
        first_diagnosis_td = wait.until(EC.presence_of_element_located((By.XPATH, "(//tr[contains(@class, 'p-selectable-row')]//td[contains(@class, 'dt-status')])[1]")))
        driver.execute_script("arguments[0].click();", first_diagnosis_td)
        time.sleep(1)
        
        formatted_subjective = old_subjective
        if old_subjective:
            y, m, d = record_date.split("-")
            roc_year = int(y) - 1911
            patterns = [f"({roc_year}/{m}/{d})", f"（{roc_year}/{m}/{d}）", f"({record_date})", f"（{record_date}）"]
            for pat in patterns:
                if pat in old_subjective:
                    after = old_subjective.split(pat)[-1].strip()
                    next_entry = re.search(r'\(\d{2,3}/\d{2}/\d{2}\)|\(\d{4}-\d{2}-\d{2}\)', after)
                    if next_entry: after = after[:next_entry.start()].strip()
                    formatted_subjective = after
                    break

        if append_discharge_note:
            discharge_note = "預計於今日出院。"
            if formatted_subjective:
                if not formatted_subjective.rstrip().endswith("今日出院。"):
                    formatted_subjective = formatted_subjective.rstrip() + "\n" + discharge_note
            else: formatted_subjective = discharge_note
            print(f"    🏥 已在 S 欄末尾補充：「{discharge_note}」")

        formatted_plan = ""
        raw_plan = opd_plan or ""
        jianbao_header = ""
        
        
        if raw_plan:
            # 擷取健保計畫頭部文字
            header_match = re.search(r'^(患者.*?符合.*?計畫(?:[(（][^)）]*[)）])?)', raw_plan)
            if header_match:
                jianbao_header = header_match.group(1).strip()
            
            # 檢查整段文字是否包含中藥關鍵字，決定是否啟動爬蟲
            # 即使沒有 #標籤，只要文字裡有提到「中藥、水煎藥」等就會觸發
            trigger_crawler = any(kw in (raw_plan + jianbao_header) for kw in herb_keywords)

            fetched_meds = ""
            if trigger_crawler:
                print("    🌿 偵測到計畫內容提及中藥，啟動藥囑紀錄爬蟲...")
                fetched_meds = fetch_and_format_chinese_medicine(driver, wait, record_date, opd_dr_name)
                if fetched_meds:
                    print("    ✅ 已成功抓取並計算藥量。")

            # 組合最終的 P 欄文字
            if keep_jianbao_var.get() and jianbao_header:
                formatted_plan += jianbao_header + "\n\n"

            # 清理原始計畫中的標籤格式
            if '#' in raw_plan:
                clean_plan = raw_plan[raw_plan.find('#'):]
                clean_plan = re.sub(r'\s+(?=#)', '\n', clean_plan)
                # 移除計畫中的時間戳記
                clean_plan = re.sub(r'\s*\(\d{2,3}/\d{2}/\d{2}\)\s*(?=Time\s*out|Sign\s*out)', '\n', clean_plan, flags=re.IGNORECASE)
                
                result_lines = [line.strip() for line in clean_plan.splitlines() if line.strip()]
                # 過濾掉可能重複的手打藥物標籤
                non_herb_lines = [l for l in result_lines if not any(k in l for k in ["#中藥", "#科學中藥", "#水藥"])]
                formatted_plan += "\n\n".join(non_herb_lines)
            
            # 將爬下來的正式藥單貼在最下面
            if fetched_meds:
                if formatted_plan and not formatted_plan.endswith("\n\n"):
                    formatted_plan += "\n\n"
                formatted_plan += fetched_meds
                    
            print(f"    📋 P 欄排版完成（已包含自動補抓的中藥內容）")

        s_box = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='B1']/following-sibling::textarea[1]")))
        o_box = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='B2']/following-sibling::textarea[1]")))
        a_box = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='B3']/following-sibling::textarea[1]")))
        p_box = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='B4']/following-sibling::textarea[1]")))
        
        if formatted_subjective: s_box.clear(); s_box.send_keys(formatted_subjective)
        if ditto_a: a_box.clear(); a_box.send_keys(ditto_a)
        if formatted_plan: p_box.clear(); p_box.send_keys(formatted_plan)
            
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
        
        # 格式化：確保望診/聞診/舌診/切診各自獨立成一行，切診後與其他內容空行分隔
        if target_o:
            for kw in ["望診：", "聞診：", "舌診：", "切診："]:
                target_o = re.sub(r'(?<!\n)(' + re.escape(kw) + r')', r'\n\1', target_o)
            # 切診行結尾後若緊接著非換行內容，插入空行
            target_o = re.sub(r'(切診：[^\n]+)(?=\n[^\n])', r'\1\n', target_o)
            target_o = target_o.strip()
        
        if target_o:
            lines = target_o.splitlines()
            new_o_lines = []
            skipping_vitals = False
            for line in lines:
                line_lower = line.lower()
                if any(k in line for k in ["望診", "聞診", "舌診", "切診"]): skipping_vitals = False
                if "vital signs" in line_lower: skipping_vitals = True; continue
                if skipping_vitals:
                    if any(k in line_lower for k in ["spo2", "rr", "fio2", "peep", "pcv", "temperature", "bp", "pulse", "blood pressure", "respiratory", "peripheral"]): continue
                    if re.search(r'\d+\.?\d*\s*/\s*\d+\.?\d*\s*=', line): continue
                    if re.search(r'\d+\.?\d*\s*(℃|mmhg|%|per min)', line, re.IGNORECASE): continue
                    skipping_vitals = False
                line = re.sub(r'^\s*\(\d{4}-\d{2}-\d{2}\)\s*', '', line)
                line = re.sub(r'^\s*\(\d{2,3}/\d{2}/\d{2}\)\s*', '', line)
                if not skipping_vitals: new_o_lines.append(line)
            
            target_o = "\n".join(new_o_lines).strip()
            if any(keyword in target_o for keyword in ["報告時間", "檢驗單", "影像部檢查"]):
                needs_new_reports = True
                match = re.search(r'(望診：.*?切診：[^\n]*)', target_o, flags=re.DOTALL)
                if match: target_o = match.group(1)
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
        
        try:
            final_s = s_box.get_attribute("value") or formatted_subjective
            final_o = o_box.get_attribute("value") or target_o
            final_a = a_box.get_attribute("value") or ditto_a
            final_p = p_box.get_attribute("value") or formatted_plan
        except Exception:
            final_s, final_o, final_a, final_p = formatted_subjective, target_o, ditto_a, formatted_plan
            
        final_soap_dict = {"S": final_s, "O": final_o, "A": final_a, "P": final_p}
        
        print(f"\n    {'='*40}")
        print(f"    📝 [最終 {action_mode} 前 SOAP 預覽]")
        print(f"    {'='*40}")
        print(f"    【S】\n{final_s}\n")
        print(f"    【O】\n{final_o}\n")
        print(f"    【A】\n{final_a}\n")
        print(f"    【P】\n{final_p}")
        print(f"    {'='*40}\n")

        step_6_submit_or_draft(driver, wait, action_mode, opd_dr_name, physician_code)

        return herb_detected_for_draft, final_soap_dict
        
    except Exception as e:
        if "使用者手動停止" in str(e): raise e
        raise e 

# ✅ 更新：接收 final_soap_dict 並往回傳遞
def step_4_write_record(driver, wait, chart_no, patient_name, physician, physician_code,
                        record_date, action_mode, draft_on_herb=False, name_ref=None):
    label = f"{chart_no}（{patient_name}）"
    print(f"    ⏳ 開始執行病歷操作 — {label}")

    mismatch_detected = False
    is_ghost_record   = False
    is_discharged     = False
    discharge_date    = ""
    old_subjective    = ""
    old_objective     = ""
    old_plan          = ""

    try:
        time.sleep(2)
        ensure_window_focus(driver)

        # 進入病歷後第一時間擷取患者的真實姓名與床位
        try:
            banner_el = driver.find_element(By.ID, "page-banner")
            banner_text = banner_el.text
            parts = banner_text.strip().split()
            name_parts = []
            for part in parts:
                if part.isdigit() and len(part) >= 7: 
                    break
                name_parts.append(part)
            
            extracted_name = " ".join(name_parts)
            if extracted_name:
                patient_name = extracted_name  
                if name_ref is not None:
                    name_ref[0] = patient_name  
                print(f"    👤 自動擷取患者床位姓名：{patient_name}")
        except Exception as e:
            print(f"    ⚠️ 無法擷取姓名床位 (將使用原本的名稱)：{e}")

        try:
            discharged_btns = driver.find_elements(
                By.XPATH,
                "//button[contains(., '已出院無法存檔')]"
            )
            if discharged_btns:
                is_discharged = True
                print(f"    🏥 偵測到「已出院無法存檔」按鈕，標記為已出院病人")
                try:
                    banner_el = driver.find_element(By.ID, "page-banner")
                    banner_text = banner_el.text
                    m_dis = re.search(r'住院期間[：:]\s*\d{4}-\d{2}-\d{2}[~～~]\s*(\d{4}-\d{2}-\d{2})', banner_text)
                    if m_dis:
                        discharge_date = m_dis.group(1).strip()
                        print(f"    📅 出院日期：{discharge_date}")
                except Exception as be:
                    print(f"    ⚠️ 無法讀取出院日期：{be}")
        except Exception as de:
            print(f"    ⚠️ 出院偵測發生狀況（略過）：{de}")

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
                            else: top.after(1000, auto_close, left - 1)
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
                    subjective_xpath = "//fieldset[.//legend[contains(., '主觀') or contains(., 'Subjective') or contains(., '會診目的') or contains(., '主訴') or contains(., '診斷主訴')]]/div/div[contains(@class, 'p-fieldset-content')]"
                    subjective_el = driver.find_element(By.XPATH, subjective_xpath)
                    old_subjective = driver.execute_script("return arguments[0].innerText;", subjective_el) or ""
                    old_subjective = '\n'.join([line.strip() for line in old_subjective.splitlines() if line.strip()])
                except Exception: pass

                try:
                    objective_xpath = "//fieldset[.//legend[contains(., '理學檢查') or contains(., 'Objective')]]/div/div[contains(@class, 'p-fieldset-content')]"
                    objective_el = driver.find_element(By.XPATH, objective_xpath)
                    old_objective = driver.execute_script("return arguments[0].innerText;", objective_el) or ""
                    old_objective = '\n'.join([line.strip() for line in old_objective.splitlines() if line.strip()])
                except Exception: pass

                try:
                    plan_xpath = "//fieldset[.//legend[contains(., '治療計畫') or contains(., 'Plan')]]/div/div[contains(@class, 'p-fieldset-content')]"
                    plan_el = driver.find_element(By.XPATH, plan_xpath)
                    old_plan = driver.execute_script("return arguments[0].innerText;", plan_el) or ""
                    old_plan = old_plan.strip()
                except Exception: pass
            else:
                raise Exception("疑似今日未上傳門診紀錄")
        
        try:
            close_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'pi-times') and contains(@class, 'p-dialog-header-close-icon')]")))
            driver.execute_script("arguments[0].click();", close_btn)
            time.sleep(2)
        except: pass
        
        safe_checkpoint("準備切換至 DITTO 操作")
        ensure_window_focus(driver)
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

        plan_for_herb_check = old_plan  # 只用門診P，ditto P 為舊病歷不代表今天
        
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
            plan_for_herb_check = ""  # ghost record 無門診P，不觸發爬蟲

        append_discharge_note = False
        if is_discharged and discharge_date and discharge_date == record_date:
            append_discharge_note = True
            print(f"    🏥 出院日期與目標日期相符，將在 S 末尾補充出院提示")
            action_mode = "暫存"   

        # ✅ 接收 step_5 傳回來的 SOAP
        herb_triggered, final_soap = step_5_add_new_record(
            driver, wait, chart_no, patient_name,
            old_subjective, old_objective, ditto_o, ditto_a,
            opd_plan=old_plan,             
            plan_for_herb_check=plan_for_herb_check,   
            record_date=record_date, action_mode=action_mode,
            opd_dr_name=physician, physician_code=physician_code,
            draft_on_herb=draft_on_herb,
            append_discharge_note=append_discharge_note
        )

        # ✅ 回傳狀態、患者姓名與最後的 SOAP
        if is_ghost_record:
            return "ghost_record", patient_name, final_soap
        if herb_triggered:
            return "chinese_herb", patient_name, final_soap
        if is_discharged:
            return "discharged", patient_name, final_soap
        return ("forced_draft" if mismatch_detected else "success"), patient_name, final_soap
        
    except Exception as e:
        if "使用者手動停止" not in str(e): print(f"    ❌ 執行病歷操作發生狀況：{e}")
        raise e

# ✅ 更新：接收從 step_4 傳上來的 final_soap 並丟給 update_ui 函數
def step_3_process_patients(driver, wait, groups, action_mode,
                            auto_add_flag=False, auto_uncheck_flag=False,
                            draft_on_herb=False, priority_mode="checked_first"):
    print("\n— 進入第三步驟：處理已勾選病人 —")

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
    chinese_herb_map  = global_report_state["chinese_herb_map"]
    discharged_map    = global_report_state["discharged_map"]

    for g in groups:
        for p in g["patients"]:
            expected_map[p["chart_no"]] = {
                "name":           p["name"],
                "physician":      g["physician"],
                "physician_code": g["code"],
            }

    print(f"📋 期望處理病人（共 {len(expected_map)} 位，跨 {len(groups)} 個群組）：")

    def get_checked_chart_nos():
        try:
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "tbody.p-datatable-tbody tr")
            ))
        except Exception:
            pass

        js_batch = """
        var results = [];
        var rows = document.querySelectorAll("tbody.p-datatable-tbody tr");
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            var cb = row.querySelector(".p-checkbox-box");
            if (!cb || !cb.className.includes("p-highlight")) continue;
            var tds = row.querySelectorAll("td");
            if (tds.length < 4) continue;
            var chart_no = tds[3].innerText.trim().replace(/^0+/, "");
            var name = tds[2].innerText.trim();
            if (chart_no) results.push([name, chart_no]);
        }
        return results;
        """
        prev_result = None
        for attempt in range(6):   
            check_stop()
            try:
                raw = driver.execute_script(js_batch)
                current = [(r[0], r[1]) for r in raw] if raw else []
            except Exception:
                current = []

            if prev_result is not None and len(current) == len(prev_result):
                prev_set = {c for _, c in prev_result}
                curr_set = {c for _, c in current}
                if prev_set == curr_set:
                    print(f"    ☑️ checkbox 讀取穩定（第 {attempt+1} 次確認，共 {len(current)} 筆勾選）")
                    return current
            prev_result = current
            time.sleep(0.8)

        print(f"    ⚠️ checkbox 狀態未完全穩定，使用最後讀取結果（共 {len(prev_result)} 筆）")
        return prev_result if prev_result else []

    checked_patients = get_checked_chart_nos()
    print(f"\n☑️  表格中已勾選病人（共 {len(checked_patients)} 位）")

    if not checked_patients:
        messagebox.showwarning("注意", "表格中找不到任何已勾選的病人！")
        return

    checked_set = {chart_no for _, chart_no in checked_patients}

    print(f"\n    ✅ 清單讀取完成，共 {len(checked_set)} 位已勾選，開始逐群組處理...")

    def _search_by_chart_no_and_get_row(chart_no):
        try:
            ensure_window_focus(driver)
            try:
                chart_radio_label = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//label[contains(normalize-space(), '病歷號')]")
                ))
                driver.execute_script("arguments[0].click();", chart_radio_label)
            except Exception:
                try:
                    radio_inputs = driver.find_elements(By.XPATH, "//input[@type='radio' and @name='admPtType']")
                    if len(radio_inputs) >= 2:
                        driver.execute_script("arguments[0].click();", radio_inputs[1])
                except Exception:
                    pass
            time.sleep(1)

            chart_input = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//input[@maxlength='10' and @type='text']")
            ))
            try: chart_input.click()
            except Exception: driver.execute_script("arguments[0].click();", chart_input)
            chart_input.send_keys(Keys.CONTROL + 'a')
            time.sleep(0.2)
            chart_input.send_keys(Keys.BACKSPACE)
            time.sleep(0.2)
            chart_input.send_keys(chart_no)
            time.sleep(0.5)

            search_btn = driver.find_element(By.XPATH, "//button[.//span[contains(., '查詢')]]")
            driver.execute_script("arguments[0].click();", search_btn)
            time.sleep(2)

            rows = driver.find_elements(By.CSS_SELECTOR, "tbody.p-datatable-tbody tr")
            inpatient_rows = []
            any_rows = []
            for row in rows:
                try:
                    tds = row.find_elements(By.TAG_NAME, "td")
                    if len(tds) < 4: continue
                    row_chart = tds[3].text.strip().lstrip("0")
                    if row_chart != chart_no.lstrip("0"): continue
                    try:
                        cb = tds[0].find_element(By.CSS_SELECTOR, "div.p-checkbox-box")
                        is_checked = "p-highlight" in cb.get_attribute("class")
                    except Exception:
                        is_checked = False
                    any_rows.append((row, is_checked))
                    if len(tds) >= 11 and tds[10].text.strip() == "住院":
                        try:
                            dt = datetime.datetime.strptime(tds[9].text.strip(), "%Y-%m-%d %H:%M")
                        except Exception:
                            dt = datetime.datetime.min
                        inpatient_rows.append((dt, row, is_checked))
                except Exception:
                    continue

            if inpatient_rows:
                inpatient_rows.sort(key=lambda x: x[0], reverse=True)
                _, target_row, is_checked = inpatient_rows[0]
                print(f"        ✅ 找到住院記錄，checkbox={'已勾選' if is_checked else '未勾選'}")
                return target_row, is_checked
            elif any_rows:
                target_row, is_checked = any_rows[0]
                print(f"        ✅ 找到記錄（非住院），checkbox={'已勾選' if is_checked else '未勾選'}")
                return target_row, is_checked
            else:
                print(f"        ❌ 病歷號搜尋查無資料：{chart_no}")
                return None, False

        except Exception as e:
            if "使用者手動停止" in str(e): raise e
            print(f"        ⚠️ 病歷號搜尋發生錯誤：{e}")
            return None, False

    def _return_to_my_list():
        try:
            my_list_label = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//label[contains(normalize-space(), '我的清單')]")
            ))
            driver.execute_script("arguments[0].click();", my_list_label)
            time.sleep(1.5)
            print("      → 已切回我的清單")
        except Exception as e:
            print(f"      ⚠️ 切回我的清單失敗：{e}")

    def _process_one_patient_via_search(chart_no, patient_name, physician, physician_code,
                                        record_date, action_mode, txt_names, txt_charts, txt_proc_names,
                                        txt_proc_charts, txt_proc_status, label_prefix="出院"):
        safe_checkpoint(f"準備處理{label_prefix}病人：{patient_name}")
        print(f"\n  ▶ [{label_prefix}] {chart_no}（{patient_name}）")

        name_ref = [patient_name]  
        try:
            found_row, is_checked = _search_by_chart_no_and_get_row(chart_no)

            if found_row is None:
                skipped_map[chart_no] = patient_name
                update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status,
                                            patient_name, chart_no, "⚠️ 查無此病歷號")
                _return_to_my_list()
                return

            if not is_checked:
                try:
                    tds = found_row.find_elements(By.TAG_NAME, "td")
                    cb = tds[0].find_element(By.CSS_SELECTOR, "div.p-checkbox-box")
                    driver.execute_script("arguments[0].click();", cb)
                    time.sleep(0.5)
                    print(f"        ✅ 已勾選病人")
                except Exception as ce:
                    print(f"        ⚠️ 勾選失敗（繼續嘗試進入）：{ce}")

            ensure_window_focus(driver)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", found_row)
            time.sleep(0.5)
            ActionChains(driver).double_click(found_row).perform()
            time.sleep(2)
            try:
                WebDriverWait(driver, 5).until(
                    EC.invisibility_of_element_located((By.CSS_SELECTOR, "tbody.p-datatable-tbody tr"))
                )
            except Exception:
                pass
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "a.al-sidebar-list-link"))
            )

            # ✅ 解包新增的 final_soap
            status, updated_name, final_soap = step_4_write_record(
                driver, wait, chart_no, patient_name,
                physician, physician_code, record_date, action_mode,
                draft_on_herb=draft_on_herb,
                name_ref=name_ref
            )
            patient_name = updated_name  

            if status == "ghost_record":
                ghost_record_map[chart_no] = patient_name
                completed_map[chart_no] = patient_name
                update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status,
                                            patient_name, chart_no, f"✅ {label_prefix}幽靈暫存", final_soap)
            elif status == "forced_draft":
                forced_draft_map[chart_no] = patient_name
                completed_map[chart_no] = patient_name
                update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status,
                                            patient_name, chart_no, f"✅ {label_prefix}醫師不符暫存", final_soap)
            elif status == "chinese_herb":
                chinese_herb_map[chart_no] = patient_name
                completed_map[chart_no] = patient_name
                update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status,
                                            patient_name, chart_no, f"🌿 {label_prefix}含中藥暫存", final_soap)
            elif status == "discharged":
                discharged_map[chart_no] = patient_name
                completed_map[chart_no] = patient_name
                update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status,
                                            patient_name, chart_no, f"🏥 {label_prefix}已出院(暫存)", final_soap)
            else:
                completed_map[chart_no] = patient_name
                update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status,
                                            patient_name, chart_no, f"✅ {label_prefix}完成", final_soap)

            return_to_patient_list(driver, wait)
            _uncheck_patient_row(chart_no)   
            
        except Exception as e:
            patient_name = name_ref[0]  
            error_msg = str(e)
            if "使用者手動停止" in error_msg: raise e
            elif "疑似新會病歷" in error_msg:
                new_consult_map[chart_no] = patient_name
                update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status,
                                            patient_name, chart_no, "🌟 疑似新會病歷")
            elif "疑似今日未上傳門診紀錄" in error_msg:
                missing_today_map[chart_no] = patient_name
                update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status,
                                            patient_name, chart_no, "⚠️ 未上傳門診")
            elif "今日病歷已存在" in error_msg:
                exist_record_map[chart_no] = patient_name
                completed_map[chart_no] = patient_name
                update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status,
                                            patient_name, chart_no, "✅ 今日病歷已存")
            else:
                skipped_map[chart_no] = patient_name
                update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status,
                                            patient_name, chart_no, f"❌ {label_prefix}處理錯誤")
            try:
                return_to_patient_list(driver, wait)
            except Exception:
                pass

    for g_idx, g in enumerate(groups):
        physician       = g["physician"]
        physician_code  = g["code"]
        record_date     = g["date"].get().strip()
        txt_names       = g["txt_names"]
        txt_charts      = g["txt_charts"]
        txt_proc_names  = g["txt_proc_names"]  
        txt_proc_charts = g["txt_proc_charts"]
        txt_proc_status = g["txt_proc_status"]
        
        print(f"\n{'='*52}")
        print(f"▶▶ 群組 {g_idx+1}｜主治醫師：{physician}（{physician_code}）｜目標日期：{record_date}")
        print(f"{'='*52}")

        print(f"  → 預載入表格 row 快取...")
        scroll_to_load_all_rows(driver)
        row_cache = {}
        try:
            _all_rows = driver.find_elements(By.CSS_SELECTOR, "tbody.p-datatable-tbody tr")
            for _r in _all_rows:
                try:
                    _tds = _r.find_elements(By.TAG_NAME, "td")
                    if len(_tds) < 4: continue
                    _rno = _tds[3].text.strip().lstrip("0")
                    if _rno: row_cache[_rno] = _r
                except Exception:
                    continue
        except Exception:
            pass
        print(f"  → 快取建立完成，共 {len(row_cache)} 筆 row")

        def _rebuild_row_cache():
            row_cache.clear()
            try:
                scroll_to_load_all_rows(driver)
                _rows = driver.find_elements(By.CSS_SELECTOR, "tbody.p-datatable-tbody tr")
                for _r in _rows:
                    try:
                        _tds = _r.find_elements(By.TAG_NAME, "td")
                        if len(_tds) < 4: continue
                        _rno = _tds[3].text.strip().lstrip("0")
                        if _rno: row_cache[_rno] = _r
                    except Exception:
                        continue
                print(f"    🔄 快取已更新，剩餘 {len(row_cache)} 筆")
            except Exception as _ce:
                print(f"    ⚠️ 快取更新失敗（下一位將重新掃表）：{_ce}")

        checked_patients_group   = []
        unchecked_patients_group = []

        for p in g["patients"]:
            chart_no     = p["chart_no"]
            patient_name = p["name"]
            if not chart_no or chart_no.strip() == "":
                missing_chart_map["[空白病歷號]"] = patient_name
                update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status, patient_name, chart_no, "❌ 格式空白")
            elif chart_no in checked_set:
                checked_patients_group.append(p)
            else:
                unchecked_patients_group.append(p)

        if priority_mode == "unchecked_first":
            print(f"\n  ⚙️ 處理順序：先 1.【未在清單(搜尋)】({len(unchecked_patients_group)}位) → 再 2.【已在清單】({len(checked_patients_group)}位)")
        else:
            print(f"\n  ⚙️ 處理順序：先 2.【已在清單】({len(checked_patients_group)}位) → 再 1.【未在清單(搜尋)】({len(unchecked_patients_group)}位)（預設）")

        # ==========================================
        # 已在勾選清單的病人群
        # ==========================================
        def _run_checked_round():
            if not checked_patients_group:
                print(f"  → 無【已在清單】病人，跳過此輪。")
                return
            print(f"\n{'─'*52}")
            print(f"  ▶▶ 【已在清單】共 {len(checked_patients_group)} 位（直接在清單中雙擊）")
            print(f"{'─'*52}")
            for p in checked_patients_group:
                safe_checkpoint(f"準備處理病人：{p.get('name', '未知名稱')}")
                chart_no     = p["chart_no"]
                patient_name = p["name"]
                label        = f"{chart_no}（{patient_name}）" if patient_name else chart_no

                print(f"\n  ▶ [已勾選] 處理：{label}")
                name_ref = [patient_name]  
                try:
                    target_row = row_cache.get(chart_no)
                    if target_row is not None:
                        try:
                            _ = target_row.is_displayed()
                        except Exception:
                            print(f"    ⚠️ 快取 row 失效，重新掃表...")
                            scroll_to_load_all_rows(driver)
                            target_row = find_row_by_chart_no(driver, chart_no)
                    else:
                        print(f"    ⚠️ 快取查無 {chart_no}，重新掃表...")
                        scroll_to_load_all_rows(driver)
                        target_row = find_row_by_chart_no(driver, chart_no)

                    if target_row is None:
                        missing_chart_map[chart_no] = patient_name
                        update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status, patient_name, chart_no, "❌ 找不到病歷")
                        continue

                    ensure_window_focus(driver)
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target_row)
                    time.sleep(0.5)
                    ActionChains(driver).double_click(target_row).perform()
                    time.sleep(2)

                    try:
                        WebDriverWait(driver, 5).until(
                            EC.invisibility_of_element_located((By.CSS_SELECTOR, "tbody.p-datatable-tbody tr"))
                        )
                    except Exception:
                        pass
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "a.al-sidebar-list-link"))
                    )

                    # ✅ 解包新增的 final_soap
                    status, updated_name, final_soap = step_4_write_record(
                        driver, wait, chart_no, patient_name,
                        physician, physician_code, record_date, action_mode,
                        draft_on_herb=draft_on_herb,
                        name_ref=name_ref
                    )
                    patient_name = updated_name  

                    if status == "ghost_record":
                        ghost_record_map[chart_no] = patient_name
                        completed_map[chart_no] = patient_name
                        update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status, patient_name, chart_no, "✅ 幽靈強制暫存", final_soap)
                    elif status == "forced_draft":
                        forced_draft_map[chart_no] = patient_name
                        completed_map[chart_no] = patient_name
                        update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status, patient_name, chart_no, "✅ 醫師不符暫存", final_soap)
                    elif status == "chinese_herb":
                        chinese_herb_map[chart_no] = patient_name
                        completed_map[chart_no] = patient_name
                        update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status, patient_name, chart_no, "🌿 含中藥暫存", final_soap)
                    elif status == "discharged":
                        discharged_map[chart_no] = patient_name
                        completed_map[chart_no] = patient_name
                        update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status, patient_name, chart_no, "🏥 已出院(暫存)", final_soap)
                    else:
                        completed_map[chart_no] = patient_name
                        update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status, patient_name, chart_no, "✅ 完成", final_soap)

                    return_to_patient_list(driver, wait)
                    _uncheck_patient_row(chart_no)   
                    _rebuild_row_cache()

                except Exception as e:
                    patient_name = name_ref[0]  
                    error_msg = str(e)
                    if "使用者手持停止" in error_msg: raise e
                    elif "疑似新會病歷" in error_msg:
                        new_consult_map[chart_no] = patient_name
                        update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status, patient_name, chart_no, "🌟 疑似新會病歷")
                    elif "疑似今日未上傳門診紀錄" in error_msg:
                        missing_today_map[chart_no] = patient_name
                        update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status, patient_name, chart_no, "⚠️ 未上傳門診")
                    elif "今日病歷已存在" in error_msg:
                        exist_record_map[chart_no] = patient_name
                        completed_map[chart_no] = patient_name
                        update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status, patient_name, chart_no, "✅ 今日病歷已存")
                    else:
                        skipped_map[chart_no] = patient_name
                        update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status, patient_name, chart_no, "❌ 處理發生錯誤")
                    return_to_patient_list(driver, wait)
                    _rebuild_row_cache()

        def _run_unchecked_round():
            if not unchecked_patients_group:
                print(f"  → 無【未在清單】病人，跳過此輪。")
                return
            if not auto_add_flag:
                print(f"\n{'─'*52}")
                print(f"  ▶▶ 【未在清單】共 {len(unchecked_patients_group)} 位 → 自動新增已關閉，全部跳過")
                print(f"{'─'*52}")
                for p in unchecked_patients_group:
                    skipped_map[p["chart_no"]] = p["name"]
                    update_ui_patient_processed(txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status,
                                                p["name"], p["chart_no"], "⏭️ 未在清單中(跳過)")
                return
            print(f"\n{'─'*52}")
            print(f"  ▶▶ 【未在清單】共 {len(unchecked_patients_group)} 位（以病歷號搜尋查詢）")
            print(f"{'─'*52}")
            for p in unchecked_patients_group:
                check_stop()
                _process_one_patient_via_search(
                    p["chart_no"], p["name"],
                    physician, physician_code, record_date, action_mode,
                    txt_names, txt_charts, txt_proc_names, txt_proc_charts, txt_proc_status,
                    label_prefix="出院/新"
                )
                _rebuild_row_cache()

        def _uncheck_patient_row(chart_no_to_uncheck):
            if not auto_uncheck_flag:
                return
            try:
                ensure_window_focus(driver)
                rows_now = driver.find_elements(By.CSS_SELECTOR, "tbody.p-datatable-tbody tr")
                for _row in rows_now:
                    try:
                        _tds = _row.find_elements(By.TAG_NAME, "td")
                        if len(_tds) < 4: continue
                        if _tds[3].text.strip().lstrip("0") != chart_no_to_uncheck: continue
                        cb = _row.find_element(By.CSS_SELECTOR, ".p-checkbox-box")
                        if "p-highlight" in cb.get_attribute("class"):
                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", _row)
                            time.sleep(0.2)
                            driver.execute_script("arguments[0].click();", cb)
                            print(f"    ☑️ 已即時反勾選：{chart_no_to_uncheck}")
                        break
                    except Exception:
                        continue
            except Exception as _ue:
                print(f"    ⚠️ 即時反勾選失敗（略過）：{_ue}")

        def _show_group_summary_popup(g_idx, physician, checked_cnt, unchecked_cnt, priority_mode):
            done_event = threading.Event()
            def _show():
                top = tk.Toplevel(root)
                top.title(f"群組 {g_idx+1} 分組預覽")
                top.attributes("-topmost", True)
                w, h = 480, 240
                sw, sh = top.winfo_screenwidth(), top.winfo_screenheight()
                top.geometry(f"{w}x{h}+{(sw-w)//2}+{sh-h-160}")

                if priority_mode == "unchecked_first":
                    order_txt = "先搜尋 1.【未在清單】→ 再處理 2.【已在清單】"
                    clr = "#B71C1C"
                else:
                    order_txt = "先處理 2.【已在清單】→ 再搜尋 1.【未在清單】"
                    clr = "#1565C0"

                tk.Label(top, text=f"群組 {g_idx+1}｜主治醫師：{physician}", font=("Arial", 12, "bold")).pack(pady=(12, 4))
                tk.Label(top, text=f"✅ 已在清單（直接雙擊）：{checked_cnt} 位", font=("Arial", 11), fg="#1B5E20").pack()
                tk.Label(top, text=f"🔍 未在清單（需搜尋）：{unchecked_cnt} 位", font=("Arial", 11), fg="#E65100").pack()
                tk.Label(top, text=f"處理順序：{order_txt}", font=("Arial", 10, "bold"), fg=clr, wraplength=440).pack(pady=(6, 2))
                tk.Label(top, text="將於 4 秒後自動開始，或點擊按鈕立即開始", font=("Arial", 9), fg="gray").pack()

                def proceed():
                    done_event.set()
                    try: top.destroy()
                    except: pass

                tk.Button(top, text="▶ 立即開始", font=("Arial", 11, "bold"), bg="#4CAF50", fg="white",
                          command=proceed).pack(pady=8)

                def countdown(left):
                    if done_event.is_set(): return
                    if left > 0:
                        try: top.after(1000, countdown, left - 1)
                        except: pass
                    else: proceed()
                countdown(4)

            root.after(0, _show)
            done_event.wait()

        _show_group_summary_popup(g_idx, physician,
                                  len(checked_patients_group), len(unchecked_patients_group),
                                  priority_mode)

        def _switch_back_to_my_list_and_rebuild():
            try:
                print("    🔄 切換回「我的清單」並重建快取...")
                my_list_label = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//label[contains(normalize-space(), '我的清單')]")
                ))
                driver.execute_script("arguments[0].click();", my_list_label)
                time.sleep(2)
                print("      → 已切回我的清單")
            except Exception as e:
                print(f"      ⚠️ 切回我的清單失敗：{e}")
            _rebuild_row_cache()

        if priority_mode == "unchecked_first":
            _run_unchecked_round()
            if unchecked_patients_group and checked_patients_group:
                _switch_back_to_my_list_and_rebuild()
            _run_checked_round()
        else:
            _run_checked_round()
            if checked_patients_group and unchecked_patients_group:
                _switch_back_to_my_list_and_rebuild()
            _run_unchecked_round()

        threading.Thread(
            target=send_discord_progress,
            args=(f"📋 群組 {g_idx+1} 完成｜{physician}",
                  f"已完成 {physician} 群組（共 {len(checked_patients_group)+len(unchecked_patients_group)} 位病人）"),
            daemon=True
        ).start()

    msg, has_warnings = generate_current_report()
    if has_warnings: final_msg = "⚠️ 執行完畢 (含例外狀況)\n\n" + msg
    else: final_msg = "✅ 所有病人處理完畢，清單核對一致！\n\n" + msg
        
    print("\n=== 第三步驟完成，準備自動關閉 ===")

    threading.Thread(
        target=send_discord_progress,
        args=("🎉 全部群組處理完畢！", ""),
        daemon=True
    ).start()

    root.after(0, lambda: btn_start.config(state="normal", text="開始執行"))
    root.after(0, lambda: btn_pause.config(state="disabled", text="暫停執行 (Alt+S)", bg="#FF9800"))
    root.after(0, lambda: btn_stop.config(state="disabled", text="停止並重置 (Alt+A)"))
    final_countdown_and_close(driver, final_msg)

def step_2_next_actions(driver, wait, groups, action_mode, auto_add_flag, auto_uncheck_flag, draft_on_herb, priority_mode="checked_first"):
    print("\n— 進入第二步驟：開啟住院醫囑系統 —")
    try:
        check_stop()
        try:
            driver.maximize_window()
        except:
            pass

        programs_btn = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//a[contains(@class, 'dropdown-toggle') and (.//span[contains(text(), '程式集')] or .//i[contains(@class, 'glyphicon-th')])]")
        ))
        driver.execute_script("arguments[0].click();", programs_btn)
        time.sleep(1.5)

        inpatient_sys_btn = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//*[contains(@class, 'app-font') and contains(., '住院醫囑系統')] | //span[contains(text(), '住院醫囑系統')]")
        ))
        driver.execute_script("arguments[0].click();", inpatient_sys_btn)
        
        countdown_popup(10)
        
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])

        step_3_process_patients(driver, wait, groups, action_mode,
                                auto_add_flag, auto_uncheck_flag,
                                draft_on_herb=draft_on_herb,
                                priority_mode=priority_mode)

    except Exception as e:
        if "使用者手動停止" in str(e): raise e
        final_countdown_and_close(driver, f"第二步驟發生錯誤：\n{e}")

def step_1_login(emp_id, emp_pwd, groups, action_mode, auto_add_flag, auto_uncheck_flag, draft_on_herb, priority_mode="checked_first"):
    print("\n— 進入第一步驟：系統登入 —")
    keep_system_awake()
    options = webdriver.EdgeOptions()
    options.add_argument("--start-maximized") 
    
    prefs = {
        "protocol_handler.excluded_schemes.runpcallmainp": False,
        "custom_handlers.enabled": True
    }
    options.add_experimental_option("prefs", prefs)

    try:
        from selenium.webdriver.edge.service import Service
        from webdriver_manager.microsoft import EdgeChromiumDriverManager
        service = Service(EdgeChromiumDriverManager().install())
        driver = webdriver.Edge(service=service, options=options)
    except Exception as e:
        print(f"    ⚠️ webdriver-manager 初始化失敗，嘗試直接啟動：{e}")
        driver = webdriver.Edge(options=options)

    try:
        driver.maximize_window()
    except:
        pass

    wait = WebDriverWait(driver, 25)

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

        step_2_next_actions(driver, wait, groups, action_mode,
                            auto_add_flag, auto_uncheck_flag,
                            draft_on_herb=draft_on_herb,
                            priority_mode=priority_mode)

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

def show_disclaimer():
    top = tk.Toplevel(root)
    top.title("系統操作說明與免責聲明")
    top.attributes("-topmost", True)
    top.geometry("850x760") 

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
                img.thumbnail((350, 250))
                photo = ImageTk.PhotoImage(img)
                lbl_img = tk.Label(top, image=photo)
                lbl_img.image = photo
                lbl_img.pack(pady=(10, 5))
        except Exception:
            pass

    tk.Label(top, text="Copyright and Developed by PBCM-38 譚皓宇", font=("Arial", 10), fg="gray").pack(pady=(0, 5))

    txt = scrolledtext.ScrolledText(top, width=100, height=22, font=("Arial", 11), bg="#fdfdfd")
    txt.pack(padx=20, pady=5, fill="both", expand=True)

    txt.tag_configure("title", font=("Arial", 14, "bold"), foreground="#2C3E50", spacing3=5)
    txt.tag_configure("step", font=("Arial", 12, "bold"), foreground="#C0392B")
    txt.tag_configure("h1", font=("Arial", 12, "bold"), foreground="#2980B9", spacing1=10, spacing3=5)
    txt.tag_configure("green_title", font=("Arial", 11, "bold"), foreground="#27AE60", spacing1=5)
    txt.tag_configure("yellow_title", font=("Arial", 11, "bold"), foreground="#D35400", spacing1=5)
    txt.tag_configure("red_title", font=("Arial", 11, "bold"), foreground="#C0392B", spacing1=5)
    txt.tag_configure("bold", font=("Arial", 11, "bold"), foreground="#333333")
    txt.tag_configure("highlight", font=("Arial", 11, "bold"), foreground="#16A085")
    txt.tag_configure("gray", font=("Arial", 10), foreground="#7F8C8D", spacing1=2)
    txt.tag_configure("example_box", font=("Consolas", 10), foreground="#34495E", background="#EAEDED", lmargin1=10, lmargin2=10)

    content = [
        ("操作與安裝說明:\n", "title"),
        ("第一步: ", "step"), ("先載壓縮檔.7z 再進去找到(執行檔)exe\n", "bold"),
        ("第二步: ", "step"), ("病患處理類型與介紹:\n", "bold"),

        ("=== Part-1：病歷處理分類類型 ===\n", "h1"),
        ("系統在處理每一位病人時，會於介面「處理狀態」欄位顯示對應結果，幫助您快速掌握病歷狀況，分為以下三大類：\n", ""),

        ("一、 🟢 順利完成類\n", "green_title"),
        (" ✅ 完成：", "highlight"), ("成功抓取門診紀錄與歷史 Ditto，並順利執行送件/暫存。\n", ""),
        (" ✅ 今日病歷已存：", "highlight"), ("系統發現今日已開立過「中醫病程記錄」，為避免重複將自動跳過。\n", ""),

        ("二、 🟡 防護與強制暫存類（觸發防呆機制）\n", "yellow_title"),
        ("這類病人雖處理完畢，但偵測到特殊狀況，系統會自動攔截並「強制改為暫存」：\n", ""),
        (" ✅ 幽靈強制暫存：", "bold"), ("因為查無該病患今日任何「門診紀錄」，可能是不符合健保使用狀況的病人所以沒有紀錄。所以，系統會從預設字庫隨機生成 S 欄防報錯，並強制暫存。\n", ""),
        (" ✅ 醫師不符暫存：", "bold"), ("送出時因門診紀錄上的醫師名稱與您輸入的不符。為防呆避免送錯，改強制暫存。\n", ""),
        (" 🌿 含中藥暫存：", "bold"), ("P 欄偵測到中藥處方，觸發中藥爬蟲。已自動抓取藥囑並排版完成，基於安全考量強制暫存。\n", ""),
        (" 🏥 已出院(暫存)：", "bold"), ("比對出院日期與”已出院無法存檔狀況”，若偵測到病患已出院，會自動在 S 欄末尾補上「預計於今日出院。」並強制暫存。\n", ""),

        ("三、 🔴 無法處理需人工介入類\n", "red_title"),
        ("系統遇到無法自動克服的障礙，將跳過該病人，請醫師手動處理：\n", ""),
        (" 🌟 疑似新會病歷：", "bold"), ("過去 7 天內找不到「中醫病程」可 Ditto，需手動從零開立。\n", ""),
        (" ⚠️ 未上傳門診：", "bold"), ("查無今日的門診文字紀錄（可能尚未打完或上傳）。\n", ""),
        (" ❌ 找不到病歷/查無此病歷號：", "bold"), ("病歷號格式錯誤，或系統搜尋不到住院資料。\n", ""),
        (" ❌ 處理發生錯誤：", "bold"), ("過程中發生非預期的網頁卡頓或系統延遲。\n", ""),
        (" ⏭️ 未在清單中(跳過)：", "bold"), ("未勾選「自動新增」，遇到不在勾選清單裡的病歷號時直接跳過。\n", ""),

        ("=== Part-2：SOAP 處理邏輯 ===\n", "h1"),
        ("本系統抓取的資料是 「今日門診紀錄」與「歷史中醫病程 (Ditto)」各自比對並提取資料，接著進行邏輯判斷來過濾資料。(如果門診紀錄不對、四診格式有錯、P沒有用#排版..等)\n", ""),
        (" 📌【S 欄】主觀 / 主訴：", "bold"), ("擷取”今日門診病歷” S 欄，過濾舊日期戳記，篩選當天日期後的S內容，假設今天是4/1 ，他就會找到門診紀錄S裡面符合 (4/1) 頭痛....腹脹...這一段。若出院則自動補上出院提示。 (門診紀錄的S日期戳記要有括號喔，不然偵測不到)\n", ""),
        (" 📌【O 欄】客觀 / 理學檢查：", "bold"), ("優先採用「歷史 Ditto」O 欄，自動刪除過時的生命徵象數值，保留中醫四診。送出前會自動點擊帶入最新生命徵象！\n", ""),
        (" 📌【A 欄】評估 / 診斷：", "bold"), ("直接沿用過去 7 天內最近一次 Ditto 的 A 欄，不做額外更動。\n", ""),
        (" 📌【P 欄】治療計畫：", "bold"), ("擷取今日門診 P 欄，清除健保宣告冗字並重新排版（自動換行與空行）；若含中藥則自動彈出「藥囑紀錄」爬取並計算日劑量。\n", ""),

        ("=== part-3：範例 ===\n", "h1"),
        ("我們假設今天是 2026-04-01（民國 115 年 4 月 1 日），系統正在處理一位有開立「科學中藥」的住院會診病人。\n以下是系統運作的解析範例：\n\n", ""),
        ("📥 【爬取前】HIS 系統上的原始資料\n", "bold"),
        ("系統會去抓取「今日的門診紀錄」以及「過去 7 天內的 Ditto 歷史中醫病程」，這時抓到的原始文字通常非常雜亂：\n", ""),
        ("【S】（來自今日門診紀錄）\n…………..(115/03/29) 納差，腹脹，睡眠不佳。(115/04/01) 胃口稍改善，大便已成形，昨晚睡得較安穩。\n【O】（來自歷史 Ditto 病程）\n(115/03/31) Vital signs: Temperature: 36.6 ℃ BP: 125 / 82 mmHg Pulse: 78 per min SpO2: 98 % 望診：神明清，面色微黃。 聞診：語聲正常，無特殊氣味。 舌診：舌淡紅，苔薄白。 切診：脈弦細。 報告時間：2026-03-31 08:30 檢驗單：WBC 6.8, Hb 12.1\n【A】（來自歷史 Ditto 病程）\nSleep disorder\nConstipation\nFunctional dyspepsia\n【P】（來自今日門診紀錄）\n患者符合中醫健保特定疾病住院會診加強照護計畫(115/01/01~115/12/31) #針灸治療：合谷、足三里、太衝。留針 15 分鐘。 #科學中藥 (115/04/01) Time out\n\n", "example_box"),
        ("📤 【爬取後】系統清洗、排版、並寫入的新病歷\n", "bold"),
        ("經過本系統的 Regex（正則表達式）與自動爬蟲邏輯處理後，貼到今日病歷上的格式會變成這樣：\n", ""),
        ("【S】\n胃口稍改善，大便已成形，昨晚睡得較安穩。\n【O】\n[系統同時偵測到原始資料裡有「報告時間、檢驗單」等字眼，於是觸發 needs_new_reports = True，在病歷送出前，機器人會自動去點擊右側的「生命徵象」與「檢查報告」按鈕，幫你把「最新、當下」的數值自動貼上！]\n望診：神明清，面色微黃。\n聞診：語聲正常，無特殊氣味。\n舌診：舌淡紅，苔薄白。\n切診：脈弦細。\n【A】\nSleep disorder\nConstipation\nFunctional dyspepsia\n【P】\n#針灸治療：合谷、足三里、太衝。留針 15 分鐘。\n#科學中藥 加味逍遙散4.00 GM BID 酸棗仁湯3.00 GM BID\n\n", "example_box"),

        ("=== part-4：進階設定與快捷鍵 ===\n", "h1"),
        (" • 快捷鍵功能：\n", "bold"),
        ("   - Alt+S：暫停 / 繼續執行\n   - Alt+A：停止並重置\n   - Ctrl+Alt+F：點擊程式空白處按下此組合鍵，可解鎖隱藏的進階設定。\n", ""),
        (" • Discord 通知：", "bold"), ("解鎖進階設定後填入 Webhook URL，系統能在執行完畢後自動推送結算報告至您的頻道。\n", ""),

        ("=== part-5：開發者免責聲明 ===\n", "h1"),
        (" 1. 本系統本為複製並撰寫針灸科病程使用，不具主動外傳病人資料或令裝置感染病毒之代碼功能。\n 2. 內容與流程(SOAP)基於開發者於 2026/03 之前的經驗實作，後續若有變動請自行修正。\n 3. 系統使用爬蟲分析病歷結構(學長姊的門診紀錄格是很重要，沒照格式打會找不到)，若遇網頁位移、醫院系統更新等干擾，可能產生 BUG，請自行留意。\n 4. 本系統僅為輔助撰寫病歷，請務必親自為您負責的病歷做最終確認！\n 5. 操作若有問題可以再聯繫 BY 譚哥 (PBCM-38)\n", "gray"),

        ("=== 其他細節(可忽略) ===\n", "h1"),
        ("1. 判斷是否「強制改暫存」 (防呆機制)\n", "bold"),
        ("這個邏輯對應到介面上「🌿 內容有中藥治療時改暫存」的打勾選項。\n觸發條件： 必須在介面上勾選該選項 (draft_on_herb=True)。\n搜尋範圍： 優先檢查「歷史 Ditto 的 P 欄」。如果歷史紀錄沒有 P 欄，則退而求其次檢查「今日門診紀錄的 P 欄」。\n關鍵字清單： 只要文字中包含以下任一個詞彙，就會判定為含中藥，並將最後的動作強制改成「暫存」：\n中藥治療 / 中藥水治療 / 中藥水 / 科學中藥 / 水藥\n\n", "gray"),
        ("2. 判斷是否「觸發自動抓取藥單」 (爬蟲機制)\n", "bold"),
        ("這個邏輯是在處理「今日門診紀錄 P 欄排版」時自動運作的，不管上有沒有勾選暫存都會執行。\n觸發條件： 掃描今日門診紀錄的 P 欄，尋找開頭帶有 # 標籤的文字行（例如：#中藥治療）。\n關鍵字清單： 只要該 # 標籤行包含以下任一個詞彙，系統就會認定這是中藥處方區塊：\n中藥 / 科學中藥 / 水藥\n執行動作： 一旦觸發，系統會自動切換去點擊「醫藥囑相關紀錄」>「藥囑紀錄」，把當天的中醫藥品明細、頻率抓下來，計算好日劑量（對齊 0.5 倍數），然後整齊地貼回 P 欄的下方。\n簡單總結： 系統是靠抓取 P 欄裡的特定字眼來判斷的。一旦看到這些字眼，排版時就會去爬藥單；如果有勾選防呆選項，送出前就會把狀態改成暫存。\n", "gray"),
    ]

    for text, tag in content:
        if tag:
            txt.insert(tk.END, text, tag)
        else:
            txt.insert(tk.END, text)

    txt.config(state="disabled")

    btn_close = tk.Button(top, text="我已了解並關閉", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", command=top.destroy)
    btn_close.pack(pady=10)

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
    txt_names = scrolledtext.ScrolledText(col_name, width=15, height=6)
    txt_names.pack()

    col_chart = tk.Frame(row_pt)
    col_chart.pack(side="left", padx=(0, 4))
    tk.Label(col_chart, text="待處理患者病歷號 *", font=("Arial", 9, "bold"), fg="#c0392b").pack(anchor="w")
    txt_charts = scrolledtext.ScrolledText(col_chart, width=23, height=6)
    txt_charts.pack()

    col_proc_name = tk.Frame(row_pt)
    col_proc_name.pack(side="left", padx=(4, 4))
    tk.Label(col_proc_name, text="已處理姓名床位", font=("Arial", 9, "bold"), fg="#4CAF50").pack(anchor="w")
    txt_proc_names = scrolledtext.ScrolledText(col_proc_name, width=23, height=6)
    txt_proc_names.pack()

    col_proc_chart = tk.Frame(row_pt)
    col_proc_chart.pack(side="left", padx=(0, 4))
    tk.Label(col_proc_chart, text="已處理病歷號", font=("Arial", 9, "bold"), fg="#4CAF50").pack(anchor="w")
    txt_proc_charts = scrolledtext.ScrolledText(col_proc_chart, width=18, height=6)
    txt_proc_charts.pack()

    col_proc_status = tk.Frame(row_pt)
    col_proc_status.pack(side="left")
    tk.Label(col_proc_status, text="處理狀態", font=("Arial", 9, "bold"), fg="#4CAF50").pack(anchor="w")
    txt_proc_status = scrolledtext.ScrolledText(col_proc_status, width=24, height=6)
    txt_proc_status.pack()

    def sync_scroll(*args):
        txt_proc_names.yview_moveto(args[0])
        txt_proc_charts.yview_moveto(args[0])
        txt_proc_status.yview_moveto(args[0])

    txt_proc_names.vbar.config(command=sync_scroll)
    txt_proc_charts.vbar.config(command=sync_scroll)
    txt_proc_status.vbar.config(command=sync_scroll)

    group_frames.append({
        "frame":         frame,
        "physician":     ent_physician,
        "code":          ent_code,
        "year":          ent_year,
        "date":          ent_date,
        "names":         txt_names,
        "charts":        txt_charts,
        "proc_names":    txt_proc_names,
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
    apply_advanced_settings()
    emp_id  = entry_id.get().strip()
    emp_pwd = entry_pwd.get().strip()
    action_mode = action_var.get()

    auto_add_flag    = auto_add_var.get()
    auto_uncheck_flag = auto_uncheck_var.get()
    draft_on_herb    = draft_on_herb_var.get()
    priority_mode    = priority_mode_var.get()

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

        grp["proc_names"].delete("1.0", tk.END)
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
                "txt_proc_names": grp["proc_names"],
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

    threading.Thread(
        target=step_1_login,
        args=(emp_id, emp_pwd, groups, action_mode,
              auto_add_flag, auto_uncheck_flag, draft_on_herb, priority_mode),
        daemon=True
    ).start()

def toggle_pause():
    global status_window
    if pause_event.is_set():
        pause_event.clear()
        btn_pause.config(text="繼續執行 (Alt+S)", bg="#2196F3")
        root.update_idletasks()
        print("\n    ⏳ 已收到暫停指令，將在到達下一個安全檢查點時暫停…")

        threading.Thread(
            target=send_discord_progress,
            args=("⏸️ HIS 系統已暫停", "使用者手動暫停執行，以下為目前處理進度："),
            daemon=True
        ).start()

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
    print("\n🛑 已收到停止指令，正在安全中止程序並重置視窗…")
    stop_event.set()
    pause_event.set()
    btn_pause.config(state="disabled")
    btn_stop.config(state="disabled", text="停止中…")
    root.update_idletasks()

    def reset_ui():
        if saved_initial_state:
            for state in saved_initial_state:
                grp = state["grp"]
                grp["names"].delete("1.0", tk.END)
                if state["names"]: grp["names"].insert(tk.END, state["names"])
                
                grp["charts"].delete("1.0", tk.END)
                if state["charts"]: grp["charts"].insert(tk.END, state["charts"])
                    
                grp["proc_names"].delete("1.0", tk.END)
                grp["proc_charts"].delete("1.0", tk.END)
                grp["proc_status"].delete("1.0", tk.END)
        
        btn_start.config(state="normal", text="開始執行")
        btn_pause.config(state="disabled", text="暫停執行 (Alt+S)", bg="#FF9800")
        btn_stop.config(state="disabled", text="停止並重置 (Alt+A)")
        print("    ✅ 已成功重置為最一開始執行前的參數狀態。")
        
    root.after(1500, reset_ui)

# ==========================================
# 建立主視窗與解鎖功能
# ==========================================

root = tk.Tk()
root.title("HIS 病歷自動化系統")
root.geometry("1500x850")
root.configure(padx=20, pady=20)

top_header_frame = tk.Frame(root)
top_header_frame.pack(fill="x", pady=(0, 15))
tk.Label(top_header_frame, text="系統登入參數", font=("Arial", 14, "bold")).pack(side="left")
btn_info = tk.Button(top_header_frame, text="📖 使用說明", font=("Arial", 9), command=show_disclaimer)
btn_info.pack(side="right")

frame_form = tk.Frame(root)
frame_form.pack(fill="x")

var_id = tk.StringVar()
var_pwd = tk.StringVar()

def check_secret_unlock(*args):
    if var_id.get() == "kronioel":
        unlock_advanced_settings("kronioel")

var_id.trace_add("write", check_secret_unlock)
var_pwd.trace_add("write", check_secret_unlock)

tk.Label(frame_form, text="員工代號:").grid(row=0, column=0, sticky="e", pady=5)
entry_id = tk.Entry(frame_form, width=20, textvariable=var_id)
entry_id.grid(row=0, column=1, padx=10)

tk.Label(frame_form, text="系統密碼:").grid(row=1, column=0, sticky="e", pady=5)
entry_pwd = tk.Entry(frame_form, width=20, show="*", textvariable=var_pwd)
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

auto_add_var = tk.BooleanVar(value=True)
tk.Checkbutton(frame_action, text="自動新增未在清單中的病歷號", variable=auto_add_var, font=("Arial", 10)).pack(side="left", padx=(15, 0))

auto_uncheck_var = tk.BooleanVar(value=False)
tk.Checkbutton(frame_action, text="病歷完成後自動反勾選", variable=auto_uncheck_var, font=("Arial", 10)).pack(side="left", padx=(15, 0))

draft_on_herb_var = tk.BooleanVar(value=False)
tk.Checkbutton(
    frame_action,
    text="🌿 內容有中藥治療時改暫存",
    variable=draft_on_herb_var,
    font=("Arial", 10)
).pack(side="left", padx=(15, 0))

# ✅ 新增：保留P欄健保計畫開頭按鈕
keep_jianbao_var = tk.BooleanVar(value=True)
tk.Checkbutton(
    frame_action,
    text="保留P欄健保頭",
    variable=keep_jianbao_var,
    font=("Arial", 10)
).pack(side="left", padx=(15, 0))


frame_priority = tk.Frame(root)
frame_priority.pack(fill="x", pady=(0, 4))

tk.Label(frame_priority, text="處理順序：", font=("Arial", 10, "bold")).pack(side="left")

priority_mode_var = tk.StringVar(value="checked_first")

tk.Radiobutton(
    frame_priority,
    text="先處理 2.【已在清單中病人】→ 再搜尋 1.【未在清單病人】（預設）",
    variable=priority_mode_var,
    value="checked_first",
    font=("Arial", 10),
    fg="#1565C0"
).pack(side="left")

tk.Radiobutton(
    frame_priority,
    text="先搜尋 1.【未在清單病人】→ 再處理 2.【已在清單中病人】",
    variable=priority_mode_var,
    value="unchecked_first",
    font=("Arial", 10),
    fg="#B71C1C"
).pack(side="left", padx=(12, 0))

# ==========================================
# 進階設定預設隱藏與解鎖邏輯
# ==========================================

adv_expanded = tk.BooleanVar(value=False)
adv_outer = tk.Frame(root, bd=1, relief="groove")

is_unlocked = False

def unlock_advanced_settings(method=""):
    global is_unlocked
    if not is_unlocked:
        is_unlocked = True
        adv_outer.pack(fill="x", pady=(6, 2), before=btn_frame)
        print(f"🔓 已觸發隱藏彩蛋：進階設定已解鎖！({method})")
    
    if method == "kronioel" and entry_discord.get().strip() == "":
        entry_discord.delete(0, tk.END)
        entry_discord.insert(0, "https://discord.com/api/webhooks/1487777904124231720/kI89PBoiQl-LsY15nPBDGZABhyEk8--rkpBL3dT6Tl6eDRQse05fC8-qewa2K--OMUWt")
        print("🔗 已為專屬帳戶自動載入 Webhook 網址。")

def toggle_advanced():
    if adv_expanded.get():
        adv_body.pack_forget()
        adv_toggle_btn.config(text="▶ 進階設定")
        adv_expanded.set(False)
    else:
        adv_body.pack(fill="x", padx=10, pady=(0, 8))
        adv_toggle_btn.config(text="▼ 進階設定")
        adv_expanded.set(True)

adv_header = tk.Frame(adv_outer)
adv_header.pack(fill="x")
adv_toggle_btn = tk.Button(
    adv_header, text="▶ 進階設定",
    font=("Arial", 10, "bold"), fg="#555", relief="flat",
    bg="#f0f0f0", activebackground="#e0e0e0",
    command=toggle_advanced, anchor="w"
)
adv_toggle_btn.pack(fill="x", padx=6, pady=4)

adv_body = tk.Frame(adv_outer)

discord_frame = tk.Frame(adv_body)
discord_frame.pack(fill="x", pady=(6, 2))

tk.Label(discord_frame, text="Discord Webhook URL：", font=("Arial", 10, "bold"), width=22, anchor="e").pack(side="left")

entry_discord = tk.Entry(discord_frame, width=80, show="")
entry_discord.pack(side="left", padx=(4, 6))

lbl_discord_status = tk.Label(discord_frame, text="", font=("Arial", 9), fg="gray")
lbl_discord_status.pack(side="left")

def test_discord_webhook():
    url = entry_discord.get().strip()
    if not url:
        lbl_discord_status.config(text="⚠️ 請先填入 Webhook URL", fg="orange")
        return
    lbl_discord_status.config(text="傳送中…", fg="gray")
    root.update_idletasks()
    def _test():
        try:
            content = "✅ HIS 病歷系統 Discord 通知測試成功！"
            payload = json.dumps({"content": content}).encode("utf-8")
            
            req = urllib.request.Request(
                url, 
                data=payload,
                headers={
                    "Content-Type": "application/json",
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
                }, 
                method="POST"
            )
            urllib.request.urlopen(req, timeout=10)
            root.after(0, lambda: lbl_discord_status.config(text="✅ 測試成功！", fg="#4CAF50"))
        except Exception as e:
            root.after(0, lambda: lbl_discord_status.config(text=f"❌ 失敗：{e}", fg="red"))
    threading.Thread(target=_test, daemon=True).start()

tk.Button(discord_frame, text="測試", font=("Arial", 9), bg="#9C27B0", fg="white",
          command=test_discord_webhook).pack(side="left")

tk.Label(adv_body,
         text="💡 到 Discord 頻道設定 → 整合 → Webhook → 新增 Webhook → 複製連結後貼上",
         font=("Arial", 8), fg="gray", justify="left"
         ).pack(anchor="w", padx=4, pady=(0, 4))

# --- 每完成一位病人通知選項 ---
notify_frame = tk.Frame(adv_body)
notify_frame.pack(fill="x", pady=(2, 6), padx=4)

notify_per_patient_var = tk.BooleanVar(value=False)
chk_notify = tk.Checkbutton(
    notify_frame,
    text="🔔 每完成一位病人就發送 Discord 通知（預設關閉）",
    variable=notify_per_patient_var,
    font=("Arial", 10)
)
chk_notify.pack(side="left")

tk.Label(
    notify_frame,
    text="  ⚠️ 訊息量大時請謹慎啟用",
    font=("Arial", 8), fg="orange"
).pack(side="left")

def apply_advanced_settings():
    global discord_webhook_url, notify_per_patient
    if is_unlocked:
        discord_webhook_url = entry_discord.get().strip()
        notify_per_patient = notify_per_patient_var.get()
    else:
        discord_webhook_url = ""
        notify_per_patient = False

# ==========================================
# 執行按鈕區塊
# ==========================================

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
    keyboard.add_hotkey('ctrl+alt+f', lambda: root.after(0, lambda: unlock_advanced_settings("hotkey")))
    print("⌨️ 全域快捷鍵已註冊成功！(Alt+S: 暫停/繼續, Alt+A: 停止並重置, Ctrl+Alt+F: 解鎖進階)")
except Exception as e:
    print(f"⚠️ 快捷鍵註冊失敗，請確認是否已安裝 keyboard 套件: {e}")

lbl_copyright = tk.Label(root, text="Copyright and Developed by PBCM-38 譚皓宇", font=("Arial", 8), fg="gray")
lbl_copyright.pack(side="bottom", pady=5)

root.mainloop()
