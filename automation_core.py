import requests
import json
import re
import base64
import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
import time
import sys
from datetime import datetime
from openpyxl import load_workbook
import multiprocessing as mp
import random
import urllib3
urllib3.disable_warnings()

BASE    = "https://10.20.129.197/ftthPortalBE"
SEARCH  = BASE + "/dataSearchAndViewServiceNumberData/searchingWOAndVC4"
REFRESH = BASE + "/refresh"

PREFIXES = {"2", "3", "13", "40", "48", "55", "62", "64", "46", "66", "68", "69",
            "50", "45", "57", "47", "92", "65", "84", "82", "86", "88", "93", "95", "96", "97"}

# ------------------------------------------------------------------ parse pasted cURL
def parse_curl(paste: str):
    """Extracts (bearer, cookie_dict) from a full 'Copy as cURL' paste, or a bare token."""
    paste = paste.strip()
    m = re.search(r"Bearer\s+([A-Za-z0-9\-_.]+)", paste)
    if m:
        bearer = m.group(1)
    elif re.fullmatch(r"[A-Za-z0-9\-_.]{80,}", paste) and paste.count(".") == 2:
        bearer = paste                      # user pasted just the token
    else:
        raise ValueError("لم يتم العثور على التوكن — الصق الـ curl كاملاً أو التوكن فقط")

    cookies = {}
    m = re.search(r"-b\s+'([^']*)'", paste)
    if not m:
        m = re.search(r'-b\s+"([^"]*)"', paste)
    if m:
        for part in m.group(1).split(";"):
            if "=" in part:
                k, v = part.split("=", 1)
                cookies[k.strip()] = v.strip()
    return bearer, cookies

def build_session(bearer, cookies):
    """Create a session ONCE for the whole run (cookies + Set-Cookie rotation handled automatically)."""
    s = requests.Session()
    s.verify = False
    for k, v in cookies.items():
        s.cookies.set(k, v)
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/155.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Referer": "https://10.20.129.197/ftthPortal/",
        "Authorization": "Bearer " + bearer,
    })
    return s

# ------------------------------------------------------------------ token refresh
def jwt_exp(token: str):
    """Return 'exp' timestamp from the JWT payload (no verification)."""
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload)).get("exp")
    except Exception:
        return None

def refresh_token(session, log_queue=None):
    """
    Call /refresh with the current session (current bearer + cookies).
    Extract new token from response JSON key 'result' ("Bearer eyJ..."),
    update the session Authorization header IN PLACE.
    Returns new token string. Raises PermissionError if refresh fails.
    """
    r = session.get(REFRESH, timeout=20)
    if r.status_code != 200:
        raise PermissionError(f"refresh failed: HTTP {r.status_code}")
    try:
        new_token = r.json().get("result", "")
    except ValueError:
        raise PermissionError("refresh response was not JSON")
    if new_token.lower().startswith("bearer "):
        new_token = new_token[7:]
    if not new_token or new_token.count(".") != 2:
        raise PermissionError("refresh response had no valid token")
    session.headers["Authorization"] = "Bearer " + new_token
    if log_queue:
        exp = jwt_exp(new_token)
        mins = int(exp - time.time()) if exp else "?"
        log_queue.put(f"تم تجديد الجلسة (صالحة ~{mins} دقيقة)")
    return new_token

def split_number(full_number):
    full_number = full_number.strip()
    for prefix in sorted(PREFIXES, key=len, reverse=True):
        if full_number.startswith(prefix):
            return "0" + prefix, full_number[len(prefix):]
    return "", full_number

# ------------------------------------------------------------------ excel writer
def excel_writer_listener(queue, log_queue, results_file, draft_file):
    headers = ["FULLNUMBER", "INFORMATION"]
    for f in [results_file, draft_file]:
        if not os.path.exists(f):
            pd.DataFrame(columns=headers).to_excel(f, index=False)
    while True:
        item = queue.get()
        if item == "STOP":
            break
        line_number, record = item
        target_row = line_number + 2          # header row + Excel 1-based
        for f in [results_file, draft_file]:
            wb = load_workbook(f)
            sheet = wb.active
            for col_idx, value in enumerate(record, start=1):
                sheet.cell(row=target_row, column=col_idx).value = value
            wb.save(f)
        log_queue.put(f"تم معالجة السطر {line_number + 1}")

def choose_excel_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
    root.destroy()
    return file_path if file_path else None

def init_results_file(original_file_path):
    dir_name = os.path.dirname(original_file_path)
    base_name = os.path.splitext(os.path.basename(original_file_path))[0]
    results_file = os.path.join(dir_name, f"{base_name}_results_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    draft_file   = os.path.join(dir_name, f"{base_name}_draft_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    pd.DataFrame(columns=["FULLNUMBER", "INFORMATION"]).to_excel(results_file, index=False)
    pd.DataFrame(columns=["FULLNUMBER", "INFORMATION"]).to_excel(draft_file, index=False)
    return results_file, draft_file

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()

# ------------------------------------------------------------------ worker
def worker_task(session, index, row_data, queue, stop_flag, log_queue):
    if stop_flag.is_set():
        return
    full_number = str(row_data[0])
    full_number = full_number if full_number.startswith("0") else "0" + full_number
    landline_code, landline_number = split_number(str(row_data[0]))

    time.sleep(random.uniform(0.1, 0.9))

    r = None
    data = None

    # ---- attempt loop: 1st try, and one retry after refresh if 401/403 ----
    for attempt in range(2):
        if stop_flag.is_set():
            return
        try:
            r = session.get(SEARCH, params={"cityCode": landline_code,
                                            "serviceNumber": landline_number}, timeout=20)
        except requests.RequestException as e:
            queue.put((index, [full_number, f"error: {e}"]))
            return

        if r.status_code in (401, 403) or "اسم المستخدم" in r.text:
            if attempt == 0:
                # token expired -> refresh and retry THIS SAME number once
                try:
                    refresh_token(session, log_queue)
                    continue          # retry the same request with the new token
                except (PermissionError, requests.RequestException) as e:
                    log_queue.put(f"فشل تجديد الجلسة: {e}")
                    break             # fall through to the "session dead" handling below
            break                     # second 401 in a row -> session really dead

        break                         # got a usable response (any status != 401/403)

    # ---- session dead -> notify and stop the whole run ----
    if r is not None and (r.status_code in (401, 403) or "اسم المستخدم" in r.text):
        log_queue.put(f"انتهت الجلسة نهائياً عند السطر {index+1}")
        log_queue.put(f"برجاء بدء العملية مرة اخرى بداية من السطر {index+1}")
        stop_flag.set()
        return

    if r.status_code != 200:
        queue.put((index, [full_number, f"error HTTP {r.status_code}: {r.text[:100]}"]))
        return

    try:
        data = r.json()
    except ValueError:
        queue.put((index, [full_number, f"error: غير JSON — {r.text[:100]}"]))
        return

    if data.get("code") == 1:
        queue.put((index, [full_number, "Orange"]))
    elif data.get("code") == 4:
        msg = data.get("errorMessage") or ""
        if "Not Accepted" in msg:
            queue.put((index, [full_number, "Free"]))
        elif "another ISP" in msg:
            queue.put((index, [full_number, "Another Isp"]))
        else:
            queue.put((index, [full_number, f"code4: {msg}"]))
    else:
        queue.put((index, [full_number, f"unknown code {data.get('code')}"]))

# ------------------------------------------------------------------ runner
def run_bot(curl_paste, start_line, excel_file, log_queue, stop_flag):
    try:
        results_file, draft_file = init_results_file(excel_file)

        try:
            df = pd.read_excel(excel_file, header=None)
        except Exception as e:
            log_queue.put(f"Could not read Excel file: {e}")
            log_queue.put("تم انهاء العملية")
            stop_flag.set()
            return

        try:
            bearer, cookies = parse_curl(curl_paste)
        except ValueError as e:
            log_queue.put(str(e))
            log_queue.put("تم انهاء العملية")
            stop_flag.set()
            return

        session = build_session(bearer, cookies)

        # sanity check before starting: try one refresh now (also confirms /refresh works)
        try:
            refresh_token(session, log_queue)
        except Exception as e:
            log_queue.put(f"فشل تجديد الجلسة عند البداية: {e} — تأكد من صحة الـ curl")
            log_queue.put("تم انهاء العملية")
            stop_flag.set()
            return

        log_queue.put("تم تجهيز الجلسة — جاري معالجة البيانات")

        start_time = time.time()
        start_index = int(start_line) - 1
        tasks = [(index, list(row)) for index, row in df.iloc[start_index:].iterrows()]

        with mp.Manager() as internal_manager:
            results_queue = internal_manager.Queue()
            watcher = mp.Process(target=excel_writer_listener,
                                 args=(results_queue, log_queue, results_file, draft_file))
            watcher.start()

            with mp.Pool(1) as pool:
                pool.starmap(worker_task, [
                    (session, t[0], t[1], results_queue, stop_flag, log_queue)
                    for t in tasks
                ])

            results_queue.put("STOP")
            watcher.join()

        log_queue.put("تم انهاء العملية")
        elapsed = (time.time() - start_time) / 60
        log_queue.put(f"وقت المعالجة {elapsed:.2f} دقيقة")
        stop_flag.set()
    except Exception as e:
        log_queue.put("خطأ اثناء اجراء العملية")
        log_queue.put(f"{e}")
        stop_flag.set()