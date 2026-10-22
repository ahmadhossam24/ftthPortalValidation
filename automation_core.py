import requests
import json
import re
import base64
import os
import pandas as pd
import time
import sys
import threading
import random
import urllib3
from datetime import datetime
from openpyxl import load_workbook

KA_LOCK = threading.Lock()   # protects Authorization header updates / request races
urllib3.disable_warnings()

BASE    = "https://10.20.129.197/ftthPortalBE"
SEARCH  = BASE + "/dataSearchAndViewServiceNumberData/searchingWOAndVC4"
REFRESH = BASE + "/refresh"

PREFIXES = {"2", "3", "13", "40", "48", "55", "62", "64", "46", "66", "68", "69",
            "50", "45", "57", "47", "92", "65", "84", "82", "86", "88", "93", "95", "96", "97"}

KA_INTERVAL = 75          # keepalive refresh every 75s (backend re-issues ~every 2 min)

# ------------------------------------------------------------------ parse pasted cURL
def parse_curl(paste: str):
    """Extracts (bearer, cookie_dict) from a full 'Copy as cURL' paste, or a bare token."""
    paste = paste.strip()
    m = re.search(r"Bearer\s+([A-Za-z0-9\-_.]+)", paste)
    if m:
        bearer = m.group(1)
    elif re.fullmatch(r"[A-Za-z0-9\-_.]{80,}", paste) and paste.count(".") == 2:
        bearer = paste
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

# ------------------------------------------------------------------ token handling
def jwt_exp(token: str):
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload)).get("exp")
    except Exception:
        return None

def refresh_token(session, log_queue=None):
    """
    Call /refresh with the current session; extract new token from JSON key 'result'
    ("Bearer eyJ...") and update the session Authorization header IN PLACE.
    Raises PermissionError on failure.
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

# ------------------------------------------------------------------ excel writer (thread inside worker)
def excel_writer_loop(queue, log_queue, results_file, draft_file):
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

def init_results_file(original_file_path):
    dir_name = os.path.dirname(original_file_path)
    base_name = os.path.splitext(os.path.basename(original_file_path))[0]
    results_file = os.path.join(dir_name, f"{base_name}_results_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    draft_file   = os.path.join(dir_name, f"{base_name}_draft_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    pd.DataFrame(columns=["FULLNUMBER", "INFORMATION"]).to_excel(results_file, index=False)
    pd.DataFrame(columns=["FULLNUMBER", "INFORMATION"]).to_excel(draft_file, index=False)
    return results_file, draft_file

# ------------------------------------------------------------------ keepalive
def keepalive_loop(session, stop_event, log_queue):
    """Refresh the token every KA_INTERVAL seconds so the session never dies, even idle."""
    counter = 0
    while not stop_event.wait(KA_INTERVAL):
        try:
            with KA_LOCK:
                refresh_token(session)
            counter += 1
            if counter % 10 == 0:
                log_queue.put("الجلسة ما زالت نشطة (keepalive)")
        except Exception as e:
            log_queue.put(f"فشل keepalive: {e}")

# ------------------------------------------------------------------ process one row
def process_row(session, row_data, index, log_queue, batch_stop):
    full_number = str(row_data[0])
    full_number = full_number if full_number.startswith("0") else "0" + full_number
    landline_code, landline_number = split_number(str(row_data[0]))

    time.sleep(random.uniform(0.1, 0.9))

    r = None
    for attempt in range(2):
        if batch_stop.is_set():
            return None
        try:
            with KA_LOCK:
                r = session.get(SEARCH, params={"cityCode": landline_code,
                                                "serviceNumber": landline_number}, timeout=20)
        except requests.RequestException as e:
            return (index, [full_number, f"error: {e}"])

        if r.status_code in (401, 403) or "اسم المستخدم" in r.text:
            if attempt == 0:
                try:
                    with KA_LOCK:
                        refresh_token(session, log_queue)
                    continue                  # retry the SAME number with the new token
                except (PermissionError, requests.RequestException) as e:
                    log_queue.put(f"فشل تجديد الجلسة: {e}")
                    break
            break
        break

    if r is not None and (r.status_code in (401, 403) or "اسم المستخدم" in r.text):
        log_queue.put(f"انتهت الجلسة نهائياً عند السطر {index+1}")
        log_queue.put(f"برجاء بدء العملية مرة اخرى بداية من السطر {index+1}")
        return "SESSION_DEAD"

    if r.status_code != 200:
        return (index, [full_number, f"error HTTP {r.status_code}: {r.text[:100]}"])

    try:
        data = r.json()
    except ValueError:
        return (index, [full_number, f"error: غير JSON — {r.text[:100]}"])

    if data.get("code") == 1:
        return (index, [full_number, "Orange"])
    elif data.get("code") == 4:
        msg = data.get("errorMessage") or ""
        if "Not Accepted" in msg:
            return (index, [full_number, "Free"])
        elif "another ISP" in msg:
            return (index, [full_number, "Another Isp"])
        else:
            return (index, [full_number, f"code4: {msg}"])
    else:
        return (index, [full_number, f"unknown code {data.get('code')}"])

# ------------------------------------------------------------------ persistent worker process
def session_worker(cmd_queue, log_queue, batch_stop, kill_event):
    """
    Lives for the whole app lifetime. Holds the session + keepalive forever.
    Commands: ("RUN", curl_paste, excel_file, start_line) | ("QUIT",)
    batch_stop: aborts current batch but keeps session. kill_event: stops keepalive.
    """
    session = None                        # <- RESTORED

    while True:
        cmd = cmd_queue.get()
        if cmd is None or cmd[0] == "QUIT":
            break
        if cmd[0] == "RUN":
            _, curl_paste, excel_file, start_line = cmd
            try:
                if session is None:
                    bearer, cookies = parse_curl(curl_paste)
                    session = build_session(bearer, cookies)
                    with KA_LOCK:         # <- was session.ka_lock
                        refresh_token(session, log_queue)
                    threading.Thread(target=keepalive_loop,
                                     args=(session, kill_event, log_queue),
                                     daemon=True).start()
                    log_queue.put("تم تجهيز الجلسة — keepalive نشط")
                else:
                    # user may paste a fresh curl after a hard logout — update creds
                    # user pasted something new — try to apply it WITHOUT destroying the live session if it's invalid
                    old_auth = None
                    with KA_LOCK:
                        old_auth = session.headers.get("Authorization")
                    try:
                        bearer, cookies = parse_curl(curl_paste)
                        with KA_LOCK:
                            session.headers["Authorization"] = "Bearer " + bearer
                            for k, v in cookies.items():
                                session.cookies.set(k, v)
                            refresh_token(session, log_queue)   # must succeed
                        log_queue.put("تم تحديث بيانات الدخول بنجاح")
                    except Exception as e:
                        with KA_LOCK:
                            session.headers["Authorization"] = old_auth   # restore live token
                        log_queue.put(f"تحذير: تعذر تحديث بيانات الدخول ({e}) — تم الرجوع للجلسة الحالية")

                batch_stop.clear()
                results_file, draft_file = init_results_file(excel_file)
                try:
                    df = pd.read_excel(excel_file, header=None)
                except Exception as e:
                    log_queue.put(f"Could not read Excel file: {e}")
                    log_queue.put("تم انهاء العملية")
                    log_queue.put("__BATCH_DONE__")
                    continue

                start_index = int(start_line) - 1
                tasks = [(index, list(row)) for index, row in df.iloc[start_index:].iterrows()]
                if not tasks:
                    log_queue.put("لا توجد بيانات لمعالجتها")
                    log_queue.put("تم انهاء العملية")
                    log_queue.put("__BATCH_DONE__")
                    continue

                log_queue.put(f"بدء المعالجة من السطر {start_index+1} — عدد الصفوف {len(tasks)}")
                start_time = time.time()
                aborted = False

                import queue as pyqueue
                writer_q = pyqueue.Queue()
                wt = threading.Thread(target=excel_writer_loop,
                                      args=(writer_q, log_queue, results_file, draft_file),
                                      daemon=True)
                wt.start()

                session_dead = False
                for index, row_data in tasks:
                    if batch_stop.is_set():
                        log_queue.put("تم ايقاف العملية بواسطة المستخدم")
                        aborted = True
                        break
                    result = process_row(session, row_data, index, log_queue, batch_stop)
                    if result == "SESSION_DEAD":
                        session_dead = True
                        break
                    if result is not None:
                        writer_q.put(result)

                writer_q.put("STOP")
                wt.join(timeout=30)

                if not aborted and not session_dead:
                    log_queue.put("تم انهاء العملية")
                    elapsed = (time.time() - start_time) / 60
                    log_queue.put(f"وقت المعالجة {elapsed:.2f} دقيقة")
                log_queue.put("__BATCH_DONE__")     # UI uses this to re-enable inputs
                # session stays alive regardless — keepalive keeps running

            except Exception as e:
                log_queue.put("خطأ اثناء اجراء العملية")
                log_queue.put(f"{e}")
                log_queue.put("__BATCH_DONE__")

    # QUIT: stop keepalive and exit
    kill_event.set()
    log_queue.put("تم ايقاف الـ worker")