import tkinter as tk
from tkinter import filedialog, messagebox
from multiprocessing import freeze_support, Manager, Process
import base64
import json
import os
import sys

import automation_core

# ---- frozen-aware base dir (works as .py and as PyInstaller --onedir exe) ----
def get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)     # exe folder (--onedir)
    return os.path.dirname(os.path.abspath(__file__))

CREDENTIALS_FILE = os.path.join(get_base_dir(), "credentials.json")

def load_credentials():
    """Returns {'username': ..., 'password': ...} or {} if missing/empty/corrupt."""
    try:
        with open(CREDENTIALS_FILE, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and data.get("username") and data.get("password"):
            return {"username": data["username"], "password": data["password"]}
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        pass
    return {}          # first run / empty file / corrupt file -> silently empty

def save_credentials(username, password):
    try:
        with open(CREDENTIALS_FILE, "w", encoding="utf-8") as f:
            json.dump({"username": username, "password": password},
                      f, ensure_ascii=False, indent=2)
    except OSError as e:
        log_text_widget.insert(tk.END, f"تعذر حفظ بيانات الدخول: {e}\n")

if __name__ == '__main__':
    freeze_support()
    manager = Manager()
    log_queue = manager.Queue()
    cmd_queue = manager.Queue()
    batch_stop = manager.Event()
    kill_event = manager.Event()

    worker_process = Process(target=automation_core.session_worker,
                             args=(cmd_queue, log_queue, batch_stop, kill_event))
    worker_process.start()

    UPDATE_INTERVAL = 100
    root = tk.Tk()
    root.title("Automation Tool for httf portal Validation")
    saved = load_credentials()          # {} on first run — no error
    saved_username = saved.get("username", "")
    saved_password = saved.get("password", "")

    def update_log_from_queue():
        while not log_queue.empty():
            msg = log_queue.get()
            if msg == "__BATCH_DONE__":
                reenable_inputs()
            elif msg.startswith("__SHOW_LOGIN__|"):
                show_login_window(msg.split("|", 1)[1])
            elif msg == "__NEED_OTP__":
                show_otp_window()
                if pending_creds["username"]:
                    save_credentials(pending_creds["username"], pending_creds["password"])
            elif msg == "__LOGGED_IN__":
                close_login_windows()
                login_button.config(state=tk.DISABLED)
                start_button.config(state=tk.NORMAL)
                log_text_widget.insert(tk.END, "تم تسجيل الدخول بنجاح\n")
            elif msg.startswith("__LOGIN_FAILED__|"):
                close_login_windows()
                log_text_widget.insert(tk.END, "فشل تسجيل الدخول: " + msg.split("|", 1)[1] + "\n")
                login_button.config(state=tk.NORMAL)
            else:
                log_text_widget.insert(tk.END, msg + "\n")
                log_text_widget.see(tk.END)
        root.after(UPDATE_INTERVAL, update_log_from_queue)

    # ---------------- login windows ----------------
    login_win = {"win": None}
    pending_creds = {"username": "", "password": ""}

    def show_login_window(captcha_b64):
        win = tk.Toplevel(root)
        win.title("تسجيل الدخول")
        login_win["win"] = win  
        
        tk.Label(win, text="Email:").grid(row=0, column=0, padx=5, pady=5)
        user_e = tk.Entry(win, width=35)
        user_e.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(win, text="Password:").grid(row=1, column=0, padx=5, pady=5)
        pass_e = tk.Entry(win, width=35)
        pass_e.grid(row=1, column=1, padx=5, pady=5)

        user_e.insert(0, saved_username)
        pass_e.insert(0, saved_password)

        img = tk.PhotoImage(data=base64.b64decode(captcha_b64))
        tk.Label(win, image=img).grid(row=2, column=0, columnspan=2, pady=5)
        win.captcha_img = img        # keep reference

        tk.Label(win, text="Captcha:").grid(row=3, column=0, padx=5, pady=5)
        cap_e = tk.Entry(win, width=20)
        cap_e.grid(row=3, column=1, padx=5, pady=5)

        def submit():
            u, p, c = user_e.get().strip(), pass_e.get(), cap_e.get().strip()
            if not u or not p or not c:
                messagebox.showerror("خطأ", "برجاء ادخال جميع البيانات", parent=win)
                return
            pending_creds["username"] = user_e.get().strip()
            pending_creds["password"] = pass_e.get()
            cmd_queue.put(("LOGIN_SUBMIT", u, p, c))
            win.destroy()

        tk.Button(win, text="Login", command=submit).grid(row=4, column=0, columnspan=2, pady=8)

    def show_otp_window():
        win = tk.Toplevel(root)
        win.title("OTP")
        login_win["win"] = win
        tk.Label(win, text="تم ارسال رمز OTP — برجاء ادخاله:").pack(padx=10, pady=5)
        otp_e = tk.Entry(win, width=15)
        otp_e.pack(padx=10, pady=5)

        def submit():
            otp = otp_e.get().strip()
            if not otp:
                messagebox.showerror("خطأ", "برجاء ادخال الرمز", parent=win)
                return
            cmd_queue.put(("LOGIN_OTP", otp))
            win.destroy()

        tk.Button(win, text="Verify", command=submit).pack(pady=8)

    def close_login_windows():
        w = login_win.get("win")
        if w is not None and w.winfo_exists():
            w.destroy()
        login_win["win"] = None

    # ---------------- main controls ----------------
    def reenable_inputs():
        root.after(0, lambda: start_button.config(state=tk.NORMAL))
        root.after(0, lambda: file_entry.config(state=tk.NORMAL))
        root.after(0, lambda: browse_button.config(state=tk.NORMAL))
        root.after(0, lambda: start_line_entry.config(state=tk.NORMAL))
        root.after(0, lambda: stop_button.config(state=tk.DISABLED))

    def on_closing():
        if not messagebox.askyesno(
                "تأكيد",
                "هل انت متأكد؟ سوف يقوم بتسجيل الخروج من الموقع"):
            return  
        cmd_queue.put(("QUIT",))
        kill_event.set()
        batch_stop.set()
        worker_process.join(timeout=5)
        root.destroy()

    def browse_file():
        file_path = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)

    def start_login():
        login_button.config(state=tk.DISABLED)
        log_text_widget.insert(tk.END, "جاري تحضير تسجيل الدخول...\n")
        cmd_queue.put(("LOGIN",))

    def start_process():
        file_path = file_entry.get().strip()
        start_line = start_line_entry.get().strip()
        if not file_path:
            messagebox.showerror("Input Error", "Please select an Excel file.")
            return
        batch_stop.clear()
        cmd_queue.put(("RUN", file_path, start_line))
        start_button.config(state=tk.DISABLED)
        stop_button.config(state=tk.NORMAL)
        file_entry.config(state=tk.DISABLED)
        browse_button.config(state=tk.DISABLED)
        start_line_entry.config(state=tk.DISABLED)

    def stop_process():
        batch_stop.set()          # abort batch; session + keepalive stay alive

    tk.Label(root, text="Excel File:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
    file_entry = tk.Entry(root, width=60)
    file_entry.grid(row=0, column=1, padx=5, pady=5)
    browse_button = tk.Button(root, text="Browse", command=browse_file)
    browse_button.grid(row=0, column=2, padx=5, pady=5)

    tk.Label(root, text="start from line:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
    start_line_entry = tk.Entry(root, width=5)
    start_line_entry.insert(0, "1")
    start_line_entry.grid(row=1, column=1, padx=5, pady=5)

    login_button = tk.Button(root, text="1. Login", command=start_login, state=tk.NORMAL)
    login_button.grid(row=2, column=0, padx=10, pady=10)

    start_button = tk.Button(root, text="2. Start Processing",
                             command=start_process, state=tk.DISABLED)
    start_button.grid(row=2, column=1, padx=10, pady=10)

    stop_button = tk.Button(root, text="Stop Processing",
                            command=stop_process, state=tk.DISABLED)
    stop_button.grid(row=2, column=2, padx=10, pady=10)

    log_text_widget = tk.Text(root, height=15, width=90)
    log_text_widget.grid(row=3, column=0, columnspan=3, padx=5, pady=5)

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.after(UPDATE_INTERVAL, update_log_from_queue)
    root.mainloop()