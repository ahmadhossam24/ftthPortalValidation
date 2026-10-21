import tkinter as tk
from tkinter import filedialog, messagebox
from multiprocessing import freeze_support, Manager, Process
import multiprocessing

import automation_core

if __name__ == '__main__':
    freeze_support()
    manager = Manager()
    log_queue = manager.Queue()
    cmd_queue = manager.Queue()        # UI -> worker commands
    batch_stop = manager.Event()       # abort current batch (session survives)
    kill_event = manager.Event()       # stop keepalive / quit worker

    # ---- persistent worker: starts with the app, keeps the session alive ----
    worker_process = Process(target=automation_core.session_worker,
                             args=(cmd_queue, log_queue, batch_stop, kill_event))
    worker_process.start()

    UPDATE_INTERVAL = 100
    PROCESS_CHECK_INTERVAL = 200
    current_run = {"running": False}

    root = tk.Tk()
    root.title("Automation Tool for httf portal Validation")

    # def update_log_from_queue():
    #     while not log_queue.empty():
    #         msg = log_queue.get()
    #         log_text_widget.insert(tk.END, msg + "\n")
    #         log_text_widget.see(tk.END)
    #     root.after(UPDATE_INTERVAL, update_log_from_queue)
    def update_log_from_queue():
        while not log_queue.empty():
            msg = log_queue.get()
            if msg == "__BATCH_DONE__":
                reenable_inputs()
                continue
            log_text_widget.insert(tk.END, msg + "\n")
            log_text_widget.see(tk.END)
        root.after(UPDATE_INTERVAL, update_log_from_queue)

    def reenable_inputs():
        root.after(0, lambda: start_button.config(state=tk.NORMAL))
        root.after(0, lambda: auth_entry.config(state=tk.NORMAL))
        root.after(0, lambda: file_entry.config(state=tk.NORMAL))
        root.after(0, lambda: browse_button.config(state=tk.NORMAL))
        root.after(0, lambda: start_line_entry.config(state=tk.NORMAL))
        root.after(0, lambda: stop_button.config(state=tk.DISABLED))
        current_run["running"] = False

    def set_running_state():
        start_button.config(state=tk.DISABLED)
        stop_button.config(state=tk.NORMAL)
        auth_entry.config(state=tk.DISABLED)
        file_entry.config(state=tk.DISABLED)
        start_line_entry.config(state=tk.DISABLED)
        browse_button.config(state=tk.DISABLED)
        current_run["running"] = True
        check_batch_finished()

    def on_closing():
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
        try:
            file_entry.delete(0, tk.END)
            file_entry.insert(0, file_path)
        except Exception as e:
            messagebox.showerror("خطأ", f"تعذر قراءة الملف: {e}")
            file_entry.delete(0, tk.END)

    def start_process():
        auth = auth_entry.get().strip()
        file_path = file_entry.get().strip()
        start_line = start_line_entry.get().strip()
        if not auth or not file_path:
            messagebox.showerror("Input Error",
                                 "Please paste the curl, and select an Excel file.")
            return
        batch_stop.clear()
        # send command to the persistent worker; session is kept/reused inside it
        cmd_queue.put(("RUN", auth, file_path, start_line))
        set_running_state()

    def check_batch_finished():
        """The worker logs 'تم انهاء العملية' / 'تم ايقاف العملية' when the batch ends;
        we re-enable inputs when the worker stops sending progress for the batch.
        Simplest reliable signal: monitor a dedicated sentinel message."""
        if not current_run["running"]:
            return
        # finished detection is driven by messages the worker pushes:
        # 'تم انهاء العملية' or 'تم ايقاف العملية بواسطة المستخدم' or 'انتهت الجلسة'
        root.after(PROCESS_CHECK_INTERVAL, check_batch_finished)

    def stop_process():
        batch_stop.set()          # abort batch; session + keepalive stay alive

    def show_context_menu(event, entry):
        menu = tk.Menu(entry, tearoff=0)
        menu.add_command(label="Cut", command=lambda: entry.event_generate("<<Cut>>"))
        menu.add_command(label="Copy", command=lambda: entry.event_generate("<<Copy>>"))
        menu.add_command(label="Paste", command=lambda: entry.event_generate("<<Paste>>"))
        menu.tk_popup(event.x_root, event.y_root)

    tk.Label(root, text="curl(bash) paste:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
    auth_entry = tk.Entry(root, width=60)
    auth_entry.grid(row=1, column=1, padx=5, pady=5)
    auth_entry.bind("<Button-3>", lambda event: show_context_menu(event, auth_entry))

    tk.Label(root, text="Excel File:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
    file_entry = tk.Entry(root, width=60)
    file_entry.grid(row=2, column=1, padx=5, pady=5)
    browse_button = tk.Button(root, text="Browse", command=browse_file)
    browse_button.grid(row=2, column=2, padx=5, pady=5)

    tk.Label(root, text="start from line:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
    start_line_entry = tk.Entry(root, width=5)
    start_line_entry.insert(0, "1")
    start_line_entry.grid(row=3, column=1, padx=5, pady=5)

    start_button = tk.Button(root, text="Start Processing", command=start_process)
    start_button.grid(row=4, column=0, padx=10, pady=10)

    stop_button = tk.Button(root, text="Stop Processing", command=stop_process, state=tk.DISABLED)
    stop_button.grid(row=4, column=1, padx=10, pady=10)

    log_text_widget = tk.Text(root, height=15, width=90)
    log_text_widget.grid(row=5, column=0, columnspan=3, padx=5, pady=5)

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.after(UPDATE_INTERVAL, update_log_from_queue)
    root.mainloop()