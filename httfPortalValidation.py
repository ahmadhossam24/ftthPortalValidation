from multiprocessing import freeze_support,Manager, Process
import multiprocessing
multiprocessing.freeze_support()
import tkinter as tk
from tkinter import filedialog, messagebox
from multiprocessing import Process 
from automation_core import run_bot

#############################################
# Tkinter GUI Setup
#############################################

if __name__ == '__main__':
    freeze_support()
    manager = Manager()
    current_process=None
    UPDATE_INTERVAL = 100  # ms
    PROCESS_CHECK_INTERVAL = 200  # ms for checking if process finished
    log_queue = manager.Queue() 
    stop_flag = manager.Event()

    def update_log_from_queue():
        while not log_queue.empty():
            msg = log_queue.get()
            log_text_widget.insert(tk.END, msg + "\n")
            log_text_widget.see(tk.END)
            
        root.after(UPDATE_INTERVAL, update_log_from_queue)

    def reenable_new_process():
        # Re-enable inputs after the process completes
        root.after(0, lambda: start_button.config(state=tk.NORMAL))
        root.after(0, lambda: auth_entry.config(state=tk.NORMAL))
        root.after(0, lambda: file_entry.config(state=tk.NORMAL))
        root.after(0, lambda: browse_button.config(state=tk.NORMAL))
        root.after(0, lambda: start_line_entry.config(state=tk.NORMAL))
        root.after(0, lambda: stop_button.config(state=tk.DISABLED))

    def on_closing():
        stop_flag.set()
        root.destroy()
            
    def browse_file():
        file_path = filedialog.askopenfilename(
            title="Select an Excel file",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not file_path:
            return
        try:
            file_entry.delete(0, tk.END)
            file_entry.insert(0, file_path)
        except Exception as e:
            messagebox.showerror("خطأ", f"تعذر قراءة الملف: {e}")
            file_entry.delete(0, tk.END)

    def start_process():
        global log_queue,stop_flag, current_process

        auth = auth_entry.get().strip()
        file_path = file_entry.get().strip()
        start_line=start_line_entry.get().strip()
        if not auth or not file_path:
            messagebox.showerror("Input Error", "Please enter authorization, and select an Excel file.")
            return
        start_button.config(state=tk.DISABLED)
        stop_button.config(state=tk.NORMAL)
        auth_entry.config(state=tk.DISABLED)
        file_entry.config(state=tk.DISABLED)
        start_line_entry.config(state=tk.DISABLED)
        browse_button.config(state=tk.DISABLED)
        stop_flag.clear()
        # print(auth,file_path,start_line)
        current_process = Process(target=run_bot, args=(auth,start_line,file_path,log_queue, stop_flag))
        current_process.start()
        # # Start monitoring the process # #
        check_process_alive()

    def check_process_alive():
        """Periodically check if the bot process has finished."""
        global current_process
        if current_process is not None:
            if not current_process.is_alive():
                # Process finished (normally or due to stop)
                reenable_new_process()
                current_process = None
                # Optionally show a message
                messagebox.showinfo("Completed", "The operation has finished.")
            else:
                # Still running – check again later
                root.after(PROCESS_CHECK_INTERVAL, check_process_alive)


    def stop_process():
        stop_flag.set()
        reenable_new_process()

    def show_context_menu(event, entry):
        menu = tk.Menu(entry, tearoff=0)
        menu.add_command(label="Cut", command=lambda: entry.event_generate("<<Cut>>"))
        menu.add_command(label="Copy", command=lambda: entry.event_generate("<<Copy>>"))
        menu.add_command(label="Paste", command=lambda: entry.event_generate("<<Paste>>"))
        menu.tk_popup(event.x_root, event.y_root)

    # Create main Tkinter window
    root = tk.Tk()
    root.title("Automation Tool for httf portal Validation")

    instructions = (
        # "في حالة وجود رقمين هاتف يرجى وضع البيانات كالاتي\n"
        # "العنوان <- رقم الهاتف الاضافي <- رقم هاتف التقديم <- الاسم <- رقم الارضي كاملا\n\n"
        # "في حالة وجود رقم هاتف واحد فقط يرجى وضع البيانات كالاتي\n"
        # "العنوان <- رقم الهاتف <- الاسم <- رقم الارضي كاملا"
    )

    tk.Label(root, text=instructions, justify="right", anchor="w").grid(
        row=0, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5
    )

    tk.Label(root, text="curl(bash) paste:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
    auth_entry = tk.Entry(root, width=30)
    auth_entry.grid(row=1, column=1, padx=5, pady=5)

    auth_entry.bind("<Button-3>", lambda event: show_context_menu(event, auth_entry))

    tk.Label(root, text="Excel File:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
    file_entry = tk.Entry(root, width=30)
    file_entry.grid(row=2, column=1, padx=5, pady=5)
    browse_button = tk.Button(root, text="Browse", command=browse_file)
    browse_button.grid(row=2, column=2, padx=5, pady=5)

    tk.Label(root, text="start from line:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
    start_line_entry = tk.Entry(root, width=5)
    start_line_entry.insert(0, "1")   # default value
    start_line_entry.grid(row=3, column=1, padx=5, pady=5)

    start_button = tk.Button(root, text="Start Processing", command=start_process)
    start_button.grid(row=4, column=0, padx=10, pady=10)

    stop_button = tk.Button(root, text="Stop Processing", command=stop_process, state=tk.DISABLED)
    stop_button.grid(row=4, column=1, padx=10, pady=10)

    log_text_widget = tk.Text(root, height=15, width=70)
    log_text_widget.grid(row=5, column=0, columnspan=3, padx=5, pady=5)

    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    root.after(UPDATE_INTERVAL, update_log_from_queue)  # Start updater
    root.mainloop()
