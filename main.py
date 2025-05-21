import tkinter as tk
from tkinter import messagebox
import subprocess
import os
import sqlite3

DB_PATH = "scripts.db"
LOG_DIR = "logs"

def run_powershell_script(script_content):
    log_file = os.path.join(LOG_DIR, "execution.log")
    try:
        completed = subprocess.run(
            ["powershell", "-Command", script_content],
            capture_output=True, text=True, check=True
        )
        with open(log_file, "a") as f:
            f.write(f"SUCCESS:\n{completed.stdout}\n")
        messagebox.showinfo("Success", "Script executed successfully!")
    except subprocess.CalledProcessError as e:
        with open(log_file, "a") as f:
            f.write(f"ERROR:\n{e.stderr}\n")
        messagebox.showerror("Error", f"Script failed:\n{e.stderr}")

def fetch_scripts():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT id, name, description FROM scripts")
    scripts = cursor.fetchall()
    conn.close()
    
    # Remove duplicates and sort scripts alphabetically by name
    unique_scripts = {script[1]: script for script in scripts}
    sorted_scripts = sorted(unique_scripts.values(), key=lambda x: x[1])
    
    return sorted_scripts

def get_script_content(script_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT script_content FROM scripts WHERE id=?", (script_id,))
    result = cursor.fetchone()
    conn.close()
    return result[0] if result else ""

def on_run():
    selected_idx = script_listbox.curselection()
    if not selected_idx:
        messagebox.showwarning("No Selection", "Please select a script.")
        return
    script_id = scripts[selected_idx[0]][0]
    script_content = get_script_content(script_id)
    run_powershell_script(script_content)

def on_select(event):
    selected_idx = script_listbox.curselection()
    if selected_idx:
        desc = scripts[selected_idx[0]][2]
        desc_var.set(desc)
    else:
        desc_var.set("")

def setup_gui():
    global script_listbox, scripts, desc_var
    root = tk.Tk()
    root.title("PowerShell-Scripts (O365/AD)")
    root.configure(bg="#222222")  # Dark background

    label_fg = "#FFFFFF"
    bg_color = "#222222"
    entry_bg = "#333333"
    button_bg = "#444444"
    button_fg = "#FFFFFF"
    listbox_bg = "#333333"
    listbox_fg = "#FFFFFF"
    scrollbar_bg = "#444444"

    tk.Label(root, text="Select a script:", bg=bg_color, fg=label_fg).pack(pady=5)

    frame = tk.Frame(root, bg=bg_color)
    frame.pack(padx=10, pady=5, fill="both", expand=True)

    script_listbox = tk.Listbox(
        frame, width=40, height=10, bg=listbox_bg, fg=listbox_fg, selectbackground="#555555", selectforeground="#FFFFFF"
    )
    script_listbox.pack(side="left", fill="both", expand=True)
    scrollbar = tk.Scrollbar(frame, orient="vertical", command=script_listbox.yview, bg=scrollbar_bg)
    scrollbar.pack(side="right", fill="y")
    script_listbox.config(yscrollcommand=scrollbar.set)

    scripts = fetch_scripts()
    for _, name, _ in scripts:
        script_listbox.insert("end", name)

    script_listbox.bind("<<ListboxSelect>>", on_select)

    desc_var = tk.StringVar()
    desc_label = tk.Label(root, textvariable=desc_var, wraplength=350, fg="#CCCCCC", bg=bg_color)
    desc_label.pack(pady=5)

    tk.Button(root, text="Run Script", command=on_run, bg=button_bg, fg=button_fg, activebackground="#666666", activeforeground="#FFFFFF").pack(pady=10)
    tk.Button(root, text="Exit", command=root.destroy, bg=button_bg, fg=button_fg, activebackground="#666666", activeforeground="#FFFFFF").pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    os.makedirs(LOG_DIR, exist_ok=True)
    setup_gui()