import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import webbrowser
import os
import time
from chatgpt_to_word import fetch_chat_content, generate_docx

root = tk.Tk()
root.title("ChatSaver - Export Chat to Word")
root.geometry("900x700")
root.configure(bg="#1e1e1e")
root.attributes("-alpha", 0.0)
root.resizable(True, True)

def open_link(url):
    webbrowser.open_new_tab(url)

def browse_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_var.set(folder_selected)

def log_message(message):
    log_text.config(state='normal')
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)
    log_text.config(state='disabled')

def download_chat():
    thread = threading.Thread(target=run_download)
    thread.start()

spinner_symbols = ['|', '/', '-', '\\']

def log_spinner_message(base_message, duration=2.0, interval=0.2):
    end_time = time.time() + duration
    i = 0
    while time.time() < end_time:
        log_text.config(state='normal')
        log_text.insert(tk.END, f"{base_message} {spinner_symbols[i % len(spinner_symbols)]}\n")
        log_text.see(tk.END)
        log_text.update()
        time.sleep(interval)
        log_text.delete("end-2l", "end-1l")
        i += 1
    log_text.config(state='disabled')

def run_download():
    shared_link = link_entry.get().strip()
    output_folder = folder_var.get().strip()
    output_filename = filename_entry.get().strip()

    log_message("[INFO] Preparing download session...")

    if not shared_link or not output_folder or not output_filename:
        messagebox.showerror("Input Error", "Please complete all fields before downloading.")
        log_message("[ERROR] Missing required fields.")
        return

    if not output_filename.endswith(".docx"):
        output_filename += ".docx"

    output_path = os.path.join(output_folder, output_filename)

    try:
        log_message("[INFO] Establishing server connection...")
        log_spinner_message("[CONNECTING] Waiting for server", duration=2)

        log_message("[STEP] Verifying link...")
        time.sleep(0.5)

        log_message("[STEP] Checking server response...")
        log_spinner_message("[WAIT] Server handshake", duration=3)

        log_message(f"[STEP] Fetching chat from:\n{shared_link}")
        log_spinner_message("[FETCHING] Downloading chat data", duration=4)

        messages = fetch_chat_content(shared_link)

        if not messages:
            messagebox.showerror("Download Error", "No chat data found. Please check the link.")
            log_message("[ERROR] No messages extracted.")
            return

        log_message("[STEP] Parsing conversation...")
        time.sleep(0.5)

        log_message("[STEP] Building Word document...")
        log_spinner_message("[BUILDING] Assembling Word file", duration=2)

        generate_docx(messages, output_path)

        log_message(f"[SUCCESS] Chat saved successfully!")
        log_message(f"[FILE] {output_path}")

        messagebox.showinfo("Success", f"File saved to:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Unexpected Error", f"An error occurred:\n{str(e)}")
        log_message(f"[ERROR] {str(e)}")

def on_enter(event):
    event.widget.config(bg="#1abc9c")

def on_leave(event):
    event.widget.config(bg="#16a085")

def create_button(master, text, command):
    button = tk.Button(master, text=text, command=command, bg="#16a085", fg="white",
                       activebackground="#1abc9c", relief="flat", font=("Segoe UI", 10))
    button.bind("<Enter>", on_enter)
    button.bind("<Leave>", on_leave)
    return button

def fade_in():
    alpha = root.attributes("-alpha")
    if alpha < 1.0:
        alpha += 0.05
        root.attributes("-alpha", alpha)
        root.after(30, fade_in)
fade_in()

header = tk.Frame(root, bg="#0f111a", height=60)
header.pack(fill="x")
tk.Label(header, text="ChatSaver", font=("Segoe UI", 20, "bold"), fg="white", bg="#0f111a").pack(padx=10, pady=10, anchor="w")

content = tk.Frame(root, bg="#1e1e1e")
content.pack(padx=25, pady=20, fill="both", expand=True)

tk.Label(content, text="🔗 Chat Link:", font=("Segoe UI", 11, "bold"), fg="white", bg="#1e1e1e").pack(anchor="w", pady=(5, 2))
link_entry = tk.Entry(content, bg="#2c2c2c", fg="white", insertbackground="white", relief="flat")
link_entry.pack(fill="x", pady=(0, 10), expand=True)

tk.Label(content, text="📁 Output Folder:", font=("Segoe UI", 11, "bold"), fg="white", bg="#1e1e1e").pack(anchor="w", pady=(10, 2))
folder_frame = tk.Frame(content, bg="#1e1e1e")
folder_frame.pack(fill='x', pady=(0, 10))
folder_var = tk.StringVar()
folder_entry = tk.Entry(folder_frame, textvariable=folder_var, bg="#2c2c2c", fg="white", insertbackground="white", relief="flat")
folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
browse_button = create_button(folder_frame, "Browse", browse_folder)
browse_button.pack(side="right")

tk.Label(content, text="📝 File Name (.docx):", font=("Segoe UI", 11, "bold"), fg="white", bg="#1e1e1e").pack(anchor="w", pady=(10, 2))
filename_entry = tk.Entry(content, bg="#2c2c2c", fg="white", insertbackground="white", relief="flat")
filename_entry.pack(fill="x", pady=(0, 10), expand=True)

download_button = create_button(content, "⬇️ Download Chat", download_chat)
download_button.config(font=("Segoe UI", 12, "bold"))
download_button.pack(pady=10)

tk.Label(content, text="🧾 Process Log:", font=("Segoe UI", 11, "bold"), fg="white", bg="#1e1e1e").pack(anchor="w", pady=(20, 2))
log_text = scrolledtext.ScrolledText(content, height=10, state='disabled', bg="#121212", fg="white", insertbackground="white", relief="flat")
log_text.pack(fill='both', expand=True, pady=(0, 10))

social_frame = tk.Frame(root, bg="#1e1e1e")
social_frame.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)  

def open_instagram():
    open_link("https://www.instagram.com/_yuvraj_panwar/") 

def open_github():
    open_link("https://github.com/Yuvi9587") 

insta_button = create_button(social_frame, "📷 Instagram", open_instagram)
insta_button.pack(side="left", padx=(0, 10))

github_button = create_button(social_frame, "🐙 GitHub", open_github)
github_button.pack(side="left")

root.mainloop()
