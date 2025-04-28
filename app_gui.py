import tkinter as tk
from tkinter import filedialog, messagebox
import webbrowser
import os
from chatgpt_to_word import fetch_chat_content, generate_docx  # <-- Import your functions

def open_link(url):
    webbrowser.open_new(url)

def browse_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_var.set(folder_selected)

def download_chat():
    shared_link = link_entry.get().strip()
    output_folder = folder_var.get().strip()
    output_filename = filename_entry.get().strip()

    if not shared_link or not output_folder or not output_filename:
        messagebox.showerror("Error", "Please fill all fields.")
        return

    if not output_filename.endswith(".docx"):
        output_filename += ".docx"

    output_path = os.path.join(output_folder, output_filename)

    try:
        messages = fetch_chat_content(shared_link)
        if not messages:
            messagebox.showerror("Error", "No messages found or failed to fetch the conversation.")
            return

        generate_docx(messages, output_path)
        messagebox.showinfo("Success", f"Word document saved to:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong:\n{str(e)}")

# ---------------- GUI Setup ----------------

root = tk.Tk()
root.title("ChatGPT to Word Converter")
root.geometry("500x400")
root.resizable(False, False)

title_label = tk.Label(root, text="ChatGPT to Word Downloader", font=("Arial", 16, "bold"))
title_label.pack(pady=10)

# Link Entry
tk.Label(root, text="Paste Chat Link:").pack(anchor='w', padx=20)
link_entry = tk.Entry(root, width=50)
link_entry.pack(padx=20, pady=5)

# Folder Selection
tk.Label(root, text="Select Download Folder:").pack(anchor='w', padx=20)
folder_frame = tk.Frame(root)
folder_frame.pack(padx=20, pady=5, fill="x")
folder_var = tk.StringVar()
folder_entry = tk.Entry(folder_frame, textvariable=folder_var, width=38)
folder_entry.pack(side="left", fill="x", expand=True)
browse_button = tk.Button(folder_frame, text="Browse", command=browse_folder)
browse_button.pack(side="right")

# File Name
tk.Label(root, text="Enter File Name (without .docx):").pack(anchor='w', padx=20)
filename_entry = tk.Entry(root, width=50)
filename_entry.pack(padx=20, pady=5)

# Download Button
download_button = tk.Button(root, text="Download Chat", command=download_chat, bg="#4CAF50", fg="white", font=("Arial", 12))
download_button.pack(pady=20)

# Social Media Links
social_frame = tk.Frame(root)
social_frame.pack(pady=10)

tk.Label(social_frame, text="Follow me: ").pack(side="left")
tk.Button(social_frame, text="GitHub", command=lambda: open_link("https://github.com/Yuvi9587")).pack(side="left", padx=5)
tk.Button(social_frame, text="Instagram", command=lambda: open_link("https://www.instagram.com/_yuvraj_panwar")).pack(side="left", padx=5)
tk.Button(social_frame, text="Gmail", command=lambda: open_link("https://mail.google.com/mail/u/8/#sent?compose=jrjtXMncShVfPFmBdSKZNVPLxJbXSQXZfNLbhcQMDmHmWjBMlJMgzNMvQMldtvCcBPGkvSGf")).pack(side="left", padx=5)

root.mainloop()
