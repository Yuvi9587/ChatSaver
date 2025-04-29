ChatGPT to Word - Export ChatGPT Shared Links to DOCX
Welcome! 👋 This project helps you convert ChatGPT shared conversations into a clean and readable Word (.docx) document, preserving:

✅ Headings

✅ Paragraphs

✅ Lists

✅ Code blocks

✅ Tables

Use it to archive, print, or share ChatGPT conversations in a professional format!

📁 Project Files

File Name	             Description
chat_to_word.py	         Core script that fetches the chat and builds the DOCX file
app_gui.py	             GUI (Graphical User Interface) to make using the script easier
ChatGPTtoWord.exe	     Compiled EXE version – for quick use without installing Python

💻 How to Use
👉 Option 1: Use the EXE (No Python Needed)
Download ChatGPTtoWord.exe

Double-click and run it 🖱️

Paste your ChatGPT shared link, click a button, and done! 🎉

🛡️ Is the EXE safe?
Yes! The .exe was generated using pyinstaller from this exact source code.
Still, if you're unsure or don't trust random executables (totally fair!), use Option 2 below.

👉 Option 2: Use the Python Scripts (For Full Transparency)
1️⃣ Clone the repo
bash
Copy
Edit
git clone https://github.com/Yuvi9587/ChatSaver
cd ChatGPT-to-Word
2️⃣ Install dependencies
Make sure you're using Python 3.8 or above.

You can install the required libraries with pip:

pip install -r requirements.txt
Or manually install them:

pip install selenium
pip install beautifulsoup4
pip install python-docx
🧩 You'll also need Google Chrome and ChromeDriver for Selenium.
Make sure chromedriver is in your system PATH or in the project directory.
Download: https://chromedriver.chromium.org/downloads

3️⃣ Run the GUI app

python app_gui.py

🛠 How It Works
chat_to_word.py:
Uses Selenium to open the ChatGPT shared link in headless Chrome

Extracts the entire conversation using BeautifulSoup

Parses HTML elements: headings, paragraphs, lists, tables, code

Creates a nicely formatted .docx file with python-docx

app_gui.py:
Provides a simple GUI built with tkinter

You just paste the link and click a button 💡

No command-line needed!

🧪 Tested On
Windows 10/11

Python 3.8–3.12

Chrome v120+

Works with newer ChatGPT shared link formats

📎 To-Do / Improvements
 Dark mode GUI 🌙

 Support for embedded images

 PDF export option

 Save multiple conversations at once

📌 GitHub Repo
[👉[ GITHUB LINK HERE](https://github.com/Yuvi9587) 👈]
 
 
 
🙌 Credits
Made with ☕, selenium, and lots of ChatGPT chats.
Feel free to open issues or contribute with PRs!