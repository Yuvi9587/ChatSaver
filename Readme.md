# ChatGPT to Word  
**Export ChatGPT Shared Links to DOCX**

This project allows you to convert shared ChatGPT conversations into well-formatted Word (.docx) documents. It preserves:

- Headings  
- Paragraphs  
- Lists  
- Code blocks  
- Tables  

Use this tool to archive, print, or share ChatGPT conversations in a clean and professional format.

---

## Features

- Converts ChatGPT shared links to Word documents  
- Clean formatting using styles  
- Easy-to-use graphical interface  
- Fully open and readable source code  

---

## Project Files

| File Name         | Description                                      |
|------------------|--------------------------------------------------|
| chat_to_word.py   | Main script to fetch and convert conversations  |
| app_gui.py        | GUI built with Tkinter for easier interaction   |

Note: The executable (.exe) version is available via GitHub Releases. It is not included in the repository directly.

---

## How to Use

### Option 1: Use the Executable (No Python Required)

1. Go to the [Releases page](https://github.com/Yuvi9587/ChatSaver/releases)
2. Download the latest ChatSaver.exe  
3. Run the application  
4. Paste your ChatGPT shared link and click the button to export  

The executable was built from this exact source using PyInstaller. For added transparency, you can also use the Python version described below.

---

### Option 2: Run with Python

**Step 1: Clone the repository**

```
git clone https://github.com/Yuvi9587/ChatSaver
cd ChatGPT-to-Word
```

**Step 2: Install the required dependencies**  
Make sure you're using Python 3.8 or newer.

You can install dependencies using:

```
pip install -r requirements.txt
```

Or individually:

```
pip install selenium
pip install beautifulsoup4
pip install python-docx
```

Make sure you have Google Chrome and ChromeDriver installed.  
ChromeDriver should be in your system PATH or placed in the project directory.  
You can download it from: https://chromedriver.chromium.org/downloads

**Step 3: Launch the GUI**

```
python app_gui.py
```

---

## How It Works

**chat_to_word.py**

- Opens the ChatGPT shared link in a headless Chrome browser using Selenium  
- Extracts the entire conversation with BeautifulSoup  
- Parses headings, paragraphs, lists, code blocks, and tables  
- Generates a formatted Word document using python-docx  

**app_gui.py**

- Provides a simple graphical interface using Tkinter  
- Lets users export chats without using the command line  

---

## Compatibility and Testing

- Windows 10 and 11  
- Python versions 3.8 through 3.12  
- Chrome version 120 and above  
- Works with the latest ChatGPT shared link format  

---

## Planned Improvements

- Dark mode for the GUI  
- Support for embedded images  
- Option to export as PDF  
- Ability to save multiple conversations at once  

---

## License
This project does not currently use an open source license.  
---

## Credits

Developed using Selenium, BeautifulSoup, python-docx, and many hours of ChatGPT conversations.  
Contributions and feedback are welcome through GitHub issues or pull requests.

---

## Repository

GitHub: https://github.com/Yuvi9587/ChatSaver
---
