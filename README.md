# 📊 Text Counter Pro: Python Excel & CSV Analyzer

**PythonExcelWordCharCounter** is a high-performance, Notion-inspired desktop application designed to accurately count words and characters within Excel (`.xlsx`) and CSV files. Whether you are auditing a single sheet or batch-processing hundreds of files, this tool provides precise analytics with professional-grade features like multiprocessing and advanced filtering.

---

## ✨ Features

- **🚀 Hyper-Fast Processing:** Utilizes Multiprocessing (Multicore) to analyze large batches of files simultaneously.
- **🖱️ Drag & Drop:** Simply drop your `.xlsx` or `.csv` files directly onto the app to start analyzing.
- **📑 Full Workbook Reports:** Generate a comprehensive breakdown of every sheet within a workbook with one click.
- **🔍 Advanced Filtering:** Ignore specific words (e.g., stop words) or characters (e.g., punctuation) from your counts.
- **🙈 Visibility Aware:** Optional setting to skip hidden rows and columns for 100% visual accuracy.
- **🌓 Dark Mode:** Sleek, Notion-inspired user interface with full support for Light and Dark themes.
- **📥 Professional Export:** Save your results to clean `.txt` or structured `.csv` files for further reporting.

---

## 🛠️ Tech Stack

- **Python 3.10+**
- **openpyxl:** High-performance Excel engine.
- **customtkinter:** Modern, responsive UI framework.
- **tkinterdnd2:** Native OS drag-and-drop integration.
- **multiprocessing:** Parallel computing for massive datasets.

---

## 🚀 Installation

Follow these steps to get the project running on your local machine:

1. **Clone the repository:**
   ```bash
   git clone https://github.com/your-username/PythonExcelWordCharCounter.git
   cd PythonExcelWordCharCounter
   ```

2. **Set up a Virtual Environment (Recommended):**
   ```bash
   python -m venv venv
   # Windows:
   .\venv\Scripts\activate
   # macOS/Linux:
   source venv/bin/activate
   ```

3. **Install Dependencies:**
   ```bash
   pip install customtkinter openpyxl tkinterdnd2
   ```

---

## 📂 Usage

To launch the application, run the following command in your terminal:

```bash
python excel_counter.py
```

### Quick Start:
1. **Select Files:** Click "Select File" or drag your files directly into the window.
2. **Choose Mode:** Select "Words", "Chars", or "Both".
3. **Auto-Detect:** Click "Auto-detect range" to instantly find the boundaries of your data.
4. **Start:** Click "Start" for a quick count or "Full Report" for a sheet-by-sheet breakdown.

---

## ⚙️ Configuration

The application uses a `settings.json` file to store your preferences automatically. You don't need to edit this file manually, as the app updates it whenever you change settings in the UI.

**What is stored:**
- `appearance_mode`: Your choice of Light or Dark mode.
- `last_file_path`: The directory of the last file you accessed.
- `ignore_words`: Your custom list of words to exclude.
- `skip_hidden`: Whether to ignore hidden Excel cells.
- `skip_parallel`: Your preference for Multicore processing.

---

## 📜 License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.

---

**Developed with ❤️ for high-precision text auditing.**
