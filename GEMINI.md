# ExcelSheetCharWordCounter

A professionally-polished Python desktop application for counting words and characters within specific ranges of Excel (`.xlsx`) files. Designed with a modern, Notion-inspired aesthetic and robust processing capabilities.

## Project Overview

*   **Purpose:** Provides a user-friendly interface to analyze text volume in Excel sheets, supporting both single-file and batch processing.
*   **Key Technologies:**
    *   `customtkinter`: For the modern, responsive UI.
    *   `tkinterdnd2`: For native drag-and-drop support (arch-aware for 64-bit systems).
    *   `openpyxl`: For memory-efficient Excel reading (`read_only=True`, `data_only=True`).
    *   `threading` & `queue`: To ensure the UI remains responsive during heavy processing tasks.
    *   `logging`: Comprehensive error tracking and audit logs in `app_debug.log`.

## Architecture & Logic

The application follows a responsive UI pattern:
1.  **UI Thread:** Handles all user interactions and updates the display.
2.  **Worker Thread:** Processes Excel files in the background to avoid "Not Responding" states.
3.  **Message Queue:** Safely communicates progress and results from the worker thread back to the UI.
4.  **Persistence:** Remembers user preferences (last file, mode, range) via `settings.json`.

## Building and Running

### Prerequisites
*   Python 3.10+
*   Dependencies:
    ```bash
    pip install customtkinter openpyxl tkinterdnd2
    ```

### Running the App
```bash
python excel_counter.py
```

### Packaging (TODO)
*   Standard practice would be to use `PyInstaller` or `Nuitka` to create a standalone executable.
    ```bash
    # Placeholder command
    pyinstaller --noconsole --onefile excel_counter.py
    ```

## Development Conventions

*   **Error Handling:** All file-level and processing-level operations are wrapped in `try-except` blocks with detailed logging.
*   **UI Style:** Adheres to a strict "Notion-inspired" theme using pure whites, deep charcoals, and specific gray borders.
*   **Performance:** Uses `openpyxl`'s `iter_rows` and `read_only` mode to minimize RAM usage on large files.
*   **Validation:** All user inputs (especially Excel Ranges) are validated via Regex before processing starts.
