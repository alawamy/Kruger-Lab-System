# -*- coding: utf-8 -*-
"""
Daily PDF Report Generator - Excel to PDF Conversion Module

This module provides automated conversion of daily laboratory Excel reports to PDF format
with intelligent worksheet selection based on report content and data availability.

The module implements a sophisticated PDF generation system that:
    1. Dynamically selects worksheets based on data presence in the main report
    2. Handles Excel COM automation with robust error recovery
    3. Creates organized PDF archives mirroring the Excel file structure
    4. Manages concurrent file access and permission conflicts
    5. Provides user feedback through interactive directory selection

PDF Generation Logic:
    - Always includes 'Main' worksheet (core report data)
    - Conditionally includes specialized worksheets based on data validation:
        * Sulfide → if L16 has value
        * COD → if J16 has value
        * TS & VS → if G16 has value
        * TSS → if H16 has value
        * BOD5 → if D34 has value
        * O&G → if M16 has value
        * Efficiency → if E25 has value
        * State Points → if E26 has value
        * TDS → if TDS worksheet F26 has value

File Organization:
    Daily Reports/
    ├── Daily Sheets/
    │   ├── 2026/
    │   │   ├── January/
    │   │   │   ├── 01-01-2026.xlsx
    │   │   │   ├── 02-01-2026.xlsx
    │   │   │   └── ...
    │   └── PDFs/ (auto-generated)
    │       ├── 2026/
    │       │   ├── January/
    │       │   │   ├── 01-01-2026.pdf
    │       │   │   ├── 02-01-2026.pdf
    │       │   │   └── ...

Excel COM Automation:
    - Uses win32com.client for Excel application control
    - Implements retry logic for file access conflicts
    - Handles OLE automation errors gracefully
    - Ensures proper cleanup of COM objects

Error Handling:
    - File permission conflicts (PDF already open)
    - Excel application busy states (OLE 0x800ac472)
    - Missing worksheets or invalid cell references
    - COM object initialization failures
    - Directory access and creation issues

Dependencies:
    - pythoncom: COM initialization and cleanup
    - win32com.client: Excel automation interface
    - tkinter.filedialog: Directory selection dialog
    - config: DAILY_PDF, DAILY_SHEET path constants

Author: Kruger Lab System Development Team
Version: 2.0
Last Updated: 2026-04-06
"""

import os
import time
import pythoncom
import win32com.client
import tkinter as tk
from tkinter import filedialog, messagebox
from config import DAILY_PDF, DAILY_SHEET


def daily_pdf_report(root):
    """
    Convert daily Excel reports to PDF format with intelligent worksheet selection.

    This function orchestrates the complete PDF generation workflow for daily laboratory
    reports. It provides an interactive directory selection interface and automatically
    processes all Excel files in the selected directory, converting them to PDF format
    with conditional worksheet inclusion based on data availability.

    Process Flow:
        1. Initialize COM for Excel automation
        2. Prompt user to select source directory (daily reports folder)
        3. Create corresponding PDF output directory structure
        4. Initialize Excel application with error handling
        5. Process each Excel file:
           - Analyze Main worksheet for data presence
           - Select relevant worksheets for PDF inclusion
           - Create temporary workbook with selected sheets
           - Export to PDF format
        6. Clean up Excel application and COM resources

    Directory Structure:
        Input:  DAILY_SHEET/year/month/ (Excel files)
        Output: DAILY_PDF/year/month/ (PDF files)

    Worksheet Selection Criteria:
        Core worksheets (always included):
            - Main: Primary report data and metadata

        Conditional worksheets (included if data exists):
            - Sulfide: L16 cell has value
            - COD: J16 cell has value
            - TS & VS: G16 cell has value
            - TSS: H16 cell has value
            - BOD5: D34 cell has value
            - O&G: M16 cell has value
            - Efficiency: E25 cell has value
            - State Points: E26 cell has value
            - TDS: TDS worksheet F26 has value

    Args:
        root (tk.Tk): Parent tkinter window for dialog ownership.
                     Used to ensure proper dialog parenting and focus management.

    Returns:
        None: Function displays success/error messages to user and returns void.

    Raises:
        None: All exceptions are caught internally and displayed to user via messagebox.

    Excel Automation Details:
        - Uses win32com.client.Dispatch for Excel application control
        - Sets Visible=False and DisplayAlerts=False for background processing
        - Implements retry logic (3 attempts) for file access conflicts
        - Handles OLE automation errors (0x800ac472) with delays
        - Properly closes workbooks and quits Excel application

    Error Recovery:
        - File locked/open: Skip with continue (permission errors)
        - Excel busy: Retry with 1-second delays up to 3 attempts
        - Missing worksheets: Gracefully skip invalid sheet references
        - COM errors: Display user-friendly error messages
        - Directory issues: Create paths with exist_ok=True

    Performance Considerations:
        - Processes files sequentially to avoid resource conflicts
        - Minimal memory footprint through immediate workbook cleanup
        - Optimized worksheet copying using Before parameter
        - Efficient PDF export using ExportAsFixedFormat

    User Experience:
        - Interactive directory selection with sensible defaults
        - Progress indication through processed file count
        - Clear success/failure messaging in Arabic
        - Graceful cancellation handling (empty directory selection)

    Examples:
        >>> # Called from main application
        >>> root = tk.Tk()
        >>> daily_pdf_report(root)  # Shows directory dialog, processes files

        >>> # Standalone execution
        >>> if __name__ == "__main__":
        ...     root = tk.Tk()
        ...     root.withdraw()
        ...     daily_pdf_report(root)

    Integration Points:
        - app.py: Called from main dashboard PDF generation button
        - main_menu.py: Accessible through advanced operations menu
        - file_checker_module.py: Processes reports created by ReportManager

    Dependencies:
        - pythoncom.CoInitialize/CoUninitialize: COM apartment threading
        - win32com.client.Dispatch: Excel application automation
        - tkinter.filedialog.askdirectory: Directory selection UI
        - config paths: DAILY_SHEET (input), DAILY_PDF (output)
    """
    # Initialize COM for Excel automation (required for win32com)
    pythoncom.CoInitialize()

    # Interactive directory selection for source Excel files
    source_dir = filedialog.askdirectory(initialdir=DAILY_SHEET)
    if not source_dir:
        # User cancelled directory selection
        pythoncom.CoUninitialize()
        return

    # Extract folder names for dynamic PDF path construction
    month_folder = os.path.basename(source_dir)
    year_folder = os.path.basename(os.path.dirname(source_dir))

    # Create PDF output directory mirroring Excel structure
    pdf_root = DAILY_PDF / year_folder / month_folder
    pdf_root.mkdir(parents=True, exist_ok=True)

    # Initialize Excel application and processing counters
    excel = None
    processed_count = 0

    try:
        # Initialize Excel application with proper error handling
        # Note: Using regular Dispatch instead of dynamic for property recognition
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Run in background
        excel.DisplayAlerts = False  # Suppress alerts for automation

        # Get list of Excel files to process (exclude temp files starting with ~$)
        excel_files = [f for f in os.listdir(source_dir)
                       if f.endswith('.xlsx') and not f.startswith('~$')]

        # Process each Excel file in the selected directory
        for excel_file in excel_files:
            # Construct corresponding PDF output path
            pdf_path = pdf_root / excel_file.replace('.xlsx', '.pdf')

            # Handle existing PDF files (remove to avoid conflicts)
            if pdf_path.exists():
                try:
                    os.remove(pdf_path)
                except PermissionError:
                    # Skip if PDF is open in another application
                    continue

            # Attempt to open Excel file with retry logic for busy states
            wb_source = None
            for attempt in range(3):  # MAX_RETRY_ATTEMPTS = 3
                try:
                    # Open workbook in read-only mode for safety
                    wb_source = excel.Workbooks.Open(
                        os.path.abspath(os.path.join(source_dir, excel_file)),
                        ReadOnly=True)
                    break
                except Exception:
                    # Retry after delay if Excel is busy (OLE error 0x800ac472)
                    time.sleep(1)  # RETRY_DELAY = 1 second

            # Skip file if unable to open after retries
            if not wb_source:
                continue

            # Access Main worksheet for data validation
            main_ws = wb_source.Worksheets('Main')

            # Initialize with core worksheet (always included)
            target_sheets = ['Main']

            # Define conditions for conditional worksheet inclusion
            # Each tuple: (cell_reference, worksheet_name)
            conditions = [
                ("L16", "Sulfide"),
                ("J16", "COD"),
                ("G16", "TS & VS"),
                ("H16", "TSS"),
                ("D34", "BOD5"),
                ("M16", "O&G"),
                ("E25", "Efficiency"),
                ("E26", "State Points")
            ]

            # Check each condition and add worksheets with data
            for cell, sheet_name in conditions:
                try:
                    if main_ws.Range(cell).Value:
                        target_sheets.append(sheet_name)
                except Exception:
                    # Skip invalid cell references gracefully
                    pass

            # Special handling for TDS worksheet (different validation logic)
            try:
                if wb_source.Worksheets("TDS").Range("F26").Value:
                    target_sheets.append("TDS")
            except Exception:
                # TDS worksheet or cell may not exist
                pass

            # Create new temporary workbook for PDF export
            wb_new = excel.Workbooks.Add()

            # Copy selected worksheets to temporary workbook
            for name in target_sheets:
                try:
                    # Copy worksheet before first sheet in new workbook
                    wb_source.Worksheets(name).Copy(Before=wb_new.Sheets(1))
                except Exception:
                    # Skip worksheets that can't be copied
                    pass

            # Remove unwanted worksheets from temporary workbook
            # Use reverse iteration by index to avoid enumeration issues
            for i in range(wb_new.Sheets.Count, 0, -1):
                sheet = wb_new.Sheets(i)
                if sheet.Name not in target_sheets:
                    sheet.Delete()

            # Export temporary workbook to PDF format
            wb_new.ExportAsFixedFormat(0, str(pdf_path))

            # Clean up temporary workbook
            wb_new.Saved = True
            wb_new.Close(False)

            # Close source workbook
            wb_source.Close(False)

            # Increment success counter
            processed_count += 1

        # Display success message with processing statistics
        messagebox.showinfo("نجاح", f"تم تحويل {processed_count} ملف بنجاح.")

    except Exception as e:
        # Display error message for any unhandled exceptions
        messagebox.showerror("خطأ", f"حدث خطأ أثناء التحويل: {str(e)}")

    finally:
        # Ensure Excel application is properly closed
        if excel:
            try:
                excel.Quit()
            except Exception:
                # Ignore errors during cleanup
                pass

        # Uninitialize COM apartment
        pythoncom.CoUninitialize()


# Standalone execution support
if __name__ == "__main__":
    # Create hidden root window for dialog ownership
    root = tk.Tk()
    root.withdraw()  # Hide main window, only show dialogs

    # Execute PDF conversion
    daily_pdf_report(root)
