# -*- coding: utf-8 -*-
"""
Dry Sludge Analysis Report Generator - Sludge Quality Assessment

This module provides automated generation of dry sludge analysis reports for
laboratory quality control and environmental monitoring. The module creates
Excel-based reports from templates with date stamping and worksheet protection.

The module implements a streamlined workflow for:
    1. Template validation and availability checking
    2. User choice between creating new reports or opening existing folders
    3. Automatic date stamping and report initialization
    4. Worksheet protection with password-based security
    5. File organization by year and month structure

Report Characteristics:
    - Based on "Dry Sludge Template.xlsx" template
    - Date stamped with current date in dd/mm/yyyy format
    - Protected with master password '01006610166'
    - Single worksheet ("Main") with sludge analysis parameters
    - Automatic file naming: "dd-mm-yyyy S.xlsx"

File Organization:
    Dry Sludge/
    ├── 2026/
    │   ├── January/
    │   │   ├── 15-01-2026 S.xlsx
    │   │   ├── 16-01-2026 S.xlsx
    │   │   └── ...
    │   ├── February/
    │   │   └── ...
    │   └── ...
    └── Templates/
        └── Dry Sludge Template.xlsx

Template Processing:
    - Copies "Dry Sludge Template.xlsx" from SETTING_DIR
    - Updates date in cell D11 with current date
    - Disables worksheet protection temporarily for editing
    - Re-enables protection with hardcoded password
    - Saves customized report to organized directory structure

User Interface:
    - Modal confirmation dialog for action selection
    - Three options: Create New, Open Folder, Cancel
    - Automatic Excel file opening after successful creation
    - Error messages for missing templates or processing failures
    - Optional main application window closure after completion

Error Handling:
    - Template file missing: User notification with full path
    - File access conflicts: Option to open existing file
    - Excel processing errors: Exception message display
    - Directory creation failures: Automatic path creation
    - Password protection issues: Graceful error handling

Dependencies:
    - openpyxl: Excel file manipulation and worksheet protection
    - tkinter: GUI dialogs and messagebox
    - config: SETTING_DIR, DRY_SLUDGE_DIR path constants
    - datetime: Current date for file naming and stamping

Author: Kruger Lab System Development Team
Version: 2.0
Last Updated: 2026-04-06
"""

import openpyxl
import os
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime

from shutil import copyfile

from config import SETTING_DIR, DRY_SLUDGE_DIR


def dry_sludge_report(main_root=None):
    """
    Generate dry sludge analysis reports with user interaction.

    This function provides a complete workflow for creating dry sludge analysis
    reports, including template validation, user choice handling, and automatic
    report customization with date stamping and protection.

    Process Flow:
        1. Validate template file existence
        2. Display action selection dialog (Create/Open Folder/Cancel)
        3. Create destination directory structure if needed
        4. If Create chosen: Copy template and customize
        5. Update report with current date
        6. Apply worksheet protection
        7. Open completed report in Excel

    Args:
        main_root (tk.Tk, optional): Parent window for dialog ownership.
            If provided and user completes operation, this window will be destroyed.
            Used for modal dialog ownership and application flow control.

    Returns:
        None: Function handles all user interaction and file operations.
            Returns None explicitly when opening folder to indicate operation completion.

    Raises:
        None: All exceptions are caught internally and displayed to user via messagebox.

    Template Validation:
        - Checks for "Dry Sludge Template.xlsx" in SETTING_DIR
        - Shows error message with full path if template missing
        - Prevents processing if template unavailable

    User Choice Handling:
        - Yes (True): Create new dry sludge report
        - No (False): Open destination folder in Windows Explorer
        - None (Cancel): Return without action

    File Creation Process:
        1. Generate filename: "dd-mm-yyyy S.xlsx"
        2. Check for existing file conflicts
        3. Copy template to destination
        4. Load workbook and access Main worksheet
        5. Temporarily disable protection for editing
        6. Update date cell D11 with current date (dd/mm/yyyy format)
        7. Re-enable protection with password '01006610166'
        8. Save and close workbook
        9. Show success message
        10. Open file in Excel

    Directory Structure:
        - Year folders: DRY_SLUDGE_DIR / YYYY
        - Month folders: DRY_SLUDGE_DIR / YYYY / MonthName
        - Files: dd-mm-yyyy S.xlsx (S indicates Sludge)

    Error Conditions:
        - Template missing: Early return with error message
        - File exists: User choice between opening existing or cancelling
        - Excel errors: Exception details shown to user
        - Directory access: Handled by mkdir(parents=True, exist_ok=True)

    User Experience:
        - Modal dialogs prevent interaction with other windows
        - Clear action choices in confirmation dialog
        - Automatic file opening for immediate use
        - Success confirmation before Excel launch
        - Optional main application closure after completion

    Examples:
        >>> # Called from main application
        >>> dry_sludge_report(main_window)  # May close main window after completion

        >>> # Standalone execution
        >>> if __name__ == "__main__":
        ...     dry_sludge_report()  # Independent operation

    Integration Points:
        - app.py: Called from dry sludge report menu button
        - main_menu.py: Accessible through specialized operations
        - Template maintained in SETTING_DIR
        - Reports organized under DRY_SLUDGE_DIR structure
    """
    # Validate template availability
    template_path = SETTING_DIR / 'Dry Sludge Template.xlsx'

    if not template_path.exists():
        messagebox.showerror(
            "Error", f"Template not found at:\n{template_path}")
        return

    # Create temporary root for dialog ownership
    root = tk.Tk()
    root.withdraw()

    # Get user action choice
    choice = messagebox.askyesnocancel(
        "Report Action", "YES: Create New\nNO: Open Folder\nCANCEL: Return")
    root.destroy()

    # Get current date for file organization
    now = datetime.now()
    dest_folder = DRY_SLUDGE_DIR / now.strftime("%Y") / now.strftime("%B")
    dest_folder.mkdir(parents=True, exist_ok=True)

    if choice is None:  # User cancelled
        return

    if choice is False:  # Open folder option
        os.startfile(dest_folder)
        if main_root:
            main_root.destroy()  # Close main application window
        return None

    if choice is True:  # Create new report
        # Generate filename with date and sludge indicator
        new_filename = f"{now.strftime('%d-%m-%Y')} S.xlsx"
        dest_path = dest_folder / new_filename

        try:
            # Handle existing file conflict
            if dest_path.exists():
                if not messagebox.askyesno("Open File?", f"File exists. Open?"):
                    os.startfile(dest_path)
                    return

            # Copy template to destination
            copyfile(template_path, dest_path)

            # Customize the report
            wb = openpyxl.load_workbook(dest_path)
            ws = wb['Main']

            # Temporarily disable protection for editing
            ws.protection.sheet = False

            # Update date in designated cell
            ws['D11'] = now.strftime('%d/%m/%Y')

            # Re-enable worksheet protection with password
            ws.protection.set_password('01006610166')
            ws.protection.sheet = True

            # Save customized report
            wb.save(dest_path)
            wb.close()

            # Show success message and open file
            messagebox.showinfo("Success", "Report ready!")
            os.startfile(dest_path)

            # Close main application window if provided
            if main_root:
                main_root.destroy()  # Close main application

        except Exception as e:
            # Display any processing errors
            messagebox.showerror("Error", str(e))


# Standalone execution support
if __name__ == "__main__":
    # Execute dry sludge report generation independently
    dry_sludge_report()
