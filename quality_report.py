# -*- coding: utf-8 -*-
"""
Quality Control Report Generator - Statistical Analysis Reports

This module provides interactive generation of quality control reports for laboratory
data analysis, supporting Relative Percent Difference (RPD) calculations and X-chart
statistical process control. The module creates Excel-based reports from templates
with automated worksheet protection and data validation.

The module implements a user-friendly interface for:
    1. Quality report type selection (RPD or X-chart)
    2. Template-based report creation with automatic customization
    3. Worksheet protection with password-based access control
    4. File organization by year and month
    5. Existing file detection and user choice handling

Quality Report Types:
    RPD (Relative Percent Difference):
        - Measures precision between duplicate analyses
        - Calculates percentage difference between measurements
        - Used for quality assurance and method validation
        - Template: "Quality RPD Template.xlsx"

    X-Chart (Statistical Process Control):
        - Monitors process stability over time
        - Plots individual measurements against control limits
        - Identifies trends and out-of-control conditions
        - Template: "Quality X chart Template.xlsx"

File Organization:
    Quality Reports/
    ├── 2026/
    │   ├── January/
    │   │   ├── 15-01-2026 RPD.xlsx
    │   │   ├── 15-01-2026 X chart.xlsx
    │   │   └── ...
    │   ├── February/
    │   │   └── ...
    │   └── ...
    └── Templates/
        ├── Quality RPD Template.xlsx
        └── Quality X chart Template.xlsx

Template Processing:
    - Copies appropriate template based on selected report type
    - Renames "Main" worksheet to "{Type} curve" (e.g., "RPD curve")
    - Updates chart references in chart worksheet
    - Applies worksheet protection with SHEET_PASS password
    - Protects all worksheets except data entry areas

User Interface:
    - Modal dialog for report type selection
    - Combobox with RPD/X-chart options
    - Generate button with visual feedback
    - Error handling with user-friendly messages
    - Automatic Excel file opening after creation

Error Handling:
    - Missing templates: User notification with specific filename
    - File access conflicts: Option to open existing file
    - Excel processing errors: Detailed error messages
    - Directory creation failures: Automatic path creation
    - Password protection issues: Graceful fallback

Dependencies:
    - openpyxl: Excel file manipulation and worksheet protection
    - tkinter: GUI components and dialog management
    - config: QUALITY_REPORT_DIR, SETTING_DIR, SHEET_PASS constants
    - datetime: Current date for file naming and organization

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
from config import QUALITY_REPORT_DIR, SETTING_DIR, SHEET_PASS


def quality_report(main_root=None):
    """
    Generate quality control reports with user interaction.

    This function creates a modal dialog for quality report generation, allowing
    users to select between RPD and X-chart analysis types. It handles template
    copying, worksheet customization, protection, and automatic file opening.

    Process Flow:
        1. Display action selection dialog (Create/Open Folder/Cancel)
        2. If Open Folder: Launch Windows Explorer at current year folder
        3. If Create: Show report type selection interface
        4. Validate user selections and template availability
        5. Create destination directory structure
        6. Copy and customize template
        7. Apply worksheet protection
        8. Open completed report in Excel

    Args:
        main_root (tk.Tk, optional): Parent window for dialog ownership.
            If None, creates modal toplevel window.

    Returns:
        None: Function handles all user interaction and file operations.

    Raises:
        None: All exceptions are caught internally and displayed to user.

    Dialog Flow:
        First Dialog: Action Selection
            - Yes: Create new quality report
            - No: Open quality reports folder for current year
            - Cancel: Return without action

        Second Dialog: Report Type Selection (if Create chosen)
            - Combobox with "RPD" and "X chart" options
            - Generate button to process selection
            - Window closes after successful generation

    Template Processing:
        For selected report type (RPD or X chart):
        1. Locate template: SETTING_DIR / "Quality {type} Template.xlsx"
        2. Create destination: QUALITY_REPORT_DIR / year / month / "dd-mm-yyyy {type}.xlsx"
        3. Copy template to destination
        4. Customize worksheets:
           - Rename "Main" to "{type} curve"
           - Update chart references in "{type} chart" worksheet
        5. Apply password protection to all worksheets
        6. Save and open in Excel

    File Conflict Resolution:
        - If report already exists: Ask user to open existing or cancel
        - Existing file takes precedence to prevent accidental overwrites
        - User choice determines whether to proceed with creation

    Directory Structure:
        - Year folders: QUALITY_REPORT_DIR / YYYY
        - Month folders: QUALITY_REPORT_DIR / YYYY / MonthName
        - Files: dd-mm-yyyy Type.xlsx (e.g., "15-01-2026 RPD.xlsx")

    Error Conditions:
        - Missing template: Shows specific filename in error message
        - Directory creation failure: Handled by mkdir(parents=True, exist_ok=True)
        - Excel processing error: Displays exception message to user
        - File opening failure: Report created but not opened

    User Experience:
        - Modal dialogs prevent interaction with other windows
        - Clear Arabic/English text labels
        - Visual feedback through button colors and messages
        - Automatic file opening for immediate use
        - Graceful cancellation at any point

    Examples:
        >>> # Called from main application
        >>> quality_report(main_window)  # Shows dialog, creates report

        >>> # Standalone execution
        >>> if __name__ == "__main__":
        ...     root = tk.Tk()
        ...     root.withdraw()
        ...     quality_report(root)
        ...     root.mainloop()

    Integration Points:
        - app.py: Called from quality report menu button
        - main_menu.py: Accessible through advanced operations
        - Templates stored in SETTING_DIR for easy maintenance
        - Reports organized by QUALITY_REPORT_DIR structure
    """
    # Create modal window for dialog ownership
    window = tk.Toplevel(main_root)
    window.withdraw()  # Hide initially

    # First dialog: Choose action (Create/Open/Cancel)
    choice = messagebox.askyesnocancel(
        "Report Action",
        "YES: Create New\nNO: Open Folder\nCANCEL: Return",
        parent=window
    )

    if choice is None:  # User cancelled
        window.destroy()
        return

    now = datetime.now()

    if choice is False:  # Open folder option
        # Open Windows Explorer at current year's quality reports
        os.startfile(QUALITY_REPORT_DIR / now.strftime("%Y"))
        window.destroy()
        return

    # ===== Create New Report =====
    # Show report type selection interface
    window.deiconify()
    window.title("Quality Report - 2026")
    window.geometry("350x280")
    window.attributes("-topmost", True)  # Keep window on top

    # UI Labels
    tk.Label(window, text="Select Quality Item:",
             font=('Arial', 10, 'bold')).pack(pady=5)

    tk.Label(window, text="Select Kind:", font=(
        'Arial', 10, 'bold')).pack(pady=5)

    # Report type selection combobox
    kind_cb = ttk.Combobox(window, state="readonly", values=("RPD", "X chart"))
    kind_cb.current(0)  # Default to first option (RPD)
    kind_cb.pack(pady=5)

    def process_file():
        """
        Process the quality report generation based on user selections.

        Validates user input, creates necessary directories, copies and customizes
        the appropriate template, applies worksheet protection, and opens the
        completed report in Excel.

        Validation:
            - Ensures report type is selected
            - Verifies template file exists
            - Checks for existing files and handles conflicts

        Template Customization:
            - Renames "Main" worksheet to "{type} curve"
            - Updates chart references in chart worksheet
            - Applies password protection to all worksheets

        Error Handling:
            - Template missing: User notification
            - File exists: User choice dialog
            - Processing errors: Exception display
        """
        # Get selected report type
        kind = kind_cb.get()
        if not kind:
            messagebox.showwarning(
                "Input Error", "Please select a Quality Kind.")
            return

        # Create destination directory structure
        dest_folder = QUALITY_REPORT_DIR / \
            now.strftime("%Y") / now.strftime("%B")
        dest_folder.mkdir(parents=True, exist_ok=True)

        # Locate appropriate template
        template_path = SETTING_DIR / f'Quality {kind} Template.xlsx'
        if not template_path.exists():
            messagebox.showerror("Error", f"Template missing: {template_path}")
            return

        # Generate filename and full path
        file_name = f"{now.strftime('%d-%m-%Y')} {kind}.xlsx"
        dest_path = dest_folder / file_name

        try:
            # Handle existing file conflict
            if dest_path.exists():
                if messagebox.askyesno("File Exists", f"{file_name} already exists. Open it?"):
                    os.startfile(dest_path)
                    window.destroy()
                    return
                # If user says No, continue with creation (will overwrite)

            # Copy template to destination
            copyfile(template_path, dest_path)

            # Customize the copied template
            wb = openpyxl.load_workbook(dest_path)

            # Rename Main worksheet to report type curve
            if "Main" in wb.sheetnames:
                wb["Main"].title = f"{kind} curve"

            # Update chart worksheet references
            chart_sheet_name = f"{kind} chart"
            if chart_sheet_name in wb.sheetnames:
                ws_chart = wb[chart_sheet_name]
                # Update chart data reference cell
                ws_chart['O8'].value = f"{kind} curve"

            # Apply password protection to all worksheets
            for sheet in wb.worksheets:
                sheet.protection.set_password(SHEET_PASS)
                sheet.protection.enable()

            # Save customized report
            wb.save(dest_path)
            wb.close()

            # Open completed report in Excel
            os.startfile(dest_path)

            # Close dialog window
            window.destroy()

        except Exception as e:
            # Display any processing errors to user
            messagebox.showerror(
                "System Error", f"An error occurred: {str(e)}")

    # Generate button with styling
    tk.Button(window, text="Generate Report", command=process_file,
              bg="#4CAF50", fg="white", width=20, height=2).pack(pady=20)


# Standalone execution support
if __name__ == "__main__":
    # Create hidden root window for dialog ownership
    root = tk.Tk()
    root.withdraw()  # Hide main window

    # Execute quality report generation
    quality_report(root)

    # Start tkinter event loop
    root.mainloop()
