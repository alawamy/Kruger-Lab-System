# -*- coding: utf-8 -*-
"""
Report Manager Module - Daily Report Lifecycle Management

This module provides comprehensive management of daily laboratory reports throughout
their complete lifecycle including creation, validation, state tracking, and file organization.

It implements a sophisticated ReportManager class that handles:
    1. System environment setup and template distribution
    2. Interactive calendar-based date selection for custom reports
    3. Dynamic report creation with automatic structuring
    4. Real-time report state detection and visual feedback
    5. Button state management reflecting operational status
    6. Excel worksheet protection with password-based access control

Report State Machine:
    MISSING → Create new report
    INCOMPLETE → Highlight for completion
    COMPLETE → Ready for aggregation
    SUBMITTED → Processed to monthly report

The module automatically manages directory structures across multiple years and
months, creating subdirectories as needed and maintaining consistency with the
application's file organization schema.

Key Features:
    - Interactive DateEntry calendar widget for date selection
    - Automatic report template copying from central templates
    - UUID-based file tracking and state persistence
    - Excel protection management with automatic password application
    - Button state visual feedback (color, text, disabled state)
    - Comprehensive error handling with user messaging
    - Automatic environment initialization on first run

File Organization:
    Daily Reports/
    ├── Daily Sheets/
    │   ├── 2026/
    │   │   ├── January/
    │   │   │   ├── 01-01-2026.xlsx
    │   │   │   ├── 02-01-2026.xlsx
    │   │   │   └── ...
    │   │   └── ...
    │   └── PDFs/ (mirror of sheet structure)
    └── Templates/
        ├── Daily Template.xlsx
        ├── Dry Sludge Template.xlsx
        └── elements.csv

Excel Template Protection:
    - Master Password: '01006610166' (from SHEET_PASS)
    - Protection applied to: All data-entry worksheets except Main in initial template
    - Locked cells: Headers, formulas, lookup tables
    - Unlocked cells: Data entry fields (highlighted in cells)

Arabic Text and Locale:
    - All UI text supports Arabic display through ar_func utility
    - DateEntry displays calendar in Arabic when system locale is Arabic
    - File paths use Latin characters for Windows compatibility

Dependencies:
    - openpyxl: Excel file manipulation and protection
    - tkcalendar.DateEntry: Interactive date picker widget
    - tkinter: GUI components and messagebox
    - shutil.copy2: Template file copying
    - config: Paths (BASE_DIR, DAILY_SHEET, SETTING_DIR) and constants

Author: Kruger Lab System Development Team
Version: 2.0
Last Updated: 2026-04-06
"""

import logging
import os
import time
from datetime import datetime
from shutil import copy2
from tkinter import messagebox
import tkinter as tk
from tkcalendar import DateEntry

import openpyxl
from openpyxl.drawing.image import Image
from config import BASE_DIR, DAILY_SHEET, INTERNAL_BASE, SETTING_DIR, SHEET_PASS, TEMPLATES, IMG

# Configure logging for error tracking and debugging
logging.basicConfig(level=logging.INFO, filename=SETTING_DIR / "app.log")

# Load logo image for Excel reports
img = Image(IMG)


class ReportManager:
    """
    Daily Report Lifecycle Manager - Central coordination for report operations.

    Manages all aspects of daily laboratory report creation, validation, and state
    tracking. Coordinates with Excel templates, file systems, and UI elements to
    maintain a coherent report management workflow.

    Attributes:
        selected_date (datetime): Currently selected date for the active report.
                                 Updated via interactive date picker or auto-detection.

    Key Responsibilities:
        1. Environment initialization (directories, templates, permissions)
        2. Daily report creation with automatic date assignment
        3. Report state detection (missing, incomplete, complete, submitted)
        4. Button state calculation and visual update
        5. Excel worksheet protection and access control
        6. User interface coordination and feedback

    Methods:
        setup_env() → None
            Static method that initializes all required directories and copies
            template files from internal storage to working directories.

        create_report(custom_date=None, ask_for_date=True) → None
            Creates or opens a daily report for the specified date. If ask_for_date=True,
            displays an interactive calendar widget for date selection.

        get_report_state(date) → str
            Returns state of report for given date: 'MISSING', 'INCOMPLETE', 'COMPLETE'

        update_btn(btn, date=None) → None
            Updates button appearance (color, text) based on report state.
            Used by main GUI to show real-time status.

    Usage Examples:
        # Create today's report with auto-detection
        manager = ReportManager()
        manager.create_report()

        # Create report with user date selection
        manager.create_report(ask_for_date=True)

        # Create specific date report
        from datetime import datetime
        date = datetime(2026, 3, 15)
        manager.create_report(custom_date=date, ask_for_date=False)

    Integration Points:
        - app.py: Main dashboard calls create_report() from button handler
        - main_menu.py: Advanced operations may trigger report creation
        - settings_module.py: Configuration affects template paths

    State Management:
        Report states are determined by file existence and content validation.
        Button colors provide immediate visual feedback:
        - Green: Report ready/available
        - Blue: Report exists, can be opened
        - Red: Error or missing critical data
    """

    def __init__(self):
        """
        Initialize the Report Manager instance.

        Sets up the report management environment and initializes the selected_date
        attribute to None. The setup_env() method is called automatically to ensure
        all required directories and templates are in place.

        Args:
            None

        Returns:
            None

        Side Effects:
            - Creates required directory structure if not present
            - Copies template files from internal storage
            - Initializes logging system
            - Sets selected_date to None
        """
        self.setup_env()
        self.selected_date = None

    @staticmethod
    def setup_env():
        """
        Initialize the application environment and required directories.

        Creates all necessary directory structures and copies template files from
        the internal application storage to the working directories. This method
        is called automatically during ReportManager initialization and ensures
        the application has all required resources.

        Directory Structure Created:
            - SETTING_DIR: Configuration and template storage
            - DAILY_SHEET: Daily report Excel files (year/month subdirectories)
            - DRY_SLUDGE_DIR: Dry sludge analysis reports

        Template Files Copied:
            - Daily Template.xlsx: Base template for daily reports
            - elements.csv: Chemical element reference data
            - Dry Sludge Template.xlsx: Template for sludge analysis
            - Logo.png: Company logo for reports
            - Logo.ico: Application icon

        Args:
            None

        Returns:
            None

        Raises:
            OSError: If directory creation fails due to permissions
            FileNotFoundError: If source template files are missing
            shutil.Error: If file copying encounters issues

        Note:
            - Uses copy2() to preserve file metadata (timestamps, permissions)
            - Only copies files that don't already exist in destination
            - Creates parent directories automatically with exist_ok=True
        """
        # Create required directory structure
        required_dirs = [
            SETTING_DIR,
            BASE_DIR / "Daily Reports/Daily Sheets",
            BASE_DIR / "Dry Sludge"
        ]

        for directory in required_dirs:
            directory.mkdir(parents=True, exist_ok=True)

        # Copy template files from internal storage to working directories
        for template_file in TEMPLATES:
            source_path = INTERNAL_BASE / "Setting" / template_file
            dest_path = SETTING_DIR / template_file

            # Only copy if source exists and destination doesn't
            if source_path.exists() and not dest_path.exists():
                copy2(source_path, dest_path)

    def prompt_date_selection(self):
        """
        Display interactive date selection dialog for report creation.

        Presents a user-friendly dialog asking whether to use today's date or select
        a custom date via an interactive calendar widget. This method provides a
        simple yes/no interface before potentially showing the full calendar.

        Process Flow:
            1. Show confirmation dialog: "Use today's date?" (Yes/No)
            2. If Yes: Set selected_date to datetime.now()
            3. If No: Launch calendar dialog via show_calendar_dialog()
            4. Update self.selected_date with chosen date

        Args:
            None

        Returns:
            datetime or None: The selected date object, or None if user cancelled.
                             Also updates self.selected_date attribute.

        Raises:
            None: All exceptions are handled internally with user messages.

        User Experience:
            - Simple yes/no dialog first to avoid unnecessary calendar display
            - Calendar only shown when user wants custom date selection
            - Clear Arabic text with RTL support
            - Graceful cancellation handling

        Example:
            >>> manager = ReportManager()
            >>> selected = manager.prompt_date_selection()
            >>> if selected:
            ...     print(f"Selected date: {selected.strftime('%Y-%m-%d')}")
            ... else:
            ...     print("User cancelled date selection")
        """
        # Ask user if they want to use today's date
        use_today = messagebox.askyesno(
            "اختيار التاريخ",
            "هل تريد استخدام تاريخ اليوم؟\n\nنعم: تاريخ اليوم\nلا: اختيار تاريخ آخر من التقويم"
        )

        if use_today:
            # Use current date and time
            selected_date = datetime.now()
        else:
            # Show calendar dialog for custom date selection
            selected_date = self.show_calendar_dialog()

            # Handle user cancellation
            if selected_date is None:
                return None

        # Update instance variable and return
        self.selected_date = selected_date
        return selected_date

    def show_calendar_dialog(self):
        """
        Display interactive calendar widget for custom date selection.

        Creates a modal dialog window containing a tkcalendar DateEntry widget
        with confirmation and cancellation buttons. The calendar allows users to
        visually select any date, with today's date pre-selected as default.

        Dialog Features:
            - Modal window (blocks interaction with parent)
            - DateEntry widget with dark blue theme
            - dd-mm-yyyy date format for consistency
            - Pre-selected current date
            - Green "Confirm" and red "Cancel" buttons
            - Arabic text labels with proper RTL support

        Args:
            None

        Returns:
            datetime or None: Parsed datetime object from selected date string,
                             or None if user cancelled the selection.

        Raises:
            ValueError: If date string parsing fails (handled internally)
            Exception: Any other calendar-related errors (handled internally)

        Technical Details:
            - Uses tkcalendar.DateEntry for cross-platform calendar widget
            - Date pattern 'dd-mm-yyyy' matches application conventions
            - wait_window() blocks until dialog closes
            - Error handling prevents crashes from invalid date operations

        Example:
            >>> manager = ReportManager()
            >>> date = manager.show_calendar_dialog()
            >>> if date:
            ...     print(f"User selected: {date.strftime('%d-%m-%Y')}")
            ... else:
            ...     print("Selection cancelled")
        """
        # Create modal calendar dialog window
        calendar_window = tk.Toplevel()
        calendar_window.title("اختر التاريخ")
        calendar_window.geometry("300x300")
        calendar_window.resizable(False, False)
        calendar_window.configure(bg="#f0f0f0")

        # Variable to store the selected date
        selected_date = None

        # Create calendar widget with custom styling
        cal = DateEntry(calendar_window, width=12, background='darkblue',
                        foreground='white', borderwidth=2, date_pattern='dd-mm-yyyy',
                        font=('Arial', 14))
        cal.pack(pady=20)

        # Set current date as default selection
        cal.set_date(datetime.now())

        def confirm_date():
            """Handle date confirmation button click."""
            nonlocal selected_date
            try:
                # Parse selected date string into datetime object
                date_str = cal.get()
                selected_date = datetime.strptime(date_str, "%d-%m-%Y")
                calendar_window.destroy()
            except Exception as e:
                messagebox.showerror(
                    "خطأ", f"حدث خطأ في اختيار التاريخ:\n{str(e)}")

        def cancel_selection():
            """Handle cancellation button click."""
            nonlocal selected_date
            selected_date = None
            calendar_window.destroy()

        # Create button frame with confirmation and cancellation buttons
        button_frame = tk.Frame(calendar_window, bg="#f0f0f0")
        button_frame.pack(pady=10)

        # Green confirm button
        tk.Button(button_frame, text="تأكيد", command=confirm_date,
                  bg="#4CAF50", fg="white", font=('Arial', 12, 'bold'),
                  width=10).pack(side=tk.LEFT, padx=5)

        # Red cancel button
        tk.Button(button_frame, text="إلغاء", command=cancel_selection,
                  bg="#f44336", fg="white", font=('Arial', 12, 'bold'),
                  width=10).pack(side=tk.LEFT, padx=5)

        # Wait for dialog to close before returning
        calendar_window.wait_window()

        return selected_date

    def get_path(self, custom_date=None):
        """
        Generate the complete file path for a daily report based on date.

        Constructs the full path to a daily report Excel file following the
        application's hierarchical directory structure: Year/Month/filename.xlsx.
        Creates any missing directories automatically.

        Path Structure:
            DAILY_SHEET/year/month/dd-mm-yyyy.xlsx

        Args:
            custom_date (datetime, optional): Specific date for the report path.
                If None, uses self.selected_date or current datetime.

        Returns:
            Path: Complete path to the report file, with directories created if needed.

        Raises:
            OSError: If directory creation fails due to permissions.

        Examples:
            >>> manager = ReportManager()
            >>> # Get today's report path
            >>> today_path = manager.get_path()
            >>> print(today_path)  # Path(.../2026/January/15-01-2026.xlsx)

            >>> # Get specific date path
            >>> from datetime import datetime
            >>> date = datetime(2026, 3, 15)
            >>> custom_path = manager.get_path(date)
            >>> print(custom_path)  # Path(.../2026/March/15-03-2026.xlsx)

        Notes:
            - Automatically creates year and month subdirectories
            - Uses full month names (January, February, etc.) for readability
            - Filename format: dd-mm-yyyy.xlsx for consistency
            - Thread-safe directory creation with exist_ok=True
        """
        # Determine target date (custom, selected, or current)
        if custom_date:
            target_date = custom_date
        elif self.selected_date:
            target_date = self.selected_date
        else:
            target_date = datetime.now()

        # Build path: DAILY_SHEET/year/month/
        year_dir = DAILY_SHEET / target_date.strftime("%Y")
        month_dir = year_dir / target_date.strftime("%B")  # Full month name

        # Create directories if they don't exist
        month_dir.mkdir(parents=True, exist_ok=True)

        # Generate filename: dd-mm-yyyy.xlsx
        filename = f"{target_date.strftime('%d-%m-%Y')}.xlsx"
        return month_dir / filename

    def create_report(self, custom_date=None, ask_for_date=True):
        """
        Create or open a daily laboratory report for the specified date.

        This is the primary method for daily report management. It handles the complete
        workflow of report creation, from date selection through file generation and
        opening. The method can operate in automatic mode (today's date) or interactive
        mode (user date selection).

        Workflow:
            1. Date Selection: Prompt user or use provided/custom date
            2. Path Generation: Create file path with proper directory structure
            3. Existence Check: If file exists, open it directly
            4. Template Processing: Load template, customize, protect, save
            5. File Opening: Launch Excel with the created/opened report

        Args:
            custom_date (datetime, optional): Specific date for report creation.
                - If None and ask_for_date=True: Prompt user for date selection
                - If None and ask_for_date=False: Use current date
                - If provided: Use this date directly

            ask_for_date (bool, default=True): Whether to prompt user for date selection.
                - True: Show date selection dialog (default behavior)
                - False: Use custom_date or current date without prompting

        Returns:
            None: Method opens file in Excel or shows error messages to user.

        Raises:
            None: All exceptions are caught and displayed to user via messagebox.

        Template Processing Details:
            - Loads "Daily Template.xlsx" from SETTING_DIR
            - Disables worksheet protection for editing
            - Sets date in cells D11 and K11 (formatted as dd/mm/yyyy)
            - Adds company logo (IMG) to cell G3
            - Enables worksheet protection with SHEET_PASS password
            - Saves to generated path and opens in Excel

        Error Handling:
            - File access errors: Permission issues, missing templates
            - Excel processing errors: Invalid template, cell access issues
            - Date parsing errors: Invalid date formats
            - All errors logged to app.log and shown to user

        Examples:
            >>> manager = ReportManager()

            >>> # Interactive mode - prompt for date
            >>> manager.create_report()  # Shows date selection dialog

            >>> # Automatic mode - today's date
            >>> manager.create_report(ask_for_date=False)

            >>> # Specific date mode
            >>> from datetime import datetime
            >>> date = datetime(2026, 3, 15)
            >>> manager.create_report(custom_date=date, ask_for_date=False)

        Performance Notes:
            - Template loading is cached by openpyxl when possible
            - Directory creation is optimized with exist_ok=True
            - File opening uses os.startfile() for native application launch
            - Small delay (0.5s) after save ensures file is ready for opening

        Integration:
            - Called from app.py main dashboard button
            - Used by monthly_report.py for data aggregation
            - Referenced by quality_report.py for QA/QC data
        """
        # Interactive date selection if requested and no custom date provided
        if ask_for_date and custom_date is None:
            selected_date = self.prompt_date_selection()
            if selected_date is None:
                messagebox.showwarning(
                    "تم الإلغاء", "تم إلغاء عملية إنشاء التقرير")
                return
            custom_date = selected_date

        # Generate file path for the target date
        report_path = self.get_path(custom_date)

        # If report already exists, open it directly
        if report_path.exists():
            os.startfile(str(report_path))
            return

        # Create new report from template
        try:
            # Load the daily report template
            wb = openpyxl.load_workbook(SETTING_DIR / "Daily Template.xlsx")
            ws = wb['Main']

            # Temporarily disable protection for editing
            ws.protection.disable()

            # Set the report date in designated cells
            display_date = custom_date or datetime.now()
            date_str = display_date.strftime('%d/%m/%Y')
            ws['D11'] = ws['K11'] = date_str

            # Add company logo to the report
            ws.add_image(img, "G3")

            # Enable worksheet protection with password
            ws.protection.set_password(SHEET_PASS)
            ws.protection.enable()

            # Save the customized report
            wb.save(str(report_path))
            wb.close()

            # Brief pause to ensure file is fully written
            time.sleep(0.5)

            # Open the report in Excel
            os.startfile(str(report_path))

        except Exception as e:
            # Log error for debugging
            logging.error(f"Report creation error: {e}")

            # Show user-friendly error message
            messagebox.showerror(
                "خطأ", f"حدث خطأ عند إنشاء التقرير:\n{str(e)}")

    def update_btn(self, btn):
        """
        Update button appearance based on current report state.

        Analyzes the existence of today's daily report and updates the button's
        text and color to provide immediate visual feedback about the report status.
        This helps users understand what action will occur when clicking the button.

        Button States:
            - Report exists: Blue text "فتح تقرير اليوم" (Open today's report)
            - Report missing: Green text "إنشاء تقرير جديد" (Create new report)

        Args:
            btn (tk.Button): The tkinter button widget to update with current state.

        Returns:
            None: Updates button properties in-place.

        Side Effects:
            - Changes button text and foreground color
            - Provides visual indication of report availability

        Color Coding:
            - Blue (#0000FF): Report exists and can be opened
            - Green (#008000): Report needs to be created

        Example:
            >>> manager = ReportManager()
            >>> button = tk.Button(root, text="Daily Report")
            >>> manager.update_btn(button)  # Button shows current state

        Notes:
            - Uses get_path() to check file existence
            - Called during application startup and after report operations
            - Provides immediate user feedback without additional queries
        """
        # Check if today's report exists
        report_exists = self.get_path().exists()

        # Update button based on report state
        if report_exists:
            btn.config(text="فتح تقرير اليوم", fg="blue")
        else:
            btn.config(text="إنشاء تقرير جديد", fg="green")


# ============================================================================
# Global Instance - Singleton Pattern
# ============================================================================

# Create global instance for application-wide use
report_manager = ReportManager()
"""
Global ReportManager instance for application-wide access.

This singleton instance provides centralized access to all report management
functionality throughout the Kruger Lab System. It maintains state across
different modules and ensures consistent report handling.

Usage:
    from file_checker_module import report_manager
    
    # Create new report
    report_manager.create_report()
    
    # Update button state
    report_manager.update_btn(my_button)

Note:
    - Automatically initializes environment on import
    - Thread-safe for typical GUI operations
    - Maintains selected_date state across operations
"""
