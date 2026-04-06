# -*- coding: utf-8 -*-
"""
Kruger Lab Management System - Main Application Dashboard

This module serves as the entry point and main dashboard for the Kruger Tamiya Lab Management System.
It provides a user-friendly GUI with buttons for accessing different lab reporting and testing
functionalities including daily reports, monthly aggregation, quality assurance, PDF generation,
dry sludge testing, and external plant reports.

The application uses tkinter for the GUI and organizes functionality into a 2x4 button grid,
with each button routing to specialized modules. The system supports both Arabic and English,
with full Excel integration for data storage and manipulation.

Key Features:
    - Daily report management with automatic state tracking
    - Monthly report aggregation and updates
    - Quality control (QA/QC) report generation
    - PDF export functionality
    - Dry sludge analysis and testing
    - External plant report management
    - Access to advanced lab menu and settings

Dependencies:
    - tkinter: GUI framework
    - openpyxl: Excel file manipulation
    - Specialized modules: monthly_report, quality_report, daily_pdf, dry_sludge_module, etc.

Author: Kruger Lab System Development Team
Last Updated: 2026-04-06
Version: 2.0
"""

import os
import tkinter as tk
from datetime import datetime
from tkinter import messagebox
import external_report
from monthly_report import MonthlyReportUpdater
import quality_report
from daily_pdf import daily_pdf_report
from dry_sludge_module import dry_sludge_report
from file_checker_module import report_manager
from main_menu import open_main_menu


# Initialize the report management environment
report_manager.setup_env()


def handle_daily_report_logic(btn):
    """
    Handle daily report button click event.

    Creates a new daily report and updates the button state to reflect the current
    operational status. This function serves as a wrapper around the report_manager's
    core functionality to ensure proper state management and user feedback.

    Args:
        btn (tk.Button): The daily report button widget whose state will be updated
                        to show current report status (e.g., pending, completed).

    Returns:
        None

    Note:
        - The report_manager automatically determines if a new report needs to be created
        - Button state (text and color) is updated to reflect the current operational status
        - Integration with Excel templates for data persistence
    """
    report_manager.create_report()  # Create new daily report or update existing
    # Update button appearance to reflect current state
    report_manager.update_btn(btn)


def on_monthly():
    """
    Handle monthly report button click event.

    Initializes and runs the monthly report updater, which aggregates data from
    daily reports and generates comprehensive monthly analysis. The function ensures
    safe error handling and proper application cleanup after completion.

    Returns:
        None

    Note:
        - Hides the main window during processing
        - Exits the entire application gracefully after completion
        - Error messages are displayed to user via messagebox before exit

    Raises:
        Exception: Any exception during monthly report processing is caught and displayed
    """
    try:
        root.withdraw()  # Hide the main window during processing
        updater = MonthlyReportUpdater()
        updater.run()
    except Exception as e:
        messagebox.showerror("System Error", f"حدث خطأ: {e}")
    finally:
        # Exit the entire app cleanly so we don't try to destroy a dead window
        os._exit(0)


# ============================================================================
# Application Setup and Initialization
# ============================================================================

# Create main application window
root = tk.Tk()
root.title("Kruger Tamiya - Lab Management System")
root.geometry("750x550")
root.configure(bg="#f0f0f0")

# Define unified Arabic font for consistent display across all widgets
arabic_font = ("Arial", 18, "bold")

# Configure grid layout for responsive, proportional button sizing
# This ensures all buttons resize uniformly when window is resized
for i in range(4):
    root.grid_rowconfigure(i, weight=1)
for i in range(2):
    root.grid_columnconfigure(i, weight=1)

# ============================================================================
# Button Definitions - Organized by Rows
# ============================================================================

# --- Row 0: Main Reports ---
# Daily Report Button - Displays current day's operational status
btn_daily = tk.Button(root, font=arabic_font, bg="#ffffff", fg="#2c3e50", relief="groove",
                      command=lambda: [handle_daily_report_logic(btn_daily), root.destroy()])
btn_daily.grid(row=0, column=0, sticky="nsew", padx=15, pady=15)
report_manager.update_btn(btn_daily)  # Initialize button state immediately

# Monthly Report Button - Updates monthly aggregated data
btn_monthly = tk.Button(root, text="تحديث التقرير الشهري", font=arabic_font, bg="#ffffff",
                        command=lambda: [on_monthly(), root.destroy()])
btn_monthly.grid(row=0, column=1, sticky="nsew", padx=15, pady=15)

# --- Row 1: Quality Assurance and Exports ---
# Quality Assurance/Quality Control Report Button
btn_quality = tk.Button(root, text="تقارير الجودة (QA/QC)", font=arabic_font, bg="#ffffff",
                        command=lambda: quality_report.quality_report(root))
btn_quality.grid(row=1, column=0, sticky="nsew", padx=15, pady=15)

# PDF Export Button - Generates formatted PDF of daily reports
btn_pdf = tk.Button(root, text="Print to PDF", font=arabic_font, bg="#ffffff", fg="#c0392b",
                    command=lambda: daily_pdf_report(root))
btn_pdf.grid(row=1, column=1, sticky="nsew", padx=15, pady=15)

# --- Row 2: Specialized Testing ---
# Dry Sludge Testing Button - Analysis and tracking of dry sludge parameters
btn_sludge = tk.Button(root, text="اختبارات الحمأة الجافة", font=arabic_font, bg="#ffffff",
                       command=lambda: dry_sludge_report(root))
btn_sludge.grid(row=2, column=0, sticky="nsew", padx=15, pady=15)

# External Plant Reports Button - Management of data from external treatment facilities
btn_external = tk.Button(root, text="تقارير المحطات الخارجية", font=arabic_font, bg="#ffffff",
                         command=lambda: external_report.external_report(root))
btn_external.grid(row=2, column=1, sticky="nsew", padx=15, pady=15)

# --- Row 3: Navigation and Utility ---
# Open Reports Folder Button - Provides quick access to stored reports directory
btn_folder = tk.Button(root, text="فتح مجلد التقارير", font=arabic_font, bg="#bdc3c7",
                       command=lambda: [os.startfile(report_manager.get_path()), root.destroy()])
btn_folder.grid(row=3, column=0, sticky="nsew", padx=15, pady=15)

# Main Menu Button - Access advanced lab operations and settings
btn_menu = tk.Button(root, text="القائمة الرئيسية", font=arabic_font, bg="#34495e", fg="white",
                     command=lambda: [open_main_menu(root, arabic_font), root.withdraw()])
btn_menu.grid(row=3, column=1, sticky="nsew", padx=15, pady=15)

# ============================================================================
# Window Event Handling and Application Launch
# ============================================================================

# Configure window close button behavior
root.protocol("WM_DELETE_WINDOW", root.destroy)

# Start the application event loop
if __name__ == "__main__":
    try:
        root.mainloop()
    except Exception as e:
        # Log uncaught exceptions to error tracking file
        with open("error_log.txt", "a", encoding="utf-8") as f:
            f.write(f"{datetime.now()}: Application Error: {str(e)}\n")

        # إظهار رسالة للمستخدم
        messagebox.showerror("خطأ في النظام", f"حدث خطأ غير متوقع:\n{e}")
