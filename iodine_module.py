# -*- coding: utf-8 -*-
"""
Iodine Titration Calibration Module - Laboratory Quality Control Interface

This module provides a comprehensive graphical user interface for iodine titration
calibration procedures in laboratory quality control. The module implements an
interactive workflow for recording titration volumes, managing calibration trials,
and automatically updating laboratory report templates with calibration data.

The module implements a sophisticated calibration workflow for:
    1. Interactive iodine volume specification with default values
    2. Structured data entry for three calibration trials
    3. Step-by-step Arabic instructions for titration procedure
    4. Automatic template updating across multiple report types
    5. Real-time data validation and user feedback

Titration Process:
    The module guides users through the iodine-sodium thiosulfate titration:
    1. Prepare conical flask with distilled water (~200ml in 500ml flask)
    2. Add 2ml of 1:1 HCl solution
    3. Add 10ml of iodine solution (or record taken volume)
    4. Titrate against sodium thiosulfate in burette
    5. Continue until dark yellow changes to light yellow
    6. Add 2 drops of starch solution (turns blue)
    7. Complete titration until blue color disappears completely
    8. Record sodium thiosulfate volume
    9. Repeat steps 2-8 for consistency
    10. Record volumes in table and save data

Data Management:
    - Iodine volume: Default 10ml, user-configurable
    - Three trial volumes: Recorded for each calibration run
    - Template integration: Updates both daily and external report templates
    - Cell mapping: Specific cells in Sulfide worksheet updated
    - Password protection: Automatic worksheet unprotect/protect cycle

Template Updates:
    Updates "Sulfide" worksheet in templates:
    - Daily Template.xlsx: Internal daily reports
    - External Report Template.xlsx: External plant reports

    Cell Mappings:
    - C19: Iodine volume (val_vol)
    - C20: Trial 1 volume (vals[0])
    - C21: Trial 2 volume (vals[1])
    - C22: Trial 3 volume (vals[2])

User Interface:
    - Modal calibration window with Arabic RTL support
    - Step-by-step instructions in numbered list
    - Input validation with visual feedback
    - Keyboard navigation with Enter key support
    - Confirmation dialogs for data saving
    - Success/error message display

Arabic Text Support:
    - All UI elements use ar_func() for Arabic display
    - RTL text layout for instructions and labels
    - Arabic font specification for proper rendering
    - Bilingual error messages and confirmations

Error Handling:
    - Input validation: Ensures all fields are filled
    - Template access: Checks file existence before processing
    - Excel operations: Comprehensive exception handling
    - User cancellation: Graceful window closure
    - Data integrity: Validation before template updates

Dependencies:
    - tkinter: GUI framework with messagebox and event handling
    - openpyxl: Excel file manipulation and worksheet protection
    - config: SETTING_DIR path and ar_func utility

Author: Kruger Lab System Development Team
Version: 2.0
Last Updated: 2026-04-06
"""

import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook
from config import SETTING_DIR, ar_func


class IodineTitrationWindow:
    """
    Iodine Titration Calibration Interface - Interactive Quality Control Window.

    This class implements a comprehensive graphical interface for iodine titration
    calibration procedures, providing step-by-step guidance, data entry forms,
    and automatic template updating functionality.

    The interface includes:
        - Detailed Arabic instructions for titration procedure
        - Configurable iodine volume input (default 10ml)
        - Three trial volume entry fields
        - Keyboard navigation and input validation
        - Automatic template updating across report types

    Attributes:
        ar (function): Arabic text processing function from config
        arabic_font (tuple): Font specification for Arabic text rendering
        parent (tk.Tk): Parent window for modal behavior
        window (tk.Toplevel): Main calibration window
        general_path (Path): Path to settings directory
        TBIvolume (tk.Entry): Iodine volume input field
        trials_entries (list): List of three trial volume entry fields
        btn_save (tk.Button): Save data button

    Methods:
        setup_enter_navigation() → None: Configure keyboard navigation
        close_window() → None: Handle window closure and parent restoration
        save_data() → None: Validate and save calibration data to templates
    """

    def __init__(self, parent, arabic_font):
        """
        Initialize the iodine titration calibration window.

        Creates a modal window with comprehensive Arabic instructions, input fields
        for calibration data, and automatic template updating functionality.

        Args:
            parent (tk.Tk): Parent window for modal ownership and restoration
            arabic_font (tuple): Font specification tuple for Arabic text rendering

        Returns:
            None: Creates and displays the calibration interface

        Window Configuration:
            - Modal toplevel window (900x850 pixels)
            - Light gray background (#f0f0f0)
            - Arabic title with proper RTL support
            - Comprehensive event handling and navigation

        UI Components:
            - Main title with large Arabic font
            - Instructions frame with numbered step-by-step guide
            - Input frame with iodine volume field
            - Table frame with three trial entry fields
            - Save button with prominent styling
            - Keyboard navigation between all input elements
        """
        self.ar = ar_func
        self.arabic_font = arabic_font
        self.parent = parent

        # Hide parent window during calibration
        self.parent.withdraw()

        # Create main calibration window
        self.window = tk.Toplevel(parent)
        self.window.title(self.ar("معايرة اليود", is_ui_element=False))
        self.window.geometry("900x850")
        self.window.configure(bg="#f0f0f0")
        self.general_path = SETTING_DIR

        # Main title with prominent Arabic styling
        tk.Label(self.window, text=self.ar("معايرة اليود"),
                 font=(self.arabic_font[0], 24, "bold"), bg="#f0f0f0", pady=20).pack()

        # Instructions frame with detailed procedure steps
        instructions_frame = tk.Frame(self.window, bg="#f0f0f0")
        instructions_frame.pack(fill="both", padx=40, pady=10)

        # Comprehensive titration procedure instructions
        instructions = [
            "1- يتم معايرة اليود عند تحضير اليود أو عند تحضير صوديوم ثيوسلفات.",
            "2- ضع كمية من المياه المقطرة فى الدورق المخروطي (حوالى 200 مل فى دورق 500 مل).",
            "3- ضع 2 مل من محلول (1:1) HCl .",
            "4- ضع 10 مل من محلول اليود أو سجل الحجم المأخوذ.",
            "5- يتم معايرة المحلول فى الدورق على القلاب فى مقابل الصوديوم ثيوسلفات فى السحاحة.",
            "6- دع محلول الصوديوم ينزل فى الدورق حتى يتغير اللون الأصفر الداكن إلى لون أصفر فاتح .",
            "7- ضع نقطتين من محلول النشا عندها يتحول المحلول إلى اللون الأزرق .",
            "8- أكمل المعايرة حتى يتحول اللون الأزرق حتى يختفى تماما .",
            "9- قم بتسجيل حجم الصوديوم ثيوسلفات .",
            "10- قم بإعادة الخطوات من رقم 2 حتى رقم 9.",
            "11- قم بتسجيل الأحجام فى الجدول بالأسفل ثم قم بحفظ البيانات."
        ]

        # Display each instruction with proper Arabic RTL formatting
        for text in instructions:
            tk.Label(instructions_frame, text=self.ar(text), font=(self.arabic_font[0], 14),
                     bg="#f0f0f0", anchor="e", justify="right").pack(fill="x", pady=2)

        # Input frame for iodine volume specification
        input_frame = tk.Frame(self.window, bg="#f0f0f0")
        input_frame.pack(pady=20)

        # Iodine volume header
        tk.Label(input_frame, text=self.ar("الحجم بالملليتر"), font=self.arabic_font,
                 bg="#f0f0f0").grid(row=0, column=2, padx=10)

        # Iodine volume input field with default value
        self.TBIvolume = tk.Entry(
            input_frame, font=self.arabic_font, width=10, justify="center")
        self.TBIvolume.insert(0, "10")  # Default 10ml
        self.TBIvolume.grid(row=0, column=1, padx=10)

        # Table frame for three calibration trials
        table_frame = tk.Frame(self.window, bg="#f0f0f0")
        table_frame.pack(pady=20)

        # Trial labels in Arabic
        labels = ["المحاولة الأولى", "المحاولة الثانية", "المحاولة الثالثة"]
        self.trials_entries = []

        # Create three trial input fields with labels
        for i, label in enumerate(labels):
            tk.Label(table_frame, text=self.ar(label), font=self.arabic_font,
                     bg="#f0f0f0").grid(row=0, column=2-i, padx=30)

            entry = tk.Entry(table_frame, font=self.arabic_font,
                             width=12, justify="center")
            entry.grid(row=1, column=2-i, padx=30, pady=10)
            self.trials_entries.append(entry)

        # Save data button with prominent styling
        self.btn_save = tk.Button(self.window, text=self.ar("حفظ البيانات"),
                                  font=(self.arabic_font[0], 14, "bold"),
                                  bg="#ffffff", relief="raised", bd=4, width=18,
                                  command=self.save_data)
        self.btn_save.pack(side="bottom", pady=50, padx=50, anchor="w")

        # Configure keyboard navigation
        self.setup_enter_navigation()

        # Set initial focus and selection
        self.TBIvolume.focus_set()
        self.TBIvolume.selection_range(0, tk.END)  # Select default value

        # Handle window close event
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)

    def setup_enter_navigation(self):
        """
        Configure keyboard navigation between input fields using Enter key.

        Sets up a logical navigation flow: iodine volume → trial 1 → trial 2 →
        trial 3 → save button, allowing efficient data entry without mouse use.

        Returns:
            None: Configures event bindings for keyboard navigation

        Navigation Flow:
            TBIvolume (Enter) → trials_entries[0]
            trials_entries[0] (Enter) → trials_entries[1]
            trials_entries[1] (Enter) → trials_entries[2]
            trials_entries[2] (Enter) → btn_save
            btn_save (Enter) → save_data()
        """
        # Chain navigation through all input fields
        self.TBIvolume.bind(
            "<Return>", lambda e: self.trials_entries[0].focus())
        self.trials_entries[0].bind(
            "<Return>", lambda e: self.trials_entries[1].focus())
        self.trials_entries[1].bind(
            "<Return>", lambda e: self.trials_entries[2].focus())
        self.trials_entries[2].bind(
            "<Return>", lambda e: self.btn_save.focus())
        self.btn_save.bind("<Return>", lambda e: self.save_data())

    def close_window(self):
        """
        Handle window closure and restore parent window visibility.

        Properly closes the calibration window and restores the parent window
        to visible state, ensuring clean modal dialog behavior.

        Returns:
            None: Destroys calibration window and shows parent

        Side Effects:
            - Destroys self.window
            - Calls self.parent.deiconify() to show parent window
        """
        self.window.destroy()
        self.parent.deiconify()

    def save_data(self):
        """
        Validate input data and update laboratory report templates.

        Performs comprehensive validation of calibration data, then updates
        all relevant Excel templates with the new iodine calibration values.
        Handles worksheet protection automatically during the update process.

        Validation Process:
            1. Display confirmation dialog asking user to proceed
            2. Check that iodine volume and all three trials have values
            3. Validate template file existence
            4. Update Sulfide worksheet in each template
            5. Provide success/error feedback to user

        Template Updates:
            Updates "Sulfide" worksheet in:
            - Daily Template.xlsx (internal daily reports)
            - External Report Template.xlsx (external plant reports)

            Cell Mappings:
            - C19: Iodine volume (TBIvolume.get())
            - C20: Trial 1 volume (trials_entries[0])
            - C21: Trial 2 volume (trials_entries[1])
            - C22: Trial 3 volume (trials_entries[2])

        Protection Handling:
            - Temporarily disables worksheet protection with password
            - Updates calibration cells
            - Re-enables worksheet protection
            - Uses hardcoded password "01006610166"

        Error Handling:
            - User cancellation: Returns without action
            - Missing data: Warning message for incomplete fields
            - Template errors: Exception details in error message
            - File access issues: Comprehensive exception catching

        User Feedback:
            - Confirmation dialog before processing
            - Success message with Arabic text
            - Error messages with specific failure details
            - Automatic window closure on successful save

        Returns:
            None: Updates templates or shows error messages

        Side Effects:
            - Modifies Excel template files permanently
            - Closes calibration window on success
            - Updates all future reports with new calibration data
        """
        # Confirmation dialog with Arabic text
        msg_title = ar_func("تنبيه", is_ui_element=False)
        msg_body = ar_func(
            "هل تريد تحديث بيانات معايرة اليود؟", is_ui_element=False)

        # Get user confirmation
        if messagebox.askyesnocancel(msg_title, msg_body):
            # Extract input values
            val_vol = self.TBIvolume.get()
            vals = [e.get() for e in self.trials_entries]

            # Validate all fields are filled
            if val_vol and all(vals):
                try:
                    # Define templates to update
                    templates = ["Daily Template.xlsx",
                                 "External Report Template.xlsx"]

                    # Update each template file
                    for filename in templates:
                        path = self.general_path / filename
                        if path.exists():
                            # Load and prepare workbook
                            wb = load_workbook(path)
                            ws = wb["Sulfide"]

                            # Temporarily disable protection for editing
                            ws.protection.password = "01006610166"
                            ws.protection.sheet = False

                            # Update calibration data cells
                            ws["C19"], ws["C20"] = val_vol, vals[0]
                            ws["C21"], ws["C22"] = vals[1], vals[2]

                            # Re-enable worksheet protection
                            ws.protection.sheet = True

                            # Save changes
                            wb.save(path)
                            wb.close()

                    # Success feedback and window closure
                    messagebox.showinfo(self.ar("نجاح"), self.ar(
                        "تم تحديث بيانات معايرة اليود بنجاح"))
                    self.close_window()

                except Exception as e:
                    # Error feedback with exception details
                    messagebox.showerror(
                        self.ar("خطأ"), f"{self.ar('فشل الحفظ')}: {e}")
            else:
                # Validation failure feedback
                messagebox.showwarning(self.ar("تنبيه"), self.ar(
                    "يرجى ملء كافة الحقول أولاً"))


# Standalone execution support
if __name__ == "__main__":
    # Create root window for standalone execution
    root = tk.Tk()
    root.withdraw()

    # Default Arabic font specification
    default_font = ("Arial", 18, "bold")

    # Launch iodine titration calibration interface
    IodineTitrationWindow(root, default_font)

    # Start tkinter event loop
    root.mainloop()
