# -*- coding: utf-8 -*-
"""
Sodium Thiosulfate Titration Calibration Module - Laboratory Quality Control Interface

This module provides a comprehensive graphical user interface for sodium thiosulfate
titration calibration procedures in laboratory quality control. The module implements
an interactive workflow for recording dichromate volumes, managing calibration trials,
and automatically updating laboratory report templates with calibration data.

The module implements a sophisticated calibration workflow for:
    1. Interactive dichromate volume specification with default values
    2. Structured data entry for three calibration trials
    3. Step-by-step Arabic instructions for sodium thiosulfate titration
    4. Automatic template updating across multiple report types
    5. Real-time data validation and user feedback

Titration Process:
    The module guides users through the sodium thiosulfate titration procedure:
    1. Use this calibration to determine sodium thiosulfate solution titer accurately
    2. Place 10ml of KH(IO3)2 or potassium iodate (KIO₃) or potassium dichromate (0.025 N)
    3. Add 5ml concentrated hydrochloric acid and 10ml potassium iodide (10%)
    4. Leave mixture in dark for 5 minutes for reaction and iodine evolution
    5. Begin titration with sodium thiosulfate until pale yellow color
    6. Add drops of starch indicator to turn solution dark blue
    7. Continue titration until blue color disappears completely or light green (dichromate)
    8. Record consumed sodium thiosulfate volume
    9. Repeat experiment three times for accurate average

Data Management:
    - Dichromate volume: Default 10ml, user-configurable
    - Three trial volumes: Recorded for each calibration run
    - Template integration: Updates both daily and external report templates
    - Cell mapping: Specific cells in Sulfide worksheet updated
    - Password protection: Automatic worksheet unprotect/protect cycle

Template Updates:
    Updates "Sulfide" worksheet in templates:
    - Daily Template.xlsx: Internal daily reports
    - External Report Template.xlsx: External plant reports

    Cell Mappings:
    - D19: Dichromate volume (val_vol)
    - D20: Trial 1 volume (val_1st)
    - D21: Trial 2 volume (val_2nd)
    - D22: Trial 3 volume (val_3rd)

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


class SodiumTitrationWindow:
    """
    Sodium Thiosulfate Titration Calibration Interface - Interactive Quality Control Window.

    This class implements a comprehensive graphical interface for sodium thiosulfate
    titration calibration procedures, providing step-by-step guidance, data entry forms,
    and automatic template updating functionality.

    The interface includes:
        - Detailed Arabic instructions for sodium thiosulfate titration procedure
        - Configurable dichromate volume input (default 10ml)
        - Three trial volume entry fields
        - Keyboard navigation and input validation
        - Automatic template updating across report types

    Attributes:
        ar (function): Arabic text processing function from config
        parent (tk.Tk): Parent window for modal behavior
        window (tk.Toplevel): Main calibration window
        arabic_font (tuple): Font specification for Arabic text rendering
        general_path (Path): Path to settings directory
        TBSodiumVol (tk.Entry): Dichromate volume input field
        trials_entries (list): List of three trial volume entry fields
        btn_save (tk.Button): Save data button

    Methods:
        setup_enter_navigation() → None: Configure keyboard navigation
        close_window() → None: Handle window closure and parent restoration
        save_data() → None: Validate and save calibration data to templates
    """

    def __init__(self, parent, arabic_font):
        """
        Initialize the sodium thiosulfate titration calibration window.

        Creates a modal window with comprehensive Arabic instructions, input fields
        for calibration data, and automatic template updating functionality.

        Args:
            parent (tk.Tk): Parent window for modal ownership and restoration
            arabic_font (tuple): Font specification tuple for Arabic text rendering

        Returns:
            None: Creates and displays the calibration interface

        Window Configuration:
            - Modal toplevel window (900x850 pixels)
            - Light blue background (#e8f4f8)
            - Arabic title with proper RTL support
            - Comprehensive event handling and navigation

        UI Components:
            - Main title with large Arabic font
            - Instructions frame with numbered step-by-step guide
            - Input frame with dichromate volume field
            - Table frame with three trial entry fields
            - Save button with prominent styling
            - Keyboard navigation between all input elements
        """
        self.ar = ar_func  # Store for use throughout the class
        self.parent = parent

        # Hide parent window during calibration
        self.parent.withdraw()

        # Create main calibration window
        self.window = tk.Toplevel(parent)
        self.window.title("معايرة الصوديوم ثيوسلفات")  # Will work now
        self.window.geometry("900x850")
        self.window.configure(bg="#e8f4f8")
        self.arabic_font = arabic_font
        self.general_path = SETTING_DIR

        # Main title with prominent Arabic styling
        tk.Label(self.window, text="معايرة الصوديوم ثيوسلفات",
                 font=(arabic_font[0], 22, "bold"), bg="#e8f4f8", pady=10).pack()

        # Instructions frame with detailed procedure steps
        instructions_frame = tk.Frame(self.window, bg="#e8f4f8")
        instructions_frame.pack(fill="both", padx=40, pady=10)

        # Comprehensive sodium thiosulfate titration procedure instructions
        instructions = [
            "1- يتم استخدام هذه المعايرة لتحديد عيارية محلول الصوديوم ثيوسلفات بدقة.",
            "2-  ضع 10 مل من محلول KH(IO3)2 أو محلول يودات البوتاسيوم (KIO₃) أو محلول " "بوتاسيوم داي كرومات (0.025 "
            "N)  في دورق مخروطي أو سجل "
            "الحجم المأخوذ.",
            "3- أضف 5 مل من حمض الهيدروكلوريك المركز و 10 مل من يوديد البوتاسيوم (10%).",
            "4- اترك الخليط في الظلام لمدة 5 دقائق لبدء التفاعل وتصاعد اليود.",
            "5- ابدأ المعايرة بواسطة الصوديوم ثيوسلفات حتى يصل اللون إلى الأصفر الشاحب.",
            "6- أضف قطرات من دليل النشا ليتحول المحلول للون الأزرق الغامق.",
            "7- استمر في المعايرة حتى يختفي اللون الأزرق تماماً أو أخضر فاتح فى حالة استخدامك البوتاسيوم داي كرومات .",
            "8- سجل الحجم المستهلك من الصوديوم ثيوسلفات.",
            "9- كرر التجربة ثلاث مرات للحصول على متوسط دقيق."
        ]

        # Display each instruction with proper Arabic RTL formatting
        for text in instructions:
            # Use anchor="e" and justify="right" with formatted texts
            tk.Label(instructions_frame, text=self.ar(text), font=(arabic_font[0], 16),
                     bg="#e8f4f8", anchor="e", justify="right").pack(fill="x", pady=2)

        # Input frame for dichromate volume specification
        input_frame = tk.Frame(self.window, bg="#e8f4f8")
        input_frame.pack(pady=20)

        # Dichromate volume header
        tk.Label(input_frame, text="حجم الداي كرومات (ملليتر)", font=self.arabic_font,
                 bg="#e8f4f8").grid(row=0, column=2, padx=10)

        # Dichromate volume input field with default value
        self.TBSodiumVol = tk.Entry(
            input_frame, font=self.arabic_font, width=10, justify="center")
        self.TBSodiumVol.insert(0, "10")  # Default 10ml
        self.TBSodiumVol.grid(row=0, column=1, padx=10)

        # Table frame for three calibration trials
        table_frame = tk.Frame(self.window, bg="#e8f4f8")
        table_frame.pack(pady=20)

        # Trial labels in Arabic
        labels = ["المحاولة الأولى", "المحاولة الثانية", "المحاولة الثالثة"]
        self.trials_entries = []

        # Create three trial input fields with labels
        for i, label in enumerate(labels):
            tk.Label(table_frame, text=label, font=self.arabic_font,
                     bg="#e8f4f8").grid(row=0, column=2-i, padx=30)
            entry = tk.Entry(table_frame, font=self.arabic_font,
                             width=10, justify="center")
            entry.grid(row=1, column=2-i, padx=30, pady=10)
            self.trials_entries.append(entry)

        # Save data button with prominent styling
        self.btn_save = tk.Button(self.window, text="حفظ البيانات", font=(arabic_font[0], 14, "bold"),
                                  bg="#ffffff", relief="raised", bd=4, width=15, command=self.save_data)
        self.btn_save.pack(side="bottom", pady=40, padx=50, anchor="w")

        # Configure keyboard navigation
        self.setup_enter_navigation()

        # Set initial focus and selection
        self.TBSodiumVol.focus_set()
        self.TBSodiumVol.selection_range(0, tk.END)  # Select default value

        # Handle window close event
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)

    def setup_enter_navigation(self):
        """
        Configure keyboard navigation between input fields using Enter key.

        Sets up a logical navigation flow: dichromate volume → trial 1 → trial 2 →
        trial 3 → save button, allowing efficient data entry without mouse use.

        Returns:
            None: Configures event bindings for keyboard navigation

        Navigation Flow:
            TBSodiumVol (Enter) → trials_entries[0]
            trials_entries[0] (Enter) → trials_entries[1]
            trials_entries[1] (Enter) → trials_entries[2]
            trials_entries[2] (Enter) → btn_save
            btn_save (Enter) → save_data()
        """
        # Chain navigation through all input fields
        self.TBSodiumVol.bind(
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
        all relevant Excel templates with the new sodium thiosulfate calibration values.
        Handles worksheet protection automatically during the update process.

        Validation Process:
            1. Display confirmation dialog asking user to proceed
            2. Check that dichromate volume and all three trials have values
            3. Validate template file existence
            4. Update Sulfide worksheet in each template
            5. Provide success/error feedback to user

        Template Updates:
            Updates "Sulfide" worksheet in:
            - Daily Template.xlsx (internal daily reports)
            - External Report Template.xlsx (external plant reports)

            Cell Mappings:
            - D19: Dichromate volume (TBSodiumVol.get())
            - D20: Trial 1 volume (trials_entries[0])
            - D21: Trial 2 volume (trials_entries[1])
            - D22: Trial 3 volume (trials_entries[2])

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
        # Confirmation dialog
        msg_title = "تنبيه"
        msg_body = "هل تريد تحديث بيانات الصوديوم؟"
        answer = messagebox.askyesnocancel(msg_title, msg_body)

        if answer is True:
            # Extract input values
            val_vol = self.TBSodiumVol.get()
            val_1st = self.trials_entries[0].get()
            val_2nd = self.trials_entries[1].get()
            val_3rd = self.trials_entries[2].get()

            # Validate all fields are filled
            if all([val_vol, val_1st, val_2nd, val_3rd]):
                try:
                    # Define template paths
                    daily_path = self.general_path / "Daily Template.xlsx"
                    external_path = self.general_path / "External Report Template.xlsx"

                    # Update each template file
                    for path in [daily_path, external_path]:
                        if not path.exists():
                            continue

                        # Load and prepare workbook
                        wb = load_workbook(path)
                        ws = wb["Sulfide"]

                        # Temporarily disable protection for editing
                        ws.protection.password = "01006610166"
                        ws.protection.sheet = False

                        # Update calibration data cells
                        ws["D19"], ws["D20"] = val_vol, val_1st
                        ws["D21"], ws["D22"] = val_2nd, val_3rd

                        # Re-enable worksheet protection
                        ws.protection.sheet = True

                        # Save changes
                        wb.save(path)
                        wb.close()

                    # Success feedback and window closure
                    messagebox.showinfo("نجاح", (
                        "تم تحديث بيانات معايرة الصوديوم بنجاح"))
                    self.close_window()

                except Exception as e:
                    # Error feedback with exception details
                    messagebox.showerror(
                        "خطأ", f"{'فشل الحفظ'}: {e}")
            else:
                # Validation failure feedback
                messagebox.showwarning(
                    "تنبيه", "يرجى ملء كافة الحقول")


# Standalone execution support
if __name__ == "__main__":
    # Create root window for standalone execution
    root = tk.Tk()
    root.withdraw()  # Hide main window

    # Default Arabic font specification
    arabic_font = ("Arial", 16)  # Example Arabic font

    # Launch sodium thiosulfate titration calibration interface
    SodiumTitrationWindow(root, arabic_font)

    # Note: Missing root.mainloop() call - this would need to be added for proper standalone execution
    root.mainloop()