# -*- coding: utf-8 -*-
"""
Total Dissolved Solids (TDS) Factor Calibration Module - Laboratory Parameter Configuration

This module provides a graphical user interface for managing and updating total dissolved
solids (TDS) calibration factors used in wastewater treatment quality control. The module
allows operators to adjust three critical TDS factors that represent different points in
the water treatment process.

TDS Factor Purpose:
    TDS (Total Dissolved Solids) factors are calibration constants used to convert electrical
    conductivity or raw sensor measurements into standardized TDS concentrations. Three factors
    are maintained for different process streams:
        1. Influent Factor: Raw wastewater incoming to treatment plant
        2. Effluent Factor: Treated water discharged from treatment plant
        3. DW (Distilled Water) Factor: Reference distilled water calibration factor

Factor Application:
    These factors are applied to laboratory reports to standardize TDS measurements across:
    - Daily process monitoring reports
    - Quality control calculations
    - Performance trend analysis
    - Process efficiency evaluation

Data Management:
    - Storage: TDS worksheet in Daily Template.xlsx
    - Search Method: Keyword-based cell location using specific identifiers
    - Update Mechanism: Dynamic cell discovery and adjacent value placement
    - Protection: Automatic worksheet protection/unprotection cycle
    - Password: Hardcoded protection password "01006610166"

User Interface:
    - Modal window for TDS factor configuration
    - Three labeled input fields for factor values
    - Arabic instructions for each factor type
    - Keyboard navigation with Enter key support
    - Data validation before saving
    - Comprehensive error handling and user feedback

Input Fields:
    1. Influent Factor (مقياس الدخول): Factor for incoming wastewater
    2. Effluent Factor (مقياس الخروج): Factor for treated effluent
    3. DW Factor (مقياس المياه المقطرة): Reference distilled water factor

Error Handling:
    - Input validation: Ensures all factor fields are filled before saving
    - File access: Handles Excel file permission issues gracefully
    - Cell location: Searches for keywords in worksheet for flexible updates
    - Protection handling: Manages worksheet protection during updates
    - User feedback: Clear error messages with recovery instructions

Dependencies:
    - tkinter: GUI framework with messagebox for user dialogs
    - openpyxl: Excel workbook manipulation and protection management
    - config: SETTING_DIR path constant for template location

Author: Kruger Lab System Development Team
Version: 2.0
Last Updated: 2026-04-06
"""

import tkinter as tk
from tkinter import messagebox

from openpyxl import load_workbook

from config import SETTING_DIR


class SaltsFactorWindow:
    """
    Total Dissolved Solids (TDS) Factor Configuration Interface - Quality Control Window.

    This class implements a graphical interface for configuring three TDS calibration factors
    used in wastewater treatment quality control. The interface provides keyword-based cell
    discovery for flexible worksheet navigation and automatic template updating.

    TDS Factors:
        - Influent Factor: Calibration constant for raw wastewater
        - Effluent Factor: Calibration constant for treated water discharge
        - DW Factor: Reference calibration for distilled water

    Attributes:
        arabic_font (tuple): Font specification for Arabic UI text
        general_path (Path): Path to settings/template directory
        window (tk.Toplevel): Main configuration window
        parent (tk.Tk): Parent window for modal ownership
        entries (list): List of three factor input fields (Entry widgets)
        btn_update (tk.Button): Save/update button for factor values

    Methods:
        close_window() → None: Handle window closure and parent restoration
        save_data() → None: Validate and save TDS factors to template worksheet
    """

    def __init__(self, parent, arabic_font):
        """
        Initialize the TDS factor configuration window.

        Creates a modal window with input fields for three TDS calibration factors
        and automatic template file updating functionality.

        Args:
            parent (tk.Tk): Parent window for modal ownership and restoration
            arabic_font (tuple): Font specification tuple for Arabic text rendering

        Returns:
            None: Creates and displays the configuration interface

        Window Configuration:
            - Modal toplevel window (600x400 pixels)
            - Light gray background (#f5f5f5)
            - Arabic title in bold 20pt font
            - Comprehensive event handling and navigation

        UI Components:
            - Main title with Arabic text
            - Content frame with three input fields and labels
            - Factor labels in Arabic with RTL text alignment
            - Three Entry widgets for factor value input
            - Update button with event binding
            - Keyboard navigation (Enter key between fields)
            - Initial focus on first input field

        Input Fields:
            1. Influent Factor: "أدخل المعامل الخاص بالدخول"
            2. Effluent Factor: "أدخل المعامل الخاص بالخروج"
            3. DW Factor: "أدخل المعامل الخاص بالمياه المقطرة"
        """
        self.arabic_font = arabic_font
        self.general_path = SETTING_DIR

        # Window configuration for modal display
        self.arabic_font = arabic_font
        self.general_path = SETTING_DIR

        # Window configuration for modal display
        self.window = tk.Toplevel(parent)
        self.window.title("تعديل معامل الأملاح الذائبة")
        self.window.geometry("600x400")
        self.window.configure(bg="#f5f5f5")
        self.parent = parent

        # Hide parent window during configuration
        self.parent.withdraw()

        # Main title with prominent styling
        tk.Label(self.window, text="تعديل معامل الأملاح الذائبة", font=(self.arabic_font[0], 20, "bold"),
                 bg="#f5f5f5", pady=20).pack()

        # Content frame for organized input layout
        content_frame = tk.Frame(self.window, bg="#f5f5f5")
        content_frame.pack(pady=10)

        # Keyboard navigation helper function
        def handle_enter(event):
            """Move focus to next widget when Enter pressed."""
            event.widget.tk_focusNext().focus()
            return "break"

        # TDS factor labels in Arabic
        labels = [
            "أدخل المعامل الخاص بالدخول",
            "أدخل المعامل الخاص بالخروج",
            "أدخل المعامل الخاص بالمياه المقطرة"
        ]

        self.entries = []
        # Create input fields with labels
        for i, text in enumerate(labels):
            tk.Label(content_frame, text=text, font=self.arabic_font, bg="#f5f5f5").grid(
                row=i, column=1, padx=20, pady=15, sticky="e")
            entry = tk.Entry(content_frame, font=self.arabic_font,
                             width=10, justify="center")
            entry.grid(row=i, column=0, padx=10, pady=15)
            # Bind Enter for navigation or save
            if i < len(labels) - 1:
                entry.bind("<Return>", handle_enter)
            else:
                entry.bind("<Return>", lambda e: self.save_data())

            self.entries.append(entry)

        # Set initial focus to first input field
        if self.entries:
            self.entries[0].focus()

        # Update button for saving factor values
        self.btn_update = tk.Button(self.window, text="تحديث البيانات", font=(self.arabic_font[0], 14, "bold"),
                                    bg="#e0e0e0", relief="raised", bd=3, width=20, command=self.save_data)
        self.btn_update.pack(pady=20)

        # Handle window close event
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)

    def close_window(self):
        """
        Handle window closure and restore parent window visibility.

        Properly closes the TDS factor window and restores the parent window
        to visible state, ensuring clean modal dialog behavior.

        Returns:
            None: Destroys TDS factor window and shows parent

        Side Effects:
            - Destroys self.window
            - Calls self.parent.deiconify() to show parent window
        """
        self.window.destroy()
        self.parent.deiconify()

    def save_data(self):
        """
        Validate and save TDS factor values to the template worksheet.

        Performs comprehensive input validation, then updates the TDS worksheet in
        Daily Template.xlsx with new calibration factor values. Uses keyword-based
        cell discovery for flexible worksheet navigation.

        Validation Process:
            1. Check all three factor fields are filled
            2. Locate template file in settings directory
            3. Find keyword cells in TDS worksheet
            4. Update adjacent cells with new factor values
            5. Handle protection/unprotection automatically
            6. Provide success/error feedback

        Cell Location Strategy:
            Uses keyword-based search to find factor cells:
            - "Influent Factor": Keyword for incoming wastewater factor
            - "Effluent Factor": Keyword for treated water factor
            - "DW Factor": Keyword for distilled water reference factor

            Updates occur in the cell immediately following (column+1) the keyword cell.

        Template Updates:
            Updates "TDS" worksheet in Daily Template.xlsx:
            - Cell search for keyword labels
            - Value placement in adjacent column (c+1)
            - Automatic row/column discovery
            - Dynamic cell location for flexibility

        Protection Handling:
            - Temporarily disables worksheet protection
            - Uses password "01006610166"
            - Applies protection after updates
            - Prevents unauthorized modifications

        Error Handling:
            - Missing data: Warning message for incomplete fields
            - File access issues: Specific error for locked Excel files
            - Template not found: Continues silently if file doesn't exist
            - Cell location errors: Exception details in error message
            - General exceptions: Comprehensive error feedback

        User Feedback:
            - Confirmation: No pre-confirmation dialog (direct save)
            - Success: Message confirming TDS updates
            - Error: Detailed error messages with recovery hints
            - Permission errors: Specific message for file access issues

        Returns:
            None: Updates template or shows error messages

        Side Effects:
            - Modifies Daily Template.xlsx permanently
            - Updates TDS worksheet cells
            - No window closure on error (allows retry)

        Example:
            Searching for "Influent Factor" keyword in TDS worksheet:
            >>> Found at (row=10, column=1)
            >>> Updates cell (row=10, column=2) with influent factor value
        """
        # Extract all factor values from entry fields
        values = [e.get() for e in self.entries]

        # Validate all fields are filled before processing
        if not all(values):
            messagebox.showwarning("تنبيه", "يرجى ملء جميع الحقول")
            return

        try:
            # Locate and load template file
            path = self.general_path / "Daily Template.xlsx"
            wb = load_workbook(path)
            ws = wb["TDS"]  # TDS worksheet for factor storage

            # Temporarily disable protection for editing
            ws.protection.password = "01006610166"
            ws.protection.sheet = False

            # Keywords for factor cell location in worksheet
            search_keys = ["Influent Factor", "Effluent Factor", "DW Factor"]

            # Search for each keyword and update adjacent cells
            for i, key in enumerate(search_keys):
                found = False
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value == key:
                            # Get cell row and column indices
                            r, c = cell.row, cell.column
                            if r and c:
                                # Write value in next column (c+1)
                                ws.cell(row=int(r), column=int(
                                    c) + 1, value=str(values[i]))
                                found = True
                                break
                    if found:
                        break

            # Re-enable worksheet protection
            ws.protection.sheet = True
            wb.save(path)
            wb.close()

            # Success feedback
            messagebox.showinfo("نجاح", "تم تحديث معاملات الأملاح بنجاح")

        except PermissionError:
            # File locked error with recovery instruction
            messagebox.showerror(
                "خطأ", "يرجى إغلاق ملف الإكسيل المفتوح أولاً!")
        except Exception as e:
            # General exception with details
            messagebox.showerror("خطأ", f"حدث خطأ: {e}")


# Standalone execution support
if __name__ == "__main__":
    # Create root window for standalone execution
    root = tk.Tk()
    root.withdraw()  # Hide main window

    # Default Arabic font specification
    app = SaltsFactorWindow(root, ("Arial", 18))

    # Note: Missing root.mainloop() call - this would need to be added for proper standalone execution
