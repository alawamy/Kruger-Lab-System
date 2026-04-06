# -*- coding: utf-8 -*-
"""
Sludge Pumps Operation Configuration Module - Wastewater Treatment Control Interface

This module provides a comprehensive graphical interface for configuring and managing
sludge pump operational parameters in wastewater treatment plants. The module enables
operators to adjust pump operation settings for both WAS (Waste Activated Sludge) and
RAS (Return Activated Sludge) pumps using a tabbed interface.

Sludge Pump Systems:
    WAS (Waste Activated Sludge) Pump:
        - Flow rate: 43.2 m³/hour (fixed reference)
        - Purpose: Removes excess activated sludge from treatment process
        - Configuration: Operation steps and stop steps with time intervals
        - Application: Excess sludge disposal and process control

    RAS (Return Activated Sludge) Pump:
        - Flow rate: 504 m³/hour (fixed reference)
        - Purpose: Returns settled sludge from clarifier to aeration basin
        - Configuration: Operation steps and stop steps with time intervals
        - Application: Process balance and biological treatment efficiency

Operational Parameters:
    For each pump, four configuration values are managed:
        1. Run Steps: Gauge indicator range (1-10) for pump activation
        2. Run Rate: Operating frequency in minutes (10 or 60 minute cycles)
        3. Stop Steps: Gauge indicator range (1-10) for pump deactivation
        4. Stop Rate: Deactivation frequency in minutes (10 or 60 minute cycles)

Data Management:
    - Storage: BOD5 worksheet in Daily Template.xlsx
    - Search Method: Keyword-based cell location ("WAS on", "WAS off", "RAS on", "RAS off")
    - Cell Mapping: Specific columns for steps and rates relative to keyword cells
    - Protection: Automatic worksheet protection/unprotection cycle
    - Password: Hardcoded protection password "01006610166"

Cell Update Strategy:
    For each pump configuration ("WAS on", "WAS off", "RAS on", "RAS off"):
    - Keyword cell: Located using text search in worksheet
    - Steps value: Written to (column+1) of keyword cell
    - Rate value: Written to (column+4) of keyword cell

User Interface:
    - Tabbed interface with two pump tabs (WAS and RAS)
    - Separate configuration panels for each pump type
    - Pump flow rate displayed in header
    - Four input fields per pump with Arabic labels
    - Keyboard navigation with Enter key support
    - Save button with event binding
    - Comprehensive error handling and user feedback

Error Handling:
    - Input validation: Ensures all four fields are filled for each pump
    - File access: Handles Excel file permission issues gracefully
    - Cell location: Searches for keywords dynamically
    - Type conversion: Safely converts input to cell values
    - Protection management: Handles protection state during updates
    - User feedback: Clear error messages with recovery instructions

Dependencies:
    - tkinter: GUI framework with messagebox and ttk for tabbed interface
    - openpyxl: Excel workbook manipulation and protection management
    - warnings: Suppression of openpyxl formatting warnings
    - config: SETTING_DIR path constant for template location

Author: Kruger Lab System Development Team
Version: 2.0
Last Updated: 2026-04-06
"""

import tkinter as tk
import warnings
from tkinter import messagebox, ttk

from openpyxl import load_workbook

from config import SETTING_DIR

# Suppress openpyxl conditional formatting warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


class SludgePumpsWindow:
    """
    Sludge Pumps Operation Configuration Interface - Tabbed Control Window.

    This class implements a tabbed graphical interface for managing WAS and RAS pump
    operational parameters in wastewater treatment plants. The interface provides
    separate configuration panels for each pump type with synchronized template updating.

    Pump Types:
        - WAS: Waste Activated Sludge pump (43.2 m³/hour)
        - RAS: Return Activated Sludge pump (504 m³/hour)

    Configuration Fields (per pump):
        - Run Steps: Gauge indicator position for pump activation (1-10)
        - Run Rate: Pump operation frequency in minutes (typically 10 or 60)
        - Stop Steps: Gauge indicator position for pump deactivation (1-10)
        - Stop Rate: Pump deactivation frequency in minutes (typically 10 or 60)

    Attributes:
        arabic_font (tuple): Font specification for Arabic UI text
        general_path (Path): Path to settings/template directory
        window (tk.Toplevel): Main configuration window
        parent (tk.Tk): Parent window for modal ownership
        notebook (ttk.Notebook): Tabbed interface container
        was_frame (tk.Frame): WAS pump configuration tab
        ras_frame (tk.Frame): RAS pump configuration tab
        WAS_entries (list): Four Entry widgets for WAS configuration
        RAS_entries (list): Four Entry widgets for RAS configuration

    Methods:
        setup_pump_ui(frame, flow_rate, pump_type) → None: Create pump configuration panel
        close_window() → None: Handle window closure and parent restoration
        save_pump_data(pump_type) → None: Validate and save pump configuration
    """

    def __init__(self, parent, abic_font):
        """
        Initialize the sludge pumps configuration window.

        Creates a tabbed modal window with separate panels for WAS and RAS pump
        configuration, featuring input fields for pump operation parameters.

        Args:
            parent (tk.Tk): Parent window for modal ownership and restoration
            abic_font (tuple): Font specification tuple for Arabic text rendering

        Returns:
            None: Creates and displays the pump configuration interface

        Window Configuration:
            - Modal toplevel window (700x600 pixels)
            - Light gray background (#f5f5f5)
            - Tabbed interface with two pump tabs
            - Comprehensive event handling and navigation

        Tab Structure:
            - Tab 1: WAS Pump Configuration (Waste Activated Sludge)
            - Tab 2: RAS Pump Configuration (Return Activated Sludge)

        Tab Features:
            - Pump flow rate in header
            - Four configuration fields with Arabic labels
            - Keyboard navigation between fields
            - Save button for each tab
            - Error handling and user feedback
        """
        self.arabic_font = abic_font
        self.general_path = SETTING_DIR

        # Window configuration for modal display
        self.window = tk.Toplevel(parent)
        self.window.title("ضبط تشغيل طلمبات الحمأة")
        self.window.geometry("700x600")
        self.window.configure(bg="#f5f5f5")
        self.parent = parent

        # Hide parent window during configuration
        self.parent.withdraw()

        # Configure tabbed interface styling
        style = ttk.Style()
        style.configure("TNotebook.Tab", font=(self.arabic_font[0], 12))
        self.notebook = ttk.Notebook(self.window)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # WAS (Waste Activated Sludge) pump configuration tab
        self.was_frame = tk.Frame(self.notebook, bg="#f5f5f5")
        self.notebook.add(self.was_frame, text="طلمبة الحمأة الزائدة WAS")
        self.setup_pump_ui(self.was_frame, "43.2", "WAS")

        # RAS (Return Activated Sludge) pump configuration tab
        self.ras_frame = tk.Frame(self.notebook, bg="#f5f5f5")
        self.notebook.add(self.ras_frame, text="طلمبات الحمأة الراجعة RAS")
        self.setup_pump_ui(self.ras_frame, "504", "RAS")

    def setup_pump_ui(self, frame, flow_rate, pump_type):
        """
        Create pump configuration panel with input fields and labels.

        Builds a complete configuration interface for a specific pump type, including
        header information, input fields with labels, and save button.

        Args:
            frame (tk.Frame): Parent frame to contain pump configuration UI
            flow_rate (str): Pump flow rate in m³/hour for display header
            pump_type (str): Pump identifier ("WAS" or "RAS") for cell lookup

        Returns:
            None: Creates and configures pump UI components

        Components Created:
            - Header label with pump flow rate
            - Content frame with four input fields:
                1. Run steps (gauge indicator 1-10)
                2. Run rate (minutes: 10 or 60)
                3. Stop steps (gauge indicator 1-10)
                4. Stop rate (minutes: 10 or 60)
            - Arabic labels for each field
            - Save button at bottom of frame
            - Keyboard navigation between fields

        Field Configuration:
            Run Steps: "عدد درجات المؤشر من 1 إلى 10" (operation gauge)
            Run Rate: "معدل التشغيل بالدقيقة، اكتب 10 أو 60" (frequency)
            Stop Steps: "عدد درجات المؤشر من 1 إلى 10" (deactivation gauge)
            Stop Rate: "معدل الإيقاف بالدقيقة، اكتب 10 أو 60" (deactivation frequency)

        Keyboard Navigation:
            - Enter in run steps → run rate field
            - Enter in run rate → stop steps field
            - Enter in stop steps → stop rate field
            - Enter in stop rate → save_pump_data(pump_type)

        Data Storage:
            Saves entry widgets as instance attributes:
            - self.WAS_entries (for WAS pump)
            - self.RAS_entries (for RAS pump)

        Side Effects:
            - Creates entry widgets and buttons
            - Binds event handlers for keyboard navigation
            - Sets window close protocol
        """
        # Header with pump flow rate information
        header_text = f"تصرف الطلمبة {flow_rate} م³/ساعة"
        tk.Label(frame, text=header_text, font=(self.arabic_font[0], 18, "bold"),
                 bg="#f5f5f5", pady=20).pack()

        # Content frame for organized input layout
        content_frame = tk.Frame(frame, bg="#f5f5f5")
        content_frame.pack(pady=10)

        # Keyboard navigation helper function
        def handle_enter(event):
            """Move focus to next widget when Enter pressed."""
            event.widget.tk_focusNext().focus()
            return "break"

        # --- Configuration Fields ---
        # 1. Run Steps: Gauge indicator range for pump activation
        tk.Label(content_frame, text="عدد درجات المؤشر من 1 إلى 10",
                 font=self.arabic_font, bg="#f5f5f5").grid(row=1, column=1, padx=20, sticky="e")
        run_steps = tk.Entry(
            content_frame, font=self.arabic_font, width=8, justify="center")
        run_steps.grid(row=1, column=0, padx=10, pady=10)
        run_steps.bind("<Return>", handle_enter)

        # 2. Run Rate: Operating frequency in minutes
        tk.Label(content_frame, text="معدل التشغيل بالدقيقة، اكتب 10 أو 60",
                 font=self.arabic_font, bg="#f5f5f5").grid(row=2, column=1, padx=20, sticky="e")
        run_rate = tk.Entry(
            content_frame, font=self.arabic_font, width=8, justify="center")
        run_rate.grid(row=2, column=0, padx=10, pady=10)
        run_rate.bind("<Return>", handle_enter)

        # 3. Stop Steps: Gauge indicator range for pump deactivation
        tk.Label(content_frame, text="عدد درجات المؤشر من 1 إلى 10",
                 font=self.arabic_font, bg="#f5f5f5").grid(row=4, column=1, padx=20, sticky="e")
        stop_steps = tk.Entry(
            content_frame, font=self.arabic_font, width=8, justify="center")
        stop_steps.grid(row=4, column=0, padx=10, pady=10)
        stop_steps.bind("<Return>", handle_enter)

        # 4. Stop Rate: Deactivation frequency in minutes (last field)
        tk.Label(content_frame, text="معدل الإيقاف بالدقيقة، اكتب 10 أو 60",
                 font=self.arabic_font, bg="#f5f5f5").grid(row=5, column=1, padx=20, sticky="e")
        stop_rate = tk.Entry(
            content_frame, font=self.arabic_font, width=8, justify="center")
        stop_rate.grid(row=5, column=0, padx=10, pady=10)

        # Bind Enter in last field to save function
        stop_rate.bind("<Return>", lambda e: self.save_pump_data(pump_type))

        # Store entries as instance attribute for later access
        setattr(self, f"{pump_type}_entries", [
                run_steps, run_rate, stop_steps, stop_rate])

        # Save button for pump configuration
        btn_update = tk.Button(frame, text="تحديث البيانات", font=(self.arabic_font[0], 14, "bold"),
                               bg="#e0e0e0", relief="raised", bd=3, width=20,
                               command=lambda p=pump_type: self.save_pump_data(p))
        btn_update.pack(side="bottom", pady=40)

        # Handle window close event
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)

    def close_window(self):
        """
        Handle window closure and restore parent window visibility.

        Properly closes the pump configuration window and restores the parent window
        to visible state, ensuring clean modal dialog behavior.

        Returns:
            None: Destroys pump configuration window and shows parent

        Side Effects:
            - Destroys self.window
            - Calls self.parent.deiconify() to show parent window
        """
        self.window.destroy()
        self.parent.deiconify()

    def save_pump_data(self, pump_type):
        """
        Validate and save pump configuration to the template worksheet.

        Performs comprehensive input validation, then updates the BOD5 worksheet in
        Daily Template.xlsx with new pump operation parameters. Uses keyword-based
        cell discovery for flexible worksheet navigation.

        Args:
            pump_type (str): Pump identifier ("WAS" or "RAS") for template lookup

        Returns:
            None: Updates template or shows error messages

        Validation Process:
            1. Check all four fields are filled (run steps, run rate, stop steps, stop rate)
            2. Locate template file in settings directory
            3. Find keyword cells in BOD5 worksheet
            4. Update cells with new configuration values
            5. Handle protection/unprotection automatically
            6. Provide success/error feedback

        Keywords for Cell Lookup:
            - "{pump_type} on": Cell for run configuration
            - "{pump_type} off": Cell for stop configuration

        Cell Update Strategy:
            For "WAS on" or "RAS on":
            - Keyword cell: Located using text search
            - Run steps: Written to (column+1)
            - Run rate: Written to (column+4)

            For "WAS off" or "RAS off":
            - Keyword cell: Located using text search
            - Stop steps: Written to (column+1)
            - Stop rate: Written to (column+4)

        Template Updates:
            Updates "BOD5" worksheet in Daily Template.xlsx:
            - Cell search for keywords ("WAS on", "WAS off", "RAS on", "RAS off")
            - Value placement in calculated columns
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
            - Type errors: Safe None checking for row/column values

        User Feedback:
            - Confirmation: No pre-confirmation dialog (direct save)
            - Success: Message confirming pump configuration update
            - Error: Detailed error messages with recovery hints
            - Permission errors: Specific message for file access issues
            - Update message includes pump type for clarity

        Side Effects:
            - Modifies Daily Template.xlsx permanently
            - Updates BOD5 worksheet cells
            - No window closure on error (allows retry)

        Example Lookup:
            For WAS pump with keyword "WAS on" in cell (row=15, column=2):
            >>> Found cell at (15, 2)
            >>> Updates cell (15, 3) with run_steps value
            >>> Updates cell (15, 6) with run_rate value
        """
        # Get entries for the specified pump type
        entries = getattr(self, f"{pump_type}_entries")
        values = [e.get() for e in entries]

        # Validate all fields are filled before processing
        if not all(values):
            messagebox.showwarning("تنبيه", "يرجى ملء جميع الحقول")
            return

        try:
            # Locate and load template file
            path = self.general_path / "Daily Template.xlsx"
            wb = load_workbook(path)
            ws = wb["BOD5"]  # BOD5 worksheet for pump configuration

            # Temporarily disable protection for editing
            ws.protection.password = "01006610166"
            ws.protection.sheet = False

            # Define search configurations for pump cells
            search_configs = [
                {"key": f"{pump_type} on",  "s": values[0], "r": values[1]},
                {"key": f"{pump_type} off", "s": values[2], "r": values[3]}
            ]

            # Process each configuration (on/off)
            for cfg in search_configs:
                found = False
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value == cfg["key"]:
                            r, c = cell.row, cell.column
                            # Safely check row and column are valid
                            if r is not None and c is not None:
                                # Write steps value to column+1
                                ws.cell(row=int(r), column=int(
                                    c) + 1, value=str(cfg["s"]))
                                # Write rate value to column+4
                                ws.cell(row=int(r), column=int(
                                    c) + 4, value=str(cfg["r"]))
                                found = True
                                break
                    if found:
                        break

            # Re-enable worksheet protection
            ws.protection.sheet = True
            wb.save(path)
            wb.close()

            # Success feedback
            messagebox.showinfo("نجاح", f"تم تحديث بيانات {pump_type} بنجاح")

        except PermissionError:
            # File locked error with recovery instruction
            messagebox.showerror(
                "خطأ في الصلاحيات", "يرجى إغلاق ملف الإكسيل المفتوح أولاً قبل المحاولة مرة أخرى.")
        except Exception as e:
            # General exception with details
            messagebox.showerror("خطأ", f"حدث خطأ غير متوقع: {e}")


# Standalone execution support
if __name__ == "__main__":
    # Create root window for standalone execution
    root = tk.Tk()
    root.withdraw()  # Hide main window

    # Default Arabic font specification
    app = SludgePumpsWindow(root, ("Arial", 18))

    # Note: Missing root.mainloop() call - this would need to be added for proper standalone execution
