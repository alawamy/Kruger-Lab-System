# -*- coding: utf-8 -*-
"""
External Plant Report Generator - Bilingual Plant Data Management

This module provides comprehensive generation of external plant analysis reports
with bilingual plant name management, history tracking, and automated report
customization. The module supports both English and Arabic plant names with
persistent history storage for efficient data entry.

The module implements a sophisticated workflow for:
    1. Interactive bilingual plant name input with history suggestions
    2. Template-based report creation with plant-specific customization
    3. Automatic date stamping and bilingual header generation
    4. Worksheet protection with password-based security
    5. File organization by year and month structure

Plant Name Management:
    - Bilingual input: English and Arabic plant names
    - History tracking: Persistent storage of previously entered plants
    - Auto-completion: Combobox with existing plant names
    - Validation: Ensures both English and Arabic names are provided
    - Title case formatting: Automatic capitalization of English names

Report Characteristics:
    - Based on "External Report Template.xlsx" template
    - Bilingual headers in cells A4 and I4
    - Date stamping in cells B6 and G6 (dd/mm/yyyy format)
    - Protected with master password '01006610166'
    - Single worksheet ("Main") with external plant parameters

File Organization:
    External Plants/
    ├── 2026/
    │   ├── January/
    │   │   ├── 15-01-2026 PlantName.xlsx
    │   │   ├── 16-01-2026 AnotherPlant.xlsx
    │   │   └── ...
    │   ├── February/
    │   │   └── ...
    │   └── ...
    └── Templates/
        ├── External Report Template.xlsx
        └── plants_history.txt (auto-generated)

History File Format:
    plants_history.txt contains comma-separated entries:
    English Plant Name,Arabic Plant Name
    Example:
    Cairo West,القاهرة الغربية
    Alexandria North,الإسكندرية الشمالية

Template Processing:
    - Copies "External Report Template.xlsx" from SETTING_DIR
    - Updates bilingual headers with plant names
    - Sets report dates in multiple locations
    - Disables protection temporarily for editing
    - Re-enables protection with hardcoded password
    - Saves customized report with plant-specific filename

User Interface:
    - PlantInputDialog: Modal dialog for bilingual name entry
    - English combobox with history auto-completion
    - Arabic text entry with right-to-left support
    - Keyboard navigation between fields
    - Validation before report creation

Error Handling:
    - Template file missing: User notification with full path
    - History file access: Graceful fallback to empty history
    - File access conflicts: Option to overwrite or open existing
    - Excel processing errors: Exception message display
    - Directory creation failures: Automatic path creation

Dependencies:
    - openpyxl: Excel file manipulation and worksheet protection
    - tkinter: GUI components, dialogs, and combobox
    - config: EXTERNAL_DIR, SETTING_DIR path constants
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
from config import EXTERNAL_DIR, SETTING_DIR


class PlantInputDialog(tk.Toplevel):
    """
    Bilingual Plant Name Input Dialog - Interactive Plant Data Entry.

    This modal dialog provides a user-friendly interface for entering bilingual
    plant names (English and Arabic) with history tracking and auto-completion.
    The dialog maintains a persistent history of previously entered plants for
    efficient data entry and consistency.

    Features:
        - English plant name combobox with history suggestions
        - Arabic plant name text entry with RTL support
        - Keyboard navigation between fields
        - Automatic history saving for future use
        - Input validation ensuring both names are provided

    Attributes:
        history_file (Path): File path for storing plant name history
        history_data (dict): In-memory cache of plant name mappings
        result (tuple): (english_name, arabic_name) or (None, None) if cancelled

    Methods:
        load_history() → dict: Load plant name history from file
        save_to_history(eng, ar) → None: Save new plant name to history
        on_select_eng(event) → None: Handle English name selection
        on_enter_eng(event) → None: Handle Enter key in English field
        on_ok() → None: Validate and accept input
    """

    def __init__(self, parent, history_file):
        """
        Initialize the plant input dialog.

        Creates a modal dialog window with bilingual input fields and loads
        the plant name history for auto-completion suggestions.

        Args:
            parent (tk.Widget): Parent widget for dialog ownership
            history_file (Path): Path to the plant history file

        Returns:
            None: Initializes dialog and loads history data

        Side Effects:
            - Creates modal dialog window
            - Loads plant history into memory
            - Sets up keyboard bindings and event handlers
        """
        super().__init__(parent)
        self.title("Plant Entry")
        self.history_file = history_file
        self.history_data = self.load_history()
        self.result = (None, None)

        # Configure dialog appearance
        self.geometry("350x220")
        self.resizable(False, False)
        self.attributes('-topmost', True)

        # English plant name input (combobox with history)
        tk.Label(self, text="Plant Name (English):").pack(pady=(10, 0))
        self.eng_entry = ttk.Combobox(self, width=35)
        self.eng_entry['values'] = list(self.history_data.keys())
        self.eng_entry.pack(pady=5)
        self.eng_entry.focus_set()

        # Arabic plant name input (RTL text entry)
        tk.Label(self, text="اسم المحطة (بالعربية):").pack(pady=(10, 0))
        self.ar_entry = tk.Entry(self, width=38, justify='right')
        self.ar_entry.pack(pady=5)

        # Set up event bindings for keyboard navigation
        self.eng_entry.bind("<<ComboboxSelected>>", self.on_select_eng)
        self.eng_entry.bind('<Return>', self.on_enter_eng)
        self.ar_entry.bind('<Return>', lambda e: self.on_ok())

        # Button frame with OK/Cancel buttons
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="OK", width=10, bg="#e1e1e1",
                  command=self.on_ok).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Cancel", width=10,
                  command=self.destroy).pack(side="left", padx=5)

    def load_history(self):
        """
        Load plant name history from persistent storage file.

        Reads the plant history file and parses comma-separated English-Arabic
        name pairs into a dictionary for auto-completion and validation.

        Returns:
            dict: Mapping of English plant names to Arabic equivalents
                 Empty dict if file doesn't exist or parsing fails

        File Format:
            Each line: "English Name,Arabic Name"
            Example: "Cairo West,القاهرة الغربية"

        Error Handling:
            - File not found: Returns empty dictionary
            - Malformed lines: Skips invalid entries
            - Encoding errors: Graceful fallback
        """
        history = {}
        if self.history_file.exists():
            try:
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        if ',' in line:
                            eng, ar = line.strip().split(',', 1)
                            history[eng.strip()] = ar.strip()
            except Exception:
                # Return empty history on any error
                pass
        return history

    def save_to_history(self, eng, ar):
        """
        Save new plant name pair to persistent history storage.

        Adds the plant name mapping to the in-memory cache and appends it to
        the history file if it's not already present or has changed.

        Args:
            eng (str): English plant name (will be title-cased)
            ar (str): Arabic plant name

        Returns:
            None: Updates history file and in-memory cache

        Side Effects:
            - Updates self.history_data dictionary
            - Appends to history file if new or changed
            - Uses UTF-8 encoding for international character support
        """
        # Only save if new or different from existing
        if eng not in self.history_data or self.history_data[eng] != ar:
            self.history_data[eng] = ar
            with open(self.history_file, 'a', encoding='utf-8') as f:
                f.write(f"{eng},{ar}\n")

    def on_select_eng(self, event):
        """
        Handle selection of English plant name from combobox dropdown.

        When user selects a plant name from the history dropdown, automatically
        fills the corresponding Arabic name and moves focus to Arabic field.

        Args:
            event: Tkinter event object (unused but required for binding)

        Returns:
            None: Updates Arabic field and focus

        Side Effects:
            - Clears and fills Arabic text entry
            - Moves keyboard focus to Arabic field
        """
        eng_val = self.eng_entry.get().strip()
        if eng_val in self.history_data:
            self.ar_entry.delete(0, tk.END)
            self.ar_entry.insert(0, self.history_data[eng_val])
            self.ar_entry.focus_set()

    def on_enter_eng(self, event):
        """
        Handle Enter key press in English plant name field.

        Processes the entered English name: if it exists in history, fills Arabic
        field and moves focus; if new, moves focus to Arabic field for manual entry.

        Args:
            event: Tkinter event object (unused but required for binding)

        Returns:
            None: Updates UI focus based on name existence

        Behavior:
            - Existing name: Auto-fill Arabic and focus Arabic field
            - New name: Focus Arabic field for manual entry
        """
        eng_val = self.eng_entry.get().strip()
        if eng_val in self.history_data:
            # Existing plant: fill Arabic and focus it
            self.ar_entry.delete(0, tk.END)
            self.ar_entry.insert(0, self.history_data[eng_val])
            self.ar_entry.focus_set()
        else:
            # New plant: focus Arabic field for entry
            self.ar_entry.focus_set()

    def on_ok(self):
        """
        Validate input and accept the plant name entry.

        Checks that both English and Arabic names are provided, saves to history
        if valid, and closes the dialog with success result.

        Returns:
            None: Closes dialog or shows validation error

        Validation:
            - Both fields must be non-empty after stripping whitespace
            - English name is title-cased before saving
            - Shows warning dialog if validation fails

        Side Effects:
            - Saves plant names to history file
            - Sets self.result with validated names
            - Destroys dialog window on success
        """
        # Get current field values
        eng = self.eng_entry.get().strip()
        ar = self.ar_entry.get().strip()

        if eng and ar:
            # Valid input: save and close
            self.save_to_history(eng.title(), ar)
            self.result = (eng.title(), ar)
            self.destroy()
        else:
            # Invalid input: show warning
            messagebox.showwarning(
                "Input Error", "Please fill both fields.", parent=self)


def get_plant_gui(history_path):
    """
    Launch plant input dialog and return the result.

    Creates a temporary root window and modal PlantInputDialog to collect
    bilingual plant name input from the user.

    Args:
        history_path (Path): Path to the plant history file

    Returns:
        tuple: (english_name, arabic_name) or (None, None) if cancelled

    Side Effects:
        - Creates and destroys temporary tkinter root window
        - Blocks until user completes or cancels dialog
    """
    root = tk.Tk()
    root.withdraw()
    dialog = PlantInputDialog(root, history_path)
    root.wait_window(dialog)
    res = dialog.result
    root.destroy()
    return res


def external_report(main_root=None):
    """
    Generate external plant analysis reports with bilingual customization.

    This function provides a complete workflow for creating external plant analysis
    reports, including interactive plant name input, template validation, and
    automatic report customization with bilingual headers and date stamping.

    Process Flow:
        1. Validate template and history file availability
        2. Display action selection dialog (Create/Open Folder/Cancel)
        3. Create destination directory structure if needed
        4. If Create chosen: Launch plant name input dialog
        5. Copy template and customize with plant names and dates
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
        - Checks for "External Report Template.xlsx" in SETTING_DIR
        - Checks for "plants_history.txt" in SETTING_DIR (created if missing)
        - Shows error message with full path if template missing

    Plant Name Input:
        - Launches PlantInputDialog for bilingual name entry
        - Supports history auto-completion for English names
        - Validates both English and Arabic names provided
        - Returns to main flow if user cancels

    File Creation Process:
        1. Generate filename: "dd-mm-yyyy PlantName.xlsx"
        2. Check for existing file conflicts (overwrite option)
        3. Copy template to destination
        4. Load workbook and access Main worksheet
        5. Temporarily disable protection for editing
        6. Update bilingual headers in cells A4 and I4
        7. Set dates in cells B6 and G6 (dd/mm/yyyy format)
        8. Re-enable protection with password '01006610166'
        9. Save and close workbook
        10. Show success message and open file

    Directory Structure:
        - Year folders: EXTERNAL_DIR / YYYY
        - Month folders: EXTERNAL_DIR / YYYY / MonthName
        - Files: dd-mm-yyyy PlantName.xlsx

    Error Conditions:
        - Template missing: Early return with error message
        - Plant input cancelled: Return without creating report
        - File exists: User choice between overwriting or opening existing
        - Excel errors: Exception details shown to user
        - Directory access: Handled by mkdir(parents=True, exist_ok=True)

    User Experience:
        - Modal dialogs prevent interaction with other windows
        - Clear action choices in confirmation dialog
        - Intuitive bilingual plant name entry
        - History-based auto-completion for efficiency
        - Automatic file opening for immediate use
        - Optional main application closure after completion

    Examples:
        >>> # Called from main application
        >>> external_report(main_window)  # May close main window after completion

        >>> # Standalone execution
        >>> if __name__ == "__main__":
        ...     external_report()  # Independent operation

    Integration Points:
        - app.py: Called from external plant report menu button
        - main_menu.py: Accessible through specialized operations
        - Template maintained in SETTING_DIR
        - History file auto-managed in SETTING_DIR
        - Reports organized under EXTERNAL_DIR structure
    """
    # Validate required files
    template_path = SETTING_DIR / 'External Report Template.xlsx'
    history_path = SETTING_DIR / 'plants_history.txt'

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
    dest_folder = EXTERNAL_DIR / now.strftime("%Y") / now.strftime("%B")
    dest_folder.mkdir(parents=True, exist_ok=True)

    if choice is None:  # User cancelled
        return

    if choice is False:  # Open folder option
        os.startfile(dest_folder)
        if main_root:
            main_root.destroy()  # Close main application window
        return None

    if choice is True:  # Create new report
        # Get plant name input from user
        plant_data = get_plant_gui(history_path)
        if not plant_data or not plant_data[0]:
            return

        plant_eng, plant_ar = plant_data

        # Generate filename with plant name
        new_filename = f"{now.strftime('%d-%m-%Y')} {plant_eng}.xlsx"
        dest_path = dest_folder / new_filename

        try:
            # Handle existing file conflict
            if dest_path.exists():
                if not messagebox.askyesno("Overwrite?", f"File exists. Overwrite?"):
                    os.startfile(dest_path)
                    return

            # Copy template to destination
            copyfile(template_path, dest_path)

            # Customize the report
            wb = openpyxl.load_workbook(dest_path)
            ws = wb['Main']

            # Temporarily disable protection for editing
            ws.protection.sheet = False

            # Update bilingual headers with plant names
            ws['A4'] = f"{plant_eng} Treatment Plant"
            ws['I4'] = f"محطة معالجة  {plant_ar}"

            # Set report dates in multiple locations
            ws['B6'] = ws['G6'] = now.strftime('%d/%m/%Y')

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
    # Execute external plant report generation independently
    external_report()
