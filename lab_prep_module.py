# -*- coding: utf-8 -*-
"""
Laboratory Materials Preparation Module - Chemical Solutions Management Interface

This module provides a central navigation interface for accessing various laboratory
chemical preparation procedures in the Kruger Lab System. It serves as a hub for
standardized chemical solution preparation including iodine, sulfuric acid, and
potassium dichromate solutions.

Laboratory Preparations:
    The module provides access to three main chemical preparation procedures:
        1. Iodine Solution: For titration and oxidation analysis
        2. Sulfuric Acid Reagent: For indicator preparation and acidic analysis
        3. Potassium Dichromate (K2Cr2O7): For COD (Chemical Oxygen Demand) analysis

User Interface:
    - Centralized preparation materials menu
    - Three main preparation option buttons
    - Professional button styling with visual feedback
    - Return to main menu navigation
    - Arabic text support for bilingual interface
    - Modal window management

Features:
    - Navigation hub for chemical preparations
    - Integration with specialized preparation modules
    - Parent window preservation for modal workflow
    - Professional layout and styling
    - Cursor feedback on button hover
    - Easy access to allthree preparation types

Menu Structure:
    - Iodine: Opens iodine_prep module
        * Interactive normality and volume calculator
        * Real-time mass calculations
        * Detailed preparation instructions

    - Sulfuric R: Opens sulfuric_prep module
        * Specific gravity and weight inputs
        * Volume and silver sulfate calculations
        * 2-3 day dissolution guidance

    - K2Cr2O7 (COD): Opens COD_range module
        * Three standard COD ranges
        * Automatic normality selection
        * Dynamic calculation based on volume

Dependencies:
    - tkinter: GUI framework with window management
    - config: ar_func utility for Arabic text processing
    - COD_range: PotassiumPrepWindow for potassium dichromate preparation
    - iodine_prep: IodinePrepWindow for iodine preparation
    - sulfuric_prep: SulfuricPrepWindow for sulfuric acid reagent preparation
    - main_menu: Available for navigation (imported but not used in this module)

Integration Points:
    - Chemical Center Menu (chemical_module.py): Parent interface
    - All three specialized preparation windows
    - Modal dialog workflow with window withdrawal/restoration

Author: Kruger Lab System Development Team
Version: 2.0
Last Updated: 2026-04-06
"""

from re import L
import tkinter as tk
from config import ar_func
from COD_range import PotassiumPrepWindow
from iodine_prep import IodinePrepWindow
from sulfuric_prep import SulfuricPrepWindow


class LabMaterialsWindow:
    """
    Laboratory Materials Preparation Navigation Interface - Chemical Solutions Hub.

    This class implements a central navigation menu for accessing three main laboratory
    chemical preparation procedures: iodine solutions, sulfuric acid reagents, and
    potassium dichromate (COD) solutions.

    Preparation Options:
        - Iodine: Normality-based solution for titration procedures
        - Sulfuric R: Reagent grade for acidic analysis and indicators
        - K2Cr2O7 (COD): Three-range potassium dichromate for oxygen demand testing

    Attributes:
        arabic_font (tuple): Font specification for Arabic text
        parent (tk.Tk): Parent window for modal ownership and restoration
        window (tk.Toplevel): Main materials preparation menu window
        btn_iodine (tk.Button): Button to open iodine preparation
        btn_sulfuric (tk.Button): Button to open sulfuric acid reagent preparation
        btn_cod (tk.Button): Button to open COD potassium dichromate preparation

    Methods:
        grid_btn_full(btn, row, col, colspan) → None: Place button across columns
        fix_text(text: str, is_ui: bool) → str: Process Arabic text for display
        open_cod() → None: Launch potassium dichromate preparation
        open_iodine() → None: Launch iodine preparation
        open_sulfuric() → None: Launch sulfuric acid reagent preparation
        close_and_back() → None: Close menu and restore parent window
    """

    def __init__(self, parent, arabic_font):
        """
        Initialize the laboratory materials preparation menu window.

        Creates a modal window with navigation buttons for three chemical preparation
        procedures, organized with professional styling and Arabic support.

        Args:
            parent (tk.Tk): Parent window for modal ownership and restoration
            arabic_font (tuple): Font specification tuple for Arabic text rendering

        Returns:
            None: Creates and displays the materials preparation menu interface

        Window Configuration:
            - Modal toplevel window (600x550 pixels)
            - Light green-gray background (#f4f7f6)
            - Arabic title with proper RTL support
            - Three button grid layout plus return button

        UI Components:
            - Title: "قائمة المواد المتوفرة للتحضير" (Available materials for preparation)
            - Button Grid:
                * Row 0: Iodine (left) and Sulfuric R (right)
                * Row 1: K2Cr2O7 (COD) - spans both columns
            - Button Style: White background, raised relief, hand cursor
            - Return Button: Red background, bottom placement
        """
        self.arabic_font = arabic_font
        self.parent = parent  # Reference to parent window
        self.window = tk.Toplevel(parent)
        self.window.title(self.fix_text(
            "التحضيرات المعملية القياسية", is_ui=False))
        self.window.geometry("600x550")
        self.window.configure(bg="#f4f7f6")

        # Title label
        tk.Label(self.window, text=self.fix_text("قائمة المواد المتوفرة للتحضير"),
                 font=(self.arabic_font[0], 18, "bold"), bg="#f4f7f6", fg="#2c3e50").pack(pady=20)

        # Button frame for organized layout
        btns_frame = tk.Frame(self.window, bg="#f4f7f6")
        btns_frame.pack(pady=10)

        # Button styling configuration
        btn_style = {
            "font": ("Arial", 12, "bold"),
            "width": 20,
            "height": 2,
            "bg": "#ffffff",
            "bd": 3,
            "relief": "raised",
            "cursor": "hand2"
        }

        # Iodine preparation button
        self.btn_iodine = tk.Button(btns_frame, text="Iodine", **btn_style,
                                    command=self.open_iodine)
        self.btn_iodine.grid(row=0, column=0, padx=15, pady=15)

        # Sulfuric acid reagent preparation button
        self.btn_sulfuric = tk.Button(btns_frame, text="Sulfuric R", **btn_style,
                                      command=self.open_sulfuric)
        self.btn_sulfuric.grid(row=0, column=1, padx=15, pady=15)

        # COD potassium dichromate preparation button (spans full width)
        self.btn_cod = tk.Button(btns_frame, text="K2Cr2O7 (COD)", **btn_style,
                                 command=self.open_cod)
        self.grid_btn_full(self.btn_cod, row=1, col=0, colspan=2)

        # Return to main menu button
        btn_back = tk.Button(self.window, text=self.fix_text("رجوع للقائمة الرئيسية"),
                             font=self.arabic_font, bg="#e74c3c", fg="white",
                             width=25, command=self.close_and_back)
        btn_back.pack(side="bottom", pady=40)

    @staticmethod
    def grid_btn_full(btn, row, col, colspan):
        """
        Place button in grid spanning multiple columns.

        Places a button widget in the grid layout with specified row, column,
        and column span for full-width buttons.

        Args:
            btn (tk.Button): Button widget to place
            row (int): Grid row position
            col (int): Grid column position
            colspan (int): Number of columns to span

        Returns:
            None: Places button in grid with configuration
        """
        btn.grid(row=row, column=col, columnspan=colspan, padx=15, pady=15)

    @staticmethod
    def fix_text(text: str, is_ui=True) -> str:
        """
        Process and reshape Arabic text for proper display.

        Converts Arabic text to properly reshaped form for display in tkinter widgets,
        ensuring correct rendering of connected characters and RTL layout.

        Args:
            text (str): Original Arabic text string
            is_ui (bool): Whether text is for UI elements (default True)

        Returns:
            str: Reshaped text ready for tkinter display
        """
        return ar_func(text, is_ui_element=is_ui)

    def open_cod(self):
        """
        Launch COD potassium dichromate preparation interface.

        Hides the current materials menu and opens the potassium dichromate
        preparation window for COD (Chemical Oxygen Demand) solution preparation.

        Returns:
            None: Displays potassium dichromate preparation interface

        Side Effects:
            - Withdraws current materials menu window
            - Creates PotassiumPrepWindow modal dialog
        """
        self.window.withdraw()  # Hide materials menu
        PotassiumPrepWindow(self.window, self.arabic_font)

    def open_iodine(self):
        """
        Launch iodine solution preparation interface.

        Hides the current materials menu and opens the iodine preparation window
        for normality-based iodine solution preparation.

        Returns:
            None: Displays iodine preparation interface

        Side Effects:
            - Withdraws current materials menu window
            - Creates IodinePrepWindow modal dialog
        """
        self.window.withdraw()  # Hide materials menu
        IodinePrepWindow(self.window, self.arabic_font)

    def open_sulfuric(self):
        """
        Launch sulfuric acid reagent preparation interface.

        Hides the current materials menu and opens the sulfuric acid reagent
        preparation window for silver sulfate indicator solution preparation.

        Returns:
            None: Displays sulfuric acid preparation interface

        Side Effects:
            - Withdraws current materials menu window
            - Creates SulfuricPrepWindow modal dialog
        """
        self.window.withdraw()  # Hide materials menu
        SulfuricPrepWindow(self.window, self.arabic_font)

    def close_and_back(self):
        """
        Close menu and restore parent window visibility.

        Properly closes the materials preparation menu and restores the parent window
        to visible state, ensuring clean modal dialog behavior in the workflow.

        Returns:
            None: Destroys materials menu and shows parent

        Side Effects:
            - Destroys self.window
            - Calls self.parent.deiconify() to show parent window if it exists
        """
        self.window.destroy()
        if self.parent:
            self.parent.deiconify()  # Restore parent window visibility


# Standalone execution support
if __name__ == "__main__":
    # Create root window for standalone execution
    root = tk.Tk()
    root.withdraw()  # Hide main window

    # Default Arabic font specification
    LabMaterialsWindow(root, ("Arial", 12, "bold"))

    # Note: Missing root.mainloop() call - this would need to be added for proper standalone execution
