# -*- coding: utf-8 -*-
"""
Chemical Center Management Module - Laboratory Chemical and Preparation Hub

This module serves as a central navigation hub for all chemical handling and laboratory
preparation procedures in the Kruger Lab System. It provides a user-friendly menu interface
for accessing chemical preparation procedures and intelligent chemical analysis tools.

Chemical Management System:
    The module organizes chemical operations into two main categories:
        1. Laboratory Material Preparations: Iodine and acid preparation procedures
        2. Smart Chemical Analyzer: Intelligent molecular weight and molarity calculations

Main Features:
    - Navigation hub for chemical operations
    - Integration with laboratory preparation module
    - Smart chemical analysis calculator
    - Modal window management with parent preservation
    - Arabic text support for bilingual interface
    - Professional menu layout with clear organization

Navigation Structure:
    Option 1: Laboratory Preparations (التحضيرات المعملية)
        - Iodine preparation procedures
        - Acid preparation procedures
        - Links to lab_prep_module.py

    Option 2: Smart Chemical Analyzer (المحلل الكيميائي الذكي)
        - Molecular weight calculations
        - Molarity and molality computations
        - Compound analysis tools
        - Links to calculate_compound_weight.py

User Interface:
    - Modal window with professional styling
    - Two main option buttons with descriptive labels
    - Return to main menu button
    - Arabic text processing for proper display
    - Color-coded buttons for intuitive navigation
    - Organized layout with proper spacing

Dependencies:
    - tkinter: GUI framework with window management
    - config: ar_func utility for Arabic text processing
    - lab_prep_module: Laboratory material preparation interface
    - calculate_compound_weight: Smart chemical analyzer module

Module Integration:
    - Imports lab_prep_module.LabMaterialsWindow on demand
    - Imports calculate_compound_weight.SmartChemicalAnalyzer on demand
    - Local imports within methods prevent circular dependencies
    - Proper window withdrawal/restoration for modal behavior

Author: Kruger Lab System Development Team
Version: 2.0
Last Updated: 2026-04-06
"""

import tkinter as tk
from config import ar_func


class ChemicalMenuWindow:
    """
    Chemical Center Management Navigation Interface - Main Chemical Operations Hub.

    This class implements the central navigation interface for all chemical handling
    and laboratory preparation operations in the Kruger Lab System. It provides access
    to laboratory material preparation procedures and intelligent chemical analysis tools.

    Menu Options:
        - Laboratory Preparations: Iodine and acid preparation procedures
        - Smart Chemical Analyzer: Molecular weight and molarity calculations

    Attributes:
        parent (tk.Tk): Parent window for modal behavior
        arabic_font (tuple): Font specification for Arabic text
        window (tk.Toplevel): Main chemical center menu window

    Methods:
        open_lab_prep() → None: Launch laboratory material preparation interface
        open_smart_analyzer() → None: Launch smart chemical analysis calculator
        close_window() → None: Handle window closure and parent restoration
    """

    def __init__(self, parent, arabic_font):
        """
        Initialize the chemical center management menu window.

        Creates a modal window with navigation buttons for chemical operations,
        including laboratory preparations and smart chemical analysis tools.

        Args:
            parent (tk.Tk): Parent window for modal ownership and restoration
            arabic_font (tuple): Font specification tuple for Arabic text rendering

        Returns:
            None: Creates and displays the chemical management menu interface

        Window Configuration:
            - Modal toplevel window (700x450 pixels)
            - Light gray background (#ecf0f1)
            - Arabic title with proper RTL support
            - Comprehensive event handling

        UI Components:
            - Title: "اختر القسم المطلوب" (Select required section)
            - Option 1: Laboratory Preparations button (green #27ae60)
            - Option 2: Smart Chemical Analyzer button (blue #2980b9)
            - Return button (gray #95a5a6)
            - Organized layout with proper spacing
        """
        self.parent = parent
        self.arabic_font = arabic_font

        # Create main chemical center window
        self.window = tk.Toplevel(parent)

        # Set window title with Arabic text processing
        self.window.title(
            ar_func("مركز إدارة الكيماويات والتحضيرات", is_ui_element=False))

        self.window.geometry("700x450")
        self.window.configure(bg="#ecf0f1")

        # Hide parent window during menu display
        self.parent.withdraw()

        # Arabic text processing function for UI elements
        def fix(text):
            """Process Arabic text for UI elements with proper reshaping."""
            return ar_func(text, is_ui_element=True)

        # Title label
        tk.Label(self.window, text=fix("اختر القسم المطلوب"),
                 font=(self.arabic_font[0], 20, "bold"), bg="#ecf0f1", fg="#2c3e50").pack(pady=30)

        # Option 1: Laboratory Material Preparations
        btn_lab = tk.Button(self.window, text=fix("1. التحضيرات المعملية (Iodine / Acids)"),
                            font=self.arabic_font, bg="#27ae60", fg="white", width=35, height=2,
                            command=self.open_lab_prep)
        btn_lab.pack(pady=15)

        # Option 2: Smart Chemical Analyzer
        btn_smart = tk.Button(self.window, text=fix("2. المحلل الكيميائي الذكي (M & N)"),
                              font=self.arabic_font, bg="#2980b9", fg="white", width=35, height=2,
                              command=self.open_smart_analyzer)
        btn_smart.pack(pady=15)

        # Return to main menu button
        tk.Button(self.window, text=fix("رجوع للقائمة الرئيسية"), font=self.arabic_font,
                  bg="#95a5a6", command=self.close_window).pack(pady=20)

        # Handle window close event
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)

    def open_lab_prep(self):
        """
        Launch laboratory material preparation interface.

        Imports and opens the laboratory material preparation window, which provides
        access to procedures for preparing iodine solutions and acidic reagents.

        Returns:
            None: Displays laboratory preparation interface

        Module Import:
            - Imports LabMaterialsWindow from lab_prep_module
            - Local import prevents circular dependencies
            - Hides chemical menu window during access

        Side Effects:
            - Withdraws current chemical menu window
            - Creates LabMaterialsWindow modal dialog
        """
        # Local import to prevent circular dependencies
        from lab_prep_module import LabMaterialsWindow

        # Hide chemical menu during laboratory preparation
        self.window.withdraw()

        # Launch laboratory material preparation interface
        LabMaterialsWindow(self.window, self.arabic_font)

    def open_smart_analyzer(self):
        """
        Launch smart chemical analysis calculator.

        Imports and opens the intelligent chemical analyzer interface, which provides
        tools for molecular weight calculations and molarity/concentration computations.

        Returns:
            None: Displays smart chemical analyzer interface

        Module Import:
            - Imports SmartChemicalAnalyzer from calculate_compound_weight
            - Local import prevents circular dependencies
            - Hides chemical menu window during access

        Features Accessed:
            - Molecular weight calculations
            - Molarity and concentration computations
            - Compound analysis tools

        Side Effects:
            - Withdraws current chemical menu window
            - Creates SmartChemicalAnalyzer modal dialog
        """
        # Local import for smart chemical analyzer module
        from calculate_compound_weight import SmartChemicalAnalyzer

        # Hide chemical menu during analyzer access
        self.window.withdraw()

        # Launch smart chemical analysis calculator
        SmartChemicalAnalyzer(self.window, self.arabic_font)

    def close_window(self):
        """
        Handle window closure and restore parent window visibility.

        Properly closes the chemical management menu window and restores the parent
        window to visible state, ensuring clean modal dialog behavior.

        Returns:
            None: Destroys chemical menu window and shows parent

        Side Effects:
            - Destroys self.window
            - Calls self.parent.deiconify() to show parent window
        """
        self.window.destroy()
        self.parent.deiconify()


# Standalone execution support
if __name__ == "__main__":
    # Create root window for standalone execution
    root = tk.Tk()
    root.withdraw()  # Hide main window

    # Default Arabic font specification
    ChemicalMenuWindow(root, ("Arial", 18, "bold"))

    # Note: Missing root.mainloop() call - this would need to be added for proper standalone execution
