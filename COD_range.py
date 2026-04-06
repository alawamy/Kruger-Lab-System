# -*- coding: utf-8 -*-
"""
Chemical Oxygen Demand (COD) Potassium Dichromate Preparation Module - Laboratory Calculator

This module provides a comprehensive graphical interface for calculating and guiding the
preparation of standardized potassium dichromate (K2Cr2O7) solutions for Chemical Oxygen
Demand (COD) analysis in wastewater treatment laboratory quality control.

COD Analysis Purpose:
    COD (Chemical Oxygen Demand) measures the amount of oxygen required to oxidize organic
    matter in wastewater using potassium dichromate as the oxidizing agent. Three standard
    ranges are maintained for different wastewater strength conditions:
        1. 0-90 mg/L: Dilute or highly treated effluent (normality 0.025)
        2. 0-400 mg/L: Standard medium strength wastewater (normality 0.1)
        3. 100-900 mg/L: Strong or industrial wastewater (normality 0.25)

Solution Preparation:
    The module calculates the exact mass of K2Cr2O7 needed for a specific preparation volume
    and provides detailed step-by-step instructions for standardized solution preparation:
    1. Precise mass calculation based on selected normality and final volume
    2. Mercury sulfate (HgSO4) addition for chloride complexation
    3. Sulfuric acid (H2SO4) addition for acidic environment
    4. Potassium dichromate dissolution and volume adjustment
    5. Temperature equilibration for accuracy
    6. Amber bottle storage for light protection

Formula Calculation:
    Mass (grams) = Normality × (Volume in ml ÷ 1000) × 49.03 (K2Cr2O7 molar mass factor)
    
    Configuration:
    - 0-90 mg/L range: N = 0.025, Produces ~1.225g per 1000ml
    - 0-400 mg/L range: N = 0.1, Produces ~4.903g per 1000ml
    - 100-900 mg/L range: N = 0.25, Produces ~12.258g per 1000ml

Preparation Components:
    - Mercury sulfate (HgSO4): 33.3 g (fixed amount)
    - Sulfuric acid (H2SO4): 167 ml concentrated (fixed amount)
    - Potassium dichromate (K2Cr2O7): Calculated based on normality and volume
    - Distilled water: Used for initial and final volume adjustment

User Interface:
    - Range selector with three predefined normality levels
    - Volume input field (default 1000 ml)
    - Real-time mass calculation display
    - Detailed step-by-step preparation instructions
    - Automatic focus management and navigation
    - Clear visual feedback and guidance

Features:
    - Interactive range selection dropdown
    - Dynamic normality calculation based on range
    - Real-time weight calculation with 4 decimal places
    - Step-by-step Arabic preparation instructions
    - Keyboard navigation with Enter key support
    - Proper text processing for Arabic display
    - Professional color scheme and layout

Error Handling:
    - Input validation for numeric volume values
    - Try-except for calculation errors
    - Graceful handling of invalid inputs
    - Clear user feedback on calculation results

Dependencies:
    - tkinter: GUI framework with ttk for combobox widgets
    - config: ar_func utility for Arabic text processing and display
    - Python math: Built-in calculations for normality conversion

Author: Kruger Lab System Development Team
Version: 2.0
Last Updated: 2026-04-06
"""

import tkinter as tk
from tkinter import ttk
from config import ar_func


class PotassiumPrepWindow:
    """
    COD Potassium Dichromate Solution Preparation Calculator - Laboratory Guide Interface.

    This class implements a comprehensive graphical interface for calculating and guiding
    the preparation of standardized potassium dichromate solutions for COD analysis. The
    interface provides interactive range selection, real-time mass calculation, and
    detailed step-by-step preparation instructions.

    COD Range Options:
        - 0-90 mg/L: Dilute effluent samples (normality 0.025)
        - 0-400 mg/L: Standard wastewater (normality 0.1)
        - 100-900 mg/L: Strong/industrial wastewater (normality 0.25)

    Solution Calculations:
        Mass (g) = Normality × (Volume ÷ 1000) × 49.03

    Attributes:
        parent (tk.Tk): Parent window for modal behavior
        arabic_font (tuple): Font specification for Arabic text
        window (tk.Toplevel): Main preparation guide window
        range_var (tk.StringVar): Selected COD range (combobox variable)
        range_combo (ttk.Combobox): Range selection dropdown
        vol_entry (tk.Entry): Final volume input field (ml)
        weight_label (tk.Label): Calculated K2Cr2O7 mass display
        norm_label (tk.Label): Normality information display
        steps_label (tk.Label): Step-by-step preparation instructions

    Methods:
        close_window() → None: Handle window closure and parent restoration
        fix_text(text: str) → str: Process Arabic text for proper display
        on_range_select(event) → None: Handle range selection and focus management
        calculate_potassium() → None: Perform mass calculation and update instructions
    """

    def __init__(self, parent, arabic_font):
        """
        Initialize the COD potassium dichromate preparation calculator window.

        Creates a modal window with interactive range selection, volume input, real-time
        calculation display, and detailed preparation instructions for standardized
        potassium dichromate solution preparation.

        Args:
            parent (tk.Tk): Parent window for modal ownership and restoration
            arabic_font (tuple): Font specification tuple for Arabic text rendering

        Returns:
            None: Creates and displays the preparation guide interface

        Window Configuration:
            - Modal toplevel window (650x750 pixels)
            - Light background (#fdfdfd)
            - Arabic title with proper RTL support
            - Comprehensive event handling and navigation

        UI Components:
            Section 1 - Range Selection:
            - Label: "اختر نطاق تجربة المطلوب" (Select required test range)
            - Combobox with three options:
                * "0 - 90 mg/L" (normality 0.025)
                * "0 - 400 mg/L" (normality 0.1)
                * "100 - 900 mg/L" (normality 0.25)

            Section 2 - Volume Input:
            - Label: "قم بتسجيل الحجم النهائي المطلوب (ml)"
            - Entry field with default value 1000 ml
            - Enter key binding for calculation trigger

            Section 3 - Results Display:
            - Weight result in large bold font
            - Normality information label
            - Results frame with labeled border

            Section 4 - Instructions:
            - Step-by-step preparation guide
            - Dynamic text based on calculated mass
            - Arabic formatted text

            Footer:
            - Return button to close and restore parent window
        """
        self.parent = parent  # Reference to parent window
        self.arabic_font = arabic_font

        # Create main window
        self.window = tk.Toplevel(parent)
        self.window.title("COD تحضير بوتاسيوم داي كرومات")
        self.window.geometry("650x750")
        self.window.configure(bg="#fdfdfd")

        # Section 1: COD Range Selection
        tk.Label(self.window, text=self.fix_text("1-  اختر نطاق تجربة المطلوب:"),
                 font=(self.arabic_font[0], 14, "bold"), bg="#fdfdfd").pack(pady=15)

        # Range selector combobox
        self.range_var = tk.StringVar()
        self.range_combo = ttk.Combobox(self.window, textvariable=self.range_var,
                                        values=[
                                            "0 - 90 mg/L", "0 - 400 mg/L", "100 - 900 mg/L"],
                                        state="readonly", font=("Arial", 12), width=25)
        self.range_combo.pack(pady=5)
        self.range_combo.current(0)  # Default to first range

        # Bind range selection to calculation and focus management
        self.range_combo.bind("<<ComboboxSelected>>", self.on_range_select)

        # Section 2: Final Volume Input
        tk.Label(self.window, text=self.fix_text("2- قم بتسجيل الحجم النهائي المطلوب (ml):"),
                 font=(self.arabic_font[0], 14, "bold"), bg="#fdfdfd").pack(pady=15)
        self.vol_entry = tk.Entry(self.window, font=(
            "Arial", 12), width=15, justify="center")
        self.vol_entry.pack()
        self.vol_entry.insert(0, "1000")  # Default 1000 ml

        # Bind Enter key to trigger calculation
        self.vol_entry.bind("<Return>", lambda e: self.calculate_potassium())

        # Section 3: Results Display Frame
        results_frame = tk.LabelFrame(self.window, text=self.fix_text(" الوزن المطلوب من K2Cr2O7 "),
                                      font=self.arabic_font, bg="#fdfdfd", fg="#2c3e50")
        results_frame.pack(pady=30, padx=20, fill="x")

        # Mass calculation result in large bold font
        self.weight_label = tk.Label(results_frame, text="0.0000 g", font=("Arial", 24, "bold"),
                                     fg="#c0392b", bg="#fdfdfd")
        self.weight_label.pack(pady=10)

        # Normality information
        self.norm_label = tk.Label(
            results_frame, text="Normality: ---", font=("Arial", 12), bg="#fdfdfd")
        self.norm_label.pack(pady=5)

        # Section 4: Step-by-Step Instructions
        self.steps_label = tk.Label(self.window, text="", font=(self.arabic_font[0], 14),
                                    bg="#fdfdfd", justify="right")
        self.steps_label.pack(pady=20, padx=20)

        # Perform initial calculation on window open
        self.calculate_potassium()

        # Footer with return button
        footer = tk.Frame(self.window, bg="#fdfdfd")
        footer.pack(side="bottom", pady=30)
        tk.Button(footer, text=self.fix_text("رجوع"), font=self.arabic_font, bg="#e74c3c",
                  fg="white", width=12, command=self.close_window).pack()

        # Set initial focus
        self.range_combo.focus()

        # Handle window close event
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)

    def close_window(self):
        """
        Handle window closure and restore parent window visibility.

        Properly closes the preparation guide window and restores the parent window
        to visible state, ensuring clean modal dialog behavior.

        Returns:
            None: Destroys preparation window and shows parent

        Side Effects:
            - Destroys self.window
            - Calls self.parent.deiconify() to show parent window
        """
        self.window.destroy()  # Close current window
        self.parent.deiconify()  # Restore parent window visibility

    @staticmethod
    def fix_text(text):
        """
        Process and reshape Arabic text for proper display.

        Converts Arabic text to properly reshaped form for display in tkinter widgets,
        ensuring correct rendering of connected characters and RTL layout.

        Args:
            text (str): Original Arabic text string

        Returns:
            str: Reshaped text ready for tkinter display

        Processing:
            - Applies ar_func() from config module
            - Handles Arabic character reshaping
            - Ensures proper text flow and display
        """
        reshaped = ar_func(text)
        return reshaped

    def on_range_select(self, event):
        """
        Handle range selection change and manage focus flow.

        When user selects a COD range, triggers calculation update and moves focus
        to the volume input field for efficient workflow.

        Args:
            event (tk.Event): Combobox selection event

        Returns:
            None: Updates calculation and manages focus

        Side Effects:
            - Calls calculate_potassium() to update results
            - Moves focus to vol_entry
            - Selects existing text in vol_entry for easy replacement
        """
        # Recalculate with new range normality
        self.calculate_potassium()

        # Move focus to volume entry for next input
        self.vol_entry.focus()
        # Select existing text for easy replacement
        self.vol_entry.selection_range(0, tk.END)

    def calculate_potassium(self):
        """
        Calculate potassium dichromate mass and generate preparation instructions.

        Performs real-time calculation of the required K2Cr2O7 mass based on selected
        normality and desired final volume, then generates detailed step-by-step
        preparation instructions.

        Calculation Formula:
            Mass (grams) = Normality × (Volume in ml ÷ 1000) × 49.03

            Where 49.03 is the molar mass factor for K2Cr2O7

        Normality by Range:
            - 0-90 mg/L: Normality = 0.025
            - 0-400 mg/L: Normality = 0.1
            - 100-900 mg/L: Normality = 0.25

        Preparation Steps Generated:
            1. Initial water volume placement (half final volume)
            2. Mercury sulfate addition and dissolution
            3. Sulfuric acid addition with safety procedures
            4. Continued mixing until complete dissolution
            5. K2Cr2O7 weighing and dissolution
            6. Final volume adjustment to target
            7. Temperature equilibration
            8. Amber bottle storage

        Error Handling:
            - Try-except for invalid numeric input
            - Graceful handling of non-numeric volume values
            - Continues without error message on invalid input

        Side Effects:
            - Updates weight_label with calculated mass (4 decimal places)
            - Updates norm_label with selected normality
            - Updates steps_label with formatted preparation instructions
            - All text properly formatted for Arabic display

        Example Calculation:
            Range: 0-400 mg/L (N=0.1)
            Volume: 1000 ml
            Mass = 0.1 × (1000/1000) × 49.03 = 4.903 g
        """
        try:
            # Parse volume input
            v_ml = float(self.vol_entry.get())
            selected_range = self.range_var.get()

            # Determine normality based on selected COD range
            if selected_range == "0 - 90 mg/L":
                n = 0.025
            elif selected_range == "0 - 400 mg/L":
                n = 0.1
            else:  # 100 - 900 mg/L
                n = 0.25

            # Calculate mass using formula: N × (V/1000) × 49.03
            weight = n * (v_ml / 1000) * 49.03

            # Update weight display with 4 decimal places
            self.weight_label.config(text=f"{weight:.4f} g")
            # Update normality display
            self.norm_label.config(text=f"Normality: {n}")

            # Generate step-by-step preparation instructions
            steps = (
                f"1- ضع {int(v_ml / 2)} مل (نصف الحجم) من الماء المقطر على القلاب المغناطيسي.\n"
                f"2- زن 33.3 جم من كبريتات الزئبق (HgSO4) وأضفها إلى الماء في القلاب.\n"
                f"3- ضع 167 مل من حمض الكبريتيك المركز (H2SO4) ببطء مع التحريك المستمر.\n"
                f"4- استمر في التحريك حتى يذوب HgSO4 تمامًا.\n"
                f"5- زن {weight:.4f} جرام من داي كرومات البوتاسيوم المجفف مسبقاً عند 103°C،\n"
                f"   ثم أذبها وأكمل الحجم النهائي حتى {int(v_ml)} مل بالماء المقطر.\n"
                f"6- اترك المحلول ليبرد إلى درجة حرارة الغرفة قبل الاستخدام.\n"
                f"7- ضع المحلول في زجاجة بنية داكنة (Amber Bottle)."
            )

            # Update instructions with formatted Arabic text
            self.steps_label.config(text=self.fix_text(steps))

        except ValueError:
            # Silently handle non-numeric input (no error message)
            pass


# Standalone execution support
if __name__ == "__main__":
    # Create root window for standalone execution
    root = tk.Tk()
    root.withdraw()  # Hide main window

    # Default Arabic font specification
    arabic_font = ("Arial", 16)

    # Launch COD preparation calculator
    PotassiumPrepWindow(root, arabic_font)

    # Note: Missing root.mainloop() call - this would need to be added for proper standalone execution
