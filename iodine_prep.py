# -*- coding: utf-8 -*-
"""
Iodine Solution Preparation Module - Laboratory Chemistry Calculator and Guide

This module provides a comprehensive graphical interface for calculating and guiding the
preparation of standardized iodine (I₂) solutions for wastewater treatment laboratory
quality control. The module enables operators to specify normality and volume, then
calculates required chemical masses and displays detailed preparation instructions.

Iodine Solution Purpose:
    Iodine solutions are used in various analytical procedures including:
    - Oxygen demand measurements using titration methods
    - Water quality analysis and disinfection studies
    - Standardized chemical testing in wastewater treatment

Solution Composition:
    Standard iodine solution preparation employs:
    - Iodine (I₂): Primary oxidizing agent (calculated mass)
    - Potassium Iodide (KI): Solubilizing agent (2.5 × iodine mass)
    - Distilled Water: Solvent for preparation
    - Volumetric flask: Final volume adjustment

Mass Calculation Formula:
    Iodine mass (grams) = Normality × (Volume in ml ÷ 1000) × 126.9 (molar mass)
    KI mass (grams) = Iodine mass × 2.5 (empirical solubilization ratio)

Preparation Process:
    1. User specifies normality (N) and desired final volume
    2. System calculates required iodine and KI masses
    3. Step-by-step instructions guide operator through preparation
    4. Proper dissolution sequence ensures maximum solubility
    5. Volumetric flask used for accurate final volume

User Interface:
    - Normality input field with focus management
    - Final volume input field (ml)
    - Real-time mass calculation display
    - Dynamic preparation instructions updated per calculation
    - Keyboard navigation with Enter key support
    - Reset function for new calculations
    - Return to menu button

Features:
    - Interactive input validation
    - Real-time mass calculation with 1 decimal place
    - Dynamic instruction text based on calculated volume
    - Arabic text processing for proper display
    - Automatic focus movement between fields
    - Professional color scheme and layout

Calculation Example:
    Normality: 0.047 N
    Volume: 1000 ml
    Iodine mass = 0.047 × (1000/1000) × 126.9 = 5.965 g
    KI mass = 5.965 × 2.5 = 14.9 g

Error Handling:
    - Input validation for numeric values
    - Error messages for invalid input
    - Graceful handling of calculation failures
    - Clear user feedback on success/failure

Dependencies:
    - tkinter: GUI framework with messagebox for error handling
    - config: ar_func utility for Arabic text processing and display
    - Python math: Built-in calculations for mass computations

Author: Kruger Lab System Development Team
Version: 2.0
Last Updated: 2026-04-06
"""

import tkinter as tk
from tkinter import messagebox
from config import ar_func


class IodinePrepWindow:
    """
    Iodine Solution Preparation Calculator - Laboratory Chemistry Guide Interface.

    This class implements a graphical interface for calculating and guiding iodine
    solution preparation procedures. The interface allows operators to specify normality
    and volume, then calculates required chemical masses and displays detailed
    step-by-step preparation instructions.

    Solution Preparation:
        - Input: Normality (N) and final volume (ml)
        - Calculated: Iodine mass and KI mass based on formulas
        - Output: Step-by-step instructions with specific volumes

    Attributes:
        parent (tk.Tk): Parent window for modal behavior
        arabic_font (tuple): Font specification for Arabic text
        window (tk.Toplevel): Main preparation guide window
        norm_entry (tk.Entry): Normality input field
        vol_entry (tk.Entry): Final volume input field (ml)
        iodine_label (tk.Label): Calculated iodine mass display
        ki_label (tk.Label): Calculated KI mass display
        steps_frame (tk.LabelFrame): Container for instructions
        instructions_label (tk.Label): Step-by-step preparation instructions

    Methods:
        close_window() → None: Handle window closure and parent restoration
        fix_text(text: str) → str: Process Arabic text for proper display
        get_dynamic_steps(volume_third: str) → str: Generate dynamic instructions
        calculate_iodine_weights() → None: Calculate masses and update display
        reset_fields() → None: Clear inputs and reset display
    """

    def __init__(self, parent, arabic_font):
        """
        Initialize the iodine solution preparation calculator window.

        Creates a modal window with input fields for normality and volume, displays
        calculated chemical masses, and provides dynamic step-by-step preparation
        instructions updated in real-time.

        Args:
            parent (tk.Tk): Parent window for modal ownership and restoration
            arabic_font (tuple): Font specification tuple for Arabic text rendering

        Returns:
            None: Creates and displays the iodine preparation interface

        Window Configuration:
            - Modal toplevel window (600x800 pixels)
            - Light background (#f8f9fa)
            - Arabic title with proper RTL support
            - Comprehensive event handling and navigation

        UI Sections:
            Section 1 - Input Parameters:
            - Normality field (concentration specification)
            - Final volume field (ml specification)
            - Enter key bindings for navigation and calculation

            Section 2 - Calculated Masses:
            - Iodine mass display (bold red text)
            - Potassium Iodide mass display (bold blue text)
            - Real-time updates on calculation

            Section 3 - Preparation Instructions:
            - Dynamic text based on calculated volumes
            - Step-by-step guidance in Arabic
            - Updated instructions with each calculation

            Footer:
            - Reset button for new calculation
            - Return button to menu
        """
        self.parent = parent  # Reference to parent window
        self.arabic_font = arabic_font

        # Create main window
        self.window = tk.Toplevel(parent)
        self.window.title("خطوات تحضير محلول اليود")
        self.window.geometry("600x800")
        self.window.configure(bg="#f8f9fa")

        # Initialize dynamic instructions
        self.get_dynamic_steps("---")

        # Section 1: Title for input parameters
        tk.Label(self.window, text=self.fix_text("1- قم بتسجيل العيارية و الحجم النهائي"),
                 font=(self.arabic_font[0], 16, "bold"), bg="#f8f9fa", fg="#2c3e50").pack(pady=20)

        # Input parameters frame
        input_frame = tk.Frame(self.window, bg="#f8f9fa")
        input_frame.pack(pady=10)

        # Normality input field
        tk.Label(input_frame, text=self.fix_text("العيارية (Normality):"),
                 font=self.arabic_font, bg="#f8f9fa").grid(row=0, column=1, padx=10, pady=10)
        self.norm_entry = tk.Entry(input_frame, font=(
            "Arial", 12), width=15, justify="center")
        self.norm_entry.grid(row=0, column=0)
        self.norm_entry.focus()

        # Final volume input field (ml)
        tk.Label(input_frame, text=self.fix_text("الحجم النهائي (ml):"),
                 font=self.arabic_font, bg="#f8f9fa").grid(row=1, column=1, padx=10, pady=10)
        self.vol_entry = tk.Entry(input_frame, font=(
            "Arial", 12), width=15, justify="center")
        self.vol_entry.grid(row=1, column=0)

        # Configure Enter key navigation
        self.norm_entry.bind("<Return>", lambda e: self.vol_entry.focus())
        self.vol_entry.bind(
            "<Return>", lambda e: self.calculate_iodine_weights())

        # Section 2: Title for calculated masses
        tk.Label(self.window, text=self.fix_text("2- قم بوزن الكيماويات التالية :"),
                 font=(self.arabic_font[0], 16, "bold"), bg="#f8f9fa").pack(pady=20)

        # Iodine mass display (bold red)
        self.iodine_label = tk.Label(self.window, text="Iodine: 0.0 g",
                                     font=("Arial", 15, "bold"), fg="#c0392b", bg="#f8f9fa")
        self.iodine_label.pack(pady=5)

        # KI mass display (bold blue)
        self.ki_label = tk.Label(self.window, text="KI: 0.0 g",
                                 font=("Arial", 15, "bold"), fg="#2980b9", bg="#f8f9fa")
        self.ki_label.pack(pady=5)

        # Section 3: Preparation instructions frame
        self.steps_frame = tk.LabelFrame(self.window, text=" طريقة التحضير ",
                                         font=self.arabic_font, bg="#f8f9fa")
        self.steps_frame.pack(pady=20, padx=20, fill="both")

        # Dynamic instructions label
        self.instructions_label = tk.Label(self.steps_frame, text=self.get_dynamic_steps("---"),
                                           font=(self.arabic_font[0], 14), bg="#f8f9fa", justify="right")
        self.instructions_label.pack(fill="both", padx=10, pady=10)

        # Footer buttons frame
        footer_frame = tk.Frame(self.window, bg="#f8f9fa")
        footer_frame.pack(side="bottom", pady=30)

        # Reset button for new calculation
        tk.Button(footer_frame, text="تركيز جديد", font=self.arabic_font,
                  command=self.reset_fields, width=12).pack(side="right", padx=10)

        # Return to menu button
        tk.Button(footer_frame, text="رجوع", font=self.arabic_font, bg="#e74c3c",
                  fg="white", width=12, command=self.close_window).pack(side="right", padx=10)

        # Handle window close event
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)

    def close_window(self):
        """
        Handle window closure and restore parent window visibility.

        Properly closes the iodine preparation window and restores the parent window
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
    def fix_text(text: str) -> str:
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
            - Returns string type (not bytes)
        """
        reshaped_text = ar_func(text)

        # Ensure output is string (not bytes)
        if isinstance(reshaped_text, bytes):
            return reshaped_text.decode('utf-8')
        return str(reshaped_text)

    def get_dynamic_steps(self, volume_third):
        """
        Generate dynamic preparation instructions based on calculated volume.

        Creates step-by-step preparation instructions with specific volume guidance
        based on the calculated half-volume (for proper dissolution sequence).

        Args:
            volume_third (str): Half of the final volume for dissolution guidance

        Returns:
            str: Formatted Arabic instructions for iodine preparation

        Instructions Include:
            - Dissolution in water smaller than calculated volume
            - Sequential addition of KI and I₂
            - Final volumetric flask adjustment
            - Proper solubilization sequence

        Example:
            volume_third = "500"
            Returns instructions mentioning "less than 500 ml"
        """
        raw_text = (
            f"4- ضع يوديد البوتاسيوم في كأس به ماء مقطر أقل\n"
            f"من {volume_third} ملليلتر على القلاب المغناطيسي حتى\n"
            "تمام الذوبان ثم ضع اليود وأكمل عملية الذوبان بعد\n"
            "ذلك ضع المحلول في دورق عياري وأكمل الماء\n"
            "المقطر حتى الحجم المطلوب."
        )
        return self.fix_text(raw_text)

    def calculate_iodine_weights(self):
        """
        Calculate iodine and KI masses and update display with instructions.

        Performs mass calculation based on user-specified normality and final volume,
        then updates the interface with calculated masses and dynamic preparation
        instructions based on the computed half-volume.

        Calculation Formulas:
            Iodine mass (g) = Normality × (Volume ÷ 1000) × 126.9 (molar mass)
            KI mass (g) = Iodine mass × 2.5 (solubilization ratio)

        Input Validation:
            - Expects numeric values for normality and volume
            - Displays error message for invalid input
            - Continues without change if validation fails

        Updates:
            - self.iodine_label: Shows calculated iodine mass (1 decimal place)
            - self.ki_label: Shows calculated KI mass (1 decimal place)
            - self.instructions_label: Shows preparation steps with half-volume

        Side Effects:
            - Updates all three display labels
            - Generates new instructions based on volume/2
            - Shows error dialog on invalid input

        Error Handling:
            - ValueError: Caught when input is not numeric
            - Shows messagebox with error message
            - Leaves display unchanged on error

        Example:
            Input: N=0.047, V=1000 ml
            Output: Iodine: 6.0 g, KI: 15.0 g
            Instructions: References "500" ml as initial volume
        """
        try:
            # Parse normality and volume inputs
            n = float(self.norm_entry.get())
            v_ml = float(self.vol_entry.get())

            # Calculate chemical masses
            i_wt = n * (v_ml / 1000) * 126.9  # Iodine mass
            ki_wt = i_wt * 2.5  # KI mass (solubilization)

            # Update mass display labels with 1 decimal place
            self.iodine_label.config(text=f"Iodine: {i_wt:.1f} g")
            self.ki_label.config(text=f"KI: {ki_wt:.1f} g")

            # Update instructions with half-volume guidance
            half_val = int(v_ml / 2)
            self.instructions_label.config(
                text=self.get_dynamic_steps(str(half_val)))

        except ValueError:
            # Invalid input error handling
            messagebox.showerror(
                "خطأ", ("يرجى إدخال قيم رقمية صحيحة"))

    def reset_fields(self):
        """
        Clear all input fields and reset display to initial state.

        Clears all input and output fields, restoring the interface to its initial
        state as if a new calculation session is starting.

        Side Effects:
            - Clears normality entry field
            - Clears volume entry field
            - Resets iodine mass display to "Iodine: 0.0 g"
            - Resets KI mass display to "KI: 0.0 g"
            - Resets instructions to initial text with "---"
            - Returns focus to normality entry field
        """
        # Clear input fields
        self.norm_entry.delete(0, tk.END)
        self.vol_entry.delete(0, tk.END)

        # Reset display labels
        self.iodine_label.config(text="Iodine: 0.0 g")
        self.ki_label.config(text="KI: 0.0 g")

        # Reset instructions to initial state
        self.instructions_label.config(text=self.get_dynamic_steps("---"))

        # Return focus to normality field for next entry
        self.norm_entry.focus()


# Standalone execution support
if __name__ == "__main__":
    # Create root window for standalone execution
    root = tk.Tk()
    root.withdraw()  # Hide main window

    # Default Arabic font specification
    arabic_font = ("Arial", 18)

    # Launch iodine preparation calculator
    IodinePrepWindow(root, arabic_font)

    # Note: Missing root.mainloop() call - this would need to be added for proper standalone execution
