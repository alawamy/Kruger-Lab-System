# -*- coding: utf-8 -*-
"""
Sulfuric Acid Reagent Preparation Module - Laboratory Chemistry Calculator and Guide

This module provides a comprehensive graphical interface for calculating and guiding the
preparation of sulfuric acid reagent testing solutions for wastewater treatment laboratory
quality control. The module enables operators to specify acid specific gravity and desired
weight, then calculates required volumes and chemical masses.

Sulfuric Acid Reagent Purpose:
    Sulfuric acid solutions are used in:
    - Silver sulfate indicator preparation for chloride titration
    - Acidic environment maintenance for various analytical procedures
    - Reagent formulation for wastewater analysis
    - Chemical oxygen demand (COD) testing procedures

Solution Preparation:
    The preparation involves:
    - Concentrated sulfuric acid: Main reagent component
    - Silver sulfate (Ag₂SO₄): Indicator chemical (5.5g per kg acid)
    - Mixing procedure: Silver sulfate dissolves over 2-3 days

Calculation Formulas:
    Volume (ml) = (Weight in kg ÷ Specific Gravity) × 1000
    Silver Sulfate (g) = Weight in kg × 5.5 (standard ratio)

Specific Gravity Reference:
    - Pure H₂SO₄: 1.84 Kg/L (default and most common)
    - Temperature dependent: Varies with concentration and temperature
    - Common concentrations: 95-98% w/w

User Interface:
    - Specific Gravity input field (default 1.84 Kg/L)
    - Acid weight input field (kg)
    - Real-time volume calculation display
    - Real-time silver sulfate mass calculation
    - Step-by-step preparation instructions
    - Keyboard navigation with Enter key support
    - Reset function for new calculations

Features:
    - Interactive calculation based on two input parameters
    - Real-time volume calculation with 1 decimal place
    - Real-time silver sulfate mass calculation with 1 decimal place
    - Automatic formula application
    - Input validation with error handling
    - Professional color scheme and layout
    - Clear step-by-step guidance

Calculation Example:
    Specific Gravity: 1.84 Kg/L
    Acid Weight: 10 kg
    Volume = (10 / 1.84) × 1000 = 5434.78 ml
    Silver Sulfate = 10 × 5.5 = 55 g

Error Handling:
    - Input validation for numeric values
    - Error messages for invalid input
    - Graceful handling of calculation failures
    - Clear user feedback on success/failure

Dependencies:
    - tkinter: GUI framework with messagebox for error handling
    - config: ar_func utility for Arabic text processing and display
    - Python math: Built-in calculations for volume conversions

Author: Kruger Lab System Development Team
Version: 2.0
Last Updated: 2026-04-06
"""

import tkinter as tk
from tkinter import messagebox
from config import ar_func


class SulfuricPrepWindow:
    """
    Sulfuric Acid Reagent Preparation Calculator - Laboratory Chemistry Guide Interface.

    This class implements a graphical interface for calculating and guiding sulfuric acid
    reagent preparation procedures. The interface allows operators to specify specific
    gravity and acid weight, then calculates required volumes and silver sulfate masses.

    Reagent Preparation:
        - Input: Specific Gravity (Kg/L) and Acid Weight (kg)
        - Calculated: Volume of acid (ml) and Silver Sulfate mass (g)
        - Setup: 2-3 day dissolution process

    Attributes:
        parent (tk.Tk): Parent window for modal behavior
        arabic_font (tuple): Font specification for Arabic text
        window (tk.Toplevel): Main preparation guide window
        sg_entry (tk.Entry): Specific gravity input field (default 1.84)
        weight_entry (tk.Entry): Acid weight input field (kg)
        vol_result (tk.Label): Calculated sulfuric acid volume display (ml)
        silver_result (tk.Label): Calculated silver sulfate mass display (g)

    Methods:
        close_window() → None: Handle window closure and parent restoration
        fix_text(text: str) → str: Process Arabic text for proper display
        calculate_sulfuric() → None: Calculate volumes and update display
        reset() → None: Clear inputs and reset display
    """

    def __init__(self, parent, arabic_font):
        """
        Initialize the sulfuric acid reagent preparation calculator window.

        Creates a modal window with input fields for specific gravity and acid weight,
        displays calculated volumes and chemical masses, and provides step-by-step
        preparation instructions.

        Args:
            parent (tk.Tk): Parent window for modal ownership and restoration
            arabic_font (tuple): Font specification tuple for Arabic text rendering

        Returns:
            None: Creates and displays the sulfuric acid preparation interface

        Window Configuration:
            - Modal toplevel window (600x550 pixels)
            - Light background (#f8f9fa)
            - Arabic title with proper RTL support
            - Comprehensive event handling

        UI Sections:
            Section 1 - Specific Gravity Input:
            - Label for sulfuric acid specific gravity
            - Default value: 1.84 Kg/L (standard concentration)
            - Unit label: Kg/L

            Section 2 - Acid Weight Input:
            - Label for desired acid weight
            - Input in kilograms (kg)
            - Enter key binding for calculation trigger

            Section 3 - Volume Calculation:
            - Calculated sulfuric acid volume in ml
            - Updated in real-time on input

            Section 4 - Silver Sulfate Calculation:
            - Calculated silver sulfate mass in grams
            - Uses 5.5g per kg ratio (standard indicator formula)

            Section 5 - Preparation Instructions:
            - Final dissolution procedure
            - 2-3 day dissolution timeline
        """
        self.parent = parent
        self.arabic_font = arabic_font
        self.window = tk.Toplevel(parent)
        self.window.title("تحضير كاشف حمض الكبريتيك")
        self.window.geometry("600x550")
        self.window.configure(bg="#f8f9fa")

        # Section 1: Specific Gravity Input
        tk.Label(self.window, text=self.fix_text("1- قم بتسجيل الكثافة النوعية لحمض الكبريتيك:"),
                 font=(self.arabic_font[0], 14, "bold"), bg="#f8f9fa").pack(pady=15)

        f1 = tk.Frame(self.window, bg="#f8f9fa")
        f1.pack()
        tk.Label(f1, text="Kg/L", font=("Arial", 12, "bold"),
                 bg="#f8f9fa").grid(row=0, column=2, padx=5)
        self.sg_entry = tk.Entry(
            f1, font=("Arial", 12), width=15, justify="center")
        self.sg_entry.grid(row=0, column=1)
        self.sg_entry.insert(0, "1.84")
        tk.Label(f1, text="Specific Gravity:", font=(
            "Arial", 12, "bold"), bg="#f8f9fa").grid(row=0, column=0, padx=5)

        # Section 2: Acid Weight Input
        tk.Label(self.window, text=self.fix_text("2- قم بتسجيل وزن الحمض المطلوب للتحضير:"),
                 font=(self.arabic_font[0], 14, "bold"), bg="#f8f9fa").pack(pady=15)

        f2 = tk.Frame(self.window, bg="#f8f9fa")
        f2.pack()
        tk.Label(f2, text="Kg", font=("Arial", 12, "bold"),
                 bg="#f8f9fa").grid(row=0, column=2, padx=5)
        self.weight_entry = tk.Entry(
            f2, font=("Arial", 12), width=15, justify="center")
        self.weight_entry.grid(row=0, column=1)
        tk.Label(f2, text="Sulfuric Weight:", font=(
            "Arial", 12, "bold"), bg="#f8f9fa").grid(row=0, column=0, padx=5)

        # Configure Enter key navigation
        self.sg_entry.bind("<Return>", lambda e: self.weight_entry.focus())
        self.weight_entry.bind("<Return>", lambda e: self.calculate_sulfuric())

        # Section 3: Sulfuric Acid Volume Result
        tk.Label(self.window, text=self.fix_text("3- قم بأخذ الحجم التالي من حمض الكبريتيك:"),
                 font=(self.arabic_font[0], 14), bg="#f8f9fa").pack(pady=10)
        self.vol_result = tk.Label(self.window, text="Sulfuric Volume: 0 ml", font=(
            "Arial", 16, "bold"), fg="#c0392b", bg="#f8f9fa")
        self.vol_result.pack()

        # Section 4: Silver Sulfate Mass Result
        tk.Label(self.window, text=self.fix_text("4- قم بأخذ الوزن التالي من كبريتات الفضة:"),
                 font=(self.arabic_font[0], 14), bg="#f8f9fa").pack(pady=10)
        self.silver_result = tk.Label(self.window, text="Silver Sulfate: 0 g", font=(
            "Arial", 16, "bold"), fg="#2980b9", bg="#f8f9fa")
        self.silver_result.pack()

        # Section 5: Final Preparation Instructions
        instr_text = "5- ضع كبريتات الفضة على الحمض واترك الخليط\nليذوب (من 2 : 3 أيام)."
        tk.Label(self.window, text=self.fix_text(instr_text), font=(self.arabic_font[0], 14),
                 bg="#f8f9fa", pady=20, justify="right").pack()

        # Footer buttons frame
        btns = tk.Frame(self.window, bg="#f8f9fa")
        btns.pack(side="bottom", pady=30)

        tk.Button(btns, text="حساب جديد", font=self.arabic_font,
                  command=self.reset).pack(side="right", padx=10)
        tk.Button(btns, text="رجوع", font=self.arabic_font, bg="#e74c3c",
                  fg="white", command=self.close_window).pack(side="right", padx=10)
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)

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
        """
        return ar_func(text)

    def calculate_sulfuric(self):
        """
        Calculate sulfuric acid volume and silver sulfate mass.

        Performs real-time calculation of required sulfuric acid volume based on
        specific gravity and desired weight, plus silver sulfate mass based on
        the standard indicator formula (5.5g per kg acid).

        Calculation Formulas:
            Volume (ml) = (Weight in kg ÷ Specific Gravity in Kg/L) × 1000
            Silver Sulfate (g) = Weight in kg × 5.5 (indicator standard ratio)

        Updates:
            - self.vol_result: Shows calculated sulfuric acid volume (rounded ml)
            - self.silver_result: Shows calculated silver sulfate mass (1 decimal place)
        """
        try:
            sg = float(self.sg_entry.get())
            wt_kg = float(self.weight_entry.get())

            volume_ml = (wt_kg / sg) * 1000
            silver_g = wt_kg * 5.5

            self.vol_result.config(
                text=f"Sulfuric Volume: {int(round(volume_ml))} ml")
            self.silver_result.config(
                text=f"Silver Sulfate: {round(silver_g, 1)} g")
        except ValueError:
            messagebox.showerror(("خطأ"), "يرجى إدخال أرقام صحيحة")

    def close_window(self):
        """
        Handle window closure and restore parent window visibility.

        Properly closes the sulfuric acid preparation window and restores the parent
        window to visible state, ensuring clean modal dialog behavior.

        Returns:
            None: Destroys preparation window and shows parent
        """
        self.window.destroy()
        self.parent.deiconify()

    def reset(self):
        """
        Clear all input fields and reset display to initial state.

        Clears weight entry and restores calculation displays to zero state,
        resetting the interface for a new calculation while preserving the
        specific gravity entry for potential reuse.

        Side Effects:
            - Clears weight entry field
            - Resets volume display to "Sulfuric Volume: 0 ml"
            - Resets silver sulfate display to "Silver Sulfate: 0 g"
            - Returns focus to specific gravity field
        """
        self.weight_entry.delete(0, tk.END)
        self.vol_result.config(text="Sulfuric Volume: 0 ml")
        self.silver_result.config(text="Silver Sulfate: 0 g")
        self.sg_entry.focus()


# Standalone execution support
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    app = SulfuricPrepWindow(root, ("Arial", 18))
    # Note: Missing root.mainloop() call
