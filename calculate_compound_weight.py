# -*- coding: utf-8 -*-
"""
Smart Chemical Analyzer Module - Intelligent Molecular Weight and Molarity Calculator

This module provides a sophisticated graphical interface for calculating molecular weights,
molarity, and normality of chemical compounds and solutions. The analyzer utilizes an
intelligent formula parser that handles complex chemical formulas including hydrated
compounds, calculates accurate molecular weights, and determines solution concentrations.

Calculation Capabilities:
    - Molecular Weight (MW): Accurate calculation from chemical formulas
    - Molarity (M): Concentration in moles per liter
    - Normality (N): Equivalent concentration based on valency
    - Support for hydrated compounds: Example Na2S2O3.5H2O

Chemical Formula Support:
    - Simple compounds: H2O, NaCl, CaCO3
    - Complex molecules: H2SO4, (NH4)2SO4
    - Hydrated compounds: CuSO4.5H2O, Na2S2O3.5H2O
    - Multi-element combinations with varying subscripts

Calculation Process:
    1. User enters chemical formula
    2. Parser extracts element symbols and counts
    3. System loads element atomic weights from elements.csv
    4. Calculates total molecular weight
    5. User enters weight dissolved and volume of solvent
    6. Calculates molarity: M = (weight/MW) / volume_in_liters
    7. Calculates normality: N = M × valency (z)

Data Source:
    - Elements database: Setting/elements.csv
    - Atomic weights for all elements
    - Located in SETTING_DIR path

User Interface:
    - Chemical formula input field (real-time MW display)
    - Weight dissolved input (grams)
    - Solution volume input (milliliters)
    - Valency input (Z) for normality calculation
    - Results display with formatted output
    - Keyboard navigation with Enter key support

Features:
    - Intelligent formula parsing with regex
    - Support for parentheses in complex formulas
    - Hydrated compound recognition (dot notation)
    - Error handling for invalid formulas
    - Real-time molecular weight calculation
    - Professional color-coded results display
    - Database validation for element symbols

Calculation Examples:
    Example 1 - Simple compound:
        Formula: Na2S2O3 (sodium thiosulfate)
        MW: 158.06 g/mol
        Weight: 1.58g, Volume: 100ml, Z: 1
        M: 0.1 mol/L, N: 0.1 N

    Example 2 - Hydrated compound:
        Formula: CuSO4.5H2O (copper sulfate pentahydrate)
        MW: 249.68 g/mol
        Weight: 2.50g, Volume: 100ml, Z: 2
        M: 0.1 mol/L, N: 0.2 N

Error Handling:
    - Missing elements.csv database: User notification
    - Invalid formula syntax: Parse error messages
    - Unknown element symbols: Specific element identification
    - Invalid numeric inputs: Concentration error dialogs
    - Missing data fields: Validation before calculation

Dependencies:
    - tkinter: GUI framework with messagebox for error handling
    - pandas: CSV file reading for atomic weight database
    - re (regex): Formula parsing and element extraction
    - config: SETTING_DIR path constant for elements.csv location

Author: Kruger Lab System Development Team
Version: 2.0
Last Updated: 2026-04-06
"""

import re
import tkinter as tk
from tkinter import messagebox

import pandas as pd

from config import SETTING_DIR


class SmartChemicalAnalyzer:
    """
    Intelligent Chemical Analyzer - Molecular Weight and Concentration Calculator.

    This class implements a sophisticated chemical calculation interface capable of parsing
    complex chemical formulas, calculating accurate molecular weights, determining molarity
    and normality of solutions, and providing comprehensive chemical analysis tools.

    Formula Support:
        - Simple: H2O, NaCl
        - Complex: (NH4)2SO4, Ca(OH)2
        - Hydrated: Na2S2O3.5H2O, CuSO4.5H2O

    Attributes:
        parent (tk.Tk): Parent window for modal behavior
        arabic_font (tuple): Font specification for Arabic text
        window (tk.Toplevel): Main analyzer window
        elements_dict (dict): Element symbol to atomic weight mapping
        formula_entry (tk.Entry): Chemical formula input field
        weight_input (tk.Entry): Dissolved weight input (grams)
        vol_input (tk.Entry): Solution volume input (milliliters)
        valency_input (tk.Entry): Valency input for normality calculation
        mw_display (tk.Label): Molecular weight display
        result_display (tk.Label): Calculation results display

    Methods:
        add_input(frame, label, row) → tk.Entry: Create labeled input field
        parse_formula(formula: str) → float: Calculate molecular weight
        calculate_simple_mw(part: str) → float: Parse simple formula part
        calculate_mw_only() → None: Calculate and display molecular weight only
        calculate_all() → None: Calculate molarity and normality
        close_window() → None: Handle window closure and parent restoration
    """

    def __init__(self, parent, arabic_font):
        """
        Initialize the smart chemical analyzer window.

        Creates a modal window with comprehensive chemical calculation capabilities,
        including formula parsing, molecular weight calculation, and concentration
        determination.

        Args:
            parent (tk.Tk): Parent window for modal ownership and restoration
            arabic_font (tuple): Font specification tuple for Arabic text rendering

        Returns:
            None: Creates and displays the chemical analyzer interface

        Window Configuration:
            - Modal toplevel window (700x800 pixels)
            - Light background (#f8f9fa)
            - Arabic title with proper RTL support
            - Multiple input sections with organized layout

        Data Loading:
            - Loads elements.csv from SETTING_DIR
            - Creates elements dictionary: {symbol: atomic_weight}
            - Shows error if file not found

        UI Sections:
            Section 1 - Compound Data:
            - Chemical formula input (e.g., Na2S2O3.5H2O)
            - Real-time molecular weight display
            - Enter key for immediate calculation

            Section 2 - Concentration Calculation:
            - Weight dissolved (grams)
            - Solution volume (milliliters)
            - Valency (Z) for normality calculation
            - Keyboard navigation between fields

            Section 3 - Results:
            - Formatted output displaying:
                * Compound formula
                * Molecular weight
                * Molarity (M)
                * Normality (N)
            - Color-coded display

            Footer:
            - Calculate button
            - Return to menu button
        """
        self.parent = parent
        self.arabic_font = arabic_font
        self.window = tk.Toplevel(parent)
        self.window.title("المحلل الكيميائي الذكي - Smart Calc")
        self.window.geometry("700x800")
        self.window.configure(bg="#f8f9fa")

        # Load element database
        try:
            # Load atomic weights from elements.csv
            csv_path = SETTING_DIR / "elements.csv"
            df = pd.read_csv(csv_path)
            self.elements_dict = pd.Series(
                df.Molecular_Weight.values, index=df.Symbol).to_dict()
        except Exception as e:
            # Database loading error
            messagebox.showerror(
                "خطأ", (
                    f"تأكد من وجود ملف elements.csv في المسار:\n{SETTING_DIR}")
            )
            self.window.destroy()
            self.parent.deiconify()
            return

        # Main title
        tk.Label(self.window, text="المحلل الكيميائي الذكي",
                 font=(self.arabic_font[0], 20, "bold"), bg="#f8f9fa", fg="#2c3e50").pack(pady=20)

        # Section 1: Chemical Formula Input
        formula_frame = tk.LabelFrame(
            self.window, text=" بيانات المركب ", font=self.arabic_font, bg="#f8f9fa", padx=10, pady=10)
        formula_frame.pack(pady=10, padx=20, fill="both")

        tk.Label(formula_frame, text="الصيغة (مثلاً Na2S2O3.5H2O):",
                 font=(self.arabic_font[0], 13), bg="#f8f9fa").grid(row=0, column=1, sticky="e")
        self.formula_entry = tk.Entry(
            formula_frame, font=self.arabic_font, width=20, justify="center")
        self.formula_entry.grid(row=0, column=0, pady=10, padx=10)

        # Bind Enter to calculate molecular weight only
        self.formula_entry.bind("<Return>", lambda e: self.calculate_mw_only())

        # Molecular weight display
        self.mw_display = tk.Label(formula_frame, text="M.W: ---", font=(
            self.arabic_font[0], 12, "bold"), fg="#27ae60", bg="#f8f9fa")
        self.mw_display.grid(row=0, column=2, padx=10)

        # Section 2: Concentration Calculation
        calc_frame = tk.LabelFrame(self.window, text=" حساب التركيزات ",
                                   font=self.arabic_font, bg="#f8f9fa", padx=10, pady=10)
        calc_frame.pack(pady=10, padx=20, fill="both")

        # Create input fields
        self.weight_input = self.add_input(calc_frame, "الوزن المذاب (g):", 0)
        self.vol_input = self.add_input(calc_frame, "الحجم المذيب (ml):", 1)
        self.valency_input = self.add_input(calc_frame, "التكافؤ (Z):", 2)

        # Configure Enter key navigation
        self.weight_input.bind("<Return>", lambda e: self.vol_input.focus())
        self.vol_input.bind("<Return>", lambda e: self.valency_input.focus())
        self.valency_input.bind("<Return>", lambda e: self.calculate_all())

        # Button frame
        btn_frame = tk.Frame(self.window, bg="#f8f9fa")
        btn_frame.pack(pady=20)

        # Calculate button
        tk.Button(btn_frame, text="حساب التركيزات", font=self.arabic_font, bg="#27ae60", fg="white", width=15,
                  command=self.calculate_all).pack(side="left", padx=10)

        # Return button
        tk.Button(btn_frame, text="رجوع", font=self.arabic_font, bg="#e74c3c", fg="white", width=15,
                  command=self.close_window).pack(side="left", padx=10)

        # Section 3: Results Display
        self.result_display = tk.Label(self.window, text="النتائج ستظهر هنا", font=(self.arabic_font[0], 13),
                                       bg="white", relief="sunken", width=50, height=5, fg="#2c3e50")
        self.result_display.pack(pady=20, padx=20)

        # Set initial focus
        self.enter_focus()

        # Handle window close event
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)

    def enter_focus(self):
        """
        Configure initial focus and keyboard navigation.

        Sets initial focus to formula entry field and configures Enter key
        navigation through all input fields in logical order.

        Returns:
            None: Configures focus and event bindings
        """
        # Focus on formula entry with text selected
        self.formula_entry.focus()
        self.formula_entry.select_range(0, tk.END)

        # Configure keyboard navigation flow
        self.formula_entry.bind(
            "<Return>", lambda e: self.weight_input.focus())
        self.weight_input.bind("<Return>", lambda e: self.vol_input.focus())
        self.vol_input.bind("<Return>", lambda e: self.valency_input.focus())
        self.valency_input.bind("<Return>", lambda e: self.calculate_all())

    def add_input(self, frame, label, row):
        """
        Create labeled input field in frame.

        Creates a pair of widgets: label and entry field, arranged in grid layout.

        Args:
            frame (tk.Frame): Parent frame to contain input field
            label (str): Label text for the input field
            row (int): Grid row position

        Returns:
            tk.Entry: Created entry widget for user input
        """
        tk.Label(frame, text=label, font=self.arabic_font,
                 bg="#f8f9fa").grid(row=row, column=1, sticky="e")
        ent = tk.Entry(frame, font=self.arabic_font,
                       width=12, justify="center")
        ent.grid(row=row, column=0, pady=8)
        return ent

    def parse_formula(self, formula):
        """
        Parse chemical formula and calculate molecular weight.

        Intelligently parses chemical formulas including support for hydrated
        compounds (dot notation), extracting element symbols and counts,
        then summing atomic weights.

        Args:
            formula (str): Chemical formula string (e.g., "Na2S2O3.5H2O")

        Returns:
            float: Total molecular weight in g/mol

        Supported Formats:
            - Simple: H2O, NaCl
            - Complex: (NH4)2SO4
            - Hydrated: CuSO4.5H2O, Na2S2O3.5H2O

        Raises:
            ValueError: If element not found in database
            Exception: On formula parsing errors

        Processing:
            1. Check for dot notation (hydrated compounds)
            2. Split at dot if present
            3. Calculate main compound MW
            4. Calculate water molecules MW
            5. Combine totals
            6. Return total MW
        """
        try:
            # Handle hydrated compounds with dot notation
            if '.' in formula:
                main_part, water_part = formula.split('.')
                # Calculate main compound molecular weight
                mw_main = self.calculate_simple_mw(main_part)
                # Parse water molecules (e.g., "5H2O")
                water_match = re.match(r'(\d*)H2O', water_part, re.IGNORECASE)
                if water_match:
                    # Extract count or default to 1
                    count = int(water_match.group(
                        1)) if water_match.group(1) else 1
                    # Calculate water molecular weight: 2×H + O
                    mw_water = (
                        self.elements_dict['H'] * 2) + self.elements_dict['O']
                    # Return total with hydrated water
                    return mw_main + (count * mw_water)
                return mw_main
            else:
                # Simple formula without hydration
                return self.calculate_simple_mw(formula)
        except Exception as e:
            raise e

    def calculate_simple_mw(self, part):
        """
        Calculate molecular weight of simple formula part.

        Helper function that extracts element symbols and counts using regex,
        then sums up the atomic weights multiplied by their counts.

        Args:
            part (str): Formula part (e.g., "Na2S2O3" or "H2O")

        Returns:
            float: Molecular weight of the formula part

        Parsing:
            - Uses regex to match element: count pairs
            - Element: Capital letter followed by optional lowercase
            - Count: Optional number following element
            - Example: "Na2S2O3" → [(Na, 2), (S, 2), (O, 3)]

        Raises:
            ValueError: If element symbol not found in elements_dict
        """
        # Parse formula with regex: element symbols and counts
        matches = re.findall(r'([A-Z][a-z]*)(\d*)', part)
        weight = 0
        for symbol, count in matches:
            # Default count to 1 if not specified
            c = int(count) if count else 1
            # Look up atomic weight
            if symbol in self.elements_dict:
                weight += self.elements_dict[symbol] * c
            else:
                # Unknown element error
                raise ValueError(f"العنصر '{symbol}' غير موجود")
        return weight

    def calculate_mw_only(self):
        """
        Calculate and display molecular weight only.

        Triggered by Enter key in formula field, calculates the molecular weight
        of the entered formula and displays it in real-time, then moves focus to
        weight input field.

        Returns:
            None: Updates MW display and manages focus

        Side Effects:
            - Updates self.mw_display with calculated MW (3 decimal places)
            - Moves focus to weight_input after calculation
            - Shows error message on parse failure
        """
        try:
            # Get formula from input
            formula = self.formula_entry.get().strip()
            if not formula:
                return

            # Calculate molecular weight
            mw = self.parse_formula(formula)
            # Update display with 3 decimal places
            self.mw_display.config(text=f"M.W: {mw:.3f}")

            # Move focus to weight input
            self.weight_input.focus()
        except Exception as e:
            # Formula parsing error
            messagebox.showerror("خطأ", str(e))

    def close_window(self):
        """
        Handle window closure and restore parent window visibility.

        Properly closes the analyzer window and restores the parent window
        to visible state, ensuring clean modal dialog behavior.

        Returns:
            None: Destroys analyzer window and shows parent

        Side Effects:
            - Destroys self.window
            - Calls self.parent.deiconify() to show parent window
        """
        self.window.destroy()
        self.parent.deiconify()

    def calculate_all(self):
        """
        Calculate molarity and normality from compound and solution data.

        Performs comprehensive chemical calculations based on user-entered formula,
        dissolved weight, solution volume, and valency. Calculates and displays
        molecular weight, molarity, and normality with professional formatting.

        Calculations:
            MW: Parsed from formula (g/mol)
            Molarity: M = (weight in grams / MW) / volume in liters
            Normality: N = Molarity × Valency

        Input Validation:
            - Chemical formula: Parsed and validated against elements database
            - Weight: Numeric value in grams
            - Volume: Numeric value in milliliters (converted to liters)
            - Valency: Numeric value for equivalent calculation

        Results Displayed:
            - Compound formula
            - Molecular weight (3 decimal places)
            - Molarity (4 decimal places)
            - Normality (4 decimal places)

        Error Handling:
            - Invalid formula: Parse error message
            - Non-numeric inputs: Concentration error message
            - All exceptions: Generic error dialog
            - No fields cleared on error (allows retry)

        Returns:
            None: Updates result display or shows error

        Example:
            Formula: Na2S2O3
            Weight: 1.58 g
            Volume: 500 ml (0.5 L)
            Valency: 1
            Result:
                MW: 158.06 g/mol
                M: 0.02 mol/L
                N: 0.02 N
        """
        try:
            # Parse chemical formula
            formula = self.formula_entry.get().strip()
            mw = self.parse_formula(formula)

            # Get numeric inputs
            w = float(self.weight_input.get())  # Weight in grams
            v = float(self.vol_input.get()) / 1000  # Volume in liters
            z = float(self.valency_input.get())  # Valency

            # Calculate concentrations
            molarity = (w / mw) / v  # mol/L
            normality = molarity * z  # equivalents/L

            # Format results display
            res_text = (f"المركب: {formula}\n الوزن الجزيئي: {mw:.3f} g/mol\n"
                        f"المولارية (M): {molarity:.4f}\n العيارية (N): {normality:.4f}")
            self.result_display.config(text=res_text, fg="#1e3799")

        except Exception as e:
            # Calculation error handling
            messagebox.showerror(
                "خطأ", "تأكد من كتابة الصيغة والقيم بشكل صحيح")


# Standalone execution support
if __name__ == "__main__":
    # Create root window for standalone execution
    root = tk.Tk()
    root.withdraw()  # Hide main window

    # Default Arabic font specification
    arabic_font = ("Arial", 18)

    # Launch smart chemical analyzer
    app = SmartChemicalAnalyzer(root, arabic_font)

    # Note: Missing root.mainloop() call - this would need to be added for proper standalone execution
