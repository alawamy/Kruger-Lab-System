# -*- coding: utf-8 -*-
"""
Laboratory Operations Menu - Advanced Testing and Procedures Hub

This module provides the main navigation menu for advanced laboratory procedures and 
chemical testing operations. It serves as the central hub connecting all specialized 
testing modules including titrations, pump adjustments, and chemical preparation procedures.

The module opens an interactive Toplevel window with organized button groups representing 
different laboratory domains, each with distinctive color coding for intuitive navigation:
    - Blue buttons (#e3f2fd): Chemical titration procedures (Iodine, Sodium Thiosulfate)
    - Orange buttons (#fff3e0): Technical operations (Sludge rates, TDS factors)
    - Green buttons (#c8e6c9): Chemical preparations and inventory
    - Light gray buttons (#f5f5f5): Administrative functions (Settings, Exit)

Key Features:
    - Organized button matrix with color-coded domains for easy navigation
    - Direct access to all titration procedures and calculations
    - Technical parameter adjustment interfaces
    - Chemical preparation calculator hub
    - Settings management with password protection
    - Responsive grid layout with proportional sizing

Author: Kruger Lab System Development Team
Version: 2.0
Last Updated: 2026-04-06
"""

import tkinter as tk
from tkinter import messagebox
# Opens the chemical operations hub
from chemical_module import ChemicalMenuWindow
from settings_module import prompt_for_password
from iodine_module import IodineTitrationWindow
from sludge_pumps_module import SludgePumpsWindow
from sodium_module import SodiumTitrationWindow
from tds_factor import SaltsFactorWindow


def open_main_menu(root_window, arabic_font):
    """
    Open the main laboratory operations menu window.

    Creates and displays an interactive Toplevel window containing organized buttons 
    for accessing all advanced laboratory procedures and operations.

    Args:
        root_window (tk.Tk): The root tkinter window. Will be hidden while menu is open
                           and restored when menu is closed.
        arabic_font (tuple): Font specification tuple (font_name, size, style) for rendering
                           Arabic text correctly. Example: ("Arial", 18, "bold").

    Returns:
        None

    Side Effects:
        - Creates new Toplevel window with menu interface
        - Hides root_window
        - Each submodule launches in its own Toplevel window
        - Restores root_window visibility when menu closes
    """
    # Create main menu window as Toplevel
    main_form = tk.Toplevel(root_window)
    main_form.title("القائمة الرئيسية للمعمل")
    main_form.geometry("900x600")
    main_form.configure(bg="#f0f0f0")

    # Hide the main dashboard window while menu is open
    root_window.withdraw()

    def close_main():
        """Close menu window and restore main dashboard visibility."""
        main_form.destroy()
        root_window.deiconify()

    # Window title label - Application name and branding
    tk.Label(main_form, text="نظام إدارة المعمل المتكامل", bg="#f0f0f0",
             font=(arabic_font[0], 26, "bold"), pady=30).pack()

    # Unified button style factory function
    def create_btn_style(bg_color="#ffffff"):
        """Create standardized button style dictionary."""
        return {
            "font": (arabic_font[0], 18, "bold"),
            "width": 28,
            "height": 2,
            "bd": 3,
            "relief": "raised",
            "bg": bg_color
        }

    # ========================================================================
    # Row 1: Chemical Titration Procedures (Blue - Precision Operations)
    # ========================================================================
    f1 = tk.Frame(main_form, bg="#f0f0f0")
    f1.pack(pady=10)
    tk.Button(f1, text="معايرة اليود", **create_btn_style("#e3f2fd"),
              command=lambda: IodineTitrationWindow(main_form, arabic_font)).pack(side="left", padx=10)
    tk.Button(f1, text="معايرة الصوديوم ثيوسلفات", **create_btn_style("#e3f2fd"),
              command=lambda: SodiumTitrationWindow(main_form, arabic_font)).pack(side="left", padx=10)

    # ========================================================================
    # Row 2: Technical Operations (Orange - System Adjustments)
    # ========================================================================
    f2 = tk.Frame(main_form, bg="#f0f0f0")
    f2.pack(pady=10)
    tk.Button(f2, text="معدل الحمأة", **create_btn_style("#fff3e0"),
              command=lambda: SludgePumpsWindow(main_form, arabic_font)).pack(side="left", padx=10)
    tk.Button(f2, text="معامل الأملاح الذائبة (TDS)", **create_btn_style("#fff3e0"),
              command=lambda: SaltsFactorWindow(main_form, arabic_font)).pack(side="left", padx=10)

    # ========================================================================
    # Row 3: Chemical Management (Green - Preparation and Support)
    # ========================================================================
    f3 = tk.Frame(main_form, bg="#f0f0f0")
    f3.pack(pady=10)
    # Opens ChemicalMenuWindow which provides hub for all chemical preparation procedures
    tk.Button(f3, text="التحضيرات الكيميائية ", **create_btn_style("#c8e6c9"),
              command=lambda: ChemicalMenuWindow(main_form, arabic_font)).pack(side="left", padx=10)
    # Inventory module under development
    tk.Button(f3, text="جرد المعمل الكيميائي", **create_btn_style("#c8e6c9"),
              command=lambda: messagebox.showinfo("قريباً", "موديول الجرد تحت التطوير")).pack(side="left", padx=10)

    # ========================================================================
    # Row 4: Administration and Navigation (Gray and Red)
    # ========================================================================
    f4 = tk.Frame(main_form, bg="#f0f0f0")
    f4.pack(pady=30)
    # Settings button - Requires password authentication
    tk.Button(f4, text="⚙ إعدادات التطبيق", **create_btn_style("#f5f5f5"),
              command=lambda: [main_form.destroy(), prompt_for_password(root_window, arabic_font)]).pack(side="right", padx=10)
    # Exit button - Closes menu and returns to main dashboard
    tk.Button(f4, text="خروج من القائمة", **create_btn_style("#ffcdd2"),
              command=close_main).pack(side="left", padx=10)

    # Configure window close button (X) to properly close menu
    main_form.protocol("WM_DELETE_WINDOW", close_main)


# Script execution for standalone testing
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    arabic_font = ("Arial", 18)
    open_main_menu(root, arabic_font)
    root.mainloop()
