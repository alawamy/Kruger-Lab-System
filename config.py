"""
Kruger Lab System - Configuration and Utility Module

This module serves as the central configuration hub for the Kruger Tamiya Lab Management System.
It manages all paths, constants, version information, and provides critical utility functions
for handling Arabic text in the tkinter GUI framework.

Key responsibilities:
    1. Application Metadata: Version, author, release information
    2. Directory Paths: All centralized path definitions for reports, settings, backups, logs
    3. Security: Encryption/protection passwords for Excel sheets
    4. Arabic Text Processing: Utilities to handle right-to-left (RTL) text and character reshaping
    5. System Detection: Locale detection and Arabic system identification
    6. Template Management: Reference to Excel and other template files

Arabic Text Handling Explanation:
    The tkinter framework has known issues with Arabic text rendering, particularly:
    - Right-to-left (RTL) text direction
    - Character substitution and reshaping (each Arabic letter has different forms based on position)
    - Bidirectional text mixing (Arabic + English)
    
    This module provides utilities using the 'arabic_reshaper' and 'python-bidi' libraries to:
    - Detect Arabic content in strings
    - Properly reshape Arabic characters for display
    - Handle bidirectional text layout
    - Automatically apply transformations only when needed

Version: 2.0
Last Updated: 2026-04-06
Author: Kruger Lab System Development Team
"""

from pathlib import Path
import sys
import locale

import arabic_reshaper
from bidi.algorithm import get_display

# ============================================================================
# Application Metadata
# ============================================================================
# Version and release information for tracking
APP_NAME = "Kruger Lab System"
APP_VERSION = "2.0.0"
APP_AUTHOR = "Kruger Lab"
APP_DESCRIPTION = "نظام إدارة المعمل المتكامل - الإصدار 2.0"
LAST_UPDATED = "2026-04-06"

# ============================================================================
# Directory Path Configuration
# ============================================================================
# All paths are centralized here for easy modification and maintenance
BASE_DIR = Path("D:/Plant Reports")
DAILY_SHEET = BASE_DIR / "Daily Reports/Daily Sheets"
DAILY_PDF = BASE_DIR / "Daily Reports/Daily PDFs"
SETTING_DIR = BASE_DIR / "Setting"
EXTERNAL_DIR = BASE_DIR / "External Plants"
DRY_SLUDGE_DIR = BASE_DIR / "Dry Sludge"
MONTHLY_REPORT_DIR = BASE_DIR / "Monthly Report"
QUALITY_REPORT_DIR = BASE_DIR / "Quality"

# Support for both development and packaged (frozen) executable environments
INTERNAL_BASE = Path(sys._MEIPASS) if getattr(  # type: ignore
    sys, 'frozen', False) else Path(__file__).parent  # type: ignore

# Logo path for application branding
IMG = SETTING_DIR / "Logo.png"

# ============================================================================
# Security and Templates
# ============================================================================
# Master password for protecting sensitive Excel sheets from unauthorized editing
SHEET_PASS = '01006610166'

# List of Excel templates and resources required by the application
TEMPLATES = [
    "Daily Template.xlsx",
    "elements.csv",
    "Dry Sludge Template.xlsx",
    "Logo.png",
    "Logo.ico"
]

# ============================================================================
# Feature Toggles (Version 2.0 Enhancements)
# ============================================================================
CALENDAR_ENABLED = True  # Interactive date picker for report date selection
AUTO_BACKUP = True       # Automatic backup of daily reports and settings
LOGGING_ENABLED = True   # Detailed operation logging for troubleshooting

# ============================================================================
# Supporting Directories
# ============================================================================
BACKUP_DIR = BASE_DIR / "Backup"  # Automatic backup storage location
TEMP_DIR = BASE_DIR / "Temp"      # Temporary file storage for processing
LOG_DIR = BASE_DIR / "Logs"       # Application event and error logs


def is_arabic_system() -> bool:
    """
    Detect if the operating system uses Arabic as the default locale.

    This function checks the system's default locale settings to determine
    if Arabic language support is enabled system-wide. This information is used
    to decide whether Arabic text transformation is necessary, as some systems
    have native Arabic support built-in.

    Returns:
        bool: True if the default system locale is Arabic (ar_*), False otherwise.

    Note:
        - Safely handles cases where locale detection fails
        - Returns False if any exception occurs during detection
        - Used to optimize text processing and skip unnecessary transformations

    Example:
        >>> if is_arabic_system():
        ...     # System has native Arabic support
        ...     arabic_text = "مرحبا"  # Display as-is
        ... else:
        ...     # Need to apply text reshaping
        ...     arabic_text = ar_func("مرحبا")
    """
    try:
        locale_name = locale.getdefaultlocale()[0] or ""
        return locale_name.lower().startswith("ar")
    except Exception:
        return False


def contains_arabic(text: str) -> bool:
    """
    Detect if a string contains any Arabic characters.

    Checks whether the input text includes Unicode characters in the Arabic ranges.
    This is used to determine if special text processing is required.

    The function checks two Unicode ranges:
    - U+0600 to U+06FF: Main Arabic block
    - U+0750 to U+077F: Arabic Supplement block (for rare characters)

    Args:
        text (str): The string to analyze for Arabic content.

    Returns:
        bool: True if the string contains at least one Arabic character, False otherwise.

    Example:
        >>> contains_arabic("مرحبا")
        True
        >>> contains_arabic("Hello")
        False
        >>> contains_arabic("Hello مرحبا")
        True
    """
    return any("\u0600" <= ch <= "\u06FF" or "\u0750" <= ch <= "\u077F" for ch in text)


def ar_func(word, is_ui_element=True):
    """
    Format Arabic text for proper display in tkinter GUI components.

    This function solves critical tkinter Arabic display issues:
    1. Character Reshaping: Arabic letters change form based on their position
       (isolated, initial, medial, final). tkinter doesn't handle this automatically.
    2. RTL Direction: tkinter doesn't properly handle right-to-left text direction.
    3. Bidirectional Mixed Text: Handles mixed Arabic-English text properly.

    The function applies transformations only when necessary:
    - Skips processing if text is English only (optimization)
    - Skips processing if system has native Arabic support
    - Applies appropriate transformation based on usage context

    Args:
        word (str or any): The text to format. Non-string types are converted to string.
        is_ui_element (bool, optional): Affects transformation strategy:
            - True (default): For UI elements (Labels, Buttons) that need full RTL handling
            - False: For window titles that don't need RTL (Windows handles these natively)

    Returns:
        str: The properly formatted text ready for tkinter display.

    Note:
        Critical distinction between UI elements vs titles:
        - UI Elements (Labels/Buttons) in tkinter have RTL display issues and need reshaping + bidi
        - Window Titles (Title Bar) are handled by Windows/Linux and need no processing

    Example:
        >>> ar_func("مرحبا", is_ui_element=True)  # For Label text
        'ابحرم'  # Reshaped and reordered for proper display

        >>> ar_func("مرحبا", is_ui_element=False)  # For window title
        'مرحبا'  # Returned as-is, OS handles it

    Raises:
        None: Function safely handles all input types and returns a string in all cases.
    """
    if not word:
        return ""

    text = str(word)

    # Optimization: English text requires no processing
    if not contains_arabic(text):
        return text

    # Optimization: Arabic system has native support, skip transformation
    if is_arabic_system():
        return text

    # UI Elements: Apply full reshaping and bidirectional algorithm for proper display
    if is_ui_element:
        try:
            # Step 1: Reshape Arabic characters for their position in the word
            reshaped = arabic_reshaper.reshape(text)
            # Step 2: Apply bidirectional text algorithm to handle mixed Arabic-English
            # This is the critical step that makes Arabic display correctly in tkinter
            return str(get_display(reshaped))
        except Exception:
            # Fallback: Return original text if reshaping fails
            return text

    # Window Titles: Return as-is since Windows/Linux handle title bar text natively
    # Applying transformations to titles would cause visual corruption
    return text
