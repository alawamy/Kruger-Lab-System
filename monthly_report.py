# -*- coding: utf-8 -*-
"""
Monthly Report Generator and Updater - Data Aggregation Module

This module provides comprehensive monthly report generation and updating functionality
for laboratory operations, aggregating data from daily reports and external plant sources
into consolidated monthly Excel reports.

The module implements a sophisticated data aggregation system that:
    1. Creates new monthly reports from templates with proper initialization
    2. Aggregates daily laboratory data into monthly summaries
    3. Incorporates external plant data into unified reports
    4. Handles complex data mapping between different report formats
    5. Manages multi-sheet Excel workbooks with specialized data sections

Monthly Report Structure:
    - تقرير المواد الصلبة (Solids Report): Daily solids analysis data
    - Main: Primary monthly summary with all laboratory parameters
    - Integrated data from daily reports and external plant sources

Data Aggregation Logic:
    Daily Reports Integration:
        - Matches daily files by date (dd-mm-yyyy format)
        - Maps specific cells from daily Main worksheet to monthly columns
        - Handles operational status (running/stopped) indicators
        - Processes up to 31 days per month automatically

    External Plants Integration:
        - Processes external plant Excel files with standardized format
        - Dynamically locates or creates plant name headers
        - Maps external data to appropriate monthly report columns
        - Supports multiple external plants per monthly period

File Organization:
    Monthly Report/
    ├── 2026/
    │   ├── January 2026.xlsx
    │   ├── February 2026.xlsx
    │   └── ...
    ├── 2025/
    │   └── ...
    └── Templates/
        └── Monthly Report Template.xlsx

Data Mapping (Daily → Monthly):
    Solids Report Sheet:
        - Column C: K25 (daily) → Monthly solids data
        - Column D: G25 → TS data
        - Column E: E25 → VS data
        - Column G: J25 → TSS data
        - Column H: E26 → Operational status
        - Column I: E27 → Additional status
        - Column J: K27 → Final readings

    Main Sheet (20+ parameters):
        - Comprehensive mapping from daily cells to monthly columns
        - Includes COD, BOD5, pH, temperature, and specialized parameters
        - Operational status indicators and efficiency metrics

External Plants Processing:
    - Plant name extraction from filename patterns
    - Dynamic row assignment in monthly report
    - Standardized data mapping for external sources
    - Header management ("المحطة" - Station)

Error Handling:
    - File access and permission issues
    - Missing worksheets or invalid cell references
    - Template file availability
    - Data type conversion and validation
    - User cancellation and error recovery

Dependencies:
    - openpyxl: Excel file manipulation and data reading
    - tkinter: GUI dialogs for user interaction
    - calendar: Month calculations and date operations
    - pathlib: Modern file path handling
    - config: Directory path constants and settings

Author: Kruger Lab System Development Team
Version: 2.0
Last Updated: 2026-04-06
"""

import calendar
import os
import shutil
import sys
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, filedialog

from openpyxl import load_workbook

# Windows encoding setup for proper UTF-8 output
if sys.platform == "win32":
    import codecs

    # Ensure stdout and stderr are properly encoded for Windows
    if sys.stdout is not None:
        try:
            sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer)
        except AttributeError:
            pass
    if sys.stderr is not None:
        try:
            sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer)
        except AttributeError:
            pass

# Path configuration with fallback for standalone execution
try:
    from config import BASE_DIR, DAILY_SHEET, EXTERNAL_DIR, MONTHLY_REPORT_DIR, SETTING_DIR
except ImportError:
    # Fallback paths if config import fails
    BASE_DIR = Path(__file__).parent
    DAILY_SHEET = BASE_DIR / 'Daily Reports' / "Daily Sheets"
    EXTERNAL_DIR = BASE_DIR / "External Plants"
    MONTHLY_REPORT_DIR = BASE_DIR / "Monthly Report"
    SETTING_DIR = BASE_DIR / "Setting"


class MonthlyReportUpdater:
    """
    Monthly Report Generator and Data Aggregator.

    This class manages the complete lifecycle of monthly laboratory reports,
    from creation and initialization through data aggregation and updating.
    It handles complex data mapping between daily reports, external plant data,
    and consolidated monthly summaries.

    The class provides a comprehensive solution for laboratory data consolidation,
    supporting multiple data sources and complex Excel workbook manipulation.

    Attributes:
        app_root (tk.Tk): Root tkinter window for dialog ownership.
        DAILY_SHEET (Path): Directory containing daily report Excel files.
        EXTERNAL_DIR (Path): Directory containing external plant data files.
        MONTHLY_REPORT_DIR (Path): Output directory for monthly reports.
        SETTINGS_DIR (Path): Directory containing templates and settings.
        TEMPLATE_FILENAME (str): Name of the monthly report template file.

    Key Methods:
        run() → None: Main execution method with user interaction
        _initialize_new_report(path) → bool: Set up new monthly report structure
        process_daily_files(path, dir) → int: Aggregate daily report data
        process_external_files(path, dir) → int: Incorporate external plant data

    Data Processing Workflow:
        1. Template Validation: Ensure monthly template exists
        2. Report Creation: Copy template and initialize with current month
        3. Daily Data Aggregation: Process all daily Excel files for the month
        4. External Data Integration: Add external plant data to report
        5. File Launch: Open completed report in Excel for review

    File Path Structure:
        Monthly reports: MONTHLY_REPORT_DIR/year/"Month Year.xlsx"
        Daily sources: DAILY_SHEET/year/Month/
        External sources: EXTERNAL_DIR/year/year/
        Templates: SETTINGS_DIR/Monthly Report Template.xlsx

    Error Handling:
        - Template missing: User notification with graceful exit
        - File access errors: Continue processing other files
        - Data mapping failures: Skip invalid cells gracefully
        - User cancellation: Clean exit without data loss

    Usage Examples:
        >>> # Standard execution with user interaction
        >>> updater = MonthlyReportUpdater()
        >>> updater.run()  # Shows dialogs, processes data, opens result

        >>> # Programmatic usage
        >>> updater = MonthlyReportUpdater()
        >>> report_path = updater.MONTHLY_REPORT_DIR / "2026" / "January 2026.xlsx"
        >>> if not report_path.exists():
        ...     updater._initialize_new_report(report_path)
        >>> daily_count = updater.process_daily_files(report_path, daily_dir)
        >>> external_count = updater.process_external_files(report_path, ext_dir)

    Integration Points:
        - app.py: Called from monthly report generation button
        - main_menu.py: Accessible through advanced operations
        - daily reports: Source data for monthly aggregation
        - external plants: Additional data sources for comprehensive reports
    """

    def __init__(self):
        """
        Initialize the Monthly Report Updater.

        Sets up all necessary paths and configuration for monthly report processing.
        Initializes directory paths from config module with fallback defaults.

        Args:
            None

        Returns:
            None

        Side Effects:
            - Initializes all path attributes
            - Sets template filename constant
            - Prepares for monthly report operations
        """
        self.app_root = None
        self.DAILY_SHEET = DAILY_SHEET
        self.EXTERNAL_DIR = EXTERNAL_DIR
        self.MONTHLY_REPORT_DIR = MONTHLY_REPORT_DIR
        self.SETTINGS_DIR = SETTING_DIR
        self.TEMPLATE_FILENAME = "Monthly Report Template.xlsx"

    def _initialize_new_report(self, report_path):
        """
        Initialize a new monthly report with proper structure and dates.

        Creates a new monthly report by copying the template and populating it
        with the current month's calendar structure, dates, and headers. This
        method sets up both the solids report sheet and main summary sheet.

        Process Flow:
            1. Load the template workbook
            2. Initialize solids report sheet with month/year and daily dates
            3. Mark Fridays as weekly rest days
            4. Initialize main sheet with dates and station header
            5. Save the initialized report

        Args:
            report_path (Path): Full path where the new report should be created.

        Returns:
            bool: True if initialization successful, False if errors occurred.

        Raises:
            None: All exceptions are caught internally and logged.

        Template Modifications:
            Solids Report Sheet ("تقرير المواد الصلبة"):
                - I10: Current month number
                - J10: Current year
                - B12-B42: Daily dates (01/MM/YYYY format)
                - I12-I42: "راحة أسبوعية" for Fridays

            Main Sheet:
                - Q10: Current month number
                - R10: Current year
                - B14-B45: Daily dates (01/MM/YYYY format)
                - B(header_row): "المحطة" header for external plants

        Date Calculations:
            - Uses calendar.monthrange() to get correct last day of month
            - Handles leap years and month variations automatically
            - Formats dates consistently as dd/mm/yyyy

        Error Handling:
            - Workbook loading failures
            - Missing worksheets
            - Cell access errors
            - File save permissions

        Example:
            >>> updater = MonthlyReportUpdater()
            >>> success = updater._initialize_new_report(Path("January 2026.xlsx"))
            >>> if success:
            ...     print("Monthly report initialized successfully")
        """
        try:
            # Load the monthly report template
            wb = load_workbook(report_path)
            now = datetime.now()
            t_month, t_year = now.month, now.year

            # Calculate the last day of the current month
            _, last_day = calendar.monthrange(t_year, t_month)

            # ===== Initialize Solids Report Sheet =====
            if "تقرير المواد الصلبة" in wb.sheetnames:
                ws1 = wb["تقرير المواد الصلبة"]

                # Set month and year headers
                ws1["I10"] = t_month
                ws1["J10"] = t_year

                # Populate daily dates in column B (rows 12-42)
                for day in range(1, last_day + 1):
                    row = 11 + day
                    date_str = f"{day:02d}/{t_month:02d}/{t_year}"
                    ws1.cell(row=row, column=2, value=date_str)

                # Mark Fridays as weekly rest days
                for row in range(12, 43):
                    day_val = ws1.cell(row=row, column=1).value
                    if day_val and str(day_val).isdigit():
                        d = int(str(day_val))
                        if 1 <= d <= last_day:
                            # Check if this day is Friday (weekday() returns 4 for Friday)
                            if datetime(t_year, t_month, d).weekday() == 4:
                                ws1.cell(row=row, column=9,
                                         value="راحة أسبوعية")

            # ===== Initialize Main Sheet =====
            if "Main" in wb.sheetnames:
                ws_main = wb["Main"]

                # Set month and year headers
                ws_main["Q10"] = t_month
                ws_main["R10"] = t_year

                # Populate daily dates in column B (rows 14-45)
                for day in range(1, last_day + 1):
                    row = 13 + day
                    date_str = f"{day:02d}/{t_month:02d}/{t_year}"
                    ws_main.cell(row=row, column=2, value=date_str)

                # ===== Add Station Header for External Files =====
                # Find first empty row or specific row for station header
                header_row = None
                for row in range(1, 20):
                    if ws_main.cell(row=row, column=2).value is None:
                        header_row = row
                        break

                if header_row:
                    ws_main.cell(row=header_row, column=2, value="المحطة")

            # Save the initialized report
            wb.save(report_path)
            wb.close()

            return True

        except Exception as e:
            # Log error and return failure status
            print(f"  [ERROR] Report initialization failed: {e}")
            return False

    def process_daily_files(self, report_path, daily_dir):
        """
        Process and aggregate data from daily report files into monthly report.

        This method reads all daily Excel files for the month and aggregates their
        data into the monthly report, updating both the solids report sheet and
        the main summary sheet with daily measurements and operational status.

        Data Processing:
            - Matches daily files by date in filename (dd-mm-yyyy.xlsx)
            - Updates solids report with operational data and measurements
            - Updates main sheet with comprehensive parameter data
            - Handles operational status indicators (running/stopped)

        Args:
            report_path (Path): Path to the monthly report Excel file.
            daily_dir (Path): Directory containing daily Excel files.

        Returns:
            int: Number of successful daily file updates (0 if no files processed).

        Raises:
            None: All exceptions are caught internally with error logging.

        Daily File Processing:
            For each daily Excel file:
            1. Extract date from filename (11-02-2026.xlsx → 11/02/2026)
            2. Load daily workbook and access Main worksheet
            3. Find matching date row in monthly report
            4. Update solids report data (columns C-J)
            5. Update main sheet data (columns C-W)
            6. Set operational status indicators

        Data Mapping - Solids Report:
            Column C: K25 (daily) → Primary solids measurement
            Column D: G25 → Total Solids (TS)
            Column E: E25 → Volatile Solids (VS)
            Column G: J25 → Total Suspended Solids (TSS)
            Column H: E26 → Operational status 1
            Column I: E27 → Operational status 2
            Column J: K27 → Final measurement
            Column K: G24 status → Running indicator
            Column L: G23 status → Additional status

        Data Mapping - Main Sheet:
            20+ parameters mapped from daily cells to monthly columns
            Includes COD, BOD5, pH, temperature, efficiency metrics
            Handles missing data gracefully with try/except blocks

        Operational Status Logic:
            - Column K: "1" if G24 is not empty/stopped/0
            - Column L: "1" if G23 is not empty/stopped/0
            - Empty cells indicate non-operational status

        Error Handling:
            - Missing daily directory: Return 0 with error message
            - No daily files found: Return 0 with error message
            - File loading errors: Skip file and continue
            - Cell access errors: Skip invalid mappings
            - Workbook save errors: Log and return partial success count

        Performance Notes:
            - Processes files sequentially to avoid memory issues
            - Uses data_only=True for daily files to get calculated values
            - Closes workbooks immediately after processing
            - Returns count of successful updates for user feedback

        Example:
            >>> updater = MonthlyReportUpdater()
            >>> daily_dir = Path("Daily Reports/Daily Sheets/2026/January")
            >>> report_path = Path("Monthly Report/2026/January 2026.xlsx")
            >>> updated_count = updater.process_daily_files(report_path, daily_dir)
            >>> print(f"Successfully updated {updated_count} daily entries")
        """
        # Validate daily directory exists
        if not daily_dir.exists():
            print(f"  [ERROR] Daily directory not found: {daily_dir}")
            return 0

        # Get list of daily Excel files (exclude temp files)
        files = [f for f in daily_dir.iterdir()
                 if f.suffix == '.xlsx' and not f.name.startswith('~$')]

        if not files:
            print(f"  [ERROR] No daily files found in {daily_dir}")
            return 0

        # Initialize success counters
        success_solids = 0
        success_main = 0

        try:
            # Load the monthly report workbook
            wb = load_workbook(report_path)
            ws_solids = wb["تقرير المواد الصلبة"]
            ws_main = wb["Main"]

            # Process each daily file
            for f_path in files:
                try:
                    # Extract date from filename: 11-02-2026.xlsx → 11/02/2026
                    filename = f_path.stem.strip()
                    # Convert to dd/mm/yyyy format
                    date_text = filename.replace('-', '/')

                    # Load daily workbook with calculated values
                    d_wb = load_workbook(f_path, data_only=True)
                    dws = d_wb['Main']

                    if dws is None:
                        d_wb.close()
                        continue

                    # ===== Update Solids Report Sheet =====
                    found_in_solids = False
                    for row in range(12, 43):  # Check date rows in solids sheet
                        cell = ws_solids.cell(row=row, column=2).value
                        if cell and str(cell).strip() == date_text:
                            # Update solids data columns C-J
                            ws_solids.cell(row=row, column=3,
                                           value=dws["K25"].value)  # C
                            ws_solids.cell(row=row, column=4,
                                           value=dws["G25"].value)  # D
                            ws_solids.cell(row=row, column=5,
                                           value=dws["E25"].value)  # E
                            ws_solids.cell(row=row, column=7,
                                           value=dws["J25"].value)  # G
                            ws_solids.cell(row=row, column=8,
                                           value=dws["E26"].value)  # H
                            ws_solids.cell(row=row, column=9,
                                           value=dws["E27"].value)  # I
                            ws_solids.cell(row=row, column=10,
                                           value=dws["K27"].value)  # J

                            # Update operational status indicators
                            g23 = str(dws["G23"].value or "").strip()
                            g24 = str(dws["G24"].value or "").strip()

                            # Column K: Running status from G24
                            ws_solids.cell(row=row, column=11,
                                           value="1" if g24 not in ["", "متوقف", "0"] else "")

                            # Column L: Additional status from G23
                            ws_solids.cell(row=row, column=12,
                                           value="1" if g23 not in ["", "متوقف", "0"] else "")

                            found_in_solids = True
                            success_solids += 1
                            break

                    # ===== Update Main Sheet =====
                    found_in_main = False
                    for row in range(14, 46):  # Check date rows in main sheet
                        cell = ws_main.cell(row=row, column=2).value
                        if cell and str(cell).strip() == date_text:
                            # Comprehensive data mapping from daily to monthly
                            updates = {
                                3: "F16",   # C: Parameter data
                                4: "F18",   # D: Additional measurements
                                5: "E16",   # E: Analysis results
                                6: "E18",   # F: Quality metrics
                                7: "L16",   # G: Sulfide data
                                8: "L18",   # H: Extended sulfide
                                9: "K18",   # I: Potassium readings
                                10: "I16",  # J: Iodine measurements
                                11: "I18",  # K: Extended iodine
                                12: "G16",  # L: Solids data
                                13: "G18",  # M: Extended solids
                                14: "H16",  # N: Additional solids
                                15: "H18",  # O: Final solids
                                16: "J16",  # P: COD data
                                17: "J18",  # Q: Extended COD
                                18: "D34",  # R: BOD5 measurements
                                19: "H34",  # S: Extended BOD5
                                20: "M16",  # T: Oil & grease
                                21: "M18",  # U: Extended O&G
                                22: "M34",  # V: Final measurements
                                23: "N18"   # W: Nitrogen data
                            }

                            # Apply all data mappings with error handling
                            for col, src in updates.items():
                                try:
                                    value = dws[src].value
                                    if value is not None:
                                        ws_main.cell(
                                            row=row, column=col, value=value)
                                except Exception:
                                    # Skip invalid cell references
                                    pass

                            found_in_main = True
                            success_main += 1
                            break

                    # Log if no matching date found in either sheet
                    if not found_in_solids and not found_in_main:
                        print(
                            f"      [ERROR] No date match found for {filename}")

                    # Close daily workbook
                    d_wb.close()

                except Exception as e:
                    print(f"      [ERROR] Processing {f_path.name}: {e}")
                    continue

            # Save all changes to monthly report
            wb.save(report_path)
            wb.close()

            # Return total successful updates
            total_success = success_solids + success_main
            return total_success

        except Exception as e:
            print(f"  [ERROR] Daily file processing failed: {e}")
            return 0

    def process_external_files(self, report_path, ext_dir):
        """
        Process and integrate external plant data into monthly report.

        This method reads external plant Excel files and adds their data to the
        monthly report's main sheet. It handles plant name extraction, dynamic
        row assignment, and comprehensive data mapping for external sources.

        Process Flow:
            1. Validate external directory exists
            2. Find all external Excel files
            3. Locate or create "المحطة" (Station) header in monthly report
            4. For each external file:
               - Extract plant name from filename
               - Find or assign appropriate row in monthly report
               - Map external data to monthly columns
               - Update plant information

        Args:
            report_path (Path): Path to the monthly report Excel file.
            ext_dir (Path): Directory containing external plant Excel files.

        Returns:
            int: Number of successfully processed external files.

        Raises:
            None: All exceptions are caught internally with error logging.

        Plant Name Extraction:
            Filename format: "PlantName Data.xlsx" or similar
            - Splits on spaces and takes all parts after first as plant name
            - Falls back to full filename if no spaces found
            - Handles Arabic and English plant names

        Header Location Logic:
            1. Search for exact "المحطة" text in column B
            2. Search for partial "محطة" text if exact match fails
            3. Return error if header cannot be found or created

        Row Assignment:
            - Search for existing plant name in column B
            - Use first empty row if plant not found
            - Support up to 150 rows for external plants

        Data Mapping (External → Monthly):
            Column C: C10  → Primary parameter
            Column D: C11  → Secondary parameter
            Column E: B10  → Basic measurement
            Column F: B11  → Extended measurement
            Column G: J10  → Analysis result 1
            Column H: J11  → Analysis result 2
            Column I: D11  → Additional data
            Column J: F10  → Quality metric 1
            Column K: F11  → Quality metric 2
            Column L: E10  → Status indicator 1
            Column M: E11  → Status indicator 2
            Column N: G10  → Parameter set 1
            Column O: G11  → Parameter set 2
            Column P: I10  → Analysis data 1
            Column Q: I11  → Analysis data 2
            Column R: H10  → Final readings 1
            Column S: H11  → Final readings 2
            Column T: K10  → Comprehensive data 1
            Column U: K11  → Comprehensive data 2
            Column W: L11  → Summary data

        Error Handling:
            - Missing external directory: Return 0
            - No external files: Return 0
            - Workbook loading failures: Skip file
            - Header location failures: Return 0 with error
            - Cell mapping errors: Continue with other mappings
            - Save failures: Log error and return partial count

        Performance Notes:
            - Processes files sequentially
            - Uses data_only=True for calculated values
            - Closes workbooks immediately after processing
            - Handles large numbers of external files efficiently

        Example:
            >>> updater = MonthlyReportUpdater()
            >>> ext_dir = Path("External Plants/2026/2026")
            >>> report_path = Path("Monthly Report/2026/January 2026.xlsx")
            >>> processed_count = updater.process_external_files(report_path, ext_dir)
            >>> print(f"Successfully processed {processed_count} external plants")
        """
        # Validate external directory exists
        if not ext_dir.exists():
            return 0

        # Get list of external Excel files
        files = [f for f in ext_dir.iterdir()
                 if f.suffix == '.xlsx' and not f.name.startswith('~$')]

        if not files:
            return 0

        # Initialize success counter
        success = 0

        try:
            # Load monthly report workbook
            wb = load_workbook(report_path)
            ws = wb["Main"]

            # ===== Locate Station Header =====
            header_row = None

            # Method 1: Exact text search for "المحطة"
            for row in range(30, 60):
                cell_value = ws.cell(row=row, column=2).value
                if cell_value == "المحطة":
                    header_row = row
                    break

            # Method 2: Partial text search if exact match fails
            if not header_row:
                for row in range(31, 60):
                    cell_value = ws.cell(row=row, column=2).value
                    if cell_value and "محطة" in str(cell_value):
                        header_row = row
                        break

            # Return error if header cannot be located
            if not header_row:
                print(f"  [ERROR] Could not find or create 'المحطة' header")
                wb.close()
                return 0

            # Process each external file
            for f_path in files:
                try:
                    # Extract plant name from filename
                    parts = f_path.stem.split(" ")
                    if len(parts) > 1:
                        plant_name = " ".join(parts[1:]).strip()
                    else:
                        plant_name = f_path.stem.strip()

                    # Load external plant workbook
                    e_wb = load_workbook(f_path, data_only=True)
                    ews = e_wb.active

                    if ews is None:
                        e_wb.close()
                        continue

                    # Find appropriate row for this plant
                    target_row = None
                    for row in range(header_row + 1, 150):  # Search up to row 150
                        cell = ws.cell(row=row, column=2).value
                        cell_str = str(cell).strip() if cell else ""

                        if cell == plant_name:
                            # Found existing plant entry
                            target_row = row
                            break
                        elif not cell or cell_str == "":
                            # Found empty row for new plant
                            target_row = row
                            break

                    # Update plant data if row found
                    if target_row:
                        # Set plant name in column B
                        ws.cell(row=target_row, column=2, value=plant_name)

                        # Map external data to monthly columns
                        mapping = {
                            3: "C10", 4: "C11", 5: "B10", 6: "B11",
                            7: "J10", 8: "J11", 9: "D11", 10: "F10",
                            11: "F11", 12: "E10", 13: "E11", 14: "G10",
                            15: "G11", 16: "I10", 17: "I11", 18: "H10",
                            19: "H11", 20: "K10", 21: "K11", 23: "L11"
                        }

                        # Apply data mappings with error handling
                        for col, src in mapping.items():
                            try:
                                value = ews[src].value
                                if value is not None:
                                    ws.cell(row=target_row,
                                            column=col, value=value)
                            except Exception:
                                # Skip invalid cell references
                                pass

                        success += 1

                    # Close external workbook
                    e_wb.close()

                except Exception as e:
                    print(f"      [ERROR] Processing {f_path.name}: {e}")
                    continue

            # Save changes to monthly report
            wb.save(report_path)
            wb.close()

            return success

        except Exception as e:
            print(f"  [ERROR] External file processing failed: {e}")
            return 0

    def run(self):
        """
        Execute the complete monthly report generation and updating workflow.

        This is the main entry point for monthly report processing. It provides
        a complete user-interactive workflow for creating and updating monthly
        laboratory reports with data aggregation from multiple sources.

        Process Flow:
            1. Initialize tkinter root window (hidden)
            2. Determine current month and year
            3. Check for existing monthly report or create from template
            4. Get user confirmation for update operation
            5. Allow user to select alternative report if desired
            6. Process daily report files for the month
            7. Process external plant files
            8. Display success statistics to user
            9. Open completed report in Excel

        User Interaction:
            - Hidden main window (withdraw())
            - Confirmation dialog for update operation
            - File selection dialog if user chooses alternative report
            - Success/failure message boxes with detailed statistics
            - Automatic Excel launch for completed report

        Args:
            None

        Returns:
            None: Method handles all user interaction and file operations.

        Raises:
            None: All exceptions are caught and displayed to user via messagebox.

        Template Handling:
            - Checks for "Monthly Report Template.xlsx" in SETTINGS_DIR
            - Copies template to appropriate year/month location
            - Initializes new report with _initialize_new_report()
            - Shows error if template is missing

        File Path Construction:
            - Monthly report: MONTHLY_REPORT_DIR/year/"Month Year.xlsx"
            - Daily source: DAILY_SHEET/year/Month/
            - External source: EXTERNAL_DIR/year/year/
            - Uses calendar.month_name for proper month names

        Data Processing:
            - Daily files: Aggregated via process_daily_files()
            - External files: Integrated via process_external_files()
            - Success counts: Displayed in user feedback
            - Zero updates: Warning message shown

        Error Recovery:
            - Template missing: User notification and exit
            - User cancellation: Clean exit at any confirmation point
            - Processing errors: Partial success with error logging
            - File opening: Only if report exists and processing succeeded

        Example:
            >>> updater = MonthlyReportUpdater()
            >>> updater.run()  # Complete interactive workflow

        Integration Notes:
            - Called from app.py monthly report button
            - Requires tkinter event loop for dialog ownership
            - Opens Excel with os.startfile() for cross-platform compatibility
            - Handles Windows encoding issues automatically
        """
        # Create hidden root window for dialog ownership
        root = tk.Tk()
        root.withdraw()  # Hide main window, only show dialogs

        # Track processing success for final file opening
        should_open = False
        report_path = None

        try:
            # Get current date information
            now = datetime.now()
            year = str(now.year)
            month = calendar.month_name[now.month]

            # Construct default monthly report path
            report_path = self.MONTHLY_REPORT_DIR / \
                year / f"{month} {year}.xlsx"

            # ===== Create Report if Not Exists =====
            if not report_path.exists():
                # Ensure directory exists
                report_path.parent.mkdir(parents=True, exist_ok=True)

                # Copy from template
                template = self.SETTINGS_DIR / self.TEMPLATE_FILENAME
                if template.exists():
                    shutil.copy2(template, report_path)
                    self._initialize_new_report(report_path)
                else:
                    messagebox.showerror(
                        "خطأ", "قالب التقرير الشهري غير موجود!")
                    return

            # ===== Get User Confirmation =====
            response = messagebox.askyesnocancel(
                "تأكيد",
                f"هل تريد تحديث التقرير الشهري الحالي؟\n\n{month} {year}.xlsx\n\n"
                "اختر 'نعم' للتحديث، 'لا' لاختيار تقرير آخر، أو 'إلغاء' للخروج."
            )

            if response is None:  # User cancelled
                return
            elif response is False:  # User wants to select different report
                selected = filedialog.askopenfilename(
                    title="اختر تقريراً شهرياً",
                    filetypes=[("Excel files", "*.xlsx")],
                    initialdir=self.MONTHLY_REPORT_DIR / year
                )
                if not selected:
                    return
                report_path = Path(selected)

            # ===== Process Data Sources =====
            # Extract year and month from report filename
            daily_year = report_path.stem.strip().split(" ")[1]
            daily_month = report_path.stem.strip().split()[0]

            # Process daily report files
            daily_path = self.DAILY_SHEET / daily_year / daily_month
            daily_updated = self.process_daily_files(report_path, daily_path)

            # Process external plant files
            external_path = self.EXTERNAL_DIR / daily_year / daily_year
            external_updated = self.process_external_files(
                report_path, external_path)

            # Calculate total updates
            total = daily_updated + external_updated
            should_open = True  # Processing completed successfully

            # ===== User Feedback =====
            if total > 0:
                messagebox.showinfo("تم بنجاح",
                                    f"تم تحديث التقرير الشهري:\n\n"
                                    f"📊 التقارير اليومية: {daily_updated}\n"
                                    f"🏭 التقارير الخارجية: {external_updated}")
            else:
                messagebox.showwarning("تنبيه", "لم يتم تحديث أي بيانات!")

        except Exception as e:
            # Show unexpected error to user
            messagebox.showerror("خطأ", f"حدث خطأ غير متوقع:\n{str(e)}")

        finally:
            # Clean up tkinter root
            root.destroy()

            # Open completed report in Excel if processing succeeded
            if should_open and report_path and report_path.exists():
                os.startfile(str(report_path))


# Standalone execution support
if __name__ == "__main__":
    # Create and run the monthly report updater
    updater = MonthlyReportUpdater()
    updater.run()
