# Excel Compare Tool

A comprehensive Python application for comparing two Excel files, including:
- Sheet data comparison
- Formula analysis
- VBA code comparison
- Visual difference highlighting
- Detailed comparison reports

## Features

- **Modern GUI Interface**: User-friendly interface built with tkinter
- **Multi-level Comparison**: Compare sheets, formulas, and VBA code
- **Visual Differences**: Highlight differences with color coding
- **Report Generation**: Export detailed comparison reports in multiple formats
- **Formula Analysis**: Detect and compare Excel formulas
- **VBA Code Comparison**: Analyze VBA modules and procedures

## Installation

1. Install Python 3.8 or higher
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Run the main application:
```bash
python excel_compare.py
```

## Project Structure

- `excel_compare.py` - Main application entry point
- `excel_parser.py` - Excel file parsing and analysis
- `comparison_engine.py` - Core comparison logic
- `gui_interface.py` - User interface components
- `report_generator.py` - Report generation utilities
- `vba_analyzer.py` - VBA code analysis module

## Supported Formats

- Excel files (.xlsx, .xlsm)
- VBA code extraction and comparison
- Formula detection and analysis
