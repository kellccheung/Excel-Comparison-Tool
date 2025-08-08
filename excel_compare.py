#!/usr/bin/env python3
"""
Excel Compare Tool - Main Application

A comprehensive tool for comparing Excel files, including:
- Sheet data comparison
- Formula analysis
- VBA code comparison
- Visual difference highlighting
- Detailed comparison reports

Author: Excel Compare Tool
Version: 1.0
"""

import tkinter as tk
import sys
import os

# Try to import tkinterdnd2 for drag and drop functionality
try:
    from tkinterdnd2 import TkinterDnD
    DRAG_DROP_AVAILABLE = True
except ImportError:
    DRAG_DROP_AVAILABLE = False
    print("Warning: tkinterdnd2 not available. Drag and drop will be disabled.")
    print("Install with: pip install tkinterdnd2")

from gui_interface import ExcelCompareGUI


def main():
    """Main application entry point."""
    try:
        # Create the main window (with drag and drop support if available)
        if DRAG_DROP_AVAILABLE:
            root = TkinterDnD.Tk()
        else:
            root = tk.Tk()
        
        # Create the GUI application
        app = ExcelCompareGUI(root)
        
        # Set up closing handler
        root.protocol("WM_DELETE_WINDOW", app.on_closing)
        
        # Center the window on screen
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f"{width}x{height}+{x}+{y}")
        
        # Start the application
        print("Excel Compare Tool started successfully!")
        print("Select two Excel files to begin comparison.")
        root.mainloop()
        
    except Exception as e:
        print(f"Error starting application: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
