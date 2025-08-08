import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
from typing import Dict, List, Any
from comparison_engine import ComparisonEngine
from report_generator import ReportGenerator
import pandas as pd

# Try to import tkinterdnd2 for drag and drop functionality
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DRAG_DROP_AVAILABLE = True
except ImportError:
    DRAG_DROP_AVAILABLE = False
    print("Warning: tkinterdnd2 not available. Drag and drop will be disabled.")
    print("Install with: pip install tkinterdnd2")


class ExcelCompareGUI:
    """Modern GUI for Excel file comparison."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Compare Tool")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f8fafc')  # Modern light background
        
        # Initialize components
        self.comparison_engine = ComparisonEngine()
        self.comparison_results = {}
        self.file1_path = ""
        self.file2_path = ""
        
        # Create GUI components
        self._create_widgets()
        self._setup_styles()
        
    def _setup_styles(self):
        """Setup custom styles for the GUI."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Modern color palette
        colors = {
            'primary': '#2563eb',      # Modern blue
            'secondary': '#64748b',    # Slate gray
            'success': '#059669',      # Emerald green
            'warning': '#d97706',      # Amber
            'error': '#dc2626',        # Red
            'background': '#f8fafc',   # Light gray background
            'surface': '#ffffff',      # White surface
            'border': '#e2e8f0',       # Light border
            'text': '#1e293b',         # Dark text
            'text_secondary': '#64748b' # Secondary text
        }
        
        # Configure styles with modern colors
        style.configure('Title.TLabel', 
                       font=('Segoe UI', 18, 'bold'), 
                       foreground=colors['primary'],
                       background=colors['background'])
        
        style.configure('Header.TLabel', 
                       font=('Segoe UI', 12, 'bold'),
                       foreground=colors['text'],
                       background=colors['background'])
        
        style.configure('Success.TLabel', 
                       foreground=colors['success'],
                       font=('Segoe UI', 11, 'bold'))
        
        style.configure('Error.TLabel', 
                       foreground=colors['error'],
                       font=('Segoe UI', 11, 'bold'))
        
        style.configure('Warning.TLabel', 
                       foreground=colors['warning'],
                       font=('Segoe UI', 11, 'bold'))
        
        # Modern drop zone styles
        style.configure('DropZone.TFrame', 
                       background='#eff6ff', 
                       relief='solid', 
                       borderwidth=2,
                       bordercolor=colors['border'])
        
        style.configure('DropZoneActive.TFrame', 
                       background='#dbeafe', 
                       relief='solid', 
                       borderwidth=2,
                       bordercolor=colors['primary'])
        
        # Modern button styles
        style.configure('Accent.TButton',
                       background=colors['primary'],
                       foreground='white',
                       font=('Segoe UI', 10, 'bold'),
                       borderwidth=0,
                       focuscolor='none')
        
        style.map('Accent.TButton',
                 background=[('active', '#1d4ed8'), ('pressed', '#1e40af')])
        
        # Modern frame styles
        style.configure('Modern.TFrame', background=colors['background'])
        style.configure('Surface.TFrame', background=colors['surface'])
        
        # Modern label frame styles
        style.configure('Modern.TLabelframe', 
                       background=colors['surface'],
                       bordercolor=colors['border'],
                       borderwidth=1,
                       relief='solid')
        
        style.configure('Modern.TLabelframe.Label', 
                       background=colors['surface'],
                       foreground=colors['text'],
                       font=('Segoe UI', 10, 'bold'))
        
    def _create_widgets(self):
        """Create all GUI widgets."""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10", style='Modern.TFrame')
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel File Comparison Tool", style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection section
        self._create_file_selection_section(main_frame)
        
        # Control section
        self._create_control_section(main_frame)
        
        # Results section
        self._create_results_section(main_frame)
        
        # Status bar
        self._create_status_bar(main_frame)
        
    def _create_file_selection_section(self, parent):
        """Create file selection widgets with drag and drop support."""
        # File selection frame
        file_frame = ttk.LabelFrame(parent, text="File Selection", padding="10", style='Modern.TLabelframe')
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # File 1 selection
        ttk.Label(file_frame, text="File 1:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.file1_entry = ttk.Entry(file_frame, width=60)
        self.file1_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(file_frame, text="Browse", command=self._browse_file1).grid(row=0, column=2)
        
        # File 2 selection
        ttk.Label(file_frame, text="File 2:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        self.file2_entry = ttk.Entry(file_frame, width=60)
        self.file2_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(10, 0))
        ttk.Button(file_frame, text="Browse", command=self._browse_file2).grid(row=1, column=2, pady=(10, 0))
        
        # Drag and drop zones (if available)
        if DRAG_DROP_AVAILABLE:
            self._create_drag_drop_zones(file_frame)
        
    def _create_drag_drop_zones(self, parent):
        """Create drag and drop zones for file selection."""
        # Drag and drop frame
        drop_frame = ttk.Frame(parent)
        drop_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(15, 0))
        drop_frame.columnconfigure(0, weight=1)
        drop_frame.columnconfigure(1, weight=1)
        
        # File 1 drop zone
        self.file1_drop_frame = ttk.Frame(drop_frame, style='DropZone.TFrame')
        self.file1_drop_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        
        file1_drop_label = ttk.Label(self.file1_drop_frame, text="Drop File 1 Here\nor click Browse", 
                                   style='Header.TLabel', anchor='center')
        file1_drop_label.pack(pady=20)
        
        # File 2 drop zone
        self.file2_drop_frame = ttk.Frame(drop_frame, style='DropZone.TFrame')
        self.file2_drop_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 0))
        
        file2_drop_label = ttk.Label(self.file2_drop_frame, text="Drop File 2 Here\nor click Browse", 
                                   style='Header.TLabel', anchor='center')
        file2_drop_label.pack(pady=20)
        
        # Configure drag and drop
        self._configure_drag_drop()
        
    def _configure_drag_drop(self):
        """Configure drag and drop functionality."""
        if not DRAG_DROP_AVAILABLE:
            return
            
        # File 1 drop zone
        self.file1_drop_frame.drop_target_register(DND_FILES)
        self.file1_drop_frame.dnd_bind('<<Drop>>', self._on_file1_drop)
        self.file1_drop_frame.dnd_bind('<<DropEnter>>', self._on_drop_enter)
        self.file1_drop_frame.dnd_bind('<<DropLeave>>', self._on_drop_leave)
        
        # File 2 drop zone
        self.file2_drop_frame.drop_target_register(DND_FILES)
        self.file2_drop_frame.dnd_bind('<<Drop>>', self._on_file2_drop)
        self.file2_drop_frame.dnd_bind('<<DropEnter>>', self._on_drop_enter)
        self.file2_drop_frame.dnd_bind('<<DropLeave>>', self._on_drop_leave)
        
    def _on_file1_drop(self, event):
        """Handle file drop for file 1."""
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0]
            if self._is_valid_excel_file(file_path):
                self.file1_path = file_path
                self.file1_entry.delete(0, tk.END)
                self.file1_entry.insert(0, file_path)
                self.status_var.set(f"File 1 loaded: {os.path.basename(file_path)}")
            else:
                messagebox.showerror("Error", "Please drop a valid Excel file (.xlsx or .xlsm)")
                
    def _on_file2_drop(self, event):
        """Handle file drop for file 2."""
        files = self.root.tk.splitlist(event.data)
        if files:
            file_path = files[0]
            if self._is_valid_excel_file(file_path):
                self.file2_path = file_path
                self.file2_entry.delete(0, tk.END)
                self.file2_entry.insert(0, file_path)
                self.status_var.set(f"File 2 loaded: {os.path.basename(file_path)}")
            else:
                messagebox.showerror("Error", "Please drop a valid Excel file (.xlsx or .xlsm)")
                
    def _on_drop_enter(self, event):
        """Handle drop enter event."""
        if hasattr(event.widget, 'configure'):
            event.widget.configure(style='DropZoneActive.TFrame')
            
    def _on_drop_leave(self, event):
        """Handle drop leave event."""
        if hasattr(event.widget, 'configure'):
            event.widget.configure(style='DropZone.TFrame')
            
    def _is_valid_excel_file(self, file_path):
        """Check if the file is a valid Excel file."""
        if not os.path.exists(file_path):
            return False
        file_ext = os.path.splitext(file_path)[1].lower()
        return file_ext in ['.xlsx', '.xlsm']
        
    def _create_control_section(self, parent):
        """Create control buttons."""
        control_frame = ttk.Frame(parent)
        control_frame.grid(row=2, column=0, columnspan=3, pady=(0, 10))
        
        # Compare button
        self.compare_button = ttk.Button(control_frame, text="Compare Files", 
                                       command=self._start_comparison, style='Accent.TButton')
        self.compare_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # Export buttons
        ttk.Button(control_frame, text="Export Excel Report", 
                  command=self._export_excel_report).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(control_frame, text="Export PDF Report", 
                  command=self._export_pdf_report).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(control_frame, text="Export Text Report", 
                  command=self._export_text_report).pack(side=tk.LEFT, padx=(0, 10))
        
        # Clear button
        ttk.Button(control_frame, text="Clear Results", 
                  command=self._clear_results).pack(side=tk.RIGHT)
        
    def _create_results_section(self, parent):
        """Create results display section."""
        # Results frame
        results_frame = ttk.LabelFrame(parent, text="Comparison Results", padding="10", style='Modern.TLabelframe')
        results_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(1, weight=1)
        
        # Summary frame
        summary_frame = ttk.Frame(results_frame)
        summary_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        summary_frame.columnconfigure(1, weight=1)
        
        # Summary labels
        self.summary_label = ttk.Label(summary_frame, text="No comparison performed yet", style='Header.TLabel')
        self.summary_label.grid(row=0, column=0, columnspan=2, sticky=tk.W)
        
        self.details_label = ttk.Label(summary_frame, text="")
        self.details_label.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        # Notebook for detailed results
        self.notebook = ttk.Notebook(results_frame)
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create tabs
        self._create_summary_tab()
        self._create_sheets_tab()
        self._create_formulas_tab()
        self._create_vba_tab()
        self._create_detailed_tab()
        
    def _create_summary_tab(self):
        """Create summary tab."""
        summary_frame = ttk.Frame(self.notebook)
        self.notebook.add(summary_frame, text="Summary")
        
        # Summary text widget
        self.summary_text = scrolledtext.ScrolledText(summary_frame, wrap=tk.WORD, height=20,
                                                    font=('Segoe UI', 9))
        self.summary_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
    def _create_sheets_tab(self):
        """Create sheets comparison tab."""
        sheets_frame = ttk.Frame(self.notebook)
        self.notebook.add(sheets_frame, text="Sheets")
        
        # Sheets treeview
        columns = ('Sheet', 'Status', 'Differences')
        self.sheets_tree = ttk.Treeview(sheets_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.sheets_tree.heading(col, text=col)
            self.sheets_tree.column(col, width=150)
        
        # Configure treeview style
        style = ttk.Style()
        style.configure("Treeview", 
                       background="#ffffff",
                       foreground="#1e293b",
                       fieldbackground="#ffffff",
                       font=('Segoe UI', 9))
        style.configure("Treeview.Heading", 
                       background="#f1f5f9",
                       foreground="#1e293b",
                       font=('Segoe UI', 9, 'bold'))
        
        self.sheets_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbar for treeview
        sheets_scrollbar = ttk.Scrollbar(sheets_frame, orient=tk.VERTICAL, command=self.sheets_tree.yview)
        sheets_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.sheets_tree.configure(yscrollcommand=sheets_scrollbar.set)
        
    def _create_formulas_tab(self):
        """Create formulas comparison tab."""
        formulas_frame = ttk.Frame(self.notebook)
        self.notebook.add(formulas_frame, text="Formulas")
        
        # Formulas treeview
        columns = ('Sheet', 'Cell', 'Type', 'Formula')
        self.formulas_tree = ttk.Treeview(formulas_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.formulas_tree.heading(col, text=col)
            self.formulas_tree.column(col, width=150)
        
        self.formulas_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbar for treeview
        formulas_scrollbar = ttk.Scrollbar(formulas_frame, orient=tk.VERTICAL, command=self.formulas_tree.yview)
        formulas_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.formulas_tree.configure(yscrollcommand=formulas_scrollbar.set)
        
    def _create_vba_tab(self):
        """Create VBA comparison tab."""
        vba_frame = ttk.Frame(self.notebook)
        self.notebook.add(vba_frame, text="VBA Code")
        
        # VBA treeview
        columns = ('Module', 'Type', 'Status')
        self.vba_tree = ttk.Treeview(vba_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.vba_tree.heading(col, text=col)
            self.vba_tree.column(col, width=150)
        
        self.vba_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbar for treeview
        vba_scrollbar = ttk.Scrollbar(vba_frame, orient=tk.VERTICAL, command=self.vba_tree.yview)
        vba_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.vba_tree.configure(yscrollcommand=vba_scrollbar.set)
        
    def _create_detailed_tab(self):
        """Create detailed differences tab with popup button only."""
        detailed_frame = ttk.Frame(self.notebook)
        self.notebook.add(detailed_frame, text="Detailed Differences")
        
        # Configure grid weights
        detailed_frame.columnconfigure(0, weight=1)
        detailed_frame.rowconfigure(1, weight=1)
        
        # Control frame for sheet selection
        control_frame = ttk.Frame(detailed_frame)
        control_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(control_frame, text="Select Sheet:").pack(side=tk.LEFT, padx=(0, 10))
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(control_frame, textvariable=self.sheet_var, state='readonly', width=30)
        self.sheet_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.sheet_combo.bind('<<ComboboxSelected>>', self._on_sheet_selected)
        
        ttk.Button(control_frame, text="Open Detailed View", command=self._open_detailed_popup, style='Accent.TButton').pack(side=tk.LEFT)
        
        # Differences summary frame
        diff_frame = ttk.LabelFrame(detailed_frame, text="Differences Summary", padding="10", style='Modern.TLabelframe')
        diff_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.diff_summary_text = scrolledtext.ScrolledText(diff_frame, wrap=tk.WORD, height=20,
                                                          font=('Segoe UI', 9))
        self.diff_summary_text.pack(fill=tk.BOTH, expand=True)
        
    def _on_sheet_selected(self, event=None):
        """Handle sheet selection change."""
        if not self.comparison_results or not self.sheet_var.get():
            return
            
        sheet_name = self.sheet_var.get()
        self._update_differences_summary(sheet_name)
            

        
    def _update_differences_summary(self, sheet_name):
        """Update the differences summary text."""
        if not self.comparison_results:
            return
            
        sheets_data = self.comparison_results.get('sheets', {})
        sheet_comparisons = sheets_data.get('sheet_comparisons', {})
        sheet_data = sheet_comparisons.get(sheet_name, {})
        
        summary_text = f"Sheet: {sheet_name}\n"
        summary_text += f"Status: {'Identical' if sheet_data.get('identical', False) else 'Different'}\n"
        
        # Get the correct difference count
        differences_count = sheet_data.get('differences', 0)
        if isinstance(differences_count, str):
            summary_text += f"Number of differences: {differences_count}\n\n"
        else:
            summary_text += f"Number of differences: {differences_count}\n\n"
        
        if 'details' in sheet_data and 'diff_locations' in sheet_data['details']:
            diff_locations = sheet_data['details']['diff_locations']
            summary_text += f"Differences found at (showing first 50):\n"
            for diff in diff_locations[:50]:  # Show first 50 differences
                summary_text += f"  Row {diff['row']}, Col {diff['col']}: "
                summary_text += f"'{diff['file1_value']}' vs '{diff['file2_value']}'\n"
            
            if len(diff_locations) > 50:
                summary_text += f"  ... and {len(diff_locations) - 50} more differences\n"
        else:
            summary_text += "No detailed difference information available.\n"
        
        self.diff_summary_text.delete(1.0, tk.END)
        self.diff_summary_text.insert(1.0, summary_text)
        
    def _open_detailed_popup(self):
        """Open detailed comparison in a popup window."""
        if not self.comparison_results or not self.sheet_var.get():
            messagebox.showwarning("Warning", "Please perform a comparison and select a sheet first")
            return
            
        # Create popup window
        popup = tk.Toplevel(self.root)
        popup.title(f"Detailed Comparison - {self.sheet_var.get()}")
        popup.geometry("1600x900")
        popup.configure(bg='#f8fafc')
        
        # Make popup modal
        popup.transient(self.root)
        popup.grab_set()
        
        # Center the popup
        popup.update_idletasks()
        x = (popup.winfo_screenwidth() // 2) - (800)
        y = (popup.winfo_screenheight() // 2) - (450)
        popup.geometry(f"1600x900+{x}+{y}")
        
        # Create main container
        main_frame = ttk.Frame(popup, padding="10", style='Modern.TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # Header
        header_frame = ttk.Frame(main_frame, style='Modern.TFrame')
        header_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        title_label = ttk.Label(header_frame, 
                               text=f"Detailed Comparison: {self.sheet_var.get()}", 
                               style='Title.TLabel')
        title_label.pack(side=tk.LEFT)
        
        close_button = ttk.Button(header_frame, text="Close", 
                                 command=popup.destroy, style='Accent.TButton')
        close_button.pack(side=tk.RIGHT)
        
        # File 1 spreadsheet view
        file1_frame = ttk.LabelFrame(main_frame, text=f"File 1: {os.path.basename(self.file1_path)}", 
                                   padding="5", style='Modern.TLabelframe')
        file1_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        file1_frame.columnconfigure(0, weight=1)
        file1_frame.rowconfigure(0, weight=1)
        
        # Create large spreadsheet widget for file 1
        popup_file1_spreadsheet = SpreadsheetWidget(file1_frame)
        popup_file1_spreadsheet.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # File 2 spreadsheet view
        file2_frame = ttk.LabelFrame(main_frame, text=f"File 2: {os.path.basename(self.file2_path)}", 
                                   padding="5", style='Modern.TLabelframe')
        file2_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        file2_frame.columnconfigure(0, weight=1)
        file2_frame.rowconfigure(0, weight=1)
        
        # Create large spreadsheet widget for file 2
        popup_file2_spreadsheet = SpreadsheetWidget(file2_frame)
        popup_file2_spreadsheet.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Differences summary frame
        diff_frame = ttk.LabelFrame(main_frame, text="Differences Summary", 
                                  padding="5", style='Modern.TLabelframe')
        diff_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        diff_summary_text = scrolledtext.ScrolledText(diff_frame, wrap=tk.WORD, height=10,
                                                    font=('Segoe UI', 9))
        diff_summary_text.pack(fill=tk.BOTH, expand=True)
        
        # Load data and highlight differences
        try:
            sheet_name = self.sheet_var.get()
            df1 = self.comparison_engine.parser1.get_sheet_data(sheet_name)
            df2 = self.comparison_engine.parser2.get_sheet_data(sheet_name)
            
            # Update spreadsheet widgets
            popup_file1_spreadsheet.load_data(df1, f"File 1 - {sheet_name}")
            popup_file2_spreadsheet.load_data(df2, f"File 2 - {sheet_name}")
            
            # Highlight differences
            self._highlight_differences_popup(df1, df2, popup_file1_spreadsheet, popup_file2_spreadsheet)
            
            # Update differences summary
            self._update_differences_summary_popup(sheet_name, diff_summary_text)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheet data: {str(e)}")
            
    def _highlight_differences_popup(self, df1, df2, spreadsheet1, spreadsheet2):
        """Highlight differences between the two dataframes in popup."""
        if df1 is None or df2 is None:
            return
            
        # Get the maximum dimensions
        max_rows = max(len(df1) if df1 is not None else 0, len(df2) if df2 is not None else 0)
        max_cols = max(len(df1.columns) if df1 is not None else 0, len(df2.columns) if df2 is not None else 0)
        
        # Pad dataframes to same size
        if df1 is not None:
            df1_padded = df1.reindex(index=range(max_rows), columns=range(max_cols), fill_value='')
        else:
            df1_padded = pd.DataFrame('', index=range(max_rows), columns=range(max_cols))
            
        if df2 is not None:
            df2_padded = df2.reindex(index=range(max_rows), columns=range(max_cols), fill_value='')
        else:
            df2_padded = pd.DataFrame('', index=range(max_rows), columns=range(max_cols))
        
        # Find differences
        differences = (df1_padded != df2_padded) & ~(df1_padded.isna() & df2_padded.isna())
        
        # Highlight differences in spreadsheet widgets
        diff_cells = []
        for row_idx in range(max_rows):
            for col_idx in range(max_cols):
                if differences.iloc[row_idx, col_idx]:
                    diff_cells.append((row_idx, col_idx))
        
        spreadsheet1.highlight_cells(diff_cells, '#dc2626')  # Modern red
        spreadsheet2.highlight_cells(diff_cells, '#dc2626')  # Modern red
        
    def _update_differences_summary_popup(self, sheet_name, text_widget):
        """Update the differences summary text in popup."""
        if not self.comparison_results:
            return
            
        sheets_data = self.comparison_results.get('sheets', {})
        sheet_data = sheets_data.get(sheet_name, {})
        
        summary_text = f"Sheet: {sheet_name}\n"
        summary_text += f"Status: {'Identical' if sheet_data.get('identical', False) else 'Different'}\n"
        summary_text += f"Number of differences: {sheet_data.get('differences', 0)}\n\n"
        
        if 'details' in sheet_data and 'diff_locations' in sheet_data['details']:
            summary_text += "Differences found at:\n"
            for diff in sheet_data['details']['diff_locations'][:50]:  # Show first 50 differences in popup
                summary_text += f"  Row {diff['row']}, Col {diff['col']}: "
                summary_text += f"'{diff['file1_value']}' vs '{diff['file2_value']}'\n"
            
            if len(sheet_data['details']['diff_locations']) > 50:
                summary_text += f"  ... and {len(sheet_data['details']['diff_locations']) - 50} more differences\n"
        
        text_widget.delete(1.0, tk.END)
        text_widget.insert(1.0, summary_text)
        
    def _create_status_bar(self, parent):
        """Create status bar."""
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        
        status_bar = ttk.Label(parent, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W,
                              background='#f1f5f9', foreground='#64748b', font=('Segoe UI', 9))
        status_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E))
        
    def _browse_file1(self):
        """Browse for first Excel file."""
        filename = filedialog.askopenfilename(
            title="Select First Excel File",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if filename:
            self.file1_path = filename
            self.file1_entry.delete(0, tk.END)
            self.file1_entry.insert(0, filename)
            
    def _browse_file2(self):
        """Browse for second Excel file."""
        filename = filedialog.askopenfilename(
            title="Select Second Excel File",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if filename:
            self.file2_path = filename
            self.file2_entry.delete(0, tk.END)
            self.file2_entry.insert(0, filename)
            
    def _start_comparison(self):
        """Start the comparison process in a separate thread."""
        if not self.file1_path or not self.file2_path:
            messagebox.showerror("Error", "Please select both Excel files")
            return
            
        if not os.path.exists(self.file1_path) or not os.path.exists(self.file2_path):
            messagebox.showerror("Error", "One or both files do not exist")
            return
        
        # Disable compare button and show progress
        self.compare_button.config(state='disabled')
        self.status_var.set("Comparing files... Please wait.")
        
        # Start comparison in separate thread
        thread = threading.Thread(target=self._perform_comparison)
        thread.daemon = True
        thread.start()
        
    def _perform_comparison(self):
        """Perform the actual comparison."""
        try:
            # Load files
            if not self.comparison_engine.load_files(self.file1_path, self.file2_path):
                self.root.after(0, lambda: messagebox.showerror("Error", "Failed to load one or both files"))
                return
            
            # Perform comparison
            self.comparison_results = self.comparison_engine.compare_all()
            
            # Update GUI with results
            self.root.after(0, self._update_results_display)
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Comparison failed: {str(e)}"))
        finally:
            self.root.after(0, self._comparison_finished)
            
    def _comparison_finished(self):
        """Called when comparison is finished."""
        self.compare_button.config(state='normal')
        self.status_var.set("Comparison completed")
        
    def _update_results_display(self):
        """Update the GUI with comparison results."""
        if not self.comparison_results:
            return
            
        # Update summary
        summary = self.comparison_results.get('summary', {})
        files_identical = summary.get('files_identical', False)
        total_differences = summary.get('total_differences', 0)
        
        if files_identical:
            self.summary_label.config(text="Files are IDENTICAL", style='Success.TLabel')
        else:
            self.summary_label.config(text=f"Files are DIFFERENT - {total_differences} differences found", 
                                    style='Error.TLabel')
        
        # Update details
        diff_by_type = summary.get('differences_by_type', {})
        details_text = []
        for diff_type, count in diff_by_type.items():
            if count > 0:
                details_text.append(f"{diff_type.title()}: {count}")
        
        if details_text:
            self.details_label.config(text=" | ".join(details_text))
        
        # Update summary tab
        self._update_summary_tab()
        
        # Update sheets tab
        self._update_sheets_tab()
        
        # Update formulas tab
        self._update_formulas_tab()
        
        # Update VBA tab
        self._update_vba_tab()
        
        # Update detailed differences tab
        self._update_detailed_tab()
        
    def _update_summary_tab(self):
        """Update summary tab content."""
        self.summary_text.delete(1.0, tk.END)
        
        if not self.comparison_results:
            return
            
        summary_text = "EXCEL FILE COMPARISON REPORT\n"
        summary_text += "=" * 50 + "\n\n"
        
        # Metadata
        summary_text += "COMPARISON INFORMATION\n"
        summary_text += "-" * 25 + "\n"
        summary_text += f"File 1: {self.file1_path}\n"
        summary_text += f"File 2: {self.file2_path}\n\n"
        
        # Overall summary
        summary = self.comparison_results.get('summary', {})
        summary_text += "OVERALL SUMMARY\n"
        summary_text += "-" * 15 + "\n"
        summary_text += f"Files Identical: {'Yes' if summary.get('files_identical', False) else 'No'}\n"
        summary_text += f"Total Differences: {summary.get('total_differences', 0)}\n\n"
        
        # Differences by type
        diff_by_type = summary.get('differences_by_type', {})
        summary_text += "DIFFERENCES BY TYPE\n"
        summary_text += "-" * 20 + "\n"
        for diff_type, count in diff_by_type.items():
            summary_text += f"{diff_type.title()}: {count}\n"
        
        # Recommendations
        recommendations = summary.get('recommendations', [])
        if recommendations:
            summary_text += "\nRECOMMENDATIONS\n"
            summary_text += "-" * 15 + "\n"
            for rec in recommendations:
                summary_text += f"• {rec}\n"
        
        self.summary_text.insert(1.0, summary_text)
        
    def _update_sheets_tab(self):
        """Update sheets tab content."""
        # Clear existing items
        for item in self.sheets_tree.get_children():
            self.sheets_tree.delete(item)
        
        sheets_data = self.comparison_results.get('sheets', {})
        
        # Add added sheets
        for sheet in sheets_data.get('added_sheets', []):
            self.sheets_tree.insert('', 'end', values=(sheet, 'ADDED', '-'))
        
        # Add removed sheets
        for sheet in sheets_data.get('removed_sheets', []):
            self.sheets_tree.insert('', 'end', values=(sheet, 'REMOVED', '-'))
        
        # Add common sheets
        sheet_comparisons = sheets_data.get('sheet_comparisons', {})
        for sheet_name, comparison in sheet_comparisons.items():
            status = 'IDENTICAL' if comparison.get('identical', False) else 'DIFFERENT'
            differences = comparison.get('differences', 0)
            self.sheets_tree.insert('', 'end', values=(sheet_name, status, differences))
            
    def _update_formulas_tab(self):
        """Update formulas tab content."""
        # Clear existing items
        for item in self.formulas_tree.get_children():
            self.formulas_tree.delete(item)
        
        formulas_data = self.comparison_results.get('formulas', {})
        
        # Add added formulas
        added_formulas = formulas_data.get('added_formulas', {})
        for sheet_name, formulas in added_formulas.items():
            for cell, formula_info in formulas.items():
                self.formulas_tree.insert('', 'end', values=(
                    sheet_name, cell, 'ADDED', formula_info.get('formula', '')
                ))
        
        # Add modified formulas
        modified_formulas = formulas_data.get('modified_formulas', {})
        for sheet_name, formulas in modified_formulas.items():
            for cell, formula_info in formulas.items():
                self.formulas_tree.insert('', 'end', values=(
                    sheet_name, cell, 'MODIFIED', 
                    f"File1: {formula_info['file1'].get('formula', '')} | File2: {formula_info['file2'].get('formula', '')}"
                ))
        
    def _update_vba_tab(self):
        """Update VBA tab content."""
        # Clear existing items
        for item in self.vba_tree.get_children():
            self.vba_tree.delete(item)
        
        vba_data = self.comparison_results.get('vba_code', {})
        modules_data = vba_data.get('modules', {})
        
        # Add module changes
        for change_type in ['added', 'removed', 'modified']:
            modules = modules_data.get(change_type, [])
            for module in modules:
                self.vba_tree.insert('', 'end', values=(module, 'Module', change_type.upper()))
                
    def _update_detailed_tab(self):
        """Update detailed differences tab content."""
        # Populate sheet combo box
        sheets_data = self.comparison_results.get('sheets', {})
        sheet_comparisons = sheets_data.get('sheet_comparisons', {})
        
        # Get all sheet names
        sheet_names = list(sheet_comparisons.keys())
        
        # Update combo box
        self.sheet_combo['values'] = sheet_names
        if sheet_names:
            self.sheet_combo.set(sheet_names[0])
            # Load first sheet data
            self._update_differences_summary(sheet_names[0])
        else:
            self.sheet_combo.set('')
            self.diff_summary_text.delete(1.0, tk.END)
                    
    def _export_excel_report(self):
        """Export Excel report."""
        if not self.comparison_results:
            messagebox.showwarning("Warning", "No comparison results to export")
            return
            
        filename = filedialog.asksaveasfilename(
            title="Save Excel Report",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if filename:
            try:
                export_data = self.comparison_engine.export_comparison_data()
                report_gen = ReportGenerator(export_data)
                if report_gen.generate_excel_report(filename):
                    messagebox.showinfo("Success", f"Excel report saved to {filename}")
                else:
                    messagebox.showerror("Error", "Failed to generate Excel report")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {str(e)}")
                
    def _export_pdf_report(self):
        """Export PDF report."""
        if not self.comparison_results:
            messagebox.showwarning("Warning", "No comparison results to export")
            return
            
        filename = filedialog.asksaveasfilename(
            title="Save PDF Report",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if filename:
            try:
                export_data = self.comparison_engine.export_comparison_data()
                report_gen = ReportGenerator(export_data)
                if report_gen.generate_pdf_report(filename):
                    messagebox.showinfo("Success", f"PDF report saved to {filename}")
                else:
                    messagebox.showerror("Error", "Failed to generate PDF report")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {str(e)}")
                
    def _export_text_report(self):
        """Export text report."""
        if not self.comparison_results:
            messagebox.showwarning("Warning", "No comparison results to export")
            return
            
        filename = filedialog.asksaveasfilename(
            title="Save Text Report",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt")]
        )
        
        if filename:
            try:
                export_data = self.comparison_engine.export_comparison_data()
                report_gen = ReportGenerator(export_data)
                if report_gen.generate_text_report(filename):
                    messagebox.showinfo("Success", f"Text report saved to {filename}")
                else:
                    messagebox.showerror("Error", "Failed to generate text report")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {str(e)}")
                
    def _clear_results(self):
        """Clear all results and reset the interface."""
        self.comparison_results = {}
        self.summary_label.config(text="No comparison performed yet", style='Header.TLabel')
        self.details_label.config(text="")
        
        # Clear all treeviews
        for tree in [self.sheets_tree, self.formulas_tree, self.vba_tree]:
            for item in tree.get_children():
                tree.delete(item)
        

        
        # Clear sheet combo box
        if hasattr(self, 'sheet_combo'):
            self.sheet_combo['values'] = []
            self.sheet_combo.set('')
        
        # Clear differences summary
        if hasattr(self, 'diff_summary_text'):
            self.diff_summary_text.delete(1.0, tk.END)
        
        # Clear summary text
        self.summary_text.delete(1.0, tk.END)
        
        self.status_var.set("Ready")
        
    def on_closing(self):
        """Handle application closing."""
        if self.comparison_engine:
            self.comparison_engine.close()
        self.root.destroy()


class SpreadsheetWidget(ttk.Frame):
    """Custom widget to display Excel data in a grid format."""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.data = None
        self.cells = {}
        self.highlighted_cells = set()
        
        # Configure grid
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        
        # Create canvas and scrollbars
        self.canvas = tk.Canvas(self, bg='#ffffff')
        self.v_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        self.h_scrollbar = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.canvas.xview)
        
        # Configure canvas
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)
        
        # Grid layout
        self.canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Create frame inside canvas for content
        self.content_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.content_frame, anchor='nw')
        
        # Bind events
        self.content_frame.bind('<Configure>', self._on_frame_configure)
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        
    def _on_frame_configure(self, event=None):
        """Update canvas scroll region when frame size changes."""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
    def _on_canvas_configure(self, event):
        """Update canvas window width when canvas is resized."""
        self.canvas.itemconfig(self.canvas.find_withtag("all")[0], width=event.width)
        
    def load_data(self, df, title=""):
        """Load data into the spreadsheet widget."""
        # Clear existing content
        for widget in self.content_frame.winfo_children():
            widget.destroy()
        self.cells.clear()
        self.highlighted_cells.clear()
        
        if df is None or df.empty:
            # Show empty message
            empty_label = ttk.Label(self.content_frame, text="No data available", 
                                  font=('Segoe UI', 12, 'italic'),
                                  foreground='#64748b')
            empty_label.grid(row=0, column=0, padx=10, pady=10)
            return
        
        # Configure grid
        rows, cols = df.shape
        for i in range(rows + 1):  # +1 for header row
            self.content_frame.rowconfigure(i, weight=1)
        for j in range(cols + 1):  # +1 for row numbers
            self.content_frame.columnconfigure(j, weight=1)
        
        # Create header row (column letters)
        for j, col in enumerate(df.columns):
            header_label = ttk.Label(self.content_frame, text=str(col), 
                                   relief='solid', borderwidth=1, 
                                   background='#f1f5f9', font=('Segoe UI', 9, 'bold'),
                                   foreground='#1e293b')
            header_label.grid(row=0, column=j+1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=1, pady=1)
        
        # Create row numbers
        for i in range(rows):
            row_label = ttk.Label(self.content_frame, text=str(i+1), 
                                relief='solid', borderwidth=1, 
                                background='#f1f5f9', font=('Segoe UI', 9, 'bold'),
                                foreground='#1e293b')
            row_label.grid(row=i+1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=1, pady=1)
        
        # Create data cells
        for i in range(rows):
            for j in range(cols):
                value = df.iloc[i, j]
                if pd.isna(value):
                    value = ""
                else:
                    value = str(value)
                
                cell_label = ttk.Label(self.content_frame, text=value, 
                                     relief='solid', borderwidth=1, 
                                     font=('Segoe UI', 9), wraplength=120,
                                     foreground='#1e293b', background='#ffffff')
                cell_label.grid(row=i+1, column=j+1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=1, pady=1)
                
                # Store cell reference
                self.cells[(i, j)] = cell_label
        
        # Update scroll region
        self.content_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
    def highlight_cells(self, cell_positions, color):
        """Highlight specific cells with a color."""
        for row, col in cell_positions:
            if (row, col) in self.cells:
                cell = self.cells[(row, col)]
                cell.configure(background=color, foreground='white')
                self.highlighted_cells.add((row, col))
                
    def clear_highlights(self):
        """Clear all cell highlights."""
        for row, col in self.highlighted_cells:
            if (row, col) in self.cells:
                cell = self.cells[(row, col)]
                cell.configure(background='#ffffff', foreground='#1e293b')
        self.highlighted_cells.clear()
