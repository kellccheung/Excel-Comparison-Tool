import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Any, Optional
from excel_parser import ExcelParser
from vba_analyzer import VBAAnalyzer
import difflib
import re


class ComparisonEngine:
    """Engine for comparing Excel files, formulas, and VBA code."""
    
    def __init__(self):
        self.parser1 = None
        self.parser2 = None
        self.vba_analyzer = VBAAnalyzer()
        self.comparison_results = {}
        
    def load_files(self, file1_path: str, file2_path: str) -> bool:
        """Load two Excel files for comparison."""
        try:
            self.parser1 = ExcelParser(file1_path)
            self.parser2 = ExcelParser(file2_path)
            
            if not self.parser1.load_workbook():
                print(f"Failed to load first file: {file1_path}")
                return False
                
            if not self.parser2.load_workbook():
                print(f"Failed to load second file: {file2_path}")
                return False
                
            return True
        except Exception as e:
            print(f"Error loading files: {e}")
            return False
    
    def compare_all(self) -> Dict:
        """Perform comprehensive comparison of both files."""
        if not self.parser1 or not self.parser2:
            return {}
        
        self.comparison_results = {
            'workbook_properties': self._compare_workbook_properties(),
            'sheets': self._compare_sheets(),
            'formulas': self._compare_formulas(),
            'vba_code': self._compare_vba_code(),
            'summary': {}
        }
        
        # Generate overall summary
        self.comparison_results['summary'] = self._generate_summary()
        
        return self.comparison_results
    
    def _compare_workbook_properties(self) -> Dict:
        """Compare workbook properties."""
        props1 = self.parser1.get_workbook_properties()
        props2 = self.parser2.get_workbook_properties()
        
        comparison = {
            'identical': True,
            'differences': {},
            'file1': props1,
            'file2': props2
        }
        
        # Compare key properties
        key_props = ['title', 'subject', 'creator', 'last_modified_by', 'sheets_count']
        for prop in key_props:
            if props1.get(prop) != props2.get(prop):
                comparison['identical'] = False
                comparison['differences'][prop] = {
                    'file1': props1.get(prop),
                    'file2': props2.get(prop)
                }
        
        return comparison
    
    def _compare_sheets(self) -> Dict:
        """Compare all sheets between the two files."""
        sheets1 = set(self.parser1.get_sheet_names())
        sheets2 = set(self.parser2.get_sheet_names())
        
        comparison = {
            'added_sheets': list(sheets2 - sheets1),
            'removed_sheets': list(sheets1 - sheets2),
            'common_sheets': list(sheets1 & sheets2),
            'sheet_comparisons': {}
        }
        
        # Compare common sheets
        for sheet_name in comparison['common_sheets']:
            comparison['sheet_comparisons'][sheet_name] = self._compare_single_sheet(sheet_name)
        
        return comparison
    
    def _compare_single_sheet(self, sheet_name: str) -> Dict:
        """Compare a single sheet between the two files."""
        df1 = self.parser1.get_sheet_data(sheet_name)
        df2 = self.parser2.get_sheet_data(sheet_name)
        
        # Handle empty DataFrames
        if df1.empty and df2.empty:
            return {'identical': True, 'differences': 0, 'details': {}}
        
        if df1.empty or df2.empty:
            return {
                'identical': False,
                'differences': 'One sheet is empty',
                'details': {
                    'file1_empty': df1.empty,
                    'file2_empty': df2.empty
                }
            }
        
        # Align DataFrames to same shape
        max_rows = max(len(df1), len(df2))
        max_cols = max(len(df1.columns), len(df2.columns))
        
        # Pad DataFrames to same size
        df1_padded = df1.reindex(index=range(max_rows), columns=range(max_cols), fill_value=None)
        df2_padded = df2.reindex(index=range(max_rows), columns=range(max_cols), fill_value=None)
        
        # Compare values
        differences = (df1_padded != df2_padded) & ~(df1_padded.isna() & df2_padded.isna())
        
        # Get difference locations
        diff_locations = []
        for row_idx in range(max_rows):
            for col_idx in range(max_cols):
                if differences.iloc[row_idx, col_idx]:
                    diff_locations.append({
                        'row': row_idx + 1,
                        'col': col_idx + 1,
                        'file1_value': df1_padded.iloc[row_idx, col_idx],
                        'file2_value': df2_padded.iloc[row_idx, col_idx]
                    })
        
        return {
            'identical': len(diff_locations) == 0,
            'differences': len(diff_locations),
            'details': {
                'diff_locations': diff_locations,
                'shape_file1': df1.shape,
                'shape_file2': df2.shape
            }
        }
    
    def _compare_formulas(self) -> Dict:
        """Compare formulas between the two files."""
        formulas1 = self.parser1.get_formulas()
        formulas2 = self.parser2.get_formulas()
        
        comparison = {
            'added_formulas': {},
            'removed_formulas': {},
            'modified_formulas': {},
            'identical_formulas': {},
            'summary': {}
        }
        
        # Get all sheet names
        all_sheets = set(formulas1.keys()) | set(formulas2.keys())
        
        for sheet_name in all_sheets:
            sheet_formulas1 = formulas1.get(sheet_name, {})
            sheet_formulas2 = formulas2.get(sheet_name, {})
            
            # Compare formulas in this sheet
            cell_addresses1 = set(sheet_formulas1.keys())
            cell_addresses2 = set(sheet_formulas2.keys())
            
            # Added formulas
            added_cells = cell_addresses2 - cell_addresses1
            if added_cells:
                comparison['added_formulas'][sheet_name] = {
                    cell: sheet_formulas2[cell] for cell in added_cells
                }
            
            # Removed formulas
            removed_cells = cell_addresses1 - cell_addresses2
            if removed_cells:
                comparison['removed_formulas'][sheet_name] = {
                    cell: sheet_formulas1[cell] for cell in removed_cells
                }
            
            # Modified formulas
            common_cells = cell_addresses1 & cell_addresses2
            modified_cells = {}
            identical_cells = {}
            
            for cell in common_cells:
                formula1 = sheet_formulas1[cell]['formula']
                formula2 = sheet_formulas2[cell]['formula']
                
                if formula1 != formula2:
                    modified_cells[cell] = {
                        'file1': sheet_formulas1[cell],
                        'file2': sheet_formulas2[cell]
                    }
                else:
                    identical_cells[cell] = sheet_formulas1[cell]
            
            if modified_cells:
                comparison['modified_formulas'][sheet_name] = modified_cells
            if identical_cells:
                comparison['identical_formulas'][sheet_name] = identical_cells
        
        # Generate summary
        total_added = sum(len(formulas) for formulas in comparison['added_formulas'].values())
        total_removed = sum(len(formulas) for formulas in comparison['removed_formulas'].values())
        total_modified = sum(len(formulas) for formulas in comparison['modified_formulas'].values())
        total_identical = sum(len(formulas) for formulas in comparison['identical_formulas'].values())
        
        comparison['summary'] = {
            'total_added': total_added,
            'total_removed': total_removed,
            'total_modified': total_modified,
            'total_identical': total_identical,
            'total_formulas': total_added + total_removed + total_modified + total_identical
        }
        
        return comparison
    
    def _compare_vba_code(self) -> Dict:
        """Compare VBA code between the two files."""
        vba1 = self.vba_analyzer.analyze_vba_from_file(self.parser1.file_path)
        vba2 = self.vba_analyzer.analyze_vba_from_file(self.parser2.file_path)
        
        return self.vba_analyzer.compare_vba_code(vba1, vba2)
    
    def _generate_summary(self) -> Dict:
        """Generate overall comparison summary."""
        summary = {
            'files_identical': False,
            'total_differences': 0,
            'differences_by_type': {},
            'recommendations': []
        }
        
        # Check if files are identical
        sheets_identical = (
            len(self.comparison_results['sheets']['added_sheets']) == 0 and
            len(self.comparison_results['sheets']['removed_sheets']) == 0
        )
        
        formulas_identical = (
            self.comparison_results['formulas']['summary']['total_added'] == 0 and
            self.comparison_results['formulas']['summary']['total_removed'] == 0 and
            self.comparison_results['formulas']['summary']['total_modified'] == 0
        )
        
        vba_identical = (
            self.comparison_results['vba_code']['summary']['modules_added'] == 0 and
            self.comparison_results['vba_code']['summary']['modules_removed'] == 0 and
            self.comparison_results['vba_code']['summary']['modules_modified'] == 0
        )
        
        summary['files_identical'] = (
            sheets_identical and 
            formulas_identical and 
            vba_identical and
            self.comparison_results['workbook_properties']['identical']
        )
        
        # Count total differences
        sheet_diffs = len(self.comparison_results['sheets']['added_sheets']) + len(self.comparison_results['sheets']['removed_sheets'])
        for sheet_comp in self.comparison_results['sheets']['sheet_comparisons'].values():
            if not sheet_comp['identical']:
                sheet_diffs += sheet_comp['differences'] if isinstance(sheet_comp['differences'], int) else 1
        
        formula_diffs = (
            self.comparison_results['formulas']['summary']['total_added'] +
            self.comparison_results['formulas']['summary']['total_removed'] +
            self.comparison_results['formulas']['summary']['total_modified']
        )
        
        vba_diffs = (
            self.comparison_results['vba_code']['summary']['modules_added'] +
            self.comparison_results['vba_code']['summary']['modules_removed'] +
            self.comparison_results['vba_code']['summary']['modules_modified']
        )
        
        summary['total_differences'] = sheet_diffs + formula_diffs + vba_diffs
        
        summary['differences_by_type'] = {
            'sheets': sheet_diffs,
            'formulas': formula_diffs,
            'vba_code': vba_diffs,
            'properties': len(self.comparison_results['workbook_properties']['differences'])
        }
        
        # Generate recommendations
        if sheet_diffs > 0:
            summary['recommendations'].append("Review sheet structure differences")
        if formula_diffs > 0:
            summary['recommendations'].append("Check formula changes for accuracy")
        if vba_diffs > 0:
            summary['recommendations'].append("Verify VBA code modifications")
        
        return summary
    
    def get_detailed_differences(self, sheet_name: str = None) -> Dict:
        """Get detailed differences for a specific sheet or all sheets."""
        if not self.comparison_results:
            return {}
        
        if sheet_name:
            return self.comparison_results['sheets']['sheet_comparisons'].get(sheet_name, {})
        
        return self.comparison_results
    
    def export_comparison_data(self) -> Dict:
        """Export comparison data in a structured format."""
        return {
            'metadata': {
                'file1': self.parser1.file_path if self.parser1 else None,
                'file2': self.parser2.file_path if self.parser2 else None,
                'comparison_timestamp': pd.Timestamp.now().isoformat()
            },
            'results': self.comparison_results
        }
    
    def close(self):
        """Close the parsers."""
        if self.parser1:
            self.parser1.close()
        if self.parser2:
            self.parser2.close()
