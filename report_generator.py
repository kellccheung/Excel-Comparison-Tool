import pandas as pd
import xlsxwriter
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from typing import Dict, List, Any
import os
from datetime import datetime


class ReportGenerator:
    """Generate comparison reports in various formats."""
    
    def __init__(self, comparison_data: Dict):
        self.comparison_data = comparison_data
        self.metadata = comparison_data.get('metadata', {})
        self.results = comparison_data.get('results', {})
        
    def generate_excel_report(self, output_path: str) -> bool:
        """Generate a detailed Excel report."""
        try:
            workbook = xlsxwriter.Workbook(output_path)
            
            # Create formats
            header_format = workbook.add_format({
                'bold': True,
                'font_size': 14,
                'bg_color': '#4F81BD',
                'font_color': 'white',
                'border': 1
            })
            
            subheader_format = workbook.add_format({
                'bold': True,
                'font_size': 12,
                'bg_color': '#8DB4E2',
                'border': 1
            })
            
            diff_format = workbook.add_format({
                'bg_color': '#FFC7CE',
                'font_color': '#9C0006',
                'border': 1
            })
            
            normal_format = workbook.add_format({
                'border': 1,
                'text_wrap': True
            })
            
            # Summary sheet
            self._create_summary_sheet(workbook, header_format, subheader_format, normal_format)
            
            # Sheets comparison
            self._create_sheets_comparison_sheet(workbook, header_format, subheader_format, normal_format, diff_format)
            
            # Formulas comparison
            self._create_formulas_comparison_sheet(workbook, header_format, subheader_format, normal_format, diff_format)
            
            # VBA comparison
            self._create_vba_comparison_sheet(workbook, header_format, subheader_format, normal_format, diff_format)
            
            # Detailed differences
            self._create_detailed_differences_sheet(workbook, header_format, subheader_format, normal_format, diff_format)
            
            workbook.close()
            return True
            
        except Exception as e:
            print(f"Error generating Excel report: {e}")
            return False
    
    def _create_summary_sheet(self, workbook, header_format, subheader_format, normal_format):
        """Create summary sheet."""
        worksheet = workbook.add_worksheet('Summary')
        
        # Title
        worksheet.merge_range('A1:D1', 'Excel File Comparison Report', header_format)
        worksheet.write('A3', 'Comparison Date:', subheader_format)
        worksheet.write('B3', self.metadata.get('comparison_timestamp', 'N/A'), normal_format)
        
        # File information
        worksheet.write('A5', 'File 1:', subheader_format)
        worksheet.write('B5', self.metadata.get('file1', 'N/A'), normal_format)
        worksheet.write('A6', 'File 2:', subheader_format)
        worksheet.write('B6', self.metadata.get('file2', 'N/A'), normal_format)
        
        # Overall summary
        summary = self.results.get('summary', {})
        worksheet.write('A8', 'Overall Summary:', subheader_format)
        worksheet.write('A9', 'Files Identical:', normal_format)
        worksheet.write('B9', 'Yes' if summary.get('files_identical', False) else 'No', normal_format)
        worksheet.write('A10', 'Total Differences:', normal_format)
        worksheet.write('B10', summary.get('total_differences', 0), normal_format)
        
        # Differences by type
        worksheet.write('A12', 'Differences by Type:', subheader_format)
        diff_by_type = summary.get('differences_by_type', {})
        row = 13
        for diff_type, count in diff_by_type.items():
            worksheet.write(f'A{row}', f'{diff_type.title()}:', normal_format)
            worksheet.write(f'B{row}', count, normal_format)
            row += 1
        
        # Recommendations
        recommendations = summary.get('recommendations', [])
        if recommendations:
            worksheet.write(f'A{row+1}', 'Recommendations:', subheader_format)
            for i, rec in enumerate(recommendations):
                worksheet.write(f'A{row+2+i}', f'• {rec}', normal_format)
        
        # Set column widths
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 50)
    
    def _create_sheets_comparison_sheet(self, workbook, header_format, subheader_format, normal_format, diff_format):
        """Create sheets comparison sheet."""
        worksheet = workbook.add_worksheet('Sheets Comparison')
        
        worksheet.write('A1', 'Sheets Comparison', header_format)
        
        sheets_data = self.results.get('sheets', {})
        
        # Sheet structure differences
        row = 3
        worksheet.write(f'A{row}', 'Sheet Structure:', subheader_format)
        row += 1
        
        added_sheets = sheets_data.get('added_sheets', [])
        removed_sheets = sheets_data.get('removed_sheets', [])
        
        if added_sheets:
            worksheet.write(f'A{row}', 'Added Sheets:', normal_format)
            for i, sheet in enumerate(added_sheets):
                worksheet.write(f'B{row+i}', sheet, diff_format)
            row += len(added_sheets)
        
        if removed_sheets:
            worksheet.write(f'A{row}', 'Removed Sheets:', normal_format)
            for i, sheet in enumerate(removed_sheets):
                worksheet.write(f'B{row+i}', sheet, diff_format)
            row += len(removed_sheets)
        
        # Common sheets comparison
        row += 1
        worksheet.write(f'A{row}', 'Common Sheets Comparison:', subheader_format)
        row += 1
        
        sheet_comparisons = sheets_data.get('sheet_comparisons', {})
        for sheet_name, comparison in sheet_comparisons.items():
            worksheet.write(f'A{row}', f'Sheet: {sheet_name}', subheader_format)
            worksheet.write(f'B{row}', 'Identical' if comparison.get('identical', False) else 'Different', 
                          diff_format if not comparison.get('identical', False) else normal_format)
            row += 1
            
            if not comparison.get('identical', False):
                differences = comparison.get('differences', 0)
                worksheet.write(f'A{row}', f'  Differences: {differences}', normal_format)
                row += 1
        
        worksheet.set_column('A:A', 25)
        worksheet.set_column('B:B', 30)
    
    def _create_formulas_comparison_sheet(self, workbook, header_format, subheader_format, normal_format, diff_format):
        """Create formulas comparison sheet."""
        worksheet = workbook.add_worksheet('Formulas Comparison')
        
        worksheet.write('A1', 'Formulas Comparison', header_format)
        
        formulas_data = self.results.get('formulas', {})
        summary = formulas_data.get('summary', {})
        
        # Summary
        row = 3
        worksheet.write(f'A{row}', 'Formulas Summary:', subheader_format)
        row += 1
        worksheet.write(f'A{row}', 'Total Added:', normal_format)
        worksheet.write(f'B{row}', summary.get('total_added', 0), normal_format)
        row += 1
        worksheet.write(f'A{row}', 'Total Removed:', normal_format)
        worksheet.write(f'B{row}', summary.get('total_removed', 0), normal_format)
        row += 1
        worksheet.write(f'A{row}', 'Total Modified:', normal_format)
        worksheet.write(f'B{row}', summary.get('total_modified', 0), normal_format)
        row += 1
        
        # Detailed formulas
        row += 1
        worksheet.write(f'A{row}', 'Detailed Formulas:', subheader_format)
        row += 1
        
        # Added formulas
        added_formulas = formulas_data.get('added_formulas', {})
        if added_formulas:
            worksheet.write(f'A{row}', 'Added Formulas:', subheader_format)
            row += 1
            for sheet_name, formulas in added_formulas.items():
                for cell, formula_info in formulas.items():
                    worksheet.write(f'A{row}', f'{sheet_name}!{cell}', normal_format)
                    worksheet.write(f'B{row}', formula_info.get('formula', ''), diff_format)
                    row += 1
        
        # Modified formulas
        modified_formulas = formulas_data.get('modified_formulas', {})
        if modified_formulas:
            worksheet.write(f'A{row}', 'Modified Formulas:', subheader_format)
            row += 1
            for sheet_name, formulas in modified_formulas.items():
                for cell, formula_info in formulas.items():
                    worksheet.write(f'A{row}', f'{sheet_name}!{cell}', normal_format)
                    worksheet.write(f'B{row}', f"File1: {formula_info['file1'].get('formula', '')}", normal_format)
                    worksheet.write(f'C{row}', f"File2: {formula_info['file2'].get('formula', '')}", diff_format)
                    row += 1
        
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 40)
        worksheet.set_column('C:C', 40)
    
    def _create_vba_comparison_sheet(self, workbook, header_format, subheader_format, normal_format, diff_format):
        """Create VBA comparison sheet."""
        worksheet = workbook.add_worksheet('VBA Comparison')
        
        worksheet.write('A1', 'VBA Code Comparison', header_format)
        
        vba_data = self.results.get('vba_code', {})
        summary = vba_data.get('summary', {})
        
        # Summary
        row = 3
        worksheet.write(f'A{row}', 'VBA Summary:', subheader_format)
        row += 1
        worksheet.write(f'A{row}', 'Modules Added:', normal_format)
        worksheet.write(f'B{row}', summary.get('modules_added', 0), normal_format)
        row += 1
        worksheet.write(f'A{row}', 'Modules Removed:', normal_format)
        worksheet.write(f'B{row}', summary.get('modules_removed', 0), normal_format)
        row += 1
        worksheet.write(f'A{row}', 'Modules Modified:', normal_format)
        worksheet.write(f'B{row}', summary.get('modules_modified', 0), normal_format)
        row += 1
        
        # Detailed VBA changes
        modules_data = vba_data.get('modules', {})
        if modules_data:
            row += 1
            worksheet.write(f'A{row}', 'VBA Module Changes:', subheader_format)
            row += 1
            
            for change_type in ['added', 'removed', 'modified']:
                modules = modules_data.get(change_type, [])
                if modules:
                    worksheet.write(f'A{row}', f'{change_type.title()} Modules:', normal_format)
                    for module in modules:
                        worksheet.write(f'B{row}', module, diff_format)
                        row += 1
        
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 30)
    
    def _create_detailed_differences_sheet(self, workbook, header_format, subheader_format, normal_format, diff_format):
        """Create detailed differences sheet."""
        worksheet = workbook.add_worksheet('Detailed Differences')
        
        worksheet.write('A1', 'Detailed Cell Differences', header_format)
        
        sheets_data = self.results.get('sheets', {})
        sheet_comparisons = sheets_data.get('sheet_comparisons', {})
        
        row = 3
        for sheet_name, comparison in sheet_comparisons.items():
            if not comparison.get('identical', False):
                worksheet.write(f'A{row}', f'Sheet: {sheet_name}', subheader_format)
                row += 1
                
                details = comparison.get('details', {})
                diff_locations = details.get('diff_locations', [])
                
                if diff_locations:
                    worksheet.write(f'A{row}', 'Row', normal_format)
                    worksheet.write(f'B{row}', 'Column', normal_format)
                    worksheet.write(f'C{row}', 'File 1 Value', normal_format)
                    worksheet.write(f'D{row}', 'File 2 Value', normal_format)
                    row += 1
                    
                    for diff in diff_locations:
                        worksheet.write(f'A{row}', diff.get('row', ''), normal_format)
                        worksheet.write(f'B{row}', diff.get('col', ''), normal_format)
                        worksheet.write(f'C{row}', str(diff.get('file1_value', '')), normal_format)
                        worksheet.write(f'D{row}', str(diff.get('file2_value', '')), diff_format)
                        row += 1
                
                row += 1
        
        worksheet.set_column('A:A', 15)
        worksheet.set_column('B:B', 15)
        worksheet.set_column('C:C', 30)
        worksheet.set_column('D:D', 30)
    
    def generate_pdf_report(self, output_path: str) -> bool:
        """Generate a PDF report."""
        try:
            doc = SimpleDocTemplate(output_path, pagesize=A4)
            styles = getSampleStyleSheet()
            story = []
            
            # Title
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=16,
                spaceAfter=30,
                alignment=1  # Center
            )
            story.append(Paragraph('Excel File Comparison Report', title_style))
            story.append(Spacer(1, 12))
            
            # Metadata
            story.append(Paragraph('Comparison Information', styles['Heading2']))
            story.append(Paragraph(f'Comparison Date: {self.metadata.get("comparison_timestamp", "N/A")}', styles['Normal']))
            story.append(Paragraph(f'File 1: {self.metadata.get("file1", "N/A")}', styles['Normal']))
            story.append(Paragraph(f'File 2: {self.metadata.get("file2", "N/A")}', styles['Normal']))
            story.append(Spacer(1, 12))
            
            # Summary
            summary = self.results.get('summary', {})
            story.append(Paragraph('Summary', styles['Heading2']))
            
            summary_data = [
                ['Files Identical', 'Yes' if summary.get('files_identical', False) else 'No'],
                ['Total Differences', str(summary.get('total_differences', 0))]
            ]
            
            diff_by_type = summary.get('differences_by_type', {})
            for diff_type, count in diff_by_type.items():
                summary_data.append([f'{diff_type.title()} Differences', str(count)])
            
            summary_table = Table(summary_data, colWidths=[2*inch, 1*inch])
            summary_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            story.append(summary_table)
            story.append(Spacer(1, 12))
            
            # Recommendations
            recommendations = summary.get('recommendations', [])
            if recommendations:
                story.append(Paragraph('Recommendations', styles['Heading2']))
                for rec in recommendations:
                    story.append(Paragraph(f'• {rec}', styles['Normal']))
                story.append(Spacer(1, 12))
            
            # Build PDF
            doc.build(story)
            return True
            
        except Exception as e:
            print(f"Error generating PDF report: {e}")
            return False
    
    def generate_text_report(self, output_path: str) -> bool:
        """Generate a text report."""
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write("=" * 60 + "\n")
                f.write("EXCEL FILE COMPARISON REPORT\n")
                f.write("=" * 60 + "\n\n")
                
                # Metadata
                f.write("COMPARISON INFORMATION\n")
                f.write("-" * 30 + "\n")
                f.write(f"Comparison Date: {self.metadata.get('comparison_timestamp', 'N/A')}\n")
                f.write(f"File 1: {self.metadata.get('file1', 'N/A')}\n")
                f.write(f"File 2: {self.metadata.get('file2', 'N/A')}\n\n")
                
                # Summary
                summary = self.results.get('summary', {})
                f.write("SUMMARY\n")
                f.write("-" * 10 + "\n")
                f.write(f"Files Identical: {'Yes' if summary.get('files_identical', False) else 'No'}\n")
                f.write(f"Total Differences: {summary.get('total_differences', 0)}\n\n")
                
                diff_by_type = summary.get('differences_by_type', {})
                for diff_type, count in diff_by_type.items():
                    f.write(f"{diff_type.title()} Differences: {count}\n")
                f.write("\n")
                
                # Detailed sections
                self._write_sheets_section(f)
                self._write_formulas_section(f)
                self._write_vba_section(f)
                
                # Recommendations
                recommendations = summary.get('recommendations', [])
                if recommendations:
                    f.write("RECOMMENDATIONS\n")
                    f.write("-" * 15 + "\n")
                    for rec in recommendations:
                        f.write(f"• {rec}\n")
                    f.write("\n")
                
                f.write("=" * 60 + "\n")
                f.write("END OF REPORT\n")
                f.write("=" * 60 + "\n")
            
            return True
            
        except Exception as e:
            print(f"Error generating text report: {e}")
            return False
    
    def _write_sheets_section(self, f):
        """Write sheets comparison section to text file."""
        f.write("SHEETS COMPARISON\n")
        f.write("-" * 20 + "\n")
        
        sheets_data = self.results.get('sheets', {})
        added_sheets = sheets_data.get('added_sheets', [])
        removed_sheets = sheets_data.get('removed_sheets', [])
        
        if added_sheets:
            f.write("Added Sheets:\n")
            for sheet in added_sheets:
                f.write(f"  + {sheet}\n")
        
        if removed_sheets:
            f.write("Removed Sheets:\n")
            for sheet in removed_sheets:
                f.write(f"  - {sheet}\n")
        
        sheet_comparisons = sheets_data.get('sheet_comparisons', {})
        for sheet_name, comparison in sheet_comparisons.items():
            status = "IDENTICAL" if comparison.get('identical', False) else "DIFFERENT"
            f.write(f"{sheet_name}: {status}\n")
            if not comparison.get('identical', False):
                differences = comparison.get('differences', 0)
                f.write(f"  Differences: {differences}\n")
        
        f.write("\n")
    
    def _write_formulas_section(self, f):
        """Write formulas comparison section to text file."""
        f.write("FORMULAS COMPARISON\n")
        f.write("-" * 20 + "\n")
        
        formulas_data = self.results.get('formulas', {})
        summary = formulas_data.get('summary', {})
        
        f.write(f"Total Added: {summary.get('total_added', 0)}\n")
        f.write(f"Total Removed: {summary.get('total_removed', 0)}\n")
        f.write(f"Total Modified: {summary.get('total_modified', 0)}\n\n")
        
        # Added formulas
        added_formulas = formulas_data.get('added_formulas', {})
        if added_formulas:
            f.write("Added Formulas:\n")
            for sheet_name, formulas in added_formulas.items():
                for cell, formula_info in formulas.items():
                    f.write(f"  {sheet_name}!{cell}: {formula_info.get('formula', '')}\n")
        
        # Modified formulas
        modified_formulas = formulas_data.get('modified_formulas', {})
        if modified_formulas:
            f.write("Modified Formulas:\n")
            for sheet_name, formulas in modified_formulas.items():
                for cell, formula_info in formulas.items():
                    f.write(f"  {sheet_name}!{cell}:\n")
                    f.write(f"    File1: {formula_info['file1'].get('formula', '')}\n")
                    f.write(f"    File2: {formula_info['file2'].get('formula', '')}\n")
        
        f.write("\n")
    
    def _write_vba_section(self, f):
        """Write VBA comparison section to text file."""
        f.write("VBA CODE COMPARISON\n")
        f.write("-" * 20 + "\n")
        
        vba_data = self.results.get('vba_code', {})
        summary = vba_data.get('summary', {})
        
        f.write(f"Modules Added: {summary.get('modules_added', 0)}\n")
        f.write(f"Modules Removed: {summary.get('modules_removed', 0)}\n")
        f.write(f"Modules Modified: {summary.get('modules_modified', 0)}\n\n")
        
        modules_data = vba_data.get('modules', {})
        for change_type in ['added', 'removed', 'modified']:
            modules = modules_data.get(change_type, [])
            if modules:
                f.write(f"{change_type.title()} Modules:\n")
                for module in modules:
                    f.write(f"  {module}\n")
        
        f.write("\n")
