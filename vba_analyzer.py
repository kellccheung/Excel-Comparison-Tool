import re
import hashlib
from typing import Dict, List, Tuple, Any
import zipfile
import xml.etree.ElementTree as ET


class VBAAnalyzer:
    """Analyzer for VBA code in Excel files."""
    
    def __init__(self):
        self.vba_modules = {}
        self.vba_procedures = {}
        self.vba_functions = {}
        self.vba_variables = {}
        
    def analyze_vba_from_file(self, file_path: str) -> Dict:
        """Analyze VBA code from an Excel file."""
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_file:
                return self._extract_vba_components(zip_file)
        except Exception as e:
            print(f"Error analyzing VBA from file: {e}")
            return {}
    
    def _extract_vba_components(self, zip_file) -> Dict:
        """Extract VBA components from the zip file."""
        vba_data = {
            'modules': {},
            'procedures': {},
            'functions': {},
            'variables': {},
            'project_info': {}
        }
        
        # Look for VBA project file
        vba_project_files = [f for f in zip_file.namelist() if 'vbaProject.bin' in f]
        if vba_project_files:
            vba_data['project_info']['has_vba'] = True
        
        # Look for VBA modules
        for file_name in zip_file.namelist():
            if 'xl/vba/' in file_name and file_name.endswith('.bin'):
                module_name = os.path.basename(file_name).replace('.bin', '')
                try:
                    module_data = zip_file.read(file_name)
                    vba_data['modules'][module_name] = {
                        'data': module_data,
                        'hash': hashlib.md5(module_data).hexdigest(),
                        'size': len(module_data)
                    }
                except:
                    continue
        
        return vba_data
    
    def parse_vba_code(self, vba_text: str) -> Dict:
        """Parse VBA code text and extract components."""
        if not vba_text:
            return {}
        
        components = {
            'procedures': [],
            'functions': [],
            'variables': [],
            'comments': [],
            'imports': []
        }
        
        lines = vba_text.split('\n')
        current_procedure = None
        current_function = None
        
        for line_num, line in enumerate(lines, 1):
            line = line.strip()
            
            # Skip empty lines
            if not line:
                continue
            
            # Check for comments
            if line.startswith("'") or line.startswith("Rem "):
                components['comments'].append({
                    'line': line_num,
                    'text': line
                })
                continue
            
            # Check for imports
            if line.startswith("Option ") or line.startswith("Imports "):
                components['imports'].append({
                    'line': line_num,
                    'text': line
                })
                continue
            
            # Check for procedure declarations
            procedure_match = re.match(r'^(Sub|Private Sub|Public Sub)\s+(\w+)', line, re.IGNORECASE)
            if procedure_match:
                current_procedure = {
                    'name': procedure_match.group(2),
                    'type': procedure_match.group(1),
                    'start_line': line_num,
                    'end_line': None,
                    'parameters': self._extract_parameters(line),
                    'body': []
                }
                components['procedures'].append(current_procedure)
                continue
            
            # Check for function declarations
            function_match = re.match(r'^(Function|Private Function|Public Function)\s+(\w+)', line, re.IGNORECASE)
            if function_match:
                current_function = {
                    'name': function_match.group(2),
                    'type': function_match.group(1),
                    'start_line': line_num,
                    'end_line': None,
                    'parameters': self._extract_parameters(line),
                    'return_type': self._extract_return_type(line),
                    'body': []
                }
                components['functions'].append(current_function)
                continue
            
            # Check for End Sub/End Function
            if line.lower() in ['end sub', 'end function']:
                if current_procedure:
                    current_procedure['end_line'] = line_num
                    current_procedure = None
                elif current_function:
                    current_function['end_line'] = line_num
                    current_function = None
                continue
            
            # Add line to current procedure/function body
            if current_procedure:
                current_procedure['body'].append(line)
            elif current_function:
                current_function['body'].append(line)
            
            # Check for variable declarations
            var_match = re.match(r'^(Dim|Private|Public|Static)\s+(\w+)', line, re.IGNORECASE)
            if var_match:
                components['variables'].append({
                    'line': line_num,
                    'declaration': var_match.group(1),
                    'name': var_match.group(2),
                    'full_line': line
                })
        
        return components
    
    def _extract_parameters(self, line: str) -> List[str]:
        """Extract parameters from a procedure/function declaration."""
        # Look for parameters in parentheses
        param_match = re.search(r'\((.*?)\)', line)
        if param_match:
            params_str = param_match.group(1)
            # Split by comma and clean up
            params = [p.strip() for p in params_str.split(',') if p.strip()]
            return params
        return []
    
    def _extract_return_type(self, line: str) -> str:
        """Extract return type from a function declaration."""
        # Look for "As Type" pattern
        return_match = re.search(r'As\s+(\w+)', line, re.IGNORECASE)
        if return_match:
            return return_match.group(1)
        return "Variant"  # Default VBA return type
    
    def compare_vba_code(self, vba1: Dict, vba2: Dict) -> Dict:
        """Compare two VBA code structures."""
        comparison = {
            'modules': self._compare_modules(vba1.get('modules', {}), vba2.get('modules', {})),
            'procedures': self._compare_procedures(vba1.get('procedures', {}), vba2.get('procedures', {})),
            'functions': self._compare_functions(vba1.get('functions', {}), vba2.get('functions', {})),
            'variables': self._compare_variables(vba1.get('variables', {}), vba2.get('variables', {})),
            'summary': {}
        }
        
        # Generate summary
        comparison['summary'] = {
            'modules_added': len(comparison['modules']['added']),
            'modules_removed': len(comparison['modules']['removed']),
            'modules_modified': len(comparison['modules']['modified']),
            'procedures_added': len(comparison['procedures']['added']),
            'procedures_removed': len(comparison['procedures']['removed']),
            'procedures_modified': len(comparison['procedures']['modified']),
            'functions_added': len(comparison['functions']['added']),
            'functions_removed': len(comparison['functions']['removed']),
            'functions_modified': len(comparison['functions']['modified'])
        }
        
        return comparison
    
    def _compare_modules(self, modules1: Dict, modules2: Dict) -> Dict:
        """Compare VBA modules."""
        keys1 = set(modules1.keys())
        keys2 = set(modules2.keys())
        
        return {
            'added': list(keys2 - keys1),
            'removed': list(keys1 - keys2),
            'modified': [key for key in keys1 & keys2 if modules1[key]['hash'] != modules2[key]['hash']],
            'unchanged': [key for key in keys1 & keys2 if modules1[key]['hash'] == modules2[key]['hash']]
        }
    
    def _compare_procedures(self, proc1: Dict, proc2: Dict) -> Dict:
        """Compare VBA procedures."""
        # This would need more sophisticated comparison logic
        # For now, return basic structure
        return {
            'added': [],
            'removed': [],
            'modified': [],
            'unchanged': []
        }
    
    def _compare_functions(self, func1: Dict, func2: Dict) -> Dict:
        """Compare VBA functions."""
        # This would need more sophisticated comparison logic
        # For now, return basic structure
        return {
            'added': [],
            'removed': [],
            'modified': [],
            'unchanged': []
        }
    
    def _compare_variables(self, var1: Dict, var2: Dict) -> Dict:
        """Compare VBA variables."""
        # This would need more sophisticated comparison logic
        # For now, return basic structure
        return {
            'added': [],
            'removed': [],
            'modified': [],
            'unchanged': []
        }
    
    def get_vba_summary(self, vba_data: Dict) -> Dict:
        """Get a summary of VBA code."""
        return {
            'has_vba': bool(vba_data.get('modules')),
            'modules_count': len(vba_data.get('modules', {})),
            'total_size': sum(module['size'] for module in vba_data.get('modules', {}).values()),
            'module_names': list(vba_data.get('modules', {}).keys())
        }


# Import os at the top level
import os
