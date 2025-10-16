"""
POWER BI COMPREHENSIVE ANALYSIS TOOLKIT
=======================================
This script generates:
1. Data Dictionary (Excel)
2. HTML Documentation
3. DAX Dependencies Analysis
4. Model Validation Report
5. All DAX Formulas Export
"""

import json
import pandas as pd
import os
from datetime import datetime
from collections import defaultdict
import re

class PowerBIAnalyzer:
    """Comprehensive Power BI model analyzer."""
    
    def __init__(self, json_file='pbix_analysis/complete_analysis.json', output_dir='analysis_output'):
        self.json_file = json_file
        self.output_dir = output_dir
        self.data = None
        
        # Create output directory
        os.makedirs(output_dir, exist_ok=True)
    
    # ========== FIX: ADD THIS NEW METHOD ==========
    def _safe_get_expression(self, obj, key='expression'):
        """Safely get expression - handle both string and list."""
        expr = obj.get(key, '')
        
        # Handle different expression formats
        if isinstance(expr, list):
            # If it's a list, join with newlines
            return '\n'.join(str(item) for item in expr)
        elif isinstance(expr, str):
            return expr
        elif expr is None:
            return ''
        else:
            return str(expr)
    # ==============================================
    
    def load_data(self):
        """Load the JSON file."""
        print("=" * 80)
        print("POWER BI COMPREHENSIVE ANALYSIS TOOLKIT")
        print("=" * 80)
        print(f"Input: {self.json_file}")
        print(f"Output Directory: {self.output_dir}\n")
        
        if not os.path.exists(self.json_file):
            print(f"❌ Error: File not found: {self.json_file}")
            return False
        
        try:
            with open(self.json_file, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
            print("✓ JSON file loaded successfully\n")
            return True
        except Exception as e:
            print(f"❌ Error reading JSON file: {e}")
            return False
    
    def run_all(self):
        """Run all analyses."""
        if not self.load_data():
            return
        
        if 'dataModel' not in self.data:
            print("❌ Error: No data model found in the JSON file")
            return
        
        print("Running all analyses...\n")
        
        self.create_data_dictionary()
        self.create_html_documentation()
        self.analyze_dax_dependencies()
        self.validate_model()
        self.export_all_dax_formulas()
        
        print("\n" + "=" * 80)
        print("✅ ALL ANALYSES COMPLETE!")
        print("=" * 80)
        print(f"\nAll output files are in: {self.output_dir}/")
        print("\nGenerated files:")
        for file in os.listdir(self.output_dir):
            if os.path.isfile(os.path.join(self.output_dir, file)):
                size = os.path.getsize(os.path.join(self.output_dir, file))
                print(f"   • {file} ({size:,} bytes)")
    
    def create_data_dictionary(self):
        """Generate Excel data dictionary."""
        print("=" * 80)
        print("1. CREATING DATA DICTIONARY (EXCEL)")
        print("=" * 80)
        
        output_file = os.path.join(self.output_dir, 'Data_Dictionary.xlsx')
        
        tables_list = []
        columns_list = []
        measures_list = []
        relationships_list = []
        hierarchies_list = []
        roles_list = []
        
        # Process Tables
        print("📊 Processing Tables...")
        for table in self.data['dataModel'].get('tables', []):
            total_columns = len(table.get('columns', []))
            visible_columns = sum(1 for col in table.get('columns', []) if not col.get('isHidden', False))
            calculated_columns = sum(1 for col in table.get('columns', []) if col.get('expression'))
            
            tables_list.append({
                'Table Name': table['name'],
                'Description': table.get('description', ''),
                'Is Hidden': 'Yes' if table.get('isHidden', False) else 'No',
                'Total Columns': total_columns,
                'Visible Columns': visible_columns,
                'Calculated Columns': calculated_columns,
                'Measures': len(table.get('measures', [])),
                'Hierarchies': len(table.get('hierarchies', []))
            })
        
        # Process Columns
        print("📋 Processing Columns...")
        for table in self.data['dataModel'].get('tables', []):
            for col in table.get('columns', []):
                # ========== FIX: USE SAFE EXPRESSION GETTER ==========
                columns_list.append({
                    'Table': table['name'],
                    'Column Name': col['name'],
                    'Data Type': col.get('dataType', ''),
                    'Source Column': col.get('sourceColumn', ''),
                    'Format String': col.get('formatString', ''),
                    'Is Hidden': 'Yes' if col.get('isHidden', False) else 'No',
                    'Is Key': 'Yes' if col.get('isKey', False) else 'No',
                    'Is Calculated': 'Yes' if col.get('expression') else 'No',
                    'DAX Expression': self._safe_get_expression(col, 'expression'),
                    'Description': col.get('description', ''),
                    'Data Category': col.get('dataCategory', ''),
                    'Summarize By': col.get('summarizeBy', 'default'),
                    'Sort By Column': col.get('sortByColumn', ''),
                    'Display Folder': col.get('displayFolder', '')
                })
                # =====================================================
        
        # Process Measures
        print("📈 Processing Measures...")
        for table in self.data['dataModel'].get('tables', []):
            for measure in table.get('measures', []):
                # ========== FIX: USE SAFE EXPRESSION GETTER ==========
                measures_list.append({
                    'Table': table['name'],
                    'Measure Name': measure['name'],
                    'Display Folder': measure.get('displayFolder', ''),
                    'Format String': measure.get('formatString', ''),
                    'Is Hidden': 'Yes' if measure.get('isHidden', False) else 'No',
                    'DAX Expression': self._safe_get_expression(measure, 'expression'),
                    'Description': measure.get('description', '')
                })
                # =====================================================
        
        # Process Relationships
        print("🔗 Processing Relationships...")
        for rel in self.data['dataModel'].get('relationships', []):
            relationships_list.append({
                'Relationship Name': rel.get('name', ''),
                'From Table': rel['fromTable'],
                'From Column': rel['fromColumn'],
                'To Table': rel['toTable'],
                'To Column': rel['toColumn'],
                'From Cardinality': rel.get('fromCardinality', ''),
                'To Cardinality': rel.get('toCardinality', ''),
                'Full Cardinality': f"{rel.get('fromCardinality', '?')}:{rel.get('toCardinality', '?')}",
                'Cross Filter Behavior': rel.get('crossFilteringBehavior', ''),
                'Is Active': 'Yes' if rel.get('isActive', True) else 'No',
                'Security Filtering': rel.get('securityFilteringBehavior', '')
            })
        
        # Process Hierarchies
        print("📐 Processing Hierarchies...")
        for table in self.data['dataModel'].get('tables', []):
            for hierarchy in table.get('hierarchies', []):
                levels = ' → '.join([level['name'] for level in hierarchy.get('levels', [])])
                hierarchies_list.append({
                    'Table': table['name'],
                    'Hierarchy Name': hierarchy['name'],
                    'Is Hidden': 'Yes' if hierarchy.get('isHidden', False) else 'No',
                    'Level Count': len(hierarchy.get('levels', [])),
                    'Levels': levels
                })
        
        # Process Security Roles
        print("🔒 Processing Security Roles...")
        for role in self.data['dataModel'].get('roles', []):
            for perm in role.get('tablePermissions', []):
                # ========== FIX: USE SAFE EXPRESSION GETTER ==========
                roles_list.append({
                    'Role Name': role['name'],
                    'Role Description': role.get('description', ''),
                    'Table': perm['name'],
                    'Filter Expression': self._safe_get_expression(perm, 'filterExpression')
                })
                # =====================================================
        
        # Summary
        summary_data = []
        model_info = self.data['dataModel']
        summary_stats = model_info.get('summary', {})
        
        summary_data.append({'Metric': 'Model Name', 'Value': model_info.get('name', 'N/A')})
        summary_data.append({'Metric': 'Culture', 'Value': model_info.get('culture', 'N/A')})
        summary_data.append({'Metric': 'Total Tables', 'Value': summary_stats.get('totalTables', 0)})
        summary_data.append({'Metric': 'Total Measures', 'Value': summary_stats.get('totalMeasures', 0)})
        summary_data.append({'Metric': 'Total Relationships', 'Value': summary_stats.get('totalRelationships', 0)})
        summary_data.append({'Metric': 'Calculated Columns', 'Value': summary_stats.get('totalCalculatedColumns', 0)})
        summary_data.append({'Metric': 'Calculated Tables', 'Value': summary_stats.get('totalCalculatedTables', 0)})
        summary_data.append({'Metric': 'Security Roles', 'Value': summary_stats.get('totalRoles', 0)})
        summary_data.append({'Metric': 'Generated On', 'Value': datetime.now().strftime('%Y-%m-%d %H:%M:%S')})
        
        # Create Excel
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                if summary_data:
                    pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
                if tables_list:
                    pd.DataFrame(tables_list).to_excel(writer, sheet_name='Tables', index=False)
                if columns_list:
                    pd.DataFrame(columns_list).to_excel(writer, sheet_name='Columns', index=False)
                if measures_list:
                    pd.DataFrame(measures_list).to_excel(writer, sheet_name='Measures', index=False)
                if relationships_list:
                    pd.DataFrame(relationships_list).to_excel(writer, sheet_name='Relationships', index=False)
                if hierarchies_list:
                    pd.DataFrame(hierarchies_list).to_excel(writer, sheet_name='Hierarchies', index=False)
                if roles_list:
                    pd.DataFrame(roles_list).to_excel(writer, sheet_name='Security Roles', index=False)
            
            print(f"✅ Data Dictionary created: {output_file}\n")
        except Exception as e:
            print(f"❌ Error creating Excel: {e}\n")
    
    def create_html_documentation(self):
        """Generate interactive HTML documentation."""
        print("=" * 80)
        print("2. CREATING HTML DOCUMENTATION")
        print("=" * 80)
        
        output_file = os.path.join(self.output_dir, 'Model_Documentation.html')
        
        html = """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Power BI Model Documentation</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            background: #f5f5f5; 
            color: #333;
        }
        .header {
            background: linear-gradient(135deg, #0078d4 0%, #106ebe 100%);
            color: white;
            padding: 40px 20px;
            text-align: center;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; }
        .header p { font-size: 1.1em; opacity: 0.9; }
        .container { 
            max-width: 1400px; 
            margin: 0 auto; 
            padding: 30px 20px;
        }
        .stats {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 30px 0;
        }
        .stat-box {
            background: white;
            padding: 25px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            text-align: center;
            border-top: 4px solid #0078d4;
        }
        .stat-number {
            font-size: 2.5em;
            color: #0078d4;
            font-weight: bold;
        }
        .stat-label {
            color: #666;
            margin-top: 8px;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        .nav {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin: 20px 0;
            position: sticky;
            top: 20px;
            z-index: 100;
        }
        .nav h3 { 
            color: #0078d4; 
            margin-bottom: 15px;
            font-size: 1.2em;
        }
        .nav-links {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
        }
        .nav-link {
            display: inline-block;
            padding: 8px 16px;
            background: #f0f0f0;
            color: #0078d4;
            text-decoration: none;
            border-radius: 4px;
            transition: all 0.3s;
            font-size: 0.9em;
        }
        .nav-link:hover {
            background: #0078d4;
            color: white;
        }
        .section {
            background: white;
            margin: 30px 0;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .section h2 {
            color: #0078d4;
            border-bottom: 3px solid #0078d4;
            padding-bottom: 15px;
            margin-bottom: 25px;
            font-size: 1.8em;
        }
        .table-card {
            background: #f9f9f9;
            border-left: 4px solid #0078d4;
            padding: 20px;
            margin: 20px 0;
            border-radius: 4px;
        }
        .table-name {
            font-size: 1.4em;
            font-weight: bold;
            color: #0078d4;
            margin-bottom: 10px;
        }
        .table-meta {
            color: #666;
            font-size: 0.9em;
            margin-bottom: 15px;
        }
        .subsection {
            margin: 20px 0;
        }
        .subsection h4 {
            color: #333;
            margin: 15px 0 10px 0;
            padding-bottom: 8px;
            border-bottom: 1px solid #ddd;
        }
        .item {
            margin: 10px 0;
            padding: 12px;
            background: white;
            border-left: 3px solid #ccc;
            border-radius: 3px;
        }
        .item-name {
            font-weight: bold;
            color: #333;
            margin-bottom: 5px;
        }
        .item-detail {
            color: #666;
            font-size: 0.9em;
            margin: 3px 0;
        }
        .dax {
            font-family: 'Consolas', 'Courier New', monospace;
            background: #f5f5f5;
            padding: 12px;
            margin: 8px 0;
            overflow-x: auto;
            white-space: pre-wrap;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 0.9em;
            line-height: 1.5;
        }
        .measure {
            background: #e8f4f8;
            border-left-color: #0078d4;
        }
        .calculated {
            background: #fff4e6;
            border-left-color: #ff8c00;
        }
        .hidden {
            color: #999;
            font-style: italic;
        }
        .badge {
            display: inline-block;
            padding: 3px 8px;
            border-radius: 3px;
            font-size: 0.8em;
            font-weight: bold;
            margin-left: 8px;
        }
        .badge-hidden { background: #ddd; color: #666; }
        .badge-calculated { background: #ff8c00; color: white; }
        .badge-key { background: #28a745; color: white; }
        .relationship {
            padding: 12px;
            margin: 8px 0;
            background: #fff4e6;
            border-left: 3px solid #ff8c00;
            border-radius: 3px;
        }
        .rel-active { border-left-color: #28a745; background: #e8f8e8; }
        .footer {
            text-align: center;
            padding: 30px;
            color: #666;
            font-size: 0.9em;
        }
        @media print {
            .nav { display: none; }
            .section { page-break-inside: avoid; }
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 Power BI Model Documentation</h1>
        <p>Generated on """ + datetime.now().strftime('%B %d, %Y at %H:%M:%S') + """</p>
    </div>
    
    <div class="container">
"""
        
        # Statistics
        if 'dataModel' in self.data:
            summary = self.data['dataModel'].get('summary', {})
            html += """
        <div class="stats">
            <div class="stat-box">
                <div class="stat-number">{}</div>
                <div class="stat-label">Tables</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">{}</div>
                <div class="stat-label">Measures</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">{}</div>
                <div class="stat-label">Relationships</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">{}</div>
                <div class="stat-label">Calculated Columns</div>
            </div>
        </div>
""".format(
                summary.get('totalTables', 0),
                summary.get('totalMeasures', 0),
                summary.get('totalRelationships', 0),
                summary.get('totalCalculatedColumns', 0)
            )
            
            # Navigation
            html += '<div class="nav"><h3>📑 Quick Navigation</h3><div class="nav-links">'
            for table in self.data['dataModel'].get('tables', []):
                table_id = table['name'].replace(' ', '_').replace("'", "").replace('[', '').replace(']', '')
                html += f'<a href="#{table_id}" class="nav-link">{table["name"]}</a>'
            html += '</div></div>'
            
            # Tables Section
            html += '<div class="section"><h2>📋 Tables and Columns</h2>'
            
            for table in self.data['dataModel'].get('tables', []):
                table_id = table['name'].replace(' ', '_').replace("'", "").replace('[', '').replace(']', '')
                hidden_badge = '<span class="badge badge-hidden">HIDDEN</span>' if table.get('isHidden') else ''
                
                html += f'<div class="table-card" id="{table_id}">'
                html += f'<div class="table-name">{table["name"]}{hidden_badge}</div>'
                
                if table.get('description'):
                    html += f'<p class="table-meta"><em>{table["description"]}</em></p>'
                
                col_count = len(table.get('columns', []))
                meas_count = len(table.get('measures', []))
                html += f'<div class="table-meta">{col_count} columns • {meas_count} measures</div>'
                
                # Columns
                if table.get('columns'):
                    html += '<div class="subsection"><h4>Columns</h4>'
                    for col in table['columns']:
                        hidden_badge = '<span class="badge badge-hidden">HIDDEN</span>' if col.get('isHidden') else ''
                        calc_badge = '<span class="badge badge-calculated">CALCULATED</span>' if col.get('expression') else ''
                        key_badge = '<span class="badge badge-key">KEY</span>' if col.get('isKey') else ''
                        
                        item_class = 'item calculated' if col.get('expression') else 'item'
                        
                        html += f'<div class="{item_class}">'
                        html += f'<div class="item-name">{col["name"]}{hidden_badge}{calc_badge}{key_badge}</div>'
                        html += f'<div class="item-detail">Type: <code>{col.get("dataType", "N/A")}</code>'
                        
                        if col.get('formatString'):
                            html += f' • Format: <code>{col["formatString"]}</code>'
                        if col.get('dataCategory'):
                            html += f' • Category: {col["dataCategory"]}'
                        
                        html += '</div>'
                        
                        if col.get('description'):
                            html += f'<div class="item-detail">{col["description"]}</div>'
                        
                        # ========== FIX: USE SAFE EXPRESSION GETTER ==========
                        expr = self._safe_get_expression(col, 'expression')
                        if expr:
                            html += f'<div class="dax">{expr}</div>'
                        # =====================================================
                        
                        html += '</div>'
                    html += '</div>'
                
                # Measures
                if table.get('measures'):
                    html += '<div class="subsection"><h4>Measures</h4>'
                    for measure in table['measures']:
                        hidden_badge = '<span class="badge badge-hidden">HIDDEN</span>' if measure.get('isHidden') else ''
                        
                        html += '<div class="item measure">'
                        html += f'<div class="item-name">{measure["name"]}{hidden_badge}</div>'
                        
                        details = []
                        if measure.get('displayFolder'):
                            details.append(f'Folder: {measure["displayFolder"]}')
                        if measure.get('formatString'):
                            details.append(f'Format: <code>{measure["formatString"]}</code>')
                        
                        if details:
                            html += f'<div class="item-detail">{" • ".join(details)}</div>'
                        
                        if measure.get('description'):
                            html += f'<div class="item-detail">{measure["description"]}</div>'
                        
                        # ========== FIX: USE SAFE EXPRESSION GETTER ==========
                        expr = self._safe_get_expression(measure, 'expression')
                        if expr:
                            html += f'<div class="dax">{expr}</div>'
                        # =====================================================
                        
                        html += '</div>'
                    html += '</div>'
                
                # Hierarchies
                if table.get('hierarchies'):
                    html += '<div class="subsection"><h4>Hierarchies</h4>'
                    for hier in table['hierarchies']:
                        html += f'<div class="item">'
                        html += f'<div class="item-name">{hier["name"]}</div>'
                        html += '<div class="item-detail">Levels: '
                        levels = [f'{level["name"]}' for level in hier.get('levels', [])]
                        html += ' → '.join(levels)
                        html += '</div></div>'
                    html += '</div>'
                
                html += '</div>'
            
            html += '</div>'
            
            # Relationships Section
            if self.data['dataModel'].get('relationships'):
                html += '<div class="section"><h2>🔗 Relationships</h2>'
                for rel in self.data['dataModel']['relationships']:
                    is_active = rel.get('isActive', True)
                    active_symbol = '✓' if is_active else '✗'
                    rel_class = 'relationship rel-active' if is_active else 'relationship'
                    
                    html += f'<div class="{rel_class}">'
                    html += f'<strong>{active_symbol} {rel["fromTable"]}[{rel["fromColumn"]}]</strong> → '
                    html += f'<strong>{rel["toTable"]}[{rel["toColumn"]}]</strong><br>'
                    html += f'<div class="item-detail">'
                    html += f'Cardinality: {rel.get("fromCardinality", "?")}:{rel.get("toCardinality", "?")} • '
                    html += f'Cross-filter: {rel.get("crossFilteringBehavior", "N/A")}'
                    html += '</div></div>'
                html += '</div>'
            
            # Security Roles
            if self.data['dataModel'].get('roles'):
                html += '<div class="section"><h2>🔒 Security Roles (RLS)</h2>'
                for role in self.data['dataModel']['roles']:
                    html += '<div class="table-card">'
                    html += f'<div class="table-name">{role["name"]}</div>'
                    if role.get('description'):
                        html += f'<p class="table-meta"><em>{role["description"]}</em></p>'
                    
                    if role.get('tablePermissions'):
                        html += '<div class="subsection"><h4>Table Permissions</h4>'
                        for perm in role['tablePermissions']:
                            html += '<div class="item">'
                            html += f'<div class="item-name">Table: {perm["name"]}</div>'
                            
                            # ========== FIX: USE SAFE EXPRESSION GETTER ==========
                            expr = self._safe_get_expression(perm, 'filterExpression')
                            if expr:
                                html += f'<div class="dax">{expr}</div>'
                            # =====================================================
                            
                            html += '</div>'
                        html += '</div>'
                    html += '</div>'
                html += '</div>'
        
        html += """
        <div class="footer">
            <p>📊 Power BI Model Documentation</p>
            <p>Generated by Power BI Analysis Toolkit</p>
        </div>
    </div>
</body>
</html>
"""
        
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html)
            print(f"✅ HTML Documentation created: {output_file}\n")
        except Exception as e:
            print(f"❌ Error creating HTML: {e}\n")
    
    def analyze_dax_dependencies(self):
        """Analyze DAX dependencies between measures."""
        print("=" * 80)
        print("3. ANALYZING DAX DEPENDENCIES")
        print("=" * 80)
        
        output_file = os.path.join(self.output_dir, 'DAX_Dependencies.txt')
        output_file_reverse = os.path.join(self.output_dir, 'DAX_Reverse_Dependencies.txt')
        
        # Build object lists
        all_tables = set()
        all_columns = {}
        all_measures = {}
        
        for table in self.data['dataModel'].get('tables', []):
            table_name = table['name']
            all_tables.add(table_name)
            all_columns[table_name] = [col['name'] for col in table.get('columns', [])]
            
            for measure in table.get('measures', []):
                # ========== FIX: USE SAFE EXPRESSION GETTER ==========
                all_measures[measure['name']] = {
                    'table': table_name,
                    'expression': self._safe_get_expression(measure, 'expression'),
                    'displayFolder': measure.get('displayFolder', '')
                }
                # =====================================================
        
        print(f"📊 Analyzing {len(all_measures)} measures...")
        
        # Analyze dependencies
        dependencies = {}
        
        for measure_name, measure_info in all_measures.items():
            dax = measure_info['expression']
            deps = {
                'tables_used': set(),
                'columns_used': set(),
                'measures_used': set()
            }
            
            # Find table references
            for table in all_tables:
                pattern1 = rf'\b{re.escape(table)}\s*\['
                pattern2 = rf"'{re.escape(table)}'\s*\["
                if re.search(pattern1, dax, re.IGNORECASE) or re.search(pattern2, dax, re.IGNORECASE):
                    deps['tables_used'].add(table)
                    
                    for col in all_columns.get(table, []):
                        col_pattern1 = rf'\b{re.escape(table)}\s*\[\s*{re.escape(col)}\s*\]'
                        col_pattern2 = rf"'{re.escape(table)}'\s*\[\s*{re.escape(col)}\s*\]"
                        if re.search(col_pattern1, dax, re.IGNORECASE) or re.search(col_pattern2, dax, re.IGNORECASE):
                            deps['columns_used'].add(f"{table}[{col}]")
            
            # Find measure references
            for other_measure in all_measures.keys():
                if other_measure != measure_name:
                    pattern = rf'\[\s*{re.escape(other_measure)}\s*\]'
                    if re.search(pattern, dax, re.IGNORECASE):
                        deps['measures_used'].add(other_measure)
            
            dependencies[measure_name] = {
                'table': measure_info['table'],
                'displayFolder': measure_info['displayFolder'],
                'tables': sorted(deps['tables_used']),
                'columns': sorted(deps['columns_used']),
                'measures': sorted(deps['measures_used'])
            }
        
        # Save forward dependencies
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("DAX MEASURE DEPENDENCIES ANALYSIS\n")
            f.write("What does each measure depend on?\n")
            f.write("=" * 80 + "\n\n")
            
            for measure_name in sorted(dependencies.keys()):
                deps = dependencies[measure_name]
                f.write(f"\n📊 {measure_name}\n")
                f.write(f"   Table: {deps['table']}\n")
                if deps['displayFolder']:
                    f.write(f"   Folder: {deps['displayFolder']}\n")
                
                if deps['tables']:
                    f.write(f"   Uses Tables: {', '.join(deps['tables'])}\n")
                if deps['columns']:
                    f.write(f"   Uses Columns:\n")
                    for col in deps['columns']:
                        f.write(f"      • {col}\n")
                if deps['measures']:
                    f.write(f"   Uses Measures:\n")
                    for meas in deps['measures']:
                        f.write(f"      • {meas}\n")
        
        # Create reverse dependencies
        reverse_deps = defaultdict(list)
        for measure_name, deps in dependencies.items():
            for used_measure in deps['measures']:
                reverse_deps[used_measure].append(measure_name)
        
        # Save reverse dependencies
        with open(output_file_reverse, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("DAX REVERSE DEPENDENCIES\n")
            f.write("What depends on each measure? (Impact Analysis)\n")
            f.write("=" * 80 + "\n\n")
            
            for measure_name in sorted(reverse_deps.keys()):
                f.write(f"\n📊 {measure_name}\n")
                f.write(f"   Used by {len(reverse_deps[measure_name])} measure(s):\n")
                for dependent in sorted(reverse_deps[measure_name]):
                    f.write(f"      • {dependent}\n")
            
            # List measures not used by any other measure
            unused_measures = set(all_measures.keys()) - set(reverse_deps.keys())
            if unused_measures:
                f.write(f"\n\n{'=' * 80}\n")
                f.write(f"LEAF MEASURES (not used by other measures): {len(unused_measures)}\n")
                f.write("=" * 80 + "\n")
                for measure in sorted(unused_measures):
                    f.write(f"   • {measure}\n")
        
        print(f"✅ DAX Dependencies analyzed:")
        print(f"   📄 {output_file}")
        print(f"   📄 {output_file_reverse}\n")
    
    def validate_model(self):
        """Validate the model for common issues."""
        print("=" * 80)
        print("4. VALIDATING MODEL")
        print("=" * 80)
        
        output_file = os.path.join(self.output_dir, 'Model_Validation.txt')
        
        issues = []
        warnings = []
        
        # Check for tables without relationships
        tables_with_rels = set()
        for rel in self.data['dataModel'].get('relationships', []):
            tables_with_rels.add(rel['fromTable'])
            tables_with_rels.add(rel['toTable'])
        
        for table in self.data['dataModel'].get('tables', []):
            table_name = table['name']
            
            # Orphaned tables
            if table_name not in tables_with_rels and not table.get('isHidden'):
                warnings.append(f"⚠️  Table '{table_name}' has no relationships (orphaned table)")
            
            # Tables without columns or measures
            if not table.get('columns') and not table.get('measures'):
                issues.append(f"❌ Table '{table_name}' has no columns or measures")
            
            # Too many calculated columns
            calc_cols = [col for col in table.get('columns', []) if col.get('expression')]
            if len(calc_cols) > 5:
                warnings.append(f"⚠️  Table '{table_name}' has {len(calc_cols)} calculated columns (consider measures for better performance)")
            
            # Measures without format strings
            for measure in table.get('measures', []):
                if not measure.get('formatString') and not measure.get('isHidden'):
                    warnings.append(f"⚠️  Measure '{table_name}[{measure['name']}]' has no format string")
            
            # Hidden tables with visible columns
            if table.get('isHidden'):
                visible_cols = [col for col in table.get('columns', []) if not col.get('isHidden')]
                if visible_cols:
                    warnings.append(f"⚠️  Hidden table '{table_name}' has {len(visible_cols)} visible columns")
        
        # Check relationships
        inactive_rels = [rel for rel in self.data['dataModel'].get('relationships', []) if not rel.get('isActive', True)]
        if inactive_rels:
            warnings.append(f"⚠️  {len(inactive_rels)} inactive relationship(s) found")
        
        bidir_rels = [rel for rel in self.data['dataModel'].get('relationships', []) 
                      if rel.get('crossFilteringBehavior') == 'bothDirections']
        if bidir_rels:
            warnings.append(f"⚠️  {len(bidir_rels)} bidirectional relationship(s) found (can impact performance)")
        
        many_to_many = [rel for rel in self.data['dataModel'].get('relationships', [])
                        if rel.get('fromCardinality') == 'many' and rel.get('toCardinality') == 'many']
        if many_to_many:
            warnings.append(f"⚠️  {len(many_to_many)} many-to-many relationship(s) found (use with caution)")
        
        # Output validation results
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("MODEL VALIDATION REPORT\n")
            f.write("=" * 80 + "\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 80 + "\n\n")
            
            if issues:
                f.write("❌ CRITICAL ISSUES:\n")
                f.write("-" * 80 + "\n")
                for issue in issues:
                    f.write(f"{issue}\n")
                f.write("\n")
            else:
                f.write("✅ No critical issues found\n\n")
            
            if warnings:
                f.write("⚠️  WARNINGS:\n")
                f.write("-" * 80 + "\n")
                for warning in warnings:
                    f.write(f"{warning}\n")
                f.write("\n")
            else:
                f.write("✅ No warnings\n\n")
            
            # Statistics
            f.write("📊 MODEL STATISTICS:\n")
            f.write("-" * 80 + "\n")
            summary = self.data['dataModel'].get('summary', {})
            f.write(f"Tables: {summary.get('totalTables', 0)}\n")
            f.write(f"Measures: {summary.get('totalMeasures', 0)}\n")
            f.write(f"Calculated Columns: {summary.get('totalCalculatedColumns', 0)}\n")
            f.write(f"Relationships: {summary.get('totalRelationships', 0)}\n")
            f.write(f"  - Active: {len([r for r in self.data['dataModel'].get('relationships', []) if r.get('isActive', True)])}\n")
            f.write(f"  - Inactive: {len(inactive_rels)}\n")
            f.write(f"  - Bidirectional: {len(bidir_rels)}\n")
            f.write(f"  - Many-to-Many: {len(many_to_many)}\n")
            f.write(f"Security Roles: {summary.get('totalRoles', 0)}\n")
            
            f.write("\n" + "=" * 80 + "\n")
            f.write("SUMMARY:\n")
            f.write("=" * 80 + "\n")
            f.write(f"Critical Issues: {len(issues)}\n")
            f.write(f"Warnings: {len(warnings)}\n")
            
            if len(issues) == 0 and len(warnings) == 0:
                f.write("\n✅ Model validation passed with no issues!\n")
            elif len(issues) == 0:
                f.write("\n⚠️  Model has warnings but no critical issues\n")
            else:
                f.write("\n❌ Model has critical issues that should be addressed\n")
        
        print(f"✅ Model validated:")
        print(f"   📄 {output_file}")
        print(f"   ❌ Issues: {len(issues)}")
        print(f"   ⚠️  Warnings: {len(warnings)}\n")
    
    def export_all_dax_formulas(self):
        """Export all DAX formulas to separate files."""
        print("=" * 80)
        print("5. EXPORTING ALL DAX FORMULAS")
        print("=" * 80)
        
        # Create DAX subfolder
        dax_dir = os.path.join(self.output_dir, 'DAX_Formulas')
        os.makedirs(dax_dir, exist_ok=True)
        
        # Export 1: All measures in one file
        all_measures_file = os.path.join(self.output_dir, 'All_DAX_Measures.txt')
        measure_count = 0
        
        with open(all_measures_file, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("ALL DAX MEASURES\n")
            f.write("=" * 80 + "\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 80 + "\n\n")
            
            for table in self.data['dataModel'].get('tables', []):
                measures = table.get('measures', [])
                if measures:
                    f.write(f"\n{'=' * 80}\n")
                    f.write(f"TABLE: {table['name']}\n")
                    f.write(f"{'=' * 80}\n\n")
                    
                    for measure in measures:
                        measure_count += 1
                        f.write(f"-- Measure: {measure['name']}\n")
                        if measure.get('displayFolder'):
                            f.write(f"-- Folder: {measure['displayFolder']}\n")
                        if measure.get('formatString'):
                            f.write(f"-- Format: {measure['formatString']}\n")
                        if measure.get('description'):
                            f.write(f"-- Description: {measure['description']}\n")
                        
                        # ========== FIX: USE SAFE EXPRESSION GETTER ==========
                        expr = self._safe_get_expression(measure, 'expression')
                        f.write(f"\n{expr}\n")
                        # =====================================================
                        
                        f.write("\n" + "-" * 80 + "\n\n")
        
        print(f"✅ All measures exported: {all_measures_file}")
        print(f"   📊 Total measures: {measure_count}")
        
        # Export 2: Calculated columns
        calc_cols_file = os.path.join(self.output_dir, 'All_Calculated_Columns.txt')
        calc_col_count = 0
        
        with open(calc_cols_file, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("ALL CALCULATED COLUMNS\n")
            f.write("=" * 80 + "\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 80 + "\n\n")
            
            for table in self.data['dataModel'].get('tables', []):
                calc_cols = [col for col in table.get('columns', []) if col.get('expression')]
                if calc_cols:
                    f.write(f"\n{'=' * 80}\n")
                    f.write(f"TABLE: {table['name']}\n")
                    f.write(f"{'=' * 80}\n\n")
                    
                    for col in calc_cols:
                        calc_col_count += 1
                        f.write(f"-- Column: {col['name']}\n")
                        f.write(f"-- Data Type: {col.get('dataType', 'N/A')}\n")
                        if col.get('formatString'):
                            f.write(f"-- Format: {col['formatString']}\n")
                        if col.get('description'):
                            f.write(f"-- Description: {col['description']}\n")
                        
                        # ========== FIX: USE SAFE EXPRESSION GETTER ==========
                        expr = self._safe_get_expression(col, 'expression')
                        f.write(f"\n{expr}\n")
                        # =====================================================
                        
                        f.write("\n" + "-" * 80 + "\n\n")
        
        print(f"✅ Calculated columns exported: {calc_cols_file}")
        print(f"   📊 Total calculated columns: {calc_col_count}")
        
        # Export 3: Individual files by folder
        print(f"✅ Creating individual DAX files by folder...")
        
        # Group measures by folder
        by_folder = defaultdict(list)
        for table in self.data['dataModel'].get('tables', []):
            for measure in table.get('measures', []):
                folder = measure.get('displayFolder', '_Root')
                by_folder[folder].append({
                    'table': table['name'],
                    'measure': measure
                })
        
        folder_count = 0
        for folder, measures in by_folder.items():
            folder_count += 1
            safe_folder_name = folder.replace('/', '_').replace('\\', '_').replace(':', '_')
            folder_file = os.path.join(dax_dir, f'{safe_folder_name}.txt')
            
            with open(folder_file, 'w', encoding='utf-8') as f:
                f.write(f"DAX MEASURES - Folder: {folder}\n")
                f.write("=" * 80 + "\n\n")
                
                for item in measures:
                    measure = item['measure']
                    f.write(f"[{item['table']}].[{measure['name']}]\n")
                    if measure.get('formatString'):
                        f.write(f"Format: {measure['formatString']}\n")
                    
                    # ========== FIX: USE SAFE EXPRESSION GETTER ==========
                    expr = self._safe_get_expression(measure, 'expression')
                    f.write(f"\n{expr}\n")
                    # =====================================================
                    
                    f.write("\n" + "-" * 80 + "\n\n")
        
        print(f"   📁 Created {folder_count} folder files in: {dax_dir}/")
        
        # Export 4: Summary file
        summary_file = os.path.join(self.output_dir, 'DAX_Summary.txt')
        with open(summary_file, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("DAX FORMULAS SUMMARY\n")
            f.write("=" * 80 + "\n\n")
            
            f.write(f"Total Measures: {measure_count}\n")
            f.write(f"Total Calculated Columns: {calc_col_count}\n")
            f.write(f"Total Display Folders: {len(by_folder)}\n\n")
            
            f.write("Measures by Folder:\n")
            f.write("-" * 80 + "\n")
            for folder in sorted(by_folder.keys()):
                f.write(f"  📁 {folder}: {len(by_folder[folder])} measures\n")
            
            f.write("\n" + "=" * 80 + "\n")
            f.write("FILES GENERATED:\n")
            f.write("=" * 80 + "\n")
            f.write(f"• All_DAX_Measures.txt - All measures in one file\n")
            f.write(f"• All_Calculated_Columns.txt - All calculated columns\n")
            f.write(f"• DAX_Summary.txt - This summary\n")
            f.write(f"• DAX_Formulas/ - Individual files by folder ({folder_count} files)\n")
        
        print(f"✅ DAX summary created: {summary_file}\n")


# =============================================================================
# MAIN EXECUTION
# =============================================================================

if __name__ == '__main__':
    import sys
    
    # Configuration
    json_input = 'pbix_analysis/complete_analysis.json'  # CHANGE THIS to your JSON file path
    output_directory = 'analysis_output'                  # CHANGE THIS for different output folder
    
    # Allow command line arguments
    if len(sys.argv) > 1:
        json_input = sys.argv[1]
    if len(sys.argv) > 2:
        output_directory = sys.argv[2]
    
    # Check if pandas is installed
    try:
        import pandas as pd
    except ImportError:
        print("❌ Error: 'pandas' library not found")
        print("   Please install it: pip install pandas openpyxl")
        sys.exit(1)
    
    # Run analysis
    analyzer = PowerBIAnalyzer(json_input, output_directory)
    analyzer.run_all()