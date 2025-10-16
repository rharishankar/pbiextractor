import zipfile
import json
import os
from pathlib import Path
from datetime import datetime
import base64

class PBIXParser:
    """Comprehensive PBIX/PBIT file parser - Pure Python, no external dependencies."""
    
    def __init__(self, file_path, output_dir='pbix_analysis'):
        self.file_path = file_path
        self.output_dir = output_dir
        self.extract_dir = os.path.join(output_dir, 'extracted')
        self.results = {}
    
    # ========== ADD THIS NEW METHOD ==========
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
        
    def parse(self):
        """Main parsing function - extracts everything possible."""
        print("=" * 80)
        print(f"COMPREHENSIVE POWER BI FILE ANALYSIS")
        print(f"File: {self.file_path}")
        print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 80)
        print()
        
        # Create output directories
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.extract_dir, exist_ok=True)
        
        # Extract the ZIP file
        self._extract_file()
        
        # Parse all components
        self._parse_report_layout()
        self._parse_data_model_schema()
        self._parse_connections()
        self._parse_metadata()
        self._parse_custom_visuals()
        self._parse_diagram_layout()
        self._parse_bookmarks()
        self._parse_report_settings()
        self._parse_mobile_layout()
        self._parse_theme()
        self._parse_static_resources()
        self._parse_version_info()
        self._list_all_files()
        
        # Generate outputs
        self._save_json_output()
        self._generate_summary_report()
        self._generate_detailed_report()
        self._generate_measures_report()
        self._generate_relationships_diagram()
        
        print("\n" + "=" * 80)
        print("✓ ANALYSIS COMPLETE!")
        print("=" * 80)
        self._print_summary()
        
        return self.results
    
    def _extract_file(self):
        """Extract PBIX/PBIT file (it's a ZIP archive)."""
        print("📦 Extracting file...")
        try:
            with zipfile.ZipFile(self.file_path, 'r') as zip_ref:
                zip_ref.extractall(self.extract_dir)
            print(f"   ✓ Extracted to: {self.extract_dir}\n")
        except Exception as e:
            print(f"   ✗ Error extracting file: {e}\n")
            raise
    
    def _parse_report_layout(self):
        """Parse Report Layout - pages, visuals, filters."""
        print("📄 Parsing Report Layout...")
        try:
            layout_path = os.path.join(self.extract_dir, 'Report', 'Layout')
            if not os.path.exists(layout_path):
                print("   ⚠ Layout file not found\n")
                return
            
            with open(layout_path, 'r', encoding='utf-16-le') as f:
                layout = json.load(f)
            
            self.results['reportLayout'] = {
                'id': layout.get('id'),
                'resourcePackages': layout.get('resourcePackages', []),
                'pages': [],
                'config': layout.get('config'),
                'layoutOptimization': layout.get('layoutOptimization')
            }
            
            # Parse each page/section
            for section in layout.get('sections', []):
                page_info = {
                    'name': section.get('name'),
                    'displayName': section.get('displayName', 'Unnamed Page'),
                    'width': section.get('width'),
                    'height': section.get('height'),
                    'displayOption': section.get('displayOption'),
                    'visualContainers': [],
                    'filters': section.get('filters', ''),
                    'config': section.get('config')
                }
                
                # Parse visuals
                for visual in section.get('visualContainers', []):
                    visual_info = {
                        'x': visual.get('x'),
                        'y': visual.get('y'),
                        'z': visual.get('z'),
                        'width': visual.get('width'),
                        'height': visual.get('height'),
                        'type': 'Unknown',
                        'title': '',
                        'config': {}
                    }
                    
                    # Parse visual config
                    config_str = visual.get('config', '')
                    if config_str:
                        try:
                            config = json.loads(config_str)
                            visual_info['config'] = config
                            
                            # Extract visual type
                            if 'singleVisual' in config:
                                visual_type = config['singleVisual'].get('visualType', 'Unknown')
                                visual_info['type'] = visual_type
                                
                                # Extract visual title
                                if 'vcObjects' in config['singleVisual']:
                                    vc_objects = config['singleVisual']['vcObjects']
                                    if 'title' in vc_objects:
                                        title_obj = vc_objects['title'][0]
                                        if 'properties' in title_obj:
                                            title_props = title_obj['properties']
                                            if 'text' in title_props:
                                                visual_info['title'] = title_props['text'].get('expr', {}).get('Literal', {}).get('Value', '')
                                
                                # Extract data roles (what fields are used)
                                if 'prototypeQuery' in config['singleVisual']:
                                    proto = config['singleVisual']['prototypeQuery']
                                    visual_info['dataRoles'] = proto
                        except json.JSONDecodeError:
                            pass
                    
                    # Parse filters
                    filters_str = visual.get('filters', '')
                    if filters_str:
                        try:
                            visual_info['filters'] = json.loads(filters_str)
                        except:
                            visual_info['filters'] = filters_str
                    
                    page_info['visualContainers'].append(visual_info)
                
                self.results['reportLayout']['pages'].append(page_info)
            
            total_visuals = sum(len(p['visualContainers']) for p in self.results['reportLayout']['pages'])
            print(f"   ✓ Found {len(self.results['reportLayout']['pages'])} pages")
            print(f"   ✓ Found {total_visuals} total visuals\n")
            
        except Exception as e:
            print(f"   ✗ Error parsing layout: {e}\n")
    
    def _parse_data_model_schema(self):
        """Parse DataModelSchema - tables, columns, measures, relationships."""
        print("🗄️  Parsing Data Model Schema...")
        try:
            schema_path = os.path.join(self.extract_dir, 'DataModelSchema')
            if not os.path.exists(schema_path):
                print("   ⚠ DataModelSchema not found (older PBIX format)")
                print("   → Recommend converting to PBIT or using pbi-tools\n")
                return
            
            with open(schema_path, 'r', encoding='utf-16-le') as f:
                schema = json.load(f)
            
            model = schema.get('model', {})
            
            self.results['dataModel'] = {
                'name': model.get('name'),
                'description': model.get('description'),
                'culture': model.get('culture'),
                'defaultMode': model.get('defaultMode'),
                'tables': [],
                'relationships': [],
                'cultures': [],
                'perspectives': [],
                'roles': [],
                'annotations': model.get('annotations', [])
            }
            
            # Parse tables
            all_measures = []
            all_calculated_columns = []
            all_calculated_tables = []
            
            for table in model.get('tables', []):
                table_info = {
                    'name': table.get('name'),
                    'description': table.get('description', ''),
                    'isHidden': table.get('isHidden', False),
                    'isPrivate': table.get('isPrivate', False),
                    'lineageTag': table.get('lineageTag'),
                    'columns': [],
                    'measures': [],
                    'hierarchies': [],
                    'partitions': [],
                    'annotations': table.get('annotations', [])
                }
                
                # Parse columns
                for col in table.get('columns', []):
                    col_info = {
                        'name': col.get('name'),
                        'dataType': col.get('dataType'),
                        'isHidden': col.get('isHidden', False),
                        'isKey': col.get('isKey', False),
                        'isNullable': col.get('isNullable', True),
                        'sourceColumn': col.get('sourceColumn'),
                        'formatString': col.get('formatString'),
                        'dataCategory': col.get('dataCategory'),
                        'summarizeBy': col.get('summarizeBy', 'default'),
                        'displayFolder': col.get('displayFolder', ''),
                        'description': col.get('description', ''),
                        'expression': col.get('expression'),  # Calculated column DAX
                        'sortByColumn': col.get('sortByColumn'),
                        'annotations': col.get('annotations', [])
                    }
                    
                    if col_info['expression']:
                        all_calculated_columns.append(f"{table['name']}[{col['name']}]")
                    
                    table_info['columns'].append(col_info)
                
                # Parse measures
                for measure in table.get('measures', []):
                    measure_info = {
                        'name': measure.get('name'),
                        'expression': measure.get('expression'),
                        'formatString': measure.get('formatString'),
                        'isHidden': measure.get('isHidden', False),
                        'displayFolder': measure.get('displayFolder', ''),
                        'description': measure.get('description', ''),
                        'lineageTag': measure.get('lineageTag'),
                        'annotations': measure.get('annotations', [])
                    }
                    table_info['measures'].append(measure_info)
                    all_measures.append({
                        'table': table['name'],
                        'measure': measure['name'],
                        'expression': measure.get('expression', ''),
                        'displayFolder': measure.get('displayFolder', '')
                    })
                
                # Parse hierarchies
                for hierarchy in table.get('hierarchies', []):
                    hier_info = {
                        'name': hierarchy.get('name'),
                        'isHidden': hierarchy.get('isHidden', False),
                        'levels': []
                    }
                    for level in hierarchy.get('levels', []):
                        hier_info['levels'].append({
                            'name': level.get('name'),
                            'column': level.get('column'),
                            'ordinal': level.get('ordinal')
                        })
                    table_info['hierarchies'].append(hier_info)
                
                # Parse partitions (data source queries)
                for partition in table.get('partitions', []):
                    part_info = {
                        'name': partition.get('name'),
                        'mode': partition.get('mode'),
                        'source': partition.get('source', {}),
                        'annotations': partition.get('annotations', [])
                    }
                    
                    # Check if it's a calculated table
                    if part_info['source'].get('type') == 'calculated':
                        all_calculated_tables.append(table['name'])
                    
                    table_info['partitions'].append(part_info)
                
                self.results['dataModel']['tables'].append(table_info)
            
            # Parse relationships
            for rel in model.get('relationships', []):
                rel_info = {
                    'name': rel.get('name'),
                    'fromTable': rel.get('fromTable'),
                    'fromColumn': rel.get('fromColumn'),
                    'toTable': rel.get('toTable'),
                    'toColumn': rel.get('toColumn'),
                    'fromCardinality': rel.get('fromCardinality'),
                    'toCardinality': rel.get('toCardinality'),
                    'crossFilteringBehavior': rel.get('crossFilteringBehavior'),
                    'securityFilteringBehavior': rel.get('securityFilteringBehavior'),
                    'isActive': rel.get('isActive', True),
                    'relyOnReferentialIntegrity': rel.get('relyOnReferentialIntegrity', False),
                    'annotations': rel.get('annotations', [])
                }
                self.results['dataModel']['relationships'].append(rel_info)
            
            # Parse cultures (translations)
            for culture in model.get('cultures', []):
                self.results['dataModel']['cultures'].append({
                    'name': culture.get('name'),
                    'linguisticMetadata': culture.get('linguisticMetadata', {})
                })
            
            # Parse perspectives
            for perspective in model.get('perspectives', []):
                self.results['dataModel']['perspectives'].append({
                    'name': perspective.get('name'),
                    'description': perspective.get('description', ''),
                    'annotations': perspective.get('annotations', [])
                })
            
            # Parse RLS roles
            for role in model.get('roles', []):
                role_info = {
                    'name': role.get('name'),
                    'description': role.get('description', ''),
                    'modelPermission': role.get('modelPermission'),
                    'tablePermissions': []
                }
                
                for perm in role.get('tablePermissions', []):
                    role_info['tablePermissions'].append({
                        'name': perm.get('name'),
                        'filterExpression': perm.get('filterExpression'),
                        'annotations': perm.get('annotations', [])
                    })
                
                self.results['dataModel']['roles'].append(role_info)
            
            # Store summary counts
            self.results['dataModel']['summary'] = {
                'totalTables': len(self.results['dataModel']['tables']),
                'totalMeasures': len(all_measures),
                'totalRelationships': len(self.results['dataModel']['relationships']),
                'totalRoles': len(self.results['dataModel']['roles']),
                'totalCalculatedColumns': len(all_calculated_columns),
                'totalCalculatedTables': len(all_calculated_tables),
                'allMeasures': all_measures,
                'calculatedColumns': all_calculated_columns,
                'calculatedTables': all_calculated_tables
            }
            
            print(f"   ✓ Found {len(self.results['dataModel']['tables'])} tables")
            print(f"   ✓ Found {len(all_measures)} measures")
            print(f"   ✓ Found {len(all_calculated_columns)} calculated columns")
            print(f"   ✓ Found {len(all_calculated_tables)} calculated tables")
            print(f"   ✓ Found {len(self.results['dataModel']['relationships'])} relationships")
            if self.results['dataModel']['roles']:
                print(f"   ✓ Found {len(self.results['dataModel']['roles'])} security roles")
            print()
            
        except Exception as e:
            print(f"   ✗ Error parsing data model: {e}\n")
    
    def _parse_connections(self):
    """Parse Connections - data sources."""
    print("🔌 Parsing Data Connections...")
    try:
        connections_path = os.path.join(self.extract_dir, 'Connections')
        if not os.path.exists(connections_path):
            print("   ⚠ Connections file not found\n")
            return
        
        # Try different encodings
        connections = None
        encodings_to_try = ['utf-16-le', 'utf-8', 'utf-16', 'utf-16-be']
        
        for encoding in encodings_to_try:
            try:
                with open(connections_path, 'r', encoding=encoding) as f:
                    connections = json.load(f)
                print(f"   ✓ Successfully read with {encoding} encoding")
                break
            except (UnicodeDecodeError, json.JSONDecodeError):
                continue
        
        if connections is None:
            print("   ⚠ Could not read Connections file with any known encoding")
            print("   Skipping connections parsing\n")
            return
        
        self.results['connections'] = []
        for conn in connections.get('Connections', []):
            conn_info = {
                'name': conn.get('Name'),
                'connectionString': conn.get('ConnectionString'),
                'connectionType': conn.get('ConnectionType'),
                'impersonationMode': conn.get('ImpersonationMode'),
                'privacy': conn.get('Privacy'),
                'annotations': conn.get('Annotations', [])
            }
            self.results['connections'].append(conn_info)
        
        print(f"   ✓ Found {len(self.results['connections'])} data source connections\n")
        
    except Exception as e:
        print(f"   ⚠ Error parsing connections: {e}")
        print("   Continuing without connections data\n")
    
      
    
    def _parse_metadata(self):
        """Parse Metadata."""
        print("ℹ️  Parsing Metadata...")
        try:
            metadata_path = os.path.join(self.extract_dir, 'Metadata', 'metadata.json')
            if not os.path.exists(metadata_path):
                print("   ⚠ Metadata file not found\n")
                return
            
            with open(metadata_path, 'r', encoding='utf-8') as f:
                metadata = json.load(f)
            
            self.results['metadata'] = metadata
            print(f"   ✓ Metadata loaded\n")
            
        except Exception as e:
            print(f"   ✗ Error parsing metadata: {e}\n")
    
    def _parse_custom_visuals(self):
        """Parse Custom Visuals."""
        print("🎨 Parsing Custom Visuals...")
        try:
            custom_visuals_dir = os.path.join(self.extract_dir, 'Report', 'CustomVisuals')
            if not os.path.exists(custom_visuals_dir):
                print("   ⚠ No custom visuals found\n")
                return
            
            self.results['customVisuals'] = []
            for item in os.listdir(custom_visuals_dir):
                if not item.startswith('.'):
                    visual_path = os.path.join(custom_visuals_dir, item)
                    visual_info = {
                        'name': item,
                        'size': os.path.getsize(visual_path)
                    }
                    
                    # Try to read package.json if it's a directory
                    if os.path.isdir(visual_path):
                        package_json = os.path.join(visual_path, 'package.json')
                        if os.path.exists(package_json):
                            try:
                                with open(package_json, 'r', encoding='utf-8') as f:
                                    package = json.load(f)
                                visual_info['package'] = package
                            except:
                                pass
                    
                    self.results['customVisuals'].append(visual_info)
            
            print(f"   ✓ Found {len(self.results['customVisuals'])} custom visuals\n")
            
        except Exception as e:
            print(f"   ✗ Error parsing custom visuals: {e}\n")
    
    def _parse_diagram_layout(self):
        """Parse DiagramLayout - model diagram."""
        print("📐 Parsing Diagram Layout...")
        try:
            diagram_path = os.path.join(self.extract_dir, 'DiagramLayout')
            if not os.path.exists(diagram_path):
                print("   ⚠ Diagram layout not found\n")
                return
            
            with open(diagram_path, 'r', encoding='utf-16-le') as f:
                diagram = json.load(f)
            
            self.results['diagramLayout'] = diagram
            print(f"   ✓ Diagram layout loaded\n")
            
        except Exception as e:
            print(f"   ✗ Error parsing diagram layout: {e}\n")
    
    def _parse_bookmarks(self):
        """Parse bookmarks."""
        print("🔖 Parsing Bookmarks...")
        try:
            # Bookmarks can be in Report/bookmarks.json or in Layout
            bookmarks_path = os.path.join(self.extract_dir, 'Report', 'bookmarks.json')
            
            if os.path.exists(bookmarks_path):
                with open(bookmarks_path, 'r', encoding='utf-8') as f:
                    bookmarks = json.load(f)
                self.results['bookmarks'] = bookmarks
                print(f"   ✓ Found bookmarks\n")
            else:
                print("   ⚠ No bookmarks found\n")
            
        except Exception as e:
            print(f"   ✗ Error parsing bookmarks: {e}\n")
    
    def _parse_report_settings(self):
        """Parse report settings."""
        print("⚙️  Parsing Report Settings...")
        try:
            settings_path = os.path.join(self.extract_dir, 'Settings')
            if not os.path.exists(settings_path):
                print("   ⚠ Settings file not found\n")
                return
            
            with open(settings_path, 'r', encoding='utf-16-le') as f:
                settings = json.load(f)
            
            self.results['settings'] = settings
            print(f"   ✓ Settings loaded\n")
            
        except Exception as e:
            print(f"   ✗ Error parsing settings: {e}\n")
    
    def _parse_mobile_layout(self):
        """Parse mobile layout if present."""
        print("📱 Parsing Mobile Layout...")
        try:
            mobile_path = os.path.join(self.extract_dir, 'Report', 'MobileState')
            if not os.path.exists(mobile_path):
                print("   ⚠ No mobile layout found\n")
                return
            
            with open(mobile_path, 'r', encoding='utf-16-le') as f:
                mobile = json.load(f)
            
            self.results['mobileLayout'] = mobile
            print(f"   ✓ Mobile layout loaded\n")
            
        except Exception as e:
            print(f"   ✗ Error parsing mobile layout: {e}\n")
    
    def _parse_theme(self):
        """Parse report theme."""
        print("🎨 Parsing Theme...")
        try:
            theme_path = os.path.join(self.extract_dir, 'Report', 'StaticResources', 'SharedResources', 'BaseThemes')
            if not os.path.exists(theme_path):
                print("   ⚠ Theme not found\n")
                return
            
            # Theme might be in various locations
            for item in os.listdir(theme_path):
                if item.endswith('.json'):
                    with open(os.path.join(theme_path, item), 'r', encoding='utf-8') as f:
                        theme = json.load(f)
                    self.results['theme'] = theme
                    print(f"   ✓ Theme loaded: {item}\n")
                    break
            
        except Exception as e:
            print(f"   ✗ Error parsing theme: {e}\n")
    
    def _parse_static_resources(self):
        """Parse static resources (images, etc.)."""
        print("🖼️  Parsing Static Resources...")
        try:
            static_path = os.path.join(self.extract_dir, 'Report', 'StaticResources')
            if not os.path.exists(static_path):
                print("   ⚠ No static resources found\n")
                return
            
            self.results['staticResources'] = []
            for root, dirs, files in os.walk(static_path):
                for file in files:
                    file_path = os.path.join(root, file)
                    rel_path = os.path.relpath(file_path, static_path)
                    self.results['staticResources'].append({
                        'path': rel_path,
                        'size': os.path.getsize(file_path),
                        'extension': os.path.splitext(file)[1]
                    })
            
            if self.results['staticResources']:
                print(f"   ✓ Found {len(self.results['staticResources'])} static resources\n")
            else:
                print("   ⚠ No static resources found\n")
            
        except Exception as e:
            print(f"   ✗ Error parsing static resources: {e}\n")
    
    def _parse_version_info(self):
        """Parse version information."""
        print("🔢 Parsing Version Info...")
        try:
            version_path = os.path.join(self.extract_dir, 'Version')
            if not os.path.exists(version_path):
                print("   ⚠ Version file not found\n")
                return
            
            with open(version_path, 'r', encoding='utf-8') as f:
                version = f.read().strip()
            
            self.results['version'] = version
            print(f"   ✓ Version: {version}\n")
            
        except Exception as e:
            print(f"   ✗ Error parsing version: {e}\n")
    
    def _list_all_files(self):
        """List all files in the extracted archive."""
        print("📂 Cataloging All Files...")
        try:
            self.results['fileStructure'] = []
            for root, dirs, files in os.walk(self.extract_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    rel_path = os.path.relpath(file_path, self.extract_dir)
                    self.results['fileStructure'].append({
                        'path': rel_path,
                        'size': os.path.getsize(file_path)
                    })
            
            print(f"   ✓ Cataloged {len(self.results['fileStructure'])} files\n")
            
        except Exception as e:
            print(f"   ✗ Error listing files: {e}\n")
    
    def _save_json_output(self):
        """Save complete results as JSON."""
        output_file = os.path.join(self.output_dir, 'complete_analysis.json')
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2, ensure_ascii=False)
        print(f"💾 Saved complete JSON: {output_file}")
    
    def _generate_summary_report(self):
        """Generate a concise summary report."""
        output_file = os.path.join(self.output_dir, 'summary_report.txt')
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("POWER BI FILE - SUMMARY REPORT\n")
            f.write("=" * 80 + "\n")
            f.write(f"File: {self.file_path}\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 80 + "\n\n")
            
            # Report Pages
            if 'reportLayout' in self.results:
                pages = self.results['reportLayout'].get('pages', [])
                f.write(f"📄 REPORT PAGES: {len(pages)}\n")
                f.write("-" * 80 + "\n")
                for page in pages:
                    f.write(f"  • {page['displayName']}\n")
                    f.write(f"    Visuals: {len(page['visualContainers'])}\n")
                f.write("\n")
            
            # Data Model
            if 'dataModel' in self.results:
                dm = self.results['dataModel']
                summary = dm.get('summary', {})
                
                f.write(f"🗄️  DATA MODEL\n")
                f.write("-" * 80 + "\n")
                f.write(f"  Tables: {summary.get('totalTables', 0)}\n")
                f.write(f"  Measures: {summary.get('totalMeasures', 0)}\n")
                f.write(f"  Calculated Columns: {summary.get('totalCalculatedColumns', 0)}\n")
                f.write(f"  Calculated Tables: {summary.get('totalCalculatedTables', 0)}\n")
                f.write(f"  Relationships: {summary.get('totalRelationships', 0)}\n")
                f.write(f"  Security Roles: {summary.get('totalRoles', 0)}\n")
                f.write("\n")
            
            # Connections
            if 'connections' in self.results:
                f.write(f"🔌 DATA CONNECTIONS: {len(self.results['connections'])}\n")
                f.write("-" * 80 + "\n")
                for conn in self.results['connections']:
                    f.write(f"  • {conn['name']} ({conn.get('connectionType', 'N/A')})\n")
                f.write("\n")
            
            # Custom Visuals
            if 'customVisuals' in self.results:
                f.write(f"🎨 CUSTOM VISUALS: {len(self.results['customVisuals'])}\n")
                f.write("-" * 80 + "\n")
                for visual in self.results['customVisuals']:
                    f.write(f"  • {visual['name']}\n")
                f.write("\n")
            
            # Version
            if 'version' in self.results:
                f.write(f"🔢 VERSION: {self.results['version']}\n\n")
            
            f.write("=" * 80 + "\n")
        
        print(f"💾 Saved summary report: {output_file}")
    
    def _generate_detailed_report(self):
        """Generate a detailed report with all information."""
        output_file = os.path.join(self.output_dir, 'detailed_report.txt')
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("POWER BI FILE - DETAILED ANALYSIS REPORT\n")
            f.write("=" * 80 + "\n")
            f.write(f"File: {self.file_path}\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 80 + "\n\n")
            
            # REPORT PAGES
            if 'reportLayout' in self.results:
                f.write("\n" + "=" * 80 + "\n")
                f.write("📄 REPORT PAGES\n")
                f.write("=" * 80 + "\n\n")
                
                for page in self.results['reportLayout'].get('pages', []):
                    f.write(f"PAGE: {page['displayName']}\n")
                    f.write("-" * 80 + "\n")
                    f.write(f"Dimensions: {page['width']} x {page['height']}\n")
                    f.write(f"Visuals: {len(page['visualContainers'])}\n\n")
                    
                    # Visual breakdown
                    visual_types = {}
                    for visual in page['visualContainers']:
                        vtype = visual['type']
                        visual_types[vtype] = visual_types.get(vtype, 0) + 1
                    
                    f.write("Visual Types:\n")
                    for vtype, count in sorted(visual_types.items()):
                        f.write(f"  • {vtype}: {count}\n")
                    
                    f.write("\nVisual Details:\n")
                    for i, visual in enumerate(page['visualContainers'], 1):
                        f.write(f"  [{i}] {visual['type']}\n")
                        if visual.get('title'):
                            f.write(f"      Title: {visual['title']}\n")
                        f.write(f"      Position: ({visual['x']}, {visual['y']})\n")
                        f.write(f"      Size: {visual['width']} x {visual['height']}\n")
                    
                    f.write("\n")
            
            # DATA MODEL
            if 'dataModel' in self.results:
                f.write("\n" + "=" * 80 + "\n")
                f.write("🗄️  DATA MODEL\n")
                f.write("=" * 80 + "\n\n")
                
                # Tables
                for table in self.results['dataModel'].get('tables', []):
                    hidden = " [HIDDEN]" if table.get('isHidden') else ""
                    f.write(f"TABLE: {table['name']}{hidden}\n")
                    f.write("-" * 80 + "\n")
                    
                    if table.get('description'):
                        f.write(f"Description: {table['description']}\n")
                    
                    f.write(f"\nColumns ({len(table['columns'])}):\n")
                    for col in table['columns']:
                        hidden_col = " [HIDDEN]" if col.get('isHidden') else ""
                        calc = " [CALCULATED]" if col.get('expression') else ""
                        f.write(f"  • {col['name']}{hidden_col}{calc}\n")
                        f.write(f"    Type: {col.get('dataType', 'N/A')}\n")
                        if col.get('formatString'):
                            f.write(f"    Format: {col['formatString']}\n")
                        if col.get('expression'):
                            f.write(f"    Expression: {col['expression']}\n")
                    
                    if table['measures']:
                        f.write(f"\nMeasures ({len(table['measures'])}):\n")
                        for measure in table['measures']:
                            hidden_meas = " [HIDDEN]" if measure.get('isHidden') else ""
                            f.write(f"  • {measure['name']}{hidden_meas}\n")
                            if measure.get('displayFolder'):
                                f.write(f"    Folder: {measure['displayFolder']}\n")
                            
                            # ========== FIX: USE SAFE EXPRESSION GETTER ==========
                            expr = self._safe_get_expression(measure, 'expression')
                            if expr:
                                expr = expr.replace('\n', '\n    ')
                                f.write(f"    DAX: {expr}\n")
                            # =====================================================
                            
                            if measure.get('formatString'):
                                f.write(f"    Format: {measure['formatString']}\n")
                    
                    if table['hierarchies']:
                        f.write(f"\nHierarchies ({len(table['hierarchies'])}):\n")
                        for hier in table['hierarchies']:
                            f.write(f"  • {hier['name']}\n")
                            for level in hier['levels']:
                                f.write(f"    → {level['name']} ({level['column']})\n")
                    
                    f.write("\n")
                
                # Relationships
                if self.results['dataModel'].get('relationships'):
                    f.write("\n" + "-" * 80 + "\n")
                    f.write("RELATIONSHIPS\n")
                    f.write("-" * 80 + "\n")
                    for rel in self.results['dataModel']['relationships']:
                        active = "✓" if rel.get('isActive', True) else "✗"
                        f.write(f"{active} {rel['fromTable']}[{rel['fromColumn']}] → ")
                        f.write(f"{rel['toTable']}[{rel['toColumn']}]\n")
                        f.write(f"   Cardinality: {rel.get('fromCardinality', 'N/A')} to ")
                        f.write(f"{rel.get('toCardinality', 'N/A')}\n")
                        f.write(f"   Cross-filtering: {rel.get('crossFilteringBehavior', 'N/A')}\n\n")
                
                # Security Roles
                if self.results['dataModel'].get('roles'):
                    f.write("\n" + "-" * 80 + "\n")
                    f.write("SECURITY ROLES (RLS)\n")
                    f.write("-" * 80 + "\n")
                    for role in self.results['dataModel']['roles']:
                        f.write(f"ROLE: {role['name']}\n")
                        if role.get('description'):
                            f.write(f"Description: {role['description']}\n")
                        f.write(f"Permissions:\n")
                        for perm in role.get('tablePermissions', []):
                            f.write(f"  • Table: {perm['name']}\n")
                            
                            # ========== FIX: USE SAFE EXPRESSION GETTER ==========
                            expr = self._safe_get_expression(perm, 'filterExpression')
                            if expr:
                                expr = expr.replace('\n', '\n    ')
                                f.write(f"    Filter: {expr}\n")
                            # =====================================================
                        f.write("\n")
            
            # CONNECTIONS
            if 'connections' in self.results:
                f.write("\n" + "=" * 80 + "\n")
                f.write("🔌 DATA CONNECTIONS\n")
                f.write("=" * 80 + "\n\n")
                for conn in self.results['connections']:
                    f.write(f"CONNECTION: {conn['name']}\n")
                    f.write("-" * 80 + "\n")
                    f.write(f"Type: {conn.get('connectionType', 'N/A')}\n")
                    if conn.get('connectionString'):
                        f.write(f"Connection String: {conn['connectionString']}\n")
                    f.write("\n")
            
            f.write("\n" + "=" * 80 + "\n")
        
        print(f"💾 Saved detailed report: {output_file}")
    
    def _generate_measures_report(self):
        """Generate a dedicated report for all measures."""
        if 'dataModel' not in self.results:
            return
        
        output_file = os.path.join(self.output_dir, 'measures_report.txt')
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("ALL DAX MEASURES\n")
            f.write("=" * 80 + "\n\n")
            
            all_measures = self.results['dataModel'].get('summary', {}).get('allMeasures', [])
            
            # Group by display folder
            by_folder = {}
            for measure in all_measures:
                folder = measure.get('displayFolder', '(Root)')
                if folder not in by_folder:
                    by_folder[folder] = []
                by_folder[folder].append(measure)
            
            for folder in sorted(by_folder.keys()):
                f.write(f"\n📁 {folder}\n")
                f.write("-" * 80 + "\n")
                
                for measure in by_folder[folder]:
                    f.write(f"\n{measure['table']}[{measure['measure']}]\n")
                    
                    # ========== FIX: USE SAFE EXPRESSION GETTER ==========
                    expr = self._safe_get_expression(measure, 'expression')
                    if expr:
                        expr = expr.strip()
                        f.write(f"{expr}\n")
                    # =====================================================
                    f.write("\n")
        
        print(f"💾 Saved measures report: {output_file}")
    
    def _generate_relationships_diagram(self):
        """Generate a text-based relationships diagram."""
        if 'dataModel' not in self.results:
            return
        
        output_file = os.path.join(self.output_dir, 'relationships_diagram.txt')
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("DATA MODEL RELATIONSHIPS\n")
            f.write("=" * 80 + "\n\n")
            
            # Group relationships by from table
            by_table = {}
            for rel in self.results['dataModel'].get('relationships', []):
                from_table = rel['fromTable']
                if from_table not in by_table:
                    by_table[from_table] = []
                by_table[from_table].append(rel)
            
            for table in sorted(by_table.keys()):
                f.write(f"\n{table}\n")
                for rel in by_table[table]:
                    active = "━━" if rel.get('isActive', True) else "┄┄"
                    cardinality = f"{rel.get('fromCardinality', '?')}:{rel.get('toCardinality', '?')}"
                    f.write(f"  [{rel['fromColumn']}] {active}({cardinality}){active}> ")
                    f.write(f"{rel['toTable']}[{rel['toColumn']}]\n")
                f.write("\n")
        
        print(f"💾 Saved relationships diagram: {output_file}")
    
    def _print_summary(self):
        """Print summary to console."""
        print(f"\n📊 ANALYSIS SUMMARY:")
        print("-" * 80)
        
        if 'reportLayout' in self.results:
            pages = len(self.results['reportLayout'].get('pages', []))
            visuals = sum(len(p['visualContainers']) for p in self.results['reportLayout'].get('pages', []))
            print(f"Report Pages: {pages}")
            print(f"Total Visuals: {visuals}")
        
        if 'dataModel' in self.results:
            summary = self.results['dataModel'].get('summary', {})
            print(f"Tables: {summary.get('totalTables', 0)}")
            print(f"Measures: {summary.get('totalMeasures', 0)}")
            print(f"Relationships: {summary.get('totalRelationships', 0)}")
            print(f"Calculated Columns: {summary.get('totalCalculatedColumns', 0)}")
            print(f"Calculated Tables: {summary.get('totalCalculatedTables', 0)}")
            if summary.get('totalRoles', 0) > 0:
                print(f"Security Roles: {summary.get('totalRoles', 0)}")
        
        if 'connections' in self.results:
            print(f"Data Connections: {len(self.results['connections'])}")
        
        if 'customVisuals' in self.results:
            print(f"Custom Visuals: {len(self.results['customVisuals'])}")
        
        if 'version' in self.results:
            print(f"Power BI Version: {self.results['version']}")
        
        print("\n📁 OUTPUT FILES:")
        print("-" * 80)
        print(f"  • {os.path.join(self.output_dir, 'complete_analysis.json')}")
        print(f"  • {os.path.join(self.output_dir, 'summary_report.txt')}")
        print(f"  • {os.path.join(self.output_dir, 'detailed_report.txt')}")
        print(f"  • {os.path.join(self.output_dir, 'measures_report.txt')}")
        print(f"  • {os.path.join(self.output_dir, 'relationships_diagram.txt')}")
        print(f"  • {os.path.join(self.output_dir, 'extracted', '...')} (raw files)")


# =============================================================================
# MAIN EXECUTION
# =============================================================================

if __name__ == '__main__':
    import sys
    
    # Check if file path provided
    if len(sys.argv) > 1:
        pbix_file = sys.argv[1]
    else:
        # Default file name - change this to your PBIX/PBIT file
        pbix_file = 'your_report.pbit'
    
    # Check if file exists
    if not os.path.exists(pbix_file):
        print(f"❌ Error: File not found: {pbix_file}")
        print(f"\nUsage: python {sys.argv[0]} <path_to_pbix_or_pbit_file>")
        print(f"   Or: Edit the script and change 'your_report.pbit' to your file name")
        sys.exit(1)
    
    # Parse the file
    parser = PBIXParser(pbix_file)
    try:
        results = parser.parse()
    except Exception as e:
        print(f"\n❌ Fatal error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)