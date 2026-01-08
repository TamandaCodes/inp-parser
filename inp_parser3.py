import os
import pandas as pd
import numpy as np
import re


class inpParser3:

    """Advanced parser for EPANET .inp files using regex to extract pipe and equipment attributes"""
    
    def __init__(self, filepath):
        """Initialize parser with file path"""
        self.filepath = filepath
        self.content = self.read_file(filepath)
        self.sections = {}
        self.data = {}
        self.units = {}
        self.node_map = {}
        self._parse_all_sections()
    
    def read_file(self, file_path):
        """Read in the .inp file from given file path"""
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"The file '{file_path}' cannot be found.")
        if not file_path.lower().endswith(".inp"):
            raise ValueError("The file must have a .inp extension.")
        
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            content = f.read()
        return content
    
    def _parse_all_sections(self):
        """Parse all sections from the INP file"""
        # Split content by section headers
        section_pattern = r'^\*{3}\s+(.+?)\s+\*{3}'
        matches = list(re.finditer(section_pattern, self.content, re.MULTILINE))
        
        for i, match in enumerate(matches):
            section_name = match.group(1).strip()
            start_pos = match.end()
            
            # Find end position (start of next section or end of file)
            end_pos = matches[i + 1].start() if i + 1 < len(matches) else len(self.content)
            section_content = self.content[start_pos:end_pos].strip()
            
            self.sections[section_name] = section_content
            
            # Parse each section based on its type
            self._parse_section(section_name, section_content)
        
        # After parsing all sections, extract network connectivity
        self._build_master_node_map()
        self.extract_connectivity()
    
    def _parse_section(self, section_name, content):
        """Route section to appropriate parser based on its structure"""
        if "Pipe Detail Summary" in section_name:
            self._parse_pipe_detail_summary(content)
        elif "Pipe Elevations" in section_name:
            self._parse_pipe_elevations(content)
        elif "Transient Data Table" in section_name:
            self._parse_transient_data(content)
        elif any(keyword in section_name for keyword in ["Table", "Branch", "Pump", "Valve", "Pressure"]):
            self._parse_table_section(section_name, content)
    
    def _parse_pipe_detail_summary(self, content):
        """Parse Pipe Detail Summary section (Type A - key-value pairs)"""
        pipe_pattern = r'Pipe\s+(\d+)\s+Detailed Input Data'
        pipes = {}
        current_units = {}
        
        # Split by pipe
        pipe_blocks = re.split(pipe_pattern, content)
        
        for i in range(1, len(pipe_blocks), 2):
            pipe_num = pipe_blocks[i]
            pipe_data_block = pipe_blocks[i + 1]
            
            pipe_info = {}
            
            # Extract Name
            name_match = re.search(r'Name:\s*(.+)', pipe_data_block)
            if name_match:
                pipe_info['Name'] = name_match.group(1).strip()
            
            # Extract Geometry
            geom_match = re.search(r'Geometry:\s*(.+)', pipe_data_block)
            if geom_match:
                pipe_info['Geometry'] = geom_match.group(1).strip()
            
            # Extract properties with values and units
            # Pattern: Property= value units
            prop_pattern = r'([A-Za-z\s&]+?)=\s*([\d.]+)\s+([a-z]+)'
            
            for match in re.finditer(prop_pattern, pipe_data_block):
                prop_name = match.group(1).strip().replace(' ', '_')
                value = float(match.group(2))
                unit = match.group(3)
                
                # Store value
                pipe_info[f"{prop_name}"] = value
                
                # Store unit
                current_units[prop_name] = unit
            
            pipes[f"Pipe_{pipe_num}"] = pipe_info
        
        if pipes:
            df = pd.DataFrame(pipes).T
            
            # Add units to column names
            renamed_cols = {}
            for col in df.columns:
                if col in ['Name', 'Geometry']:
                    renamed_cols[col] = col
                else:
                    unit = current_units.get(col, '')
                    renamed_cols[col] = f"{col.replace('_', ' ')} ({unit})" if unit else col.replace('_', ' ')
            
            df.rename(columns=renamed_cols, inplace=True)
            self.data['Pipe_Detail_Summary'] = df
            self.units['Pipe_Detail_Summary'] = current_units
    
    def _parse_pipe_elevations(self, content):
        """Parse Pipe Elevations section"""
        pipe_pattern = r'^([A-Z]\d+)\s+\(([^)]+)\)'
        lines = content.split('\n')
        
        current_pipe = None
        segment_data = []
        all_pipes = {}
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Check for new pipe header
            pipe_match = re.match(pipe_pattern, line)
            if pipe_match:
                # Save previous pipe
                if current_pipe and segment_data:
                    all_pipes[current_pipe] = pd.DataFrame(segment_data)
                
                current_pipe = f"{pipe_match.group(1)} ({pipe_match.group(2)})"
                segment_data = []
                continue
            
            # Skip header lines
            if 'Length' in line and 'Elevation' in line:
                continue
            if 'Along Pipe' in line or 'of Segment' in line:
                continue
            if line.startswith('(feet)') or line == 'n/a':
                continue
            
            # Parse segment data
            parts = line.split()
            if len(parts) >= 3 and current_pipe:
                try:
                    segment_data.append({
                        'Length Along Pipe (feet)': float(parts[0]) if parts[0] != 'n/a' else None,
                        'Length of Segment (feet)': float(parts[1]) if parts[1] != 'n/a' else None,
                        'Elevation (feet)': float(parts[2]) if parts[2] != 'n/a' else None
                    })
                except ValueError:
                    continue
        
        # Save last pipe
        if current_pipe and segment_data:
            all_pipes[current_pipe] = pd.DataFrame(segment_data)
        
        # Create summary
        summary = {}
        for pipe_id, segments_df in all_pipes.items():
            if not segments_df.empty:
                elevations = segments_df['Elevation (feet)'].dropna()
                lengths = segments_df['Length of Segment (feet)'].dropna()
                
                summary[pipe_id] = {
                    'Num Segments': len(segments_df),
                    'Total Length (feet)': lengths.sum() if not lengths.empty else None,
                    'Start Elevation (feet)': elevations.iloc[0] if not elevations.empty else None,
                    'End Elevation (feet)': elevations.iloc[-1] if not elevations.empty else None,
                    'Elevation Change (feet)': elevations.iloc[-1] - elevations.iloc[0] if len(elevations) >= 2 else None,
                    'Min Elevation (feet)': elevations.min() if not elevations.empty else None,
                    'Max Elevation (feet)': elevations.max() if not elevations.empty else None
                }
        
        if summary:
            self.data['Pipe_Elevations_Summary'] = pd.DataFrame(summary).T
        
        self.data['Pipe_Elevations_Detailed'] = all_pipes
        self.units['Pipe_Elevations'] = {'default': 'feet'}
    
    def _parse_transient_data(self, content):
        """Parse Transient Data Table section (Type C)"""
        # Pattern: Equipment_ID (Type) Transient Data:
        equipment_pattern = r'([A-Z]\d+)\s+\(([^)]+)\)\s+Transient Data:'
        
        blocks = re.split(equipment_pattern, content)
        transient_data = {}
        
        for i in range(1, len(blocks), 3):
            equipment_id = blocks[i]
            equipment_type = blocks[i + 1]
            data_block = blocks[i + 2]
            
            # Look for Time Data section
            if 'Time Data' in data_block:
                # Extract column headers with units
                header_match = re.search(r'Time\s+(.+?)(?=\n\s*\d)', data_block, re.DOTALL)
                
                if header_match:
                    header_line = header_match.group(0)
                    
                    # Extract column name and unit
                    unit_pattern = r'(.+?)\s*\(([^)]+)\)'
                    unit_match = re.search(unit_pattern, header_line)
                    
                    col_name = unit_match.group(1).strip() if unit_match else "Value"
                    unit = unit_match.group(2) if unit_match else ""
                    
                    # Extract data rows
                    data_rows = []
                    for line in data_block.split('\n'):
                        parts = line.split()
                        if len(parts) == 2:
                            try:
                                data_rows.append({
                                    'Time': float(parts[0]),
                                    f'{col_name} ({unit})': float(parts[1])
                                })
                            except ValueError:
                                continue
                    
                    if data_rows:
                        transient_data[f"{equipment_id} ({equipment_type})"] = pd.DataFrame(data_rows)
        
        if transient_data:
            self.data['Transient_Data'] = transient_data
    
    def _parse_table_section(self, section_name, content):
        """Parse table sections (Type B - columnar data) with multi-part headers"""
        lines = [l.strip() for l in content.split('\n') if l.strip()]
        
        if not lines:
            return
        
        # Special handling for sections with multiple data blocks
        if any(keyword in section_name for keyword in ['Control Valve', 'Assigned Pressure']):
            self._parse_multi_block_table(section_name, content, lines)
            return
        
        # Find header row(s)
        header_lines = []
        data_start_idx = 0
        
        for i, line in enumerate(lines):
            # Header lines typically contain words, not just numbers
            if re.search(r'[A-Za-z]', line) and not line[0].isdigit():
                header_lines.append(line)
            else:
                data_start_idx = i
                break
        
        if not header_lines:
            return
        
        # Parse headers and extract units
        headers = []
        units_dict = {}
        
        # Check if multi-row header
        if len(header_lines) > 1:
            # Combine multi-row headers
            primary_headers = header_lines[0].split()
            
            # Check for "Units" row
            if any('Units' in h or 'units' in h for h in header_lines):
                for h_line in header_lines:
                    if 'Units' in h_line or 'units' in h_line:
                        units_line = h_line.split()
                        # Map units to columns
                        for idx, unit in enumerate(units_line):
                            if unit.lower() not in ['units', 'unit'] and idx < len(primary_headers):
                                units_dict[primary_headers[idx]] = unit
            
            # Build combined headers
            for i, h in enumerate(primary_headers):
                if h in units_dict:
                    headers.append(f"{h} ({units_dict[h]})")
                else:
                    headers.append(h)
        else:
            # Single row header - check for embedded units
            header_parts = header_lines[0].split()
            for part in header_parts:
                headers.append(part)
        
        # Parse data rows
        data_rows = []
        for line in lines[data_start_idx:]:
            if not line or line.startswith('---'):
                continue
            
            parts = line.split()
            if parts and len(parts) <= len(headers):
                data_rows.append(parts)
        
        if data_rows:
            # Create DataFrame
            df = pd.DataFrame(data_rows)
            
            # Adjust column count if needed
            if len(df.columns) < len(headers):
                headers = headers[:len(df.columns)]
            elif len(df.columns) > len(headers):
                # Pad headers
                for i in range(len(headers), len(df.columns)):
                    headers.append(f"Column_{i}")
            
            df.columns = headers
            
            # Convert numeric columns
            for col in df.columns:
                try:
                    df[col] = pd.to_numeric(df[col])
                except (ValueError, TypeError):
                    pass
            
            # Clean section name for storage
            clean_name = section_name.replace(' ', '_').replace('*', '').strip()
            self.data[clean_name] = df
            self.units[clean_name] = units_dict
    
    def _parse_multi_block_table(self, section_name, content, lines):
        """Parse tables with multiple data blocks separated by header rows"""
        blocks = []
        current_block_headers = []
        current_block_data = []
        current_block_name = None
        in_data = False
        
        i = 0
        while i < len(lines):
            line = lines[i]
            
            # Check if this is a header line
            is_header = (re.search(r'[A-Za-z]', line) and 
                        not line[0].isdigit() and 
                        not line.startswith('('))
            
            if is_header:
                # Save previous block if it exists
                if current_block_data and current_block_headers:
                    blocks.append({
                        'name': current_block_name,
                        'headers': current_block_headers,
                        'data': current_block_data
                    })
                
                # Start new block
                current_block_name = line.split()[0] if not current_block_name else current_block_name
                current_block_headers.append(line)
                current_block_data = []
                in_data = False
            else:
                # This is a data line
                parts = line.split()
                if parts and parts[0].replace('.', '').replace('-', '').isdigit():
                    current_block_data.append(parts)
                    in_data = True
            
            i += 1
        
        # Save last block
        if current_block_data and current_block_headers:
            blocks.append({
                'name': current_block_name,
                'headers': current_block_headers,
                'data': current_block_data
            })
        
        # Process each block and create combined DataFrame
        all_data = []
        all_headers = set()
        
        for block in blocks:
            # Parse headers for this block
            headers = self._parse_block_headers(block['headers'])
            all_headers.update(headers)
            
            # Create DataFrame for this block
            block_df = pd.DataFrame(block['data'])
            
            # Ensure correct number of columns
            if len(block_df.columns) > len(headers):
                headers.extend([f'Col_{i}' for i in range(len(headers), len(block_df.columns))])
            elif len(block_df.columns) < len(headers):
                headers = headers[:len(block_df.columns)]
            
            block_df.columns = headers
            
            # Convert numeric columns
            for col in block_df.columns:
                try:
                    block_df[col] = pd.to_numeric(block_df[col])
                except (ValueError, TypeError):
                    pass
            
            all_data.append(block_df)
        
        # Combine all blocks
        if all_data:
            # Concatenate dataframes horizontally, handling duplicate columns
            combined_df = all_data[0]
            id_col = combined_df.columns[0]
            
            for block_idx, df in enumerate(all_data[1:], start=1):
                df_id_col = df.columns[0]
                
                # Rename columns in the new df to avoid duplicates (except ID column)
                new_cols = {}
                for col in df.columns:
                    if col == df_id_col:
                        continue  # Skip ID column
                    # Check if column already exists
                    if col in combined_df.columns:
                        # Add suffix to make it unique
                        new_cols[col] = f"{col}_block{block_idx}"
                    else:
                        new_cols[col] = col
                
                df_renamed = df.rename(columns=new_cols)
                
                # Merge
                combined_df = combined_df.merge(
                    df_renamed, 
                    left_on=id_col, 
                    right_on=df_id_col, 
                    how='outer', 
                    suffixes=('', f'_b{block_idx}')
                )
                
                # Remove duplicate ID columns if they exist
                dup_cols = [c for c in combined_df.columns if c.endswith(f'_b{block_idx}') and c.replace(f'_b{block_idx}', '') == id_col]
                if dup_cols:
                    combined_df = combined_df.drop(columns=dup_cols)
            
            # Clean section name for storage
            clean_name = section_name.replace(' ', '_').replace('*', '').strip()
            self.data[clean_name] = combined_df
    
    def _parse_block_headers(self, header_lines):
        """Parse headers from a block, handling multi-row headers with units"""
        headers = []
        
        if not header_lines:
            return headers
        
        # Join all header lines to handle wrapped headers
        full_header_text = ' '.join(header_lines)
        
        # Try to extract structured column names
        # Pattern for complex headers like "(Pipe #1) K In, K Out"
        pipe_pattern = r'\(Pipe #(\d+)\)\s+([\w\s,]+)'
        pipe_matches = re.findall(pipe_pattern, full_header_text)
        
        if pipe_matches:
            # This is the Assigned Pressure table with pipe columns
            # Start with basic columns
            basic_headers = header_lines[0].split()
            
            # Filter out pipe-related text from basic headers
            filtered_basic = []
            skip_next = False
            for i, h in enumerate(basic_headers):
                if '(Pipe' in h or skip_next:
                    skip_next = ')' not in h
                    continue
                filtered_basic.append(h)
            
            headers.extend(filtered_basic)
            
            # Add pipe columns
            for pipe_num, pipe_cols in pipe_matches:
                cols = [c.strip() for c in pipe_cols.split(',')]
                for col in cols:
                    headers.append(f"Pipe_{pipe_num}_{col.replace(' ', '_')}")
            
            return headers
        
        # Standard multi-row header parsing
        all_parts = []
        for line in header_lines:
            all_parts.append(line.split())
        
        # Find the line with most parts (main header)
        main_header_idx = 0
        max_parts = 0
        for idx, parts in enumerate(all_parts):
            if len(parts) > max_parts:
                max_parts = len(parts)
                main_header_idx = idx
        
        main_headers = all_parts[main_header_idx]
        
        # Look for units line
        units_line = None
        for line in header_lines:
            if 'Units' in line or 'units' in line:
                units_line = line.split()
                break
        
        # Build headers with units
        for i, header in enumerate(main_headers):
            if units_line and i < len(units_line):
                unit = units_line[i]
                if unit.lower() not in ['units', 'unit']:
                    # Check if unit looks like a unit (lowercase, common units)
                    if unit in ['feet', 'inches', 'psia', 'psig', 'barrels/day', 'gpm', 'N/A']:
                        if unit != 'N/A':
                            headers.append(f"{header} ({unit})")
                        else:
                            headers.append(header)
                    else:
                        headers.append(header)
                else:
                    headers.append(header)
            else:
                headers.append(header)
        
        return headers
    
    def _build_master_node_map(self):
        """
        Scans all component tables to create a dictionary mapping Node IDs to descriptive names.
        Example: Map ID '30' -> 'Junction 30 (J 84)' or 'Source 6 (Sullivan)'
        """
        # Map Section Headers to readable Component Types
        type_map = {
            'Branch Table': 'Junction',
            'Assigned Pressure Table': 'Boundary',
            'Assigned Flow Table': 'Source',
            'Valve Table': 'Valve',
            'Control Valve Table': 'Control Valve',
            'Reservoir Table': 'Reservoir',
            'Check Valve Table': 'Check Valve',
            'Relief Valve Table': 'Relief Valve',
            'Surge Tank Table': 'Surge Tank',
            'Pump Table': 'Pump'
        }

        for section_key, content in self.sections.items():
            # Find the component type based on the section name
            comp_type = next((v for k, v in type_map.items() if k in section_key), None)
            if not comp_type:
                continue

            lines = content.split('\n')
            for line in lines:
                line = line.strip()
                # Skip headers and empty lines. Data lines start with a digit.
                if not line or not line[0].isdigit():
                    continue
                
                parts = line.split()
                if len(parts) >= 2:
                    node_id = parts[0]
                    
                    # --- FIX: Prevent Overwriting ---
                    # The Definition table (with the Name) is always the first block in the section.
                    # Subsequent blocks list attributes (like "No", "(P17)", etc.).
                    # If we have already named this node, skip it to avoid overwriting with garbage.
                    if node_id in self.node_map:
                        continue

                    # Logic: In AFT, Name is often the 2nd column.
                    # If the Name is "J" (common for Junctions), check if parts[2] is a number (e.g. "J 84")
                    raw_name = parts[1]
                    
                    # Special handling for names with spaces (like "J 84") 
                    # If part[1] is 'J' and part[2] is a number, combine them.
                    if len(parts) > 2 and parts[1] == 'J' and parts[2].isdigit():
                         raw_name = f"J {parts[2]}"
                    # Special handling for Multi-Word names (e.g., "Wiedeman Federal")
                    # If there is a 3rd part and it's NOT "Yes"/"No" or a number, append it.
                    elif len(parts) > 2 and parts[2] not in ['Yes', 'No'] and not parts[2].replace('.','').isdigit():
                         raw_name = f"{parts[1]} {parts[2]}"
                    
                    # Store in map
                    self.node_map[node_id] = f"{comp_type} {node_id} ({raw_name})"

    def pipeNames(self):
        """Extract names of all pipes in the network"""
        pipe_summary = self.data.get('Pipe_Detail_Summary')
        if pipe_summary is not None and 'Name' in pipe_summary.columns:
            return pd.Series(pipe_summary['Name'].values, index=pipe_summary.index, name='Pipe_Name')
        return pd.Series(dtype=str)
    
    def pipeDiameter(self):
        """Extract diameter of each pipe"""
        pipe_summary = self.data.get('Pipe_Detail_Summary')
        if pipe_summary is not None:
            diameter_col = [col for col in pipe_summary.columns if 'Diameter' in col]
            if diameter_col:
                return pd.Series(pipe_summary[diameter_col[0]].values, index=pipe_summary.index, name='Diameter')
        return pd.Series(dtype=float)
    
    def pipeTotal_Length(self):
        """Extract total length of each pipe"""
        pipe_summary = self.data.get('Pipe_Detail_Summary')
        if pipe_summary is not None:
            length_col = [col for col in pipe_summary.columns if 'Length' in col]
            if length_col:
                return pd.Series(pipe_summary[length_col[0]].values, index=pipe_summary.index, name='Length')
        return pd.Series(dtype=float)
    
    def pipeRoughness(self):
        """Extract roughness coefficient of each pipe"""
        pipe_summary = self.data.get('Pipe_Detail_Summary')
        if pipe_summary is not None:
            roughness_col = [col for col in pipe_summary.columns if 'Roughness' in col]
            if roughness_col:
                return pd.Series(pipe_summary[roughness_col[0]].values, index=pipe_summary.index, name='Roughness')
        return pd.Series(dtype=float)
    
    def pipeLen_Elev(self):
        """Extract length vs elevation data for pipes from Pipe Elevations section"""
        return self.data.get('Pipe_Elevations_Summary', pd.DataFrame())
    
    def getPumps(self):
        """Extract pump names and properties"""
        return self.data.get('Pump_Table', pd.DataFrame())
    
    def getJunctions(self):
        """Extract junction names"""
        return self.data.get('Branch_Table', pd.DataFrame())
    
    def getReservoirs(self):
        """Extract reservoir names and pressures"""
        return self.data.get('Assigned_Pressure_Table', pd.DataFrame())
    
    def getValves(self):
        """Extract control valve names and flowrates"""
        valve_table = self.data.get('Valve_Table', pd.DataFrame())
        control_valve_table = self.data.get('Control_Valve_Table', pd.DataFrame())
        
        if not valve_table.empty:
            return valve_table
        elif not control_valve_table.empty:
            return control_valve_table
        return pd.DataFrame()
    
    def extract_connectivity(self):
        """
        Extracts pipe connections prioritizing the 'Junctions (Up,Down)' table column.
        """
        connections = []
        target_section = None
        
        # --- Strategy 1: Find the Table with "Junctions (Up,Down)" (High Accuracy) ---
        for name, content in self.sections.items():
            if "Junctions" in content and "Up,Down" in content:
                target_section = content
                break
        
        if target_section:
            lines = target_section.split('\n')
            # Regex to find: PipeID ... UpNode, DownNode
            # Matches lines like: "17   30, 280"
            pattern = re.compile(r'^\s*(\d+)\s+.*?(\d+)\s*,\s*(\d+)')
            
            for line in lines:
                match = pattern.search(line)
                if match:
                    pipe_id = match.group(1)
                    up_id = match.group(2)
                    down_id = match.group(3)
                    
                    # Use the Master Node Map to get friendly names
                    up_name = self.node_map.get(up_id, f"Node {up_id}")
                    down_name = self.node_map.get(down_id, f"Node {down_id}")
                    
                    connections.append({
                        'Name': f"Pipe_{pipe_id}",
                        'Upstream Node': up_name,
                        'Downstream Node': down_name,
                        'Notes': ''
                    })
        
        # --- Strategy 2: Fallback (Scan component tables) ---
        if not connections:
            pipe_pattern = re.compile(r'\(P(\d+)\)')
            pipe_raw_links = {}

            # Reuse the type map logic to scan tables
            for section_key, content in self.sections.items():
                if "Table" not in section_key: continue
                
                lines = content.split('\n')
                for line in lines:
                    line = line.strip()
                    if not line or not line[0].isdigit(): continue
                    
                    parts = line.split()
                    node_id = parts[0]
                    node_label = self.node_map.get(node_id, f"Node {node_id}")
                    
                    found_pipes = pipe_pattern.findall(line)
                    for pid in found_pipes:
                        pkey = f"Pipe_{pid}"
                        if pkey not in pipe_raw_links: pipe_raw_links[pkey] = []
                        if node_label not in pipe_raw_links[pkey]:
                            pipe_raw_links[pkey].append(node_label)
            
            # Format fallback connections
            for pkey, nodes in pipe_raw_links.items():
                row = {'Name': pkey}
                if len(nodes) > 0: row['Upstream Node'] = nodes[0]
                else: row['Upstream Node'] = "Unknown"
                
                if len(nodes) > 1: row['Downstream Node'] = nodes[1]
                else: row['Downstream Node'] = "Boundary/Open"
                
                row['Notes'] = "Inferred from Components"
                connections.append(row)

        self.data['Network_Connectivity'] = pd.DataFrame(connections)
    
    def get_network_connectivity(self):
        """Get network connectivity DataFrame showing pipe connections to junctions"""
        return self.data.get('Network_Connectivity')
    
    def extract_all(self):
        """Extract all pipe and equipment data into a comprehensive dataset"""
        # Pipe data
        pipe_data = pd.DataFrame({
            'Pipe_Name': self.pipeNames(),
            'Diameter': self.pipeDiameter(),
            'Length': self.pipeTotal_Length(),
            'Roughness': self.pipeRoughness()
        })
        
        # Add elevation data
        len_elev = self.pipeLen_Elev()
        if not len_elev.empty:
            pipe_data = pipe_data.join(len_elev[['Start Elevation (feet)', 'End Elevation (feet)', 'Elevation Change (feet)']])
        
        # Extract network connectivity (already done in _parse_all_sections)
        connectivity = self.get_network_connectivity()
        
        # Equipment data
        equipment_data = {
            'Pumps': self.getPumps(),
            'Junctions': self.getJunctions(),
            'Reservoirs': self.getReservoirs(),
            'Valves': self.getValves()
        }
        
        return pipe_data, equipment_data
    
    def export_to_excel(self, output_path=None, include_detailed_segments=True):
        """Export all parsed data to Excel file"""
        if output_path is None:
            base_name = os.path.splitext(self.filepath)[0]
            output_path = f"{base_name}_parsed.xlsx"
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:

            #  Connectivity first
            if 'Network_Connectivity' in self.data:
                self.data['Network_Connectivity'].to_excel(
                    writer, sheet_name='Network_Connectivity', index=False
                )

            # Core tables
            for section, df in self.data.items():
                if section in ['Network_Connectivity', 'Pipe_Elevations_Detailed', 'Transient_Data']:
                    continue

                if isinstance(df, pd.DataFrame) and not df.empty:
                    sheet = section[:31]
                    df.to_excel(writer, sheet_name=sheet, index=True)

            # Detailed segments LAST
            if include_detailed_segments:
                elev = self.data.get('Pipe_Elevations_Detailed', {})
                for pipe, df in elev.items():
                    df.to_excel(writer, sheet_name=f"Elev_{pipe}"[:31], index=False)

                trans = self.data.get('Transient_Data', {})
                for eq, df in trans.items():
                    df.to_excel(writer, sheet_name=f"Trans_{eq}"[:31], index=False)

        print(f"✓ Data successfully exported to: {output_path}")
        print(f"✓ Exported {len([d for d in self.data.values() if isinstance(d, pd.DataFrame)])} main sections")
        return output_path

    # --- GUI MAIN EXECUTION ---

# GUI implementation
if __name__ == "__main__":
    import tkinter as tk
    from tkinter import filedialog, messagebox
    import os

    # 1. Hide the main tkinter window (we only want the pop-ups)
    root = tk.Tk()
    root.withdraw()

    # 2. Ask user to select the INP file
   
    inp_path = filedialog.askopenfilename(
        title="Select your AFT Impulse .inp file",
        filetypes=[("AFT Input Files", "*.inp"), ("All Files", "*.*")]
    )

    # Check if they clicked Cancel
    if not inp_path:
        print("No file selected. Exiting.")
        exit()

    # 3. Ask user where to save the Excel file
    default_save = os.path.splitext(inp_path)[0] + "_parsed.xlsx"
    save_path = filedialog.asksaveasfilename(
        title="Save Excel Report As...",
        defaultextension=".xlsx",
        initialfile=os.path.basename(default_save),
        filetypes=[("Excel Files", "*.xlsx")]
    )

    if not save_path:
        print("No save location selected. Exiting.")
        exit()

    # 4. Run the Process
    try:
        
        parser = inpParser3(inp_path)
        parser.extract_connectivity()
        parser.export_to_excel(output_path=save_path, include_detailed_segments=True)
        
        # 5. Success Popup
        messagebox.showinfo(
            title="Success", 
            message=f"Extraction Complete!\n\nSaved to:\n{save_path}"
        )
        
    except Exception as e:
        # Error Popup
        messagebox.showerror(
            title="Error", 
            message=f"An error occurred during extraction:\n\n{str(e)}"
        )