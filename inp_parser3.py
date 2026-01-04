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
        """Parse table sections (Type B - columnar data)"""
        lines = [l.strip() for l in content.split('\n') if l.strip()]
        
        if not lines:
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
    
    def get_section(self, section_name):
        """Get parsed data for a specific section"""
        return self.data.get(section_name)
    
    def get_pipe_detail_summary(self):
        """Get pipe detail summary DataFrame"""
        return self.data.get('Pipe_Detail_Summary')
    
    def get_pipe_elevations_summary(self):
        """Get pipe elevations summary DataFrame"""
        return self.data.get('Pipe_Elevations_Summary')
    
    def get_pipe_elevations_detailed(self, pipe_id=None):
        """Get detailed pipe elevation segments"""
        detailed = self.data.get('Pipe_Elevations_Detailed', {})
        if pipe_id:
            return detailed.get(pipe_id)
        return detailed
    
    def get_transient_data(self, equipment_id=None):
        """Get transient data for equipment"""
        transient = self.data.get('Transient_Data', {})
        if equipment_id:
            return transient.get(equipment_id)
        return transient
    
    def list_sections(self):
        """List all available sections"""
        return list(self.data.keys())
    
    def export_to_excel(self, output_path=None, include_detailed_segments=True):
        """Export all parsed data to Excel file
        
        Args:
            output_path: Path for output Excel file (optional)
            include_detailed_segments: If True, export detailed segment data for pipes
        """
        if output_path is None:
            base_name = os.path.splitext(self.filepath)[0]
            output_path = f"{base_name}_parsed.xlsx"
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Export main sections
            for section_name, df in self.data.items():
                if isinstance(df, pd.DataFrame) and not df.empty:
                    # Clean sheet name (Excel 31 char limit, no special chars)
                    sheet_name = section_name[:31].replace('/', '_').replace('\\', '_')
                    df.to_excel(writer, sheet_name=sheet_name, index=True)
                
                elif isinstance(df, dict):
                    # Handle nested data (like transient data, detailed elevations)
                    if section_name == 'Pipe_Elevations_Detailed' and include_detailed_segments:
                        for pipe_id, pipe_df in df.items():
                            if isinstance(pipe_df, pd.DataFrame) and not pipe_df.empty:
                                sheet_name = f"Elev_{pipe_id}"[:31].replace('/', '_').replace('\\', '_')
                                pipe_df.to_excel(writer, sheet_name=sheet_name, index=True)
                    
                    elif section_name == 'Transient_Data':
                        for equipment_id, equip_df in df.items():
                            if isinstance(equip_df, pd.DataFrame) and not equip_df.empty:
                                sheet_name = f"Trans_{equipment_id}"[:31].replace('/', '_').replace('\\', '_')
                                equip_df.to_excel(writer, sheet_name=sheet_name, index=True)
        
        print(f"✓ Data successfully exported to: {output_path}")
        print(f"✓ Exported {len([d for d in self.data.values() if isinstance(d, pd.DataFrame)])} main sections")
        return output_path
    
    def summary(self):
        """Print summary of parsed data"""
        print("\n" + "="*60)
        print("INP PARSER SUMMARY")
        print("="*60)
        print(f"File: {os.path.basename(self.filepath)}")
        print(f"\nSections found: {len(self.sections)}")
        print(f"Sections parsed: {len(self.data)}")
        print("\nAvailable data:")
        for section_name, data in self.data.items():
            if isinstance(data, pd.DataFrame):
                print(f"  • {section_name}: {len(data)} rows")
            elif isinstance(data, dict):
                print(f"  • {section_name}: {len(data)} items")
        print("="*60 + "\n")


# Example usage
if __name__ == "__main__":
    parser = inpParser3("Sample.inp")
    parser.summary()
    parser.export_to_excel(include_detailed_segments=False)
    pass