# EPANET INP Parser

A Python tool to extract pipe and equipment attributes from EPANET .inp files and export them to Excel for easier readability and analysis.

## Features

✨ **Comprehensive Data Extraction:**
- Pipe details (name, diameter, length, roughness, geometry)
- Pipe elevation profiles with statistics (min/max elevation, elevation change)
- Equipment data (pumps, valves, control valves, junctions)
- Branch and pressure tables
- Transient data for equipment
- Connection information

📊 **Smart Parsing:**
- Uses regular expressions for robust parsing
- Automatically detects section types
- Extracts units from data and includes them in column headers
- Handles multi-row table headers

📁 **Excel Export:**
- Clean, readable Excel format
- Units embedded in column headers (e.g., "Length (feet)")
- Multiple sheets for different data types
- Optional detailed segment data export

## Installation

### Prerequisites
```bash
pip install pandas openpyxl numpy
```

### Download
Clone or download this repository:
```bash
git clone https://github.com/yourusername/inp-parser.git
cd inp-parser
```

Or install dependencies from requirements.txt:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```python
from inpParser3 import inpParser3

# Initialize parser with your INP file
parser = inpParser3("path/to/your/network.inp")

# View summary of parsed data
parser.summary()

# Export all data to Excel
parser.export_to_excel("output.xlsx", include_detailed_segments=False)
```

### Advanced Usage

```python
# Access specific sections
pipe_summary = parser.get_pipe_detail_summary()
print(pipe_summary.head())

# Get elevation summary with statistics
elevation_summary = parser.get_pipe_elevations_summary()
print(elevation_summary)

# Access specific pipe elevation details
pipe_details = parser.get_pipe_elevations_detailed("P2 (Gustafson1-1)")

# Get transient data for equipment
transient_data = parser.get_transient_data("J272 (Control Valve)")

# List all available sections
sections = parser.list_sections()
print(sections)

# Get any section by name
branch_table = parser.get_section('Branch_Table')
pump_table = parser.get_section('Pump_Table')
```

### Export Options

```python
# Export without detailed pipe segments (recommended for large networks)
parser.export_to_excel("output.xlsx", include_detailed_segments=False)

# Export with detailed pipe segments (creates separate sheet per pipe)
parser.export_to_excel("output.xlsx", include_detailed_segments=True)

# Custom output path
parser.export_to_excel("/custom/path/results.xlsx")
```

## Supported Sections

The parser automatically detects and extracts data from these sections:

- **Pipe Detail Summary** - Pipe properties with units
- **Pipe Elevations** - Segment-by-segment elevation data
  - Summary statistics (min/max elevation, elevation change)
  - Detailed segments (optional)
- **Branch Table** - Branch connections and elevations
- **Pump Table** - Pump locations and properties
- **Valve Table** - Valve information
- **Control Valve Table** - Control valve settings
- **Assigned Pressure Table** - Pressure assignments
- **Transient Data Table** - Time-series data for equipment

## Output Format

### Excel Structure

**Main Sheets:**
- `Pipe_Detail_Summary` - All pipe properties
- `Pipe_Elevations_Summary` - Elevation statistics per pipe
- `Branch_Table` - Branch data
- `Pump_Table` - Pump data
- `Control_Valve_Table` - Control valve data
- `Valve_Table` - Valve data
- `Assigned_Pressure_Table` - Pressure data

**Optional Sheets (if include_detailed_segments=True):**
- `Elev_P2 (Gustafson1-1)` - Detailed segments for Pipe 2
- `Trans_J272 (Control Valve)` - Transient data for equipment J272

### Column Headers with Units

All numeric columns include units in the header:
- `Hydraulic Diameter (inches)`
- `Length (feet)`
- `Absolute Roughness (inches)`
- `Elevation (feet)`

## Example Output

### Pipe Detail Summary
| Pipe Name | Hydraulic Diameter (inches) | Length (feet) | Absolute Roughness (inches) |
|-----------|----------------------------|---------------|----------------------------|
| Gustafson1-1 | 7.63 | 50.39 | 0.0018 |
| XTO Dwyer 1 | 3.67 | 12778.39 | 0.0018 |

### Pipe Elevations Summary
| Pipe | Num Segments | Total Length (feet) | Start Elevation (feet) | End Elevation (feet) | Elevation Change (feet) | Min Elevation (feet) | Max Elevation (feet) |
|------|--------------|---------------------|----------------------|---------------------|------------------------|---------------------|---------------------|
| P2 (Gustafson1-1) | 5 | 50.39 | 2240.16 | 2248 | 7.84 | 2240.16 | 2248.555 |

## Requirements

- Python 3.7+
- pandas >= 1.3.0
- openpyxl >= 3.0.0
- numpy >= 1.20.0

## File Structure

```
inp-parser/
├── inpParser3.py          # Main parser class
├── README.md              # This file
├── requirements.txt       # Python dependencies
└── examples/
    └── example_usage.py   # Usage examples
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT License - feel free to use this in your projects!

## Acknowledgments

Built for TSnet simulations and EPANET network analysis.

## Support

If you encounter any issues or have questions, please open an issue on GitHub.

---

**Happy Parsing! 🚀**