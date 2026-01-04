"""
Example usage of inpParser3 for parsing EPANET .inp files
"""

from inpParser3 import inpParser3

def basic_example():
    """Basic usage example"""
    print("="*60)
    print("BASIC USAGE EXAMPLE")
    print("="*60)
    
    # Initialize parser
    parser = inpParser3("your_network.inp")
    
    # Display summary
    parser.summary()
    
    # Export to Excel (without detailed segments for cleaner output)
    parser.export_to_excel("network_output.xlsx", include_detailed_segments=False)
    
    print("\n✓ Basic example complete!")


def advanced_example():
    """Advanced usage with data access"""
    print("\n" + "="*60)
    print("ADVANCED USAGE EXAMPLE")
    print("="*60)
    
    # Initialize parser
    parser = inpParser3("your_network.inp")
    
    # Access specific sections
    print("\n1. Accessing Pipe Detail Summary:")
    pipe_summary = parser.get_pipe_detail_summary()
    if pipe_summary is not None:
        print(f"   Found {len(pipe_summary)} pipes")
        print("\n   First 3 pipes:")
        print(pipe_summary.head(3))
    
    # Access elevation data
    print("\n2. Accessing Pipe Elevation Summary:")
    elev_summary = parser.get_pipe_elevations_summary()
    if elev_summary is not None:
        print(f"   Found elevation data for {len(elev_summary)} pipes")
        print("\n   Sample elevation statistics:")
        print(elev_summary.head(3))
    
    # Access equipment data
    print("\n3. Accessing Equipment Tables:")
    sections = parser.list_sections()
    for section in sections:
        if 'Table' in section:
            data = parser.get_section(section)
            if data is not None and not data.empty:
                print(f"   • {section}: {len(data)} entries")
    
    # Access detailed segments for a specific pipe
    print("\n4. Accessing Detailed Pipe Segments:")
    detailed = parser.get_pipe_elevations_detailed()
    if detailed:
        first_pipe = list(detailed.keys())[0]
        pipe_segments = detailed[first_pipe]
        print(f"   Pipe: {first_pipe}")
        print(f"   Segments: {len(pipe_segments)}")
        print("\n   First 3 segments:")
        print(pipe_segments.head(3))
    
    print("\n✓ Advanced example complete!")


def custom_analysis_example():
    """Example of custom data analysis"""
    print("\n" + "="*60)
    print("CUSTOM ANALYSIS EXAMPLE")
    print("="*60)
    
    # Initialize parser
    parser = inpParser3("your_network.inp")
    
    # Get pipe data
    pipes = parser.get_pipe_detail_summary()
    elevations = parser.get_pipe_elevations_summary()
    
    if pipes is not None:
        # Find pipes with largest diameter
        print("\n1. Top 5 pipes by diameter:")
        diameter_col = [col for col in pipes.columns if 'Diameter' in col][0]
        top_diameter = pipes.nlargest(5, diameter_col)
        print(top_diameter[[pipes.columns[0], diameter_col]])
        
        # Find longest pipes
        print("\n2. Top 5 longest pipes:")
        length_col = [col for col in pipes.columns if 'Length' in col][0]
        longest = pipes.nlargest(5, length_col)
        print(longest[[pipes.columns[0], length_col]])
    
    if elevations is not None:
        # Find pipes with largest elevation change
        print("\n3. Top 5 pipes by elevation change:")
        elev_change_col = [col for col in elevations.columns if 'Change' in col][0]
        top_change = elevations.nlargest(5, elev_change_col)
        print(top_change[[elev_change_col]])
        
        # Statistics
        print("\n4. Network elevation statistics:")
        start_col = [col for col in elevations.columns if 'Start' in col][0]
        end_col = [col for col in elevations.columns if 'End' in col][0]
        print(f"   Min start elevation: {elevations[start_col].min():.2f} feet")
        print(f"   Max end elevation: {elevations[end_col].max():.2f} feet")
        print(f"   Avg elevation change: {elevations[elev_change_col].mean():.2f} feet")
    
    print("\n✓ Custom analysis complete!")


def export_options_example():
    """Example of different export options"""
    print("\n" + "="*60)
    print("EXPORT OPTIONS EXAMPLE")
    print("="*60)
    
    parser = inpParser3("your_network.inp")
    
    # Option 1: Export without detailed segments (recommended)
    print("\n1. Exporting without detailed segments...")
    parser.export_to_excel("output_summary.xlsx", include_detailed_segments=False)
    
    # Option 2: Export with detailed segments (for small networks)
    print("\n2. Exporting with detailed segments...")
    parser.export_to_excel("output_detailed.xlsx", include_detailed_segments=True)
    
    # Option 3: Custom path
    print("\n3. Exporting to custom path...")
    parser.export_to_excel("/custom/path/results.xlsx", include_detailed_segments=False)
    
    print("\n✓ Export examples complete!")


if __name__ == "__main__":
    # Run the example you want
    print("\nINP PARSER EXAMPLES")
    print("Uncomment the example you want to run:\n")
    
    # Uncomment one of these:
    # basic_example()
    # advanced_example()
    # custom_analysis_example()
    # export_options_example()
    
    print("\nTo run an example, uncomment it in the __main__ section.")
