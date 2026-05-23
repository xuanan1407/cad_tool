# main.py
# CAD Column Inspector Pro - Main Entry Point
# Copyright (c) 2026 Tran Xuan An

import sys
import os
from data_loader import DataLoader
from excel_exporter import ExcelExporter
from shape_calculator import calculate_perimeter


class ShapeProcessor:
    """Process polygon data from AutoLISP and export to Excel"""
    
    def __init__(self, json_file):
        self.json_file = json_file
        self.output_folder = os.path.dirname(json_file)
        self.loader = DataLoader(json_file)
        self.polygons = []
    
    def load_data(self):
        """Load polygon data from JSON file"""
        if self.loader.load():
            self.polygons = self.loader.get_polygons()
            return True
        return False
    
    def process(self):
        """Process polygon data and display summary"""
        if not self.polygons:
            print("✗ No data to process")
            return False
        
        print("\n=== Processing Shape Data ===")
        for idx, poly in enumerate(self.polygons, start=1):
            perimeter = calculate_perimeter(poly)
            print(f"\nShape #{idx}:")
            print(f"  Name: {poly.get('name', 'Unknown')}")
            print(f"  Type: {poly.get('type')}")
            print(f"  Point Count: {poly.get('point_count')}")
            print(f"  Area: {poly.get('area'):.2f} sq units")
            print(f"  Perimeter: {perimeter:.2f} units")
            if poly.get('radius'):
                print(f"  Radius: {poly.get('radius'):.2f} units")
            print(f"  Centroid: {poly.get('centroid')}")
            print(f"  Drawing: {poly.get('drawing')}")
        
        return True
    
    def export_to_excel(self):
        """Export data to Excel file"""
        exporter = ExcelExporter(self.polygons, self.output_folder)
        return exporter.export()


def main():
    """Main entry point"""
    print("=" * 60)
    print("   CAD Column Inspector Pro - Python Processor")
    print("=" * 60)
    
    # Check arguments
    if len(sys.argv) < 2:
        print("\nUsage: python main.py <json_file>")
        sys.exit(1)
    
    json_file = sys.argv[1]
    
    # Process
    processor = ShapeProcessor(json_file)
    
    if not processor.load_data():
        sys.exit(1)
    
    if not processor.process():
        sys.exit(1)
    
    output_file = processor.export_to_excel()
    
    if output_file:
        print("\n✓ Processing completed successfully!")
        print("=" * 60)
        sys.exit(0)
    else:
        print("\n✗ Processing failed!")
        sys.exit(1)


if __name__ == "__main__":
    main()
