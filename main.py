# main.py
# CAD Column Inspector Pro - Main Entry Point
# Copyright (c) 2026 Tran Xuan An

import sys
import os
import json
import math
from data_loader import DataLoader
from excel_exporter import ExcelExporter
from shape_calculator import calculate_perimeter


def filter_data(json_filepath, tolerance=20.0):
    """
    Reads JSON file and filters duplicate shapes based on centroid distance.
    Keeps the first scanned shape and removes subsequent duplicates.
    """
    if not os.path.exists(json_filepath):
        print(f"✗ Data file not found: {json_filepath}")
        return

    # 1. Read and load data from JSON file
    try:
        with open(json_filepath, 'r', encoding='utf-8') as f:
            shapes = json.load(f)
    except json.JSONDecodeError:
        print(f"✗ Error: File '{json_filepath}' is not a valid JSON structure.")
        return
    except Exception as e:
        print(f"✗ Error reading file: {e}")
        return

    if not isinstance(shapes, list):
        print("✗ Error: JSON root structure is not an Array.")
        return

    unique_shapes = []
    duplicate_count = 0

    # 2. Duplicate detection algorithm
    for new_shape in shapes:
        is_duplicate = False
        new_centroid = new_shape.get("centroid")
        
        # If centroid data is missing, skip validation and keep raw data
        if not new_centroid or len(new_centroid) < 2:
            unique_shapes.append(new_shape)
            continue

        for existing_shape in unique_shapes:
            exist_centroid = existing_shape.get("centroid")
            
            # Calculate Euclidean distance between two centroids
            dist = math.sqrt(
                (new_centroid[0] - exist_centroid[0]) ** 2 + 
                (new_centroid[1] - exist_centroid[1]) ** 2
            )
            
            # If distance is within the allowed tolerance -> Mark as duplicate
            if dist < tolerance:
                is_duplicate = True
                break
        
        if not is_duplicate:
            unique_shapes.append(new_shape)
        else:
            duplicate_count += 1

    # 3. Normalize IDs sequentially (e.g., Column_001, Column_002,...)
    for index, shape in enumerate(unique_shapes, start=1):
        name = shape.get("name", "Column")
        shape["id"] = f"{name}_{index:03d}"

    # 4. Write back clean data with indentation (Pretty Print)
    try:
        with open(json_filepath, 'w', encoding='utf-8') as f:
            json.dump(unique_shapes, f, indent=2, ensure_ascii=False)
    except Exception as e:
        print(f"✗ Failed to save cleaned data to file: {e}")
        return

    # 5. Output processing results to Terminal
    print("\n" + "="*40)
    print("      DATA FILTERING RESULTS (PYTHON)     ")
    print("="*40)
    print(f" ✓ Total shapes imported:      {len(shapes)}")
    print(f" ✓ Duplicate shapes removed:   {duplicate_count}")
    print(f" ✓ Unique shapes remaining:    {len(unique_shapes)}")
    print(f" ✓ File status:                Cleaned & Sync'd IDs.")
    print("="*40 + "\n")

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

    filter_data(json_file, tolerance=20.0)
    
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
