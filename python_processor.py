# python_processor.py
# CAD Column Inspector Pro - Python Processor
# Copyright (c) 2026 Tran Xuan An

import sys
import json
import os
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import math


class PolygonProcessor:
    """Process polygon data from AutoLISP and export to Excel"""
    
    def __init__(self, json_file):
        self.json_file = json_file
        self.polygons = []
        self.output_folder = os.path.dirname(json_file)
        
    def load_data(self):
        """Load polygon data from JSON file"""
        try:
            with open(self.json_file, 'r') as f:
                data = json.load(f)
            
            # Handle both single object (old format) and array (new format)
            if isinstance(data, list):
                self.polygons = data
            else:
                # Old format: single polygon object
                self.polygons = [data]
            
            print(f"✓ Loaded {len(self.polygons)} polygon(s) from: {self.json_file}")
            return True
        except Exception as e:
            print(f"✗ Error loading data: {e}")
            return False
    
    def calculate_perimeter(self, polygon):
        """Calculate perimeter of a polygon"""
        shape_type = polygon.get('type', '')
        
        # For circles
        if shape_type == 'circle':
            radius = polygon.get('radius', 0)
            return 2 * math.pi * radius
        
        # For polygons
        points = polygon.get('points', [])
        if len(points) < 2:
            return 0
        
        perimeter = 0
        for i in range(len(points)):
            p1 = points[i]
            p2 = points[(i + 1) % len(points)]  # Wrap around to first point
            
            # Calculate 2D distance
            dx = p2[0] - p1[0]
            dy = p2[1] - p1[1]
            distance = math.sqrt(dx * dx + dy * dy)
            perimeter += distance
        
        return perimeter
    
    def process(self):
        """Process polygon data"""
        if not self.polygons:
            print("✗ No data to process")
            return False
        
        print("\n=== Processing Shape Data ===")
        for idx, poly in enumerate(self.polygons, start=1):
            perimeter = self.calculate_perimeter(poly)
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
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(
                self.output_folder,
                f"Column_Inspector_{timestamp}.xlsx"
            )
            
            # Create workbook
            wb = Workbook()
            
            # Create Summary sheet first
            ws_summary = wb.active
            ws_summary.title = "Summary"
            self._create_summary_sheet(ws_summary)
            
            # Create Detail sheet
            ws_detail = wb.create_sheet("Shape Details")
            self._create_detail_sheet(ws_detail)
            
            # Save workbook
            wb.save(output_file)
            print(f"\n✓ Excel file created with {len(self.polygons)} shape(s): {output_file}")
            print(f"✓ Summary sheet created with statistics by name")
            
            return output_file
            
        except Exception as e:
            print(f"✗ Error creating Excel file: {e}")
            return None
    
    def _create_summary_sheet(self, ws):
        """Create summary statistics sheet"""
        # Group shapes by name
        stats = {}
        for polygon in self.polygons:
            name = polygon.get('name', 'Unknown')
            area = polygon.get('area', 0)
            perimeter = self.calculate_perimeter(polygon)
            
            if name not in stats:
                stats[name] = {
                    'count': 0,
                    'total_area': 0,
                    'total_perimeter': 0,
                    'areas': [],
                    'perimeters': []
                }
            
            stats[name]['count'] += 1
            stats[name]['total_area'] += area
            stats[name]['total_perimeter'] += perimeter
            stats[name]['areas'].append(area)
            stats[name]['perimeters'].append(perimeter)
        
        # Styles
        title_font = Font(bold=True, size=16, color="FFFFFF")
        title_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(bold=True, size=12, color="FFFFFF")
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Title
        ws.merge_cells('A1:F1')
        cell = ws.cell(row=1, column=1)
        cell.value = "SHAPE STATISTICS SUMMARY"
        cell.font = title_font
        cell.fill = title_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30
        
        # Headers
        headers = ["Name", "Count", "Total Area", "Avg Area", "Total Perimeter", "Avg Perimeter"]
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=3, column=col)
            cell.value = header
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
        
        # Data
        row = 4
        for name in sorted(stats.keys()):
            stat = stats[name]
            avg_area = stat['total_area'] / stat['count']
            avg_perimeter = stat['total_perimeter'] / stat['count']
            
            ws.cell(row=row, column=1).value = name
            ws.cell(row=row, column=1).border = border
            
            ws.cell(row=row, column=2).value = stat['count']
            ws.cell(row=row, column=2).alignment = Alignment(horizontal="center")
            ws.cell(row=row, column=2).border = border
            
            ws.cell(row=row, column=3).value = round(stat['total_area'], 2)
            ws.cell(row=row, column=3).alignment = Alignment(horizontal="right")
            ws.cell(row=row, column=3).border = border
            
            ws.cell(row=row, column=4).value = round(avg_area, 2)
            ws.cell(row=row, column=4).alignment = Alignment(horizontal="right")
            ws.cell(row=row, column=4).border = border
            
            ws.cell(row=row, column=5).value = round(stat['total_perimeter'], 2)
            ws.cell(row=row, column=5).alignment = Alignment(horizontal="right")
            ws.cell(row=row, column=5).border = border
            
            ws.cell(row=row, column=6).value = round(avg_perimeter, 2)
            ws.cell(row=row, column=6).alignment = Alignment(horizontal="right")
            ws.cell(row=row, column=6).border = border
            
            row += 1
        
        # Totals row
        row += 1
        ws.cell(row=row, column=1).value = "TOTAL"
        ws.cell(row=row, column=1).font = Font(bold=True)
        ws.cell(row=row, column=1).fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        ws.cell(row=row, column=1).border = border
        
        total_count = sum(s['count'] for s in stats.values())
        total_area = sum(s['total_area'] for s in stats.values())
        total_perimeter = sum(s['total_perimeter'] for s in stats.values())
        
        ws.cell(row=row, column=2).value = total_count
        ws.cell(row=row, column=2).font = Font(bold=True)
        ws.cell(row=row, column=2).fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        ws.cell(row=row, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=2).border = border
        
        ws.cell(row=row, column=3).value = round(total_area, 2)
        ws.cell(row=row, column=3).font = Font(bold=True)
        ws.cell(row=row, column=3).fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        ws.cell(row=row, column=3).alignment = Alignment(horizontal="right")
        ws.cell(row=row, column=3).border = border
        
        ws.cell(row=row, column=4).value = round(total_area / total_count if total_count > 0 else 0, 2)
        ws.cell(row=row, column=4).font = Font(bold=True)
        ws.cell(row=row, column=4).fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        ws.cell(row=row, column=4).alignment = Alignment(horizontal="right")
        ws.cell(row=row, column=4).border = border
        
        ws.cell(row=row, column=5).value = round(total_perimeter, 2)
        ws.cell(row=row, column=5).font = Font(bold=True)
        ws.cell(row=row, column=5).fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        ws.cell(row=row, column=5).alignment = Alignment(horizontal="right")
        ws.cell(row=row, column=5).border = border
        
        ws.cell(row=row, column=6).value = round(total_perimeter / total_count if total_count > 0 else 0, 2)
        ws.cell(row=row, column=6).font = Font(bold=True)
        ws.cell(row=row, column=6).fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        ws.cell(row=row, column=6).alignment = Alignment(horizontal="right")
        ws.cell(row=row, column=6).border = border
        
        # Adjust column widths
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        ws.column_dimensions['E'].width = 18
        ws.column_dimensions['F'].width = 18
    
    def _create_detail_sheet(self, ws):
        """Create detailed shape data sheet"""
    def _create_detail_sheet(self, ws):
        """Create detailed shape data sheet"""
        # Header style
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=12)
        
        row = 1
        
        # Process each polygon
        for poly_idx, polygon in enumerate(self.polygons, start=1):
            # Calculate perimeter
            perimeter = self.calculate_perimeter(polygon)
            
            # Section header
            if poly_idx > 1:
                row += 2  # Add space between shapes
            
            ws.cell(row=row, column=1).value = f"SHAPE #{poly_idx}"
            ws.cell(row=row, column=1).font = Font(bold=True, size=14, color="FF0000")
            row += 1
            
            # Property headers
            headers = ["Property", "Value"]
            for col, header in enumerate(headers, start=1):
                cell = ws.cell(row=row, column=col)
                cell.value = header
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            row += 1
            
            # Write data
            ws.cell(row=row, column=1).value = "Name"
            ws.cell(row=row, column=2).value = polygon.get('name', 'Unknown')
            ws.cell(row=row, column=2).font = Font(bold=True, color="0000FF")
            row += 1
            
            ws.cell(row=row, column=1).value = "Type"
            ws.cell(row=row, column=2).value = polygon.get('type', 'N/A')
            row += 1
            
            ws.cell(row=row, column=1).value = "Point Count"
            ws.cell(row=row, column=2).value = polygon.get('point_count', 0)
            row += 1
            
            ws.cell(row=row, column=1).value = "Area (sq units)"
            ws.cell(row=row, column=2).value = round(polygon.get('area', 0), 2)
            row += 1
            
            ws.cell(row=row, column=1).value = "Perimeter (units)"
            ws.cell(row=row, column=2).value = round(perimeter, 2)
            row += 1
            
            # Add radius for circles
            if polygon.get('radius'):
                ws.cell(row=row, column=1).value = "Radius"
                ws.cell(row=row, column=2).value = round(polygon.get('radius', 0), 2)
                row += 1
            
            centroid = polygon.get('centroid', [0, 0])
            ws.cell(row=row, column=1).value = "Centroid X"
            ws.cell(row=row, column=2).value = round(centroid[0], 2)
            row += 1
            
            ws.cell(row=row, column=1).value = "Centroid Y"
            ws.cell(row=row, column=2).value = round(centroid[1], 2)
            row += 1
            
            ws.cell(row=row, column=1).value = "Drawing"
            ws.cell(row=row, column=2).value = polygon.get('drawing', 'N/A')
            row += 1
            
            ws.cell(row=row, column=1).value = "Timestamp"
            ws.cell(row=row, column=2).value = polygon.get('timestamp', 'N/A')
            row += 2
            
            # Points section
            ws.cell(row=row, column=1).value = "Points:"
            ws.cell(row=row, column=1).font = Font(bold=True)
            row += 1
            
            # Points header
            point_headers = ["Point #", "X", "Y", "Z"]
            for col, header in enumerate(point_headers, start=1):
                cell = ws.cell(row=row, column=col)
                cell.value = header
                cell.font = Font(bold=True)
            row += 1
            
            # Points data
            points = polygon.get('points', [])
            for idx, point in enumerate(points, start=1):
                ws.cell(row=row, column=1).value = idx
                ws.cell(row=row, column=2).value = round(point[0], 2)
                ws.cell(row=row, column=3).value = round(point[1], 2)
                ws.cell(row=row, column=4).value = round(point[2], 2) if len(point) > 2 else 0
                row += 1
        
        # Adjust column widths
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15


def main():
    """Main entry point"""
    print("=" * 60)
    print("   CAD Column Inspector Pro - Python Processor")
    print("=" * 60)
    
    # Check arguments
    if len(sys.argv) < 2:
        print("\nUsage: python python_processor.py <json_file>")
        sys.exit(1)
    
    json_file = sys.argv[1]
    
    # Check file exists
    if not os.path.exists(json_file):
        print(f"✗ Error: File not found: {json_file}")
        sys.exit(1)
    
    # Process
    processor = PolygonProcessor(json_file)
    
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
