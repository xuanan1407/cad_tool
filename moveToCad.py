import ezdxf
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pyautocad import Autocad, APoint
import time
import array
import pythoncom
from win32com.client import VARIANT
import win32com.client
import comtypes.automation as automation
import ctypes
import math

class CadColumnInspector:
    def __init__(self, root):
        self.root = root
        self.root.title("CAD Column Inspector Pro - Automation Tool")
        self.root.geometry("900x700")
        
        self.acad = None
        self.dxf_doc = None
        self.all_data = []  # Store all detected objects

        # --- Control UI ---
        top_frame = tk.Frame(root, pady=10)
        top_frame.pack(side=tk.TOP, fill=tk.X)

        btn_frame1 = tk.Frame(top_frame)
        btn_frame1.pack(side=tk.LEFT, padx=10)
        
        tk.Button(btn_frame1, text="📂 1. Load DXF File", command=self.load_dxf, width=20, bg="#e1e1e1").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame1, text="🔗 2. Connect AutoCAD", command=self.connect_to_cad, width=20, bg="#d4edda").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame1, text="📊 3. Export Excel", command=self.export_to_excel, width=15).pack(side=tk.LEFT, padx=5)
        
        # Filter frame
        filter_frame = tk.LabelFrame(top_frame, text="Filter", padx=10, pady=5)
        filter_frame.pack(side=tk.RIGHT, padx=10)
        
        self.show_rect = tk.BooleanVar(value=True)
        self.show_circle = tk.BooleanVar(value=True)
        
        tk.Checkbutton(filter_frame, text="Show Rectangles", variable=self.show_rect, 
                      command=self.refresh_display).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(filter_frame, text="Show Circles", variable=self.show_circle,
                      command=self.refresh_display).pack(side=tk.LEFT, padx=5)

        # --- Data display table (Treeview) ---
        columns = ("ID", "Type", "Layer", "Size (WxH/Dia)", "Coordinate (X,Y)", "Area")
        self.tree = ttk.Treeview(root, columns=columns, show='headings', height=20)
        
        # Define column headings
        col_widths = {"ID": 60, "Type": 100, "Layer": 150, "Size (WxH/Dia)": 120, "Coordinate (X,Y)": 150, "Area": 120}
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor=tk.CENTER, width=col_widths.get(col, 120))
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5, side=tk.LEFT)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        
        # Double-click event to jump to CAD
        self.tree.bind("<Double-1>", self.zoom_to_cad_object)

        # --- Status bar ---
        self.status_var = tk.StringVar(value="Status: Ready")
        status_bar = tk.Label(root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def connect_to_cad(self):
        """Connect to a running AutoCAD instance on Windows"""
        try:
            self.acad = Autocad(create_if_not_exists=True)
            self.status_var.set(f"Connected to: {self.acad.doc.Name}")
            messagebox.showinfo("Success", "Successfully connected to AutoCAD!")
        except Exception as e:
            messagebox.showerror("Error", f"Cannot connect to AutoCAD. Make sure it's open!\n{e}")

    def is_rectangle(self, points, tolerance=0.01):
        """Check if 4 points form a rectangle"""
        if len(points) != 4:
            return False
        
        # Calculate all distances between points
        dists = []
        for i in range(4):
            for j in range(i+1, 4):
                dx = points[i][0] - points[j][0]
                dy = points[i][1] - points[j][1]
                dists.append(dx*dx + dy*dy)
        
        dists.sort()
        # For rectangle: 4 sides equal (opposite sides) + 2 diagonals equal
        # Should have 4 equal smallest distances (sides) and 2 equal largest distances (diagonals)
        return (abs(dists[0] - dists[1]) < tolerance and 
                abs(dists[1] - dists[2]) < tolerance and 
                abs(dists[2] - dists[3]) < tolerance and
                abs(dists[4] - dists[5]) < tolerance and
                dists[4] > dists[3])

    def extract_rectangle_info(self, pline):
        """Extract rectangle information from LWPOLYLINE"""
        points = list(pline.get_points())
        # Remove duplicate first point if exists
        if len(points) == 5 and points[0] == points[-1]:
            points = points[:4]
        
        if len(points) != 4:
            return None
        
        xs = [p[0] for p in points]
        ys = [p[1] for p in points]
        
        w = round(max(xs) - min(xs), 2)
        h = round(max(ys) - min(ys), 2)
        cx = round((max(xs) + min(xs)) / 2, 2)
        cy = round((max(ys) + min(ys)) / 2, 2)
        area = round(w * h, 2)
        
        return {
            'type': 'Rectangle',
            'layer': pline.dxf.layer,
            'size': f"{w} x {h}",
            'width': w,
            'height': h,
            'cx': cx,
            'cy': cy,
            'area': area,
            'raw_data': pline
        }

    def extract_circle_info(self, circle):
        """Extract circle information from CIRCLE entity"""
        center = circle.dxf.center
        radius = round(circle.dxf.radius, 2)
        cx = round(center.x, 2)
        cy = round(center.y, 2)
        diameter = round(radius * 2, 2)
        area = round(math.pi * radius * radius, 2)
        
        return {
            'type': 'Circle',
            'layer': circle.dxf.layer,
            'size': f"D={diameter}",
            'radius': radius,
            'diameter': diameter,
            'cx': cx,
            'cy': cy,
            'area': area,
            'raw_data': circle
        }

    def load_dxf(self):
        """Read DXF file and extract rectangles and circles"""
        file_path = filedialog.askopenfilename(filetypes=[("CAD DXF Files", "*.dxf")])
        if not file_path:
            return

        try:
            self.status_var.set("Loading DXF file...")
            self.dxf_doc = ezdxf.readfile(file_path, errors='ignore')
            msp = self.dxf_doc.modelspace()
            
            # Clear old data
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.all_data = []

            rect_count = 0
            circle_count = 0
            
            # ========== 1. FIND RECTANGLES ==========
            # Method 1: LWPOLYLINE closed with 4 points
            for pline in msp.query('LWPOLYLINE'):
                points = list(pline.get_points())
                is_closed = pline.closed or (len(points) > 2 and points[0] == points[-1])
                
                # Check if closed
                if not is_closed:
                    continue
                
                # Get unique points
                unique_points = points[:4] if len(points) == 5 else points
                
                # Must have exactly 4 points for rectangle
                if len(unique_points) == 4:
                    # Check if it's a rectangle (angles ~90 degrees)
                    info = self.extract_rectangle_info(pline)
                    if info:
                        rect_count += 1
                        self.all_data.append(info)
            
            # Method 2: Also check POLYLINE (old format)
            for pline in msp.query('POLYLINE'):
                points = list(pline.get_points())
                is_closed = pline.is_closed or (len(points) > 2 and points[0] == points[-1])
                
                if is_closed and len(points) >= 4:
                    # Get first 4 points
                    unique_points = points[:4] if len(points) == 5 else points[:4]
                    if len(unique_points) == 4:
                        info = self.extract_rectangle_info(pline)
                        if info:
                            rect_count += 1
                            self.all_data.append(info)
            
            # ========== 2. FIND CIRCLES ==========
            # Method 1: CIRCLE entity
            for circle in msp.query('CIRCLE'):
                info = self.extract_circle_info(circle)
                if info:
                    circle_count += 1
                    self.all_data.append(info)
            
            # Method 2: ARC that forms a full circle (angle 360 degrees)
            for arc in msp.query('ARC'):
                if abs(arc.dxf.end_angle - arc.dxf.start_angle) >= 359.9:
                    center = (arc.dxf.center.x, arc.dxf.center.y)
                    radius = round(arc.dxf.radius, 2)
                    cx = round(center[0], 2)
                    cy = round(center[1], 2)
                    diameter = round(radius * 2, 2)
                    area = round(math.pi * radius * radius, 2)
                    
                    circle_count += 1
                    self.all_data.append({
                        'type': 'Circle (Arc)',
                        'layer': arc.dxf.layer,
                        'size': f"D={diameter}",
                        'radius': radius,
                        'diameter': diameter,
                        'cx': cx,
                        'cy': cy,
                        'area': area,
                        'raw_data': arc
                    })
            
            # Method 3: ELLIPSE that is a circle (equal axes)
            for ellipse in msp.query('ELLIPSE'):
                if abs(ellipse.dxf.major_axis.magnitude - ellipse.dxf.ratio) < 0.001:
                    center = ellipse.dxf.center
                    radius = round(ellipse.dxf.major_axis.magnitude, 2)
                    cx = round(center.x, 2)
                    cy = round(center.y, 2)
                    diameter = round(radius * 2, 2)
                    area = round(math.pi * radius * radius, 2)
                    
                    circle_count += 1
                    self.all_data.append({
                        'type': 'Circle (Ellipse)',
                        'layer': ellipse.dxf.layer,
                        'size': f"D={diameter}",
                        'radius': radius,
                        'diameter': diameter,
                        'cx': cx,
                        'cy': cy,
                        'area': area,
                        'raw_data': ellipse
                    })
            
            # Update status
            self.status_var.set(f"Found: {rect_count} rectangles, {circle_count} circles (Total: {len(self.all_data)})")
            
            # Refresh display with current filters
            self.refresh_display()
            
            if not self.all_data:
                messagebox.showinfo("No Data", "No rectangles or circles found in the DXF file!")
                
        except Exception as e:
            messagebox.showerror("File Error", f"Failed to parse DXF file: {e}")
            self.status_var.set("Error loading file")
    
    def refresh_display(self):
        """Refresh treeview based on current filters"""
        # Clear current display
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Filter and display data
        display_id = 1
        for data in self.all_data:
            # Apply filters
            if 'Rectangle' in data['type'] and not self.show_rect.get():
                continue
            if 'Circle' in data['type'] and not self.show_circle.get():
                continue
            
            # Insert into tree
            coord_str = f"{data['cx']},{data['cy']}"
            self.tree.insert("", tk.END, values=(
                display_id,
                data['type'],
                data['layer'],
                data['size'],
                coord_str,
                data['area']
            ))
            display_id += 1
        
        # Store mapping between display ID and actual data for zoom function
        self.display_data = self.all_data.copy()
        self.status_var.set(f"Displaying: {display_id-1} items (Filtered)")

    def zoom_to_cad_object(self, event):
        if not self.acad:
            messagebox.showwarning("Connection", "Please click 'Connect AutoCAD' first!")
            return
    
        selected_item = self.tree.selection()
        if not selected_item:
            return
    
        values = self.tree.item(selected_item[0], "values")
        display_id = int(values[0])
        coord_raw = values[4]
        x, y = map(float, coord_raw.split(','))
    
        try:
            # Get AutoCAD application directly
            acad_app = win32com.client.GetActiveObject("AutoCAD.Application")
            doc = acad_app.ActiveDocument
            
            # Focus AutoCAD
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.AppActivate("AutoCAD")
            time.sleep(0.3)
            
            # Create a point at the coordinates
            center_point = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [x, y, 0])
            
            # Method 1: Draw a temporary red circle and blink it
            mspace = doc.ModelSpace
            
            # Draw temporary circle
            temp_circle = mspace.AddCircle(center_point, 1.0)
            temp_circle.Color = 1  # Red color
            temp_circle.LineWeight = 40  # Thick line
            
            # Draw crosshair lines
            line1_start = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [x - 2, y, 0])
            line1_end = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [x + 2, y, 0])
            temp_line1 = mspace.AddLine(line1_start, line1_end)
            temp_line1.Color = 1
            
            line2_start = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [x, y - 2, 0])
            line2_end = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [x, y + 2, 0])
            temp_line2 = mspace.AddLine(line2_start, line2_end)
            temp_line2.Color = 1
            
            # Zoom to see the highlight
            acad_app.ZoomCenter(center_point, 10)
            
            # Regenerate to show
            doc.Regen(0)
            
            # Blink effect: change color multiple times
            for _ in range(3):
                time.sleep(0.2)
                temp_circle.Color = 2  # Yellow
                temp_line1.Color = 2
                temp_line2.Color = 2
                doc.Regen(0)
                time.sleep(0.2)
                temp_circle.Color = 1  # Red
                temp_line1.Color = 1
                temp_line2.Color = 1
                doc.Regen(0)
            
            # Keep highlight for 1 second then delete
            time.sleep(0.5)
            
            # Delete temporary objects
            temp_circle.Delete()
            temp_line1.Delete()
            temp_line2.Delete()
            doc.Regen(0)
            
            self.status_var.set(f"✅ Highlighted position: {x}, {y}")
            
        except Exception as e:
            print(f"Connection error: {e}")
            self.status_var.set("❌ Cannot connect to AutoCAD")
            messagebox.showwarning("Connection", "Please make sure AutoCAD is running and try 'Connect AutoCAD' again")

    def export_to_excel(self):
        """Export displayed data to Excel file, grouped by type and size"""
        # Get current displayed data
        data = []
        for child in self.tree.get_children():
            values = self.tree.item(child)["values"]
            data.append({
                'ID': int(values[0]),  # Convert to int
                'Type': values[1],
                'Layer': values[2],
                'Size': values[3],
                'Coordinate_X': float(values[4].split(',')[0]),
                'Coordinate_Y': float(values[4].split(',')[1]),
                'Area': float(values[5])  # Convert to float
            })
        
        if not data:
            messagebox.showwarning("Data", "No data to export!")
            return
    
        # Convert to DataFrame
        df = pd.DataFrame(data)
        
        # Group by Type, Size and Layer
        grouped = df.groupby(['Type', 'Size', 'Layer']).agg(
            Count=('Size', 'size'),
            IDs=('ID', lambda x: ', '.join(map(str, x))),
            Area_avg=('Area', 'mean'),  # Now Area is float, can calculate mean
            Area_min=('Area', 'min'),
            Area_max=('Area', 'max'),
            Coordinates=('Coordinate_X', lambda x: ', '.join([f"({x.iloc[i]},{df.iloc[i]['Coordinate_Y']})" for i in range(len(x))]))
        ).reset_index()
        
        # Round area columns
        grouped['Area_avg'] = grouped['Area_avg'].round(2)
        grouped['Area_min'] = grouped['Area_min'].round(2)
        grouped['Area_max'] = grouped['Area_max'].round(2)
        
        # Save to Excel
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if save_path:
            try:
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    # Sheet 1: Grouped by Type, Size and Layer
                    grouped.to_excel(writer, sheet_name='Summary_by_Type_Size', index=False)
                    
                    # Sheet 2: Detailed data
                    df.to_excel(writer, sheet_name='Detailed_Data', index=False)
                    
                    # Sheet 3: Statistics
                    stats_data = []
                    
                    # Overall statistics
                    stats_data.append(['=== OVERALL STATISTICS ===', ''])
                    stats_data.append(['Total Objects', len(df)])
                    stats_data.append(['Rectangles', len(df[df['Type'].str.contains('Rectangle')])])
                    stats_data.append(['Circles', len(df[df['Type'].str.contains('Circle')])])
                    stats_data.append(['Unique Sizes', df['Size'].nunique()])
                    stats_data.append(['', ''])
                    
                    # Statistics by Type
                    stats_data.append(['=== STATISTICS BY TYPE ===', ''])
                    for obj_type in df['Type'].unique():
                        type_df = df[df['Type'] == obj_type]
                        stats_data.append([f'Type: {obj_type}', ''])
                        stats_data.append(['  Count', len(type_df)])
                        stats_data.append(['  Area - Min', round(type_df['Area'].min(), 2)])
                        stats_data.append(['  Area - Max', round(type_df['Area'].max(), 2)])
                        stats_data.append(['  Area - Avg', round(type_df['Area'].mean(), 2)])
                        stats_data.append(['', ''])
                    
                    # Statistics by Size
                    stats_data.append(['=== STATISTICS BY SIZE ===', ''])
                    size_stats = df.groupby('Size').agg(
                        Count=('Size', 'size'),
                        Type=('Type', lambda x: ', '.join(x.unique())),
                        Area_avg=('Area', 'mean')
                    ).reset_index()
                    size_stats['Area_avg'] = size_stats['Area_avg'].round(2)
                    
                    # Convert to list for Excel
                    stats_data.append(['Size', 'Count', 'Types', 'Avg Area'])
                    for _, row in size_stats.iterrows():
                        stats_data.append([row['Size'], row['Count'], row['Type'], row['Area_avg']])
                    
                    stats_df = pd.DataFrame(stats_data)
                    stats_df.to_excel(writer, sheet_name='Statistics', index=False, header=False)
                    
                messagebox.showinfo("Success", 
                                   f"Report saved to:\n{save_path}\n\n"
                                   f"Summary: {len(grouped)} unique groups\n"
                                   f"Total items: {len(df)}\n"
                                   f"- Rectangles: {len(df[df['Type'].str.contains('Rectangle')])}\n"
                                   f"- Circles: {len(df[df['Type'].str.contains('Circle')])}")
                
                self.status_var.set(f"Exported to: {os.path.basename(save_path)}")
                
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export: {str(e)}")
                self.status_var.set("Export failed!")

if __name__ == "__main__":
    root = tk.Tk()
    app = CadColumnInspector(root)
    root.mainloop()