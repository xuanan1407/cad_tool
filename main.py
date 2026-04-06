# main.py
import os
import tkinter as tk
from tkinter import messagebox, filedialog
import ezdxf

from ui_components import ButtonPanel, FilterFrame, DataTable, StatusBar
from cad_processor import CadProcessor
from cad_connector import CadConnector
from excel_exporter import ExcelExporter
from config import WINDOW_SIZE

class CadColumnInspector:
    def __init__(self, root):
        self.root = root
        self.root.title("CAD Column Inspector Pro - Automation Tool")
        self.root.geometry(WINDOW_SIZE)
        
        self.acad_connector = CadConnector()
        self.dxf_doc = None
        self.all_data = []  # Store all detected objects
        self.display_data = []  # Store current display mapping
        
        # Create UI components
        self._create_ui()
    
    def _create_ui(self):
        # Top frame for buttons and filters
        top_frame = tk.Frame(self.root, pady=10)
        top_frame.pack(side=tk.TOP, fill=tk.X)

        # Button panel
        self.button_panel = ButtonPanel(
            top_frame,
            on_load_dxf=self.load_dxf,
            on_connect_cad=self.connect_to_cad,
            on_export=self.export_to_excel,
            on_remove_duplicates=self.remove_duplicates
        )

        # Filter frame
        self.filter_frame = FilterFrame(top_frame, on_filter_change=self.refresh_display)

        # Data table
        self.data_table = DataTable(self.root, on_double_click=self.zoom_to_cad_object)

        # Status bar
        self.status_bar = StatusBar(self.root)

    def remove_duplicates(self):
        """Remove duplicate shapes with the same code and all properties, and shapes whose centers are too close (<20mm)"""
        seen = set()
        unique_data = []
        min_dist_sq = 20 * 20  # 20mm squared
        for data in self.all_data:
            # Use tuple of all relevant properties as key
            key = (
                data.get('type'),
                data.get('layer'),
                data.get('size'),
                data.get('cx'),
                data.get('cy'),
                data.get('area'),
                data.get('code', '')
            )
            if key in seen:
                continue
            cx, cy = data.get('cx'), data.get('cy')
            too_close = False
            for d in unique_data:
                dcx, dcy = d.get('cx'), d.get('cy')
                if dcx is not None and dcy is not None and cx is not None and cy is not None:
                    dist_sq = (cx - dcx) ** 2 + (cy - dcy) ** 2
                    if dist_sq < min_dist_sq:
                        too_close = True
                        break
            if not too_close:
                seen.add(key)
                unique_data.append(data)
        removed = len(self.all_data) - len(unique_data)
        self.all_data = unique_data
        self.refresh_display()
        self.status_bar.set_status(f"Removed {removed} duplicates/overlaps. Remaining: {len(self.all_data)}")
    
    def connect_to_cad(self):
        success, message = self.acad_connector.connect()
        if success:
            self.status_bar.set_status(message)
            messagebox.showinfo("Success", "Successfully connected to AutoCAD!")
        else:
            messagebox.showerror("Error", message)
    
    def load_dxf(self):
        """Read DXF file and extract rectangles and circles, assign nearest code string to each shape"""
        file_path = filedialog.askopenfilename(filetypes=[("CAD DXF Files", "*.dxf")])
        if not file_path:
            return

        try:
            self.status_bar.set_status("Loading DXF file...")
            self.dxf_doc = ezdxf.readfile(file_path, errors='ignore')
            msp = self.dxf_doc.modelspace()

            # Clear old data
            self.data_table.clear()
            self.all_data = []

            # Get all texts (codes)
            texts = CadProcessor.extract_texts_from_dxf(msp)

            # Process rectangles
            rect_count, rectangles = CadProcessor.process_rectangles_from_dxf(msp)
            # Process circles
            circle_count, circles = CadProcessor.process_circles_from_dxf(msp)

            # Combine all shapes
            all_shapes = rectangles + circles

            # For each shape, find nearest code (text)
            for shape in all_shapes:
                cx, cy = shape['cx'], shape['cy']
                min_dist = float('inf')
                nearest_code = ''
                for t in texts:
                    dist = (cx - t['x'])**2 + (cy - t['y'])**2
                    if dist < min_dist:
                        min_dist = dist
                        nearest_code = t['text']
                shape['code'] = nearest_code
            self.all_data = all_shapes

            # Update status
            self.status_bar.set_status(
                f"Found: {rect_count} rectangles, {circle_count} circles (Total: {len(self.all_data)})"
            )

            # Refresh display with current filters
            self.refresh_display()

            if not self.all_data:
                messagebox.showinfo("No Data", "No rectangles or circles found in the DXF file!")

        except Exception as e:
            messagebox.showerror("File Error", f"Failed to parse DXF file: {e}")
            self.status_bar.set_status("Error loading file")
    
    def refresh_display(self):
        """Refresh treeview based on current filters"""
        self.data_table.clear()
        
        filters = self.filter_frame.get_filters()
        display_id = 1
        self.display_data = []
        
        for data in self.all_data:
            # Apply filters
            if 'Rectangle' in data['type'] and not filters['show_rect']:
                continue
            if 'Circle' in data['type'] and not filters['show_circle']:
                continue

            # Insert into tree (thêm cột Code)
            coord_str = f"{data['cx']},{data['cy']}"
            self.data_table.add_item((
                display_id,
                data['type'],
                data['layer'],
                data['size'],
                coord_str,
                data['area'],
                data.get('code', '')
            ))

            # Store mapping
            self.display_data.append(data)
            display_id += 1
        
        self.status_bar.set_status(f"Displaying: {display_id-1} items (Filtered)")
    
    def zoom_to_cad_object(self, event):
        if not self.acad_connector.is_connected():
            messagebox.showwarning("Connection", "Please click 'Connect AutoCAD' first!")
            return
        
        selected_values = self.data_table.get_selected()
        if not selected_values:
            return
        
        display_id = int(selected_values[0])
        coord_raw = selected_values[4]
        x, y = map(float, coord_raw.split(','))
        
        success, message = self.acad_connector.zoom_to_point(x, y)
        if success:
            self.status_bar.set_status(f"✅ {message}")
        else:
            self.status_bar.set_status("❌ Cannot connect to AutoCAD")
            messagebox.showwarning("Connection", "Please make sure AutoCAD is running and try 'Connect AutoCAD' again")
    
    def export_to_excel(self):
        """Export displayed data to Excel file"""
        # Get current displayed data
        items = self.data_table.get_all_items()
        
        if not items:
            messagebox.showwarning("Data", "No data to export!")
            return
        
        # Format data for export
        export_data = []
        for values in items:
            export_data.append({
                'ID': int(values[0]),
                'Type': values[1],
                'Layer': values[2],
                'Size': values[3],
                'Coordinate_X': float(values[4].split(',')[0]),
                'Coordinate_Y': float(values[4].split(',')[1]),
                'Area': float(values[5]),
                'code': values[6] if len(values) > 6 else ''
            })
        
        # Save to Excel
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if save_path:
            success, result = ExcelExporter.export_data(export_data, save_path)
            
            if success:
                messagebox.showinfo("Success", 
                                   f"Report saved to:\n{save_path}\n\n"
                                   f"Summary: {result['unique_groups']} unique groups\n"
                                   f"Total items: {result['total_items']}\n"
                                   f"- Rectangles: {result['rectangles']}\n"
                                   f"- Circles: {result['circles']}")
                self.status_bar.set_status(f"Exported to: {os.path.basename(save_path)}")
            else:
                messagebox.showerror("Export Error", result)
                self.status_bar.set_status("Export failed!")

if __name__ == "__main__":
    root = tk.Tk()
    app = CadColumnInspector(root)
    root.mainloop()