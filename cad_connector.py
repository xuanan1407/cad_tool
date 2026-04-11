# Copyright (c) 2026 Tran Xuan An
# This file is part of CAD Column Inspector Pro.
# Licensed under the MIT License.

# cad_connector.py
import time
import pythoncom
from win32com.client import VARIANT
import win32com.client
from pyautocad import Autocad
from config import (
    COLOR_RED,
    COLOR_YELLOW,
    ZOOM_DISTANCE,
    HIGHLIGHT_RADIUS,
    CROSSHAIR_LENGTH,
    BLINK_DURATION,
    HIGHLIGHT_DURATION,
)


class CadConnector:
    @staticmethod
    def highlight_object(x, y, all_data):
        """Highlight the outline of the shape (rectangle/circle) at (x, y) from all_data."""
        try:
            acad_app = win32com.client.GetActiveObject("AutoCAD.Application")
            doc = acad_app.ActiveDocument
            mspace = doc.ModelSpace

            # Find the correct data object with (x, y)
            target_obj = None
            for d in all_data:
                if abs(d.get("cx", 0) - x) < 1e-3 and abs(d.get("cy", 0) - y) < 1e-3:
                    target_obj = d
                    break
            if not target_obj:
                return

            temp_entity = None
            if "Rectangle" in target_obj["type"]:
                # Draw the outline of the rectangle
                size = target_obj["size"].replace(" ", "")
                if "x" in size:
                    w, h = [float(v) for v in size.split("x")]
                else:
                    w = h = 0
                cx, cy = target_obj["cx"], target_obj["cy"]
                # 4 points in order
                pts = [
                    (cx - w / 2, cy - h / 2, 0),
                    (cx + w / 2, cy - h / 2, 0),
                    (cx + w / 2, cy + h / 2, 0),
                    (cx - w / 2, cy + h / 2, 0),
                    (cx - w / 2, cy - h / 2, 0),
                ]
                arr = VARIANT(
                    pythoncom.VT_ARRAY | pythoncom.VT_R8,
                    [coord for pt in pts for coord in pt],
                )
                temp_entity = mspace.AddPolyline(arr)
                temp_entity.Closed = True
                temp_entity.Color = COLOR_YELLOW
                temp_entity.LineWeight = 40
            elif "Circle" in target_obj["type"]:
                # Draw the outline of the circle
                radius = target_obj.get("radius") or (target_obj.get("diameter", 0) / 2)
                center_point = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [x, y, 0])
                temp_entity = mspace.AddCircle(center_point, radius)
                temp_entity.Color = COLOR_YELLOW
                temp_entity.LineWeight = 40

            doc.Regen(0)
            time.sleep(HIGHLIGHT_DURATION)
            if temp_entity:
                temp_entity.Delete()
                doc.Regen(0)
        except Exception as e:
            import logging

            logging.error(f"Error in highlight_object: {e}")

    def __init__(self):
        self.acad = None

    def connect(self):
        """Connect to a running AutoCAD instance on Windows"""
        try:
            self.acad = Autocad(create_if_not_exists=True)
            return True, f"Connected to: {self.acad.doc.Name}"
        except Exception as e:
            return False, f"Cannot connect to AutoCAD. Make sure it's open!\n{e}"

    def is_connected(self):
        return self.acad is not None

    def zoom_to_point(self, x, y):
        """Zoom to a specific point and highlight it"""
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

            # Draw temporary objects
            mspace = doc.ModelSpace
            temp_circle = mspace.AddCircle(center_point, HIGHLIGHT_RADIUS)
            temp_circle.Color = COLOR_RED
            temp_circle.LineWeight = 40  # Thick line

            # Draw crosshair lines
            line1_start = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8, [x - CROSSHAIR_LENGTH, y, 0]
            )
            line1_end = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8, [x + CROSSHAIR_LENGTH, y, 0]
            )
            temp_line1 = mspace.AddLine(line1_start, line1_end)
            temp_line1.Color = COLOR_RED

            line2_start = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8, [x, y - CROSSHAIR_LENGTH, 0]
            )
            line2_end = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8, [x, y + CROSSHAIR_LENGTH, 0]
            )
            temp_line2 = mspace.AddLine(line2_start, line2_end)
            temp_line2.Color = COLOR_RED

            # Zoom to see the highlight
            acad_app.ZoomCenter(center_point, ZOOM_DISTANCE)

            # Regenerate to show
            doc.Regen(0)

            # Blink effect: change color multiple times
            for _ in range(3):
                time.sleep(BLINK_DURATION)
                temp_circle.Color = COLOR_YELLOW
                temp_line1.Color = COLOR_YELLOW
                temp_line2.Color = COLOR_YELLOW
                doc.Regen(0)
                time.sleep(BLINK_DURATION)
                temp_circle.Color = COLOR_RED
                temp_line1.Color = COLOR_RED
                temp_line2.Color = COLOR_RED
                doc.Regen(0)

            # Keep highlight for a moment then delete
            time.sleep(HIGHLIGHT_DURATION)

            # Delete temporary objects
            temp_circle.Delete()
            temp_line1.Delete()
            temp_line2.Delete()
            doc.Regen(0)

            return True, f"Highlighted position: {x}, {y}"

        except Exception as e:
            return False, f"Cannot connect to AutoCAD: {e}"
