# cad_processor.py
import math
from config import RECTANGLE_TOLERANCE

class CadProcessor:
    @staticmethod
    def is_rectangle(points, tolerance=RECTANGLE_TOLERANCE):
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

    @staticmethod
    def extract_rectangle_info(pline):
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

    @staticmethod
    def extract_circle_info(circle):
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

    @staticmethod
    def process_rectangles_from_dxf(msp):
        """Extract rectangles from DXF modelspace"""
        rect_count = 0
        rectangles = []
        
        # Method 1: LWPOLYLINE closed with 4 points
        for pline in msp.query('LWPOLYLINE'):
            points = list(pline.get_points())
            is_closed = pline.closed or (len(points) > 2 and points[0] == points[-1])
            
            if not is_closed:
                continue
            
            unique_points = points[:4] if len(points) == 5 else points
            
            if len(unique_points) == 4:
                info = CadProcessor.extract_rectangle_info(pline)
                if info:
                    rect_count += 1
                    rectangles.append(info)
        
        # Method 2: Also check POLYLINE (old format)
        for pline in msp.query('POLYLINE'):
            points = list(pline.get_points())
            is_closed = pline.is_closed or (len(points) > 2 and points[0] == points[-1])
            
            if is_closed and len(points) >= 4:
                unique_points = points[:4] if len(points) == 5 else points[:4]
                if len(unique_points) == 4:
                    info = CadProcessor.extract_rectangle_info(pline)
                    if info:
                        rect_count += 1
                        rectangles.append(info)
        
        return rect_count, rectangles

    @staticmethod
    def process_circles_from_dxf(msp):
        """Extract circles from DXF modelspace"""
        circle_count = 0
        circles = []
        
        # Method 1: CIRCLE entity
        for circle in msp.query('CIRCLE'):
            info = CadProcessor.extract_circle_info(circle)
            if info:
                circle_count += 1
                circles.append(info)
        
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
                circles.append({
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
                circles.append({
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
        
        return circle_count, circles