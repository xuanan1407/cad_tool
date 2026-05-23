# shape_calculator.py
# Shape calculation utilities
# Copyright (c) 2026 Tran Xuan An

import math


def calculate_perimeter(polygon):
    """Calculate perimeter of a polygon or circle"""
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


def calculate_statistics(polygons):
    """Calculate statistics grouped by shape name"""
    stats = {}
    
    for polygon in polygons:
        name = polygon.get('name', 'Unknown')
        area = polygon.get('area', 0)
        perimeter = calculate_perimeter(polygon)
        
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
    
    return stats
