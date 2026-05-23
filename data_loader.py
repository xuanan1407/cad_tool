# data_loader.py
# JSON data loading utilities
# Copyright (c) 2026 Tran Xuan An

import json
import os


class DataLoader:
    """Load and validate polygon data from JSON"""
    
    def __init__(self, json_file):
        self.json_file = json_file
        self.polygons = []
        
    def load(self):
        """Load polygon data from JSON file"""
        if not os.path.exists(self.json_file):
            print(f"✗ Error: File not found: {self.json_file}")
            return False
        
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
            
        except json.JSONDecodeError as e:
            print(f"✗ Error: Invalid JSON format: {e}")
            return False
        except Exception as e:
            print(f"✗ Error loading data: {e}")
            return False
    
    def get_polygons(self):
        """Get loaded polygons"""
        return self.polygons
