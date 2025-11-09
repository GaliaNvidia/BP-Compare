"""
Sample data generator for testing the PO Line Comparison Tool
Run this script to generate sample Excel files for testing
"""

import pandas as pd
from datetime import datetime, timedelta
import random

def generate_sample_data():
    """Generate sample Excel files for testing"""
    
    # Previous week data
    prev_data = {
        'Purch.doc.': ['PO001', 'PO001', 'PO002', 'PO003', 'PO004', 'PO005'],
        'Item': ['10', '20', '10', '10', '10', '10'],
        'Short text': ['PN-12345', 'PN-67890', 'PN-11111', 'PN-22222', 'PN-33333', 'PN-44444'],
        'Order': ['PWO-001', 'PWO-002', 'PWO-003', 'PWO-004', 'PWO-005', 'PWO-006'],
        'Type': ['Standard', 'Standard', 'Rush', 'Standard', 'Rush', 'Standard'],
        'ComDate': [
            datetime(2024, 11, 1),
            datetime(2024, 11, 2),
            datetime(2024, 11, 3),
            datetime(2024, 11, 4),
            datetime(2024, 11, 5),
            datetime(2024, 10, 25)  # This one will be pushed >7 days
        ]
    }
    
    # Current week data (with changes)
    curr_data = {
        'Purch.doc.': [
            'PO001',  # Same
            'PO001',  # Same
            'PO001',  # NEW - Split line
            'PO002',  # Pushed 3 days
            'PO003',  # Pushed 5 days
            'PO004',  # Pushed 2 days
            'PO005',  # Pushed 10 days - ALERT!
            'PO006'   # New PO
        ],
        'Item': ['10', '20', '30', '10', '10', '10', '10', '10'],
        'Short text': [
            'PN-12345',
            'PN-67890',
            'PN-12345',  # Split - same PN as PO001/10
            'PN-11111',
            'PN-22222',
            'PN-33333',
            'PN-44444',
            'PN-55555'
        ],
        'Order': ['PWO-001', 'PWO-002', 'PWO-001', 'PWO-003', 'PWO-004', 'PWO-005', 'PWO-006', 'PWO-007'],
        'Type': ['Standard', 'Standard', 'Standard', 'Rush', 'Standard', 'Rush', 'Standard', 'Standard'],
        'ComDate': [
            datetime(2024, 11, 1),   # No change
            datetime(2024, 11, 2),   # No change
            datetime(2024, 11, 8),   # Split line
            datetime(2024, 11, 6),   # Pushed 3 days
            datetime(2024, 11, 9),   # Pushed 5 days
            datetime(2024, 11, 7),   # Pushed 2 days
            datetime(2024, 11, 4),   # Pushed 10 days - ALERT!
            datetime(2024, 11, 8)    # New
        ]
    }
    
    # Create DataFrames
    prev_df = pd.DataFrame(prev_data)
    curr_df = pd.DataFrame(curr_data)
    
    # Save to Excel
    prev_file = 'sample_previous_week.xlsx'
    curr_file = 'sample_current_week.xlsx'
    
    prev_df.to_excel(prev_file, index=False)
    curr_df.to_excel(curr_file, index=False)
    
    print("✅ Sample files generated successfully!")
    print(f"   - {prev_file}")
    print(f"   - {curr_file}")
    print("\nExpected Results:")
    print("   - 1 Split line (PO001 Item 30)")
    print("   - 4 Pushed lines (PO002, PO003, PO004, PO005)")
    print("   - 1 Alert (PO005 pushed 10 days)")
    print("\nYou can now upload these files to the Streamlit app for testing.")

if __name__ == "__main__":
    generate_sample_data()

