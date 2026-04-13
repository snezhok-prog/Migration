#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Test script to verify migration structure without API calls
"""
import sys
import os
import pandas as pd
import json
import logging

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _config import EXCEL_FILE_NAME, SCRIPT_DIR
from _logger import setup_logger
from fair_migration import (
    process_fair_sheet, process_mesto_sheet, process_permits_sheet
)

excel_file = os.path.join(SCRIPT_DIR, EXCEL_FILE_NAME)
logger = setup_logger()

def test_migration():
    """Test migration with Excel data"""
    try:
        # Read Excel file  
        print("Reading Excel file...")
        excel_fair = pd.read_excel(excel_file, sheet_name="4. Реестр ярмарок", header=3)
        excel_mesto = pd.read_excel(excel_file, sheet_name="2. Реестр мест", header=3)
        excel_permits = pd.read_excel(excel_file, sheet_name="3. Реестр разрешений", header=3)
        
        print(f"Fair sheet rows: {len(excel_fair)}")
        print(f"Mesto sheet rows: {len(excel_mesto)}")
        print(f"Permits sheet rows: {len(excel_permits)}")
        
        # Test fair sheet
        if len(excel_fair) > 0:
            print("\n=== Testing Fair Sheet ===")
            row = excel_fair.iloc[0].to_dict()
            result = process_fair_sheet(row, logger)
            print(json.dumps(result, ensure_ascii=False, indent=2)[:500] + "...\n")
        
        # Test mesto sheet
        if len(excel_mesto) > 0:
            print("\n=== Testing Mesto Sheet ===")
            row = excel_mesto.iloc[0].to_dict()
            result = process_mesto_sheet(row, logger)
            
            # Check for mapping functions
            if "placeMarketInfo" in result:
                pmi = result["placeMarketInfo"]
                print(f"Market Type: {pmi.get('marketType')}")
                print(f"Market Specialization: {pmi.get('marketSpecialization')}")
                print(f"Status Fair Place: {pmi.get('statusFairPlace')}")
                
            if result.get('marketInfo') and len(result['marketInfo']) > 0:
                print(f"Market Status: {result['marketInfo'][0].get('marketStatus')}")
            
            print("\nFull structure (first 1500 chars):")
            print(json.dumps(result, ensure_ascii=False, indent=2)[:1500] + "...")
        
        # Test permits sheet
        if len(excel_permits) > 0:
            print("\n=== Testing Permits Sheet ===")
            row = excel_permits.iloc[0].to_dict()
            result = process_permits_sheet(row, logger)
            print(json.dumps(result, ensure_ascii=False, indent=2)[:500] + "...\n")
        
        print("✓ All tests completed successfully!")
        
    except Exception as e:
        print(f"✗ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_migration()
