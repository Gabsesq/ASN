#!/usr/bin/env python3
"""
Test script for Murdochs ASN processing with carton labels.

This demonstrates the new two-step workflow:
1. Process order file to create ASN
2. Add carton labels from EDI file
"""

import os
from processors.MurdochsASN import process_murdochs_asn_with_labels

def test_step1_only():
    """Test Step 1: Process order file only"""
    print("=== Testing Step 1 Only ===")
    
    # Example order file path (you'll need to provide an actual file)
    order_file = "uploads/example_order.xlsx"  # Replace with actual file path
    
    if os.path.exists(order_file):
        try:
            asn_file, po_number = process_murdochs_asn_with_labels(order_file)
            print(f"✅ Step 1 completed successfully!")
            print(f"   ASN file: {asn_file}")
            print(f"   PO Number: {po_number}")
        except Exception as e:
            print(f"❌ Step 1 failed: {e}")
    else:
        print(f"⚠️  Order file not found: {order_file}")
        print("   Please provide a valid order file path")

def test_step1_and_step2():
    """Test both steps: Process order file and add carton labels"""
    print("\n=== Testing Step 1 + Step 2 ===")
    
    # Example file paths (you'll need to provide actual files)
    order_file = "uploads/example_order.xlsx"  # Replace with actual file path
    edi_file = "uploads/example_edi_labels.xlsx"  # Replace with actual file path
    
    if os.path.exists(order_file) and os.path.exists(edi_file):
        try:
            final_asn, po_number = process_murdochs_asn_with_labels(order_file, edi_file)
            print(f"✅ Both steps completed successfully!")
            print(f"   Final ASN file: {final_asn}")
            print(f"   PO Number: {po_number}")
        except Exception as e:
            print(f"❌ Processing failed: {e}")
    else:
        print(f"⚠️  Files not found:")
        print(f"   Order file: {order_file}")
        print(f"   EDI file: {edi_file}")
        print("   Please provide valid file paths")

def main():
    """Main test function"""
    print("Murdochs ASN Processing with Carton Labels - Test Script")
    print("=" * 60)
    
    # Test Step 1 only
    test_step1_only()
    
    # Test both steps
    test_step1_and_step2()
    
    print("\n" + "=" * 60)
    print("Test completed!")
    print("\nTo use this functionality:")
    print("1. Place your order file in the uploads/ folder")
    print("2. Place your EDI file with carton labels in the uploads/ folder")
    print("3. Update the file paths in this script")
    print("4. Run: python test_murdochs_with_labels.py")

if __name__ == "__main__":
    main() 