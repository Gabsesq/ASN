# upc_counts.py

counts = {
    # EDI Products
    "850016364883": 24,  # Edi-HJ-PB-LRG
    "850016364913": 24,  # Edi-DR-SP-LRG
    "864178000275": 24,  # Edi-DR-BC-LRG
    "850016364876": 24,  # Edi-HJ-PB-SML
    "850016364890": 12,  # Edi-HJ-PB-FAM
    "850016364982": 24,  # Edi-DR-BC-SML
    "850016364906": 24,  # Edi-DR-SP-SML
    "850016364968": 40,  # TS-Edi-HJ-PB
    "850016364944": 40,  # TS-Edi-STRESS-P
    "850016364951": 40,  # TS-Edi-STRESS-P
    "850016364821": 24,  # EDI-STRESS-PEPP-LRG
    "850016364838": 24,  # EDI-STRESS-PB-SML
    "850016364845": 24,  # EDI-STRESS-PEPP-SML
    "850016364852": 24,  # EDI-STRESS-PB-LRG
    "850016364869": 12,  # EDI-STRESS-PB-FAM

    # Products Starting with Numbers
    "860008203465": 30,  # 300-HJR-HO
    "860008203441": 30,  # 300-SR-HO
    "860008203458": 30,  # 600-SR-HO
    "860008203472": 30,  # 600-HJR-HO
    "860008221988": 30,  # 180-CAT-SR
    "860009592575": 40,  # 150-Mini-Stress
    "860008203403": 30,  # 100-DR-HO
    "860008876744": 6,   # 2IN1-SK-CT

    # Other Products
    "860009592568": 30,  # Post-Bio-GH: POST-BIO-GH
    "860009592551": 30,  # Omega-Alg: OMEGA-ALG
    "860008203410": 30,  # 200-DR-HO
    "860008203427": 30,  # 500-DR-HO
    "860008203434": 30,  # 750-DR-HO
    "860008876775": 30,  # 100-Lipe-Ultra: 100-LIPE-ULTRA
    "860008876768": 30,  # 300-Lipe-Ultra: 300-LIPE-ULTRA
    "860009592513": 30,  # 600-Lipe-Ultra: 600-LIPE-ULTRA
    "861109000304": 18,  # CAP450: 450 CAP
    "850016364586": 18,  # SNT30: SNT30
    "860009592537": 6,  # TS-Itchy-Dry-Shampoo: TS-Itchy & Dry Shampoo
    "860009592520": 6,  # Itchy-Dry-Shampoo-Gallon: ITCHY-DRY SHAMPOO-GALLON
    "860008876720": 6,  # Sensitive-SK-CT: SENSITIVE SK-CT
    "860008876737": 6,  # Conditioner-SK-CT: CONDITIONER SK-CT
    "860009592544": 40,  # TS-2in1-Shampoo: TS-2 IN 1 SHAMPOO
    "860008221971": 30,  # SK-PW-RL: SK-PW-RL
}

def calculate_total_cases(sheet, start_row=15, upc_col=5, qty_col=1):
    """
    Calculates the total number of cases based on UPC and QTY columns.
    
    Args:
        sheet: The Excel sheet object to read data from.
        start_row (int): The row index from which to start reading data.
        upc_col (int): The column index for the UPC.
        qty_col (int): The column index for the quantity.

    Returns:
        int: The total number of cases.
    """
    total_cases = 0

    for row in range(start_row, sheet.nrows):  # Iterate over rows starting at start_row
        try:
            upc = str(int(sheet.cell_value(row, upc_col)))  # Read UPC and convert to string
            qty = int(sheet.cell_value(row, qty_col))       # Read QTY and convert to int

            if upc in counts:
                items_per_case = counts[upc]
                cases = qty // items_per_case  # Calculate the number of cases
                total_cases += cases           # Add to the total cases
                
                # Log message for this row
                print(f"Row {row}: UPC {upc}, QTY {qty}, Items/Case {items_per_case}, Cases {cases}")
            else:
                print(f"Warning: UPC {upc} not found in counts dictionary (Row {row}).")
        except (ValueError, IndexError) as e:
            print(f"Error processing row {row}: {e}")

    print(f"Total calculated cases: {total_cases}")
    return total_cases
