# Calendar Modules

This directory contains company-specific calendar functionality for ship date calculations and order processing.

## Structure

- `__init__.py` - Package initialization
- `chewy_calendar.py` - Chewy-specific ship date logic with location tracking
- `pet_supermarket_calendar.py` - Pet Supermarket-specific ship date logic
- `README.md` - This documentation file

## Adding a New Company

To add calendar functionality for a new company:

1. Create a new file: `company_name_calendar.py`
2. Implement the required functions:
   - `get_company_ship_date_advanced(file_path, **kwargs)` - Main ship date calculation
   - `get_company_order_details(file_path)` - Extract order details (optional)
   - Any other company-specific functions

3. Update `calendar_helpers.py` to import and use the new module

## Example Company Module Structure

```python
"""
Company-specific calendar functionality and ship date calculations.
"""
import datetime
from openpyxl import load_workbook
import xlrd

def get_company_ship_date_advanced(file_path, **kwargs):
    """
    Advanced company ship date calculation with business rules.
    
    Args:
        file_path: Path to the Excel file
        **kwargs: Additional parameters for future expansion
    
    Returns:
        datetime.datetime: Recommended ship date or None
    """
    # Implementation here
    pass

def get_company_order_details(file_path):
    """
    Extract additional order details from company files.
    
    Args:
        file_path: Path to the Excel file
    
    Returns:
        dict: Order details
    """
    # Implementation here
    pass
```

## Business Rules by Company

### Chewy
- **Ship Date Logic**: If total quantity > 100, ship next day (FedEx)
- **Priority Levels**: 
  - High: > 200 units
  - Medium: 100-200 units  
  - Low: < 100 units
- **Weekend Handling**: Adjust ship date to avoid weekends
- **Location Tracking**: Reads Chewy location from cell B16
- **Additional Info**: 
  - Location name
  - Total quantity
  - Line item count
  - FedEx requirement flag

### Pet Supermarket
- **Ship Date Logic**: 12 business days before due date (cell C7)
- **Rush Orders**: 8 business days before due date (if rush flag)
- **Weekend Handling**: Adjust ship date to avoid weekends
- **Additional Info**: Due date, days until due, order type

## Calendar Event Details

Each company can provide rich calendar event information:

### Chewy Events Include:
- Company name with location (e.g., "Chewy Phoenix - PO 12345")
- Ship date recommendation
- Priority level with color coding
- Total quantity
- Shipping method (FedEx for large orders)
- Location name

### Pet Supermarket Events Include:
- Company name and PO number
- Ship date recommendation
- Due date
- Days until due
- Order type information

## Integration with Main Application

The calendar modules are integrated through `calendar_helpers.py`, which provides:

- `get_ship_date_recommendation(company, file_path, **kwargs)` - Main entry point
- `get_order_details(company, file_path)` - Get company-specific details
- `get_calendar_event_details(company, file_path, po_number)` - Get complete event info

## Testing

To test a new company module:

1. Create test Excel files with the company's format
2. Test the ship date calculation function directly
3. Test integration through the main application
4. Verify calendar event details are correct

## Future Enhancements

- Add holiday calendar support
- Implement shipping method detection
- Add order priority algorithms
- Support for multiple file formats
- Integration with external shipping APIs
- Location-based shipping rules
- Customer-specific business rules 