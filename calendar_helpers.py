import datetime
import calendar
from openpyxl import load_workbook
import xlrd

# Import company-specific modules
from calendar_modules.chewy_calendar import get_chewy_ship_date_advanced, get_chewy_order_priority, get_chewy_order_details, get_chewy_event_details
from calendar_modules.pet_supermarket_calendar import get_pet_supermarket_ship_date_advanced, get_pet_supermarket_order_details

def calculate_business_days(start_date, days):
    """Calculate a date that is a certain number of business days before the start date."""
    current_date = start_date
    business_days_counted = 0
    
    while business_days_counted < days:
        current_date -= datetime.timedelta(days=1)
        # Check if it's a weekday (Monday = 0, Sunday = 6)
        if current_date.weekday() < 5:  # Monday to Friday
            business_days_counted += 1
    
    return current_date

def parse_date_string(date_string):
    """Parse various date string formats and return a datetime object."""
    if not date_string:
        return None
    
    # Remove any extra whitespace
    date_string = str(date_string).strip()
    
    # Try different date formats
    date_formats = [
        '%m/%d/%Y',
        '%m/%d/%y',
        '%m-%d-%Y',
        '%m-%d-%y',
        '%Y-%m-%d',
        '%m/%d/%Y %H:%M:%S',
        '%m/%d/%y %H:%M:%S'
    ]
    
    for fmt in date_formats:
        try:
            return datetime.datetime.strptime(date_string, fmt)
        except ValueError:
            continue
    
    # If none of the formats work, try to extract date from Excel serial number
    try:
        # Excel serial number (days since 1900-01-01)
        excel_date = float(date_string)
        # Excel epoch is 1900-01-01, but Excel incorrectly treats 1900 as a leap year
        # So we need to adjust for dates after 1900-02-28
        if excel_date > 59:  # After 1900-02-28
            excel_date -= 1
        return datetime.datetime(1900, 1, 1) + datetime.timedelta(days=excel_date - 1)
    except (ValueError, TypeError):
        pass
    
    return None

def get_chewy_ship_date(file_path):
    """Calculate ship date for Chewy files based on quantities and FedEx shipping."""
    return get_chewy_ship_date_advanced(file_path)

def get_pet_supermarket_ship_date(file_path):
    """Calculate ship date for Pet Supermarket files based on due date in cell C7."""
    return get_pet_supermarket_ship_date_advanced(file_path)

def get_ship_date_recommendation(company, file_path, **kwargs):
    """
    Get ship date recommendation based on company and file content.
    
    Args:
        company: Company name
        file_path: Path to the Excel file
        **kwargs: Additional parameters for company-specific logic
    
    Returns:
        datetime.datetime: Recommended ship date or None
    """
    if company == "Chewy":
        return get_chewy_ship_date_advanced(file_path, **kwargs)
    elif company == "Pet Supermarket":
        return get_pet_supermarket_ship_date_advanced(file_path, **kwargs)
    # Add more companies here as needed
    else:
        return None

def get_order_details(company, file_path):
    """
    Get additional order details based on company and file content.
    
    Args:
        company: Company name
        file_path: Path to the Excel file
    
    Returns:
        dict: Order details specific to the company
    """
    if company == "Chewy":
        return get_chewy_order_details(file_path)
    elif company == "Pet Supermarket":
        return get_pet_supermarket_order_details(file_path)
    else:
        return {}

def format_ship_date_for_calendar(ship_date):
    """Format ship date for calendar display."""
    if ship_date:
        return ship_date.strftime('%Y-%m-%d')
    return None

def get_calendar_event_details(company, file_path, po_number):
    """
    Get comprehensive calendar event details for a company.
    
    Args:
        company: Company name
        file_path: Path to the Excel file
        po_number: Purchase order number
    
    Returns:
        dict: Calendar event details
    """
    # Initialize event_details first
    ship_date = get_ship_date_recommendation(company, file_path)
    order_details = get_order_details(company, file_path)
    event_details = {
        'company': company,
        'po_number': po_number,
        'ship_date': format_ship_date_for_calendar(ship_date),
        'title': f"{company} - PO {po_number}",
        'description': f"Ship date: {format_ship_date_for_calendar(ship_date) if ship_date else 'Manual review needed'}"
    }
    
    if company == "Chewy":
        chewy = get_chewy_event_details(file_path, po_number)
        event_details.update({
            'location': chewy['location'],
            'ship_date': chewy['ship_date'].strftime('%Y-%m-%d') if chewy['ship_date'] else None,
            'title': f"Chewy {chewy['location']} - PO {po_number}",
            'description': f"Ship Date: {chewy['ship_date'].strftime('%Y-%m-%d') if chewy['ship_date'] else ''}"
        })
        
        # Add company-specific details
        priority = order_details.get('priority', 'low')
        location = order_details.get('location', 'Unknown Location')
        total_qty = order_details.get('total_quantity', 0)
        
        event_details['title'] = f"Chewy {location} - PO {po_number}"
        event_details['description'] = f"Location: {location} | Ship date: {format_ship_date_for_calendar(ship_date) if ship_date else 'Manual review needed'} | Priority: {priority} | Qty: {total_qty}"
        
        # Add all order details to event_details for template access
        event_details.update(order_details)
        
        if order_details.get('priority') == 'high':
            event_details['color'] = 'red'
        elif order_details.get('priority') == 'medium':
            event_details['color'] = 'orange'
        else:
            event_details['color'] = 'green'
            
    elif company == "Pet Supermarket":
        details = get_order_details(company, file_path)
        if details.get('due_date'):
            event_details['description'] += f" | Due: {details['due_date'].strftime('%m/%d/%Y')}"
        if details.get('days_until_due'):
            event_details['description'] += f" | Days until due: {details['days_until_due']}"
        
        # Add all order details to event_details for template access
        event_details.update(details)
        
        event_details['color'] = 'blue'
    
    return event_details 