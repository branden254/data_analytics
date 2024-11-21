import openpyxl
from collections import defaultdict
from decimal import Decimal

def parse_kes(value):
    if isinstance(value, str):
        return Decimal(value.replace('KES ', '').replace(',', ''))
    elif isinstance(value, (int, float)):
        return Decimal(str(value))
    return Decimal('0')

def generate_report(data):
    report = []
    total_amount = Decimal('0')
    
    for person, uploads in data.items():
        upload_details = []
        person_total = Decimal('0')
        
        for source, info in uploads.items():
            upload_count = info['uploads']
            payment = info['payment']
            upload_details.append(f"{upload_count} on {source}")
            person_total += payment
        
        upload_str = " and ".join(upload_details)
        report.append(f"{person} has done {upload_str} and should be paid KES {person_total:.2f}.")
        total_amount += person_total
    
    report.append(f"Total amount to be paid: KES {total_amount:.2f}.")
    return "\n".join(report)

def process_excel(filename):
    data = defaultdict(lambda: defaultdict(lambda: {'uploads': 0, 'payment': Decimal('0')}))
    
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    
    headers = [cell.value for cell in sheet[1]]
    
    for row in sheet.iter_rows(min_row=2, values_only=True):
        row_data = dict(zip(headers, row))
        
        name = row_data['Name']
        if not name:  # Skip rows without a name
            continue
        
        source = row_data['Source'] or 'Unknown'
        uploads = row_data['No of uploads']
        uploads = int(uploads) if uploads is not None else 0
        payment = parse_kes(row_data['Accrued Pay'])
        
        data[name][source]['uploads'] += uploads
        data[name][source]['payment'] += payment
    
    return dict(data)

def main():
    filename = r"C:\Users\carso\Documents\october\website uploads.xlsx"
    processed_data = process_excel(filename)
    report = generate_report(processed_data)
    
    # Save report to a text file
    with open("payment_report.txt", "w") as f:
        f.write(report)
    
    print("Report generated and saved to payment_report.txt")
    print("\nReport preview:")
    print(report)

if __name__ == "__main__":
    main()