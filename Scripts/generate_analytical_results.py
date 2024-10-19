from openpyxl import Workbook

def generate_analytical_results(data, country):
    # Create a new Excel workbook
    wb = Workbook()
    ws = wb.active

    # Save the new Excel file
    output_path = '../Output/Analytical Results.xlsx'
    wb.save(output_path)
    print(f"Analytical Results saved to {output_path}")