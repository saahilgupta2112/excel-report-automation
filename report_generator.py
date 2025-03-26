import pandas as pd
from openpyxl import load_workbook

def generate_report(input_file, output_file):
    # Load Excel data
    df = pd.read_excel(input_file)

    # Clean data (example: drop empty rows and fill NaNs)
    df.dropna(how='all', inplace=True)
    df.fillna('', inplace=True)

    # Add a summary row
    summary = df.describe(include='all')

    # Save cleaned data to new file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Cleaned Data', index=False)
        summary.to_excel(writer, sheet_name='Summary')

    print(f"Report saved to {output_file}")

# Example usage
# generate_report('input.xlsx', 'output.xlsx')
