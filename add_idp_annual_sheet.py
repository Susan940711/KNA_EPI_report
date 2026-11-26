"""Add IDP_annual sheet to KNA_EPI_long.xlsx by aggregating quarterly data."""

import pandas as pd
from openpyxl import load_workbook
from pathlib import Path

def create_idp_annual_sheet(input_path: str, output_path: str = None):
    """
    Create IDP_annual sheet by summing quarterly IDP data.
    
    Args:
        input_path: Path to input Excel file (KNA_EPI_long.xlsx)
        output_path: Path to output Excel file (defaults to input_path)
    """
    if output_path is None:
        output_path = input_path
    
    # Read the IDP sheet
    df_idp = pd.read_excel(input_path, sheet_name='IDP')
    
    # Create a copy for annual data
    df_annual = df_idp[['Period', 'Organization', 'Project Name', 'indicator']].copy()
    
    # Sum quarterly columns for each category
    # IDP Male: Q1 + Q2 + Q3 + Q4
    df_annual['IDP Male'] = (
        df_idp['Q1 IDP Male'] + 
        df_idp['Q2 IDP Male'] + 
        df_idp['Q3 IDP Male'] + 
        df_idp['Q4 IDP Male']
    )
    
    # IDP Female: Q1 + Q2 + Q3 + Q4
    df_annual['IDP Female'] = (
        df_idp['Q1 IDP Female'] + 
        df_idp['Q2 IDP Female'] + 
        df_idp['Q3 IDP Female'] + 
        df_idp['Q4 IDP Female']
    )
    
    # non-IDP Male: Q1 + Q2 + Q3 + Q4
    df_annual['non-IDP Male'] = (
        df_idp['Q1 non-IDP Male'] + 
        df_idp['Q2 non-IDP Male'] + 
        df_idp['Q3 non-IDP Male'] + 
        df_idp['Q4 non-IDP Male']
    )
    
    # non-IDP Female: Q1 + Q2 + Q3 + Q4
    df_annual['non-IDP Female'] = (
        df_idp['Q1 non-IDP Female'] + 
        df_idp['Q2 non-IDP Female'] + 
        df_idp['Q3 non-IDP Female'] + 
        df_idp['Q4 non-IDP Female']
    )
    
    # Load workbook and add new sheet
    wb = load_workbook(input_path)
    
    # Remove sheet if it already exists
    if 'IDP_annual' in wb.sheetnames:
        del wb['IDP_annual']
    
    # Add new sheet
    ws = wb.create_sheet('IDP_annual')
    
    # Write column headers
    headers = ['Period', 'Organization', 'Project Name', 'indicator', 
               'IDP Male', 'IDP Female', 'non-IDP Male', 'non-IDP Female']
    for col_idx, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_idx, value=header)
    
    # Write data
    for row_idx, row_data in enumerate(df_annual.values, 2):
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            # Format as integer if it's numeric
            if isinstance(value, (int, float)) and not pd.isna(value):
                if value == int(value):
                    cell.value = int(value)
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 25
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 18
    ws.column_dimensions['H'].width = 18
    
    # Save workbook
    wb.save(output_path)
    print(f"Successfully created IDP_annual sheet in {output_path}")
    print(f"  - Columns: {', '.join(headers)}")
    print(f"  - Rows: {len(df_annual)}")


if __name__ == '__main__':
    input_file = Path(__file__).parent / 'KNA_EPI_long.tmp.xlsx'
    
    if not input_file.exists():
        print(f"Error: {input_file} not found")
        exit(1)
    
    create_idp_annual_sheet(str(input_file))
