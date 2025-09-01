import pandas as pd
import re
import colorsys
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter

def process_excel_report(df, excel_filename):
    """
    Main function that creates the Excel file with both sheets.
    This is the ONLY function accessible to main in app.py.
    """
    wb = Workbook()
    
    # Create Sheet1 and process it
    ws1 = wb.active
    ws1.title = "Sheet1"
    process_sheet1_data(ws1)
    
    # Create Sheet2 and process it
    ws2 = wb.create_sheet("Sheet2")
    process_sheet2_data(df, ws2)
    
    # Save the workbook
    wb.save(excel_filename)

def process_sheet1_data(ws1):
    """
    Process Sheet1 - currently empty for future implementation.
    """
    # TODO: Implement Sheet1 processing logic
    # For now, Sheet1 remains empty
    pass

def process_sheet2_data(df, ws2):
    """
    Process and format Sheet2 with data processing, coloring, and additional columns.
    """
    processed_df = process_original_excel_data(df)

    # Write DataFrame to Sheet2
    for r in dataframe_to_rows(processed_df, index=False, header=True):
        ws2.append(r)
    
    # Add additional columns to the right
    add_additional_columns_to_sheet2(ws2)
    
    # Apply formatting (bold headers, center alignment, column widths)
    apply_formatting(ws2)
    
    # Apply coloring to Data Source and Weight bearing columns
    apply_coloring(ws2)

def process_original_excel_data(df):
    """Process and filter the original Excel data."""
    processed_df = df.copy()

    # Filter out rows with '.dat' in "File short name"
    processed_df = processed_df[~processed_df["File short name"].str.endswith(".dat", na=False)]

    # Select and rename required columns
    column_mapping = {
        "File comment": "Data Source",
        "Maximum force (normalized to BW) /Total object/ [%BW]": "Maximum force [%BW]",
        "Force-time integral (normalized to BW) /Total object/ [%BW*s]": "Force-time integral [%BW*s]",
        "Contact time/TO [ms]": "Contact time/TO [ms]",
    }
    
    processed_df = processed_df[list(column_mapping.keys())].rename(columns=column_mapping)

    # Convert LF1 -> LF_1, LH2 -> LH_2, etc.
    processed_df["Data Source"] = processed_df["Data Source"].apply(
        lambda x: re.sub(r'([A-Z]+)(\d+)', r'\1_\2', str(x))
    )

    return processed_df

def add_additional_columns_to_sheet2(ws2):
    """Add additional columns to the right of Sheet2."""
    current_max_col = ws2.max_column
    
    # Add empty column with arrow in center
    arrow_col_idx = current_max_col + 1
    num_data_rows = ws2.max_row - 1
    if num_data_rows > 0:
        arrow_row_idx = (num_data_rows // 2) + 2
        arrow_cell = ws2.cell(row=arrow_row_idx, column=arrow_col_idx, value="→")
        arrow_cell.alignment = Alignment(horizontal="center", vertical="center")
        arrow_cell.font = Font(size=14)

    # Add two more columns with headings
    new_columns = [
        (current_max_col + 2, "Weight bearing [%]"),
        (current_max_col + 3, "Asymmetery Index (L to R: SI)")
    ]
    
    for col_idx, heading in new_columns:
        header_cell = ws2.cell(row=1, column=col_idx, value=heading)
        header_cell.font = Font(bold=True)
        header_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        # Center-align all data cells in the new columns
        for row_idx in range(2, ws2.max_row + 1):
            ws2.cell(row=row_idx, column=col_idx).alignment = Alignment(horizontal="center", vertical="center")

def apply_formatting(ws2):
    """Apply formatting: bold headers, center alignment, and auto-adjust column widths."""
    # Apply bold and center alignment to headers
    for cell in ws2[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Center-align all data cells
    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Auto-adjust column widths
    for column in ws2.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        
        for cell in column:
            try:
                if cell.value:
                    cell_length = len(str(cell.value)) + 2  # Add padding
                    max_length = max(max_length, cell_length)
            except:
                pass
        
        # Set column width with limits
        adjusted_width = min(max(max_length, 10), 50)
        ws2.column_dimensions[column_letter].width = adjusted_width

def apply_coloring(ws2):
    """Apply coloring to Data Source and Weight bearing columns."""
    # Color cache for consistent coloring
    color_cache = {}
    
    # Apply coloring to Data Source (column 1) and Weight bearing columns
    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row):
        for cell in row:
            if cell.column in [1, 6]:  # Data Source and Weight bearing columns
                match = re.search(r'_(\d+)$', str(ws2.cell(row=cell.row, column=1).value))
                if match:
                    num = int(match.group(1))
                    if num not in color_cache:
                        color_cache[num] = {
                            'bright': get_color_for_number(num),
                            'dim': get_dim_color_for_number(num)
                        }
                    
                    # Apply bright color to Data Source, dim color to Weight bearing
                    color = color_cache[num]['bright'] if cell.column == 1 else color_cache[num]['dim']
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

def get_color_for_number(n):
    """Generate a unique pastel-like color for each number using HSL."""
    hue = (n * 137) % 360 / 360.0  # golden angle for good distribution
    lightness = 0.75
    saturation = 0.9
    r, g, b = colorsys.hls_to_rgb(hue, lightness, saturation)
    return f"{int(r*255):02X}{int(g*255):02X}{int(b*255):02X}"

def get_dim_color_for_number(n):
    """Generate a dimmer version of the color for each number using HSL."""
    hue = (n * 137) % 360 / 360.0  # golden angle for good distribution
    lightness = 0.75  # Much lighter (dimmer)
    saturation = 0.5  # Less saturated
    r, g, b = colorsys.hls_to_rgb(hue, lightness, saturation)
    return f"{int(r*255):02X}{int(g*255):02X}{int(b*255):02X}"
