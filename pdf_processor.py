import tempfile
import os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

def process_pdf_report(excel_file_path, original_filename):
    """
    Create a PDF report from the Excel file.
    Input: excel_file_path (str) - path to the Excel file
    Output: PDF data as bytes
    """
    # Get styles - define this at the beginning so it's available everywhere
    styles = getSampleStyleSheet()
    
    try:
        # Read the Excel file to get the DataFrame
        import pandas as pd
        df = pd.read_excel(excel_file_path, sheet_name="FILES_DAT")
        
        # Create a temporary file for the PDF
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_file.close()
        
        # Create PDF document
        doc = SimpleDocTemplate(temp_file.name, pagesize=A4)
        story = []
        
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=18,
            spaceAfter=30,
            alignment=1  # Center alignment
        )
        
        # Title
        title = Paragraph("Data Processing Report", title_style)
        story.append(title)
        story.append(Spacer(1, 20))
        
        # Processing information
        info_style = ParagraphStyle(
            'InfoStyle',
            parent=styles['Normal'],
            fontSize=12,
            spaceAfter=12
        )
        
        story.append(Paragraph(f"<b>Original File:</b> {original_filename}", info_style))
        story.append(Paragraph(f"<b>Processing Date:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", info_style))
        story.append(Paragraph(f"<b>Total Rows:</b> {len(df)}", info_style))
        story.append(Paragraph(f"<b>Total Columns:</b> {len(df.columns)}", info_style))
        story.append(Spacer(1, 20))
        
        # Data summary table
        story.append(Paragraph("<b>Data Summary</b>", styles['Heading2']))
        story.append(Spacer(1, 12))
        
        # Create summary table
        summary_data = [['Metric', 'Value']]
        summary_data.append(['Total Rows', str(len(df))])
        summary_data.append(['Total Columns', str(len(df.columns))])
        
        # Add numeric column statistics
        numeric_columns = df.select_dtypes(include=['number']).columns
        if len(numeric_columns) > 0:
            for col in numeric_columns[:3]:  # Show first 3 numeric columns
                summary_data.append([f'{col} - Mean', f"{df[col].mean():.2f}"])
                summary_data.append([f'{col} - Max', f"{df[col].max():.2f}"])
                summary_data.append([f'{col} - Min', f"{df[col].min():.2f}"])
        
        # Create table
        summary_table = Table(summary_data)
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(summary_table)
        story.append(Spacer(1, 20))
        
        # Column information
        story.append(Paragraph("<b>Column Information</b>", styles['Heading2']))
        story.append(Spacer(1, 12))
        
        # Create column info table
        col_data = [['Column Name', 'Data Type', 'Non-Null Count', 'Null Count']]
        for col in df.columns:
            col_data.append([
                col,
                str(df[col].dtype),
                str(df[col].count()),
                str(df[col].isnull().sum())
            ])
        
        col_table = Table(col_data)
        col_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        story.append(col_table)
        
        # Build PDF
        doc.build(story)
        
        # Read the PDF file
        with open(temp_file.name, 'rb') as f:
            pdf_data = f.read()
        
        # Clean up temporary file
        os.unlink(temp_file.name)
        
        return pdf_data
        
    except Exception as e:
        # Return a simple error PDF if something goes wrong
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_file.close()
        
        doc = SimpleDocTemplate(temp_file.name, pagesize=A4)
        story = []
        
        error_style = ParagraphStyle(
            'Error',
            parent=styles['Normal'],
            fontSize=14,
            textColor=colors.red
        )
        
        story.append(Paragraph("Error Generating PDF", styles['Heading1']))
        story.append(Spacer(1, 20))
        story.append(Paragraph(f"An error occurred while converting Excel to PDF: {str(e)}", error_style))
        
        doc.build(story)
        
        with open(temp_file.name, 'rb') as f:
            pdf_data = f.read()
        
        os.unlink(temp_file.name)
        return pdf_data
