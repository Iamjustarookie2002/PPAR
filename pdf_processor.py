import tempfile
import os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

def process_pdf_report(processed_df, original_filename):
    """
    Create a PDF report from the processed Excel data.
    Input: processed_df (DataFrame) - the processed Excel data
    Output: PDF data as bytes
    """
    # Create a temporary file for the PDF
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_file.close()
    
    # Create PDF document
    doc = SimpleDocTemplate(temp_file.name, pagesize=A4)
    story = []
    
    # Get styles
    styles = getSampleStyleSheet()
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
    story.append(Paragraph(f"<b>Total Rows:</b> {len(processed_df)}", info_style))
    story.append(Paragraph(f"<b>Total Columns:</b> {len(processed_df.columns)}", info_style))
    story.append(Spacer(1, 20))
    
    # Data summary table
    story.append(Paragraph("<b>Data Summary</b>", styles['Heading2']))
    story.append(Spacer(1, 12))
    
    # Create summary table
    summary_data = [['Metric', 'Value']]
    summary_data.append(['Total Rows', str(len(processed_df))])
    summary_data.append(['Total Columns', str(len(processed_df.columns))])
    
    # Add numeric column statistics
    numeric_columns = processed_df.select_dtypes(include=['number']).columns
    if len(numeric_columns) > 0:
        for col in numeric_columns[:3]:  # Show first 3 numeric columns
            summary_data.append([f'{col} - Mean', f"{processed_df[col].mean():.2f}"])
            summary_data.append([f'{col} - Max', f"{processed_df[col].max():.2f}"])
            summary_data.append([f'{col} - Min', f"{processed_df[col].min():.2f}"])
    
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
    for col in processed_df.columns:
        col_data.append([
            col,
            str(processed_df[col].dtype),
            str(processed_df[col].count()),
            str(processed_df[col].isnull().sum())
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
