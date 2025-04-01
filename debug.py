import pandas as pd
import os
import sys
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.units import inch

def generate_simple_invoice():
    """
    Simplified function to generate a basic invoice PDF to test if PDF generation is working properly.
    """
    print("Starting simple invoice generation...")
    
    # Create output directory if it doesn't exist
    output_dir = "Output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Define PDF path
    pdf_path = os.path.join(output_dir, "test_invoice.pdf")
    print(f"PDF will be saved to: {pdf_path}")

    # Create a simple PDF document
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        rightMargin=0.75*inch,
        leftMargin=0.75*inch,
        topMargin=0.75*inch,
        bottomMargin=0.75*inch
    )

    # Create content for the PDF
    elements = []
    styles = getSampleStyleSheet()
    
    # Add title
    title_style = ParagraphStyle(
        'Title',
        parent=styles['Heading1'],
        fontSize=20,
        alignment=1  # Center alignment
    )
    elements.append(Paragraph("Test Invoice", title_style))
    elements.append(Spacer(1, 20))
    
    # Add invoice details
    data = [
        ["Description", "Amount"],
        ["Test Item", "$100.00"],
        ["Tax", "$10.00"],
        ["Total", "$110.00"]
    ]
    
    table = Table(data, colWidths=[4*inch, 2*inch])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    elements.append(table)
    
    # Build the PDF
    try:
        print("Building PDF...")
        doc.build(elements)
        print(f"Successfully generated test invoice at: {pdf_path}")
        return True
    except Exception as e:
        print(f"Error generating PDF: {e}")
        import traceback
        print(traceback.format_exc())
        return False

def check_reportlab_installation():
    """
    Check if ReportLab is installed and working properly.
    """
    print("\nChecking ReportLab installation...")
    try:
        import reportlab
        print(f"ReportLab version: {reportlab.Version}")
        return True
    except ImportError:
        print("ReportLab is not installed. Please install it with: pip install reportlab")
        return False

def check_file_permissions():
    """
    Check if we can write to the Output directory.
    """
    print("\nChecking file permissions...")
    output_dir = "Output"
    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            print(f"Created Output directory: {output_dir}")
        except Exception as e:
            print(f"Error creating Output directory: {e}")
            return False
    
    # Try to create a test file
    test_file = os.path.join(output_dir, "test_permissions.txt")
    try:
        with open(test_file, 'w') as f:
            f.write("Testing write permissions")
        print(f"Successfully wrote to test file: {test_file}")
        
        # Try to delete the test file
        os.remove(test_file)
        print("Successfully deleted test file")
        
        return True
    except Exception as e:
        print(f"Error checking file permissions: {e}")
        return False

def check_excel_file():
    """
    Check if we can read the Excel file.
    """
    print("\nChecking for Excel files...")
    input_dir = "Input"
    if not os.path.exists(input_dir):
        print(f"Input directory does not exist: {input_dir}")
        return False
    
    excel_files = [f for f in os.listdir(input_dir) if f.lower().endswith('.xlsx')]
    if not excel_files:
        print("No Excel files found in Input directory")
        return False
    
    excel_file = os.path.join(input_dir, excel_files[0])
    print(f"Found Excel file: {excel_file}")
    
    try:
        df = pd.read_excel(excel_file)
        print(f"Successfully read Excel file with {len(df)} rows")
        print(f"Columns: {', '.join(df.columns.tolist())}")
        return True
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return False

def main():
    print("=== Invoice Generator Diagnostic Tool ===")
    print(f"Python version: {sys.version}")
    print(f"Current directory: {os.getcwd()}")
    
    # Run checks
    reportlab_ok = check_reportlab_installation()
    permissions_ok = check_file_permissions()
    excel_ok = check_excel_file()
    
    # If all checks pass, try to generate a simple invoice
    if reportlab_ok and permissions_ok:
        print("\nAttempting to generate a simple test invoice...")
        invoice_ok = generate_simple_invoice()
    else:
        invoice_ok = False
        print("\nSkipping invoice generation due to failed checks")
    
    # Print summary
    print("\n=== Diagnostic Summary ===")
    print(f"ReportLab check: {'PASS' if reportlab_ok else 'FAIL'}")
    print(f"File permissions check: {'PASS' if permissions_ok else 'FAIL'}")
    print(f"Excel file check: {'PASS' if excel_ok else 'FAIL'}")
    print(f"Test invoice generation: {'PASS' if invoice_ok else 'FAIL'}")
    
    if not (reportlab_ok and permissions_ok and excel_ok and invoice_ok):
        print("\nDiagnostic failed. Please address the issues above.")
    else:
        print("\nAll diagnostics passed! The system should be able to generate invoices.")

if __name__ == "__main__":
    main()