from generator import InvoiceGenerator
from validator import InvoiceValidator
import os
import pandas as pd

def collect_verification_data(generator):
    """
    Collect verification data from generated PDFs and original data.
    For grouped invoices by invoice number.
    
    Args:
        generator: InvoiceGenerator instance that contains the generated mappings
        
    Returns:
        list: Verification data for each invoice
    """
    verification_data = []
    
    # Check if mappings exist
    if not hasattr(generator, 'generated_mappings') or not generator.generated_mappings:
        print("Warning: No generated mappings found in the generator instance")
        return verification_data
    
    # For each invoice number
    for invoice_id, _ in generator.generated_mappings.items():
        # Get the data for this invoice
        invoice_data = generator.df[generator.df['Invoice'] == invoice_id]
        
        if invoice_data.empty:
            print(f"Warning: No data found for invoice {invoice_id}")
            continue
        
        # Calculate totals for this invoice
        amount = invoice_data['Amount'].sum()
        
        # Calculate total tax from all tax fields
        tax = 0.0
        tax_fields = ['State-Tax', 'County-Tax', 'City-Tax']
        
        # Add HST/GST and PST/QST taxes if they exist
        additional_tax_fields = ['HST/GST Tax', 'PST/QST Tax']
        for field in additional_tax_fields:
            if field in invoice_data.columns:
                tax_fields.append(field)
        
        # Sum all tax fields
        for field in tax_fields:
            if field in invoice_data.columns:
                tax += invoice_data[field].sum()
        
        # Get facility/division info from first row
        first_row = invoice_data.iloc[0]
        
        # Create verification entry
        verification_entry = {
            'new_invoice_number': invoice_id,
            'original_invoice_numbers': [invoice_id],
            'facility_name': first_row.get('Division', ''),
            'total_amount': float(amount),
            'total_tax': float(tax),
            'total_invoice': float(amount + tax),
            'number_of_charges': len(invoice_data)
        }
        
        verification_data.append(verification_entry)
        
        # Print validation information
        print(f"\nProcessed Invoice: {invoice_id}")
        print(f"Amount: ${amount:,.2f}")
        print(f"Tax: ${tax:,.2f}")
    
    # Final validation
    if verification_data:
        total_processed_amount = sum(entry['total_amount'] for entry in verification_data)
        total_processed_tax = sum(entry['total_tax'] for entry in verification_data)
        
        # Calculate original totals
        original_total_amount = generator.df['Amount'].sum()
        
        # Calculate total tax from all tax fields
        original_total_tax = 0.0
        tax_fields = ['State-Tax', 'County-Tax', 'City-Tax']
        
        # Add HST/GST and PST/QST taxes if they exist
        additional_tax_fields = ['HST/GST Tax', 'PST/QST Tax']
        for field in additional_tax_fields:
            if field in generator.df.columns:
                tax_fields.append(field)
        
        # Sum all tax fields
        for field in tax_fields:
            if field in generator.df.columns:
                original_total_tax += generator.df[field].sum()
        
        print("\nValidation Totals:")
        print(f"Original Total Amount: ${original_total_amount:,.2f}")
        print(f"Processed Total Amount: ${total_processed_amount:,.2f}")
        print(f"Original Total Tax: ${original_total_tax:,.2f}")
        print(f"Processed Total Tax: ${total_processed_tax:,.2f}")
        
        if abs(original_total_amount - total_processed_amount) > 0.01:
            print("\nWARNING: Total amounts don't match!")
        if abs(original_total_tax - total_processed_tax) > 0.01:
            print("\nWARNING: Total taxes don't match!")
    
    return verification_data

def analyze_discrepancies(generator, verification_data):
    """
    Analyze discrepancies between original and generated invoice data.
    For invoices grouped by invoice number.
    """
    print("\n=== Discrepancy Analysis Report ===\n")
    
    # Check if verification data is empty
    if not verification_data:
        print("No verification data available for analysis.")
        return None
    
    # Create sets of original and processed invoice numbers
    original_invoices = set(generator.df['Invoice'].unique())
    processed_invoices = set()
    for entry in verification_data:
        processed_invoices.add(entry['new_invoice_number'])
    
    # Find missing invoices
    missing_invoices = original_invoices - processed_invoices
    
    # Analyze missing data
    if missing_invoices:
        missing_df = generator.df[generator.df['Invoice'].isin(missing_invoices)]
        
        # Calculate missing amount
        missing_amount = missing_df['Amount'].sum()
        
        # Calculate missing tax
        missing_tax = 0.0
        tax_fields = ['State-Tax', 'County-Tax', 'City-Tax']
        
        # Add HST/GST and PST/QST taxes if they exist
        additional_tax_fields = ['HST/GST Tax', 'PST/QST Tax']
        for field in additional_tax_fields:
            if field in missing_df.columns:
                tax_fields.append(field)
        
        # Sum all tax fields
        for field in tax_fields:
            if field in missing_df.columns:
                missing_tax += missing_df[field].sum()
        
        print("Missing Data Summary:")
        print(f"Number of missing invoices: {len(missing_invoices)}")
        print(f"Missing amount total: ${missing_amount:,.2f}")
        print(f"Missing tax total: ${missing_tax:,.2f}")
        return missing_df
    else:
        print("No missing invoices found. All records processed successfully.")
        return None

def main():
    try:
        print("\n=== Starting Invoice Processing ===\n")
        
        # Define common paths
        current_dir = os.getcwd()
        reports_folder = os.path.join(current_dir, 'Reports')
        
        # Create reports folder if it doesn't exist
        if not os.path.exists(reports_folder):
            os.makedirs(reports_folder)
        
        # Initialize generator and validator
        print("Initializing Invoice Generator...")
        generator = InvoiceGenerator()
        
        print("Initializing Validator...")
        validator = InvoiceValidator(reports_folder)
        
        # Validate original data
        print("\nValidating original data...")
        original_validation = validator.validate_original_data(generator.df)
        print("Original data validation complete.")
        
        # Generate invoices
        print("\nGenerating invoices grouped by invoice number...")
        generator.generate_all_invoices()
        print("Invoice generation complete.")
        
        # Collect verification data after generation
        print("\nCollecting verification data...")
        verification_data = collect_verification_data(generator)
        print(f"Collected verification data for {len(verification_data)} invoices.")
        
        # Validate generated invoices
        print("\nValidating generated invoices...")
        validator.validate_generated_invoices(verification_data)
        print("Invoice validation complete.")
        
        # Analyze discrepancies
        print("\nAnalyzing discrepancies...")
        missing_data = analyze_discrepancies(generator, verification_data)
        
        if missing_data is not None and not missing_data.empty:
            # Save missing data to Excel for further analysis
            report_path = os.path.join(reports_folder, 'missing_data_analysis.xlsx')
            missing_data.to_excel(report_path, index=False)
            print(f"\nDetailed missing data report saved to: {report_path}")
        
        # Generate validation report
        print("\nGenerating validation report...")
        report_file = validator.generate_validation_report()
        
        # Print validation summary
        validator.print_validation_summary()

        print("\nGenerating invoice mapping report...")
        mapping_report = validator.generate_invoice_mapping_report()
        if mapping_report:
            print(f"Invoice mapping report available at: {mapping_report}")
        else:
            print("No invoice mapping report was generated.")
        
        print(f"\nValidation report available at: {report_file}")
        print("\n=== Process completed successfully! ===")
        
    except Exception as e:
        print(f"\nAn error occurred: {str(e)}")
        print("\nFull error details:")
        import traceback
        print(traceback.format_exc())
        raise

if __name__ == "__main__":
    main()