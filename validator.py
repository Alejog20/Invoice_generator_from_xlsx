import pandas as pd
import os
from datetime import datetime

class InvoiceValidator:
    def __init__(self, reports_folder='Reports'):
        """
        Initialize the invoice validator.

        Args:
            reports_folder (str): Path to the folder where validation reports will be saved
        """
        self.reports_folder = reports_folder
        if not os.path.exists(reports_folder):
            os.makedirs(reports_folder)

        # Store validation results
        self.validation_results = {
            'original_totals': {},
            'new_totals': {},
            'discrepancies': [],
            'validation_details': []
        }
        
        # Add property to store invoice mappings
        self.invoice_mappings = {}

    def validate_original_data(self, df):
        """
        Validate the original data and store the validation results.
        
        Args:
            df (pandas.DataFrame): The original invoice data
            
        Returns:
            dict: Summary of the original data totals
        """
        if df.empty:
            print("Warning: No data to validate.")
            self.validation_results['original_totals'] = {
                'total_amount': 0.0,
                'total_tax': 0.0,
                'total_invoices': 0,
                'record_count': 0
            }
            return self.validation_results['original_totals']

        # Calculate total tax from all tax fields
        total_tax = 0.0
        tax_fields = ['State-Tax', 'County-Tax', 'City-Tax']
        
        # Add HST/GST and PST/QST taxes if they exist
        additional_tax_fields = ['HST/GST Tax', 'PST/QST Tax']
        for field in additional_tax_fields:
            if field in df.columns:
                tax_fields.append(field)
        
        # Sum all tax fields
        for field in tax_fields:
            if field in df.columns:
                total_tax += df[field].sum()

        self.validation_results['original_totals'] = {
            'total_amount': df['Amount'].sum() if 'Amount' in df.columns else 0.0,
            'total_tax': total_tax,
            'total_invoices': len(df['Invoice'].unique()) if 'Invoice' in df.columns else 0,
            'record_count': len(df)
        }

        # Validate data quality
        validation_details = {
            'null_amounts': df['Amount'].isnull().sum() if 'Amount' in df.columns else 0,
            'negative_amounts': (df['Amount'] < 0).sum() if 'Amount' in df.columns else 0,
            'zero_amounts': (df['Amount'] == 0).sum() if 'Amount' in df.columns else 0,
            'null_invoice_numbers': df['Invoice'].isnull().sum() if 'Invoice' in df.columns else 0,
            'duplicate_invoices': df['Invoice'].duplicated().sum() if 'Invoice' in df.columns else 0
        }

        # Add tax field validation if they exist
        null_taxes = {}
        for field in tax_fields:
            if field in df.columns:
                null_taxes[field] = df[field].isnull().sum()
        
        validation_details['null_taxes'] = null_taxes

        self.validation_results['validation_details'].append(validation_details)
        return self.validation_results['original_totals']

    def validate_generated_invoices(self, verification_data):
        """
        Validate the generated invoices against the original data.
        
        Args:
            verification_data (list): List of dictionaries containing verification data for each invoice
        """
        # Handle empty verification data
        if not verification_data:
            print("Warning: No verification data to validate.")
            self.validation_results['new_totals'] = {
                'total_amount': 0.0,
                'total_tax': 0.0,
                'total_invoices': 0,
                'record_count': 0
            }
            return
            
        # Store the mapping data for later reporting
        self.invoice_mappings = {entry['new_invoice_number']: entry['original_invoice_numbers'] 
                                for entry in verification_data}

        self.validation_results['new_totals'] = {
            'total_amount': sum(entry['total_amount'] for entry in verification_data),
            'total_tax': sum(entry['total_tax'] for entry in verification_data),
            'total_invoices': len(verification_data),
            'record_count': sum(entry['number_of_charges'] for entry in verification_data)
        }

        # Compare totals and identify discrepancies
        self._compare_totals()

    def _compare_totals(self, tolerance=0.01):
        """
        Compare original and new totals to identify discrepancies.

        Args:
            tolerance (float): Acceptable difference in monetary values
        """
        original = self.validation_results['original_totals']
        new = self.validation_results['new_totals']

        # Check monetary values
        for field in ['total_amount', 'total_tax']:
            if abs(original.get(field, 0) - new.get(field, 0)) > tolerance:
                self.validation_results['discrepancies'].append({
                    'field': field,
                    'original_value': original.get(field, 0),
                    'new_value': new.get(field, 0),
                    'difference': original.get(field, 0) - new.get(field, 0)
                })

        # Check counts
        for field in ['record_count']:
            if original.get(field, 0) != new.get(field, 0):
                self.validation_results['discrepancies'].append({
                    'field': field,
                    'original_value': original.get(field, 0),
                    'new_value': new.get(field, 0),
                    'difference': original.get(field, 0) - new.get(field, 0)
                })

    def generate_invoice_mapping_report(self):
        """
        Generate an Excel report showing mappings between original and new invoice numbers.
        
        Returns:
            str: Path to the generated mapping report
        """
        if not self.invoice_mappings:
            print("No invoice mappings found. Run validate_generated_invoices first.")
            return None
            
        # Create a flattened list for the report
        mapping_data = []
        for new_invoice, original_invoices in self.invoice_mappings.items():
            for orig_invoice in original_invoices:
                mapping_data.append({
                    'Invoice Number': new_invoice,
                    'Original Invoice': orig_invoice
                })
                
        # Convert to DataFrame
        mapping_df = pd.DataFrame(mapping_data)
        
        # Generate timestamp for the report
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(self.reports_folder, f'invoice_mapping_report_{timestamp}.xlsx')
        
        # Write to Excel
        mapping_df.to_excel(output_path, index=False, sheet_name='Invoice Mappings')
        
        return output_path

    def generate_validation_report(self, output_path=None):
        """
        Generate an Excel report with validation results.
        
        Args:
            output_path (str, optional): Path to save the report. If None, a default path is used.
            
        Returns:
            str: Path to the generated report
        """
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        if output_path is None:
            output_path = os.path.join(self.reports_folder, f'validation_report_{timestamp}.xlsx')

        try:
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                workbook = writer.book

                # Summary sheet
                self._create_summary_sheet(writer, workbook)

                # Discrepancies sheet
                self._create_discrepancies_sheet(writer, workbook)

                # Data quality sheet
                self._create_data_quality_sheet(writer, workbook)
        except Exception as e:
            print(f"Error generating validation report: {e}")
            import traceback
            print(f"Error details:\n{traceback.format_exc()}")
            return None

        return output_path

    def _create_summary_sheet(self, writer, workbook):
        """Create the summary sheet in the Excel report."""
        summary_data = {
            'Metric': [
                'Total Original Amount',
                'Total New Amount',
                'Total Original Tax',
                'Total New Tax',
                'Original Record Count',
                'New Record Count',
                'Original Invoice Count',
                'New Invoice Count'
            ],
            'Value': [
                self.validation_results['original_totals'].get('total_amount', 0),
                self.validation_results['new_totals'].get('total_amount', 0),
                self.validation_results['original_totals'].get('total_tax', 0),
                self.validation_results['new_totals'].get('total_tax', 0),
                self.validation_results['original_totals'].get('record_count', 0),
                self.validation_results['new_totals'].get('record_count', 0),
                self.validation_results['original_totals'].get('total_invoices', 0),
                self.validation_results['new_totals'].get('total_invoices', 0)
            ]
        }

        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

        # Format the summary sheet
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#D3D3D3',
            'border': 1
        })

        worksheet = writer.sheets['Summary']
        for col_num, value in enumerate(summary_df.columns.values):
            worksheet.write(0, col_num, value, header_format)

    def _create_discrepancies_sheet(self, writer, workbook):
        """Create the discrepancies sheet in the Excel report."""
        if self.validation_results['discrepancies']:
            discrepancies_df = pd.DataFrame(self.validation_results['discrepancies'])
            discrepancies_df.to_excel(writer, sheet_name='Discrepancies', index=False)

            worksheet = writer.sheets['Discrepancies']
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#FFB6C1',  # Light red for discrepancies
                'border': 1
            })

            for col_num, value in enumerate(discrepancies_df.columns.values):
                worksheet.write(0, col_num, value, header_format)
        else:
            # Create an empty discrepancies sheet with headers
            empty_df = pd.DataFrame(columns=['field', 'original_value', 'new_value', 'difference'])
            empty_df.to_excel(writer, sheet_name='Discrepancies', index=False)
            
            worksheet = writer.sheets['Discrepancies']
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#FFB6C1',
                'border': 1
            })
            
            for col_num, value in enumerate(empty_df.columns.values):
                worksheet.write(0, col_num, value, header_format)

    def _create_data_quality_sheet(self, writer, workbook):
        """Create the data quality sheet in the Excel report."""
        if self.validation_results['validation_details']:
            quality_df = pd.DataFrame(self.validation_results['validation_details'])
            
            # Convert nested dictionary to columns
            if 'null_taxes' in quality_df.columns:
                # Extract null_taxes dictionary into separate columns
                for idx, row in quality_df.iterrows():
                    null_taxes = row['null_taxes']
                    if isinstance(null_taxes, dict):
                        for tax_field, count in null_taxes.items():
                            quality_df.at[idx, f'null_{tax_field}'] = count
                
                # Drop the original null_taxes column
                quality_df = quality_df.drop('null_taxes', axis=1)
            
            quality_df.to_excel(writer, sheet_name='Data Quality', index=False)

            worksheet = writer.sheets['Data Quality']
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#98FB98',  # Light green for data quality
                'border': 1
            })

            for col_num, value in enumerate(quality_df.columns.values):
                worksheet.write(0, col_num, value, header_format)
        else:
            # Create an empty data quality sheet with a message
            empty_df = pd.DataFrame(['No validation details available'])
            empty_df.to_excel(writer, sheet_name='Data Quality', index=False, header=False)

    def print_validation_summary(self):
        """Print a summary of validation results to console."""
        print("\n=== Validation Summary ===")

        print("\nOriginal Totals:")
        for key, value in self.validation_results['original_totals'].items():
            print(f"{key}: {value}")

        print("\nNew Totals:")
        for key, value in self.validation_results['new_totals'].items():
            print(f"{key}: {value}")

        if self.validation_results['discrepancies']:
            print("\nDiscrepancies Found:")
            for disc in self.validation_results['discrepancies']:
                print(f"- {disc['field']}: Original={disc['original_value']}, "
                      f"New={disc['new_value']}, Diff={disc['difference']}")
        else:
            print("\nNo discrepancies found!")