import pandas as pd
import os
import re
import random
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, KeepTogether
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

class InvoiceGenerator:
    def __init__(self):
        current_dir = os.getcwd()
        self.input_folder = os.path.join(current_dir, 'Input')
        self.output_folder = os.path.join(current_dir, 'Output')
        self.used_invoice_numbers = set()
        self.margin = 0.75 * inch
        self.content_width = letter[0] - (2 * self.margin)
        self.generated_mappings = {}
        self.conversion_log = []  # Initialize conversion log
      
        # Create input and output folders if they don't exist
        if not os.path.exists(self.input_folder):
            os.makedirs(self.input_folder)
            print(f"Created input folder: {self.input_folder}")
            
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)
            print(f"Created output folder: {self.output_folder}")
        
        # Check if there are any files in the input folder
        input_files = [f for f in os.listdir(self.input_folder) if f.lower().endswith('.xlsx')]
        if not input_files:
            print(f"Warning: No Excel files found in {self.input_folder}")
            # Create a placeholder dataframe with the required columns
            self.df = self._create_empty_dataframe()
        else:
            # Read the first Excel file
            excel_path = self._get_first_file(self.input_folder, ['.xlsx'])
            self.df = self._read_input_file(excel_path)
        
        # Initialize styles first
        self.styles = getSampleStyleSheet()
        
        # Then define all style attributes
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=20,
            fontName='Helvetica-Bold',
            spaceAfter=10,
            alignment=0
        )

        self.header_style = ParagraphStyle(
            'HeaderStyle',
            parent=self.styles['Normal'],
            fontSize=10,
            leading=14,
            leftIndent=0,
            spaceBefore=6,
            spaceAfter=6,
            wordWrap='LTR'
        )
        
        self.detail_style = ParagraphStyle(
            'DetailStyle',
            parent=self.styles['Normal'],
            fontSize=8,
            leading=10,
            spaceBefore=2,
            spaceAfter=2,
            wordWrap='LTR'
        )
        
        self.summary_style = ParagraphStyle(
            'Summary',
            parent=self.styles['Normal'],
            fontSize=10,
            fontName='Helvetica-Bold',
            alignment=2,
            spaceBefore=6,
            spaceAfter=6
        )

    def _create_empty_dataframe(self):
        """Create an empty dataframe with the required columns"""
        columns = [
            'Invoice', 
            'Invoice Date',
            'Division',
            'Correct Division',
            'Correct Dept ID',
            'Correct Dept Name',
            'Bill Code',
            'Bill Code Description',
            'Description',
            'Unit-Qty',
            'Amount', 
            'State-Tax', 
            'County-Tax', 
            'City-Tax',
            'HST/GST Tax',
            'PST/QST Tax',
            'Addr1',
            'addr2',
            'addr3',
            'City',
            'State',
            'Zip'
        ]
        df = pd.DataFrame(columns=columns)
        
        # Set appropriate data types
        numeric_columns = ['Amount', 'State-Tax', 'County-Tax', 'City-Tax', 'HST/GST Tax', 'PST/QST Tax', 'Unit-Qty']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = df[col].astype(float)
            
        return df

    def _get_first_file(self, folder, extensions):
        for ext in extensions:
            files = [f for f in os.listdir(folder) if f.lower().endswith(ext.lower())]
            if files:
                return os.path.join(folder, files[0])
        raise Exception(f"No file with extensions {extensions} found in {folder}")

    def _convert_european_number(self, value):
        """
        Convert string number representations to float with detailed error tracking.
        """
        try:
            if pd.isna(value) or value == '':
                self.conversion_log.append({
                    'value': value,
                    'issue': 'empty_or_null',
                    'converted_to': 0.0
                })
                return 0.0
                
            cleaned_value = str(value).replace(',', '.').replace(' ', '')
            converted = float(cleaned_value)
            
            # Track any potential precision loss
            if abs(float(cleaned_value) - converted) > 1e-10:
                self.conversion_log.append({
                    'value': value,
                    'issue': 'precision_loss',
                    'original': cleaned_value,
                    'converted': converted
                })
                
            return converted
            
        except (ValueError, TypeError) as e:
            self.conversion_log.append({
                'value': value,
                'issue': 'conversion_error',
                'error': str(e),
                'converted_to': 0.0
            })
            return 0.0

    def format_date(self, date_str):
        """
        Format the date to YYYY-MM-DD format
        """
        if pd.isna(date_str) or date_str == '':
            return ''
            
        try:
            # Try to parse the date
            date = pd.to_datetime(date_str)
            return date.strftime('%Y-%m-%d')
        except:
            # If parsing fails, return the original string
            return str(date_str)

    def get_invoice_mapping(self):    
        # For grouped invoices by invoice number
        mapping = {}
        if self.df.empty:
            return mapping
            
        # Group by Invoice number
        invoice_groups = self.df.groupby('Invoice')
        for invoice_num, group in invoice_groups:
            mapping[str(invoice_num)] = [str(invoice_num)]
        
        return mapping

    def _read_input_file(self, path):
        print(f"Reading input file: {path}")
        
        try:
            # Fixed: Removed 'sep' parameter as it's not valid for read_excel
            df = pd.read_excel(path, dtype=str)
            print("Successfully read Excel file")
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            try:
                # Try with different engine
                df = pd.read_excel(path, engine='openpyxl', dtype=str)
                print("Successfully read file using openpyxl engine")
            except Exception as e2:
                print(f"Second attempt failed: {e2}")
                # Return empty dataframe with required columns
                return self._create_empty_dataframe()

        # Add this new code block after reading the Excel
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].apply(lambda x: ''.join(char for char in str(x) if char.isprintable()))
        
        print("Available columns in Excel file:", df.columns.tolist())
    
        # Handle missing columns
        needed_columns = [
            'Invoice', 
            'Invoice Date',
            'Division',
            'Correct Division',
            'Correct Dept ID',
            'Correct Dept Name',
            'Bill Code',
            'Bill Code Description',
            'Description',
            'Unit-Qty',
            'Amount', 
            'State-Tax', 
            'County-Tax', 
            'City-Tax',
            'HST/GST Tax',
            'PST/QST Tax',
            'Addr1',
            'addr2',
            'addr3',
            'City',
            'State',
            'Zip'
        ]
        
        print("\nChecking for missing columns...")
        missing_columns = []
        for col in needed_columns:
            if col not in df.columns:
                print(f"Missing column: '{col}'")
                missing_columns.append(col)
                
        # Add missing columns with default values
        for col in missing_columns:
            if col in ['Amount', 'State-Tax', 'County-Tax', 'City-Tax', 'HST/GST Tax', 'PST/QST Tax', 'Unit-Qty']:
                df[col] = 0.0
            else:
                df[col] = ''
        
        # Convert numeric columns
        numeric_columns = ['Amount', 'State-Tax', 'County-Tax', 'City-Tax', 'HST/GST Tax', 'PST/QST Tax', 'Unit-Qty']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = df[col].apply(self._convert_european_number)
        
        # Fill missing values
        df = df.fillna({
            'Bill Code': '',
            'Bill Code Description': '',
            'Description': '',
            'Division': '',
            'Correct Division': '',
            'Correct Dept ID': '',
            'Correct Dept Name': '',
            'Invoice': '',
            'Invoice Date': '',
            'Amount': 0.0,
            'State-Tax': 0.0,
            'County-Tax': 0.0,
            'City-Tax': 0.0,
            'HST/GST Tax': 0.0,
            'PST/QST Tax': 0.0,
            'Unit-Qty': 0.0,
            'Addr1': '',
            'addr2': '',
            'addr3': '',
            'City': '',
            'State': '',
            'Zip': ''
        })
        
        print(f"\nSuccessfully read {len(df)} records from input file")
        return df

    def format_currency(self, value):
        try:
            return f"${float(value):,.2f}"
        except (ValueError, TypeError):
            return "$0.00"

    def sanitize_text(self, text):
       
        if text is None or str(text).strip() == '':
            return " "  # Return a space instead of empty string
        
        # Convert to string and clean the text
        try:
            # First, encode to ASCII, then decode back to remove non-ASCII chars
            text = str(text).encode('ascii', 'ignore').decode('ascii')
            # Remove any control characters
            text = ''.join(char for char in text if ord(char) >= 32)
            # Normalize whitespace
            text = ' '.join(text.split())
            return text if text else " "  # Return a space if text is empty
        except Exception as e:
            print(f"Warning: Error sanitizing text: {e}")
            return " " 
    
    def prepare_invoice_groups(self):
        """
        Group invoices by Invoice number as per requirements.
        """
        if self.df.empty:
            print("Warning: No data to process.")
            return []
            
        # Group dataframe by Invoice number
        invoice_groups = []
        for invoice_num, group_df in self.df.groupby('Invoice'):
            if pd.isna(invoice_num) or str(invoice_num).strip() == '':
                print(f"Warning: Skipping group with empty invoice number")
                continue
                
            invoice_groups.append({
                'invoice_num': str(invoice_num),
                'data': group_df
            })
            
        return invoice_groups

    def generate_invoice(self, invoice_group):
        """Generate an invoice PDF for a grouped set of charges with the same invoice number"""
        invoice_num = invoice_group['invoice_num']
        group_df = invoice_group['data']
        
        if group_df.empty:
            print(f"Warning: Empty data for invoice {invoice_num}")
            return None
            
        filename = os.path.join(self.output_folder, f"Invoice_{invoice_num}.pdf")
        
        doc = SimpleDocTemplate(
            filename,
            pagesize=letter,
            rightMargin=self.margin,
            leftMargin=self.margin,
            topMargin=self.margin,
            bottomMargin=self.margin
        )
        
        elements = []

        # Company address definition
        company_address = "New York City, NY, 18900-1656"
        company_address_style = ParagraphStyle(
            'CompanyAddress',
            parent=self.styles['Normal'],
            fontSize=8,
            fontName='Helvetica-Bold',
            alignment=0,
            leading=10,
            spaceBefore=2,
            spaceAfter=6
        )

        # Bill to address style
        billing_address_style = ParagraphStyle(
            'BillingAddress',
            parent=self.styles['Normal'],
            fontSize=8,
            fontName='Helvetica',
            alignment=2,
            leading=10,
            spaceBefore=0,
            spaceAfter=0
        )
        
        # Logo and address handling
        logo_path = os.path.join(os.getcwd(), 'logo.png')
        if os.path.exists(logo_path):
            try:
                img = Image(logo_path)
                img.drawHeight = 1*inch
                img.drawWidth = 2*inch
                
                # Get the first row for billing information
                first_row = group_df.iloc[0]
                
                # Create billing address content
                billing_lines = []
                
                # Add Bill to header
                billing_lines.append(Paragraph('Bill to:', company_address_style))
                
                # Add address lines if they exist and are not zero
                addr1 = self.sanitize_text(first_row.get('Addr1', ''))
                addr2 = self.sanitize_text(first_row.get('addr2', ''))
                addr3 = self.sanitize_text(first_row.get('addr3', ''))
                
                # Only add non-empty address lines
                if addr1 and addr1.strip() != '0':
                    billing_lines.append(Paragraph(addr1, billing_address_style))
                if addr2 and addr2.strip() != '0':
                    billing_lines.append(Paragraph(addr2, billing_address_style))
                if addr3 and addr3.strip() != '0':
                    billing_lines.append(Paragraph(addr3, billing_address_style))
                
                # Combine City, State, Zip on one line
                city = self.sanitize_text(first_row.get('City', ''))
                state = self.sanitize_text(first_row.get('State', ''))
                zip_code = self.sanitize_text(first_row.get('Zip', ''))
                
                # Create combined address line with non-empty, non-zero components
                address_parts = []
                if city and city.strip() != '0':
                    address_parts.append(city)
                if state and state.strip() != '0':
                    address_parts.append(state)
                if zip_code and zip_code.strip() != '0':
                    address_parts.append(zip_code)
                
                if address_parts:
                    combined_address = ', '.join(address_parts)
                    billing_lines.append(Paragraph(combined_address, billing_address_style))
                
                # Create the combined table with logo and billing address
                top_row = [[img, billing_lines]]
                bottom_row = [[Paragraph(company_address, company_address_style), '']]
                
                combined_table = Table(
                    top_row + bottom_row,
                    colWidths=[self.content_width * 0.6, self.content_width * 0.4]
                )
                
                combined_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                    ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('LEFTPADDING', (0, 0), (-1, -1), 0),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                    ('TOPPADDING', (0, 0), (-1, -1), 0),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                ]))
                
                elements.append(combined_table)
                elements.append(Spacer(1, 20))
            except Exception as e:
                print(f"Error loading logo: {e}")
                # Continue without the logo

        def create_safe_paragraph(text, style):
            """
            Creates a paragraph with enhanced safety measures to prevent ReportLab rendering issues.
            """
            try:
                # Handle empty or None input
                if text is None or str(text).strip() == '':
                    return Paragraph('&nbsp;', style)
                
                # Prepare the text
                safe_text = str(text)
                
                # Convert to plain ASCII and remove problematic characters
                safe_text = safe_text.encode('ascii', 'ignore').decode('ascii')
                
                # Remove control characters while preserving spaces
                safe_text = ''.join(char if ord(char) >= 32 else ' ' for char in safe_text)
                
                # Normalize whitespace while ensuring content
                safe_text = ' '.join(filter(None, safe_text.split()))
                
                # Ensure minimum content and proper HTML encoding
                if not safe_text or safe_text.isspace():
                    return Paragraph('&nbsp;', style)
                    
                # Add safety padding to prevent layout issues
                safe_text = f"{safe_text}&nbsp;"
                
                return Paragraph(safe_text, style)
            except Exception as e:
                print(f"Error in paragraph creation: {e}")
                return Paragraph('&nbsp;', style)
                
        
        self.header_style_with_spacing = ParagraphStyle(
            'HeaderStyleWithSpacing',
            parent=self.styles['Normal'],
            fontSize=10,
            leading=12,
            spaceBefore=1,
            spaceAfter=1,
            wordWrap='LTR',
            alignment=0,
            allowWidows=0,
            allowOrphans=0,
            splitLongWords=1,
            firstLineIndent=0,
            leftIndent=0
        )

        self.header_style_with_spacing2 = ParagraphStyle(
            'HeaderStyleWithSpacing2',
            parent=self.styles['Normal'],
            fontSize=10,
            leading=12,
            spaceBefore=1,
            spaceAfter=1,
            wordWrap='LTR',
            alignment=2,
            allowWidows=0,
            allowOrphans=0,
            splitLongWords=1,
            firstLineIndent=0,
            leftIndent=0
        )

        self.payment_terms_style = ParagraphStyle(
            'PaymentTerms',
            parent=self.styles['Normal'],
            fontSize=8,
            fontName='Helvetica-Bold',
            leading=10,
            spaceBefore=1,
            spaceAfter=1,
            alignment=2,
            allowWidows=0,
            allowOrphans=0,
            splitLongWords=0
        )

        # Get the first row for header information
        first_row = group_df.iloc[0]
        
        # Format the date to YYYY-MM-DD
        formatted_date = self.format_date(first_row.get('Invoice Date', ''))
        
        # Use 'Correct Division' field for the Division display
        division_value = first_row.get('Correct Division', '')
        
        # Header contents based on project requirements
        header_contents = {
            'left': [
                ('Division', division_value),
                ('Invoice Date', formatted_date),
                ('Department', first_row.get('Correct Dept ID', '')),
                ('Department Name', first_row.get('Correct Dept Name', ''))
            ],
            'right': [
                ('Payment Terms', 'Due Immediate'),
                (None, None),
                (None, None),
                None
            ]
        }

        # Safe paragraph creation for both columns
        left_col = []
        right_col = []

        # Process header columns
        for label, value in header_contents['left']:
            try:
                label = self.sanitize_text(label or "")
                value = self.sanitize_text(value or "")
                text = f"{label}: {value}" if label.strip() and value.strip() else (label or value or " ")
                para = create_safe_paragraph(text, self.header_style_with_spacing)
                left_col.append(para)
            except Exception as e:
                print(f"Error creating left paragraph: {e}")
                left_col.append(create_safe_paragraph(" ", self.header_style_with_spacing))

            
       # Then modify the right column processing code:
        for item in header_contents['right']:
            try:
                if item is None:
                    # When we encounter None, create an empty paragraph
                    para = create_safe_paragraph(" ", self.header_style_with_spacing2)
                else:
                    # When we have a tuple, process it as before
                    label, value = item
                    if label is None and value is None:
                        # Skip creating a paragraph with colons when both are None
                        para = create_safe_paragraph(" ", self.header_style_with_spacing2)
                    else:
                        label = self.sanitize_text(label or "")
                        value = self.sanitize_text(value or "")
                        text = f"{label}: {value}" if label and value else (label or value or " ")
                        style = self.payment_terms_style if label == 'Payment Terms' else self.header_style_with_spacing2
                        para = create_safe_paragraph(text, style)
                right_col.append(para)
            except Exception as e:
                print(f"Error creating right paragraph: {e}")
                right_col.append(create_safe_paragraph(" ", self.header_style_with_spacing2))

        header_data = [
            [left_col[0], right_col[0]],  # Division + Payment Terms
            [left_col[1], right_col[1]],  # Invoice Date
            [left_col[2], right_col[2]],  # Department
            [left_col[3], right_col[3]]   # Department Name
        ]

        try:
            # Create header table
            header_table = Table(
                header_data,
                colWidths=[self.content_width * 0.6, self.content_width * 0.4],
                rowHeights=[12] * 4
            )
            
            header_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('TOPPADDING', (0, 0), (-1, -1), 1),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
                ('LEFTPADDING', (0, 0), (-1, -1), 2),
                ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ]))

        except Exception as e:
            print(f"Error creating header table: {e}")
            # Create an extremely simple fallback table
            header_table = Table(
                [[Paragraph('&nbsp;', self.header_style)] * 2] * 4,
                colWidths=[self.content_width * 0.5] * 2,
                rowHeights=[12] * 4
            )
    
        elements.extend([
            create_safe_paragraph(f"Invoice #{invoice_num}", self.title_style),
            Spacer(1, 15),
            header_table,
            Spacer(1, 30)
        ])

        # Detail table setup - revised columns as requested (removed Charge Code column)
        col_widths = [
            float(self.content_width * 0.50),  # Charge Description
            float(self.content_width * 0.15),  # Quantity
            float(self.content_width * 0.15),  # Amount
            float(self.content_width * 0.20),  # Total Tax
        ]

        detail_data = [['Charge Description', 'Quantity', 'Amount', 'Total Tax']]
        
        total_amount = 0.0
        total_tax = 0.0
        
        # Process rows for detail table
        for _, row in group_df.iterrows():
            try:
                # Calculate total tax for this row
                row_tax = 0.0
                tax_fields = ['State-Tax', 'County-Tax', 'City-Tax', 'HST/GST Tax', 'PST/QST Tax']
                for tax_field in tax_fields:
                    if tax_field in row and not pd.isna(row[tax_field]):
                        tax_value = self._convert_european_number(row[tax_field])
                        row_tax += tax_value
                
                # Get amount
                amount = self._convert_european_number(row['Amount'])
                
                # Get quantity
                quantity = self._convert_european_number(row['Unit-Qty']) if 'Unit-Qty' in row else 1
                
                # Use Bill Code Description instead of Description as requested
                description = self.sanitize_text(row.get('Bill Code Description', ''))
                
                # Add row to detail data (without Charge Code column)
                detail_data.append([
                    description,
                    str(int(quantity)) if quantity == int(quantity) else str(quantity),
                    self.format_currency(amount),
                    self.format_currency(row_tax)
                ])
                
                # Update totals
                total_amount += amount
                total_tax += row_tax
            except (ValueError, TypeError) as e:
                print(f"Warning: Error processing row data: {e}")
                continue
                    
        # Create and style detail table
        detail_table = Table(detail_data, colWidths=col_widths)
        detail_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#6AACC1')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),  # Charge Description (left-aligned)
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),  # Quantity
            ('ALIGN', (2, 0), (2, -1), 'RIGHT'),  # Amount
            ('ALIGN', (3, 0), (3, -1), 'RIGHT'),  # Total Tax
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('GRID', (0, 0), (-1, -1), 1, colors.grey),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
        ]))

        elements.append(detail_table)
        elements.append(Spacer(1, 30))

        # Add totals
        elements.extend([
            create_safe_paragraph(f"Subtotal: {self.format_currency(total_amount)}", self.summary_style),
            create_safe_paragraph(f"Total Tax: {self.format_currency(total_tax)}", self.summary_style),
            create_safe_paragraph(f"Total Invoice Amount: {self.format_currency(total_amount + total_tax)}", self.summary_style),
            Spacer(1, 30)
        ])
        
        # Footer handling
        footer_path = os.path.join(os.getcwd(), 'foot.png')
        if os.path.exists(footer_path):
            try:
                footer_img = Image(footer_path)
                desired_width = (299/72) * inch
                desired_height = (130/72) * inch
                footer_img.drawWidth = desired_width
                footer_img.drawHeight = desired_height
                
                padding = float((self.content_width - desired_width) / 2)
                footer_table = Table([[footer_img]], colWidths=[self.content_width])
                footer_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('LEFTPADDING', (0, 0), (-1, -1), padding),
                    ('RIGHTPADDING', (0, 0), (-1, -1), padding),
                    ('TOPPADDING', (0, 0), (-1, -1), 0),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                ]))
                elements.append(footer_table)
            except Exception as e:
                print(f"Error loading footer image: {e}")

        # Build the PDF
        try:
            doc.build(elements)
            print(f"Generated PDF for Invoice: {invoice_num}")
            # Store the mapping
            self.generated_mappings[invoice_num] = [invoice_num]
            return invoice_num
        except Exception as e:
            print(f"Error generating PDF for invoice {invoice_num}: {str(e)}")
            import traceback
            print(f"Full error details:\n{traceback.format_exc()}")
            return None

    def generate_all_invoices(self):
        """
        Generate invoices grouped by invoice number.
        """
        invoice_groups = self.prepare_invoice_groups()
        print(f"Total invoice groups to process: {len(invoice_groups)}")
        
        # Handle case with no groups
        if not invoice_groups:
            print("No invoice data to process. Check input data.")
            return self.generated_mappings
            
        batch_size = 100
        for i in range(0, len(invoice_groups), batch_size):
            batch = invoice_groups[i:i + batch_size]
            print(f"Processing batch {i//batch_size + 1} of {(len(invoice_groups) + batch_size - 1)//batch_size}")
            
            for group_idx, group in enumerate(batch):
                try:
                    invoice_number = self.generate_invoice(group)
                    if invoice_number:
                        print(f"Successfully generated invoice {invoice_number}")
                except Exception as e:
                    print(f"Error generating invoice for group at index {i + group_idx}: {str(e)}")
                    import traceback
                    print(f"Error details:\n{traceback.format_exc()}")
        
        print(f"Total invoices generated: {len(self.generated_mappings)}")
        return self.generated_mappings

if __name__ == "__main__":
    try:
        generator = InvoiceGenerator()
        generator.generate_all_invoices()
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        import traceback
        print(f"Full error details:\n{traceback.format_exc()}")