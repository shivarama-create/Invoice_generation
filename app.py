import os
import zipfile
import shutil
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT, TA_JUSTIFY
from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory, send_file
from io import BytesIO
import datetime

# --- Core Invoice Generation Logic ---

def get_safe_value(df_row, key, default=''):
    """
    Safely retrieves a value from a DataFrame row, returning a default
    if the key is missing or the value is NaN/NaT.
    """
    value = df_row.get(key, default)
    if pd.isna(value) or pd.isnull(value):
        return default
    return value

def format_invoice_date(date_val):
    """
    Formats a date like 5092025 into 'DD Mon YY' format.
    """
    if not date_val or pd.isna(date_val):
        return datetime.date.today().strftime('%d %b %y')
    try:
        # Assuming format MMDDYYYY or MDDYYYY
        date_str = str(int(date_val))
        if len(date_str) == 7: # MDDYYYY
             date_str = '0' + date_str
        dt = datetime.datetime.strptime(date_str, '%m%d%Y')
        return dt.strftime('%d %b %y')
    except (ValueError, TypeError):
        return datetime.date.today().strftime('%d %b %y')


def generate_individual_invoice(df_row, output_dir):
    """
    Generates a single PDF invoice for a given customer row, matching the sample format.
    """
    # --- Data Extraction ---
    invoice_number = get_safe_value(df_row, 'Invoice Number')
    awb_number = get_safe_value(df_row, 'House AWB No.', 'N/A')
    export_ref = get_safe_value(df_row, 'Reference_1')
    invoice_date_raw = get_safe_value(df_row, 'Invoice Date')
    invoice_date = format_invoice_date(invoice_date_raw)

    # Recipient is the 'buyer' or 'importer other than consignee'
    recipient_name = get_safe_value(df_row, 'Recipient_Contact Name')
    recipient_address1 = get_safe_value(df_row, 'Recipient_Address Line 1')
    recipient_address2 = get_safe_value(df_row, 'Recipient_Address Line 2')
    recipient_city = get_safe_value(df_row, 'Recipient_City')
    recipient_state = get_safe_value(df_row, 'Recipient_State')
    recipient_postal_code = get_safe_value(df_row, 'Recipient_Postal code')
    recipient_country_code = get_safe_value(df_row, 'Recipient_Country', '').split('-')[0]
    recipient_phone = get_safe_value(df_row, 'Recipient_Phone Number')
    recipient_email = get_safe_value(df_row, 'Recipient_Email')

    # Build recipient addresses, filtering out empty lines
    recipient_address_lines = [recipient_address1, recipient_address2, f"{recipient_city} {recipient_state} {recipient_postal_code} {recipient_country_code}"]
    recipient_full_address = "<br/>".join(filter(None, recipient_address_lines))


    # Item Details
    description = get_safe_value(df_row, 'COMMODITY')
    state_origin = get_safe_value(df_row, 'St. of Origin of goods')
    district_origin = get_safe_value(df_row, 'Dis. Of Origin of goods')
    hs_code = get_safe_value(df_row, 'HS CODE 1')
    country_mfg = get_safe_value(df_row, 'Country of Manufacture').replace('-INDIA', '')
    net_weight = float(get_safe_value(df_row, 'UNIT_Weight 1', 0.0))
    quantity = int(get_safe_value(df_row, 'QUANTITY 1', 0))
    uom = get_safe_value(df_row, 'UOM1')
    unit_value = float(get_safe_value(df_row, 'UNIT_VALUE 1', 0.0))
    total_value = float(get_safe_value(df_row, 'Invoice Value', 0.0))
    currency = get_safe_value(df_row, 'CURRENCY', 'USD').split('-')[0]
    freight_charges = float(get_safe_value(df_row, 'Freight_charges', 0.0))
    total_invoice_amount = total_value + freight_charges


    # --- PDF Generation ---
    # Create a safe filename. The AWB number can be 'N/A', which contains '/', an invalid character for filenames.
    # We fall back to the invoice number if the AWB number is not available or is 'N/A'.
    if awb_number and awb_number != 'N/A':
        base_filename = str(awb_number).replace('/', '_').replace('\\', '_')
    else:
        base_filename = f"invoice_{invoice_number}"

    filename = os.path.join(output_dir, f'{base_filename}.pdf')
    doc = SimpleDocTemplate(filename, pagesize=A4, topMargin=0.5*cm, bottomMargin=0.5*cm, leftMargin=1*cm, rightMargin=1*cm)

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Bold', fontName='Helvetica-Bold'))
    styles.add(ParagraphStyle(name='RightAlign', alignment=TA_RIGHT))
    styles.add(ParagraphStyle(name='CenterAlign', alignment=TA_CENTER))
    styles.add(ParagraphStyle(name='Small', fontSize=8, leading=10))
    styles.add(ParagraphStyle(name='SmallBold', fontName='Helvetica-Bold', fontSize=8, leading=10))

    story = []

    # Header Section
    header_data = [
        ['', Paragraph('<b>COMMERCIAL INVOICE</b>', styles['CenterAlign']), Paragraph(f'<b>FedEx INTERNATIONAL AIRWAYBILL<br/>{awb_number}</b>', styles['CenterAlign'])],
        [Paragraph('<b>DATE OF EXPORT</b>', styles['SmallBold']), invoice_date, Paragraph('<b>EXPORT REFERENCES</b>', styles['SmallBold']), export_ref],
        [Paragraph('<b>INVOICE NUMBER</b>', styles['SmallBold']), invoice_number, Paragraph('<b>INVOICE DATE</b>', styles['SmallBold']), invoice_date]
    ]
    header_table = Table(header_data, colWidths=[doc.width/4]*4)
    header_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('SPAN', (1,0), (2,0)),
        ('SPAN', (3,0), (-1,0)),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(header_table)

    # Shipper / Consignee Section
    shipper_info = Paragraph("""<b>SHIPPER/EXPORTER</b><br/><br/>
    Mitul Sanghvi<br/>
    Fabrics and More<br/>
    RAMANI COMPUND OPP HP PETROL<br/>
    PUMP,SV ROAD<br/>
    DAHISAR E<br/>
    MUMBAI MH 400068 IN<br/>
    TEL: 7021460762<br/>
    SHIPPER'S TAX NUMBER: 27CTWPR7908H1ZQ
    """, styles['Small'])

    consignee_info = Paragraph("""<b>RECIPIENT/CONSIGNEE</b><br/><br/>
    Cozy Corner Patios LLC<br/>
    Cozy Corner Patios LLC<br/>
    1499 W 120th Ave<br/>
    Unit 110<br/>
    Westminster CO 80241 US<br/>
    TEL: 7206277225
    """, styles['Small'])

    importer_info = Paragraph(f"""<b>IMPORTER OTHER THAN C/NEE OR BILL TO PARTY</b><br/><br/>
    {recipient_name}<br/>
    {recipient_full_address}<br/>
    TEL: {recipient_phone}<br/>
    EMAIL: {recipient_email}
    """, styles['Small'])

    contact_data = [[shipper_info, consignee_info], ['', importer_info]]
    contact_table = Table(contact_data, colWidths=[doc.width/2, doc.width/2])
    contact_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    story.append(contact_table)
    story.append(Spacer(1, 0.2*cm))

    # Details Table
    item_table_data = [
        [Paragraph('<b>S.N-O</b>', styles['SmallBold']), Paragraph('<b>FULL DESCRIPTION OF GOODS</b>', styles['SmallBold']), Paragraph('<b>STATE OF ORIGIN GOODS</b>', styles['SmallBold']), Paragraph('<b>HS CODE</b>', styles['SmallBold']), Paragraph('<b>COUNTRY OF MFG</b>', styles['SmallBold']), Paragraph('<b>NET WGT KG</b>', styles['SmallBold']), Paragraph('<b>QTY</b>', styles['SmallBold']), Paragraph('<b>UOM</b>', styles['SmallBold']), Paragraph('<b>UNIT VALUE</b>', styles['SmallBold']), Paragraph('<b>TOTAL VALUE</b>', styles['SmallBold'])],
        ['1', Paragraph(description, styles['Small']), Paragraph(state_origin, styles['Small']), Paragraph(str(hs_code), styles['Small']), Paragraph(country_mfg, styles['Small']), f'{net_weight:.2f}', str(quantity), Paragraph(uom, styles['Small']), f'{unit_value:.2f}', f'{total_value:.2f}']
    ]
    item_table = Table(item_table_data, colWidths=[0.5*inch, 2*inch, 0.8*inch, 0.8*inch, 0.7*inch, 0.6*inch, 0.4*inch, 0.5*inch, 0.6*inch, 0.6*inch])
    item_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('LEFTPADDING', (0,0), (-1,-1), 2),
        ('RIGHTPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(item_table)

    # Footer section
    declaration = Paragraph("I DECLARE ALL THE INFORMATION CONTAINED IN THIS INVOICE IS TRUE AND CORRECT TO THE BEST OF MY KNOWLEDGE.", styles['Small'])
    signature = Paragraph("<b>NAME (PLEASE PRINT)</b> Fabrics and More<br/><br/><br/>Mitul Sanghvi", styles['Small'])
    
    footer_data = [
        ['', Paragraph('<b>TOTAL</b>', styles['RightAlign']), f'{total_value:.2f}'],
        [Paragraph(f'<b>CURRENCY IN WORDS:</b> {currency}', styles['SmallBold']), Paragraph('<b>TOTAL FREIGHT CHARGES</b>', styles['RightAlign']), f'{freight_charges:.2f}'],
        [declaration, Paragraph('<b>TOTAL INVOICE AMOUNT {currency}</b>'.format(currency=currency), styles['RightAlign']), f'{total_invoice_amount:.2f}'],
        [signature, '', '']
    ]
    footer_table = Table(footer_data, colWidths=[4*inch, 2*inch, 1.5*inch])
    footer_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('SPAN', (0,2), (0,3)),
    ]))

    story.append(footer_table)
    
    doc.build(story)
    print(f"Generated individual invoice: {os.path.basename(filename)}")
    return os.path.basename(filename)


def draw_mother_invoice_header(canvas, doc):
    """ Draws the header on each page of the mother invoice. """
    canvas.saveState()
    styles = getSampleStyleSheet()
    # Reduced leading for more compact header text
    styles.add(ParagraphStyle(name='Small', fontSize=8, leading=9))
    styles.add(ParagraphStyle(name='SmallRight', fontSize=8, leading=9, alignment=TA_RIGHT))
    
    # Define a pixel unit for easier padding calculation
    px = inch / 72

    # Available width for content
    width = doc.width

    shipper_info = Paragraph("""<b>Shipper</b><br/>
        Fabrics and More<br/>
        Mitul Sanghvi<br/>
        RAMANI COMPUND OPP HP PETROL, PUMP,SV ROAD, DAHISAR E, 400068, MUMBAI, MAHARASHTRA, IN<br/>
        7021460762<br/>
        GSTIN: 27CTWPR7908H1ZQ""", styles['Small'])
    
    consignee_info = Paragraph("""<b>Consignee</b><br/>
        Cozy Corner Patios LLC<br/>
        Cozy Corner Patios LLC<br/>
        1499 W 120th Ave, Unit 110, 80241, Westminster, COLORADO, UNITED STATES<br/>
        7206277225<br/>
        cozycornerpatios@gmail.com""", styles['Small'])

    invoice_date = datetime.date.today().strftime("%d-%b-%Y")
    invoice_no = f"FAM/{datetime.date.today().strftime('%d%m%Y')}/{doc.page}"

    invoice_details_content = [
        [Paragraph('<b>Invoice No:</b>', styles['Small']), Paragraph(invoice_no, styles['SmallRight'])],
        [Paragraph('<b>Date:</b>', styles['Small']), Paragraph(invoice_date, styles['SmallRight'])],
        [Paragraph('<b>Place of Receipt By Shipper:</b>', styles['Small']), Paragraph('N/A', styles['SmallRight'])],
        [Paragraph('<b>City/Port Of Loading:</b>', styles['Small']), Paragraph('N/A', styles['SmallRight'])],
        [Paragraph('<b>City/Port of Discharge:</b>', styles['Small']), Paragraph('N/A', styles['SmallRight'])],
        [Paragraph('<b>Reason for Export:</b>', styles['Small']), Paragraph('N/A', styles['SmallRight'])],
        [Paragraph('<b>Terms Of Trade:</b>', styles['Small']), Paragraph('CIF', styles['SmallRight'])],
        [Paragraph('<b>Place of Supply:</b>', styles['Small']), Paragraph('N/A', styles['SmallRight'])],
        [Paragraph('<b>AD Code:</b>', styles['Small']), Paragraph('6390614-291009', styles['SmallRight'])],
        [Paragraph('<b>IEC:</b>', styles['Small']), Paragraph('CTWPR7908H', styles['SmallRight'])],
    ]
    invoice_details_table = Table(invoice_details_content, colWidths=[1.4*inch, 1.0*inch])

    header_data = [
        [shipper_info, consignee_info, invoice_details_table]
    ]
    header_table = Table(header_data, colWidths=[width*0.35, width*0.35, width*0.30])
    header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP')]))
    
    # --- Corrected Positioning Logic ---
    
    # 1. Get actual height of the header table
    _w, actual_header_height = header_table.wrapOn(canvas, width, 0)
    
    # 2. Calculate the Y position for the bottom of the header table.
    # This places its top edge 10px from the top of the page.
    header_y_pos = doc.pagesize[1] - (10 * px) - actual_header_height
    
    # 3. Draw the header table
    header_table.drawOn(canvas, doc.leftMargin, header_y_pos)

    # 4. Position title above the main content table
    title = Paragraph("<b>Commercial Invoice cum Packing List</b>", styles['h2'])
    _w, title_h = title.wrapOn(canvas, width, 0) # Get actual height
    
    # 5. Calculate Y position for the title, placing it 15px above the main content table.
    main_content_start_y = doc.pagesize[1] - doc.topMargin
    title_y_pos = main_content_start_y + (15 * px)
    
    # 6. Draw the title
    title.drawOn(canvas, doc.leftMargin, title_y_pos)

    canvas.restoreState()


def generate_mother_invoice(df, output_dir):
    """ Generates a single 'mother invoice' PDF that summarizes all individual invoices. """
    filename = os.path.join(output_dir, 'mother_invoice.pdf')
    # Adjusted top margin to ensure space for the manually drawn header and title
    doc = SimpleDocTemplate(filename, pagesize=A4, topMargin=2.2*inch, bottomMargin=1*inch, leftMargin=1*cm, rightMargin=1*cm)

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Small', fontSize=7, leading=9))
    styles.add(ParagraphStyle(name='SmallBold', fontName='Helvetica-Bold', fontSize=7, leading=9, alignment=TA_CENTER))
    styles.add(ParagraphStyle(name='RightAlign', alignment=TA_RIGHT))


    story = []
    
    header = [Paragraph(h, styles['SmallBold']) for h in ['S.No', 'Buyer', 'Invoice No.', 'Order ID', 'AWB No.', 'Description HSN UOM', 'Net Wt (KG)', 'Qty', 'Unit Val', 'Total Val', 'IGST %', 'IGST Paid']]
    data = [header]

    for idx, row in df.iterrows():
        description_text = f"{get_safe_value(row, 'COMMODITY')} <br/> {get_safe_value(row, 'HS CODE 1')} <br/> {get_safe_value(row, 'UOM1')}"
        data_row = [
            Paragraph(str(idx + 1), styles['Small']),
            Paragraph(get_safe_value(row, 'Recipient_Contact Name'), styles['Small']),
            Paragraph(str(get_safe_value(row, 'Invoice Number')), styles['Small']),
            Paragraph(str(get_safe_value(row, 'Reference_1')), styles['Small']),
            Paragraph(str(get_safe_value(row, 'House AWB No.')), styles['Small']),
            Paragraph(description_text, styles['Small']),
            Paragraph(f"{float(get_safe_value(row, 'Total Shipment weight', 0.0)):.2f}", styles['Small']),
            Paragraph(str(get_safe_value(row, 'QUANTITY 1', 0)), styles['Small']),
            Paragraph(f"{float(get_safe_value(row, 'UNIT_VALUE 1', 0.0)):.2f}", styles['Small']),
            Paragraph(f"{float(get_safe_value(row, 'Invoice Value', 0.0)):.2f}", styles['Small']),
            Paragraph('0', styles['Small']),
            Paragraph('0.0', styles['Small'])
        ]
        data.append(data_row)
    
    # Define column widths to fit the page
    col_widths = [0.4*inch, 1.1*inch, 0.7*inch, 0.7*inch, 0.8*inch, 1.5*inch, 0.5*inch, 0.4*inch, 0.5*inch, 0.5*inch, 0.4*inch, 0.5*inch]
    summary_table = Table(data, colWidths=col_widths, repeatRows=1)
    
    summary_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('LEFTPADDING', (0,0), (-1,-1), 2),
        ('RIGHTPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(summary_table)

    # Footer section with totals
    total_packages = len(df)
    total_value = df['Invoice Value'].astype(float).sum()
    total_weight = df['Total Shipment weight'].astype(float).sum()
    currency = get_safe_value(df.iloc[0], 'CURRENCY', 'USD').split('-')[0]

    footer_text = f"""
    <b>Total Packages:</b> {total_packages}<br/>
    <b>Total Invoice Value:</b> {total_value:.2f}<br/>
    <b>Total Weight (Kg):</b> {total_weight:.2f}<br/>
    <b>Currency:</b> {currency}
    """
    
    signature = """<br/><br/><br/><br/>
    For Fabrics and More<br/><br/><br/>
    Authorised Signatory
    """

    footer_data = [
        [Paragraph(footer_text, styles['Normal']), '', Paragraph(signature, styles['RightAlign'])]
    ]
    footer_table = Table(footer_data, colWidths=[doc.width/3, doc.width/3, doc.width/3])
    story.append(Spacer(1, 0.5*cm))
    story.append(footer_table)

    doc.build(story, onFirstPage=draw_mother_invoice_header, onLaterPages=draw_mother_invoice_header)
    print(f"Generated mother invoice: {os.path.basename(filename)}")
    return os.path.basename(filename)


def create_child_invoices_zip(output_dir, child_invoice_files):
    """
    Creates a zip file containing all child invoices (excluding mother invoice).
    Returns the path to the created zip file.
    """
    zip_filename = os.path.join(output_dir, 'child_invoices.zip')
    
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for invoice_file in child_invoice_files:
            if invoice_file != 'mother_invoice.pdf':  # Exclude mother invoice
                file_path = os.path.join(output_dir, invoice_file)
                if os.path.exists(file_path):
                    zipf.write(file_path, invoice_file)
    
    print(f"Created zip file: {os.path.basename(zip_filename)}")
    return os.path.basename(zip_filename)


# --- Flask Web App Logic ---

app = Flask(__name__, static_folder='static')
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['GENERATED_PDFS_FOLDER'] = 'static/invoices'
app.secret_key = 'super_secret_key'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['GENERATED_PDFS_FOLDER'], exist_ok=True)

# These are the columns we hope to find, but the code will handle their absence
EXPECTED_COLUMNS = [
    'Invoice Number', 'Recipient_Contact Name', 'Recipient_Address Line 1',
    'Recipient_City', 'Recipient_State', 'Recipient_Country', 'Recipient_Postal code',
    'COMMODITY', 'QUANTITY 1', 'UNIT_VALUE 1', 'CURRENCY', 'Invoice Value', 
    'Total Shipment weight', 'HS CODE 1', 'UOM1'
]

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)

        # Clear previous invoices
        for item in os.listdir(app.config['GENERATED_PDFS_FOLDER']):
            os.remove(os.path.join(app.config['GENERATED_PDFS_FOLDER'], item))

        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)

        df = None
        try:
            if file.filename.lower().endswith('.csv'):
                df = pd.read_csv(filepath)
            elif file.filename.lower().endswith(('.xls', '.xlsx')):
                dfs_from_file = []
                with pd.ExcelFile(filepath) as xls:
                    for sheet_name in xls.sheet_names:
                        if 'recipient and invoice data' in sheet_name.lower():
                            temp_df = pd.read_excel(xls, sheet_name=sheet_name)
                            dfs_from_file.append(temp_df)
                
                if dfs_from_file:
                    df = pd.concat(dfs_from_file, ignore_index=True)
                else:
                    flash("No sheets named 'Recipient and Invoice Data' found in the Excel file.")
                    return redirect(request.url)
            else:
                 flash("Unsupported file type. Please upload CSV, XLS, or XLSX.")
                 return redirect(request.url)

        except Exception as e:
            flash(f"Error reading the file: {e}")
            return redirect(request.url)

        # Basic validation
        if df is None or df.empty:
            flash("Could not find any valid data in the uploaded file.")
            return redirect(request.url)
        
        # Data Cleaning
        df.dropna(subset=['Invoice Number'], inplace=True)
        # Convert key columns to string to avoid scientific notation on numbers
        for col in ['Invoice Number', 'Recipient_Postal code', 'Reference_1', 'House AWB No.']:
             if col in df.columns:
                 df[col] = df[col].astype(str).replace('\.0', '', regex=True)
        
        # This will hold the data to display in the table
        invoice_list_data = []
        generated_files = []
        child_invoice_files = []
        
        for _, row in df.iterrows():
            try:
                filename = generate_individual_invoice(row, app.config['GENERATED_PDFS_FOLDER'])
                generated_files.append(filename)
                child_invoice_files.append(filename)
                
                # Extract data for the table
                invoice_list_data.append({
                    'invoice_number': get_safe_value(row, 'Invoice Number'),
                    'recipient_name': get_safe_value(row, 'Recipient_Contact Name'),
                    'recipient_phone': get_safe_value(row, 'Recipient_Phone Number'),
                    'recipient_state': get_safe_value(row, 'Recipient_State'),
                    'download_link': url_for('serve_invoice', filename=filename)
                })

            except Exception as e:
                inv_num = get_safe_value(row, 'Invoice Number', 'Unknown')
                flash(f"Could not generate invoice {inv_num}. Error: {e}")

        zip_filename = None
        if child_invoice_files:
            try:
                zip_filename = create_child_invoices_zip(app.config['GENERATED_PDFS_FOLDER'], child_invoice_files)
                generated_files.append(zip_filename)
            except Exception as e:
                flash(f"Could not create zip file. Error: {e}")
                
        mother_invoice_filename = None
        if not df.empty:
            try:
                mother_invoice_filename = generate_mother_invoice(df, app.config['GENERATED_PDFS_FOLDER'])
                generated_files.append(mother_invoice_filename)
            except Exception as e:
                flash(f"Could not generate mother invoice. Error: {e}")
                
        # Clean up uploaded file
        os.remove(filepath)
        
        # Pass the extracted data to the template
        return render_template('index.html', 
                               files=sorted(generated_files), 
                               has_zip=zip_filename is not None, 
                               mother_invoice_file=mother_invoice_filename,
                               invoice_list=invoice_list_data)

    return render_template('index.html', files=None, has_zip=False)

@app.route('/static/invoices/<filename>')
def serve_invoice(filename):
    return send_from_directory(app.config['GENERATED_PDFS_FOLDER'], filename)

@app.route('/download/child-invoices-zip')
def download_child_invoices_zip():
    """Download all child invoices as a zip file"""
    zip_path = os.path.join(app.config['GENERATED_PDFS_FOLDER'], 'child_invoices.zip')
    if os.path.exists(zip_path):
        return send_file(zip_path, as_attachment=True, download_name='child_invoices.zip')
    else:
        flash('Zip file not found. Please regenerate invoices.')
        return redirect(url_for('index'))
        
@app.route('/download/mother-invoice')
def download_mother_invoice():
    """Download the mother invoice"""
    mother_invoice_path = os.path.join(app.config['GENERATED_PDFS_FOLDER'], 'mother_invoice.pdf')
    if os.path.exists(mother_invoice_path):
        return send_file(mother_invoice_path, as_attachment=True, download_name='mother_invoice.pdf')
    else:
        flash('Mother invoice not found. Please regenerate invoices.')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
