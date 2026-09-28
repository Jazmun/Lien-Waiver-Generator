import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from datetime import date
import io
import os

st.set_page_config(page_title="Lien Waiver Generator", page_icon="📄", layout="centered")

COMPANY_NAME = "La Salle Landscaping and Lawn Care"
COMPANY_ADDRESS = "PO BOX 68, Bellaire, TX 77402"
COMPANY_PHONE = "713-657-0875"
COMPANY_EMAIL = "accounting@lasallelandscaping.com"
OFFICER_NAME = "Ben Gil"
OFFICER_TITLE = "Controller"

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    xml_str = f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>'
    tcPr.append(parse_xml(xml_str))

def set_cell_border(cell, **kwargs):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    for edge, border_args in kwargs.items():
        edge_data = f'<w:{edge} {nsdecls("w")} w:val="{border_args.get("val", "single")}" w:sz="{border_args.get("sz", 4)}" w:space="0" w:color="{border_args.get("color", "CCCCCC")}"/>'
        tcBorders.append(parse_xml(edge_data))
    tcPr.append(tcBorders)

def number_to_words(n):
    ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
            "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
            "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
    
    def _to_words(num):
        if num < 20:
            return ones[num]
        elif num < 100:
            return tens[num // 10] + (" " + ones[num % 10] if num % 10 != 0 else "")
        elif num < 1000:
            return ones[num // 100] + " Hundred" + (" " + _to_words(num % 100) if num % 10 != 0 else "")
        elif num < 1000000:
            return _to_words(num // 1000) + " Thousand" + (" " + _to_words(num % 1000) if num % 1000 != 0 else "")
        elif num < 1000000000:
            return _to_words(num // 1000000) + " Million" + (" " + _to_words(num % 1000000) if num % 1000000 != 0 else "")
        return str(num)
    
    return _to_words(n) if n > 0 else "Zero"

def format_currency_words(amount):
    dollars = int(amount)
    cents = int(round((amount - dollars) * 100))
    words = number_to_words(dollars)
    return f"{words} and {cents:02d}/100 Dollars"

def generate_docx(data):
    doc = docx.Document()
    
    # Clean page margins for expanded breathing room
    for section in doc.sections:
        section.top_margin = Inches(0.55)
        section.bottom_margin = Inches(0.55)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # 1. Header (Logo & Company Details)
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False

    col_widths = [Inches(3.3), Inches(3.7)]
    row = header_table.rows[0]
    row.cells[0].width, row.cells[1].width = col_widths

    left_p = row.cells[0].paragraphs[0]
    left_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    left_run = left_p.add_run()
    if os.path.exists('logo.png'):
        left_run.add_picture('logo.png', width=Inches(2.5))
    else:
        left_run.text = COMPANY_NAME
        left_run.bold = True
        left_run.font.size = Pt(14)
        left_run.font.color.rgb = RGBColor(0, 128, 55)

    right_p = row.cells[1].paragraphs[0]
    right_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r1 = right_p.add_run(f"{COMPANY_NAME.upper()}\n")
    r1.bold = True
    r1.font.size = Pt(10.5)
    r1.font.color.rgb = RGBColor(0, 102, 51)
    r2 = right_p.add_run(f"{COMPANY_ADDRESS}\nPhone: {COMPANY_PHONE}  |  {COMPANY_EMAIL}\nRepresentative: {OFFICER_NAME} – {OFFICER_TITLE}")
    r2.font.size = Pt(9)
    r2.font.color.rgb = RGBColor(80, 80, 80)

    # Decorative dividing rule
    div = doc.add_paragraph()
    div.paragraph_format.space_before = Pt(4)
    div.paragraph_format.space_after = Pt(10)
    div._p.get_or_add_pPr().append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="16" w:space="1" w:color="008037"/></w:pBdr>'))

    # 2. Document Title
    is_progress = (data["waiver_type"] == "Progress Payment")
    title_text = "CONDITIONAL WAIVER AND RELEASE ON PROGRESS PAYMENT" if is_progress else "CONDITIONAL WAIVER AND RELEASE ON FINAL PAYMENT"
    statute_code = "TEXAS PROPERTY CODE § 53.284(b)" if is_progress else "TEXAS PROPERTY CODE § 53.284(d)"

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    t_run = title_p.add_run(title_text)
    t_run.bold = True
    t_run.font.size = Pt(12.5)
    t_run.font.color.rgb = RGBColor(15, 34, 64)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.paragraph_format.space_before = Pt(0)
    sub_p.paragraph_format.space_after = Pt(8)
    s_run = sub_p.add_run(statute_code)
    s_run.bold = True
    s_run.font.size = Pt(9.5)
    s_run.font.color.rgb = RGBColor(100, 100, 100)

    # 3. Clean Notice Box (Expanded 8.5pt font, open line spacing)
    notice_table = doc.add_table(rows=1, cols=1)
    notice_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    n_cell = notice_table.rows[0].cells[0]
    n_cell.width = Inches(7.0)
    set_cell_background(n_cell, "F5F8F5")
    set_cell_margins(n_cell, top=70, bottom=70, left=120, right=120)
    set_cell_border(n_cell, 
                    top=dict(val='single', sz=4, color='B8D8C2'),
                    bottom=dict(val='single', sz=4, color='B8D8C2'),
                    left=dict(val='single', sz=22, color='008037'),
                    right=dict(val='single', sz=4, color='B8D8C2'))
    np = n_cell.paragraphs[0]
    np.paragraph_format.space_before = Pt(0)
    np.paragraph_format.space_after = Pt(0)
    np.paragraph_format.line_spacing = 1.2
    nr1 = np.add_run("NOTICE: ")
    nr1.bold = True
    nr1.font.size = Pt(8.5)
    nr1.font.color.rgb = RGBColor(0, 102, 51)
    nr2 = np.add_run("This document waives and releases lien, stop payment notice, and payment bond rights unconditionally and states that you have been paid for giving up those rights. It is enforceable against you if you sign it, even if you have not been paid. If you have not been paid, use a conditional release form.")
    nr2.font.size = Pt(8.5)
    nr2.font.color.rgb = RGBColor(60, 60, 60)

    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(6)
    sp.paragraph_format.space_after = Pt(0)

    # 4. Details Grid (Comfortable 9.5pt font)
    table_rows = [
        ("Project / Property:", f"{data['property_name']}\n{data['property_address']}"),
        ("Property Owner / Customer:", data['owner_info']),
        ("Claimant / Contractor:", f"{COMPANY_NAME}\n{COMPANY_ADDRESS}  |  {COMPANY_PHONE}"),
        ("Invoice Number(s) & Date:", f"Invoice(s): #{data['invoices']}  |  Date: {data['doc_date']}"),
        ("Payment Amount Claimed:", f"${data['amount']:,.2f} ({format_currency_words(data['amount'])})")
    ]

    details_table = doc.add_table(rows=len(table_rows), cols=2)
    details_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    details_table.autofit = False

    for idx, (label, val) in enumerate(table_rows):
        r = details_table.rows[idx]
        c1, c2 = r.cells[0], r.cells[1]
        c1.width, c2.width = Inches(2.2), Inches(4.8)
        set_cell_background(c1, "F8FAFC")
        set_cell_background(c2, "FFFFFF")
        set_cell_margins(c1, top=60, bottom=60, left=100, right=100)
        set_cell_margins(c2, top=60, bottom=60, left=100, right=100)
        for c in (c1, c2):
            set_cell_border(c, top=dict(sz=4, color="D8DEE4"), bottom=dict(sz=4, color="D8DEE4"), left=dict(sz=4, color="D8DEE4"), right=dict(sz=4, color="D8DEE4"))
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        rn1 = p1.add_run(label)
        rn1.bold = True
        rn1.font.size = Pt(9.5)
        rn1.font.color.rgb = RGBColor(30, 41, 59)
        
        p2 = c2.paragraphs[0]
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        rn2 = p2.add_run(val)
        rn2.font.size = Pt(9.5)
        rn2.font.color.rgb = RGBColor(0, 102, 51) if "Amount" in label else RGBColor(30, 41, 59)
        if "Amount" in label:
            rn2.bold = True

    # 5. Statutory Legal Body (Expanded 9.5pt font with 1.22 line spacing)
    owner_short = data['owner_info'].split('\n')[0]
    p_body1 = doc.add_paragraph()
    p_body1.paragraph_format.space_before = Pt(9)
    p_body1.paragraph_format.space_after = Pt(5)
    p_body1.paragraph_format.line_spacing = 1.22
    rb1 = p_body1.add_run(
        f"On receipt by the signer of this document of a check or electronic funds transfer from "
        f"{owner_short} in the sum of ${data['amount']:,.2f} payable to {COMPANY_NAME}, "
        f"and when the check or electronic payment has been properly endorsed and cleared by the bank on which it is drawn, "
        f"this document becomes effective to release any mechanic's lien, stop payment notice, or any right against a payment bond "
        f"that the signer has on the property referenced above to the following extent:"
    )
    rb1.font.size = Pt(9.5)
    rb1.font.color.rgb = RGBColor(35, 40, 50)

    p_body2 = doc.add_paragraph()
    p_body2.paragraph_format.space_before = Pt(0)
    p_body2.paragraph_format.space_after = Pt(8)
    p_body2.paragraph_format.line_spacing = 1.22

    if is_progress:
        rb2_text = (
            f"This release covers a progress payment for all labor, services, equipment, or materials furnished to the property or to "
            f"{owner_short} as documented under Invoice(s) #{data['invoices']} through the effective date of this document, "
            f"except for unpaid retention, pending modifications, and changes, or items furnished after said date. "
            f"Before any recipient of this document relies on this document, the recipient should verify evidence of payment to the signer."
        )
    else:
        rb2_text = (
            f"This release covers the final payment to the undersigned for all labor, services, equipment, or materials furnished to the property or to "
            f"{owner_short} as documented under Invoice(s) #{data['invoices']}. Before any recipient of this document relies on it, the recipient should verify evidence of payment to the signer."
        )
    rb2 = p_body2.add_run(rb2_text)
    rb2.font.size = Pt(9.5)
    rb2.font.color.rgb = RGBColor(35, 40, 50)

    # 6. Signatures Grid (Open, easily readable blocks)
    include_notary = data["include_notary"]

    if include_notary:
        sig_table = doc.add_table(rows=1, cols=2)
        sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        sig_table.autofit = False
        c_claim, c_notary = sig_table.rows[0].cells[0], sig_table.rows[0].cells[1]
        c_claim.width, c_notary.width = Inches(3.45), Inches(3.55)
        
        for c in (c_claim, c_notary):
            set_cell_background(c, "FAFBFC")
            set_cell_margins(c, top=70, bottom=70, left=100, right=100)
            set_cell_border(c, top=dict(sz=6, color="CFD6DF"), bottom=dict(sz=6, color="CFD6DF"), left=dict(sz=6, color="CFD6DF"), right=dict(sz=6, color="CFD6DF"))
        
        # Claimant Execution Block
        cp = c_claim.paragraphs[0]
        cp.paragraph_format.space_before = Pt(0)
        cp.paragraph_format.space_after = Pt(3)
        cr1 = cp.add_run("CLAIMANT EXECUTION")
        cr1.bold = True
        cr1.font.size = Pt(9.5)
        cr1.font.color.rgb = RGBColor(0, 102, 51)

        cp2 = c_claim.add_paragraph()
        cp2.paragraph_format.space_before = Pt(0)
        cp2.paragraph_format.space_after = Pt(0)
        cp2.paragraph_format.line_spacing = 1.25
        cr2 = cp2.add_run(
            f"Company: {COMPANY_NAME}\n\n"
            f"By: ____________________________________\n"
            f"Name:  {OFFICER_NAME}\n"
            f"Title: {OFFICER_TITLE}\n"
            f"Date:  __________________________________"
        )
        cr2.font.size = Pt(9.5)
        cr2.font.color.rgb = RGBColor(35, 40, 50)

        # Notary Block
        np = c_notary.paragraphs[0]
        np.paragraph_format.space_before = Pt(0)
        np.paragraph_format.space_after = Pt(3)
        nr1 = np.add_run("NOTARY ACKNOWLEDGMENT (TEXAS)")
        nr1.bold = True
        nr1.font.size = Pt(9.5)
        nr1.font.color.rgb = RGBColor(0, 102, 51)

        np2 = c_notary.add_paragraph()
        np2.paragraph_format.space_before = Pt(0)
        np2.paragraph_format.space_after = Pt(0)
        np2.paragraph_format.line_spacing = 1.25
        nr2 = np2.add_run(
            "State of Texas, County of Harris\n\n"
            "Sworn to and subscribed before me on this\n"
            "_____ day of __________________, 20____,\n"
            f"by {OFFICER_NAME}, {OFFICER_TITLE} of {COMPANY_NAME}.\n\n\n"
            "________________________________________\n"
            "Notary Public, State of Texas"
        )
        nr2.font.size = Pt(9.5)
        nr2.font.color.rgb = RGBColor(35, 40, 50)

    else:
        sig_table = doc.add_table(rows=1, cols=1)
        sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        sig_table.autofit = False
        c_claim = sig_table.rows[0].cells[0]
        c_claim.width = Inches(7.0)
        set_cell_background(c_claim, "FAFBFC")
        set_cell_margins(c_claim, top=80, bottom=80, left=120, right=120)
        set_cell_border(c_claim, top=dict(sz=6, color="CFD6DF"), bottom=dict(sz=6, color="CFD6DF"), left=dict(sz=6, color="CFD6DF"), right=dict(sz=6, color="CFD6DF"))
        
        cp = c_claim.paragraphs[0]
        cp.paragraph_format.space_before = Pt(0)
        cp.paragraph_format.space_after = Pt(3)
        cr1 = cp.add_run("CLAIMANT EXECUTION")
        cr1.bold = True
        cr1.font.size = Pt(10)
        cr1.font.color.rgb = RGBColor(0, 102, 51)

        cp2 = c_claim.add_paragraph()
        cp2.paragraph_format.space_before = Pt(0)
        cp2.paragraph_format.space_after = Pt(0)
        cp2.paragraph_format.line_spacing = 1.3
        cr2 = cp2.add_run(
            f"Company: {COMPANY_NAME}\n\n"
            f"By: ___________________________________________________\n"
            f"Name:  {OFFICER_NAME}\n"
            f"Title: {OFFICER_TITLE}\n"
            f"Date:  ___________________________________________________"
        )
        cr2.font.size = Pt(9.5)
        cr2.font.color.rgb = RGBColor(35, 40, 50)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- Streamlit Web UI ---
st.title("📄 Lien Waiver Generator")
st.caption("Texas Property Code Lien Waiver generation for La Salle Landscaping.")

with st.form("waiver_form"):
    st.subheader("1. Payment & Notary Options")
    col1, col2 = st.columns(2)
    with col1:
        waiver_type = st.radio(
            "Payment Structure", 
            options=["Progress Payment", "Typical / Final Payment"],
            help="Select Progress Payment for billing installments or Final Payment for single/completed contracts."
        )
    with col2:
        doc_date = st.date_input("Effective Date", value=date.today())
        include_notary = st.checkbox("Include Notary Acknowledgment Block", value=True)

    st.subheader("2. Invoice Information")
    col3, col4 = st.columns(2)
    with col3:
        invoices = st.text_input("Invoice Number(s)", placeholder="e.g. 1261 or 1226, 1227")
    with col4:
        amount = st.number_input("Payment Amount ($)", min_value=0.01, step=50.0, format="%.2f")

    st.subheader("3. Property & Customer Details")
    prop_name = st.text_input("Project / Property Name", placeholder="e.g. GWR Sabo Road")
    prop_address = st.text_input("Property Address", placeholder="e.g. 2000 W Loop S. Suite #1050, Houston, TX 77027")
    owner_info = st.text_area("Owner / Customer Entity", placeholder="e.g. GWR Sabo Road Owner, LLC\nc/o GWR Management, LLC", height=70)

    submitted = st.form_submit_button("Generate Lien Waiver (.docx)")

if submitted:
    if not invoices or not prop_name or amount <= 0:
        st.error("Please fill in all required fields (Invoice #, Amount, and Property Name).")
    else:
        waiver_data = {
            "waiver_type": waiver_type,
            "include_notary": include_notary,
            "doc_date": doc_date.strftime("%B %d, %Y"),
            "invoices": invoices,
            "amount": amount,
            "property_name": prop_name,
            "property_address": prop_address,
            "owner_info": owner_info
        }
        
        docx_buffer = generate_docx(waiver_data)
        clean_inv = invoices.replace(" ", "").replace(",", "_")
        clean_prop = prop_name.strip().replace(" ", "_")
        filename = f"Lien_Waiver_{clean_prop}_Inv_{clean_inv}.docx"
        
        st.success("Lien waiver generated successfully!")
        st.download_button(
            label=f"⬇️ Download {filename}",
            data=docx_buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
