import streamlit as st
import shutil
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from datetime import datetime
import os
import re

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def all_paras(doc):
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    for elem in doc.element.body:
        tag = elem.tag.split('}')[-1]
        if tag == 'p':
            yield Paragraph(elem, doc)
        elif tag == 'tbl':
            for row in Table(elem, doc).rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p

def fill(run, text):
    run.text = text
    run.font.highlight_color = None

def hl_runs(para):
    return [r for r in para.runs if r.font.highlight_color is not None]

def number_to_words(n_str):
    """Convert a numeric string like '3,218,545' to English words."""
    ones = ['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine',
            'Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen',
            'Seventeen','Eighteen','Nineteen']
    tens = ['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety']

    def say(n):
        if n < 20:
            return ones[n]
        elif n < 100:
            return tens[n // 10] + (' ' + ones[n % 10] if n % 10 else '')
        elif n < 1000:
            return ones[n // 100] + ' Hundred' + (' ' + say(n % 100) if n % 100 else '')
        elif n < 1_000_000:
            return say(n // 1000) + ' Thousand' + (' ' + say(n % 1000) if n % 1000 else '')
        elif n < 1_000_000_000:
            return say(n // 1_000_000) + ' Million' + (' ' + say(n % 1_000_000) if n % 1_000_000 else '')
        else:
            return say(n // 1_000_000_000) + ' Billion' + (' ' + say(n % 1_000_000_000) if n % 1_000_000_000 else '')

    try:
        num = int(n_str.replace(',', '').replace('.', '').strip())
        return say(num)
    except:
        return n_str

# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Siemens Offer Generator", page_icon="⚡", layout="centered")
st.title("⚡ Siemens Offer Letter Generator")
st.markdown("---")

# ── OFFER TYPE ──────────────────────────────────────────────────────────────
offer_type = st.radio("Offer Type", ["Firm", "Budgetary"], horizontal=True)
is_firm = offer_type == "Firm"

st.markdown("---")
st.subheader("📋 Customer Information")

# ── CUSTOMER FIELDS ─────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    customer_company_options = [
        "Electro Mechanical Co. LLC (ELMEC)",
        "ADNOC",
        "Nofal for Trade & Agencies",
        "Siemens Energy LLC",
        "Other"
    ]
    customer_company_sel = st.selectbox("Customer Company", customer_company_options)
    if customer_company_sel == "Other":
        customer_company = st.text_input("Enter Company Name")
    else:
        customer_company = customer_company_sel

    # Company display name (without abbreviation for address block)
    company_display = re.sub(r'\s*\(.*?\)', '', customer_company).strip()

    customer_attn = st.text_input("Contact Person (e.g. Mr. John Smith)")
    customer_po_box = st.text_input("P.O. Box Number (or leave blank)")
    customer_tel = st.text_input("Tel (e.g. +971 2 6262 800 Ext:33, or leave blank)")

with col2:
    country_options = ["UAE", "Saudi Arabia", "Algeria", "Yemen", "Kuwait", "Oman", "Qatar", "Other"]
    country_sel = st.selectbox("Country", country_options)
    if country_sel == "Other":
        country = st.text_input("Enter Country")
    else:
        country = country_sel

    city_input = st.text_input("City (e.g. Abu Dhabi)", value="Abu Dhabi")
    customer_city_full = f"{city_input}, {country}" if city_input else country

    customer_fax = st.text_input("Fax (or leave blank)")
    if not is_firm:
        customer_mob = st.text_input("Mobile (or leave blank)")
    else:
        customer_mob = ""

st.markdown("---")
st.subheader("📄 Offer Details")

col3, col4 = st.columns(2)
with col3:
    reference_no = st.text_input("Reference No. (or leave blank)")
    offer_date = st.text_input("Offer Date (e.g. 15 April 2025)")
    subject = st.text_input("Subject (e.g. 33KV Switchgear Supply)")
    project_name = st.text_input("Project Name")

with col4:
    end_user = st.text_input("End User")
    currency_options = {
        "EUR - Euro": ("EUR", "Euro"),
        "USD - US Dollar": ("USD", "US Dollar"),
        "GBP - British Pound": ("GBP", "British Pound"),
        "AED - UAE Dirham": ("AED", "UAE Dirham"),
        "SAR - Saudi Riyal": ("SAR", "Saudi Riyal"),
    }
    currency_sel = st.selectbox("Currency", list(currency_options.keys()))
    currency_code, currency_full = currency_options[currency_sel]
    total_price_num = st.text_input("Total Price (e.g. 3,218,545)")

st.markdown("---")
st.subheader("🚚 Commercial Terms")

col5, col6 = st.columns(2)
with col5:
    incoterm_options = [
        "FCA, Bremerhaven, Germany",
        "CIF, Abu Dhabi seaport",
        "DAP, Abu Dhabi",
        "EXW, Bremerhaven, Germany",
        "Other"
    ]
    incoterm_sel = st.selectbox("Incoterm", incoterm_options)
    if incoterm_sel == "Other":
        incoterm_name = st.text_input("Enter Incoterm")
    else:
        incoterm_name = incoterm_sel

    delivery_options = ["06", "09", "12", "15", "18", "24", "Other"]
    delivery_sel = st.selectbox("Delivery (months)", delivery_options)
    if delivery_sel == "Other":
        delivery_months = st.text_input("Enter delivery months")
    else:
        delivery_months = delivery_sel

    warranty_options = ["12", "18", "24", "Other"]
    warranty_sel = st.selectbox("Warranty (months)", warranty_options)
    if warranty_sel == "Other":
        warranty_months = st.text_input("Enter warranty months")
    else:
        warranty_months = warranty_sel

with col6:
    payment_option = st.radio(
        "Payment Option",
        ["Option A — 20% down + 80% shipping docs",
         "Option B — 20% advance + 10% drawings + 20% MFC + 50% bill of lading"],
        index=0
    )
    payment_option_code = "A" if "Option A" in payment_option else "B"

    payment_days_options = ["30", "45", "60", "90", "120", "Other"]
    payment_days_sel = st.selectbox("Payment Days", payment_days_options, index=3)
    if payment_days_sel == "Other":
        payment_days = st.text_input("Enter payment days")
    else:
        payment_days = payment_days_sel

    # ── CANCELLATION TABLE CHOICE (replaces old National/International) ──────
    cancel_choice = st.radio(
        "Cancellation Table",
        [
            "National  →  -5% / -45% / -80% / -100%",
            "International  →  5% / 45% / 90% / 100%"
        ],
        index=0
    )
    is_national = "National" in cancel_choice

st.markdown("---")
st.subheader("👤 Sender / Sales Contact")

sender_options = {
    "Ahmad Awny": ("Ahmad Awny", "RC-AE SI EA S VD-V-D", "+971 55 2003541", "ahmad.awny@siemens.com"),
    "Vivek Gopalakrishnan": ("Vivek Gopalakrishnan", "RC-AE SI EA S", "+971 55 2002583", "vivek.gopalakrishnan@siemens.com"),
    "Madiha Khan": ("Madiha Khan", "RC-AE SI EA S", "+971 55 2003818", "madiha.khan@siemens.com"),
    "Hyun-Sik Kim": ("Hyun-Sik Kim", "RC-AE SI EA S VD-V-D", "+971 55 2002693", "hyun-sik.kim@siemens.com"),
    "Other": ("", "", "", ""),
}
sender_sel = st.selectbox("Sales Contact", list(sender_options.keys()))
if sender_sel == "Other":
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        sender_name = st.text_input("Sender Name")
        sender_dept = st.text_input("Sender Department")
    with col_s2:
        sender_mobile = st.text_input("Sender Mobile")
        sender_email = st.text_input("Sender Email")
else:
    sender_name, sender_dept, sender_mobile, sender_email = sender_options[sender_sel]
    st.info(f"📧 {sender_name} | {sender_dept} | {sender_mobile} | {sender_email}")

# ── FIRM-ONLY FIELDS ─────────────────────────────────────────────────────────
if is_firm:
    st.markdown("---")
    st.subheader("🏗️ Firm Offer — Additional Details")

    col7, col8 = st.columns(2)
    with col7:
        install_options = {
            "UAE - Abu Dhabi (Dubai or Northern Emirates)": "the Emirate of Abu Dhabi (Dubai or Northern Emirates)",
            "UAE - Dubai / Jebel Ali": "the Emirate of Dubai / Jebel Ali",
            "Saudi Arabia": "a Saudi Arabian seaport",
            "Algeria": "an Algerian seaport",
            "Yemen": "a Yemeni seaport",
            "Other": None,
        }
        install_sel = st.selectbox("Country / Port of Installation", list(install_options.keys()))
        if install_sel == "Other":
            import_port = st.text_input("Enter port / country")
        else:
            import_port = install_options[install_sel]

        mfc_date = st.text_input("MFC Date (e.g. 26th February 2026)")

    with col8:
        offer_validity = st.text_input("Offer Validity (e.g. March 30, 2026)")

        signatory_options = [
            "Robert Hennig + Rodrigo Fernandes",
            "Robert Hennig + Rodrigo Fernandes + Prashanth Parameswaran + Paul Fairweather",
            "Robert Hennig + Rodrigo Fernandes + Prashanth Parameswaran + Paul Fairweather + Nasir Merchant",
            "Other"
        ]
        signatory_sel = st.selectbox("Signatories", signatory_options)
        if signatory_sel == "Other":
            signatories = st.text_input("Enter signatory names")
        else:
            signatories = signatory_sel
else:
    import_port = ""
    mfc_date = ""
    offer_validity = ""
    signatories = ""

# ── SCOPE ITEMS ──────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("📦 Scope of Supply")

st.markdown("Enter each scope item below:")
num_scope = st.number_input("Number of scope items", min_value=1, max_value=20, value=1, step=1)
scope_items = []
for i in range(int(num_scope)):
    with st.expander(f"Scope Item {i+1}", expanded=(i == 0)):
        c1, c2, c3 = st.columns([4, 1, 2])
        with c1:
            desc = st.text_input(f"Description", key=f"scope_desc_{i}")
        with c2:
            qty = st.text_input(f"Qty", key=f"scope_qty_{i}", value="1")
        with c3:
            price = st.text_input(f"Total Price", key=f"scope_price_{i}", placeholder=f"{currency_code} 0")
        scope_items.append({"no": str(i+1), "desc": desc, "qty": qty, "total": price})

st.markdown("**Optional Items** (leave blank if none)")
num_opt = st.number_input("Number of optional items", min_value=0, max_value=10, value=0, step=1)
optional_items = []
for i in range(int(num_opt)):
    with st.expander(f"Optional Item {i+1}"):
        c1, c2, c3 = st.columns([4, 1, 2])
        with c1:
            desc = st.text_input(f"Description", key=f"opt_desc_{i}")
        with c2:
            qty = st.text_input(f"Qty", key=f"opt_qty_{i}", value="1")
        with c3:
            price = st.text_input(f"Total Price", key=f"opt_price_{i}", placeholder=f"{currency_code} 0")
        optional_items.append({"no": str(i+1), "desc": desc, "qty": qty, "total": price})

# ─────────────────────────────────────────────────────────────────────────────
# GENERATE BUTTON
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
if st.button("🚀 Generate Offer Letter", type="primary", use_container_width=True):

    # ── Auto-generate price in words
    total_price_words = number_to_words(total_price_num) + f" {currency_full} Only"

    # ── Payment option text
    if payment_option_code == "A":
        pay_header = f"Payment terms: {payment_days} days from shipping documents"
        pay_lines = (
            f"20% of the contract value to be paid as down payment within 15 days after placement of the order.\n"
            f"80% of the contract value payable against presentation of shipping documents within {payment_days} days."
        )
    else:
        pay_header = f"Payment terms: {payment_days} days from shipping documents"
        pay_lines = (
            f"20% of the contract value to be paid as Advance Payment within 15 days after Order Confirmation.\n"
            f"10% of the contract value against submission of Drawings (SLD), payable within {payment_days} days.\n"
            f"20% of the contract value upon Manufacturing Clearance, payable within 45 days.\n"
            f"50% of the contract value against Bill of Lading, payable within {payment_days} days."
        )

    import_port_sentence = (
        f"The scope shall be offloaded in and imported via a port of {import_port}."
        if is_firm and import_port else ""
    )

    # ── Tel/Fax/Mob formatted lines
    tel_line  = f"Tel : {customer_tel}"  if customer_tel  else ""
    fax_line  = f"Fax: {customer_fax}"  if customer_fax  else ""
    mob_line  = f"Mob: {customer_mob}"  if customer_mob  else ""

    # ── Cancellation high value
    cancel_high = "-80%" if is_national else "90%"

    # ── Replacements map (using new fixed placeholder names)
    replacements = {
        "INSERT_CUSTOMER_COMPANY":      company_display,
        "INSERT_CUSTOMER_PO_BOX_NUM":   customer_po_box,
        "INSERT_CUSTOMER_CITY_FULL":    customer_city_full,
        "INSERT_CUSTOMER_ATTN":         f"Kind Attn: {customer_attn}",
        "INSERT_CUSTOMER_TEL_LINE":     tel_line,
        "INSERT_CUSTOMER_FAX_LINE":     fax_line,
        "INSERT_CUSTOMER_MOB_LINE":     mob_line,
        "INSERT_SENDER_NAME":           sender_name,
        "INSERT_SENDER_DEPT":           sender_dept,
        "INSERT_SENDER_MOBILE":         sender_mobile,
        "INSERT_SENDER_EMAIL":          sender_email,
        "INSERT_OFFER_DATE":            offer_date,
        "INSERT_REFERENCE_NO":          reference_no,
        "INSERT_SUBJECT_FULL":          subject,
        "INSERT_PROJECT_NAME":          project_name,
        "INSERT_END_USER":              end_user,
        "INSERT_CURRENCY_FULL":         currency_full,
        "INSERT_TOTAL_PRICE_NUM":       total_price_num,
        "INSERT_CURRENCY_CODE":         currency_code,
        "INSERT_TOTAL_PRICE_WORDS":     f"{currency_code} {total_price_words}).",
        "INSERT_INCOTERM_NAME":         incoterm_name,
        "INSERT_DELIVERY_MONTHS":       f"{delivery_months} month(s)",
        "INSERT_WARRANTY_MONTHS":       f"{warranty_months} months",
        "INSERT_PAYMENT_DAYS":          payment_days,
        "INSERT_PAYMENT_OPTION_HEADER": pay_header,
        "INSERT_PAYMENT_OPTION_LINES":  pay_lines,
        "INSERT_MFC_DATE":              mfc_date if is_firm else "",
        "INSERT_IMPORT_PORT_SENTENCE":  import_port_sentence,
        "INSERT_OFFER_VALIDITY":        offer_validity if is_firm else "",
        "INSERT_CANCEL_HIGH":           cancel_high,
    }

    # ── Open correct fixed template
    TEMPLATE = "template_Firm_FIXED.docx" if is_firm else "template_Budgetary_FIXED.docx"
    if not os.path.exists(TEMPLATE):
        st.error(f"Template not found: {TEMPLATE}")
        st.stop()

    offer_short = "FIRM" if is_firm else "BUD"
    cust_short  = company_display.split()[0].upper().strip(".,")
    SKIP = {"FOR","THE","OF","AND","IN","A","AN","PKG","WORKS","-","PROJECT"}
    proj_short  = "".join(w for w in project_name.split() if w.upper() not in SKIP)[:12].upper()
    try:
        date_str = datetime.strptime(offer_date.strip(), "%d %B %Y").strftime("%Y%m%d")
    except:
        date_str = offer_date.replace(" ", "")[:8]

    filename = f"{offer_short}_{cust_short}_{proj_short}_{date_str}_v01.docx"
    shutil.copy(TEMPLATE, filename)
    doc = Document(filename)

    # ── Apply replacements
    for para in all_paras(doc):
        hrs = hl_runs(para)
        if not hrs:
            continue
        # Merge split runs (e.g. INSERT_SENDER_EMAIL split across 2 runs)
        for i in range(len(hrs) - 1):
            combined = (hrs[i].text or "") + (hrs[i+1].text or "")
            if combined.strip() in replacements:
                hrs[i].text = combined
                hrs[i+1].text = ""
                hrs[i+1].font.highlight_color = None
        for r in hrs:
            key = (r.text or "").strip()
            if key in replacements:
                fill(r, replacements[key])
            # Handle compound run: "P. O. Box INSERT_CUSTOMER_PO_BOX_NUM"
            if "INSERT_CUSTOMER_PO_BOX_NUM" in (r.text or ""):
                r.text = r.text.replace("INSERT_CUSTOMER_PO_BOX_NUM", customer_po_box)
                r.font.highlight_color = None

    # ── Firm-only style patches
    if is_firm:
        for para in all_paras(doc):
            if para.text.strip().startswith("The offer currency is"):
                for r in para.runs:
                    if r.font.highlight_color is not None and not r.bold:
                        fill(r, currency_full)
                    if r.font.highlight_color is not None and r.bold and not r.underline:
                        fill(r, mfc_date)
                    if r.bold and r.underline and r.font.highlight_color is None:
                        r.text = incoterm_name
                    if r.text in ("{", "}"):
                        r.text = ""
                break

        for para in all_paras(doc):
            if para.text.strip().startswith("The delivery period of the offered equipment is estimated to be"):
                for r in para.runs:
                    if r.bold and r.font.highlight_color is None:
                        if r.text and (r.text.strip()[:2].isdigit() or "month" in r.text):
                            r.text = f"{delivery_months} month(s)"
                            break
                break

        for para in all_paras(doc):
            if "valid until" in para.text:
                replaced = False
                for r in para.runs:
                    if r.bold and r.underline and r.font.highlight_color is None and not replaced:
                        r.text = offer_validity
                        replaced = True
                    elif r.bold and r.underline and r.font.highlight_color is None and replaced:
                        r.text = ""
                break

    # ── Scope tables
    def fill_table(table, items):
        for ri, item in enumerate(items):
            while ri + 1 >= len(table.rows):
                table.add_row()
            row = table.rows[ri + 1]
            for ci, val in enumerate([item.get("no", str(ri+1)), item.get("desc", ""),
                                       item.get("qty", ""), item.get("total", "")]):
                if ci < len(row.cells):
                    p = row.cells[ci].paragraphs[0]
                    if p.runs:
                        p.runs[0].text = val
                    else:
                        r2 = p.add_run(val)
                        r2.font.name = "Arial"
                        r2.font.size = Pt(9)

    if len(doc.tables) > 1:
        fill_table(doc.tables[1], scope_items)
    if optional_items and len(doc.tables) > 2:
        fill_table(doc.tables[2], optional_items)

    # ── Post-generation: clear any remaining INSERT_ tokens
    for para in all_paras(doc):
        for r in para.runs:
            if r.text and "INSERT_" in r.text:
                r.text = re.sub(r'INSERT_\w+', '', r.text)
                r.font.highlight_color = None

    doc.save(filename)

    with open(filename, "rb") as f:
        st.download_button(
            label="📥 Download Offer Letter",
            data=f,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    st.success(f"✅ Offer letter generated: {filename}")
