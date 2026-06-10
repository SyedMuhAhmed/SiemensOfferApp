import streamlit as st
import shutil
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from datetime import datetime
import os
import re

st.set_page_config(page_title="Siemens Offer Letter Generator", page_icon="⚡", layout="wide")

st.markdown("""
<style>
    .main-header { font-size: 2.2rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0.2rem; }
    .section-header { font-size: 1.3rem; font-weight: 600; color: #16213e; margin-top: 1.5rem; margin-bottom: 0.5rem; }
    hr { margin: 1rem 0; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">⚡ Siemens Offer Letter Generator</div>', unsafe_allow_html=True)
st.markdown("---")

# ─────────────────────────────────────────────────────────────
# BASE DIRECTORY
# ─────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ─────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────

def number_to_words(n):
    ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
            "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
            "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]

    def helper(num):
        if num == 0:
            return ""
        elif num < 20:
            return ones[num]
        elif num < 100:
            return tens[num // 10] + (" " + ones[num % 10] if num % 10 != 0 else "")
        elif num < 1000:
            return ones[num // 100] + " Hundred" + (" " + helper(num % 100) if num % 100 != 0 else "")
        elif num < 1_000_000:
            return helper(num // 1000) + " Thousand" + (" " + helper(num % 1000) if num % 1000 != 0 else "")
        elif num < 1_000_000_000:
            return helper(num // 1_000_000) + " Million" + (" " + helper(num % 1_000_000) if num % 1_000_000 != 0 else "")
        else:
            return helper(num // 1_000_000_000) + " Billion" + (" " + helper(num % 1_000_000_000) if num % 1_000_000_000 != 0 else "")

    try:
        clean = str(n).replace(",", "").replace(" ", "").strip()
        num = int(float(clean))
        if num == 0:
            return "Zero"
        return helper(num)
    except:
        return ""

def clear_highlight(run):
    rPr = run._r.get_or_add_rPr()
    for hl in rPr.findall(qn("w:highlight")):
        rPr.remove(hl)

def fill(run, text):
    run.text = text
    run.font.highlight_color = None
    clear_highlight(run)

def all_paras(doc):
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    for elem in doc.element.body:
        tag = elem.tag.split("}")[-1]
        if tag == "p":
            yield Paragraph(elem, doc)
        elif tag == "tbl":
            for row in Table(elem, doc).rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p

def fill_table(table, items):
    for ri, item in enumerate(items):
        while ri + 1 >= len(table.rows):
            table.add_row()
        row = table.rows[ri + 1]
        vals = [item.get("no", str(ri+1)), item.get("desc", ""), item.get("qty", ""), item.get("total", "")]
        for ci, val in enumerate(vals):
            if ci < len(row.cells):
                p = row.cells[ci].paragraphs[0]
                if p.runs:
                    p.runs[0].text = val
                else:
                    r2 = p.add_run(val)
                    r2.font.name = "Arial"
                    r2.font.size = Pt(9)

# ─────────────────────────────────────────────────────────────
# SECTION 1 - OFFER TYPE
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="section-header">📄 Offer Type</div>', unsafe_allow_html=True)
offer_type = st.radio("", ["Firm", "Budgetary"], horizontal=True, label_visibility="collapsed")
is_firm = (offer_type == "Firm")
st.markdown("---")

# ─────────────────────────────────────────────────────────────
# SECTION 2 - CUSTOMER INFORMATION
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="section-header">📋 Customer Information</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    company_options = ["Electro Mechanical Co. LLC (ELMEC)", "ADNOC", "Nofal for Trade & Agencies", "Siemens Energy LLC", "Other"]
    customer_company_sel = st.selectbox("Customer Company", company_options)
    customer_company = st.text_input("Enter Company Name") if customer_company_sel == "Other" else customer_company_sel

with col2:
    country_options = ["UAE", "Saudi Arabia", "Algeria", "Yemen", "Kuwait", "Oman", "Qatar", "Other"]
    country_sel = st.selectbox("Country", country_options)
    country = st.text_input("Enter Country Name") if country_sel == "Other" else country_sel

col3, col4 = st.columns(2)
with col3:
    customer_attn = st.text_input("Contact Person (e.g. Mr. John Smith)")
with col4:
    city_default = "Abu Dhabi" if country in ["UAE", ""] else country
    customer_city_input = st.text_input("City", value=city_default)

col5, col6 = st.columns(2)
with col5:
    customer_po_box = st.text_input("P.O. Box (or leave blank)")
with col6:
    customer_fax = st.text_input("Fax (or leave blank)")

col7, col8 = st.columns(2)
with col7:
    customer_tel = st.text_input("Tel (or leave blank)")
with col8:
    if not is_firm:
        customer_mob = st.text_input("Mobile (or leave blank)")
    else:
        customer_mob = ""

col9, col10 = st.columns(2)
with col9:
    reference_no = st.text_input("Reference No. (or leave blank)")
with col10:
    offer_date = st.text_input("Offer Date (e.g. 15 April 2025)")

st.markdown("---")

# ─────────────────────────────────────────────────────────────
# SECTION 3 - OFFER DETAILS
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="section-header">📝 Offer Details</div>', unsafe_allow_html=True)
subject      = st.text_input("Subject (e.g. 33KV Switchgear Supply)")
project_name = st.text_input("Project Name")
end_user     = st.text_input("End User")
st.markdown("---")

# ─────────────────────────────────────────────────────────────
# SECTION 4 - COMMERCIAL TERMS
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="section-header">💰 Commercial Terms</div>', unsafe_allow_html=True)

col_c1, col_c2 = st.columns(2)
with col_c1:
    currency_options = {
        "EUR - Euro": ("EUR", "Euro"),
        "USD - US Dollar": ("USD", "US Dollar"),
        "GBP - British Pound": ("GBP", "British Pound"),
        "AED - UAE Dirham": ("AED", "UAE Dirham"),
        "SAR - Saudi Riyal": ("SAR", "Saudi Riyal"),
    }
    currency_sel = st.selectbox("Currency", list(currency_options.keys()))
    currency_code, currency_full = currency_options[currency_sel]
with col_c2:
    total_price_num = st.text_input("Total Price (e.g. 3,218,545)")

col_c3, col_c4 = st.columns(2)
with col_c3:
    incoterm_options = ["FCA, Bremerhaven, Germany", "CIF, Abu Dhabi seaport",
                        "DAP, Abu Dhabi", "EXW, Bremerhaven, Germany", "Other"]
    incoterm_sel = st.selectbox("Incoterm", incoterm_options)
    incoterm_name = st.text_input("Enter Incoterm") if incoterm_sel == "Other" else incoterm_sel
with col_c4:
    delivery_options = ["06", "09", "12", "15", "18", "24", "Other"]
    delivery_sel = st.selectbox("Delivery Period (months)", delivery_options)
    delivery_months = st.text_input("Enter months") if delivery_sel == "Other" else delivery_sel

col_c5, col_c6 = st.columns(2)
with col_c5:
    warranty_options = ["12", "18", "24", "Other"]
    warranty_sel = st.selectbox("Warranty Period (months)", warranty_options)
    warranty_months = st.text_input("Enter warranty months") if warranty_sel == "Other" else warranty_sel
with col_c6:
    payment_days_options = ["30", "45", "60", "90", "120", "Other"]
    payment_days_sel = st.selectbox("Payment Days", payment_days_options)
    payment_days = st.text_input("Enter days") if payment_days_sel == "Other" else payment_days_sel

# ── Payment Option ───────────────────────────────────────────
payment_option_options = [
    "Option A - 20% down payment + 80% against shipping documents",
    "Option B - 20% advance + 10% drawings + 20% MFC + 50% bill of lading",
    "Other"
]
payment_option_sel = st.selectbox("Payment Option", payment_option_options)
if payment_option_sel.startswith("Option A"):
    payment_option = "A"
    custom_payment = ""
elif payment_option_sel.startswith("Option B"):
    payment_option = "B"
    custom_payment = ""
else:
    payment_option = "Other"
    custom_payment = st.text_area("Enter custom payment terms")

# ── Cancellation High % (Firm only) ─────────────────────────
if is_firm:
    cancel_high = st.selectbox(
        "Cancellation Charge (high bracket)",
        ["-80%", "-90%"],
        help="Select -80% for National customers, -90% for International / Siemens Energy / Critical Country"
    )
else:
    cancel_high = ""

st.markdown("---")

# ─────────────────────────────────────────────────────────────
# SECTION 5 - FIRM ONLY FIELDS
# ─────────────────────────────────────────────────────────────

if is_firm:
    st.markdown('<div class="section-header">🏗️ Firm Offer Details</div>', unsafe_allow_html=True)

    install_options = {
        "UAE - Abu Dhabi (Dubai or Northern Emirates)": "the Emirate of Abu Dhabi (Dubai or Northern Emirates)",
        "UAE - Dubai / Jebel Ali": "the Emirate of Dubai / Jebel Ali",
        "Saudi Arabia": "a Saudi Arabian seaport",
        "Algeria": "an Algerian seaport",
        "Yemen": "a Yemeni seaport",
        "Other": None
    }
    install_sel = st.selectbox("Country / Port of Installation", list(install_options.keys()))
    import_port = st.text_input("Enter country / port") if install_sel == "Other" else install_options[install_sel]

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        mfc_date = st.text_input("MFC Date (e.g. 26th February 2026)")
    with col_f2:
        offer_validity = st.text_input("Offer Validity (e.g. March 30, 2026)")

    signatory_options = [
        "Robert Hennig + Rodrigo Fernandes",
        "Robert Hennig + Rodrigo Fernandes + Prashanth Parameswaran + Paul Fairweather",
        "Robert Hennig + Rodrigo Fernandes + Prashanth Parameswaran + Paul Fairweather + Nasir Merchant",
        "Other"
    ]
    signatory_sel = st.selectbox("Signatories", signatory_options)
    signatories = st.text_input("Enter signatory names") if signatory_sel == "Other" else signatory_sel
    st.markdown("---")
else:
    import_port    = ""
    mfc_date       = ""
    offer_validity = ""
    signatories    = ""

# ─────────────────────────────────────────────────────────────
# SECTION 6 - SALES CONTACT
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="section-header">👤 Sales Contact</div>', unsafe_allow_html=True)

sales_contacts = {
    "Ahmad Awny | RC-AE SI EA S VD-V-D | +971 55 2003541 | ahmad.awny@siemens.com":
        {"name": "Ahmad Awny", "dept": "RC-AE SI EA S VD-V-D", "mobile": "+971 55 2003541", "email": "ahmad.awny@siemens.com"},
    "Vivek Gopalakrishnan | RC-AE SI EA S | +971 55 2002583 | vivek.gopalakrishnan@siemens.com":
        {"name": "Vivek Gopalakrishnan", "dept": "RC-AE SI EA S", "mobile": "+971 55 2002583", "email": "vivek.gopalakrishnan@siemens.com"},
    "Madiha Khan | RC-AE SI EA S | +971 55 2003818 | madiha.khan@siemens.com":
        {"name": "Madiha Khan", "dept": "RC-AE SI EA S", "mobile": "+971 55 2003818", "email": "madiha.khan@siemens.com"},
    "Hyun-Sik Kim | RC-AE SI EA S VD-V-D | +971 55 2002693 | hyun-sik.kim@siemens.com":
        {"name": "Hyun-Sik Kim", "dept": "RC-AE SI EA S VD-V-D", "mobile": "+971 55 2002693", "email": "hyun-sik.kim@siemens.com"},
    "Other": None
}

sales_sel = st.selectbox("Sales Contact", list(sales_contacts.keys()))
if sales_sel == "Other":
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        sender_name   = st.text_input("Sender Name")
        sender_mobile = st.text_input("Sender Mobile")
    with col_s2:
        sender_dept  = st.text_input("Sender Department")
        sender_email = st.text_input("Sender Email")
else:
    sc = sales_contacts[sales_sel]
    sender_name   = sc["name"]
    sender_dept   = sc["dept"]
    sender_mobile = sc["mobile"]
    sender_email  = sc["email"]

st.markdown("---")

# ─────────────────────────────────────────────────────────────
# SECTION 7 - SCOPE OF SUPPLY
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="section-header">📦 Scope of Supply</div>', unsafe_allow_html=True)

num_scope = st.number_input("Number of scope items", min_value=1, max_value=20, value=1, step=1)
scope_items = []
for i in range(int(num_scope)):
    st.markdown(f"**Item {i+1}**")
    col_sc1, col_sc2, col_sc3 = st.columns([4, 1, 2])
    with col_sc1: desc  = st.text_input("Description", key=f"scope_desc_{i}")
    with col_sc2: qty   = st.text_input("Qty",         key=f"scope_qty_{i}")
    with col_sc3: total = st.text_input("Total Price", key=f"scope_total_{i}")
    scope_items.append({"no": str(i+1), "desc": desc, "qty": qty, "total": total})

st.markdown("**Optional Items**")
num_opt = st.number_input("Number of optional items", min_value=0, max_value=10, value=0, step=1)
optional_items = []
for i in range(int(num_opt)):
    st.markdown(f"**Optional Item {i+1}**")
    col_o1, col_o2, col_o3 = st.columns([4, 1, 2])
    with col_o1: odesc  = st.text_input("Description", key=f"opt_desc_{i}")
    with col_o2: oqty   = st.text_input("Qty",         key=f"opt_qty_{i}")
    with col_o3: ototal = st.text_input("Total Price", key=f"opt_total_{i}")
    optional_items.append({"no": str(i+1), "desc": odesc, "qty": oqty, "total": ototal})

st.markdown("---")

# ─────────────────────────────────────────────────────────────
# GENERATE
# ─────────────────────────────────────────────────────────────

if st.button("🚀 Generate Offer Letter", type="primary", use_container_width=True):

    errors = []
    if not customer_company or customer_company == "Other": errors.append("Customer Company")
    if not customer_attn:   errors.append("Contact Person")
    if not offer_date:      errors.append("Offer Date")
    if not subject:         errors.append("Subject")
    if not project_name:    errors.append("Project Name")
    if not end_user:        errors.append("End User")
    if not total_price_num: errors.append("Total Price")
    if is_firm and not mfc_date:       errors.append("MFC Date")
    if is_firm and not offer_validity: errors.append("Offer Validity")
    if errors:
        st.error(f"Please fill in: {', '.join(errors)}")
        st.stop()

    # ── Derived values ───────────────────────────────────────
    customer_city     = f"{customer_city_input}, {country}"
    total_price_words = number_to_words(total_price_num)

    # Payment text
    if payment_option == "A":
        pay_header = f"Payment terms: {payment_days} days from shipping documents"
        pay_lines  = (
            f"20% of the contract value to be paid as down payment within 15 days "
            f"after placement of the order.\n"
            f"80% of the contract value payable against presentation of shipping "
            f"documents or warehouse receipt within {payment_days} days."
        )
    elif payment_option == "B":
        pay_header = f"Payment terms: {payment_days} days from shipping documents"
        pay_lines  = (
            f"20% of the contract value to be paid as Advance Payment within 15 days "
            f"after Order Confirmation.\n"
            f"10% of the contract value against submission of Drawings (SLD - Single "
            f"Line Diagrams), payable within {payment_days} days.\n"
            f"20% of the contract value upon Manufacturing Clearance, payable within "
            f"45 days.\n"
            f"50% of the contract value against presentation of shipping documents "
            f"(Bill of Lading), payable within {payment_days} days."
        )
    else:
        pay_header = "Payment terms: As per agreement"
        pay_lines  = custom_payment

    # Import port sentence (Firm only)
    import_port_sentence = (
        f"The scope shall be offloaded in and imported via a port of {import_port}."
        if is_firm and import_port else ""
    )

    # ── Build formatted line strings ─────────────────────────
    # P. O. Box line: whole run is the tag, replace entire run text
    po_box_full   = f"P. O. Box {customer_po_box}" if customer_po_box.strip() else ""
    tel_str       = f"Tel : {customer_tel}"         if customer_tel.strip()    else ""
    fax_str       = f"Fax: {customer_fax}"          if customer_fax.strip()    else ""
    mob_str       = f"Mob: {customer_mob}"          if customer_mob.strip()    else ""

    # Price words run value  →  "EUR Three Million ... Only)."
    price_words_run = f"{currency_code} {total_price_words} Only)."

    # ── Replacements map (keys = EXACT tag text in template) ─
    # NOTE: "P. O. Box INSERT_CUSTOMER_PO_BOX_NUM" is the full run text
    replacements = {
        # ── Customer block ──────────────────────────────────
        "INSERT_CUSTOMER_COMPANY":                  customer_company,
        "P. O. Box INSERT_CUSTOMER_PO_BOX_NUM":     po_box_full,
        "INSERT_CUSTOMER_CITY_FULL":                customer_city,
        "INSERT_CUSTOMER_ATTN":                     f"Kind Attn: {customer_attn}",
        "INSERT_CUSTOMER_TEL_LINE":                 tel_str,
        "INSERT_CUSTOMER_FAX_LINE":                 fax_str,
        "INSERT_CUSTOMER_MOB_LINE":                 mob_str,          # Budgetary only
        # ── Sender block ─────────────────────────────────────
        "INSERT_SENDER_NAME":                       sender_name,
        "INSERT_SENDER_DEPT":                       sender_dept,
        "INSERT_SENDER_MOBILE":                     sender_mobile,
        "INSERT_SENDER_EMAIL":                      sender_email,
        # ── Header meta ──────────────────────────────────────
        "INSERT_OFFER_DATE":                        offer_date,
        "INSERT_REFERENCE_NO":                      reference_no.strip(),
        # ── Subject / project ────────────────────────────────
        "INSERT_SUBJECT_FULL":                      subject,
        "INSERT_PROJECT_NAME":                      project_name,
        "INSERT_END_USER":                          end_user,
        # ── Price block ──────────────────────────────────────
        "INSERT_CURRENCY_FULL":                     currency_full,
        "INSERT_TOTAL_PRICE_NUM":                   total_price_num,
        "INSERT_CURRENCY_CODE":                     currency_code,
        # Combined price-words run (same text in both templates, different suffix)
        "INSERT_CURRENCY_CODE INSERT_TOTAL_PRICE_WORDS).":          price_words_run,   # Firm
        "INSERT_CURRENCY_CODE INSERT_TOTAL_PRICE_WORDS Only).":     price_words_run,   # Budgetary
        # ── Commercial terms ─────────────────────────────────
        "INSERT_INCOTERM_NAME":                     incoterm_name,
        "INSERT_DELIVERY_MONTHS":                   f"{delivery_months} month(s)",
        "INSERT_WARRANTY_MONTHS":                   f"{warranty_months} months",
        "INSERT_PAYMENT_OPTION_HEADER":             pay_header,
        "INSERT_PAYMENT_OPTION_LINE":               pay_lines,        # Firm  (singular)
        "INSERT_PAYMENT_OPTION_LINES":              pay_lines,        # Budgetary (plural)
        # ── Firm-only ────────────────────────────────────────
        "INSERT_MFC_DATE":                          mfc_date if is_firm else "",
        "INSERT_IMPORT_PORT_SENTENCE":              import_port_sentence,
        "INSERT_OFFER_VALIDITY":                    offer_validity if is_firm else "",
        "INSERT_CANCEL_HIGH":                       cancel_high,
    }

    # ── File setup ───────────────────────────────────────────
    TEMPLATE_PATH = (
        os.path.join(BASE_DIR, "template_Firm_FIXED.docx")
        if is_firm else
        os.path.join(BASE_DIR, "template_Budgetary_FIXED.docx")
    )

    offer_short = "FIRM" if is_firm else "BUD"
    cust_short  = customer_company.split()[0].upper().strip(".,)")
    SKIP        = {"FOR","THE","OF","AND","IN","A","AN","PKG","WORKS","-","PROJECT"}
    proj_short  = "".join(w for w in project_name.split() if w.upper() not in SKIP)[:12].upper()
    try:    date_str = datetime.strptime(offer_date.strip(), "%d %B %Y").strftime("%Y%m%d")
    except: date_str = offer_date.replace(" ", "")[:8]
    filename = f"{offer_short}_{cust_short}_{proj_short}_{date_str}_v01.docx"
    filepath = os.path.join(BASE_DIR, filename)

    shutil.copy(TEMPLATE_PATH, filepath)
    doc = Document(filepath)

    # ════════════════════════════════════════════════════════
    # PASS 1 — Highlighted runs (standard tags)
    # Strategy:
    #   a) Collect all highlighted runs in the paragraph
    #   b) Merge adjacent highlighted runs whose combined text is a known key
    #   c) Replace any run whose stripped text matches a key exactly
    #   d) For runs that still contain INSERT_ fragments, do regex sub
    # ════════════════════════════════════════════════════════
    for para in all_paras(doc):
        all_runs = para.runs
        # Merge ALL adjacent highlighted runs first (handles split email & others)
        i = 0
        while i < len(all_runs) - 1:
            ri  = all_runs[i]
            ri1 = all_runs[i + 1]
            if ri.font.highlight_color is not None and ri1.font.highlight_color is not None:
                combined = (ri.text or "") + (ri1.text or "")
                # Only merge if combined is a known key OR either piece has INSERT_
                if combined.strip() in replacements or (
                    "INSERT_" in (ri.text or "") and "INSERT_" in (ri1.text or "")
                ):
                    ri.text  = combined
                    ri1.text = ""
                    clear_highlight(ri1)
                    ri1.font.highlight_color = None
                    # Re-check same index (in case of triple split)
                    continue
            i += 1

        # Now replace
        for r in all_runs:
            if not r.text:
                continue
            key = r.text.strip()
            if key in replacements:
                fill(r, replacements[key])
            elif r.font.highlight_color is not None and "INSERT_" in r.text:
                # Regex sub for any remaining INSERT_ tokens inside a run
                new_text = re.sub(
                    r"INSERT_\w+",
                    lambda m: replacements.get(m.group(0), ""),
                    r.text
                )
                fill(r, new_text)

    # ════════════════════════════════════════════════════════
    # PASS 2 — Non-highlighted runs that carry INSERT_ tags
    # Covers: INSERT_INCOTERM_NAME, INSERT_DELIVERY_MONTHS,
    #         INSERT_OFFER_VALIDITY  (no highlight in Firm template)
    #         INSERT_DELIVERY_MONTHS split across runs in Budgetary
    # ════════════════════════════════════════════════════════
    for para in all_paras(doc):
        # Merge adjacent non-highlighted runs that together form a key
        all_runs = para.runs
        i = 0
        while i < len(all_runs) - 1:
            ri  = all_runs[i]
            ri1 = all_runs[i + 1]
            if "INSERT_" in (ri.text or "") and "INSERT_" in (ri1.text or ""):
                combined = (ri.text or "") + (ri1.text or "")
                if combined.strip() in replacements:
                    ri.text  = combined
                    ri1.text = ""
                    continue
            i += 1

        for r in all_runs:
            if not r.text or "INSERT_" not in r.text:
                continue
            key = r.text.strip()
            if key in replacements:
                r.text = replacements[key]
                clear_highlight(r)
            else:
                new_text = re.sub(
                    r"INSERT_\w+",
                    lambda m: replacements.get(m.group(0), ""),
                    r.text
                )
                if new_text != r.text:
                    r.text = new_text
                    clear_highlight(r)

    # ════════════════════════════════════════════════════════
    # PASS 3 — Budgetary: strip the stray "{" before the price line
    # Template has: ": {INSERT_CURRENCY_FULL ..."
    # The "{" is in run[5] with no highlight — just clear it.
    # ════════════════════════════════════════════════════════
    if not is_firm:
        for para in all_paras(doc):
            if "Non-binding" in para.text and "Budgetary" in para.text and "Price" in para.text:
                for r in para.runs:
                    if r.text and r.text.strip() == "{":
                        r.text = ""
                break

    # ════════════════════════════════════════════════════════
    # PASS 4 — Scope tables
    # Table[0] = header/info table (skip)
    # Table[1] = scope of supply
    # Table[2] = optional items (if present)
    # ════════════════════════════════════════════════════════
    if len(doc.tables) > 1:
        fill_table(doc.tables[1], scope_items)
    if optional_items and len(doc.tables) > 2:
        fill_table(doc.tables[2], optional_items)

    # ════════════════════════════════════════════════════════
    # PASS 5 — Final safety sweep
    # Remove any leftover INSERT_ tokens and clear all highlights
    # ════════════════════════════════════════════════════════
    for para in all_paras(doc):
        for r in para.runs:
            if r.text and "INSERT_" in r.text:
                r.text = re.sub(r"INSERT_\w+", "", r.text)
            if r.font.highlight_color is not None:
                clear_highlight(r)
                r.font.highlight_color = None

    doc.save(filepath)

    st.success("✅ Offer letter generated successfully!")
    with open(filepath, "rb") as f:
        st.download_button(
            label=f"📥 Download {filename}",
            data=f,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
