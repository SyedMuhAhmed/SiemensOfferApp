
import streamlit as st
import shutil
import re
import os
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.shared import Pt
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Siemens Offer Generator", page_icon="⚡", layout="wide")

st.markdown("""
<style>
    .main { background-color: #f5f5f5; }
    .stButton>button {
        background-color: #009999;
        color: white;
        border-radius: 4px;
        font-weight: bold;
        padding: 0.5rem 2rem;
    }
    .stButton>button:hover { background-color: #007777; }
    h1, h2, h3 { color: #009999; }
    .section-header {
        background-color: #009999;
        color: white;
        padding: 8px 16px;
        border-radius: 4px;
        margin: 16px 0 8px 0;
        font-weight: bold;
    }
    .stSelectbox label, .stTextInput label, .stRadio label { color: #333; font-weight: 500; }
    .stProgress > div > div { background-color: #009999; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# TEMPLATE PATHS
# ─────────────────────────────────────────────────────────────────────────────
FIRM_TEMPLATE      = "/tmp/tmeplate_-Firm.docx"
BUDGETARY_TEMPLATE = "/tmp/tempate_-Budgetary.docx"

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def all_paras(doc):
    for elem in doc.element.body:
        tag = elem.tag.split("}")[-1]
        if tag == "p":
            yield Paragraph(elem, doc)
        elif tag == "tbl":
            for row in Table(elem, doc).rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p

def number_to_words(n_str):
    """Convert a numeric string (with commas) to English words."""
    ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven",
            "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen",
            "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty",
            "Sixty", "Seventy", "Eighty", "Ninety"]

    def below_1000(n):
        if n == 0:
            return ""
        elif n < 20:
            return ones[n]
        elif n < 100:
            return tens[n // 10] + (" " + ones[n % 10] if n % 10 else "")
        else:
            rest = below_1000(n % 100)
            return ones[n // 100] + " Hundred" + (" " + rest if rest else "")

    try:
        n = int(n_str.replace(",", "").replace(".", "").strip())
    except Exception:
        return n_str

    if n == 0:
        return "Zero"

    parts = []
    billions  = n // 1_000_000_000
    millions  = (n % 1_000_000_000) // 1_000_000
    thousands = (n % 1_000_000) // 1_000
    remainder = n % 1_000

    if billions:
        parts.append(below_1000(billions) + " Billion")
    if millions:
        parts.append(below_1000(millions) + " Million")
    if thousands:
        parts.append(below_1000(thousands) + " Thousand")
    if remainder:
        parts.append(below_1000(remainder))

    return " ".join(parts)

def build_replacements(d):
    """Build the INSERT_ → value mapping from collected data dict."""
    is_firm = (d["OFFER_TYPE"] == "Firm")

    # Payment text
    pay_days = d.get("PAYMENT_DAYS", "")
    if d.get("PAYMENT_OPTION") == "A":
        pay_header = "Payment terms: %s days from shipping documents" % pay_days
        pay_lines  = (
            "20%% of the contract value to be paid as down payment within 15 days "
            "after placement of the order.\n"
            "80%% of the contract value payable against presentation of shipping "
            "documents within %s days." % pay_days
        )
    else:
        pay_header = "Payment terms: %s days from shipping documents" % pay_days
        pay_lines  = (
            "20%% of the contract value to be paid as Advance Payment within 15 days "
            "after Order Confirmation.\n"
            "10%% of the contract value against submission of Drawings (SLD - Single "
            "Line Diagrams), payable within %s days.\n"
            "20%% of the contract value upon Manufacturing Clearance, payable within "
            "45 days.\n"
            "50%% of the contract value against presentation of shipping documents "
            "(Bill of Lading), payable within %s days." % (pay_days, pay_days)
        )

    import_port_sentence = ""
    if is_firm:
        import_port_sentence = (
            "The scope shall be offloaded in and imported via a port of %s."
            % d.get("IMPORT_PORT", "")
        )

    currency_code  = d.get("CURRENCY_CODE", "")
    currency_full  = d.get("CURRENCY_FULL", "")
    price_num      = d.get("TOTAL_PRICE_NUM", "")
    price_words    = d.get("TOTAL_PRICE_WORDS", "")

    return {
        "INSERT_CUSTOMER_COMPANY":      d.get("CUSTOMER_COMPANY", ""),
        "INSERT_CUSTOMER_PO_BOX":       d.get("CUSTOMER_PO_BOX", ""),
        "INSERT_CUSTOMER_CITY":         d.get("CUSTOMER_CITY", ""),
        "INSERT_CUSTOMER_ATTN":         "Kind Attn: " + d.get("CUSTOMER_ATTN", ""),
        "INSERT_CUSTOMER_TEL":          d.get("CUSTOMER_TEL", ""),
        "INSERT_CUSTOMER_FAX":          d.get("CUSTOMER_FAX", ""),
        "INSERT_CUSTOMER_MOB":          d.get("CUSTOMER_MOB", ""),
        "INSERT_SENDER_NAME":           d.get("SENDER_NAME", ""),
        "INSERT_SENDER_DEPT":           d.get("SENDER_DEPT", ""),
        "INSERT_SENDER_MOBILE":         d.get("SENDER_MOBILE", ""),
        "INSERT_SENDER_EMAIL":          d.get("SENDER_EMAIL", ""),
        "INSERT_OFFER_DATE":            d.get("OFFER_DATE", ""),
        "INSERT_REFERENCE_NO":          d.get("REFERENCE_NO", ""),
        "INSERT_SUBJECT":               d.get("SUBJECT", ""),
        "INSERT_PROJECT_NAME":          d.get("PROJECT_NAME", ""),
        "INSERT_END_USER":              d.get("END_USER", ""),
        "INSERT_CURRENCY_FULL":         currency_full,
        "INSERT_TOTAL_PRICE_NUM":       price_num,
        "INSERT_CURRENCY_CODE":         currency_code,
        # Template already has opening "(" before this token, so we close with "Only)."
        "INSERT_CURRENCY_CODE INSERT_TOTAL_PRICE_WORDS Only).": (
            "%s %s Only)." % (currency_code, price_words)
        ),
        "INSERT_INCOTERM_NAME":         d.get("INCOTERM_NAME", ""),
        "INSERT_DELIVERY_MONTHS":       "%s month(s)" % d.get("DELIVERY_MONTHS", ""),
        "INSERT_WARRANTY_MONTHS":       "%s months" % d.get("WARRANTY_MONTHS", ""),
        "INSERT_PAYMENT_DAYS":          pay_days,
        "INSERT_PAYMENT_OPTION_HEADER": pay_header,
        "INSERT_PAYMENT_OPTION_LINES":  pay_lines,
        "INSERT_MFC_DATE":              d.get("MFC_DATE", "") if is_firm else "",
        "INSERT_IMPORT_PORT_SENTENCE":  import_port_sentence,
        "INSERT_OFFER_VALIDITY":        d.get("OFFER_VALIDITY", "") if is_firm else "",
    }

def sub_tokens(text, sorted_keys, replacements):
    for k in sorted_keys:
        if k in text:
            text = text.replace(k, replacements[k])
    return text

def generate_document(d):
    """Fill template and return (filepath, filename)."""
    is_firm     = (d["OFFER_TYPE"] == "Firm")
    is_national = (d.get("CUSTOMER_TYPE", "") == "National")

    replacements = build_replacements(d)
    # Sort longest key first so compound keys match before sub-keys
    sorted_keys = sorted(replacements.keys(), key=len, reverse=True)

    def sub(text):
        return sub_tokens(text, sorted_keys, replacements)

    # ── Output filename ──────────────────────────────────────────────────────
    offer_short = "FIRM" if is_firm else "BUD"
    cust_short  = (d.get("CUSTOMER_COMPANY", "CUST").split()[0]
                   .upper().strip(".,()"))
    SKIP = {"FOR", "THE", "OF", "AND", "IN", "A", "AN", "PKG", "WORKS", "-", "PROJECT"}
    proj_short  = "".join(
        w for w in d.get("PROJECT_NAME", "PROJ").split()
        if w.upper() not in SKIP
    )[:12].upper()
    raw_date = d.get("OFFER_DATE", "")
    try:
        date_str = datetime.strptime(raw_date.strip(), "%d %B %Y").strftime("%Y%m%d")
    except Exception:
        date_str = raw_date.replace(" ", "")[:8]
    filename = "%s_%s_%s_%s_v01.docx" % (offer_short, cust_short, proj_short, date_str)
    filepath = "/tmp/%s" % filename

    # ── Copy template ────────────────────────────────────────────────────────
    template = FIRM_TEMPLATE if is_firm else BUDGETARY_TEMPLATE
    shutil.copy(template, filepath)
    doc = Document(filepath)

    # ── PASS 1: Replace all INSERT_ tokens ──────────────────────────────────
    for para in all_paras(doc):
        runs = para.runs
        n = len(runs)

        # Step A: merge consecutive highlighted runs, then replace
        i = 0
        while i < n:
            if runs[i].font.highlight_color is None:
                i += 1
                continue
            j = i + 1
            while j < n and runs[j].font.highlight_color is not None:
                j += 1
            combined = "".join(r.text or "" for r in runs[i:j])
            replaced = sub(combined)
            runs[i].text = replaced
            runs[i].font.highlight_color = None
            for x in range(i + 1, j):
                runs[x].text = ""
                runs[x].font.highlight_color = None
            i = j

        # Step B: non-highlighted runs that still contain INSERT_
        for r in runs:
            if r.font.highlight_color is None and r.text and "INSERT_" in r.text:
                r.text = sub(r.text)

    # ── PASS 2: Cancellation table patch ────────────────────────────────────
    if is_firm and is_national:
        national_map = {"5%": "-5%", "45%": "-45%", "90%": "-80%", "100%": "-100%"}
        for para in all_paras(doc):
            t = para.text.strip()
            if t in national_map:
                for r in para.runs:
                    if r.text.strip() in national_map:
                        r.text = r.text.replace(r.text.strip(), national_map[r.text.strip()])

    # ── PASS 3: Signatories table (Firm only) ────────────────────────────────
    if is_firm:
        sig = d.get("SIGNATORIES", "a")
        # Table[4] = 2-signatory (Robert + Rodrigo)
        # Table[5] = 4/5-signatory (Prashanth + Paul + Nasir + Rodrigo)
        # Default template already has Robert+Rodrigo in Table[4]
        # and Prashanth+Paul+Nasir+Rodrigo in Table[5]
        # We hide/show tables by clearing content based on selection
        if sig == "a":
            # Keep Table[4] (2 sigs), clear Table[5]
            for row in doc.tables[5].rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for r in p.runs:
                            r.text = ""
        elif sig == "b":
            # Keep Table[4] and Table[5] rows 0 only (4 sigs: Robert+Rodrigo+Prashanth+Paul)
            # Clear Nasir Merchant from Table[5] row[1] col[1]
            row1 = doc.tables[5].rows[1]
            for cell in [row1.cells[1]]:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.text = ""
        # sig == "c": keep all (5 sigs) — default template already has all 5
        # sig == "d": custom — already filled via replacements if user typed names

    # ── PASS 4: Scope table ──────────────────────────────────────────────────
    scope_items    = d.get("SCOPE_ITEMS", [])
    optional_items = d.get("OPTIONAL_ITEMS", [])

    def fill_scope_table(table, items):
        # Find first data row (row index 1)
        header_rows = 1
        # Add rows if needed
        while len(table.rows) < header_rows + len(items):
            table.add_row()
        for ri, item in enumerate(items):
            row = table.rows[header_rows + ri]
            vals = [
                item.get("no", str(ri + 1)),
                item.get("desc", ""),
                item.get("qty", ""),
                item.get("total", ""),
            ]
            for ci, val in enumerate(vals):
                if ci < len(row.cells):
                    p = row.cells[ci].paragraphs[0]
                    if p.runs:
                        p.runs[0].text = val
                    else:
                        r2 = p.add_run(val)
                        r2.font.name = "Arial"
                        r2.font.size = Pt(9)

    if len(doc.tables) > 1:
        fill_scope_table(doc.tables[1], scope_items)
    if optional_items and len(doc.tables) > 2:
        fill_scope_table(doc.tables[2], optional_items)

    # ── PASS 5: Post-scan — catch any remaining INSERT_ tokens ───────────────
    for para in all_paras(doc):
        for r in para.runs:
            if r.text and "INSERT_" in r.text:
                r.text = sub(r.text)
                if "INSERT_" in r.text:
                    r.text = re.sub(r"INSERT_\w+", "", r.text)

    doc.save(filepath)
    return filepath, filename

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────────────────────────────────────
if "step" not in st.session_state:
    st.session_state.step = 1          # 1=offer type, 2=all questions, 3=confirm, 4=done
if "data" not in st.session_state:
    st.session_state.data = {}
if "generated_file" not in st.session_state:
    st.session_state.generated_file = None
if "generated_name" not in st.session_state:
    st.session_state.generated_name = None

# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────
col_logo, col_title = st.columns([1, 5])
with col_logo:
    st.markdown("## ⚡")
with col_title:
    st.markdown("# Siemens Offer Letter Generator")
    st.markdown("*Smart Infrastructure — UAE*")

st.divider()

# Progress bar
progress_map = {1: 0.1, 2: 0.4, 3: 0.8, 4: 1.0}
st.progress(progress_map.get(st.session_state.step, 0.1))

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — OFFER TYPE
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.step == 1:
    st.markdown("### Step 1 of 4 — Select Offer Type")
    offer_type = st.radio(
        "What type of offer would you like to generate?",
        options=["Firm", "Budgetary"],
        horizontal=True
    )
    if st.button("Next →"):
        st.session_state.data["OFFER_TYPE"] = offer_type
        # Set N/A fields based on offer type
        if offer_type == "Budgetary":
            st.session_state.data["MFC_DATE"]       = "N/A"
            st.session_state.data["OFFER_VALIDITY"]  = "N/A"
            st.session_state.data["SIGNATORIES"]     = "N/A"
            st.session_state.data["CUSTOMER_MOB"]    = ""
            st.session_state.data["IMPORT_PORT"]     = "N/A"
        else:
            st.session_state.data["CUSTOMER_MOB"]    = ""
        st.session_state.step = 2
        st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — ALL QUESTIONS
# ─────────────────────────────────────────────────────────────────────────────
elif st.session_state.step == 2:
    d = st.session_state.data
    is_firm = (d.get("OFFER_TYPE") == "Firm")

    st.markdown("### Step 2 of 4 — Offer Details")
    st.info("Fill in all fields below, then click **Generate Offer** at the bottom.")

    # ── CUSTOMER INFORMATION ─────────────────────────────────────────────────
    st.markdown('<div class="section-header">📋 Customer Information</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        country_opts = ["UAE", "Saudi Arabia", "Algeria", "Yemen", "Kuwait", "Oman", "Qatar", "Other"]
        country = st.selectbox("Country *", country_opts,
                               index=country_opts.index(d.get("COUNTRY", "UAE")) if d.get("COUNTRY") in country_opts else 0)
        if country == "Other":
            country = st.text_input("Specify country *", value=d.get("COUNTRY", "") if d.get("COUNTRY") not in country_opts else "")

    with col2:
        ctype_opts = ["National (UAE-based customer)", "International (non-UAE customer)",
                      "Siemens Energy", "Critical Country"]
        ctype = st.selectbox("Customer Type *", ctype_opts,
                             index=ctype_opts.index(d.get("CUSTOMER_TYPE_LABEL", ctype_opts[0]))
                             if d.get("CUSTOMER_TYPE_LABEL") in ctype_opts else 0)

    col1, col2 = st.columns(2)
    with col1:
        company_opts = [
            "Electro Mechanical Co. LLC (ELMEC)",
            "ADNOC",
            "Nofal for Trade & Agencies",
            "Siemens Energy LLC",
            "Other"
        ]
        company = st.selectbox("Customer Company *", company_opts,
                               index=company_opts.index(d.get("CUSTOMER_COMPANY", company_opts[0]))
                               if d.get("CUSTOMER_COMPANY") in company_opts else 0)
        if company == "Other":
            company = st.text_input("Specify company name *",
                                    value=d.get("CUSTOMER_COMPANY", "") if d.get("CUSTOMER_COMPANY") not in company_opts else "")

    with col2:
        attn = st.text_input("Contact Person (Attn) *", value=d.get("CUSTOMER_ATTN", ""),
                             placeholder="e.g. Mr. John Smith")

    col1, col2, col3 = st.columns(3)
    with col1:
        po_box = st.text_input("P.O. Box", value=d.get("CUSTOMER_PO_BOX", ""),
                               placeholder="e.g. P. O. Box 732 (or leave blank)")
    with col2:
        tel = st.text_input("Telephone", value=d.get("CUSTOMER_TEL", ""),
                            placeholder="e.g. Tel: +971 2 6262 800")
    with col3:
        fax = st.text_input("Fax", value=d.get("CUSTOMER_FAX", ""),
                            placeholder="e.g. Fax: +971 2 6269 871")

    if not is_firm:
        mob = st.text_input("Mobile", value=d.get("CUSTOMER_MOB", ""),
                            placeholder="e.g. Mob: +971 50 1234567 (or leave blank)")
    else:
        mob = ""

    # ── OFFER META ───────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">📄 Offer Details</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        ref_no = st.text_input("Reference No.", value=d.get("REFERENCE_NO", ""),
                               placeholder="e.g. O-U276-AE-22-0102 (or leave blank)")
    with col2:
        offer_date = st.text_input("Offer Date *", value=d.get("OFFER_DATE", ""),
                                   placeholder="e.g. 15 April 2025")

    subject = st.text_input("Subject *", value=d.get("SUBJECT", ""),
                            placeholder="e.g. 33KV Switchgear Supply")
    project = st.text_input("Project Name *", value=d.get("PROJECT_NAME", ""),
                            placeholder="e.g. ADNOC Habshan Gas Compressor Plant")
    end_user = st.text_input("End User *", value=d.get("END_USER", ""),
                             placeholder="e.g. ADNOC Gas Processing")

    # ── COMMERCIAL TERMS ─────────────────────────────────────────────────────
    st.markdown('<div class="section-header">💰 Commercial Terms</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        currency_opts = ["EUR - Euro", "USD - US Dollar", "GBP - British Pound",
                         "AED - UAE Dirham", "SAR - Saudi Riyal"]
        currency_label = d.get("CURRENCY_LABEL", currency_opts[0])
        currency_sel = st.selectbox("Currency *", currency_opts,
                                    index=currency_opts.index(currency_label)
                                    if currency_label in currency_opts else 0)
    with col2:
        price_num = st.text_input("Total Price (numeric) *", value=d.get("TOTAL_PRICE_NUM", ""),
                                  placeholder="e.g. 3,218,545")

    col1, col2 = st.columns(2)
    with col1:
        inco_opts = [
            "FCA, Bremerhaven, Germany",
            "CIF, Abu Dhabi seaport",
            "DAP, Abu Dhabi",
            "EXW, Bremerhaven, Germany",
            "Other"
        ]
        inco_val = d.get("INCOTERM_NAME", inco_opts[0])
        inco_sel = st.selectbox("Incoterm *", inco_opts,
                                index=inco_opts.index(inco_val) if inco_val in inco_opts else 0)
        if inco_sel == "Other":
            inco_sel = st.text_input("Specify Incoterm *",
                                     value=inco_val if inco_val not in inco_opts else "")
    with col2:
        delivery_opts = ["06", "09", "12", "15", "18", "24", "Other"]
        del_val = d.get("DELIVERY_MONTHS", "09")
        del_sel = st.selectbox("Delivery (months) *", delivery_opts,
                               index=delivery_opts.index(del_val) if del_val in delivery_opts else 0)
        if del_sel == "Other":
            del_sel = st.text_input("Specify delivery months *",
                                    value=del_val if del_val not in delivery_opts else "")

    col1, col2 = st.columns(2)
    with col1:
        warranty_opts = ["12", "18", "24", "Other"]
        war_val = d.get("WARRANTY_MONTHS", "12")
        war_sel = st.selectbox("Warranty (months) *", warranty_opts,
                               index=warranty_opts.index(war_val) if war_val in warranty_opts else 0)
        if war_sel == "Other":
            war_sel = st.text_input("Specify warranty months *",
                                    value=war_val if war_val not in warranty_opts else "")
    with col2:
        pay_opt_sel = st.radio(
            "Payment Option *",
            options=[
                "Option A — 20% down + 80% shipping docs",
                "Option B — 20% advance + 10% drawings + 20% MFC + 50% bill of lading"
            ],
            index=0 if d.get("PAYMENT_OPTION", "A") == "A" else 1,
            horizontal=False
        )

    col1, col2 = st.columns(2)
    with col1:
        pay_days_opts = ["30", "45", "60", "90", "120", "Other"]
        pay_days_val = d.get("PAYMENT_DAYS", "90")
        pay_days_sel = st.selectbox("Payment Days *", pay_days_opts,
                                    index=pay_days_opts.index(pay_days_val)
                                    if pay_days_val in pay_days_opts else 0)
        if pay_days_sel == "Other":
            pay_days_sel = st.text_input("Specify payment days *",
                                         value=pay_days_val if pay_days_val not in pay_days_opts else "")

    # ── FIRM-ONLY FIELDS ─────────────────────────────────────────────────────
    if is_firm:
        st.markdown('<div class="section-header">🏗️ Firm Offer — Additional Details</div>', unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        with col1:
            install_opts = [
                "UAE - Abu Dhabi (Dubai or Northern Emirates)",
                "UAE - Dubai / Jebel Ali",
                "Saudi Arabia",
                "Algeria",
                "Yemen",
                "Other"
            ]
            install_val = d.get("COUNTRY_OF_INSTALL", install_opts[0])
            install_sel = st.selectbox("Country / Port of Installation *", install_opts,
                                       index=install_opts.index(install_val)
                                       if install_val in install_opts else 0)
            if install_sel == "Other":
                install_sel = st.text_input("Specify installation country/port *",
                                            value=install_val if install_val not in install_opts else "")
        with col2:
            mfc_date = st.text_input("MFC Date *", value=d.get("MFC_DATE", ""),
                                     placeholder="e.g. 26th February 2026")

        offer_validity = st.text_input("Offer Validity *", value=d.get("OFFER_VALIDITY", ""),
                                       placeholder="e.g. March 30, 2026")

        sig_opts = [
            "a — Robert Hennig + Rodrigo Fernandes",
            "b — Robert Hennig + Rodrigo Fernandes + Prashanth Parameswaran + Paul Fairweather",
            "c — Robert Hennig + Rodrigo Fernandes + Prashanth Parameswaran + Paul Fairweather + Nasir Merchant",
        ]
        sig_val = d.get("SIGNATORIES", "a")
        sig_sel = st.selectbox("Signatories *", sig_opts,
                               index=["a", "b", "c"].index(sig_val) if sig_val in ["a", "b", "c"] else 0)
        sig_key = sig_sel[0]  # 'a', 'b', or 'c'

    # ── SENDER / SALES CONTACT ───────────────────────────────────────────────
    st.markdown('<div class="section-header">👤 Sales Contact (Sender)</div>', unsafe_allow_html=True)

    sender_opts = [
        "Ahmad Awny | RC-AE SI EA S VD-V-D | +971 55 2003541 | ahmad.awny@siemens.com",
        "Vivek Gopalakrishnan | RC-AE SI EA S | +971 55 2002583 | vivek.gopalakrishnan@siemens.com",
        "Madiha Khan | RC-AE SI EA S | +971 55 2003818 | madiha.khan@siemens.com",
        "Hyun-Sik Kim | RC-AE SI EA S VD-V-D | +971 55 2002693 | hyun-sik.kim@siemens.com",
        "Other"
    ]
    sender_label = d.get("SENDER_LABEL", sender_opts[0])
    sender_sel = st.selectbox("Sales Contact *", sender_opts,
                              index=sender_opts.index(sender_label)
                              if sender_label in sender_opts else 0)

    if sender_sel == "Other":
        col1, col2 = st.columns(2)
        with col1:
            s_name  = st.text_input("Sender Name *",   value=d.get("SENDER_NAME", ""))
            s_dept  = st.text_input("Sender Dept *",   value=d.get("SENDER_DEPT", ""))
        with col2:
            s_mob   = st.text_input("Sender Mobile *", value=d.get("SENDER_MOBILE", ""))
            s_email = st.text_input("Sender Email *",  value=d.get("SENDER_EMAIL", ""))
    else:
        parts = [p.strip() for p in sender_sel.split("|")]
        s_name, s_dept, s_mob, s_email = parts[0], parts[1], parts[2], parts[3]

    # ── SCOPE OF SUPPLY ──────────────────────────────────────────────────────
    st.markdown('<div class="section-header">📦 Scope of Supply</div>', unsafe_allow_html=True)
    st.caption("Enter each line as: **Description | Qty | Total Price** — one item per line")

    # Pre-fill from existing scope_items
    existing_scope = d.get("SCOPE_ITEMS", [])
    scope_default = "\n".join(
        "%s | %s | %s" % (it.get("desc",""), it.get("qty",""), it.get("total",""))
        for it in existing_scope
    ) if existing_scope else ""

    scope_raw = st.text_area("Scope Items *", value=scope_default, height=150,
                             placeholder="33KV GIS Switchgear Panel | 10 | EUR 3,218,545\nControl Cable | 500m | EUR 25,000")

    st.caption("Optional items — same format, or leave blank")
    existing_opt = d.get("OPTIONAL_ITEMS", [])
    opt_default = "\n".join(
        "%s | %s | %s" % (it.get("desc",""), it.get("qty",""), it.get("total",""))
        for it in existing_opt
    ) if existing_opt else ""

    opt_raw = st.text_area("Optional Items", value=opt_default, height=100,
                           placeholder="Spare Parts Kit | 1 | EUR 5,000")

    # ── GENERATE BUTTON ──────────────────────────────────────────────────────
    st.divider()
    col_btn, col_reset = st.columns([3, 1])
    with col_btn:
        generate_clicked = st.button("⚡ Generate Offer Document", use_container_width=True)
    with col_reset:
        if st.button("🔄 Start Over", use_container_width=True):
            st.session_state.step = 1
            st.session_state.data = {}
            st.session_state.generated_file = None
            st.session_state.generated_name = None
            st.rerun()

    if generate_clicked:
        # ── Validation ───────────────────────────────────────────────────────
        errors = []
        if not country:               errors.append("Country is required.")
        if not company or company == "Other": errors.append("Customer Company is required.")
        if not attn:                  errors.append("Contact Person (Attn) is required.")
        if not offer_date:            errors.append("Offer Date is required.")
        if not subject:               errors.append("Subject is required.")
        if not project:               errors.append("Project Name is required.")
        if not end_user:              errors.append("End User is required.")
        if not price_num:             errors.append("Total Price is required.")
        if not scope_raw.strip():     errors.append("At least one Scope Item is required.")
        if is_firm:
            if not mfc_date:          errors.append("MFC Date is required.")
            if not offer_validity:    errors.append("Offer Validity is required.")

        if errors:
            for e in errors:
                st.error(e)
        else:
            # ── Parse currency ────────────────────────────────────────────────
            cur_code = currency_sel.split(" - ")[0].strip()
            cur_full = currency_sel.split(" - ")[1].strip()

            # ── Parse customer type ───────────────────────────────────────────
            ctype_map = {
                "National (UAE-based customer)":    "National",
                "International (non-UAE customer)": "International",
                "Siemens Energy":                   "Siemens Energy",
                "Critical Country":                 "Critical Country",
            }
            ctype_key = ctype_map.get(ctype, "National")

            # ── Import port mapping (Firm only) ───────────────────────────────
            import_port = ""
            if is_firm:
                port_map = {
                    "UAE - Abu Dhabi (Dubai or Northern Emirates)": "the Emirate of Abu Dhabi (Dubai or Northern Emirates)",
                    "UAE - Dubai / Jebel Ali":                      "the Emirate of Dubai / Jebel Ali",
                    "Saudi Arabia":                                  "a Saudi Arabian seaport",
                    "Algeria":                                       "an Algerian seaport",
                    "Yemen":                                         "a Yemeni seaport",
                }
                import_port = port_map.get(install_sel, install_sel)

            # ── Customer city ─────────────────────────────────────────────────
            city_map = {
                "Electro Mechanical Co. LLC (ELMEC)": "Abu Dhabi",
                "ADNOC":                              "Abu Dhabi",
                "Nofal for Trade & Agencies":         "Riyadh",
                "Siemens Energy LLC":                 "Abu Dhabi",
            }
            city = city_map.get(company, "Abu Dhabi")
            customer_city = "%s, %s" % (city, country)

            # ── Parse scope items ─────────────────────────────────────────────
            scope_items = []
            for idx, line in enumerate(scope_raw.strip().splitlines()):
                line = line.strip()
                if not line:
                    continue
                parts = [p.strip() for p in line.split("|")]
                scope_items.append({
                    "no":    str(idx + 1),
                    "desc":  parts[0] if len(parts) > 0 else "",
                    "qty":   parts[1] if len(parts) > 1 else "",
                    "total": parts[2] if len(parts) > 2 else "",
                })

            opt_items = []
            for idx, line in enumerate(opt_raw.strip().splitlines()):
                line = line.strip()
                if not line:
                    continue
                parts = [p.strip() for p in line.split("|")]
                opt_items.append({
                    "no":    str(idx + 1),
                    "desc":  parts[0] if len(parts) > 0 else "",
                    "qty":   parts[1] if len(parts) > 1 else "",
                    "total": parts[2] if len(parts) > 2 else "",
                })

            # ── Payment option ────────────────────────────────────────────────
            pay_opt = "A" if "Option A" in pay_opt_sel else "B"

            # ── Price in words ────────────────────────────────────────────────
            price_words = number_to_words(price_num)

            # ── Build data dict ───────────────────────────────────────────────
            data = {
                "OFFER_TYPE":       d["OFFER_TYPE"],
                "CUSTOMER_TYPE":    ctype_key,
                "COUNTRY":          country,
                "CUSTOMER_COMPANY": company,
                "CUSTOMER_CITY":    customer_city,
                "CUSTOMER_ATTN":    attn,
                "CUSTOMER_PO_BOX":  po_box,
                "CUSTOMER_TEL":     tel,
                "CUSTOMER_FAX":     fax,
                "CUSTOMER_MOB":     mob,
                "SENDER_NAME":      s_name,
                "SENDER_DEPT":      s_dept,
                "SENDER_MOBILE":    s_mob,
                "SENDER_EMAIL":     s_email,
                "OFFER_DATE":       offer_date,
                "REFERENCE_NO":     ref_no,
                "SUBJECT":          subject,
                "PROJECT_NAME":     project,
                "END_USER":         end_user,
                "CURRENCY_CODE":    cur_code,
                "CURRENCY_FULL":    cur_full,
                "TOTAL_PRICE_NUM":  price_num,
                "TOTAL_PRICE_WORDS": price_words,
                "INCOTERM_NAME":    inco_sel,
                "DELIVERY_MONTHS":  del_sel,
                "WARRANTY_MONTHS":  war_sel,
                "PAYMENT_OPTION":   pay_opt,
                "PAYMENT_DAYS":     pay_days_sel,
                "IMPORT_PORT":      import_port if is_firm else "N/A",
                "MFC_DATE":         mfc_date if is_firm else "N/A",
                "OFFER_VALIDITY":   offer_validity if is_firm else "N/A",
                "SIGNATORIES":      sig_key if is_firm else "N/A",
                "COUNTRY_OF_INSTALL": install_sel if is_firm else "N/A",
                "SCOPE_ITEMS":      scope_items,
                "OPTIONAL_ITEMS":   opt_items,
                # Labels for re-population
                "CURRENCY_LABEL":   currency_sel,
                "SENDER_LABEL":     sender_sel,
                "CUSTOMER_TYPE_LABEL": ctype,
            }

            st.session_state.data = data
            st.session_state.step = 3
            st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3 — CONFIRMATION SUMMARY
# ─────────────────────────────────────────────────────────────────────────────
elif st.session_state.step == 3:
    d = st.session_state.data
    is_firm = (d.get("OFFER_TYPE") == "Firm")

    st.markdown("### Step 3 of 4 — Confirm Details")
    st.success("Please review all details below before generating the document.")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Offer Type:**  " + d.get("OFFER_TYPE", ""))
        st.markdown("**Customer:**  " + d.get("CUSTOMER_COMPANY", "") + " — " + d.get("COUNTRY", ""))
        st.markdown("**Customer Type:**  " + d.get("CUSTOMER_TYPE", ""))
        st.markdown("**Contact:**  " + d.get("CUSTOMER_ATTN", ""))
        st.markdown("**P.O. Box:**  " + (d.get("CUSTOMER_PO_BOX", "") or "—"))
        st.markdown("**Tel:**  " + (d.get("CUSTOMER_TEL", "") or "—"))
        st.markdown("**Fax:**  " + (d.get("CUSTOMER_FAX", "") or "—"))
        if not is_firm:
            st.markdown("**Mobile:**  " + (d.get("CUSTOMER_MOB", "") or "—"))
        st.markdown("**Project:**  " + d.get("PROJECT_NAME", ""))
        st.markdown("**End User:**  " + d.get("END_USER", ""))
        st.markdown("**Subject:**  " + d.get("SUBJECT", ""))
        st.markdown("**Reference No.:**  " + (d.get("REFERENCE_NO", "") or "—"))
        st.markdown("**Offer Date:**  " + d.get("OFFER_DATE", ""))
    with col2:
        st.markdown("**Sender:**  " + d.get("SENDER_NAME", "") + " | " + d.get("SENDER_DEPT", ""))
        st.markdown("**Currency:**  " + d.get("CURRENCY_FULL", "") + " (" + d.get("CURRENCY_CODE", "") + ")")
        st.markdown("**Total Price:**  " + d.get("TOTAL_PRICE_NUM", ""))
        st.markdown("**Price in Words:**  " + d.get("TOTAL_PRICE_WORDS", ""))
        st.markdown("**Incoterm:**  " + d.get("INCOTERM_NAME", ""))
        st.markdown("**Delivery:**  " + d.get("DELIVERY_MONTHS", "") + " months")
        st.markdown("**Warranty:**  " + d.get("WARRANTY_MONTHS", "") + " months")
        st.markdown("**Payment:**  Option " + d.get("PAYMENT_OPTION", "") + " — " + d.get("PAYMENT_DAYS", "") + " days")
        if is_firm:
            st.markdown("**Installation:**  " + d.get("COUNTRY_OF_INSTALL", ""))
            st.markdown("**MFC Date:**  " + d.get("MFC_DATE", ""))
            st.markdown("**Offer Validity:**  " + d.get("OFFER_VALIDITY", ""))
            sig_labels = {
                "a": "Robert Hennig + Rodrigo Fernandes",
                "b": "Robert Hennig + Rodrigo Fernandes + Prashanth Parameswaran + Paul Fairweather",
                "c": "Robert Hennig + Rodrigo Fernandes + Prashanth Parameswaran + Paul Fairweather + Nasir Merchant",
            }
            st.markdown("**Signatories:**  " + sig_labels.get(d.get("SIGNATORIES", "a"), ""))
        st.markdown("**Scope Items:**  " + str(len(d.get("SCOPE_ITEMS", []))) + " item(s)")
        st.markdown("**Optional Items:**  " + str(len(d.get("OPTIONAL_ITEMS", []))) + " item(s)")

    st.divider()
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        if st.button("✅ Confirm & Generate Document", use_container_width=True):
            with st.spinner("Generating your offer document..."):
                try:
                    filepath, filename = generate_document(d)
                    st.session_state.generated_file = filepath
                    st.session_state.generated_name = filename
                    st.session_state.step = 4
                    st.rerun()
                except Exception as ex:
                    st.error("Error generating document: %s" % str(ex))
    with col2:
        if st.button("✏️ Go Back & Edit", use_container_width=True):
            st.session_state.step = 2
            st.rerun()
    with col3:
        if st.button("🔄 Start Over", use_container_width=True):
            st.session_state.step = 1
            st.session_state.data = {}
            st.session_state.generated_file = None
            st.session_state.generated_name = None
            st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4 — DOWNLOAD
# ─────────────────────────────────────────────────────────────────────────────
elif st.session_state.step == 4:
    d = st.session_state.data
    st.markdown("### Step 4 of 4 — Download Your Offer")
    st.balloons()
    st.success("✅ Your offer document has been generated successfully!")

    st.markdown("---")
    st.markdown("**Document:** `%s`" % st.session_state.generated_name)
    st.markdown("**Customer:** %s" % d.get("CUSTOMER_COMPANY", ""))
    st.markdown("**Project:** %s" % d.get("PROJECT_NAME", ""))
    st.markdown("**Date:** %s" % d.get("OFFER_DATE", ""))
    st.markdown("---")

    filepath = st.session_state.generated_file
    if filepath and os.path.exists(filepath):
        with open(filepath, "rb") as f:
            file_bytes = f.read()
        st.download_button(
            label="📥 Download Offer Letter (.docx)",
            data=file_bytes,
            file_name=st.session_state.generated_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    else:
        st.error("File not found. Please generate again.")

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("✏️ Edit This Offer", use_container_width=True):
            st.session_state.step = 2
            st.session_state.generated_file = None
            st.session_state.generated_name = None
            st.rerun()
    with col2:
        if st.button("🆕 Generate New Offer", use_container_width=True):
            st.session_state.step = 1
            st.session_state.data = {}
            st.session_state.generated_file = None
            st.session_state.generated_name = None
            st.rerun()
