import streamlit as st
import shutil
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from datetime import datetime
import os
import re

st.set_page_config(page_title="Siemens Offer Letter Generator", page_icon="⚡", layout="wide")

# ─────────────────────────────────────────────────────────────
# SIEMENS PROFESSIONAL THEME
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* ── Google Font ── */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    /* ── Global reset ── */
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* ── Page background ── */
    .stApp {
        background-color: #f4f6f9;
    }

    /* ── Hide Streamlit default chrome ── */
    #MainMenu, footer, header { visibility: hidden; }

    /* ── Top banner ── */
    .siemens-banner {
        background: linear-gradient(135deg, #009999 0%, #007a7a 60%, #005f5f 100%);
        border-radius: 12px;
        padding: 28px 36px;
        margin-bottom: 24px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        box-shadow: 0 4px 20px rgba(0,153,153,0.25);
    }
    .siemens-banner-left h1 {
        color: #ffffff;
        font-size: 1.9rem;
        font-weight: 700;
        margin: 0;
        letter-spacing: -0.3px;
    }
    .siemens-banner-left p {
        color: rgba(255,255,255,0.80);
        font-size: 0.88rem;
        margin: 4px 0 0 0;
        font-weight: 400;
    }
    .siemens-logo-box {
        background: rgba(255,255,255,0.15);
        border-radius: 8px;
        padding: 10px 20px;
        color: #ffffff;
        font-size: 1.5rem;
        font-weight: 800;
        letter-spacing: 2px;
        border: 1px solid rgba(255,255,255,0.3);
    }

    /* ── Section cards ── */
    .section-card {
        background: #ffffff;
        border-radius: 10px;
        padding: 24px 28px;
        margin-bottom: 20px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 6px rgba(0,0,0,0.06);
    }

    /* ── Section header ── */
    .section-header {
        font-size: 1.0rem;
        font-weight: 700;
        color: #009999;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        margin-bottom: 16px;
        padding-bottom: 10px;
        border-bottom: 2px solid #009999;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    /* ── Divider ── */
    hr { border: none; border-top: 1px solid #e2e8f0; margin: 20px 0; }

    /* ── Input labels ── */
    label, .stSelectbox label, .stTextInput label,
    .stNumberInput label, .stTextArea label, .stRadio label {
        font-size: 0.82rem !important;
        font-weight: 600 !important;
        color: #374151 !important;
        text-transform: uppercase;
        letter-spacing: 0.4px;
    }

    /* ── Input boxes ── */
    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea,
    .stNumberInput > div > div > input {
        border: 1.5px solid #d1d5db !important;
        border-radius: 7px !important;
        font-size: 0.9rem !important;
        color: #111827 !important;
        background: #fafafa !important;
        transition: border-color 0.2s;
    }
    .stTextInput > div > div > input:focus,
    .stTextArea > div > div > textarea:focus {
        border-color: #009999 !important;
        background: #ffffff !important;
        box-shadow: 0 0 0 3px rgba(0,153,153,0.12) !important;
    }

    /* ── Selectbox ── */
    .stSelectbox > div > div {
        border: 1.5px solid #d1d5db !important;
        border-radius: 7px !important;
        background: #fafafa !important;
    }

    /* ── Radio buttons ── */
    .stRadio > div {
        gap: 6px;
    }
    .stRadio > div > label {
        background: #f8fafc;
        border: 1.5px solid #e2e8f0;
        border-radius: 8px;
        padding: 8px 16px !important;
        font-size: 0.85rem !important;
        font-weight: 500 !important;
        text-transform: none !important;
        letter-spacing: 0 !important;
        color: #374151 !important;
        transition: all 0.2s;
        cursor: pointer;
    }
    .stRadio > div > label:hover {
        border-color: #009999;
        background: #f0fafa;
        color: #009999 !important;
    }
    [data-testid="stRadio"] [aria-checked="true"] + div label,
    .stRadio > div > label[data-checked="true"] {
        border-color: #009999 !important;
        background: #e6f7f7 !important;
        color: #009999 !important;
    }

    /* ── Primary button ── */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #009999 0%, #007a7a 100%) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px !important;
        font-size: 1.0rem !important;
        font-weight: 700 !important;
        padding: 14px 0 !important;
        letter-spacing: 0.3px;
        box-shadow: 0 4px 14px rgba(0,153,153,0.35) !important;
        transition: all 0.2s !important;
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #00b3b3 0%, #009999 100%) !important;
        box-shadow: 0 6px 20px rgba(0,153,153,0.45) !important;
        transform: translateY(-1px);
    }

    /* ── Download button ── */
    .stDownloadButton > button {
        background: #ffffff !important;
        color: #009999 !important;
        border: 2px solid #009999 !important;
        border-radius: 8px !important;
        font-size: 0.95rem !important;
        font-weight: 600 !important;
        padding: 12px 0 !important;
        transition: all 0.2s !important;
    }
    .stDownloadButton > button:hover {
        background: #009999 !important;
        color: #ffffff !important;
        box-shadow: 0 4px 14px rgba(0,153,153,0.3) !important;
    }

    /* ── Success / error / info boxes ── */
    .stSuccess {
        background: #ecfdf5 !important;
        border-left: 4px solid #10b981 !important;
        border-radius: 8px !important;
        color: #065f46 !important;
    }
    .stAlert[data-baseweb="notification"] {
        border-radius: 8px !important;
    }

    /* ── Info box ── */
    .stInfo {
        background: #eff6ff !important;
        border-left: 4px solid #3b82f6 !important;
        border-radius: 8px !important;
    }

    /* ── Number input ── */
    .stNumberInput > div { border-radius: 7px !important; }

    /* ── Caption / helper text ── */
    .stCaption, small {
        color: #6b7280 !important;
        font-size: 0.80rem !important;
    }

    /* ── Scope item cards ── */
    .scope-item-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-left: 3px solid #009999;
        border-radius: 8px;
        padding: 14px 18px;
        margin-bottom: 12px;
    }

    /* ── Filename preview pill ── */
    .filename-pill {
        display: inline-block;
        background: #e6f7f7;
        border: 1px solid #009999;
        border-radius: 20px;
        padding: 5px 16px;
        font-size: 0.82rem;
        font-weight: 600;
        color: #007a7a;
        margin-top: 6px;
    }

    /* ── Footer ── */
    .siemens-footer {
        text-align: center;
        color: #9ca3af;
        font-size: 0.75rem;
        padding: 24px 0 8px 0;
        border-top: 1px solid #e2e8f0;
        margin-top: 32px;
    }
    .siemens-footer span {
        color: #009999;
        font-weight: 600;
    }

    /* ── Step badge ── */
    .step-badge {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 24px;
        height: 24px;
        background: #009999;
        color: white;
        border-radius: 50%;
        font-size: 0.72rem;
        font-weight: 700;
        margin-right: 8px;
        flex-shrink: 0;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# TOP BANNER
# ─────────────────────────────────────────────────────────────
st.markdown("""
<div class="siemens-banner">
    <div class="siemens-banner-left">
        <h1>⚡ Offer Letter Generator</h1>
        <p>Siemens Industrial LLC &nbsp;·&nbsp; Smart Document Automation &nbsp;·&nbsp; RC-AE SI EA S</p>
    </div>
    <div class="siemens-logo-box">SIEMENS</div>
</div>
""", unsafe_allow_html=True)

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

    for section in doc.sections:
        for hdr in [section.header, section.first_page_header, section.even_page_header]:
            if hdr is None:
                continue
            for p in hdr.paragraphs:
                yield p
            for tbl in hdr.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            yield p

def merge_and_replace(para, replacements):
    runs = para.runs
    if not runs:
        return

    changed = True
    while changed:
        changed = False
        runs = para.runs
        for i in range(len(runs) - 1):
            r0, r1 = runs[i], runs[i + 1]
            if r0.font.highlight_color is not None and r1.font.highlight_color is not None:
                combined = (r0.text or "") + (r1.text or "")
                if combined.strip() in replacements or "INSERT_" in combined:
                    r0.text = combined
                    r1.text = ""
                    clear_highlight(r1)
                    r1.font.highlight_color = None
                    changed = True
                    break

    for r in para.runs:
        if not r.text:
            continue
        key = r.text.strip()
        if r.font.highlight_color is not None:
            if key in replacements:
                fill(r, replacements[key])
            elif "INSERT_" in r.text:
                new_text = re.sub(
                    r"INSERT_\w+",
                    lambda m: replacements.get(m.group(0), ""),
                    r.text
                )
                fill(r, new_text)

def fill_table(table, items):
    for ri, item in enumerate(items):
        while ri + 1 >= len(table.rows):
            table.add_row()
        row = table.rows[ri + 1]
        vals = [item.get("no", str(ri + 1)), item.get("desc", ""),
                item.get("qty", ""), item.get("total", "")]
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

st.markdown('<div class="section-card"><div class="section-header"><span class="step-badge">1</span> Offer Type</div>', unsafe_allow_html=True)
offer_type = st.radio("Select the type of commercial offer to generate:", ["Firm", "Budgetary"], horizontal=True)
is_firm = (offer_type == "Firm")
if is_firm:
    st.info("📋 **Firm Offer** — Legally binding commercial offer with full terms, MFC date, validity, and signatories.")
else:
    st.info("📋 **Budgetary Offer** — Non-binding indicative price offer for budgeting purposes.")
st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SECTION 2 - CUSTOMER INFORMATION
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="section-card"><div class="section-header"><span class="step-badge">2</span> Customer Information</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    company_options = ["Electro Mechanical Co. LLC (ELMEC)", "ADNOC", "Nofal for Trade & Agencies", "Siemens Energy LLC", "Other"]
    customer_company_sel = st.selectbox("Customer Company", company_options)
    customer_company = st.text_input("Enter Company Name", placeholder="Type company name...") if customer_company_sel == "Other" else customer_company_sel

with col2:
    country_options = ["UAE", "Saudi Arabia", "Algeria", "Yemen", "Kuwait", "Oman", "Qatar", "Other"]
    country_sel = st.selectbox("Country", country_options)
    country = st.text_input("Enter Country Name", placeholder="Type country name...") if country_sel == "Other" else country_sel

col3, col4 = st.columns(2)
with col3:
    customer_attn = st.text_input("Contact Person", placeholder="e.g. Mr. John Smith")
with col4:
    city_default = "Abu Dhabi" if country in ["UAE", ""] else country
    customer_city_input = st.text_input("City", value=city_default)

col5, col6 = st.columns(2)
with col5:
    customer_po_box = st.text_input("P.O. Box", placeholder="Leave blank if not applicable")
with col6:
    customer_fax = st.text_input("Fax", placeholder="Leave blank if not applicable")

col7, col8 = st.columns(2)
with col7:
    customer_tel = st.text_input("Tel", placeholder="Leave blank if not applicable")
with col8:
    customer_mob = st.text_input("Mobile", placeholder="Leave blank if not applicable")

col9, col10 = st.columns(2)
with col9:
    reference_no = st.text_input("Reference No.", placeholder="Leave blank if not applicable")
with col10:
    offer_date = st.text_input("Offer Date", placeholder="e.g. 15 April 2025")

st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SECTION 3 - OFFER DETAILS
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="section-card"><div class="section-header"><span class="step-badge">3</span> Offer Details</div>', unsafe_allow_html=True)
subject      = st.text_input("Subject", placeholder="e.g. 33KV Switchgear Supply")
col_od1, col_od2 = st.columns(2)
with col_od1:
    project_name = st.text_input("Project Name", placeholder="e.g. PDHPP Project")
with col_od2:
    end_user = st.text_input("End User", placeholder="e.g. STEP Spa")
st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SECTION 4 - COMMERCIAL TERMS
# ─────────────────────────────────────────────────────────────

st.markdown('<div class="section-card"><div class="section-header"><span class="step-badge">4</span> Commercial Terms</div>', unsafe_allow_html=True)

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
    total_price_num = st.text_input("Total Price", placeholder="e.g. 3,218,545")

col_c3, col_c4 = st.columns(2)
with col_c3:
    incoterm_options = ["FCA, Bremerhaven, Germany", "CIF, Abu Dhabi seaport",
                        "DAP, Abu Dhabi", "EXW, Bremerhaven, Germany", "Other"]
    incoterm_sel = st.selectbox("Incoterm", incoterm_options)
    incoterm_name = st.text_input("Enter Incoterm", placeholder="e.g. FCA, Hamburg, Germany") if incoterm_sel == "Other" else incoterm_sel
with col_c4:
    delivery_options = ["06", "09", "12", "15", "18", "24", "Other"]
    delivery_sel = st.selectbox("Delivery Period (months)", delivery_options)
    delivery_months = st.text_input("Enter months", placeholder="e.g. 10") if delivery_sel == "Other" else delivery_sel

col_c5, col_c6 = st.columns(2)
with col_c5:
    warranty_options = ["12", "18", "24", "Other"]
    warranty_sel = st.selectbox("Warranty Period (months)", warranty_options)
    warranty_months = st.text_input("Enter warranty months", placeholder="e.g. 24") if warranty_sel == "Other" else warranty_sel

st.markdown("---")
st.markdown("**💳 Payment Structure**")

payment_option_labels = [
    "Option A — 20% down payment + 80% against shipping documents",
    "Option B — 20% advance + 10% drawings + 20% MFC + 50% bill of lading",
    "Option C — Custom payment terms",
]
payment_option_sel = st.radio(
    "Select payment structure",
    payment_option_labels,
    label_visibility="collapsed",
)

if payment_option_sel.startswith("Option A"):
    payment_option = "A"
    custom_payment_text = ""
    col_pd1, col_pd2 = st.columns([1, 3])
    with col_pd1:
        payment_days_options = ["30", "45", "60", "90", "120", "Other"]
        payment_days_sel = st.selectbox("Payment Days", payment_days_options, key="pd_A")
        payment_days = st.text_input("Enter days", key="pd_A_custom") if payment_days_sel == "Other" else payment_days_sel

elif payment_option_sel.startswith("Option B"):
    payment_option = "B"
    custom_payment_text = ""
    col_pd1, col_pd2 = st.columns([1, 3])
    with col_pd1:
        payment_days_options = ["30", "45", "60", "90", "120", "Other"]
        payment_days_sel = st.selectbox("Payment Days", payment_days_options, key="pd_B")
        payment_days = st.text_input("Enter days", key="pd_B_custom") if payment_days_sel == "Other" else payment_days_sel

else:
    payment_option = "C"
    custom_payment_text = st.text_area(
        "Enter custom payment terms",
        placeholder=(
            "e.g.\n"
            "100% of the contract value payable against the date of "
            "delivery documents within 90 days."
        ),
        height=120,
        key="custom_payment_text",
    )
    col_pd1, col_pd2 = st.columns([1, 3])
    with col_pd1:
        payment_days_options = ["30", "45", "60", "90", "120", "Other"]
        payment_days_sel = st.selectbox(
            "Payment Days (for header)",
            payment_days_options,
            key="pd_C",
            help="Used in the section header line only",
        )
        payment_days = (
            st.text_input("Enter days", key="pd_C_custom")
            if payment_days_sel == "Other"
            else payment_days_sel
        )
    with col_pd2:
        st.info(
            f"📌 Header will read: **Payment terms: {payment_days} days from delivery documents**\n\n"
            "The text above will be inserted verbatim into the document."
        )

if is_firm:
    st.markdown("---")
    cancel_high = st.selectbox(
        "Cancellation Charge — High Bracket",
        ["80%", "90%"],
        help="Select -80% for National customers, -90% for International / Siemens Energy / Critical Country"
    )
else:
    cancel_high = ""

st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SECTION 5 - FIRM ONLY FIELDS
# ─────────────────────────────────────────────────────────────

if is_firm:
    st.markdown('<div class="section-card"><div class="section-header"><span class="step-badge">5</span> Firm Offer Details</div>', unsafe_allow_html=True)

    install_options = {
        "UAE - Abu Dhabi (Dubai or Northern Emirates)": "the Emirate of Abu Dhabi (Dubai or Northern Emirates)",
        "UAE - Dubai / Jebel Ali": "the Emirate of Dubai / Jebel Ali",
        "Saudi Arabia": "a Saudi Arabian seaport",
        "Algeria": "an Algerian seaport",
        "Yemen": "a Yemeni seaport",
        "Other": None
    }
    install_sel = st.selectbox("Country / Port of Installation", list(install_options.keys()))
    import_port = st.text_input("Enter country / port", placeholder="e.g. a Kuwaiti seaport") if install_sel == "Other" else install_options[install_sel]

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        mfc_date = st.text_input("MFC Date", placeholder="e.g. 26th February 2026")
    with col_f2:
        offer_validity = st.text_input("Offer Validity", placeholder="e.g. March 30, 2026")

    signatory_options = [
        "Robert Hennig + Rodrigo Fernandes",
        "Robert Hennig + Rodrigo Fernandes + Prashanth Parameswaran + Paul Fairweather",
        "Robert Hennig + Rodrigo Fernandes + Prashanth Parameswaran + Paul Fairweather + Nasir Merchant",
        "Other"
    ]
    signatory_sel = st.selectbox("Signatories", signatory_options)
    signatories = st.text_input("Enter signatory names", placeholder="Name 1 + Name 2 + ...") if signatory_sel == "Other" else signatory_sel
    st.markdown('</div>', unsafe_allow_html=True)
else:
    import_port    = ""
    mfc_date       = ""
    offer_validity = ""
    signatories    = ""

# ─────────────────────────────────────────────────────────────
# SECTION 6 - SALES CONTACT
# ─────────────────────────────────────────────────────────────

step_num = "6" if is_firm else "5"
st.markdown(f'<div class="section-card"><div class="section-header"><span class="step-badge">{step_num}</span> Sales Contact</div>', unsafe_allow_html=True)

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
        sender_name   = st.text_input("Sender Name", placeholder="Full name")
        sender_mobile = st.text_input("Sender Mobile", placeholder="+971 XX XXXXXXX")
    with col_s2:
        sender_dept  = st.text_input("Sender Department", placeholder="e.g. RC-AE SI EA S")
        sender_email = st.text_input("Sender Email", placeholder="name@siemens.com")
else:
    sc = sales_contacts[sales_sel]
    sender_name   = sc["name"]
    sender_dept   = sc["dept"]
    sender_mobile = sc["mobile"]
    sender_email  = sc["email"]
    col_si1, col_si2, col_si3 = st.columns(3)
    with col_si1:
        st.markdown(f"<div style='background:#f0fafa;border:1px solid #009999;border-radius:7px;padding:10px 14px;font-size:0.85rem;'><b>📧</b> {sender_email}</div>", unsafe_allow_html=True)
    with col_si2:
        st.markdown(f"<div style='background:#f0fafa;border:1px solid #009999;border-radius:7px;padding:10px 14px;font-size:0.85rem;'><b>📱</b> {sender_mobile}</div>", unsafe_allow_html=True)
    with col_si3:
        st.markdown(f"<div style='background:#f0fafa;border:1px solid #009999;border-radius:7px;padding:10px 14px;font-size:0.85rem;'><b>🏢</b> {sender_dept}</div>", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SECTION 7 - SCOPE OF SUPPLY
# ─────────────────────────────────────────────────────────────

step_num2 = "7" if is_firm else "6"
st.markdown(f'<div class="section-card"><div class="section-header"><span class="step-badge">{step_num2}</span> Scope of Supply</div>', unsafe_allow_html=True)

num_scope = st.number_input("Number of scope items", min_value=1, max_value=20, value=1, step=1)
scope_items = []
for i in range(int(num_scope)):
    st.markdown(f'<div class="scope-item-card">', unsafe_allow_html=True)
    st.markdown(f"**Item {i+1}**")
    col_sc1, col_sc2, col_sc3 = st.columns([4, 1, 2])
    with col_sc1: desc  = st.text_input("Description", key=f"scope_desc_{i}", placeholder="e.g. 33KV 8DA10 AIS NXAIR Switchgear Panel")
    with col_sc2: qty   = st.text_input("Qty",         key=f"scope_qty_{i}",  placeholder="e.g. 10")
    with col_sc3: total = st.text_input("Total Price", key=f"scope_total_{i}", placeholder="e.g. EUR 3,218,545")
    scope_items.append({"no": str(i+1), "desc": desc, "qty": qty, "total": total})
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")
st.markdown("**Optional Items**")
num_opt = st.number_input("Number of optional items", min_value=0, max_value=10, value=0, step=1)
optional_items = []
for i in range(int(num_opt)):
    st.markdown(f'<div class="scope-item-card">', unsafe_allow_html=True)
    st.markdown(f"**Optional Item {i+1}**")
    col_o1, col_o2, col_o3 = st.columns([4, 1, 2])
    with col_o1: odesc  = st.text_input("Description", key=f"opt_desc_{i}")
    with col_o2: oqty   = st.text_input("Qty",         key=f"opt_qty_{i}")
    with col_o3: ototal = st.text_input("Total Price", key=f"opt_total_{i}")
    optional_items.append({"no": str(i+1), "desc": odesc, "qty": oqty, "total": ototal})
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SECTION 8 - OUTPUT FILENAME
# ─────────────────────────────────────────────────────────────

step_num3 = "8" if is_firm else "7"
st.markdown(f'<div class="section-card"><div class="section-header"><span class="step-badge">{step_num3}</span> Output Filename</div>', unsafe_allow_html=True)

offer_short_prev = "FIRM" if is_firm else "BUD"
cust_short_prev  = customer_company.split()[0].upper().strip(".,)") if customer_company else "CUSTOMER"
SKIP_WORDS       = {"FOR", "THE", "OF", "AND", "IN", "A", "AN", "PKG", "WORKS", "-", "PROJECT"}
proj_short_prev  = "".join(w for w in project_name.split() if w.upper() not in SKIP_WORDS)[:12].upper() if project_name else "PROJECT"
try:    date_str_prev = datetime.strptime(offer_date.strip(), "%d %B %Y").strftime("%Y%m%d")
except: date_str_prev = offer_date.replace(" ", "")[:8] if offer_date else "DATE"

auto_filename = f"{offer_short_prev}_{cust_short_prev}_{proj_short_prev}_{date_str_prev}_v01"

filename_mode = st.radio(
    "Filename mode",
    ["🔄 Auto", "✏️ Custom"],
    horizontal=True,
    label_visibility="collapsed",
)

if filename_mode == "✏️ Custom":
    custom_filename_input = st.text_input(
        "Enter filename (without extension)",
        placeholder=f"e.g. {auto_filename}",
    )
    raw = custom_filename_input.strip().removesuffix(".docx")
    final_filename = (raw if raw else auto_filename) + ".docx"
else:
    final_filename = auto_filename + ".docx"

st.markdown(f'<div class="filename-pill">📄 {final_filename}</div>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# GENERATE
# ─────────────────────────────────────────────────────────────

st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 Generate Offer Letter", type="primary", use_container_width=True):

    errors = []
    if not customer_company or customer_company == "Other": errors.append("Customer Company")
    if not customer_attn:   errors.append("Contact Person")
    if not offer_date:      errors.append("Offer Date")
    if not subject:         errors.append("Subject")
    if not project_name:    errors.append("Project Name")
    if not end_user:        errors.append("End User")
    if not total_price_num: errors.append("Total Price")
    if payment_option == "C" and not custom_payment_text.strip():
        errors.append("Custom Payment Terms (Option C cannot be empty)")
    if is_firm and not mfc_date:       errors.append("MFC Date")
    if is_firm and not offer_validity: errors.append("Offer Validity")
    if errors:
        st.error(f"⚠️ Please fill in the following required fields: **{', '.join(errors)}**")
        st.stop()

    customer_city     = f"{customer_city_input}, {country}"
    total_price_words = number_to_words(total_price_num)

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
        pay_header = f"Payment terms: {payment_days} days from delivery documents"
        pay_lines  = custom_payment_text.strip()

    import_port_sentence = (
        f"The scope shall be offloaded in and imported via a port of {import_port}."
        if is_firm and import_port else ""
    )

    po_box_full = f"P. O. Box {customer_po_box}" if customer_po_box.strip() else ""
    tel_str     = f"Tel : {customer_tel}"         if customer_tel.strip()    else ""
    fax_str     = f"Fax: {customer_fax}"          if customer_fax.strip()    else ""
    mob_str     = f"Mob: {customer_mob}"          if customer_mob.strip()    else ""

    price_words_run = f"{currency_code} {total_price_words} Only)."

    replacements = {
        "INSERT_CUSTOMER_COMPANY":                          customer_company,
        "P. O. Box INSERT_CUSTOMER_PO_BOX_NUM":            po_box_full,
        "INSERT_CUSTOMER_CITY_FULL":                        customer_city,
        "INSERT_CUSTOMER_ATTN":                             f"Kind Attn: {customer_attn}",
        "INSERT_CUSTOMER_TEL_LINE":                         tel_str,
        "INSERT_CUSTOMER_FAX_LINE":                         fax_str,
        "INSERT_CUSTOMER_MOB_LINE":                         mob_str,
        "INSERT_SENDER_NAME":                               sender_name,
        "INSERT_SENDER_DEPT":                               sender_dept,
        "INSERT_SENDER_MOBILE":                             sender_mobile,
        "INSERT_SENDER_EMAIL":                              sender_email,
        "INSERT_OFFER_DATE":                                offer_date,
        "INSERT_REFERENCE_NO":                              reference_no.strip(),
        "INSERT_SUBJECT_FULL":                              subject,
        "INSERT_PROJECT_NAME":                              project_name,
        "INSERT_END_USER":                                  end_user,
        "INSERT_CURRENCY_FULL":                             currency_full,
        "INSERT_TOTAL_PRICE_NUM":                           total_price_num,
        "INSERT_CURRENCY_CODE":                             currency_code,
        "INSERT_CURRENCY_CODE INSERT_TOTAL_PRICE_WORDS).":  price_words_run,
        "INSERT_CURRENCY_CODE INSERT_TOTAL_PRICE_WORDS Only).": price_words_run,
        "INSERT_INCOTERM_NAME":                             incoterm_name,
        "INSERT_DELIVERY_MONTHS":                           f"{delivery_months} month(s)",
        "INSERT_WARRANTY_MONTHS":                           f"{warranty_months} months",
        "INSERT_PAYMENT_OPTION_HEADER":                     pay_header,
        "INSERT_PAYMENT_OPTION_LINE":                       pay_lines,
        "INSERT_PAYMENT_OPTION_LINES":                      pay_lines,
        "INSERT_MFC_DATE":                                  mfc_date if is_firm else "",
        "INSERT_IMPORT_PORT_SENTENCE":                      import_port_sentence,
        "INSERT_OFFER_VALIDITY":                            offer_validity if is_firm else "",
        "INSERT_CANCEL_HIGH":                               cancel_high,
    }

    TEMPLATE_PATH = (
        os.path.join(BASE_DIR, "template_Firm_FIXED.docx")
        if is_firm else
        os.path.join(BASE_DIR, "template_Budgetary_FIXED.docx")
    )

    filepath = os.path.join(BASE_DIR, final_filename)
    shutil.copy(TEMPLATE_PATH, filepath)
    doc = Document(filepath)

    for para in all_paras(doc):
        merge_and_replace(para, replacements)

    for para in all_paras(doc):
        runs = para.runs
        i = 0
        while i < len(runs) - 1:
            r0, r1 = runs[i], runs[i + 1]
            if "INSERT_" in (r0.text or "") and "INSERT_" in (r1.text or ""):
                combined = (r0.text or "") + (r1.text or "")
                if combined.strip() in replacements:
                    r0.text = combined
                    r1.text = ""
                    continue
            i += 1
        for r in para.runs:
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

    if len(doc.tables) > 1:
        fill_table(doc.tables[1], scope_items)
    if optional_items and len(doc.tables) > 2:
        fill_table(doc.tables[2], optional_items)

    for para in all_paras(doc):
        for r in para.runs:
            if r.text and "INSERT_" in r.text:
                r.text = re.sub(r"INSERT_\w+", "", r.text)
            if r.font.highlight_color is not None:
                clear_highlight(r)
                r.font.highlight_color = None

    doc.save(filepath)

    st.success(f"✅ **Offer letter generated successfully!** Your document is ready for download.")
    with open(filepath, "rb") as f:
        st.download_button(
            label=f"📥 Download {final_filename}",
            data=f,
            file_name=final_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

# ─────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────
st.markdown("""
<div class="siemens-footer">
    <span>Siemens Industrial LLC</span> &nbsp;·&nbsp; Offer Letter Generator &nbsp;·&nbsp;
    RC-AE SI EA S &nbsp;·&nbsp; Internal Use Only
</div>
""", unsafe_allow_html=True)
