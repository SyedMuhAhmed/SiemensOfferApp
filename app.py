
import streamlit as st
import shutil
import re
from docx import Document
from docx.shared import Pt
from datetime import datetime
import os

st.set_page_config(page_title="Siemens Offer Letter Generator", page_icon="⚡", layout="wide")

# ── Siemens branding ──────────────────────────────────────────────────────────
st.markdown("""
<style>
  .main { background-color: #f5f5f5; }
  .stButton>button { background-color: #009999; color: white; border-radius: 6px; font-weight: bold; }
  .stButton>button:hover { background-color: #007777; }
  h1 { color: #009999; }
  h2, h3 { color: #333333; }
  .block-container { padding-top: 2rem; }
</style>
""", unsafe_allow_html=True)

st.title("⚡ Siemens Offer Letter Generator")
st.markdown("---")

# ── Helpers ───────────────────────────────────────────────────────────────────
def num_to_words(n_str):
    """Convert a numeric string like 3,218,545 to English words."""
    ones = ["","One","Two","Three","Four","Five","Six","Seven","Eight","Nine",
            "Ten","Eleven","Twelve","Thirteen","Fourteen","Fifteen","Sixteen",
            "Seventeen","Eighteen","Nineteen"]
    tens = ["","","Twenty","Thirty","Forty","Fifty","Sixty","Seventy","Eighty","Ninety"]
    def say(n):
        if n == 0: return ""
        elif n < 20: return ones[n]
        elif n < 100: return tens[n//10] + (" " + ones[n%10] if n%10 else "")
        elif n < 1000:
            return ones[n//100] + " Hundred" + (" " + say(n%100) if n%100 else "")
        elif n < 1_000_000:
            return say(n//1000) + " Thousand" + (" " + say(n%1000) if n%1000 else "")
        elif n < 1_000_000_000:
            return say(n//1_000_000) + " Million" + (" " + say(n%1_000_000) if n%1_000_000 else "")
        else:
            return say(n//1_000_000_000) + " Billion" + (" " + say(n%1_000_000_000) if n%1_000_000_000 else "")
    try:
        num = int(n_str.replace(",","").replace(" ","").split(".")[0])
        return say(num).strip()
    except:
        return n_str

def hl_runs(para):
    return [r for r in para.runs if r.font.highlight_color is not None]

def fill(run, text):
    run.text = text
    run.font.highlight_color = None

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
        for ci, val in enumerate([item.get("no", str(ri+1)), item.get("desc",""),
                                   item.get("qty",""), item.get("total","")]):
            if ci < len(row.cells):
                p = row.cells[ci].paragraphs[0]
                if p.runs:
                    p.runs[0].text = val
                else:
                    r2 = p.add_run(val)
                    r2.font.name = "Arial"
                    r2.font.size = Pt(9)

# ── Step 1: Offer Type ────────────────────────────────────────────────────────
st.header("Step 1 — Offer Type")
offer_type = st.radio("Select offer type:", ["Firm", "Budgetary"], horizontal=True)
is_firm = offer_type == "Firm"

st.markdown("---")
st.header("Step 2 — Customer Details")

col1, col2 = st.columns(2)

with col1:
    country_options = ["UAE","Saudi Arabia","Algeria","Yemen","Kuwait","Oman","Qatar","Other"]
    country = st.selectbox("Country", country_options)
    if country == "Other":
        country = st.text_input("Enter country name")

    customer_type = st.selectbox("Customer Type", [
        "National (UAE-based customer)",
        "International (non-UAE customer)",
        "Siemens Energy",
        "Critical Country"
    ])

    company_options = [
        "Electro Mechanical Co. LLC (ELMEC)",
        "ADNOC",
        "Nofal for Trade & Agencies",
        "Siemens Energy LLC",
        "Other"
    ]
    customer_company_sel = st.selectbox("Customer Company", company_options)
    if customer_company_sel == "Other":
        customer_company = st.text_input("Enter company name")
    else:
        customer_company = customer_company_sel

    customer_attn = st.text_input("Attention (e.g. Mr. John Smith)")
    customer_po_box = st.text_input("P.O. Box (or leave blank)")
    customer_tel = st.text_input("Tel (or leave blank)")

with col2:
    customer_fax = st.text_input("Fax (or leave blank)")
    if not is_firm:
        customer_mob = st.text_input("Mobile (or leave blank)")
    else:
        customer_mob = ""

    reference_no = st.text_input("Reference No. (or leave blank)")
    offer_date = st.text_input("Offer Date (e.g. 15 April 2025)")
    subject = st.text_input("Subject (e.g. 33KV Switchgear Supply)")
    project_name = st.text_input("Project Name")
    end_user = st.text_input("End User")

st.markdown("---")
st.header("Step 3 — Commercial Terms")

col3, col4 = st.columns(2)

with col3:
    currency_map = {
        "EUR - Euro": ("EUR","Euro"),
        "USD - US Dollar": ("USD","US Dollar"),
        "GBP - British Pound": ("GBP","British Pound"),
        "AED - UAE Dirham": ("AED","UAE Dirham"),
        "SAR - Saudi Riyal": ("SAR","Saudi Riyal"),
    }
    currency_sel = st.selectbox("Currency", list(currency_map.keys()))
    currency_code, currency_full = currency_map[currency_sel]

    total_price_num = st.text_input("Total Price (e.g. 3,218,545)")

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

    delivery_options = ["06","09","12","15","18","24","Other"]
    delivery_sel = st.selectbox("Delivery (months)", delivery_options)
    if delivery_sel == "Other":
        delivery_months = st.text_input("Enter number of months")
    else:
        delivery_months = delivery_sel

with col4:
    warranty_options = ["12","18","24","Other"]
    warranty_sel = st.selectbox("Warranty (months)", warranty_options)
    if warranty_sel == "Other":
        warranty_months = st.text_input("Enter warranty months")
    else:
        warranty_months = warranty_sel

    payment_option = st.selectbox("Payment Option", [
        "Option A - 20% down payment + 80% against shipping documents",
        "Option B - 20% advance + 10% drawings + 20% MFC + 50% bill of lading",
    ])
    payment_option_key = "A" if "Option A" in payment_option else "B"

    payment_days_options = ["30","45","60","90","120","Other"]
    payment_days_sel = st.selectbox("Payment Days", payment_days_options)
    if payment_days_sel == "Other":
        payment_days = st.text_input("Enter number of days")
    else:
        payment_days = payment_days_sel

st.markdown("---")
st.header("Step 4 — Sales Contact")

sales_contacts = {
    "Ahmad Awny | RC-AE SI EA S VD-V-D | +971 55 2003541 | ahmad.awny@siemens.com": {
        "name":"Ahmad Awny","dept":"RC-AE SI EA S VD-V-D","mob":"+971 55 2003541","email":"ahmad.awny@siemens.com"},
    "Vivek Gopalakrishnan | RC-AE SI EA S | +971 55 2002583 | vivek.gopalakrishnan@siemens.com": {
        "name":"Vivek Gopalakrishnan","dept":"RC-AE SI EA S","mob":"+971 55 2002583","email":"vivek.gopalakrishnan@siemens.com"},
    "Madiha Khan | RC-AE SI EA S | +971 55 2003818 | madiha.khan@siemens.com": {
        "name":"Madiha Khan","dept":"RC-AE SI EA S","mob":"+971 55 2003818","email":"madiha.khan@siemens.com"},
    "Hyun-Sik Kim | RC-AE SI EA S VD-V-D | +971 55 2002693 | hyun-sik.kim@siemens.com": {
        "name":"Hyun-Sik Kim","dept":"RC-AE SI EA S VD-V-D","mob":"+971 55 2002693","email":"hyun-sik.kim@siemens.com"},
    "Other": {"name":"","dept":"","mob":"","email":""},
}
contact_sel = st.selectbox("Sales Contact", list(sales_contacts.keys()))
contact = sales_contacts[contact_sel]
if contact_sel == "Other":
    c1,c2,c3,c4 = st.columns(4)
    sender_name  = c1.text_input("Name")
    sender_dept  = c2.text_input("Department")
    sender_mob   = c3.text_input("Mobile")
    sender_email = c4.text_input("Email")
else:
    sender_name  = contact["name"]
    sender_dept  = contact["dept"]
    sender_mob   = contact["mob"]
    sender_email = contact["email"]

# ── Firm-only fields ──────────────────────────────────────────────────────────
if is_firm:
    st.markdown("---")
    st.header("Step 5 — Firm Offer Details")

    col5, col6 = st.columns(2)
    with col5:
        install_options = {
            "UAE - Abu Dhabi (Dubai or Northern Emirates)": "the Emirate of Abu Dhabi (Dubai or Northern Emirates)",
            "UAE - Dubai / Jebel Ali": "the Emirate of Dubai / Jebel Ali",
            "Saudi Arabia": "a Saudi Arabian seaport",
            "Algeria": "an Algerian seaport",
            "Yemen": "a Yemeni seaport",
            "Other": ""
        }
        install_sel = st.selectbox("Country of Installation", list(install_options.keys()))
        if install_sel == "Other":
            import_port = st.text_input("Enter country / port")
            country_of_install = import_port
        else:
            import_port = install_options[install_sel]
            country_of_install = install_sel

        mfc_date = st.text_input("MFC Date (e.g. 26th February 2026)")

    with col6:
        offer_validity = st.text_input("Offer Validity (e.g. March 30, 2026)")

        signatories_options = {
            "Robert Hennig + Rodrigo Fernandes": "a",
            "Robert Hennig + Rodrigo Fernandes + Prashanth Parameswaran + Paul Fairweather": "b",
            "Robert Hennig + Rodrigo Fernandes + Prashanth Parameswaran + Paul Fairweather + Nasir Merchant": "c",
            "Other": "d"
        }
        signatories_sel = st.selectbox("Signatories", list(signatories_options.keys()))
        if signatories_sel == "Other":
            signatories = st.text_input("Enter signatory names")
        else:
            signatories = signatories_sel
else:
    import_port = ""
    country_of_install = ""
    mfc_date = ""
    offer_validity = ""
    signatories = "N/A"

# ── Scope Items ───────────────────────────────────────────────────────────────
st.markdown("---")
st.header("Step 6 — Scope of Supply")

st.markdown("Enter each scope item below:")
num_scope = st.number_input("Number of scope items", min_value=1, max_value=20, value=1, step=1)
scope_items = []
for i in range(int(num_scope)):
    st.markdown(f"**Item {i+1}**")
    sc1, sc2, sc3, sc4 = st.columns([1,4,2,2])
    no   = sc1.text_input("No", value=str(i+1), key=f"sc_no_{i}")
    desc = sc2.text_input("Description", key=f"sc_desc_{i}")
    qty  = sc3.text_input("Qty", key=f"sc_qty_{i}")
    tot  = sc4.text_input("Total Price", key=f"sc_tot_{i}")
    scope_items.append({"no":no,"desc":desc,"qty":qty,"total":tot})

st.markdown("**Optional Items** (leave blank if none)")
num_opt = st.number_input("Number of optional items", min_value=0, max_value=10, value=0, step=1)
optional_items = []
for i in range(int(num_opt)):
    st.markdown(f"**Optional Item {i+1}**")
    oc1, oc2, oc3, oc4 = st.columns([1,4,2,2])
    no   = oc1.text_input("No", value=str(i+1), key=f"opt_no_{i}")
    desc = oc2.text_input("Description", key=f"opt_desc_{i}")
    qty  = oc3.text_input("Qty", key=f"opt_qty_{i}")
    tot  = oc4.text_input("Total Price", key=f"opt_tot_{i}")
    optional_items.append({"no":no,"desc":desc,"qty":qty,"total":tot})

# ── Generate Button ───────────────────────────────────────────────────────────
st.markdown("---")
if st.button("🚀 Generate Offer Letter"):

    # Derive values
    total_price_words = num_to_words(total_price_num)
    is_national = "National" in customer_type

    # Payment text
    if payment_option_key == "A":
        pay_header = f"Payment terms: {payment_days} days from shipping documents"
        pay_lines  = (
            f"20% of the contract value to be paid as down payment within 15 days after placement of the order.\n"
            f"80% of the contract value payable against presentation of shipping documents within {payment_days} days."
        )
    else:
        pay_header = f"Payment terms: {payment_days} days from shipping documents"
        pay_lines  = (
            f"20% of the contract value to be paid as Advance Payment within 15 days after Order Confirmation.\n"
            f"10% of the contract value against submission of Drawings (SLD), payable within {payment_days} days.\n"
            f"20% of the contract value upon Manufacturing Clearance, payable within 45 days.\n"
            f"50% of the contract value against Bill of Lading, payable within {payment_days} days."
        )

    import_port_sentence = (
        f"The scope shall be offloaded in and imported via a port of {import_port}."
        if is_firm and import_port else ""
    )

    # Customer city
    customer_city = f"{country}"

    replacements = {
        "INSERT_CUSTOMER_COMPANY":      customer_company,
        "INSERT_CUSTOMER_PO_BOX":       customer_po_box,
        "INSERT_CUSTOMER_CITY":         customer_city,
        "INSERT_CUSTOMER_ATTN":         f"Kind Attn: {customer_attn}",
        "INSERT_CUSTOMER_TEL":          customer_tel,
        "INSERT_CUSTOMER_FAX":          customer_fax,
        "INSERT_CUSTOMER_MOB":          customer_mob,
        "INSERT_SENDER_NAME":           sender_name,
        "INSERT_SENDER_DEPT":           sender_dept,
        "INSERT_SENDER_MOBILE":         sender_mob,
        "INSERT_SENDER_EMAIL":          sender_email,
        "INSERT_OFFER_DATE":            offer_date,
        "INSERT_REFERENCE_NO":          reference_no,
        "INSERT_SUBJECT":               subject,
        "INSERT_PROJECT_NAME":          project_name,
        "INSERT_END_USER":              end_user,
        "INSERT_CURRENCY_FULL":         currency_full,
        "INSERT_TOTAL_PRICE_NUM":       total_price_num,
        "INSERT_CURRENCY_CODE":         currency_code,
        "INSERT_TOTAL_PRICE_WORDS":     f"{currency_code} {total_price_words} Only).",
        "INSERT_INCOTERM_NAME":         incoterm_name,
        "INSERT_DELIVERY_MONTHS":       f"{delivery_months} month(s)",
        "INSERT_WARRANTY_MONTHS":       f"{warranty_months} months",
        "INSERT_PAYMENT_DAYS":          payment_days,
        "INSERT_PAYMENT_OPTION_HEADER": pay_header,
        "INSERT_PAYMENT_OPTION_LINES":  pay_lines,
        "INSERT_MFC_DATE":              mfc_date if is_firm else "",
        "INSERT_IMPORT_PORT_SENTENCE":  import_port_sentence,
        "INSERT_OFFER_VALIDITY":        offer_validity if is_firm else "",
    }

    import os

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    TEMPLATE_PATH = os.path.join(BASE_DIR, 'tmeplate_-Firm.docx') if is_firm else os.path.join(BASE_DIR, 'tempate_-Budgetary.docx')

    offer_short = "FIRM" if is_firm else "BUD"
    cust_short  = customer_company.split()[0].upper().strip(".,")
    SKIP = {"FOR","THE","OF","AND","IN","A","AN","PKG","WORKS","-","PROJECT"}
    proj_short  = "".join(w for w in project_name.split() if w.upper() not in SKIP)[:12].upper()
    try:    date_str = datetime.strptime(offer_date.strip(),"%d %B %Y").strftime("%Y%m%d")
    except: date_str = offer_date.replace(" ","")[:8]
    filename = f"{offer_short}_{cust_short}_{proj_short}_{date_str}_v01.docx"
    filepath = os.path.join(BASE_DIR, filename)

    shutil.copy(TEMPLATE_PATH, filepath)
    doc = Document(filepath)

    for para in all_paras(doc):
        hrs = hl_runs(para)
        if not hrs:
            continue
        for i in range(len(hrs)-1):
            combined = (hrs[i].text or "") + (hrs[i+1].text or "")
            if combined.strip() in replacements:
                hrs[i].text = combined
                hrs[i+1].text = ""
                hrs[i+1].font.highlight_color = None
        for r in hrs:
            key = (r.text or "").strip()
            if key in replacements:
                fill(r, replacements[key])

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
                    if r.text in ("{","}"):
                        r.text = ""
                break

        for para in all_paras(doc):
            if para.text.strip().startswith("The delivery period of the offered equipment is estimated to be"):
                for r in para.runs:
                    if r.bold and r.font.highlight_color is None:
                        if r.text and (r.text.strip()[:2].isdigit() or "month" in r.text):
                            r.text = f"{delivery_months} month(s)"
                            break
                    if r.text in ("{","}"):
                        r.text = ""
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
                    if r.text in ("{","}"):
                        r.text = ""
                break

    if is_firm and is_national:
        national_map = {"5%":"-5%","45%":"-45%","90%":"-80%","100%":"-100%"}
        for para in all_paras(doc):
            t = para.text.strip()
            if t in national_map:
                for r in para.runs:
                    if r.text.strip() in national_map:
                        r.text = r.text.replace(r.text.strip(), national_map[r.text.strip()])

    if len(doc.tables) > 1:
        fill_table(doc.tables[1], scope_items)
    if optional_items and len(doc.tables) > 2:
        fill_table(doc.tables[2], optional_items)

    # Post-generation cleanup
    for para in all_paras(doc):
        for r in para.runs:
            if r.text and "INSERT_" in r.text:
                for key, val in replacements.items():
                    r.text = r.text.replace(key, val)
                r.font.highlight_color = None

    doc.save(filepath)

    st.success(f"✅ Offer letter generated: **{filename}**")
    with open(filepath, "rb") as f:
        st.download_button(
            label="📥 Download Offer Letter",
            data=f,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
