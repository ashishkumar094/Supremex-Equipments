# import pandas as pd
# import smtplib
# import re
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.utils import formataddr, formatdate, make_msgid, getaddresses
# from pdf2image import convert_from_path
# import pytesseract
#
# # -------------------------------
# # CONFIGURATION
# # -------------------------------
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# POPPLER_PATH = r"C:\Users\ashis\Downloads\Release-25.12.0-0\poppler-25.12.0\Library\bin"
#
# MY_EMAIL = "collections2@supremexfire.com"
# APP_PASSWORD = "u8FdmPNSrFnh"
# FROM_NAME = "Kavita Patil"
#
# PDF_FILE = "mail.pdf"
# EXCEL_FILE = "Collection Data for trail.xlsx"
#
# # -------------------------------
# # READ PDF TEMPLATE
# # -------------------------------
# print("\n📄 Reading PDF Template...")
# pages = convert_from_path(PDF_FILE, 200, poppler_path=POPPLER_PATH)
#
# pdf_text = ""
# for page in pages:
#     pdf_text += pytesseract.image_to_string(page)
#
# sig_match = re.search(r"Thanks & Regards,(.*?)Notice\s*:", pdf_text, re.DOTALL)
# if sig_match:
#     sig_raw = sig_match.group(1).strip()
#     sig_lines = [l.strip() for l in sig_raw.split("\n") if l.strip()]
#     signature = "<br>".join(sig_lines[:5])
# else:
#     signature = "Collection Executive<br>SUPREMEX EQUIPMENTS"
#
# notice_match = re.search(r"Notice\s*:(.*?)(?=Thanks & Regards,|\Z)", pdf_text, re.DOTALL)
# notice = notice_match.group(1).strip().replace("\n", " ") if notice_match else ""
#
# print("✅ PDF Template Loaded")
#
# # -------------------------------
# # READ EXCEL DATA
# # -------------------------------
# print("\n📊 Reading Excel Data...")
# df = pd.read_excel(EXCEL_FILE)
# df.columns = df.columns.str.strip()
# df["customer_name"] = df["customer_name"].astype(str).str.strip()
#
# # Detect CC column automatically
# cc_column_name = None
# for col in df.columns:
#     if col.lower() == "cc email":
#         cc_column_name = col
#         break
#
# grouped = df.groupby("customer_name")
# client_data = []
#
# for customer, group in grouped:
#
#     # TO
#     raw_to = str(group["Email ID"].iloc[0])
#     to_header = [e.strip() for e in raw_to.split(",") if e.strip()]
#
#     # ✅ Collect ALL CC values for that client
#     #cc_header = []
#
#     if cc_column_name:
#         cc_header = []
#
#         if cc_column_name:
#             cc_values = group[cc_column_name].dropna().astype(str)
#             combined_cc_string = ",".join(cc_values.tolist())
#
#             parsed_cc = getaddresses([combined_cc_string])
#
#             seen = set()
#
#             for name, addr in parsed_cc:
#                 if addr and addr not in seen:
#                     seen.add(addr)
#                     if name:
#                         cc_header.append(f"{name} <{addr}>")
#                     else:
#                         cc_header.append(addr)
#
#     client_data.append({
#         "customer": customer,
#         "to_header": to_header,
#         "cc_header": cc_header,
#         "invoices": group.copy()
#     })
#
# # -------------------------------
# # CLIENT SELECTION
# # -------------------------------
# print("\n🔍 Select:")
# print("   1. Send to ALL clients")
# print("   2. Send to SPECIFIC client")
#
# choice = input("\nEnter (1/2): ").strip()
#
# if choice == "2":
#     print()
#     for i, n in enumerate(df["customer_name"].unique(), 1):
#         print(f"{i}. {n}")
#
#     name_input = input("\nEnter client name(s) separated by comma: ").strip().lower()
#     selected_names = [n.strip() for n in name_input.split(",")]
#
#     client_data = [
#         c for c in client_data
#         if c["customer"].strip().lower() in selected_names
#     ]
#
#     if not client_data:
#         print("❌ Client not found.")
#         exit()
#
# print(f"\n✅ Sending to {len(client_data)} client(s)")
# confirm = input("📨 Confirm? (yes/no): ").strip().lower()
# if confirm != "yes":
#     print("❌ Cancelled.")
#     exit()
#
# # -------------------------------
# # CONNECT TO SMTP
# # -------------------------------
# smtp = smtplib.SMTP_SSL("smtp.zoho.com", 465)
# smtp.login(MY_EMAIL, APP_PASSWORD)
# print("\n✅ Connected to Zoho Mail!\n")
#
# # -------------------------------
# # SEND EMAILS
# # -------------------------------
# for client in client_data:
#
#     customer = client["customer"]
#     to_header = client["to_header"]
#     cc_header = client["cc_header"]
#     invoices = client["invoices"]
#
#     total_balance = invoices["bcy_balance"].sum()
#
#     rows_html = ""
#     for _, inv in invoices.iterrows():
#         rows_html += f"""
#         <tr>
#             <td>{str(inv['date'])[:10]}</td>
#             <td>{inv['invoice_number']}</td>
#             <td>{customer}</td>
#             <td>{inv['invoice.CF.PO NO']}</td>
#             <td>{inv['invoice.CF.Payment T&C']}</td>
#             <td style="text-align:right;">₹ {inv['bcy_total']:,.2f}</td>
#         </tr>
#         """
#
#     html_body = f"""
#     <html>
#     <body style="font-family: Arial; font-size:14px;">
#         <p><strong>Dear {customer},</strong></p>
#         <p>Please confirm payment status of pending invoices.</p>
#
#         <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width:100%;">
#             <tr style="background:#f2f2f2;">
#                 <th>Date</th><th>Invoice</th><th>Customer</th>
#                 <th>PO</th><th>Terms</th><th>Total</th>
#             </tr>
#             {rows_html}
#         </table>
#
#         <p><strong>Total Outstanding: ₹ {total_balance:,.2f}</strong></p>
#         <br>
#         Thanks & Regards,<br>{signature}<br><br>
#         SUPREMEX EQUIPMENTS
#         <hr>
#         <small>Notice: {notice}</small>
#     </body>
#     </html>
#     """
#
#     msg = MIMEMultipart("alternative")
#     msg["From"] = formataddr((FROM_NAME, MY_EMAIL))
#     msg["To"] = ", ".join(to_header)
#
#     if cc_header:
#         msg["Cc"] = ", ".join(cc_header)
#
#     msg["Subject"] = "Regarding Outstanding Payment"
#     msg["Date"] = formatdate(localtime=True)
#     msg["Message-ID"] = make_msgid(domain="supremexfire.com")
#
#     plain_text = f"""Dear {customer},
#
# Please check pending invoices.
#
# Total Outstanding Balance: ₹ {total_balance:,.2f}
#
# Thanks & Regards
# SUPREMEX EQUIPMENTS"""
#
#     msg.attach(MIMEText(plain_text, "plain", "utf-8"))
#     msg.attach(MIMEText(html_body, "html", "utf-8"))
#
#     # SMTP envelope
#     all_addresses = to_header + cc_header
#     parsed = getaddresses(all_addresses)
#     recipients = [addr for name, addr in parsed]
#
#     smtp.sendmail(MY_EMAIL, recipients, msg.as_bytes())
#
#     print(f"✅ Sent → {customer}")
#     print(f"   To : {msg['To']}")
#     if cc_header:
#         print(f"   Cc : {msg['Cc']}")
#
# smtp.quit()
# print("\n🎉 All emails processed successfully!")




# import pandas as pd
# import smtplib
# import re
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.utils import formataddr, formatdate, make_msgid, getaddresses
# from pdf2image import convert_from_path
# import pytesseract
#
# # -------------------------------
# # CONFIGURATION
# # -------------------------------
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# POPPLER_PATH = r"C:\Users\ashis\Downloads\Release-25.12.0-0\poppler-25.12.0\Library\bin"
#
# MY_EMAIL = "collections2@supremexfire.com"
# APP_PASSWORD = "u8FdmPNSrFnh"
# FROM_NAME = "Kavita Patil"
#
# PDF_FILE = "mail.pdf"
# EXCEL_FILE = "Collection.xlsx"
#
# # -------------------------------
# # READ PDF TEMPLATE
# # -------------------------------
# print("\n📄 Reading PDF Template...")
# pages = convert_from_path(PDF_FILE, 200, poppler_path=POPPLER_PATH)
#
# pdf_text = ""
# for page in pages:
#     pdf_text += pytesseract.image_to_string(page)
#
# sig_match = re.search(r"Thanks & Regards,(.*?)Notice\s*:", pdf_text, re.DOTALL)
# if sig_match:
#     sig_raw = sig_match.group(1).strip()
#     sig_lines = [l.strip() for l in sig_raw.split("\n") if l.strip()]
#     signature = "<br>".join(sig_lines[:5])
# else:
#     signature = "Collection Executive<br>SUPREMEX EQUIPMENTS"
#
# notice_match = re.search(r"Notice\s*:(.*?)(?=Thanks & Regards,|\Z)", pdf_text, re.DOTALL)
# notice = notice_match.group(1).strip().replace("\n", " ") if notice_match else ""
#
# print("✅ PDF Template Loaded")
#
# # -------------------------------
# # READ EXCEL DATA
# # -------------------------------
# print("\n📊 Reading Excel Data...")
# df = pd.read_excel(EXCEL_FILE)
# df.columns = df.columns.str.strip()
# df["customer_name"] = df["customer_name"].fillna("").astype(str).str.strip()
# df = df[df["customer_name"] != ""]
#
# # Detect CC column
# cc_column_name = None
# for col in df.columns:
#     if col.lower() == "cc email":
#         cc_column_name = col
#         break
#
# grouped = df.groupby("customer_name")
# client_data = []
#
# for customer, group in grouped:
#
#     raw_to = str(group["Email ID"].iloc[0])
#     to_header = [e.strip() for e in raw_to.split(",") if e.strip()]
#
#     cc_header = []
#     if cc_column_name:
#         cc_values = group[cc_column_name].dropna().astype(str)
#         combined_cc_string = ",".join(cc_values.tolist())
#         parsed_cc = getaddresses([combined_cc_string])
#
#         seen = set()
#         for name, addr in parsed_cc:
#             if addr and addr not in seen:
#                 seen.add(addr)
#                 cc_header.append(f"{name} <{addr}>" if name else addr)
#
#     client_data.append({
#         "customer": customer,
#         "to_header": to_header,
#         "cc_header": cc_header,
#         "invoices": group.copy()
#     })
#
# # -------------------------------
# # CLIENT SELECTION
# # -------------------------------
# print("\n🔍 Select:")
# print("   1. Send to ALL clients")
# print("   2. Send to SPECIFIC client")
#
# choice = input("\nEnter (1/2): ").strip()
#
# if choice == "2":
#     print()
#     for i, n in enumerate(df["customer_name"].unique(), 1):
#         print(f"{i}. {n}")
#
#     name_input = input("\nEnter client name(s) separated by comma: ").strip().lower()
#     selected_names = [n.strip() for n in name_input.split(",")]
#
#     client_data = [
#         c for c in client_data
#         if c["customer"].strip().lower() in selected_names
#     ]
#
#     if not client_data:
#         print("❌ Client not found.")
#         exit()
#
# print(f"\n✅ Sending to {len(client_data)} client(s)")
# confirm = input("📨 Confirm? (yes/no): ").strip().lower()
# if confirm != "yes":
#     print("❌ Cancelled.")
#     exit()
#
# # -------------------------------
# # CONNECT TO SMTP
# # -------------------------------
# smtp = smtplib.SMTP_SSL("smtp.zoho.com", 465)
# smtp.login(MY_EMAIL, APP_PASSWORD)
# print("\n✅ Connected to Zoho Mail!\n")
#
# # -------------------------------
# # SEND EMAILS
# # -------------------------------
# for client in client_data:
#
#     customer = client["customer"]
#     to_header = client["to_header"]
#     cc_header = client["cc_header"]
#     invoices = client["invoices"]
#
#     total_balance = invoices["bcy_balance"].sum()
#
#     rows_html = ""
#     for _, inv in invoices.iterrows():
#         rows_html += f"""
#         <tr>
#             <td>{str(inv['date'])[:10]}</td>
#             <td>{inv['invoice_number']}</td>
#             <td>{customer}</td>
#             <td>{inv['invoice.CF.PO NO']}</td>
#             <td style="text-align:right;">₹ {inv['bcy_total']:,.2f}</td>
#             <td style="text-align:right;">₹ {inv['bcy_balance']:,.2f}</td>
#         </tr>
#         """
#
#     html_body = f"""
#     <html>
#     <body style="font-family: Arial, sans-serif; font-size:14px;">
#         <p><strong>Dear {customer},</strong></p>
#
#         <p>
#         Please find below the details of your pending invoices
#         & confirm on payment status as per agreed terms.
#         </p>
#
#         <table border="1" cellpadding="8" cellspacing="0"
#                style="border-collapse: collapse; width:100%;">
#             <tr style="background-color:#f2f2f2;">
#                 <th>Date</th>
#                 <th>Invoice Number</th>
#                 <th>Customer Name</th>
#                 <th>PO Number</th>
#                 <th>Total (INR)</th>
#                 <th>Outstanding (INR)</th>
#             </tr>
#             {rows_html}
#         </table>
#
#         <br>
#
#         <p style="font-size:16px;">
#             <strong>Total Outstanding Balance:
#             ₹ {total_balance:,.2f}</strong>
#         </p>
#
#         <br>
#
#         <p>
#         Thanks & Regards,<br>
#         {signature}<br><br>
#
#         <strong>SUPREMEX EQUIPMENTS</strong><br>
#         HO: D-2A, Ghatkopar Industrial Estate | LBS Marg | Mumbai - 400086<br>
#         Factory: Tarapur MIDC | Palghar - 401502<br>
#         Website: www.supremexfire.com
#         </p>
#
#         <hr>
#
#         <p style="font-size:11px; color:gray;">
#         Notice: {notice}
#         </p>
#
#     </body>
#     </html>
#     """
#
#     msg = MIMEMultipart("alternative")
#     msg["From"] = formataddr((FROM_NAME, MY_EMAIL))
#     msg["To"] = ", ".join(to_header)
#
#     if cc_header:
#         msg["Cc"] = ", ".join(cc_header)
#
#     msg["Subject"] = "Regarding Outstanding Payment"
#     msg["Date"] = formatdate(localtime=True)
#     msg["Message-ID"] = make_msgid(domain="supremexfire.com")
#
#     plain_text = f"""Dear {customer},
#
# Please check pending invoices.
#
# Total Outstanding Balance: ₹ {total_balance:,.2f}
#
# Thanks & Regards
# SUPREMEX EQUIPMENTS"""
#
#     msg.attach(MIMEText(plain_text, "plain", "utf-8"))
#     msg.attach(MIMEText(html_body, "html", "utf-8"))
#
#     # SMTP envelope
#     all_addresses = to_header + cc_header
#     parsed = getaddresses(all_addresses)
#     recipients = [addr for name, addr in parsed]
#
#     smtp.sendmail(MY_EMAIL, recipients, msg.as_bytes())
#
#     print(f"✅ Sent → {customer}")
#     print(f"   To : {msg['To']}")
#     if cc_header:
#         print(f"   Cc : {msg['Cc']}")
#
# smtp.quit()
# print("\n🎉 All emails processed successfully!")



import pandas as pd
import smtplib
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr, formatdate, make_msgid, getaddresses
from pdf2image import convert_from_path
import pytesseract

# -------------------------------
# CONFIGURATION
# -------------------------------
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH = r"C:\Users\ashis\Downloads\Release-25.12.0-0\poppler-25.12.0\Library\bin"

MY_EMAIL = "collections2@supremexfire.com"
APP_PASSWORD = "u8FdmPNSrFnh"
FROM_NAME = "Kavita Patil"

PDF_FILE = "mail.pdf"
EXCEL_FILE = "Collection.xlsx"

# -------------------------------
# READ PDF TEMPLATE
# -------------------------------
print("\n📄 Reading PDF Template...")
pages = convert_from_path(PDF_FILE, 200, poppler_path=POPPLER_PATH)

pdf_text = ""
for page in pages:
    pdf_text += pytesseract.image_to_string(page)

sig_match = re.search(r"Thanks & Regards,(.*?)Notice\s*:", pdf_text, re.DOTALL)
if sig_match:
    sig_raw = sig_match.group(1).strip()
    sig_lines = [l.strip() for l in sig_raw.split("\n") if l.strip()]
    signature = "<br>".join(sig_lines[:5])
else:
    signature = "Collection Executive<br>SUPREMEX EQUIPMENTS"

notice_match = re.search(r"Notice\s*:(.*?)(?=Thanks & Regards,|\Z)", pdf_text, re.DOTALL)
notice = notice_match.group(1).strip().replace("\n", " ") if notice_match else ""

print("✅ PDF Template Loaded")

# -------------------------------
# READ EXCEL DATA
# -------------------------------
print("\n📊 Reading Excel Data...")
df = pd.read_excel(EXCEL_FILE)
df.columns = df.columns.str.strip()
df["customer_name"] = df["customer_name"].fillna("").astype(str).str.strip()
df = df[df["customer_name"] != ""]

# Detect CC column
cc_column_name = None
for col in df.columns:
    if "cc" in col.lower():
        cc_column_name = col
        break

grouped = df.groupby("customer_name")
client_data = []

for customer, group in grouped:

    raw_to = str(group["Email ID"].iloc[0])
    to_header = [e.strip() for e in raw_to.split(",") if e.strip()]

    # -------------------------------
    # CLEAN CC PROPERLY (ZOHO SAFE)
    # -------------------------------
    cc_header = []

    if cc_column_name:
        cc_values = group[cc_column_name].dropna().astype(str)

        combined = ",".join(cc_values.tolist())

        # Remove line breaks and extra spaces
        combined = re.sub(r"\s+", " ", combined)

        # Remove angle brackets
        combined = combined.replace("<", "").replace(">", "")

        # Remove trailing commas
        combined = combined.strip().rstrip(",")

        # Split cleanly
        emails = [e.strip() for e in combined.split(",") if "@" in e]

        # Remove duplicates
        seen = set()
        for email in emails:
            if email not in seen:
                seen.add(email)
                cc_header.append(email)

    client_data.append({
        "customer": customer,
        "to_header": to_header,
        "cc_header": cc_header,
        "invoices": group.copy()
    })

# -------------------------------
# CLIENT SELECTION
# -------------------------------
print("\n🔍 Select:")
print("   1. Send to ALL clients")
print("   2. Send to SPECIFIC client")

choice = input("\nEnter (1/2): ").strip()

if choice == "2":
    print()
    for i, n in enumerate(df["customer_name"].unique(), 1):
        print(f"{i}. {n}")

    name_input = input("\nEnter client name(s) separated by comma: ").strip().lower()
    selected_names = [n.strip() for n in name_input.split(",")]

    client_data = [
        c for c in client_data
        if c["customer"].strip().lower() in selected_names
    ]

    if not client_data:
        print("❌ Client not found.")
        exit()

print(f"\n✅ Sending to {len(client_data)} client(s)")
confirm = input("📨 Confirm? (yes/no): ").strip().lower()
if confirm != "yes":
    print("❌ Cancelled.")
    exit()

# -------------------------------
# CONNECT TO SMTP
# -------------------------------
smtp = smtplib.SMTP_SSL("smtp.zoho.com", 465)
smtp.login(MY_EMAIL, APP_PASSWORD)
print("\n✅ Connected to Zoho Mail!\n")

# -------------------------------
# SEND EMAILS
# -------------------------------
for client in client_data:

    customer = client["customer"]
    to_header = client["to_header"]
    cc_header = client["cc_header"]
    invoices = client["invoices"]

    total_balance = invoices["bcy_balance"].sum()

    rows_html = ""
    for _, inv in invoices.iterrows():
        rows_html += f"""
        <tr>
            <td>{str(inv['date'])[:10]}</td>
            <td>{inv['invoice_number']}</td>
            <td>{customer}</td>
            <td>{inv['invoice.CF.PO NO']}</td>
            <td style="text-align:right;">₹ {inv['bcy_total']:,.2f}</td>
            <td style="text-align:right;">₹ {inv['bcy_balance']:,.2f}</td>
        </tr>
        """

    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; font-size:14px;">
        <p><strong>Dear {customer},</strong></p>

        <p>Please find below the details of your pending invoices.</p>

        <table border="1" cellpadding="8" cellspacing="0"
               style="border-collapse: collapse; width:100%;">
            <tr style="background-color:#f2f2f2;">
                <th>Date</th>
                <th>Invoice Number</th>
                <th>Customer Name</th>
                <th>PO Number</th>
                <th>Total (INR)</th>
                <th>Outstanding (INR)</th>
            </tr>
            {rows_html}
        </table>

        <br>

        <p style="font-size:16px;">
            <strong>Total Outstanding Balance:
            ₹ {total_balance:,.2f}</strong>
        </p>

        <br>

        <p>
        Thanks & Regards,<br>
        {signature}<br><br>

        <strong>SUPREMEX EQUIPMENTS</strong><br>
        HO: D-2A, Ghatkopar Industrial Estate | LBS Marg | Mumbai - 400086<br>
        Factory: Tarapur MIDC | Palghar - 401502<br>
        Website: www.supremexfire.com
        </p>

        <hr>

        <p style="font-size:11px; color:gray;">
        Notice: {notice}
        </p>

    </body>
    </html>
    """

    msg = MIMEMultipart("alternative")
    msg["From"] = formataddr((FROM_NAME, MY_EMAIL))
    msg["To"] = ", ".join(to_header)

    if cc_header:
        msg["Cc"] = ", ".join(cc_header)

    msg["Subject"] = "Regarding Outstanding Payment"
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="supremexfire.com")

    plain_text = f"""Dear {customer},

Please check pending invoices.

Total Outstanding Balance: ₹ {total_balance:,.2f}

Thanks & Regards
SUPREMEX EQUIPMENTS"""

    msg.attach(MIMEText(plain_text, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    all_addresses = to_header + cc_header
    parsed = getaddresses(all_addresses)
    recipients = [addr for name, addr in parsed]

    smtp.sendmail(MY_EMAIL, recipients, msg.as_bytes())

    print(f"✅ Sent → {customer}")
    print(f"   To : {msg['To']}")
    if cc_header:
        print(f"   Cc : {msg['Cc']}")

smtp.quit()
print("\n🎉 All emails processed successfully!")