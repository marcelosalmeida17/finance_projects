import win32com.client as win32
from datetime import date, timedelta
import os

#Creating blank email
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
namespace = outlook.GetNamespace("MAPI")

#Defining dates
today = date.today()

if today.weekday() == 0:
    yesterday = today-timedelta(days=3)
else:
    yesterday = today-timedelta(days=1)

yesterdays_date = yesterday.strftime("%Y%m%d")
todays_date = date.today().strftime("%m-%d-%y")
todays_date_EOD = date.today().strftime("%d/%m/%Y")

mtm_email_date = today.strftime("%d-%b-%Y").lower()
cash_email_date = yesterday.strftime("%d-%b-%Y").lower()
reconciliation_email_date = yesterday.strftime("%d.%m.%Y").lower()
EOD_date = today.strftime("%d%m%Y")

#Attachment Paths
attached_files = [
    f"./data/reports/reconciliation_breaks_{todays_date}.xls",
    f"./data/reports/mtm_breaks_{todays_date}.xls",
    f"./data/reports/cash_balances_{yesterdays_date}.xlsx"
]

for attachment in attached_files:
    if os.path.isfile(attachment):
        mail.Attachments.Add(os.path.normpath(attachment))
    else:
        print(f"Attachment Skipped: {attachment}")

#Finding Emails
def find_emails(mailbox_name, folder_name, subject_text):
    folder = namespace.Folders[mailbox_name].Folders["Inbox"].Folders[folder_name]
    items = folder.items
    items.Sort("[ReceivedTime]", True)

    subject_text = subject_text.replace("'", "''")
    filtered_items = items.Restrict(f'@SQL="urn:schemas:httpmail:subject" LIKE \'%{subject_text}%\'')
    
    if filtered_items.Count > 0:
        return filtered_items.Item(1)
    return None


MTM_email = find_emails(
    "SHARED_MAILBOX",
    "REPORTS_FOLDER",
    f"MTM Alerts - {mtm_email_date}"
)

Cashpool_email = find_emails(
    "SHARED_MAILBOX",
    "REPORTS_FOLDER",
    f"Cash Reconciliation - {cash_email_date}"
)

Reconciliation_email = find_emails(
    "SHARED_MAILBOX",
    "REPORTS_FOLDER",
    f"Reconciliation Breaks - {reconciliation_email_date}"
)

if MTM_email:
    mtm_text = MTM_email.HTMLBody
    start = mtm_text.find("Please find")
    cutoff = mtm_text.lower().find("kind regards")
    if start != -1 and cutoff != -1:
        mtm_text = mtm_text[start:cutoff]
    else:
        mtm_text = "MTM report not found"

if Cashpool_email:
    cashpool_text = Cashpool_email.HTMLBody
    start = cashpool_text.find("Below is the")
    cutoff = cashpool_text.lower().find("regards")
    if start != -1 and cutoff != -1:
        cashpool_text = cashpool_text[start:cutoff]
    else:
        cashpool_text = "Cashpool report not found"

if Reconciliation_email:
    reconciliation_text = Reconciliation_email.HTMLBody
    start = reconciliation_text.find("Below is the")
    cutoff = reconciliation_text.lower().find("regards")
    if start != -1 and cutoff != -1:
        reconciliation_text = reconciliation_text[start:cutoff]
    else:
        matisse_text = "Reconciliation breaks report not found"

#Email Details
mail.To = "team@example.com"
mail.Cc = "manager@example.com"
mail.Subject = f"Daily Cash EOD Report {todays_date_EOD}"

body_style = "color: #003366; font-size:14pt; font-weight:bold"

body = f"""
<p>Hi All,</p>

<p>Please find below and attached the latest cash End of Day Report.</p>

<p><span style = "{body_style}">MTM Breaks:</span></p>
<p>{mtm_text}</p>

<p><span style = "{body_style}">Cashpool Breaks & Balance:</span></p>
<p>{cashpool_text}</p>

<p><span style = "{body_style}">Reconciliation Breaks:</span></p>
<p>{reconciliation_text}</p>
"""

mail.HTMLBody = body

#Display Email on Screen
mail.Display()

#Save email
output_path = "./output"
os.makedirs(output_path, exist_ok=True)
output_file = os.path.join(output_path, f"Daily_Cash_EOD_Report_{EOD_date}.msg")
mail.SaveAs(output_file)
