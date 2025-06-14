from openpyxl import load_workbook
import shutil
import os

# List of emails to remove
emails_to_remove = [
    "heathbyerly761@gmail.com",
    "williams.jason251@gmail.com",
    "medmonds53@gmail.com",
    "iteachicoach7575@gmail.com",
    "crobertoholmes@gmail.com",
    "cristal00s@gmail.com",
    "shannantx@icloud.com",
    "0424hud@gmail.com",
    "ruby.campbell@att.net",
    "dademon062@icloud.com",
    "tlazauskas77@yahoo.com",
    "chuckhenderson8857@gmail.com",
    "ingrid2451@gmail.com",
    "taijemarshall22@gmail.com",
    "suburbanscott473@gmail.com",
    "catherinelaw228@gmail.com",
    "alissah503@gmail.com",
    "jessicastarks_49@yahoo.com",
    "amymoore0972@gmail.com",
    "poroletty2010@icloud.com",
    "jacquiece.toomer@yahoo.com",
    "thatrooper@icloud.com"


]

# Main function
def remove_emails_from_excel(file_path):
    if not os.path.exists(file_path):
        print("❌ File not found.")
        return

    # Backup original
    backup_path = file_path.replace(".xlsx", "_backup.xlsx")
    shutil.copy(file_path, backup_path)
    print(f"📦 Backup created: {backup_path}")

    # Normalize emails
    emails_to_remove_clean = [email.strip().lower() for email in emails_to_remove]

    # Load Excel
    wb = load_workbook(file_path)
    ws = wb.active

    # Find 'Email' column
    email_col_index = None
    for cell in ws[1]:  # Header row
        if str(cell.value).strip().lower() == "email":
            email_col_index = cell.column
            break

    if not email_col_index:
        print("❌ 'Email' column not found.")
        return

    # Scan for matching emails
    found_emails = []
    for row in range(2, ws.max_row + 1):
        cell_value = ws.cell(row=row, column=email_col_index).value
        if cell_value and cell_value.strip().lower() in emails_to_remove_clean:
            found_emails.append(cell_value.strip())

    # Show matched emails
    if not found_emails:
        print("⚠️ No matching emails found.")
        return

    print("🟡 Found the following email(s) in the file:")
    for email in found_emails:
        print(f"   ➤ {email}")

    input("🔔 Press Enter to remove these emails from the file...")

    try:  # Delete rows from bottom up
        rows_deleted = 0
        for row in range(ws.max_row, 1, -1):
            cell_value = ws.cell(row=row, column=email_col_index).value
            if cell_value and cell_value.strip().lower() in emails_to_remove_clean:
                ws.delete_rows(row)
                rows_deleted += 1
        wb.save(file_path)
    except PermissionError as e:
        print(f"❌ Error while deleting rows: {e}")
        return


    print(f"✅ Removed {rows_deleted} email(s) from '{file_path}':")
    for email in found_emails:
        print(f"   ✔️ {email}")

# 📁 Call the function
remove_emails_from_excel("jeki.xlsx")
