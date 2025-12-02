import smtplib
import pandas as pd
from datetime import datetime as dt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox


# -----------------------------
# MAIN TKINTER WINDOW
# -----------------------------
root = tk.Tk()
root.title("Email Automation App")
root.geometry("650x600")
root.resizable(False, False)


# -----------------------------
# GLOBAL VARIABLES
# -----------------------------
df = None
csv_path = ""
attachment_path = ""
body_path = ""


# -----------------------------
# Browse CSV
# -----------------------------
def browse_csv():
    global csv_path, df

    csv_path = filedialog.askopenfilename(
        title="Select CSV File",
        filetypes=[("CSV Files", "*.csv")]
    )
    csv_entry.delete(0, tk.END)
    csv_entry.insert(0, csv_path)

    if csv_path:
        df = pd.read_csv(csv_path)
        columns = list(df.columns)

        # Update dropdowns dynamically
        name_dropdown["values"] = columns
        email_dropdown["values"] = columns
        date_dropdown["values"] = columns
        time_dropdown["values"] = columns


# -----------------------------
# Browse Attachment
# -----------------------------
def browse_attachment():
    global attachment_path
    attachment_path = filedialog.askopenfilename(
        title="Select Attachment",
        filetypes=[("All Files", "*.*")]
    )
    attach_entry.delete(0, tk.END)
    attach_entry.insert(0, attachment_path)


# -----------------------------
# Browse Mail Body
# -----------------------------
def browse_body():
    global body_path
    body_path = filedialog.askopenfilename(
        title="Select Mail Body Text File",
        filetypes=[("Text Files", "*.txt")]
    )
    body_entry.delete(0, tk.END)
    body_entry.insert(0, body_path)


# -----------------------------
# Email Sending Function
# -----------------------------
def send_email(to_address, subject, body, attachment=None):
    msg = MIMEMultipart()
    msg["From"] = email_id.get().strip()
    msg["To"] = to_address
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain", "utf-8"))

    # Attachment
    if attachment:
        with open(attachment, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{os.path.basename(attachment)}"'
        )
        msg.attach(part)

    with smtplib.SMTP("smtp.gmail.com") as connection:
        connection.starttls()
        connection.login(email_id.get().strip(), app_pass.get().strip())
        connection.sendmail(email_id.get().strip(), to_address, msg.as_string())


# -----------------------------
# Main Process Function
# -----------------------------
def start_process():
    global df

    if df is None:
        messagebox.showerror("Error", "CSV file not loaded!")
        return

    if not body_path:
        messagebox.showerror("Error", "Mail body file not selected!")
        return

    with open(body_path, "r", encoding="utf-8") as f:
        template = f.read()

    subject = subject_entry.get()

    now = dt.now()
    today_date = now.strftime("%d-%m-%Y")
    current_time = now.strftime("%H:%M")

    sent_count = 0

    for index, row in df.iterrows():
        name = str(row[name_dropdown.get()])
        email = str(row[email_dropdown.get()])
        date_val = str(row[date_dropdown.get()])
        time_val = str(row[time_dropdown.get()])

        if today_date == date_val and current_time == time_val:
            mail_body = template.replace("[Name]", name)

            send_email(
                to_address=email,
                subject=subject,
                body=mail_body,
                attachment=attachment_path if attachment_path else None
            )

            sent_count += 1

    messagebox.showinfo("Success", f"Emails Sent: {sent_count}")


# -----------------------------
# GUI LAYOUT
# -----------------------------
tk.Label(root, text="Email Automation System", font=("Arial", 18, "bold")).pack(pady=10)

frame = tk.Frame(root)
frame.pack(pady=10)

# Gmail ID
tk.Label(frame, text="Your Gmail ID:").grid(row=0, column=0, sticky="w")
email_id = tk.Entry(frame, width=40)
email_id.grid(row=0, column=1, pady=5)

# App Password
tk.Label(frame, text="Gmail App Password:").grid(row=1, column=0, sticky="w")
app_pass = tk.Entry(frame, width=40, show="*")
app_pass.grid(row=1, column=1, pady=5)

# CSV File
tk.Label(frame, text="CSV File:").grid(row=2, column=0, sticky="w")
csv_entry = tk.Entry(frame, width=40)
csv_entry.grid(row=2, column=1, pady=5)
tk.Button(frame, text="Browse", command=browse_csv).grid(row=2, column=2, padx=5)

# Attachment
tk.Label(frame, text="Attachment (Optional):").grid(row=3, column=0, sticky="w")
attach_entry = tk.Entry(frame, width=40)
attach_entry.grid(row=3, column=1, pady=5)
tk.Button(frame, text="Browse", command=browse_attachment).grid(row=3, column=2, padx=5)

# Mail Body
tk.Label(frame, text="Mail Body (.txt):").grid(row=4, column=0, sticky="w")
body_entry = tk.Entry(frame, width=40)
body_entry.grid(row=4, column=1, pady=5)
tk.Button(frame, text="Browse", command=browse_body).grid(row=4, column=2, padx=5)

# Subject
tk.Label(frame, text="Email Subject:").grid(row=5, column=0, sticky="w")
subject_entry = tk.Entry(frame, width=40)
subject_entry.grid(row=5, column=1, pady=5)

# Column selectors
tk.Label(frame, text="Name Column:").grid(row=6, column=0, sticky="w")
name_dropdown = ttk.Combobox(frame, width=37, state="readonly")
name_dropdown.grid(row=6, column=1, pady=5)

tk.Label(frame, text="Email Column:").grid(row=7, column=0, sticky="w")
email_dropdown = ttk.Combobox(frame, width=37, state="readonly")
email_dropdown.grid(row=7, column=1, pady=5)

tk.Label(frame, text="Date Column:").grid(row=8, column=0, sticky="w")
date_dropdown = ttk.Combobox(frame, width=37, state="readonly")
date_dropdown.grid(row=8, column=1, pady=5)

tk.Label(frame, text="Time Column:").grid(row=9, column=0, sticky="w")
time_dropdown = ttk.Combobox(frame, width=37, state="readonly")
time_dropdown.grid(row=9, column=1, pady=5)

# Start Button
tk.Button(root, text="START AUTOMATION", font=("Arial", 14, "bold"),
          bg="green", fg="white", width=25, command=start_process).pack(pady=20)

root.mainloop()
