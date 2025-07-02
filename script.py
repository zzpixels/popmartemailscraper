import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
import imaplib
import email
import re
import csv
from bs4 import BeautifulSoup
from email.header import decode_header
from datetime import datetime, timedelta
from email.utils import parseaddr

ALLOWED_SENDERS = ["popmart"]
SUPPORTED_DOMAINS = {
    "@gmail.com": "imap.gmail.com",
    "@icloud.com": "imap.mail.me.com",
    "@me.com": "imap.mail.me.com",
    "@mac.com": "imap.mail.me.com"
}


def deduplicate_entries(entries):
    best_entries = {}
    for entry in entries:
        oid = entry["Order ID"]
        if oid not in best_entries:
            best_entries[oid] = entry
        elif best_entries[oid]["Tracking Numbers"] == "N/A" and entry["Tracking Numbers"] != "N/A":
            best_entries[oid] = entry
    return list(best_entries.values())

def export_to_csv(entries, email_user, initial_dir="."):
    fieldnames = [
        "To Email", "Order ID", "Order Date", "Tracking Numbers",
        "Subject", "Item Name", "Qty", "Total Price", "Ship To ZIP"
    ]
    sanitized_email = email_user.replace("@", "_at_").replace(".", "_")
    filename = f"popmart_report_{sanitized_email}.csv"
    filepath = os.path.join(initial_dir, filename)

    with open(filepath, mode="w", newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(entries)

    return filepath

def extract_rfc822(fetch_data):
    for part in fetch_data:
        if isinstance(part, tuple):
            return part[1]
    return None

def extract_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            disp = str(part.get("Content-Disposition") or "")
            if "attachment" in disp.lower():
                continue
            if content_type in ["text/plain", "text/html"]:
                try:
                    charset = part.get_content_charset() or "utf-8"
                    body += part.get_payload(decode=True).decode(charset, errors="ignore")
                except Exception:
                    pass
    else:
        try:
            charset = msg.get_content_charset() or "utf-8"
            body = msg.get_payload(decode=True).decode(charset, errors="ignore")
        except Exception:
            pass
    return body

def connect_and_fetch(email_user, email_pass, days_back=21, batch_size=50):
    imap_server = None
    for domain, server in SUPPORTED_DOMAINS.items():
        if domain in email_user:
            imap_server = server
            break
    if not imap_server:
        raise ValueError("Unsupported email domain")

    is_icloud = any(d in email_user for d in ["@icloud.com", "@me.com", "@mac.com"])

    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(email_user, email_pass)
    mail.select("inbox")

    since_date = (datetime.today() - timedelta(days=days_back)).strftime("%d-%b-%Y")
    status, messages = mail.search(None, f'(SINCE {since_date})')
    email_ids = messages[0].split()

    raw_msgs = []

    def fetch_email_bytes(eid_batch):
        if is_icloud:
            for eid in eid_batch:
                status, data = mail.fetch(eid, "(BODY.PEEK[])")
                if status == "OK" and data and isinstance(data[0], tuple):
                    yield data[0][1]
        else:
            id_string = ",".join(e.decode() for e in eid_batch)
            status, fetched_data = mail.fetch(id_string, "(RFC822)")
            for part in fetched_data:
                raw = extract_rfc822([part])
                if raw:
                    yield raw

    for i in range(0, len(email_ids), batch_size):
        batch_ids = email_ids[i:i + batch_size]
        raw_msgs.extend(fetch_email_bytes(batch_ids))

    mail.logout()

    order_confirmations = {}
    results = []

    def parse_email(raw_msg_bytes):
        try:
            msg = email.message_from_bytes(raw_msg_bytes)
            from_name, from_email = parseaddr(msg.get("From"))
            if not any(sender in (from_name or '').lower() or sender in (from_email or '').lower() for sender in ALLOWED_SENDERS):
                return None, False, False

            subject_raw = msg.get("Subject", "")
            decoded_subject = decode_header(subject_raw)[0][0]
            subject = decoded_subject.decode(errors="ignore") if isinstance(decoded_subject, bytes) else decoded_subject

            match = re.search(r"#(O\d+)", subject, re.IGNORECASE)
            if not match:
                return None, False, False
            order_id = match.group(1)

            to_email = parseaddr(msg.get("To"))[1]
            body = extract_body(msg)
            soup = BeautifulSoup(body, "html.parser")

            tracking_matches = re.findall(r"(YT\d{16}|1Z[0-9A-Z]{10,22})", body, re.IGNORECASE)

            item_name = qty = total_price = order_date = zip_code = ""
            item_row = soup.find("ul", class_="table-list")
            if item_row:
                li = item_row.find("li")
                if li:
                    spans = li.find_all("span")
                    if len(spans) >= 4:
                        item_name = spans[1].get_text(strip=True)
                        qty = spans[3].get_text(strip=True)

            address_block = soup.find(string=re.compile(r"Order number: #O\d+"))
            if address_block:
                parent = address_block.find_parent()
                if parent:
                    for line in parent.get_text(separator="\n").splitlines():
                        m = re.search(r"\b\d{5}\b", line)
                        if m:
                            zip_code = m.group(0)
                            break

            for label in soup.find_all("span"):
                if "TOTAL" in label.get_text(strip=True).upper():
                    sibling = label.find_next("span")
                    if sibling:
                        total_price = sibling.get_text(strip=True)
                        if not total_price.startswith("$"):
                            total_price = f"${total_price}"
                        break

            date_match = re.search(r'Date:\s*([A-Za-z]{3,}\s+\d{1,2},\s*\d{4})', body)
            if date_match:
                order_date = date_match.group(1)

            is_shipping = "shipment from order" in subject.lower() or "on the way" in subject.lower()
            is_confirmation = "order #" in subject.lower() and "confirmed" in subject.lower()

            entry = {
                "To Email": to_email,
                "Order ID": order_id,
                "Order Date": order_date,
                "Tracking Numbers": ", ".join(tracking_matches) if tracking_matches else "N/A",
                "Subject": subject,
                "Item Name": item_name,
                "Qty": qty,
                "Total Price": total_price,
                "Ship To ZIP": zip_code
            }
            return entry, is_confirmation, is_shipping
        except Exception:
            return None, False, False

    all_entries = []
    for raw in raw_msgs:
        entry, is_conf, _ = parse_email(raw)
        if entry:
            if is_conf:
                order_confirmations[entry["Order ID"]] = entry
            all_entries.append(entry)

    for entry in all_entries:
        oid = entry["Order ID"]
        if oid in order_confirmations:
            for field in ["Item Name", "Qty", "Total Price", "Order Date", "Ship To ZIP"]:
                if not entry.get(field):
                    entry[field] = order_confirmations[oid].get(field)
        results.append(entry)

    return deduplicate_entries(results)


class PopmartGUI(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        parent.title("Popmart Email Scanner")
        parent.geometry("1100x600")
        self.pack(fill="both", expand=True)

        input_frame = ttk.LabelFrame(self, text="Credentials & Settings")
        input_frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(input_frame, text="Email:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.email_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.email_var, width=40).grid(row=0, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(input_frame, text="App Password:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.pass_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.pass_var, show="*", width=40).grid(row=1, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(input_frame, text="Days Back:").grid(row=0, column=2, sticky="e", padx=5, pady=5)
        self.days_var = tk.IntVar(value=21)
        ttk.Spinbox(input_frame, from_=1, to=60, textvariable=self.days_var, width=5).grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(input_frame, text="Batch Size:").grid(row=1, column=2, sticky="e", padx=5, pady=5)
        self.batch_var = tk.IntVar(value=50)
        ttk.Spinbox(input_frame, from_=10, to=200, increment=10, textvariable=self.batch_var, width=5).grid(row=1, column=3, padx=5, pady=5)

        self.scan_btn = ttk.Button(input_frame, text="Start Scan", command=self.start_scan)
        self.scan_btn.grid(row=0, column=4, rowspan=2, padx=10)

        self.status_var = tk.StringVar(value="Idle")
        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.status_label = ttk.Label(self, textvariable=self.status_var)
        self.status_label.pack(fill="x", padx=10)
        self.progress.pack(fill="x", padx=10)

        columns = ("To Email", "Order ID", "Order Date", "Tracking Numbers", "Item Name", "Qty", "Total Price", "Ship To ZIP")
        self.tree = ttk.Treeview(self, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130 if col != "Item Name" else 220)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        self.tree.pack(fill="both", expand=True, padx=10, pady=(5, 0))
        vsb.place(in_=self.tree, relx=1.0, rely=0, relheight=1.0, anchor="ne")
        hsb.pack(fill="x", padx=10, pady=(0, 10))

        self.export_btn = ttk.Button(self, text="Export to CSV", command=self.export_csv, state="disabled")
        self.export_btn.pack(pady=(0, 10))

        self.entries = []


    def start_scan(self):
        email_user = self.email_var.get().strip()
        email_pass = self.pass_var.get().strip()
        if not email_user or not email_pass:
            messagebox.showerror("Input Error", "Email and password are required.")
            return

        self.scan_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        self.status_var.set("Scanning emails… This may take a moment.")
        self.progress.start(10)
        self.tree.delete(*self.tree.get_children())

        threading.Thread(target=self._scan_worker, args=(email_user, email_pass, self.days_var.get(), self.batch_var.get()), daemon=True).start()

    def _scan_worker(self, email_user, email_pass, days_back, batch_size):
        try:
            results = connect_and_fetch(email_user, email_pass, days_back, batch_size)
            self.entries = results
            self.parent.after(0, self._on_scan_complete, None)
        except Exception as exc:
            self.parent.after(0, self._on_scan_complete, exc)

    def _on_scan_complete(self, error):
        self.progress.stop()
        self.scan_btn.config(state="normal")
        if error:
            self.status_var.set("Error: " + str(error))
            messagebox.showerror("Scan Failed", str(error))
            return
        self.status_var.set(f"Scan complete. {len(self.entries)} entries found.")
        for entry in self.entries:
            values = (
                entry["To Email"], entry["Order ID"], entry["Order Date"],
                entry["Tracking Numbers"], entry["Item Name"], entry["Qty"],
                entry["Total Price"], entry["Ship To ZIP"]
            )
            self.tree.insert("", "end", values=values)
        if self.entries:
            self.export_btn.config(state="normal")

    def export_csv(self):
        if not self.entries:
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=f"popmart_report.csv"
        )
        if file_path:
            export_to_csv(self.entries, self.email_var.get(), initial_dir=os.path.dirname(file_path))
            messagebox.showinfo("Export Complete", f"CSV saved to {file_path}")


if __name__ == "__main__":
    root = tk.Tk()

    style = ttk.Style()
    style.theme_use("clam")

    app = PopmartGUI(root)
    root.mainloop()
