"""
📧 Smart Terminal Assistant - Email Report Module (Fixed Sender)
---------------------------------------------------------------
Now automatically uses a preset sender email + app password.
Only asks for receiver email when sending.
"""

import os
import smtplib
import pandas as pd
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from rich.console import Console
from rich.prompt import Prompt

console = Console()

# ------------------------------------------------------------
# ✉️ CONFIGURATION - Set your Gmail credentials here once
# ------------------------------------------------------------
SENDER_EMAIL = "smartterminalst@gmail.com"           # 👉 Replace with your Gmail
SENDER_PASSWORD = "ynqd cdus bshd npbg"     # 👉 Replace with your 16-digit app password
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587


# ------------------------------------------------------------
# 🧾 Generate CSV Report (optional)
# ------------------------------------------------------------
def generate_report(folder: Path) -> Path:
    """Create a CSV report listing all files in the folder."""
    console.print(f"\n🧾 [bold cyan]Generating report for {folder}...[/bold cyan]")
    data = []

    for file in folder.iterdir():
        if file.is_file():
            info = {
                "File Name": file.name,
                "Size (KB)": round(file.stat().st_size / 1024, 2),
                "Modified": pd.to_datetime(file.stat().st_mtime, unit="s"),
            }
            data.append(info)

    if not data:
        console.print("[yellow]⚠️ No files found in this folder.[/yellow]")
        return None

    df = pd.DataFrame(data)
    report_path = folder / "file_report.csv"
    df.to_csv(report_path, index=False)
    console.print(f"✅ Report saved as [green]{report_path}[/green]\n")
    return report_path


# ------------------------------------------------------------
# 📂 Let user choose specific files
# ------------------------------------------------------------
def choose_files(folder: Path):
    """Display all files and let user select which ones to email."""
    files = [f for f in folder.iterdir() if f.is_file()]
    if not files:
        console.print("[yellow]⚠️ No files found in this folder.[/yellow]")
        return []

    console.print("\n📂 [bold cyan]Files in folder:[/bold cyan]")
    for i, f in enumerate(files, start=1):
        console.print(f"{i}. {f.name}")

    console.print("\n[dim]Example: 1,3,5 or 2-6 or all[/dim]")
    selection = Prompt.ask("Select files to send", default="all").strip().lower()

    selected = []
    try:
        if selection == "all":
            return files

        parts = [s.strip() for s in selection.split(",")]
        for p in parts:
            if "-" in p:
                start, end = map(int, p.split("-"))
                selected.extend(files[start - 1:end])
            else:
                idx = int(p)
                selected.append(files[idx - 1])
    except Exception:
        console.print("[red]⚠️ Invalid input. Selecting all files instead.[/red]")
        selected = files

    return selected


# ------------------------------------------------------------
# ✉️ Send selected files as attachments
# ------------------------------------------------------------
def send_selected_files(files, receiver_email):
    """Send multiple selected files as attachments."""
    if not files:
        console.print("[red]❌ No files selected.[/red]")
        return

    try:
        console.print(f"\n📧 [bold cyan]Preparing to send {len(files)} file(s)...[/bold cyan]")
        msg = MIMEMultipart()
        msg["From"] = SENDER_EMAIL
        msg["To"] = receiver_email
        msg["Subject"] = f"Smart Assistant - {len(files)} file(s) attached"
        body = f"Attached file(s): {[f.name for f in files]}"
        msg.attach(MIMEText(body, "plain"))

        for f in files:
            with open(f, "rb") as file:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={f.name}")
            msg.attach(part)
            console.print(f"✅ Attached: {f.name}")

        console.print("[yellow]Connecting to Gmail SMTP server...[/yellow]")
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)

        console.print(f"\n🎉 [bold green]Email sent successfully to {receiver_email}[/bold green]\n")

    except smtplib.SMTPAuthenticationError:
        console.print("[red]❌ Authentication failed.[/red]")
        console.print("[dim]Double-check your Gmail App Password or replace it in the config section.[/dim]")
    except Exception as e:
        console.print(f"[red]❌ Failed to send email: {e}[/red]")


# ------------------------------------------------------------
# 📧 Email Report Menu
# ------------------------------------------------------------
def email_menu():
    """Menu for generating reports and sending files."""
    while True:
        console.print("""
📧 [bold cyan]Email Report Generator[/bold cyan]
1. Generate CSV Report
2. Send Selected Files by Email
3. Back to Main Menu
""")

        choice = Prompt.ask("[bold green]Choose an option[/bold green]", choices=["1", "2", "3"])

        if choice == "1":
            folder = Path(Prompt.ask("Enter folder path", default=".")).resolve()
            if folder.exists():
                generate_report(folder)
            else:
                console.print("[red]❌ Folder not found.[/red]")

        elif choice == "2":
            folder = Path(Prompt.ask("Enter folder path", default=".")).resolve()
            if not folder.exists():
                console.print("[red]❌ Folder not found.[/red]")
                continue

            files = choose_files(folder)
            if not files:
                console.print("[red]No files selected.[/red]")
                continue

            receiver = Prompt.ask("Enter receiver email")
            send_selected_files(files, receiver)

        elif choice == "3":
            return

        Prompt.ask("\nPress Enter to continue...")
