import os
import shutil
from datetime import datetime
from pathlib import Path
from rich.console import Console
from rich.prompt import Prompt

console = Console()

# ------------------------------------------------------------
# 🔹 Helper: make a safe directory
# ------------------------------------------------------------
def make_dir(base: Path, name: str) -> Path:
    path = base / name
    path.mkdir(parents=True, exist_ok=True)
    return path


# ------------------------------------------------------------
# 🧾 Helper: choose files interactively
# ------------------------------------------------------------
def choose_files(folder: Path):
    """Display files and let the user select one or more by index."""
    files = [f for f in folder.iterdir() if f.is_file()]
    if not files:
        console.print("[yellow]⚠️ No files found in this folder.[/yellow]")
        return []

    console.print("\n📂 Files in folder:")
    for i, f in enumerate(files, start=1):
        console.print(f"{i}. {f.name}")

    console.print("\n[dim]Example: 1,3,5 or 2-6 or all[/dim]")
    selection = Prompt.ask("Enter file numbers to sort", default="all").strip().lower()

    if selection == "all":
        return files

    selected = []
    try:
        parts = [s.strip() for s in selection.split(",")]
        for p in parts:
            if "-" in p:
                start, end = map(int, p.split("-"))
                selected.extend(files[start - 1:end])
            else:
                idx = int(p)
                selected.append(files[idx - 1])
    except Exception:
        console.print("[red]⚠️ Invalid input, selecting all files instead.[/red]")
        selected = files

    return selected


# ------------------------------------------------------------
# 🗂 Sort selected files by type
# ------------------------------------------------------------
def sort_by_type(folder: Path):
    selected_files = choose_files(folder)
    if not selected_files:
        return

    console.print(f"\n🗂 [bold cyan]Sorting selected files by type...[/bold cyan]")
    moved = 0
    for file in selected_files:
        ext = file.suffix.lower().strip(".") or "others"
        dest_dir = make_dir(folder, ext.capitalize())
        shutil.move(str(file), dest_dir / file.name)
        moved += 1
        console.print(f"✅ Moved {file.name} → {dest_dir.name}/")

    console.print(f"\n🎉 Done! Moved [bold]{moved}[/bold] files.\n")


# ------------------------------------------------------------
# 📅 Sort selected files by date
# ------------------------------------------------------------
def sort_by_date(folder: Path):
    selected_files = choose_files(folder)
    if not selected_files:
        return

    console.print(f"\n📅 [bold cyan]Sorting selected files by modified date...[/bold cyan]")
    moved = 0
    for file in selected_files:
        mtime = datetime.fromtimestamp(file.stat().st_mtime)
        subfolder = mtime.strftime("%Y-%m")
        dest_dir = make_dir(folder, subfolder)
        shutil.move(str(file), dest_dir / file.name)
        moved += 1
        console.print(f"✅ Moved {file.name} → {subfolder}/")

    console.print(f"\n🎉 Done! Moved [bold]{moved}[/bold] files.\n")


# ------------------------------------------------------------
# 📁 File Organizer Menu
# ------------------------------------------------------------
def organizer_menu():
    while True:
        console.print("""
🗂 [bold cyan]File Organizer Tools[/bold cyan]
1. Sort selected files by type
2. Sort selected files by date
3. Back to Main Menu
""")
        choice = Prompt.ask("[bold green]Choose an option[/bold green]", choices=["1", "2", "3"])

        if choice in ["1", "2"]:
            folder = Path(Prompt.ask("Enter folder path", default=".")).resolve()
            if not folder.exists():
                console.print("[red]❌ Folder not found.[/red]")
                continue

            if choice == "1":
                sort_by_type(folder)
            else:
                sort_by_date(folder)

        elif choice == "3":
            return

        Prompt.ask("\nPress Enter to continue...")
