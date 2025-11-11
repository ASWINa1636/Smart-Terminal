"""
💻 Smart Terminal Automation Assistant (Offline Version)
--------------------------------------------------------
A command-line automation tool built with Python.
Performs real-world tasks like:
 - PDF Automation (merge/split)
 - File Organizer
 - Email Report Generator
 - Voice Assistant Mode (new!)
All offline, no API keys needed.
"""

import os
import sys
from rich.console import Console
from rich.prompt import Prompt

# Import local modules
from modules import pdf_tools, file_organizer, email_report, voice_assistant

console = Console()


def clear_screen():
    """Clear the terminal screen for a cleaner UI."""
    os.system('cls' if os.name == 'nt' else 'clear')


def main_menu():
    """Display the main menu."""
    while True:
        clear_screen()
        console.rule("[bold cyan]💻 Smart Terminal Automation Assistant[/bold cyan]")
        console.print("""
[bold yellow]Select an option:[/bold yellow]

1. PDF & Word Tools
2. File Organizer
3. Email Report Generator
4. 🎙️ Voice Assistant Mode
5. Exit
        """)

        choice = Prompt.ask("[bold green]Enter your choice[/bold green]", choices=["1", "2", "3", "4", "5"])

        if choice == "1":
            pdf_tools.pdf_menu()
        elif choice == "2":
            file_organizer.organizer_menu()
        elif choice == "3":
            email_report.email_menu()
        elif choice == "4":
            console.print("\n🎤 [bold cyan]Starting Voice Assistant Mode...[/bold cyan]\n")
            voice_assistant.start_voice_assistant()
        elif choice == "5":
            console.print("\n👋 [bold cyan]Goodbye![/bold cyan]")
            sys.exit(0)


if __name__ == "__main__":
    main_menu()
