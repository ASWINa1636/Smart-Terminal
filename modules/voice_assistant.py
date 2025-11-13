import os, sys, warnings
try:
    devnull_fd = os.open(os.devnull, os.O_WRONLY)
    os.dup2(devnull_fd, 2)  # redirect FD 2 (stderr)
except Exception:
    pass

os.environ['PYGAME_HIDE_SUPPORT_PROMPT'] = '1'
os.environ["AUDIODEV"] = "null"
warnings.filterwarnings("ignore")


import speech_recognition as sr
from gtts import gTTS
import tempfile, time
try:
    import vlc
except Exception:
    vlc = None
from rich.console import Console
from modules import pdf_tools

console = Console()


def speak(text: str):
    console.print(f"[bold cyan]{text}[/bold cyan]")
    try:
        tts = gTTS(text=text, lang='en', slow=False)
        with tempfile.NamedTemporaryFile(delete=True, suffix=".mp3") as tmp:
            tts.save(tmp.name)
            if vlc:
                player = vlc.MediaPlayer(tmp.name)
                player.play()
                time.sleep(0.5)
                while player.is_playing():
                    time.sleep(0.1)
            else:
                console.print("[yellow]VLC not found — skipping audio output.[/yellow]")
    except Exception as e:
        console.print(f"[red]TTS error:[/red] {e}")

def listen():
    try:
        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            console.print("[dim]Listening...[/dim]")
            recognizer.adjust_for_ambient_noise(source, duration=0.5)
            audio = recognizer.listen(source)

        try:
            return recognizer.recognize_google(audio).lower()
        except sr.UnknownValueError:
            return ""
        except sr.RequestError:
            speak("Speech recognition service is unavailable.")
            return ""
    except Exception:
        speak("Microphone is unavailable.")
        return ""

def start_voice_assistant():
    speak("Hello, I’m your Smart Terminal Voice Assistant!")
    speak("You can say commands like merge PDFs, split PDF, or exit to quit.")

    while True:
        command = listen()
        if not command:
            continue

        if "exit" in command or "quit" in command:
            speak("Goodbye!")
            break

        elif "merge pdf" in command:
            speak("Opening PDF merge tool.")
            pdf_tools.merge_pdfs()

        elif "split pdf" in command:
            speak("Opening PDF splitter.")
            pdf_tools.split_pdf()

        elif "protect pdf" in command:
            speak("Opening PDF protection tool.")
            pdf_tools.protect_pdf()

        elif "unlock pdf" in command:
            speak("Opening PDF unlock tool.")
            pdf_tools.unlock_pdf()

        elif "word to pdf" in command or "convert word" in command:
            speak("Converting Word document to PDF.")
            pdf_tools.word_to_pdf()

        elif "image to pdf" in command:
            speak("Converting image files to PDF.")
            pdf_tools.image_to_pdf()

        else:
            speak("Sorry, I don't have a command for that yet.")
