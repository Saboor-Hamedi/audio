import pyttsx3
import docx
import pdfplumber
from tkinter import Tk, Button, filedialog, Label, Scale, StringVar, messagebox, Entry, OptionMenu, ttk
import threading

def extract_text_from_docx(text_path):
    try:
        doc = docx.Document(text_path)
        full_text = [paragraph.text for paragraph in doc.paragraphs]
        return '\n'.join(full_text)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to extract text from DOCX: {e}")
        return ""

def extract_text_from_pdf(text_path):
    try:
        with pdfplumber.open(text_path) as pdf:
            full_text = [page.extract_text() for page in pdf.pages]
            return '\n'.join(full_text)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to extract text from PDF: {e}")
        return ""

def extract_text_from_txt(text_path):
    try:
        with open(text_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to extract text from TXT: {e}")
        return ""

def text_to_audio(text, output_path, voice_id, rate, volume, pitch, progress_bar, convert_button):
    try:
        engine = pyttsx3.init()
        engine.setProperty('rate', rate)
        engine.setProperty('volume', volume)
        engine.setProperty('voice', voice_id)
        engine.setProperty('pitch', pitch)

        # Simulate progress update
        total_size = len(text)
        chunk_size = total_size // 100
        progress = 0

        def on_word(word):
            nonlocal progress
            progress += len(word)
            progress_percentage = min(100, (progress * 100) // total_size)
            progress_bar['value'] = progress_percentage
            root.update_idletasks()

        engine.connect('started-word', on_word)
        engine.save_to_file(text, output_path)
        engine.runAndWait()
        progress_bar['value'] = 100
        root.update_idletasks()
        messagebox.showinfo("Success", f'Converted {file_path} to {output_path} successfully!')
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert text to audio: {e}")
    finally:
        progress_bar['value'] = 0
        convert_button['state'] = 'normal'

def select_file():
    global file_path, extract_text
    file_types = [('All Files', '*.*'), ('Text Files', '*.txt'), ('PDF Files', '*.pdf'), ('Word Documents', '*.docx')]
    file_path = filedialog.askopenfilename(filetypes=file_types)
    file_path_entry.delete(0, 'end')
    file_path_entry.insert(0, file_path)
    if not file_path:
        messagebox.showerror("Error", "No file selected.")
        return

    extract_text = {
        'txt': extract_text_from_txt,
        'pdf': extract_text_from_pdf,
        'docx': extract_text_from_docx
    }.get(file_path.split('.')[-1], None)

    if not extract_text:
        messagebox.showerror("Error", "Unsupported file type.")
        file_path = None

def get_audio_settings():
    output_file = filedialog.asksaveasfilename(defaultextension=".mp3")
    return output_file

def convert_file():
    global file_path, extract_text, progress_bar, convert_button, voice_options, voice_var

    if not file_path or not extract_text:
        messagebox.showerror("Error", "No file selected or failed to extract text.")
        return

    output_file = get_audio_settings()
    if not output_file:
        return

    text = extract_text(file_path)
    if not text:
        messagebox.showinfo("Error", "Failed to extract text from the file.")
        return

    voice_id = voice_options[voice_var.get()]
    rate = rate_scale.get()
    volume = volume_scale.get()
    pitch = pitch_scale.get()

    convert_button['state'] = 'disabled'
    threading.Thread(target=text_to_audio, args=(text, output_file, voice_id, rate, volume, pitch, progress_bar, convert_button)).start()

def main():
    global file_path, extract_text, voice_var, rate_scale, volume_scale, pitch_scale, file_path_entry, progress_bar, convert_button, root, voice_options

    file_path = None
    extract_text = None

    engine = pyttsx3.init()
    voices = engine.getProperty('voices')

    voice_options = {voice.name: voice.id for voice in voices}

    root = Tk()
    root.title("Text-to-Speech Converter")
    root.geometry("800x500")
    root.resizable(False, False)

    # Create a frame for the form elements
    form_frame = ttk.Frame(root, padding="20")
    form_frame.pack(fill='both', expand=True)

    # File path entry
    file_path_label = Label(form_frame, text="File Path:", font=("Arial", 12, "bold"))
    file_path_label.grid(row=0, column=0, padx=10, pady=5, sticky='w')
    file_path_entry = Entry(form_frame, width=50, font=("Arial", 12))
    file_path_entry.grid(row=0, column=1, padx=10, pady=5, sticky='ew')
    upload_button = Button(form_frame, text="Upload File", command=select_file, font=("Arial", 12, "bold"), bg="#4CAF50", fg="white")
    upload_button.grid(row=0, column=2, padx=10, pady=5)

    # Progress Bar
    progress_bar = ttk.Progressbar(form_frame, orient='horizontal', length=500, mode='determinate')
    progress_bar.grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky='ew')

    # Voice selection
    voice_var = StringVar(root)
    voice_var.set(next(iter(voice_options.keys())))
    voice_label = Label(form_frame, text="Select Voice:", font=("Arial", 12, "bold"))
    voice_label.grid(row=2, column=0, padx=10, pady=5, sticky='w')
    voice_menu = OptionMenu(form_frame, voice_var, *voice_options.keys())
    voice_menu.grid(row=2, column=1, padx=10, pady=5, columnspan=2, sticky='ew')

    # Rate selection
    rate_label = Label(form_frame, text="Select Rate:", font=("Arial", 12, "bold"))
    rate_label.grid(row=3, column=0, padx=10, pady=5, sticky='w')
    rate_scale = Scale(form_frame, from_=100, to=200, orient='horizontal', length=400)
    rate_scale.set(150)
    rate_scale.grid(row=3, column=1, padx=10, pady=5, columnspan=2, sticky='ew')

    # Volume selection
    volume_label = Label(form_frame, text="Select Volume:", font=("Arial", 12, "bold"))
    volume_label.grid(row=4, column=0, padx=10, pady=5, sticky='w')
    volume_scale = Scale(form_frame, from_=0.0, to=1.0, resolution=0.1, orient='horizontal', length=400)
    volume_scale.set(1.0)
    volume_scale.grid(row=4, column=1, padx=10, pady=5, columnspan=2, sticky='ew')

    # Pitch selection
    pitch_label = Label(form_frame, text="Select Pitch:", font=("Arial", 12, "bold"))
    pitch_label.grid(row=5, column=0, padx=10, pady=5, sticky='w')
    pitch_scale = Scale(form_frame, from_=0.5, to=2.0, resolution=0.1, orient='horizontal', length=400)
    pitch_scale.set(1.0)
    pitch_scale.grid(row=5, column=1, padx=10, pady=5, columnspan=2, sticky='ew')

    convert_button = Button(form_frame, text="Convert File", command=convert_file, font=("Arial", 12, "bold"), bg="#4CAF50", fg="white")
    convert_button.grid(row=6, column=0, columnspan=3, padx=10, pady=20, sticky='ew')

    # Make the rows and columns resize dynamically
    for i in range(7):
        form_frame.rowconfigure(i, weight=1)
    for i in range(3):
        form_frame.columnconfigure(i, weight=1)

    root.mainloop()

if __name__ == '__main__':
    main()
