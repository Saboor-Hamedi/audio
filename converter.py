import pyttsx3
import docx
import pdfplumber
from tkinter import Tk, Button, filedialog, Label, Scale, OptionMenu, StringVar, messagebox

def extract_text_from_docx(text_path):
    doc = docx.Document(text_path)
    full_text = [paragraph.text for paragraph in doc.paragraphs]
    return '\n'.join(full_text)

def extract_text_from_pdf(text_path):
    with pdfplumber.open(text_path) as pdf:
        full_text = [page.extract_text() for page in pdf.pages]
        return '\n'.join(full_text)

def extract_text_from_txt(text_path):
    with open(text_path, 'r', encoding='utf-8') as f:
        return f.read()

def text_to_audio(text, output_path, voice_id, rate, volume, pitch):
    engine = pyttsx3.init()
    engine.setProperty('rate', rate)
    engine.setProperty('volume', volume)
    engine.setProperty('voice', voice_id)
    engine.setProperty('pitch', pitch)
    engine.save_to_file(text, output_path)
    engine.runAndWait()

def select_file():
    file_types = [('Text Files', '*.txt'), ('PDF Files', '*.pdf'), ('Word Documents', '*.docx')]
    file_path = filedialog.askopenfilename(filetypes=file_types)
    if not file_path:
        return None, None

    extract_text = {
        'txt': extract_text_from_txt,
        'pdf': extract_text_from_pdf,
        'docx': extract_text_from_docx
    }[file_path.split('.')[-1]]

    return file_path, extract_text

def get_audio_settings():
    output_file = filedialog.asksaveasfilename(defaultextension=".mp3")
    return output_file

def convert_file():
    file_path, extract_text = select_file()
    if not file_path:
        return

    output_file = get_audio_settings()

    text = extract_text(file_path)
    if not text:
        messagebox.showinfo("Error", "Failed to extract text from the file.")
        return

    voice_id = voice_var.get()
    rate = rate_scale.get()
    volume = volume_scale.get()
    pitch = pitch_scale.get()

    text_to_audio(text, output_file, voice_id, rate, volume, pitch)
    messagebox.showinfo("Success", f'Converted {file_path} to {output_file} successfully!')

def main():
    global voice_var, rate_scale, volume_scale, pitch_scale
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    default_voice_name = "Microsoft Hazel Desktop - English (Great Britain)"
    voice_id = None
    for voice in voices:
        if default_voice_name in voice.name:
            voice_id = voice.id
            break
    if not voice_id:
        messagebox.showinfo("Error", f"Default voice '{default_voice_name}' not found. Please check the available voices.")
        return

    root = Tk()
    root.title("Text-to-Speech Converter")

    # Voice selection
    voice_var = StringVar(root)
    voice_var.set(voice_id)
    voice_label = Label(root, text="Select Voice:")
    voice_label.pack()
    voice_menu = OptionMenu(root, voice_var, *voices)
    voice_menu.pack()

    # Rate selection
    rate_label = Label(root, text="Select Rate:")
    rate_label.pack()
    rate_scale = Scale(root, from_=100, to=200, orient='horizontal')
    rate_scale.set(150)
    rate_scale.pack()

    # Volume selection
    volume_label = Label(root, text="Select Volume:")
    volume_label.pack()
    volume_scale = Scale(root, from_=0.0, to=1.0, resolution=0.1, orient='horizontal')
    volume_scale.set(1.0)
    volume_scale.pack()

    # Pitch selection
    pitch_label = Label(root, text="Select Pitch:")
    pitch_label.pack()
    pitch_scale = Scale(root, from_=0.5, to=2.0, resolution=0.1, orient='horizontal')
    pitch_scale.set(1.0)
    pitch_scale.pack()

    convert_button = Button(root, text="Convert File", command=convert_file)
    convert_button.pack(pady=10)

    root.mainloop()

if __name__ == '__main__':
    main()