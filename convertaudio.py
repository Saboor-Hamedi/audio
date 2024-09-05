import pyttsx3
import docx
import pdfplumber
import os
from prompt_toolkit import prompt

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

def text_to_audio(text, output_path, voice_id, rate, volume):
    engine = pyttsx3.init()
    engine.setProperty('rate', rate)
    engine.setProperty('volume', volume)
    engine.setProperty('voice', voice_id)
    engine.save_to_file(text, output_path)
    engine.runAndWait()

def select_file():
    print("Select a file type:")
    print("1. TXT")
    print("2. PDF")
    print("3. DOCX")
    
    file_type = input("Enter your choice (1/2/3): ").strip()
    if file_type not in {'1', '2', '3'}:
        print("Invalid choice. Exiting.")
        return None, None

    file_name = prompt("Enter the file name: ").strip()
    file_extension = {'1': '.txt', '2': '.pdf', '3': '.docx'}[file_type]
    file_path = file_name + file_extension
    
    if not os.path.isfile(file_path):
        print(f"File {file_path} not found. Exiting.")
        return None, None

    extract_text = {
        '1': extract_text_from_txt,
        '2': extract_text_from_pdf,
        '3': extract_text_from_docx
    }[file_type]
    
    return file_path, extract_text

def get_audio_settings():
    output_file = prompt("Audio name: ").strip() + '.mp3'
    rate = prompt("Enter the speech rate (default: 150): ").strip()
    volume = prompt("Enter the volume (0.0 to 1.0, default: 1.0): ").strip()

    try:
        rate = int(rate) if rate else 150
        volume = float(volume) if volume else 1.0
    except ValueError:
        print("Invalid rate or volume. Using default values.")
        rate, volume = 150, 1.0

    return output_file, rate, volume

def main():
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    default_voice_name = "Microsoft Hazel Desktop - English (Great Britain)"
    voice_id = None
    for voice in voices:
        if default_voice_name in voice.name:
            voice_id = voice.id
            break
    if not voice_id:
        print(f"Default voice '{default_voice_name}' not found. Please check the available voices.")
        return

    file_path, extract_text = select_file()
    if not file_path:
        return

    output_file, rate, volume = get_audio_settings()

    text = extract_text(file_path)
    if not text:
        print("Failed to extract text from the file. Exiting.")
        return

    text_to_audio(text, output_file, voice_id, rate, volume)
    print(f'Converted {file_path} to {output_file} successfully!')

if __name__ == '__main__':
    main()
