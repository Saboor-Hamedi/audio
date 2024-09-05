import pyttsx3
from docx import Document
from prompt_toolkit import prompt

def extract_text(text_path):
    doc = Document(text_path)
    full_text = []
    for paragraph in doc.paragraphs:
        full_text.append(paragraph.text)
    return '\n'.join(full_text)

def text_to_audio(text, output_path, voice_id, rate, volume):
    engine = pyttsx3.init()
    engine.setProperty('rate', rate)
    engine.setProperty('volume', volume)
    
    # Set the default voice (Hazel)
    engine.setProperty('voice', voice_id)

    engine.save_to_file(text, output_path)
    engine.runAndWait()

def main():
    engine = pyttsx3.init()
    
    # List available voices
    voices = engine.getProperty('voices')
    

    # Default voice ID (Microsoft Hazel Desktop)
    default_voice_name = "Microsoft Hazel Desktop - English (Great Britain)"
    voice_id = None
    for voice in voices:
        if default_voice_name in voice.name:
            voice_id = voice.id
            break
    
    if not voice_id:
        print(f"Default voice '{default_voice_name}' not found. Please check the available voices.")
        return

    # Get user inputs
    input_file = prompt("Enter the path to the .docx file: ").strip()
    output_file = prompt("Enter the path to the output .mp3 file: ").strip()
    rate = int(prompt("Enter the speech rate (e.g., 150): "))
    volume = float(prompt("Enter the volume (0.0 to 1.0): "))

    # Extract text and convert to audio
    text = extract_text(input_file)
    text_to_audio(text, output_file, voice_id, rate, volume)
    print(f'Converted {input_file} to {output_file} successfully!')

if __name__ == '__main__':
    main()
