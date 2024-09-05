import pyttsx3
from docx import Document
import argparse

def extract_text(text_path):
    # Open the document
    doc = Document(text_path)
    full_text = []
    for paragraph in doc.paragraphs:
        full_text.append(paragraph.text)
    return '\n'.join(full_text)

def text_to_audio(text, output_path):
    # Initialize the text-to-speech engine
    engine = pyttsx3.init()
    
    # Adjust the speech rate
    rate = engine.getProperty('rate')
    engine.setProperty('rate', rate - 50)  # Slows down the speech rate
    
    # Adjust the volume
    engine.setProperty('volume', 1.0)  # Max volume
    
    # Find and set the British female voice (Microsoft Hazel)
    voices = engine.getProperty('voices')
    british_female_voice_id = None
    for voice in voices:
        if 'hazel' in voice.name.lower():
            british_female_voice_id = voice.id
            break

    if british_female_voice_id:
        engine.setProperty('voice', british_female_voice_id)  # Set to British female voice
    
    # Save to audio file
    engine.save_to_file(text, output_path)
    engine.runAndWait()

def main():
    parser = argparse.ArgumentParser(description='Convert a .docx file to an audio file.')
    parser.add_argument('input_file', help='Path to the input .docx file')
    parser.add_argument('output_file', help='Path to the output .mp3 file')
    
    args = parser.parse_args()
    
    # Extract text from the provided file
    text = extract_text(args.input_file)
    
    # Convert text to audio
    text_to_audio(text, args.output_file)
    
    print(f'Converted {args.input_file} to {args.output_file} successfully!')

if __name__ == '__main__':
    main()
