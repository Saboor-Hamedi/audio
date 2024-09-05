import pyttsx3
def list_voices():
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    
    for idx, voice in enumerate(voices):
        print(f"Voice {idx}:")
        print(f"  - Name: {voice.name}")
        print(f"  - Gender: {voice.gender}")
        print(f"  - ID: {voice.id}")
        print(f"  - Languages: {voice.languages}\n")

list_voices()