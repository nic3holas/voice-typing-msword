import speech_recognition as sr
import win32com.client

# Initialize the recognizer
recognizer = sr.Recognizer()

# Initialize Microsoft Word
word = win32com.client.Dispatch("Word.Application")
word.Visible = True
doc = word.Documents.Add()

# Capture audio from the microphone
with sr.Microphone() as source:
    print("Please start speaking...")
    
    while True:
        audio = recognizer.listen(source)
        
        # Use Google Web Speech API to recognize the audio
        try:
            print("Recognizing...")
            text = recognizer.recognize_google(audio)
            print("You said:", text)
            
            # Type recognized text into Microsoft Word
            selection = word.Selection
            selection.TypeText(text)
            
        except sr.UnknownValueError:
            print("Sorry, I couldn't understand what you said.")
        except sr.RequestError:
            print("Sorry, I'm unable to access the Google Speech Recognition API. Please check your internet connection.")