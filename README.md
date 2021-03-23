# voice_assistant

#### Simple voice assistant created using python

Packages  required to install to run this program:

- SpeechRecognition
- pyttsx3
- spacy
- PyAudio

To run voice assistant on your computer first you must provide the program with absolute path to your web browser in main.py file:

```python
firefox_path = r"C:\\Program Files\\Mozilla Firefox\\firefox.exe"
webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(firefox_path))
```

Also you must update absolute paths to apps installed on your computer which you want voice assistant to be able to run:

```python
aplikacje = {    
    'spotify': r'C:\\Users\\Marta Roszak\\AppData\\Roaming\\Spotify\\Spotify.exe',
    # etc.
    'notatnik': r'C:\\WINDOWS\\system32\\notepad.exe'
}
```