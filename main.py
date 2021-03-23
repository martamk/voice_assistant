import speech_recognition as sr
import time
import webbrowser
import subprocess
import spacy
import pyttsx3 as tts
import os
import wmi
import re

r = sr.Recognizer()

engine = tts.init()
engine.setProperty('rate', 150)

firefox_path = r"C:\\Program Files\\Mozilla Firefox\\firefox.exe"
webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(firefox_path))

appCatalogue = r'C:\\Users\\Marta Roszak\\Desktop\\sztudia\\PJN_cwiczenia\\ludzka_konsola\\LK-zapisane'

webCommands = {"puść", "wyszukaj", "znajdź", "poszukaj", "otwórz"}  # słowo "wyszukaj" chyba nie działa w tym core
appCommandsOpen = {"uruchom", "odpal", "włącz"}
appCommandsClose = {"zamknij", "zakończ", "wyłącz"}
fileCommands = {"otwórz plik", "pokaż"}
noteCommands = {"zrób", "napisz", "zapisz", "zanotuj"}

nlp = spacy.load('pl_core_news_sm')
# nlp = spacy.load('pl_core_news_lg')     # python -m spacy download pl_core_news_lg (już pobrane)

noteCount = 1

aplikacje = {
    'spotify': r'C:\\Users\\Marta Roszak\\AppData\\Roaming\\Spotify\\Spotify.exe',
    'microsoft word': r'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\WINWORD.exe',
    'microsoft excel': r'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.exe',
    'microsoft powerpoint': r'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.exe',
    'visual studio': r'C:\\Program Files (x86)\\Microsoft Visual Studio\\2019\\Enterprise\\Common7\\IDE\\devenv.exe',
    'discord': r'C:\\Users\Marta Roszak\\AppData\\Local\\Discord\\Update.exe',
    'microsoft teams': r'C:\\Users\\Marta Roszak\\AppData\\Local\\Microsoft\\Teams\\Update.exe',
    'notatnik': r'C:\\WINDOWS\\system32\\notepad.exe'
}


def say(text):
    engine.say(text)
    engine.runAndWait()


def getText():
    with sr.Microphone() as source:
        try:
            r.adjust_for_ambient_noise(source, duration=0.3)
            say("Co mogę dla Ciebie zrobić?")
            audio = r.listen(source, phrase_time_limit=5)   # limit 5 sekund
            text = r.recognize_google(audio, language='pl-PL')
            return text
        except:
            return 0


def getNote():
    with sr.Microphone() as source:
        try:
            r.adjust_for_ambient_noise(source, duration=0.3)
            say("Co zanotować?")
            audio = r.listen(source, phrase_time_limit=6)   # limit 6 sekund
            text = r.recognize_google(audio, language='pl-PL')
            return text
        except:
            return 0


def searchWeb(command):
    if "youtube" in command:
        say("Otwieram w YouTube")
        indeks = command.split().index("youtube")
        query = command.split()[indeks + 1:]
        adres = "http://www.youtube.com/results?search_query=" + '+'.join(query)
        webbrowser.get('firefox').open(adres, new=2)
    elif "google" in command:
        say("Szukam w Google")
        indeks = command.split().index("google")
        query = command.split()[indeks + 1:]
        adres = "https://www.google.com/search?q=" + '+'.join(query)
        webbrowser.get('firefox').open(adres, new=2)
    else:
        try:
            say("Otwieram stronę")
            indeks = command.split().index("otwórz")
            query = command.split()[indeks + 1]
            adres = "http://www." + '/'.join(query) + ".pl"
            webbrowser.get('firefox').open(adres, new=2)
        except:
            say("Chyba się nie zrozumieliśmy jednak")


def openFile(command):
    say('Otwieram plik')
    indeks = command.split().index("plik")
    filename = command.split()[indeks + 1:]
    filename = ' '.join(filename)
    ext = ''
    for root, dirs, files in os.walk(appCatalogue):
        for name in files:
            plik = name.split('.')
            nazwa = plik[0]
            rozsz = plik[1]
            if nazwa == filename:
                ext = rozsz
    fileLocalization = appCatalogue + '\\' + filename + '.' + ext
    print(fileLocalization)
    os.startfile('"' + fileLocalization + '"')


def runApp(command, polecenie):
    say('Otwieram program')
    indeks = command.split().index(polecenie)
    app = command.split()[indeks + 1:]
    app = ' '.join(app)
    appLocalizztion = aplikacje[app]
    subprocess.run(('cmd', '/C', 'start', '', appLocalizztion))


def closeApp(command, polecenie):
    indeks = command.split().index(polecenie)
    app = command.split()[indeks + 1:]
    app = ''.join(app)
    if os.system("taskkill /f /im  " + app + ".exe") == 0:
        say("Zamykam proces.")
    else:
        say("Nie znajduję tego procesu.)")
        print("Czy chodziło Ci o któryś z tych?")
        f = wmi.WMI()
        for process in f.Win32_Process():
            p = re.search(app, str(process.Name), flags=re.IGNORECASE)
            if p:
                print(process.Name)
        app = input("Wprowadź nazwę procesu, który chcesz zamknąć (Włącznie z '.exe')")
        os.system("taskkill /f /im  " + app)


def makeNote():
    global noteCount
    noteLocalization = appCatalogue + "\\notatka " + str(noteCount) + ".txt"
    with open(noteLocalization, 'w', encoding='utf-8') as note:
        notatka = getNote()
        print(notatka)
        if notatka != 0:
            note.write(notatka)
            noteCount += 1
        else:
            say("Nie zrozumiałam, co chcesz zanotować,")
    note.close()


def processCommand(command):
    try:
        # dzielenie na tokeny i postagging w celu skategoryzowania polecenia
        tokens = nlp(command)
        polecenie = ''
        indeksPolecenia = 0
        for i in range(0, len(tokens)):
            if tokens[i].tag_ == "IMPT":
                polecenie = tokens[i].text
                indeksPolecenia = i
                break
        # print("-> polecenie to: " + polecenie)
        if polecenie == "otwórz" and tokens[indeksPolecenia + 1].text == "plik" or polecenie == "pokaż":
            openFile(command)
        elif polecenie in webCommands:
            searchWeb(command)
        elif polecenie in appCommandsOpen:
            runApp(command, polecenie)
        elif polecenie in appCommandsClose:
            closeApp(command, polecenie)
        elif polecenie in noteCommands:
            makeNote()
        else:
            pass
    except:
        say("Przepraszam, nie mogę wykonać polecenia.")
        return


if __name__ == '__main__':
    txt = "start"
    key = input("Wpisz 1 by rozpocząć  / 0 aby zakończyć")
    while key == '1':
        txt = getText()
        # txt = input("Co mogę dla Ciebie zrobić?")
        print(txt.lower())
        if txt != 0:
            # print("...")
            processCommand(txt.lower())
            time.sleep(3)
        else:
            pass
        key = input("Wpisz 1 by rozpocząć / 0 aby zakończyć")


