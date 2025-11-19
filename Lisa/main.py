import speech_recognition as sr
import win32com.client
import webbrowser
import time
import os
import datetime
import urllib.parse
import requests
import uuid
import re
import screen_brightness_control as sbc
from ctypes import cast, POINTER
from deep_translator import GoogleTranslator
from gtts import gTTS
import subprocess
import playsound, tempfile, os
import keyboard
import dateparser
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import pickle
import pyautogui
from PIL import Image
import pytesseract
import pyttsx3
import comtypes
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import POINTER, cast

SCOPES = ['https://www.googleapis.com/auth/gmail.modify']
apikey = "Add your api key"
GEMINI_API_KEY = "Add your api key"
GEMINI_URL = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={GEMINI_API_KEY}"

# -------- ENHANCED TTS SETUP --------
# Original Windows SAPI setup
speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Voice = speaker.GetVoices().Item(1)

# Additional pyttsx3 engine for enhanced features
engine = pyttsx3.init()

# Recognizer configuration
r = sr.Recognizer()
r.energy_threshold = 1000
r.dynamic_energy_threshold = True
r.pause_threshold = 0.5

def set_indian_english_voice():
    voices = engine.getProperty('voices')
    indian_voice_found = False
    for voice in voices:
        # Pick voice that matches Indian English or fallback to English female
        if ("india" in voice.name.lower() or "en-in" in voice.id.lower()) and "english" in voice.name.lower():
            engine.setProperty('voice', voice.id)
            indian_voice_found = True
            break
    if not indian_voice_found:
        for voice in voices:
            if "english" in voice.name.lower() and "female" in voice.name.lower():
                engine.setProperty('voice', voice.id)
                break
    engine.setProperty('rate', 165)


set_indian_english_voice()


def enhanced_speak(text):
    """Enhanced speak function with pyttsx3"""
    print(f"Lisa A.I.: {text}")
    engine.say(text)
    engine.runAndWait()


# In-memory events list
events = []


def get_gmail_service():
    creds = None
    if os.path.exists('gmail_token.pickle'):
        with open('gmail_token.pickle', 'rb') as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                speaker.Speak("Please place your Gmail API credentials.json in the folder.")
                return None
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('gmail_token.pickle', 'wb') as f:
            pickle.dump(creds, f)

    return build('gmail', 'v1', credentials=creds)


def read_unread_emails(max_messages=5):
    """
    Fetch up to max_messages unread emails, speak subject and snippet,
    then mark them as read.
    """
    service = get_gmail_service()
    if not service:
        return

    # List unread messages
    results = service.users().messages().list(
        userId='me',
        labelIds=['INBOX', 'UNREAD'],
        maxResults=max_messages
    ).execute()
    msgs = results.get('messages', [])

    if not msgs:
        speaker.Speak("You have no unread emails.")
        return

    for msg_meta in msgs:
        msg = service.users().messages().get(
            userId='me',
            id=msg_meta['id'],
            format='metadata',
            metadataHeaders=['Subject', 'From']
        ).execute()

        headers = {h['name']: h['value'] for h in msg.get('payload', {}).get('headers', [])}
        subject = headers.get('Subject', '(no subject)')
        sender = headers.get('From', '(unknown sender)')

        # You can also fetch a snippet or full body if you like:
        snippet = msg.get('snippet', '')

        # Speak it
        speaker.Speak(f"Email from {sender}. Subject: {subject}. {snippet}")

        # Mark as read
        service.users().messages().modify(
            userId='me',
            id=msg_meta['id'],
            body={'removeLabelIds': ['UNREAD']}
        ).execute()

    speaker.Speak("Those were your unread emails.")


def create_event(duration_minutes=60):
    """
    Keeps asking until the user gives a valid "Title at TIME" sentence,
    then parses, adjusts for past-today times, and stores in global events.
    """
    while True:
        # 1) prompt for input
        speaker.Speak("Please tell me the event and time, like 'Team meeting at 3 PM today'.")
        with sr.Microphone() as src:
            r.adjust_for_ambient_noise(src, duration=0.5)
            try:
                audio = r.listen(src, timeout=7, phrase_time_limit=7)
                sentence = r.recognize_google(audio, language="en-in").lower()
                print(f"Event sentence: {sentence}")
            except (sr.UnknownValueError, sr.WaitTimeoutError):
                speaker.Speak("I didn't catch that. Let's try again.")
                continue

        # 2) validate format
        if " at " not in sentence:
            speaker.Speak("I need you to include 'at', for example: 'Lunch at 1 PM'.")
            continue

        # 3) parse title/time
        title_part, time_part = sentence.split(" at ", 1)
        event_title = title_part.strip()
        time_str = time_part.strip().lower()

        # bare-numeric times like "237" → "2:37 pm"
        if re.fullmatch(r'\d{3,4}', time_str):
            hh, mm = time_str[:-2], time_str[-2:]
            time_str = f"{int(hh)}:{mm} pm"
        else:
            time_str = re.sub(
                r'(\d{1,2})(\d{2})\s*(am|pm)',
                r'\1:\2 \3',
                time_str,
                flags=re.IGNORECASE
            )

        # 4) convert to datetime
        event_time = dateparser.parse(time_str)
        if not event_time:
            speaker.Speak(f"I couldn't understand the time '{time_str}'. Let's try again.")
            continue

        # 5) adjust for past-today
        now = datetime.datetime.now()
        if event_time < now and event_time.date() == now.date():
            event_time += datetime.timedelta(days=1)
            speaker.Speak("That time has passed today; I'm scheduling it for tomorrow.")

        # 6) success! store and break loop
        events.append({
            "title": event_title,
            "time": event_time,
            "notified": False
        })
        formatted = event_time.strftime('%I:%M %p on %A, %d %B %Y')
        speaker.Speak(f"Event created: {event_title} at {formatted}")
        print(f"✓ Event created: {event_title} at {formatted}")
        break


def check_events():
    """
    Call this frequently (e.g. once per loop). It:
      1) Grabs the current datetime
      2) Finds any event where now >= event time and not yet notified
      3) Speaks a reminder and marks it notified
    """
    now = datetime.datetime.now()
    for event in events:
        if not event["notified"] and now >= event["time"]:
            # Speak the reminder
            formatted = event["time"].strftime('%I:%M %p on %A')
            speaker.Speak(f"Reminder: {event['title']} at {formatted} is starting now.")
            # Prevent speaking again
            event["notified"] = True


def play_pause_media():
    keyboard.send('play/pause media')


app_processes = {
    "notepad": "notepad.exe",
    "calculator": "calc.exe",
    "paint": "mspaint.exe",
    "word": "winword.exe",
    "excel": "excel.exe",
    "powerpoint": "powerpnt.exe",
    "chrome": "chrome.exe",
    "firefox": "firefox.exe",
    "control panel": "control.exe",
    "task manager": "taskmgr.exe",
    "cmd": "cmd.exe",
    "mail": "outlook.exe",
    "spotify": "Spotify.exe",
    "vlc": "vlc.exe",
    "steam": "Steam.exe",
    "vs code": "Code.exe",
    "youtube": "chrome.exe"
}


def close_application(app_name):
    """
    Close an application by its process name.
    app_name: string, e.g. 'notepad.exe', 'chrome.exe'
    """
    try:
        subprocess.run(["taskkill", "/f", "/im", app_name], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        speaker.Speak(f"Closed {app_name.replace('.exe', '')}.")
    except subprocess.CalledProcessError:
        speaker.Speak(f"Could not close {app_name}. It might not be running.")


def choose_voice_gender(max_tries: int = 3):
    voices = speaker.GetVoices()

    for _ in range(max_tries):
        speaker.Speak("Would you like a women voice or a man voice?")
        with sr.Microphone() as source:
            r.adjust_for_ambient_noise(source, duration=0.5)
            try:
                audio = r.listen(source, timeout=5, phrase_time_limit=4)
                reply = r.recognize_google(audio, language="en-in").lower()
                print("Voice-choice reply:", reply)
            except sr.WaitTimeoutError:
                speaker.Speak("No answer heard.")
                continue
            except sr.UnknownValueError:
                speaker.Speak("Sorry, I couldn't understand.")
                continue
            except sr.RequestError:
                speaker.Speak("Speech service error. I'll keep the default voice.")
                return

        if any(word in reply for word in ("girl", "female", "woman")):
            desired_gender = "Female"
        elif any(word in reply for word in ("boy", "male", "man", "men")):
            desired_gender = "Male"
        else:
            speaker.Speak("Please say girl or boy.")
            continue

        # Try to find the first installed voice that matches the gender
        selected = None
        for v in voices:
            try:
                if v.GetAttribute("Gender") == desired_gender:
                    selected = v
                    break
            except Exception:
                # Older SAPI versions: fall back to description text
                if desired_gender.lower() in v.GetDescription().lower():
                    selected = v
                    break

        if selected:
            speaker.Voice = selected
            speaker.Speak(f"Okay, I will speak with a {desired_gender.lower()} voice.")
            return
        else:
            speaker.Speak(f"Sorry, no {desired_gender.lower()} voice is installed.")
            return

    speaker.Speak("I'll keep the current voice.")





def google_audio(text: str, lang_code: str = "en"):
    """
    Synthesise <text> with Google TTS and play it back.
    lang_code examples: 'en', 'hi', 'te', 'fr', …
    """
    try:
        tts = gTTS(text=text, lang=lang_code)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as fp:
            temp_mp3 = fp.name
        tts.save(temp_mp3)
        playsound.playsound(temp_mp3, block=True)
    finally:
        # remove the temp file even if playback fails
        try:
            os.remove(temp_mp3)
        except:
            pass


# spoken language names → gTTS language codes
LANG_CODES = {
    'afrikaans': 'af', 'arabic': 'ar', 'bengali': 'bn', 'chinese': 'zh-cn',
    'dutch': 'nl', 'english': 'en', 'french': 'fr', 'german': 'de',
    'hindi': 'hi', 'italian': 'it', 'japanese': 'ja', 'korean': 'ko',
    'marathi': 'mr', 'portuguese': 'pt', 'punjabi': 'pa', 'russian': 'ru',
    'spanish': 'es', 'tamil': 'ta', 'telugu': 'te', 'urdu': 'ur'
}


def translate_text(raw_cmd: str):
    """
    Accepts strings like
        "hello how are you to telugu"
        "hi everyone in spanish"
    """
    if " to " in raw_cmd:
        text_part, lang_name = raw_cmd.rsplit(" to ", 1)
    elif " in " in raw_cmd:
        text_part, lang_name = raw_cmd.rsplit(" in ", 1)
    else:
        speaker.Speak("Please say, for example: translate hello to Hindi.")
        return

    lang_name = lang_name.strip().lower()
    text_part = text_part.strip()

    lang_code = LANG_CODES.get(lang_name)
    if not lang_code:
        speaker.Speak(f"Sorry, I don't support the language {lang_name}.")
        return

    try:
        translated = GoogleTranslator(source="auto",
                                      target=lang_code).translate(text_part)
        print(f"TRANSLATION ({lang_name.title()}): {translated}")
        google_audio(translated, lang_code)
    except Exception as e:
        print("Translation error:", e)
        speaker.Speak("Something went wrong while translating.")



# Function to set the volume to a specific value (0 to 100)
def set_volume(volume_level):
    from ctypes import POINTER, cast
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume_interface = cast(interface, POINTER(IAudioEndpointVolume))
    volume_interface.SetMasterVolumeLevelScalar(volume_level / 100, None)


def get_brightness():
    current_brightness = sbc.get_brightness()[0]
    return current_brightness


# Function to increase brightness by 10%
def increase_brightness():
    current_brightness = get_brightness()
    new_brightness = min(current_brightness + 10, 100)
    sbc.set_brightness(new_brightness)
    speaker.Speak(f"Brightness increased to {new_brightness}%")


# Function to decrease brightness by 10%
def decrease_brightness():
    current_brightness = get_brightness()
    new_brightness = max(current_brightness - 10, 0)
    sbc.set_brightness(new_brightness)
    speaker.Speak(f"Brightness decreased to {new_brightness}%")


# Function to get the current volume
def get_volume():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume_interface = cast(interface, POINTER(IAudioEndpointVolume))
    volume = volume_interface.GetMasterVolumeLevelScalar() * 100
    return round(volume)



# Function to increase the volume by 10%
def increase_volume():
    current_volume = get_volume()
    new_volume = min(current_volume + 10, 100)
    set_volume(new_volume)
    speaker.Speak(f"Volume increased to {new_volume}%")


# Function to decrease the volume by 10%
def decrease_volume():
    current_volume = get_volume()
    new_volume = max(current_volume - 10, 0)
    set_volume(new_volume)
    speaker.Speak(f"Volume decreased to {new_volume}%")



# -------- ENHANCED GEMINI API CALL --------
def call_gemini_api(user_input):
    """Enhanced Gemini API call with better prompting"""
    # Direct prompt for crisp, concise Indian English answers
    prompt_text = (
            "You are Lisa A.I., a helpful and crisp Indian English voice assistant. "
            "Respond concisely and directly, no fluff.\n\nUser: " + user_input + "\nLisa:"
    )
    headers = {"Content-Type": "application/json"}
    payload = {
        "contents": [{"parts": [{"text": prompt_text}]}]
    }
    try:
        response = requests.post(GEMINI_URL, headers=headers, json=payload, timeout=15)
        response.raise_for_status()
        result = response.json()
        # Extract the first candidate text
        answer = result["candidates"][0]["content"]["parts"][0]["text"].strip()
        # Clean any trailing instructions or placeholders if any
        answer = answer.replace("\n", " ").strip()
        if not answer:
            return "Sorry, I couldn't find an answer to that."
        return answer
    except Exception as e:
        print(f"API error: {e}")
        return "Sorry, I am unable to fetch the information at the moment."


# -------- SCREEN TEXT READING --------
def read_screen_text():
    """Capture and read text from screen using OCR"""
    try:
        screenshot = pyautogui.screenshot()
        gray_img = screenshot.convert('L')
        text = pytesseract.image_to_string(gray_img).strip()
        return text
    except Exception as e:
        print(f"Screen reading error: {e}")
        return ""


def ai(prompt):
    """Original AI function - now enhanced"""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={apikey}"

    headers = {
        "Content-Type": "application/json"
    }

    data = {
        "contents": [
            {
                "parts": [
                    {"text": prompt}
                ]
            }
        ]
    }

    try:
        response = requests.post(url, headers=headers, json=data)
        response.raise_for_status()
        result = response.json()

        # Extract the response text
        output = result["candidates"][0]["content"]["parts"][0]["text"]
        full_text = f"Gemini API response for Prompt: {prompt}\n*\n\n{output}"
        print(full_text)

        # Save response to file
        output_dir = "GeminiResponses"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        safe_prompt = re.sub(r'[<>:"/\\|?*]', '_', prompt[:50])
        filename = f"{safe_prompt}_{uuid.uuid4().hex[:8]}.txt"
        file_path = os.path.join(output_dir, filename)

        with open(file_path, "w", encoding="utf-8") as f:
            f.write(full_text)

        print(f"Saved response to {file_path}")
        speaker.Speak("I have fetched and saved the response.")

    except Exception as e:
        print(f"Error fetching Gemini response: {e}")
        speaker.Speak("There was a problem fetching the AI response.")


def greet_user():
    current_hour = time.localtime().tm_hour
    tf = "morning" if 5 <= current_hour < 12 else "afternoon" if 12 <= current_hour <= 18 else "evening"
    speaker.Speak(f"Good {tf}, this is Lisa. How can I help you?")


def handle_youtube():
    speaker.Speak("Opening YouTube. Do you want to play any Video?")
    while True:
        with sr.Microphone() as source2:
            print("Listening for Video request...")
            r.adjust_for_ambient_noise(source2, duration=0.5)
            try:
                audio2 = r.listen(source2, timeout=5, phrase_time_limit=5)
                response = r.recognize_google(audio2, language="en-in").lower()
                print(f"User said: {response}")

                if response == "no thanks":
                    webbrowser.open("https://youtube.com")
                    speaker.Speak("Opening YouTube.")
                    break
                elif response in ["exit", "quit"]:
                    speaker.Speak("Okay, canceling YouTube.")
                    break
                else:
                    from pytube import Search
                    s = Search(response)
                    if s.results:
                        first_video = s.results[0]
                        webbrowser.open(first_video.watch_url)
                        speaker.Speak(f"Playing {response} on YouTube")
                    else:
                        speaker.Speak("Sorry, couldn't find the video")

                    time.sleep(1)
                    speaker.Speak("Would you like to play another Video?")
                    try:
                        audio3 = r.listen(source2, timeout=5, phrase_time_limit=5)
                        another_response = r.recognize_google(audio3, language="en-in").lower()
                        if "no thanks" in another_response:
                            speaker.Speak("Okay, enjoy your Video!")
                            break
                        elif "yes" in another_response:
                            speaker.Speak("What Video would you like to play?")
                            continue
                    except:
                        speaker.Speak("I didn't understand. Exiting YouTube mode.")
                        break
            except sr.WaitTimeoutError:
                speaker.Speak("No input received. Please try again.")
                continue
            except sr.UnknownValueError:
                speaker.Speak("I didn't understand. Could you repeat?")
                continue
            except sr.RequestError:
                speaker.Speak("I am unable to process your request. Check your internet connection.")
                break


def sendmail():
    while True:
        try:
            speaker.Speak("Please say the username: ")
            with sr.Microphone() as source:
                audio = r.listen(source, timeout=5, phrase_time_limit=5)
                user_part = r.recognize_google(audio, language="en-in").lower().replace(" ", "")
                print(f"Username part: {user_part}")
                break
        except Exception as e:
            print(f"Error getting username: {e}")
            speaker.Speak("Sorry, I didn't catch that.")
    while True:
        try:
            speaker.Speak("Please say the domain of the recipient email. For example, gmail.com or iiitl.ac.in")
            with sr.Microphone() as source:
                audio = r.listen(source, timeout=5, phrase_time_limit=5)
                domain = r.recognize_google(audio, language="en-in").lower().replace(" ", "")
                print(f"Domain: {domain}")
                break
        except Exception as e:
            print(f"Error getting domain: {e}")
            speaker.Speak("Sorry, I didn't catch that. Please say the domain again.")

    recipient = f"{user_part}@{domain}"

    while True:
        try:
            speaker.Speak("What should be the subject?")
            with sr.Microphone() as source:
                audio = r.listen(source, timeout=5, phrase_time_limit=5)
                subject = r.recognize_google(audio, language="en-in")
                print(f"Subject: {subject}")
                break
        except Exception as e:
            print(f"Error getting subject: {e}")
            speaker.Speak("Sorry, I didn't catch that. Please say the subject again.")

    while True:
        try:
            speaker.Speak("What should I write in the email?")
            with sr.Microphone() as source:
                audio = r.listen(source, timeout=5, phrase_time_limit=5)
                content = r.recognize_google(audio, language="en-in")
                print(f"Content: {content}")
                break
        except Exception as e:
            print(f"Error getting content: {e}")
            speaker.Speak("Sorry, I didn't catch that. Please say the message again.")

    try:
        email_url = f"mailto:{recipient}?subject={urllib.parse.quote(subject)}&body={urllib.parse.quote(content)}"
        webbrowser.open(email_url)
        speaker.Speak("I've opened the email draft for you. Please review and send it.")

    except Exception as e:
        print(f"Error opening email draft: {e}")
        speaker.Speak("Something went wrong while opening the email draft.")


def open_website(query):
    websites = {
        "open google": "https://www.google.com",
        "open facebook": "https://www.facebook.com",
        "open twitter": "https://www.twitter.com",
        "open linkedin": "https://www.linkedin.com",
        "open instagram": "https://www.instagram.com",
        "open reddit": "https://www.reddit.com",
        "open amazon": "https://www.amazon.com",
        "open ebay": "https://www.ebay.com",
        "open wikipedia": "https://www.wikipedia.org",
        "open yahoo": "https://www.yahoo.com",
        "open bing": "https://www.bing.com",
        "open netflix": "https://www.netflix.com",
        "open spotify": "https://www.spotify.com",
        "open github": "https://www.github.com",
        "open stack overflow": "https://stackoverflow.com",
        "open medium": "https://www.medium.com",
        "open quora": "https://www.quora.com",
        "open pinterest": "https://www.pinterest.com",
        "open tiktok": "https://www.tiktok.com",
        "open tumblr": "https://www.tumblr.com",
        "open microsoft": "https://www.microsoft.com",
        "open apple": "https://www.apple.com",
        "open adobe": "https://www.adobe.com",
        "open coursera": "https://www.coursera.org",
        "open udemy": "https://www.udemy.com",
        "open edx": "https://www.edx.org",
        "open cnn": "https://www.cnn.com",
        "open bbc": "https://www.bbc.com",
        "open nytimes": "https://www.nytimes.com",
        "open forbes": "https://www.forbes.com"
    }

    applications = {
        "open notepad": "notepad.exe",
        "open calculator": "calc.exe",
        "open paint": "mspaint.exe",
        "open word": "winword.exe",
        "open excel": "excel.exe",
        "open powerpoint": "powerpnt.exe",
        "open chrome": r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "open firefox": r"C:\Program Files\Mozilla Firefox\firefox.exe",
        "open control panel": "control.exe",
        "open task manager": "taskmgr.exe",
        "open cmd": "cmd.exe",
        "open camera": "microsoft.windows.camera:",
        "open settings": "ms-settings:",
        "open clock": "ms-clock:",
        "open calendar": "outlookcal:",
        "open mail": "outlook.exe",
        "open photos": "ms-photos:",
        "open spotify": r"C:\Users\{username}\AppData\Roaming\Spotify\Spotify.exe",
        "open vlc": r"C:\Program Files\VideoLAN\VLC\vlc.exe",
        "open steam": r"C:\Program Files (x86)\Steam\Steam.exe",
        "open vs code": r"C:\Users\{username}\AppData\Local\Programs\Microsoft VS Code\Code.exe"
    }

    for key, url in websites.items():
        if query == key:
            webbrowser.open(url)
            speaker.Speak(f"Opening {key.replace('open ', '')}.")
            return True

    try:
        for app_name, app_path in applications.items():
            if query == app_name:
                os.startfile(app_path)
                speaker.Speak(f"Opening {app_name.replace('open ', '')}")
                return True
    except Exception:
        speaker.Speak(f"Sorry, I couldn't open {query.replace('open ', '')}. The application might not be installed.")
        return False

    return False


# -------- ENHANCED COMMAND PROCESSING --------
def process_enhanced_commands(command):
    """Process enhanced commands like weather, news, and screen reading"""
    command = command.lower()

    if any(word in command for word in ["weather", "temperature", "forecast"]):
        prompt = "What is the current weather and short forecast for India? Answer crisply."
        response = call_gemini_api(prompt)
        speaker.Speak(response)
        return True

    elif "news" in command:
        prompt = "Give me the latest top news headlines in India and globally, briefly."
        response = call_gemini_api(prompt)
        speaker.Speak(response)
        return True

    elif any(phrase in command for phrase in ["read screen", "read text", "read this"]):
        speaker.Speak("Reading the text on your screen now.")
        screen_text = read_screen_text()
        if screen_text:
            speaker.Speak(screen_text)
        else:
            speaker.Speak("No readable text found on the screen.")
        return True

    return False


def main():
    greet_user()
    #choose_voice_gender()  # ask once: girl or boy voice?

    while True:
        check_events()
        with sr.Microphone() as source:
            print("Listening...")
            check_events()
            r.adjust_for_ambient_noise(source, duration=0.5)
            try:
                audio = r.listen(source, timeout=5, phrase_time_limit=5)
                query = r.recognize_google(audio, language="en-in").lower()
                print(f"User said: {query}")

                # Check for enhanced commands first
                if process_enhanced_commands(query):
                    continue

                # Original commands
                if "read my emails" in query or "read unread emails" in query:
                    read_unread_emails()
                    continue
                if "create event" in query:
                    create_event()  # this will prompt repeatedly until it gets a valid "TITLE at TIME"
                    continue
                if "close " in query:
                    app_to_close = query.replace("close ", "").strip()
                    if app_to_close in app_processes:
                        close_application(app_processes[app_to_close])
                    else:
                        speaker.Speak(f"I don't know how to close {app_to_close}.")
                    continue
                if query == "exit":
                    speaker.Speak("Thank you, have a nice day!")
                    break
                if "increase brightness" in query:
                    increase_brightness()
                if "decrease brightness" in query:
                    decrease_brightness()

                if "send email" in query or "send mail" in query:
                    sendmail()
                    continue

                if "the time" in query:
                    now = datetime.datetime.now()
                    speaker.Speak(f"The time is {now.strftime('%H')} hours and {now.strftime('%M')} minutes")
                    continue
                if "increase volume" in query:
                    increase_volume()
                    continue
                if "decrease volume" in query:
                    decrease_volume()
                    continue
                if query == "open youtube":
                    handle_youtube()
                    continue
                if "play" in query or "pause" in query:
                    play_pause_media()
                    speaker.Speak("Toggling play and pause.")
                    continue
                if query.startswith("translate "):
                    translate_text(query[len("translate "):])
                    continue
                if "using artificial intelligence" in query:
                    ai(prompt=query)
                if not open_website(query):
                    speaker.Speak(query)

            except sr.WaitTimeoutError:
                continue
            except sr.UnknownValueError:
                speaker.Speak("I couldn't understand. Please say that again.")
                continue
            except sr.RequestError:
                speaker.Speak("There was a problem connecting to the internet. Please check your connection.")
                continue

# Start the program
if __name__ == "__main__":
    main()
