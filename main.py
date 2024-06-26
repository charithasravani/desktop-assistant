import speech_recognition as sr
import google.generativeai as genai
import os
from dotenv import load_dotenv

import win32com.client
import webbrowser
import datetime
load_dotenv()


# Initialize the speech synthesis engine
speaker = win32com.client.Dispatch("SAPI.SpVoice")

# Initialize a global chat string to store the conversation
chatStr = ""

# Replace with your actual Gemini API key
apikey = os.getenv('API_KEY')

# Function to capture voice input
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.8  # Set to a sensible default
        r.non_speaking_duration = 0.5  # Explicitly set non_speaking_duration
        print("Listening...")
        audio = r.listen(source)
        try:
            print("Recognizing...")
            req = r.recognize_google(audio, language="en-in")
            print(f"User said: {req}")
            return req
        except sr.UnknownValueError:
            print("Sorry, I did not understand that.")
            return "Sorry, I did not understand that."
        except sr.RequestError as e:
            print(f"Could not request results; {e}")
            return "Sorry, the service is down."

# Function to write responses to a file
def write_response_to_file(query, response_text):
    directory = "genAI"
    if not os.path.exists(directory):
        os.makedirs(directory)
    with open(f"{directory}/{query[:30]}.txt", "w", encoding="utf-8") as file:
        file.write(response_text)

# Function to handle chat interactions
def chat(query):
    global chatStr
    print(chatStr)
    genai.configure(api_key=apikey)
    chatStr += f"User: {query}\n Nova: "
    model = genai.GenerativeModel('gemini-1.5-flash')
    response = model.generate_content(query)
    response_text = response.text.replace('*', '').strip()
    response_text = response_text.replace("#",'')
    print(f"Nova: {response_text}")
    speaker.Speak(response_text)
    chatStr += f"{response_text}\n"
    return response_text

# Function to handle AI interactions
def ai(query):
    genai.configure(api_key=apikey)
    text = f"Gemini AI response for Query: {query}\n*************************\n\n"
    model = genai.GenerativeModel('gemini-1.5-flash')
    response = model.generate_content(query)
    response_text = response.text.replace('*', '').strip()
    response_text = response_text.replace("#", '')
    text += response_text
    if not os.path.exists("genAI"):
        os.mkdir("genAI")
    with open(f"genAI/{''.join(query.split('intelligence')[1:]).strip()}.txt", "w") as f:
        f.write(text)
    speaker.Speak(response_text)
    return response_text

# Main loop
while True:
    print("Speak something to interact with the AI...")
    speaker.Speak("I am Nova AI")

    query = takeCommand()
    speaker.Speak(query)

    # List of websites to open
    sites = [
        ["youtube", "https://www.youtube.com"],
        ["wikipedia", "https://www.wikipedia.com"],
        ["google", "https://www.google.com"],
        ["instagram", "https://www.instagram.com"]
    ]

    # Check if the query matches any site commands
    for site in sites:
        if f"open {site[0]}".lower() in query.lower():
            speaker.Speak(f"Opening {site[0]} sir...")
            webbrowser.open(site[1])

    # Check if the query is about the time
    if "the time".lower() in query.lower():
        strfTime = datetime.datetime.now().strftime("%H:%M:%S")
        hour = datetime.datetime.now().strftime("%H")
        minute = datetime.datetime.now().strftime("%M")
        second = datetime.datetime.now().strftime("%S")
        speaker.Speak(f"The time is {hour} hours {minute} minutes {second} seconds")

    # Pass the query to the AI function if "using AI" is in the query, else to chat function
    if "using Artificial intelligence".lower() in query.lower():
        ai(query)
    elif "Quit chat".lower() in query.lower():
        break
    elif "chat reset".lower() in query.lower():
        chatStr = ""
    else:
        print("chatting...")
        chat(query)

