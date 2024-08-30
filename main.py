import os
import random
import speech_recognition as sr
import win32com.client
from openai import OpenAI
from dotenv import load_dotenv


speaker = win32com.client.Dispatch("SAPI.SpVoice")

def openAi(prompt):
    load_dotenv()
    openai.api_key = os.getenv('API_KEY')

    text = f"OpenAi response for Prompt: {prompt}\n ***********\n\n"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    try:
        print(response["choices"][0]["text"])
        text += response["choices"][0]["text"]
        if not os.path.exists("Openai"):
            os.mkdir("Openai")
        with open(f"Openai/{prompt[0:30]}.txt", "w") as f:
            f.write((text))
        return response["choices"][0]["text"]
    except Exception as e:
        return("sorry i am unable to understant what you just said can you please try again")

chatStr = ""
def chat(prompt):
    load_dotenv()
    openai.api_key = os.getenv('API_KEY')
    global  chatStr
    chatStr += f"user: {prompt}\n Jarvis: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=chatStr,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    try:
        print(response["choices"][0]["text"])
        chatStr += f"{response['choices'][0]['text']}\n"
        return response["choices"][0]["text"]
    except Exception as e:
        return("sorry i am unable to understant what you just said can you please try again")
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.6
        r.energy_threshold = 1000
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said:{query}")
            return  query
        except Exception as e:
            return("Some error occured")

if __name__ == '__main__':
    print('MOBAIS')
    speaker.Speak("Hello  How Can I Help You ")
    while True:
        print("Listening...")
        # query = takeCommand()
        query = ("how can you help me ?")
        if "using ai".lower() in query.lower():
            output = openAi(query)
            print(output)
            speaker.Speak(output)

        else:
            output = chat(query)
            speaker.Speak(output)