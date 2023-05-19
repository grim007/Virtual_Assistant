import win32com.client
import speech_recognition as sr
import webbrowser
import openai
import os
from dotenv import load_dotenv
from googlesearch import*
from time import *
import requests
load_dotenv()
speaker=win32com.client.Dispatch("SAPI.SpVoice")
chatStr=''

def chat(text):
    global chatStr
    openai.api_key = os.getenv("api_key")
    chatStr+=f"User: {text} \n Assistant: "
    response = openai.Completion.create(
    model="text-davinci-003",
    prompt=text,
    max_tokens=500,
    top_p=1,
    frequency_penalty=0,
    presence_penalty=0
    )
    chatStr+=f"{response['choices'][0]['text']}\n"
    speaker.Speak(response["choices"][0]["text"])
    return response["choices"][0]["text"]
        
  
def ai(prompt):
    openai.api_key = os.getenv("api_key")
    result=''
    response = openai.Completion.create(
    model="text-davinci-003",
    prompt=prompt,
    max_tokens=100,
    top_p=1,
    frequency_penalty=0,
    presence_penalty=0
    )
    try:
        print(response["choices"][0]["text"])
        result=response["choices"][0]["text"]
    except Exception as e:
        print("Please try again later!")

    if not os.path.exists("Open AI"):
        os.mkdir("Open AI")

    with open (f"Open AI/ {''.join(prompt.split('intelligence')[1:])}.txt","w" ) as f:
        f.write(f"The response of {prompt}\n {result}")    


def take_command():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        audio=r.listen(source)
        while True:
            try:
                query=r.recognize_google(audio, language="en-US")
                print(f"User said: {query}")
                return query
            except Exception as e:
                print("Speak clearly you bitch!")
                speaker.Speak("Speak clearly you bitch!")
                return "start over"


if __name__=="__main__":
    speaker=win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak("Hello! I am a voice assistant at your service")
    speaker.Speak("Here are the list of things that I can help you with")
    print(''' \n              ***** List of things that I can help you with *****   
|-------------------------------------------------------------------------------| \n 
| \tOPENING DIFFERENT SITES                                                 | \n
| \tOPENING DIFFERENT APPS                                                  | \n
| \tPROVIDING RESPONSE TO YOUR QUERY USING ARTIFICIAL INTELLIGENCE          | \n
| \tCHATTING WITH YOU                                                       | \n
| \tDO A QUICK GOOGLE SEARCH                                                | \n
|-------------------------------------------------------------------------------|''') 
    speaker.Speak("How can i serve you master?")
    while True:
        print("The AI is listening to you.... ")
        text=take_command().lower()
        
        sites=[['youtube','https://www.youtube.com'],['google','https://www.google.com'],['wikipedia','https://www.wikipedia.com']]
        
        if "open".lower() in text.lower():
            if 'open microsoft word'.lower() in text.lower():
                speaker.Speak("Opening Microsoft Word")
                os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE")
            for site in sites:
                try:
                    if f"Open {site[0]}".lower() in text.lower():
                        speaker.Speak(f"Opening {site[0]} master")
                        webbrowser.open(f"{site[1]}")
                except Exception as e:
                    speaker.Speak("Please request a valid site")       
            
        
        elif 'play music'.lower() in text.lower():
            speaker.Speak("Playing music master")
            os.startfile("C:\Python\Projects\JVS\Alan Walker - Dreamer [NCS Release].mp3")

        elif 'using artificial intelligence'.lower() in text.lower(): 
            prompt=text
            ai(prompt)   
        
        elif "reset chat".lower() in text.lower():
            chatStr+=''
        
        elif "exit".lower() in text.lower():
            speaker.Speak("Exiting the assistant")
            exit()
        
        elif "start over".lower() in text.lower():
            speaker.Speak("Say it again!")
            take_command()    
               
        elif "search for".lower() in text.lower():
           to_search=text[10:]
           speaker.Speak(f"Searching for {to_search} in google")
           google_api = os.getenv("google_api")
           search_engine_id = os.getenv("search_id")
           api_url = f"https://www.googleapis.com/customsearch/v1?key={google_api}&cx={search_engine_id}&q={text}"
           response=requests.get(api_url)
           if response.status_code == 200:
                data = response.json()
                first_result = data["items"][0]["link"]
                print("First result:", first_result)
                webbrowser.open_new_tab(first_result)
           else:
                print("Error:", response.status_code)

        else:
            print("Chatting....")
            chat(text)        
        
                