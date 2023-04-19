import requests
import json
import win32com.client as wincom

speak = wincom.Dispatch("SAPI.SpVoice")

city = input("Enter the name of the city\n")

url = f"https://api.weatherapi.com/v1/current.json?key=21e93cf090c74d32aad204853231404&q={city}"

r = requests.get(url)
wdic = json.loads(r.text)
wc = wdic["current"]["temp_c"]
wf = wdic["current"]["temp_f"]
time = wdic["location"]["localtime"]
speak.Speak(f"The current weather in {city} is {wc} degree celcius and {wf} degree fahrenheit on {time}")