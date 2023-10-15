import requests
import json
import win32com.client as wincom

# Initialize the SAPI voice
speak = wincom.Dispatch("SAPI.SpVoice")

city = input("Enter the name of the city: ")
url = f"https://api.weatherapi.com/v1/current.json?key=e658119fa5f04995a3b124345231510&q={city}"
try:
    r = requests.get(url)
    r.raise_for_status()  # Check for request errors
    wdic = r.json()
    temperature = wdic['current']['temp_c']  # Assuming you want temperature in Celsius
    weather = wdic['current']['condition']['text']
    text = f"The current weather in {city} is {temperature} degrees Celsius, and the condition is {weather}."
    speak.Speak(text)
except requests.exceptions.RequestException as e:
    print("An error occurred while making the request:", e)
except KeyError:
    print("Unable to retrieve weather information for the provided city.")
