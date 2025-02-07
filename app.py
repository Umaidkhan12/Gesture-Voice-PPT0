import cv2
import pyttsx3
import speech_recognition as sr
import keyboard
import win32com.client
import numpy as np
from cvzone.HandTrackingModule import HandDetector
import os
import aspose.slides as slides
import aspose.pydrawing as drawing

# Initialize PowerPoint
ppt_path = "C:\\Users\\Umaid khan\\OneDrive\\Desktop\\SIES Nerul\\Onion Routing.pptx"  # Change this to your file path
Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Open(ppt_path)
Presentation.SlideShowSettings.Run()

# Initialize text-to-speech
engine = pyttsx3.init()

# Camera and Hand Detection Setup
width, height = 900, 720
gestureThreshold = 300

cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)

detectorHand = HandDetector(detectionCon=0.8, maxHands=1)

# Variables
buttonPressed = False
counter = 0
delay = 30
imgNumber = 20

# Speech Recognition Function
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.adjust_for_ambient_noise(source, duration=1)  # Adjust for background noise
        r.pause_threshold = 0.5  # Reduce the delay before recognizing speech
        
        try:
            audio = r.listen(source, timeout=5)  # Listen for 5 seconds max
            query = r.recognize_google(audio, language='en-in').lower()
            print(f"You said: {query}")
            return query
        except sr.WaitTimeoutError:
            print("Listening timed out. Try speaking again.")
            return "None"
        except sr.UnknownValueError:
            print("Could not understand audio, please repeat.")
            return "None"
        except sr.RequestError:
            print("Could not request results from Google Speech Recognition.")
            return "None"


# Function to Control PPT via Voice
def controlPpt():
    speak("Voice control activated. Say 'quit' to exit voice control.")
    while True:
        query = takeCommand()
        
        if query == "None":  
            continue  # Ignore empty inputs

        if 'start' in query:
            keyboard.send('F5')
            speak("Starting presentation")
        
        elif 'next' in query:
            Presentation.SlideShowWindow.View.Next()
            speak("Next slide")

        elif 'previous' in query:
            Presentation.SlideShowWindow.View.Previous()
            speak("Previous slide")

        elif 'stop' in query:
            keyboard.send('escape')
            speak("Stopping presentation")

        elif 'quit' in query:
            speak("Exiting voice control mode")
            break  # Exit voice control mode

        elif 'close' in query:
            speak("Closing PowerPoint")
            os.system("TASKKILL /F /IM POWERPNT.exe")
            return

# Function for Text-to-Speech
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

# Main Loop for Hand Gesture and Voice Control
while True:
    success, img = cap.read()
    hands, img = detectorHand.findHands(img)

    if hands and not buttonPressed:
        hand = hands[0]
        lmList = hand["lmList"]
        cx, cy = hand["center"]
        fingers = detectorHand.fingersUp(hand)

        if cy <= gestureThreshold:  # Ensure hand is at the correct height
            if fingers == [1, 1, 1, 1, 1]:  # Open palm → Next slide
                print("Next Slide")
                buttonPressed = True
                Presentation.SlideShowWindow.View.Next()
            elif fingers == [1, 0, 0, 0, 0]:  # Thumbs up → Previous slide
                print("Previous Slide")
                buttonPressed = True
                Presentation.SlideShowWindow.View.Previous()
            elif fingers == [0, 1, 1, 0, 0]:  # "V" shape with fingers → Activate voice control
                print("Voice Control Activated")
                speak("Voice control activated")
                controlPpt()

    if buttonPressed:
        counter += 1
        if counter > delay:
            counter = 0
            buttonPressed = False

    cv2.imshow("Gesture Control", img)

    key = cv2.waitKey(1)

    if key == ord('q'):  # Press 'q' to quit
        break

# Release resources
cap.release()
cv2.destroyAllWindows()
