import sys
from googlesearch import search
import psutil
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QPoint
from PyQt5.QtGui import QPainter, QColor, QFont
import win32com.client
import webbrowser as wb
from Body import Listen
from features.Chatbot.openAI_chatbot import reply_chat
from features.connect_with_mobile import connect
from features.playsongs import play_songs
from features.openApplication import openapp
from features.playsongs.play_on_youtube import search_youtube, play_youtube_video
from features.volume_control.volume import volumeAdjust
from features.weather.weather import get_weather
from features.windows_features_on_off import close_window
from features.Image_Generation.AI_image import generate_image
import pyautogui
import datetime
import socket
import time
import re
import os
from PyQt5.QtCore import pyqtSignal, pyqtSlot

# Initialize the text-to-speech engine
speaker = win32com.client.Dispatch("SAPI.SpVoice")


def say(text, rate=-1):
    speaker.Rate = rate
    speaker.speak(text)


def greet_user():
    """Function to greet the user based on the time of day."""
    current_time = datetime.datetime.now().strftime("%H.%M")
    if "06.00" <= current_time < "12.00":
        say("Good morning sir.")
    elif current_time == "12.00":
        say("Good noon sir.")
    elif "12.00" < current_time < "18.00":
        say("Good afternoon sir.")
    elif "18.00" <= current_time < "21.00":
        say("Good evening sir.")
    else:
        say("Welcome sir.")


class VoiceAssistantThread(QThread):
    """Thread class to handle voice assistant operations."""
    new_response = pyqtSignal(str)
    listening = pyqtSignal(bool, str)  # Update to pass the response text as well
    flag = 1  # Flag to control if the chatbot should respond
    previous_response = ""
    previous_query = ""
    stop_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self._is_running = True
        self.stop_signal.connect(self.stop)

    def run(self):
        """Main thread function to handle voice interactions."""
        greet_user()
        while self._is_running:
            self.listening.emit(True, "")  # Indicate that the assistant is listening
            query = Listen.MicExecution()
            self.listening.emit(False, ' Response...')  # Indicate that the assistant is processing
            response = self.handle_query(query)
            self.listening.emit(False, response)  # Indicate the actual response

    @pyqtSlot()
    def stop(self):
        self._is_running = False

    def handle_query(self, query):
        """Function to handle user queries and provide appropriate responses."""
        if "open google" in query.lower() or "search google" in query.lower():
            say("Launching Google...")
            wb.open("www.google.com")
            time.sleep(2)
            return "Is there anything else, sir?"
        elif "search youtube" in query.lower() or "search on youtube" in query.lower() or "on youtube" in query.lower():
            if self.flag == 1:
                self.flag = 0
            say("okay ")
            query = query.replace("search youtube", "")
            query = query.replace("search on youtube", "")
            query = query.replace("on youtube", "")
            query = query.replace("search", "")
            video_id = search_youtube(query)
            say("Playing the first video on YouTube...")
            play_youtube_video(video_id)
        elif "open website" in query.lower() or "show " in query.lower() or "show today's news" in query.lower() or "show today news" in query.lower() or "match update" in query.lower() or "ipl" in query.lower():
            query = re.sub(r"(open website|search|show)", "", query, flags=re.IGNORECASE)
            for url in search(query, tld="co.in", num=10, stop=1, pause=2):
                site = r"https?://(www\.)?([^/]+)"
                matches = re.search(site, url)
                if matches:
                    domain = matches.group(2)
                    say(f"Opening {query} at {domain} ...")
                wb.open(url)
                say("Is there anything else, sir?")
        elif "open application" in query.lower():
            query = re.sub(r"(Park|open|application)", "", query, flags=re.IGNORECASE)
            openapp.openapp(query)
            say("Is there anything else, sir?")
        elif "play song" in query.lower() or "play the song" in query.lower() or "Play a song" in query.lower() or "play any song" in query.lower() or "play any songs" in query.lower() or "playlist" in query.lower():
            if "play song" in query.lower() or "play the song" in query.lower() or "play any song" in query.lower() or "play any songs" in query.lower():
                say("sir what song should i play")
                song = Listen.MicExecution()
                play_songs.play_songs(song)
            elif "playlist" in query.lower():
                # from features.playsongs.play_songs import playlist_chat
                reply_from_playlist = play_songs.playlist_chat(query)
                say(reply_from_playlist)
        elif "connect my phone" in query.lower():
            say("okay, trying to connect your Phone..")
            connect.connectMobile(60)
        elif "create image" in query.lower() or "generate image" in query.lower():
            try:
                query = query.replace("create image", "").strip()
                print("wait few seconds")
                say("Please wait a few seconds.")
                generate_image(query)
                say("An image has been created. Please check it out.")
            except:
                say("sorry sir ..")
        elif "weather" in query.lower() or "update weather" in query.lower() or "current weather" in query.lower():
            # api_key = '44d1d069131442e6ade125836241806'  # Replace with your Weather API key
            match = re.search(r"weather in (\w+)", query.lower())
            if match:
                city = match.group(1)
            else:
                city = 'Kalyani'  # Replace with a default city if not specified in the query
            weather_info = get_weather(city)
            say(weather_info)
            return weather_info
        elif "time now" in query:
            strTime = datetime.datetime.now().strftime("%H:%M")
            say(f"sir time is {strTime} now")
        elif "delete the chat log" in query.lower() or "delete chat log" in query.lower() or "clear chat log" in query.lower():
            say("sir it is an important file.there is had your previous saved data ......are you sure you can "
                "delete"
                "this file , yes or No ?")
            worn = Listen.MicExecution()
            if "yes" in worn.lower():
                f = open("chat_log.txt", "w")
                f.write("")
                print("clearing the logfile....")
                time.sleep(5)
                say("done sir. all deleted")
            else:
                say("sorry not delete the chat log, thankyou sir..")
        elif "show text" in query.lower():
            query = query.replace("show text", "")
            result = reply_chat(query)
            print(result)
        elif "save data" in query.lower():
            if not os.path.exists("C:\\Users\\rimpa\\Desktop\\ai respond save data"):
                os.mkdir("C:\\Users\\rimpa\\Desktop\\ai respond save data")
            if self.previous_response != "":
                # Create a text string with both the previous and current responses
                text = f"OpenAI Response for prompt :\n***************************************************\n\n"
                text += self.previous_response
                with open(f"C:\\Users\\rimpa\\Desktop\\ai respond save data/{''.join(self.previous_query)}.txt",
                          "w") as f:
                    f.write(text)
                    say("done sir")
            query = query.replace("save", "")
            query = query.replace("data", "")
            if query.strip() != '':
                reply = reply_chat(query)
                print(reply)
                # Create a text string current responses
                text = f"OpenAI Response for prompt :\n***************************************************\n\n"
                text += reply
                with open(f"C:\\Users\\rimpa0\\Desktop\\ai respond save data/{''.join(query)}.txt", "w") as f:
                    f.write(text)
                say("saved it..")
        elif "shut down my computer" in query.lower() or "shutdown my computer" in query.lower() or "shutdown computer" in query.lower():
            say("shutting down your computer sir..")
            os.system("shutdown /s /t 2")
        elif "restart my computer" in query.lower() or "restart computer" in query.lower():
            say("okay sir i will restart your computer..")
            os.system("shutdown /r /t 2")
        elif "log out my computer" in query.lower() or "log out computer" in query.lower():
            say("okay sir ..")
            os.system("shutdown /l ")
        elif "charge on my computer" in query.lower() or "charge status" in query.lower() or "battery status" in query.lower():
            battery = psutil.sensors_battery()
            if battery is None:
                print("No battery found.")
                exit()
            percentage = battery.percent
            print(f"Battery Percentage: {percentage}%")
            say(f"sir Battery {percentage} percentage are available now ")
        elif "give me prompt" in query.lower():
            query = input("You: ")
            reply = reply_chat(query)
            print(reply)
            say(reply)
        elif "volume" in query.lower() or "current volume" in query.lower():
            vol = volumeAdjust(query)
            print(vol)
            say(vol)
        elif "what is my ip" in query.lower():
            hostname = socket.gethostname()
            ip = socket.gethostbyname(hostname)
            print(ip)
            say(f"your ip address is {ip} ")
        elif "pause" in query.lower() or "resume" in query.lower():
            pyautogui.hotkey('playpause')
        elif "page down" in query.lower():
            pyautogui.hotkey('pagedown')
        elif "page up" in query.lower():
            pyautogui.hotkey('pageup')
        elif "close this" in query.lower() or "close window" in query.lower():
            say("closing sir..")
            pyautogui.hotkey('alt', 'f4')
        elif "close" in query.lower() or "exit" in query.lower():
            query = query.replace("close", "")
            query = query.replace("exit", "")
            close_window.close_window(query)
        elif "quit" in query.lower():
            if ("21.00" <= datetime.datetime.now().strftime("%H.%M") <= "24.00") or (
                    "01.00" <= datetime.datetime.now().strftime("%H.%M") < "05.00"):
                say("bye, good night sir.")
            else:
                say("bye sir, thank you.")
            clearfile = open("chat_log.txt", "w")
            clearfile.write("")
            clearfile.close()  # Close the file after writing to it
            self.stop_signal.emit()  # Signal the thread to stop
            QApplication.instance().quit()  # Quit the application
        elif "stop park" in query.lower() or "stop answer" in query.lower() or "answer mode stop" in query.lower():
            if self.flag == 0:
                say("answer mode already stopped.")
            else:
                self.flag = 0
                say("okay sir. answer mode stopped now.")
        elif "hello park" in query.lower() or "ready for answer" in query.lower() or "start answer" in query.lower():
            self.flag = 1
            say("hello sir ")
        elif "minimise window" in query.lower() or "minimais this" in query.lower() or "minimise this" in query.lower() or "minimize window" in query.lower():
            pyautogui.hotkey('win', 'm')
        else:
            try:
                if self.flag == 1:
                    reply = reply_chat(query)
                    print(reply)
                    say(reply)
                    self.previous_response = reply
                    self.previous_query = query
                    return reply
            except Exception as e:
                return str(e)


class VoiceAssistantGUI(QWidget):
    """Main GUI class for the voice assistant."""

    def __init__(self):
        super().__init__()
        self.initUI()
        self.voice_assistant_thread = VoiceAssistantThread()
        self.voice_assistant_thread.new_response.connect(self.update_output)
        self.voice_assistant_thread.listening.connect(self.update_listening_status)
        self.voice_assistant_thread.start()

    def closeEvent(self, event):
        """Handle the close event to ensure the thread stops."""
        self.voice_assistant_thread.stop_signal.emit()  # Signal the thread to stop
        self.voice_assistant_thread.wait()  # Wait for the thread to finish
        event.accept()

    def initUI(self):
        """Initialize the GUI components."""
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setGeometry(100, 100, 400, 400)

        self.output_label = QLabel(self)
        self.output_label.setFont(QFont('Arial', 14))
        self.output_label.setStyleSheet("QLabel { color : white; }")
        self.output_label.setAlignment(Qt.AlignTop)

        self.status_label = QLabel(self)
        self.status_label.setFont(QFont('Arial', 14))
        self.status_label.setStyleSheet("QLabel { color : green; }")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setText("Listening...")

        layout = QVBoxLayout()
        layout.addWidget(self.status_label)
        layout.addWidget(self.output_label)
        self.setLayout(layout)

        self.moving = False
        self.offset = QPoint()

    def paintEvent(self, event):
        """Custom paint event to draw the ellipse."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(QColor(0, 255, 0))  # Green color
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(75, 75, 60, 60)  # Adjusted position and size

    def mousePressEvent(self, event):
        """Event to handle mouse press for moving the window."""
        if event.button() == Qt.LeftButton:
            self.moving = True
            self.offset = event.pos()

    def mouseMoveEvent(self, event):
        """Event to handle mouse move for moving the window."""
        if self.moving:
            self.move(event.globalPos() - self.offset)

    def mouseReleaseEvent(self, event):
        """Event to handle mouse release after moving the window."""
        self.moving = False

    def update_output(self, text):
        """Update the output label with new text."""
        self.output_label.setText(text)
        self.output_label.adjustSize()

    def update_listening_status(self, listening, response_text):
        """Update the status label based on the listening status."""
        if listening:
            self.status_label.setText("Listening...")
        else:
            self.status_label.setText(response_text)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    gui = VoiceAssistantGUI()
    gui.show()
    sys.exit(app.exec_())
