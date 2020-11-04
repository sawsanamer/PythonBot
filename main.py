import os
import pyautogui
import pandas as pd
import time
from datetime import datetime
from win10toast import ToastNotifier
from datetime import date


def pressWindowsAndG():
    pyautogui.keyDown('winleft')
    pyautogui.press('g')
    pyautogui.keyUp('winleft')

def  record_meeting(meeting_finishes):
    pressWindowsAndG()
    time.sleep(2)
    findButtonAndClick('recording_images/record_button.PNG', False)
    pressWindowsAndG()
    while True:
        currentTime = datetime.now().strftime("%H:%M")
        if (meeting_finishes==currentTime):
            findButtonAndClick('recording_images/stop_recording_button.PNG', False)
            break


def stopRecordingandExitMeeting(application):
    findButtonAndClick('recording_images/stop_recording_button.PNG', False)

    if application=="zoom":
        findButtonAndClick('zoom_images/close_meeting.PNG', False)
        findButtonAndClick('zoom_images/leave_meeting.PNG', False)
    else:
        findButtonAndClick('microsoft_teams_images/leave_meeting.PNG', False)

def findButtonAndClick(image, sleep):
    if (sleep):
        time.sleep(8)
    button = pyautogui.locateOnScreen(image)
    pyautogui.moveTo(button)
    pyautogui.click()

def waitThenType(val):
    time.sleep(4)
    pyautogui.write(val)

def MapDaysToIntWeekdayValue(meeting_days):
    days=[]
    for day in meeting_days:
        if(day=="M"):
            days.append(0)
        elif(day=="T"):
            days.append(1)
        elif (day == "W"):
            days.append(2)
        elif (day == "TH"):
            days.append(3)
        elif (day == "F"):
            days.append(4)
        elif (day == "SU"):
            days.append(5)
        elif (day == "SAT"):
            days.append(6)
    return days


def sendNotificationReminder( meetingTime, reminderTime, subject):
    difference = datetime.strptime(meetingTime, "%H:%M") - datetime.strptime(reminderTime, "%H:%M")
    hr = ToastNotifier()
    hr.show_toast('Meeting coming up', 'Time left for '+subject+ ' meeting to start: '+str(difference))


def automate_zoom(meetingid, password):
    os.startfile(
        "C:\\Users\\sawsan\\AppData\\Roaming\\Zoom\\bin\\Zoom" + '.exe')

    time.sleep(5)
    loc = pyautogui.locateOnScreen("zoom_images/join_meeting.PNG", grayscale=True, confidence=.5)
    pyautogui.click(loc)

    waitThenType(meetingid)
    findButtonAndClick('zoom_images/checkbox_button.PNG', False)
    findButtonAndClick('zoom_images/checkbox_button.PNG', False)
    findButtonAndClick('zoom_images/join_button2.PNG', False)
    waitThenType(password)
    findButtonAndClick('zoom_images/join_meeting_button3.PNG', False)

    join_with_computer_audio = None
    while join_with_computer_audio is None:
        join_with_computer_audio = pyautogui.locateOnScreen('zoom_images/join_with_computer_audio.PNG')
    pyautogui.moveTo(join_with_computer_audio)
    pyautogui.click()
    findButtonAndClick('zoom_images/mute_mic.PNG', False)



def automate_microsoft_teams (subject):
    # os.startfile(
    #     "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    import webbrowser
    url = 'https://outlook.office.com/mail/inbox'
    webbrowser.register('chrome',
                        None,
                        webbrowser.BackgroundBrowser(
                            "C://Program Files (x86)//Google//Chrome//Application//chrome.exe"))
    webbrowser.get('chrome').open(url)

    route="microsoft_teams_images/"
    findButtonAndClick(route+'calendar.PNG', True)
    time.sleep(1)
    findButtonAndClick(route+subject+".PNG", True)
    findButtonAndClick(route+"join.PNG", True)
    findButtonAndClick(route+"open_microsoft_teams_button.PNG", True)
    findButtonAndClick(route+"mic_checkbox.PNG", True)
    findButtonAndClick(route+"join_now_button2.PNG", True)





def scheduleMeetings ():
    meetings_data = {'program':          ['zoom', 'teams', "zoom"],
                     'meeting_time':     ['08:30', "10:00", "11:30"],
                     'id':               ['719 8322 8149', '', "762 1983 4641" ],
                     'password':         ['blah', '',"hey"],
                     'subject':          ['chemistry','multimedia_lecture',''],
                     'record':           [True, True, False]    ,
                     'meeting_finishes': ['9:45', '11:15', '12:45'],
                     'notification_time':['08:00', '09:00', ''],
                     'meeting_days':     [['M','T','W', 'TH', 'F', 'SU', 'SAT'], ['TH',"T"],['M', "T"]]}

    df = pd.DataFrame(meetings_data, columns=['program', 'meeting_time', 'id', 'password', 'subject', 'record', 'meeting_finishes',  'notification_time', 'meeting_days'])
    flag=True
    while True:
        currentTime=datetime.now().strftime("%H:%M")
        start_time=""
        #check if current time is equal to notf time, flag is there for the notification to appear only once
        if (currentTime in meetings_data['notification_time'] and flag):
            row = df.loc[df['notification_time'] == currentTime]
            #check if current day is equal to the notification day
            if date.today().weekday() in MapDaysToIntWeekdayValue(row.iloc[0,8]):
                sendNotificationReminder( row.iloc[0,1], row.iloc[0,7], row.iloc[0,4])
                flag=False
        if currentTime in meetings_data['meeting_time']:
            row = df.loc[df['meeting_time'] == currentTime]
            if date.today().weekday() in MapDaysToIntWeekdayValue(row.iloc[0,8]):
                flag=True
                start_time = time.time()
                if row.iloc[0,0]=="teams":
                   automate_microsoft_teams(row.iloc[0, 4])
                else:
                   automate_zoom(row.iloc[0,2], row.iloc[0,3])
                   time.sleep(3)
                   if (row.iloc[0, 5]):
                       record_meeting(row.iloc[0, 6])
                   time.sleep(3)
        if (currentTime in meetings_data['meeting_finishes']):
            row = df.loc[df['meeting_finishes'] == currentTime]
            if date.today().weekday() in MapDaysToIntWeekdayValue(row.iloc[0, 8]):
                stopRecordingandExitMeeting(row.iloc[0,0])



scheduleMeetings()
