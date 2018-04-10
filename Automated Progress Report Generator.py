#! python3

## This program uses gui automation to access an internal medical equipment management system on my desktop 
## in order to generate a scheduled maintenance report that I use to track the performance of my department.
## Out of the box, the report is not very useful, so I have this program initiate some excel
## macros to clean up the report and save it to a directory on my PC.



import pyautogui, sys, time, datetime, pyperclip, webbrowser


phase = 0

def cycle(phase):  ## Provides visual feedback of distinct program phases in console window for review and troubleshooting purposes
    phase = int(phase)
    if phase == 0:
        Start = "START"
        print(Start.center(100, "-"))
    elif 0 < phase < 14:
        Sequence = "PHASE_" + str(phase)
        print(Sequence.center(100, "-"))
    else:
        Complete = "COMPLETE"
        print(Complete.center(100, "-"))

cycle(phase) ##Start

now = datetime.datetime.now()
print(str(now))
Start_Time = time.time()
print(Start_Time)

phase +=1 ##Phase 1
cycle(phase) 


Entry_Date = (now.strftime("%m/01/%Y")) ## This selects a report period dating back to the beginning of the present month
print("Entry Date = " + Entry_Date)
File_Name = (now.strftime("%m%d%Y_Incomplete_PMs"))
print("File Name = " + File_Name)
File_Path = r"C:\Users\richarl2\Documents\My Docs\CaroMont Health\Operations\Incomplete PMs\2018"
print("File Path = " + File_Path)

phase +=1 ##Phase 2
cycle(phase) 

webbrowser.open("https://idesk.aramark.net/Login")
print("URL Command Sent for iDESK.Aramark.com")

phase +=1 ##Phase 3
cycle(phase)

time.sleep(4.0)
Idesk_Maximize = pyautogui.locateOnScreen("Idesk_Maximize.png", grayscale=True)
Idesk_Minimize = pyautogui.locateOnScreen("Idesk_Minimize.png", grayscale=True)

while Idesk_Minimize == None or Idesk_Minimize == None:
        Idesk_Maximize = pyautogui.locateOnScreen("Idesk_Maximize.png", grayscale=True)
        Idesk_Minimize = pyautogui.locateOnScreen("Idesk_Minimize.png", grayscale=True)
        print("Still searching for Idesk Minimize or Maximize Button...")

        if Idesk_Maximize != None:
            #When program finds Idesk_Maximize image, print the location
            time.sleep(2.0)         
            print("Idesk Maximize is located at {0}".format(Idesk_Maximize))
            Idesk_Maximize_X, Idesk_Maximize_Y = pyautogui.center(Idesk_Maximize)
            print("Center point of Idesk_Maximize is located at ({0}, {1})".format(Idesk_Maximize_X, Idesk_Maximize_Y))
            pyautogui.moveTo(Idesk_Maximize_X, Idesk_Maximize_Y)
            pyautogui.click(clicks=1) # click Idesk Maximize
            break               
        
        
phase +=1 ##Phase 4
cycle(phase)

Idesk_Login = 'username' # Will need to be updated prior to implementation
Idesk_Password = "password" # Will need to be updated prior to implementation

time.sleep(3.0)
pyperclip.copy(Idesk_Login)
time.sleep(0.5)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
pyperclip.copy(Idesk_Password)
time.sleep(0.5)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(3.0)


Scheduled_Work_Orders_Button = pyautogui.locateOnScreen("Scheduled_Work_Orders_Button.png", grayscale=True)

#Searches for the Scheduled_Work_Orders_Button Icon
while Scheduled_Work_Orders_Button == None:
    Scheduled_Work_Orders_Button = pyautogui.locateOnScreen("Scheduled_Work_Orders_Button.png", grayscale=True)
    print("Still searching for Scheduled Work Orders Button...")

#When program finds image, print the location
print("Scheduled_Work_Orders_Button is located at {0}".format(Scheduled_Work_Orders_Button))
Scheduled_Work_Orders_Button_X, Scheduled_Work_Orders_Button_Y = pyautogui.center(Scheduled_Work_Orders_Button)
print("Center point of Scheduled_Work_Orders_Button is located at ({0}, {1})".format(Scheduled_Work_Orders_Button_X, Scheduled_Work_Orders_Button_Y))

phase +=1 ##Phase 5
cycle(phase)

pyautogui.moveTo(Scheduled_Work_Orders_Button_X, Scheduled_Work_Orders_Button_Y)
pyautogui.click(clicks=1) # click Idesk Shortcut
time.sleep(3.0) 


Start_Date_Input_Field = (830, 353)
pyautogui.moveTo(Start_Date_Input_Field)
pyautogui.click(clicks=1)  # click in the Start Date Field

pyautogui.pause = 0.5
pyautogui.typewrite(Entry_Date, interval=0.12)
pyautogui.press('enter')

pyautogui.pause = 1.2
Export_Button = (1400, 181)
pyautogui.moveTo(Export_Button)
pyautogui.click(clicks=1)  # click the Export Button
print("Export_Button is located at {}".format(Export_Button))
time.sleep(3.0) 

phase +=1  ##Phase 6
cycle(phase) 

Open_Button = pyautogui.locateOnScreen("Open_Button.png", grayscale=True)
Open_Button_2 = pyautogui.locateOnScreen("Open_Button_2.png", grayscale=True)


#Searches for the Open_Button Icon
while Open_Button == None and Open_Button_2 == None:
    Open_Button = pyautogui.locateOnScreen("Open_Button.png", grayscale=True)
    Open_Button_2 = pyautogui.locateOnScreen("Open_Button_2.png", grayscale=True)
    print("Still searching for Open_Button 1 or 2...")

if Open_Button != None:
        time.sleep(1.0)
        print("Open_Button is located at {0}".format(Open_Button)) 
        Open_Button_X, Open_Button_Y = pyautogui.center(Open_Button)
        print("Center point of Open_Button is located at ({0}, {1})".format(Open_Button_X, Open_Button_Y))
        pyautogui.moveTo(Open_Button_X, Open_Button_Y)
        pyautogui.click(clicks=1) # click Maximize_Open_Button    

elif Open_Button_2 != None:
        time.sleep(1.0)
        print("Open_Button_2 is located at {0}".format(Open_Button_2)) 
        Open_Button_2_X, Open_Button_2_Y = pyautogui.center(Open_Button_2)
        print("Center point of Open_Button_2 is located at {(0}, {1})".format(Open_Button_2_X, Open_Button_2_Y))
        pyautogui.moveTo(Open_Button_2_X, Open_Button_2_Y)
        pyautogui.click(clicks=1) # click Maximize_Open_Button_2    

else:
        print("This just is not working!")    

time.sleep(4.0)


pyautogui.pause = 1.5
Maximize_Excel_Button = pyautogui.locateOnScreen("Maximize_Excel_Button.png", grayscale=True)
Minimize_Excel_Button = pyautogui.locateOnScreen("Minimize_Excel_Button.png", grayscale=True)


while Minimize_Excel_Button == None and Maximize_Excel_Button == None:
        Maximize_Excel_Button = pyautogui.locateOnScreen("Maximize_Excel_Button.png", grayscale=True)
        Minimize_Excel_Button = pyautogui.locateOnScreen("Minimize_Excel_Button.png", grayscale=True)
        print("Still searching for Excel Minimize or Maximize Button...")
        

if Maximize_Excel_Button != None:
    #When program finds Maximize_Excel_Button image, print the location
    time.sleep(1.0)         
    print("Maximize_Excel_Button is located at {0}".format(Maximize_Excel_Button))
    Maximize_Excel_Button_X, Maximize_Excel_Button_Y = pyautogui.center(Maximize_Excel_Button)
    print("Center point of Maximize_Excel_Button is located at ({0}, {1})".format(Maximize_Excel_Button_X, Maximize_Excel_Button_Y))
    pyautogui.moveTo(Maximize_Excel_Button_X, Maximize_Excel_Button_Y)
    pyautogui.click(clicks=1) # click Maximize_Excel_Button

                         
phase +=1  ##Phase 7
cycle(phase)

Cell_A1 = (55,317) ## Finds cell A1 of excel file
pyautogui.moveTo(Cell_A1)
pyautogui.click(clicks=1)
print("Found Cell A1")
pyautogui.PAUSE = 0.5


pyautogui.hotkey('ctrl', 'shift', 'i') # Initializes excel Macro to clean up incomplete work orders report
print("Initiated Excel Macro #1")
time.sleep(1.0)

File_Tab = pyautogui.locateOnScreen("File_Tab.png", grayscale=True)

while File_Tab == None:
        File_Tab = pyautogui.locateOnScreen("File_Tab.png", grayscale=True)
        print("Still searching for File_Tab...")
    
#When program finds image, print the location
print("File_Tab is located at {0}".format(File_Tab))
File_Tab_X, File_Tab_Y = pyautogui.center(File_Tab)
print("Center point of File_Tab is located at ({0}, {1})".format(File_Tab_X, File_Tab_Y))


phase +=1 ##Phase 8
cycle(phase) 

pyautogui.moveTo(File_Tab_X, File_Tab_Y)
pyautogui.click(clicks=1) # click File_Tab
time.sleep(1.0)

Save_As_Button = pyautogui.locateOnScreen("Save_As_Button.png", grayscale=True)

while Save_As_Button == None:
        Save_As_Button = pyautogui.locateOnScreen("Save_As_Button.png", grayscale=True)
        print("Still searching for Save_As_Button...")
    
#When program finds image, print the location
print("Save_As_Button is located at {0}".format(Save_As_Button))
print(Save_As_Button)
Save_As_Button_X, Save_As_Button_Y = pyautogui.center(Save_As_Button)
print("Center point of Save_As_Button is located at ({0}, {1})".format(Save_As_Button_X, Save_As_Button_Y))
phase +=1

cycle(phase) ##Phase 9
pyautogui.moveTo(Save_As_Button_X, Save_As_Button_Y)
pyautogui.click(clicks=1) # click File_Tab
time.sleep(0.5)


Computer_Button = pyautogui.locateOnScreen("Computer_Button.png", grayscale=True)

while Computer_Button == None:
        Computer_Button = pyautogui.locateOnScreen("Computer_Button.png", grayscale=True)
        print("Still searching for Computer_Button...")
    
#When program finds Computer_Button image, print the location
print("Computer_Button is located at {0}".format(Computer_Button))
print(Computer_Button)
Computer_Button_X, Computer_Button_Y = pyautogui.center(Computer_Button)
print("Center point of Computer_Button is located at ({0}, {1})".format(Computer_Button_X, Computer_Button_Y))


phase +=1  ##Phase 10
cycle(phase)

pyautogui.moveTo(Computer_Button_X, Computer_Button_Y)
pyautogui.click(clicks=2) # click Computer_Button
time.sleep(1.5)


pyperclip.copy(File_Name)
time.sleep(0.5)
pyautogui.hotkey("ctrl", "v")
time.sleep(0.5)
pyautogui.press('tab')
time.sleep(1.0)


Save_As_Type_Highlighted = pyautogui.locateOnScreen("Save_As_Type_Highlighted.png", grayscale=True)

while Save_As_Type_Highlighted == None:
        Save_As_Type_Highlighted = pyautogui.locateOnScreen("Save_As_Type_Highlighted.png", grayscale=True)
        print("Still searching for Save_As_Type_Highlighted...")
    
#When program finds Save_As_Type_Highlighted image, print the location
print("Save_As_Type_Highlighted is located at {0}".format(Save_As_Type_Highlighted))
Save_As_Type_Highlighted_X, Save_As_Type_Highlighted_Y = pyautogui.center(Save_As_Type_Highlighted)
print("Center point of Save_As_Type_Highlighted is located at ({0}, {1})".format(Save_As_Type_Highlighted_X, Save_As_Type_Highlighted_Y))


phase +=1  ##Phase 11
cycle(phase)

pyautogui.moveTo(Save_As_Type_Highlighted_X, Save_As_Type_Highlighted_Y)
pyautogui.click(clicks=1) # click Save_As_Type_Highlighted
time.sleep(1.0)

pyautogui.pause = 0.25
for i in range (1,15):  ## Arrows up to the excel workbook selection (.xlsx)
        pyautogui.press('up')
pyautogui.press('tab')
time.sleep(1.0)

pyautogui.pause = 0.75
Path_Bar = pyautogui.locateOnScreen("Path_Bar.png", grayscale=True)

while Path_Bar == None:
        Path_Bar = pyautogui.locateOnScreen("Path_Bar.png", grayscale=True)
        print("Still searching for Path_Bar...")
    
#When program Path_Bar image, print the location
print("Path_Bar is located at {0}".format(Path_Bar))
Path_Bar_X, Path_Bar_Y = pyautogui.center(Path_Bar)
print("Center point of Path_Bar is located at ({0}, {1})".format(Path_Bar_X, Path_Bar_Y))


phase +=1 ##Phase 12
cycle(phase) 

pyautogui.moveTo(Path_Bar_X + 680, Path_Bar_Y)
pyautogui.click(clicks=1) # click Path_Bar
time.sleep(0.5)
pyautogui.press('delete')
pyperclip.copy(File_Path)
time.sleep(0.5)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

pyautogui.pause = 0.25
for i in range (1,17):  ## Tabs to the save_button
        pyautogui.press('tab')
pyautogui.press('enter')


for i in range (1,6):
       Yes_Button = pyautogui.locateOnScreen("Yes_Button.png", grayscale=True) # Yes_Button will only appear as an inquiry to the question, "This file already exists. Do you want to replace it?"


if Yes_Button == None:
        Yes_Button = pyautogui.locateOnScreen("Yes_Button.png", grayscale=True)
        print("Yes_Button not found. Moving On...")

elif Yes_Button != None:
        print("Yes_Button is located at {0}".format(Yes_Button))
        Yes_Button_X, Yes_Button_Y = pyautogui.center(Yes_Button)
        print("Center point of Yes_Button is located at ({0}, {1})".format(Yes_Button_X, Yes_Button_Y))
        pyautogui.moveTo(Yes_Button_X, Yes_Button_Y)
        pyautogui.click(clicks=1) # click Yes_Button



phase +=1  ##Phase 13
cycle(phase)


pyautogui.hotkey('ctrl', 'shift', 'p')
print("Initiated Excel Macro #2")
pyautogui.pause = 0.75
time.sleep(0.5)

pyperclip.copy("") # Clears Clipboard
print("Clipboard cleared")
now = datetime.datetime.now()
print(str(now))
print("Program Run Time = %s seconds " % (time.time() - Start_Time))


phase +=1  ##Complete
cycle(phase)
exit()
