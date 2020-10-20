# Python script for semi-automatic scheduling for support using Chrome remote desktop
 
# Dependencies:
#   Webdriver for Chrome (version must match the version of chrome)
#   Available at https://chromedriver.chromium.org/
#   Python 3.7 32-bit and the following modules (Which are downloaded on first time setup):
#   selenium
#   pyperclip
#   pynput
#   pywin32

# Copy the 'chromium' directory to C:\Users\Public\


###########################################################################

# Imports standard modules
import subprocess
import sys
from importlib import reload
import time
import os
import ctypes

## Imports module for sending email vai gmail API
#import apitest

# Creates location for additional modules
if not os.path.exists('modules'):
    os.makedirs('modules')

# Adds module location to path
module_location = os.getcwd()+ "\\modules"
sys.path += [module_location]

# Installs modules for first time setup
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package, "-t", "modules\\"])

# Function to calculate and print the time remaining on the remote desktop session
def timeleft(t,l):
    mins, secs = divmod(l-t, 60)
    timeformat = 'This remote desktop session will expire in {:g} mins and {:g} seconds\n'.format(mins, secs)
    _ = os.system('cls')
    print(timeformat)
    print('Please leave this window open unless you intend to close the session\n')
    print('For interactive sessions, while your browser is the active window you can control the mouse and keyboard\n')
    print('For demonstration sessions, avoid using the mouse and keyboard while the browser is your active window\n')
    print('If a message popup appears asking if you want to continue this session, please click continue\n')
    print('Please do not use (or close!) the Chromium browser on the remote computer.\n')
    print('Before your session ends, please save your work and close the applications used during your practical.\n')

# Function to send email through the Gmail server
def send_gmail(email, code):
    
    # Go to Gmail page
    shell.SendKeys('%')
    browser.switch_to.window(browser.window_handles[0])
    time.sleep(0.5)

    # Send guest email with access code
    while True:
        try:
            browser.find_element_by_class_name("z0").click() # Compose button
        except:
            time.sleep(0.5)
        else:
            break
    while True:
        try:
            browser.find_element_by_css_selector("[aria-label='To']").send_keys(email)
        except:
            time.sleep(0.5)
        else: 
            # Writes email using Selenium (more robust to user mouse and keyboard intervention
            browser.find_element_by_css_selector("[aria-label='Subject']").send_keys("Code: {:s}".format(code))
            browser.find_element_by_css_selector(".Am.Al.editable.LW-avf").click()
            browser.find_element_by_css_selector(".Am.Al.editable.LW-avf").send_keys(("Here is your code for remote support: {:s}".format(code),
                                                                                     "\nThis code will only be available for 5 minutes.",
                                                                                     "\nPlease login at https://remotedesktop.google.com/support/ and enter the access code where it says 'Give Support'."))
            browser.find_element_by_class_name("dC").click() # Send button
            ## Writes email using win32gui
            #keyboard.press(Key.tab)
            #keyboard.release(Key.tab)
            #time.sleep(0.5)
            #keyboard.press(Key.tab)
            #keyboard.release(Key.tab)
            #time.sleep(0.5)
            #keyboard.type("Code: "+ code)
            #time.sleep(0.5)
            #keyboard.press(Key.tab)
            #keyboard.release(Key.tab)
            #keyboard.type('Hey there!\nHere is your code for remote support: '+ code +'\nThis code will only be available for 5 minutes')
            #time.sleep(0.5)
            #keyboard.press(Key.tab)
            #keyboard.release(Key.tab)
            #time.sleep(0.5)
            #keyboard.press(Key.enter)
            #keyboard.release(Key.enter)
            #time.sleep(1.0)
            break
    time.sleep(1)

# Toggles on/off mouse if BlueLife KeyFreeze is loaded
# Make sure settings in KeyFreeze apply just to the mouse!
def toggle_freeze():
    # Set focus on desktop 
    win32gui.SetForegroundWindow(win32gui.GetDesktopWindow())
    # Trigger KeyFreeze
    shell.SendKeys('^%f') 
        
# Function for the generation of access codes and email
def generate(email):

    # Initialize flag for successful login
    flag = False
    
    # Load Chrome Remote Desktop support server page
    settings['winnum'] = settings['winnum']+1
    browser.execute_script("window.open('https://remotedesktop.google.com/support');")
    browser.switch_to.window(browser.window_handles[settings['winnum']])
    #browser.get("https://remotedesktop.google.com/support")

    # Click 'No thanks' (we already added the extension)
    try:
        browser.find_element_by_css_selector("[aria-label='No thanks']").click()
    except:
        pass

    # Try to Click 'Generate Code'
    while True:
        try:
            browser.find_element_by_css_selector("[aria-label='Generate Code']").click()
        except:
            time.sleep(0.5)
        else:
            break

    # Copy code to clipboard
    while True:     
        try:
            browser.find_element_by_css_selector("[aria-label='Copy access code to clipboard.']").click()
        except:
            time.sleep(0.5)
        else:
            code = pyperclip.paste()
            break
    time.sleep(1)

    ## Send email using Gmail API     
    #apitest.SendMessage(master_email, email, "Code: "+ code, 'Hey there!<br/>Here is your code for remote support: '+ code +'<br/>This code will only be available for 5 minutes', "null")
        
    # Send email using Gmail server  
    send_gmail(email, code)
   

    # Clear console and explain what is happening
    _ = os.system('cls') 
    print('The setup for Chrome Remote Desktop will be automated.\n')
    print('Please do NOT interact with the mouse or keyboard while connections are being established.\n')
    print('Email invitations will be sent to:')
    print(email_list)
    print('\nWe are currently inviting:')
    print(email+'\n')
    
    # Wait for confirmation for 5 mins and accept
    d = 5  # lag, must be integer
    l = int(min((300-d),session['time'])-30)
    for t in range(l):
        time.sleep(1)
        mins, secs = divmod(l-(t+1), 60)
        sys.stdout.write('\rTheir access code will expire in approximately:  {:g} mins and {:g} seconds'.format(mins, secs))
        handle = win32gui.FindWindow(None,'Chrome Remote Desktop')
        if handle == 0:
            pass
            sys.stdout.flush()
        elif handle != 0:
            root.update() # Make sure splash screen is showing
            time.sleep(0.6)
            shell.SendKeys('%')
            win32gui.SetForegroundWindow(handle)
            time.sleep(0.2)
            keyboard.press(Key.tab)
            keyboard.release(Key.tab)
            win32gui.SetForegroundWindow(handle)
            time.sleep(0.2)
            keyboard.press(Key.enter)
            keyboard.release(Key.enter)
            handle = 0
            flag = True
            break
        
    # Give it a couple of seconds to complete establishing the connection    
    time.sleep(2.0)
    
    # If code was used, 'stop sharing' control bar pops up and takes the focus
    # This control bar gets in the way and guests may accidently press it
    # Hide it since we can terminate the session from withing Chromium
    if flag:
        share_bar[email] = win32gui.GetForegroundWindow()
        win32gui.SetWindowPos(share_bar[email],win32con.HWND_TOP,0,-100,500,100,win32con.SWP_SHOWWINDOW)
        
    return flag


# Setting up for the first time
first_time = 0
def first():
        os.system('cmd /c "python -m pip install --upgrade pip"')
        print('PERFORMING FIRST TIME SETUP')

try:
    from selenium import webdriver
except:
    first_time = first_time + 1
    if first_time == 1:
        first()
    install('selenium')
    from selenium import webdriver
    
try:
    import pyperclip
except: 
    first_time = first_time + 1
    if first_time == 1:
        first()
    install('pyperclip')
    import pyperclip

try:
    from pynput.keyboard import Key, Controller
except: 
    first_time = first_time + 1
    if first_time == 1:
        first()
    install('pynput')
    from pynput.keyboard import Key, Controller

#try:
#    import chromedriver_autoinstaller
#except: 
#    first_time = first_time + 1
#    if first_time == 1:
#        first()
#    install('chromedriver_autoinstaller')
#    import chromedriver_autoinstaller
    
#try:
#    import oauth2client
#    from oauth2client import client, tools, file
#except:
#    install('oauth2client')
#    import oauth2client
#    from oauth2client import client, tools, file

try:
    import win32gui, win32api, win32con, win32com.client
except: 
    first_time = first_time + 1
    if first_time == 1:
        first()
    os.system('cmd /c "pip install pywin32"')
    print('FIRST TIME SETUP ENDED\nPLEASE RESTART THE SCRIPT')
    exit()

try:
    from cryptography.fernet import Fernet
except: 
    first_time = first_time + 1
    if first_time == 1:
        first()
    install('cryptography')
    from cryptography.fernet import Fernet
    
if first_time > 0:
    print('FIRST TIME SETUP ENDED')

# Maximise command window
cmd_handle = win32gui.GetForegroundWindow()
win32gui.ShowWindow(cmd_handle,win32con.SW_MAXIMIZE)

# Password encryption/decryption (depends on serial number of current windows volume)
cwd = os.path.abspath('.')
volume_info=win32api.GetVolumeInformation("{:s}:\\".format(cwd[0]))
volume_serial = str(volume_info[1])
temp = volume_serial + '-' * 43
key = bytes(temp[0:43]+'=','utf-8')
cipher_suite = Fernet(key) 
#encryptedpwd = cipher_suite.encrypt(b"password") # substitute word 'password' for google account password
#with open('./passcode.bin', 'wb') as file_object:  file_object.write(encryptedpwd) # write passcode.bin file 
with open('./passcode.bin', 'rb') as file_object:
    for line in file_object:
        passcode = line
try:
    pw = (cipher_suite.decrypt(passcode)).decode('utf-8')
except:
    print("The passcode does not match the encryption key")
    input("Press any key to close the terminal window")
    quit()

# Initialize winnum and share_bar dictionaries
settings = {'winnum':0}
share_bar = {}

# Setup keyboard controller
keyboard = Controller()

# Checks if the chrome + chromedriver versions are compatable
#chromedriver_autoinstaller.install()

# Load email text file
# Clear console and explain what is happening
#gmail_dir = input('Enter the path to email text file:\n')
_ = os.system('cls')
gmail_dir = ''
if gmail_dir == '':
    gmail_dir = '.\email.txt'
f = open(gmail_dir, 'r')
line = f.readlines()
master_email = line[0]
f.close()
l = input('Enter the maximum duration of your remote desktop session in minutes (default is 60):\n')
if l == '':
    l = 60.0
else:
    l = float(l)
l = l*60 # convert maximum duration to seconds
l -= 30  # terminate 30 seconds earlier in case another remote desktop session is scheduled after this one
session={}
session['time']=l # required for generate function in case code expires after session end time

# Guest email address(es)
email_list = input('\nEnter (space-separated) list of email addresses of the guests:\n') 
# Schedule a future CRD session
sched = input('\nEnter local time to start sending email invitations to the guests (hh:mm):\n') 
if sched:
    now = time.localtime()
    then = list(now)
    sched_list = sched.split(":")
    then[3:5] = [int(i) for i in (sched.split(":")+[0])]
    then = time.struct_time(then)
    if (then < now):
        print("\nThe scheduled time must be today sometime in the future.")
        print("\nPress any key to exit.")
        input()
        quit()
    now_str = "{:d}".format(now.tm_hour).rjust(2,"0") + ":" + "{:d}".format(now.tm_min).rjust(2,"0")
    now_day = now.tm_yday
    print('\nCRDX is scheduled to invite guests for a remote desktop session at {:s}.'.format(sched))
    print('\nPlease leave this window open.')
    time.sleep(5)
    # Minimize command window
    win32gui.ShowWindow(cmd_handle,win32con.SW_MINIMIZE)
    while True:
        if (now_str == sched):
            break
        else:
            time.sleep(1)
            # Update what the time is now
            now = time.localtime()
            now_str = "{:d}".format(now.tm_hour).rjust(2,"0") + ":" + "{:d}".format(now.tm_min).rjust(2,"0")
            # Check date
            if (now_day == now.tm_yday):
                pass
            else:
                quit()
else:
    pass

# Freeze mouse input
shell = win32com.client.Dispatch("WScript.Shell")
toggle_freeze()

# Maximise command window
win32gui.ShowWindow(cmd_handle,win32con.SW_MAXIMIZE)

# Create splash screen
import tkinter as tk
root = tk.Tk()
root.attributes('-topmost', True)
# Get screen info
root.overrideredirect(True)
w = root.winfo_screenwidth()
h = root.winfo_screenheight()
# Get image and set size and position
image_file = "crd_600x400.png"
root.geometry('%dx%d+%d+%d' % (600, 400, w-650, h-450))
image = tk.PhotoImage(file=image_file)
canvas = tk.Canvas(root, height=400, width=600, bg="green")
canvas.create_image(300, 200, image=image)
canvas.pack()
# Display the splash screen
root.update()

# Start timer
tic = time.time()

# Start chrome webdriver with chrome remote desktop extension
# Drive chromium since it does not have automatic updates
chromeOptions = webdriver.ChromeOptions()
if os.path.exists("C:/Users/Public/Documents/chromium"):
    chromeOptions.binary_location = "C:/Users/Public/Documents/chromium/bin/chrome.exe"
    chromeOptions.add_extension("C:/Users/Public/Documents/chromium/ext/inomeogfingihgjfjlpeplalcfajhgai.crx") # Chrome remote desktop extension
else:
    chromeOptions.binary_location = "S:/crd/chromium/bin/chrome.exe"
    chromeOptions.add_extension("S:/crd/chromium/ext/inomeogfingihgjfjlpeplalcfajhgai.crx") # Chrome remote desktop extension
    chromeOptions.add_argument("--no-sandbox")
chromeOptions.add_argument("--window-size=800,600")
chromeOptions.add_argument("--window-position=-2000,0")
browser = webdriver.Chrome(options=chromeOptions)
browser_window = win32gui.GetForegroundWindow()

## Log into remote.test.student google account
#browser.get('https://stackoverflow.com/users/signup?ssrc=head&returnurl=%2fusers%2fstory%2fcurrent')
#time.sleep(0.5)
#browser.find_element_by_xpath('//*[@id="openid-buttons"]/button[1]').click()
## Tries to click "Use Another Account"
#try:
#    new_acc = browser.find_element_by_css_selector("[jsname='rwl3qc']").click()
#except:
#    pass
#time.sleep(0.5)

# Log into remote.test.student google account
# Load login page
browser.get('https://accounts.google.com/signin/v2/identifier?')
time.sleep(0.5)

# Tries to type in email and password
try:
    master = browser.find_element_by_xpath('//*[@id="identifierId"]')
except:
    pass
else:
    master.clear()
    master.send_keys(master_email)
    try:
        browser.find_element_by_id("identifierNext").click()
        time.sleep(1)
    except:
        pass
    for _ in range(60):
        try:
            passw = browser.find_element_by_name("password")
            browser.find_element_by_name("password").clear()
        except:
            time.sleep(0.5)
        else:
            browser.find_element_by_name("password").send_keys(pw)
            try:
                browser.find_element_by_id("passwordNext").click()
            except:
                pass
            break

# Go to Gmail when sign-in is complete
for _ in range(60):
    try:
        # Check sign-in is complete and Google account settings page has loaded
        browser.find_element_by_xpath("//a[@title='Google Account settings']")
    except:
        time.sleep(0.5)
    else:
        browser.get("https://mail.google.com/mail/u/0/#inbox")
        break

# Tries to click "Close" if security warning appears
try:
    browser.find_element_by_css_selector("[aria-label='Close']").click()
except:
    pass
time.sleep(0.5)

# Loop through email addresses establishing remote desktop support
email_list = email_list.split()
for email in email_list:
    generate(email)
toggle_freeze()
    
# If no connections were made the share_bar dictionary will be empty
# If share_bar is empty then quit
if share_bar:
    pass
else:
    root.destroy()
    browser.quit()
    quit()

# Minimize command window
win32gui.ShowWindow(cmd_handle,win32con.SW_MINIMIZE)
 
# Hide splash screen (so it can be retrieved later if necessary
root.withdraw()

# --- we need to keep the google account logged in if we wish to re-invite guests automatically during the session
## Log out of gmail and close tab
#browser.switch_to.window(browser.window_handles[0])
#browser.get("https://mail.google.com/mail/?logout&hl=en")
#browser.close()
#browser.switch_to.window(browser.window_handles[0])

# Clear console from the block of errors that happen
_ = os.system('cls')

# Stop timer and subtract delay from duration of the remote desktop session
toc = time.time()
d = toc-tic
l = int(l-d)
session['time']=l

# 30 second final countdown function
import threading
def final_countdown():
    for i in range(30):
        timeleft(i+1,30)
        time.sleep(1)
    browser.quit()
    quit()
    
# Limit the time of the remote desktop session
t = 0
d = 0
while True:
    handle = win32gui.FindWindow(None,'Chrome Remote Desktop')
    if handle == 0:
        time.sleep(1-d)
        tic = time.time()
        t += 1
        timeleft(t,l)
        session['time'] = int(l-t)
        # Check if browser is still open, if not then quit
        if win32gui.IsWindow(browser_window):
            pass
        else:
            quit()
        if t >= l:
            browser.quit()
            quit() 
        # 30 second warning
        # Message box will timeout in just under 30 seconds
        if (l-t) < 30:
            win32gui.ShowWindow(cmd_handle,win32con.SW_MAXIMIZE)
            win32gui.MoveWindow(cmd_handle,0,0,1000,600,True)
            countdown = threading.Thread(target=final_countdown)
            countdown.start()
            tle = "Chrome Remote Desktop Session Alert"
            msg = "This remote desktop session will terminate within 30 seconds.\nPlease save and close your work."
            ctypes.windll.user32.MessageBoxTimeoutW(0, msg, tle, 0x42030,0, 29000)
            break
        # Check if share_bar is empty, if it is then quit
        if share_bar:
            pass
        else:
            browser.quit()
            quit()
        # Check if each guest is still logged in
        for key in share_bar.keys():
            if win32gui.IsWindow(share_bar[key]):
                pass
            else:
                if win32gui.IsWindow(browser_window):
                    email = key
                    share_bar[email] = False # This guest has disconnected
                    any_guest = False
                    for value in share_bar.values():
                        if value:
                            any_guest = True
                    if any_guest:
                        # Ask remaining guests if they want to reconnect missing guest
                        # Message box will timeout in 60 seconds
                        tle = "Chrome Remote Desktop Session Alert"
                        msg = "Guest {:s} has disconnected.\nDo you want to invite them back?".format(email)
                        ans = ctypes.windll.user32.MessageBoxTimeoutW(0, msg, tle, 0x42034,0, 60000)
                        if (ans == 32000): # Message box timed-out
                            ans = 6 # Set to Yes if Message box timed-out
                    else:
                        ans = 6 # Yes since there are no other guests
                    if ans == 6: # If Yes
                        # Prepare to show splash screen again
                        root.deiconify()
                        # Freeze mouse
                        toggle_freeze()
                        # Load splash screen
                        root.update()
                        # Maximise command window
                        win32gui.ShowWindow(cmd_handle,win32con.SW_MAXIMIZE)
                        # Send invitation to missing guest
                        flag = generate(email)
                        if (flag == False):
                            share_bar[email] = False
                        # Hide splash screen 
                        root.withdraw()
                        # Unfreeze mouse
                        toggle_freeze()
                        # Minimize command window
                        win32gui.ShowWindow(cmd_handle,win32con.SW_MINIMIZE) 
                        toc = time.time()
                        l -= int(toc-tic) # Correct time remaining
                        session['time'] = l
                        tic = time.time()
                    else:
                        toc = time.time()
                        l -= int(toc-tic) # Correct time remaining
                        session['time'] = l
                        tic = time.time()
                else:
                    quit()
        # Remove guests who are not reinvited or do not log back in within 5 minutes
        [share_bar.pop(key) for key,val in tuple(share_bar.items()) if (val == False)]
        toc = time.time()
        d = min(1,toc-tic) # calculate software delay (max 1 sec)
    else:
        while handle:
            # Try accepting to continue allowing access until we are successfull
            try:
                time.sleep(0.2)
                shell.SendKeys('%')
                win32gui.SetForegroundWindow(handle)
                time.sleep(0.1) # if guests click during this time it will break the script
                win32api.SetCursorPos((200,100))
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,200,100,0,0)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,200,100,0,0)
                time.sleep(0.2)
            except:
                time.sleep(0.3)
                pass
            handle = win32gui.FindWindow(None,'Chrome Remote Desktop')
        print("Allowing access")
        handle = 0
        time.sleep(2)
        t += 2
        
