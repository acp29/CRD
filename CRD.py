# Python script for semi-automatic setup for support using Chrome remote desktop
 
# Dependencies:
#   Webdriver for Chrome (version must match the version of chrome)
#   Available at https://chromedriver.chromium.org/
#   Python 3 and the following modules (Which are downloaded on first time setup):
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
    win32gui.SetForegroundWindow(browser_window)
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
    d = 6  # lag, must be integer
    l = 300 - d # must be integer
    for t in range(l):
        time.sleep(1)
        mins, secs = divmod(l-t, 60)
        sys.stdout.write('\rTheir access code will expire in approximately:  {:g} mins and {:g} seconds'.format(mins, secs))
        handle = win32gui.FindWindow(None,'Chrome Remote Desktop')
        if handle == 0:
            pass
            sys.stdout.flush()
        elif handle != 0:
            time.sleep(0.3)
            time.sleep(0.3)
            shell.SendKeys('%')
            print(handle)
            win32gui.SetForegroundWindow(handle)
            time.sleep(0.2)
            keyboard.press(Key.tab)
            keyboard.release(Key.tab)
            time.sleep(0.2)
            keyboard.press(Key.enter)
            keyboard.release(Key.enter)
            handle = 0
            flag = True
            break
        
    # Give it a few second to complete establishing the connection    
    time.sleep(5.0)
    
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

try:
    import chromedriver_autoinstaller
except: 
    first_time = first_time + 1
    if first_time == 1:
        first()
    install('chromedriver_autoinstaller')
    import chromedriver_autoinstaller
    
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
chromedriver_autoinstaller.install()

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
# Guest email address(es)
email_list = input('\nEnter (space-separated) list of email addresses of the guests:\n') 

print('\nPlease wait while we load the browser and log into the remote.test.student Google account...\n')

# Start chrome webdriver with chrome remote desktop extension
# Drive chromium since it does not have automatic updates
chromeOptions = webdriver.ChromeOptions()
if os.path.exists("C:/Users/Public/chromium"):
    chromeOptions.binary_location = "C:/Users/Public/chromium/bin/chrome.exe"
    chromeOptions.add_extension("C:/Users/Public/chromium/ext/inomeogfingihgjfjlpeplalcfajhgai.crx") # Chrome remote desktop extension
else:
    chromeOptions.binary_location = "S:/crd/chromium/bin/chrome.exe"
    chromeOptions.add_extension("S:/crd/chromium/ext/inomeogfingihgjfjlpeplalcfajhgai.crx") # Chrome remote desktop extension
    chromeOptions.add_argument("--no-sandbox")
chromeOptions.add_argument("--window-size=800,600")
chromeOptions.add_argument("--window-position=-2000,0")
browser = webdriver.Chrome(options=chromeOptions)
browser_window = win32gui.GetForegroundWindow()

# Log into remote.test.student google account
browser.get('https://stackoverflow.com/users/signup?ssrc=head&returnurl=%2fusers%2fstory%2fcurrent')
browser.find_element_by_xpath('//*[@id="openid-buttons"]/button[1]').click()

# Tries to click "Use Another Account"
try:
    new_acc = browser.find_element_by_css_selector("[jsname='rwl3qc']").click()
except:
    pass
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
time.sleep(2)

# Google to Gmail 
browser.get("https://mail.google.com/mail/u/0/#inbox")

# Tries to click "Close" if security warning appears
try:
    browser.find_element_by_css_selector("[aria-label='Close']").click()
except:
    pass
time.sleep(0.5)

# Loop through email addresses establishing remote desktop support
shell = win32com.client.Dispatch("WScript.Shell")
email_list = email_list.split()
for email in email_list:
    generate(email)
    
# If no connections were made the share_bar dictionary will be empty
# If share_bar is empty then quit
if share_bar:
    pass
else:
    browser.quit()
    quit()

# --- we need to keep the google account logged in if we wish to re-invite guests automatically during the session
## Log out of gmail and close tab
#browser.switch_to.window(browser.window_handles[0])
#browser.get("https://mail.google.com/mail/?logout&hl=en")
#browser.close()
#browser.switch_to.window(browser.window_handles[0])

# Clear console from the block of errors that happen
_ = os.system('cls') 

# Limit the time of the remote desktop session
t = 0
l = l*60 # convert maximum duration to seconds
while True:
    handle = win32gui.FindWindow(None,'Chrome Remote Desktop')
    if handle == 0:
        time.sleep(1)
        t += 1
        timeleft(t,l)
        if t >= l:
            browser.quit()
            quit()
        # Check if browser is still open, if not then quit
        if win32gui.IsWindow(browser_window):
            pass
        else:
            quit()
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
                    # Send invitation to missing guest
                    email = key
                    tic = time.time()
                    flag = generate(email)
                    if (flag == False):
                        share_bar[email] = False
                    toc = time.time()
                    t += int(toc-tic) # Correct time remaining
                else:
                    quit()
        # Remove guests who do not log back in within 5 minutes
        [share_bar.pop(key) for key,val in tuple(share_bar.items()) if (val == False)]
    else:
        time.sleep(0.2)
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(handle)
        time.sleep(0.1) # if guests click during this time it will break the script
        win32api.SetCursorPos((200,100))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,200,100,0,0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,200,100,0,0)
        print("Allowing access")
        handle = 0
        time.sleep(2)
        t += 2
        
