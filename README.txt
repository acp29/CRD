CRD: Chrome Remote Desktop Automation script
Version 1.0.0 (2020/09/20)
Copyright © 2020 Andrew Penn and Samuel Liu, University of Sussex. All Rights Reserved.
Contact: Andrew Penn (A.C.Penn@sussex.ac.uk) or Samuel Liu (samuelliuuk@gmail.com)

Description:
This script grants control and use of this computer to an email. 
This is achieved through Chrome Remote Desktop which generates a code for another user to enter and gain remote control. 
The code expires 5 minutes after generation, but once the student has access to the computer, it will not expire and the session can be terminated by closing the command prompt.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dependancies:

Python 3 in %PATH%
pip in %PATH%
pywin32 in path_to_python/Lib/site-packages (achieveable by typing "pip install pywin32" in a command prompt -ADMIN)
command prompt QuickEdit Mode DISABLED (Right click, propertise)
Give Gmail API permissions (Detailed below) + "pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib" in a command prompt -ADMIN
Create a txt file containing master email detailts (Detailed below. Recommeded to be named email.txt and put in this directory)

(NB. tested with Google Chrome Version 85.0.4183.83 (Official Build) (64-bit))

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Giving gmail API permissions. This allows for automated sending of emails 
THIS IS NO LONGER USED BY CRD 

To do this for a new master email:
0. Delete credentials.json + token.pickle from this directory. You will be generating your own for the new email.
1. Click "Enable the Gmail API" here: https://developers.google.com/gmail/api/quickstart/python
2. Enter project name Quickstart
3. Select Desktop App
4. Click Download Client Configuration and save the file in the same directory as the script (credentials.json)
5. In a command prompt, run "pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib"
6. Run the script named quickstart.py in this directory (this should generate token.pickle)
7. A window will open asking you to log in and grant quickstart access to view your email.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

On each new PC:
8. Upon launching the chrome remote desktop script (control.py) on a computer for the first time, a new window will appear asking you to log in and grant quickstart permission to send an email.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Using the chrome remote desktop script

0. Before you run the script you first need to create a file with the master gmail address on the first line and the password on the second line. Its advised you make this file hidden.

1. Given all the API permission are completed, on the first time setup, many modules will be downloaded into a local folder named modules (this is to get around any permissions). You may need to restart the script after module installation
2. Upon launching the script, you will be prompted to enter the path to the file with the gmail details. If no path is given, it will attempt to find a file named "email.txt" in this directory
3. Then you will be prompted to enter the student email addresses (each seperated by a space)

The console will be cleared when the student gains access to the computer so location of the gmail details will remain hidden
Remind the students not to close google chrome or the console until they are done with the computer

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

-Enter the path to gmail address
C:\Users\Admin\Documents\email.txt
(The encrypted password should be stored in passcode.bin)

-Enter the email address of the student
Student123@sussex.ac.uk

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


