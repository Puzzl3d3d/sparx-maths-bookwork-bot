import os
from ctypes import windll, Structure, c_long, byref
import time

class POINT(Structure):
    _fields_ = [("x", c_long), ("y", c_long)]
def queryMousePosition():
    pt = POINT()
    windll.user32.GetCursorPos(byref(pt))
    return { "x": pt.x, "y": pt.y}


# Auto-download libraries
try:
    from PIL import ImageGrab, Image, ImageShow
except:
    os.system("pip install Pillow")
    from PIL import ImageGrab, Image
try:
    import pytesseract
except:
    os.system("pip install pytesseract")
    import pytesseract
try:
    import numpy
except:
    os.system("pip install numpy")
    import numpy
try:
    import cv2
except:
    os.system("pip install cv2")
    import cv2
try:
    import keyboard
except:
    os.system("pip install keyboard")
    import keyboard
try:
    from pynput.keyboard import Key, Listener
except:
    os.system("pip install pynput")
    from pynput.keyboard import Key, Listener

pytesseract.pytesseract.tesseract_cmd =rf'{os.path.abspath(os.path.dirname(__file__))}\Tesseract-OCR\tesseract.exe'

settings_file = open("settings.txt", "r+")
lines = settings_file.readlines()

x_q_coord = lines[1].strip()
y_q_coord = lines[2].strip()
x_c_coord = lines[4].strip()
y_c_coord = lines[5].strip()
x_a_coord = lines[7].strip()
y_a_coord = lines[8].strip()

q_coords = (x_q_coord.split(":")[0], x_q_coord.split(":")[1], y_q_coord.split(":")[0], y_q_coord.split(":")[1])
c_coords = (x_c_coord.split(":")[0], x_c_coord.split(":")[1], y_c_coord.split(":")[0], y_c_coord.split(":")[1])
a_coords = (x_a_coord.split(":")[0], x_a_coord.split(":")[1], y_a_coord.split(":")[0], y_a_coord.split(":")[1])

q_coords = (int(q_coords[0]), int(q_coords[1]), int(q_coords[2]), int(q_coords[3]))
c_coords = (int(c_coords[0]), int(c_coords[1]), int(c_coords[2]), int(c_coords[3]))
a_coords = (int(a_coords[0]), int(a_coords[1]), int(a_coords[2]), int(a_coords[3]))

settings_file.close()

print(
"""INPUTS:
    --     Main    --
    F2  - Automatically screenshots current Bookcode, saves image to folder.
    F4  - Find the bookcode of the check and answer that is detected, and opens the image. (For Bookwork Checks)

    -- Calibration --
    F8  - Manually calibrate screen coords of Bookwork Code (for questions)
    F9  - Manually calibrate screen coords of Bookwork Check Code (for bookwork checks)
    F10 - Manually calibrate screen coords of Answer section (for saving answers)

    END - Clear any manual calibration (resets to normal)
    
    DISCLAIMER: If you don't calibrate the screen coordinates correctly, the bot will not work until fixed, and you don't need to calibrate it if it already works.
    """
    
)

DIR = os.path.abspath(os.path.dirname(__file__))+r"\BookworkCodes"

calibrating = {
    "q": {
        "i": (0,0),
        "j": (0,0)
    },
    "c": {
        "i": (0,0),
        "j": (0,0)
    },
    "a": {
        "i": (0,0),
        "j": (0,0)
    }
}
def GetBookworkCodeImage():
    unencoded = ImageGrab.grab(bbox=q_coords)
    screen = numpy.array(unencoded)
    return screen, unencoded
def GetBookworkCheckImage():
    unencoded = ImageGrab.grab(bbox=c_coords)
    screen = numpy.array(unencoded)
    return screen, unencoded
def GetAnswerImage():
    unencoded = ImageGrab.grab(bbox=a_coords)
    return unencoded
def OpenImage(screen):
    cv2.imshow('Bookwork Check Answer', screen)
def Update(data):
    global q_coords
    global c_coords
    global a_coords
    x_q_coord = data[1].strip()
    y_q_coord = data[2].strip()
    x_c_coord = data[4].strip()
    y_c_coord = data[5].strip()
    x_a_coord = data[7].strip()
    y_a_coord = data[8].strip()

    q_coords = (x_q_coord.split(":")[0], x_q_coord.split(":")[1], y_q_coord.split(":")[0], y_q_coord.split(":")[1])
    c_coords = (x_c_coord.split(":")[0], x_c_coord.split(":")[1], y_c_coord.split(":")[0], y_c_coord.split(":")[1])
    a_coords = (x_a_coord.split(":")[0], x_a_coord.split(":")[1], y_a_coord.split(":")[0], y_a_coord.split(":")[1])

    q_coords = (int(q_coords[0]), int(q_coords[1]), int(q_coords[2]), int(q_coords[3]))
    c_coords = (int(c_coords[0]), int(c_coords[1]), int(c_coords[2]), int(c_coords[3]))
    a_coords = (int(a_coords[0]), int(a_coords[1]), int(a_coords[2]), int(a_coords[3]))

    settings_file.close()
def ImgToText(screen):
    return pytesseract.image_to_string(
                cv2.cvtColor(screen, cv2.COLOR_BGR2GRAY), 
                lang ='eng')
def SaveImage(name, screen):
    try:
        screen.save(rf"{DIR}\{name.strip()}.png")
    except:
        screen.save(rf"{DIR}\Manual Read Required.png")
def on_press(key):
    if key == Key.f2:
        # Screenshot
        screen,debug = GetBookworkCodeImage()
        SaveImage("!LastBookworkCheck", debug)
        answer_screen = GetAnswerImage()
        code = ImgToText(screen).strip()[-3:]
        SaveImage(code, answer_screen)
        print("Saved","--"+code+"--")
    elif key == Key.f4:
        # Get image
        
        try:
            screen,debug = GetBookworkCheckImage()
            ImgToText(screen)
            code = ImgToText(screen).strip()[-3:]
            img = Image.open(rf"{DIR}\{code}.png")
            ImageShow.show(img, title="Bookwork Check Answer")
            print(f"Opened image {code}")
        except:
            print("Couldnt open image.")
    elif key == Key.f8:
        dictionary = calibrating["q"]
        first = dictionary["i"] == (0,0)
        if first:
            pos = queryMousePosition()
            dictionary["i"] = pos
        else:
            pos = queryMousePosition()
            dictionary["j"] = pos

            # set the file pos
            settings_file = open("settings.txt", "r")
            data = settings_file.readlines()
            settings_file.close()
            new_1st_line = str(dictionary["i"]["x"])+":"+str(dictionary["i"]["y"])+"\n"
            new_2nd_line = str(dictionary["j"]["x"])+":"+str(dictionary["j"]["y"])+"\n"
            data[1] = new_1st_line
            data[2] = new_2nd_line

            settings_file = open("settings.txt", "w")
            settings_file.write("".join(data))
            settings_file.close()

            Update(data)
            
            dictionary["i"] = (0, 0)
            dictionary["j"] = (0, 0)
        print("Calibration","done" if not first else "almost done, press F8 on the opposite corner to complete")
    elif key == Key.f9:
        dictionary = calibrating["c"]
        first = dictionary["i"] == (0,0)
        if first:
            pos = queryMousePosition()
            dictionary["i"] = pos
        else:
            pos = queryMousePosition()
            dictionary["j"] = pos

            # set the file pos
            settings_file = open("settings.txt", "r")
            data = settings_file.readlines()
            settings_file.close()
            new_1st_line = str(dictionary["i"]["x"])+":"+str(dictionary["i"]["y"])+"\n"
            new_2nd_line = str(dictionary["j"]["x"])+":"+str(dictionary["j"]["y"])+"\n"
            data[4] = new_1st_line
            data[5] = new_2nd_line

            settings_file = open("settings.txt", "w")
            settings_file.write("".join(data))
            settings_file.close()

            Update(data)
            
            dictionary["i"] = (0, 0)
            dictionary["j"] = (0, 0)
        print("Calibration","done" if not first else "almost done, press F9 on the opposite corner to complete")
    elif key == Key.f10:
        dictionary = calibrating["a"]
        first = dictionary["i"] == (0,0)
        if first:
            pos = queryMousePosition()
            dictionary["i"] = pos
        else:
            pos = queryMousePosition()
            dictionary["j"] = pos

            # set the file pos
            settings_file = open("settings.txt", "r")
            data = settings_file.readlines()
            settings_file.close()
            new_1st_line = str(dictionary["i"]["x"])+":"+str(dictionary["i"]["y"])+"\n"
            new_2nd_line = str(dictionary["j"]["x"])+":"+str(dictionary["j"]["y"])+"\n"
            data[7] = new_1st_line
            data[8] = new_2nd_line

            

            settings_file = open("settings.txt", "w")
            settings_file.write("".join(data))
            settings_file.close()

            Update(data)
            
            dictionary["i"] = (0, 0)
            dictionary["j"] = (0, 0)
        print("Calibration","done" if not first else "almost done, press F10 on the opposite corner to complete")
    elif key == Key.end:
        data = """Questions:
798:177
1001:220
Check:
858:392
1060:445
Answer:
358:170
1564:1027"""
        settings_file = open("settings.txt", "w")
        settings_file.write(data)
        print("Successfully overwritten old data.")
def on_release(key):
    pass
# Collect events until released
with Listener(
        on_press=on_press,
        on_release=on_release) as listener:
    listener.join()

