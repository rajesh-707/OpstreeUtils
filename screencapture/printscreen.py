from pynput.keyboard import Key, Listener
from PIL import ImageGrab
from docx import Document
from docx.shared import Inches
import sys
from Xlib.display import Display

document = Document()
counter = 0
filename_prefix = 'demo'


def on_press(key):
    if key == Key.space:
       print('{0} pressed'.format(
        key))

def on_release(key):
#    counter = 0
    global counter
    if key == Key.space :
       print('{0} release'.format(
        key))
       snapshot = ImageGrab.grab()
       str_counter = str(counter)
       save_path = 'C:\\python\\' + filename_prefix + str_counter + '.jpg'
       snapshot.save(save_path)
       p = document.add_paragraph()
       r = p.add_run()
       r.add_picture(save_path, width=Inches(7.0))
       
       counter = counter + 1
    document.save('C:\\python\\' + sys.argv[1])
    if key == Key.esc:
        # Stop listener
        return False

# Collect events until released
with Listener(
        on_press=on_press,
        on_release=on_release) as listener:
    listener.join()
