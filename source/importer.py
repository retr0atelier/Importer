import time
import tkinter as tk
import win32gui
import re
import pyautogui
import win32api

#gui#

root = tk.Tk()
root.title('Fontforge Import')
root.iconbitmap('icon/retr0.ico')
root.geometry("300x150")

canvas = tk.Canvas(root,height=400,width=300)
canvas.pack()

doc_name = tk.Entry(root,font=('helvetica',10))
nr_glyphs = tk.Entry(root,font=('helvetica',10))

label2 = tk.Label(root, text='Document name:')
label2.config(font=('helvetica', 10))
canvas.create_window(150, 20, window=label2)

doc_name_box = canvas.create_window(150,45, window=doc_name)

label3 = tk.Label(root, text='Number of glyphs:')
label3.config(font=('helvetica', 10))
canvas.create_window(150, 70, window=label3)

nr_glyphs_box = canvas.create_window(150,96,window=nr_glyphs)


#select window#

def foreground():
    document_name = doc_name.get()

    class WindowMgr:
        def __init__ (self):
            self._handle = None
            
        def find_window(self, class_name, window_name=None):
            self._handle = win32gui.FindWindow(class_name, window_name)
            
        def _window_enum_callback(self, hwnd, wildcard):
            if re.match(wildcard, str(win32gui.GetWindowText(hwnd))) is not None:
                self._handle = hwnd
                
        def find_window_wildcard(self, wildcard):
            self._handle = None
            win32gui.EnumWindows(self._window_enum_callback, wildcard)
            
        def set_foreground(self):
            win32gui.SetForegroundWindow(self._handle)
        
    w = WindowMgr()
    w.find_window_wildcard(document_name)
    w.set_foreground()


#import glyphs#

def importglyphs():
    glyph = 1
    int_nr_glyphs = int(nr_glyphs.get())
    state_left = win32api.GetAsyncKeyState(0x01)
    while glyph <= int_nr_glyphs and state_left <=0:
        state_left = win32api.GetAsyncKeyState(0x01)
        if state_left < 0:
            break
        else:
            glyph_name = '_'+str(glyph)+'.'
            state_left = win32api.GetAsyncKeyState(0x01)
        if state_left < 0:
            break
        else:
            pyautogui.keyDown('ctrl')
            state_left = win32api.GetAsyncKeyState(0x01)
        if state_left < 0:
            break
        else:
            pyautogui.keyDown('shift')
            state_left = win32api.GetAsyncKeyState(0x01)
        if state_left < 0:
            break
        else:
            pyautogui.keyDown('i')
            state_left = win32api.GetAsyncKeyState(0x01)
        if state_left < 0:
            break
        else:
            pyautogui.keyUp('ctrl')
            state_left = win32api.GetAsyncKeyState(0x01)
        if state_left < 0:
            break
        else:
            pyautogui.keyUp('shift')
            state_left = win32api.GetAsyncKeyState(0x01)
        if state_left < 0:
            break
        else:
            pyautogui.keyUp('i')
            state_left = win32api.GetAsyncKeyState(0x01)
        if state_left < 0:
            break
        else:
            pyautogui.typewrite(glyph_name)
            state_left = win32api.GetAsyncKeyState(0x01)
        if state_left < 0:
            break
        else:
            pyautogui.press('enter')
            state_left = win32api.GetAsyncKeyState(0x01)
        if state_left < 0:
            break
        else:
            pyautogui.press('right')
            glyph += 1

#submit#

def execute():
    foreground()
    importglyphs()

#gui#

button1=tk.Button(text='Submit',command=execute,font=('helvetica',10))
canvas.create_window(150,130,window=button1)

root.mainloop()