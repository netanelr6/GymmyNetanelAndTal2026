# -*- coding: utf-8 -*-
import time
import tkinter as tk
from PIL import Image, ImageTk
import Settings as s
import random


class Screen(tk.Tk):
    def __init__(self):
        print("screen start")
        tk.Tk.__init__(self, className='Poppy')
        self._frame = None

        # self.switch_frame(SelectPage)
        self.switch_frame(EyesPage) #--------------->maya original

        self["bg"] = "#F3FCFB"



        #--------------------------------n&t CODE--------------------------------------------------
        # גורם לחלון לקפוץ קדימה מעל ה-VS Code
        self.lift()
        self.attributes("-topmost", True)
        self.focus_force()
        self.bind("<Escape>", lambda e: self.attributes("-fullscreen", False))

        self.bind_all("<Button-3>", self.reboot_from_mouse)   # קליק ימני
        #self.bind_all("<Button-1>", self.reboot_from_mouse)
        #--------------------------------------------------------------------------------------------


    #------------------------------------n&t CODE-------------------    
    def reboot_from_mouse(self, event=None):
        print("RIGHT CLICK REBOOT TRIGGERED")
        # מבטיח פוקוס (ליתר ביטחון)
        self.focus_force()
        s.reboot_flag = True
    #    -------------------------------------------------------------




    def switch_frame(self, frame_class):
        """Destroys current frame and replaces it with a new one."""
        new_frame = frame_class(self)
        if self._frame is not None:
            if hasattr(self._frame, 'background_label'):
                self._frame.background_label.destroy()
            self._frame.destroy()
        self._frame = new_frame
        self._frame.pack()


class EyesPage(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        image = Image.open('pictures//eyes.png')
        self.photo_image = ImageTk.PhotoImage(image)  # self. - for keeping the photo in memory so it will be shown
        tk.Label(self, image=self.photo_image).pack()


class FullScreenApp(object):
    def __init__(self, master, **kwargs):
        self.master = master
        pad = 3
        self._geom = '200x200+0+0'
        master.geometry("{0}x{1}+0+0".format(
            master.winfo_screenwidth()-pad, master.winfo_screenheight()-pad))
        master.bind('<Escape>', self.toggle_geom)

    def toggle_geom(self, event):
        geom=self.master.winfo_geometry()
        print(geom, self._geom)
        self.master.geometry(self._geom)
        self._geom = geom


if __name__ == "__main__":
    s.screen = Screen()
    app = FullScreenApp(s.screen)
    s.screen.mainloop()
