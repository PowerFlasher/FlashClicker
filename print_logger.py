import tkinter as tk

class PrintLogger(): # create file like object
    def __init__(self, textbox): # pass reference to text widget
        self.textbox = textbox # keep ref

    def write(self, text):
        self.textbox.configure(state='normal')
        self.textbox.insert(tk.END, text) # write text to textbox
        self.textbox.configure(state='disabled')
            # could also scroll to end of textbox here to make sure always visible

    def flush(self): # needed for file like object
        pass