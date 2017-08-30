import win32con
import win32api
import win32gui
import win32process
import time
import threading
import os
from tkinter import *
from tkinter import ttk
from tkinter import font
from tkinter import filedialog
from window_manager import WindowManager
from script import Script

class MainActivity(Frame):

    def __init__(self, *args, **kwargs):
        Frame.__init__(self, *args, **kwargs)
        # 
        self.initialdir = str(os.getcwd())
        self.__window_manager = WindowManager()
        self.__script = Script()
        self.sevenFont = font.Font(family="Helvetica", size=7)
        self.eightFont = font.Font(family="Helvetica", size=8)
        self.tenFont = font.Font(family="Helvetica", size=10)
        s = ttk.Style()
        s.configure('Green.TLabelframe.Label', font=('courier', 8, 'bold'))
        s.configure('Green.TLabelframe.Label', foreground ='green')
        top_frame = Frame()
        top_frame.pack(side=TOP)
        center_frame = Frame(width=150, height=166, background="#e6ffe6")
        # Target Program Label
        target_labelsFrame = ttk.Labelframe(center_frame, text='Target Program', labelanchor=N+S, width=150, height=50, style="Green.TLabelframe")
        ttk.Label(target_labelsFrame, text="Name:", font=self.sevenFont).grid(column=0, row=0, sticky='W')
        self.target_name = StringVar()
        ttk.Label(target_labelsFrame, textvariable=self.target_name, font=self.sevenFont).grid(column=1, row=0, sticky='W')
        ttk.Label(target_labelsFrame, text="PID:", font=self.sevenFont).grid(column=0, row=1, sticky='W')
        self.target_pid = StringVar()
        ttk.Label(target_labelsFrame, textvariable=self.target_pid, font=self.sevenFont).grid(column=1, row=1, sticky='W')
        target_labelsFrame.pack(fill="both")
        # End Target Program Label
        # Script data Label
        data_labelsFrame = ttk.Labelframe(center_frame, text='Script Data', labelanchor=N+S, width=150, height=50, style="Green.TLabelframe")
        ttk.Label(data_labelsFrame, text="Title:", font=self.sevenFont).grid(column=0, row=0, sticky='W')
        self.data_title = StringVar()
        ttk.Label(data_labelsFrame, textvariable=self.data_title, font=self.sevenFont).grid(column=1, row=0, sticky='W')
        ttk.Label(data_labelsFrame, text="Shortcuts:", font=self.sevenFont).grid(column=0, row=1, sticky='W')
        self.data_shortcuts = StringVar()
        ttk.Label(data_labelsFrame, textvariable=self.data_shortcuts, font=self.sevenFont).grid(column=1, row=1, sticky='W')
        data_labelsFrame.pack(fill="both")
        # End Script data Label
        # Info data Label
        status_labelsFrame = ttk.Labelframe(center_frame, text='Status Info', labelanchor=N+S, width=150, height=50, style="Green.TLabelframe")
        ttk.Label(status_labelsFrame, text="Macro:", font=self.sevenFont).grid(column=0, row=0, sticky='W')
        self.macro_title = StringVar()
        ttk.Label(status_labelsFrame, textvariable=self.macro_title, font=self.sevenFont).grid(column=1, row=0, sticky='W')
        ttk.Label(status_labelsFrame, text="Key:", font=self.sevenFont).grid(column=0, row=1, sticky='W')
        self.macro_key = StringVar()
        ttk.Label(status_labelsFrame, textvariable=self.macro_key, font=self.sevenFont).grid(column=1, row=1, sticky='W')
        ttk.Label(status_labelsFrame, text="Status:", font=self.sevenFont).grid(column=0, row=2, sticky='W')
        self.status_text = StringVar()
        ttk.Label(status_labelsFrame, textvariable=self.status_text, font=self.sevenFont).grid(column=1, row=2, sticky='W')
        status_labelsFrame.pack(fill="both")
        # End Info data Label
        center_frame.pack_propagate(False)
        center_frame.pack(side=TOP)
        bottom_frame = Frame()
        bottom_frame.pack(side=BOTTOM)
        # text="Open"
        open_icon = PhotoImage(file="icons/folder-open.png")
        button_open = Button(top_frame, image=open_icon, command=self.open_file)
        button_open.icon = open_icon
        # text="Target Program"
        self.target_icon = PhotoImage(file="icons/target.png")
        button_target_program = Button(top_frame, image=self.target_icon, command=self.target_program)
        button_target_program.icon = self.target_icon
        # text="Show Program"
        icon = PhotoImage(file="icons/window-restore.png")
        button_show_program = Button(top_frame, image=icon, command=self.show_program)
        button_show_program.icon = icon
        # text="Start/Stop"
        self.numeric_icon = PhotoImage(file="icons/numeric.png")
        self.infinity_icon = PhotoImage(file="icons/infinity.png")
        self.button_toggle = Button(top_frame, image=self.numeric_icon, relief="raised", command=self.toggle_button)
        self.button_toggle.numeric_icon = self.numeric_icon
        self.button_toggle.infinity_icon = self.infinity_icon
        # text="Info"
        icon = PhotoImage(file="icons/information-outline.png")
        button_info = Button(top_frame, image=icon)
        button_info.icon = icon
        # ProgressBar
        progress_bar = ttk.Progressbar(bottom_frame, orient="horizontal", length=150, mode="determinate")
        # GridLayout
        button_open.grid(row=0, column=0)
        button_target_program.grid(row=0, column=1)
        button_show_program.grid(row=0, column=2)
        self.button_toggle.grid(row=0, column=3)
        button_info.grid(row=0, column=4)
        progress_bar.pack()
        self.__script.progress_bar = progress_bar
        self.__script.set_status_info(self.macro_title, self.macro_key, self.status_text)


    def open_file(self):
        self.__script.load_json(filedialog.askopenfilename(initialdir=self.initialdir, title="Select macros file", filetypes=[("JSON files","*.json")]))
        self.data_title.set(self.__script.title)
        self.data_shortcuts.set(', '.join(str(x[0]).replace("VK_", "") for x in self.__script.shortcuts))


    def target_program(self):
        target_window = Toplevel(self)
        target_window.tk.call('wm', 'iconphoto', target_window._w, self.target_icon)
        target_window.wm_title("Target Program")
        # listbox = MultiColumnListbox(target_window)
        windows = self.__window_manager.getWindows()
        tree = ttk.Treeview(target_window, columns=("pid"))
        vsb = ttk.Scrollbar(target_window, orient="vertical", command=tree.yview)
        vsb.pack(side='right', fill='y')
        tree.configure(yscrollcommand=vsb.set)
        # tree["columns"] = ("pid")
        # tree.column("two", width=100)
        tree.column("#0", width=200)
        tree.heading("#0", text="Name", anchor=W)
        tree.column("pid", width=100)
        tree.heading("pid", text="PID", anchor=W)
        # tree.heading("two", text="column B")
        for window in windows:
            tree.insert("", 'end', str(window[0]), text=window[1], values=(window[0]))
            child = self.__window_manager.getChildWindow(window[0])
            while child and child != 0:
                tree.insert(str(window[0]), 'end', text=str(window[1]), values=(child))
                child = self.__window_manager.getChildWindow(child)

        tree.pack(expand=True)
        self.target_icon = PhotoImage(file="icons/target.png")
        button_target_program = Button(target_window, compound=LEFT, text="Target Program", image=self.target_icon, command= lambda: self.set_handle(tree.item(tree.selection()), target_window))
        button_target_program.icon = self.target_icon
        button_target_program.pack()

    def set_handle(self, selection, window):
        if selection['values'] != '':
            self.__window_manager.set_handle(int(selection['values'][0]))
            self.__script.set_handle(int(selection['values'][0]))
            self.target_name.set((selection['text'][:23] + '..') if len(selection['text']) > 23 else selection['text'])
            self.target_pid.set(selection['values'][0])
            window.destroy()


    def show_program(self):
        self.__window_manager.set_foreground()


    def toggle_button(self):
        if self.__script.infinity:
            self.button_toggle.config(image=self.numeric_icon)
        else:
            self.button_toggle.config(image=self.infinity_icon)
        self.__script.infinity = not self.__script.infinity
        # if self.button_toggle.config('relief')[-1] == 'sunken':
        #     self.button_toggle.config(relief="raised", image=self.numeric_icon)
        #     self.__script.stop_macro()
        # else:
        #     self.button_toggle.config(relief="sunken", image=self.infinity_icon)
        #     # self.key_press(win32con.VK_F5, 1.0)
        #     handle = int(self.__window_manager.get_handle())
        #     self.__script.press_shortcut(0)


if __name__ == '__main__':
    root = Tk()
    root.resizable(width=False, height=False)
    root.geometry('{}x{}'.format(250, 220))
    icon = PhotoImage(file="icons/flash.png")
    root.tk.call('wm', 'iconphoto', root._w, icon)
    root.title("FlashKeys")
    main = MainActivity(root)
    root.mainloop()