import win32con
import win32api
import win32gui
import win32process
import time
import threading
import os
import sys
from tkinter import *
from tkinter import ttk
from tkinter import font
from tkinter import filedialog
from window_manager import WindowManager
from script import Script
from hotkey import values, dik, mouse_buttons_values
from JsonEditor import JsonEditor
from print_logger import PrintLogger
# from macro_recorder import record as record_macro

class MainActivity(Frame):

    def __init__(self, *args, **kwargs):
        Frame.__init__(self, *args, **kwargs)
        # 
        self.initialdir = str(os.path.join(os.getcwd(), 'scripts'))
        self.__window_manager = WindowManager()
        self.__script = Script()
        self.__script_path = ""
        self.labelFont = font.Font(family="Calibri", size=8)
        s = ttk.Style()
        s.configure('Green.TLabelframe.Label', font=('Calibri', 10, 'bold'))
        s.configure('Green.TLabelframe.Label', foreground ='green')
        left_panel = Frame()
        left_panel.pack(side=LEFT)
        center_frame = Frame(width=400, height=240)
        # Target Program Label
        target_labelsFrame = ttk.Labelframe(center_frame, text='Target Program', labelanchor=N+S, width=350, height=60, style="Green.TLabelframe")
        ttk.Label(target_labelsFrame, text="Name:", font=self.labelFont).grid(column=0, row=0, sticky='W')
        self.target_name = StringVar()
        ttk.Label(target_labelsFrame, textvariable=self.target_name, font=self.labelFont).grid(column=1, row=0, sticky='W')
        ttk.Label(target_labelsFrame, text="PID:", font=self.labelFont).grid(column=0, row=1, sticky='W')
        self.target_pid = StringVar()
        ttk.Label(target_labelsFrame, textvariable=self.target_pid, font=self.labelFont).grid(column=1, row=1, sticky='W')
        target_labelsFrame.pack(fill="both", padx=10, pady=5)
        # End Target Program Label
        # Script data Label
        data_labelsFrame = ttk.Labelframe(center_frame, text='Script Data', labelanchor=N+S, width=350, height=60, style="Green.TLabelframe")
        ttk.Label(data_labelsFrame, text="Title:", font=self.labelFont).grid(column=0, row=0, sticky='W')
        self.data_title = StringVar()
        ttk.Label(data_labelsFrame, textvariable=self.data_title, font=self.labelFont).grid(column=1, row=0, columnspan=3, sticky='W')
        ttk.Label(data_labelsFrame, text="Shortcuts:", font=self.labelFont).grid(column=0, row=1, sticky='W')
        self.data_shortcuts = StringVar()
        ttk.Label(data_labelsFrame, textvariable=self.data_shortcuts, font=self.labelFont).grid(column=1, row=1, columnspan=3, sticky='W')
        # text="Edit file"
        self.button_edit_file = Button(data_labelsFrame, text="Edit", command=self.edit_file)
        self.button_edit_file['font'] = font.Font(family="Calibri", size=8)
        self.button_edit_file.config(state="disabled")
        self.button_edit_file.grid(row=3, column=0, sticky="ew")
        # text="Total Time"
        self.button_total_time = Button(data_labelsFrame, text="Total Time", command=self.total_time)
        self.button_total_time['font'] = font.Font(family="Calibri", size=8)
        self.button_total_time.config(state="disabled")
        self.button_total_time.grid(row=3, column=1, sticky="ew")
        data_labelsFrame.pack(fill="both", padx=10)
        # End Script data Label
        # Info data Label
        status_labelsFrame = ttk.Labelframe(center_frame, text='Status Info', labelanchor=N+S, width=350, height=75, style="Green.TLabelframe")
        ttk.Label(status_labelsFrame, text="Macro:", font=self.labelFont).grid(column=0, row=0, sticky='W')
        self.macro_title = StringVar()
        ttk.Label(status_labelsFrame, textvariable=self.macro_title, font=self.labelFont).grid(column=1, row=0, sticky='W')
        ttk.Label(status_labelsFrame, text="Key:", font=self.labelFont).grid(column=0, row=1, sticky='W')
        self.macro_key = StringVar()
        ttk.Label(status_labelsFrame, textvariable=self.macro_key, font=self.labelFont).grid(column=1, row=1, sticky='W')
        ttk.Label(status_labelsFrame, text="Status:", font=self.labelFont).grid(column=0, row=2, sticky='W')
        self.status_text = StringVar()
        ttk.Label(status_labelsFrame, textvariable=self.status_text, font=self.labelFont).grid(column=1, row=2, sticky='W')
        status_labelsFrame.pack(fill="both", padx=10)
        # End Info data Label
        center_frame.pack_propagate(False)
        center_frame.pack(side=TOP)
        bottom_frame = Frame()
        bottom_frame.pack(side=BOTTOM)
        self.edit_icon = PhotoImage(file="icons/edit.png")
        # text="Import File"
        open_icon = PhotoImage(file="icons/import-file.png")
        self.button_import_file = Button(left_panel, compound=LEFT, text="Import File", anchor="w", image=open_icon, command=self.open_file)
        self.button_import_file.icon = open_icon
        # text="Refresh File"
        self.reload_icon = PhotoImage(file="icons/refresh.png")
        self.button_reload = Button(left_panel, compound=LEFT, text=" ", anchor="w", image=self.reload_icon, command=self.reload)
        self.button_reload.config(state="disabled")
        self.button_reload.icon = self.reload_icon
        # text="Target Program"
        self.target_icon = PhotoImage(file="icons/target.png")
        self.button_target_program = Button(left_panel, compound=LEFT, text="Target Program", anchor="w", image=self.target_icon, command=self.target_program)
        self.button_target_program.icon = self.target_icon
        # text="Show Program"
        icon = PhotoImage(file="icons/window-search.png")
        self.button_show_program = Button(left_panel, compound=LEFT, text="Show Window", anchor="w", image=icon, command=self.show_program)
        self.button_show_program.icon = icon
        # text="Finite/Infinite"
        self.numeric_icon = PhotoImage(file="icons/countdown-clock.png")
        self.infinity_icon = PhotoImage(file="icons/infinite.png")
        self.button_toggle = Button(left_panel, compound=LEFT, text="Finite", anchor="w", image=self.numeric_icon, relief="raised", command=self.toggle_button)
        self.button_toggle.numeric_icon = self.numeric_icon
        self.button_toggle.infinity_icon = self.infinity_icon
        # text="Help"
        self.help_icon = PhotoImage(file="icons/help.png")
        self.button_help = Button(left_panel, compound=LEFT, text="Help", anchor="w", image=self.help_icon, command=self.help)
        self.button_help.icon = self.help_icon
        # text="Force Stop"
        self.force_stop_icon = PhotoImage(file="icons/force-stop.png")
        self.button_force_stop = Button(left_panel, compound=LEFT, text="Force Stop", anchor="w", image=self.force_stop_icon, command=self.stop_macro)
        self.button_force_stop.icon = self.force_stop_icon
        # text="Record Macro"
        self.record_macro_icon = PhotoImage(file="icons/record-macro.png")
        self.button_record_macro = Button(left_panel, compound=LEFT, text="Record Macro", anchor="w", image=self.record_macro_icon, command=self.record_macro)
        self.button_record_macro.icon = self.record_macro_icon
        # text="Settings"
        self.console_icon = PhotoImage(file="icons/console.png")
        self.button_console = Button(left_panel, compound=LEFT, text="Console Logs", anchor="w", image=self.console_icon, command=self.open_console)
        self.button_console.icon = self.console_icon
        # ProgressBar
        progress_bar = ttk.Progressbar(bottom_frame,length=400, orient="horizontal", mode="determinate")
        # GridLayout
        self.button_import_file.grid(row=0, column=0, sticky="ew")
        self.button_reload.grid(row=0, column=1, sticky="ew")
        self.button_target_program.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.button_show_program.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.button_toggle.grid(row=3, column=0, columnspan=2, sticky="ew")
        self.button_force_stop.grid(row=4, column=0, columnspan=2, sticky="ew")
        self.button_record_macro.grid(row=5, column=0, columnspan=2, sticky="ew")
        self.button_help.grid(row=6, column=0, columnspan=2, sticky="ew")
        self.button_console.grid(row=7, column=0, columnspan=2, sticky="ew")
        progress_bar.pack()
        self.__script.progress_bar = progress_bar
        self.__script.set_status_info(self.macro_title, self.macro_key, self.status_text)


    def open_file(self):
        self.__script_path = filedialog.askopenfilename(initialdir=self.initialdir, title="Select Script File", filetypes=[("JSON files","*.json")])
        if self.__script_path:
            self.__script.load_json(self.__script_path)
            self.data_title.set(self.__script.title)
            self.data_shortcuts.set(', '.join(str(x[0]).replace("VK_", "") for x in self.__script.shortcuts))
            self.button_reload.config(state="normal")
            self.button_edit_file.config(state="normal")
            self.button_total_time.config(state="normal")

    def reload(self):
        if self.__script_path:
            self.__script.load_json(self.__script_path)
            self.data_title.set(self.__script.title)
            self.data_shortcuts.set(', '.join(str(x[0]).replace("VK_", "") for x in self.__script.shortcuts))

    def close_and_reload(self, window):
        self.reload()
        window.destroy()

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

        tree.pack(fill=BOTH, expand=True)
        self.target_icon = PhotoImage(file="icons/target.png")
        button_target_program = Button(target_window, compound=LEFT, text="Target Program", image=self.target_icon, command= lambda: self.set_handle(tree.item(tree.selection()), target_window))
        button_target_program.icon = self.target_icon
        button_target_program.pack()

    def edit_file(self):
        if self.__script_path:
            edit_file_window = Toplevel(self)
            edit_file_window.geometry('700x500')
            edit_file_window.tk.call('wm', 'iconphoto', edit_file_window._w, self.edit_icon)
            edit_file_window.wm_title("Edit Script")
            editor = JsonEditor(edit_file_window, self.__script_path)
            edit_file_window.protocol("WM_DELETE_WINDOW", lambda: self.close_and_reload(edit_file_window))

    def total_time(self):
        statistics = self.__script.get_statistics()
        total_time_window = Toplevel(self)
        total_time_window.tk.call('wm', 'iconphoto', total_time_window._w, self.target_icon)
        total_time_window.wm_title("Total Time - Macros")
        # listbox = MultiColumnListbox(total_time_window)
        windows = self.__window_manager.getWindows()
        tree = ttk.Treeview(total_time_window, columns=("milliseconds", "seconds", "minutes", "hours", "days"))
        vsb = ttk.Scrollbar(total_time_window, orient="vertical", command=tree.yview)
        vsb.pack(side='right', fill='y')
        tree.configure(yscrollcommand=vsb.set)
        # tree["columns"] = ("pid")
        # tree.column("two", width=100)
        tree.column("#0", width=200)
        tree.heading("#0", text="Macro", anchor=W)
        tree.column("milliseconds", width=100)
        tree.heading("milliseconds", text="Milliseconds", anchor=W)
        tree.column("seconds", width=100)
        tree.heading("seconds", text="Seconds", anchor=W)
        tree.column("minutes", width=100)
        tree.heading("minutes", text="Minutes", anchor=W)
        tree.column("hours", width=100)
        tree.heading("hours", text="Hours", anchor=W)
        tree.column("days", width=100)
        tree.heading("days", text="Days", anchor=W)
        for macro in statistics:
            tree.insert("", 'end', text=macro[0], values=(macro[1],macro[2],macro[3],macro[4],macro[5]))
        tree.pack(expand=True)

    def help(self):
        help_window = Toplevel(self)
        help_window.tk.call('wm', 'iconphoto', help_window._w, self.help_icon)
        help_window.wm_title("Help Keys ID")
        tree = ttk.Treeview(help_window, columns=("value"))
        vsb = ttk.Scrollbar(help_window, orient="vertical", command=tree.yview)
        vsb.pack(side='right', fill='y')
        tree.configure(yscrollcommand=vsb.set)
        tree.column("#0", width=150)
        tree.heading("#0", text="Key", anchor=W)
        tree.column("value", width=450)
        tree.heading("value", text="Value", anchor=W)
        tree.insert("", 'end', "GCP", text="Get Cursor Position", values="Use\ the\ F10\ key\ to\ record\ the\ cursor\ coordinate\ (print\ is\ in\ the\ Console\ Logs)")
        tree.insert("", 'end', "IF", text="Import File", values="Import\ the\ correct\ json\ file\ with\ the\ macro")
        tree.insert("", 'end', "RB", text="Refresh Button", values="If\ you\ make\ a\ change\ to\ the\ file,\ just\ reload\ the\ file")
        tree.insert("", 'end', "TP", text="Target Program", values="Target\ the\ program\ to\ which\ the\ macro\ will\ be\ sent")
        tree.insert("", 'end', "SW", text="Show Window", values="Show\ the\ currently\ targeted\ window")
        tree.insert("", 'end', "FI", text="Finite/Infinite", values="Mode\ between\ final\ and\ infinite\ macro\ execution")
        tree.insert("", 'end', "FS", text="Force Stop", values="Stop\ the\ currently\ running\ macro")
        tree.insert("", 'end', "RM", text="Record Macro", values="Press\ the\ ESCAPE\ key\ (ESC)\ to\ stop\ recording")
        tree.insert("", 'end', "CL", text="Console Logs", values="Displays\ auxiliary\ printout")
        tree.insert("", 'end', "Keys", text="Keys for macro", values="")
        categories = [('Virtual Keys', values), ('Direct Keys', dik), ('Mouse X Buttons', mouse_buttons_values)];
        for category in categories:
            tree.insert("Keys", 'end', str(category[0]), text=category[0], values=(category[0]))
            for key in category[1]:
                tree.insert(str(category[0]), 'end', text=str(key), values=(str(category[1][key])))

        tree.pack(fill=BOTH, expand=True)

    def open_console(self):
        console_window = Toplevel(self)
        console_window.tk.call('wm', 'iconphoto', console_window._w, self.console_icon)
        console_window.wm_title("Console Logs")
        text_box = Text(console_window)
        text_box.config(state="disabled")
        text_box.pack()
        pl = PrintLogger(text_box)
        # replace sys.stdout with our object
        # sys.stdout = pl

    def stop_macro(self):
        self.__script.stop_macro()

    def record_macro(self):
        self.button_import_file.config(state="disabled")
        self.button_reload.config(state="disabled")
        self.button_target_program.config(state="disabled")
        self.button_show_program.config(state="disabled")
        self.button_toggle.config(state="disabled")
        self.button_force_stop.config(state="disabled")
        self.button_record_macro.config(state="disabled")
        self.button_help.config(state="disabled")
        self.button_console.config(state="disabled")
        self.__script.load_recorded_macro = self.load_recorded_macro
        self.__script.record_macro = True

    def set_handle(self, selection, window):
        if selection['values'] != '':
            self.__window_manager.set_handle(int(selection['values'][0]))
            self.__script.set_handle(int(selection['values'][0]))
            self.target_name.set((selection['text'][:35] + '..') if len(selection['text']) > 35 else selection['text'])
            self.target_pid.set(selection['values'][0])
            window.destroy()


    def show_program(self):
        self.__window_manager.set_foreground()


    def toggle_button(self):
        if self.__script.infinity:
            self.button_toggle.config(text="Finite", image=self.numeric_icon)
        else:
            self.button_toggle.config(text="Infinite", image=self.infinity_icon)
        self.__script.infinity = not self.__script.infinity
        # if self.button_toggle.config('relief')[-1] == 'sunken':
        #     self.button_toggle.config(relief="raised", image=self.numeric_icon)
        #     self.__script.stop_macro()
        # else:
        #     self.button_toggle.config(relief="sunken", image=self.infinity_icon)
        #     # self.key_press(win32con.VK_F5, 1.0)
        #     handle = int(self.__window_manager.get_handle())
        #     self.__script.press_shortcut(0)


    def load_recorded_macro(self, file_path):
        self.button_import_file.config(state="normal")
        self.button_reload.config(state="normal")
        self.button_target_program.config(state="normal")
        self.button_show_program.config(state="normal")
        self.button_toggle.config(state="normal")
        self.button_force_stop.config(state="normal")
        self.button_record_macro.config(state="normal")
        self.button_help.config(state="normal")
        self.button_console.config(state="normal")
        self.__script.record_macro = False
        if file_path:
            self.__script.load_json(file_path)
            self.__script_path = file_path
            self.data_title.set(self.__script.title)
            self.data_shortcuts.set(', '.join(str(x[0]).replace("VK_", "") for x in self.__script.shortcuts))
            self.button_reload.config(state="normal")
            self.button_edit_file.config(state="normal")
            self.button_total_time.config(state="normal")


if __name__ == '__main__':
    root = Tk()
    root.resizable(width=False, height=False)
    # root.geometry('{}x{}'.format(300, 240))
    icon = PhotoImage(file="icons/flash.png")
    root.tk.call('wm', 'iconphoto', root._w, icon)
    root.title("Flicker")
    main = MainActivity(root)
    root.mainloop()