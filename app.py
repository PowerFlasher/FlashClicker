import win32con
import win32api
import win32gui
import win32process
import time
import threading
import psutil
from tkinter import *
from tkinter import ttk
from tkinter import font

class MainActivity(Frame):

    def __init__(self, *args, **kwargs):
        Frame.__init__(self, *args, **kwargs)
        # handle
        self.handle = 264094
        top_frame = Frame()
        top_frame.pack(side=TOP)
        center_frame = Frame(width=150, height=150, background="#e6ffe6")
        center_frame.pack(side=TOP)
        bottom_frame = Frame()
        bottom_frame.pack(side=BOTTOM)
        # text="Open"
        open_icon = PhotoImage(file="icons/folder-open.png")
        button_open = Button(top_frame, image=open_icon)
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
        self.play_icon = PhotoImage(file="icons/play.png")
        self.stop_icon = PhotoImage(file="icons/stop.png")
        self.button_toggle = Button(top_frame, image=self.play_icon, relief="raised", command=self.toggle_button)
        self.button_toggle.play_icon = self.play_icon
        self.button_toggle.stop_icon = self.play_icon
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

    def target_program(self):
        target_window = Toplevel(self)
        target_window.tk.call('wm', 'iconphoto', target_window._w, self.target_icon)
        target_window.wm_title("Target Program")
        # listbox = MultiColumnListbox(target_window)
        tree = ttk.Treeview(target_window)

        tree["columns"] = ("one", "two")
        tree.column("one", width=100)
        tree.column("two", width=100)
        tree.heading("one", text="coulmn A")
        tree.heading("two", text="column B")

        tree.insert("", 0, text="Line 1", values=("1A", "1b"))

        id2 = tree.insert("", 1, "dir2", text="Dir 2")
        tree.insert(id2, "end", "dir 2", text="sub dir 2", values=("2A", "2B"))

        ##alternatively:
        tree.insert("", 3, "dir3", text="Dir 3")
        tree.insert("dir3", 3, text=" sub dir 3", values=("3A", " 3B"))

        tree.pack()

    def show_program(self):
        # win32gui.ShowWindow(self.handle, 5)
        print('alala')

    def key_press(self, key, time_press):
        win32api.PostMessage(self.handle, win32con.WM_KEYDOWN, key, 0)
        time.sleep(time_press)
        win32api.PostMessage(self.handle, win32con.WM_KEYUP, key, 0)

    def toggle_button(self):
        if self.button_toggle.config('relief')[-1] == 'sunken':
            self.button_toggle.config(relief="raised", image=self.play_icon)
            self.key_press(win32con.VK_F5, 1.0)
        else:
            self.button_toggle.config(relief="sunken", image=self.stop_icon)

class MultiColumnListbox(object):
    """use a ttk.TreeView as a multicolumn ListBox"""
    def __init__(self, master):
        self.tree = None
        self.master = master
        self.header = ['Name', 'PID']
        self._setup_widgets()
        self._build_tree()

    def _setup_widgets(self):
        s = """\click on header to sort by that column to change width of column drag boundary"""
        msg = ttk.Label(wraplength="4i", justify="left", anchor="n", padding=(10, 2, 10, 6), text=s)
        msg.pack(fill='x')
        container = ttk.Frame()
        container.pack(fill='both', expand=True)
        # create a treeview with dual scrollbars
        self.tree = ttk.Treeview(self.master, columns=self.header, show="headings")
        vsb = ttk.Scrollbar(orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(column=0, row=0, sticky='nsew', in_=container)
        vsb.grid(column=1, row=0, sticky='ns', in_=container)
        hsb.grid(column=0, row=1, sticky='ew', in_=container)
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)

    def _build_tree(self):
        for col in self.header:
            self.tree.heading(col, text=col.title(), command=lambda c=col: self.sortby(self.tree, c, 0))
            # adjust the column's width to the header string
            self.tree.column(col, width=font.Font().measure(col.title()))

        for proc in psutil.process_iter():
            try:
                pinfo = proc.as_dict(attrs=['pid', 'name'])
            except psutil.NoSuchProcess:
                pass
            else:
                self.tree.insert('', 'end', values=(pinfo['name'], pinfo['pid']))

        # for item in car_list:
        #     self.tree.insert('', 'end', values=item)
        #     # adjust column's width if necessary to fit each value
        #     for ix, val in enumerate(item):
        #         col_w = font.Font().measure(val)
        #         if self.tree.column(self.header[ix],width=None)<col_w:
        #             self.tree.column(self.header[ix], width=col_w)

    def sortby(self, tree, col, descending):
        """sort tree contents when a column header is clicked on"""
        # grab values to sort
        data = [(tree.set(child, col), child) \
                for child in tree.get_children('')]
        # if the data to be sorted is numeric change to float
        # data =  change_numeric(data)
        # now sort the data in place
        data.sort(reverse=descending)
        for ix, item in enumerate(data):
            tree.move(item[1], '', ix)
        # switch the heading so it will sort in the opposite direction
        tree.heading(col, command=lambda col=col: self.sortby(tree, col, int(not descending)))



root = Tk()
root.resizable(width=False, height=False)
root.geometry('{}x{}'.format(250, 200))
icon = PhotoImage(file="icons/flash.png")
root.tk.call('wm', 'iconphoto', root._w, icon)
root.title("FlashKeyMacro")
main = MainActivity(root)
root.mainloop()