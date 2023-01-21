#PinFav_GUI.py

#TODO:
# All the commented texts in the script ¯\_(ツ)_/¯

# Standard library imports
import os
import re
import time
import enum
import ctypes
import textwrap

# Third-party library imports
import win32con
import win32gui
import pywintypes
import win32process
import tkinter as tk
from tkinter import ttk
from pathlib import Path
from functools import partial
from psutil import Process, pid_exists
from pywinauto.handleprops import has_exstyle
from pywinauto.win32functions import win32defines


def show_message_box(title, message, style):
    return ctypes.windll.user32.MessageBoxW(0, message, title, style)

def get_handle_from_pid(pid:int):
    def callback(hwnd, pid):
        if pid == win32process.GetWindowThreadProcessId(hwnd)[1]: # [0], [1] = tid, pid
            hwnd_windows.append(hwnd)
    hwnd_windows = []
    win32gui.EnumWindows(callback, pid)
    for hwnd in hwnd_windows:
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd): # DwmGetWindowAttribute # win32gui.IsWindowEnabled(hwnd)
            return hwnd

def get_pid_from_handle(hwnd:int):
    return win32process.GetWindowThreadProcessId(hwnd)[1]

def get_window_title_from_handle(hwnd:int):
    return win32gui.GetWindowText(hwnd)

def get_name_from_pid(pid:int):
    return Process(pid).name()


class Msgbox(enum.IntFlag):
    # https://stackoverflow.com/questions/50086178/python-how-to-keep-messageboxw-on-top-of-all-other-windows
    # https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-messageboxw
    # https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/msgbox-function
    OKOnly = 0  # Display OK button only.
    OKCancel = 1  # Display OK and Cancel buttons.
    AbortRetryIgnore = 2  # Display Abort, Retry, and Ignore buttons.
    YesNoCancel = 3  # Display Yes, No, and Cancel buttons.
    YesNo = 4  # Display Yes and No buttons.
    RetryCancel = 5  # Display Retry and Cancel buttons.
    Critical = 16  # Display Critical Message icon.
    Question = 32  # Display Warning Query icon.
    Exclamation = 48  # Display Warning Message icon.
    Information = 64  # Display Information Message icon.
    DefaultButton1 = 0  # First button is default.
    DefaultButton2 = 256  # Second button is default.
    DefaultButton3 = 512  # Third button is default.
    DefaultButton4 = 768  # Fourth button is default.
    ApplicationModal = 0  # Application modal; the user must respond to the message box before continuing work in the current application.
    SystemModal = 4096  # System modal; all applications are suspended until the user responds to the message box.
    MsgBoxHelpButton = 16384  # Adds Help button to the message box.
    MsgBoxSetForeground = 65536  # Specifies the message box window as the foreground window.
    MsgBoxRight = 524288  # Text is right-aligned.
    MsgBoxRtlReading = 1048576  # Specifies text should appear as right-to-left reading on Hebrew and Arabic systems.

class ProcessManager(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(TITLE)
        self.geometry("550x400")
        self.minsize(width=292, height=273)

        self.t1 = None
        self.pinfav_handle = None
        self.update_pinfav_in_gui = None

        # for some reasons I had the need to pre-define variables here in order to get rid of stderr
        self.search_var = tk.StringVar()
        self.search_var.set("Enter a search term here")
        self.search_var.trace("w", partial(self.update_list, "user_is_using_searchbar"))
        self.button_refresh_timer = tk.StringVar()
        self.button_refresh_timer.set(f"Refresh (5)")
        self.search_frame = ttk.Frame(self, relief="flat", borderwidth=5) # flat, groove, raised, ridge, solid, or sunken
        self.clear_button = ttk.Button(self.search_frame, text="Clear", command=self.clear_search_bar)
        self.process_frame = ttk.Frame(self, relief="flat", borderwidth=5) # flat, groove, raised, ridge, solid, or sunken
        self.process_list = tk.Listbox(self.process_frame)
        self.process_list.bind("<<ListboxSelect>>", self.select_listbox_item)
        self.process_list.bind("<Double-Button-1>", partial(self.pin_or_unpin_process, "pin_process"))
        self.process_scrollbar = ttk.Scrollbar(self.process_frame, orient="vertical", command=self.process_list.yview)
        self.process_list.config(yscrollcommand=self.process_scrollbar.set)
        self.pinned_list = tk.Listbox(self.process_frame)
        self.pinned_list.bind("<<ListboxSelect>>", self.select_listbox_item)
        self.pinned_list.bind("<Double-Button-1>", partial(self.pin_or_unpin_process, "unpin_process"))
        self.pinned_scrollbar = ttk.Scrollbar(self.process_frame, orient="vertical", command=self.pinned_list.yview)
        self.pinned_list.config(yscrollcommand=self.pinned_scrollbar.set)
        self.search_entry = ttk.Entry(self.search_frame, textvariable=self.search_var, width=self.pinned_list.winfo_reqwidth() + self.process_list.winfo_reqwidth())
        self.search_entry.bind("<FocusIn>", self.on_focus_in)
        self.search_entry.bind("<FocusOut>", self.on_focus_out)
        self.empty_frame3 = ttk.Frame(self)
        self.refresh_button = ttk.Button(self, textvariable=self.button_refresh_timer, command=partial(self.update_list, "refresh_button_clicked"))
        self.pin_button = ttk.Button(self, text="Pin", command=partial(self.pin_or_unpin_process, "pin_process"))
        self.unpin_button = ttk.Button(self, text="Unpin", command=partial(self.pin_or_unpin_process, "unpin_process"))
        self.empty_frame4 = ttk.Frame(self)

        # for some reasons I had to pack everything lately in order to get rid of stderr
        self.search_frame.pack(side="top", fill="none", expand=False)
        self.clear_button.pack(side="right")
        self.process_frame.pack(side="top", fill="both", expand=True)
        #self.process_list_label = ttk.Label(self.process_frame, text="Process List")
        #self.process_list_label.pack(side="top", fill="x")
        #self.process_list_label.pack(side="top", padx=10, pady=10)
        self.process_list.pack(side="left", fill="both", expand=True)
        self.process_scrollbar.pack(side="left", fill="y")
        #self.pinned_list_label = ttk.Label(self.process_frame, text="Pinned Processes")
        #self.pinned_list_label.pack(side="top", fill="x")
        #self.pinned_list_label.pack(side="top", padx=10, pady=10)
        self.pinned_list.pack(side="right", fill="both", expand=True)
        self.pinned_scrollbar.pack(side="right", fill="y")
        self.search_entry.pack(side="left", fill="x", expand=True)
        self.empty_frame3.pack(side="top", fill="none", expand=False, pady=5)
        self.refresh_button.pack(side="top", fill="x")
        self.pin_button.pack(side="top", fill="x")
        self.unpin_button.pack(side="top", fill="x")
        self.empty_frame4.pack(side="top", fill="none", expand=False, pady=1)

        self.update_list()

    def on_focus_in(self, event):
        if self.search_var.get() == "Enter a search term here":
            self.search_entry.delete(0, "end")

    def on_focus_out(self, event):
        if self.search_var.get() in ["", "Enter a search term here"]:
            self.search_entry.delete(0, "end")
            self.search_entry.insert(0, "Enter a search term here")

    def clear_search_bar(self, event=None):
        if self.search_var.get() == "Enter a search term here":
            return
        self.search_var.set("Enter a search term here")
        self.update_list()

    def select_listbox_item(self, event=None):
        self.t1 = time.perf_counter()
        self.button_refresh_timer.set(f"Refresh (5)")

    def pin_or_unpin_process(self, pin_or_unpin, event=None):
        def get_selected_from(most_likely_selected_item_list:list):
            def get_it(list):
                try:
                    return (list, str(list.get(list.curselection())), list.curselection())
                except tk.TclError:
                    try:
                        return (list, str(list.get(list.curselection())), list.curselection())
                    except tk.TclError:
                        return None

            for list in most_likely_selected_item_list:
                selected = get_it(list)
                if selected is not None:
                    return selected

            return None

        if pin_or_unpin == "pin_process":
            selected = get_selected_from([self.process_list, self.pinned_list])
        elif pin_or_unpin == "unpin_process":
            selected = get_selected_from([self.pinned_list, self.process_list])
        if selected is None:
            return

        selected_list = selected[0]
        selected_item = selected[1]
        selected_index = selected[2]

        name = str(re.search(r"^(.*?\.(exe|con)) \(", selected_item).group(1))
        title = str(re.search(r" \((.*?)\) \(", selected_item).group(1))
        hwnd = int(re.search(r"\) \(([0-9]+)\) \(", selected_item).group(1))
        pid = int(re.search(r"\) \(([0-9]+)\)$", selected_item).group(1))

        if not pid_exists(pid):
            self.t1 = None
            return

        if pin_or_unpin == "pin_process":
            state = win32con.HWND_TOPMOST
        elif pin_or_unpin == "unpin_process":
            state = win32con.HWND_NOTOPMOST
        flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE

        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        try:
            win32gui.SetWindowPos(hwnd, state, 0,0,0,0, flags)
        except pywintypes.error as exception:
            msgbox_title = TITLE
            msgbox_text = f"""
                [ERROR (pywintypes.error)]:

                "{exception.args[2]}"

                {TITLE} could not 'pin' process:
                name: {name}
                title: {title}
                hwnd: {hwnd}
                pid: {pid}
            """
            msgbox_text = textwrap.dedent(msgbox_text).removeprefix("\n").removesuffix("\n")
            msgbox_style = Msgbox.OKOnly | Msgbox.Exclamation |Msgbox.SystemModal | Msgbox.MsgBoxSetForeground
            show_message_box(msgbox_title, msgbox_text, msgbox_style)
            return

        win32gui.SetForegroundWindow(hwnd)
        if self.pinfav_handle is not None:
            win32gui.SetForegroundWindow(self.pinfav_handle)

        if pin_or_unpin == "pin_process":
            if has_exstyle(hwnd, win32defines.WS_EX_TOPMOST):
                if selected_item not in self.pinned_list.get(0, tk.END):
                    self.pinned_list.insert(tk.END, selected_item)
                if selected_item in self.process_list.get(0, tk.END):
                    self.process_list.delete(selected_index)
            else:
                msgbox_title = TITLE
                msgbox_text = f"""
                    [ERROR ({TITLE})]:

                    {TITLE} could not 'pin' process
                    name: {name}
                    title: {title}
                    hwnd: {hwnd}
                    pid: {pid}
                """
                msgbox_text = textwrap.dedent(msgbox_text).removeprefix("\n").removesuffix("\n")
                msgbox_style = Msgbox.OKOnly | Msgbox.Exclamation |Msgbox.SystemModal | Msgbox.MsgBoxSetForeground
                show_message_box(msgbox_title, msgbox_text, msgbox_style)
                return
        elif pin_or_unpin == "unpin_process":
            if has_exstyle(hwnd, win32defines.WS_EX_TOPMOST):
                msgbox_title = TITLE
                msgbox_text = f"""
                    [ERROR ({TITLE})]:

                    {TITLE} could not 'unpin' process:
                    name: {name}
                    title: {title}
                    hwnd: {hwnd}
                    pid: {pid}
                """
                msgbox_text = textwrap.dedent(msgbox_text).removeprefix("\n").removesuffix("\n")
                msgbox_style = Msgbox.OKOnly | Msgbox.Exclamation |Msgbox.SystemModal | Msgbox.MsgBoxSetForeground
                show_message_box(msgbox_title, msgbox_text, msgbox_style)
                return
            else:
                if selected_item not in self.process_list.get(0, tk.END):
                    self.process_list.insert(tk.END, selected_item)
                if selected_item in self.pinned_list.get(0, tk.END):
                    self.pinned_list.delete(selected_index)

    #def highlight_button(self, pinned):
    #    if pinned:
    #        self.pin_button.config(bg="light blue")
    #    else:
    #        self.pin_button.config(bg="white")

    def update_list(self, *args):
        def get_handles():
            def callback(hwnd:int, handles:list):
                if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd): # or win32gui.IsIconic(hwnd)
                    handles.append(hwnd)
            handles = []
            win32gui.EnumWindows(callback, handles)
            return handles

        def gui_need_to_refresh():
            def get_seconds_elapsed_from_last_refresh():
                t2 = time.perf_counter()
                seconds_elapsed = round(t2 - self.t1)
                seconds_left = 5 - seconds_elapsed
                return seconds_left

            if self.t1 is None:
                self.t1 = time.perf_counter()
                return True

            if self.update_pinfav_in_gui is None:
                seconds_left = get_seconds_elapsed_from_last_refresh()
                self.button_refresh_timer.set(f"Refresh ({seconds_left})")
                return True
            elif self.update_pinfav_in_gui is True:
                seconds_left = get_seconds_elapsed_from_last_refresh()
                self.button_refresh_timer.set(f"Refresh ({seconds_left})")
                self.update_pinfav_in_gui = False
                return True

            if args and any(item in args for item in ["user_is_using_searchbar", "refresh_button_clicked"]):
                self.t1 = time.perf_counter()
                self.button_refresh_timer.set(f"Refresh (5)")
                return True

            seconds_left = get_seconds_elapsed_from_last_refresh()
            self.button_refresh_timer.set(f"Refresh ({seconds_left})")
            if seconds_left > 0:
                return False
            self.t1 = time.perf_counter()
            return True

        self.after(1000, self.update_list)
        if not gui_need_to_refresh():
            return

        self.pinned_list.delete(0, tk.END)
        self.process_list.delete(0, tk.END)
        handles = get_handles()
        for hwnd in handles:
            if not hwnd:
                print("debug [hwnd]:", hwnd)
                continue

            title = get_window_title_from_handle(hwnd)
            if not title:
                print("debug [title]:", hwnd)
                continue

            pid = get_pid_from_handle(hwnd)
            if not pid:
                print("debug [pid]:", hwnd, title)
                continue

            name = get_name_from_pid(pid)
            if not name.lower().endswith((".exe", ".com")):
                print("debug [name]:", hwnd, pid, title)
                continue

            process = Process(pid)
            item = f"{name} ({title}) ({hwnd}) ({pid})"
            #print("-"*10)
            #print(item)
            #print(process.exe())
            #print(process.cmdline())
            #print(process.ppid())
            #print(process.status())
            #print(process.username())
            #print("-"*10)
            if (
                name.lower() == R"explorer.exe".lower()
                and title.lower() == R"Program Manager".lower()
                and process.exe().lower() == R"C:\Windows\explorer.exe".lower()
                and process.cmdline()[0].lower() == R"C:\Windows\Explorer.EXE".lower()
                and process.status() == "running".lower()
            ):
                print("debug [explorer.exe]:", hwnd)
                continue


            if not self.search_var.get().lower() in str(["", "Enter a search term here", item]).lower():
                print("debug [search_term]:", hwnd)
                continue

            if self.update_pinfav_in_gui is None:
                if pid == os.getpid():
                    if self.pinfav_handle is None:
                        self.pinfav_handle = hwnd
                    if not has_exstyle(hwnd, win32defines.WS_EX_TOPMOST):
                        self.update_pinfav_in_gui = True
                        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0,0,0,0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

            if has_exstyle(hwnd, win32defines.WS_EX_TOPMOST):
                if item not in self.pinned_list.get(0, tk.END):
                    self.pinned_list.insert(tk.END, item)
            else:
                if item not in self.pinned_list.get(0, tk.END):
                    self.process_list.insert(tk.END, item)

if __name__ == "__main__":
    TITLE = "PinFav GUI"
    CURRENT_SCRIPT_DIRECTORY = Path(__file__).parent
    ICON_FILE = Path(fR"{CURRENT_SCRIPT_DIRECTORY}\icon.ico")

    app = ProcessManager()
    #METHOD
    #from PIL import Image, ImageTk
    #app.wm_iconphoto(True, ImageTk.PhotoImage(Image.open(icon_path)))
    #METHOD
    #app.wm_iconphoto(False, tk.PhotoImage(file=icon_path))
    #METHOD
    #app.wm_iconbitmap(bitmap=icon_path)
    #app.wm_iconbitmap(default=resource_path("thenounproject.ico"))
    app.wm_iconbitmap(default=ICON_FILE)
    app.mainloop()