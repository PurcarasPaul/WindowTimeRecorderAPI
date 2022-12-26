from tkinter import Listbox, Entry, Scrollbar
from tkinter import NORMAL, DISABLED, LEFT, BOTTOM, RIGHT, BOTH, END
from customtkinter import CTk, CTkButton, CTkLabel, CTkScrollbar, CTkToplevel
from customtkinter import CTkEntry
from customtkinter import set_appearance_mode, set_default_color_theme
from PIL import Image
from WindowTimeRecorder import WindowTimeRecorder
from threading import Thread
from pystray import MenuItem, Icon
from datetime import timedelta
from json import load, dump
from os import path
from sys import exit
from win32api import GetModuleHandle
from win32gui import WNDCLASS, RegisterClass, CreateWindowEx
from win32gui import PumpWaitingMessages
from win32con import WM_QUERYENDSESSION, WM_ENDSESSION, WM_QUIT, WM_DESTROY
from win32con import WM_CLOSE, WS_EX_LEFT, CW_USEDEFAULT
from time import sleep


class WindowTimeRecorderGUI:
    __NUMBER_OF_TOP_USED = 50
    __ICON_NAME = "icon.ico"
    __SETTINGS_FILE = "settings.json"

    def __init__(self):
        self.__load_settings()
        self.__time_recorder_thread()
        self.__set_init_design()
        self.__create_main_window()
        self.__create_event_system()

    def __load_settings(self):
        file = path.join(path.dirname(__file__), self.__SETTINGS_FILE)

        if path.exists(file):
            with open(file) as settings:
                self.__NUMBER_OF_TOP_USED = load(settings)

    def __time_recorder_thread(self):
        self.__window_time_recorder = WindowTimeRecorder()
        self.__window_time_recorder.load_saved_total_time()
        self.__window_time_recorder.start_background_service()
        time_recorder_thread = Thread(
            target=self.__window_time_recorder.record_total_time, daemon=True,
            name='Time Recorder thread')
        time_recorder_thread.start()

    def __set_init_design(self):
        set_appearance_mode("system")
        set_default_color_theme("blue")

    def __create_main_window(self):
        self.__init_main_window()
        self.__create_main_window_list_box()
        self.__create_main_window_scroll_bar()
        self.__create_main_window_search_box()
        self.__create_top_apps_window()
        self._window.withdraw()
        self.__settings.withdraw()
        self.__create_icon_thread()

    def __init_main_window(self):
        self._window = CTk()
        self._window.title("Window Time Recorder")
        self._window.geometry("700x350")
        self._window.protocol('WM_DELETE_WINDOW', self._window.withdraw)
        self._window.iconbitmap(self.__ICON_NAME)

    def __create_main_window_list_box(self):
        self.__list_box = Listbox(self._window, state=DISABLED)
        self.__list_box.pack(fill="both", expand=True)

    def __create_main_window_scroll_bar(self):
        self.__scroll_bar = CTkScrollbar(self.__list_box)
        self.__scroll_bar = CTkScrollbar(self.__list_box)
        self.__scroll_bar.pack(side=RIGHT, fill=BOTH)
        self.__list_box.config(yscrollcommand=self.__scroll_bar.set)
        self.__search_box = CTkEntry(self._window, font=('calibre', 10,
                                     'normal'))

    def __create_main_window_search_box(self):
        self.__search_box = CTkEntry(self._window)
        self.__search_box.insert(0, "Type to search...")
        self.__search_box.pack(side=BOTTOM, anchor="e", padx=8, pady=8)
        self.__search_box.bind("<Button-1>", self.__click_search_box)
        self.__search_box.bind("<Leave>", self.__leave_search_box)

    def __click_search_box(self, *args):
        self.__search_box.delete(0, 'end')

    def __leave_search_box(self, *args):
        if not self.__search_box.get():
            self.__search_box.insert(0, 'Type to search...')

    def __create_top_apps_window(self):
        self.__init_settings_window()
        self.__create_change_top_label()
        self.__create_change_top_apps_entry()
        self.__create_change_top_apps_button()

    def __init_settings_window(self):
        self.__settings = CTkToplevel(self._window)
        self.__settings.title("Settings")
        self.__settings.geometry("480x200")
        self.__settings.protocol('WM_DELETE_WINDOW',
                                 self.__settings.withdraw)
        self.__settings.iconbitmap(self.__ICON_NAME)

    def __create_change_top_label(self):
        change_top_apps_label = CTkLabel(self.__settings,
                                         text='Change top apps number: ',
                                         font=('calibre', 10, 'bold'))
        change_top_apps_label.pack(padx=5, pady=15, side=LEFT)

    def __create_change_top_apps_entry(self):
        self.__change_top_apps_entry = CTkEntry(self.__settings,
                                                font=('calibre', 10, 'normal'))
        self.__change_top_apps_entry.insert(0, str(self.__NUMBER_OF_TOP_USED))
        self.__change_top_apps_entry.pack(padx=5, pady=15, side=LEFT)

    def __create_change_top_apps_button(self):
        change_button = CTkButton(self.__settings, text='Change',
                                  command=self.__change_number)
        change_button.pack(side=BOTTOM, anchor="e", padx=8, pady=8)

    def __change_number(self):
        self.__NUMBER_OF_TOP_USED = int(self.__change_top_apps_entry.get())
        self.__refresh_window()
        self.__settings.withdraw()

    def __create_icon_thread(self):
        icon_thread = Thread(target=self.__create_icon, daemon=True,
                             name='Icon thread')
        icon_thread.start()

    def __create_icon(self):
        image = Image.open(self.__ICON_NAME)
        menu = (MenuItem('Show', self.__show_window),
                MenuItem('Settings',
                self.__change_top_apps_number),
                MenuItem('Quit', self.__quit_window))
        self.__icon = Icon("name", image, "Window Time Recorder", menu)
        self.__icon.run()

    def __show_window(self):
        self._window.deiconify()
        self.__refresh_window()

    def __refresh_window(self):
        scroll_position = self.__scroll_bar.get()[0]
        self.__clear_list_box()
        apps = self.__get_search_box_items()
        self.__insert_apps(apps)
        self.__list_box.yview_moveto(scroll_position)
        self.__repaint_window()

    def __clear_list_box(self):
        self.__list_box['state'] = NORMAL
        self.__list_box.delete(0, END)

    def __get_search_box_items(self):
        search_box_text = self.__search_box.get()
        return (
            self.__window_time_recorder.get_searched_item(search_box_text)
            if search_box_text and search_box_text != "Type to search..."
            else self.__window_time_recorder.get_top_most_used(
                self.__NUMBER_OF_TOP_USED
            )
        )

    def __insert_apps(self, apps):
        for i, app in enumerate(apps):
            time = "{:0>8}".format(str(timedelta(seconds=app[1])))
            self.__list_box.insert(i, str(f"{app[0]} - {time}"))

    def __repaint_window(self):
        if self._window.state() == NORMAL:
            self._window.after(1000, self.__refresh_window)
        self.__list_box['state'] = DISABLED

    def __change_top_apps_number(self):
        self.__refresh_window()
        self._window.deiconify()
        self.__settings.deiconify()
        self.__settings.attributes("-topmost", True)

    def __quit_window(self):
        self.__icon.stop()
        self.__save_settings()
        self.__window_time_recorder.stop_background_service()
        self.__window_time_recorder.save_total_time()
        self._window.destroy()
        exit(0)

    def __save_settings(self):
        file = path.join(path.dirname(__file__), self.__SETTINGS_FILE)

        with open(file, 'w') as settings:
            dump(self.__NUMBER_OF_TOP_USED, settings)

    def __create_event_system(self):
        hinst = GetModuleHandle(None)
        wndclass = WNDCLASS()
        wndclass.hInstance = hinst
        wndclass.lpszClassName = "testWindowClass"
        messageMap = {WM_QUERYENDSESSION: self.__wndproc,
                      WM_ENDSESSION: self.__wndproc,
                      WM_QUIT: self.__wndproc,
                      WM_DESTROY: self.__wndproc,
                      WM_CLOSE: self.__wndproc}

        wndclass.lpfnWndProc = messageMap

        myWindowClass = RegisterClass(wndclass)
        hwnd = CreateWindowEx(WS_EX_LEFT, myWindowClass, "testMsgWindow", 0, 0,
                              0, CW_USEDEFAULT, CW_USEDEFAULT, 0, 0, hinst,
                              None)

        event_thread = Thread(
            target=self.__create_event_thread,
            daemon=True, name='Event thread')
        event_thread.start()

    def __wndproc(self, hwnd, msg, wparam, lparam):
        self.__quit_window()

    def __create_event_thread(self):
        while True:
            PumpWaitingMessages()
            sleep(1)
