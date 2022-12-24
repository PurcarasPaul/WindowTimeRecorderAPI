from win32gui import GetWindowText, GetForegroundWindow
from time import sleep
from json import load, dump
from os import path
from operator import itemgetter


class WindowTimeRecorder:
    __total_time = {}
    __is_app_alive = True

    def load_saved_total_time(self):
        self.__total_time = {}
        file = path.join(path.dirname(__file__), 'totalTime.json')

        if path.exists(file):
            with open(file) as json_file:
                self.__total_time = load(json_file)

    def record_total_time(self):
        try:
            while self.__is_app_alive is True:
                if curr_window := GetWindowText(GetForegroundWindow()):
                    if curr_window not in self.__total_time:
                        self.__total_time[curr_window] = 1
                    else:
                        curr_seconds = self.__total_time.get(curr_window)
                        if curr_seconds is not None:
                            self.__total_time[curr_window] = curr_seconds + 1
                sleep(1)
        finally:
            self.save_total_time()

    def save_total_time(self):
        file = path.join(path.dirname(__file__), 'totalTime.json')
        sorted_total_time = dict(sorted(self.__total_time.items(),
                                        key=itemgetter(1),
                                        reverse=True))
        with open(file, 'w') as json_file:
            dump(sorted_total_time, json_file)

    def stop_background_service(self):
        self.__is_app_alive = False

    def start_background_service(self):
        self.__is_app_alive = True

    def get_top_most_used(self, top_number):
        sorted_total_time = sorted(self.__total_time.items(),
                                   key=itemgetter(1), reverse=True)

        dictionary_length = len(sorted_total_time)
        top_number = min(top_number, dictionary_length)
        return sorted_total_time[:top_number]

    def get_searched_item(self, item):
        total_time = {item: 0}
        for key, value in self.__total_time.items():
            if item.lower() in key.lower():
                total_time[item] += value

        return sorted(total_time.items(), key=itemgetter(1))
