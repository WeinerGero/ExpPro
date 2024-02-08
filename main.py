from sys import argv
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QApplication, QButtonGroup
from PyQt5 import uic
from PyQt5.QtCore import Qt, QUrl, QTimer
from PyQt5.QtMultimedia import QMediaPlayer, QMediaPlaylist, QMediaContent
from PyQt5.QtGui import QIcon, QFont

from random import shuffle, randint
from webbrowser import open_new

#from openpyxl import Workbook, load_workbook
# import xlwings as xw
import csv

import win32gui
import win32con
import ctypes

from functools import partial

import pyautogui
# Отключить fail-safe
pyautogui.FAILSAFE = False



# from record_stream import Interceptor

# import UI in Form
Form, Window = uic.loadUiType("MainMenu.ui")


class MainWindow(QMainWindow, Form):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)
        self.font = QFont("Arial", 14)
        self.volume = 50
        self.path_filenames = ''
        self.num_man = 0
        self.setting_stars = 5
        self.rating_song = 0
        self.valence_mood_song = 5
        self.arosual_mood_song = 5
        self.index_song = 0
        self.order_songs = []
        self.time_sleep = 1 * 1000   # 15 * 1000 = 60 seconds
        self.state_rest_duration = 1 * 1000   # 60 * 1000 = 60 seconds

        self.setWindowTitle('ExpPro')
        icon = QIcon("icon.png")
        self.setWindowIcon(icon)

        self.add_functions()

        # add media player for system
        self.system_media_player = QMediaPlayer(self)
        self.system_playlist = QMediaPlaylist(self)
        self.system_media_player.setPlaylist(self.system_playlist)

        # add media player for experiment
        self.experiment_media_player = QMediaPlayer(self)
        self.experiment_playlist = QMediaPlaylist(self)
        self.experiment_media_player.setPlaylist(self.experiment_playlist)

        # setting base volume
        self.volume = 50

        # crate playlist for system sounds
        self.path_system_sounds = ['calibration_sound.mp3',
                                   'Калибровка.mp3',
                                   'Конец исследования.mp3',
                                   'Конец оценки.mp3',
                                   'Начало исследования.mp3',
                                   'Опрос.mp3',
                                   'Оценка песни.mp3']
        self.path_system_sounds = ['system_sounds/' + path for path in self.path_system_sounds]
        for path_song in self.path_system_sounds:
            song = QMediaContent(QUrl(path_song))
            self.system_playlist.addMedia(song)
        self.system_playlist.setPlaybackMode(QMediaPlaylist.Sequential)
        self.system_media_player.setVolume(self.volume)

    def add_functions(self):
        # MainMenuWidget
        self.chooseSongsButton.clicked.connect(self.choose_songs)
        self.exitButton.clicked.connect(exit_app)

        # SettingsExperimentWidget
        self.returnMenuButton.clicked.connect(self.return_main_menu)
        self.startButton.clicked.connect(self.start_experiment)

        # self.linkLineEdit.setFont(self.font)
        self.linkLineEdit.textChanged.connect(lambda: self.on_text_changed(self.linkLineEdit))
        self.linkLineEdit.setPlaceholderText('Укажите путь до pdf файла')

        # SettingPersonWidget
        self.openQuestionnairePushButton.clicked.connect(self.open_questionnaire)
        self.continuePushButton.clicked.connect(self.open_prepare_equipment)

        # CalibrationWidget
        self.continuePushButton_2.clicked.connect(self.start_listening)
        self.downVolumePushButton.clicked.connect(self.turn_down_volume)
        self.upVolumePushButton.clicked.connect(self.turn_up_volume)

        # PreparationEquipmentWidget
        self.continuePushButton_4.clicked.connect(self.open_calibration)
        self.continuePushButton_4.clicked.connect(self.test_comments)
        self.backQuestionnairePushButton.clicked.connect(self.open_setting_person_widget)

        # End Person Widget
        self.nextPersonPushButton.clicked.connect(self.record_next_person)
        self.endPushButton.clicked.connect(self.open_cancel_experiment)

        # End Experiment Widget
        self.cancelEndPushButton.clicked.connect(self.open_end_person_widget)
        self.endExperimentPushButton.clicked.connect(self.end_experiment)

    def test_comments(self):
        self.windows = get_all_window_titles()
        for hwnd, title in self.windows:
            print(f"HWND: {hwnd}, Title: {title}")
            if 'Нейрон-Спектр.NET' in title:
                self.Neurosoft_hwnd = hwnd
            elif 'ExpPro' == title and is_window_not_a_folder(hwnd):
                self.ExpPro_hwnd = hwnd

        self.write_comment()

    def write_comment(self, code: int = 0):
        # code 0: test
        # code 1: start run song
        # code 2: end song
        # code 3: start_relax_state
        try:
            set_window_to_fullscreen(self.Neurosoft_hwnd)

            pyautogui.hotkey('1')
            pyautogui.write(str(code))
            pyautogui.hotkey(['ctrl', 'enter'])

            set_window_to_fullscreen(self.ExpPro_hwnd)
        except Exception as e: print(code, e)

    def on_text_changed(self, field):
        field.setFont(self.font)

    def open_prepare_equipment(self):

        # end player
        self.stackedWidget.setCurrentWidget(self.PreparationEquipmentWidget)

    def end_experiment(self):
        print("End")

        # need to save data

        # open main menu widget
        self.stackedWidget.setCurrentWidget(self.MainMenuWidget)

    def open_cancel_experiment(self):
        # open end experiment widget
        self.stackedWidget.setCurrentWidget(self.EndExperimentWidget)

    def record_next_person(self):
        # change number of person
        self.num_man += 1

        # open the next person
        self.open_setting_person_widget()

    def choose_songs(self):
        # get paths of songs for the experiment
        self.path_filenames, _ = QFileDialog.getOpenFileNames(self, 'Choose Files', '', 'mp3 Files (*.mp3)')
        print('path of songs are', self.path_filenames)

        self.order_songs = [i for i in range(len(self.path_filenames))]

        # open the settings experiment widget
        self.open_settings_experiment()

    def open_settings_experiment(self):
        # open settings experiment widget
        self.stackedWidget.setCurrentWidget(self.SettingsExperimentWidget)

        # Add UI
        self.linkLineEdit.setPlaceholderText('Вставьте ссылку на опрос')

        self.timeSpinBox.setMinimum(15)
        self.timeSpinBox.setMaximum(120)

        self.numManSpinBox.setMinimum(0)

        # setting ratting stars
        self.setting_stars_group = QButtonGroup()
        self.setting_stars_group.buttonClicked[int].connect(self.choose_setting_stars)

        self.setting_stars_group.addButton(self.ratting_5_RadioButton)
        self.setting_stars_group.addButton(self.ratting_10_RadioButton)

    def choose_setting_stars(self, id: int):
        # select a count of stars for rating
        if id == -3:
            self.setting_stars = 10
        else:
            self.setting_stars = 5

    def return_main_menu(self):
        # open the main menu widget
        self.stackedWidget.setCurrentWidget(self.MainMenuWidget)

    def open_questionnaire(self):
        # open the questionnaire through the link from settings experiment
        link = self.linkLineEdit.text()
        if link:
            open_new(link)  # ADD TRY

    def open_listening_mode(self):
        # run current song
        self.index_song = self.order_songs.pop(0)
        self.experiment_playlist.setCurrentIndex(self.index_song)

        # open the listen widget
        self.stackedWidget.setCurrentWidget(self.ListeningModeWidget)

        # start run song
        QTimer.singleShot(self.time_sleep, partial(self.write_comment, 1))

        # create pause
        QTimer.singleShot(self.time_sleep, self.experiment_media_player.play)
        # self.experiment_media_player.play()

        print('start index is', self.experiment_playlist.currentIndex())
        # print(self.experiment_media_player.errorString())

        # stop the accumulation of modified indexes
        self.system_playlist.currentIndexChanged.disconnect()

        # open rating widget when song end
        self.experiment_playlist.currentIndexChanged.connect(self.open_ratting_mode)

        # to avoid calling the function again
        self.rating_selected_flag = False

    def start_listening(self):
        # начинает цикл песня-рейтинг-песня-рейтинг
        # stop calibration sound
        self.system_media_player.stop()
        self.system_playlist.setPlaybackMode(QMediaPlaylist.CurrentItemOnce)

        # open the listen widget
        self.stackedWidget.setCurrentWidget(self.ListeningModeWidget)

        # play current song of the experiment
        self.experiment_media_player.setVolume(self.volume)

        # start_relax_state
        QTimer.singleShot(23000, partial(self.write_comment, 3))

        # pause for calibration EEG before open the listen widget
        QTimer.singleShot(self.state_rest_duration + 23000 - self.time_sleep, self.open_listening_mode)

        # change system sound for instruction experiment
        self.system_playlist.setCurrentIndex(4)
        self.system_media_player.play()

        # stop the accumulation of modified indexes
        self.system_playlist.currentIndexChanged.disconnect()

        # end instruction of experiment
        self.system_playlist.currentIndexChanged.connect(self.system_media_player.stop)

    def open_ratting_mode(self):

        if not self.rating_selected_flag:
            self.write_comment(2)   # end run song

            self.rating_selected_flag = True

            # stopping play current song of the experiment
            self.experiment_media_player.stop()

            # open the ratting widget
            self.stackedWidget.setCurrentWidget(self.RattingModeWidget)

            # end instruction rating
            self.system_playlist.currentIndexChanged.connect(self.system_media_player.stop)

            # create group of the stars radio buttons
            self.rating_stars_group = QButtonGroup()

            # add radio buttons 1-5 stars
            self.rating_stars_group.addButton(self.rattingRadioButton_1)
            self.rating_stars_group.addButton(self.rattingRadioButton_2)
            self.rating_stars_group.addButton(self.rattingRadioButton_3)
            self.rating_stars_group.addButton(self.rattingRadioButton_4)
            self.rating_stars_group.addButton(self.rattingRadioButton_5)

            if self.setting_stars == 5:
                # hide 6-10 stars
                self.rattingFrame_6_10.hide()
            else:
                # add radio buttons 6-10 stars
                self.rating_stars_group.addButton(self.rattingRadioButton_6)
                self.rating_stars_group.addButton(self.rattingRadioButton_7)
                self.rating_stars_group.addButton(self.rattingRadioButton_8)
                self.rating_stars_group.addButton(self.rattingRadioButton_9)
                self.rating_stars_group.addButton(self.rattingRadioButton_10)

            # hide buttons
            self.enable_button_group(enabled=False)

            # start instruction
            self.system_playlist.setCurrentIndex(6)
            self.system_media_player.play()

            # stop the accumulation of modified indexes
            self.system_playlist.currentIndexChanged.disconnect()

            self.system_playlist.currentIndexChanged.connect(self.system_media_player.stop)
            # self.system_playlist.currentIndexChanged.connect(lambda x: print('changed index', self.system_playlist.currentIndex()))
            QTimer.singleShot(8000, self.enable_button_group)

            self.rating_stars_group.buttonClicked[int].connect(self.choose_rating_stars)

    def disable_button_group(self, enabled: bool = False):
        for button in self.rating_stars_group.buttons():
            button.setEnabled(enabled)

    def enable_button_group(self, enabled: bool = True):
        for button in self.rating_stars_group.buttons():
            button.setEnabled(enabled)
            button.setVisible(enabled)

        # print('enable button group')

    def radio_button_reset(self, id: int):
        self.rating_stars_group.setExclusive(False)
        for btn in self.rating_stars_group.buttons():
            btn.setChecked(False)
        self.rating_stars_group.setExclusive(True)

    def choose_rating_stars(self, id: int):
        # unclick radio buttons
        self.radio_button_reset(id)

        # check id of the pressed radio button
        if id == -2:
            self.rating_song = 1
        elif id == -3:
            self.rating_song = 2
        elif id == -4:
            self.rating_song = 3
        elif id == -5:
            self.rating_song = 4
        elif id == -6:
            self.rating_song = 5
        elif id == -7:
            self.rating_song = 6
        elif id == -8:
            self.rating_song = 7
        elif id == -9:
            self.rating_song = 8
        elif id == -10:
            self.rating_song = 9
        elif id == -11:
            self.rating_song = 10

        self.list_rating_songs_person.append(self.rating_song)
        print(f'{self.index_song}:', self.path_filenames[self.index_song], 'is', self.rating_song)

        self.disable_button_group()

        self.open_choose_mood_widget()

    def open_end_person_widget(self):
        # stop songs playing
        self.experiment_media_player.stop()

        # create data for saving  in excel
        data = self.create_data_for_saving()

        # saving data in excel format
        for i in range(len(data)):
            write_to_csv(self.excel_file, data[i])

        # open the end person widget
        self.stackedWidget.setCurrentWidget(self.EndPersonWidget)

    def create_data_for_saving(self):
        # sort lists
        name_songs = [file_path.split('/')[-1].split('.')[0] for file_path in self.path_filenames]

        sorted_list_rating_songs_person = [item[1] for item in sorted(zip(self.order_songs_copy, self.list_rating_songs_person))]
        sorted_list_mood_songs_person = [item[1] for item in sorted(zip(self.order_songs_copy, self.list_mood_songs_person))]
        print(sorted_list_mood_songs_person)

        data = []
        for i in range(len(self.order_songs_copy)):
            data.append([self.num_man, name_songs[i], sorted_list_rating_songs_person[i], sorted_list_mood_songs_person[i][0], sorted_list_mood_songs_person[i][1]])
            print(f'{i}, {name_songs[i]}:', sorted_list_rating_songs_person[i])

        return data

    def start_experiment(self):
        # change time sleep
        self.time_sleep = self.timeSpinBox.value() * 1000

        # create a playlist for this experiment
        try:
            self.experiment_playlist.clear()
            # self.path_filenames = ['chin-chan-chon-chi-chicha-chochi.mp3',
            #                        'arigato.mp3',
            #                        'u-menya-est-takoi-plan.mp3',
            #                        'tutututu-mem-demotivator.mp3',
            #                        'muzyika-s-proydennoy-missiey-iz-gta-san-andreas.mp3']
            for path_song in self.path_filenames:
                song = QMediaContent(QUrl(path_song))
                self.experiment_playlist.addMedia(song)
            self.experiment_playlist.setPlaybackMode(QMediaPlaylist.CurrentItemOnce)
        except Exception as e: print(1, e)

        # change the number of person
        self.num_man = self.numManSpinBox.value()

        # create excel file for the experiment
        try:
            # self.excel_file = create_excel_file(self.path_filenames, self.num_man)
            self.excel_file = make_name_cvs(self.path_filenames, self.num_man)
        except Exception as e: print(2, e)

        # start the experiment for one person
        self.open_setting_person_widget()

    def open_setting_person_widget(self):
        # change order of songs for one person
        self.order_songs = [i for i in range(len(self.path_filenames))]
        for i in range(10):
            shuffle(self.order_songs)
        self.order_songs_copy = self.order_songs[:]
        print('after shuffle', self.order_songs)

        # take num man from the spin box
        text_num_man = 'Ваш номер: '

        # change number in the text field
        self.numManTextBrowser.setText(text_num_man + str(self.num_man))
        self.numManTextBrowser.setAlignment(Qt.AlignmentFlag.AlignCenter)
        style_sheet = "QTextBrowser { font-family: Arial; font-size: 14pt; text-align: center; }"
        self.numManTextBrowser.setStyleSheet(style_sheet)

        # add person, open SettingPersonWidget
        self.stackedWidget.setCurrentWidget(self.SettingPersonWidget)

        # add list of rating for this person
        self.list_rating_songs_person = []
        self.list_mood_songs_person = []

        # change system sound for setting person
        self.system_playlist.setCurrentIndex(5)
        self.system_media_player.play()

        # end sound setting person
        self.system_playlist.currentIndexChanged.connect(self.system_media_player.stop)

    def open_calibration(self):
        # open the calibration widget
        self.stackedWidget.setCurrentWidget(self.CalibrationWidget)

        # create the calibration playlist
        # media_content = QMediaContent(QUrl.fromLocalFile("calibration_sound.wav"))
        # self.system_playlist.addMedia(media_content)

        # disable button for continuous so that there is no bug with voice accompaniment and sound for calibration
        self.continuePushButton_2.setEnabled(False)

        # start instruction of calibration
        self.volume = 50
        self.system_media_player.setVolume(self.volume)
        self.volumeLevelProgressBar.setValue(self.volume)
        self.system_playlist.setPlaybackMode(QMediaPlaylist.CurrentItemOnce)
        self.system_playlist.setCurrentIndex(1)
        self.system_media_player.play()

        # end instruction of calibration and change mode of playlist
        QTimer.singleShot(4500, self.play_calibration_sound)
        # self.system_playlist.currentIndexChanged.connect(self.play_calibration_sound)

    def play_calibration_sound(self):
        self.system_media_player.stop()
        self.system_playlist.setPlaybackMode(QMediaPlaylist.CurrentItemInLoop)

        # change system sound for calibration
        self.system_playlist.setCurrentIndex(0)
        self.system_media_player.play()

        # able button for continuous so that there is no bug with voice accompaniment and sound for calibration
        self.continuePushButton_2.setEnabled(True)

        # self.continuePushButton_2.clicked.connect(self.system_media_player.stop)

    def turn_down_volume(self):
        # decrease volume for calibration widget
        if self.volume > 5:
            self.volume -= 5
            self.system_media_player.setVolume(self.volume)
            print(self.system_media_player.volume())
            self.volumeLevelProgressBar.setValue(self.volume)

    def turn_up_volume(self):
        # increase volume for calibration widget
        if self.volume < 100:
            self.volume += 5
            self.system_media_player.setVolume(self.volume)
            print(self.system_media_player.volume())
            self.volumeLevelProgressBar.setValue(self.volume)

    def open_choose_mood_widget(self):
        # open the calibration widget
        self.stackedWidget.setCurrentWidget(self.ChooseMoodWidget)
        self.continuePushButton_3.clicked.connect(self.detect_end_song)
        self.ValenceHorizontalSlider.setValue(10)
        self.ArosualHorizontalSlider.setValue(10)

        print("mood opened")

        self.flag_mood = True

    def detect_end_song(self):
        if self.flag_mood:
            self.flag_mood = False

            self.valence_mood_song = self.ArosualHorizontalSlider.value()
            self.arosual_mood_song = self.ValenceHorizontalSlider.value()
            self.list_mood_songs_person.append((self.valence_mood_song, self.arosual_mood_song))
            print(f"mood songs are {self.list_mood_songs_person}")

            # check if it was last song in playlist of this experiment
            print('Empty order song?', not self.order_songs, self.order_songs)
            if not self.order_songs:
                # change system sound for end of experiment
                self.system_playlist.setCurrentIndex(2)
                self.system_media_player.play()

                # stop the accumulation of modified indexes
                self.system_playlist.currentIndexChanged.disconnect()

                # end sound end of experiment
                self.system_playlist.currentIndexChanged.connect(self.system_media_player.stop)

                # end of the person if order is empty
                self.open_end_person_widget()  # add pause

                print('End of experiment')
            else:
                # change system sound for end of rating
                self.system_playlist.setCurrentIndex(3)
                self.system_media_player.play()

                # stop the accumulation of modified indexes
                self.system_playlist.currentIndexChanged.disconnect()

                # end sound end of rating
                self.system_playlist.currentIndexChanged.connect(self.system_media_player.stop)

                # switch to the listening widget if order is fully
                QTimer.singleShot(3000, self.open_listening_mode)

                print("Cont")


def change_widgets(old_widget, new_widget):
    old_widget.hide()
    new_widget.show()


def exit_app():
    exit()


def write_to_csv(file_path, data: list = ['Участник', 'Песня', 'Рейтинг', 'Направление эмоций','Сила эмоции']):
    with open(file_path, 'a', newline='') as csv_file:
        writer = csv.writer(csv_file, delimiter='@')
        writer.writerow(data)


def make_name_cvs(path_filenames: list, num_man: int = 0):
    # Создание названия файла
    random_index = randint(0, 1000)
    first_elements = [file_path.split('/')[-1].split('.')[0] for file_path in path_filenames]
    first_chars = [item[0] for item in first_elements]
    file_path = 'out_data/' + ''.join(first_chars) + '_' + str(num_man) + str(random_index) + '.csv'

    write_to_csv(file_path)
    return file_path


def get_all_window_titles():
    def is_real_window(hwnd):
        if not win32gui.IsWindowVisible(hwnd):
            return False
        if win32gui.GetParent(hwnd) != 0:
            return False
        if win32gui.GetWindowText(hwnd) == "":
            return False
        return True

    window_titles = []
    win32gui.EnumWindows(lambda hwnd, param: param.append((hwnd, win32gui.GetWindowText(hwnd))), window_titles)
    return [(hwnd, title) for hwnd, title in window_titles if is_real_window(hwnd)]


def set_window_to_fullscreen(goal_hwnd):
    win32gui.ShowWindow(goal_hwnd, win32con.SW_MAXIMIZE)
    win32gui.SetForegroundWindow(goal_hwnd)


def is_window_not_a_folder(goal_hwnd):
    class_name = win32gui.GetClassName(goal_hwnd)
    return class_name != "CabinetWClass"


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long),
                ("top", ctypes.c_long),
                ("right", ctypes.c_long),
                ("bottom", ctypes.c_long)]


if __name__ == '__main__':
    from sys import exit

    app = QApplication(argv)
    window = MainWindow()
    window.showMaximized()
    exit(app.exec())
