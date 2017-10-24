#!/usr/bin/python3

import tkinter
from tkinter.filedialog import askopenfilename
import konwerter
import obrazy
import sys
import subprocess
import os


class Application(tkinter.Frame):
    def __init__(self, master=None):
        tkinter.Frame.__init__(self, master)

        self.pack(fill=tkinter.BOTH, expand=True, padx=5, pady=5)
        self.master.title('Pocztowy konwerter')

        self.button_result_status = False
        self.create_widgets()

    def create_widgets(self):
        self.label_info = tkinter.Label(self, text="Program gotowy do działania.")
        self.label_info.pack(fill=tkinter.BOTH, side=tkinter.TOP, padx=5, pady=5)

        self.button_result = tkinter.Button(self, text="Wybierz plik", command=self.button_result_action)
        self.button_result.pack(fill=tkinter.BOTH, side=tkinter.TOP, padx=5, pady=5)

        quit_button = tkinter.Button(self, text="Wyjdź", command=self.master.quit)
        quit_button.pack(fill=tkinter.BOTH, side=tkinter.RIGHT, expand=True, padx=5, pady=5)

        choose_button = tkinter.Button(self, text='Pomoc', command=self.help)
        choose_button.pack(fill=tkinter.BOTH, side=tkinter.LEFT, expand=True, padx=5, pady=5)

        self.calculate_button = tkinter.Button(self, text='Konwertuj', command=self.convert, state=tkinter.DISABLED)
        self.calculate_button.pack(fill=tkinter.BOTH, side=tkinter.LEFT, expand=True, padx=5, pady=5)

        #self.open_folder_button = tkinter.Button(self, text='otworz', command=self.open_file_browser)
        #self.open_folder_button.pack(fill=tkinter.BOTH, side=tkinter.LEFT, expand=True, padx=5, pady=5)

    def button_result_action(self):
        if self.button_result_status:
            self.open_file_browser()
        else:
            self.choose_file()

    def choose_file(self):
        self.name = askopenfilename(
            initialdir=".",
            filetypes=(("Plik xml", "*.xml"), ("Wszystkie pliki", "*.*")), title="Choose a file.")

        if self.name != "":
            self.label_info['text'] = 'Plik wybrany do konwersji:'
            self.button_result['text'] = self.name
            self.calculate_button['state'] = 'normal'

    def convert(self):
        self.calculate_button['text'] = 'Konwertuj'
        self.button_result_status = False
        try:
            if self.name:
                self.res_file = konwerter.convert_to_template(self.name)
                if self.res_file:
                    if self.res_file == -1:
                        self.label_info['text'] = 'Błąd przy konwertowaniu.'
                        self.button_result['text'] = 'Wybrano niepoprawny plik.'
                    else:
                        self.label_info['text'] = 'Konwertowanie przebiegło pomyślnie. Zapisano jako:'
                        self.button_result['text'] = self.res_file
                        self.name = None
                        self.button_result_status = True
                        self.calculate_button['text'] = 'Reset'
                else:
                    self.label_info['text'] = 'Błąd przy konwertowaniu.'
            else:
                self.calculate_button['state'] = 'disabled'
                self.button_result['text'] = 'Wybierz plik.'
                self.label_info['text'] = 'Program gotowy do działania.'
        except AttributeError:
            self.calculate_button['state'] = 'disabled'
            self.button_result['text'] = 'Wybierz plik.'
            self.label_info['text'] = 'Program gotowy do działania.'

    def open_file_browser(self):
        d = os.path.split(self.res_file)[0]
        if sys.platform == 'win32':
            subprocess.Popen(['start', d], shell=True)
            #subprocess.Popen(r'explorer /select,{}'.format(d), shell=True)

        elif sys.platform == 'darwin':
            subprocess.Popen(['open', d])

        else:
            try:
                subprocess.Popen(['xdg-open', d])
            except OSError:
                pass

    def help(self):
        top_help = tkinter.Toplevel(self)
        top_help.wm_title('Pomoc')
        top_help.geometry("850x700")

        self.label_help_text = tkinter.Label(top_help, text=obrazy.img_description[0])
        self.label_help_text.pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=True, padx=5, pady=5)
        self.label_help_text.config(font='-size 14')

        help_img = tkinter.PhotoImage(data=obrazy.images[0])
        self.num_img = 0

        self.label_img = tkinter.Label(top_help, image=help_img)
        self.label_img.image = help_img
        self.label_img.pack(side=tkinter.TOP, fill="both", expand=True, padx=5, pady=5)

        button_prev = tkinter.Button(top_help, text='<', command=self.prev_img)
        button_prev.pack(fill=tkinter.X, side=tkinter.LEFT, expand=True, padx=5, pady=5)

        self.label_help_num = tkinter.Label(top_help, text='Strona 1')
        self.label_help_num.pack(fill=tkinter.X, side=tkinter.LEFT, expand=True, padx=5, pady=5)

        button_next = tkinter.Button(top_help, text='>', command=self.next_img)
        button_next.pack(fill=tkinter.X, side=tkinter.LEFT, expand=True, padx=5, pady=5)

        button_exit_help = tkinter.Button(top_help, text='Zamknij', command=top_help.destroy)
        button_exit_help.pack(fill=tkinter.X, side=tkinter.LEFT, expand=True, padx=5, pady=5)

    def show_img(self, n):
        help_img = tkinter.PhotoImage(data=obrazy.images[n])
        self.label_img['image'] = help_img
        self.label_img.image = help_img

    def next_img(self):
        self.num_img += 1
        if self.num_img >= len(obrazy.images):
            self.num_img = 0
        self.show_img(self.num_img)
        self.label_help_num["text"] = "Strona " + str(self.num_img + 1)
        try:
            self.label_help_text['text'] = obrazy.img_description[self.num_img]
        except IndexError:
            self.label_help_text['text'] = ''

    def prev_img(self):
        self.num_img -= 1
        if self.num_img < 0:
            self.num_img = len(obrazy.images) - 1
        self.show_img(self.num_img)
        self.label_help_num["text"] = "Strona " + str(self.num_img + 1)
        try:
            self.label_help_text['text'] = obrazy.img_description[self.num_img]
        except IndexError:
            self.label_help_text['text'] = ''

root = tkinter.Tk()
root.geometry("340x120")

app = Application(root)

# app = Application()

app.mainloop()

