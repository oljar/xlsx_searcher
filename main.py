import  os
from tkinter import *
from tkinter import messagebox
import tkinter as tk, threading, tkinter.scrolledtext as ScrolledText, sys, webbrowser
from tkinter import filedialog as fd
import time
import pandas as pd



wynik = str()
b = str()
path = str()
sciezka = ''
lista_plikow = str()
_is_running = 1
count = 0
password = ''



class WorkThread(threading.Thread):

    def __init__(self):
        global count
        super().__init__()
        but = {}
        count = count + 1

        self.count = count
        self.next_page = tk.IntVar()
        self.next_page.set(0)

        self.work_indicator = tk.BooleanVar()


    def next_move(self):
        self.next_page.set(0)

    def work(self):
        if self.work_indicator:
            self.work_indicator.set(False)
        else:
            self.work_indicator.set(True)

    def run(self):
        lista_plikow = str()
        text = str()
        d = ()
        res = None
        score_frame = tk.LabelFrame(window, text='Wyniki')
        score_frame.grid(column=2, row=7)
        lbl.set('szukanie trwa ...')
        wynik = str()
        lista = str()

        try:
            lista_plikow = os.listdir(entry02.get())
        except:
            messagebox.showinfo('info', 'Nieprawidłowa ścieżka dostępu_01')
            window.destroy()

        if self.count != 1:
            res = messagebox.askquestion('Exit Applikation', 'Czy kontynuować')
        if res == 'yes':
            (os.execl)(sys.executable, sys.executable, *sys.argv)
            window.destroy()
        elif res == 'no':
            window.destroy()


        for i in lista_plikow:
            self.next_page.set(0)

            while True:
                if self.work_indicator:
                    break

            def callback(event):
                webbrowser.open_new('file://' + os.path.join(entry02.get(), event.widget['text']))

            if str(i[(-3)]) == 'x' and str(i[(-2)]) == 'l' and str(i[(-1)]) == 's' or str(i[(-4)]) == 'x' and str(
                    i[(-3)]) == 'l' and str(i[(-2)]) == 's' and str(i[(-1)]) == 'x':

                pa = str(os.path.join(str(entry02.get()), i))
                wb = pd.ExcelFile(pa)

            else:
                continue


            sheets_dict = pd.read_excel(pa, sheet_name=None)

            if str(entry.get()) in str(i) or str(entry.get()).lower() in str(i) or str(entry.get()).upper() in str(i):
                nazwa = i
                t = tk.Label(score_frame, text=(nazwa.upper()))
                t.pack()
                t.bind('<Button-1>', callback)
                score_frame_height = score_frame.winfo_height()
                if score_frame_height >= 250:
                    if score_frame_height % 250 < 250:
                        self.next_page.set(1)
                        cont = tk.Button(score_frame, text='Dalej', command=self.next_move)
                        cont.pack()
                        while True:
                            if self.next_page.get() == 0:
                                score_frame.destroy()
                                score_frame = tk.LabelFrame(window, text='Wyniki')
                                score_frame.grid(column=2, row=7)
                                break

            for sheet_name, df in sheets_dict.items():


                # Sprawdzenie, czy znak znajduje się w DataFrame
                value_to_check = str(entry.get())
                result_interior = df.map(lambda x: value_to_check in str(x)).any().any() or value_to_check in df.columns

                if result_interior:
                    wynik = i
                    t = tk.Label(score_frame, text=(wynik.lower()))
                    t.pack()
                    t.bind('<Button-1>', callback)
                    o = tk.Label(score_frame, text=('(' + sheet_name.lower() + ')'))
                    o.pack()
                    score_frame_height = score_frame.winfo_height()
                    if score_frame_height >= 250:
                        if score_frame_height % 250 < 250:
                            self.next_page.set(1)
                            cont = tk.Button(score_frame, text='Dalej', command=self.next_move)
                            cont.pack()
                            while True:
                                if self.next_page.get() == 0:
                                    score_frame.destroy()
                                    score_frame = tk.LabelFrame(window, text='Wyniki')
                                    score_frame.grid(column=2, row=7)
                                    break
        lbl.set('Koniec')



def dialog_window():
    sciezka = fd.askdirectory(initialdir="\\Plgamx2\\dane\\Badanie_Rozwoj\\Dane\\Specyfikacje Techniczne")
    entry02.delete(0, 'end')
    entry02.insert(tk.INSERT, sciezka)


def znajdz():
    thread = WorkThread()
    thread.daemon = True
    thread.start()


def info():
    wininfo = Toplevel()
    wininfo.geometry('661x410')
    wininfo.title('Info')
    scrollbar = ScrolledText.ScrolledText(wininfo)
    scrollbar.pack()
    scrollbar.insert(INSERT,
                     'Program WYSZUKIWARKA XLS/X  - OPIS DZIAŁANIA  \n\nautor  - Jarosław Olszewski RD Klimor\nemail  - jolszewski@klimor.com\n\nWyszukiwarka XLS/X -  wyszukuje nazwy plików które zawierają podane frazy \nlub numery w plikach excel o rozszerzeniach -.xls oraz -.xlsx \n\nOpis :\n\nSCIEŻKA - podaj ścieżkę dostępu wybranego folderu\nHASŁO   – podaj wyszukiwaną frazę lub numer - WPISZ RĘCZNIE\nSZUKAJ/STOP – 1-click start  wyszukiwania/2-gi  click  otwarcie okna dialogowego\nCZY KONTYNUOWAĆ ? - Komunikat okna dialogowego\nTAK – reset programu\nNIE – zamknięcie programu\n\nKliknięcie na wyniki powoduje otwarcie znalezionego pliku\n\nmałymi literami - wyświetlane są nazwy plików gdzie wyszukiwana fraza jest\n                  w ich zawartości\n(w nawiasie)    - wyświetlana jest nazwa zakładki\nWIELKIMI LITERAMI - wyświetlane sa nazwy plików gdzie wyszukiwana fraza jest\n                  w nazwie pliku\nNależy zwrócić uwage na poprawność wpisania hasła (wpisz ręcznie) , gdyż błąd \nznacząco zmieni wynik wyszukiwania.')


def work():
    WorkThread.work()


def exit():
    window.destroy()

window = tk.Tk()
window.title('wyszukiwarka XLS/X')
window.geometry('500x530')
dist = tk.Label(window, width=3)
dist.grid(column=0, row=0)
lab01 = tk.Label(window, text='Scieżka ST', width=10)
lab01.grid(column=1, row=1)
dist = tk.Label(window, width=3)
dist.grid(column=2, row=1)
entry02 = tk.Entry(window, width=45)
entry02.grid(column=2, row=1)
entry02.insert(tk.INSERT, sciezka)
btn_sciezka = tk.Button(window, text='Scieżka', command=dialog_window, width=9)
btn_sciezka.grid(column=4, row=1)
label = tk.Label(window, text='Hasło')
label.grid(column=1, row=3)
entry = tk.Entry(window, width=45)
entry.grid(column=2, row=3)
dist = tk.Label(window, width=3)
dist.grid(column=3, row=1)
btn_szukaj = tk.Button(window, text='Szukaj/Stop', command=znajdz, width=9)
btn_szukaj.grid(column=4, row=3)

# btn_stop = tk.Button(window, text='Stop', command = work,width=9)
# btn_stop.grid(column=4, row=4)

lbl = tk.StringVar()
label = tk.Label(window, textvariable=lbl, font=('Helvetica', 16))
lbl.set('Wpisz hasło')
label.grid(column=2, row=5)
btn_info = tk.Button(window, text='info', command=info, padx='21')
btn_info.grid(column=4, row=5)
pathinfo = tk.Label(window, text=(str(sciezka)))
pathinfo.grid(column=2, row=6)

btn_exit = tk.Button(window, text='Exit', command=exit, width=9)
btn_exit.grid(column=4, row=6)

tk.mainloop()
