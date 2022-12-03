import tkinter as tk
from tkinter import ttk
from connection import ConnectionWindow
from analitics import AnaliticsWindow
from query import Query
from datetime import datetime

MONTHS = {
    'styczeń': 1,
    'luty': 2,
    'marzec': 3,
    'kwiecień': 4,
    'maj': 5,
    'czerwiec': 6,
    'lipiec': 7,
    'sierpień': 8,
    'wrzesień': 9,
    'październik': 10,
    'listopad': 11,
    'grudzień': 12
}


class NavigationBar(tk.Frame):
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.grid(column=0, row=0, sticky='WE')
        self.variables = None
        self.month_cb = None
        self.create_nav_widgets()

    def create_nav_widgets(self):
        tk.Label(self, text='Miesiąc:').grid(column=0, row=0, padx=3)

        self.variables = {}
        month_var = tk.StringVar()
        self.variables['month_var'] = month_var

        self.month_cb = ttk.Combobox(self, textvariable=month_var, width=15)
        self.month_cb.grid(column=1, row=0)
        self.month_cb['values'] = [month for month in MONTHS]
        self.month_cb.set('styczeń')
        self.month_cb.configure(state='readonly')

        tk.Label(self, text='Rok:').grid(column=2, row=0, padx=3)

        year_var = tk.IntVar()
        year_var.set(datetime.now().year)
        self.variables['year_var'] = year_var

        tk.Entry(
            self, textvariable=year_var, relief='solid', borderwidth=0, highlightbackground='#A0A0A0',
            highlightthickness=1, width=15
        ).grid(column=3, row=0)

        tk.Button(
            self, text='Wylicz', relief='raised', width=10, borderwidth=1, command=self.show_raports
        ).grid(column=4, row=0, padx=10)

        self.columnconfigure(5, weight=1, minsize=130)
        analitics = tk.Label(self, text='Analityki działów', anchor='e', fg='blue', cursor='hand2')
        analitics.grid(column=5, row=0, pady=5, padx=21, sticky='E')
        analitics.bind('<Button-1>', lambda e: self.show_analitics(e))

    def show_analitics(self, event):
        if StatusBar.connection is not None:
            analitics_window = AnaliticsWindow(self, StatusBar.connection)
            analitics_window.query_analitics()

    def show_raports(self):
        tree_rows = DataFrame.tree.get_children()
        for row in tree_rows:
            DataFrame.tree.delete(row)

        if StatusBar.connection is not None:
            query = Query(
                connection=StatusBar.connection,
                month=MONTHS[self.variables['month_var'].get()],
                year=self.variables['year_var'].get()
            )

            query.execute_queries()
            for index, row in enumerate(query.rows):
                if (
                    'Osob. fund. pł. z ZUS - obcy zlecenia - inne odprawy ekon./emer. '
                    'Wartość brutto duże bez wynagrodzenia chorobowego' in row
                ) or (
                    'Suma składników związanych z KZP' in row
                ) or (
                    'Razem potrącenia podatek od bonów' in row
                ):
                    DataFrame.tree.insert('', 'end', values=row, tags='warning')
                else:
                    DataFrame.tree.insert('', 'end', values=row)


class DataFrame(tk.Frame):
    tree = None

    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.grid(column=0, row=1, sticky='NSWE', padx=5)
        self.create_widgets()

    def create_widgets(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Treeview', rowheight=25)

        columns = ('konto', 'kwota_wn', 'kwota_ma', 'opis', 'opis2', 'analityka')
        DataFrame.tree = ttk.Treeview(self, columns=columns, show='headings')
        for column in columns:
            DataFrame.tree.heading(column, text=column.replace('_', ' ').title())
            DataFrame.tree.column(f'{column}', width=1, stretch=True)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        DataFrame.tree.grid(column=0, row=0, sticky='NSWE')

        DataFrame.tree.tag_configure('warning', background='#F0FA56')

        DataFrame.scrollbar = ttk.Scrollbar(self, orient='vertical')
        DataFrame.scrollbar.grid(column=1, row=0, sticky='NS')
        DataFrame.tree.configure(yscrollcommand=DataFrame.scrollbar.set)
        DataFrame.scrollbar.configure(command=DataFrame.tree.yview)


class StatusBar(tk.Frame):
    connection = None

    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.master = master
        self.grid(column=0, row=2, sticky='WE')
        self.variables = None
        self.db_label = None
        self.create_widgets()

    def create_widgets(self):
        connection_lbl = tk.Label(self, text='Połączenie:', fg='blue', cursor='hand2')
        connection_lbl.grid(column=0, row=0, padx=3, pady=5)
        connection_lbl.bind('<Button-1>', lambda e: ConnectionWindow(self))

        self.variables = {}
        db_var = tk.StringVar()
        db_var.set('brak połączenia')
        self.variables['db_var'] = db_var

        self.db_label = tk.Label(self, textvariable=db_var, fg='red')
        self.db_label.grid(column=1, row=0)

        self.columnconfigure(2, weight=1)
        tk.Button(
            self, text='Skopiuj', relief='raised', width=10, borderwidth=1, command=self.copy_rows
        ).grid(column=2, row=0, padx=22, sticky='E')

    def copy_rows(self):
        self.master.master.clipboard_clear()

        headings = []

        for i in range(1, len(DataFrame.tree.cget('columns')) + 1):
            headings.append(DataFrame.tree.heading(f'#{i}')['text'])

        self.master.master.clipboard_append(
            f'{headings[0]}\t{headings[1]}\t{headings[2]}\t{headings[3]}\t{headings[4]}\t{headings[5]}\n'
        )

        rows = DataFrame.tree.get_children()

        for row in rows:
            row_values = DataFrame.tree.item(row)['values']

            self.master.master.clipboard_append(
                f'{row_values[0]}\t{row_values[1]}\t{row_values[2]}\t{row_values[3]}\t{row_values[4]}\t{row_values[5]}\n'
            )


class MainFrame(tk.Frame):
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.master = master
        self.grid(column=0, row=0, sticky='NSWE')
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        self.nav_bar = NavigationBar(self)
        self.data_frame = DataFrame(self)
        self.status_bar = StatusBar(self)
