import tkinter as tk
from tkinter import ttk


class AnaliticsWindow(tk.Toplevel):
    def __init__(self, master, connection, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.connection = connection
        self.cursor = self.connection.cursor()
        self.title('Analityki działów')
        self.geometry('500x500')
        self.resizable(True, True)
        self.tree = None
        self.create_widgets()

    def create_widgets(self):
        label_frame = tk.LabelFrame(self, text='Działy')
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        label_frame.grid(column=0, row=0, sticky='NSWE', padx=5, pady=5)

        columns = ('dział', 'analityka')
        self.tree = ttk.Treeview(label_frame, columns=columns, show='headings')
        for column in columns:
            self.tree.heading(column, text=column.replace('_', ' ').title())

        label_frame.columnconfigure(0, weight=1)
        label_frame.rowconfigure(0, weight=1)
        self.tree.grid(column=0, row=0, padx=5, pady=5, sticky='NSWE')

        scrollbar = ttk.Scrollbar(self, orient='vertical')
        scrollbar.grid(column=1, row=0, sticky='NS')
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.configure(command=self.tree.yview)

    def query_analitics(self):
        tree_rows = self.tree.get_children()
        for row in tree_rows:
            self.tree.delete(row)

        query = '''
            SELECT IsNull(DZ.dzi_Nazwa, ''), dzi_Analityka
            FROM sl_Dzial DZ
            LEFT JOIN twsf_DzialAnalityka A ON DZ.dzi_Id = A.dzi_Id
            ORDER BY DZ.dzi_Nazwa ASC
        '''
        self.cursor.execute(query)
        departments = self.cursor.fetchall()

        for department in departments:
            self.tree.insert('', 'end', values=tuple(department))
