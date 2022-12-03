import tkinter as tk
from tkinter import ttk
import pyodbc
import pickle
from cryptography.fernet import Fernet


class ConnectionWindow(tk.Toplevel):
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.master = master
        self.title('Parametry połączenia')
        self.geometry('300x230')
        self.resizable(False, False)
        self.variables = None
        self.dbs = None
        self.dbs_cb = None
        self.create_widgets()

    def connect(self, db_name=None):
        driver = 'SQL Server'
        server = self.variables['Serwer:'].get()
        username = self.variables['Użytkownik:'].get()
        password = self.variables['Hasło:'].get()

        if db_name is not None:
            db = db_name.get()
        else:
            db = ''

        conn = pyodbc.connect(
            f'DRIVER={driver};'
            f'SERVER={server};'
            f'DATABASE={db};'
            f'UID={username};'
            f'PWD={password};'
        )

        if db and conn:
            self.master.variables['db_var'].set(server + '    ' + db)
            self.master.db_label.configure(fg='green')
            self.master.__class__.connection = conn
            self.save_config()
            self.destroy()

        return conn

    def get_dbs(self):
        conn = self.connect()
        cursor = conn.cursor()

        db_query = '''
            SELECT name
            FROM sys.databases
            WHERE name NOT IN ('master', 'tempdb', 'model', 'msdb');
        '''
        cursor.execute(db_query)

        self.dbs = []
        for name in cursor.fetchall():
            self.dbs.append(name[0])

        self.dbs_cb.configure(state='readonly')
        self.dbs_cb['values'] = self.dbs
        self.dbs_cb.set(self.dbs[0])

        conn.close()

    def create_widgets(self):
        label_frame = tk.LabelFrame(self, text='Logowanie')
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        label_frame.grid(column=0, row=0, sticky='NSWE', padx=5, pady=5)

        self.variables = {}
        for index, label_text in enumerate(['Serwer:', 'Użytkownik:', 'Hasło:'], start=0):
            tk.Label(label_frame, text=label_text, anchor='e').grid(column=0, row=index, sticky='E', padx=10, pady=1)

            var = tk.StringVar()
            self.variables[label_text] = var
            if label_text == 'Hasło:':
                tk.Entry(
                    label_frame, textvariable=var, relief='solid',
                    borderwidth=0, highlightbackground='#A0A0A0', highlightthickness=1, width=30, show='*'
                ).grid(column=1, row=index, sticky='W')
            else:
                tk.Entry(
                    label_frame, textvariable=var, relief='solid',
                    borderwidth=0, highlightbackground='#A0A0A0', highlightthickness=1, width=30
                ).grid(column=1, row=index, sticky='W')

        tk.Button(
            label_frame, text='Pobierz podmioty', command=self.get_dbs, width=25
        ).grid(column=1, row=3, sticky='W', pady=5, ipady=2)

        tk.Label(label_frame, text='Baza:', anchor='e').grid(column=0, row=4, sticky='E', padx=10, pady=1)

        db_var = tk.StringVar()
        self.variables['db_var'] = db_var
        self.dbs_cb = ttk.Combobox(label_frame, textvariable=db_var, state='disabled', width=27)
        self.dbs_cb.grid(column=1, row=4, sticky='W')

        tk.Button(
            label_frame, text='Połącz', command=lambda: self.connect(db_name=db_var), width=10
        ).grid(column=1, row=5, sticky='E', pady=30, ipady=2)

        self.load_config()

    def save_config(self):
        key = Fernet.generate_key()
        fernet = Fernet(key)

        with open('cfg.bin', 'wb') as file:
            cfg = {
                'serwer': self.variables['Serwer:'].get(),
                'uzytkownik': self.variables['Użytkownik:'].get(),
                'haslo': bytes(self.variables['Hasło:'].get(), 'utf-8'),
                'db': self.variables['db_var'].get(),
                'fkey': key
            }
            cfg['haslo'] = fernet.encrypt(cfg['haslo'])
            pickled_cfg = pickle.dumps(cfg)
            file.write(pickled_cfg)

    def load_config(self):
        try:
            with open('cfg.bin', 'rb') as file:
                cfg = pickle.load(file, encoding='utf-8')
                fernet = Fernet(cfg['fkey'])
                decr_pass = fernet.decrypt(cfg['haslo']).decode(encoding='utf-8')
                self.variables['Serwer:'].set(cfg['serwer'])
                self.variables['Użytkownik:'].set(cfg['uzytkownik'])
                self.variables['Hasło:'].set(decr_pass)
                self.variables['db_var'].set(cfg['db'])
        except FileNotFoundError as e:
            print("cfg.bin file not found, can't load connection config.")
