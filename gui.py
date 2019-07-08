import lookup_barcodes as lb
from tkinter import ttk, Tk, N,S,W,E, NSEW, NS

class GUI:

    def __init__(self):
        self.root = Tk()
        # главное окно с Treeview и кнопками
        self.frame = ttk.Frame(self.root)
        self.frame.pack(fill='both', expand=True)
        self.frame.columnconfigure(1, weight=1) # column with treeview
        self.frame.rowconfigure(1, weight=1) 
        # кнопка для ввода штрих-кодов
        self.input_bc_btn = ttk.Button(self.frame, text='Ввод штрикодов', 
            command=lambda: lb.get_barcode(self.tv, self.root)) 
        self.input_bc_btn.grid(row=0, column=0)
        # Treeview для отображения таблицы со штрих-кодами и описаниями    
        self.tv = ttk.Treeview(self.frame, show='headings')
        self.tv['columns'] = ('Штрих-код', 'Описание', 'Не найдено')
        self.tv.grid(row=1, column=1, sticky=NSEW)
        for col in self.tv['columns']:
            if col == 'Описание':                
                self.tv.heading(col, text=col)
                self.tv.column(col, width=350)
                continue
            self.tv.heading(col, text=col)
            self.tv.column(col, width=100)
        # Scrollbar для tv
        self.scrollbar = ttk.Scrollbar(self.frame, orient='vertical', command=self.tv.yview)
        self.scrollbar.grid(row=1, column=2, sticky=NS)
        self.tv['yscrollcommand'] = self.scrollbar.set
        # Кнопка для сохранения в excel-файл
        self.save_btn = ttk.Button(self.frame, text='Сохранить',
            command=lambda: lb.save(self.tv)) 
        self.save_btn.grid(row=0, column=1)
        


if __name__ == '__main__':
    my_app = GUI()
    my_app.root.mainloop()