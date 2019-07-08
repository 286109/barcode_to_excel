from tkinter import ttk, Toplevel, filedialog as fd
import requests
from bs4 import BeautifulSoup
import openpyxl

sources = {
    1: {'url': 'https://barcode-list.ru/barcode/RU/Поиск.htm',
        'cols': ['№', 'Штрих-код', 'Наименование', 'Единица измерения', 'Рейтинг*']},
    2: {'url': 'http://www.goodsmatrix.ru/',
        'cols':[]}
}


# Input barcodes
def get_barcode(widget, root):
    """
        Получение штрих-кода от пользователя.
        Предполагает использование сканера штрих-кодов.
        Создаёт окно для считывания кода, осуществляет считывание (сканер/клавиатура)
        и после перевода каретки вставляет код в таблицу основного окна программы. 
    """
    window = Toplevel()
    bc_entry = ttk.Entry(window, text='Введите штрихкод')
    bc_entry.bind('<Return>', lambda event, tree=widget: enter_pressed(event, tree))
    bc_entry.pack()
    
def enter_pressed(event, tree):
    bc = event.widget.get()
    try:
        tree.insert('', 'end', values=[bc, load_descriprtion(bc), ''])
    except:
        tree.insert('', 'end', values=[bc, '', 'Не найдено'])
    event.widget.delete(first=0, last='end')
    

def load_descriprtion(barcode, url=sources[1]['url']):
    """
    It loads item description in form of attributes searched by barcode from a 
    remote source.
    Returns list of attributes for barcode. 
    """
    bc = {'barcode': barcode}
    r = requests.get(url, params=bc)
    print(r.url)
    soup = BeautifulSoup(r.text, 'html.parser')
    data = []
    table = soup.find('table', attrs={'class':'randomBarcodes'})
    rows = table.find_all('tr')
    # get headers
    cols = rows[0].find_all('th')
    cols = [ele.text.strip() for ele in cols]
    data.append([ele for ele in cols if ele])
    # get rows
    # почему пустая строка появляется в dat
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        data.append([ele for ele in cols if ele])
    return data[2][2]


def save(tv):

    fname = fd.asksaveasfilename(defaultextension='.xlsx', 
        filetypes=[( 'Excel', '*.xlsx')])
    if fname is None:
        return
    # Создание книги Excel, получение листа книги, запись в книгу
    wb = openpyxl.Workbook()
    ws = wb.active
    n_rows = len(tv.get_children())
    for _ in zip(tv.get_children(), range(1, n_rows+1)):
        vals=tv.item(_[0])['values']
        for val_row in zip(vals, range(1,4)):
            ws.cell(row=_[1], column=val_row[1], value=val_row[0])
    wb.save(fname)