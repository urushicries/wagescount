# Standard libraries
import logging
import tkinter as tk
import re
import sys
import os

# Third-party libraries
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build

# Your own modules
from module_with_functions import (
    ffcwpend,
    ffcwp15,
    find_cells_by_type_content,
    get_dataset_name,
    is_valid_price,
    update_table_from_lists,
    parseDataNamesShift,
    parseINCOMEfromSHEETS,
    makeDictEmpTot,
    update_info_WAGES,
    update_info_everyday,
    update_info_everyday_TRADEPLACES,
    makeDataFromSheets,
    clear_wgslist_ranges,
    toggle_cell_value,
    logger
)
# Get the path to the bundled files
if getattr(sys, 'frozen', False):
    # If the app is frozen (packaged with PyInstaller)
    bundle_dir = sys._MEIPASS
else:
    # If running in a normal Python environment
    bundle_dir = os.path.abspath(os.path.dirname(__file__))



# Создание основного окна
root = tk.Tk()
root.resizable(False, False)
root.title("Программа для расчет З/П Another World")
# Размер окна
window_width = 1170
window_height = 890

# Получаем размеры экрана
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Вычисляем позицию для центрирования окна
position_x = (screen_width - window_width) // 2
position_y = (screen_height - window_height-100) // 2

# Устанавливаем размер окна и позицию
root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
# Добавим новый виджет для логов
log_text = tk.Text(root, height=10, width=100, wrap=tk.WORD, bg="black", fg="white", font=("Roboto", 10))
log_text.grid(row=9, column=0, columnspan=3, pady=20, padx=10)

class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        try:
            msg = self.format(record)
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.yview(tk.END)  # Прокрутка вниз
        except Exception:
            self.handleError(record)

# Настройка логгера
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
text_handler = TextHandler(log_text)
text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logger.addHandler(text_handler)






# Пример использования логгера
# Build the correct path to the JSON file
# Path to your JSON key file
SERVICE_ACCOUNT_FILE = json_path   
# Define the scope for Google Sheets API
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", 
          "https://www.googleapis.com/auth/drive"]
# Authenticate and initialize the client
credentials = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, SCOPES)

client = gspread.authorize(credentials)
service = build('sheets', 'v4', credentials=credentials)

sheetWAGES = client.open("! Таблица расчета зарплаты").worksheet("WGSlist")
# Изначальное значение переменной
days_in_month = 15
#SHEETS in beginning 
sheetKOM =  None
sheetPIK = None
sheetJUNE = None
sheetLM = None

dataKOM = None
dataPIK = None
dataJUNE = None
dataLM = None



def delete_ranges():
    clear_wgslist_ranges(service,sheet_id)
    succes()

def toggle_days():
    """Функция для переключения значения переменной"""
    if tWAGESWHOLEMONTH_var1.get() != False or tSetUpShiftsForAllDays_var3.get() != False or  tIncomeFromShops_var2.get() != False:
        label5.config(text=" ")
    global days_in_month

    if days_in_month == 15:
        days_in_month = 31
        label.config(text=f"Рассчитывать от {days_in_month-15}   до:    {days_in_month} ")
        logger.info("Поменял РП с \"1 до 15\" на \" 16 до 31\"")
        toggle_RP_buton(days_in_month)
    else:
        days_in_month = 15
        label.config(text=f"Рассчитывать от 0{days_in_month-14}    до:    {days_in_month}")
        logger.info("Поменял РП с \" 16 до 31\" на \"1 до 15\" ")
        toggle_RP_buton(days_in_month)

def nothing_picked():
    # Настраиваем текст для label5
    label5.config(text="Выберите хотя бы одну функцию❎", font=("Arial", 15, "bold"))
    # Убираем текст через 3 секунды
    root.after(3000, lambda: label5.config(text=""))

def succes():
    # Настраиваем текст для label6
    label6.config(text="Успех!☑", font=("Arial", 15, "bold"))
    # Убираем текст через 3 секунды
    root.after(3000, lambda: label6.config(text=""))

def toggle_RP_buton(days_in_month):
    toggle_cell_value(sheetWAGES,days_in_month)

def on_button_click(month):
    months_data = {
        "Январь": {"sheet_suffix": "Январь25", "days": days_in_month},
        "Февраль": {"sheet_suffix": "Февраль25", "days": days_in_month},  
        "Март": {"sheet_suffix": "Март25", "days": days_in_month},
        "Апрель": {"sheet_suffix": "Апрель25", "days": days_in_month},
        "Май": {"sheet_suffix": "Май25", "days": days_in_month},
        "Июнь": {"sheet_suffix": "Июнь25", "days": days_in_month},
        "Июль": {"sheet_suffix": "Июль25", "days": days_in_month},
        "Август": {"sheet_suffix": "Август25", "days": days_in_month},
        "Сентябрь": {"sheet_suffix": "Сентябрь25", "days": days_in_month},
        "Октябрь": {"sheet_suffix": "Октябрь25", "days": days_in_month},
        "Ноябрь": {"sheet_suffix": "Ноябрь25", "days": days_in_month},
        "Декабрь": {"sheet_suffix": "Декабрь25", "days": days_in_month}
    }

    month_data = months_data.get(month)

    if month_data:
        try:  
            if tWAGESWHOLEMONTH_var1.get() == False and tSetUpShiftsForAllDays_var3.get() == False and  tIncomeFromShops_var2.get() == False:
                nothing_picked()
            else:
                sheetKOM = client.open("1 отчет").worksheet(month_data["sheet_suffix"])
                sheetPIK = client.open("2 отчет").worksheet(month_data["sheet_suffix"])
                sheetJUNE = client.open("3 отчет").worksheet(month_data["sheet_suffix"])
                sheetLM = client.open("4 отчет").worksheet(month_data["sheet_suffix"])
                
                if tWAGESWHOLEMONTH_var1.get() or tSetUpShiftsForAllDays_var3.get():
                    dataKOM, dataPIK, dataJUNE, dataLM = makeDataFromSheets(month_data["days"], sheetKOM, sheetPIK, sheetJUNE, sheetLM)
                    emp_shiftLST = parseDataNamesShift(dataKOM, dataPIK, dataJUNE, dataLM)
                    dictEMPSHIFT = makeDictEmpTot(emp_shiftLST)

                if tIncomeFromShops_var2.get():
                    incomeKOM, incomePIK, incomeJUNE, incomeLM = parseINCOMEfromSHEETS(client, month_data["sheet_suffix"], shtKOM_id, shtPIK_id, shtJUN_id, shtLM_id)

                # Обработка в зависимости от флажков
                if tWAGESWHOLEMONTH_var1.get():
                    if dictEMPSHIFT and emp_shiftLST:
                        update_info_WAGES(dictEMPSHIFT, emp_shiftLST, sheetWAGES)

                if tSetUpShiftsForAllDays_var3.get():
                    if emp_shiftLST and dictEMPSHIFT:
                        update_info_everyday_TRADEPLACES(month_data["days"], emp_shiftLST, sheetWAGES)
                        update_info_everyday(month_data["days"], emp_shiftLST, sheetWAGES)

                if tIncomeFromShops_var2.get():
                    if incomeKOM and incomePIK and incomeJUNE and incomeLM:
                        update_table_from_lists(sheetWAGES, incomeKOM, incomePIK, incomeJUNE, incomeLM)
                succes()




        except Exception as e:
            logger.error(f"Error occurred while processing sheets for {month}: {e}")

    else:
        logger.error(f"Unknown month: {month}")

  
if __name__ == "__main__":
    # Список месяцев
    months = [
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
    ]




# Создание кнопок для каждого месяца
for i, month in enumerate(months):
    button = tk.Button(root, text=month, width=20, command=lambda m=month: on_button_click(m), bg="black", fg="white",font=("Arial",17),relief="sunken")
    button.grid(row=i//3, column=i%3, padx=10, pady=10)  # Размещение кнопок в сетке




# Создание метки для отображения текущего значения переменной
days_in_month = 31  # Для примера

label = tk.Label(root, text=f"Рассчитывать от {days_in_month-15}    до:    {days_in_month}", font=("Roboto", 14,"bold"), bg="black", fg="white")
label.grid(row=5, column=1, pady=10, padx=10, columnspan=1)
label3 = tk.Label(root, text=f"Кнопки сверху⬆️ включают расчет по условиям(галочкам)\n Кнопка снизу⬇️ сменяет расчетный период\n(не применяется на приход с арен, он рассчитывается за весь месяц)", font=("Roboto", 11,"bold"), bg="black", fg="white")
label3.grid(row=4, column=0, pady=10, ipadx=100, columnspan=3)
# Кнопка для переключения значения
toggle_button = tk.Button(root, text="Сменить расчет🔄", command=toggle_days, font=("Roboto", 14), bg="black", fg="white")
toggle_button.grid(row=6, column=1, pady=30, columnspan=1)
delete_button = tk.Button(root, text="Очистить диапозоны 🗑️", command=delete_ranges, font=("Roboto", 14), bg="black", fg="white")
delete_button.grid(row=6, column=2, pady=30, columnspan=1)
# Переключатель 1
labelWAGES = tk.Label(root,text="Рассчитает зарплату за весь расчетный период(РП)\n1)Будут рассчитаны все отработанные смены со всех арен\n2)Только смены, без бонусов",font=("Roboto", 9 ,"bold"), bg="black", fg="white")
labelWAGES.grid(row=8,column=0,pady=20)
tWAGESWHOLEMONTH_var1 = tk.BooleanVar()
tWAGESWHOLEMONTH_var = tk.Checkbutton(root, text="Рассчитать зарплату за весь РП?", variable=tWAGESWHOLEMONTH_var1, bg="white", fg="black",font=("Roboto",13,"bold"))
tWAGESWHOLEMONTH_var.grid(row=7, column=0, pady=10, padx=10)

# Переключатель 2
labelINCOME = tk.Label(root,text="Рассчитает приход за весь месяц со всех арен\n1)Данная функция работает на весь месяц\n2)В таблице будут рассчитаны бонусы для сотрудников",font=("Roboto", 9,"bold"), bg="black", fg="white")
labelINCOME.grid(row=8,column=1,pady=20)
tIncomeFromShops_var2 = tk.BooleanVar()
tIncomeFromShops_var = tk.Checkbutton(root, text="Рассчитать приход с точек?", variable=tIncomeFromShops_var2, bg="white", fg="black",font=("Roboto",13,"bold"))
tIncomeFromShops_var.grid(row=7, column=1, pady=10, padx=10)

# Переключатель 3 
labelALLDAYS = tk.Label(root,text=" Рассчитает З/П за каждый день\n1) Т.е. 1, 0,5 и тд за каждый рабочий день\n2)Покажется на какой точке был работник в свою смену\n3)Смены по дням расставляются этой функцией",font=("Roboto", 9, "bold"), bg="black", fg="white")
labelALLDAYS.grid(row=8,column=2,pady=20)
tSetUpShiftsForAllDays_var3 = tk.BooleanVar()
tSetUpShiftsForAllDays_var = tk.Checkbutton(root, text="Расставить смены на каждый день?", variable=tSetUpShiftsForAllDays_var3, bg="white", fg="black",font=("Roboto",13,"bold"))
tSetUpShiftsForAllDays_var.grid(row=7, column=2, pady=10, padx=10)


# Создание метки для отображения текущего значения переменной
label2 = tk.Label(root, text="клава    кока    x    feduk\nкабы не было тебя\nver 0.0.3", bg="black", fg="white",font=("Arial",7,"bold"))
label2.grid(row=10, column=1, pady=0, padx=50, columnspan=1)

label5 = tk.Label(root, text="", bg="black", fg="red",font=15)
label5.grid(row=5, column=0)
label6 = tk.Label(root, text="", bg="black", fg="lightgreen",font=15)
label6.grid(row=5, column=2)
# Настройки окна
root.iconbitmap(ico_path)
root.configure(bg="black")

# Запуск главного цикла приложения
root.mainloop()
