import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import logging

# Логгер для модуля
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

def ffcwp15(sheet) -> int | None:
    """Find the first cell with the '15.xx.xxxx' format in the first column."""
    pattern = re.compile(r"^15\.\d{2}\.\d{4}$")
    column_data = sheet.col_values(1)  # Get all values in the first column

    return next(
        (row_num for row_num, value in enumerate(column_data, start=1) if pattern.match(value)),
        None
    )

def ffcwpend(sheet) -> int | None:
    """
    Find the first cell in the first column with one of the following formats:
    '31.xx.xxxx', '30.xx.xxxx', '29.xx.xxxx', or '28.xx.xxxx'.
    Priority is given in that order.
    """
    # Compile regexes once
    pattern31 = re.compile(r"^31\.\d{2}\.\d{4}$")
    pattern30 = re.compile(r"^30\.\d{2}\.\d{4}$")
    pattern29 = re.compile(r"^29\.\d{2}\.\d{4}$")
    pattern28 = re.compile(r"^28\.\d{2}\.\d{4}$")
    
    column_data = sheet.col_values(1)  # Get all values in the first column

    # Initialize cell id holders
    cell_id_31 = None
    cell_id_30 = None
    cell_id_29 = None
    cell_id_28 = None

    # Single pass: record the first occurrence for each pattern if not already found
    for row_num, value in enumerate(column_data, start=1):
        if cell_id_31 is None and pattern31.match(value):
            cell_id_31 = row_num
        if cell_id_30 is None and pattern30.match(value):
            cell_id_30 = row_num
        if cell_id_29 is None and pattern29.match(value):
            cell_id_29 = row_num
        if cell_id_28 is None and pattern28.match(value):
            cell_id_28 = row_num

    # Return according to priority
    if cell_id_31 is not None:
        return cell_id_31
    if cell_id_30 is not None:
        return cell_id_30
    if cell_id_29 is not None:
        return cell_id_29
    if cell_id_28 is not None:
        return cell_id_28

    return None  # No match found

def makeDataFromSheets(pattern : int,*sheets):
    """data for different time\nsheets: KOM | PIK | JUNE | LONDONMALL"""
    sheetKOM, sheetPIK, sheetJUNE, sheetLM = sheets
    if pattern == 15:
        #1-15 числа
        data15KOMENDA = sheetKOM.get(f'A1:M{ffcwp15(sheetKOM)}')
        data15PIK = sheetPIK.get(f'A1:M{ffcwp15(sheetPIK)}')
        data15JUNE = sheetJUNE.get(f'A1:M{ffcwp15(sheetJUNE)}')
        data15LM = sheetLM.get(f'A1:M{ffcwp15(sheetLM)}')
        return  data15KOMENDA, data15PIK, data15LM, data15JUNE

    if pattern == 31:
        #15-31 числа
        data31KOMENDA = sheetKOM.get(f'A{ffcwp15(sheetKOM)}:M{ffcwpend(sheetKOM)+20}')
        data31PIK = sheetPIK.get(f'A{ffcwp15(sheetPIK)}:M{ffcwpend(sheetPIK)+20}')
        data31JUNE = sheetJUNE.get(f'A{ffcwp15(sheetJUNE)}:M{ffcwpend(sheetJUNE)+20}')
        data31LM = sheetLM.get(f'A{ffcwp15(sheetLM)}:M{ffcwpend(sheetLM)+20}')
        return  data31KOMENDA, data31PIK, data31LM, data31JUNE
    
    return None

def is_valid_price(cell_value):
    """Проверяет, соответствует ли строка формату цены с ,00 в конце."""
    if ",00"  in cell_value:
        return True
    else:
        return False

def find_cells_by_type_content(client, spreadsheet_id: str, sheet_name: str) -> list:
    """
    Находит ячейки, содержащие числовые (финансовые) значения в Google Sheets, и возвращает список кортежей (day_index, cell_value).

    Args:
        spreadsheet_id: ID таблицы Google Sheets.
        sheet_name: Название листа.

    Returns:
        Список кортежей (day_index, cell_value) или пустой список, если ячейки не найдены.
        Возвращает None в случае ошибки подключения или доступа.

    Raises:
        ValueError: Если spreadsheet_id или sheet_name не являются строками.
    """

    if not isinstance(spreadsheet_id, str):
        raise ValueError("spreadsheet_id должен быть строкой.")
    if not isinstance(sheet_name, str):
        raise ValueError("sheet_name должен быть строкой.")

    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
        sheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.SpreadsheetNotFound:
        logger.error(f"Ошибка: Таблица с ID '{spreadsheet_id}' не найдена.")
        return None
    except gspread.exceptions.WorksheetNotFound:
        logger.error(f"Ошибка: Лист с именем '{sheet_name}' не найден в таблице.")
        return None
    except Exception as e:
        logger.error(f"Произошла ошибка при доступе к Google Sheets: {e}")
        return None

    cells_with_money_type = []
    day_index = 1

    try:
        all_values = sheet.get_all_values()
    except Exception as e:
        logger.error(f"Произошла ошибка при получении данных с листа: {e}")
        return None
    for row_index, row in enumerate(all_values):
        for col_index, cell_value in enumerate(row):
            if is_valid_price(cell_value):
                cells_with_money_type.append((day_index, float(cell_value.replace(",", ".").replace("\xa0", ""))))
                logger.debug(f"adding this thing to income table - {cell_value.replace(",", ".").replace("\xa0", "")}")
                day_index += 1

    return cells_with_money_type

def get_dataset_name(dataset):
    # Ищем имя переменной, соответствующей переданному объекту
    dataset_name = next(
        (name for name, value in globals().items() if value is dataset),
        None
    )
    # Проверяем, найдено ли имя и соответствует ли оно паттерну
    if dataset_name and re.match(r'^data[A-Z]+$', dataset_name):
        return dataset_name
    return None

def parseINCOMEfromSHEETS(client, month, *sheet_ids) :
    """sheetlinks must be in this order: \n
      sheetKOM, sheetPIK, sheetJUNE, sheetLM 
      return: KOM, PIK, JUNE, LM"""
    sheetKOM, sheetPIK, sheetJUNE, sheetLM = sheet_ids

    logger.debug("INFORMATION ABOUT income IN KOMENDA")
    list_with_income_KOM = find_cells_by_type_content(client,sheetKOM, month)
    logger.debug("INFORMATION ABOUT income IN PIK")
    list_with_income_PIK = find_cells_by_type_content(client,sheetPIK, month)
    logger.debug("INFORMATION ABOUT income IN JUNE")
    list_with_income_JUNE = find_cells_by_type_content(client,sheetJUNE, month)
    logger.debug("INFORMATION ABOUT income IN LM")
    list_with_income_LM = find_cells_by_type_content(client,sheetLM, month)
    return list_with_income_KOM, list_with_income_PIK, list_with_income_JUNE, list_with_income_LM

def parseDataNamesShift(*datasets) -> list:
    """
    Parses employee shifts from multiple datasets.
    
    Args:
        datasets: Variable number of datasets (e.g., dataKOM, data15PIK, etc.).
                  Each dataset is a list of rows, where each row is a list of cell strings.
    
    Returns:
        list: A list of tuples (name, shift, day_index, dataset_name), where:
              - name (str): Employee name.
              - shift (float): Employee shift.
              - day_index (int): The sequential day index within the dataset.
              - dataset_name (str): Name associated with the dataset.
    """
    employee_shifts = []
    dataset_names = {1: "KOMENDA", 2: "PIK", 3: "LM", 4: "JUNE"}

    for sheet_number, dataset in enumerate(datasets, start=1):
        dataset_name = dataset_names.get(sheet_number, "UNKNOWN")
        logger.debug(f"INFORMATION ABOUT SHIFTS IN {dataset_name}")
        day_index = 0  # reset day index for each dataset

        for row in dataset:
            for cell in row:
                if "На смене:" in cell:
                    day_index += 1  # increment day index for every cell indicating a day
                    if "(" in cell and ")" in cell:
                        # Split the cell by whitespace and process each entry
                        entries = cell.split()
                        for entry in entries:
                            if "(" in entry and ")" in entry:
                                try:
                                    name, shift_str = entry.split("(", 1)
                                except ValueError:
                                    continue  # Skip if splitting fails
                                name = name.strip()
                                shift_str = shift_str.strip(")").replace(",", ".")
                                try:
                                    shift = float(shift_str)
                                except ValueError:
                                    continue  # Skip invalid shift values
                                logger.debug(f"name: {name} | shift: {shift} | on day {day_index}")
                                employee_shifts.append((name, shift, day_index, dataset_name))

    return employee_shifts

def makeDictEmpTot (emp_shift:list) -> dict:
    """emp_shift list to dictionary"""
    employee_totals = {}

    for employee, shift, j_ , i_ in emp_shift:
        if employee in employee_totals:
            employee_totals[employee] += shift  
        else:
            employee_totals[employee] = shift  
    return employee_totals

def update_info_WAGES(employee_shift_dict: dict, employee_shiftsList: list, sheetLink) -> None:
    """Обновляем в Таблице ЗП смены за месяц все. (dict with shifts, sheetlink)"""
    logger.debug("starting update_info_WAGES")
    rangeEMPNAMES = 'C97:D109'  # Assuming these are employee names

    
    # Get current data from the sheet
    cell_values = sheetLink.get(rangeEMPNAMES)

    # Prepare batch updates for the employee shifts
    updates = []

    # Iterate through employee shifts for the whole month
    for i_, row in enumerate(cell_values, start=97):
        if row:
            name = row[0]
        else:
            logger.debug("Empty row encountered")
            continue

        # Update employee's shift for the entire month (using employee_shift_dict)
        if name in employee_shift_dict:
            updates.append({
                'range': f'D{i_}',
                'values': [[employee_shift_dict[name]]]
            })
            logger.debug(f"adding shifts {employee_shift_dict[name]} for {name}")
        # Batch update the sheet with the new values
    if updates:
        sheetLink.batch_update(updates)

def update_info_everyday(days_in_month: int, employee_shiftsList: list, sheetLink) -> None:
    """Обновляем в Таблице ЗП смены за каждый день месяц. (dict with shifts, sheetlink)"""
    logger.debug("starting update_info_everyday")
    rangeEMPnamesDAYS = 'D21:P51'  # Assuming this is the range for day shifts

    # Map employee names to column indices in the range
    employee_to_column = {
        'Вова': 1, '__': 2, 'Саша': 3, 'Даня': 4, '_': 5,
        'Илья': 6, 'Костя': 7, 'Максим': 8, 'Никита': 9, 'Павел': 10,
        'Ришат': 11, 'Рома': 12, 'Сева': 13
    }

    # Extract the starting cell's coordinates
    start_row15, start_col15 = 21, ord('D')  # Row 21 and column 'D'
    # Extract the starting cell's coordinates
    start_rowend, start_colend = 36, ord('D')  # Row 35 and column 'D'
    # Prepare updates based on employee_shiftsList
    updates = []
    if days_in_month == 15:
        for employee, value, day, dataset in employee_shiftsList:
            if employee in employee_to_column:
                column_offset = employee_to_column[employee]
                col_letter = chr(start_col15 + column_offset - 1)  # Convert column index to letter
                row = start_row15 + day - 1  # Map the day to the corresponding row
                cell_address = f"{col_letter}{row}"
                updates.append({"range": cell_address, "values": [[value]]})
                logger.debug(f" {employee} |  смена типа {value} | числа: {day} | на арене {dataset}")
    elif days_in_month == 31:
        for employee, value, day, dataset in employee_shiftsList:
            if employee in employee_to_column:
                column_offset = employee_to_column[employee]
                col_letter = chr(start_colend + column_offset - 1)  # Convert column index to letter
                row = start_rowend + day - 1  # Map the day to the corresponding row
                cell_address = f"{col_letter}{row}"
                updates.append({"range": cell_address, "values": [[value]]})
                logger.debug(f" {employee} |  смена типа {value} | числа: {day} | на арене {dataset}")
    # Batch update the sheet with the new values
    if updates:
        sheetLink.batch_update(updates)

def update_info_everyday_TRADEPLACES(days_in_month: int, employee_shiftsList: list, sheetLink) -> None:
    """Обновляем в Таблице ЗП смены "на какой арене" за каждый день в месяце выбранном. (days in month, dict with shifts, sheetlink)"""
    logger.debug("Starting to do update_info_everyday_TRADEPLACES")
    rangeEMPnamesDAYS = 'D59:P89'  # Assuming this is the range for day shifts

    # Map employee names to column indices in the range
    employee_to_column = {
        'Вова': 1, '__': 2, 'Саша': 3, 'Даня': 4, '_': 5,
        'Илья': 6, 'Костя': 7, 'Максим': 8, 'Никита': 9, 'Павел': 10,
        'Ришат': 11, 'Рома': 12, 'Сева': 13
    }

    # Extract the starting cell's coordinates
    start_row15, start_col15 = 59, ord('D')  # Row 59 and column 'D'
    # Extract the starting cell's coordinates
    start_rowend, start_colend = 74, ord('D')  # Row 74 and column 'D'
    # Prepare updates based on employee_shiftsList
    updates = []
    if days_in_month == 15:
        for employee, value, day, dataset in employee_shiftsList:
            if employee in employee_to_column:
                column_offset = employee_to_column[employee]
                col_letter = chr(start_col15 + column_offset - 1)  # Convert column index to letter
                row = start_row15 + day - 1  # Map the day to the corresponding row
                cell_address = f"{col_letter}{row}"
                updates.append({"range": cell_address, "values": [[dataset]]})
                logger.debug(f"{employee} смена в арене {dataset} числа: {day}")
    elif days_in_month == 31:
        for employee, value, day, dataset in employee_shiftsList:
            if employee in employee_to_column:
                column_offset = employee_to_column[employee]
                col_letter = chr(start_colend + column_offset - 1)  # Convert column index to letter
                row = start_rowend + day - 1  # Map the day to the corresponding row
                cell_address = f"{col_letter}{row}"
                updates.append({"range": cell_address, "values": [[dataset]]})
                logger.debug(f"{employee} смена в арене {dataset} числа: {day+15}")
    # Batch update the sheet with the new values
    if updates:
        sheetLink.batch_update(updates)

def update_table_from_lists(sheetLink,*lists) -> None:
    """
    Updates table data for columns with INCOME from TRADEPLACES based on provided lists.
    
    Args:
        lists (list): A list of four lists, each containing tuples with (day_index, value).
        sheetLink: Google Sheets link object to interact with.
    """
    incomeLSTKOM, incomeLSTPIK, incomeLSTJUNE, incomeLSTLM = lists
    fullincomeList = [incomeLSTKOM, incomeLSTPIK, incomeLSTJUNE, incomeLSTLM]
    # Define the starting row and columns for the range
    start_row = 21
    columns = ['Q', 'R', 'S', 'T']  # Corresponding columns for each list
    
    # Prepare batch updates
    updates = []
    logger.debug("starting update_income")
    for i, data_list in enumerate(fullincomeList):
        column_letter = columns[i]  # Determine the column based on the list index
        for day_index, value in data_list:
            # Calculate the row number based on the day index
            row = start_row + day_index - 1
            # Prepare the cell address
            cell_address = f"{column_letter}{row}"
            # Append the update to the batch
            updates.append({"range": cell_address, "values": [[value]]})
            logger.debug(f"В список обновлений добавлено {value} из списка {i}")
    # Perform batch update on the sheet
    if updates:
        sheetLink.batch_update(updates)

def clear_wgslist_ranges(service, spreadsheet_id, ranges = [
        "WGSlist!D21:P51",
        "WGSlist!D97:D109",
        "WGSlist!Q21:T51",
        "WGSlist!D59:P89"
    ]):
    """
    Удаляет данные из заданных диапазонов в таблице WGSlist.

    :param service: Авторизованный объект сервиса Google Sheets API.
    :param spreadsheet_id: ID таблицы Google Sheets.
    """

    try:
        body = {
            "ranges": ranges
        }
        service.spreadsheets().values().batchClear(
            spreadsheetId=spreadsheet_id, body=body
        ).execute()
        logger.debug("Данные успешно удалены из указанных диапазонов.")
    except Exception as e:
        logger.error(f"Ошибка при удалении данных: {e}")

def toggle_cell_value(sheet,days_in_month) -> None:
    """
    Функция принимает объект листа и меняет значение ячейки E93 на листе WGSlist:
    если текущее значение равно "31", меняет на "15", иначе – на "31".

    Аргументы:
        sheet: Объект листа (например, gspread.Worksheet), содержащий лист WGSlist.
    """
    # Предполагается, что объект sheet уже ссылается на лист WGSlist.
    try:
        # Получаем текущее значение ячейки E93
        cell = sheet.acell('E93')
        current_value = cell.value.strip() if cell.value else ""

        if str(days_in_month) != current_value:
            new_value = "31" if str(days_in_month) == "31" else "15"
            # Обновляем значение ячейки E93
            sheet.update_acell('E93', new_value)
            print(f"Значение ячейки E93 изменено с {current_value} на {new_value}")
    except Exception as e:
        print(f"Ошибка при обновлении ячейки: {e}")