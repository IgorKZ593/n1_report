import xlwings as xw
import os

# Устанавливаем путь к папке, где хранятся файлы
folder_path = r"F:\Python Projets\Report_N1\Data"

# Находим первый файл, имя которого начинается с 'src_' и оканчивается на '.xlsx'
file_name = None
for f in os.listdir(folder_path):
    if f.startswith('src_') and f.endswith('.xlsx'):
        file_name = f
        break

if not file_name:
    raise FileNotFoundError("Файл, начинающийся с 'src_' не найден в указанной папке.")

# Формируем имя нового файла по заданным правилам
# Убираем 'src_' и добавляем 'отчет_'
base_name = file_name.replace('src_', '', 1)
new_file_name = f"отчет_{base_name}"
new_file_path = os.path.join(folder_path, new_file_name)
#+++++++++++++++++++++++++++++
print('test_123')
#++++++++++++++++++++++++++++++
# Открываем исходный файл и создаём новый файл на его основе
with xw.App(visible=False) as app:
    wb_src = app.books.open(os.path.join(folder_path, file_name))

    # Создаём новую книгу (отчет)
    wb_new = app.books.add()

    # Проверяем, существует ли лист 'общие данные'. Если да – удаляем
    if 'общие данные' in [s.name for s in wb_new.sheets]:
        wb_new.sheets['общие данные'].delete()

    # Создаем новый лист с именем 'общие данные'
    sht = wb_new.sheets.add('общие данные', before=wb_new.sheets[0])

    # Окрашиваем ярлык листа в светло-синий (RGB = 173, 216, 230)
    sht.api.Tab.Color = 16764057  # системное значение цвета light blue в Excel

    # Сохраняем новую книгу по сформированному пути
    wb_new.save(new_file_path)
    wb_new.close()
    wb_src.close()

# Выводим финальное сообщение в консоль
print(f"✅ Новый файл '{new_file_name}' создан и сохранен по пути: {new_file_path}")

# === Добавление листа "текущие цены" ===

# Открываем ранее созданный файл для модификации
with xw.App(visible=False) as app:
    wb = app.books.open(new_file_path)

    # Проверяем наличие листа 'текущие цены' и удаляем, если есть
    if 'текущие цены' in [s.name for s in wb.sheets]:
        wb.sheets['текущие цены'].delete()

    # Создаем новый лист 'текущие цены'
    sht = wb.sheets.add('текущие цены')

    # Окрашиваем ярлык в светло-коричневый
    # Пример RGB (210, 180, 140) = tan, Excel BGR Decimal = 9234160
    sht.api.Tab.Color = 9234160

    # Формируем перечень всех листов
    sheets_list = [s.name for s in wb.sheets]

    # Выводим информацию в консоль
    print("\n📊 В книге:")
    print(f"Полное имя: {wb.name}")
    print(f"Путь к файлу: {new_file_path}")
    print(f"Созданы листы: {', '.join(sheets_list)}")
    print("✅ Книга сохранена.")

    # Сохраняем и закрываем книгу
    wb.save()
    wb.close()

    # === Создание таблицы на листе "текущие цены" ===

    # Открываем книгу снова для добавления таблицы
    with xw.App(visible=False) as app:
        wb = app.books.open(new_file_path)
        sht = wb.sheets['текущие цены']

        # Задаем заголовки таблицы в первую строку
        headers = ['ISIN', 'Актив', 'Тикер', 'Тип', 'Дата', 'Цена']
        sht.range('A1').value = headers

        # Применяем жирное форматирование к заголовкам
        sht.range('A1:F1').api.Font.Bold = True

        # Автоширина для всех столбцов таблицы
        sht.range('A:F').autofit()

        # Выводим финальное сообщение в консоль
        print("✅ Таблица 'текущие цены' создана. Требует заполнения.")

        # Сохраняем и закрываем книгу
        wb.save()
        wb.close()

import xlwings as xw
import os

import re

# === Блок сопоставления ISIN с тикерами и заполнения листа 'текущие цены' ===

folder_path = r"F:\Python Projets\Report_N1\Data"

# 🔷 Функция для проверки валидности ISIN
def is_valid_isin(isin):
    """
    Проверяет ISIN по стандарту:
    - 12 символов
    - Первые 2 – буквы (код страны)
    - Далее 9 букв/цифр
    - Последний символ – цифра
    """
    pattern = r'^[A-Z]{2}[A-Z0-9]{9}[0-9]$'
    return bool(re.match(pattern, isin))

# Формируем имя src_ файла на основе ранее созданного отчет_ файла
src_file_name = new_file_name.replace('отчет_', 'src_')
src_file_path = os.path.join(folder_path, src_file_name)

# Путь к справочнику tickers
tickers_file_path = os.path.join(folder_path, 'tickers.xlsx')

with xw.App(visible=False) as app:
    # 🔷 Открываем src_ файл
    wb_src = app.books.open(src_file_path)
    ws_portfolio = wb_src.sheets['портфель']

    # 🔷 Находим колонку ISIN по заголовку
    header_row = ws_portfolio.range('1:1').value
    if 'ISIN' not in header_row:
        raise ValueError("❌ В файле src_ не найден столбец 'ISIN'.")

    isin_col_index = header_row.index('ISIN') + 1  # Excel columns are 1-based

    # 🔷 Собираем все ISIN в список, оставляя только валидные по формату
    last_row = ws_portfolio.range('A' + str(ws_portfolio.cells.last_cell.row)).end('up').row
    isin_values = ws_portfolio.range((2, isin_col_index), (last_row, isin_col_index)).value

    # Приводим к списку
    if isinstance(isin_values, str):
        isin_values = [isin_values]
    elif isinstance(isin_values, tuple):
        isin_values = list(isin_values)

    # 🔷 Фильтруем ISIN: исключаем None, пустые строки и невалидные по формату
    isins = []
    for idx, isin in enumerate(isin_values, start=2):
        if isin is None:
            continue
        isin_str = str(isin).strip().upper()
        if isin_str == "":
            continue
        if not is_valid_isin(isin_str):
            print(f"⚠️ Строка {idx}: '{isin_str}' не соответствует формату ISIN. Игнорирована.")
            continue
        isins.append(isin_str)

    wb_src.close()

    # 🔷 Открываем файл tickers
    wb_tickers = app.books.open(tickers_file_path)
    ws_tickers = wb_tickers.sheets['акции_фонды']

    # 🔷 Формируем словарь ISIN → (Наименование, Тикер, Тип)
    ticker_data = {}
    tickers_range = ws_tickers.range('A1').expand('table').value  # предполагается компактная таблица
    tickers_headers = tickers_range[0]

    isin_idx = tickers_headers.index('ISIN')
    name_idx = tickers_headers.index('наименование')
    ticker_idx = tickers_headers.index('тикер')
    type_idx = tickers_headers.index('тип')

    for row in tickers_range[1:]:
        ticker_data[row[isin_idx]] = (row[name_idx], row[ticker_idx], row[type_idx])

    wb_tickers.close()

    # 🔷 Открываем файл отчет_ для внесения данных в лист 'текущие цены'
    wb_report = app.books.open(new_file_path)
    ws_prices = wb_report.sheets['текущие цены']

    # 🔷 Заполняем таблицу текущих цен
    row_idx = 2
    not_found = []

    for isin in isins:
        if isin in ticker_data:
            name, ticker, type_ = ticker_data[isin]
            ws_prices.range(f"A{row_idx}").value = isin
            ws_prices.range(f"B{row_idx}").value = name
            ws_prices.range(f"C{row_idx}").value = ticker
            ws_prices.range(f"D{row_idx}").value = type_
        else:
            not_found.append(isin)  # ISIN без совпадений
        row_idx += 1

    # 🔷 Добавляем ISIN без совпадений внизу таблицы
    for isin in not_found:
        ws_prices.range(f"A{row_idx}").value = isin
        ws_prices.range(f"B{row_idx}").value = "НЕ найдено"
        row_idx += 1

    # 🔷 Автоширина столбцов после заполнения
    ws_prices.range('A:F').autofit()

    # 🔷 Сохраняем и закрываем файл отчет_
    wb_report.save()
    wb_report.close()

# === Финальное сообщение ===

if not_found:
    print("⚠️ ISIN без совпадений:", ', '.join(str(isin) for isin in not_found))
else:
    print("✅ Все ISIN нашли совпадения. ISIN без совпадений не обнаружены.")


# === Блок создания листа "неопознанные ISIN" ===

with xw.App(visible=False) as app:
    wb_report = app.books.open(new_file_path)

    # Удаляем лист, если уже существует
    try:
        ws_unrecognized = wb_report.sheets['неопознанные ISIN']
        ws_unrecognized.delete()
    except:
        pass  # если листа нет, ничего не делаем

    # Создаем новый лист
    ws_unrecognized = wb_report.sheets.add(name='неопознанные ISIN', after=wb_report.sheets.count)

    # Окрашиваем ярлык в ярко-красный
    ws_unrecognized.api.Tab.Color = 255  # RGB (255,0,0)

    # Заголовки
    ws_unrecognized.range("A1").value = "ISIN"
    ws_unrecognized.range("B1").value = "Комментарий"

    # Заполняем данными
    if not_found:
        row_idx = 2
        for isin in not_found:
            ws_unrecognized.range(f"A{row_idx}").value = isin
            ws_unrecognized.range(f"B{row_idx}").value = "Нет в справочнике"
            row_idx += 1
        print(f"✅ Создан лист 'неопознанные ISIN' с {len(not_found)} записями.")
    else:
        print("✅ Все ISIN найдены. Лист 'неопознанные ISIN' пуст.")

    # Автоширина
    ws_unrecognized.range("A:B").autofit()

    # Сохраняем и закрываем книгу
    wb_report.save()
    wb_report.close()


# === Блок получения цен по тикерам с проверкой даты и удалением пустых строк ===

import yfinance as yf
import datetime
import holidays
import time

# === Блок получения цен по тикерам с проверкой даты и удалением пустых строк ===

us_holidays = holidays.US()

# 🔷 Функция проверки, является ли дата рабочим днем в США
def is_valid_us_trading_day(date_obj):
    return date_obj.weekday() < 5 and date_obj not in us_holidays

# 🔷 Запрашиваем дату у пользователя
while True:
    date_input = input("Введите дату в формате dd/mm/yyyy: ")
    try:
        query_date = datetime.datetime.strptime(date_input, "%d/%m/%Y").date()
        if not is_valid_us_trading_day(query_date):
            print("⚠️ Введенная дата приходится на выходной или государственный праздник США. Попробуйте снова.")
            continue
        break
    except ValueError:
        print("❌ Неверный формат даты. Попробуйте снова.")

# 🔷 Открываем файл отчет_ для внесения цен
with xw.App(visible=False) as app:
    wb_report = app.books.open(new_file_path)
    ws_prices = wb_report.sheets['текущие цены']

    # Определяем последнюю заполненную строку
    last_row = ws_prices.range('A' + str(ws_prices.cells.last_cell.row)).end('up').row

    # 🔷 Проходим по каждой строке таблицы текущих цен (снизу вверх)
    for row in range(last_row, 1, -1):  # от last_row до 2
        ticker = ws_prices.range(f"C{row}").value  # колонка C - тикер

        if not ticker or str(ticker).strip() == "":
            print(f"⚠️ Строка {row}: тикер отсутствует, строка будет удалена.")
            ws_prices.api.Rows(row).Delete()
            continue

        success = False
        attempts = 0
        price = None

        # 🔷 Пытаемся получить цену до 3 раз
        while attempts < 3 and not success:
            try:
                df = yf.download(ticker, start=query_date, end=query_date + datetime.timedelta(days=1), progress=False)
                if not df.empty:
                    close_price = float(df['Close'].iloc[0])

                    # 🔷 Вставляем дату в колонку E
                    ws_prices.range(f"E{row}").value = query_date.strftime("%d.%m.%Y")

                    # 🔷 Вставляем цену как число
                    ws_prices.range(f"F{row}").value = close_price

                    # 🔷 Устанавливаем NumberFormat Excel
                    ws_prices.range(f"F{row}").api.NumberFormat = '# ##0,000'

                    print(f"✅ {ticker}: цена получена {close_price:.3f}")
                    success = True
                else:
                    attempts += 1
                    print(f"⚠️ {ticker}: цена не получена, попытка {attempts}/3")
                    time.sleep(1)
            except Exception as e:
                attempts += 1
                print(f"❌ {ticker}: ошибка при получении цены ({e}), попытка {attempts}/3")
                time.sleep(1)

        # 🔷 Если после 3 попыток цена не получена – записываем 'NA'
        if not success:
            ws_prices.range(f"E{row}").value = query_date.strftime("%d.%m.%Y")
            ws_prices.range(f"F{row}").value = "NA"
            print(f"❌ {ticker}: цена не получена после 3 попыток, записано 'NA'.")

    # 🔷 Автоширина столбцов E и F после заполнения
    ws_prices.range('E:F').autofit()

    # 🔷 Сохраняем и закрываем файл отчет_
    wb_report.save()
    wb_report.close()

print("✅ Получение цен завершено.")


# === Блок поиска неопознанных ISIN в "структурные_продукты" и обновления комментариев ===

import xlwings as xw

folder_path = r"F:\Python Projets\Report_N1\Data"
tickers_file_path = os.path.join(folder_path, 'tickers.xlsx')
report_file_path = new_file_path  # путь к твоему отчетному файлу

with xw.App(visible=False) as app:
    # 🔷 Открываем отчетный файл
    wb_report = app.books.open(report_file_path)
    ws_unrec = wb_report.sheets['неопознанные ISIN']

    # Собираем ISIN с их строками
    isins_range = ws_unrec.range('A2').expand('down')
    isins = isins_range.value

    # === 🔷 ИСПРАВЛЕНИЕ: обработка случая отсутствия неопознанных ISIN ===
    if isins is None:
        print("✅ Непознанных ISIN нет. Проверка структурных продуктов не требуется.")
    else:
        if isinstance(isins, str):
            isins = [isins]

        # 🔷 Открываем файл tickers
        wb_tickers = app.books.open(tickers_file_path)
        ws_struct = wb_tickers.sheets['структурные_продукты']

        # Загружаем всю таблицу в память
        struct_table = ws_struct.range('A1').expand('table').value
        headers = struct_table[0]

        isin_idx = headers.index('ISIN')
        issuer_idx = headers.index('эмитент')

        # Формируем словарь ISIN → эмитент
        struct_dict = {row[isin_idx]: row[issuer_idx] for row in struct_table[1:]}

        # 🔷 Проверяем каждый ISIN и обновляем комментарий в листе
        found_count = 0
        not_found_count = 0

        for i, isin in enumerate(isins, start=2):  # начинаем с 2-й строки
            issuer = struct_dict.get(isin)
            if issuer:
                print(f"✅ Неопознанный ISIN {isin} является Структурным продуктом. Эмитент: {issuer}")
                ws_unrec.range(f"B{i}").value = "Структурный продукт"
                found_count += 1
            else:
                not_found_count += 1

        wb_tickers.close()

        # === Финальное сообщение ===
        print(f"\n🔎 Результат проверки структурных продуктов:")
        print(f"✔️ Найдено структурных продуктов: {found_count}")
        print(f"❌ Не удалось опознать: {not_found_count}")

    # 🔷 Сохраняем и закрываем отчетную книгу
    wb_report.save()
    wb_report.close()

print(f"✅ Файл '{report_file_path}' сохранен.")



# === Минималистичное форматирование листа 'неопознанные ISIN' ===

with xw.App(visible=False) as app:
    wb = app.books.open(report_file_path)

    # Проверяем, есть ли лист 'неопознанные ISIN'
    if 'неопознанные ISIN' in [sheet.name for sheet in wb.sheets]:
        ws_unrec = wb.sheets['неопознанные ISIN']

        # Определяем последнюю строку с данными
        last_row = ws_unrec.range('A' + str(ws_unrec.cells.last_cell.row)).end('up').row

        # 🔷 Делаем заголовок жирным
        ws_unrec.range("A1:B1").api.Font.Bold = True

        # 🔷 Применяем автоширину ко всем столбцам с данными
        ws_unrec.range(f"A1:B{last_row}").columns.autofit()

        print("✅ Лист 'неопознанные ISIN' отформатирован (заголовок жирный, автоширина применена).")

    wb.save()
    wb.close()

import xlwings as xw

# === Блок создания листа акции_etf ===

with xw.App(visible=False) as app:
    wb = app.books.open(report_file_path)

    # Удаляем лист 'акции_etf', если он существует
    if 'акции_etf' in [sheet.name for sheet in wb.sheets]:
        wb.sheets['акции_etf'].delete()

    # Создаем новый лист 'акции_etf'
    ws_ae = wb.sheets.add('акции_etf', after=wb.sheets[-1])

    # Окрашиваем ярлык в #65bd93
    ws_ae.api.Tab.Color = 0x65bd93

    # 🔷 Формируем заголовок в строке 1
    headers = ["N", "Актив", "Тикер", "Количество", "Цена входа", "Объем входа",
               "Цена текущая", "Объем текущий", "Разница USD", "Разница %"]
    ws_ae.range("A1:J1").value = headers

    # 🔷 Применяем жирный шрифт, выравнивание по центру и цвет заливки
    header_range = ws_ae.range("A1:J1")
    header_range.api.Font.Bold = True
    header_range.api.HorizontalAlignment = -4108  # xlCenter
    header_range.api.VerticalAlignment = -4108    # xlCenter
    header_range.color = '#d5d3af'

    # 🔷 Устанавливаем высоту строки заголовка +5 пунктов
    current_height = ws_ae.range("1:1").row_height
    ws_ae.range("1:1").row_height = current_height + 5

    # 🔷 Автоширина + 4 pt для удобства чтения
    ws_ae.range("A:J").columns.autofit()
    for col in ws_ae.range("A1:J1").columns:
        current_width = col.column_width
        col.column_width = current_width + 4

    # 🔷 Сохраняем и закрываем книгу
    wb.save()
    wb.close()

print("✅ Таблица 'акции_etf' создана.")


# === Блок заполнения таблицы акции_etf ===
#=====================================================

import xlwings as xw
import os

# === Функция форматирования ширины столбца ===
def adjust_column_width(ws, col_letter, extra_pts=3):
    ws.range(f"{col_letter}:{col_letter}").columns.autofit()
    current_width = ws.range(f"{col_letter}1").column_width
    ws.range(f"{col_letter}:{col_letter}").column_width = current_width + extra_pts

# === Блок заполнения таблицы акции_etf (столбцы A, B, C, D, E, F, G) ===

with xw.App(visible=False) as app:
    wb = app.books.open(report_file_path)
    ws_ae = wb.sheets['акции_etf']
    ws_prices = wb.sheets['текущие цены']

    # 🔷 Определяем количество строк в таблице 'текущие цены'
    last_row_prices = ws_prices.range('A' + str(ws_prices.cells.last_cell.row)).end('up').row

    # ===================== 🔷 Блок создания словаря ISIN → количество и ISIN → Баланс. цена из src_ файла =====================
    src_file_name = new_file_name.replace('отчет_', 'src_')
    src_file_path = os.path.join(folder_path, src_file_name)
    wb_src = app.books.open(src_file_path)
    ws_portfolio = wb_src.sheets['портфель']

    header_row = ws_portfolio.range('1:1').value
    isin_col_index = header_row.index('ISIN') + 1
    qty_col_index = 5  # колонка E (Количество)
    price_col_index = 8  # колонка H (Баланс. цена)

    qty_dict = {}
    price_dict = {}

    last_row_src = ws_portfolio.range('A' + str(ws_portfolio.cells.last_cell.row)).end('up').row

    for i in range(2, last_row_src + 1):
        isin = ws_portfolio.range((i, isin_col_index)).value
        qty = ws_portfolio.range((i, qty_col_index)).value
        price = ws_portfolio.range((i, price_col_index)).value
        if isin:
            qty_dict[isin] = qty
            price_dict[isin] = price

    wb_src.close()
    # ==========================================================================================

    # ===================== 🔷 Заполнение столбца A (нумерация) =====================
    for i in range(2, last_row_prices + 1):
        cell_num = ws_ae.range(f"A{i}")
        cell_num.value = int(i - 1)
        cell_num.api.NumberFormat = '# ##0'
        cell_num.api.HorizontalAlignment = -4108  # xlCenter
        cell_num.api.VerticalAlignment = -4108    # xlCenter
    adjust_column_width(ws_ae, 'A', extra_pts=3)
    # ==========================================================================

    # ===================== 🔷 Заполнение столбца B (активы) =====================
    for i in range(2, last_row_prices + 1):
        active_name = ws_prices.range(f"B{i}").value
        cell_B = ws_ae.range(f"B{i}")
        cell_B.value = active_name
        cell_B.api.HorizontalAlignment = -4131  # xlLeft
        cell_B.api.VerticalAlignment = -4108    # xlCenter
    adjust_column_width(ws_ae, 'B', extra_pts=3)
    # ==========================================================================

    # ===================== 🔷 Заполнение столбца C (тикеры) =====================
    for i in range(2, last_row_prices + 1):
        ticker = ws_prices.range(f"C{i}").value
        cell_C = ws_ae.range(f"C{i}")
        cell_C.value = ticker
        cell_C.api.HorizontalAlignment = -4108  # xlCenter
        cell_C.api.VerticalAlignment = -4108    # xlCenter
    adjust_column_width(ws_ae, 'C', extra_pts=3)
    # ==========================================================================

    # ===================== 🔷 Заполнение столбца D (количество) =====================
    for i in range(2, last_row_prices + 1):
        isin = ws_prices.range(f"A{i}").value
        qty = qty_dict.get(isin, 0)
        cell_D = ws_ae.range(f"D{i}")
        cell_D.value = int(qty) if isinstance(qty, (int, float)) else 0
        cell_D.api.NumberFormat = '# ##0'
        cell_D.api.HorizontalAlignment = -4152  # xlRight
        cell_D.api.VerticalAlignment = -4108    # xlCenter
    adjust_column_width(ws_ae, 'D', extra_pts=3)
    # ==========================================================================

    # ===================== 🔷 Заполнение столбца E (Цена входа) из src_ (Баланс. цена) =====================
    for i in range(2, last_row_prices + 1):
        isin = ws_prices.range(f"A{i}").value
        price = price_dict.get(isin, 0)
        cell_E = ws_ae.range(f"E{i}")
        if isinstance(price, (int, float)):
            cell_E.value = price
            cell_E.api.NumberFormat = '# ##0,000'
        else:
            cell_E.value = 0
            cell_E.api.NumberFormat = '# ##0,000'
        cell_E.api.HorizontalAlignment = -4152  # xlRight
        cell_E.api.VerticalAlignment = -4108    # xlCenter
    adjust_column_width(ws_ae, 'E', extra_pts=3)
    print("✅ Столбец 'Цена входа' заполнен.")
    # ==========================================================================

    # ===================== 🔷 Заполнение столбца F (Объем входа = Количество * Цена входа) =====================
    for i in range(2, last_row_prices + 1):
        qty = ws_ae.range(f"D{i}").value
        price_in = ws_ae.range(f"E{i}").value
        volume_in = qty * price_in if isinstance(qty, (int, float)) and isinstance(price_in, (int, float)) else 0
        cell_F = ws_ae.range(f"F{i}")
        cell_F.value = volume_in
        cell_F.api.NumberFormat = '# ##0,00'
        cell_F.api.HorizontalAlignment = -4152  # xlRight
        cell_F.api.VerticalAlignment = -4108    # xlCenter
    adjust_column_width(ws_ae, 'F', extra_pts=3)
    print("✅ Столбец 'Объем входа' заполнен.")
    # ==========================================================================

    # ===================== 🔷 Заполнение столбца G (Цена текущая) из листа 'текущие цены' =====================
    for i in range(2, last_row_prices + 1):
        price_current = ws_prices.range(f"F{i}").value  # колонка F на листе текущие_цены
        cell_G = ws_ae.range(f"G{i}")
        if isinstance(price_current, (int, float)):
            cell_G.value = price_current
            cell_G.api.NumberFormat = '# ##0,000'
        else:
            cell_G.value = 0
            cell_G.api.NumberFormat = '# ##0,000'
        cell_G.api.HorizontalAlignment = -4152  # xlRight
        cell_G.api.VerticalAlignment = -4108    # xlCenter
    adjust_column_width(ws_ae, 'G', extra_pts=3)
    print("✅ Столбец 'Цена текущая' заполнен.")
    # ==========================================================================

    # ===================== 🔷 Заполнение столбца H (Объем текущий = Количество * Цена текущая) =====================
    for i in range(2, last_row_prices + 1):
        qty = ws_ae.range(f"D{i}").value
        current_price = ws_ae.range(f"G{i}").value
        cell_H = ws_ae.range(f"H{i}")

        # Проверяем, что оба значения являются числами
        if isinstance(qty, (int, float)) and isinstance(current_price, (int, float)):
            volume_current = qty * current_price
        else:
            volume_current = 0

        cell_H.value = volume_current

        # Формат числа: разделение на разряды, два знака после запятой
        cell_H.api.NumberFormat = '# ##0,00'
        cell_H.api.HorizontalAlignment = -4152  # xlRight
        cell_H.api.VerticalAlignment = -4108    # xlCenter

    adjust_column_width(ws_ae, 'H', extra_pts=3)
    print("✅ Столбец 'Объем текущий' (H) успешно заполнен.")
    # ==========================================================================

    # ===================== 🔷 Заполнение столбца I (Разница USD = Объем текущий - Объем входа) =====================
    for i in range(2, last_row_prices + 1):
        volume_current = ws_ae.range(f"H{i}").value
        volume_entry = ws_ae.range(f"F{i}").value
        cell_I = ws_ae.range(f"I{i}")

        # Проверяем, что оба значения являются числами
        if isinstance(volume_current, (int, float)) and isinstance(volume_entry, (int, float)):
            difference_usd = volume_current - volume_entry
            cell_I.value = difference_usd
            # Применяем формат только если значение – число
            cell_I.api.NumberFormat = '# ##0,00'
        else:
            cell_I.value = 0
            cell_I.api.NumberFormat = '# ##0,00'

        cell_I.api.HorizontalAlignment = -4152  # xlRight
        cell_I.api.VerticalAlignment = -4108    # xlCenter

    adjust_column_width(ws_ae, 'I', extra_pts=3)
    print("✅ Столбец 'Разница USD' (I) успешно заполнен.")
    # ==========================================================================

    # ===================== 🔷 Заполнение столбца J (Разница % = Разница USD / Объем входа) =====================
    for i in range(2, last_row_prices + 1):
        diff_usd = ws_ae.range(f"I{i}").value
        volume_entry = ws_ae.range(f"F{i}").value
        cell_J = ws_ae.range(f"J{i}")

        # Проверяем, что оба значения являются числами и объем входа не равен нулю
        if isinstance(diff_usd, (int, float)) and isinstance(volume_entry, (int, float)) and volume_entry != 0:
            diff_percent = diff_usd / volume_entry
            cell_J.value = diff_percent
            # Применяем формат процента: два знака после запятой и разделение разрядов
            cell_J.api.NumberFormat = '0,00%'
        else:
            cell_J.value = 0
            cell_J.api.NumberFormat = '0,00%'

        cell_J.api.HorizontalAlignment = -4152  # xlRight
        cell_J.api.VerticalAlignment = -4108    # xlCenter

    adjust_column_width(ws_ae, 'J', extra_pts=3)
    print("✅ Столбец 'Разница %' (J) успешно заполнен.")
    # ==========================================================================

    # ===================== 🔷 Формирование итогов в таблице акции_etf =====================
    # Определяем строку после последней записи
    total_row = last_row_prices + 1

    # Вписываем текст "ИТОГО" в столбец B
    cell_total_label = ws_ae.range(f"B{total_row}")
    cell_total_label.value = "ИТОГО"
    cell_total_label.api.Font.Bold = True
    cell_total_label.api.HorizontalAlignment = -4131  # xlLeft
    cell_total_label.api.VerticalAlignment = -4108    # xlCenter

    # Формируем формулы для суммы в столбцах F, H, I
    columns_sum = {'F': 'Объем входа', 'H': 'Объем текущий', 'I': 'Разница USD'}
    for col, name in columns_sum.items():
        cell = ws_ae.range(f"{col}{total_row}")
        formula = f"=SUM({col}2:{col}{last_row_prices})"
        cell.api.Formula = formula

        # Применяем формат '# ##0,00' и жирный шрифт
        cell.api.NumberFormat = '# ##0,00'
        cell.api.Font.Bold = True
        cell.api.HorizontalAlignment = -4152  # xlRight
        cell.api.VerticalAlignment = -4108    # xlCenter

    # 🔷 Расчет Разница % в столбце J как Разница USD / Объем входа
    cell_J = ws_ae.range(f"J{total_row}")
    formula_J = f"=I{total_row}/F{total_row}"
    cell_J.api.Formula = formula_J

    # Применяем формат процента с двумя знаками после запятой и жирный шрифт
    cell_J.api.NumberFormat = '0,00%'
    cell_J.api.Font.Bold = True
    cell_J.api.HorizontalAlignment = -4152  # xlRight
    cell_J.api.VerticalAlignment = -4108    # xlCenter

    print("✅ Итоговые данные по акциям и ETF сформированы.")
    # ==========================================================================

    # ===================== 🔷 Цветовое оформление результатов таблицы акции_etf (столбцы I и J) =====================
    from xlwings.utils import rgb_to_int

    # Определяем последнюю строку с данными и строку итогов
    total_row = last_row_prices + 1

    # Список колонок для окрашивания
    color_columns = ['I', 'J']

    for col in color_columns:
        for i in range(2, total_row + 1):
            cell = ws_ae.range(f"{col}{i}")
            value = cell.value

            # Приведение к числу, если формула возвращает результат
            try:
                val = float(value)
            except (TypeError, ValueError):
                val = None

            if val is not None:
                if val > 0:
                    # Светло-зеленый (например, RGB 198, 239, 206)
                    cell.color = (198, 239, 206)
                elif val < 0:
                    # Светло-красный (например, RGB 255, 199, 206)
                    cell.color = (255, 199, 206)
                else:
                    # Светло-желтый (например, RGB 255, 235, 156)
                    cell.color = (255, 235, 156)
            else:
                # Если значение не является числом, оставляем без заливки
                pass

    print("✅ Цветовое оформление результатов таблицы акции_etf завершено.")
    # ==========================================================================




    # 🔷 Сохраняем и закрываем книгу
    wb.save()
    wb.close()

print("✅ Таблица 'акции_etf' обновлена и заполнена имеющимися данными.")
