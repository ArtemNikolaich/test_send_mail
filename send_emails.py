import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib


TABLE_NAME = 'Название вашей таблицы'
CREDENTIALS_FILE = 'credentials.json'
CONFIG_FILE = '1.txt'


def read_google_sheet():
    # Подключаемся к Google Sheets
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        CREDENTIALS_FILE,
        scope
    )
    client = gspread.authorize(credentials)

    # Открываем таблицу
    sheet = client.open(TABLE_NAME).sheet1

    # Получаем данные по порядку из таблицы
    data = sheet.get_all_values()

    if not data:
        raise ValueError('Таблица пустая.')

    # Извлекаем адреса и сообщения
    email_addresses = []
    messages = []
    for row in data[1:]:  # Пропускаем заголовки (первая строка)
        email_addresses.append(row[0])  # Первый столбец
        messages.append(row[1])  # Второй столбец

    return email_addresses, messages, client


def read_config_file():
    with open(CONFIG_FILE, 'r', encoding='utf-8') as file:
        lines = file.readlines()
        login = lines[0].strip()
        password = lines[1].strip()

    return login, password


def send_emails(email_addresses, messages, login, password, sheet):
    if not email_addresses:
        raise ValueError('Нет адресов для отправки сообщений.')

    smtp_server = 'smtp.gmail.com'
    port = 587

    try:
        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls()
            server.login(login, password)

            for i, email in enumerate(email_addresses):
                if not messages[i]:
                    continue

                try:
                    server.sendmail(login, email, messages[i].encode('utf-8'))
                except Exception as e:
                    sheet.update_cell(i + 2, 3, str(e))
    except Exception as e:
        raise ValueError(f'Ошибка соединения с SMTP сервером: {e}')


email_addresses, messages, client = read_google_sheet()
login, password = read_config_file()
sheet = client.open(TABLE_NAME).sheet1
send_emails(email_addresses, messages, login, password, sheet)
