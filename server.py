import PySimpleGUI as sg
import threading
import json
import socket
import platform
def extract_user_data(username):
    with open("data/user_data/" + username + ".txt", "r", encoding="utf-8") as file:
        data = []
        data = file.read().split('\n')
    new_data = []
    for i in data:
        n = i.find(' [')
        new_data.append([i[:n], i[n+1:]])

    return new_data
def start_server(window):
    host = '127.0.0.1'
    port = 12345

    server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server_socket.bind((host, port))
    server_socket.listen(5)

    print(f"Сервер слушает на {host}:{port}")

    try:
        while True:
            client_socket, addr = server_socket.accept()
            print(f"Получено соединение от {addr}")

            with client_socket:
                try:
                    data = receive_data(client_socket)
                    print(f"Полученные данные: ")
                    print(data)
                    print(json.dumps(data, indent=2, ensure_ascii=False))

                    # Отправить успешный ответ клиенту
                    send_response(client_socket, "Данные успешно получены")

                    # Обновление графического интерфейса
                    window.write_event_value('-UPDATE-', json.dumps(data, indent=2, ensure_ascii=False))
                except Exception as e:
                    print(f"Ошибка при обработке данных: {str(e)}")
    except KeyboardInterrupt:
        print("Сервер завершает работу.")
    finally:
        server_socket.close()


def receive_data(client_socket):
    data = b""
    while True:
        chunk = client_socket.recv(1024)
        if not chunk:
            break
        data += chunk

    return json.loads(data.decode('cp1251'))


def send_response(client_socket, response):
    client_socket.send(response.encode('utf-8'))


def start_check(window):
    layout1 = [[sg.Text("Отправлен запрос на начало проверки")],
               [sg.Text("Ожидание ответа...")]];
    window1 = sg.Window("Запрос на проверку", layout1)
    while True:
        event, value = window1.read()
        if event in (sg.WIN_CLOSED, '-EXIT-'):
            window1.close()
            break;


def update_table(window, data):
    if isinstance(data, (list, dict)):
        # Очистить существующую таблицу
        window['-TABLE-'].update(values=[])
        # Выровнять данные для отображения в таблице
        flattened_data = []
        if isinstance(data, list):
            for i in enumerate(data):
                j = i[0]
                k = i[1]
                n = k.find(' [')
                keywords = k[n:]
                path = k[:n]
                flattened_data.append([str(j), str(path), str(keywords)])

        # Обновить таблицу новыми данными
        window['-TABLE-'].update(values=flattened_data)


def create_table(username):

    layout = [
        [sg.Table(values=extract_user_data(username), headings=['Путь', 'Ключи'], display_row_numbers=False, auto_size_columns=False,
                  justification='left', key='-USER_TABLE-', enable_events=True, bind_return_key=True,
                  vertical_scroll_only=False, num_rows=30, def_col_width=20, max_col_width=100,
                  col_widths=[120, 50])],
        [sg.Button("Завершить сервер",  key='-EXIT-', auto_size_button=True)],
        [sg.Button("Начать проверку", key='-START_CHECK-', auto_size_button=True)]
    ]
    table_window = sg.Window(username, layout, resizable=True, auto_size_buttons=True)
    while True:
        event, values = table_window.read()

        if event in (sg.WIN_CLOSED, '-EXIT-'):
            table_window.close()
            break

def create_main_table():
    with open("data/users.csv", "r") as file:
        data = []
        data = file.read().split('\n')
        data = [x.split(';') for x in data]
        cnt = len(data)
        header = ["Пользователь", "Дата проверки", "Найдено файлов"]
        rows = [[data[i][0], data[i][1], data[i][2]] for i in range(len(data)) if len(data[i]) >= 3]

        layout = [
            [sg.Table(values=rows, headings=header, display_row_numbers=False, auto_size_columns=False,
                      justification='left', enable_events=True, key='-TABLE-',
                      select_mode=sg.TABLE_SELECT_MODE_EXTENDED, col_widths=[30, 30, 30]),
             sg.Button("Просмотр", key="-OPEN_USER-")]
        ]

        window = sg.Window("Сервер", layout, resizable=True, auto_size_buttons=True, auto_size_text=True)

        while True:
            event, values = window.read()

            if event in (sg.WIN_CLOSED, '-EXIT-'):
                window.close()
                break
            elif event == '-OPEN_USER-':
                selected_rows = values['-TABLE-']
                print(selected_rows)
                if (selected_rows == []) or (len(selected_rows) > 1):
                    continue
                #print(data[selected_rows[0]])
                create_table(data[selected_rows[0]][0])

def main():
    username = 'K.txt'
    layout = [[sg.Button("Пользователи", key='-USERS-')],
              [sg.Button("Назначить всем проверку", key='-CHECK_ALL-')],
              [sg.Button("Уведомления", key="-NOTIFICATIONS'")],
              [sg.Button("Выход", key='-EXIT')]
              ]
    window = sg.Window("Сервер", layout, resizable=True, auto_size_buttons=True)
    server_thread = threading.Thread(target=start_server, args=(window,))
    server_thread.start()

    while True:
        event, values = window.read()

        if event in (sg.WIN_CLOSED, '-EXIT-'):
            # Подтверждение выхода с помощью всплывающего окна
            if sg.popup_yes_no("Вы уверены, что хотите завершить сервер?", title="Подтверждение выхода") == 'Yes':
                window.close()  # Закрываем окно PySimpleGUI
                break

        elif event == '-UPDATE-':
            data = json.loads(values[event])
            update_table(window, data)
        elif event == '-USERS-':
            create_main_table()

    server_thread.join()  # Дождаться завершения серверного потока


if __name__ == "__main__":
    if platform.system() == 'Windows':
        sg.theme('DarkBlue')
        sg.SetOptions(font=('Courier New', 10, 'bold'))
    else:
        sg.theme('DarkBlue3')
    main()
