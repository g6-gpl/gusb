import wmi
import time
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
import json
from datetime import datetime
import pandas as pd
import keyboard  
import os


console = Console()

def save_data(data, filename):
    with open(filename, "w") as file:
        json.dump(data, file, indent=4)

def load_data(filename):
    try:
        with open(filename, "r") as file:
            return json.load(file)
    except FileNotFoundError:
        return {}


def _set(obj, attribute, value):
    """Helper function to add an attribute directly into the instance
    dictionary, bypassing possible `__getattr__` calls

    :param obj: Any python object
    :param attribute: String containing attribute name
    :param value: Any python object
    """
    obj.__dict__[attribute] = value
    
    
def get_removable_drives():
    c = wmi.WMI()
    removable_drives = []
        
    for drive in c.Win32_DiskDrive():
        
        if drive.MediaType == "Removable Media" or drive.MediaType == "External hard disk media":
            for partition in drive.associators("Win32_DiskDriveToDiskPartition"):
                for logical_disk in partition.associators("Win32_LogicalDiskToPartition"):
                    _set(drive,"Letter",logical_disk.DeviceID)
                    removable_drives.append(drive) 

    return removable_drives

def get_drives_hash(drives):
    return hash(tuple((drive.Model, drive.SerialNumber) for drive in drives))

owners = load_data("owners.json")
history = load_data("history.json")

def get_owner(serial_number):
    if serial_number not in owners:
        owner = console.input(f"[yellow]Флеш-карта с серийным номером {serial_number} подключена впервые. Введите имя владельца: ")
        owners[serial_number] = owner
        save_data(owners, "owners.json")
    return owners[serial_number]

def update_history(serial_number, event, file_changes=None):
    if serial_number not in history:
        history[serial_number] = []
    
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    history[serial_number].append({
        "event": event,
        "timestamp": timestamp,
        "file_changes": file_changes  # Добавляем информацию о файлах
    })
    save_data(history, "history.json")


def get_user_by_serial_number(serial_number):
    users = load_data("owners.json")
    try:
        return users[serial_number]
    except Exception:
        return None

def scan_files_on_drive(drive_letter):
    """Сканирует файлы на съемном носителе."""
    files = []
    for root, _, filenames in os.walk(drive_letter):
        for filename in filenames:
            files.append(os.path.join(root, filename))
    return files


def log_file_changes(serial_number, drive_letter):
    """Логирует изменения файлов на съемном носителе."""
    log_file = f"file_changes_{serial_number}.log"
    current_files = set(scan_files_on_drive(drive_letter))

    if not os.path.exists(log_file):
        # Если файл лога не существует, создаем его и записываем текущие файлы
        with open(log_file, "w", encoding="utf-8") as file:
            for filepath in current_files:
                file.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Добавлен: {filepath}\n")
    else:
        # Читаем предыдущий список файлов
        previous_files = set()
        with open(log_file, "r", encoding="utf-8") as file:
            for line in file:
                if "Добавлен:" in line:
                    previous_files.add(line.split("Добавлен: ")[1].strip())

        # Находим новые, измененные и удаленные файлы
        new_files = current_files - previous_files
        removed_files = previous_files - current_files

        # Записываем новые файлы в лог
        with open(log_file, "a", encoding="utf-8") as file:
            for filepath in new_files:
                file.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Добавлен: {filepath}\n")
            for filepath in removed_files:
                file.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Удален: {filepath}\n")

        # Возвращаем изменения файлов
        return {
            "new_files": list(new_files),
            "removed_files": list(removed_files)
        }

def export_to_timeline_html():
    # Создаем HTML-файл с временной линией
    html_content = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Timeline of USB Connections</title>
        <style>
            /* Ваши стили для timeline */
        </style>
    </head>
    <body>
        <div class="timeline">
    """

    # Добавляем события в timeline
    for serial_number, events in history.items():
        owner = get_user_by_serial_number(serial_number)
        for event in events:
            timestamp = event["timestamp"]
            event_type = event["event"]
            file_changes = event.get("file_changes", {})
            html_content += f"""
            <div class="container {'left' if event_type == 'подключен' else 'right'}">
                <div class="content">
                    <h2>{timestamp}</h2>
                    <p><strong>Владелец:</strong> {owner}</p>
                    <p><strong>Серийный номер:</strong> {serial_number}</p>
                    <p><strong>Событие:</strong> {event_type}</p>
                    <p><strong>Изменения файлов:</strong></p>
                    <ul>
                        {"".join(f"<li>Добавлен: {file}</li>" for file in file_changes.get("new_files", []))}
                        {"".join(f"<li>Удален: {file}</li>" for file in file_changes.get("removed_files", []))}
                    </ul>
                </div>
            </div>
            """

    html_content += """
        </div>
    </body>
    </html>
    """

    # Сохраняем HTML-файл
    with open("timeline.html", "w", encoding="utf-8") as file:
        file.write(html_content)

    console.print("[green]Данные успешно экспортированы в файл 'timeline.html'.")
    
def export_to_excel():
    owners_df = pd.DataFrame(list(owners.items()), columns=["Серийный номер", "Владелец"])

    history_data = []
    for serial_number, events in history.items():
        for event in events:
            file_changes = event.get("file_changes", {})
            history_data.append({
                "Владелец": get_user_by_serial_number(serial_number),
                "Серийный номер": serial_number,
                "Событие": event["event"],
                "Время": event["timestamp"],
                "Добавленные файлы": ", ".join(file_changes.get("new_files", [])),
                "Удаленные файлы": ", ".join(file_changes.get("removed_files", []))
            })
    history_df = pd.DataFrame(history_data)

    with pd.ExcelWriter("подключения_и_пользователи.xlsx") as writer:
        owners_df.to_excel(writer, sheet_name="Владельцы", index=False)
        history_df.to_excel(writer, sheet_name="История подключений", index=False)

    console.print("[green]Данные успешно экспортированы в файл 'подключения_и_пользователи.xlsx'.")

    with pd.ExcelWriter("подключения_и_пользователи.xlsx") as writer:
        owners_df.to_excel(writer, sheet_name="Владельцы", index=False)
        history_df.to_excel(writer, sheet_name="История подключений", index=False)

    console.print("[green]Данные успешно экспортированы в файл 'подключения_и_пользователи.xlsx'.")

if __name__ == "__main__":
    try:
        previous_hash = None
        connected_drives = set()



        while True:
            removable_drives = get_removable_drives()
            current_hash = get_drives_hash(removable_drives)
            if current_hash != previous_hash:
                console.clear()
                
                current_serials = {drive.SerialNumber for drive in removable_drives}
                new_drives = current_serials - connected_drives
                removed_drives = connected_drives - current_serials

                for serial in new_drives:
                    for removable_drive in removable_drives:
                        if removable_drive.SerialNumber == serial:
                            file_changes = log_file_changes(serial, removable_drive.Letter)
                            update_history(serial, "подключен", file_changes)

                for serial in removed_drives:
                    update_history(serial, "отключен")

                connected_drives = current_serials

                if removable_drives:
                            # Оработчик нажатия клавиши "1"
                    keyboard.on_press_key("1", lambda _: export_to_excel())
                    keyboard.on_press_key("2", lambda _: export_to_timeline_html())

                    console.print("[cyan]Нажмите '1' для экспорта данных в Excel.")
                    console.print("[cyan]Нажмите '2' для экспорта данных в timeline.html.")
                    
                    console.print("[cyan]Нажмите Ctrl+C для завершения программы.")
                    table = Table(title="ПОДКЛЮЧЕННЫЕ СЪЕМНЫЕ НОСИТЕЛИ")
                    table.add_column("Модель", justify="left")
                    table.add_column("Серийный номер", justify="left")
                    table.add_column("Владелец", justify="left")
                    table.add_column("Изменения файлов", justify="left")
                    
                    for drive in removable_drives:
                        owner = get_owner(drive.SerialNumber)
                        try:
                            # Получаем историю для текущего устройства
                            device_history = history.get(drive.SerialNumber, [])

                            # Проверяем, есть ли история для устройства
                            if device_history:
                                # Берем последнее событие и извлекаем изменения файлов
                                file_changes = device_history[-1].get("file_changes", {})
                                file_changes_str = "\n".join(
                                    [f"Добавлен: {file}" for file in file_changes.get("new_files", [])] +
                                    [f"Удален: {file}" for file in file_changes.get("removed_files", [])]
                                )
                            else:
                                # Если история пуста, указываем, что изменений нет
                                file_changes_str = "Нет изменений файлов"

                            # Добавляем строку в таблицу
                            table.add_row(drive.Model, drive.SerialNumber, owner, file_changes_str)
                        except Exception as e:
                            # Логируем ошибку, если что-то пошло не так
                            console.print(f"[red]Ошибка при обработке устройства {drive.SerialNumber}: {e}")
                        

                    console.print(table)
                else:
                    console.print(Panel("Съемные носители не найдены.", style="red"))

                previous_hash = current_hash

            time.sleep(0.3)
    except KeyboardInterrupt:
        console.print("[red]Программа завершена.")
    finally:
        keyboard.unhook_all()