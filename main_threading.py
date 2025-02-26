import wmi
import multiprocessing as mp
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from time import sleep

process = True

def _set(obj, attribute, value):
    """Helper function to add an attribute directly into the instance
    dictionary, bypassing possible `__getattr__` calls

    :param obj: Any python object
    :param attribute: String containing attribute name
    :param value: Any python object
    """
    obj.__dict__[attribute] = value


def get_removable_drives(queue):
    """Функция для получения съемных дисков и помещения их в очередь."""
    while process:
        try:
            c = wmi.WMI()
            removable_drives = []

            for drive in c.Win32_DiskDrive():
                if drive.MediaType == "Removable Media" or drive.MediaType == "External hard disk media":
                    for partition in drive.associators("Win32_DiskDriveToDiskPartition"):
                        for logical_disk in partition.associators("Win32_LogicalDiskToPartition"):
                            drive_info = {
                                "BytesPerSector": drive.BytesPerSector,
                                "Capabilities": drive.Capabilities,
                                "CapabilityDescriptions": drive.CapabilityDescriptions,
                                "Caption": drive.Caption,
                                "ConfigManagerErrorCode": drive.ConfigManagerErrorCode,
                                "ConfigManagerUserConfig": drive.ConfigManagerUserConfig,
                                "CreationClassName": drive.CreationClassName,
                                "Description": drive.Description,
                                "DeviceID": drive.DeviceID,
                                "FirmwareRevision": drive.FirmwareRevision,
                                "Index": drive.Index,
                                "InterfaceType": drive.InterfaceType,
                                "Manufacturer": drive.Manufacturer,
                                "MediaLoaded": drive.MediaLoaded,
                                "MediaType": drive.MediaType,
                                "Model": drive.Model,
                                "Name": drive.Name,
                                "Partitions": drive.Partitions,
                                "PNPDeviceID": drive.PNPDeviceID,
                                "SCSIBus": drive.SCSIBus,
                                "SCSILogicalUnit": drive.SCSILogicalUnit,
                                "SCSIPort": drive.SCSIPort,
                                "SCSITargetId": drive.SCSITargetId,
                                "SectorsPerTrack": drive.SectorsPerTrack,
                                "SerialNumber": drive.SerialNumber,
                                "Signature": drive.Signature,
                                "Size": drive.Size,
                                "Status": drive.Status,
                                "SystemCreationClassName": drive.SystemCreationClassName,
                                "SystemName": drive.SystemName,
                                "TotalCylinders": drive.TotalCylinders,
                                "TotalHeads": drive.TotalHeads,
                                "TotalSectors": drive.TotalSectors,
                                "TotalTracks": drive.TotalTracks,
                                "TracksPerCylinder": drive.TracksPerCylinder,
                                "Letter": logical_disk.DeviceID,
                                "Owner": None  # Изначально владелец неизвестен
                            }
                            removable_drives.append(drive_info)

            queue.put(removable_drives)
        except Exception as e:
            print(f"An error occurred: {e}")
            queue.put([])
        sleep(2)  # Обновление каждые 2 секунды


def display_devices(queue):
    """Функция для отображения подключенных устройств с использованием Rich."""
    console = Console()

    while process:
        removable_drives = queue.get()
        console.clear()

        if removable_drives:
            table = Table(title="Connected Removable Drives", show_header=True, header_style="bold magenta")
            table.add_column("Letter", style="cyan")
            table.add_column("Caption", style="green")
            table.add_column("Size", style="yellow")
            table.add_column("Owner", style="bold magenta")
            table.add_column("Media Type", style="blue")
            table.add_column("Status", style="red")

            for drive in removable_drives:
                table.add_row(
                    drive["Letter"],
                    drive["Caption"],
                    str(drive["Size"]),
                    drive["Owner"],
                    drive["MediaType"],
                    drive["Status"]
                )

            console.print(Panel.fit(table, title="Removable Drives", border_style="bold blue"))
        else:
            console.print(Panel.fit("No removable drives found.", title="Removable Drives", border_style="bold red"))

main_loop = True

if __name__ == '__main__':
    queue = mp.Queue()

    # Запуск процесса для получения съемных дисков
    processing_drives = mp.Process(target=get_removable_drives, args=(queue,))
    processing_drives.start()

    # Запуск процесса для отображения устройств
    display_process = mp.Process(target=display_devices, args=(queue,))
    display_process.start()

    try:
        # Основной процесс остается активным
        while main_loop:
            removable_drives = queue.get()

            # Проверяем, есть ли устройства без владельца
            unknown_owner_devices = [drive for drive in removable_drives if drive["Owner"] is None]

            if unknown_owner_devices:
                # Если есть устройства без владельца, запрашиваем ввод
                main_loop = False
                
                for drive in unknown_owner_devices:
                    owner = input(f"Введите имя владельца для устройства {drive['Letter']}: ")
                    drive["Owner"] = owner  # Сохраняем введенное значение
                    break
            else:
                True
            # Отправляем обновленные данные обратно в очередь
            queue.put(removable_drives)

            sleep(1)
    except KeyboardInterrupt:
        # Остановка процессов при завершении
        process = False
        processing_drives.terminate()
        display_process.terminate()
        processing_drives.join()
        display_process.join()