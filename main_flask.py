import multiprocessing
import os
import time
import sqlite3
import json
from fastapi import FastAPI, Query
from fastapi.responses import JSONResponse
import uvicorn

# База данных
DB_PATH = "usb_log.db"

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS usb_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            event_type TEXT,
            device TEXT,
            owner TEXT,
            file_changes TEXT
        )''')
        conn.commit()

# Функция мониторинга USB
import psutil  # Используем psutil для обнаружения новых устройств

def monitor_usb():
    known_devices = set()
    while True:
        devices = {disk.device for disk in psutil.disk_partitions(all=True) if 'removable' in disk.opts}
        added = devices - known_devices
        removed = known_devices - devices
        
        for device in added:
            log_event("connect", device, "Unknown", "{} подключен".format(device))
        for device in removed:
            log_event("disconnect", device, "Unknown", "{} отключен".format(device))
        
        known_devices = devices
        time.sleep(2)

# Логирование изменений файлов
import hashlib

def hash_file(filepath):
    hasher = hashlib.md5()
    with open(filepath, 'rb') as f:
        hasher.update(f.read())
    return hasher.hexdigest()

def monitor_files(device):
    file_hashes = {}
    while os.path.exists(device):
        current_files = {f: hash_file(os.path.join(device, f)) for f in os.listdir(device)}
        
        for f, h in current_files.items():
            if f not in file_hashes:
                log_event("file_added", device, "Unknown", f"Файл добавлен: {f}")
            elif file_hashes[f] != h:
                log_event("file_modified", device, "Unknown", f"Файл изменён: {f}")
        
        for f in file_hashes.keys() - current_files.keys():
            log_event("file_deleted", device, "Unknown", f"Файл удалён: {f}")
        
        file_hashes = current_files
        time.sleep(2)

# Логирование событий

def log_event(event_type, device, owner, file_changes):
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute("INSERT INTO usb_events (timestamp, event_type, device, owner, file_changes) VALUES (?, ?, ?, ?, ?)",
                       (timestamp, event_type, device, owner, file_changes))
        conn.commit()

# Веб-интерфейс
app = FastAPI()

@app.get("/")
def index():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM usb_events ORDER BY timestamp DESC")
    events = cursor.fetchall()
    conn.close()
    return JSONResponse(content={"events": events})

@app.get("/filter")
def filter_events(event_type: str = Query("", alias="event_type")):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    query = "SELECT * FROM usb_events WHERE event_type LIKE ? ORDER BY timestamp DESC"
    cursor.execute(query, (f"%{event_type}%",))
    events = cursor.fetchall()
    conn.close()
    return JSONResponse(content={"events": events})

if __name__ == "__main__":
    init_db()
    usb_monitor = multiprocessing.Process(target=monitor_usb)
    usb_monitor.start()
    uvicorn.run(app, host="0.0.0.0", port=5000)

