import cv2
from pyzbar.pyzbar import decode
import sqlite3
from datetime import datetime
import openpyxl
import winsound  # Use for beep sound on Windows
import os
import time  # Import time module for sleep function

# Generate a unique database name based on the current timestamp
db_name = f'barcode_scans_{datetime.now().strftime("%Y%m%d_%H%M%S")}.db'

def create_database():
    try:
        conn = sqlite3.connect(db_name)
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS scans
                     (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                      date TEXT,
                      barcode TEXT, 
                      intime TEXT,
                      outtime TEXT)''')
        conn.commit()
        conn.close()
        print("Database and table created.")
    except Exception as e:
        print(f"Error creating database: {e}")

def insert_or_update_scan(barcode):
    try:
        conn = sqlite3.connect(db_name)
        c = conn.cursor()
        current_date = datetime.now().strftime('%Y-%m-%d')
        current_time = datetime.now().strftime('%H:%M:%S')

        # Check if the barcode already exists for the current date
        c.execute("SELECT id, intime, outtime FROM scans WHERE barcode = ? AND date = ? ORDER BY id DESC LIMIT 1", (barcode, current_date))
        result = c.fetchone()

        if result and result[2] is None:  # If barcode exists and outtime is None, update outtime
            c.execute("UPDATE scans SET outtime = ? WHERE id = ?", (current_time, result[0]))
            print(f"Updated scan: Barcode = {barcode}, Out Time = {current_time}")
        else:  # Otherwise, insert a new record with date and intime
            c.execute("INSERT INTO scans (date, barcode, intime) VALUES (?, ?, ?)", (current_date, barcode, current_time))
            print(f"Inserted scan: Date = {current_date}, Barcode = {barcode}, In Time = {current_time}")

        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Error inserting or updating scan: {e}")

def beep():
    try:
        winsound.Beep(1000, 500)  # Frequency = 1000 Hz, Duration = 500 ms
    except ImportError:
        print("\nBeep sound not available on this platform.")

def scan_barcodes():
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("Error: Could not open video capture.")
        return

    try:
        while True:
            ret, frame = cap.read()
            if not ret:
                print("Failed to capture image.")
                break

            barcodes = decode(frame)
            if barcodes:
                for barcode in barcodes:
                    barcode_data = barcode.data.decode('utf-8')
                    print(f"Detected Barcode: {barcode_data}")
                    insert_or_update_scan(barcode_data)
                    beep()  # Produce beep sound

                    # Draw rectangle around the barcode and display the barcode data
                    (x, y, w, h) = barcode.rect
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
                    cv2.putText(frame, barcode_data, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)

                    # Wait for 10 seconds after a successful scan
                    time.sleep(10)

            cv2.imshow('Barcode Scanner', frame)
            
            if cv2.waitKey(1) & 0xFF == ord('q'):
                print("Exiting program.")
                break

    except KeyboardInterrupt:
        print("Program interrupted by user.")
    finally:
        cap.release()
        cv2.destroyAllWindows()
        print("Video capture and window closed.")

def fetch_scans():
    try:
        conn = sqlite3.connect(db_name)
        c = conn.cursor()
        c.execute("SELECT * FROM scans")
        rows = c.fetchall()
        conn.close()
        return rows
    except Exception as e:
        print(f"Error fetching scans: {e}")
        return []

def save_to_excel(scans):
    filename = f'scans_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Scans"

    ws.append(["Serial Number", "Date", "Roll Number", "In Time", "Out Time"])

    for scan in scans:
        ws.append(scan)

    wb.save(filename)
    print(f"Data saved to {filename}")

if __name__ == '__main__':
    create_database()
    scan_barcodes()
    scans = fetch_scans()
    save_to_excel(scans)
