import asyncio
import threading
from bleak import BleakClient
import matplotlib.pyplot as plt
from queue import Queue
import time
import struct
import numpy as np
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import datetime

# Device address and characteristic UUID
DEVICE_ADDRESS = "EB:6F:FC:C9:7B:2F"
FSM_DATA_CHAR_UUID = "c042d543-3929-48c9-af11-cf252f5ce7f4"

# Global variables for data
time_data = []
array_a_data = [[] for _ in range(21)]  # 6 wavelengths * 3 PDs + 1 LED OFF * 3 PDs
array_b_data = [[] for _ in range(21)]  # 6 wavelengths * 3 PDs + 1 LED OFF * 3 PDs
data_queue = Queue()
start_time = None
running = True

# Set up the plots
fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 10))
lines_a = [ax1.plot([], [], label=f'W{i//3+1} PD{i%3+1}')[0] for i in range(18)] + \
          [ax1.plot([], [], label=f'LED OFF PD{i+1}', linestyle='--')[0] for i in range(3)]
lines_b = [ax2.plot([], [], label=f'W{i//3+1} PD{i%3+1}')[0] for i in range(18)] + \
          [ax2.plot([], [], label=f'LED OFF PD{i+1}', linestyle='--')[0] for i in range(3)]

for ax in (ax1, ax2):
    ax.set_xlabel('Time (s)')
    ax.set_ylabel('Voltage (V)')
    ax.grid(True)
    ax.legend()
    ax.set_ylim(0.05, 3)  # Set y-axis limits from 0.05V to 3V

ax1.set_title('Array A Data')
ax2.set_title('Array B Data')

def parse_fsm_data(data):
    # Parse time data (bytes 1-6)
    time_min, time_us = struct.unpack_from('<HI', data, 0)
    total_time = time_min * 60 + time_us / 1_000_000

    # Function to parse PD values (28 unsigned integers)
    def parse_array_data(start_index):
        values = struct.unpack_from('<28I', data, start_index)
        return np.array(values) / 1e9  # Convert to volts

    # Parse Array A data (bytes 9-93)
    array_a_data = parse_array_data(9)

    # Parse Array B data (bytes 96-180)
    array_b_data = parse_array_data(96)

    return total_time, array_a_data, array_b_data

def generate_headers():
    headers = ["Time (s)"]
    for array in ["A", "B"]:
        for w in range(1, 7):
            for pd in range(1, 4):
                headers.append(f"Array {array} W{w} PD{pd}")
        for pd in range(1, 4):
            headers.append(f"Array {array} LED OFF PD{pd}")
    return headers

def notification_handler(sender, data):
    global start_time
    if start_time is None:
        start_time = time.time()
    
    try:
        total_time, array_a_data, array_b_data = parse_fsm_data(data)
        timestamp = time.time() - start_time
        data_queue.put((timestamp, total_time, array_a_data, array_b_data))
        print(f"Time: {timestamp:.2f}s, Data received and queued")
    except Exception as e:
        print(f"Error parsing data: {e}")

def update_plot():
    global running
    while running:
        while not data_queue.empty():
            timestamp, total_time, a_data, b_data = data_queue.get()
            time_data.append(timestamp)
            for i in range(21):
                array_a_data[i].append(a_data[i])
                array_b_data[i].append(b_data[i])
        
        if time_data:
            for ax in (ax1, ax2):
                ax.set_xlim(0, max(30, time_data[-1]))
            
            for i in range(21):
                lines_a[i].set_data(time_data, array_a_data[i])
                lines_b[i].set_data(time_data, array_b_data[i])
        
        fig.canvas.draw()
        fig.canvas.flush_events()
        plt.pause(0.1)
        
        if not plt.get_fignums():
            running = False
            break

def save_to_excel(data, filename=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "FSM Data"

    # Write headers
    headers = generate_headers()
    ws.append(headers)

    # Write data
    for row in data:
        ws.append(row)

    # Set column widths
    for col in range(1, 58):  # 57 columns in total
        ws.column_dimensions[get_column_letter(col)].width = 15

    # Set number format for data columns
    for col in range(1, 58):
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=col).number_format = '0.000000'

    # Generate filename with timestamp if not provided
    if filename is None:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"fsm_data_{timestamp}.xlsx"

    # Save the workbook
    wb.save(filename)
    print(f"Data saved to {filename}")

async def run_ble_client(address):
    global running
    async with BleakClient(address) as client:
        try:
            print(f"Connected: {client.is_connected}")
            await client.start_notify(FSM_DATA_CHAR_UUID, notification_handler)
            print("FSM Data Monitoring Started. Close the plot window to stop.")
            
            while running:
                await asyncio.sleep(1)
        
        except Exception as e:
            print(f"Error: {e}")
        finally:
            try:
                await client.stop_notify(FSM_DATA_CHAR_UUID)
            except Exception as e:
                print(f"Error stopping notifications: {e}")
            print("FSM Data Monitoring Stopped.")

def run_asyncio_loop(loop):
    asyncio.set_event_loop(loop)
    loop.run_until_complete(run_ble_client(DEVICE_ADDRESS))

def main():
    global running
    loop = asyncio.new_event_loop()
    ble_thread = threading.Thread(target=run_asyncio_loop, args=(loop,))
    ble_thread.start()

    try:
        plt.ion()
        update_plot()
        plt.ioff()
    except Exception as e:
        print(f"Error in main loop: {e}")
    finally:
        running = False
        ble_thread.join()
        collected_data = []
        for i in range(len(time_data)):
            row = [time_data[i]] + [array_a_data[j][i] for j in range(21)] + [array_b_data[j][i] for j in range(21)]
            collected_data.append(row)
        save_to_excel(collected_data)
        print("Data saved to Excel file.")

if __name__ == "__main__":
    main()