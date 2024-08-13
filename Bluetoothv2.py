import asyncio
import threading
from bleak import BleakClient
import matplotlib.pyplot as plt
from queue import Queue
import time
import struct
import numpy as np

# Device address
DEVICE_ADDRESS = "FF:CD:AA:44:31:1D"
FSM_DATA_CHAR_UUID = "c042d543-3929-48c9-af11-cf252f5ce7f4"

# Global variables for data
time_data = []
avg_data = []
data_queue = Queue()
start_time = None

# Set up the plot
fig, ax = plt.subplots()
line, = ax.plot([], [], 'b-')
ax.set_xlabel('Time (s)')
ax.set_ylabel('Average Voltage (nV)')
ax.set_title('FSM Data Monitor')
ax.grid(True)

def parse_fsm_data(data):
    time_min, time_us = struct.unpack_from('<HI', data, 0)
    total_time = time_min * 60 + time_us / 1_000_000
    
    array_a_data = struct.unpack_from('<21I', data, 9)
    array_b_data = struct.unpack_from('<21I', data, 96)
    
    all_data = array_a_data + array_b_data
    avg_value = np.mean(all_data)
    
    return total_time, avg_value

def notification_handler(sender, data):
    global start_time
    if start_time is None:
        start_time = time.time()
    
    print(f"Received data (length: {len(data)}): {data.hex()}")
    
    try:
        total_time, avg_value = parse_fsm_data(data)
        timestamp = time.time() - start_time
        data_queue.put((timestamp, avg_value))
        print(f"Time: {timestamp:.2f}s, Avg Value: {avg_value:.2f}")
    except Exception as e:
        print(f"Error parsing data: {e}")

def update_plot():
    print("Update plot function called")
    while True:
        while not data_queue.empty():
            timestamp, avg_value = data_queue.get()
            time_data.append(timestamp)
            avg_data.append(avg_value)
        
        if time_data:
            ax.set_xlim(0, max(30, time_data[-1]))
            ax.set_ylim(min(avg_data) - 5, max(avg_data) + 5)
            line.set_data(time_data, avg_data)
        
        plt.pause(0.1)
        if not plt.get_fignums():
            break

async def run_ble_client(address):
    async with BleakClient(address) as client:
        try:
            print(f"Connected: {client.is_connected}")
            
            await client.start_notify(FSM_DATA_CHAR_UUID, notification_handler)
            print("FSM Data Monitoring Started. Close the plot window to stop.")
            
            counter = 0
            while plt.get_fignums():
                await asyncio.sleep(1)
                counter += 1
                if counter % 5 == 0:  # Print every 5 seconds
                    print(f"Waiting for data... (Time elapsed: {counter} seconds)")
        
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
    loop = asyncio.new_event_loop()
    ble_thread = threading.Thread(target=run_asyncio_loop, args=(loop,))
    ble_thread.start()

    plt.ion()
    update_plot()
    plt.ioff()

    ble_thread.join()

if __name__ == "__main__":
    main()