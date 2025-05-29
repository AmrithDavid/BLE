# Medical Device BLE Data Processing System

Real-time data acquisition and visualisation system for wireless BLE medical sensors, developed during my internship at IDE Group.

## Overview
Python-based system for continuous monitoring of fetal blood oxygen saturation sensors via Bluetooth Low Energy (BLE). Handles multi-channel sensor data with real-time visualisation and data logging.

## Tech Stack
- **Language:** Python
- **BLE Communication:** Bleak (async BLE library)
- **Visualization:** Matplotlib with Tkinter GUI
- **Data Processing:** NumPy, multi-threading
- **Data Export:** OpenPyXL for Excel generation

## Key Features
- Asynchronous BLE communication for continuous data streaming
- Real-time visualisation of 21 sensor channels (6 wavelengths × 3 photodiodes + LED OFF state)
- Multi-threaded architecture for non-blocking UI
- Configurable wavelength selection (784-881 nm)
- Automatic data logging with Excel export
- Linear/logarithmic plot scaling

## Architecture
- **BLE Handler:** Async notification handler for sensor data packets
- **Data Parser:** Extracts timestamp and 42 voltage readings from binary packets
- **GUI Thread:** Real-time plot updates without blocking data collection
- **Queue System:** Thread-safe data passing between BLE and GUI threads

## Status
Developed for verification and validation testing of medical devices at IDE Group (July-Nov 2024)
