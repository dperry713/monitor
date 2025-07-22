# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import messagebox, ttk
import obd
import colorsys
import threading
import random
import time
import serial.tools.list_ports
import traceback
import datetime
import csv
import os
import subprocess
import platform
import re
import ctypes
import math

# Try to import Bluetooth support (optional)
try:
    import bleak
    import asyncio
    BLUETOOTH_AVAILABLE = True
except ImportError:
    BLUETOOTH_AVAILABLE = False
    print("Bluetooth support not available. Install bleak for Bluetooth functionality.")

# Try to import WMI for Windows device management (optional)
try:
    import wmi  # type: ignore
    WMI_AVAILABLE = True
except ImportError:
    wmi = None
    WMI_AVAILABLE = False

# Async function for Bluetooth device scanning


async def scan_bluetooth_devices():
    """Scan for Bluetooth devices using multiple methods"""
    if not BLUETOOTH_AVAILABLE:
        return []

    devices = []
    try:
        # Method 1: Try BLE scanning first with shorter timeout
        print("Scanning for BLE devices...")

        # Use shorter timeout to prevent hanging
        timeout = 3.0
        ble_devices = await bleak.BleakScanner.discover(timeout=timeout)
        devices.extend(ble_devices)
        print(f"Found {len(ble_devices)} BLE devices")

        # Filter for potential OBD devices
        obd_devices = []
        for device in ble_devices:
            device_name = device.name or "Unknown"
            if any(keyword in device_name.upper() for keyword in ['OBD', 'ELM', 'OBDII', 'DIAGNOSTIC', 'ECU']):
                obd_devices.append(device)
                print(
                    f"Found potential OBD device: {device_name} ({device.address})")

        if obd_devices:
            print(f"Filtered to {len(obd_devices)} potential OBD devices")
            return obd_devices

    except Exception as e:
        print(f"BLE scan error: {e}")
        # Don't fail completely if BLE scan fails - continue with empty list

    return devices


def scan_paired_bluetooth_devices():
    """Scan for paired Bluetooth devices using Windows commands"""
    try:
        if platform.system() != "Windows":
            print("Paired device scanning only supported on Windows")
            return []

        paired_devices = []

        # Method 1: Use PowerShell to get Bluetooth devices
        try:
            cmd = 'powershell "Get-PnpDevice -Class Bluetooth | Where-Object {$_.Status -eq \'OK\'} | Select-Object FriendlyName, InstanceId"'
            result = subprocess.run(
                cmd, shell=True, capture_output=True, text=True, timeout=15)

            if result.returncode == 0 and result.stdout:
                lines = result.stdout.strip().split('\n')
                for line in lines[3:]:  # Skip header lines
                    if line.strip() and not line.startswith('-'):
                        try:
                            # Split by multiple spaces to separate name and instance ID
                            parts = re.split(r'\s{2,}', line.strip())
                            if len(parts) >= 2:
                                name = parts[0].strip()
                                instance_id = parts[-1].strip()

                                # Extract MAC address from instance ID if possible
                                if 'DEV_' in instance_id:
                                    mac_match = re.search(
                                        r'DEV_([0-9A-F]{12})', instance_id)
                                    if mac_match:
                                        mac_raw = mac_match.group(1)
                                        # Convert to standard MAC format
                                        mac = ':'.join(
                                            [mac_raw[i:i+2] for i in range(0, len(mac_raw), 2)])
                                        paired_devices.append(
                                            {'name': name, 'address': mac})
                        except Exception as parse_error:
                            print(
                                f"Error parsing device line: {line} - {parse_error}")
                            continue
        except Exception as ps_error:
            print(f"PowerShell method failed: {ps_error}")

        # Method 2: Try alternative Windows command for Bluetooth COM ports
        if not paired_devices:
            try:
                cmd = 'reg query "HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Services\\BTHPORT\\Parameters\\Devices" /s'
                result = subprocess.run(
                    cmd, shell=True, capture_output=True, text=True, timeout=10)

                if result.returncode == 0:
                    lines = result.stdout.split('\n')
                    for line in lines:
                        if 'Name' in line and 'REG_SZ' in line:
                            try:
                                name = line.split('REG_SZ')[-1].strip()
                                if name and name != 'Name' and any(keyword in name.upper() for keyword in ['OBD', 'ELM', 'OBDII']):
                                    paired_devices.append(
                                        {'name': name, 'address': 'Unknown'})
                            except:
                                continue
            except Exception as reg_error:
                print(f"Registry method failed: {reg_error}")

        # Method 3: Try WMI if available
        if not paired_devices and WMI_AVAILABLE and wmi is not None:
            try:
                c = wmi.WMI()
                for device in c.Win32_PnPEntity():
                    if device.Name and 'bluetooth' in device.Name.lower():
                        if any(keyword in device.Name.upper() for keyword in ['OBD', 'ELM', 'OBDII', 'DIAGNOSTIC']):
                            paired_devices.append(
                                {'name': device.Name, 'address': 'Unknown'})
            except Exception as wmi_error:
                print(f"WMI method failed: {wmi_error}")
        elif not paired_devices:
            print("WMI not available")

        print(f"Found {len(paired_devices)} paired devices")
        return paired_devices
    except Exception as e:
        print(f"Error scanning paired devices: {e}")
        return []


def find_com_ports_for_bluetooth():
    """Find COM ports associated with Bluetooth devices"""
    try:
        bluetooth_ports = []
        ports = serial.tools.list_ports.comports()

        for port in ports:
            # Check if this is a Bluetooth serial port
            description_upper = port.description.upper()
            hwid_upper = port.hwid.upper() if port.hwid else ""

            # Look for Bluetooth indicators
            bluetooth_keywords = ['BLUETOOTH', 'BT',
                                  'RFCOMM', 'SPP', 'SERIAL PORT PROFILE']
            obd_keywords = ['OBD', 'ELM', 'OBDII', 'DIAGNOSTIC', 'ECU']

            is_bluetooth = any(
                keyword in description_upper for keyword in bluetooth_keywords)
            is_bluetooth = is_bluetooth or any(
                keyword in hwid_upper for keyword in bluetooth_keywords)

            # Also check for "Standard Serial over Bluetooth" pattern
            is_bluetooth = is_bluetooth or 'STANDARD SERIAL' in description_upper

            if is_bluetooth:
                # Check if it might be an OBD device
                is_likely_obd = any(
                    keyword in description_upper for keyword in obd_keywords)

                bluetooth_ports.append({
                    'port': port.device,
                    'description': port.description,
                    'hwid': port.hwid or 'Unknown',
                    'is_likely_obd': is_likely_obd
                })

        # Sort by likely OBD devices first, with COM7 prioritized for Bluetooth
        bluetooth_ports.sort(key=lambda x: (
            x['port'] == 'COM7',  # COM7 gets highest priority
            x['is_likely_obd']     # Then by OBD likelihood
        ), reverse=True)

        print(f"Found {len(bluetooth_ports)} Bluetooth COM ports")
        for port in bluetooth_ports:
            obd_indicator = " (Likely OBD)" if port['is_likely_obd'] else ""
            com7_indicator = " ⭐ COMMON BLUETOOTH OBD PORT" if port['port'] == 'COM7' else ""
            print(
                f"  {port['port']}: {port['description']}{obd_indicator}{com7_indicator}")

        return bluetooth_ports
    except Exception as e:
        print(f"Error finding Bluetooth COM ports: {e}")
        return []


def scan_all_com_ports():
    """Scan all COM ports and identify which might be OBD adapters"""
    try:
        all_ports = []
        ports = serial.tools.list_ports.comports()

        for port in ports:
            description_upper = port.description.upper()
            hwid_upper = port.hwid.upper() if port.hwid else ""

            # Check for OBD-related keywords
            obd_keywords = ['OBD', 'ELM', 'OBDII', 'DIAGNOSTIC', 'ECU', 'SCAN']
            usb_keywords = ['USB', 'FTDI', 'CH340', 'PL2303', 'CP210']
            bluetooth_keywords = ['BLUETOOTH', 'BT', 'RFCOMM']

            is_obd = any(
                keyword in description_upper for keyword in obd_keywords)
            is_usb = any(
                keyword in description_upper or keyword in hwid_upper for keyword in usb_keywords)
            is_bluetooth = any(
                keyword in description_upper for keyword in bluetooth_keywords)

            port_type = "Unknown"
            if is_bluetooth:
                port_type = "Bluetooth"
            elif is_usb:
                port_type = "USB/Serial"

            all_ports.append({
                'port': port.device,
                'description': port.description,
                'hwid': port.hwid or 'Unknown',
                'type': port_type,
                'is_likely_obd': is_obd,
                'is_bluetooth': is_bluetooth
            })

        # Sort by likely OBD devices first, then by type
        all_ports.sort(key=lambda x: (
            x['is_likely_obd'], x['is_bluetooth']), reverse=True)

        print(f"Found {len(all_ports)} total COM ports:")
        for port in all_ports:
            obd_indicator = " (Likely OBD)" if port['is_likely_obd'] else ""
            print(
                f"  {port['port']} ({port['type']}): {port['description']}{obd_indicator}")

        return all_ports
    except Exception as e:
        print(f"Error scanning COM ports: {e}")
        return []


# Define table axes
# 400 to 8000 RPM in 400 RPM increments
RPM_VALUES = [400 + 400 * i for i in range(20)]
# 15 to 105 kPa in 5 kPa increments
MAP_VALUES = [15 + 5 * i for i in range(19)]

# Function to map VE to color (blue to red)


def ve_to_color(ve):
    ve = max(0, min(3, ve))  # Clamp VE to 0-3
    h = (1 - ve / 3) * (240 / 360)  # Hue from blue (240°) to red (0°)
    r, g, b = colorsys.hsv_to_rgb(h, 1, 1)
    return f'#{int(r*255):02x}{int(g*255):02x}{int(b*255):02x}'


# Create main window
window = tk.Tk()
window.title("VE Table Monitor - Professional Engine Tuning Tool")
window.configure(bg='#2b2b2b')
window.geometry("1400x900")

# Configure modern styling
style = ttk.Style()
style.theme_use('clam')
style.configure('TLabel', background='#2b2b2b',
                foreground='white', font=('Segoe UI', 9))
style.configure('TButton', font=('Segoe UI', 9))
style.configure('TEntry', font=('Segoe UI', 9))
style.configure('TNotebook', background='#2b2b2b', tabposition='n')
style.configure('TNotebook.Tab', background='#404040', foreground='white',
                padding=[20, 8], font=('Segoe UI', 10, 'bold'))
style.map('TNotebook.Tab', background=[
          ('selected', '#4CAF50'), ('active', '#666666')])

# Create main frame with padding
main_frame = tk.Frame(window, bg='#2b2b2b', padx=10, pady=10)
main_frame.pack(fill='both', expand=True)

# Create notebook for tabs
notebook = ttk.Notebook(main_frame)
notebook.pack(fill='both', expand=True)

# Create tabs
ve_table_tab = ttk.Frame(notebook)
connection_tab = ttk.Frame(notebook)
pid_monitor_tab = ttk.Frame(notebook)
dtc_tab = ttk.Frame(notebook)
settings_tab = ttk.Frame(notebook)

# Add tabs to notebook
notebook.add(ve_table_tab, text='📊 VE Table Monitor')
notebook.add(connection_tab, text='🔌 Connection')
notebook.add(pid_monitor_tab, text='📡 PID Monitor')
notebook.add(dtc_tab, text='🚨 DTC Scanner')
notebook.add(settings_tab, text='⚙️ Settings')

# Configure tab backgrounds
for tab in [ve_table_tab, connection_tab, pid_monitor_tab, dtc_tab, settings_tab]:
    tab.configure(style='Tab.TFrame')

style.configure('Tab.TFrame', background='#2b2b2b')

# ===== VE TABLE TAB =====
ve_main_frame = tk.Frame(ve_table_tab, bg='#2b2b2b', padx=20, pady=15)
ve_main_frame.pack(fill='both', expand=True)

# Title section for VE Table tab
ve_title_frame = tk.Frame(ve_main_frame, bg='#2b2b2b')
ve_title_frame.pack(fill='x', pady=(0, 20))

ve_title_label = tk.Label(ve_title_frame, text="Volumetric Efficiency Table",
                          font=('Segoe UI', 16, 'bold'), fg='#4CAF50', bg='#2b2b2b')
ve_title_label.pack(side='left')

ve_subtitle_label = tk.Label(ve_title_frame, text="Real-time Engine Performance Monitoring",
                             font=('Segoe UI', 10), fg='#888888', bg='#2b2b2b')
ve_subtitle_label.pack(side='left', padx=(20, 0))

# Create table frame with border
table_frame = tk.Frame(ve_main_frame, bg='#1e1e1e', relief='raised', bd=2)
table_frame.pack(side='left', padx=(0, 20), pady=(0, 20))

# Create headers with modern styling
header_font = ('Segoe UI', 8, 'bold')
for i, rpm in enumerate(RPM_VALUES):
    header = tk.Label(table_frame, text=str(rpm), font=header_font,
                      bg='#404040', fg='#FFD700', relief='flat', bd=1, padx=3, pady=2)
    header.grid(row=0, column=i+1, sticky='nsew', padx=1, pady=1)

for j, map_val in enumerate(MAP_VALUES):
    header = tk.Label(table_frame, text=str(map_val), font=header_font,
                      bg='#404040', fg='#FFD700', relief='flat', bd=1, padx=3, pady=2)
    header.grid(row=j+1, column=0, sticky='nsew', padx=1, pady=1)

# RPM and MAP axis labels
rpm_label = tk.Label(table_frame, text="RPM →", font=('Segoe UI', 9, 'bold'),
                     bg='#2b2b2b', fg='white')
rpm_label.grid(row=0, column=0, sticky='nsew')

# Create table cells with modern styling
cell_font = ('Segoe UI', 8)
cells = [[tk.Label(table_frame, text="0.00", width=6, font=cell_font,
                   bg='#333333', fg='white', relief='flat', bd=1)
          for _ in range(20)] for _ in range(19)]
for j in range(19):
    for i in range(20):
        cells[j][i].grid(row=j+1, column=i+1, sticky='nsew', padx=1, pady=1)

# Create right panel for sensor data in VE table tab
ve_right_panel = tk.Frame(ve_main_frame, bg='#2b2b2b', width=350)
ve_right_panel.pack(side='right', fill='y', padx=(20, 0))
ve_right_panel.pack_propagate(False)

# Sensor data section
sensor_frame = tk.LabelFrame(ve_right_panel, text="Real-time Sensor Data",
                             font=('Segoe UI', 11, 'bold'), fg='#4CAF50', bg='#2b2b2b',
                             relief='flat', bd=2)
sensor_frame.pack(fill='x', pady=(0, 20))

# Create sensor value labels with modern styling
labels = {}
sensor_names = ["RPM", "MAP", "IAT", "MAF", "g/cyl", "VE"]
sensor_units = ["rpm", "kPa", "°K", "g/s", "g/cyl", "ratio"]
sensor_colors = ["#FF6B6B", "#4ECDC4",
                 "#45B7D1", "#96CEB4", "#FECA57", "#6C5CE7"]

for idx, (key, unit, color) in enumerate(zip(sensor_names, sensor_units, sensor_colors)):
    row_frame = tk.Frame(sensor_frame, bg='#2b2b2b')
    row_frame.pack(fill='x', pady=5, padx=10)

    name_label = tk.Label(row_frame, text=f"{key}:", font=('Segoe UI', 10, 'bold'),
                          fg=color, bg='#2b2b2b', width=8, anchor='w')
    name_label.pack(side='left')

    labels[key] = tk.Label(row_frame, text="0", font=('Segoe UI', 12, 'bold'),
                           fg='white', bg='#404040', relief='flat', bd=1,
                           width=10, anchor='center')
    labels[key].pack(side='left', padx=(10, 5))

    unit_label = tk.Label(row_frame, text=unit, font=('Segoe UI', 9),
                          fg='#888888', bg='#2b2b2b', anchor='w')
    unit_label.pack(side='left')

# Add legend to VE table tab
ve_legend_frame = tk.LabelFrame(ve_right_panel, text="VE Color Legend",
                                font=('Segoe UI', 11, 'bold'), fg='#4CAF50', bg='#2b2b2b',
                                relief='flat', bd=2)
ve_legend_frame.pack(fill='x', pady=(0, 10))

# Color legend
legend_colors = [(0, "Poor"), (0.75, "Fair"), (1.5, "Good"),
                 (2.25, "Excellent"), (3, "Outstanding")]
for ve_val, label in legend_colors:
    color_frame = tk.Frame(ve_legend_frame, bg='#2b2b2b')
    color_frame.pack(fill='x', pady=2, padx=10)

    color_box = tk.Label(color_frame, text="  ", bg=ve_to_color(ve_val),
                         relief='flat', bd=1, width=3)
    color_box.pack(side='left')

    tk.Label(color_frame, text=f"{label} (VE: {ve_val})",
             font=('Segoe UI', 9), fg='white', bg='#2b2b2b').pack(side='left', padx=(10, 0))

# ===== CONNECTION TAB =====
conn_main_frame = tk.Frame(connection_tab, bg='#2b2b2b', padx=20, pady=15)
conn_main_frame.pack(fill='both', expand=True)

# Connection title
conn_title_frame = tk.Frame(conn_main_frame, bg='#2b2b2b')
conn_title_frame.pack(fill='x', pady=(0, 20))

conn_title_label = tk.Label(conn_title_frame, text="Vehicle Connection",
                            font=('Segoe UI', 16, 'bold'), fg='#4CAF50', bg='#2b2b2b')
conn_title_label.pack(side='left')

conn_subtitle_label = tk.Label(conn_title_frame, text="Configure OBD-II Connection Settings",
                               font=('Segoe UI', 10), fg='#888888', bg='#2b2b2b')
conn_subtitle_label.pack(side='left', padx=(20, 0))

# Connection control section
connection_frame = tk.LabelFrame(conn_main_frame, text="Connection Control",
                                 font=('Segoe UI', 11, 'bold'), fg='#4CAF50', bg='#2b2b2b',
                                 relief='flat', bd=2)
connection_frame.pack(fill='x', pady=(0, 20))

# Status label with modern styling
status_label = tk.Label(connection_frame, text="● Not Connected",
                        font=('Segoe UI', 11, 'bold'), fg='#FF4444', bg='#2b2b2b')
status_label.pack(pady=10)

# COM port selection with modern styling
port_frame = tk.Frame(connection_frame, bg='#2b2b2b')
port_frame.pack(fill='x', pady=5, padx=10)

# Connection type selection
type_frame = tk.Frame(connection_frame, bg='#2b2b2b')
type_frame.pack(fill='x', pady=5, padx=10)

tk.Label(type_frame, text="Connection:", font=('Segoe UI', 10),
         fg='white', bg='#2b2b2b').pack(side='left')

connection_type = tk.StringVar(value="Serial")
bt_options = ["Serial", "Bluetooth"] if BLUETOOTH_AVAILABLE else ["Serial"]
type_combo = ttk.Combobox(type_frame, textvariable=connection_type,
                          values=bt_options, state="readonly", width=10)
type_combo.pack(side='right')

# Add callback to refresh device list when connection type changes


def on_connection_type_change(*args):
    """Called when connection type changes"""
    populate_device_list()


connection_type.trace('w', on_connection_type_change)

# Device selection dropdown
tk.Label(port_frame, text="Device:", font=('Segoe UI', 10),
         fg='white', bg='#2b2b2b').pack(side='left')

# Global variables for device selection
available_devices = []
device_var = tk.StringVar(value="No devices found")
device_combo = ttk.Combobox(port_frame, textvariable=device_var,
                            values=["No devices found"], state="readonly", width=30)
device_combo.pack(side='right')

# Legacy port entry (hidden by default, shown for manual entry)
com_port_var = tk.StringVar(value="COM3")
com_port_entry = ttk.Entry(
    port_frame, textvariable=com_port_var, width=15, font=('Segoe UI', 10))
# Keep entry hidden initially - will be shown if needed

# Add scan button for Bluetooth
scan_frame = tk.Frame(connection_frame, bg='#2b2b2b')
scan_frame.pack(fill='x', pady=5, padx=10)

# Add a test Bluetooth button to the connection tab
test_bt_button = tk.Button(scan_frame, text="🧪 Test Bluetooth",
                           font=('Segoe UI', 9), bg='#9C27B0', fg='white',
                           relief='flat', bd=0, padx=15, pady=4,
                           command=lambda: threading.Thread(target=test_bluetooth_setup).start())
test_bt_button.pack(side='left', padx=(5, 0))

test_ports_button = tk.Button(scan_frame, text="🔌 Test COM3/COM7",
                              font=('Segoe UI', 9), bg='#FF5722', fg='white',
                              relief='flat', bd=0, padx=15, pady=4,
                              command=lambda: threading.Thread(target=test_connection_ports).start())
test_ports_button.pack(side='left', padx=(5, 0))

# Function to test Bluetooth setup


def test_bluetooth_setup():
    """Test the current Bluetooth setup and provide diagnostics"""
    try:
        status_label.config(text="● Testing Bluetooth setup...", fg="#FFA726")

        # Test 1: Check if bleak is available
        test_results = []
        test_results.append("=== BLUETOOTH DIAGNOSTICS ===")

        if BLUETOOTH_AVAILABLE:
            test_results.append("✅ Bleak library: INSTALLED")
        else:
            test_results.append("❌ Bleak library: NOT INSTALLED")
            test_results.append("   → Run: pip install bleak")

        # Test 2: Check Windows Bluetooth service
        try:
            cmd = 'sc query bthserv'
            result = subprocess.run(
                cmd, shell=True, capture_output=True, text=True, timeout=5)
            if 'RUNNING' in result.stdout:
                test_results.append("✅ Windows Bluetooth service: RUNNING")
            else:
                test_results.append("❌ Windows Bluetooth service: NOT RUNNING")
                test_results.append("   → Start Bluetooth service in Windows")
        except:
            test_results.append("⚠️ Windows Bluetooth service: UNKNOWN")

        # Test 3: Check COM ports
        all_ports = scan_all_com_ports()
        bt_ports = [p for p in all_ports if p['is_bluetooth']]
        obd_ports = [p for p in all_ports if p['is_likely_obd']]

        test_results.append(f"✅ Total COM ports found: {len(all_ports)}")
        test_results.append(f"📶 Bluetooth COM ports: {len(bt_ports)}")
        test_results.append(f"🚗 Likely OBD ports: {len(obd_ports)}")

        if obd_ports:
            test_results.append("\n📋 RECOMMENDED PORTS:")
            for port in obd_ports[:3]:  # Show top 3
                test_results.append(
                    f"   • {port['port']} - {port['description']}")

        # Test 4: Try PowerShell access
        try:
            cmd = 'powershell "Get-Service -Name bthserv"'
            result = subprocess.run(
                cmd, shell=True, capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                test_results.append("✅ PowerShell access: WORKING")
            else:
                test_results.append("❌ PowerShell access: LIMITED")
        except:
            test_results.append("⚠️ PowerShell access: FAILED")

        # Test 5: Administrator rights
        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin()
            if is_admin:
                test_results.append("✅ Administrator rights: YES")
            else:
                test_results.append(
                    "⚠️ Administrator rights: NO (may limit Bluetooth access)")
        except Exception:
            test_results.append("⚠️ Administrator rights: UNKNOWN")

        test_results.append("\n🔧 RECOMMENDATIONS:")

        if not BLUETOOTH_AVAILABLE:
            test_results.append("1. Install bleak: pip install bleak")

        if obd_ports:
            test_results.append(
                "2. Try connecting with the recommended COM port")
        else:
            test_results.append(
                "2. Pair your OBD adapter in Windows Bluetooth settings")
            test_results.append(
                "3. Make sure vehicle is running to power the adapter")

        test_results.append("4. Use 'Scan Devices' to find your adapter")
        test_results.append(
            "5. Prefer COM ports over MAC addresses for connection")

        # Show results
        result_text = "\n".join(test_results)
        messagebox.showinfo("Bluetooth Diagnostics", result_text)

    except Exception as e:
        messagebox.showerror("Diagnostic Error",
                             f"Error running diagnostics:\n{str(e)}")
    finally:
        status_label.config(text="● Not Connected", fg="#FF4444")


scan_button = tk.Button(scan_frame, text="🔍 Scan Devices",
                        font=('Segoe UI', 9), bg='#607D8B', fg='white',
                        relief='flat', bd=0, padx=15, pady=4,
                        command=lambda: threading.Thread(target=scan_for_devices).start())
scan_button.pack(side='right')

# Buttons with modern styling
button_frame = tk.Frame(connection_frame, bg='#2b2b2b')
button_frame.pack(fill='x', pady=10, padx=10)

connect_button = tk.Button(button_frame, text="🔌 Connect Vehicle",
                           font=('Segoe UI', 10, 'bold'), bg='#4CAF50', fg='white',
                           relief='flat', bd=0, padx=20, pady=8,
                           command=lambda: threading.Thread(target=connect_to_vehicle).start())
connect_button.pack(fill='x', pady=(0, 5))

# Quick connect buttons for specific ports
quick_connect_frame = tk.Frame(button_frame, bg='#2b2b2b')
quick_connect_frame.pack(fill='x', pady=2)

# First row - Standard ports
com3_button = tk.Button(quick_connect_frame, text="🔌 COM3",
                        font=('Segoe UI', 8, 'bold'), bg='#607D8B', fg='white',
                        relief='flat', bd=0, padx=10, pady=6,
                        command=lambda: threading.Thread(target=lambda: quick_connect_port("COM3")).start())
com3_button.pack(side='left', fill='x', expand=True, padx=(0, 1))

com7_button = tk.Button(quick_connect_frame, text="📱 COM7",
                        font=('Segoe UI', 8, 'bold'), bg='#9C27B0', fg='white',
                        relief='flat', bd=0, padx=10, pady=6,
                        command=lambda: threading.Thread(target=lambda: quick_connect_port("COM7")).start())
com7_button.pack(side='left', fill='x', expand=True, padx=(1, 1))

# Second row - OBDXPROVX specific ports
obdx_connect_frame = tk.Frame(button_frame, bg='#2b2b2b')
obdx_connect_frame.pack(fill='x', pady=2)

com5_button = tk.Button(obdx_connect_frame, text="🔵 COM5 (OBDX)",
                        font=('Segoe UI', 8, 'bold'), bg='#FF5722', fg='white',
                        relief='flat', bd=0, padx=10, pady=6,
                        command=lambda: threading.Thread(target=lambda: quick_connect_port("COM5")).start())
com5_button.pack(side='left', fill='x', expand=True, padx=(0, 1))

com6_button = tk.Button(obdx_connect_frame, text="🔵 COM6 (OBDX)",
                        font=('Segoe UI', 8, 'bold'), bg='#FF5722', fg='white',
                        relief='flat', bd=0, padx=10, pady=6,
                        command=lambda: threading.Thread(target=lambda: quick_connect_port("COM6")).start())
com6_button.pack(side='left', fill='x', expand=True, padx=(1, 1))

# Emergency direct COM7 button


def emergency_com7_connect():
    """Emergency direct connection to COM7 bypassing all checks"""
    global connection
    try:
        status_label.config(
            text="● EMERGENCY COM7 CONNECTION...", fg="#FF9800")

        # Close any existing connection
        if connection:
            try:
                connection.close()
            except:
                pass
            connection = None

        print("🚨 EMERGENCY COM7 CONNECTION ATTEMPT")

        # Try direct connection with most common Bluetooth OBD settings
        try:
            print("Trying 38400 baud...")
            connection = obd.OBD('COM7', baudrate=38400,
                                 timeout=10, fast=False)
            if connection and connection.is_connected():
                print("✅ EMERGENCY COM7 CONNECTION SUCCESSFUL!")
                status_label.config(
                    text="● Connected: COM7 (Emergency)", fg="#4CAF50")
                messagebox.showinfo(
                    "Emergency Connection", "✅ COM7 Connected!\n\nEmergency connection to COM7 successful.\nYou can now use PID scanning and monitoring.")
                update()
                return
        except Exception as e:
            print(f"38400 baud failed: {e}")

        # Try 9600 baud
        try:
            print("Trying 9600 baud...")
            connection = obd.OBD('COM7', baudrate=9600, timeout=15, fast=False)
            if connection and connection.is_connected():
                print("✅ EMERGENCY COM7 CONNECTION SUCCESSFUL!")
                status_label.config(
                    text="● Connected: COM7 (Emergency)", fg="#4CAF50")
                messagebox.showinfo(
                    "Emergency Connection", "✅ COM7 Connected!\n\nEmergency connection to COM7 successful.\nYou can now use PID scanning and monitoring.")
                update()
                return
        except Exception as e:
            print(f"9600 baud failed: {e}")

        # If all attempts failed
        status_label.config(text="● Emergency COM7 failed", fg="#FF4444")
        messagebox.showerror("Emergency Connection Failed",
                             "❌ Emergency COM7 connection failed!\n\n"
                             "Troubleshooting:\n"
                             "• Ensure vehicle is running\n"
                             "• Check COM7 exists in Device Manager\n"
                             "• Verify Bluetooth adapter is paired\n"
                             "• Close other OBD software\n"
                             "• Try running as Administrator")

    except Exception as e:
        status_label.config(text="● Emergency connection error", fg="#FF4444")
        messagebox.showerror("Emergency Connection Error", f"Error: {str(e)}")


emergency_button = tk.Button(quick_connect_frame, text="🚨 FORCE COM7",
                             font=('Segoe UI', 8, 'bold'), bg='#FF5722', fg='white',
                             relief='flat', bd=0, padx=10, pady=6,
                             command=lambda: threading.Thread(target=emergency_com7_connect).start())
emergency_button.pack(side='right', fill='x', expand=True, padx=(1, 0))

# Add COM port diagnostic button


def diagnose_com_ports():
    """Diagnose available COM ports and Bluetooth status"""
    try:
        status_label.config(text="● Diagnosing COM ports...", fg="#FFA726")

        # Get all available COM ports
        ports = serial.tools.list_ports.comports()

        diagnostic_msg = "🔍 COM PORT DIAGNOSTIC REPORT\n\n"
        diagnostic_msg += f"📊 Found {len(ports)} total COM ports:\n\n"

        bluetooth_ports = []
        obd_ports = []
        other_ports = []

        for port in ports:
            port_info = f"• {port.device}: {port.description}\n"
            port_info += f"  HWID: {port.hwid or 'Unknown'}\n"

            desc_upper = port.description.upper()

            # Check if it's a Bluetooth port
            is_bluetooth = any(keyword in desc_upper for keyword in
                               ['BLUETOOTH', 'BT', 'RFCOMM', 'SPP', 'STANDARD SERIAL'])

            # Check if it's likely an OBD port
            is_obd = any(keyword in desc_upper for keyword in
                         ['OBD', 'ELM', 'OBDII', 'DIAGNOSTIC'])

            if is_bluetooth:
                port_info += "  🔵 TYPE: BLUETOOTH\n"
                bluetooth_ports.append(port.device)
            elif is_obd:
                port_info += "  🚗 TYPE: LIKELY OBD\n"
                obd_ports.append(port.device)
            else:
                port_info += "  📱 TYPE: SERIAL/USB\n"
                other_ports.append(port.device)

            diagnostic_msg += port_info + "\n"

        # Summary
        diagnostic_msg += f"📋 SUMMARY:\n"
        diagnostic_msg += f"• Bluetooth COM ports: {bluetooth_ports if bluetooth_ports else 'None found'}\n"
        diagnostic_msg += f"• Likely OBD ports: {obd_ports if obd_ports else 'None found'}\n"
        diagnostic_msg += f"• Other ports: {other_ports if other_ports else 'None found'}\n\n"

        # Recommendations
        diagnostic_msg += f"💡 RECOMMENDATIONS:\n"
        if bluetooth_ports:
            recommended_port = bluetooth_ports[0]
            diagnostic_msg += f"• Try connecting to: {recommended_port}\n"
            diagnostic_msg += f"• Use the 'Quick Connect' button for {recommended_port}\n"
        else:
            diagnostic_msg += f"• No Bluetooth COM ports found!\n"
            diagnostic_msg += f"• Check Windows Device Manager → Ports\n"
            diagnostic_msg += f"• Ensure OBD adapter is paired in Windows Bluetooth settings\n"
            diagnostic_msg += f"• Look for 'Standard Serial over Bluetooth' entries\n"

        diagnostic_msg += f"\n🔧 NEXT STEPS:\n"
        diagnostic_msg += f"1. Pair OBD adapter in Windows Bluetooth settings\n"
        diagnostic_msg += f"2. Check Device Manager → Ports → look for Bluetooth entries\n"
        diagnostic_msg += f"3. Note the COM port number assigned\n"
        diagnostic_msg += f"4. Use that COM port number in the app\n"

        messagebox.showinfo("COM Port Diagnostic", diagnostic_msg)

        # Update the quick connect buttons if we found Bluetooth ports
        if bluetooth_ports:
            recommended_port = bluetooth_ports[0]
            print(f"✅ Found Bluetooth COM port: {recommended_port}")

            # Update button text to show actual port
            if recommended_port != "COM7":
                com7_button.config(
                    text=f"📱 Connect {recommended_port}",
                    command=lambda: threading.Thread(
                        target=lambda: quick_connect_port(recommended_port)).start()
                )
                emergency_button.config(
                    text=f"🚨 FORCE {recommended_port}",
                    command=lambda: threading.Thread(
                        target=lambda: force_connect_port(recommended_port)).start()
                )

    except Exception as e:
        messagebox.showerror("Diagnostic Error",
                             f"Error diagnosing COM ports:\n{str(e)}")
    finally:
        status_label.config(text="● Not Connected", fg="#FF4444")


def force_connect_port(port):
    """Force connection to a specific port with minimal error checking"""
    global connection
    try:
        status_label.config(
            text=f"● FORCING connection to {port}...", fg="#FF9800")

        # Close any existing connection
        if connection:
            try:
                connection.close()
            except:
                pass
            connection = None

        print(f"🚨 FORCING CONNECTION TO {port}")

        # Check if port actually exists first
        existing_ports = [p.device for p in serial.tools.list_ports.comports()]
        if port not in existing_ports:
            raise Exception(
                f"❌ {port} does not exist! Available ports: {existing_ports}")

        # Try direct connection
        connection = obd.OBD(port, baudrate=38400, timeout=5, fast=False)

        if connection and connection.is_connected():
            print(f"✅ FORCE CONNECTION SUCCESSFUL!")
            status_label.config(
                text=f"● Connected: {port} (FORCED)", fg="#4CAF50")
            messagebox.showinfo("Force Connection Success",
                                f"✅ Successfully connected to {port}!")
            update()
        else:
            raise Exception(f"Connection object created but not connected")

    except Exception as e:
        status_label.config(text=f"● Force connection failed", fg="#FF4444")
        messagebox.showerror("Force Connection Failed",
                             f"❌ Could not force connect to {port}\n\n"
                             f"Error: {str(e)}\n\n"
                             f"Try:\n"
                             f"• Check that {port} exists in Device Manager\n"
                             f"• Ensure the device is properly paired\n"
                             f"• Run 'COM Port Diagnostic' to see available ports")


# Add diagnostic button
diagnostic_button = tk.Button(scan_frame, text="🔍 COM Diagnostic",
                              font=('Segoe UI', 9), bg='#795548', fg='white',
                              relief='flat', bd=0, padx=15, pady=4,
                              command=lambda: threading.Thread(target=diagnose_com_ports).start())
diagnostic_button.pack(side='left', padx=(5, 0))

# Add OBDXPROVX specific diagnostic
obdx_diagnostic_button = tk.Button(scan_frame, text="🔵 OBDX Diagnostic",
                                   font=('Segoe UI', 9), bg='#FF5722', fg='white',
                                   relief='flat', bd=0, padx=15, pady=4,
                                   command=lambda: threading.Thread(target=diagnose_obdxprovx).start())
obdx_diagnostic_button.pack(side='left', padx=(5, 0))

# Add OBDXPROVX connection stability test
stability_test_button = tk.Button(scan_frame, text="🔄 Stability Test",
                                  font=('Segoe UI', 9), bg='#9C27B0', fg='white',
                                  relief='flat', bd=0, padx=15, pady=4,
                                  command=lambda: threading.Thread(target=test_obdxprovx_stability).start())
stability_test_button.pack(side='left', padx=(5, 0))


def test_obdxprovx_stability():
    """Test OBDXPROVX connection stability before running full PID scan"""
    global connection

    if not connection or not connection.is_connected():
        messagebox.showerror("Connection Required",
                             "Please connect to your OBDXPROVX first using the quick connect buttons!")
        return

    try:
        status_label.config(
            text="● Testing OBDXPROVX stability...", fg="#FFA726")

        stability_msg = "🔄 OBDXPROVX CONNECTION STABILITY TEST\n\n"
        stability_msg += "📊 Running 10 connection health checks...\n\n"

        healthy_count = 0
        total_tests = 10
        response_times = []

        for test_num in range(1, total_tests + 1):
            print(f"🔍 Stability test {test_num}/{total_tests}")

            # Update status
            status_label.config(
                text=f"● Testing stability... {test_num}/{total_tests}", fg="#FFA726")

            is_healthy, health_msg = check_obdxprovx_connection_health(
                connection)

            if is_healthy:
                healthy_count += 1
                # Extract response time if available
                if "response in" in health_msg:
                    try:
                        time_str = health_msg.split("response in ")[
                            1].split("s")[0]
                        response_times.append(float(time_str))
                    except:
                        pass
                stability_msg += f"✅ Test {test_num}: {health_msg}\n"
            else:
                stability_msg += f"❌ Test {test_num}: {health_msg}\n"

            # Brief pause between tests
            time.sleep(1.0)

            # Update UI
            try:
                window.update_idletasks()
            except tk.TclError:
                return

        # Calculate statistics
        success_rate = (healthy_count / total_tests) * 100
        avg_response_time = sum(response_times) / \
            len(response_times) if response_times else 0

        stability_msg += f"\n📊 STABILITY RESULTS:\n"
        stability_msg += f"• Successful tests: {healthy_count}/{total_tests}\n"
        stability_msg += f"• Success rate: {success_rate:.1f}%\n"
        if response_times:
            stability_msg += f"• Average response time: {avg_response_time:.2f}s\n"
            stability_msg += f"• Response time range: {min(response_times):.2f}s - {max(response_times):.2f}s\n"

        stability_msg += f"\n💡 RECOMMENDATIONS:\n"

        if success_rate >= 90:
            stability_msg += f"🟢 EXCELLENT! Your OBDXPROVX is very stable.\n"
            stability_msg += f"• Safe to run full PID scans\n"
            stability_msg += f"• Connection should handle long monitoring sessions\n"
            stability_msg += f"• Consider using Data Logger for extended captures\n"
        elif success_rate >= 70:
            stability_msg += f"🟡 GOOD. Your OBDXPROVX is mostly stable.\n"
            stability_msg += f"• PID scanning should work but may have occasional drops\n"
            stability_msg += f"• Monitor connection during long sessions\n"
            stability_msg += f"• Close other OBD software for best results\n"
        elif success_rate >= 50:
            stability_msg += f"🟠 MARGINAL. Connection has some instability.\n"
            stability_msg += f"• Expect some connection drops during PID scanning\n"
            stability_msg += f"• Check vehicle engine is fully warmed up\n"
            stability_msg += f"• Verify OBDXPROVX is securely connected to OBD port\n"
            stability_msg += f"• Try repositioning vehicle or reducing interference\n"
        else:
            stability_msg += f"🔴 POOR. Connection is unstable.\n"
            stability_msg += f"• NOT recommended for PID scanning yet\n"
            stability_msg += f"• Check vehicle is running (not just ignition)\n"
            stability_msg += f"• Verify OBDXPROVX LED status\n"
            stability_msg += f"• Try reconnecting or using different COM port\n"
            stability_msg += f"• Close all other OBD software\n"
            stability_msg += f"• Consider restarting Bluetooth or rebooting\n"

        if avg_response_time > 5:
            stability_msg += f"\n⚠️ Response times are high ({avg_response_time:.1f}s average)\n"
            stability_msg += f"• This may indicate Bluetooth interference\n"
            stability_msg += f"• Try moving closer to vehicle\n"
            stability_msg += f"• Check for other Bluetooth devices nearby\n"

        title = "OBDXPROVX Stability Test Results"
        messagebox.showinfo(title, stability_msg)

        # Set status based on results
        if success_rate >= 80:
            status_label.config(
                text="● OBDXPROVX stable - ready for scanning", fg="#4CAF50")
        elif success_rate >= 60:
            status_label.config(text="● OBDXPROVX mostly stable", fg="#FF9800")
        else:
            status_label.config(
                text="● OBDXPROVX unstable - fix connection", fg="#FF4444")

    except Exception as e:
        status_label.config(text="● Stability test failed", fg="#FF4444")
        messagebox.showerror("Stability Test Error",
                             f"Error testing OBDXPROVX stability:\n{str(e)}\n\n"
                             f"This may indicate connection issues.\n"
                             f"Try reconnecting and test again.")
        print(f"🔵 Stability test error: {e}")


def diagnose_obdxprovx():
    """OBDXPROVX-specific diagnostic tool"""
    try:
        status_label.config(text="● Diagnosing OBDXPROVX...", fg="#FFA726")

        # Get all available COM ports
        ports = serial.tools.list_ports.comports()

        diagnostic_msg = "🔵 OBDXPROVX BLUETOOTH DIAGNOSTIC\n\n"
        diagnostic_msg += f"📊 Found {len(ports)} total COM ports:\n\n"

        obdx_ports = []
        bluetooth_ports = []

        for port in ports:
            desc = port.description.upper()
            hwid = (port.hwid or "").upper()

            # Check for Bluetooth keywords
            is_bluetooth = any(keyword in desc for keyword in
                               ['BLUETOOTH', 'BT', 'RFCOMM', 'SPP', 'STANDARD SERIAL'])

            # Check for OBDX specific identifiers
            is_obdx = any(keyword in desc for keyword in
                          ['OBDX', 'PROV', 'ELM', 'OBD']) or any(keyword in hwid for keyword in
                                                                 ['OBDX', 'PROV', 'ELM'])

            port_info = f"• {port.device}: {port.description}\n"

            if is_bluetooth:
                bluetooth_ports.append(port.device)
                port_info += "  🔵 BLUETOOTH PORT\n"

                # Your detected ports
                if is_obdx or port.device in ['COM5', 'COM6']:
                    obdx_ports.append(port.device)
                    port_info += "  🎯 POTENTIAL OBDXPROVX PORT!\n"

            diagnostic_msg += port_info + "\n"

        # Summary and recommendations
        diagnostic_msg += f"📋 SUMMARY:\n"
        diagnostic_msg += f"• Total COM ports: {len(ports)}\n"
        diagnostic_msg += f"• Bluetooth ports: {bluetooth_ports}\n"
        diagnostic_msg += f"• OBDXPROVX ports: {obdx_ports}\n\n"

        if obdx_ports:
            diagnostic_msg += f"🎯 OBDXPROVX RECOMMENDATIONS:\n"
            diagnostic_msg += f"• Primary: Try {obdx_ports[0]} first\n"
            if len(obdx_ports) > 1:
                diagnostic_msg += f"• Backup: Try {obdx_ports[1]} if primary fails\n"
            diagnostic_msg += f"• Use 'COM{obdx_ports[0][-1]} (OBDX)' quick connect button\n"
            diagnostic_msg += f"• Ensure vehicle ignition is ON\n"
            diagnostic_msg += f"• OBDXPROVX LED should be solid/blinking\n\n"
        else:
            diagnostic_msg += f"❌ NO OBDXPROVX PORTS DETECTED!\n\n"
            diagnostic_msg += f"🔧 TROUBLESHOOTING:\n"
            diagnostic_msg += f"• Check if OBDXPROVX is paired in Windows\n"
            diagnostic_msg += f"• Ensure vehicle is running (powers the device)\n"
            diagnostic_msg += f"• Re-pair device if necessary\n"
            diagnostic_msg += f"• Look for 'Standard Serial over Bluetooth'\n\n"

        diagnostic_msg += f"💡 OBDXPROVX TIPS:\n"
        diagnostic_msg += f"• Works best with 38400 or 9600 baud rate\n"
        diagnostic_msg += f"• Compatible with ELM327 protocol\n"
        diagnostic_msg += f"• Pair BEFORE plugging into OBD port\n"
        diagnostic_msg += f"• Use PIN '1234' or '0000' if prompted\n"
        diagnostic_msg += f"• Close other OBD software before connecting"

        messagebox.showinfo("OBDXPROVX Diagnostic", diagnostic_msg)

        # Update quick connect buttons if OBDX ports found
        if obdx_ports:
            print(f"✅ OBDXPROVX ports detected: {obdx_ports}")

    except Exception as e:
        messagebox.showerror("OBDXPROVX Diagnostic Error",
                             f"Error diagnosing OBDXPROVX:\n{str(e)}")
    finally:
        status_label.config(text="● Not Connected", fg="#FF4444")


# Global connection variable
connection = None
demo_mode = False

# Function to populate device dropdown


def populate_device_list():
    """Populate the device dropdown with available devices"""
    global available_devices
    try:
        available_devices.clear()
        device_list = []

        if connection_type.get() == "Serial":
            # Get all COM ports
            all_ports = scan_all_com_ports()
            for port in all_ports:
                # Create user-friendly display name
                display_name = f"{port['port']} - {port['description']}"
                if port['is_likely_obd']:
                    display_name += " ⭐"

                device_list.append(display_name)
                available_devices.append({
                    'display': display_name,
                    'port': port['port'],
                    'type': 'serial',
                    'info': port
                })

        elif connection_type.get() == "Bluetooth":
            if not BLUETOOTH_AVAILABLE:
                device_list = ["Bluetooth not available - Install bleak"]
            else:
                # ONLY use Bluetooth COM ports - avoid MAC addresses completely
                bt_ports = find_com_ports_for_bluetooth()

                # Verify that COM ports actually exist before adding them
                existing_ports = [
                    p.device for p in serial.tools.list_ports.comports()]
                verified_bt_ports = []

                for port in bt_ports:
                    if port['port'] in existing_ports:
                        verified_bt_ports.append(port)
                        print(f"✅ Verified Bluetooth COM port: {port['port']}")
                    else:
                        print(f"❌ Skipping non-existent port: {port['port']}")

                # Force COM7 to be first if it actually exists
                com7_port = None
                other_ports = []

                for port in verified_bt_ports:
                    if port['port'].upper() == 'COM7':
                        com7_port = port
                    else:
                        other_ports.append(port)

                # Add COM7 first if it actually exists
                if com7_port:
                    display_name = f"COM7 - 📱 BLUETOOTH OBD (RECOMMENDED) - {com7_port['description']}"
                    device_list.append(display_name)
                    available_devices.append({
                        'display': display_name,
                        'port': 'COM7',
                        'type': 'bluetooth_com',
                        'info': com7_port
                    })
                    print("✅ COM7 found and prioritized for Bluetooth connection")

                # Add other verified Bluetooth COM ports
                for port in other_ports:
                    display_name = f"{port['port']} - {port['description']}"
                    if port['is_likely_obd']:
                        display_name += " ⭐"

                    device_list.append(display_name)
                    available_devices.append({
                        'display': display_name,
                        'port': port['port'],
                        'type': 'bluetooth_com',
                        'info': port
                    })

                # DO NOT add MAC address devices - they don't work reliably on Windows
                # Skip paired devices that only have MAC addresses
                print(
                    f"🚫 Skipping MAC address devices - only using COM ports for Bluetooth")

                # If no verified Bluetooth COM ports found, provide helpful guidance
                if not verified_bt_ports:
                    device_list.append("❌ No active Bluetooth COM ports found")
                    device_list.append(
                        "💡 Click 'COM Diagnostic' to check available ports")
                    device_list.append(
                        "🔧 Pair OBD adapter in Windows Bluetooth settings")

        if not device_list:
            device_list = ["No devices found - Click 'Scan Devices'"]

        # Update dropdown - prioritize COM7 for Bluetooth
        device_combo['values'] = device_list
        if device_list and device_list[0] != "No devices found - Click 'Scan Devices'":
            # Try to set COM7 as default for Bluetooth if available
            com7_device = None
            for device in device_list:
                if "COM7" in device and connection_type.get() == "Bluetooth":
                    com7_device = device
                    break

            if com7_device:
                device_var.set(com7_device)
                print("Auto-selected COM7 for Bluetooth connection")
            else:
                device_var.set(device_list[0])  # Select first device
        else:
            device_var.set(device_list[0])

    except Exception as e:
        print(f"Error populating device list: {e}")
        device_combo['values'] = ["Error scanning devices"]
        device_var.set("Error scanning devices")

# Function to get selected device info


def get_selected_device():
    """Get the port/address of the currently selected device"""
    try:
        selected_display = device_var.get()
        print(f"🔍 Device Selection Debug:")
        print(f"  Selected Display: '{selected_display}'")
        print(f"  Available Devices: {len(available_devices)}")

        for i, device in enumerate(available_devices):
            print(
                f"  Device {i}: '{device['display']}' -> {device['port']} ({device['type']})")
            if device['display'] == selected_display:
                print(f"  ✅ MATCH FOUND: {device}")
                return device

        print(f"  ❌ NO MATCH FOUND for '{selected_display}'")
        return None
    except Exception as e:
        print(f"❌ Error in get_selected_device: {e}")
        return None

# Function to scan for available devices


def scan_for_devices():
    """Scan for available serial ports and Bluetooth devices"""
    try:
        status_label.config(text="● Scanning for devices...", fg="#FFA726")

        # Populate the device list
        populate_device_list()

        if connection_type.get() == "Serial":
            # Scan for serial ports
            ports = serial.tools.list_ports.comports()
            available_ports = [port.device for port in ports]

            if available_ports:
                message = f"Found {len(available_devices)} serial devices:\n\n"
                for device in available_devices:
                    port_info = device['info']
                    obd_indicator = " (Likely OBD)" if port_info['is_likely_obd'] else ""
                    message += f"• {port_info['port']}: {port_info['description']}{obd_indicator}\n"

                message += "\n💡 Devices marked with ⭐ are likely OBD adapters"
                messagebox.showinfo("Serial Devices Found", message)
            else:
                messagebox.showwarning(
                    "No Serial Ports", "No serial ports found.")

        elif connection_type.get() == "Bluetooth":
            # Enhanced Bluetooth scanning with device dropdown update
            if not BLUETOOTH_AVAILABLE:
                messagebox.showerror("Bluetooth Not Available",
                                     "Bluetooth support not installed.\n\nTo add Bluetooth support:\n1. pip install bleak\n2. Restart the application")
                return

            try:
                status_label.config(
                    text="● Scanning Bluetooth devices...", fg="#FFA726")

                # Method 1: Check ALL COM ports first (most comprehensive)
                all_com_ports = scan_all_com_ports()
                bt_com_ports = [
                    port for port in all_com_ports if port['is_bluetooth']]

                # Method 2: Get paired devices
                status_label.config(
                    text="● Scanning paired devices...", fg="#FFA726")
                paired_devices = scan_paired_bluetooth_devices()

                # Method 3: BLE scan for discoverable devices
                status_label.config(
                    text="● Scanning for BLE devices...", fg="#FFA726")

                nearby_devices = []
                try:
                    # Create and run event loop with timeout protection
                    loop = asyncio.new_event_loop()
                    asyncio.set_event_loop(loop)

                    # Use asyncio.wait_for to add an overall timeout
                    nearby_devices = loop.run_until_complete(
                        asyncio.wait_for(scan_bluetooth_devices(), timeout=8.0))

                except asyncio.TimeoutError:
                    print("BLE scan timed out - continuing with available devices")
                except Exception as ble_error:
                    print(f"BLE scan failed: {ble_error}")
                finally:
                    try:
                        loop.close()
                    except:
                        pass

                device_list = []
                found_obd_device = False

                # Show results and update dropdown
                total_found = len(available_devices)

                if total_found > 0:
                    device_list.append(
                        f"=== FOUND {total_found} BLUETOOTH COM PORTS ===")

                    # Only show COM devices (MAC addresses are not supported)
                    com_devices = [
                        d for d in available_devices if d['type'] == 'bluetooth_com']

                    if com_devices:
                        device_list.append(
                            f"\n📱 Bluetooth COM Ports ({len(com_devices)}):")
                        for device in com_devices:
                            obd_indicator = " ⭐ LIKELY OBD" if device['info']['is_likely_obd'] else ""
                            device_list.append(
                                f"  • {device['display']}{obd_indicator}")

                    message = "\n".join(device_list)
                    message += "\n\n💡 TIP: Select COM port from dropdown and click 'Connect Vehicle'"
                    message += "\n📱 COM ports are the most reliable connection method"

                    if len(com_devices) == 0:
                        message += "\n\n⚠️ No Bluetooth COM ports found!"
                        message += "\n� Check Device Manager → Ports for 'Standard Serial over Bluetooth'"

                    messagebox.showinfo("Bluetooth COM Ports Found", message)
                else:
                    # No devices found at all
                    troubleshoot_msg = "No Bluetooth devices found!\n\n"
                    troubleshoot_msg += "TROUBLESHOOTING STEPS:\n"
                    troubleshoot_msg += "1. 🔌 Ensure OBD adapter is plugged into vehicle\n"
                    troubleshoot_msg += "2. 🚗 Start the vehicle (powers the adapter)\n"
                    troubleshoot_msg += "3. 🔵 Check Windows Bluetooth settings\n"
                    troubleshoot_msg += "4. 📱 Pair the adapter in Windows first\n"
                    troubleshoot_msg += "5. 🔄 Try 'Scan Devices' again\n"
                    troubleshoot_msg += "6. 🛠️ Check Device Manager → Ports\n"
                    troubleshoot_msg += "7. 🔧 Try running as Administrator\n\n"
                    troubleshoot_msg += "📋 Manual Steps:\n"
                    troubleshoot_msg += "• Windows Settings → Devices → Bluetooth\n"
                    troubleshoot_msg += "• Add Bluetooth device\n"
                    troubleshoot_msg += "• Look for device named 'OBD' or 'ELM327'"

                    messagebox.showwarning(
                        "No Bluetooth Devices Found", troubleshoot_msg)

            except Exception as e:
                error_details = str(e)
                messagebox.showerror(
                    "Bluetooth Scan Error",
                    f"Error scanning for Bluetooth devices:\n{error_details}\n\n"
                    f"SOLUTIONS:\n"
                    f"1. Run application as Administrator\n"
                    f"2. Ensure Bluetooth is enabled in Windows\n"
                    f"3. Check Windows Bluetooth settings\n"
                    f"4. Try restarting Windows Bluetooth service\n"
                    f"5. Update Bluetooth drivers"
                )

    except Exception as e:
        messagebox.showerror(
            "Scan Error", f"Error during device scan:\n{str(e)}")
    finally:
        status_label.config(text="● Not Connected", fg="#FF4444")

# Function to toggle demo mode


def toggle_demo_mode():
    global demo_mode
    demo_mode = not demo_mode
    if demo_mode:
        demo_button.config(text="🛑 Exit Demo", bg='#FF5722')
        status_label.config(text="● Demo Mode Active", fg="#2196F3")
        print("Demo mode activated - starting updates...")
        update()  # Start updating with simulated data
    else:
        demo_button.config(text="🎯 Demo Mode", bg='#2196F3')
        status_label.config(text="● Not Connected", fg="#FF4444")
        print("Demo mode deactivated")


# Demo button (defined after toggle function)
demo_button = tk.Button(button_frame, text="🎯 Demo Mode",
                        font=('Segoe UI', 10, 'bold'), bg='#2196F3', fg='white',
                        relief='flat', bd=0, padx=20, pady=8,
                        command=toggle_demo_mode)
demo_button.pack(fill='x')

# Bluetooth setup instructions in connection tab
bt_frame = tk.LabelFrame(conn_main_frame, text="Bluetooth Setup Instructions",
                         font=('Segoe UI', 11, 'bold'), fg='#2196F3', bg='#2b2b2b',
                         relief='flat', bd=2)
bt_frame.pack(fill='x', pady=(0, 10))

bt_instructions = [
    "🔧 Setup Instructions:",
    "1. Install: pip install bleak",
    "2. Pair OBD adapter in Windows Bluetooth settings",
    "3. Note the assigned COM port (e.g., COM4)",
    "4. Select 'Bluetooth' connection type",
    "5. Enter COM port or click 'Scan Devices'",
    "",
    "💡 Tips for Success:",
    "• Use COM port instead of MAC address when possible",
    "• Check Device Manager → Ports for Bluetooth COM ports",
    "• Ensure adapter supports ELM327 protocol",
    "• Close other OBD software before connecting",
    "• Try restarting Bluetooth service if issues persist"
]

for instruction in bt_instructions:
    style = 'bold' if instruction.startswith(
        '🔧') or instruction.startswith('💡') else 'normal'
    color = '#4CAF50' if instruction.startswith(
        '🔧') else '#2196F3' if instruction.startswith('💡') else '#BBBBBB'
    tk.Label(bt_frame, text=instruction, font=('Segoe UI', 8, style),
             fg=color, bg='#2b2b2b', anchor='w').pack(fill='x', padx=10, pady=1)

# Add helper information for Bluetooth troubleshooting
help_frame = tk.LabelFrame(conn_main_frame, text="Bluetooth Troubleshooting",
                           font=('Segoe UI', 11, 'bold'), fg='#FF9800', bg='#2b2b2b',
                           relief='flat', bd=2)
help_frame.pack(fill='x', pady=(0, 10))

help_text = [
    "🔍 Finding Your OBD Adapter:",
    "• Windows Settings → Devices → Bluetooth → Check paired devices",
    "• Device Manager → Ports → Look for 'Standard Serial over Bluetooth'",
    "• Use 'Scan Devices' button to find COM ports automatically",
    "",
    "⚠️ Common Issues:",
    "• Adapter not responding: Try power cycling the adapter",
    "• Connection timeout: Ensure vehicle is running (powers adapter)",
    "• COM port busy: Close Torque, OBD Auto Doctor, or similar apps",
    "• Pairing failed: Remove device and pair again in Windows settings"
]

for help_item in help_text:
    style = 'bold' if help_item.startswith(
        '🔍') or help_item.startswith('⚠️') else 'normal'
    color = '#FF9800' if help_item.startswith(
        '🔍') or help_item.startswith('⚠️') else '#BBBBBB'
    tk.Label(help_frame, text=help_item, font=('Segoe UI', 8, style),
             fg=color, bg='#2b2b2b', anchor='w').pack(fill='x', padx=10, pady=1)

# Global variables for PID monitoring and logging
available_pids = {}
pid_monitor_window = None
pid_logger_window = None
pid_monitoring_active = False
monitored_pids = []
pid_value_labels = {}

# Enhanced connection health monitoring for OBDXPROVX


def check_obdxprovx_connection_health(connection):
    """Comprehensive connection health check specifically for OBDXPROVX"""
    if not connection:
        return False, "No connection object"

    try:
        # Basic connection check
        if not connection.is_connected():
            return False, "Connection not established"

        # Try a simple query to test responsiveness
        test_start = time.time()
        try:
            # Use a lightweight command that most vehicles support
            test_cmd = getattr(obd.commands, 'RPM', None) or getattr(
                obd.commands, 'ENGINE_LOAD', None) or getattr(obd.commands, 'SPEED', None)
            if test_cmd:
                test_response = connection.query(test_cmd, force=True)
                test_duration = time.time() - test_start

                # Check response validity and timing
                if test_response is None:
                    return False, "No response to test query"
                elif test_response.is_null():
                    return False, "Null response - adapter may be disconnected"
                elif test_duration > 6:  # Reduced threshold for faster detection
                    return False, f"Slow response ({test_duration:.1f}s) - connection degrading"
                elif test_duration > 3:  # Warning for moderate delays
                    return True, f"Slow but OK (response in {test_duration:.1f}s)"
                else:
                    return True, f"Healthy (response in {test_duration:.1f}s)"
            else:
                # If no test command available, just check connection state
                return True, "Connection state OK"

        except Exception as query_error:
            error_msg = str(query_error).lower()
            if any(keyword in error_msg for keyword in ['timeout', 'connection', 'serial', 'bluetooth', 'device', 'port']):
                return False, f"Connection error: {str(query_error)[:30]}"
            else:
                # Non-connection error might just be unsupported command
                return True, "Connection OK (test command unsupported)"

    except Exception as health_error:
        return False, f"Health check failed: {str(health_error)[:30]}"

# Function to scan for available PIDs


def scan_available_pids():
    """Robust PID scanner optimized for Bluetooth adapters like OBDXPROVX"""
    global available_pids, connection

    if not connection or not connection.is_connected():
        messagebox.showerror("Connection Error",
                             "Please connect to a vehicle first!")
        return

    try:
        # Pre-scan connection health check
        pid_status_label.config(
            text="🔍 Checking OBDXPROVX connection health...", fg="#FFA726")

        is_healthy, health_msg = check_obdxprovx_connection_health(connection)
        print(f"🔵 OBDXPROVX Health Check: {health_msg}")

        if not is_healthy:
            error_msg = f"🔵 OBDXPROVX Connection Health Check Failed\n\n"
            error_msg += f"❌ Issue: {health_msg}\n\n"
            error_msg += f"🔧 SOLUTIONS:\n"
            error_msg += f"• Check vehicle engine is running (not just ignition)\n"
            error_msg += f"• Verify OBDXPROVX LED is solid blue\n"
            error_msg += f"• Try disconnecting and reconnecting\n"
            error_msg += f"• Use 'COM6 (OBDX)' quick connect button\n"
            error_msg += f"• Close other OBD software if running\n"
            error_msg += f"• Check vehicle OBD port connection\n\n"
            error_msg += f"💡 A healthy connection is important for stable PID scanning!"

            result = messagebox.askyesno("OBDXPROVX Health Warning",
                                         error_msg + "\n\nContinue scan anyway?")
            if not result:
                pid_status_label.config(
                    text="❌ Scan cancelled - fix connection first", fg="#FF4444")
                return

        # Run quick stability test before full scan
        pid_status_label.config(
            text="🔄 Running OBDXPROVX stability test...", fg="#FFA726")

        print(f"🔵 Running pre-scan stability verification...")
        stability_passed = 0
        stability_tests = 10

        for test_num in range(1, stability_tests + 1):
            is_stable, stability_msg = check_obdxprovx_connection_health(
                connection)
            if is_stable:
                stability_passed += 1
            print(
                f"🔍 Stability test {test_num}/{stability_tests}: {stability_msg}")
            time.sleep(0.5)  # Brief pause between tests

        stability_rate = (stability_passed / stability_tests) * 100
        print(
            f"🔵 Pre-scan stability: {stability_passed}/{stability_tests} ({stability_rate:.1f}%)")

        if stability_rate < 70:
            warning_msg = f"⚠️ OBDXPROVX Stability Warning\n\n"
            warning_msg += f"📊 Stability Test Results:\n"
            warning_msg += f"• Passed: {stability_passed}/{stability_tests} tests\n"
            warning_msg += f"• Success rate: {stability_rate:.1f}%\n\n"
            warning_msg += f"🔧 RECOMMENDATIONS:\n"
            warning_msg += f"• Connection is unstable for PID scanning\n"
            warning_msg += f"• Check OBDXPROVX is securely connected\n"
            warning_msg += f"• Verify vehicle engine is fully warmed up\n"
            warning_msg += f"• Close other OBD applications\n"
            warning_msg += f"• Try repositioning closer to vehicle\n\n"
            warning_msg += f"Continue with potentially unstable connection?"

            result = messagebox.askyesno(
                "OBDXPROVX Stability Warning", warning_msg)
            if not result:
                pid_status_label.config(
                    text="❌ Scan cancelled - improve stability first", fg="#FF4444")
                return
        else:
            print(f"✅ OBDXPROVX stability test passed - ready for PID scanning!")

        pid_status_label.config(
            text="🔍 Starting OBDXPROVX-optimized PID scan...", fg="#FFA726")
        available_pids.clear()

        # Priority PIDs - most commonly used ones first
        priority_pids = [
            'RPM', 'SPEED', 'ENGINE_LOAD', 'COOLANT_TEMP', 'INTAKE_TEMP', 'MAF',
            'THROTTLE_POS', 'FUEL_PRESSURE', 'INTAKE_PRESSURE', 'O2_B1S1',
            'FUEL_LEVEL', 'BAROMETRIC_PRESSURE', 'AMBIENT_AIR_TEMP', 'FUEL_TRIM_SHORT_B1',
            'FUEL_TRIM_LONG_B1', 'ENGINE_TIME', 'DISTANCE_W_MIL', 'CATALYST_TEMP_B1S1'
        ]

        # Get all available OBD commands efficiently
        all_commands = []
        for cmd_name in dir(obd.commands):
            if not cmd_name.startswith('_'):
                try:
                    cmd = getattr(obd.commands, cmd_name)
                    if cmd is not None and hasattr(cmd, 'pid'):
                        all_commands.append(cmd)
                except:
                    continue

        # Smart PID scanning - focus on commonly supported PIDs first
        # Most vehicles only support 20-80 PIDs out of 300+ total possible PIDs
        
        # Common PIDs that most vehicles support (Mode 01)
        commonly_supported_pids = [
            'RPM', 'SPEED', 'ENGINE_LOAD', 'COOLANT_TEMP', 'INTAKE_TEMP', 'MAF',
            'THROTTLE_POS', 'FUEL_PRESSURE', 'INTAKE_PRESSURE', 'O2_B1S1',
            'FUEL_LEVEL', 'BAROMETRIC_PRESSURE', 'AMBIENT_AIR_TEMP', 'FUEL_TRIM_SHORT_B1',
            'FUEL_TRIM_LONG_B1', 'ENGINE_TIME', 'DISTANCE_W_MIL', 'CATALYST_TEMP_B1S1',
            'FUEL_TRIM_SHORT_B2', 'FUEL_TRIM_LONG_B2', 'FUEL_RAIL_PRESSURE_VAC',
            'FUEL_RAIL_PRESSURE_DIRECT', 'O2_B1S2', 'O2_B2S1', 'O2_B2S2',
            'COMMANDED_EGR', 'EGR_ERROR', 'COMMANDED_EVAP_PURGE', 'FUEL_TANK_LEVEL',
            'WARMUPS_SINCE_DTC_CLEAR', 'DISTANCE_SINCE_DTC_CLEAR', 'EVAP_VAPOR_PRESSURE',
            'COMMANDED_INTAKE_PRESSURE', 'TIMING_ADVANCE', 'RUN_TIME_MIL',
            'FUEL_TYPE', 'ETHANOL_PERCENT', 'FUEL_RAIL_PRESSURE_ABS'
        ]
        
        # Separate commands into common vs extended
        common_commands = []
        priority_commands = []
        extended_commands = []

        for cmd in all_commands:
            if cmd.name in priority_pids:
                priority_commands.append(cmd)
            elif cmd.name in commonly_supported_pids:
                common_commands.append(cmd)
            else:
                extended_commands.append(cmd)

        # Smart scanning strategy: Priority → Common → Ask user about extended scan
        print(f"🔍 SMART SCAN MODE:")
        print(f"  • Priority PIDs: {len(priority_commands)}")
        print(f"  • Common PIDs: {len(common_commands)}")  
        print(f"  • Extended PIDs: {len(extended_commands)}")
        print(f"  • Total available: {len(all_commands)}")
        
        # Start with priority + common PIDs (most likely to be supported)
        obd_commands = priority_commands + common_commands
        
        # Keep extended commands for optional scanning
        remaining_commands = extended_commands

        total_commands = len(obd_commands)
        scanned_count = 0
        working_pids = 0
        failed_reads = 0
        consecutive_failures = 0
        max_consecutive_failures = 5  # Increased tolerance for better coverage
        connection_lost_count = 0
        max_connection_errors = 4  # Slightly reduced for faster recovery
        last_successful_query = time.time()  # Track time since last successful query

        # Adaptive scanning - if connection is stable, we can scan more PIDs
        connection_stability_score = 0  # Track connection quality

        print(
            f"� FULL SCAN: Scanning all {total_commands} PIDs (priority first, then comprehensive scan)...")

        # Full PID scanning - no batch expansion needed since we're scanning everything

        for cmd in obd_commands:
            try:
                scanned_count += 1
                progress = int((scanned_count / total_commands) * 100)

                # Update UI less frequently for speed
                if scanned_count % 3 == 0 or scanned_count <= 5:  # More frequent updates for user feedback
                    pid_status_label.config(
                        text=f"🔍 Scanning... {progress}% ({working_pids} found)")
                    try:
                        window.update_idletasks()
                    except tk.TclError:
                        print("Window closed during scan")
                        return

                # Enhanced connection check every 2 commands for maximum Bluetooth stability
                if scanned_count % 2 == 0:
                    # Check if too much time has passed since last successful query (timeout detection)
                    time_since_success = time.time() - last_successful_query
                    if time_since_success > 20:  # Reduced to 20 seconds for faster detection
                        print(
                            f"⚠️ No successful queries for {time_since_success:.1f}s - connection may be stale")
                        connection_lost_count += 1

                    # Comprehensive connection health check
                    if not connection or not connection.is_connected():
                        connection_lost_count += 1
                        print(
                            f"⚠️ Basic connection check failed (attempt {connection_lost_count})")
                    else:
                        # More frequent deep health check every 3 commands (was 6)
                        if scanned_count % 3 == 0:
                            is_healthy, health_msg = check_obdxprovx_connection_health(
                                connection)
                            print(f"🔍 Health check: {health_msg}")
                            if not is_healthy:
                                connection_lost_count += 1
                                print(f"⚠️ Health check failed: {health_msg}")

                        # Add brief stability pause every 5 successful checks
                        if scanned_count % 10 == 0 and connection_lost_count == 0:
                            print(
                                f"💙 OBDXPROVX stability pause at {progress}%...")
                            # Give OBDXPROVX a moment to stabilize
                            time.sleep(1.5)

                    if connection_lost_count > 0:
                        # More aggressive recovery attempts
                        if connection_lost_count <= 6:  # Increased recovery attempts
                            print(
                                f"🔵 Attempting OBDXPROVX connection recovery #{connection_lost_count}...")
                            pid_status_label.config(
                                text=f"🔄 OBDXPROVX reconnecting... {progress}%", fg="#FF9800")

                            try:
                                # Close current connection more thoroughly
                                if connection:
                                    try:
                                        connection.close()
                                        # Force garbage collection to clean up connection
                                        import gc
                                        gc.collect()
                                    except:
                                        pass
                                    connection = None

                                # Progressive pause - longer for repeated failures
                                pause_time = min(
                                    3.0 + (connection_lost_count * 0.5), 6.0)
                                print(
                                    f"🔄 Recovery pause: {pause_time:.1f}s...")
                                time.sleep(pause_time)

                                # Try multiple connection strategies
                                # Primary OBDXPROVX ports first, then fallbacks
                                ports_to_try = ['COM6', 'COM5', 'COM7', 'COM3']
                                # Try both common rates
                                baud_rates = [38400, 9600]

                                for port in ports_to_try:
                                    for baud in baud_rates:
                                        try:
                                            print(
                                                f"🔄 Trying {port} at {baud} baud...")
                                            # Use longer timeout for recovery connections
                                            connection = obd.OBD(
                                                port, baudrate=baud, timeout=25, fast=False)
                                            if connection and connection.is_connected():
                                                # Validate the recovered connection with a test query
                                                try:
                                                    test_cmd = getattr(
                                                        obd.commands, 'RPM', None)
                                                    if test_cmd:
                                                        test_response = connection.query(
                                                            test_cmd, force=True)
                                                        if test_response and not test_response.is_null():
                                                            print(
                                                                f"✅ OBDXPROVX reconnected and validated on {port} at {baud} baud!")
                                                            connection_lost_count = 0  # Reset counter
                                                            last_successful_query = time.time()  # Reset timer
                                                            consecutive_failures = 0  # Reset failure counter
                                                            pid_status_label.config(
                                                                text=f"🔍 Scanning... {progress}% ({working_pids} found)")
                                                            # Brief stabilization pause after recovery
                                                            time.sleep(1.0)
                                                            break
                                                        else:
                                                            print(
                                                                f"⚠️ {port}@{baud} connected but failed test query")
                                                            connection.close()
                                                            connection = None
                                                    else:
                                                        print(
                                                            f"✅ OBDXPROVX reconnected on {port} at {baud} baud (no test available)!")
                                                        connection_lost_count = 0  # Reset counter
                                                        last_successful_query = time.time()  # Reset timer
                                                        consecutive_failures = 0  # Reset failure counter
                                                        pid_status_label.config(
                                                            text=f"🔍 Scanning... {progress}% ({working_pids} found)")
                                                        time.sleep(1.0)
                                                        break
                                                except Exception as test_error:
                                                    print(
                                                        f"⚠️ Test query failed on {port}@{baud}: {test_error}")
                                                    if connection:
                                                        connection.close()
                                                        connection = None
                                        except Exception as port_error:
                                            print(
                                                f"❌ {port}@{baud} failed: {str(port_error)[:50]}")
                                            continue
                                    if connection and connection.is_connected():
                                        break

                                if not connection or not connection.is_connected():
                                    print(
                                        f"❌ Recovery attempt #{connection_lost_count} failed")

                            except Exception as recovery_error:
                                print(
                                    f"❌ Recovery attempt failed: {recovery_error}")

                        if connection_lost_count >= max_connection_errors:
                            print("❌ Connection persistently lost during PID scan")
                            pid_status_label.config(
                                text="❌ Connection lost - manual reconnect needed", fg="#FF4444")

                            error_msg = "🔵 OBDXPROVX Connection Persistently Lost\n\n"
                            error_msg += f"📊 Scan Progress: {progress}% complete\n"
                            error_msg += f"✅ PIDs found so far: {working_pids}\n"
                            error_msg += f"🔄 Recovery attempts: {connection_lost_count}\n\n"
                            error_msg += "🔧 MANUAL TROUBLESHOOTING REQUIRED:\n"
                            error_msg += "• Turn vehicle OFF and ON again\n"
                            error_msg += "• Unplug and replug OBDXPROVX from OBD port\n"
                            error_msg += "• Check OBDXPROVX LED: should be solid blue when connected\n"
                            error_msg += "• Close all other OBD software (Torque, OBD Fusion, etc.)\n"
                            error_msg += "• Try Windows Device Manager → disable/enable Bluetooth\n"
                            error_msg += "• Use 'COM6 (OBDX)' quick connect to reconnect\n"
                            error_msg += "• Consider running 'Stability Test' before next scan\n"
                            error_msg += "• Check vehicle is fully warmed up and running smooth\n\n"
                            error_msg += f"💡 Found {working_pids} working PIDs - these are still available!"

                            messagebox.showerror(
                                "OBDXPROVX Needs Manual Recovery", error_msg)
                            return
                        else:
                            # Progressive delay to let OBDXPROVX stabilize
                            stabilize_time = min(
                                1.0 + (connection_lost_count * 0.3), 2.5)
                            time.sleep(stabilize_time)
                            continue
                    else:
                        connection_lost_count = 0  # Reset counter on successful check

                try:
                    # Pre-query connection validation
                    if not connection or not connection.is_connected():
                        print(
                            f"⚠️ Connection invalid before {cmd.name}, skipping")
                        failed_reads += 1
                        consecutive_failures += 1
                        continue

                    # OBDXPROVX-optimized query with enhanced timeout handling
                    query_start_time = time.time()
                    try:
                        response = connection.query(cmd, force=True)
                        query_duration = time.time() - query_start_time

                        # Check for suspiciously long query times (indicates connection issues)
                        if query_duration > 8:  # Reduced threshold for faster detection
                            print(
                                f"⚠️ Slow query for {cmd.name}: {query_duration:.1f}s - connection may be degrading")
                            connection_lost_count += 1

                    except Exception as query_error:
                        query_duration = time.time() - query_start_time
                        error_msg = str(query_error).lower()

                        # Detect connection-related errors
                        if any(keyword in error_msg for keyword in ['timeout', 'connection', 'serial', 'bluetooth', 'device']):
                            print(
                                f"🔵 Connection error detected in {cmd.name}: {str(query_error)[:50]}")
                            connection_lost_count += 1
                            failed_reads += 1
                            consecutive_failures += 1

                            # If we get multiple connection errors quickly, trigger recovery
                            if connection_lost_count >= 2:
                                print(
                                    f"🔄 Multiple connection errors - triggering immediate recovery")
                                continue  # This will trigger the connection check on next iteration
                            else:
                                # Brief pause before continuing
                                time.sleep(0.5)
                                continue
                        else:
                            # Non-connection error, just continue
                            print(
                                f"⚠️ Non-connection error for {cmd.name}: {str(query_error)[:50]}")
                            failed_reads += 1
                            consecutive_failures += 1
                            response = None

                    if response and not response.is_null() and response.value is not None:
                        # Successful query - update success timer
                        last_successful_query = time.time()

                        # Extract unit information safely
                        unit_str = 'No unit'
                        try:
                            if hasattr(response.value, 'units'):
                                unit_str = str(response.value.units)
                            elif hasattr(response.value, 'unit'):
                                unit_str = str(response.value.unit)
                            elif hasattr(cmd, 'unit'):
                                unit_str = str(cmd.unit)
                        except:
                            unit_str = 'Unknown'

                        pid_info = {
                            'command': cmd,
                            'name': cmd.name,
                            'description': getattr(cmd, 'desc', 'No description'),
                            'unit': unit_str,
                            'value': response.value,
                            'pid_hex': f"0x{cmd.pid:02X}",
                            'priority': cmd.name in priority_pids
                        }
                        available_pids[cmd.name] = pid_info
                        working_pids += 1
                        consecutive_failures = 0
                        # Reduce error count on success
                        connection_lost_count = max(
                            0, connection_lost_count - 1)

                        # Show priority PIDs as they're found
                        if cmd.name in priority_pids:
                            print(
                                f"✅ Priority PID: {cmd.name} = {response.value} {unit_str}")
                        else:
                            print(f"✓ {cmd.name}")

                        # Adaptive delay based on connection stability
                        if connection_lost_count > 0:
                            # Longer delay when connection is unstable
                            time.sleep(0.5)  # Increased for better stability
                        else:
                            # Standard delay for stable connections
                            time.sleep(0.35)  # Slightly increased base delay

                    else:
                        failed_reads += 1
                        consecutive_failures += 1
                        # Adaptive delay on failures
                        if consecutive_failures > 2:
                            # Longer delay after multiple failures
                            time.sleep(0.3)
                        else:
                            # Shorter delay for occasional failures
                            time.sleep(0.15)

                except Exception as cmd_error:
                    failed_reads += 1
                    consecutive_failures += 1
                    print(
                        f"⚠️ Command processing error for {cmd.name}: {str(cmd_error)[:50]}")

                    # Enhanced error categorization
                    error_msg = str(cmd_error).lower()
                    if any(keyword in error_msg for keyword in ['connection', 'timeout', 'serial', 'bluetooth', 'device', 'port']):
                        connection_lost_count += 1
                        print(
                            f"🔵 Connection-related error detected (total: {connection_lost_count})")

                        # More aggressive response to connection errors
                        if connection_lost_count >= 3:
                            print(
                                "❌ Multiple connection errors - triggering recovery on next check")
                            # The connection check at the start of the loop will handle recovery

                    # Adaptive delay based on error type and frequency
                    if consecutive_failures > 3:
                        time.sleep(0.5)  # Longer delay after many failures
                    else:
                        time.sleep(0.2)  # Standard delay

                    # Stop scanning if too many consecutive failures indicate fundamental issues
                    if consecutive_failures >= max_consecutive_failures:
                        print(
                            f"⚠️ {consecutive_failures} consecutive failures - may indicate vehicle/adapter issues")

                        # Check if we have any successful PIDs to show
                        if working_pids > 0:
                            pid_status_label.config(
                                text=f"⚠️ Partial scan: {working_pids} PIDs found", fg="#FF9800")

                            warning_msg = f"🔵 OBDXPROVX Scan Partially Complete\n\n"
                            warning_msg += f"📊 STATUS:\n"
                            warning_msg += f"• PIDs successfully found: {working_pids}\n"
                            warning_msg += f"• Consecutive failures: {consecutive_failures}\n"
                            warning_msg += f"• Connection errors: {connection_lost_count}\n"
                            warning_msg += f"• Scan stopped at {progress}% to prevent issues\n\n"
                            warning_msg += f"💡 RECOMMENDATIONS:\n"
                            warning_msg += f"• Use the {working_pids} working PIDs found\n"
                            warning_msg += f"• Check vehicle engine is fully warmed up\n"
                            warning_msg += f"• Some vehicles need to be driving for all PIDs\n"
                            warning_msg += f"• Try again after vehicle reaches operating temperature\n"
                            warning_msg += f"• OBDXPROVX works best with engine running under load\n\n"
                            warning_msg += f"✅ You can now use PID Monitor or Data Logger!"

                            messagebox.showinfo(
                                "OBDXPROVX Partial Scan Complete", warning_msg)
                            break
                        else:
                            print(
                                f"❌ No working PIDs found after {consecutive_failures} failures")
                            break

            except KeyboardInterrupt:
                print("PID scan interrupted by user")
                break
            except Exception as cmd_error:
                print(f"Command processing error: {cmd_error}")
                failed_reads += 1
                continue

        # Update final status
        success_rate = (working_pids / total_commands *
                        100) if total_commands > 0 else 0
        
        # Offer extended scan if there are remaining PIDs and we found some working PIDs
        extended_scan_performed = False
        if remaining_commands and working_pids > 5:  # Only offer if we found some basic PIDs
            extended_msg = f"🔍 SMART SCAN COMPLETE!\n\n"
            extended_msg += f"📊 INITIAL RESULTS:\n"
            extended_msg += f"• Priority + Common PIDs tested: {total_commands}\n"
            extended_msg += f"• Working PIDs found: {working_pids}\n"
            extended_msg += f"• Success rate: {success_rate:.1f}%\n\n"
            extended_msg += f"🔄 EXTENDED SCAN AVAILABLE:\n"
            extended_msg += f"• Additional PIDs available: {len(remaining_commands)}\n"
            extended_msg += f"• These are less common/specialized PIDs\n"
            extended_msg += f"• Expected success rate: 1-5% (normal for most vehicles)\n\n"
            extended_msg += f"💡 RECOMMENDATION:\n"
            if success_rate >= 20:
                extended_msg += f"• Your vehicle has good PID support ({working_pids} PIDs found)\n"
                extended_msg += f"• Extended scan may find 5-15 additional PIDs\n"
                extended_msg += f"• This will take longer but gives maximum coverage\n\n"
                extended_msg += f"Run extended scan for maximum PID discovery?"
                
                result = messagebox.askyesno("Extended PID Scan Available", extended_msg)
                if result:
                    # Perform extended scan
                    pid_status_label.config(text="🔍 Extended scan: Testing specialized PIDs...", fg="#FFA726")
                    print(f"🔍 Starting extended PID scan: {len(remaining_commands)} additional PIDs...")
                    
                    initial_working_pids = working_pids
                    extended_scanned = 0
                    
                    for cmd in remaining_commands:
                        try:
                            extended_scanned += 1
                            progress = int((extended_scanned / len(remaining_commands)) * 100)
                            
                            if extended_scanned % 5 == 0:
                                pid_status_label.config(
                                    text=f"🔍 Extended scan... {progress}% ({working_pids - initial_working_pids} new found)")
                                try:
                                    window.update_idletasks()
                                except tk.TclError:
                                    break
                                    
                            # Quick test without extensive error handling for extended scan
                            if connection and connection.is_connected():
                                try:
                                    response = connection.query(cmd, force=True)
                                    if response and not response.is_null() and response.value is not None:
                                        # Extract unit information
                                        unit_str = 'No unit'
                                        try:
                                            if hasattr(response.value, 'units'):
                                                unit_str = str(response.value.units)
                                            elif hasattr(response.value, 'unit'):
                                                unit_str = str(response.value.unit)
                                            elif hasattr(cmd, 'unit'):
                                                unit_str = str(cmd.unit)
                                        except:
                                            unit_str = 'Unknown'

                                        pid_info = {
                                            'command': cmd,
                                            'name': cmd.name,
                                            'description': getattr(cmd, 'desc', 'No description'),
                                            'unit': unit_str,
                                            'value': response.value,
                                            'pid_hex': f"0x{cmd.pid:02X}",
                                            'priority': False
                                        }
                                        available_pids[cmd.name] = pid_info
                                        working_pids += 1
                                        print(f"✓ Extended: {cmd.name}")
                                    
                                    time.sleep(0.2)  # Shorter delay for extended scan
                                except:
                                    pass  # Ignore failures in extended scan
                        except:
                            pass
                    
                    extended_scan_performed = True
                    extended_found = working_pids - initial_working_pids
                    total_commands += len(remaining_commands)
                    failed_reads += (len(remaining_commands) - extended_found)
                    success_rate = (working_pids / total_commands * 100) if total_commands > 0 else 0
                    
                    print(f"🔍 Extended scan complete: {extended_found} additional PIDs found")
            else:
                extended_msg += f"• Limited PID support detected ({working_pids} PIDs)\n"
                extended_msg += f"• Extended scan likely to find few additional PIDs\n"
                extended_msg += f"• Recommend using the {working_pids} PIDs already found\n\n"
                extended_msg += f"Continue with current results?"
                
                messagebox.showinfo("Extended Scan Not Recommended", extended_msg)

        pid_status_label.config(
            text=f"✅ Found {working_pids} PIDs ({success_rate:.1f}% success)", fg="#4CAF50")

        # Show comprehensive results with OBDXPROVX-specific messaging
        if extended_scan_performed:
            result_msg = f"🔵 OBDXPROVX Smart + Extended PID Scan Complete!\n\n"
            result_msg += f"📊 COMPREHENSIVE SCAN RESULTS:\n"
            result_msg += f"• Total PIDs tested: {total_commands} (smart + extended scan)\n"
        else:
            result_msg = f"🔵 OBDXPROVX Smart PID Scan Complete!\n\n"
            result_msg += f"📊 SMART SCAN RESULTS:\n"
            result_msg += f"• PIDs tested: {total_commands} (priority + commonly supported PIDs)\n"
            
        result_msg += f"• Working PIDs found: {working_pids}\n"
        result_msg += f"• Failed/Unsupported: {failed_reads}\n"
        result_msg += f"• Success rate: {success_rate:.1f}%\n"
        result_msg += f"• Connection errors: {connection_lost_count}\n\n"
        
        # Add context about success rates
        if success_rate >= 30:
            result_msg += f"🎉 EXCELLENT RESULTS!\n"
            result_msg += f"• Your vehicle has very good PID support\n"
            result_msg += f"• {success_rate:.1f}% is well above average (20-30%)\n\n"
        elif success_rate >= 15:
            result_msg += f"✅ GOOD RESULTS!\n"
            result_msg += f"• Your vehicle has good PID support\n"
            result_msg += f"• {success_rate:.1f}% is above average for most vehicles\n\n"
        elif success_rate >= 8:
            result_msg += f"👍 NORMAL RESULTS!\n"
            result_msg += f"• Your vehicle has typical PID support\n"
            result_msg += f"• {success_rate:.1f}% is normal for many vehicles\n\n"
        else:
            result_msg += f"⚠️ LIMITED RESULTS\n"
            result_msg += f"• Your vehicle has basic PID support\n"
            result_msg += f"• {success_rate:.1f}% suggests older or limited OBD implementation\n\n"

        if working_pids > 0:
            result_msg += f"🎯 PRIORITY PIDS FOUND:\n"
            priority_found = [
                name for name, info in available_pids.items() if info.get('priority', False)]
            if priority_found:
                result_msg += f"• {', '.join(priority_found[:10])}\n"
                if len(priority_found) > 10:
                    result_msg += f"• ...and {len(priority_found) - 10} more priority PIDs\n"
            else:
                result_msg += f"• No priority PIDs found\n"

            result_msg += f"\n💡 OBDXPROVX Ready! You can now:\n"
            result_msg += f"• Use 'Monitor PIDs' to view real-time data\n"
            result_msg += f"• Use 'Data Logger' for CSV export\n"
            result_msg += f"• View all {working_pids} PIDs in the monitor window\n"
            result_msg += f"• OBDXPROVX works best with engine running!\n\n"

            if extended_scan_performed:
                result_msg += f"✅ COMPREHENSIVE SCAN COMPLETE:\n"
                result_msg += f"• All {total_commands} available PIDs were tested\n"
                result_msg += f"• Maximum possible coverage achieved\n"
                result_msg += f"• {working_pids} PIDs are supported by your vehicle\n"
            else:
                result_msg += f"✅ SMART SCAN COMPLETE:\n"
                result_msg += f"• Priority and common PIDs tested for optimal efficiency\n"
                result_msg += f"• {working_pids} PIDs found with {success_rate:.1f}% success rate\n"
                if len(remaining_commands) > 0:
                    result_msg += f"• {len(remaining_commands)} specialized PIDs available for extended scan\n"
        else:
            result_msg += f"❌ No working PIDs found!\n\n"
            result_msg += f"🔧 OBDXPROVX TROUBLESHOOTING:\n"
            result_msg += f"• Ensure vehicle engine is running (not just ignition)\n"
            result_msg += f"• Check OBDXPROVX LED is solid blue (connected)\n"
            result_msg += f"• Verify OBDXPROVX is securely in OBD port\n"
            result_msg += f"• Try reconnecting: use 'COM6 (OBDX)' button\n"
            result_msg += f"• Close Torque, OBD Fusion, or other OBD apps\n"
            result_msg += f"• Try 'Demo Mode' to test interface functionality\n"
            result_msg += f"• Some vehicles require drive for full PID support"

        scan_title = "OBDXPROVX PID Scan Results" if working_pids > 0 else "OBDXPROVX Scan Issues"
        messagebox.showinfo(scan_title, result_msg)

        print(f"\n=== SMART PID SCAN SUMMARY ===")
        if extended_scan_performed:
            print(f"Total tested: {total_commands} (smart + extended scan)")
        else:
            print(f"Total tested: {total_commands} (priority + common PIDs)")
        print(f"Working PIDs: {working_pids}")
        print(f"Failed reads: {failed_reads}")
        print(f"Success rate: {success_rate:.1f}%")
        if available_pids:
            priority_count = len(
                [p for p in available_pids.values() if p.get('priority', False)])
            print(
                f"Priority PIDs found: {priority_count}/{len(priority_pids)}")
            print(
                f"Available for monitoring: {list(available_pids.keys())[:10]}...")
            if extended_scan_performed:
                print(f"COMPREHENSIVE SCAN: Maximum coverage achieved - {working_pids} PIDs supported by vehicle")
            else:
                print(f"SMART SCAN: Efficient discovery - {working_pids} PIDs found with {success_rate:.1f}% success rate")
                if len(remaining_commands) > 0:
                    print(f"Extended scan available: {len(remaining_commands)} additional specialized PIDs")

    except Exception as e:
        pid_status_label.config(text="❌ OBDXPROVX scan failed", fg="#FF4444")
        error_msg = f"🔵 OBDXPROVX PID Scan Error:\n{str(e)}\n\n"
        error_msg += f"🔧 SOLUTIONS:\n"
        error_msg += f"• Check OBDXPROVX connection to vehicle\n"
        error_msg += f"• Ensure vehicle engine is running\n"
        error_msg += f"• Try reconnecting: 'COM6 (OBDX)' button\n"
        error_msg += f"• Close other OBD software (Torque, etc.)\n"
        error_msg += f"• Check OBDXPROVX LED status\n"
        error_msg += f"• Try 'Demo Mode' to test functionality\n"
        error_msg += f"• Restart application if needed\n\n"
        error_msg += f"💡 OBDXPROVX works best with vehicle fully warmed up!"

        messagebox.showerror("OBDXPROVX Scan Error", error_msg)
        print(f"🔵 OBDXPROVX scan error: {e}")
        traceback.print_exc()


# Function to show PID monitor window


def show_pid_monitor():
    """Show the PID monitoring window"""
    global pid_monitor_window, pid_monitoring_active

    if not available_pids:
        messagebox.showwarning("No PIDs Available",
                               "Please scan for PIDs first using the 'Scan PIDs' button.")
        return

    # Close existing window if open
    if pid_monitor_window:
        pid_monitor_window.destroy()

    # Create new monitor window
    pid_monitor_window = tk.Toplevel(window)
    pid_monitor_window.title("OBD PID Monitor")
    pid_monitor_window.configure(bg='#2b2b2b')
    pid_monitor_window.geometry("800x600")

    # Create main frame
    main_frame = tk.Frame(pid_monitor_window, bg='#2b2b2b', padx=20, pady=15)
    main_frame.pack(fill='both', expand=True)

    # Title
    title_label = tk.Label(main_frame, text="OBD PID Monitor",
                           font=('Segoe UI', 14, 'bold'), fg='#FF9800', bg='#2b2b2b')
    title_label.pack(pady=(0, 10))

    # Create left frame for PID list
    left_frame = tk.Frame(main_frame, bg='#2b2b2b')
    left_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))

    # Available PIDs section
    pids_label = tk.Label(left_frame, text=f"Available PIDs ({len(available_pids)})",
                          font=('Segoe UI', 11, 'bold'), fg='#4CAF50', bg='#2b2b2b')
    pids_label.pack(anchor='w', pady=(0, 5))

    # Create scrollable frame for PID list
    canvas = tk.Canvas(left_frame, bg='#333333', highlightthickness=0)
    scrollbar = tk.Scrollbar(
        left_frame, orient='vertical', command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg='#333333')

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # PID checkboxes
    pid_vars = {}
    for pid_name, pid_info in available_pids.items():
        frame = tk.Frame(scrollable_frame, bg='#333333')
        frame.pack(fill='x', pady=2, padx=5)

        var = tk.BooleanVar()
        pid_vars[pid_name] = var

        checkbox = tk.Checkbutton(frame, variable=var, bg='#333333', fg='white',
                                  selectcolor='#404040', font=('Segoe UI', 9))
        checkbox.pack(side='left')

        info_text = f"{pid_name} ({pid_info['pid_hex']}) - {pid_info['unit']}"
        label = tk.Label(frame, text=info_text, font=('Segoe UI', 9),
                         fg='white', bg='#333333', anchor='w')
        label.pack(side='left', fill='x', expand=True)

    # Right frame for monitoring
    right_frame = tk.Frame(main_frame, bg='#2b2b2b', width=300)
    right_frame.pack(side='right', fill='y')
    right_frame.pack_propagate(False)

    # Monitor controls
    control_label = tk.Label(right_frame, text="Monitor Control",
                             font=('Segoe UI', 11, 'bold'), fg='#9C27B0', bg='#2b2b2b')
    control_label.pack(pady=(0, 10))

    # Start/Stop monitoring button
    monitor_button = tk.Button(right_frame, text="🟢 Start Monitoring",
                               font=('Segoe UI', 10, 'bold'), bg='#4CAF50', fg='white',
                               relief='flat', bd=0, padx=20, pady=8)
    monitor_button.pack(pady=5)

    # Selected PIDs display area
    selected_label = tk.Label(right_frame, text="Monitored PIDs:",
                              font=('Segoe UI', 10, 'bold'), fg='#4CAF50', bg='#2b2b2b')
    selected_label.pack(pady=(20, 5), anchor='w')

    # Scrollable area for monitored PID values
    monitor_canvas = tk.Canvas(
        right_frame, bg='#1e1e1e', height=300, highlightthickness=0)
    monitor_scrollbar = tk.Scrollbar(
        right_frame, orient='vertical', command=monitor_canvas.yview)
    monitor_frame = tk.Frame(monitor_canvas, bg='#1e1e1e')

    monitor_frame.bind(
        "<Configure>",
        lambda e: monitor_canvas.configure(
            scrollregion=monitor_canvas.bbox("all"))
    )

    monitor_canvas.create_window((0, 0), window=monitor_frame, anchor="nw")
    monitor_canvas.configure(yscrollcommand=monitor_scrollbar.set)

    monitor_canvas.pack(side="left", fill="both", expand=True)
    monitor_scrollbar.pack(side="right", fill="y")

    # PID value labels for monitoring
    local_pid_value_labels = {}

    def toggle_monitoring():
        global pid_monitoring_active, monitored_pids, pid_value_labels

        if not pid_monitoring_active:
            # Start monitoring
            selected_pids = [pid for pid, var in pid_vars.items() if var.get()]

            if not selected_pids:
                messagebox.showwarning(
                    "No PIDs Selected", "Please select at least one PID to monitor.")
                return

            monitored_pids = selected_pids
            pid_monitoring_active = True
            monitor_button.config(text="🔴 Stop Monitoring", bg='#FF5722')

            # Clear previous labels
            for widget in monitor_frame.winfo_children():
                widget.destroy()
            local_pid_value_labels.clear()
            pid_value_labels.clear()

            # Create labels for selected PIDs
            for pid in monitored_pids:
                pid_frame = tk.Frame(monitor_frame, bg='#1e1e1e')
                pid_frame.pack(fill='x', pady=2, padx=5)

                name_label = tk.Label(pid_frame, text=f"{pid}:",
                                      font=('Segoe UI', 9, 'bold'),
                                      fg='#FF9800', bg='#1e1e1e', anchor='w', width=15)
                name_label.pack(side='left')

                value_label = tk.Label(pid_frame, text="Waiting...",
                                       font=('Segoe UI', 9),
                                       fg='white', bg='#404040', relief='flat',
                                       anchor='center', width=15)
                value_label.pack(side='right')

                local_pid_value_labels[pid] = value_label
                pid_value_labels[pid] = value_label

            # Start monitoring updates
            update_pid_monitoring()

        else:
            # Stop monitoring
            pid_monitoring_active = False
            monitored_pids = []
            monitor_button.config(text="🟢 Start Monitoring", bg='#4CAF50')

    monitor_button.config(command=toggle_monitoring)

    # Handle window close
    def on_close():
        global pid_monitoring_active, pid_monitor_window
        pid_monitoring_active = False
        if pid_monitor_window:
            pid_monitor_window.destroy()
        pid_monitor_window = None

    pid_monitor_window.protocol("WM_DELETE_WINDOW", on_close)

# Function to update PID monitoring


def update_pid_monitoring():
    """Update monitored PID values with robust error handling"""
    global pid_monitoring_active

    if not pid_monitoring_active or not pid_monitor_window or not monitored_pids:
        return

    # Check connection status
    if not connection or not connection.is_connected():
        print("Connection lost during PID monitoring")
        pid_monitoring_active = False

        # Update all labels to show connection lost
        for pid_name in monitored_pids:
            if pid_name in pid_value_labels:
                pid_value_labels[pid_name].config(
                    text="Disconnected", fg='#FF4444')
        return

    consecutive_errors = 0
    max_errors = 3  # Stop monitoring if too many consecutive errors

    try:
        for pid_name in monitored_pids:
            if not pid_monitoring_active:  # Check if monitoring was stopped
                break

            if pid_name in available_pids:
                pid_info = available_pids[pid_name]
                cmd = pid_info['command']

                try:
                    # Query with timeout protection
                    response = connection.query(cmd, force=False)

                    if response and not response.is_null() and response.value is not None:
                        # Format the value nicely
                        try:
                            if hasattr(response.value, 'magnitude'):
                                value_str = f"{response.value.magnitude:.2f}"
                            elif hasattr(response.value, '__float__'):
                                value_str = f"{float(response.value):.2f}"
                            else:
                                value_str = str(response.value)

                            # Update the label
                            if pid_name in pid_value_labels:
                                pid_value_labels[pid_name].config(
                                    text=value_str, fg='#4CAF50')
                                consecutive_errors = 0  # Reset error counter on success
                        except Exception as format_error:
                            print(
                                f"Error formatting value for {pid_name}: {format_error}")
                            if pid_name in pid_value_labels:
                                pid_value_labels[pid_name].config(
                                    text="Format error", fg='#FFA726')
                    else:
                        if pid_name in pid_value_labels:
                            pid_value_labels[pid_name].config(
                                text="No data", fg='#FF4444')

                except Exception as query_error:
                    error_msg = str(query_error).lower()

                    if "timeout" in error_msg or "read" in error_msg:
                        if pid_name in pid_value_labels:
                            pid_value_labels[pid_name].config(
                                text="Timeout", fg='#FFA726')
                        consecutive_errors += 1
                    elif "permission" in error_msg:
                        if pid_name in pid_value_labels:
                            pid_value_labels[pid_name].config(
                                text="Access denied", fg='#FF4444')
                    else:
                        if pid_name in pid_value_labels:
                            pid_value_labels[pid_name].config(
                                text="Error", fg='#FF4444')
                        consecutive_errors += 1

                    print(f"PID query error for {pid_name}: {query_error}")

                    # If too many consecutive errors, something is wrong
                    if consecutive_errors >= max_errors:
                        print(
                            f"Too many consecutive PID monitoring errors ({consecutive_errors}), checking connection...")

                        # Test connection
                        try:
                            if not connection.is_connected():
                                print("Connection lost, stopping PID monitoring")
                                pid_monitoring_active = False
                                return
                        except:
                            print(
                                "Cannot check connection status, stopping PID monitoring")
                            pid_monitoring_active = False
                            return

    except Exception as monitoring_error:
        print(f"PID monitoring error: {monitoring_error}")

        # Update all labels to show error state
        for pid_name in monitored_pids:
            if pid_name in pid_value_labels:
                pid_value_labels[pid_name].config(
                    text="Monitor error", fg='#FF4444')

    # Schedule next update if monitoring is still active
    if pid_monitoring_active and pid_monitor_window:
        try:
            # Slightly slower update rate for stability
            pid_monitor_window.after(200, update_pid_monitoring)
        except tk.TclError:
            # Window was destroyed
            pid_monitoring_active = False


# Comprehensive PID Logger
def create_pid_logger():
    """Create a comprehensive PID logging window with CSV export"""
    global pid_logger_window

    if not available_pids:
        messagebox.showwarning("No PIDs Available",
                               "Please scan for PIDs first using the 'Scan PIDs' button.")
        return

    # Close existing window if open
    if pid_logger_window:
        pid_logger_window.destroy()

    # Create logger window
    pid_logger_window = tk.Toplevel(window)
    pid_logger_window.title("PID Data Logger & Analyzer")
    pid_logger_window.configure(bg='#2b2b2b')
    pid_logger_window.geometry("1000x700")

    # Main frame
    main_frame = tk.Frame(pid_logger_window, bg='#2b2b2b', padx=20, pady=15)
    main_frame.pack(fill='both', expand=True)

    # Title
    title_label = tk.Label(main_frame, text="📊 Comprehensive PID Data Logger",
                           font=('Segoe UI', 16, 'bold'), fg='#2196F3', bg='#2b2b2b')
    title_label.pack(pady=(0, 15))

    # Create top frame for controls
    control_frame = tk.Frame(main_frame, bg='#2b2b2b')
    control_frame.pack(fill='x', pady=(0, 15))

    # Logging controls
    log_frame = tk.LabelFrame(control_frame, text="Logging Control",
                              font=('Segoe UI', 11, 'bold'), fg='#2196F3', bg='#2b2b2b')
    log_frame.pack(side='left', fill='y', padx=(0, 10))

    # Create frames for PID selection and logging display
    content_frame = tk.Frame(main_frame, bg='#2b2b2b')
    content_frame.pack(fill='both', expand=True)

    # Left frame for PID selection
    left_frame = tk.Frame(content_frame, bg='#2b2b2b', width=400)
    left_frame.pack(side='left', fill='y', padx=(0, 10))
    left_frame.pack_propagate(False)

    # PID selection
    pid_select_label = tk.Label(left_frame, text="Select PIDs to Log:",
                                font=('Segoe UI', 12, 'bold'), fg='#4CAF50', bg='#2b2b2b')
    pid_select_label.pack(anchor='w', pady=(0, 10))

    # Scrollable PID list
    pid_canvas = tk.Canvas(left_frame, bg='#333333',
                           highlightthickness=0, height=400)
    pid_scrollbar = tk.Scrollbar(
        left_frame, orient='vertical', command=pid_canvas.yview)
    pid_scrollable_frame = tk.Frame(pid_canvas, bg='#333333')

    pid_scrollable_frame.bind(
        "<Configure>",
        lambda e: pid_canvas.configure(scrollregion=pid_canvas.bbox("all"))
    )

    pid_canvas.create_window((0, 0), window=pid_scrollable_frame, anchor="nw")
    pid_canvas.configure(yscrollcommand=pid_scrollbar.set)

    pid_canvas.pack(side="left", fill="both", expand=True)
    pid_scrollbar.pack(side="right", fill="y")

    # Right frame for data display
    right_frame = tk.Frame(content_frame, bg='#2b2b2b')
    right_frame.pack(side='right', fill='both', expand=True)

    # Data display area
    data_label = tk.Label(right_frame, text="Logged Data:",
                          font=('Segoe UI', 12, 'bold'), fg='#FF9800', bg='#2b2b2b')
    data_label.pack(anchor='w', pady=(0, 10))

    # Create text widget for data display
    data_text = tk.Text(right_frame, bg='#1e1e1e', fg='#ffffff',
                        font=('Consolas', 9), wrap='none')
    data_scrollbar_v = tk.Scrollbar(
        right_frame, orient='vertical', command=data_text.yview)
    data_scrollbar_h = tk.Scrollbar(
        right_frame, orient='horizontal', command=data_text.xview)

    data_text.configure(yscrollcommand=data_scrollbar_v.set,
                        xscrollcommand=data_scrollbar_h.set)

    data_text.pack(side='left', fill='both', expand=True)
    data_scrollbar_v.pack(side='right', fill='y')
    data_scrollbar_h.pack(side='bottom', fill='x')

    # Logger state variables
    logging_active = False
    logged_data = []
    selected_log_pids = []
    log_pid_vars = {}

    # Create PID checkboxes
    for pid_name, pid_info in available_pids.items():
        frame = tk.Frame(pid_scrollable_frame, bg='#333333')
        frame.pack(fill='x', pady=1, padx=5)

        var = tk.BooleanVar()
        log_pid_vars[pid_name] = var

        checkbox = tk.Checkbutton(frame, variable=var, bg='#333333', fg='white',
                                  selectcolor='#404040', font=('Segoe UI', 9))
        checkbox.pack(side='left')

        # Show priority PIDs first
        priority_indicator = " 🏆" if pid_info.get('priority', False) else ""
        info_text = f"{pid_name} ({pid_info['unit']}){priority_indicator}"
        label = tk.Label(frame, text=info_text, font=('Segoe UI', 8),
                         fg='white', bg='#333333', anchor='w')
        label.pack(side='left', fill='x', expand=True)

    # Logging control buttons
    def start_logging():
        nonlocal logging_active, selected_log_pids, logged_data

        selected_log_pids = [pid for pid,
                             var in log_pid_vars.items() if var.get()]
        if not selected_log_pids:
            messagebox.showwarning(
                "No PIDs Selected", "Please select at least one PID to log.")
            return

        logging_active = True
        logged_data = []
        start_log_button.config(text="🔴 Stop Logging", bg='#FF5722')
        data_text.delete(1.0, tk.END)

        # Add CSV header
        header = "Timestamp," + ",".join(selected_log_pids) + "\n"
        data_text.insert(tk.END, header)

        update_logger()

    def stop_logging():
        nonlocal logging_active
        logging_active = False
        start_log_button.config(text="🟢 Start Logging", bg='#4CAF50')

    def export_csv():
        if not logged_data:
            messagebox.showwarning(
                "No Data", "No data to export. Start logging first.")
            return

        from tkinter import filedialog
        import csv
        import datetime

        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Save PID Log Data"
        )

        if filename:
            try:
                with open(filename, 'w', newline='') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(["Timestamp"] + selected_log_pids)
                    writer.writerows(logged_data)
                messagebox.showinfo("Export Complete",
                                    f"Data exported to:\n{filename}")
            except Exception as e:
                messagebox.showerror(
                    "Export Error", f"Failed to export data:\n{str(e)}")

    def clear_data():
        nonlocal logged_data
        logged_data = []
        data_text.delete(1.0, tk.END)

    start_log_button = tk.Button(log_frame, text="🟢 Start Logging",
                                 font=('Segoe UI', 10, 'bold'), bg='#4CAF50', fg='white',
                                 relief='flat', bd=0, padx=15, pady=8,
                                 command=start_logging)
    start_log_button.pack(pady=5)

    export_button = tk.Button(log_frame, text="💾 Export CSV",
                              font=('Segoe UI', 9, 'bold'), bg='#607D8B', fg='white',
                              relief='flat', bd=0, padx=15, pady=6,
                              command=export_csv)
    export_button.pack(pady=2)

    clear_button = tk.Button(log_frame, text="🗑️ Clear Data",
                             font=('Segoe UI', 9, 'bold'), bg='#FF5722', fg='white',
                             relief='flat', bd=0, padx=15, pady=6,
                             command=clear_data)
    clear_button.pack(pady=2)

    # Logger update function
    def update_logger():
        nonlocal logging_active, logged_data, selected_log_pids

        if not logging_active or not connection or not connection.is_connected():
            return

        try:
            import datetime
            timestamp = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
            values = [timestamp]
            display_values = [timestamp]

            for pid_name in selected_log_pids:
                if pid_name in available_pids:
                    try:
                        cmd = available_pids[pid_name]['command']
                        response = connection.query(cmd, force=False)

                        if response and not response.is_null() and response.value is not None:
                            if hasattr(response.value, 'magnitude'):
                                value = round(response.value.magnitude, 3)
                            elif hasattr(response.value, '__float__'):
                                value = round(float(response.value), 3)
                            else:
                                value = str(response.value)
                            # Convert to string for CSV
                            values.append(str(value))
                            display_values.append(str(value))
                        else:
                            values.append("N/A")
                            display_values.append("N/A")
                    except Exception as e:
                        values.append("ERROR")
                        display_values.append("ERROR")

            # Store data
            logged_data.append(values)

            # Display data (show last 50 lines)
            display_line = ",".join(display_values) + "\n"
            data_text.insert(tk.END, display_line)

            # Keep only last 100 lines visible
            lines = data_text.get(1.0, tk.END).split('\n')
            if len(lines) > 100:
                data_text.delete(1.0, tk.END)
                data_text.insert(1.0, '\n'.join(lines[-100:]))

            data_text.see(tk.END)

        except Exception as e:
            print(f"Logger error: {e}")

        # Schedule next update
        if logging_active and pid_logger_window:
            try:
                pid_logger_window.after(250, update_logger)  # 4Hz logging rate
            except tk.TclError:
                # Window was destroyed
                logging_active = False

    # Handle window close
    def on_close():
        global pid_logger_window
        nonlocal logging_active
        logging_active = False
        if pid_logger_window:
            pid_logger_window.destroy()
            pid_logger_window = None

    pid_logger_window.protocol("WM_DELETE_WINDOW", on_close)


# ===== PID MONITOR TAB =====
pid_main_frame = tk.Frame(pid_monitor_tab, bg='#2b2b2b', padx=20, pady=15)
pid_main_frame.pack(fill='both', expand=True)

# PID Monitor title
pid_title_frame = tk.Frame(pid_main_frame, bg='#2b2b2b')
pid_title_frame.pack(fill='x', pady=(0, 20))

pid_title_label = tk.Label(pid_title_frame, text="PID Monitor & Scanner",
                           font=('Segoe UI', 16, 'bold'), fg='#FF9800', bg='#2b2b2b')
pid_title_label.pack(side='left')

pid_subtitle_label = tk.Label(pid_title_frame, text="Scan and Monitor All Available OBD PIDs",
                              font=('Segoe UI', 10), fg='#888888', bg='#2b2b2b')
pid_subtitle_label.pack(side='left', padx=(20, 0))

# PID Monitor section
pid_frame = tk.LabelFrame(pid_main_frame, text="PID Scanner & Monitor",
                          font=('Segoe UI', 11, 'bold'), fg='#FF9800', bg='#2b2b2b',
                          relief='flat', bd=2)
pid_frame.pack(fill='x', pady=(0, 20))

# PID control buttons
pid_button_frame = tk.Frame(pid_frame, bg='#2b2b2b')
pid_button_frame.pack(fill='x', pady=5, padx=10)

scan_pids_button = tk.Button(pid_button_frame, text="📡 Smart Scan PIDs",
                             font=('Segoe UI', 9, 'bold'), bg='#FF9800', fg='white',
                             relief='flat', bd=0, padx=15, pady=4,
                             command=lambda: threading.Thread(target=scan_available_pids).start())
scan_pids_button.pack(side='left', padx=(0, 5))

show_monitor_button = tk.Button(pid_button_frame, text="📊 Monitor PIDs",
                                font=('Segoe UI', 9, 'bold'), bg='#9C27B0', fg='white',
                                relief='flat', bd=0, padx=15, pady=4,
                                command=show_pid_monitor)
show_monitor_button.pack(side='left', padx=(5, 0))

# Add data logger button
logger_button = tk.Button(pid_button_frame, text="📊 Data Logger",
                          font=('Segoe UI', 9, 'bold'), bg='#2196F3', fg='white',
                          relief='flat', bd=0, padx=15, pady=4,
                          command=create_pid_logger)
logger_button.pack(side='right')

# PID status
pid_status_label = tk.Label(pid_frame, text="No PIDs scanned",
                            font=('Segoe UI', 9), fg='#888888', bg='#2b2b2b')
pid_status_label.pack(pady=5)

# PID Features section in PID tab
pid_features_frame = tk.LabelFrame(pid_main_frame, text="Available Features",
                                   font=('Segoe UI', 11, 'bold'), fg='#4CAF50', bg='#2b2b2b',
                                   relief='flat', bd=2)
pid_features_frame.pack(fill='x', pady=(0, 20))

# Feature descriptions
features_text = [
    "• 📡 Smart Scan: Intelligent PID scanning for optimal results",
    "• 📊 Monitor PIDs: Real-time monitoring of selected PIDs with units",
    "• � Data Logger: Comprehensive CSV logging with export functionality",
    "• �🔍 PID Details: View PID hex codes, units, and descriptions",
    "• ⚡ High-Speed Updates: 100ms refresh rate for real-time data",
    "• 📈 Value Tracking: Monitor multiple PIDs simultaneously",
    "• 🏆 Priority Detection: Automatically identifies common PIDs first"
]

for feature in features_text:
    feature_label = tk.Label(pid_features_frame, text=feature,
                             font=('Segoe UI', 9), fg='#BBBBBB', bg='#2b2b2b', anchor='w')
    feature_label.pack(fill='x', padx=10, pady=1)

# ===== DTC TAB =====
dtc_main_frame = tk.Frame(dtc_tab, bg='#2b2b2b', padx=20, pady=15)
dtc_main_frame.pack(fill='both', expand=True)

# DTC title
dtc_title_frame = tk.Frame(dtc_main_frame, bg='#2b2b2b')
dtc_title_frame.pack(fill='x', pady=(0, 20))

dtc_title_label = tk.Label(dtc_title_frame, text="🚨 Diagnostic Trouble Codes",
                           font=('Segoe UI', 16, 'bold'), fg='#FF5722', bg='#2b2b2b')
dtc_title_label.pack(side='left')

dtc_subtitle_label = tk.Label(dtc_title_frame, text="Read and Clear Engine Fault Codes",
                              font=('Segoe UI', 10), fg='#888888', bg='#2b2b2b')
dtc_subtitle_label.pack(side='left', padx=(20, 0))

# DTC Control section
dtc_control_frame = tk.LabelFrame(dtc_main_frame, text="DTC Scanner Control",
                                  font=('Segoe UI', 11, 'bold'), fg='#FF5722', bg='#2b2b2b',
                                  relief='flat', bd=2)
dtc_control_frame.pack(fill='x', pady=(0, 20))

# Control buttons
dtc_button_frame = tk.Frame(dtc_control_frame, bg='#2b2b2b')
dtc_button_frame.pack(fill='x', pady=10, padx=10)

read_dtc_button = tk.Button(dtc_button_frame, text="📖 Read DTCs",
                            font=('Segoe UI', 10, 'bold'), bg='#FF9800', fg='white',
                            relief='flat', bd=0, padx=20, pady=8,
                            command=lambda: threading.Thread(target=read_dtcs).start())
read_dtc_button.pack(side='left', padx=(0, 10))

clear_dtc_button = tk.Button(dtc_button_frame, text="🗑️ Clear DTCs",
                             font=('Segoe UI', 10, 'bold'), bg='#F44336', fg='white',
                             relief='flat', bd=0, padx=20, pady=8,
                             command=lambda: threading.Thread(target=clear_dtcs).start())
clear_dtc_button.pack(side='left', padx=(0, 10))

dtc_status_button = tk.Button(dtc_button_frame, text="🔍 Check MIL Status",
                              font=('Segoe UI', 10, 'bold'), bg='#9C27B0', fg='white',
                              relief='flat', bd=0, padx=20, pady=8,
                              command=lambda: threading.Thread(target=check_mil_status).start())
dtc_status_button.pack(side='right')

# DTC Status display
dtc_status_frame = tk.Frame(dtc_control_frame, bg='#2b2b2b')
dtc_status_frame.pack(fill='x', pady=5, padx=10)

dtc_status_label = tk.Label(dtc_status_frame, text="● Ready to scan DTCs",
                            font=('Segoe UI', 10), fg='#4CAF50', bg='#2b2b2b')
dtc_status_label.pack()

# DTC Results section
dtc_results_frame = tk.LabelFrame(dtc_main_frame, text="DTC Results",
                                  font=('Segoe UI', 11, 'bold'), fg='#4CAF50', bg='#2b2b2b',
                                  relief='flat', bd=2)
dtc_results_frame.pack(fill='both', expand=True, pady=(0, 20))

# Create scrollable text area for DTC results
dtc_canvas = tk.Canvas(dtc_results_frame, bg='#1e1e1e', highlightthickness=0)
dtc_scrollbar = tk.Scrollbar(
    dtc_results_frame, orient='vertical', command=dtc_canvas.yview)
dtc_content_frame = tk.Frame(dtc_canvas, bg='#1e1e1e')

dtc_content_frame.bind(
    "<Configure>",
    lambda e: dtc_canvas.configure(scrollregion=dtc_canvas.bbox("all"))
)

dtc_canvas.create_window((0, 0), window=dtc_content_frame, anchor="nw")
dtc_canvas.configure(yscrollcommand=dtc_scrollbar.set)

dtc_canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
dtc_scrollbar.pack(side="right", fill="y", pady=10)

# Default message
default_dtc_label = tk.Label(dtc_content_frame,
                             text="📋 Click 'Read DTCs' to scan for diagnostic trouble codes\n\n"
                             "🔍 The scanner will check for:\n"
                             "• Active DTCs (current faults)\n"
                             "• Pending DTCs (intermittent faults)\n"
                             "• Permanent DTCs (emissions-related)\n"
                             "• MIL (Check Engine Light) status\n\n"
                             "⚠️ Note: Vehicle must be connected first",
                             font=('Segoe UI', 10), fg='#BBBBBB', bg='#1e1e1e',
                             anchor='w', justify='left')
default_dtc_label.pack(fill='both', expand=True, padx=20, pady=20)

# DTC Information section
dtc_info_frame = tk.LabelFrame(dtc_main_frame, text="DTC Information",
                               font=('Segoe UI', 11, 'bold'), fg='#2196F3', bg='#2b2b2b',
                               relief='flat', bd=2)
dtc_info_frame.pack(fill='x')

dtc_info_text = [
    "🔧 DTC Code Format:",
    "• P0XXX: Powertrain (Engine/Transmission)",
    "• B0XXX: Body (Air bags, A/C, etc.)",
    "• C0XXX: Chassis (ABS, Steering, etc.)",
    "• U0XXX: Network (Communication)",
    "",
    "📊 DTC Types:",
    "• Active: Current faults affecting performance",
    "• Pending: Intermittent faults (1 failure cycle)",
    "• Permanent: Emissions faults requiring drive cycle",
    "",
    "⚠️ Safety Notes:",
    "• Clear codes only after repairs are completed",
    "• Some codes may return if fault persists",
    "• Clearing codes resets readiness monitors"
]

for info in dtc_info_text:
    style = 'bold' if info.startswith('🔧') or info.startswith(
        '📊') or info.startswith('⚠️') else 'normal'
    color = '#FF9800' if info.startswith('🔧') else '#4CAF50' if info.startswith(
        '📊') else '#F44336' if info.startswith('⚠️') else '#BBBBBB'
    tk.Label(dtc_info_frame, text=info, font=('Segoe UI', 8, style),
             fg=color, bg='#2b2b2b', anchor='w').pack(fill='x', padx=10, pady=1)

# ===== SETTINGS TAB =====
settings_main_frame = tk.Frame(settings_tab, bg='#2b2b2b', padx=20, pady=15)
settings_main_frame.pack(fill='both', expand=True)

# Settings title
settings_title_frame = tk.Frame(settings_main_frame, bg='#2b2b2b')
settings_title_frame.pack(fill='x', pady=(0, 20))

settings_title_label = tk.Label(settings_title_frame, text="Application Settings",
                                font=('Segoe UI', 16, 'bold'), fg='#9C27B0', bg='#2b2b2b')
settings_title_label.pack(side='left')

settings_subtitle_label = tk.Label(settings_title_frame, text="Configure Application Preferences",
                                   font=('Segoe UI', 10), fg='#888888', bg='#2b2b2b')
settings_subtitle_label.pack(side='left', padx=(20, 0))

# Update rate settings
update_frame = tk.LabelFrame(settings_main_frame, text="Update Settings",
                             font=('Segoe UI', 11, 'bold'), fg='#2196F3', bg='#2b2b2b',
                             relief='flat', bd=2)
update_frame.pack(fill='x', pady=(0, 20))

tk.Label(update_frame, text="VE Table Update Rate:", font=('Segoe UI', 10),
         fg='white', bg='#2b2b2b').pack(anchor='w', padx=10, pady=(10, 5))

update_rate_var = tk.StringVar(value="50ms")
update_rates = ["25ms (Very Fast)", "50ms (Fast)",
                "100ms (Normal)", "250ms (Slow)", "500ms (Very Slow)"]
update_combo = ttk.Combobox(update_frame, textvariable=update_rate_var,
                            values=update_rates, state="readonly", width=20)
update_combo.pack(padx=10, pady=(0, 10))

# Application info
info_frame = tk.LabelFrame(settings_main_frame, text="Application Information",
                           font=('Segoe UI', 11, 'bold'), fg='#4CAF50', bg='#2b2b2b',
                           relief='flat', bd=2)
info_frame.pack(fill='x', pady=(0, 10))

info_text = [
    "VE Table Monitor v2.0",
    "Professional Engine Tuning Tool",
    "",
    "Features:",
    "• Real-time VE table visualization",
    "• OBD-II communication (Serial/Bluetooth)",
    "• Comprehensive PID scanning & monitoring",
    "• Modern tabbed interface",
    "• Demo mode for testing"
]

for info in info_text:
    style = 'bold' if info.startswith(
        'VE Table') or info == "Features:" else 'normal'
    color = '#4CAF50' if info.startswith(
        'VE Table') else '#FF9800' if info == "Features:" else '#BBBBBB'
    tk.Label(info_frame, text=info, font=('Segoe UI', 9, style),
             fg=color, bg='#2b2b2b', anchor='w').pack(fill='x', padx=10, pady=1)

# Global variables for PID monitoring (consolidated)
available_pids = {}
pid_monitor_window = None
pid_monitoring_active = False
monitored_pids = []
pid_value_labels = {}


def quick_connect_port(port):
    """Quick connect to a specific COM port with optimized settings for OBDXPROVX"""
    global connection
    try:
        status_label.config(text=f"● Connecting to {port}...", fg="#FFA726")

        # Close existing connection
        if connection:
            try:
                connection.close()
            except:
                pass
            connection = None

        print(f"🔵 OBDXPROVX Quick Connection to {port}")

        # Special handling for OBDXPROVX Bluetooth OBD ports (COM5/COM6 primarily)
        if port.upper() in ["COM5", "COM6"]:
            print(
                f"🔵 OBDXPROVX Detected - Using optimized settings for {port}")
            status_label.config(
                text=f"● {port} OBDXPROVX connection...", fg="#FFA726")

            # OBDXPROVX specific optimized settings (based on actual device testing)
            obdx_settings = [
                {'baudrate': 38400, 'timeout': 15, 'fast': False,
                    'check_voltage': False},  # Primary OBDXPROVX setting
                {'baudrate': 9600, 'timeout': 18, 'fast': False,
                    'check_voltage': False},   # Backup OBDXPROVX setting
                {'baudrate': 115200, 'timeout': 10, 'fast': False,
                    'check_voltage': False},  # High-speed test
            ]

            connection_successful = False
            for i, settings in enumerate(obdx_settings):
                try:
                    print(
                        f"   Attempt {i+1}/3: {settings['baudrate']} baud, {settings['timeout']}s timeout")
                    status_label.config(
                        text=f"● {port} OBDXPROVX ({settings['baudrate']} baud)...", fg="#FFA726")

                    # Create connection with OBDXPROVX optimizations
                    connection = obd.OBD(port,
                                         baudrate=settings['baudrate'],
                                         timeout=settings['timeout'],
                                         fast=settings['fast'],
                                         check_voltage=settings.get('check_voltage', False))

                    if connection and connection.is_connected():
                        print(
                            f"   ✅ OBDXPROVX connected successfully at {settings['baudrate']} baud!")
                        connection_successful = True
                        break
                    else:
                        print(f"   ❌ Failed at {settings['baudrate']} baud")
                        if connection:
                            connection.close()
                            connection = None

                except Exception as attempt_error:
                    print(
                        f"   ❌ Error at {settings['baudrate']} baud: {attempt_error}")
                    if connection:
                        try:
                            connection.close()
                        except:
                            pass
                        connection = None

            if not connection_successful:
                print(f"❌ All OBDXPROVX connection attempts failed for {port}")
                status_label.config(
                    text=f"● OBDXPROVX {port} failed", fg="#FF4444")
                return

        elif port.upper() == "COM7":
            print(f"🔵 Using standard Bluetooth OBD settings for {port}")
            status_label.config(
                text=f"● {port} Bluetooth OBD connection...", fg="#FFA726")

            # Standard Bluetooth OBD settings for COM7
            standard_settings = [
                {'baudrate': 38400, 'timeout': 12, 'fast': False},
                {'baudrate': 9600, 'timeout': 15, 'fast': False},
                {'baudrate': 115200, 'timeout': 8, 'fast': False},
            ]

            connection_successful = False
            for i, settings in enumerate(standard_settings):
                try:
                    print(
                        f"COM7 quick connect attempt {i+1}: baudrate={settings['baudrate']}")
                    status_label.config(
                        text=f"● COM7 attempt {i+1}: {settings['baudrate']} baud...", fg="#FFA726")

                    connection = obd.OBD(
                        port,
                        baudrate=settings['baudrate'],
                        timeout=settings['timeout'],
                        fast=settings['fast']
                    )

                    if connection and connection.is_connected():
                        print(
                            f"✓ COM7 quick connected with {settings['baudrate']} baud")
                        connection_successful = True
                        break
                    else:
                        if connection:
                            connection.close()
                        connection = None

                except Exception as com7_error:
                    print(f"COM7 quick attempt {i+1} failed: {com7_error}")
                    if connection:
                        try:
                            connection.close()
                        except:
                            pass
                        connection = None
                    continue

            if not connection_successful:
                raise Exception(f"All COM7 quick connection attempts failed")
        else:
            # Standard connection for other ports
            connection = obd.OBD(port, baudrate=38400, timeout=15, fast=False)

        if connection and connection.is_connected():
            # Determine connection type and show success
            if port.upper() in ["COM5", "COM6"]:
                conn_type = "OBDXPROVX Bluetooth"
                device_info = "OBDX Pro VX"
            elif port.upper() == "COM7":
                conn_type = "Bluetooth OBD"
                device_info = "Standard Bluetooth"
            else:
                conn_type = "Serial"
                device_info = "Standard Serial"

            status_label.config(
                text=f"● Connected: {device_info} ({port})", fg="#4CAF50")

            # Show success message with device-specific info
            success_msg = f"✅ OBDXPROVX Connection Successful!\n\n" if port.upper(
            ) in ["COM5", "COM6"] else f"✅ Quick connection successful!\n\n"
            success_msg += f"🔵 Device: {device_info}\n"
            success_msg += f"📡 Port: {port}\n"
            success_msg += f"🔗 Type: {conn_type}\n"
            success_msg += f"⚙️ Protocol: {getattr(connection, 'protocol', 'Auto-detected')}\n\n"

            if port.upper() in ["COM5", "COM6"]:
                success_msg += f"🎯 OBDXPROVX Ready! You can now:\n"
                success_msg += f"• Use 'Scan PIDs' to find available parameters\n"
                success_msg += f"• Monitor real-time engine data\n"
                success_msg += f"• Read/Clear diagnostic trouble codes\n"
                success_msg += f"• Export data to CSV for analysis\n\n"
                success_msg += f"💡 Tip: Your OBDXPROVX works best with vehicle running!"
            else:
                success_msg += f"🎯 You can now:\n"
                success_msg += f"• Use 'Scan PIDs' to find available parameters\n"
                success_msg += f"• Open 'Comprehensive Logger' for detailed monitoring\n"
                success_msg += f"• View real-time VE table data"

            messagebox.showinfo("OBDXPROVX Connected!" if port.upper() in [
                                "COM5", "COM6"] else "Quick Connection Successful", success_msg)

            # Start updating
            print(f"🚀 Starting data updates for {device_info}")
            update()
        else:
            status_label.config(
                text=f"● Failed to connect to {port}", fg="#FF4444")

            # Device-specific error messages
            if port.upper() in ["COM5", "COM6"]:
                error_msg = f"❌ OBDXPROVX connection to {port} failed\n\n"
                error_msg += f"🔧 OBDXPROVX TROUBLESHOOTING:\n"
                error_msg += f"• Ensure vehicle ignition is ON (powers OBDXPROVX)\n"
                error_msg += f"• Check OBDXPROVX LED is solid/blinking blue\n"
                error_msg += f"• Verify OBDXPROVX is plugged into OBD port\n"
                error_msg += f"• Try unplugging and reconnecting OBDXPROVX\n"
                error_msg += f"• Close Torque, OBD Auto Doctor, or similar apps\n"
                error_msg += f"• Run this application as Administrator\n\n"
                error_msg += f"💡 Your OBDXPROVX is paired correctly but may need\n"
                error_msg += f"   the vehicle running to establish communication."
            elif port.upper() == "COM7":
                error_msg = f"❌ Quick connection to {port} failed\n\n"
                error_msg += f"🔧 TROUBLESHOOTING:\n"
                error_msg += f"• Ensure vehicle is running (powers OBD adapter)\n"
                error_msg += f"• Check that {port} is the correct port\n"
                error_msg += f"• Close other OBD software first\n"
                error_msg += f"📱 COM7 BLUETOOTH SPECIFIC:\n"
                error_msg += f"• Check Bluetooth pairing in Windows settings\n"
                error_msg += f"• Verify OBD adapter is powered and paired\n"
                error_msg += f"• Try removing and re-pairing the device\n"
                error_msg += f"• Check Device Manager for 'Standard Serial over Bluetooth'\n"
            else:
                error_msg = f"❌ Quick connection to {port} failed\n\n"
                error_msg += f"🔧 TROUBLESHOOTING:\n"
                error_msg += f"• Ensure vehicle is running (powers OBD adapter)\n"
                error_msg += f"• Check that {port} is the correct port\n"
                error_msg += f"• Close other OBD software first\n"
                error_msg += f"• Use 'Scan Devices' to find available ports"

            messagebox.showerror("OBDXPROVX Connection Failed" if port.upper() in [
                                 "COM5", "COM6"] else "Quick Connection Failed", error_msg)

    except Exception as e:
        status_label.config(text=f"● Error connecting to {port}", fg="#FF4444")
        error_msg = f"❌ Error connecting to {port}\n\n"
        error_msg += f"Error: {str(e)}\n\n"
        error_msg += f"💡 SOLUTIONS:\n"

        if port.upper() in ["COM5", "COM6"]:
            error_msg += f"🔵 OBDXPROVX TROUBLESHOOTING:\n"
            error_msg += f"• Check if {port} exists in Device Manager\n"
            error_msg += f"• Ensure OBDXPROVX is paired in Windows Bluetooth\n"
            error_msg += f"• Verify vehicle ignition is ON\n"
            error_msg += f"• Try restarting Windows Bluetooth service\n"
            error_msg += f"• Run application as Administrator\n"
            error_msg += f"• Close other OBD software (Torque, etc.)\n"
            error_msg += f"• Check OBDXPROVX LED status on device"
        elif port.upper() == "COM7":
            error_msg += f"📱 COM7 BLUETOOTH TROUBLESHOOTING:\n"
            error_msg += f"• Check if COM7 exists in Device Manager\n"
            error_msg += f"• Ensure Bluetooth adapter is paired correctly\n"
            error_msg += f"• Try restarting Windows Bluetooth service\n"
            error_msg += f"• Run application as Administrator\n"
            error_msg += f"• Check if other Bluetooth OBD software is running\n"
        else:
            error_msg += f"• Check if {port} exists in Device Manager\n"
            error_msg += f"• Ensure no other software is using the port\n"
            error_msg += f"• Try running as Administrator\n"
            error_msg += f"• Use 'Scan Devices' to find available ports"

        messagebox.showerror("OBDXPROVX Connection Error" if port.upper() in [
                             "COM5", "COM6"] else "Connection Error", error_msg)


# Function to test specific connection ports


def test_connection_ports():
    """Test connection to specific COM ports (COM3 and COM7)"""
    try:
        status_label.config(text="● Testing COM ports...", fg="#FFA726")

        test_ports = ["COM3", "COM7"]
        results = []

        for port in test_ports:
            try:
                status_label.config(text=f"● Testing {port}...", fg="#FFA726")
                print(f"Testing connection to {port}...")

                # Try to connect with a timeout
                test_conn = obd.OBD(port, timeout=10, fast=False)

                if test_conn and test_conn.is_connected():
                    protocol = getattr(test_conn, 'protocol', 'Unknown')
                    results.append(
                        f"✅ {port}: Connected (Protocol: {protocol})")
                    print(f"✅ {port} connected successfully")

                    # Test a basic command
                    try:
                        # Try to get available commands and test one
                        all_commands = [getattr(obd.commands, cmd) for cmd in dir(obd.commands)
                                        if not cmd.startswith('_') and hasattr(obd.commands, cmd)]
                        test_commands = [
                            cmd for cmd in all_commands if cmd is not None and hasattr(cmd, 'pid')]

                        if test_commands:
                            # Use first available command
                            test_cmd = test_commands[0]
                            test_response = test_conn.query(test_cmd)
                            if test_response and not test_response.is_null():
                                results.append(
                                    f"   • Test command ({test_cmd.name}): Success")
                            else:
                                results.append(
                                    f"   • Test command: No response")
                        else:
                            results.append(f"   • No test commands available")
                    except Exception as cmd_error:
                        results.append(
                            f"   • Command test failed: {cmd_error}")

                    test_conn.close()
                else:
                    results.append(f"❌ {port}: No connection")
                    print(f"❌ {port} failed to connect")

            except Exception as port_error:
                results.append(f"❌ {port}: Error - {str(port_error)}")
                print(f"❌ {port} error: {port_error}")

        # Show results
        result_text = "COM PORT CONNECTION TEST\n\n" + "\n".join(results)
        result_text += "\n\n💡 Recommendations:\n"
        result_text += "• Use the port that shows 'Connected'\n"
        result_text += "• Make sure vehicle is running\n"
        result_text += "• Close other OBD software first"

        messagebox.showinfo("Connection Test Results", result_text)

    except Exception as e:
        messagebox.showerror("Test Error", f"Error testing ports:\n{str(e)}")
    finally:
        status_label.config(text="● Not Connected", fg="#FF4444")


# Specialized PID definitions for logging
PRIORITY_PIDS = {
    # Key PIDs for engine tuning
    'IAC_POSITION': {'names': ['IAC_POSITION', 'IDLE_AIR_CONTROL', 'IAC'], 'unit': '%', 'description': 'Idle Air Control Position'},
    'IDLE_RPM': {'names': ['IDLE_RPM', 'ENGINE_RPM', 'RPM'], 'unit': 'rpm', 'description': 'Engine RPM (Idle)'},
    'SPARK_ADVANCE': {'names': ['TIMING_ADVANCE', 'SPARK_ADVANCE', 'IGNITION_TIMING'], 'unit': '°', 'description': 'Spark Advance'},
    'TPS': {'names': ['THROTTLE_POS', 'TPS', 'THROTTLE_POSITION'], 'unit': '%', 'description': 'Throttle Position Sensor'},
    'MAP': {'names': ['INTAKE_PRESSURE', 'MAP', 'MANIFOLD_PRESSURE'], 'unit': 'kPa', 'description': 'Manifold Absolute Pressure'},
    'STIT': {'names': ['SHORT_FUEL_TRIM_1', 'STIT', 'ST_FUEL_TRIM'], 'unit': '%', 'description': 'Short Term Idle Trim'},
    'LTIT': {'names': ['LONG_FUEL_TRIM_1', 'LTIT', 'LT_FUEL_TRIM'], 'unit': '%', 'description': 'Long Term Idle Trim'},
    'MAF': {'names': ['MAF', 'MASS_AIR_FLOW', 'AIR_FLOW_RATE'], 'unit': 'g/s', 'description': 'Mass Air Flow'},
    'IAT': {'names': ['INTAKE_TEMP', 'IAT', 'AIR_TEMP'], 'unit': '°C', 'description': 'Intake Air Temperature'},
    'COOLANT_TEMP': {'names': ['COOLANT_TEMP', 'ECT', 'ENGINE_COOLANT_TEMP'], 'unit': '°C', 'description': 'Engine Coolant Temperature'},
    'FUEL_PRESSURE': {'names': ['FUEL_PRESSURE', 'FUEL_RAIL_PRESSURE'], 'unit': 'kPa', 'description': 'Fuel Rail Pressure'},
    'OXYGEN_SENSOR': {'names': ['O2_S1_WR_VOLTAGE', 'OXYGEN_SENSOR', 'O2_SENSOR'], 'unit': 'V', 'description': 'Oxygen Sensor'},
}


def find_priority_pids():
    """Find priority PIDs from available commands"""
    priority_found = {}

    try:
        # Get all available commands
        all_commands = [getattr(obd.commands, cmd) for cmd in dir(obd.commands)
                        if not cmd.startswith('_') and hasattr(obd.commands, cmd)]

        # Filter for actual OBD commands
        obd_commands = [
            cmd for cmd in all_commands if cmd is not None and hasattr(cmd, 'pid')]

        # Match priority PIDs with available commands
        for priority_key, priority_info in PRIORITY_PIDS.items():
            for cmd in obd_commands:
                if hasattr(cmd, 'name'):
                    cmd_name_upper = cmd.name.upper()
                    for search_name in priority_info['names']:
                        if search_name in cmd_name_upper:
                            priority_found[priority_key] = {
                                'command': cmd,
                                'name': cmd.name,
                                'expected_unit': priority_info['unit'],
                                'description': priority_info['description'],
                                'pid_hex': f"0x{cmd.pid:02X}" if hasattr(cmd, 'pid') else 'Unknown'
                            }
                            break
                    if priority_key in priority_found:
                        break

        return priority_found

    except Exception as e:
        print(f"Error finding priority PIDs: {e}")
        return {}


# ===== DTC FUNCTIONS =====

def read_dtcs():
    """Read all diagnostic trouble codes from the vehicle"""
    global dtc_content_frame, dtc_status_label

    if not connection or not connection.is_connected():
        messagebox.showerror("Connection Error",
                             "Please connect to a vehicle first before reading DTCs!")
        return

    try:
        dtc_status_label.config(text="● Reading DTCs...", fg="#FFA726")

        # Clear previous results
        for widget in dtc_content_frame.winfo_children():
            widget.destroy()

        # Create results display
        results_label = tk.Label(dtc_content_frame, text="🔍 DTC SCAN RESULTS\n",
                                 font=('Segoe UI', 12, 'bold'), fg='#4CAF50', bg='#1e1e1e')
        results_label.pack(anchor='w', padx=20, pady=(10, 5))

        total_dtcs = 0
        scan_results = []

        # Read different types of DTCs using safe command access
        dtc_types = []

        # Check for available DTC commands using getattr
        get_dtc_cmd = getattr(obd.commands, 'GET_DTC', None)
        if get_dtc_cmd:
            dtc_types.append(("Active DTCs", get_dtc_cmd, "🔴"))

        get_current_dtc_cmd = getattr(obd.commands, 'GET_CURRENT_DTC', None)
        if get_current_dtc_cmd:
            dtc_types.append(("Current DTCs", get_current_dtc_cmd, "🟡"))

        freeze_dtc_cmd = getattr(obd.commands, 'FREEZE_DTC', None)
        if freeze_dtc_cmd:
            dtc_types.append(("Freeze Frame DTCs", freeze_dtc_cmd, "🟠"))

        # If no DTC commands found, try alternative approach
        if not dtc_types:
            # Add message about unsupported commands
            error_label = tk.Label(dtc_content_frame,
                                   text="⚠️ No DTC commands available\nYour OBD library version may not support DTC scanning",
                                   font=('Segoe UI', 11),
                                   fg='#FF9800', bg='#1e1e1e')
            error_label.pack(pady=20)
            dtc_status_label.config(
                text="● DTC commands not available", fg="#FF9800")
            return

        for dtc_type_name, dtc_command, icon in dtc_types:
            try:
                dtc_status_label.config(
                    text=f"● Reading {dtc_type_name}...", fg="#FFA726")

                # Query for DTCs
                response = connection.query(dtc_command)

                if response and not response.is_null() and response.value:
                    dtcs = response.value

                    if dtcs:
                        # Display category header
                        category_frame = tk.Frame(
                            dtc_content_frame, bg='#333333', relief='raised', bd=1)
                        category_frame.pack(fill='x', padx=20, pady=(10, 5))

                        category_label = tk.Label(category_frame,
                                                  text=f"{icon} {dtc_type_name} ({len(dtcs)} found)",
                                                  font=('Segoe UI',
                                                        11, 'bold'),
                                                  fg='#FFD700', bg='#333333')
                        category_label.pack(pady=5)

                        # Display each DTC
                        for dtc in dtcs:
                            dtc_frame = tk.Frame(
                                dtc_content_frame, bg='#404040', relief='flat', bd=1)
                            dtc_frame.pack(fill='x', padx=30, pady=2)

                            # DTC code and description
                            dtc_code = str(dtc)
                            dtc_desc = get_dtc_description(dtc_code)

                            code_label = tk.Label(dtc_frame, text=f"Code: {dtc_code}",
                                                  font=('Segoe UI',
                                                        10, 'bold'),
                                                  fg='#FF5722', bg='#404040')
                            code_label.pack(anchor='w', padx=10, pady=2)

                            desc_label = tk.Label(dtc_frame, text=f"Description: {dtc_desc}",
                                                  font=('Segoe UI', 9),
                                                  fg='white', bg='#404040', wraplength=600)
                            desc_label.pack(anchor='w', padx=10, pady=(0, 5))

                            total_dtcs += 1
                    else:
                        # No DTCs in this category
                        no_dtc_frame = tk.Frame(
                            dtc_content_frame, bg='#2d5a2d', relief='flat', bd=1)
                        no_dtc_frame.pack(fill='x', padx=20, pady=2)

                        no_dtc_label = tk.Label(no_dtc_frame, text=f"✅ No {dtc_type_name.lower()} found",
                                                font=('Segoe UI', 10),
                                                fg='#4CAF50', bg='#2d5a2d')
                        no_dtc_label.pack(pady=5)
                else:
                    # Command not supported or failed
                    error_frame = tk.Frame(
                        dtc_content_frame, bg='#5a2d2d', relief='flat', bd=1)
                    error_frame.pack(fill='x', padx=20, pady=2)

                    error_label = tk.Label(error_frame, text=f"⚠️ {dtc_type_name} scan failed or not supported",
                                           font=('Segoe UI', 10),
                                           fg='#FF9800', bg='#5a2d2d')
                    error_label.pack(pady=5)

            except Exception as dtc_error:
                print(f"Error reading {dtc_type_name}: {dtc_error}")
                error_frame = tk.Frame(
                    dtc_content_frame, bg='#5a2d2d', relief='flat', bd=1)
                error_frame.pack(fill='x', padx=20, pady=2)

                error_label = tk.Label(error_frame, text=f"❌ Error reading {dtc_type_name}: {str(dtc_error)}",
                                       font=('Segoe UI', 10),
                                       fg='#F44336', bg='#5a2d2d')
                error_label.pack(pady=5)

        # Summary
        summary_frame = tk.Frame(
            dtc_content_frame, bg='#1e1e1e', relief='raised', bd=2)
        summary_frame.pack(fill='x', padx=20, pady=(20, 10))

        if total_dtcs > 0:
            summary_text = f"📊 SCAN SUMMARY: Found {total_dtcs} total DTCs"
            summary_color = '#FF9800'
            dtc_status_label.config(
                text=f"● Found {total_dtcs} DTCs", fg="#FF9800")
        else:
            summary_text = "✅ SCAN SUMMARY: No DTCs found - System OK"
            summary_color = '#4CAF50'
            dtc_status_label.config(text="● No DTCs found", fg="#4CAF50")

        summary_label = tk.Label(summary_frame, text=summary_text,
                                 font=('Segoe UI', 12, 'bold'),
                                 fg=summary_color, bg='#1e1e1e')
        summary_label.pack(pady=10)

        # Recommendations
        if total_dtcs > 0:
            rec_label = tk.Label(summary_frame,
                                 text="💡 Recommendations:\n"
                                 "• Document all codes before clearing\n"
                                 "• Address underlying issues before clearing\n"
                                 "• Consult service manual for specific codes",
                                 font=('Segoe UI', 9),
                                 fg='#BBBBBB', bg='#1e1e1e', justify='left')
            rec_label.pack(pady=(0, 10))

        print(f"DTC scan complete: {total_dtcs} codes found")

    except Exception as e:
        dtc_status_label.config(text="● DTC scan failed", fg="#F44336")
        messagebox.showerror("DTC Scan Error",
                             f"Error reading DTCs:\n{str(e)}\n\n"
                             f"Solutions:\n"
                             f"• Check vehicle connection\n"
                             f"• Ensure vehicle is running\n"
                             f"• Try reconnecting to vehicle")
        print(f"DTC scan error: {e}")


def clear_dtcs():
    """Clear all diagnostic trouble codes"""
    global dtc_status_label

    if not connection or not connection.is_connected():
        messagebox.showerror("Connection Error",
                             "Please connect to a vehicle first before clearing DTCs!")
        return

    # Confirmation dialog
    confirm = messagebox.askyesno("Clear DTCs",
                                  "⚠️ WARNING: Clear All DTCs?\n\n"
                                  "This will:\n"
                                  "• Clear all diagnostic trouble codes\n"
                                  "• Reset readiness monitors\n"
                                  "• Turn off Check Engine Light\n"
                                  "• Reset freeze frame data\n\n"
                                  "Only proceed if repairs have been completed!\n\n"
                                  "Continue?")

    if not confirm:
        return

    try:
        dtc_status_label.config(text="● Clearing DTCs...", fg="#FFA726")

        # Clear DTCs command
        clear_dtc_cmd = getattr(obd.commands, 'CLEAR_DTC', None)
        if not clear_dtc_cmd:
            dtc_status_label.config(
                text="● Clear DTC command not available", fg="#FF9800")
            messagebox.showerror("Command Not Available",
                                 "Clear DTC command is not available in your OBD library version.")
            return

        response = connection.query(clear_dtc_cmd)

        if response and not response.is_null():
            dtc_status_label.config(
                text="● DTCs cleared successfully", fg="#4CAF50")

            # Clear the display
            for widget in dtc_content_frame.winfo_children():
                widget.destroy()

            # Show success message
            success_frame = tk.Frame(
                dtc_content_frame, bg='#2d5a2d', relief='raised', bd=2)
            success_frame.pack(fill='both', expand=True, padx=20, pady=20)

            success_label = tk.Label(success_frame,
                                     text="✅ DTCs CLEARED SUCCESSFULLY\n\n"
                                     "🔧 What was cleared:\n"
                                     "• All active DTCs\n"
                                     "• All pending DTCs\n"
                                     "• Freeze frame data\n"
                                     "• Readiness monitors reset\n\n"
                                     "📋 Next steps:\n"
                                     "• Start engine and let it warm up\n"
                                     "• Drive vehicle through normal cycle\n"
                                     "• Monitor for returning codes\n"
                                     "• Readiness monitors will set over time",
                                     font=('Segoe UI', 11),
                                     fg='#4CAF50', bg='#2d5a2d',
                                     justify='left')
            success_label.pack(expand=True, padx=20, pady=20)

            messagebox.showinfo("Clear DTCs Successful",
                                "✅ DTCs cleared successfully!\n\n"
                                "The Check Engine Light should turn off.\n"
                                "Drive the vehicle normally to reset readiness monitors.")

            print("DTCs cleared successfully")
        else:
            dtc_status_label.config(text="● Clear DTCs failed", fg="#F44336")
            messagebox.showerror("Clear DTCs Failed",
                                 "Failed to clear DTCs.\n\n"
                                 "Possible causes:\n"
                                 "• Command not supported by vehicle\n"
                                 "• Communication error\n"
                                 "• Vehicle in wrong state\n\n"
                                 "Try:\n"
                                 "• Ensure engine is running\n"
                                 "• Check connection stability")

    except Exception as e:
        dtc_status_label.config(text="● Clear DTCs error", fg="#F44336")
        messagebox.showerror("Clear DTCs Error",
                             f"Error clearing DTCs:\n{str(e)}\n\n"
                             f"Solutions:\n"
                             f"• Check vehicle connection\n"
                             f"• Ensure vehicle is running\n"
                             f"• Try reconnecting to vehicle")
        print(f"Clear DTCs error: {e}")


def check_mil_status():
    """Check Malfunction Indicator Lamp (Check Engine Light) status"""
    global dtc_status_label

    if not connection or not connection.is_connected():
        messagebox.showerror("Connection Error",
                             "Please connect to a vehicle first!")
        return

    try:
        dtc_status_label.config(text="● Checking MIL status...", fg="#FFA726")

        # Clear previous results
        for widget in dtc_content_frame.winfo_children():
            widget.destroy()

        # Create status display
        status_label = tk.Label(dtc_content_frame, text="🔍 MIL & READINESS STATUS\n",
                                font=('Segoe UI', 12, 'bold'), fg='#4CAF50', bg='#1e1e1e')
        status_label.pack(anchor='w', padx=20, pady=(10, 5))

        # Check MIL status
        try:
            status_cmd = getattr(obd.commands, 'STATUS', None)
            if not status_cmd:
                error_label = tk.Label(dtc_content_frame,
                                       text="⚠️ STATUS command not available\nMIL status cannot be read with this OBD library version",
                                       font=('Segoe UI', 11),
                                       fg='#FF9800', bg='#1e1e1e')
                error_label.pack(pady=20)
                dtc_status_label.config(
                    text="● STATUS command not available", fg="#FF9800")
                return

            mil_response = connection.query(status_cmd)

            if mil_response and not mil_response.is_null() and mil_response.value:
                status_data = mil_response.value

                # MIL status
                mil_frame = tk.Frame(
                    dtc_content_frame, bg='#404040', relief='raised', bd=1)
                mil_frame.pack(fill='x', padx=20, pady=5)

                mil_on = getattr(status_data, 'MIL', False)
                dtc_count = getattr(status_data, 'DTC_count', 0)

                mil_icon = "🔴" if mil_on else "✅"
                mil_text = "ON (Check Engine Light)" if mil_on else "OFF"
                mil_color = "#F44336" if mil_on else "#4CAF50"

                mil_label = tk.Label(mil_frame,
                                     text=f"{mil_icon} MIL Status: {mil_text}",
                                     font=('Segoe UI', 11, 'bold'),
                                     fg=mil_color, bg='#404040')
                mil_label.pack(pady=5)

                count_label = tk.Label(mil_frame,
                                       text=f"📊 DTC Count: {dtc_count}",
                                       font=('Segoe UI', 10),
                                       fg='white', bg='#404040')
                count_label.pack(pady=(0, 5))

                # Check readiness monitors if available
                try:
                    monitors = getattr(status_data, 'available_tests', [])
                    completed = getattr(status_data, 'complete_tests', [])

                    if monitors:
                        readiness_frame = tk.Frame(
                            dtc_content_frame, bg='#333333', relief='raised', bd=1)
                        readiness_frame.pack(fill='x', padx=20, pady=10)

                        readiness_title = tk.Label(readiness_frame,
                                                   text="📋 READINESS MONITORS",
                                                   font=('Segoe UI',
                                                         11, 'bold'),
                                                   fg='#FFD700', bg='#333333')
                        readiness_title.pack(pady=5)

                        for monitor in monitors:
                            monitor_frame = tk.Frame(
                                readiness_frame, bg='#333333')
                            monitor_frame.pack(fill='x', padx=10, pady=2)

                            is_ready = monitor in completed
                            status_icon = "✅" if is_ready else "⏳"
                            status_text = "Ready" if is_ready else "Not Ready"
                            status_color = "#4CAF50" if is_ready else "#FF9800"

                            monitor_label = tk.Label(monitor_frame,
                                                     text=f"{status_icon} {monitor}: {status_text}",
                                                     font=('Segoe UI', 9),
                                                     fg=status_color, bg='#333333')
                            monitor_label.pack(anchor='w')

                except Exception as monitor_error:
                    print(f"Error reading readiness monitors: {monitor_error}")

                dtc_status_label.config(
                    text=f"● MIL: {mil_text}, DTCs: {dtc_count}", fg=mil_color)

            else:
                error_label = tk.Label(dtc_content_frame,
                                       text="⚠️ Unable to read MIL status\nCommand may not be supported",
                                       font=('Segoe UI', 11),
                                       fg='#FF9800', bg='#1e1e1e')
                error_label.pack(pady=20)
                dtc_status_label.config(
                    text="● MIL status unavailable", fg="#FF9800")

        except Exception as mil_error:
            print(f"Error checking MIL status: {mil_error}")
            error_label = tk.Label(dtc_content_frame,
                                   text=f"❌ Error reading MIL status:\n{str(mil_error)}",
                                   font=('Segoe UI', 11),
                                   fg='#F44336', bg='#1e1e1e')
            error_label.pack(pady=20)
            dtc_status_label.config(text="● MIL check failed", fg="#F44336")

    except Exception as e:
        dtc_status_label.config(text="● MIL check error", fg="#F44336")
        messagebox.showerror("MIL Status Error",
                             f"Error checking MIL status:\n{str(e)}")
        print(f"MIL status error: {e}")


def get_dtc_description(dtc_code):
    """Get description for DTC code"""
    dtc_descriptions = {
        # Common P0 codes (Powertrain)
        'P0000': 'No fault found',
        'P0001': 'Fuel Volume Regulator Control Circuit/Open',
        'P0002': 'Fuel Volume Regulator Control Circuit Range/Performance',
        'P0003': 'Fuel Volume Regulator Control Circuit Low',
        'P0004': 'Fuel Volume Regulator Control Circuit High',
        'P0005': 'Fuel Shutoff Valve A Control Circuit/Open',
        'P0010': 'A Camshaft Position Actuator Circuit (Bank 1)',
        'P0011': 'A Camshaft Position - Timing Over-Advanced or System Performance (Bank 1)',
        'P0012': 'A Camshaft Position - Timing Over-Retarded (Bank 1)',
        'P0013': 'B Camshaft Position - Actuator Circuit (Bank 1)',
        'P0014': 'B Camshaft Position - Timing Over-Advanced or System Performance (Bank 1)',
        'P0015': 'B Camshaft Position - Timing Over-Retarded (Bank 1)',
        'P0016': 'Crankshaft Position Camshaft Position Correlation (Bank 1 Sensor A)',
        'P0017': 'Crankshaft Position Camshaft Position Correlation (Bank 1 Sensor B)',
        'P0018': 'Crankshaft Position Camshaft Position Correlation (Bank 2 Sensor A)',
        'P0019': 'Crankshaft Position Camshaft Position Correlation (Bank 2 Sensor B)',
        'P0020': 'A Camshaft Position Actuator Circuit (Bank 2)',
        'P0100': 'Mass or Volume Air Flow Circuit Malfunction',
        'P0101': 'Mass or Volume Air Flow Circuit Range/Performance Problem',
        'P0102': 'Mass or Volume Air Flow Circuit Low Input',
        'P0103': 'Mass or Volume Air Flow Circuit High Input',
        'P0104': 'Mass or Volume Air Flow Circuit Intermittent',
        'P0105': 'Manifold Absolute Pressure/Barometric Pressure Circuit Malfunction',
        'P0106': 'Manifold Absolute Pressure/Barometric Pressure Circuit Range/Performance Problem',
        'P0107': 'Manifold Absolute Pressure/Barometric Pressure Circuit Low Input',
        'P0108': 'Manifold Absolute Pressure/Barometric Pressure Circuit High Input',
        'P0109': 'Manifold Absolute Pressure/Barometric Pressure Circuit Intermittent',
        'P0110': 'Intake Air Temperature Circuit Malfunction',
        'P0111': 'Intake Air Temperature Circuit Range/Performance Problem',
        'P0112': 'Intake Air Temperature Circuit Low Input',
        'P0113': 'Intake Air Temperature Circuit High Input',
        'P0114': 'Intake Air Temperature Circuit Intermittent',
        'P0115': 'Engine Coolant Temperature Circuit Malfunction',
        'P0116': 'Engine Coolant Temperature Circuit Range/Performance Problem',
        'P0117': 'Engine Coolant Temperature Circuit Low Input',
        'P0118': 'Engine Coolant Temperature Circuit High Input',
        'P0119': 'Engine Coolant Temperature Circuit Intermittent',
        'P0120': 'Throttle/Pedal Position Sensor/Switch A Circuit Malfunction',
        'P0121': 'Throttle/Pedal Position Sensor/Switch A Circuit Range/Performance Problem',
        'P0122': 'Throttle/Pedal Position Sensor/Switch A Circuit Low Input',
        'P0123': 'Throttle/Pedal Position Sensor/Switch A Circuit High Input',
        'P0124': 'Throttle/Pedal Position Sensor/Switch A Circuit Intermittent',
        'P0125': 'Insufficient Coolant Temperature for Closed Loop Fuel Control',
        'P0130': 'O2 Sensor Circuit Malfunction (Bank 1 Sensor 1)',
        'P0131': 'O2 Sensor Circuit Low Voltage (Bank 1 Sensor 1)',
        'P0132': 'O2 Sensor Circuit High Voltage (Bank 1 Sensor 1)',
        'P0133': 'O2 Sensor Circuit Slow Response (Bank 1 Sensor 1)',
        'P0134': 'O2 Sensor Circuit No Activity Detected (Bank 1 Sensor 1)',
        'P0135': 'O2 Sensor Heater Circuit Malfunction (Bank 1 Sensor 1)',
        'P0171': 'System Too Lean (Bank 1)',
        'P0172': 'System Too Rich (Bank 1)',
        'P0174': 'System Too Lean (Bank 2)',
        'P0175': 'System Too Rich (Bank 2)',
        'P0300': 'Random/Multiple Cylinder Misfire Detected',
        'P0301': 'Cylinder 1 Misfire Detected',
        'P0302': 'Cylinder 2 Misfire Detected',
        'P0303': 'Cylinder 3 Misfire Detected',
        'P0304': 'Cylinder 4 Misfire Detected',
        'P0305': 'Cylinder 5 Misfire Detected',
        'P0306': 'Cylinder 6 Misfire Detected',
        'P0307': 'Cylinder 7 Misfire Detected',
        'P0308': 'Cylinder 8 Misfire Detected',
        'P0420': 'Catalyst System Efficiency Below Threshold (Bank 1)',
        'P0430': 'Catalyst System Efficiency Below Threshold (Bank 2)',
        'P0440': 'Evaporative Emission Control System Malfunction',
        'P0441': 'Evaporative Emission Control System Incorrect Purge Flow',
        'P0442': 'Evaporative Emission Control System Leak Detected (small leak)',
        'P0443': 'Evaporative Emission Control System Purge Control Valve Circuit Malfunction',
        'P0446': 'Evaporative Emission Control System Vent Control Circuit Malfunction',
        'P0455': 'Evaporative Emission Control System Leak Detected (gross leak)',
        'P0500': 'Vehicle Speed Sensor Malfunction',
        'P0505': 'Idle Control System Malfunction',
        'P0506': 'Idle Control System RPM Lower Than Expected',
        'P0507': 'Idle Control System RPM Higher Than Expected',
        'P0600': 'Serial Communication Link Malfunction',
        'P0601': 'Internal Control Module Memory Check Sum Error',
        'P0602': 'Control Module Programming Error',
        'P0603': 'Internal Control Module Keep Alive Memory (KAM) Error',
        'P0604': 'Internal Control Module Random Access Memory (RAM) Error',
        'P0605': 'Internal Control Module Read Only Memory (ROM) Error',
        'P0700': 'Transmission Control System Malfunction',
        'P0701': 'Transmission Control System Range/Performance',
        'P0702': 'Transmission Control System Electrical',
        'P0703': 'Torque Converter/Brake Switch B Circuit Malfunction',
        'P0704': 'Clutch Switch Input Circuit Malfunction',
        'P0705': 'Transmission Range Sensor Circuit Malfunction (PRNDL Input)',
        # Add more as needed
    }

    # Clean the DTC code (remove any extra formatting)
    clean_code = str(dtc_code).strip().upper()

    # Return description if found, otherwise return generic description
    return dtc_descriptions.get(clean_code, f"Unknown DTC - Consult service manual for {clean_code}")


# Function to connect to the vehicle


def connect_to_vehicle():
    global connection
    try:
        status_label.config(text="● Connecting...", fg="#FFA726")

        # Get selected device from dropdown
        selected_device = get_selected_device()
        conn_type = connection_type.get()

        # Validate connection type and device selection
        if conn_type == "Bluetooth" and not BLUETOOTH_AVAILABLE:
            messagebox.showerror("Bluetooth Not Available",
                                 "Bluetooth support not installed.\n\nTo add Bluetooth support:\n1. pip install bleak\n2. Restart the application")
            status_label.config(text="● Not Connected", fg="#FF4444")
            return

        if not selected_device:
            # Fallback to manual entry if no device selected
            port_or_address = com_port_var.get().strip()
            if not port_or_address:
                messagebox.showerror("No Device Selected",
                                     "Please select a device from the dropdown or scan for devices first.\n\n"
                                     "Click 'Scan Devices' to find available adapters.")
                status_label.config(text="● Not Connected", fg="#FF4444")
                return
            print(f"Using manual entry: {port_or_address}")
        else:
            port_or_address = selected_device['port']
            device_type = selected_device['type']
            print(
                f"Using selected device: {selected_device['display']} ({device_type})")

        # Validate port/address format
        if not port_or_address or port_or_address.strip() == "":
            messagebox.showerror("Invalid Device",
                                 "Invalid port or address. Please select a valid device.")
            status_label.config(text="● Not Connected", fg="#FF4444")
            return

        if conn_type == "Serial":
            # Serial/USB connection
            port_or_address = port_or_address.upper()
            try:
                status_label.config(
                    text=f"● Connecting to {port_or_address}...", fg="#FFA726")
                connection = obd.OBD(port_or_address)  # Try specified COM port
            except:
                status_label.config(
                    text="● Auto-detecting serial ports...", fg="#FFA726")
                connection = obd.OBD()  # Fall back to auto-detect

        elif conn_type == "Bluetooth":
            # Bluetooth connection
            try:
                if not port_or_address:
                    raise ValueError("Bluetooth device required")

                print(f"🔵 BLUETOOTH CONNECTION ATTEMPT:")
                print(f"  Port/Address: {port_or_address}")
                print(
                    f"  Device Type: {selected_device['type'] if selected_device else 'Manual Entry'}")
                print(
                    f"  Device Info: {selected_device['display'] if selected_device else 'N/A'}")

                # Force COM7 connection if detected
                if port_or_address.upper() == 'COM7' or 'COM7' in str(port_or_address).upper():
                    print("🎯 Detected COM7 - forcing optimized Bluetooth connection")
                    port_or_address = 'COM7'  # Normalize the port name

                    status_label.config(
                        text="● Connecting to COM7 (Bluetooth OBD)...", fg="#FFA726")

                    # COM7 optimized connection sequence
                    connection_successful = False
                    optimal_settings = [
                        {'baudrate': 38400, 'timeout': 10, 'fast': False},
                        {'baudrate': 9600, 'timeout': 15, 'fast': False},
                        {'baudrate': 115200, 'timeout': 8, 'fast': False},
                        {'baudrate': 57600, 'timeout': 12, 'fast': False}
                    ]

                    for i, settings in enumerate(optimal_settings):
                        try:
                            print(
                                f"COM7 attempt {i+1}: baudrate={settings['baudrate']}, timeout={settings['timeout']}")
                            status_label.config(
                                text=f"● COM7 attempt {i+1}: {settings['baudrate']} baud...", fg="#FFA726")

                            connection = obd.OBD(
                                'COM7',
                                baudrate=settings['baudrate'],
                                timeout=settings['timeout'],
                                fast=settings['fast']
                            )

                            if connection and connection.is_connected():
                                print(
                                    f"✅ COM7 connected successfully with {settings['baudrate']} baud")
                                connection_successful = True
                                break
                            else:
                                if connection:
                                    connection.close()
                                connection = None
                                print(
                                    f"❌ COM7 attempt {i+1} failed - no connection established")

                        except Exception as com7_error:
                            print(f"❌ COM7 attempt {i+1} failed: {com7_error}")
                            if connection:
                                try:
                                    connection.close()
                                except:
                                    pass
                                connection = None
                            continue

                    if not connection_successful:
                        raise Exception(
                            f"All COM7 connection attempts failed. Ensure vehicle is running and COM7 is available.")

                # Determine connection method based on device type
                elif selected_device and selected_device['type'] == 'bluetooth_com':
                    # Bluetooth COM port - most reliable method
                    status_label.config(
                        text=f"● Connecting via Bluetooth COM port {port_or_address}...", fg="#FFA726")

                    try:
                        # Special handling for COM7 (common Bluetooth OBD port)
                        if port_or_address.upper() == 'COM7':
                            print(
                                f"Detected COM7 - using optimized Bluetooth connection settings")
                            status_label.config(
                                text="● Connecting to COM7 with Bluetooth settings...", fg="#FFA726")

                            # COM7 optimized connection sequence
                            connection_successful = False

                            # Try the most common working settings for COM7 first
                            optimal_settings = [
                                {'baudrate': 38400, 'timeout': 10, 'fast': False},
                                {'baudrate': 9600, 'timeout': 15, 'fast': False},
                                {'baudrate': 115200, 'timeout': 8, 'fast': False},
                                {'baudrate': 57600, 'timeout': 12, 'fast': False}
                            ]

                            for i, settings in enumerate(optimal_settings):
                                try:
                                    print(
                                        f"COM7 attempt {i+1}: baudrate={settings['baudrate']}, timeout={settings['timeout']}")
                                    status_label.config(
                                        text=f"● COM7 attempt {i+1}: {settings['baudrate']} baud...", fg="#FFA726")

                                    connection = obd.OBD(
                                        port_or_address,
                                        baudrate=settings['baudrate'],
                                        timeout=settings['timeout'],
                                        fast=settings['fast']
                                    )

                                    if connection and connection.is_connected():
                                        print(
                                            f"✓ COM7 connected successfully with {settings['baudrate']} baud")
                                        connection_successful = True
                                        break
                                    else:
                                        if connection:
                                            connection.close()
                                        connection = None

                                except Exception as com7_error:
                                    print(
                                        f"COM7 attempt {i+1} failed: {com7_error}")
                                    if connection:
                                        try:
                                            connection.close()
                                        except:
                                            pass
                                        connection = None
                                    continue

                            if not connection_successful:
                                raise Exception(
                                    f"All COM7 connection attempts failed. Check that vehicle is running and Bluetooth adapter is paired.")

                        else:
                            # Standard Bluetooth COM port connection for other ports
                            # Try with standard settings first
                            connection = obd.OBD(
                                port_or_address, fast=False, timeout=10)
                            if connection and connection.is_connected():
                                print(
                                    f"✓ Successfully connected to {port_or_address}")
                            else:
                                if connection:
                                    connection.close()
                                connection = None

                                # Try with different baudrates if standard connection failed
                                print(
                                    f"Standard connection failed, trying different baudrates...")
                                baudrates = [38400, 9600, 115200, 57600]

                                for baudrate in baudrates:
                                    try:
                                        print(f"Trying baudrate {baudrate}...")
                                        status_label.config(
                                            text=f"● Trying {baudrate} baud on {port_or_address}...", fg="#FFA726")

                                        connection = obd.OBD(
                                            port_or_address, baudrate=baudrate, fast=False, timeout=10)
                                        if connection and connection.is_connected():
                                            print(
                                                f"✓ Connected with baudrate {baudrate}")
                                            break
                                        else:
                                            if connection:
                                                connection.close()
                                            connection = None
                                    except Exception as baud_error:
                                        print(
                                            f"Baudrate {baudrate} failed: {baud_error}")
                                        if connection:
                                            try:
                                                connection.close()
                                            except:
                                                pass
                                            connection = None
                                        continue

                                if not connection:
                                    raise Exception(
                                        f"Failed to connect to Bluetooth COM port {port_or_address}")

                    except Exception as com_error:
                        print(
                            f"Bluetooth COM port connection error: {com_error}")
                        raise Exception(
                            f"Failed to connect to Bluetooth COM port {port_or_address}: {str(com_error)}")

                elif selected_device and selected_device['type'] == 'bluetooth_mac':
                    # MAC addresses are not supported - redirect to COM port
                    status_label.config(
                        text="● MAC address not supported", fg="#FF4444")

                    error_msg = f"❌ MAC address connection not supported!\n\n"
                    error_msg += f"MAC Address: {port_or_address}\n\n"
                    error_msg += f"📱 SOLUTION - Use COM port instead:\n"
                    error_msg += f"1. Open Device Manager\n"
                    error_msg += f"2. Look for 'Ports (COM & LPT)'\n"
                    error_msg += f"3. Find 'Standard Serial over Bluetooth'\n"
                    error_msg += f"4. Note the COM port (e.g., COM7)\n"
                    error_msg += f"5. Use 'Quick Connect COM7' button\n\n"
                    error_msg += f"💡 TIP: COM ports are much more reliable than MAC addresses!"

                    messagebox.showerror(
                        "MAC Address Not Supported", error_msg)
                    return
                else:
                    # Manual address entry
                    if port_or_address.upper().startswith('COM'):
                        connection = obd.OBD(port_or_address)
                    else:
                        connection = obd.OBD(port_or_address)

            except Exception as bt_error:
                print(f"Bluetooth connection failed: {bt_error}")
                status_label.config(
                    text="● Bluetooth connection failed", fg="#FF4444")

                # Enhanced error message with troubleshooting
                error_msg = f"Failed to connect via Bluetooth:\n{str(bt_error)}\n\n"
                error_msg += "Troubleshooting Steps:\n"
                error_msg += "1. Ensure OBD adapter is powered and paired in Windows\n"
                error_msg += "2. Check Device Manager for Bluetooth COM ports\n"
                error_msg += "3. Try using the COM port instead of MAC address\n"
                error_msg += "4. Verify adapter compatibility with ELM327 protocol\n"
                error_msg += "5. Try restarting Windows Bluetooth service\n"
                error_msg += "6. Ensure no other software is using the adapter"

                messagebox.showerror("Bluetooth Connection Error", error_msg)
                return

        # Final connection validation
        if connection and connection.is_connected():
            # Verify the connection is actually working by testing a basic command
            try:
                status_label.config(
                    text="● Verifying connection...", fg="#FFA726")

                # Try to query a basic OBD command to verify the connection is working
                test_response = None
                try:
                    # Use getattr to avoid type checking issues
                    rpm_cmd = getattr(obd.commands, 'RPM', None)
                    if rpm_cmd:
                        test_response = connection.query(rpm_cmd)
                    else:
                        speed_cmd = getattr(obd.commands, 'SPEED', None)
                        if speed_cmd:
                            test_response = connection.query(speed_cmd)
                        else:
                            print("⚠ No standard commands found for verification")

                except Exception as cmd_error:
                    print(f"Command query failed: {cmd_error}")

                if test_response and test_response.value is not None:
                    print(
                        f"✓ OBD connection verified - Test response: {test_response.value}")
                else:
                    print(
                        "⚠ Connection established but test query failed - proceeding anyway")

            except Exception as verify_error:
                print(f"Connection verification warning: {verify_error}")
                # Don't fail the connection for verification issues - just warn

            # Get display name for connection info
            if selected_device:
                device_display = selected_device['display']
                port_info = f"{conn_type} - {device_display[:30]}..."
            else:
                port_info = f"{conn_type} ({port_or_address})"

            status_label.config(
                text=f"● Connected: {port_info}", fg="#4CAF50")

            # Show connection success message
            success_msg = f"✅ Successfully connected!\n\n"
            success_msg += f"Connection Type: {conn_type}\n"
            if selected_device:
                success_msg += f"Device: {selected_device['display']}\n"
                success_msg += f"Port: {selected_device['port']}\n"
            else:
                success_msg += f"Port/Address: {port_or_address}\n"

            success_msg += f"\n🔧 You can now:\n"
            success_msg += f"• View real-time VE table data\n"
            success_msg += f"• Scan and monitor OBD PIDs\n"
            success_msg += f"• Access all sensor readings"

            messagebox.showinfo("Connection Successful", success_msg)

            # Print available commands for debugging
            try:
                available_commands = [cmd for cmd in dir(
                    obd.commands) if not cmd.startswith('_')]
                print(f"Available OBD commands: {len(available_commands)}")
                # Show first 10 commands
                print(f"Sample commands: {available_commands[:10]}")

                # Look for pressure-related commands
                pressure_commands = [
                    cmd for cmd in available_commands if 'PRESSURE' in cmd or 'MANIFOLD' in cmd or 'MAP' in cmd]
                print(f"Pressure-related commands: {pressure_commands}")

                # Look for intake-related commands
                intake_commands = [
                    cmd for cmd in available_commands if 'INTAKE' in cmd]
                print(f"Intake-related commands: {intake_commands}")

            except:
                print("Could not list OBD commands")

            update()  # Start updating the table and sensor data
        else:
            status_label.config(text="● Connection Failed", fg="#FF4444")
            error_msg = f"❌ Failed to connect to device\n\n"

            if selected_device:
                error_msg += f"Device: {selected_device['display']}\n"
                error_msg += f"Port: {selected_device['port']}\n"
                error_msg += f"Type: {selected_device['type']}\n\n"

            error_msg += f"🔧 TROUBLESHOOTING:\n"
            error_msg += f"1. Ensure vehicle is running (powers OBD adapter)\n"
            error_msg += f"2. Check that no other OBD software is connected\n"
            error_msg += f"3. Try a different device from the dropdown\n"
            error_msg += f"4. For Bluetooth: Check pairing in Windows settings\n"
            error_msg += f"5. Try 'Scan Devices' to refresh the list\n"
            error_msg += f"6. Restart the OBD adapter (unplug/replug)\n"

            messagebox.showerror("Connection Failed", error_msg)

    except Exception as e:
        status_label.config(text=f"● Error: Connection Failed", fg="#FF4444")

        error_details = str(e)
        error_msg = f"❌ Connection Error\n\n"
        error_msg += f"Error: {error_details}\n\n"

        if selected_device:
            error_msg += f"Selected Device: {selected_device['display']}\n"
            error_msg += f"Port: {selected_device['port']}\n\n"

        error_msg += f"💡 SOLUTIONS:\n"
        if "No such file or directory" in error_details or "could not open port" in error_details:
            error_msg += f"• Port may be in use by another application\n"
            error_msg += f"• Check Device Manager for correct COM port\n"
            error_msg += f"• Try scanning for devices again\n"
        elif "timeout" in error_details.lower():
            error_msg += f"• Ensure vehicle is running\n"
            error_msg += f"• Check OBD adapter power LED\n"
            error_msg += f"• Try a different baud rate\n"
        elif "bluetooth" in error_details.lower():
            error_msg += f"• Install Bluetooth support: pip install bleak\n"
            error_msg += f"• Check Windows Bluetooth settings\n"
            error_msg += f"• Ensure adapter is paired in Windows\n"
        else:
            error_msg += f"• Check device connection\n"
            error_msg += f"• Try selecting a different device\n"
            error_msg += f"• Restart the application\n"

        messagebox.showerror("Connection Error", error_msg)

# Function to update sensor data and table


def update():
    # Check which tab is currently selected
    current_tab = notebook.index(notebook.select())
    ve_table_tab_index = 0  # VE Table tab is the first tab (index 0)

    # Update tab text to show monitoring status
    if current_tab == ve_table_tab_index and (demo_mode or (connection and connection.is_connected())):
        # Active indicator
        notebook.tab(ve_table_tab_index, text='📊 VE Table Monitor ●')
    else:
        notebook.tab(ve_table_tab_index,
                     text='📊 VE Table Monitor')   # Inactive

    if demo_mode:
        # Only generate VE table data if VE table tab is active
        if current_tab == ve_table_tab_index:
            # Generate simulated data for demo
            rpm = random.uniform(800, 6000)
            map_kpa = random.uniform(25, 95)
            iat_c = random.uniform(15, 50)
            maf_gps = random.uniform(2, 15)

            # Calculations
            iat_k = iat_c + 273.15  # Convert IAT to Kelvin
            # Air mass per cylinder (g/cyl) for 8-cylinder engine
            g_cyl = (maf_gps * 15) / rpm if rpm > 0 else 0
            # VE in Grams*Kelvin/kPa
            ve = (g_cyl * iat_k) / map_kpa if map_kpa > 0 else 0

            print(f"DEMO - RPM: {rpm:.0f}, MAP: {map_kpa:.1f}, VE: {ve:.2f}")

            # Update sensor value labels
            labels["RPM"].config(text=f"{rpm:.0f}")
            labels["MAP"].config(text=f"{map_kpa:.1f}")
            labels["IAT"].config(text=f"{iat_k:.1f}")
            labels["MAF"].config(text=f"{maf_gps:.2f}")
            labels["g/cyl"].config(text=f"{g_cyl:.4f}")
            labels["VE"].config(text=f"{ve:.2f}")

            # Find nearest cell in the table
            i = min(range(20), key=lambda k: abs(rpm - RPM_VALUES[k]))
            j = min(range(19), key=lambda k: abs(map_kpa - MAP_VALUES[k]))

            print(f"DEMO - Updating cell [{j}][{i}] with VE: {ve:.2f}")

            # Update cell with VE value and color
            cells[j][i].config(text=f"{ve:.2f}", bg=ve_to_color(ve))

            # Highlight current cell by making text bold and adding border
            for row in cells:
                for cell in row:
                    cell.config(font=('Segoe UI', 8), relief='flat')
            cells[j][i].config(font=('Segoe UI', 8, 'bold'),
                               relief='raised', bd=2)

        # Schedule next update - faster for VE table, slower for other tabs
        update_interval = 500 if current_tab == ve_table_tab_index else 2000
        window.after(update_interval, update)

    elif connection and connection.is_connected():
        # Only query VE table sensors when VE table tab is active
        if current_tab == ve_table_tab_index:
            try:
                # Query sensors with safe command access for VE table calculations
                rpm_resp = None
                map_resp = None
                iat_resp = None
                maf_resp = None

                try:
                    # Try to access RPM command safely
                    rpm_resp = connection.query(
                        getattr(obd.commands, 'RPM', None))
                    if rpm_resp is None:
                        print("RPM command not available")
                except Exception as e:
                    print(f"RPM query failed: {e}")

                try:
                    # Try different MAP command names - use the ones we found
                    map_cmd = None
                    possible_names = [
                        'INTAKE_PRESSURE',  # This one was found in available commands
                        'INTAKE_MANIFOLD_PRESSURE',
                        'INTAKE_MANIFOLD_ABS_PRESSURE',
                        'MANIFOLD_PRESSURE',
                        'INTAKE_MANIFOLD_ABSOLUTE_PRESSURE',
                        'MAP'
                    ]

                    for cmd_name in possible_names:
                        cmd = getattr(obd.commands, cmd_name, None)
                        if cmd is not None:
                            map_cmd = cmd
                            print(f"Found MAP command: {cmd_name}")
                            break

                    if map_cmd:
                        map_resp = connection.query(map_cmd)
                    else:
                        print("MAP command not available - searching all commands...")
                        # If still not found, search all available commands for pressure-related ones
                        all_commands = [cmd for cmd in dir(
                            obd.commands) if not cmd.startswith('_')]
                        pressure_commands = [cmd for cmd in all_commands if 'PRESSURE' in cmd.upper(
                        ) or 'MANIFOLD' in cmd.upper()]
                        print(
                            f"Available pressure commands: {pressure_commands}")

                except Exception as e:
                    print(f"MAP query failed: {e}")

                try:
                    # Try to access IAT command safely
                    iat_resp = connection.query(
                        getattr(obd.commands, 'INTAKE_TEMP', None))
                    if iat_resp is None:
                        print("IAT command not available")
                except Exception as e:
                    print(f"IAT query failed: {e}")

                try:
                    # Try to access MAF command safely
                    maf_resp = connection.query(
                        getattr(obd.commands, 'MAF', None))
                    if maf_resp is None:
                        print("MAF command not available")
                except Exception as e:
                    print(f"MAF query failed: {e}")

                # Extract values safely, default to 0 if unavailable
                def safe_extract_value(response):
                    """Safely extract magnitude from OBD response"""
                    try:
                        if response and not response.is_null() and response.value is not None:
                            if hasattr(response.value, 'magnitude'):
                                return response.value.magnitude
                            else:
                                # If no magnitude attribute, try to convert to float
                                return float(response.value)
                    except:
                        pass
                    return 0

                rpm = safe_extract_value(rpm_resp)
                map_kpa = safe_extract_value(map_resp)
                iat_c = safe_extract_value(iat_resp)
                maf_gps = safe_extract_value(maf_resp)

                # Debug output to see what values we're getting
                print(
                    f"VE TABLE - RPM: {rpm}, MAP: {map_kpa}, IAT: {iat_c}, MAF: {maf_gps}")

                # Calculations
                iat_k = iat_c + 273.15  # Convert IAT to Kelvin
                # Air mass per cylinder (g/cyl) for 8-cylinder engine
                g_cyl = (maf_gps * 15) / rpm if rpm > 0 else 0
                # VE in Grams*Kelvin/kPa
                ve = (g_cyl * iat_k) / map_kpa if map_kpa > 0 else 0

                # Update sensor value labels
                labels["RPM"].config(text=f"{rpm:.0f}")
                labels["MAP"].config(text=f"{map_kpa:.1f}")
                labels["IAT"].config(text=f"{iat_k:.1f}")
                labels["MAF"].config(text=f"{maf_gps:.2f}")
                labels["g/cyl"].config(text=f"{g_cyl:.4f}")
                labels["VE"].config(text=f"{ve:.2f}")

                # Find nearest cell in the table
                i = min(range(20), key=lambda k: abs(rpm - RPM_VALUES[k]))
                j = min(range(19), key=lambda k: abs(map_kpa - MAP_VALUES[k]))

                # Update cell with VE value and color
                cells[j][i].config(text=f"{ve:.2f}", bg=ve_to_color(ve))

                # Highlight current cell by making text bold and adding border
                for row in cells:
                    for cell in row:
                        cell.config(font=('Segoe UI', 8), relief='flat')
                cells[j][i].config(font=('Segoe UI', 8, 'bold'),
                                   relief='raised', bd=2)

            except Exception as e:
                print(f"Error in VE table update: {e}")
        else:
            # Not on VE table tab - just maintain basic connection status
            print(f"Skipping VE table queries - user on tab {current_tab}")

        # Schedule next update with adaptive timing
        if current_tab == ve_table_tab_index:
            # Fast updates for VE table when it's active
            # Increased from 50ms to reduce OBDX load
            window.after(200, update)
        else:
            # Much slower updates when not on VE table tab to reduce connection load
            window.after(5000, update)  # Only check connection every 5 seconds
    else:
        if not demo_mode:
            status_label.config(text="● Disconnected", fg="#FF4444")
        # Check for reconnection less frequently when disconnected
        window.after(2000, update)


# Handle window close to ensure OBD connection is closed
window.protocol("WM_DELETE_WINDOW", lambda: (
    connection.close() if connection else None, window.destroy()))

# Initialize the device list on startup
try:
    populate_device_list()
except:
    print("Could not populate initial device list")

# Start the GUI event loop
window.mainloop()
