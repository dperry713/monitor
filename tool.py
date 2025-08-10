import sys
import time
import json
import math
import queue
import threading
import csv
import platform
import re
import subprocess
from collections import deque
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QLineEdit, QSpinBox,
    QVBoxLayout, QHBoxLayout, QGridLayout, QTableWidget, QTableWidgetItem, QFileDialog,
    QMessageBox, QHeaderView, QComboBox
)
from PyQt5.QtCore import Qt, QTimer, pyqtSlot as Slot
import pyqtgraph as pg
from typing import Optional, Union, Dict, List
from obd import OBDStatus
import serial
import serial.tools.list_ports
import obd
import os
os.environ['PYQTGRAPH_QT_LIB'] = 'PyQt5'

# Bluetooth imports (platform-specific)
try:
    if platform.system() == "Windows":
        import socket
        import subprocess
        try:
            import bluetooth  # PyBluez  # type: ignore
            PYBLUEZ_AVAILABLE = True
        except ImportError:
            PYBLUEZ_AVAILABLE = False
    else:
        # For Linux/macOS, we'd use different libraries
        PYBLUEZ_AVAILABLE = False
except ImportError:
    PYBLUEZ_AVAILABLE = False

"""
Advanced Python OBD-II Tuning Tool (J1850 support + real-time VE table + Bluetooth RFCOMM)

Features implemented in this file:
- Support forcing ELM327 protocol to SAE J1850 PWM (1) or VPW (2) via AT SP n.
  References for ELM327 protocol numbers and AT SP usage: ELM327 documentation and reference guides.
  See: ELM327 protocol numbers and ATSP usage. :contentReference[oaicite:0]{index=0}

- Connect to ELM327-compatible adapter over serial (USB), Bluetooth RFCOMM, or TCP (WiFi ELM327).
  For serial, optionally send "ATSPn" before python-OBD connect to force J1850 protocol.
  
- Bluetooth RFCOMM support for wireless OBD adapters:
    * Automatic discovery of Bluetooth devices
    * Smart detection of OBD-likely devices
    * RFCOMM connection testing and establishment
    * Cross-platform Bluetooth support (Windows optimized)

- Real-time Volumetric Efficiency (VE) table population:
    VE units: grams * Kelvin / kPa (g*K/kPa)
    Formula used (per user spec):
        g_per_cyl = (MAF_g_per_s * 120) / (RPM * cylinders)
        VE = (g_per_cyl * Charge_Temperature_K) / MAP_kPa
    Where:
      - MAF_g_per_s: Mass Air Flow in grams/second (requires MAF PID)
      - RPM: engine speed in revolutions/minute
      - cylinders: user-configurable number of cylinders
      - Charge_Temperature_K = IAT_C + 273.15
      - MAP_kPa: Manifold Absolute Pressure in kilopascal (kPa)
    The table grid (rows=MAP bins, cols=RPM bins) is updated in real time when required inputs exist.

- GUI: PyQt5 + pyqtgraph, includes QTableWidget that displays VE values (formatted) and highlights most-recent bin.
- CSV Export: Export VE tables and data logs for analysis in external tools.

Requirements:
  pip install pyserial pyqt5 pyqtgraph python-OBD pybluez (Windows)

Run:
  python tool.py
"""


# GUI

# Plotting

# OBD & serial/tcp

# ----------------------------
# Utility: force ELM327 protocol
# ----------------------------

def scan_available_ports():
    """
    Scan for available serial ports and return list of port info dictionaries.
    Returns list of dicts with 'device', 'description', 'hwid' keys.
    """
    ports = []
    try:
        for port in serial.tools.list_ports.comports():
            port_info = {
                'device': port.device,
                'description': port.description or 'Unknown',
                'hwid': port.hwid or 'Unknown',
                'type': 'serial'
            }
            ports.append(port_info)
    except Exception as e:
        print(f"Error scanning serial ports: {e}")
    return ports


def scan_bluetooth_devices():
    """
    Scan for available Bluetooth devices that might be OBD adapters.
    Includes both discoverable and paired devices.
    Returns list of dicts with 'device', 'description', 'hwid' keys.
    """
    bluetooth_devices = []
    found_addresses = set()  # Track found devices to avoid duplicates

    if platform.system() == "Windows":
        # Method 1: Try PyBluez for discoverable devices
        if PYBLUEZ_AVAILABLE:
            try:
                import bluetooth  # type: ignore
                print("Scanning for discoverable Bluetooth devices...")
                devices = bluetooth.discover_devices(
                    lookup_names=True, lookup_class=True, duration=8)

                for addr, name, device_class in devices:
                    if addr in found_addresses:
                        continue
                    found_addresses.add(addr)

                    # Check if this might be an OBD device (be more permissive)
                    device_name_lower = name.lower() if name else ''
                    is_obd_likely = any(keyword in device_name_lower for keyword in
                                        ['obd', 'elm', 'obdii', 'eobd', 'diagnostic', 'v1.5', 'v2.1'])

                    device_info = {
                        'device': f"BT:{addr}",
                        'description': f"Bluetooth: {name or 'Unknown'}",
                        'hwid': f"Bluetooth_{addr}",
                        'type': 'bluetooth',
                        'connected': False,  # Will check later
                        'bt_address': addr,
                        'bt_name': name or 'Unknown',
                        'source': 'discoverable'
                    }

                    if is_obd_likely:
                        device_info['description'] += " (OBD likely)"

                    bluetooth_devices.append(device_info)
                    print(f"Found discoverable device: {name} ({addr})")

            except Exception as e:
                print(f"PyBluez discovery error: {e}")

        # Method 2: Use Windows netsh to find paired devices (more comprehensive)
        try:
            print("Scanning for paired Bluetooth devices...")
            result = subprocess.run(
                ['netsh', 'bluetooth', 'show', 'device'],
                capture_output=True, text=True, timeout=15
            )

            if result.returncode == 0:
                lines = result.stdout.split('\n')
                current_device = {}

                for line in lines:
                    line = line.strip()
                    if 'Device name:' in line:
                        current_device['name'] = line.split(':', 1)[1].strip()
                    elif 'Device address:' in line:
                        current_device['address'] = line.split(':', 1)[
                            1].strip()
                    elif 'Device type:' in line:
                        current_device['type'] = line.split(':', 1)[1].strip()
                    elif 'Connected:' in line:
                        current_device['connected'] = 'Yes' in line

                        # Process device if we have complete info
                        if 'name' in current_device and 'address' in current_device:
                            bt_address = current_device.get('address', '')

                            # Skip if already found by PyBluez
                            if bt_address in found_addresses:
                                current_device = {}
                                continue
                            found_addresses.add(bt_address)

                            device_name = current_device.get(
                                'name', '').lower()
                            # Be more permissive with OBD detection
                            is_obd_likely = any(keyword in device_name for keyword in
                                                ['obd', 'elm', 'obdii', 'eobd', 'diagnostic', 'v1.5', 'v2.1', 'scanner'])

                            if bt_address:
                                device_info = {
                                    'device': f"BT:{bt_address}",
                                    'description': f"Bluetooth: {current_device.get('name', 'Unknown')}",
                                    'hwid': f"Bluetooth_{bt_address}",
                                    'type': 'bluetooth',
                                    'connected': current_device.get('connected', False),
                                    'bt_address': bt_address,
                                    'bt_name': current_device.get('name', 'Unknown'),
                                    'source': 'paired'
                                }

                                if is_obd_likely:
                                    device_info['description'] += " (OBD likely)"

                                bluetooth_devices.append(device_info)
                                print(
                                    f"Found paired device: {current_device.get('name')} ({bt_address}) - Connected: {current_device.get('connected', False)}")

                        current_device = {}

        except Exception as e:
            print(f"Error scanning paired Bluetooth devices with netsh: {e}")

        # Method 3: Try PowerShell for additional device discovery
        try:
            print("Scanning with PowerShell for additional Bluetooth devices...")
            ps_result = subprocess.run([
                'powershell', '-Command',
                'Get-PnpDevice | Where-Object {$_.Class -eq "Bluetooth" -and $_.Status -eq "OK"} | Select-Object Name, InstanceId'
            ], capture_output=True, text=True, timeout=10)

            if ps_result.returncode == 0:
                lines = ps_result.stdout.split('\n')
                for line in lines[3:]:  # Skip header lines
                    if line.strip() and not line.startswith('-'):
                        parts = line.strip().split(None, 1)
                        if len(parts) >= 2:
                            name = parts[0]
                            # Extract address if available in InstanceId
                            if 'BTHENUM' in line:
                                # Parse Bluetooth address from Windows device instance
                                try:
                                    # Look for pattern like DEV_XXXXXXXXXXXX
                                    import re
                                    match = re.search(
                                        r'DEV_([0-9A-F]{12})', line, re.IGNORECASE)
                                    if match:
                                        addr_raw = match.group(1)
                                        # Convert to standard MAC format
                                        addr = ':'.join(
                                            addr_raw[i:i+2] for i in range(0, 12, 2))

                                        if addr not in found_addresses:
                                            found_addresses.add(addr)
                                            device_info = {
                                                'device': f"BT:{addr}",
                                                'description': f"Bluetooth: {name}",
                                                'hwid': f"Bluetooth_{addr}",
                                                'type': 'bluetooth',
                                                'connected': True,  # PowerShell shows active devices
                                                'bt_address': addr,
                                                'bt_name': name,
                                                'source': 'powershell'
                                            }
                                            bluetooth_devices.append(
                                                device_info)
                                            print(
                                                f"Found PowerShell device: {name} ({addr})")
                                except Exception:
                                    pass

        except Exception as e:
            print(f"PowerShell Bluetooth scan error: {e}")

    print(f"Total Bluetooth devices found: {len(bluetooth_devices)}")
    return bluetooth_devices


def scan_all_ports():
    """
    Scan for both serial and Bluetooth ports.
    Returns combined list of all available ports.
    """
    all_ports = []

    # Get serial ports
    serial_ports = scan_available_ports()
    all_ports.extend(serial_ports)

    # Get Bluetooth devices
    bluetooth_devices = scan_bluetooth_devices()
    all_ports.extend(bluetooth_devices)

    return all_ports


def test_obd_connection(port_name, timeout=2.0):
    """
    Test if a port has a working OBD adapter by attempting a quick connection.
    Supports both serial and Bluetooth RFCOMM connections.
    Returns True if OBD adapter responds, False otherwise.
    """
    try:
        # Handle Bluetooth RFCOMM connections
        if port_name.startswith('BT:'):
            return test_bluetooth_obd_connection(port_name, timeout)

        # Handle regular serial connections
        connection = obd.OBD(port_name, fast=True, timeout=timeout)
        if connection and connection.status() == OBDStatus.CAR_CONNECTED:
            connection.close()
            return True
        elif connection:
            connection.close()
        return False
    except Exception:
        return False


def test_bluetooth_obd_connection(bt_device, timeout=2.0):
    """
    Test Bluetooth RFCOMM OBD connection.
    bt_device format: "BT:XX:XX:XX:XX:XX:XX"
    """
    try:
        if platform.system() == "Windows":
            # Extract Bluetooth address
            bt_address = bt_device.replace('BT:', '')

            # Try PyBluez first if available
            if PYBLUEZ_AVAILABLE:
                try:
                    import bluetooth  # type: ignore

                    # Try to find the RFCOMM service
                    services = bluetooth.find_service(address=bt_address)
                    rfcomm_channel = None

                    for service in services:
                        if service.get('protocol') == 'RFCOMM':
                            rfcomm_channel = service.get('port', 1)
                            break

                    if rfcomm_channel is None:
                        rfcomm_channel = 1  # Default RFCOMM channel for OBD

                    # Create RFCOMM socket
                    sock = bluetooth.BluetoothSocket(bluetooth.RFCOMM)
                    sock.settimeout(timeout)
                    sock.connect((bt_address, rfcomm_channel))

                    # Send basic ELM327 test command
                    sock.send("ATZ\r")
                    response = sock.recv(1024)

                    sock.close()

                    # If we get any response, consider it working
                    return len(response) > 0

                except Exception as e:
                    print(f"PyBluez connection test failed: {e}")
                    # Fall back to basic socket approach

            # Fall back to basic socket approach
            try:
                import socket
                sock = socket.socket(socket.AF_BLUETOOTH,
                                     socket.SOCK_STREAM, socket.BTPROTO_RFCOMM)
                sock.settimeout(timeout)

                # Try common RFCOMM channels (1 is most common for OBD)
                for channel in [1, 2, 3]:
                    try:
                        sock.connect((bt_address, channel))

                        # Send basic ELM327 test command
                        sock.send(b"ATZ\r")
                        response = sock.recv(1024)

                        sock.close()

                        # If we get any response, consider it working
                        if response:
                            return True

                    except Exception:
                        continue

                try:
                    sock.close()
                except:
                    pass

            except Exception as e:
                print(f"Socket Bluetooth test failed: {e}")

        return False

    except Exception as e:
        print(f"Bluetooth test error: {e}")
        return False


def create_rfcomm_port_string(bt_address, channel=1):
    """
    Create a port string for Bluetooth RFCOMM connection.
    This will be used with python-OBD library.
    """
    if platform.system() == "Windows":
        # For Windows, we need to find the COM port assigned to the Bluetooth device
        # or use a format that python-OBD can understand
        return f"BT:{bt_address}:{channel}"
    else:
        # For Linux, RFCOMM devices appear as /dev/rfcomm*
        return f"/dev/rfcomm{channel}"


def force_elm327_protocol_serial(port: str, protocol_number: int, baudrate: int = 38400, timeout: float = 1.0) -> bool:
    """
    Open serial port, send ELM327 AT commands to set protocol, then close.
    protocol_number: integer as per ELM327 (1 = J1850 PWM, 2 = J1850 VPW, 0 = Auto)
    Returns True if adapter replied OK to ATSP command.
    """
    try:
        ser = serial.Serial(port, baudrate=baudrate, timeout=timeout)
    except Exception as e:
        print("Serial open error:", e)
        return False

    def send(cmd):
        ser.write((cmd + "\r").encode('ascii'))
        # read lines until we get a response or timeout
        resp_lines = []
        t0 = time.time()
        while True:
            if ser.in_waiting:
                line = ser.readline().decode('ascii', errors='ignore').strip()
                if line:
                    resp_lines.append(line)
                    # ELM327 usually responds with "OK" or "ERROR"
                    if line.upper() in ("OK", "ERROR"):
                        break
            if time.time() - t0 > timeout:
                break
        return resp_lines

    # reset echo off and linefeed off for clearer responses
    send("ATE0")    # echo off
    send("ATL0")    # linefeeds off
    send("ATS0")    # spaces off
    # set protocol (do not save unless user wants to; to save use ATSPn without preceding ATSP0)
    resp = send(f"ATSP{protocol_number}")
    ser.close()
    # check for OK in response lines
    return any("OK" in r.upper() for r in resp)

# ----------------------------
# Polling worker (uses python-OBD)
# ----------------------------


class PollWorker(threading.Thread):
    def __init__(self, port=None, tcp_host=None, tcp_port=35000, protocol_force=None,
                 poll_commands=None, rate_hz=10, out_q=None, stop_event=None, smooth_win=3):
        super().__init__(daemon=True)
        self.port = port
        self.tcp_host = tcp_host
        self.tcp_port = tcp_port
        self.protocol_force = protocol_force  # integer or None
        self.poll_commands = poll_commands or [
            obd.commands.RPM,  # type: ignore
            obd.commands.MAF,  # type: ignore
            # type: ignore  # MAP (kPa on many vehicles)
            obd.commands.INTAKE_PRESSURE,  # type: ignore
            obd.commands.INTAKE_TEMP,      # type: ignore  # IAT in deg C
            obd.commands.THROTTLE_POS,     # type: ignore
            obd.commands.SPEED,            # type: ignore
            obd.commands.COOLANT_TEMP      # type: ignore
        ]
        self.rate_hz = max(1, rate_hz)
        self.out_q = out_q or queue.Queue()
        self.stop_event = stop_event or threading.Event()
        self.smooth_win = max(1, smooth_win)
        self.connection = None
        self.buffers = {cmd.name: deque(maxlen=self.smooth_win)
                        for cmd in self.poll_commands}
        self.connected = False

    def _connect_elm(self):
        # If tcp_host provided, use socket-like port string for python-OBD
        try:
            if self.protocol_force is not None and self.port:
                # Force protocol by sending ATSPn via serial first (ELM327)
                ok = force_elm327_protocol_serial(
                    self.port, self.protocol_force)
                # proceed to python-OBD connect
            if self.tcp_host:
                portstr = f"socket://{self.tcp_host}:{self.tcp_port}"
                self.connection = obd.OBD(portstr, fast=False)
            elif self.port:
                self.connection = obd.OBD(self.port, fast=False)
            else:
                self.connection = obd.OBD(fast=False)  # auto
            self.connected = self.connection is not None and self.connection.status(
            ) == OBDStatus.CAR_CONNECTED
            return self.connected
        except Exception as e:
            print("OBD connect exception:", e)
            self.connected = False
            return False

    def run(self):
        # Attempt connect
        if not self._connect_elm():
            # try reconnect loop until stopped
            while not self.stop_event.is_set():
                time.sleep(1.0)
                if self._connect_elm():
                    break

        poll_interval = 1.0 / self.rate_hz
        last_time = time.time()

        while not self.stop_event.is_set():
            if not self.connected:
                time.sleep(0.5)
                continue

            datapoint = {"timestamp": int(time.time()*1000)}
            for cmd in self.poll_commands:
                try:
                    if self.connection is not None:
                        resp = self.connection.query(
                            cmd, force=True)  # type: ignore
                    else:
                        continue
                    v = None
                    if resp is not None and resp.value is not None:
                        # safe extraction
                        try:
                            if hasattr(resp.value, 'magnitude'):
                                v = float(resp.value.magnitude)
                            else:
                                v = float(resp.value)
                        except Exception:
                            v = None
                except Exception:
                    v = None

                # smoothing & buffer
                buf = self.buffers.setdefault(
                    cmd.name, deque(maxlen=self.smooth_win))
                if v is not None:
                    buf.append(v)
                    sm = sum(buf)/len(buf)
                else:
                    sm = None
                # store with friendly key names lowercased (rpm, maf, intake_pressure->map, intake_temp->iat)
                key = cmd.name.lower()
                # normalize common names
                if key == 'intake_pressure':
                    key = 'map'
                if key == 'intake_temp':
                    key = 'iat'
                if key == 'coolant_temp':
                    key = 'clt'
                if key == 'throttle_pos':
                    key = 'tps'
                if key == 'mass_air_flow':
                    key = 'maf'
                datapoint[key] = sm if sm is not None else 0.0  # type: ignore

            # enqueue datapoint
            self.out_q.put(datapoint)

            # sleep to maintain poll rate
            elapsed = time.time() - last_time
            to_sleep = poll_interval - elapsed
            if to_sleep > 0:
                time.sleep(to_sleep)
            last_time = time.time()

        # cleanup
        try:
            if self.connection:
                self.connection.close()
        except Exception:
            pass

# ----------------------------
# VE Table manager
# ----------------------------


class VETable:
    def __init__(self, rpm_bins=None, map_bins=None, cylinders=8):
        # rpm_bins: list of RPM centers (e.g., [400, 800, 1200, ...])
        # map_bins: list of MAP centers in kPa (e.g., [15, 20, 25, ...])
        self.rpm_bins = rpm_bins or list(
            range(400, 8001, 400))  # default 400..8000 step 400
        # default 15..105 kPa step 5
        self.map_bins = map_bins or list(range(15, 106, 5))
        self.cylinders = cylinders
        # store VE values in 2D dict keyed by (map_bin, rpm_bin) -> latest VE (float)
        self.table: Dict[int, Dict[int, Optional[float]]] = {m: {r: None for r in self.rpm_bins}
                                                             for m in self.map_bins}
        # additional: count or timestamp of last update for smoothing/aging
        self.last_updated: Dict[int, Dict[int, Optional[float]]] = {m: {r: None for r in self.rpm_bins}
                                                                    for m in self.map_bins}

    @staticmethod
    def compute_g_per_cyl_from_maf(maf_g_s: float, rpm: float, cylinders: int) -> Optional[float]:
        """
        Compute grams per cylinder per intake event using:
          g_per_cyl = (MAF_g_per_s * 120) / (RPM * cylinders)
        Derivation:
          MAF (g/s) -> g/min = MAF*60
          intake events per minute = (RPM/2) * cylinders
          g per event per cylinder = (MAF*60) / (RPM/2 * cylinders) = (MAF*120)/(RPM*cylinders)
        """
        if maf_g_s is None or rpm is None or rpm <= 0 or cylinders <= 0:
            return None
        return (maf_g_s * 120.0) / (rpm * cylinders)

    @staticmethod
    def compute_ve_from_g_cyl(g_per_cyl: float, charge_temp_k: float, map_kpa: float) -> Optional[float]:
        """
        VE = (g_per_cyl * ChargeTemp_K) / MAP_kPa
        Units:
          g_per_cyl: grams
          ChargeTemp_K: Kelvin
          MAP_kPa: kilopascal
        VE units: g*K/kPa
        """
        if g_per_cyl is None or charge_temp_k is None or map_kpa is None or map_kpa == 0:
            return None
        return (g_per_cyl * charge_temp_k) / map_kpa

    def find_nearest_bins(self, rpm_value, map_value):
        # choose nearest rpm bin and map bin centers
        if rpm_value is None or map_value is None:
            return None, None
        # clamp and find closest
        rpm_bin = min(self.rpm_bins, key=lambda r: abs(r - rpm_value))
        map_bin = min(self.map_bins, key=lambda m: abs(m - map_value))
        return rpm_bin, map_bin

    def update_from_measurement(self, maf_g_s, rpm, iat_c, map_kpa, timestamp=None):
        """
        Compute VE from measurement and update appropriate table cell.
        - maf_g_s: grams/sec (MAF)
        - rpm: RPM (rev/min)
        - iat_c: intake air temp in Celsius (converted to K)
        - map_kpa: MAP in kPa
        """
        if maf_g_s is None or rpm is None or iat_c is None or map_kpa is None:
            return None  # insufficient data

        g_per_cyl = self.compute_g_per_cyl_from_maf(
            maf_g_s, rpm, self.cylinders)
        if g_per_cyl is None:
            return None

        charge_k = iat_c + 273.15
        ve = self.compute_ve_from_g_cyl(g_per_cyl, charge_k, map_kpa)
        if ve is None:
            return None

        rpm_bin, map_bin = self.find_nearest_bins(rpm, map_kpa)
        if rpm_bin is None or map_bin is None:
            return None

        # store (optionally smooth with previous)
        prev = self.table[map_bin][rpm_bin]
        if prev is None:
            new_val = ve
        else:
            # simple exponential smoothing to avoid wild jumps (alpha configurable)
            alpha = 0.4
            new_val = prev * (1 - alpha) + ve * alpha

        self.table[map_bin][rpm_bin] = new_val
        self.last_updated[map_bin][rpm_bin] = timestamp or time.time()
        return (rpm_bin, map_bin, new_val)

# ----------------------------
# Main GUI application (PySide6)
# ----------------------------


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("OBD-II Advanced VE Tuner (J1850 compatible)")
        self.resize(1300, 900)

        # Queue for worker -> GUI
        self.poll_q = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None

        # Data logging for CSV export
        self.data_log = []  # Store all datapoints for CSV export
        self.max_log_entries = 10000  # Limit to prevent memory issues

        # VE table: default bins as previously described
        rpm_bins = list(range(400, 8001, 400))   # 400..8000 step 400
        map_bins = list(range(15, 106, 5))       # 15..105 kPa step 5
        self.ve_table = VETable(
            rpm_bins=rpm_bins, map_bins=map_bins, cylinders=8)

        # GUI layout
        central = QWidget()
        self.setCentralWidget(central)
        vbox = QVBoxLayout(central)

        # Connection controls row
        h_conn = QHBoxLayout()

        # Port selection
        port_layout = QVBoxLayout()
        port_row = QHBoxLayout()
        self.port_combo = QComboBox()
        self.port_combo.setMinimumWidth(200)
        self.refresh_ports_btn = QPushButton("🔄 Scan Ports")
        self.refresh_ports_btn.setMaximumWidth(120)
        self.test_connection_btn = QPushButton("🔧 Test OBD")
        self.test_connection_btn.setMaximumWidth(100)
        port_row.addWidget(QLabel("Port:"))
        port_row.addWidget(self.port_combo)
        port_row.addWidget(self.refresh_ports_btn)
        port_row.addWidget(self.test_connection_btn)
        port_layout.addLayout(port_row)

        # Manual port input (fallback)
        manual_row = QHBoxLayout()
        self.port_input = QLineEdit()
        self.port_input.setPlaceholderText(
            "Manual port (e.g., COM3 or BT:XX:XX:XX:XX:XX:XX)")
        manual_row.addWidget(QLabel("Manual:"))
        manual_row.addWidget(self.port_input)
        port_layout.addLayout(manual_row)

        h_conn.addLayout(port_layout)

        # Other connection settings
        self.protocol_combo = QComboBox()
        self.protocol_combo.addItems(
            ["Auto(0)", "J1850 PWM(1)", "J1850 VPW(2)", "ISO 9141-2(3)", "KWP(4)", "CAN(6)"])
        # default to J1850 VPW (GM) common for older GM
        self.protocol_combo.setCurrentIndex(2)
        self.tcp_input = QLineEdit()
        self.tcp_input.setPlaceholderText(
            "TCP host (optional, e.g., 192.168.0.10)")
        self.poll_hz_spin = QSpinBox()
        self.poll_hz_spin.setRange(1, 50)
        self.poll_hz_spin.setValue(10)
        self.smooth_spin = QSpinBox()
        self.smooth_spin.setRange(1, 20)
        self.smooth_spin.setValue(3)
        self.cyl_spin = QSpinBox()
        self.cyl_spin.setRange(1, 16)
        self.cyl_spin.setValue(self.ve_table.cylinders)

        self.connect_btn = QPushButton("Connect & Start")
        self.disconnect_btn = QPushButton("Stop")
        self.disconnect_btn.setEnabled(False)

        h_conn.addWidget(QLabel("TCP:"))
        h_conn.addWidget(self.tcp_input)
        h_conn.addWidget(QLabel("Protocol:"))
        h_conn.addWidget(self.protocol_combo)
        h_conn.addWidget(QLabel("Poll Hz:"))
        h_conn.addWidget(self.poll_hz_spin)
        h_conn.addWidget(QLabel("Smoothing:"))
        h_conn.addWidget(self.smooth_spin)
        h_conn.addWidget(QLabel("Cylinders:"))
        h_conn.addWidget(self.cyl_spin)
        h_conn.addWidget(self.connect_btn)
        h_conn.addWidget(self.disconnect_btn)
        vbox.addLayout(h_conn)

        # Top plots and live readouts
        top_h = QHBoxLayout()
        # plots
        plot_layout = QVBoxLayout()
        self.plot_widget = pg.GraphicsLayoutWidget()
        self.plot_widget.setBackground('w')
        self.rpm_plot = self.plot_widget.addPlot(title="RPM")
        self.rpm_curve = self.rpm_plot.plot(pen='r')
        self.plot_widget.nextRow()
        self.afr_plot = self.plot_widget.addPlot(title="MAF(g/s) and MAP(kPa)")
        self.maf_curve = self.afr_plot.plot(pen='g')
        self.map_curve = self.afr_plot.plot(pen='b')
        plot_layout.addWidget(self.plot_widget)
        top_h.addLayout(plot_layout, 3)

        # live labels
        live_layout = QVBoxLayout()
        self.rpm_label = QLabel("RPM: —")
        self.maf_label = QLabel("MAF (g/s): —")
        self.map_label = QLabel("MAP (kPa): —")
        self.iat_label = QLabel("IAT (°C): —")
        for w in (self.rpm_label, self.maf_label, self.map_label, self.iat_label):
            w.setStyleSheet(
                "font-size:16px; padding:6px; background:#fff; border:1px solid #ddd;")
            live_layout.addWidget(w)
        top_h.addLayout(live_layout, 1)

        vbox.addLayout(top_h)

        # VE table grid and controls
        mid_h = QHBoxLayout()
        # VE QTableWidget
        self.ve_table_widget = QTableWidget()
        self._init_ve_table_widget()
        mid_h.addWidget(self.ve_table_widget, 3)

        # controls for VE table
        ctrl_layout = QVBoxLayout()
        self.save_btn = QPushButton("Save VE Table (JSON)")
        self.load_btn = QPushButton("Load VE Table (JSON)")
        self.clear_btn = QPushButton("Clear VE Table")

        # CSV Export buttons
        self.export_ve_csv_btn = QPushButton("Export VE Table (CSV)")
        self.export_data_csv_btn = QPushButton("Export Data Log (CSV)")
        self.clear_log_btn = QPushButton("Clear Data Log")

        ctrl_layout.addWidget(self.save_btn)
        ctrl_layout.addWidget(self.load_btn)
        ctrl_layout.addWidget(self.clear_btn)
        ctrl_layout.addWidget(QLabel(""))  # Spacer
        ctrl_layout.addWidget(self.export_ve_csv_btn)
        ctrl_layout.addWidget(self.export_data_csv_btn)
        ctrl_layout.addWidget(self.clear_log_btn)
        ctrl_layout.addStretch()
        mid_h.addLayout(ctrl_layout, 1)

        vbox.addLayout(mid_h)

        # bottom: status log & connection info
        bottom_h = QHBoxLayout()
        self.log_label = QLabel("Status: Idle")
        self.connection_status_label = QLabel("Connection: Not connected")
        self.connection_status_label.setStyleSheet("color: red;")
        self.data_log_status_label = QLabel("Data Log: 0 entries")
        self.data_log_status_label.setStyleSheet("color: blue;")
        bottom_h.addWidget(self.log_label)
        bottom_h.addWidget(self.connection_status_label)
        bottom_h.addWidget(self.data_log_status_label)
        vbox.addLayout(bottom_h)

        # internal buffers for plotting
        self.max_points = 500
        self.buffers = {'rpm': deque(maxlen=self.max_points), 'maf': deque(
            maxlen=self.max_points), 'map': deque(maxlen=self.max_points)}

        # timers and signals
        self.gui_timer = QTimer()
        self.gui_timer.setInterval(100)  # 10 Hz refresh
        self.gui_timer.timeout.connect(self._on_gui_timer)
        self.gui_timer.start()

        # connect signals
        self.connect_btn.clicked.connect(self._on_connect)
        self.disconnect_btn.clicked.connect(self._on_disconnect)
        self.refresh_ports_btn.clicked.connect(self._on_refresh_ports)
        self.test_connection_btn.clicked.connect(self._on_test_connection)
        self.save_btn.clicked.connect(self._on_save_ve)
        self.load_btn.clicked.connect(self._on_load_ve)
        self.clear_btn.clicked.connect(self._on_clear_ve)
        self.export_ve_csv_btn.clicked.connect(self._on_export_ve_csv)
        self.export_data_csv_btn.clicked.connect(self._on_export_data_csv)
        self.clear_log_btn.clicked.connect(self._on_clear_data_log)
        self.cyl_spin.valueChanged.connect(self._on_cylinder_change)

        # Initialize port list
        self._on_refresh_ports()

    def _init_ve_table_widget(self):
        rpm_bins = self.ve_table.rpm_bins
        map_bins = self.ve_table.map_bins
        # columns = RPM bins, rows = MAP bins
        self.ve_table_widget.setColumnCount(len(rpm_bins))
        self.ve_table_widget.setRowCount(len(map_bins))
        self.ve_table_widget.setHorizontalHeaderLabels(
            [str(r) for r in rpm_bins])
        self.ve_table_widget.setVerticalHeaderLabels(
            [str(m) for m in map_bins])
        self.ve_table_widget.horizontalHeader().setSectionResizeMode(  # type: ignore
            QHeaderView.Stretch)
        self.ve_table_widget.verticalHeader().setSectionResizeMode(  # type: ignore
            QHeaderView.Stretch)
        # initialize with empty items
        for r_idx, m in enumerate(map_bins):
            for c_idx, r in enumerate(rpm_bins):
                item = QTableWidgetItem("")
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.ve_table_widget.setItem(r_idx, c_idx, item)

    @Slot()
    def _on_refresh_ports(self):
        """Scan for available serial and Bluetooth ports and populate the combo box."""
        self.port_combo.clear()
        self.port_combo.addItem("Auto-detect", "")

        try:
            # Get both serial and Bluetooth ports
            ports = scan_all_ports()
            if not ports:
                self.port_combo.addItem("No ports found", "")
                self._append_log("No serial or Bluetooth ports found")
                return

            # Categorize ports
            serial_obd_ports = []
            bluetooth_obd_ports = []
            other_serial_ports = []
            other_bluetooth_ports = []

            for port_info in ports:
                device = port_info['device']
                description = port_info['description']
                port_type = port_info.get('type', 'unknown')

                # Check if this looks like an OBD adapter
                desc_lower = description.lower()
                is_obd_likely = any(keyword in desc_lower for keyword in
                                    ['elm327', 'obd', 'obdii', 'usb-serial', 'ch340', 'cp210', 'ftdi']) or \
                    'obd likely' in desc_lower

                display_text = f"{device} - {description}"

                # Add connection status for Bluetooth devices
                if port_type == 'bluetooth':
                    connected_status = " (Connected)" if port_info.get(
                        'connected', False) else " (Not Connected)"

                    # Add source info for debugging/transparency
                    source = port_info.get('source', 'unknown')
                    source_map = {
                        'discoverable': 'Disc',
                        'paired': 'Paired',
                        'powershell': 'PS'
                    }
                    source_text = source_map.get(source, source)
                    display_text += f"{connected_status} [{source_text}]"

                # Categorize by type and OBD likelihood
                if port_type == 'bluetooth':
                    if is_obd_likely:
                        bluetooth_obd_ports.append((display_text, device))
                    else:
                        other_bluetooth_ports.append((display_text, device))
                else:  # serial
                    if is_obd_likely:
                        display_text += " (OBD likely)"
                        serial_obd_ports.append((display_text, device))
                    else:
                        other_serial_ports.append((display_text, device))

            # Add ports in priority order: BT OBD, Serial OBD, Other BT, Other Serial
            all_categorized = [
                ("🔵 Bluetooth OBD Devices:", bluetooth_obd_ports),
                ("🔌 Serial OBD Devices:", serial_obd_ports),
                ("🔵 Other Bluetooth:", other_bluetooth_ports),
                ("🔌 Other Serial:", other_serial_ports)
            ]

            for category_name, category_ports in all_categorized:
                if category_ports:
                    # Add category separator (disabled item)
                    self.port_combo.addItem(f"--- {category_name} ---", "")

                    for display_text, device in category_ports:
                        self.port_combo.addItem(display_text, device)

            serial_count = len([p for p in ports if p.get('type') == 'serial'])
            bluetooth_count = len(
                [p for p in ports if p.get('type') == 'bluetooth'])
            self._append_log(
                f"Found {serial_count} serial ports, {bluetooth_count} Bluetooth devices")

        except Exception as e:
            self._append_log(f"Error scanning ports: {e}")
            self.port_combo.addItem("Error scanning ports", "")

    @Slot()
    def _on_test_connection(self):
        """Test OBD connection on selected or all available ports."""
        selected_port = self.port_combo.currentData()
        manual_port = self.port_input.text().strip()

        if manual_port:
            port_to_test = manual_port
        # Skip separator items
        elif selected_port and not selected_port.startswith("---"):
            port_to_test = selected_port
        else:
            # Test all available ports (both serial and Bluetooth)
            ports = scan_all_ports()
            if not ports:
                QMessageBox.information(
                    self, "Test Results", "No ports available to test.")
                return

            working_ports = []
            self._append_log("Testing all available ports for OBD adapters...")

            for port_info in ports:
                port_name = port_info['device']
                port_type = port_info.get('type', 'unknown')

                self._append_log(f"Testing {port_type} port {port_name}...")

                try:
                    if test_obd_connection(port_name):
                        working_ports.append(
                            f"{port_name} - {port_info['description']} ({port_type})")
                        self._append_log(
                            f"✓ {port_name} has working OBD adapter")
                    else:
                        self._append_log(f"✗ {port_name} no OBD response")
                except Exception as e:
                    self._append_log(f"✗ {port_name} test error: {e}")

            if working_ports:
                msg = "Working OBD ports found:\n\n" + "\n".join(working_ports)
                QMessageBox.information(self, "OBD Test Results", msg)
            else:
                QMessageBox.warning(self, "OBD Test Results",
                                    "No working OBD adapters found.")
            return

        # Test single port
        port_type = "Bluetooth" if port_to_test.startswith('BT:') else "Serial"
        self._append_log(
            f"Testing {port_type} OBD connection on {port_to_test}...")

        try:
            if test_obd_connection(port_to_test):
                QMessageBox.information(self, "OBD Test Results",
                                        f"✓ OBD adapter working on {port_to_test} ({port_type})")
                self._append_log(f"✓ OBD adapter working on {port_to_test}")
            else:
                QMessageBox.warning(self, "OBD Test Results",
                                    f"✗ No OBD response on {port_to_test} ({port_type})")
                self._append_log(f"✗ No OBD response on {port_to_test}")
        except Exception as e:
            QMessageBox.warning(self, "OBD Test Results",
                                f"✗ Test error on {port_to_test}: {e}")
            self._append_log(f"✗ Test error on {port_to_test}: {e}")

    @Slot()
    def _on_cylinder_change(self):
        self.ve_table.cylinders = int(self.cyl_spin.value())

    @Slot()
    def _on_connect(self):
        # Get selected port from combo or manual input
        selected_port = self.port_combo.currentData()
        manual_port = self.port_input.text().strip()

        if manual_port:
            port = manual_port
        elif selected_port:
            port = selected_port
        else:
            port = None

        tcp = self.tcp_input.text().strip() or None
        poll_hz = int(self.poll_hz_spin.value())
        smooth = int(self.smooth_spin.value())
        protocol_idx = self.protocol_combo.currentIndex()
        # map combo index to ELM327 protocol number: as per ELM327 docs -> Auto=0, J1850 PWM=1, J1850 VPW=2, ISO=3, KWP=4, CAN=6
        protocol_map = {0: 0, 1: 1, 2: 2, 3: 3, 4: 4, 5: 6}
        protocol_number = protocol_map.get(protocol_idx, 0)

        # Validate connection settings
        if not port and not tcp:
            QMessageBox.warning(self, "Connection Error",
                                "Please select a port or enter TCP host.")
            return

        # create and start worker
        self.stop_event.clear()
        self.worker = PollWorker(port=port, tcp_host=tcp, protocol_force=(protocol_number if protocol_number != 0 else None),
                                 poll_commands=None, rate_hz=poll_hz, out_q=self.poll_q, stop_event=self.stop_event, smooth_win=smooth)
        self.worker.start()
        self.connect_btn.setEnabled(False)
        self.disconnect_btn.setEnabled(True)
        self.refresh_ports_btn.setEnabled(False)
        connection_info = f"Port: {port or 'Auto'}, TCP: {tcp or 'None'}, Protocol: {protocol_number}"
        self.log_label.setText(f"Status: Connecting... ({connection_info})")
        self._append_log("Started poll worker")

    @Slot()
    def _on_disconnect(self):
        if self.worker:
            self.stop_event.set()
            self.worker.join(timeout=2.0)
            self.worker = None
        self.connect_btn.setEnabled(True)
        self.disconnect_btn.setEnabled(False)
        self.refresh_ports_btn.setEnabled(True)
        self.log_label.setText("Status: Stopped")
        self._append_log("Stopped poll worker")

    @Slot()
    def _on_save_ve(self):
        fn, _ = QFileDialog.getSaveFileName(
            self, "Save VE Table", os.path.expanduser("~"), "JSON Files (*.json)")
        if not fn:
            return
        payload = {"rpm_bins": self.ve_table.rpm_bins, "map_bins": self.ve_table.map_bins,
                   "cylinders": self.ve_table.cylinders, "table": self.ve_table.table}
        with open(fn, "w") as f:
            json.dump(payload, f, indent=2)
        self._append_log(f"Saved VE table to {fn}")

    @Slot()
    def _on_load_ve(self):
        fn, _ = QFileDialog.getOpenFileName(
            self, "Load VE Table", os.path.expanduser("~"), "JSON Files (*.json)")
        if not fn:
            return
        with open(fn, "r") as f:
            payload = json.load(f)
        # basic validation
        if "rpm_bins" in payload and "map_bins" in payload and "table" in payload:
            self.ve_table.rpm_bins = payload["rpm_bins"]
            self.ve_table.map_bins = payload["map_bins"]
            self.ve_table.table = payload["table"]
            self.ve_table.cylinders = payload.get(
                "cylinders", self.ve_table.cylinders)
            self.cyl_spin.setValue(self.ve_table.cylinders)
            # re-init GUI table
            self._init_ve_table_widget()
            # populate GUI from table
            for r_idx, m in enumerate(self.ve_table.map_bins):
                for c_idx, r in enumerate(self.ve_table.rpm_bins):
                    val = self.ve_table.table.get(m, {}).get(r)  # type: ignore
                    if val is None:
                        text = ""
                    else:
                        text = f"{val:.3f}"
                    self.ve_table_widget.setItem(
                        r_idx, c_idx, QTableWidgetItem(text))
            self._append_log(f"Loaded VE table from {fn}")
        else:
            QMessageBox.warning(
                self, "Invalid file", "Selected JSON does not contain a valid VE table structure")

    @Slot()
    def _on_clear_ve(self):
        # clear internal and GUI table
        for m in self.ve_table.map_bins:
            for r in self.ve_table.rpm_bins:
                self.ve_table.table[m][r] = None
                self.ve_table.last_updated[m][r] = None
        self._init_ve_table_widget()
        self._append_log("Cleared VE table")

    @Slot()
    def _on_export_ve_csv(self):
        """Export VE table to CSV format."""
        fn, _ = QFileDialog.getSaveFileName(
            self, "Export VE Table to CSV", os.path.expanduser("~"), "CSV Files (*.csv)")
        if not fn:
            return

        try:
            with open(fn, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile)

                # Write header row with RPM bins
                header = ['MAP\\RPM'] + [str(rpm)
                                         for rpm in self.ve_table.rpm_bins]
                writer.writerow(header)

                # Write VE data rows
                for map_bin in self.ve_table.map_bins:
                    row = [str(map_bin)]
                    for rpm_bin in self.ve_table.rpm_bins:
                        ve_value = self.ve_table.table.get(
                            map_bin, {}).get(rpm_bin)
                        if ve_value is not None:
                            row.append(f"{ve_value:.6f}")
                        else:
                            row.append("")
                    writer.writerow(row)

            self._append_log(f"Exported VE table to CSV: {fn}")

        except Exception as e:
            QMessageBox.critical(self, "Export Error",
                                 f"Failed to export VE table:\n{e}")
            self._append_log(f"Error exporting VE table: {e}")

    @Slot()
    def _on_export_data_csv(self):
        """Export logged data to CSV format."""
        if not self.data_log:
            QMessageBox.information(
                self, "No Data", "No data available to export. Start logging first.")
            return

        fn, _ = QFileDialog.getSaveFileName(
            self, "Export Data Log to CSV", os.path.expanduser("~"), "CSV Files (*.csv)")
        if not fn:
            return

        try:
            with open(fn, 'w', newline='') as csvfile:
                if not self.data_log:
                    return

                # Get all unique keys from all datapoints
                all_keys = set()
                for dp in self.data_log:
                    all_keys.update(dp.keys())

                # Sort keys for consistent column order
                sorted_keys = sorted(all_keys)

                writer = csv.DictWriter(csvfile, fieldnames=sorted_keys)
                writer.writeheader()

                # Write all data rows
                for dp in self.data_log:
                    # Fill missing values with empty strings
                    row = {key: dp.get(key, '') for key in sorted_keys}
                    writer.writerow(row)

            self._append_log(
                f"Exported {len(self.data_log)} data points to CSV: {fn}")

        except Exception as e:
            QMessageBox.critical(self, "Export Error",
                                 f"Failed to export data log:\n{e}")
            self._append_log(f"Error exporting data log: {e}")

    @Slot()
    def _on_clear_data_log(self):
        """Clear the data log."""
        self.data_log.clear()
        self._append_log(f"Cleared data log")
        QMessageBox.information(self, "Data Log Cleared",
                                "Data log has been cleared.")

    def _append_log(self, s):
        t = time.strftime("%H:%M:%S")
        self.log_label.setText(f"Status: {s} ({t})")

    def _on_gui_timer(self):
        # Check worker connection status
        if self.worker:
            if self.worker.connected:
                self.connection_status_label.setText("Connection: Connected ✓")
                self.connection_status_label.setStyleSheet("color: green;")
            else:
                self.connection_status_label.setText(
                    "Connection: Trying to connect...")
                self.connection_status_label.setStyleSheet("color: orange;")
        else:
            self.connection_status_label.setText("Connection: Not connected")
            self.connection_status_label.setStyleSheet("color: red;")

        # Update data log status
        log_count = len(self.data_log)
        self.data_log_status_label.setText(f"Data Log: {log_count} entries")
        if log_count > self.max_log_entries * 0.9:  # Warn when approaching limit
            self.data_log_status_label.setStyleSheet("color: orange;")
        else:
            self.data_log_status_label.setStyleSheet("color: blue;")

        # drain poll queue
        updated = False
        last_dp = None
        while not self.poll_q.empty():
            try:
                dp = self.poll_q.get_nowait()
            except queue.Empty:
                break
            last_dp = dp
            updated = True

            # Add timestamp in readable format for CSV export
            dp_for_log = dp.copy()
            timestamp_ms = dp.get('timestamp', int(time.time() * 1000))
            dp_for_log['datetime'] = time.strftime(
                '%Y-%m-%d %H:%M:%S', time.localtime(timestamp_ms / 1000))

            # Store datapoint for CSV export (with size limit)
            self.data_log.append(dp_for_log)
            if len(self.data_log) > self.max_log_entries:
                self.data_log.pop(0)  # Remove oldest entry

            # update plotting buffers
            if dp.get('rpm') is not None:
                self.buffers['rpm'].append(dp.get('rpm'))
            else:
                self.buffers['rpm'].append(0)
            if dp.get('maf') is not None:
                self.buffers['maf'].append(dp.get('maf'))
            else:
                self.buffers['maf'].append(0)
            if dp.get('map') is not None:
                self.buffers['map'].append(dp.get('map'))
            else:
                self.buffers['map'].append(0)

            # update live labels
            self.rpm_label.setText(f"RPM: {dp.get('rpm', '—')}")
            self.maf_label.setText(f"MAF (g/s): {dp.get('maf', '—')}")
            self.map_label.setText(f"MAP (kPa): {dp.get('map', '—')}")
            self.iat_label.setText(f"IAT (°C): {dp.get('iat', '—')}")

            # if we have maf, rpm, iat, map -> compute VE and update table
            maf = dp.get('maf')
            rpm = dp.get('rpm')
            iat = dp.get('iat')
            map_kpa = dp.get('map')
            timestamp = dp.get('timestamp', time.time())
            result = self.ve_table.update_from_measurement(
                maf, rpm, iat, map_kpa, timestamp=timestamp)
            if result:
                rpm_bin, map_bin, ve_val = result
                # update GUI cell
                try:
                    r_idx = self.ve_table.map_bins.index(map_bin)
                    c_idx = self.ve_table.rpm_bins.index(rpm_bin)
                    item = QTableWidgetItem(f"{ve_val:.4f}")
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.ve_table_widget.setItem(r_idx, c_idx, item)
                    # highlight updated cell
                    for rr in range(self.ve_table_widget.rowCount()):
                        for cc in range(self.ve_table_widget.columnCount()):
                            it = self.ve_table_widget.item(rr, cc)
                            if it:
                                it.setBackground(QColor("white"))
                    item.setBackground(QColor("yellow"))
                except Exception:
                    pass

        if updated and last_dp is not None:
            # update plots (simple indexing x axis)
            xs = list(range(len(self.buffers['rpm'])))
            try:
                self.rpm_curve.setData(xs, list(self.buffers['rpm']))
                self.maf_curve.setData(xs, list(self.buffers['maf']))
                self.map_curve.setData(xs, list(self.buffers['map']))
            except Exception:
                pass

    def closeEvent(self, event):
        self._on_disconnect()
        event.accept()


def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
