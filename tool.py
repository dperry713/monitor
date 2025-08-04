# Diagnostic: Print all available OBD command attributes to help avoid attribute errors
# Fix for Python 3.12 compatibility with pint library
import collections
import collections.abc
try:
    collections.MutableMapping = collections.abc.MutableMapping
    collections.Mapping = collections.abc.Mapping
    collections.MutableSet = collections.abc.MutableSet
    collections.Set = collections.abc.Set
    collections.MutableSequence = collections.abc.MutableSequence
    collections.Sequence = collections.abc.Sequence
    collections.Iterable = collections.abc.Iterable
    collections.Iterator = collections.abc.Iterator
    collections.Callable = collections.abc.Callable
except AttributeError:
    pass  # Already exists in older Python versions

import warnings
import time
import random
import numpy as np
import pandas as pd
import serial.tools.list_ports
import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QComboBox, QTableWidget, QTableWidgetItem,
    QFrame, QGridLayout, QGroupBox, QProgressBar, QSplitter, QLineEdit, QMessageBox
)
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QFont, QPalette, QColor
from pyqtgraph import PlotWidget
import pyqtgraph as pg
import obd


def print_available_obd_commands():
    print("Available OBD commands:")
    for cmd in dir(obd.commands):
        # Only print public attributes (skip __dunder__ and private)
        if not cmd.startswith("_"):
            print(cmd)


# Call the diagnostic function at startup
print_available_obd_commands()

# Suppress deprecation warnings from libraries
warnings.filterwarnings('ignore', category=UserWarning)
warnings.filterwarnings('ignore', message='pkg_resources is deprecated')

# Only use the real OBD library
OBD_AVAILABLE = True

# Modern color scheme and styling
MODERN_STYLE = {
    'background': '#2b2b2b',
    'surface': '#3c3c3c',
    'primary': '#0078d4',
    'secondary': '#00bcf2',
    'accent': '#00d7ff',
    'success': '#107c10',
    'warning': '#ff8c00',
    'error': '#d13438',
    'text': '#ffffff',
    'text_secondary': '#cccccc',
    'border': '#5a5a5a'
}

# Professional UI stylesheet
STYLESHEET = f"""
QMainWindow {{
    background-color: {MODERN_STYLE['background']};
    color: {MODERN_STYLE['text']};
}}

QTabWidget::pane {{
    border: 1px solid {MODERN_STYLE['border']};
    background-color: {MODERN_STYLE['surface']};
    border-radius: 8px;
}}

QTabWidget::tab-bar {{
    alignment: center;
}}

QTabBar::tab {{
    background-color: {MODERN_STYLE['surface']};
    color: {MODERN_STYLE['text_secondary']};
    padding: 12px 24px;
    margin: 2px;
    border-radius: 6px;
    min-width: 120px;
    font-weight: 500;
}}

QTabBar::tab:selected {{
    background-color: {MODERN_STYLE['primary']};
    color: {MODERN_STYLE['text']};
    font-weight: 600;
}}

QTabBar::tab:hover {{
    background-color: {MODERN_STYLE['secondary']};
    color: {MODERN_STYLE['text']};
}}

QPushButton {{
    background-color: {MODERN_STYLE['primary']};
    border: none;
    color: {MODERN_STYLE['text']};
    padding: 12px 24px;
    border-radius: 6px;
    font-weight: 600;
    font-size: 14px;
    min-height: 20px;
}}

QPushButton:hover {{
    background-color: {MODERN_STYLE['secondary']};
}}

QPushButton:pressed {{
    background-color: {MODERN_STYLE['accent']};
}}

QPushButton:disabled {{
    background-color: {MODERN_STYLE['border']};
    color: {MODERN_STYLE['text_secondary']};
}}

QComboBox {{
    background-color: {MODERN_STYLE['surface']};
    border: 2px solid {MODERN_STYLE['border']};
    border-radius: 6px;
    padding: 8px 12px;
    color: {MODERN_STYLE['text']};
    font-size: 14px;
    min-height: 20px;
}}

QComboBox:hover {{
    border-color: {MODERN_STYLE['primary']};
}}

QComboBox::drop-down {{
    border: none;
    width: 30px;
}}

QComboBox::down-arrow {{
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 5px solid {MODERN_STYLE['text']};
    margin-right: 10px;
}}

QLabel {{
    color: {MODERN_STYLE['text']};
    font-size: 14px;
    padding: 4px;
}}

QGroupBox {{
    font-weight: 600;
    font-size: 16px;
    color: {MODERN_STYLE['text']};
    border: 2px solid {MODERN_STYLE['border']};
    border-radius: 8px;
    margin: 10px 0;
    padding-top: 20px;
}}

QGroupBox::title {{
    subcontrol-origin: margin;
    left: 15px;
    padding: 0 10px;
    color: {MODERN_STYLE['primary']};
}}

QTableWidget {{
    background-color: {MODERN_STYLE['surface']};
    alternate-background-color: {MODERN_STYLE['background']};
    color: {MODERN_STYLE['text']};
    gridline-color: {MODERN_STYLE['border']};
    border: 1px solid {MODERN_STYLE['border']};
    border-radius: 6px;
    font-size: 12px;
}}

QTableWidget::item {{
    padding: 8px;
    border: none;
}}

QTableWidget::item:selected {{
    background-color: {MODERN_STYLE['primary']};
}}

QHeaderView::section {{
    background-color: {MODERN_STYLE['primary']};
    color: {MODERN_STYLE['text']};
    padding: 8px;
    border: none;
    font-weight: 600;
}}

QFrame {{
    background-color: {MODERN_STYLE['surface']};
    border: 1px solid {MODERN_STYLE['border']};
    border-radius: 8px;
}}

QProgressBar {{
    border: 2px solid {MODERN_STYLE['border']};
    border-radius: 6px;
    text-align: center;
    background-color: {MODERN_STYLE['surface']};
    color: {MODERN_STYLE['text']};
    font-weight: 600;
}}

QProgressBar::chunk {{
    background-color: {MODERN_STYLE['success']};
    border-radius: 4px;
}}
"""

# OBD Parameter Definitions
# MAP = Manifold Absolute Pressure = Intake manifold pressure measured in kPa
# This represents the absolute pressure in the intake manifold, which is used
# to calculate engine load and volumetric efficiency

# RPM / MAP Axis setup
# MAP = Manifold Absolute Pressure (intake manifold pressure in kPa)
rpm_axis = np.arange(400, 8200, 400)  # Engine RPM range
# MAP range in kPa (intake manifold pressure)
map_axis = np.arange(15, 110, 5)
# Volumetric Efficiency table (g*K/kPa)
ve_table = np.full((len(map_axis), len(rpm_axis)), 0.8)


class OBDApp(QMainWindow):

    def setup_logging(self):
        self.log_data = []
        # Dynamic headers based on O2 display format
        display_format = getattr(self, 'o2_display_combo', None)
        current_format = display_format.currentText() if display_format else "Lambda (λ)"

        if "Lambda" in current_format:
            o2_suffix = "Lambda"
        elif "Equivalence" in current_format:
            o2_suffix = "Phi"
        else:
            o2_suffix = "Voltage"

        self.log_headers = [
            'Timestamp', 'RPM', 'Speed', 'Coolant Temp', 'MAP', 'IAT', 'Throttle', 'MAF', 'Timing Advance',
            f'O2 B1S1 {o2_suffix}', f'O2 B2S1 {o2_suffix}'
        ]

    def toggle_logging(self):
        self.logging_enabled = not self.logging_enabled
        if self.logging_enabled:
            self.log_button.setText("Stop Logging")
        else:
            self.log_button.setText("Start Logging")

    def export_log(self):
        import csv
        from PyQt5.QtWidgets import QFileDialog
        if not self.log_data:
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Export Log", "obd_log.csv", "CSV Files (*.csv)")
        if not path:
            return
        with open(path, 'w', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(self.log_headers)
            writer.writerows(self.log_data)
        from PyQt5.QtWidgets import QMessageBox
        QMessageBox.information(self, "Export Complete",
                                f"Log exported to {path}")

    def __init__(self):
        super().__init__()
        self.setWindowTitle("OBD-II Professional Monitor & VE Calculator")
        self.setGeometry(100, 100, 1400, 900)

        # Ensure logging_enabled is always set before any method uses it
        self.logging_enabled = False
        self.log_data = []
        self.log_headers = [
            'Timestamp', 'RPM', 'Speed', 'Coolant Temp', 'MAP', 'IAT', 'Throttle', 'MAF', 'Timing Advance', 'O2 B1S1', 'O2 B2S1'
        ]

        # Apply modern styling
        self.setStyleSheet(STYLESHEET)

        # Set modern font
        font = QFont("Segoe UI", 10)
        self.setFont(font)

        # Create main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Add title header
        self.create_header(main_layout)

        # Create tabs with modern styling
        self.tabs = QTabWidget()
        self.tabs.setTabPosition(QTabWidget.North)
        main_layout.addWidget(self.tabs)

        self.connection = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_pids)

        self.build_connection_tab()
        self.build_gauge_tab()
        self.build_ve_tab()
        self.build_visual_tab()

    def create_header(self, layout):
        """Create a professional header with title and status, always sized and centered correctly"""
        header_frame = QFrame()
        header_frame.setMinimumHeight(90)
        header_frame.setMaximumHeight(160)
        header_frame.setStyleSheet(f"""
            QFrame {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #1a365d, 
                    stop:0.5 {MODERN_STYLE['primary']},
                    stop:1 #2a69ac);
                border-radius: 12px;
                margin-bottom: 10px;
                border: 2px solid rgba(255, 255, 255, 0.1);
            }}
        """)

        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(0)

        # Title
        title_label = QLabel("OBD-II Professional Monitor")
        title_label.setWordWrap(True)
        title_label.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        title_label.setStyleSheet("""
            QLabel {
                color: #ffffff;
                font-size: 2.2em;
                font-weight: 700;
                background: transparent;
                border: none;
                padding-top: 18px;
                padding-bottom: 2px;
            }
        """)

        # Subtitle
        subtitle_label = QLabel(
            "Real-time engine diagnostics & volumetric efficiency analysis")
        subtitle_label.setWordWrap(True)
        subtitle_label.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        subtitle_label.setStyleSheet("""
            QLabel {
                color: #e0e6ed;
                font-size: 1.1em;
                font-weight: 400;
                background: transparent;
                border: none;
                padding-bottom: 12px;
            }
        """)

        header_layout.addWidget(title_label, stretch=2)
        header_layout.addWidget(subtitle_label, stretch=1)

        layout.addWidget(header_frame)

    def build_connection_tab(self):
        self.conn_tab = QWidget()
        main_layout = QVBoxLayout(self.conn_tab)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(20)

        # Connection Group
        connection_group = QGroupBox("OBD Device Connection")
        connection_layout = QVBoxLayout(connection_group)
        connection_layout.setSpacing(15)

        # Port selection
        port_container = QWidget()
        port_layout = QHBoxLayout(port_container)
        port_layout.setContentsMargins(0, 0, 0, 0)

        port_label = QLabel("Select Port:")
        port_label.setStyleSheet("font-weight: 600; color: #cccccc;")
        port_layout.addWidget(port_label)

        self.port_select = QComboBox()
        self.port_select.setMinimumWidth(200)
        self.port_select.setEditable(True)  # Allow manual entry
        self.refresh_ports()
        port_layout.addWidget(self.port_select)
        port_layout.addStretch()

        # Manual port entry
        self.manual_port_input = QLineEdit()
        self.manual_port_input.setPlaceholderText(
            "Enter port manually (e.g., COM4)")
        self.manual_port_input.setVisible(False)
        port_layout.addWidget(self.manual_port_input)

        # Refresh button
        refresh_btn = QPushButton("🔄 Refresh Ports")
        refresh_btn.clicked.connect(self.refresh_ports)
        refresh_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {MODERN_STYLE['surface']};
                border: 2px solid {MODERN_STYLE['border']};
                color: {MODERN_STYLE['text']};
            }}
            QPushButton:hover {{
                border-color: {MODERN_STYLE['primary']};
                background-color: {MODERN_STYLE['primary']};
            }}
        """)
        port_layout.addWidget(refresh_btn)

        # Bluetooth scan button
        bt_scan_btn = QPushButton("🔵 Scan Bluetooth")
        bt_scan_btn.clicked.connect(self.scan_bluetooth)
        bt_scan_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {MODERN_STYLE['secondary']};
                border: 2px solid {MODERN_STYLE['border']};
                color: {MODERN_STYLE['text']};
            }}
            QPushButton:hover {{
                border-color: {MODERN_STYLE['accent']};
                background-color: {MODERN_STYLE['accent']};
            }}
        """)
        port_layout.addWidget(bt_scan_btn)

        # Connect port selection change to show/hide manual input
        self.port_select.currentTextChanged.connect(
            self.on_port_selection_changed)

        connection_layout.addWidget(port_container)

        # Connection button and status
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container)
        button_layout.setContentsMargins(0, 0, 0, 0)

        self.connect_btn = QPushButton("🔌 Connect to ELM327")
        self.connect_btn.setMinimumHeight(50)
        self.connect_btn.clicked.connect(self.connect_obd)
        button_layout.addWidget(self.connect_btn)

        # Add disconnect button
        self.disconnect_btn = QPushButton("🔌 Disconnect")
        self.disconnect_btn.setMinimumHeight(50)
        self.disconnect_btn.setEnabled(False)
        self.disconnect_btn.clicked.connect(self.disconnect_obd)
        self.disconnect_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {MODERN_STYLE['error']};
                color: {MODERN_STYLE['text']};
            }}
            QPushButton:hover {{
                background-color: #ff4757;
            }}
            QPushButton:disabled {{
                background-color: {MODERN_STYLE['border']};
                color: {MODERN_STYLE['text_secondary']};
            }}
        """)
        button_layout.addWidget(self.disconnect_btn)

        connection_layout.addWidget(button_container)

        # Status section
        status_group = QGroupBox("Connection Status")
        status_layout = QVBoxLayout(status_group)

        self.status_label = QLabel("Not Connected")
        self.status_label.setStyleSheet(f"""
            QLabel {{
                color: {MODERN_STYLE['error']};
                font-size: 16px;
                font-weight: 600;
                padding: 10px;
                background-color: rgba(209, 52, 56, 0.1);
                border: 1px solid {MODERN_STYLE['error']};
                border-radius: 6px;
            }}
        """)
        status_layout.addWidget(self.status_label)

        # Connection info
        self.info_label = QLabel(
            "Select a port and click connect to start monitoring")
        self.info_label.setStyleSheet(
            "color: #cccccc; font-style: italic; padding: 10px;")
        status_layout.addWidget(self.info_label)

        # Bluetooth instructions
        bt_info_group = QGroupBox("Bluetooth Connection Help")
        bt_info_layout = QVBoxLayout(bt_info_group)

        bt_instructions = QLabel("""
<b>📱 Bluetooth OBD Connection Guide:</b><br/>
<b>1. Pairing (First Time):</b><br/>
   • Go to Windows Settings → Devices → Bluetooth<br/>
   • Make sure Bluetooth is ON<br/>
   • Put your OBD adapter in pairing mode (usually automatic)<br/>
   • Click "Add Bluetooth or other device" → Bluetooth<br/>
   • Select your OBD device (usually ELM327, OBD2, etc.)<br/>
   • Enter PIN: 1234 or 0000 (most common)<br/><br/>

<b>2. Before Connecting:</b><br/>
   • Plug OBD adapter into your vehicle's diagnostic port<br/>
   • Turn vehicle ignition to ON position (engine can be off)<br/>
   • Wait 10-15 seconds for adapter to initialize<br/>
   • Some adapters require engine to be running<br/><br/>

<b>3. Troubleshooting:</b><br/>
   • Use "Scan Bluetooth" to find paired devices automatically<br/>
   • Check Windows Device Manager for COM port assignment<br/>
   • Try "Refresh Ports" if device doesn't appear<br/>
   • Connection may take 10-15 seconds for Bluetooth<br/>
   • If connection fails, try turning Bluetooth off/on in Windows
        """)
        bt_instructions.setWordWrap(True)
        bt_instructions.setStyleSheet("""
            QLabel {
                color: #cccccc;
                background-color: rgba(120, 120, 120, 0.1);
                border-radius: 6px;
                padding: 15px;
                margin: 5px;
                line-height: 1.4;
            }
        """)
        bt_info_layout.addWidget(bt_instructions)

        main_layout.addWidget(connection_group)
        main_layout.addWidget(status_group)
        main_layout.addWidget(bt_info_group)
        main_layout.addStretch()

        self.tabs.addTab(self.conn_tab, "🔌 Connection")

    def build_gauge_tab(self):
        self.gauge_tab = QWidget()
        self.gauge_tab.setMinimumSize(1200, 600)
        main_layout = QVBoxLayout(self.gauge_tab)
        main_layout.setContentsMargins(40, 40, 40, 40)

        # Simple grid for digital values
        grid_layout = QGridLayout()
        grid_layout.setSpacing(18)
        grid_layout.setContentsMargins(0, 0, 0, 0)

        self.labels = {}
        pid_list = [
            ("RPM", obd.commands.RPM, "#ff6b6b"),
            ("Speed", obd.commands.SPEED, "#4ecdc4"),
            ("Coolant Temp", obd.commands.COOLANT_TEMP, "#45b7d1"),
            ("MAP (Intake Manifold Pressure)",
             obd.commands.INTAKE_PRESSURE, "#96ceb4"),
            ("IAT", obd.commands.INTAKE_TEMP, "#feca57"),
            ("Throttle", obd.commands.THROTTLE_POS, "#ff9ff3"),
            ("MAF", obd.commands.MAF, "#54a0ff"),
            ("Timing Advance", obd.commands.TIMING_ADVANCE, "#5f27cd"),
            ("O2 B1S1", obd.commands.O2_B1S1, "#f7b731"),
            ("O2 B2S1", obd.commands.O2_B2S1, "#8854d0"),
        ]
        self.pid_list = pid_list

        font_label = QFont("Segoe UI", 22, QFont.Bold)
        font_value = QFont("Consolas", 38, QFont.Bold)

        for i, (label, cmd, color) in enumerate(pid_list):
            row = i // 3
            col = (i % 3) * 2
            label_widget = QLabel(label)
            label_widget.setFont(font_label)
            label_widget.setStyleSheet(f"color: #cccccc; padding-right: 18px;")
            value_widget = QLabel("---")
            value_widget.setFont(font_value)
            value_widget.setStyleSheet(f"color: {color}; padding-left: 8px;")
            value_widget.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            self.labels[cmd] = value_widget
            grid_layout.addWidget(label_widget, row, col,
                                  alignment=Qt.AlignRight)
            grid_layout.addWidget(value_widget, row, col +
                                  1, alignment=Qt.AlignLeft)

        main_layout.addLayout(grid_layout)

        # Add O2 sensor display selection
        o2_control_row = QHBoxLayout()
        o2_label = QLabel("O2 Sensor Display:")
        o2_label.setStyleSheet(
            "font-weight: 600; color: #cccccc; font-size: 14px;")
        o2_control_row.addWidget(o2_label)

        self.o2_display_combo = QComboBox()
        self.o2_display_combo.addItems(
            ["Lambda (λ)", "Equivalence Ratio (φ)", "Voltage (V)"])
        self.o2_display_combo.setCurrentText("Lambda (λ)")
        self.o2_display_combo.setMinimumWidth(200)
        self.o2_display_combo.currentTextChanged.connect(
            self.on_o2_display_changed)
        o2_control_row.addWidget(self.o2_display_combo)
        o2_control_row.addStretch()
        main_layout.addLayout(o2_control_row)

        # Add Start/Stop Logging and Export Log buttons at the bottom of the gauge tab
        btn_row = QHBoxLayout()
        self.log_button = QPushButton("Start Logging")
        self.log_button.setMinimumHeight(40)
        self.log_button.clicked.connect(self.toggle_logging)
        btn_row.addWidget(self.log_button)
        export_btn = QPushButton("Export Log to CSV")
        export_btn.setMinimumHeight(40)
        export_btn.clicked.connect(self.export_log)
        btn_row.addWidget(export_btn)
        main_layout.addLayout(btn_row)
        main_layout.addStretch()
        self.tabs.addTab(self.gauge_tab, "📊 Real-Time Gauges")

    def open_customize_dashboard(self):
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QCheckBox, QComboBox, QLabel, QPushButton, QGroupBox, QFormLayout
        dialog = QDialog(self)
        dialog.setWindowTitle("Customize Dashboard")
        dialog.setMinimumWidth(600)
        layout = QVBoxLayout(dialog)
        # ...existing code...

        # Gauge size selection
        size_group = QGroupBox("Gauge Card Size")
        size_layout = QHBoxLayout(size_group)
        size_combo = QComboBox()
        size_combo.addItems(["Small", "Medium", "Large"])
        size_combo.setCurrentText(self.gauge_card_size)
        size_layout.addWidget(QLabel("Size:"))
        size_layout.addWidget(size_combo)
        layout.addWidget(size_group)

        # Gauge visibility and type
        vis_group = QGroupBox("Gauges")
        vis_layout = QFormLayout(vis_group)
        checkboxes = {}
        type_combos = {}
        for label, cmd, icon, color in self.pid_list:
            cb = QCheckBox(f"{icon} {label}")
            cb.setChecked(self.gauge_card_visible[cmd])
            checkboxes[cmd] = cb
            type_combo = QComboBox()
            type_combo.addItems(["Numeric", "Progress Bar"])
            type_combo.setCurrentText(self.gauge_card_type.get(cmd, "Numeric"))
            type_combos[cmd] = type_combo
            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.addWidget(cb)
            row_layout.addWidget(QLabel("Type:"))
            row_layout.addWidget(type_combo)
            vis_layout.addRow(row_widget)
        layout.addWidget(vis_group)

        # Save and Cancel buttons
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        def save_customization():
            self.gauge_card_size = size_combo.currentText()
            for cmd in checkboxes:
                self.gauge_card_visible[cmd] = checkboxes[cmd].isChecked()
                self.gauge_card_type[cmd] = type_combos[cmd].currentText()
            self.refresh_gauge_cards()
            dialog.accept()

        save_btn.clicked.connect(save_customization)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()

    def refresh_gauge_cards(self):
        # Remove all widgets from grid
        for i in reversed(range(self.gauge_grid_layout.count())):
            widget = self.gauge_grid_layout.itemAt(i).widget()
            if widget:
                self.gauge_grid_layout.removeWidget(widget)
                widget.setParent(None)
        # Determine card size
        size_map = {
            'Small': (220, 90, 18, 28),
            'Medium': (300, 130, 28, 38),
            'Large': (400, 180, 38, 48)
        }
        w, h, icon_size, value_size = size_map.get(
            self.gauge_card_size, (400, 180, 38, 48))
        # Recreate cards
        self.labels.clear()
        self.gauge_cards.clear()
        visible_pids = [(label, cmd, icon, color) for (
            label, cmd, icon, color) in self.pid_list if self.gauge_card_visible[cmd]]
        for i, (label, cmd, icon, color) in enumerate(visible_pids):
            card = self.create_gauge_card(
                label, cmd, icon, color, w, h, icon_size, value_size, self.gauge_card_type[cmd])
            row, col = divmod(i, 3)
            self.gauge_grid_layout.addWidget(card, row, col)
        self.gauge_main_layout.update()

    # No longer needed: create_gauge_card

    def build_ve_tab(self):
        self.ve_tab = QWidget()
        main_layout = QVBoxLayout(self.ve_tab)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(20)

        # VE Table Group
        ve_group = QGroupBox("Volumetric Efficiency Table (g/cyl)")
        ve_layout = QVBoxLayout(ve_group)

        # Table description
        desc_label = QLabel(
            "Real-time calculated air mass per cylinder based on MAP, RPM, and IAT")
        desc_label.setStyleSheet(f"""
            QLabel {{
                color: {MODERN_STYLE['text_secondary']};
                font-style: italic;
                padding: 10px;
                background-color: rgba(120, 120, 120, 0.1);
                border-radius: 6px;
                margin-bottom: 10px;
            }}
        """)
        ve_layout.addWidget(desc_label)

        # Create modern table
        self.ve_table_widget = QTableWidget(len(map_axis), len(rpm_axis))
        self.ve_table_widget.setHorizontalHeaderLabels(
            [f"{int(r)}" for r in rpm_axis])
        self.ve_table_widget.setVerticalHeaderLabels(
            [f"{int(m)}" for m in map_axis])

        # Style the table headers
        self.ve_table_widget.horizontalHeader().setStyleSheet(f"""
            QHeaderView::section {{
                background-color: {MODERN_STYLE['primary']};
                color: white;
                padding: 8px;
                border: 1px solid {MODERN_STYLE['border']};
                font-weight: 600;
                font-size: 12px;
            }}
        """)

        self.ve_table_widget.verticalHeader().setStyleSheet(f"""
            QHeaderView::section {{
                background-color: {MODERN_STYLE['secondary']};
                color: white;
                padding: 8px;
                border: 1px solid {MODERN_STYLE['border']};
                font-weight: 600;
                font-size: 12px;
            }}
        """)

        # Set table properties
        self.ve_table_widget.setAlternatingRowColors(True)
        self.ve_table_widget.setSelectionBehavior(QTableWidget.SelectItems)

        # Initialize table with placeholder values
        self.initialize_ve_table()

        ve_layout.addWidget(self.ve_table_widget)

        # Add control buttons for the VE table
        ve_controls = QHBoxLayout()

        clear_table_btn = QPushButton("Clear VE Table")
        clear_table_btn.clicked.connect(self.clear_ve_table)
        clear_table_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {MODERN_STYLE['warning']};
                color: {MODERN_STYLE['text']};
                font-size: 12px;
                padding: 8px 16px;
            }}
            QPushButton:hover {{
                background-color: {MODERN_STYLE['error']};
            }}
        """)
        ve_controls.addWidget(clear_table_btn)

        export_ve_btn = QPushButton("Export VE Table")
        export_ve_btn.clicked.connect(self.export_ve_table)
        export_ve_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {MODERN_STYLE['secondary']};
                color: {MODERN_STYLE['text']};
                font-size: 12px;
                padding: 8px 16px;
            }}
        """)
        ve_controls.addWidget(export_ve_btn)

        ve_controls.addStretch()
        ve_layout.addLayout(ve_controls)
        main_layout.addWidget(ve_group)

        self.tabs.addTab(self.ve_tab, "📋 VE Table")

    def initialize_ve_table(self):
        """Initialize the VE table with placeholder values"""
        for m_idx in range(len(map_axis)):
            for r_idx in range(len(rpm_axis)):
                item = QTableWidgetItem("---")
                item.setTextAlignment(Qt.AlignCenter)
                # Dark gray for uninitialized
                item.setBackground(QColor(60, 60, 60, 50))
                self.ve_table_widget.setItem(m_idx, r_idx, item)

    def clear_ve_table(self):
        """Clear all VE table data"""
        reply = QMessageBox.question(self, 'Clear VE Table',
                                     'Are you sure you want to clear all VE table data?',
                                     QMessageBox.Yes | QMessageBox.No,
                                     QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.initialize_ve_table()
            print("[VE TABLE] Table cleared")

    def export_ve_table(self):
        """Export VE table data to CSV"""
        from PyQt5.QtWidgets import QFileDialog, QMessageBox
        import csv

        path, _ = QFileDialog.getSaveFileName(
            self, "Export VE Table", "ve_table.csv", "CSV Files (*.csv)")

        if not path:
            return

        try:
            with open(path, 'w', newline='') as f:
                writer = csv.writer(f)

                # Write header row with RPM values
                header = ['MAP/RPM'] + [f"{int(r)}" for r in rpm_axis]
                writer.writerow(header)

                # Write data rows
                for m_idx in range(len(map_axis)):
                    row = [f"{int(map_axis[m_idx])}"]
                    for r_idx in range(len(rpm_axis)):
                        item = self.ve_table_widget.item(m_idx, r_idx)
                        if item and item.text() != "---":
                            row.append(item.text())
                        else:
                            row.append("")
                    writer.writerow(row)

            QMessageBox.information(self, "Export Complete",
                                    f"VE Table exported to {path}")
            print(f"[VE TABLE] Exported to {path}")

        except Exception as e:
            QMessageBox.critical(self, "Export Error",
                                 f"Failed to export: {str(e)}")
            print(f"[VE ERROR] Export failed: {e}")

    def build_visual_tab(self):
        self.visual_tab = QWidget()
        main_layout = QVBoxLayout(self.visual_tab)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(20)

        # Chart Group
        chart_group = QGroupBox("Real-Time Data Visualization")
        chart_layout = QVBoxLayout(chart_group)

        # Create enlarged plot widget with modern styling
        self.plot_widget = PlotWidget()
        self.plot_widget.setMinimumHeight(600)
        self.plot_widget.setMinimumWidth(1200)
        self.plot_widget.setBackground('#2b2b2b')
        self.plot_widget.setLabel('left', 'Value', color='white', size='22pt')
        self.plot_widget.setLabel(
            'bottom', 'Time (samples)', color='white', size='22pt')

        # Style the plot
        self.plot_widget.getAxis('left').setPen(color='white', width=4)
        self.plot_widget.getAxis('bottom').setPen(color='white', width=4)
        self.plot_widget.getAxis('left').setTextPen(color='white')
        self.plot_widget.getAxis('bottom').setTextPen(color='white')

        # Add grid
        self.plot_widget.showGrid(x=True, y=True, alpha=0.3)

        # Add legend with modern styling
        self.plot_widget.addLegend(offset=(30, 30), labelTextSize='20pt')

        # Initialize data storage for multiple parameters
        self.plot_data = {
            'rpm': [],
            'map': [],
            'timing_advance': [],
            'throttle': []
        }

        # Create plot curves with modern colors and thicker lines
        pen_styles = {
            'rpm': pg.mkPen(color='#ff6b6b', width=6),
            'map': pg.mkPen(color='#4ecdc4', width=6),
            'timing_advance': pg.mkPen(color='#45b7d1', width=6),
            'throttle': pg.mkPen(color='#feca57', width=6)
        }

        self.plot_curves = {
            'rpm': self.plot_widget.plot([], [], pen=pen_styles['rpm'], name='RPM'),
            'map': self.plot_widget.plot([], [], pen=pen_styles['map'], name='MAP (kPa)'),
            'timing_advance': self.plot_widget.plot([], [], pen=pen_styles['timing_advance'], name='Timing Advance (°)'),
            'throttle': self.plot_widget.plot([], [], pen=pen_styles['throttle'], name='Throttle (%)')
        }

        chart_layout.addWidget(self.plot_widget)
        main_layout.addWidget(chart_group)

        self.tabs.addTab(self.visual_tab, "📈 Visualizations")

    def on_port_selection_changed(self):
        """Handle port selection changes to show/hide manual input"""
        if self.port_select.currentData() == "MANUAL":
            self.manual_port_input.setVisible(True)
            self.manual_port_input.setFocus()
        else:
            self.manual_port_input.setVisible(False)

    def convert_o2_sensor_value(self, voltage, display_format):
        """Convert O2 sensor voltage to the selected display format"""
        if voltage is None:
            return "---"

        # Standard conversion: 0.45V = stoichiometric (14.7:1 AFR)
        # Lower voltage = rich mixture, higher voltage = lean mixture
        afr = 14.7 * (voltage / 0.45)

        if display_format == "Lambda (λ)":
            # Lambda: λ = 14.7 / AFR
            # λ = 1.0 is stoichiometric, λ < 1.0 is rich, λ > 1.0 is lean
            lambda_ratio = 14.7 / afr
            return f"{lambda_ratio:.3f} λ"
        elif display_format == "Equivalence Ratio (φ)":
            # Equivalence ratio: φ = 1 / λ
            # φ = 1.0 is stoichiometric, φ > 1.0 is rich, φ < 1.0 is lean
            lambda_ratio = 14.7 / afr
            phi_ratio = 1.0 / lambda_ratio
            return f"{phi_ratio:.3f} φ"
        elif display_format == "Voltage (V)":
            # Raw voltage display
            return f"{voltage:.3f} V"
        else:
            return f"{voltage:.3f} V"

    def on_o2_display_changed(self):
        """Handle O2 sensor display format change"""
        # Update log headers to reflect the new format
        current_format = self.o2_display_combo.currentText()

        if "Lambda" in current_format:
            o2_suffix = "Lambda"
        elif "Equivalence" in current_format:
            o2_suffix = "Phi"
        else:
            o2_suffix = "Voltage"

        self.log_headers = [
            'Timestamp', 'RPM', 'Speed', 'Coolant Temp', 'MAP', 'IAT', 'Throttle', 'MAF', 'Timing Advance',
            f'O2 B1S1 {o2_suffix}', f'O2 B2S1 {o2_suffix}'
        ]

        print(
            f"[O2 DISPLAY] Changed to {current_format}, logging headers updated")

    def scan_bluetooth(self):
        """Scan for paired Bluetooth OBD devices and add them to the port list"""
        import platform
        if platform.system() != "Windows":
            self.update_status(
                "Bluetooth scan only supported on Windows", "warning")
            return

        self.update_status(
            "Scanning for paired Bluetooth OBD devices...", "warning")

        try:
            import subprocess
            import re

            # Enhanced PowerShell command to find paired Bluetooth devices with COM ports
            ps_command = """
            # Get paired Bluetooth devices with COM ports
            $devices = @()
            
            # Method 1: Check WMI for Bluetooth COM devices
            Get-WmiObject -Class Win32_PnPEntity | Where-Object {
                $_.Name -match "COM\d+" -and ($_.Name -match "Bluetooth|BT" -or $_.DeviceID -match "BTHENUM")
            } | ForEach-Object {
                if ($_.Name -match "(COM\d+)") {
                    $devices += [PSCustomObject]@{
                        Port = $matches[1]
                        Name = $_.Name
                        Type = "Bluetooth"
                        Status = "Paired"
                    }
                }
            }
            
            # Method 2: Check registry for Bluetooth serial ports
            try {
                $regKey = Get-ItemProperty "HKLM:\HARDWARE\DEVICEMAP\SERIALCOMM" -ErrorAction SilentlyContinue
                if ($regKey) {
                    $regKey.PSObject.Properties | Where-Object {
                        $_.Name -match "BthModem|Bluetooth" -and $_.Value -match "COM\d+"
                    } | ForEach-Object {
                        $comPort = $_.Value
                        if (-not ($devices | Where-Object { $_.Port -eq $comPort })) {
                            $devices += [PSCustomObject]@{
                                Port = $comPort
                                Name = "Bluetooth Serial Port"
                                Type = "Bluetooth"
                                Status = "Paired"
                            }
                        }
                    }
                }
            } catch {}
            
            # Method 3: Check common OBD Bluetooth ports by attempting connection
            1..20 | ForEach-Object {
                $port = "COM$_"
                if (-not ($devices | Where-Object { $_.Port -eq $port })) {
                    try {
                        $serial = New-Object System.IO.Ports.SerialPort($port, 38400)
                        $serial.ReadTimeout = 100
                        $serial.WriteTimeout = 100
                        $serial.Open()
                        Start-Sleep -Milliseconds 50
                        $serial.Close()
                        
                        # If we can open the port, it might be a Bluetooth device
                        $devices += [PSCustomObject]@{
                            Port = $port
                            Name = "Available Serial Port (Potential OBD)"
                            Type = "Serial"
                            Status = "Available"
                        }
                    } catch {
                        # Port not available or in use
                    }
                }
            }
            
            # Output results
            $devices | Sort-Object Port | ForEach-Object {
                Write-Output "$($_.Port)|$($_.Name)|$($_.Type)|$($_.Status)"
            }
            """

            result = subprocess.run(
                ["powershell", "-Command", ps_command],
                capture_output=True,
                text=True,
                timeout=15
            )

            bt_devices_found = 0
            existing_ports = [self.port_select.itemData(
                j) for j in range(self.port_select.count())]

            if result.returncode == 0 and result.stdout:
                lines = result.stdout.strip().split('\n')
                for line in lines:
                    if '|' in line:
                        try:
                            parts = line.split('|')
                            if len(parts) >= 4:
                                port, name, device_type, status = parts[:4]
                                port = port.strip()
                                name = name.strip()
                                device_type = device_type.strip()
                                status = status.strip()

                                if port and port not in existing_ports:
                                    if device_type == "Bluetooth":
                                        display_name = f"{port} (🔵 Bluetooth - {name})"
                                        self.port_select.addItem(
                                            display_name, port)
                                        bt_devices_found += 1
                                        print(
                                            f"[BLUETOOTH] Found paired device: {port} - {name}")
                                    elif "OBD" in name.upper() or "ELM" in name.upper():
                                        display_name = f"{port} (🔧 Potential OBD - {name})"
                                        self.port_select.addItem(
                                            display_name, port)
                                        bt_devices_found += 1
                                        print(
                                            f"[BLUETOOTH] Found potential OBD: {port} - {name}")
                        except Exception as parse_error:
                            print(
                                f"[BLUETOOTH] Parse error for line '{line}': {parse_error}")
                            continue

            # Additional fallback: Check for common OBD Bluetooth patterns
            try:
                # Look for devices with OBD-related keywords
                obd_keywords = ["OBD", "ELM327", "ELM",
                                "OBDII", "OBD2", "DIAGNOSTIC"]

                for i in range(1, 21):
                    port_name = f"COM{i}"
                    if port_name not in existing_ports:
                        try:
                            import serial as pyserial
                            # Try to open with common OBD settings
                            test_serial = pyserial.Serial(
                                port_name,
                                baudrate=38400,  # Common OBD baud rate
                                timeout=0.5,
                                write_timeout=0.5
                            )

                            # Send AT command to test if it's an OBD device
                            test_serial.write(b'ATZ\r')  # Reset command
                            test_serial.flush()
                            time.sleep(0.2)

                            response = test_serial.read(100)
                            test_serial.close()

                            if response and (b'ELM' in response or b'OK' in response or b'>' in response):
                                display_name = f"{port_name} (🔧 Detected OBD Device)"
                                self.port_select.addItem(
                                    display_name, port_name)
                                bt_devices_found += 1
                                print(
                                    f"[BLUETOOTH] Detected OBD device on {port_name}: {response}")
                            elif len(response) > 0:  # Some response, might be Bluetooth
                                display_name = f"{port_name} (🔵 Bluetooth Device)"
                                self.port_select.addItem(
                                    display_name, port_name)
                                bt_devices_found += 1
                                print(
                                    f"[BLUETOOTH] Detected Bluetooth device on {port_name}")

                        except Exception as test_error:
                            # Port not available, in use, or not responsive
                            continue

            except Exception as fallback_error:
                print(f"[BLUETOOTH] Fallback scan error: {fallback_error}")

            if bt_devices_found > 0:
                self.update_status(
                    f"Found {bt_devices_found} Bluetooth/OBD devices", "success")
                print(
                    f"[BLUETOOTH] Successfully found {bt_devices_found} devices")
            else:
                self.update_status(
                    "No Bluetooth OBD devices found. Check pairing and try manual entry.", "warning")
                print("[BLUETOOTH] No devices found - check if OBD adapter is paired")

        except Exception as e:
            print(f"[BLUETOOTH] Scan error: {e}")
            self.update_status(
                "Bluetooth scan failed. Try manual COM port entry.", "error")

    def refresh_ports(self):
        """Refresh and populate the port selection with all available serial ports"""
        self.port_select.clear()

        # Get standard serial ports with enhanced descriptions
        ports = serial.tools.list_ports.comports()
        bluetooth_ports = []
        standard_ports = []

        for port in ports:
            description = f"{port.device}"
            port_info = ""

            # Enhanced description with device details
            if port.description and port.description != "n/a":
                port_info = f" ({port.description})"

            # Check for Bluetooth indicators
            is_bluetooth = any(bt_keyword in port.description.lower() if port.description else ""
                               for bt_keyword in ["bluetooth", "bth", "bt"])

            # Check for OBD indicators
            is_obd = any(obd_keyword in port.description.lower() if port.description else ""
                         for obd_keyword in ["elm327", "elm", "obd", "diagnostic"])

            if is_bluetooth or is_obd:
                if is_obd:
                    display_name = f"{description} 🔧 OBD{port_info}"
                else:
                    display_name = f"{description} 🔵 Bluetooth{port_info}"
                bluetooth_ports.append((display_name, port.device))
            else:
                display_name = f"{description}{port_info}"
                standard_ports.append((display_name, port.device))

        # Add Bluetooth ports first (higher priority)
        for display_name, device in bluetooth_ports:
            self.port_select.addItem(display_name, device)

        # Add standard ports
        for display_name, device in standard_ports:
            self.port_select.addItem(display_name, device)

        # Windows-specific Bluetooth port detection
        import platform
        if platform.system() == "Windows":
            try:
                import winreg
                existing_ports = [self.port_select.itemData(
                    j) for j in range(self.port_select.count())]

                # Check registry for additional Bluetooth COM ports
                key_path = r"HARDWARE\DEVICEMAP\SERIALCOMM"
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                    i = 0
                    while True:
                        try:
                            name, value, _ = winreg.EnumValue(key, i)
                            if ("BthModem" in name or "Bluetooth" in name) and value not in existing_ports:
                                self.port_select.addItem(
                                    f"{value} 🔵 Registry Bluetooth", value)
                                print(
                                    f"[PORTS] Found registry Bluetooth port: {value}")
                            i += 1
                        except WindowsError:
                            break
            except Exception as e:
                print(f"[PORTS] Registry scan error: {e}")

            # Enhanced Bluetooth port probing with better detection
            try:
                existing_ports = [self.port_select.itemData(
                    j) for j in range(self.port_select.count())]

                for i in range(1, 21):  # COM1 to COM20
                    port_name = f"COM{i}"
                    if port_name not in existing_ports:
                        try:
                            import serial as pyserial
                            import time

                            # Try multiple common OBD/Bluetooth baud rates
                            baud_rates = [38400, 9600, 115200]
                            device_detected = False
                            device_type = "Unknown"

                            for baud_rate in baud_rates:
                                if device_detected:
                                    break

                                try:
                                    test_serial = pyserial.Serial(
                                        port_name,
                                        baudrate=baud_rate,
                                        timeout=0.3,
                                        write_timeout=0.3
                                    )

                                    # Test if it's an OBD device
                                    test_serial.write(b'ATZ\r\n')
                                    test_serial.flush()
                                    time.sleep(0.2)

                                    response = test_serial.read(50)
                                    test_serial.close()

                                    if response:
                                        response_str = response.decode(
                                            'ascii', errors='ignore').upper()
                                        if any(obd_resp in response_str for obd_resp in ['ELM327', 'ELM', 'OK', '>']):
                                            device_type = "OBD Device"
                                            device_detected = True
                                            print(
                                                f"[PORTS] Detected OBD device on {port_name} at {baud_rate} baud: {response_str[:20]}")
                                        # Some response indicates active device
                                        elif len(response) > 2:
                                            device_type = "Bluetooth Device"
                                            device_detected = True
                                            print(
                                                f"[PORTS] Detected Bluetooth device on {port_name} at {baud_rate} baud")

                                except Exception as test_error:
                                    continue  # Try next baud rate

                            if device_detected:
                                if "OBD" in device_type:
                                    display_name = f"{port_name} 🔧 Detected {device_type}"
                                else:
                                    display_name = f"{port_name} 🔵 Detected {device_type}"
                                self.port_select.addItem(
                                    display_name, port_name)

                        except Exception as port_error:
                            continue  # Port not available or permission denied

            except Exception as probe_error:
                print(f"[PORTS] Bluetooth probe error: {probe_error}")

        # If no ports found, show helpful message
        if self.port_select.count() == 0:
            self.port_select.addItem(
                "⚠️ No ports detected - Try pairing Bluetooth OBD or manual entry", "")

        # Add manual entry option at the end
        self.port_select.addItem("✏️ Manual Entry (Type COM port)", "MANUAL")

        print(
            f"[PORTS] Refresh complete: {self.port_select.count()-1} ports found")

    def connect_obd(self):
        """Enhanced OBD connection with robust Bluetooth support and retry logic"""
        # Get the selected port
        if self.port_select.currentData() == "MANUAL" or self.manual_port_input.isVisible():
            port = self.manual_port_input.text().strip()
        else:
            port = self.port_select.currentData(
            ) or self.port_select.currentText().split()[0]

        if not port:
            self.update_status("No port selected", "error")
            return

        self.connect_btn.setText("🔄 Connecting...")
        self.connect_btn.setEnabled(False)

        # Determine if this is a Bluetooth connection
        is_bluetooth = any(bt_indicator in self.port_select.currentText().lower()
                           for bt_indicator in ["bluetooth", "🔵", "bth"])

        # Determine if this is likely an OBD device
        is_obd_device = any(obd_indicator in self.port_select.currentText().lower()
                            for obd_indicator in ["obd", "elm", "🔧", "diagnostic"])

        print(f"[CONNECTION] Attempting to connect to {port}")
        print(
            f"[CONNECTION] Bluetooth: {is_bluetooth}, OBD Device: {is_obd_device}")

        if OBD_AVAILABLE:
            # Enhanced connection attempt with multiple strategies
            connection_attempts = []

            if is_bluetooth or is_obd_device:
                # Strategy 1: Bluetooth-optimized settings
                connection_attempts.append({
                    'name': 'Bluetooth Optimized',
                    'timeout': 15,
                    'check_voltage': False,
                    'fast': False,
                    'protocol': None
                })

                # Strategy 2: ELM327 specific settings
                connection_attempts.append({
                    'name': 'ELM327 Specific',
                    'timeout': 10,
                    'check_voltage': False,
                    'fast': True,
                    'protocol': '6'  # CAN 11-bit 500kb
                })

                # Strategy 3: Basic Bluetooth settings
                connection_attempts.append({
                    'name': 'Basic Bluetooth',
                    'timeout': 8,
                    'check_voltage': False,
                    'fast': False,
                    'protocol': None
                })
            else:
                # Strategy for standard serial connections
                connection_attempts.append({
                    'name': 'Standard Serial',
                    'timeout': 5,
                    'check_voltage': True,
                    'fast': False,
                    'protocol': None
                })

            # Try each connection strategy
            for i, strategy in enumerate(connection_attempts):
                try:
                    self.update_status(
                        f"Trying {strategy['name']} connection... ({i+1}/{len(connection_attempts)})", "warning")
                    print(f"[CONNECTION] Strategy {i+1}: {strategy['name']}")

                    # Pre-connection test for Bluetooth devices
                    if is_bluetooth:
                        if not self.test_bluetooth_port(port):
                            print(
                                f"[CONNECTION] Bluetooth port {port} pre-test failed")
                            continue

                    # Build connection parameters
                    connection_params = {
                        'portstr': port,
                        'timeout': strategy['timeout'],
                        'check_voltage': strategy['check_voltage'],
                        'fast': strategy['fast']
                    }

                    if strategy['protocol']:
                        connection_params['protocol'] = strategy['protocol']

                    # Attempt connection
                    self.connection = obd.OBD(**connection_params)

                    if self.connection and self.connection.is_connected():
                        # Connection successful
                        self.connect_btn.setText("✅ Connected")
                        connection_type = "Bluetooth" if is_bluetooth else "Serial"
                        self.update_status(
                            f"Connected to {port} via {connection_type} - REAL DATA", "success")

                        # Get supported commands info
                        supported_commands = len(
                            self.connection.supported_commands)
                        protocol = getattr(
                            self.connection, 'protocol', 'Unknown')

                        self.info_label.setText(
                            f"Successfully connected to OBD device on {port} using {strategy['name']}.\n"
                            f"Protocol: {protocol}, Supported Commands: {supported_commands}\n"
                            f"Using REAL sensor data only."
                        )

                        print(
                            f"[CONNECTION] Success with {strategy['name']} - Protocol: {protocol}, Commands: {supported_commands}")

                        # Start data collection
                        self.timer.start(500)
                        self.disconnect_btn.setEnabled(True)
                        return
                    else:
                        print(
                            f"[CONNECTION] Strategy {strategy['name']} failed - not connected")
                        if self.connection:
                            self.connection.close()
                            self.connection = None
                        continue

                except Exception as e:
                    print(
                        f"[CONNECTION] Strategy {strategy['name']} exception: {e}")
                    if hasattr(self, 'connection') and self.connection:
                        try:
                            self.connection.close()
                        except:
                            pass
                        self.connection = None
                    continue

            # All connection attempts failed
            self.connect_btn.setText("❌ Connection Failed")
            self.connect_btn.setEnabled(True)

            if is_bluetooth:
                error_msg = (
                    f"Failed to connect to Bluetooth OBD device on {port}.\n\n"
                    "Troubleshooting:\n"
                    "• Ensure the OBD adapter is paired in Windows Bluetooth settings\n"
                    "• Check that the device is plugged into your vehicle's OBD port\n"
                    "• Turn on your vehicle's ignition (engine doesn't need to run)\n"
                    "• Try disconnecting and reconnecting the Bluetooth adapter\n"
                    "• Some adapters require the engine to be running"
                )
            else:
                error_msg = (
                    f"Failed to establish connection on {port}.\n"
                    "Check device connection, port selection, and try again."
                )

            self.update_status(
                "Connection failed - See troubleshooting tips", "error")
            self.info_label.setText(error_msg)
        else:
            # Demo mode fallback
            self.connection = "DEMO"
            self.connect_btn.setText("🎮 Demo Mode")
            self.update_status("Demo Mode Active - SIMULATED DATA", "warning")
            self.info_label.setText(
                "Running in demo mode with simulated data for testing purposes only")
            self.timer.start(500)

    def disconnect_obd(self):
        """Disconnect from OBD device and clean up connection"""
        try:
            # Stop the data timer first
            if hasattr(self, 'timer') and self.timer.isActive():
                self.timer.stop()

            # Close the OBD connection
            if hasattr(self, 'connection') and self.connection and self.connection != "DEMO":
                try:
                    self.connection.close()
                    print("[DISCONNECT] OBD connection closed successfully")
                except Exception as e:
                    print(f"[DISCONNECT] Error closing connection: {e}")

            # Reset connection state
            self.connection = None

            # Update UI
            self.connect_btn.setText("🔌 Connect to ELM327")
            self.connect_btn.setEnabled(True)
            self.disconnect_btn.setEnabled(False)
            self.update_status("Disconnected", "warning")
            self.info_label.setText(
                "Select a port and click connect to start monitoring")

            # Clear gauge displays
            for cmd, label in self.labels.items():
                label.setText("---")

            # Update window title
            self.setWindowTitle("OBD-II Professional Monitor & VE Calculator")

            print("[DISCONNECT] Disconnection completed successfully")

        except Exception as e:
            print(f"[DISCONNECT] Error during disconnection: {e}")
            self.update_status("Disconnect error", "error")

    def test_bluetooth_port(self, port):
        """Test if a Bluetooth port is responsive before attempting OBD connection"""
        try:
            import serial as pyserial
            import time

            print(f"[BT_TEST] Testing Bluetooth port {port}")

            # Try to establish basic serial communication
            test_serial = pyserial.Serial(
                port,
                baudrate=38400,  # Common OBD baud rate
                timeout=2,
                write_timeout=2,
                bytesize=8,
                parity='N',
                stopbits=1
            )

            time.sleep(0.5)  # Allow connection to stabilize

            # Clear any existing data
            test_serial.reset_input_buffer()
            test_serial.reset_output_buffer()

            # Send a basic AT command
            test_serial.write(b'ATZ\r')
            test_serial.flush()
            time.sleep(1)

            # Read response
            response = test_serial.read(100)
            test_serial.close()

            if response:
                response_str = response.decode('ascii', errors='ignore')
                print(f"[BT_TEST] Port {port} responded: {response_str[:50]}")
                return True
            else:
                print(f"[BT_TEST] Port {port} no response")
                return False

        except Exception as e:
            print(f"[BT_TEST] Port {port} test failed: {e}")
            return False

    def update_status(self, text, status_type):
        """Update connection status with appropriate styling"""
        colors = {
            'success': MODERN_STYLE['success'],
            'error': MODERN_STYLE['error'],
            'warning': MODERN_STYLE['warning']
        }

        color = colors.get(status_type, MODERN_STYLE['text_secondary'])
        self.status_label.setText(text)
        self.status_label.setStyleSheet(f"""
            QLabel {{
                color: {color};
                font-size: 16px;
                font-weight: 600;
                padding: 10px;
                background-color: rgba({color[1:3]}, {color[3:5]}, {color[5:7]}, 0.1);
                border: 1px solid {color};
                border-radius: 6px;
            }}
        """)

    def update_pids(self):
        if not self.connection:
            return

        # Update header with data source indicator
        if self.connection == "DEMO":
            self.setWindowTitle(
                "OBD-II Professional Monitor & VE Calculator - DEMO MODE (SIMULATED DATA)")
        else:
            self.setWindowTitle(
                "OBD-II Professional Monitor & VE Calculator - CONNECTED (REAL DATA)")

        rpm = self.get_value(obd.commands.RPM)
        map_kpa = self.get_value(obd.commands.INTAKE_PRESSURE)
        temp_k = self.get_value(obd.commands.INTAKE_TEMP, "kelvin")
        timing_advance = self.get_value(obd.commands.TIMING_ADVANCE)
        throttle = self.get_value(obd.commands.THROTTLE_POS)
        o2_b1s1 = self.get_value(obd.commands.O2_B1S1)
        o2_b2s1 = self.get_value(obd.commands.O2_B2S1)

        # Prepare log row
        log_row = [time.strftime("%Y-%m-%d %H:%M:%S")]
        values = {}
        for cmd, label in self.labels.items():
            val = self.get_value(cmd)
            values[cmd] = val
            if val is not None:
                if cmd == obd.commands.TIMING_ADVANCE:
                    label.setText(f"{val:.1f}°")
                elif cmd == obd.commands.INTAKE_PRESSURE:
                    label.setText(f"{val:.1f}")
                elif cmd == obd.commands.THROTTLE_POS:
                    label.setText(f"{val:.1f}%")
                elif cmd == obd.commands.RPM:
                    label.setText(f"{val:.0f}")
                elif cmd == obd.commands.SPEED:
                    label.setText(f"{val:.0f}")
                elif cmd == obd.commands.COOLANT_TEMP:
                    label.setText(f"{val:.1f}°C")
                elif cmd == obd.commands.INTAKE_TEMP:
                    if val > 200:  # Kelvin
                        label.setText(f"{val-273.15:.1f}°C")
                    else:
                        label.setText(f"{val:.1f}°C")
                elif cmd == obd.commands.MAF:
                    label.setText(f"{val:.2f}")
                elif cmd == obd.commands.O2_B1S1:
                    # Use the selected display format for O2 sensors
                    display_format = getattr(self, 'o2_display_combo', None)
                    if display_format:
                        formatted_value = self.convert_o2_sensor_value(
                            val, display_format.currentText())
                        label.setText(formatted_value)
                    else:
                        # Fallback to lambda if combo box not available
                        afr = 14.7 * (val / 0.45)
                        lambda_ratio = 14.7 / afr
                        label.setText(f"{lambda_ratio:.3f} λ")
                elif cmd == obd.commands.O2_B2S1:
                    # Use the selected display format for O2 sensors
                    display_format = getattr(self, 'o2_display_combo', None)
                    if display_format:
                        formatted_value = self.convert_o2_sensor_value(
                            val, display_format.currentText())
                        label.setText(formatted_value)
                    else:
                        # Fallback to lambda if combo box not available
                        afr = 14.7 * (val / 0.45)
                        lambda_ratio = 14.7 / afr
                        label.setText(f"{lambda_ratio:.3f} λ")
                else:
                    label.setText(f"{val:.2f}")
            else:
                label.setText("---")

        # Only log if logging is enabled
        if self.logging_enabled:
            def safe(val):
                return "" if val is None else val
            log_row.append(safe(values.get(obd.commands.RPM)))
            log_row.append(safe(values.get(obd.commands.SPEED)))
            log_row.append(safe(values.get(obd.commands.COOLANT_TEMP)))
            log_row.append(safe(values.get(obd.commands.INTAKE_PRESSURE)))
            log_row.append(safe(values.get(obd.commands.INTAKE_TEMP)))
            log_row.append(safe(values.get(obd.commands.THROTTLE_POS)))
            log_row.append(safe(values.get(obd.commands.MAF)))
            log_row.append(safe(values.get(obd.commands.TIMING_ADVANCE)))
            # O2 sensors in the selected format
            o2b1 = values.get(obd.commands.O2_B1S1)
            o2b2 = values.get(obd.commands.O2_B2S1)

            # Get current display format for logging
            display_format = getattr(self, 'o2_display_combo', None)
            current_format = display_format.currentText() if display_format else "Lambda (λ)"

            # Convert O2 values according to selected format
            if o2b1 is not None:
                if "Lambda" in current_format:
                    afr1 = 14.7 * (o2b1 / 0.45)
                    o2_value1 = 14.7 / afr1  # Lambda
                elif "Equivalence" in current_format:
                    afr1 = 14.7 * (o2b1 / 0.45)
                    lambda1 = 14.7 / afr1
                    o2_value1 = 1.0 / lambda1  # Phi (equivalence ratio)
                else:  # Voltage
                    o2_value1 = o2b1
            else:
                o2_value1 = ""

            if o2b2 is not None:
                if "Lambda" in current_format:
                    afr2 = 14.7 * (o2b2 / 0.45)
                    o2_value2 = 14.7 / afr2  # Lambda
                elif "Equivalence" in current_format:
                    afr2 = 14.7 * (o2b2 / 0.45)
                    lambda2 = 14.7 / afr2
                    o2_value2 = 1.0 / lambda2  # Phi (equivalence ratio)
                else:  # Voltage
                    o2_value2 = o2b2
            else:
                o2_value2 = ""

            log_row.append(o2_value1)
            log_row.append(o2_value2)
            self.log_data.append(log_row)

        # --- VE Table Update: Only update the cell for the current real-time sensor values ---
        maf = self.get_value(obd.commands.MAF)  # Mass Air Flow in grams/sec
        num_cyl = 8  # Number of cylinders (8-cylinder engine)

        # Debug: Print all sensor values for troubleshooting with data source indicator
        data_source = "DEMO" if self.connection == "DEMO" else "REAL"
        print(
            f"[SENSOR DEBUG - {data_source}] RPM={rpm}, MAP={map_kpa}, MAF={maf}, TempK={temp_k}")

        # Ensure all variables are valid and positive for VE calculation
        if (maf is not None and maf > 0 and
            rpm is not None and rpm > 0 and
            map_kpa is not None and map_kpa > 0 and
            temp_k is not None and temp_k > 0 and
                num_cyl > 0):

            # Find closest indices for current RPM and MAP in the table
            r_idx = int(np.abs(rpm_axis - rpm).argmin())
            m_idx = int(np.abs(map_axis - map_kpa).argmin())

            try:
                # VE Calculation Formula:
                # g_per_cyl = (MAF * 60) / (RPM/2 * num_cyl)  [corrected formula]
                # The factor of 60 converts from grams/sec to grams/min
                # RPM/2 because there are 2 crankshaft revolutions per engine cycle (4-stroke)
                g_per_cyl = (maf * 60) / ((rpm / 2) * num_cyl)

                # VE = (g_per_cyl * Temp_K) / MAP_kPa
                # This gives volumetric efficiency as a dimensionless ratio
                ve = (g_per_cyl * temp_k) / map_kpa

                # Update the specific cell in the VE table
                item = QTableWidgetItem(f"{ve:.3f}")
                item.setTextAlignment(Qt.AlignCenter)

                # Color code the cell based on VE value for visual feedback (adjusted for 8-cylinder engine)
                if ve > 0.45:
                    # Green for high VE (8-cyl: >0.45 vs 4-cyl: >0.9)
                    item.setBackground(QColor(16, 124, 16, 100))
                elif ve > 0.35:
                    # Orange for medium VE (8-cyl: >0.35 vs 4-cyl: >0.7)
                    item.setBackground(QColor(255, 140, 0, 100))
                else:
                    # Red for low VE (8-cyl: ≤0.35 vs 4-cyl: ≤0.7)
                    item.setBackground(QColor(209, 52, 56, 100))

                # Add visual indicator for demo vs real data
                if self.connection == "DEMO":
                    # Gray overlay for demo data
                    item.setBackground(QColor(128, 128, 128, 80))

                self.ve_table_widget.setItem(m_idx, r_idx, item)

                print(
                    f"[VE CALCULATED - {data_source}] RPM={rpm}, MAP={map_kpa}kPa, MAF={maf}g/s, TempK={temp_k}K, VE={ve:.3f} at table[{m_idx},{r_idx}]")

            except Exception as ve_ex:
                print(f"[VE ERROR] Calculation failed: {ve_ex}")
        else:
            # Log why VE calculation was skipped
            missing_sensors = []
            if maf is None or maf <= 0:
                missing_sensors.append("MAF")
            if rpm is None or rpm <= 0:
                missing_sensors.append("RPM")
            if map_kpa is None or map_kpa <= 0:
                missing_sensors.append("MAP")
            if temp_k is None or temp_k <= 0:
                missing_sensors.append("IAT")

            if missing_sensors:
                print(
                    f"[VE SKIP - {data_source}] Missing or invalid sensors: {', '.join(missing_sensors)}")

        # Do not clear or overwrite the rest of the table - preserve historical data

        # --- Update visualizations with multiple data series ---
        max_samples = 100

        # Update RPM data
        if rpm is not None:
            self.plot_data['rpm'].append(rpm)
            if len(self.plot_data['rpm']) > max_samples:
                self.plot_data['rpm'] = self.plot_data['rpm'][-max_samples:]
            self.plot_curves['rpm'].setData(self.plot_data['rpm'])

        # Update MAP data
        if map_kpa is not None:
            self.plot_data['map'].append(map_kpa)
            if len(self.plot_data['map']) > max_samples:
                self.plot_data['map'] = self.plot_data['map'][-max_samples:]
            self.plot_curves['map'].setData(self.plot_data['map'])

        # Update Timing Advance data
        if timing_advance is not None:
            self.plot_data['timing_advance'].append(timing_advance)
            if len(self.plot_data['timing_advance']) > max_samples:
                self.plot_data['timing_advance'] = self.plot_data['timing_advance'][-max_samples:]
            self.plot_curves['timing_advance'].setData(
                self.plot_data['timing_advance'])

        # Update Throttle data
        if throttle is not None:
            self.plot_data['throttle'].append(throttle)
            if len(self.plot_data['throttle']) > max_samples:
                self.plot_data['throttle'] = self.plot_data['throttle'][-max_samples:]
            self.plot_curves['throttle'].setData(self.plot_data['throttle'])

        # Update Throttle data
        if throttle is not None:
            self.plot_data['throttle'].append(throttle)
            if len(self.plot_data['throttle']) > max_samples:
                self.plot_data['throttle'] = self.plot_data['throttle'][-max_samples:]
            self.plot_curves['throttle'].setData(self.plot_data['throttle'])

    def get_value(self, cmd, to_unit=None):
        if not self.connection:
            return None

        # Demo mode ONLY - never use simulated data when connected to real device
        if self.connection == "DEMO":
            # Return simulated values that correlate for realistic VE calculations
            if cmd == obd.commands.RPM:
                return random.randint(800, 6000)
            elif cmd == obd.commands.SPEED:
                return random.randint(0, 120)
            elif cmd == obd.commands.COOLANT_TEMP:
                return random.randint(80, 95)
            elif cmd == obd.commands.MAP or cmd == obd.commands.INTAKE_PRESSURE:
                return random.randint(20, 100)
            elif cmd == obd.commands.INTAKE_TEMP or cmd == obd.commands.INTAKE_AIR_TEMP:
                temp_c = random.randint(20, 60)
                return temp_c + 273.15 if to_unit == "kelvin" else temp_c
            elif cmd == obd.commands.THROTTLE_POS:
                return random.uniform(0, 100)
            elif cmd == obd.commands.MAF:
                # More realistic MAF values that correlate with typical engine operation
                return random.uniform(3.0, 45.0)
            elif cmd == obd.commands.TIMING_ADVANCE:
                # Typical timing advance range in degrees
                return random.uniform(-5, 35)
            elif cmd == obd.commands.O2_B1S1 or cmd == obd.commands.O2_B2S1:
                # Realistic O2 sensor voltage (0.1-0.9V)
                return random.uniform(0.1, 0.9)
            return random.uniform(10, 100)

        # Real OBD connection - ONLY use actual sensor data
        try:
            if not self.connection.supports(cmd):
                print(f"[OBD WARNING] Command {cmd} not supported by vehicle")
                return None

            response = self.connection.query(cmd)
            if response.is_null():
                print(f"[OBD WARNING] No data received for {cmd}")
                return None

            if to_unit:
                return float(response.value.to(to_unit).magnitude)
            return float(response.value.magnitude)

        except Exception as e:
            print(f"[OBD ERROR] Failed to get {cmd}: {e}")
            return None


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        win = OBDApp()
        win.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"Application startup error: {e}")
        import traceback
        traceback.print_exc()
        input("Press Enter to exit...")  # Keep console open to see error
