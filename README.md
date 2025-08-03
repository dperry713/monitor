# OBD-II Professional Monitor & VE Calculator

A professional-grade OBD-II diagnostic tool with real-time engine monitoring, volumetric efficiency analysis, and data logging capabilities.

![OBD-II Professional Monitor](https://img.shields.io/badge/Status-Ready%20for%20Production-brightgreen)
![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![License](https://img.shields.io/badge/License-Open%20Source-green)

## 🚗 Overview

The OBD-II Professional Monitor is a comprehensive diagnostic application designed for automotive enthusiasts, mechanics, and engineers. It provides real-time monitoring of engine parameters with advanced volumetric efficiency calculations optimized for 8-cylinder engines.

### Key Features

- **Real-time Engine Monitoring** - Live display of critical engine parameters
- **Volumetric Efficiency Analysis** - Automated VE table generation and color-coded visualization
- **Professional Data Logging** - CSV export with customizable O2 sensor formats
- **Multi-format O2 Display** - Lambda (λ), Equivalence Ratio (φ), or Voltage (V)
- **Bluetooth & Serial Support** - Compatible with ELM327 adapters
- **Modern Dark Theme** - Professional interface optimized for automotive environments
- **8-Cylinder Optimization** - VE calculations and color coding specifically tuned for V8 engines

## 📊 Screenshots

### Real-Time Dashboard

The main gauge display shows all critical engine parameters in an easy-to-read format with color-coded values.

### VE Table Analysis

Real-time volumetric efficiency calculation with color-coded cells:

- 🟢 **Green**: High efficiency (>0.45 for 8-cyl)
- 🟠 **Orange**: Medium efficiency (>0.35 for 8-cyl)
- 🔴 **Red**: Low efficiency (≤0.35 for 8-cyl)

### Data Visualization

Live plotting of RPM, MAP, timing advance, and throttle position with professional styling.

## 🔧 Installation & Setup

### Option 1: Executable (Recommended)

1. Download `OBD_Monitor.exe` from the `dist` folder
2. Connect your ELM327 OBD-II adapter
3. Run the executable - no installation required!

### Option 2: Run from Source

```bash
# Clone or download the repository
git clone <repository-url>
cd NEW_obd

# Install required dependencies
pip install PyQt5 numpy pandas pyqtgraph python-obd pyserial

# Run the application
python tool.py
```

### Dependencies

- **PyQt5** - Modern GUI framework
- **numpy** - Numerical computations
- **pandas** - Data analysis and export
- **pyqtgraph** - Real-time plotting
- **python-obd** - OBD-II communication library
- **pyserial** - Serial port communication

## 🔌 Hardware Compatibility

### Supported OBD-II Adapters

- **ELM327 Bluetooth** - Wireless connection (recommended)
- **ELM327 USB** - Direct serial connection
- **ELM327 WiFi** - Network-based adapters

### Connection Types

- **Bluetooth COM Ports** - Automatically detected and scanned
- **USB Serial Ports** - Direct USB-to-serial adapters
- **Manual Port Entry** - Custom port specification

### Vehicle Compatibility

- **OBD-II Compliant Vehicles** (1996+ in US, 2001+ in EU)
- **Optimized for 8-cylinder engines** but works with any configuration
- **Gasoline engines** with standard O2 sensors

## 🚀 Quick Start Guide

### 1. Connect Hardware

1. Plug your ELM327 adapter into the vehicle's OBD-II port
2. For Bluetooth: Pair the adapter in Windows Bluetooth settings
3. Note the assigned COM port (usually COM3-COM20)

### 2. Launch Application

1. Run `OBD_Monitor.exe`
2. Navigate to the **Connection** tab
3. Select your adapter from the port dropdown or use "Scan Bluetooth"
4. Click **Connect to ELM327**

### 3. Monitor Data

- **Real-Time Gauges**: View live engine parameters
- **VE Table**: Watch volumetric efficiency calculations populate
- **Visualizations**: Monitor trends with real-time graphs

### 4. Log Data (Optional)

1. In the **Real-Time Gauges** tab, click **Start Logging**
2. Choose O2 sensor display format (Lambda, Equivalence Ratio, or Voltage)
3. Click **Export Log to CSV** to save data

## 📋 Monitored Parameters

### Engine Parameters

| Parameter          | Description                   | Units              |
| ------------------ | ----------------------------- | ------------------ |
| **RPM**            | Engine speed                  | revolutions/minute |
| **Speed**          | Vehicle speed                 | km/h or mph        |
| **Coolant Temp**   | Engine coolant temperature    | °C                 |
| **MAP**            | Manifold Absolute Pressure    | kPa                |
| **IAT**            | Intake Air Temperature        | °C                 |
| **Throttle**       | Throttle position             | %                  |
| **MAF**            | Mass Air Flow                 | g/s                |
| **Timing Advance** | Ignition timing               | degrees            |
| **O2 B1S1**        | Oxygen sensor bank 1 sensor 1 | λ/φ/V              |
| **O2 B2S1**        | Oxygen sensor bank 2 sensor 1 | λ/φ/V              |

### Calculated Values

- **Volumetric Efficiency (VE)** - Real-time calculation using the formula:
  ```
  g_per_cyl = (MAF × 60) / (RPM/2 × num_cyl)
  VE = (g_per_cyl × Temp_K) / MAP_kPa
  ```

## ⚙️ Advanced Features

### O2 Sensor Display Formats

#### Lambda (λ) - Default

- **λ = 1.0**: Perfect stoichiometric ratio (14.7:1 AFR)
- **λ < 1.0**: Rich mixture (excess fuel)
- **λ > 1.0**: Lean mixture (excess air)

#### Equivalence Ratio (φ)

- **φ = 1.0**: Perfect stoichiometric ratio
- **φ > 1.0**: Rich mixture (excess fuel)
- **φ < 1.0**: Lean mixture (excess air)

#### Voltage (V)

- **~0.45V**: Stoichiometric ratio
- **Lower voltage**: Rich mixture
- **Higher voltage**: Lean mixture

### VE Table Color Coding (8-Cylinder Optimized)

The application automatically adjusts VE thresholds for 8-cylinder engines:

- **4-cylinder thresholds**: >0.9 (green), >0.7 (orange)
- **8-cylinder thresholds**: >0.45 (green), >0.35 (orange)

### Data Logging & Export

- **CSV Format**: Standard comma-separated values
- **Dynamic Headers**: Column names update based on O2 display format
- **Timestamp**: Accurate time logging for data analysis
- **Real-time Updates**: Log data while monitoring

## 🛠️ Troubleshooting

### Connection Issues

1. **Bluetooth Pairing**: Ensure adapter is paired in Windows settings first
2. **COM Port**: Try manual port entry if auto-detection fails
3. **Driver Issues**: Install proper drivers for your ELM327 adapter
4. **Timeout**: Bluetooth connections may take 5-10 seconds

### Performance Issues

1. **Demo Mode**: Application includes demo mode for testing without hardware
2. **Slow Updates**: Check connection quality and adapter compatibility
3. **Missing Data**: Some vehicles may not support all OBD-II parameters

### Data Quality

1. **VE Calculations**: Require valid MAF, RPM, MAP, and IAT readings
2. **O2 Sensors**: Ensure sensors are warmed up for accurate readings
3. **Engine Load**: Best VE data obtained under various load conditions

## 📁 File Structure

```
NEW_obd/
├── tool.py                 # Main application source code
├── tool.spec              # PyInstaller build configuration
├── dist/
│   └── OBD_Monitor.exe    # Compiled executable (83.9 MB)
├── build/                 # PyInstaller build files
└── README.md             # This file
```

## 🔧 Development

### Building from Source

```bash
# Install PyInstaller
pip install pyinstaller

# Build executable
pyinstaller tool.spec

# Executable will be created in dist/OBD_Monitor.exe
```

### Customization

- **Engine Configuration**: Modify `num_cyl = 8` in the code for different engines
- **VE Thresholds**: Adjust color coding thresholds in the `update_pids()` method
- **Display Parameters**: Add/remove monitored parameters in the `pid_list`

### Code Structure

- **Modern PyQt5 GUI** with professional dark theme
- **Object-oriented design** with clear separation of concerns
- **Error handling** for robust operation
- **Comprehensive logging** for debugging

## 📝 Technical Specifications

### System Requirements

- **Windows 10/11** (primary support)
- **Python 3.8+** (for source code execution)
- **4GB RAM** minimum
- **USB or Bluetooth** port for OBD-II adapter

### Performance

- **Update Rate**: 500ms (2 Hz) for real-time monitoring
- **Data Precision**: 3 decimal places for calculated values
- **Memory Usage**: ~85MB executable size
- **CPU Usage**: Low impact on system resources

### Data Accuracy

- **OBD-II Standard**: Full compliance with SAE J1979
- **Sensor Calibration**: Uses industry-standard conversion formulas
- **VE Calculations**: Validated against automotive engineering references

## 🤝 Contributing

This project welcomes contributions! Areas for improvement:

- **Additional vehicle support** and parameter monitoring
- **Enhanced data analysis** features
- **Mobile platform** compatibility
- **Advanced tuning** tools integration

## 📄 License

This project is open source and available under standard licensing terms.

## 🔗 Resources

### OBD-II Information

- [SAE J1979 Standard](https://www.sae.org/standards/content/j1979_202104/)
- [OBD-II Parameter List](https://en.wikipedia.org/wiki/OBD-II_PIDs)

### ELM327 Documentation

- [ELM327 Command Reference](https://www.elmelectronics.com/wp-content/uploads/2017/01/ELM327DS.pdf)
- [OBD-II Adapter Compatibility](https://python-obd.readthedocs.io/en/latest/)

### Automotive Engineering

- [Volumetric Efficiency Explained](https://www.enginebasics.com/Advanced%20Engine%20Tuning/Volumetric%20Efficiency.html)
- [Air-Fuel Ratio Fundamentals](https://www.tuningblog.eu/en/categories/tuning-wiki/lambda-air-fuel-ratio-explained/)

---

**Built with ❤️ for automotive enthusiasts and professionals**

_For support or questions, please check the troubleshooting section above or consult the OBD-II and ELM327 documentation._
