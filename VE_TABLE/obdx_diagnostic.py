#!/usr/bin/env python3
"""
OBDXPROVX Bluetooth Connection Diagnostic Tool
"""
import serial.tools.list_ports
import subprocess
import re


def diagnose_obdxprovx():
    """Comprehensive diagnostic for OBDXPROVX device"""
    print("🔍 OBDXPROVX BLUETOOTH DIAGNOSTIC")
    print("=" * 50)

    # Step 1: Check COM ports
    print("\n1. SCANNING COM PORTS...")
    ports = list(serial.tools.list_ports.comports())
    bluetooth_ports = []
    obdx_ports = []

    for port in ports:
        desc = port.description.upper()
        hwid = (port.hwid or "").upper()

        print(f"   {port.device}: {port.description}")
        print(f"   HWID: {port.hwid or 'N/A'}")

        # Check for Bluetooth keywords
        is_bluetooth = any(keyword in desc for keyword in
                           ['BLUETOOTH', 'BT', 'RFCOMM', 'SPP', 'STANDARD SERIAL'])

        # Check for OBDX specific identifiers
        is_obdx = any(keyword in desc for keyword in
                      ['OBDX', 'PROV', 'ELM', 'OBD'])

        if is_bluetooth:
            bluetooth_ports.append(port.device)
            if is_obdx:
                obdx_ports.append(port.device)
                print(f"   >>> POTENTIAL OBDXPROVX PORT! <<<")
        print()

    print(f"📊 SUMMARY:")
    print(f"   Total COM ports: {len(ports)}")
    print(f"   Bluetooth ports: {bluetooth_ports}")
    print(f"   Potential OBDX ports: {obdx_ports}")

    # Step 2: Check Windows Bluetooth paired devices
    print("\n2. CHECKING PAIRED BLUETOOTH DEVICES...")
    try:
        result = subprocess.run([
            'powershell',
            'Get-PnpDevice | Where-Object {$_.Class -eq "Bluetooth" -and $_.Status -eq "OK"} | Select-Object Name, InstanceId'
        ], capture_output=True, text=True, timeout=10)

        if result.returncode == 0:
            output = result.stdout
            if 'OBDX' in output.upper() or 'PROV' in output.upper():
                print("   ✅ OBDXPROVX device found in paired devices!")
                print("   Device details:")
                for line in output.split('\n'):
                    if 'OBDX' in line.upper() or 'PROV' in line.upper():
                        print(f"      {line.strip()}")
            else:
                print("   ⚠️ OBDXPROVX not found in paired devices")
                print("   Available Bluetooth devices:")
                print(output)
        else:
            print("   ❌ Could not check paired devices")
    except Exception as e:
        print(f"   ❌ Error checking paired devices: {e}")

    # Step 3: OBDXPROVX specific recommendations
    print("\n3. OBDXPROVX SPECIFIC RECOMMENDATIONS:")

    if obdx_ports:
        recommended_port = obdx_ports[0]
        print(f"   🎯 RECOMMENDED: Try connecting to {recommended_port}")
        print(f"   📱 OBDXPROVX typically uses:")
        print(f"      • Baud rate: 38400 or 9600")
        print(f"      • Protocol: ELM327 compatible")
        print(f"      • Connection: Standard Serial over Bluetooth")
    elif bluetooth_ports:
        print(f"   🔧 Try these Bluetooth ports: {bluetooth_ports}")
        print(f"   💡 OBDXPROVX should appear as 'Standard Serial over Bluetooth'")
    else:
        print("   ❌ NO BLUETOOTH PORTS FOUND!")
        print("   🚨 TROUBLESHOOTING STEPS:")
        print("      1. Ensure OBDXPROVX is powered on (vehicle running)")
        print("      2. Pair OBDXPROVX in Windows Bluetooth settings:")
        print("         • Settings → Devices → Bluetooth")
        print("         • Add Bluetooth device")
        print("         • Look for 'OBDXPROVX' or similar")
        print("      3. After pairing, check Device Manager → Ports")
        print("      4. Look for 'Standard Serial over Bluetooth link'")

    print("\n4. OBDXPROVX CONNECTION TIPS:")
    print("   • Make sure vehicle ignition is ON")
    print("   • OBDXPROVX LED should be blinking/solid")
    print("   • Pair device BEFORE plugging into OBD port")
    print("   • Use PIN '1234' or '0000' if prompted")
    print("   • After pairing, note the assigned COM port number")
    print("   • In the app, select 'Bluetooth' connection type")
    print("   • Choose the assigned COM port from dropdown")

    return bluetooth_ports, obdx_ports


if __name__ == "__main__":
    bt_ports, obdx_ports = diagnose_obdxprovx()

    print("\n" + "=" * 50)
    print("🎯 NEXT STEPS:")
    if obdx_ports:
        print(f"   1. In the app, select 'Bluetooth' connection")
        print(f"   2. Choose {obdx_ports[0]} from the device dropdown")
        print(
            f"   3. Click 'Connect Vehicle' or use 'Quick Connect {obdx_ports[0]}'")
    elif bt_ports:
        print(f"   1. Try connecting to: {bt_ports}")
        print(f"   2. If none work, re-pair your OBDXPROVX device")
    else:
        print("   1. Pair your OBDXPROVX device first!")
        print("   2. Run this diagnostic again after pairing")
