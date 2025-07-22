#!/usr/bin/env python3
"""
OBDXPROVX Connection Test Script
"""
import obd
import time


def test_obdxprovx_connection():
    """Test connection to OBDXPROVX on COM5 and COM6"""
    print("🔵 OBDXPROVX CONNECTION TEST")
    print("=" * 40)

    ports_to_test = ['COM5', 'COM6']

    for port in ports_to_test:
        print(f"\n🔍 Testing {port}...")

        # Test different baud rates commonly used by OBDXPROVX
        baud_rates = [38400, 9600, 115200, 57600]

        for baud in baud_rates:
            try:
                print(f"   Trying {baud} baud...", end=' ')

                # Attempt connection
                connection = obd.OBD(port, baudrate=baud,
                                     timeout=5, fast=False)

                if connection and connection.is_connected():
                    print(f"✅ SUCCESS!")
                    print(f"   Connected to {port} at {baud} baud")

                    # Test basic communication
                    try:
                        # Test with a simple PID command
                        supported_cmds = connection.supported_commands
                        if supported_cmds:
                            print(
                                f"   📡 Connection verified - {len(supported_cmds)} PIDs supported")
                        else:
                            print(f"   ⚠️ Connected but no supported PIDs found")
                    except Exception as query_error:
                        print(
                            f"   ⚠️ Connected but verification failed: {query_error}")

                    # Close connection
                    connection.close()
                    print(
                        f"   🎯 RECOMMENDED: Use {port} with {baud} baud rate\n")
                    return port, baud

                else:
                    print(f"❌ Failed")
                    if connection:
                        connection.close()

            except Exception as e:
                print(f"❌ Error: {str(e)}")
                continue

    print(f"\n❌ Could not connect to OBDXPROVX on any port")
    print(f"🔧 Make sure:")
    print(f"   • Vehicle ignition is ON")
    print(f"   • OBDXPROVX is plugged into OBD port")
    print(f"   • Device is paired in Windows Bluetooth")
    print(f"   • No other OBD software is running")

    return None, None


if __name__ == "__main__":
    port, baud = test_obdxprovx_connection()

    if port and baud:
        print(
            f"\n🎉 SUCCESS! Your OBDXPROVX is working on {port} at {baud} baud")
        print(f"💡 In the VE Table Monitor app:")
        print(f"   1. Select 'Bluetooth' connection type")
        print(f"   2. Choose '{port}' from device dropdown")
        print(f"   3. Click 'Connect Vehicle'")
        print(f"   4. Or use the '{port} (OBDX)' quick connect button")
    else:
        print(f"\n❌ Connection failed. Check troubleshooting steps above.")
