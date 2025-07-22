#!/usr/bin/env python3
"""
Build script for VE Table Monitor executable
Creates a standalone .exe with all dependencies included
"""

import PyInstaller.__main__
import os
import sys


def build_exe():
    """Build the VE Table Monitor as a standalone executable"""

    # Get the current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    tool_path = os.path.join(current_dir, 'tool.py')

    # PyInstaller arguments
    args = [
        '--onefile',  # Create a single executable file
        '--windowed',  # Hide console window (for GUI apps)
        '--name=VE_Table_Monitor',  # Name of the executable
        '--icon=NONE',  # No icon for now
        '--add-data=requirements.txt;.',  # Include requirements.txt
        '--hidden-import=bleak',  # Ensure bleak is included
        '--hidden-import=serial',  # Ensure pyserial is included
        '--hidden-import=serial.tools.list_ports',  # Serial port tools
        '--hidden-import=obd',  # OBD library
        '--hidden-import=asyncio',  # Async support
        '--hidden-import=tkinter',  # GUI library
        '--hidden-import=tkinter.ttk',  # Modern tkinter widgets
        '--collect-all=bleak',  # Collect all bleak modules
        '--collect-all=obd',  # Collect all OBD modules
        '--distpath=dist',  # Output directory
        '--workpath=build',  # Build directory
        '--specpath=.',  # Spec file location
        tool_path  # Main Python file
    ]

    print("Building VE Table Monitor executable...")
    print("This may take a few minutes...")

    # Run PyInstaller
    PyInstaller.__main__.run(args)

    print("\nBuild complete!")
    print(
        f"Executable created: {os.path.join(current_dir, 'dist', 'VE_Table_Monitor.exe')}")
    print("\nThe .exe file includes:")
    print("- Python runtime")
    print("- All required libraries (OBD, Bleak, tkinter, etc.)")
    print("- No external dependencies needed")
    print("\nYou can distribute this single .exe file to run on any Windows machine!")


if __name__ == "__main__":
    build_exe()
