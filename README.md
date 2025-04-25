Agilent 34970A 4-Wire Resistor Measurement Tool

Description: This tool is designed for calibration and measurement of resistors using the Agilent 34970A device. It supports 4-wire measurements and provides a graphical user interface to display and manage the measurement results.

Features:
- Resistor Measurement: Reads values from 30 resistors and displays them in textboxes.
- Device Information: Displays the hardware ID of the connected device.
- Data Display: 
  - Scanned resistor values are shown in labels after pressing the read button.
  - The channels of the device could be activated by activating the checkboxes. 
- Interactive Controls:
  - Measured values are automatically transferred into the 'DataGridView'-table.
  - The row of the table can be selected by the 'NumericUpDown'.

Requirements
- .NET Framework: 4.7.2
- Development Environment: Visual Studio 2022
- Hardware: Agilent 34970A device
- Communication: The tool communicates with the device via a serial interface (COM port).
- Operating System: Windows 11
- Libraries: IVI.Visa.NET for device communication

Prerequisites
1. Keysight IO Libraries Suite:
   - Install the Keysight IO Libraries Suite to set up the required VISA drivers.
   - Download it from the official Keysight website: [Keysight IO Libraries Suite](https://www.keysight.com/).
2. FTDI Drivers:
   - Install the FTDI drivers for COM port communication under Windows 11, if they are not already installed.
   - Download the drivers from the official FTDI website: [FTDI Drivers](https://ftdichip.com/drivers/).
   - Ensure the COM port for the Agilent 34970A device is correctly recognized in the Windows Device Manager.
