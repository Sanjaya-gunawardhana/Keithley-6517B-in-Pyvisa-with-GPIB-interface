import pyvisa
import time
import numpy as np
import pandas as pd

data_points = 5000 # N value needs to update with this value
time_s = 15


# Connect to the Keithley 6517B (replace with your actual VISA address)
visa_address = 'GPIB0::27::INSTR'  # Replace with the actual address (GPIB, USB, etc.)
rm = pyvisa.ResourceManager()
keithley = rm.open_resource(visa_address)

# Reset the instrument to its default settings
keithley.write('*RST')

# Check for any existing errors or status from the instrument
error_status = keithley.query('SYSTEM:ERROR?')
print('Initial Error Status:', error_status)

# Configure the instrument for current measurement
keithley.write(':SENSE:FUNC "CURR"')  # Set to current measurement mode
keithley.write(':SENSE:CURR:RANGE 20E-6')  # Set current range (e.g., 2uA range)
keithley.write(':SENSE:CURR:NPLC 0.01')  # Set NPLC for averaging (default to 1)
keithley.write(':SENS:CURR:MED:STAT OFF')  # Turn off median filtering
keithley.write(':SYST:ZCH OFF')  # Disable zero check


# Disable the display for better performance (if needed)
keithley.write(':DISP:ENABLE OFF')
keithley.write(':SYST:LSYNC:STAT 0')

# Clear the trace and set up data collection
keithley.write(':TRACe:CLEAR')
keithley.write(':TRACE:ELEM TST')
keithley.write(':TRACE:POINTS 5000')  # Collect N data points
keithley.write(':TRIG:COUNT 5000')  # Trigger N measurements
keithley.write(':TRIG:DELAY 0')  # No delay between triggers
keithley.write(':TRACE:FEED:CONTROL NEXT')  # Feed data point by point

# Initialize the measurement process
keithley.write(':INIT')
# Allow some time for measurements to be taken
time.sleep(time_s)  # Set this time based on number of data points and time

# Query the data from the trace buffer
trace_data = keithley.query("TRAC:DATA?")
buffer_array = np.array([trace_data.split('#')])

flattened_data = buffer_array.flatten()
filtered_data = [entry for entry in flattened_data if entry.strip()]  # Remove newlines or empty entries

# Extract current (NADC) and time (secs) values
current = []
time = []

for entry in filtered_data:
    if entry[0] != ',':
        entry = ',' + entry
    parts = entry.split(',')
    nadc_value = parts[1].replace('NADC', '').strip()
    secs_value = parts[2].replace('secs', '').strip()
    current.append(float(nadc_value))  # Convert to float
    time.append(float(secs_value))  # Convert to float

# Create a DataFrame
df = pd.DataFrame({
    'Time(s)': time,
    'Current(A)': current
})

# Save the DataFrame to an Excel file
output_file = 'Current.xlsx'
df.to_excel(output_file, index=False)
print(f"Data saved to {output_file}")
keithley.write(':SYST:ZCH ON')
keithley.write(':DISP:ENABLE ON')
# Check the status of the instrument after the measurement
final_error_status = keithley.query('SYSTEM:ERROR?')
print('Final Error Status:', final_error_status)

# Close the connection to the instrument
keithley.close()
