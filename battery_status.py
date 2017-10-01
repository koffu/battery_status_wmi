
# install python
# pip install pypiwin32 wmi

import wmi
import time
import os
from shutil import move

c = wmi.WMI()
t = wmi.WMI(moniker = "//./root/wmi")

battery_voltage_levels_file = 'battery_voltage_levels'

# Load Voltage Levels from file
try:
	battery_voltage_levels_file_fd = open(battery_voltage_levels_file, 'r')
except IOError, error_message:
	min_voltage = 0
	max_voltage = 0
else:
	voltage_levels = battery_voltage_levels_file_fd.read().split(':')
	min_voltage = int(voltage_levels[0])
	max_voltage = int(voltage_levels[1])
	print "Min Voltage: %d, Max Voltage %d loaded from file" % (int(min_voltage), int(max_voltage))
	battery_voltage_levels_file_fd.close()

# Immediately write Voltage Levels
def write_levels (min_voltage, max_voltage):
	move(battery_voltage_levels_file, battery_voltage_levels_file + '.backup')
	with open(battery_voltage_levels_file, 'w') as battery_voltage_levels_file_fd:
		voltage_levels = '%d:%d' % (min_voltage, max_voltage)
		print "Writing Voltage Levels {v_min:.3f}:{v_max:.3f}".format(v_min = float(min_voltage)/1000, v_max = float(max_voltage)/1000) 
		battery_voltage_levels_file_fd.write(voltage_levels) 

counter = 0
timestamp_start = int(time.time())
battery_level_start = 0
while True:
	counter = counter + 1
	batts = t.ExecQuery('Select * from BatteryStatus where Voltage > 0')
	for i, b in enumerate(batts):
		current_voltage = int(b.Voltage)
		# print 'Curent Voltage:           ' + str(current_voltage)
	if (int(b.Voltage) > max_voltage or max_voltage == 0):
		max_voltage = int(b.Voltage)
		#print "Max Voltage %d" % max_voltage
		write_levels (min_voltage, max_voltage)
	if (int(b.Voltage) < min_voltage or min_voltage == 0):
		min_voltage = int(b.Voltage)
		#print "Min Voltage %d" % min_voltage
		write_levels (min_voltage, max_voltage)
	if ( counter % 6 == 0 or counter == 1 ):
		# Battery Level calculation
		voltage_diff_battery_full = max_voltage - min_voltage
		voltage_diff_current = current_voltage - min_voltage
		battery_level = voltage_diff_current * 100 / voltage_diff_battery_full
		
		# Remaining time calculation
		if ( battery_level_start == 0 ):
			battery_level_start = battery_level
		battery_level_discharging_from_start = battery_level_start - battery_level
		battery_time_discharging_from_start = int(time.time()) - timestamp_start

		# Todo when charing
		if (b.Charging == True):
			remaining_time = 'charging'
			timestamp_start = int(time.time())
		# Wait 60 seconds for battery voltage level establishing
		elif (int(time.time()) - timestamp_start < 120):
			battery_level_start = battery_level
			remaining_time = 'estimating in %d s' % int(120 - (int(time.time()) - int(timestamp_start)))
			# print int(time.time()) - timestamp_start
		elif ( battery_level_discharging_from_start > 0 ):
			remaining_time_seconds = battery_time_discharging_from_start * battery_level / battery_level_discharging_from_start
			remaining_time = time.strftime("%H:%M:%S", time.gmtime(remaining_time_seconds))
		elif ( battery_level < 5 ):
			shutdown /s /hybrid
		else:
			remaining_time = '-'
		os.system('mode con: cols=40 lines=8')
		os.system('cls')
		print "Battery Level: {battery_level_percent}\n\rRemaining time: {battery_remaining_time}\n\rVoltages - Now:{voltage_now:.3f}V {v1_now:.2f}V/1\n\rMin:{voltage_min:.3f}V {v1_min:.2f}V/1\n\rMax:{voltage_max:.3f}V {v1_max:.2f}/1\n\r".format(
			battery_level_percent = battery_level,
			battery_remaining_time = remaining_time,
			voltage_now = float(current_voltage) / 1000,
			voltage_min = float(min_voltage) / 1000,
			voltage_max = float(max_voltage) / 1000,
			v1_now = float(current_voltage) / 3000,
			v1_min = float(min_voltage) / 3000,
			v1_max = float(max_voltage) / 3000,
		)
	time.sleep(5	)
	
