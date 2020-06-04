"""
This short example shows how the Oscilloscope class can be used. Assuming a sinusoid signal of ~200-400 mV amplitude and up to a few kHz frequency as input on channel 4 (C4) and an external trigger,
this script asks for the ip address, attempts to connect to the oscilloscope with that address, creates a single beep sound, then sets up a 1s measurement to read out the amplitude of the
aforementioned signal. Returns a list containing a single float, the amplitude in V.
"""

from ofho import Oscilloscope
from time import sleep

osci = Oscilloscope(input("Enter oscilloscope IP (e.g. 123.45.6.7):\n"))
print() #empty line
osci.Beep() #test beeping

osci.ShowChannel("C1", False) #hide channel 1
osci.ShowChannel("C4", True) #show channel 4
osci.SetVoltRange("C4", 0.1) #set voltage units to 100 mV
osci.SetTrigger(trigger_name="Ext", trigger_mode="Normal", trigger_coupling="DC", trigger_type="Edge", trigger_slope_type="Positive", trigger_level="0.00")
osci.HorizontalSetup(horizontal_scale="1.0e-3", horizontal_offset="0.0") #time setup

param = [["P1", "C4", "Amplitude"]] #the parameter to be measured: P1 of the oscilloscope parameters should record the amplitude of C4

osci.SetUpMeasurements(param) #
sleep(1000) #measure for 1 s
osci.ReadOut(param)
