# OFHO
Oscilloscope from Home Office is a python 3 class utilizing ActiveDSO to make remote measurements with LeCroy Waverunner oscilloscopes.
## Requirements
The ActiveX standard is handled by the python library pywin32, which can be installed in most cases with pip directly from PowerShell:
`pip install pywin32`
## Usage
The class Oscilloscope from the file ofho.py (which should be copied to the working directory) needs to be imported for use:
`from ofho import Oscilloscope`
A new instance can be created when the ip address of the oscilloscope is known:
`oscilloscope = Oscilloscope(ip_address)`
