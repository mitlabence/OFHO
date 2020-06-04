import win32com.client #import the pywin32 library
import time #for sleep()
from typing import List #annotation: return List[float]
"""
Documentation:
(short) http://cdn.teledynelecroy.com/files/manuals/activedso-developers-guide.pdf
(longer) http://cdn.teledynelecroy.com/files/manuals/entire_x-stream_automation_manual.pdf
(command reference manual, very useful) http://cdn.teledynelecroy.com/files/manuals/automation_command_ref_manual_wr.pdf
"""
class Oscilloscope:
        """
        See http://cdn.teledynelecroy.com/files/manuals/automation_command_ref_manual_wr.pdf for commands
        """
        def __init__(self, ip: str):
            self.ip = ip
            self.osci=win32com.client.Dispatch("LeCroy.ActiveDSOCtrl.1")
            self.osci.MakeConnection("IP:{}".format(self.ip))
            print("Connection to oscilloscope @ {} successful.".format(self.ip))

        def Beep(self):
            self.osci.WriteString("BUZZ BEEP", 1) #http://www.manualsdir.com/manuals/353435/teledyne-lecroy-x-stream-oscilloscopes-remote-control.html?page=278

        def TypeCommand(self):
            inp = input("Enter command to be sent to oscilloscope. Type '? <command>' if answer from osci is expected! ")
            if(inp[0] == "?"):
                self.osci.WriteString("""VBS? {} """.format(inp[2:]), 1)
                return float(self.osci.ReadString(16))
            self.osci.WriteString("""VBS {} """.format(inp), 1)
            return None #TODO: is this here for timing reasons?

        def ShowChannel(self, channel: str="C1", to_state: bool=True) -> None:
            """
            Turn channel on (to_state = True) or off (to_state = False)
            """
            self.osci.WriteString("""VBS 'app.Acquisition.{}.View = {}' """.format(channel, str(to_state)), 1)

        def SetVoltRange(self, channel: str="C1", voltage_scale: float=0.5) -> None:
            self.osci.WriteString("""VBS 'app.Acquisition.{}.VerScale = {}' """.format(channel, voltage_scale), 1)

        def VerticalSetup(self, channel: str="C1", vertical_scale: str="0.5", vertical_offset: str="0.0") -> None:
            self.osci.WriteString("""VBS 'app.Acquisition.{}.VerScale = {}'  """.format(channel, vertical_scale),1)
            self.osci.WriteString("""VBS 'app.Acquisition.{}.VerOffset = {}'  """.format(channel, vertical_offset),1)

        def HorizontalSetup(self, horizontal_scale: str="1.0e-6", horizontal_offset: str="0.0") -> None:
            self.osci.WriteString("""VBS 'app.Acquisition.Horizontal.HorScale = {}'  """.format(horizontal_scale),1)
            self.osci.WriteString("""VBS 'app.Acquisition.Horizontal.HorOffset = {}'  """.format(horizontal_offset),1)
        def SetUpMeasurements(self, params = [["P1", "C1", "Amplitude"]]) -> None:
            """
            Takes an array consisting of [param_name, channel, param_type] triplets. For each triplet, a new entry in the measurement table of the oscilloscope is created,
            and at the end, measurement is started (Clear Sweeps). Use the read_out() function to read these values.
            Example: for channel C1, set up amplitude measurement as P1, for C2, mean measurement as P2:
            set_up_measurements([["P1", "C1", "Amplitude"],["P2", "C2", "Mean"]])
            Some useful param_type parameters: Amplitude, Mean, Minimum, Median, Maximum, PeakToPeak, Rise2080, Fall8020, Period, Frequency
            """
            self.osci.WriteString("""VBS 'app.Measure.ShowMeasure = true'  """,1) #show measurement table
            for parameter in params:
                param_name, channel, variable = parameter
                self.osci.WriteString("""VBS 'app.Measure.{}.View = True'""".format(param_name), 1) #make measured parameter visible
                self.osci.WriteString("""VBS 'app.{}.MeasurementType = "measure" '""".format(param_name), 1) #TODO: not sure if needed
                self.osci.WriteString("""VBS 'app.Measure.{}.Source1="{}" ' """.format(param_name, channel),1)
                self.osci.WriteString("""VBS 'app.Measure.{}.ParamEngine="{}"' """.format(param_name, variable),1)
            self.osci.WriteString("VBS 'app.ClearSweeps'", 1) #Reset statistics

        def ReadOut(self, params = [["P1", "C1", "Amplitude"]]) -> List[float]: #TODO: not only return mean, but also standard deviation, if needed!
            results = []
            for param in params:
                self.osci.WriteString("""VBS? 'return= app.Measure.{}.mean.Result.Value'""".format(param[0]),1)
                res = float(self.osci.ReadString(64))
                results.append(res) #TODO: is 64 bytes of string too much?
            return results
        def ClearSweeps(self) -> None:
            self.osci.WriteString("VBS 'app.ClearSweeps'", 1)

        def SetTrigger(self, trigger_name: str="Ext", trigger_mode: str="Auto", trigger_coupling: str="DC", trigger_type: str="Edge", trigger_slope_type: str="Positive", trigger_level: str="0.00") -> None:
            """
            trigger_mode: can be Auto, Normal, Single, Stopped
            trigger_coupling: DC, not tested with others
            trigger_type: Edge, not tested with others
            TODO: other useful commands might be:
            app.Acquisition.Trigger.Ext.Slope = "Positive"
            app.Acquisition.Trigger.Ext.Level = 0.00 (-0.41 to 0.41 in 0.01 steps, 0 mV works fine on oscilloscope)
            """
            self.osci.WriteString("""VBS 'app.Acquisition.Trigger.Source = "{}"'""".format(trigger_name), 1)
            self.osci.WriteString("""VBS 'app.Acquisition.TriggerMode = "{}"'""".format(trigger_mode), 1)
            self.osci.WriteString("""VBS 'app.Acquisition.Trigger.{}.Coupling = "{}"' """.format(trigger_name, trigger_coupling), 1)
            self.osci.WriteString("""VBS 'app.Acquisition.Trigger.Type = "{}"' """.format(trigger_type), 1)
            self.osci.WriteString("""VBS 'app.Acquisition.Trigger.{}.Slope = "{}"' """.format(trigger_name, trigger_slope_type), 1)
            self.osci.WriteString("""VBS 'app.Acquisition.Trigger.{}.Level = {}' """.format(trigger_name, trigger_level), 1)
            print(f"Trigger is now {trigger_name}, mode is {trigger_mode}, coupling {trigger_coupling}, type {trigger_type}, slope {trigger_slope_type}, level {trigger_level}\n")

        def GetMeanOf(self, channel: str="C4", param_name: str="P1", variable: str="Amplitude") -> float:
            """
            Generalized version of get_mean(), see below; returns *mean value* of specified variable (Amplitude, Peak to peak, 20-80 rise time, ...). That means possible bugs need to be corrected in both!
            Measurement types: see 1-151 to 1-154 (ParamEngine) in http://cdn.teledynelecroy.com/files/manuals/automation_command_ref_manual_wr.pdf
            Some useful ones: Amplitude, Mean, Minimum, Median, Maximum, PeakToPeak, Rise2080, Fall8020, Period, Frequency
            """
            self.osci.WriteString("""VBS 'app.Measure.ShowMeasure = true'  """,1) #show measurement table
            self.osci.WriteString("""VBS 'app.Measure.{}.View = True'""".format(param_name), 1) #make measured parameter visible
            self.osci.WriteString("""VBS 'app.{}.MeasurementType = "measure" '""".format(param_name), 1)
            self.osci.WriteString("""VBS 'app.Measure.{}.Source1="{}" ' """.format(param_name, channel),1) #set source
            self.osci.WriteString("""VBS 'app.Measure.{}.ParamEngine="{}"' """.format(param_name, variable),1) #change "P1" measurement type to "Amplitude", for example
            self.osci.WriteString("VBS 'app.ClearSweeps'", 1) #Reset statistics
            time.sleep(3)
            self.osci.WriteString("""VBS? 'return= app.Measure.{}.mean.Result.Value'""".format(param_name),1)
            return float(self.osci.ReadString(80)) #reads a maximum of 80 bytes

        def GetMeanAmplitude(self, channel: str="C4", param_name: str="P1") -> float:
            """
            Sets up a mean measurement with input at channel (e.g. "C1" - ... "C8" the 8 analog inputs), as parameter param_name ("P1" - "P12")
            """
            self.osci.WriteString("""VBS 'app.Measure.ShowMeasure = true'  """,1) #show measurement table
            self.osci.WriteString("""VBS 'app.Measure.{}.View = True'""".format(param_name), 1) #make measured parameter visible
            #self.osci.WriteString("""VBS 'app.Measure.MeasureMode = "MyMeasure" ' """,1) # TODO: what does this do?
            self.osci.WriteString("""VBS 'app.{}.MeasurementType = "measure"'""".format(param_name), 1)
            self.osci.WriteString("""VBS 'app.Measure.{}.Source1="{}" ' """.format(param_name, channel),1) #set source
            self.osci.WriteString("""VBS 'app.Measure.{}.ParamEngine="Amplitude" ' """.format(param_name),1) #Automation command to change P1 to Mean
            self.osci.WriteString("VBS 'app.ClearSweeps'", 1) #Reset statistics
            time.sleep(3)
            self.osci.WriteString("""VBS? 'return= app.Measure.{}.mean.Result.Value'""".format(param_name),1)
            return float(self.osci.ReadString(80)) #reads a maximum of 80 bytes
        #TODO: not tested!
        def __del__(self):
            self.osci.Disconnect() #Disconnects from the oscilloscope
            print("Oscilloscope @ {} disconnected.\n".format(self.ip))

#Some commands for future reference:

#scope.WriteString("VBS app.Measure.ShowMeasure = true",1) #Automation command to show measurement table
#scope.WriteString("""VBS 'app.Measure.P1.ParamEngine="Mean" ' """,1) #Automation command to change P1 to Mean
#scope.WriteString("VBS? 'return=app.Measure.P1.Out.Result.Value' ",1) #Queries the P1 parameter
#value = scope.ReadString(80)#reads a maximum of 80 bytes
#print(value) #Print value to Interactive Window

#'Perform a recall defaults, which usually puts all channels into DC1M mode.
#app.SetToDefaultSetup
#app.Acquisition.C1.View = True
#app.Acquisition.C1.VerScale = 0.05 #50 mV range
#status = app.Acquisition.Acquire(0,true)
#app.Acquisition.C1.Coupling = "DC50"
