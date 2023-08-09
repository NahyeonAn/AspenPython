#%%
"""
Libray import
"""
import win32com.client as win32
import os
import numpy as np
import pandas as pd
#%%
filename = '../aspen/Wet_flue_gas_desulfurization(WOS).apw'
sim = win32.Dispatch("Apwn.Document")
sim.InitFromArchive2(os.path.abspath(filename))
sim.Visible = True
# sim.Visible = False

#%%
MyBlocks = sim.Tree.Elements("Data").Elements("Blocks")
MyStreams = sim.Tree.Elements("Data").Elements("Streams")

#%% Material stream
wos = MyStreams.Elements("WOS")

temp = wos.Elements("Output").Elements("RES_TEMP").Value
pres = wos.Elements("Output").Elements("RES_PRES").Value

molar_flow = wos.Elements("Output").Elements("RES_MOLEFLOW").Value
mass_flow = wos.Elements("Output").Elements("RES_MASSFLOW").Value

print("WOS temperature (C): ", temp)
print("WOS pressure (bar): ", pres)

print("WOS mole flows (kmol/hr): ", molar_flow)
print("WOS mass flows (kg/hr): ", mass_flow)

#%% Components
wos_massflow = wos.Elements("Output").Elements("MASSFLOW").Elements("MIXED")    

components_list = []
components_flow_list = []

for comp in wos_massflow.Elements:
    Compoundname = comp.Name
    components_list.append(Compoundname)
    _massflow = wos_massflow.Elements(Compoundname).Value
    components_flow_list.append(_massflow)
    
comp_output = {comp : value for comp,value in zip(components_list, components_flow_list)}
print(comp_output)

#%% Blocks
scrubber = MyBlocks.Elements('SCRUBBER')

heat_duty = scrubber.Elements("Output").Elements("QCALC").Value
total_vol = scrubber.Elements("Output").Elements("TOT_VOL").Value

print("Scrubber heat duty (cal/sec): ", heat_duty)
print("Scrubber reactor volume (L): ", total_vol)

#%% Input change

# Material stream
wos_total_flow = wos.Elements("Input").Elements("TOTFLOW").Elements("MIXED")
wos_total_flow.Value = 375

# Blocks setting
gas_sp_temp = MyBlocks.Elements("GAS-SP").Elements("Input").Elements("TEMP")
gas_sp_temp.Value = 50
#%% Simulation run

sim.Reinit()
sim.Run2()

#%% Save and close Aspen file
sim.Save()
AspenFileName = sim.FullName
sim.Close(AspenFileName)