from __future__ import absolute_import
from __future__ import print_function

import os
import sys
import optparse
import openpyxl
import traci
import numpy
# import simpla

# we need to import python modules from the $SUMO_HOME/tools directory
if 'SUMO_HOME' in os.environ:
    tools = os.path.join(os.environ['SUMO_HOME'], 'tools')
    sys.path.append(tools)
else:
    sys.exit("please declare environment variable 'SUMO_HOME'")

from sumolib import checkBinary  # noqa

def get_options():
    optParser = optparse.OptionParser()
    optParser.add_option("--nogui", action="store_true",
                         default=False, help="run the commandline version of sumo")
    options, args = optParser.parse_args()
    return options

def write_excel_xlsx(path, sheet_name, value):
    index = len(value)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    for i in range(0, index):
        # for j in range(0, len(value[i])):
        sheet.cell(row=i+1, column=1, value=str(value[i]))
    workbook.save(path)
    
# this is the main entry point of this script
if __name__ == "__main__":
    options = get_options()

    # this script has been called from the command line. It will start sumo as a
    # server, then connect and run
    if options.nogui:
        sumoBinary = checkBinary('sumo')
    else:
        sumoBinary = checkBinary('sumo-gui')

    # this is the normal way of using traci. sumo is started as a
    # subprocess and then the python script connects and runs
    traci.start([sumoBinary, "-c", "Platooning.sumocfg"])
    # simpla.load("Platooning.cfg.xml")
    speedlist = []
    CO2list = []
    Fuellist = []
    for step in range(30):
        vehlist = traci.vehicle.getIDList()
        if 'veh4' in vehlist:
            speedlist.append(traci.vehicle.getSpeed("veh4"))
            CO2list.append(traci.vehicle.getCO2Emission("veh4"))
            Fuellist.append(traci.vehicle.getFuelConsumption("veh4"))
        else:
            speedlist.append(0)
            CO2list.append(0)
            Fuellist.append(0)
        # print("Speed =",traci.vehicle.getSpeed("veh1"))
        # print("Acc =",traci.vehicle.getAcceleration("veh1"))
        # print("Pos =",traci.vehicle.getPosition("veh1"))
        traci.simulationStep()
        # simpla.update()
    print("Average CO2 Emission =",numpy.mean(CO2list))
    print("Average Fuel Consumption =",numpy.mean(Fuellist))
    write_excel_xlsx('result.xlsx','Feuil1',speedlist)
    traci.close()
    

