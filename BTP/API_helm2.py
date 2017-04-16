# This file is part of GridCal.
#
# GridCal is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# GridCal is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with GridCal.  If not, see <http://www.gnu.org/licenses/>.

import numpy as np


# from grid.CircuitOO import *
#from grid.CalculationEngine import *
from CalculationEngine import *
np.set_printoptions(precision=4)
grid = MultiCircuit()



#grid.load_file('case5.m')
#grid.load_file('case6ww.m')
#grid.load_file('case9.m')
#grid.load_file('case14.m')



#grid.load_file('case30.m')

#grid.load_file('case39.m')
#grid.load_file('case30pwl.m')
#grid.load_file('case69.xlsx')
#grid.load_file('case57.m')

#grid.load_file('case89pegase.m')
#grid.load_file('case118.m')

#grid.load_file('case145.m')

#grid.load_file('case300.m')
#grid.load_file('case1354pegase.m')
grid.load_file('case1888rte.m')
#grid.load_file('case2383wp.m')
#grid.load_file('case9241pegase.m')

grid.compile()
#compile()

#print('Ybus:\n', grid.circuits[0].power_flow_input.Ybus.todense())

options = PowerFlowOptions(SolverType.HELM, verbose=False, robust=False, tolerance=1e-9)
power_flow = PowerFlow(grid, options)
power_flow.run()

print('\n\n', grid.name)
print('\t|V|:', abs(grid.power_flow_results.voltage))

data = abs(grid.power_flow_results.voltage)
#print(data)
import io, json
f = open("data.txt", "w")
f.write( str(data)  )      # str() converts to string
f.close()

a = np.asarray(data)
np.savetxt("result.csv", a, delimiter=",")

#import xlsxwriter
    # Create an new Excel file and add a worksheet.
#workbook = xlsxwriter.Workbook('demo.xlsx')
#worksheet = workbook.add_worksheet()
#n = grid.power_flow_results.voltage.shape[0]
#print(n) 
#for x in abs(grid.power_flow_results.voltage)
#	worksheet.write(x)

#workbook.close()


#print('\t|angle|:', abs(grid.power_flow_results.angle(V)))
#print('\t|Sbranch|:', abs(grid.power_flow_results.Sbranch))
#print('\t|loading|:', abs(grid.power_flow_results.loading) * 100)
#print('\terr:', grid.power_flow_results.error)
#print('\tConv:', grid.power_flow_results.converged)


