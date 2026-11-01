import math
import colorama
import pandas as pd
from PIL.ImageQt import align8to32
from docx import Document
from mpl_toolkits import axes_grid1
from openpyxl import load_workbook
from colorama import Fore, Style
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import time
colorama.init()

df = pd.read_excel("Material Specifications.xlsx")
df2 = load_workbook("Material Specifications.xlsx")
sheet = df2["Material Specifications"]

print(df.head())
Material_choice_lining = int(input("Enter the selection number for lining material: "))
Material_choice_gasket = int(input("Enter the selection number for gasket material: "))
Material_choice_outer_shell = int(input("Enter the selection number for outer shell material: "))

#gasket
row_gask = 2 + Material_choice_gasket
#Lining
row = 2 + Material_choice_lining
#Mill shell
row_mill = 2 + Material_choice_outer_shell

cell_B4_ref = "B" + str(row) #Name of the lining
cell_D4_ref = "D" + str(row) #Density of lining
#Young's modulus
cell_L4_ref = "L" + str(row_gask) #young's modulus

#density ref
cell_D4 = str(sheet[cell_D4_ref].value)
#material name ref
cell_B4 = sheet[cell_B4_ref].value
#Young's module ref
cell_L4 = sheet[cell_L4_ref].value

cell_shell_density_ref = "D" + str(row_mill) #density of shell
cell_gask_density_ref = "D" + str(row_gask) #density of gasket

cell_mill_cost_ref = "N" + str(row_mill)
cell_lining_cost_ref = "N" + str(row)
cell_gask_cost_ref = "N" + str(row_gask)

cell_shell_density = sheet[cell_shell_density_ref].value
cell_gask_density = sheet[cell_gask_density_ref].value
cell_lining_density = float(cell_D4)

cell_mill_cost = sheet[cell_mill_cost_ref].value
cell_lining_cost = sheet[cell_lining_cost_ref].value
cell_gask_cost = sheet[cell_gask_cost_ref].value

#Name of the materials
cell_mill_name_ref = "B" + str(row_mill)
cell_linings_name_ref = "B" + str(row)
cell_rubber_name_ref = "B" + str(row_gask)

cell_mill_name = sheet[cell_mill_name_ref].value
cell_linings_name = sheet[cell_linings_name_ref].value
cell_rubber_name = sheet[cell_rubber_name_ref].value

print(Fore.YELLOW+"Enter the 30mm ball-mill radius(from mill center to the center of 30 mm ball) and"+Style.RESET_ALL)
print(Fore.YELLOW+"and the 20mm ball-mill radius(from mill center to the center of 20 mm ball) specifications "+Style.RESET_ALL)
inner_d = float(input("Enter value for Ball mill radius - 6mm radius value: "))
outer_d = float(input("Enter value for Ball mill radius - 15mm radius Value: "))
start_time = time.perf_counter()
print(Fore.YELLOW+ "Materials Selected: "+Style.RESET_ALL)
print("Metal Linings Material Name: ",Fore.BLUE + cell_linings_name + Style.RESET_ALL)
print("Rubber Linings Material Name: ",Fore.BLUE + cell_rubber_name + Style.RESET_ALL)
print("Mill Shell Material Name: ",Fore.BLUE + cell_mill_name + Style.RESET_ALL)
print("")


Bolt_list = []
i = int(1)

while i<=20:
    b = str("M")
    c = b + str(i)
    Bolt_list.append(c)

    i= (i+1)

#Kinematics

N_cID = float((60/(2*math.pi))*(math.sqrt((9.81/inner_d))))
N_cOD = float((60/(2*math.pi))*(math.sqrt((9.81/outer_d))))

N_c = float((N_cID + N_cOD)/2)
print(Fore.YELLOW+"The speed at which the charge centrifuge the inner linings of the Ball Mill"+Style.RESET_ALL)
print("Critical speed: ", N_c, "RPM")

N_operation = float(N_c * 0.68)
print(Fore.YELLOW+"Operational speed (68% of critical speed) required to result in cascade motion of the charge"+Style.RESET_ALL)
print("operational speed: ", N_operation, "RPM")
print("")
#Bolt pretension calculations
mass = float()
Fg = float(mass*9.81)*outer_d
w = float((N_operation*2*math.pi)/60)

F_c = mass*math.pow(w,2)
F_total = float(Fg + F_c)

Fsum_bolt = 2*F_total
num_bolts = int(20)

Fbolt = Fsum_bolt/num_bolts
yield_bolt = float(275) #MPa low carbon steel

A_tensile = float(Fbolt/(0.7*yield_bolt))
d_bolt = float(math.sqrt((4*A_tensile)/math.pi))
pitch = float(1.5)

D_bolt = d_bolt + 0.932*1.5

if D_bolt<=8:
    bolt_choice_OD =8
else:
    while i<= D_bolt:
        bolt_choice = Bolt_list[i]
        bolt_choice_OD = int(i-1)
        i=+ i

#table with bolts, pre torque and k factors


K = float(4)
T = float(8.2)
print(Fore.YELLOW+"The Bolt type with the preload that will counter the operational forces to result in complete seal"+Style.RESET_ALL)
print("Bolt type: M",bolt_choice_OD, "x 1.5")
print("")
P_required = T/(K*float(bolt_choice_OD))
print("Bolt Pre-Load: ",P_required*1000, "N")
print("")
#Young's modulus gasket
cell_gas_F4_ref = "F" + str(row_gask)
cell_gas_F4 = sheet[cell_gas_F4_ref].value
#material type
cell_gas_B4_ref = "B" + str(row_gask)
cell_gas_B4 = sheet[cell_gas_B4_ref].value

print("Young's modulus of ",cell_gas_B4,": ", cell_gas_F4, "MPa")

A_gasket = float(0.0294)
A_lid = float(0.0735)

P_lid = (P_required*num_bolts)/A_gasket - Fsum_bolt/A_lid

Lid_stress = (P_required*num_bolts)/A_gasket

#strain

strain = Lid_stress/cell_gas_F4
Original_gasket_length = float(0.007)
#deformation along the central axis
deformatin_gask = float(Original_gasket_length/strain)
print(Fore.YELLOW+"The Strain and directional deformation across the central axis of the axis after pre tension and operational forces"+Style.RESET_ALL)
print("Gasket Deformation: ",deformatin_gask,"mm")
print("Gasket Strain: ",strain)
print("")
print(Fore.YELLOW+"Contact Mechanics (calculated Contact presurre and stress)"+Style.RESET_ALL)
print("Contact Pressure: ",P_lid,"MPa")
print("stress :",Lid_stress,"MPa")
#Motor selection

mill_ring = float(550)
mill_teeth = int(72)

module_gear = math.ceil(float(mill_ring/mill_teeth))
#use module to get the table from excel file (Z = teeth)

Zmotor = int(18)

Motor_dia = mill_ring*(Zmotor/mill_teeth)
N_motor = (mill_teeth/Zmotor)*N_operation

Bulk_volume = float(0.007)
density_balls = int(7850)
mass_bulk = Bulk_volume*density_balls

Tmill = float(mass_bulk*9.81*0.1325*math.sin(45))
Tout = float((mill_teeth/Zmotor)*Tmill)
Tout_desired = Tout*2
print(Fore.YELLOW+"Calculated power and rotational speed for the motor"+Style.RESET_ALL)
w_motor = float((2*math.pi*N_operation)/60)
print("Motor rotational speed",w_motor,"rad/s")
Pow_out = Tout_desired*w_motor
print("Motor Power Output: ",Pow_out, "W")

#thermal control

Qtotal = float((0.9)*990)
Aext = 0.42
Tamb = 295.15 #22 in degrees celsius
hconv = 14 #convection at 61 RPM
e = 0.85 #stainless steel
boltzman_constant = 5.6*math.pow(10,-8)


a = e*boltzman_constant*Aext
b = hconv*Aext
c = (hconv*Aext*Tamb+e*boltzman_constant*Aext*math.pow(Tamb,4))+ Qtotal

#print(a/a)
#print(b/a)
#print(c/a)
print("")
Tliner_uncooled_kelvin = float(math.cbrt((b/a)/4))
Tliner_uncooled = float(Tliner_uncooled_kelvin - 273)
print(Fore.YELLOW+"The temperature of the liners before pneumatic air cooling"+Style.RESET_ALL)
print("Uncooled liner: ",Tliner_uncooled, "Degree Celsius")
print("")
#Cooled Liner

h_int = 56.5
h_ext = 22.5
m_mass_charge = 0.03546
Cp_air = 1005

#ext
m = h_ext*0.42
l = h_ext*0.42*22

#int

o = h_int*0.35
p = m_mass_charge*Cp_air*2

#o*Tliner_cooled -(o+p)*T_bulk = p*22*(-1)
#(m+o)*Tliner_cooled - o*T_bulk = l + 990

T_bulk = (l+990+((m+o)*(p*22))/o)/(((m+o)*(o+p)-o*o)/o)
Tliner_cooled = (((o+p)*T_bulk)/o) - (p*22)/o
print(Fore.YELLOW+"The temperature of the liners after pneumatic air cooling"+Style.RESET_ALL)
print("Bulked cooled liner: ",Tliner_cooled, "Degrees Celsius")
print("Cooled liner: ",T_bulk, "Degrees Celsius")
print("")
#Volumes
Vol_mill_shell = float(0.00401)
Vol_rubber_linings = float(0.000258)
Vol_metal_linings = float(0.00295)
Vol_ring_gear = float(0.00436)
Vol_gasket = float(0.000206)

#mass and Weight of the outershell, linings and material
m_mill_shell = float(Vol_mill_shell*cell_shell_density)
m_rubber_linings = float(Vol_rubber_linings*cell_gask_density)
m_metal_linings = float(Vol_metal_linings*cell_lining_density)
m_ring_gear = float(Vol_ring_gear*cell_shell_density)
m_gasket = float(Vol_gasket*cell_gask_density)
print(Fore.YELLOW+"Components Masses"+Style.RESET_ALL)
print("Mill shell mass: ",m_mill_shell, "Kg")
print("Rubber linings mass: ",m_rubber_linings, "Kg")
print("Metal linings mass: ",m_metal_linings, "Kg")
print("Ring gear mass: ",m_ring_gear, "Kg")
print("Gasket mass: ",m_gasket, "Kg")
print("")
W_mill_shell = float(m_mill_shell*9.81)
W_rubber_linings = float(m_rubber_linings*9.81)
W_metal_linings = float(m_metal_linings*9.81)
W_ring_gear = float(m_ring_gear*9.81)
W_gasket = float(m_gasket*9.81)
Total_weight = W_gasket+W_metal_linings+W_ring_gear+W_mill_shell+W_rubber_linings

print(Fore.YELLOW+"Components Weights"+Style.RESET_ALL)
print("Mill shell weight: ",W_mill_shell, "N")
print("Rubber linings weight: ",W_rubber_linings, "N")
print("Metal linings weight: ",W_metal_linings, "N")
print("Ring gear weight: ",W_ring_gear, "N")
print("Gasket weight: ",W_gasket, "N")
print("Total weight: ", Total_weight, "N")
print("")
#cost using volumes (outer shell, linings, gasket)
print(Fore.YELLOW+"Components Costs"+Style.RESET_ALL)
cost_gask = m_gasket*cell_gask_cost
cost_metal_lining = m_metal_linings*cell_lining_cost
cost_mill_shell = m_mill_shell*cell_mill_cost
cost_rubber_linings = m_rubber_linings*cell_gask_cost
cost_ring_gear = m_ring_gear*cell_mill_cost
Total_cost = cost_gask+cost_metal_lining+cost_mill_shell+cost_rubber_linings+cost_ring_gear

print("Gasket cost: R",cost_gask)
print("Metal lining cost: R",cost_metal_lining)
print("Mill shell cost: R",cost_mill_shell)
print("Rubber linings cost: R",cost_rubber_linings)
print("Ring gear cost: R",cost_ring_gear)
print("Total cost: R",Total_cost)

#motor selection

#justifications of the materials (max service, hardness, performance)
print("")
print("Results interpretations")
print("")
if float(deformatin_gask)<float(1):
    print(Fore.GREEN+"The deformation",deformatin_gask,"mm being approximately 0 concludes that the gasket material"+Style.RESET_ALL)
    print(Fore.GREEN+"and dimensions choices result in the complete seal during operation reducing any leaks"+Style.RESET_ALL)
else:
    print(Fore.RED + "The deformation", deformatin_gask, "mm being more than 0 concludes that the gasket material"+Style.RESET_ALL)
    print(Fore.RED + "and dimensions choices does not result in the complete seal during operation causing leaks"+Style.RESET_ALL)
print("")
if Lid_stress<=250:
    print(Fore.GREEN + "The internal stress on the contact area of the lid and gasket being lower than 250 MPa guarantees the structural integrity of the bolts" + Style.RESET_ALL)
else:
    print(
        Fore.RED + "The internal stress on the contact area of the lid and gasket being more than 250 MPa guarantees the failure of structural integrity of the bolts" + Style.RESET_ALL)
print("")
cell_mill_yield_ref = "F" + str(row_mill)
cell_mill_yield = sheet[cell_mill_yield_ref].value
if Lid_stress<=float(cell_mill_yield):
    print(Fore.GREEN + "The internal stress on the contact area of the lid and gasket being lower than" + str(cell_mill_yield) + "gurantees the structural integrity of the bolts" + Style.RESET_ALL)
else:
    print(
        Fore.RED + "The internal stress on the contact area of the lid and gasket being more than" + str(cell_mill_yield) +" gurantees the failure of structural integrity of the mill shell body" + Style.RESET_ALL)
print("")
cell_lining_max_temp_ref = "Q" + str(row)
cell_lining_max_temp = sheet[cell_lining_max_temp_ref].value
Temp_cond = float(cell_lining_max_temp-float(300))
if Tliner_uncooled>Temp_cond:
    print(Fore.RED + "Max service temperature for "+cell_linings_name+" being lower than or close to the uncooled temperature will result in corrosion in long term of operations" + Style.RESET_ALL)
else:
    print(Fore.GREEN + "Max service temperature for " + cell_linings_name + " being more than the uncooled temperature will result not corrosion in long term of operations" + Style.RESET_ALL)

cell_lining_hardness_ref = "O" + str(row)
cell_lining_hardness = sheet[cell_lining_hardness_ref].value
print("")
print("Hardness: ",)
print(Fore.GREEN + str(cell_linings_name) +" Lining Hardness: " + str(cell_lining_hardness) +"HB"+ Style.RESET_ALL)

# End the timer
end_time = time.perf_counter()
execution_time = end_time - start_time
print("")
print(f"Execution time: {execution_time:.6f} seconds")

