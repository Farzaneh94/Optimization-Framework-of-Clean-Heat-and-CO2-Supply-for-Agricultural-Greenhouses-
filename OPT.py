#Optimization Framework of Clean Heat and CO2 Supply
# for Agricultural Greenhouses Exploiting Industrial Symbiosis
#------------------------------------------------------------------------------------------------
# Developers: Farzaneh Rezaei, Vanessa Burg, Stephan Pfister, Stefanie Hellweg and Ramin Roshandel
# Department of Energy Engineering, Sharif University of Technology, Tehran, Iran
# Institute of Environmental Engineering, ETH Zürich, Zürich, Switzerland
# August 2024
#________________________________________________________________________________________________
#This optimization code uses an EXCEL data file including Heat and CO2 supplieres,
#suitable land area, heat and lighting demand for possible tomato, cucumber and lettuce greenhouses
#and outputs the optimal heat and CO2 exchanges between supplieres and greenhouses to minimize
#total annulized cost of IS network
#for simplicity the effects of (1) Network Scale on Investment Costs is dropped in this version.
#__________________________________________________________________________________________________

# Calling libraries
import pulp
import numpy
import pandas

# Defining Optimization Model Type
model = pulp.LpProblem("Cost_minimising", pulp.LpMinimize)

# Defining of Set Lists
H_supplier=list(range(1,190))       # Waste heat suppliers
CO2_supplier=list(range(1,38))      # Waste CO2 suppliers
land=list(range(1,218))             # Suitable lands
croptype=list(range(1,4))           # 1: tomato   2: cucumber   3: lettuce
H_tech=list(range(1,3))             # 1: with ORC    2: without ORC

#******************************
biogasplant=list(range(1,154))      # Biogas Plants
MSWI= list(range(154,184))          # Municipal Solid Waste Incinerators
cement= list(range(184,190))        # Cement Production Plants



# Reading data from excel
df1 = pandas.read_excel('NEWINPUT.xlsx', sheet_name= 'Sheet1',usecols='B:C')
A = numpy.array(df1.loc[:,"Area (ha)"])   # total area of each suitable land

#*********************************************************************************************
df2 = pandas.read_excel('NEWINPUT.xlsx', sheet_name= 'Sheet2',usecols='A:B')
E = numpy.array(df2.loc[:,"E (MW)"])      # power of suppliers   (MW)

#*********************************************************************************************
df3 = pandas.read_excel('NEWINPUT.xlsx', sheet_name= 'Sheet3',usecols='A:B')
PCO2 = numpy.array(df3.loc[:,"CO2 (ton/y)"])      # CO2 of suppliers   (ton per year)

#*********************************************************************************************
df4 = pandas.read_excel('NEWINPUT.xlsx', sheet_name= 'Sheet4',usecols='B:HJ',header=1)
d=numpy.array(df4)                  # distance km
d_H=d
#*********************************************************************************************
df5 = pandas.read_excel('NEWINPUT.xlsx', sheet_name= 'Sheet5',usecols='C:E',header=0)
PD=numpy.array(df5)                  # energy required for crops   (MW per hec)

#*********************************************************************************************
df6 = pandas.read_excel('NEWINPUT.xlsx', sheet_name= 'Sheet6',usecols='C:E',header=0)
LD=(numpy.array(df6) )*10                 # lighting required for crops   (MWh per hec)

#*********************************************************************************************
df7 = pandas.read_excel('NEWINPUT.xlsx', sheet_name= 'Sheet7',usecols='B:HJ',header=1)
dc=numpy.array(df7)                  # distance km
d_CO2=dc
# Optimization Model Parameters:
# total greenhouse area demand for tomato and cucumber and lettuce (ha)
TA={
    "1" : 72,
    "2" : 30,
    "3" : 93,
}

#CO2 demand for crops   (ton per ha)
CD={
    "1" : 160,
    "2" : 200,
    "3" : 150,
}


#In this version the fixed investment cost is considered
INV_H=128000   #CHF per km per MW
OPH =0.0292    # OPERATION COST OF PIPING : CHF per km.MWh
h1=5000        # hours that greenhouse needs heat
hlc=0.01       # HEAT LOSS COEFFICIENT : %/km

#CHF per km per ton CO2 Pipeline
INV_CO2=700 # CHF/km.tone
OP_CO2=0.04*INV_CO2


# Greenhouse structure
INV_G=1160000  #CHF per ha

INV_light=1100000 #CHF per ha


#In this version the fixed investment cost is considered

INV_ORC=7423500 #CHF/MW
OM_ORC=15 #CHF/MWh



h2=8000        # hours that ORC generates elec
Pr_el=150      # ELECTRICITY PRICE : doller per MWh
R_HE=4         # heat per elec RATIO

dis=0.105   # discount rate
n=20        # lifetime of system
CRF=(dis*(1+dis)**n)/((1+dis)**n-1)

BigM=100000000



# Definition of continuous variables

x = pulp.LpVariable.dicts("greenhouse area ha",((i,z,j,l,k) for i in H_supplier
                                                          for j in land
                                                          for l in croptype
                                                          for k in H_tech
                                                          for z in CO2_supplier),lowBound=0, cat='Continuous')


# Definition of binary variables

Y = pulp.LpVariable.dicts("Pathway",((i,z,j,l,k) for i in H_supplier
                                                          for j in land
                                                          for l in croptype
                                                          for k in H_tech
                                                          for z in CO2_supplier),lowBound=0, cat='Binary')


# Objective Function

INCOST_H=pulp.lpSum((INV_H * d_H[i-1][j-1] * x[i, z , j, l, k] * PD[j-1] [l-1]) for i in H_supplier for j in land for l in croptype for k in H_tech for z in CO2_supplier)
INCOST_ORC=pulp.lpSum((INV_ORC *  x[i, z , j, l, 1] * PD[j-1] [l-1]/R_HE)for i in H_supplier for j in land for l in croptype for z in CO2_supplier)
INCOST_CO2=pulp.lpSum((INV_CO2 * d_CO2[z-1][j-1] * x[i, z , j, l, k] * CD[str(l)]) for i in H_supplier for j in land for l in croptype for k in H_tech for z in CO2_supplier)
INCOST_STRUCTURE=pulp.lpSum((INV_G * x[i, z , j, l, k]) for i in H_supplier for j in land for l in croptype for k in H_tech for z in CO2_supplier)
INCOST_LIGHT=pulp.lpSum((INV_light * x[i, z , j, l, k]) for i in H_supplier for j in land for l in croptype for k in H_tech for z in CO2_supplier)

OPERATION_HEAT=pulp.lpSum((OPH * d_H[i-1][j-1] * h1 * x[i, z , j, l, k] * PD[j-1] [l-1])  for i in H_supplier for j in land for l in croptype for k in H_tech for z in CO2_supplier)
OPERATION_ORC=pulp.lpSum((OM_ORC * h2 * x[i, z , j, l, 1] * PD[j-1] [l-1]/R_HE) for i in H_supplier for j in land for l in croptype for z in CO2_supplier)
OPERATION_CO2=pulp.lpSum((OP_CO2 * d_CO2[z-1][j-1] * x[i, z , j, l, k] * CD[str(l)] )  for i in H_supplier for j in land for l in croptype for k in H_tech for z in CO2_supplier)
OPERATION_LIGHT=pulp.lpSum((Pr_el *LD[j-1][l-1]*x[i, z , j, l, k]) for i in H_supplier for j in land for l in croptype for k in H_tech for z in CO2_supplier)

INCOME=pulp.lpSum((Pr_el * h2 * x[i, z , j, l, 1] * PD[j-1] [l-1]/R_HE) for i in H_supplier for j in land for l in croptype for z in CO2_supplier)


model += (INCOST_H+INCOST_ORC+INCOST_CO2+INCOST_STRUCTURE+INCOST_LIGHT)*CRF+(OPERATION_HEAT+OPERATION_ORC+OPERATION_CO2+OPERATION_LIGHT)-INCOME


# Constraints

# meeting_demand
for l in croptype:
    model += pulp.lpSum([x[i, z , j, l, k] for i in H_supplier for j in land for k in H_tech for z in CO2_supplier]) >= (TA[str(l)] )

# proper land area
for j in land:
    model += pulp.lpSum([x[i, z , j, l, k] for i in H_supplier for l in croptype for k in H_tech for z in CO2_supplier]) <= A[j-1]

# heat supply restriction

for i in H_supplier:
    model += pulp.lpSum([((1/0.35) * x[i, z , j, l, 1] + x[i, z , j, l, 2]) * PD[j-1] [l-1]* (1 + hlc * d_H[i - 1][j - 1]) for j in land for l in croptype for z in CO2_supplier]) <= E[i-1]

# CO2 supply restriction
for z in CO2_supplier:
    model += pulp.lpSum([x[i, z , j, l, k] * CD[str(l)]  for i in H_supplier for j in land for l in croptype for k in H_tech]) <= PCO2[z-1]

# Path tech
for i in H_supplier:
  for z in CO2_supplier:
      for j in land:
           for l in croptype:
                   model += (Y[i, z , j, l, 1] + Y[i, z , j, l, 2]) <= 1


# all pathways
for i in H_supplier:
  for z in CO2_supplier:
      for j in land:
           for l in croptype:
              for k in H_tech:
                        model += x[i, z , j, l, k]  <= Y[i, z , j, l, k] * BigM

#feasible 1  We can not have ORC in BIOgas plants
for i in biogasplant:
  for z in CO2_supplier:
      for j in land:
           for l in croptype:
                 model += Y[i, z , j, l, 1]  == 0

#feasible 2 We can not have ORC in MSWI
for i in MSWI:
  for z in CO2_supplier:
      for j in land:
           for l in croptype:
            model += Y[i, z , j, l, 1]  == 0




#print(model.solve())
print(model.solve())
print(pulp.LpStatus[model.status])

# model output
# all pathways
for i in H_supplier:
  for z in CO2_supplier:
      for j in land:
           for l in croptype:
              for k in H_tech:
                       if x[i, z , j, l, k].varValue > 0:
                            print("x","[",i, z , j, l, k,"]", "=" , x[i, z , j, l, k].varValue)




print("INCOST_H","=", INCOST_H.value())
print("INCOST_ORC","=",INCOST_ORC.value())
print("INCOST_CO2","=",INCOST_CO2.value())
print("INCOST_STRUCTURE","=",INCOST_STRUCTURE.value())
print("INCOST_LIGHT","=",INCOST_LIGHT.value())
print("OPERATION_HEAT","=",OPERATION_HEAT.value())
print("OPERATION_ORC","=",OPERATION_ORC.value())
print("OPERATION_CO2","=",OPERATION_CO2.value())
print("OPERATION_LIGHT","=",OPERATION_LIGHT.value())
print("INCOME","=",INCOME.value())

print("objective function","=",model.objective.value())