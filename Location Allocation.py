#pip install xpress

#import xpress
import xpress as xp
#import pandas
import pandas as pd
#import os
import os
#%%
#getting current working directory
print(os.getcwd())
#setting working directory to a particular folder
os.chdir('C:/Users/bsts22/OneDrive - Loughborough University/Modules/Logistics Modelling and Operations Analytics/Coursework')
#checking new working directory
print(os.getcwd())
#%% 
#no of potential dc locations
no_dc = 20
#no of user locations
demand_locations = 20
#maximum no of DCs
max_dcs = 3
#%% 
#getting distance matrix for DCs and users
distance = pd.read_excel('Location_Allocation.xlsx', index_col = None, nrows= no_dc, usecols='B:U', names = ['j'+ str(x + 1) for x in range(demand_locations)])
#getting the capacities of the DCs
capacity = pd.read_excel('Location_Allocation.xlsx', index_col = None, nrows= no_dc, usecols='V')
#getting the annual fixed operating fixed costs 
operating_cost = pd.read_excel('Location_Allocation.xlsx', index_col = None, nrows= no_dc, usecols='W')
#getting the annual transportation demand
transport_demand = pd.read_excel('Location_Allocation.xlsx', index_col = None, skiprows = demand_locations, nrows= 1, usecols='B:U', names = ['r'+ str(x + 1) for x in range(demand_locations)])
#getting the unit-distance transportation costs
unit_cost = pd.read_excel('Location_Allocation.xlsx', index_col = None, skiprows = demand_locations + 1, nrows= 1, usecols='B:U', names = ['w'+ str(x + 1) for x in range(demand_locations)])
#%% 
#Creating the problem
model1 = xp.problem(name = "A location allocation problem")
#%% define the decision variables
#decision variable for each candidate DC location
z = {i : xp.var(vartype=xp.binary, name = 'z{0}'.format(i + 1)) for i in range(no_dc)}
#decision variable for each candidate DC location serving a particular user
x = {(i,j): xp.var(vartype=xp.binary, name='x{0}_{1}'.format(i + 1,j + 1)) for i in range(no_dc) for j in range(demand_locations)}
#%%
#add the decision variables to the problem
model1.addVariable(z,x)
#%% write the constraints
#first constraint: each user must be served
constraint1 = [xp.Sum(x[i,j] for i in range(no_dc)) == 1 for j in range(demand_locations)]
#second constraint: the capacity of each DC cannot be exceeded
constraint2 = [xp.Sum(transport_demand.iloc[0,j] * x[i,j] for j in range(demand_locations)) <= capacity.iloc[i,0] * z[i] for i in range(no_dc)]
#third constraint: at most 3 DCs can be set up
constraint3 = [xp.Sum(z[i] for i in range(no_dc)) <= max_dcs]
#%%
#adding the constraints to the problem
model1.addConstraint(constraint1, constraint2, constraint3)
#%%
#writing the objective function 
obj_function = [xp.Sum(operating_cost.iloc[i,0] * z[i] for i in range(no_dc)) + xp.Sum(unit_cost.iloc[0,j] * distance.iloc[i,j] * x[i,j] for i in range(no_dc) for j in range(demand_locations))]
#%%
#setting the objective function
model1.setObjective(obj_function, sense=xp.minimize)
#%%
#Call the optimizer to solve the problem
model1.solve()
#%%
#Write the lp formulation
model1.write("location-lp", "lp")
#%% Print the solutions
print('--------------------------------------------------------------------------')
#Print the objective function
print("The annual operating cost is",model1.getObjVal())
#Get the optimal solution
print("The optimal solution is as follows:")
for i in range(no_dc):
    if model1.getSolution(z[i]) >0:
        print(z[i],"=", model1.getSolution(z[i]))
        print("This DC supplies to:")
        for j in range(demand_locations):
            if model1.getSolution(x[i,j]) > 0:
                print(x[i,j], " = ", model1.getSolution(x[i,j]))