#pip install xpress

#import xpress
import xpress as xp
#import pandas
import pandas as pd
#import os
import os
#%%
#print working directory
print(os.getcwd())
#set working directory to a particular folder you want
os.chdir('C:/Users/bsts22/OneDrive - Loughborough University/Modules/Logistics Modelling and Operations Analytics/Coursework')
#check new working directory
print(os.getcwd())
#%% 
#no of potential dc locations
no_dc = 20
#no of user locations
demand_locations = 20
#maximum no of DCs
max_dcs = 20
#longest distance from the previous solution model
longest_distance = 14.1
#80% of the longest distance
max_distance= (longest_distance*0.8)
#%% 
#get distance matrix for DCs and users
distance = pd.read_excel('Location_Allocation.xlsx', index_col = None, nrows= no_dc, usecols='B:U', names = ['j'+ str(x + 1) for x in range(demand_locations)])
#get the capacities of the DCs
capacity = pd.read_excel('Location_Allocation.xlsx', index_col = None, nrows= no_dc, usecols='V')
#get the annual fixed operating fixed costs 
operating_cost = pd.read_excel('Location_Allocation.xlsx', index_col = None, nrows= no_dc, usecols='W')
#get the annual transportation demand
transport_demand = pd.read_excel('Location_Allocation.xlsx', index_col = None, skiprows = demand_locations, nrows= 1, usecols='B:U', names = ['r'+ str(x + 1) for x in range(demand_locations)])
#get the unit-distance transportation costs
unit_cost = pd.read_excel('Location_Allocation.xlsx', index_col = None, skiprows = demand_locations + 1, nrows= 1, usecols='B:U', names = ['w'+ str(x + 1) for x in range(demand_locations)])
#%% 
#unit_cost.head()
#Create the problem and give it a name
model1 = xp.problem(name = "A location allocation problem")
#%% defining the decision variables
#decision variable for each candidate DC location
z = {i : xp.var(vartype=xp.binary, name = 'z{0}'.format(i + 1)) for i in range(no_dc)}
#decision variable for each candidate DC location serving a particular user
x = {(i,j): xp.var(vartype=xp.binary, name='x{0}_{1}'.format(i + 1,j + 1)) for i in range(no_dc) for j in range(demand_locations)}
#%%
#adding decision variables to the problem
model1.addVariable(z,x)
#%%Constraints
#first constraint: each user must be served
constraint1 = [xp.Sum(x[i,j] for i in range(no_dc)) == 1 for j in range(demand_locations)]
#second constraint: the capacity of each DC cannot be exceeded
constraint2 = [xp.Sum(transport_demand.iloc[0,j] * x[i,j] for j in range(demand_locations)) <= capacity.iloc[i,0] * z[i] for i in range(no_dc)]
#third constraint: at most 20 DCs can be set up if required
constraint3 = [xp.Sum(z[i] for i in range(no_dc)) <= max_dcs]
#fourth constraint: the distance cannot exceed 80% of the longest distance
constraint4 = [distance.iloc[i,j] * x[i,j]  <= max_distance * z[i] for i in range(no_dc) for j in range(demand_locations)]
#%%
#adding constraints to the problem
model1.addConstraint(constraint1, constraint2, constraint3 ,constraint4)
#%%
#writing the objective function 
obj_function = [xp.Sum(operating_cost.iloc[i,0] * z[i] for i in range(no_dc)) + xp.Sum(unit_cost.iloc[0,j] * distance.iloc[i,j] * x[i,j] for i in range(no_dc) for j in range(demand_locations))]
#%%
#setting the objective function
model1.setObjective(obj_function, sense=xp.minimize)
#%%
#solving the problem
model1.solve()
#%%
#Writing the lp formulation
model1.write("location-lp", "lp")
#%% Print the solutions
print('--------------------------------------------------------------------------')
#Print the objective function
print("The annual operating cost is",model1.getObjVal())
#Get the optimal solution
print("The optimal solution is:")
for i in range(no_dc):
    if model1.getSolution(z[i]) >0:
        print(z[i],"=", model1.getSolution(z[i]))
        print("This DC supplies to:")
        for j in range(demand_locations):
            if model1.getSolution(x[i,j]) > 0:
                print(x[i,j], " = ", model1.getSolution(x[i,j]))
