import distance as distance
from pulp import *
import numpy as np
import pandas as pd
from scipy import spatial
from itertools import combinations
import math as mt

if __name__ == "__main__":


    # Lets read the excel file which contain all the necessary data to start our OR

    # Input Data Preparation #
    InputData = "InputDataEnergySmallInstance.xlsx"


    # Input Data Preparation #
    def read_excel_data(filename, sheet_name):
        data = pd.read_excel(filename, sheet_name=sheet_name, header=None)
        values = data.values
        if min(values.shape) == 1:  # This If is to make the code insensitive to column-wise or row-wise expression #
            if values.shape[0] == 1:
                values = values.tolist()
            else:
                values = values.transpose()
                values = values.tolist()
            return values[0]        
        else:
            data_dict = {}
            if min(values.shape) == 1:  # For single-dimension parameters in Excel
                if values.shape[0] == 2:
                    for i in range(values.shape[1]):
                        data_dict[i+1] = values[1][i]
                else:
                    for i in range(values.shape[0]):
                        data_dict[i+1] = values[i][1]
             
            else:  # For two-dimension (matrix) parameters in Excel
                for i in range(values.shape[0]):
                    for j in range(values.shape[1]):
                        data_dict[(i+1, j+1)] = values[i][j]
            return data_dict

            ### Data importation ###


    ## Date will be imported in a list if they are one dimension or in a dictionary if they are in two dimensions
    ## with tuples as keys

    v0 =  read_excel_data(InputData, "SourceNum")  # This the Cord of the heat source

    # The excel file that we want to read
    InputData = "InputDataEnergySmallInstance.xlsx"

    
    card_V = read_excel_data(InputData, "Nodes")
    set_V = [i for i in range(1, card_V[0] + 1)]
    #print('card_V', set_V)

    # Coords of the different nodes
    NodesCord = read_excel_data(InputData, "NodesCord")
    print(NodesCord)
    #print("Cords: ", NodesCord)
    #print(len(NodesCord))

    # Sommet source
    v_0 = read_excel_data(InputData, "SourceNum")
    v_0 = v_0[0]

    # Pertes thermiques fixes (thetaijfix)
    thetafix = read_excel_data(InputData, "vfix(thetaijfix)")
    # print("Thermic loss fix:", thetafix)

    # Pertes thermiques variables (thetaijvar)
    thetavar = read_excel_data(InputData, "vvar(thetaijvar)")
    #print("Thermic loss", thetavar)

    # Fixed Unit Cost (?)
    cfix = read_excel_data(InputData, "FixedUnitCost")
    #print("Fixed Unit cost", fixedUnitCost)

    # Heat generation cost of source
    cheat = read_excel_data(InputData, "cheat(ciheat)")
    #print("heat generation:", cheat)

    # Coûts variables dinvestissement
    cvar = read_excel_data(InputData, "cvar(cijvar)")
    #print("Investissement", cvar)

    # Couts de maintenance
    com = read_excel_data(InputData, "com(cijom)")
    #print("Cout de maintenance", com)

    # Revenus
    crev = read_excel_data(InputData, "crev(cijrev)")
    #print("Revenu", crev)

    # Temps de fonctionnement de la production
    Tflh = read_excel_data(InputData, "Tflh(Tiflh)")
    #print("Temps de fonction:", Tflh[0])

    # Betta
    betta = read_excel_data(InputData, "Betta")
    #print("betta", betta[0])

    # Lambda
    Lambda = read_excel_data(InputData, "Lambda")
    #print("Lambda", Lambda[0])

    # Alpha
    alpha = read_excel_data(InputData, "Alpha")
    #print("alpha", alpha[0])

    # Edges demand peak
    d = read_excel_data(InputData, "EdgesDemandPeak(dij)")
    #print("d:", d)

    # Edges demand annual
    D = read_excel_data(InputData, "EdgesDemandAnnual(Dij)")
    #print("D:", D)

    # Capacité maximale des tuyaux
    Cmax = read_excel_data(InputData, "Cmax(cijmax)")
    #print("Cmax", Cmax)

    # Capacité maximale de la source
    Qmax = read_excel_data(InputData, "SourceMaxCap(Qimax)")
    #print("Qmax:", Qmax)

    # Pénalités de non satisfaction de la demande
    pumd = read_excel_data(InputData, "pumd(pijumd)")
    #print("pumd", pumd)

    # longueurs des canaux
    l = {}
    for i in set_V:
        for j in set_V:
            l[(i, j)] = np.sqrt(((NodesCord[j, 1] - NodesCord[i, 1]) ** 2) + ((NodesCord[j, 2] - NodesCord[i, 2]) ** 2))
    #print("lengths", l)

    # Here we create the heatProb variable which it will contain all the problem data
    heatProb = LpProblem("Heat_Network_Optimization", LpMinimize)

    # Create the decision variables
    x_var = LpVariable.dicts('x', (set_V, set_V), lowBound=0, upBound=1, cat='Binary')
    print("x_varia", x_var)
    Pin_var = LpVariable.dicts('Pin', (set_V, set_V), lowBound=0, cat='Continous')
    print("Pin_var", Pin_var)
    Pout_var = LpVariable.dicts('Pout', (set_V, set_V), lowBound=0, cat='Continous')
    print("Pout_var", Pout_var)

    # Now i will write the objective function, to minimize the cosst
    # To make it easier to write I will write each constraint in a variable and the functions in the objective function
    # and seperate the functions in the objections into seperate variables so it can be easier to understand

    # Here are the functions in the objectif functions

    # Heat Generation cost
    heat_generation_cost = lpSum((Tflh[0] * cheat[v_0] / betta[0]) * Pin_var[4][i] for i in set_V)

    # Revenue
    revenue_lis = []
    for i in set_V:
        for j in set_V:
            revenue_lis.append((-Lambda[0]) * x_var[i][j] * cvar[i, j] * D[i, j])

    revenue = lpSum(revenue_lis)

    # Maintenance cost
    maintenance_cost_lis = []
    for i in set_V:
        for j in set_V:
            maintenance_cost_lis.append(x_var[i][j] * com[i, j] * l[i, j])
    maintenance_cost = lpSum(maintenance_cost_lis)

    # Fixed Investment Cost
    fixed_investment_cost_lis = []
    for i in set_V:
        for j in set_V:
            fixed_investment_cost_lis.append(alpha[0] * x_var[i][j] * cfix[0] * l[i, j])
    fixed_investement_cost = lpSum(fixed_investment_cost_lis)

    # Variable Investment Cost
    variable_Inv_cost_lis = []
    for i in set_V:
        for j in set_V:
            variable_Inv_cost_lis.append(alpha[0] * cvar[i, j] * l[i, j] *
                                         Pin_var[i][j])

    variable_investment_cost = lpSum(variable_Inv_cost_lis)


    # Unmet Demand Penalty
    unmet_demand_penalty_lis = []
    for i in set_V:
        for j in set_V:
            unmet_demand_penalty_lis.append(0.5*(1-x_var[i][j]-x_var[j][i])*D[i, j]*pumd[i, j])
    unmet_demand_penalty = lpSum(unmet_demand_penalty_lis)
    # After writing the different terms of the objective function, I will add it in the prob data

    heatProb += (heat_generation_cost + revenue + maintenance_cost + fixed_investement_cost + variable_investment_cost
                    + unmet_demand_penalty)

    ## Here are the constraints of the Heat Network Problem

    for i in set_V:
        # Tree structure constraint
        heatProb += lpSum([x_var[i][j] for j in set_V]) == abs(len(card_V)) - 1

    # Unidirectionality
    for i in set_V:
        for j in set_V:
            heatProb += x_var[i][j] + x_var[j][i] <= 1

    # The demand satisfaction
    for i in set_V:
        for j in set_V:
            heatProb += (((1 - l[i, j]) * thetavar[i, j] * Pin_var[i][j]) - Pout_var[i][j] -
                    x_var[i][j] * (betta[0] * Lambda[0] * d[i, j] + l[i, j] * thetafix[i, j]) == 0)

    # Flow equilibrium at each vertex,

    for j in set_V:
        if j != v_0:
            lis_pin = []
            lis_pout = []
            for i in set_V:
                if i != j:
                    lis_pin.append(Pin_var[i][j])
                    lis_pout.append(Pout_var[i][j])
                if i == j:
                    continue
            heatProb += lpSum(lis_pin) == lpSum(lis_pout)

    # Edge capacity constraint
    for i in set_V:
        for j in set_V:
            heatProb += Pin_var[i][j] <= x_var[i][j] * Cmax[i, j]

    # Source Structural constraint
    struc_constraint_lis = []
    for i in set_V:
        if i != v_0:
            struc_constraint_lis.append(x_var[i][v_0])
    heatProb += lpSum(struc_constraint_lis) == 0

    # Source's heat  generation capacity
    heat_capacity_lis = []
    for j in set_V:
        if j != v_0:
            heat_capacity_lis.append(Pin_var[v_0][j])
    heatProb += lpSum(heat_capacity_lis) <= Qmax[v_0]

    # Tour elimination,
   
    for i in set_V:
        if i != v0:
            tour_lis = []
            for j in set_V:
                tour_lis.append(x_var[i][j])
    heatProb += lpSum(tour_lis) >= 1

    # I will write all the problem data in a file

    heatProb.writeLP("HeatProblem.lp")

    # Now that all the date are putted in we will solve this problem

    heatProb.solve()

    # The status of the problem, if the solution is Infeasible, Optimal, Unbounded etc

    print("Status",LpStatus[heatProb.status])
    print("The solution for the objective function is:", value(heatProb.objective))