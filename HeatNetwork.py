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
            if min(values.shape) == 2:  # For single-dimension parameters in Excel
                if values.shape[0] == 2:
                    for i in range(values.shape[1]):
                        data_dict[i + 1] = values[1][i]
                else:
                    for i in range(values.shape[0]):
                        data_dict[i + 1] = values[i][1]

            else:  # For two-dimension (matrix) parameters in Excel
                for i in range(values.shape[0]):
                    for j in range(values.shape[1]):
                        data_dict[(i + 1, j + 1)] = values[i][j]
            return data_dict

            ### Data importation ###


    InputData = "InputDataEnergySmallInstance.xlsx"
    card_V = read_excel_data(InputData, "Nodes")
    set_V = [i for i in range(1, card_V[0] + 1)]
    print('card_V', set_V)

    # Coords of the different nodes
    NodesCord = read_excel_data(InputData, "NodesCord")
    print("Cords: ", NodesCord)
    print(len(NodesCord))
    # Sommet source
    v_0 = read_excel_data(InputData, "SourceNum")

    # Pertes thermiques fixes (thetaijfix)
    thetafix = read_excel_data(InputData, "vfix(thetaijfix)")
    print("Thermic loss fix:", thetafix)

    # Pertes thermiques variables (thetaijvar)
    thetavar = read_excel_data(InputData, "vvar(thetaijvar)")
    print("Thermic loss", thetavar)

    # Fixed Unit Cost (?)
    fixedUnitCost = read_excel_data(InputData, "FixedUnitCost")
    print("Fixed Unit cost", fixedUnitCost)

    # Heat generation cost of source
    cheat = read_excel_data(InputData, "cheat(ciheat)")
    print("heat generation:", cheat)

    # Coûts variables dinvestissement
    cvar = read_excel_data(InputData, "cvar(cijvar)")
    print("Investissement", cvar)

    # Couts de maintenance
    cfix = read_excel_data(InputData, "com(cijom)")
    print("Cout de maintenance", cfix)

    # Revenus
    crev = read_excel_data(InputData, "crev(cijrev)")
    print("Revenu", crev)

    # Temps de fonctionnement de la production
    Tflh = read_excel_data(InputData, "Tflh(Tiflh)")
    print("Temps de fonction:", Tflh)

    # Betta
    betta = read_excel_data(InputData, "Betta")
    print("betta", betta[0])

    # Lambda
    Lambda = read_excel_data(InputData, "Lambda")
    print("Lambda", Lambda[0])

    # Alpha
    alpha = read_excel_data(InputData, "Alpha")
    print("alpha", alpha[0])

    # Edges demand peak
    d = read_excel_data(InputData, "EdgesDemandPeak(dij)")
    print("d:", d)

    # Edges demand annual
    D = read_excel_data(InputData, "EdgesDemandAnnual(Dij)")
    print("D:", D)

    # Capacité maximale des tuyaux
    Cmax = read_excel_data(InputData, "Cmax(cijmax)")
    print("Cmax", Cmax)

    # Capacité maximale de la source
    Qmax = read_excel_data(InputData, "SourceMaxCap(Qimax)")
    print("Qmax:", Qmax)

    # Pénalités de non satisfaction de la demande
    pumd = read_excel_data(InputData, "pumd(pijumd)")
    print("pumd", pumd)

    # longueurs des canaux
    l = {}
    for i in range(1, 9):
        for j in range(1, 9):
            l[(i, j)] = np.sqrt(((NodesCord[j, 1] - NodesCord[i, 1]) ** 2) + ((NodesCord[j, 2] - NodesCord[i, 2]) ** 2))
    print("lengths", l)

    # Here we create the heatProb variable which it will contain all the problem data
    heatProb = LpProblem("Heat_Network_Optimization", LpMinimize)

    # Create the decision variables
    x_var = LpVariable.dicts('x', (set_V, set_V), lowBound=0, upBound=1, cat='Binary')
    print("x_varia", x_var)
    Pin_var = LpVariable.dicts('Pin', (set_V, set_V), lowBound=0, cat='Continous')
    print("Pin_var", Pin_var)
    Pout_var = LpVariable.dicts('Pout', (set_V, set_V), lowBound=0, cat='Continous')
    print("Pout_var", Pout_var)
    # To make it easier to write I will write each constraint in a variable and the functions in the objective function

    for i in range(1, 9):
        # Tree structure constraint
        heatProb += lpSum([x_var[i][j] for j in range(1, 9)]) == abs(len(card_V)) - 1

    # Unidirectionality
    for i in range(1, 9):
        for j in range(1, 9):
            heatProb += x_var[i][j] + x_var[j][i] <= 1

    # The demand satisfaction
    for i in range(1, 9):
        for j in range(1, 9):

            heatProb += ((1 - l[i, j]) * thetavar[i, j] * Pin_var[i][j]) - Pout_var[i][j] - (
                        x_var[i][j] * (betta[0] * Lambda[0] * d[i, j] + l[i, j] * thetafix[i, j]
                                       ) == 0)

    # Flow equilibrium at each vertex, I supposed that the source node is in 1

    for i in range(1, 9):
        for j in range(1, 9):
            if i != j and j != 1:
                heatProb += lpSum([Pin_var[i][j]]) == lpSum([Pout_var[i][j]])
            if i == j:
                continue

    # Edge capacity constraint
    for i in range(1, 9):
        for j in range(1, 9):
            heatProb += Pin_var[i][j] <= x_var[i][j]*Cmax[i, j]

    # Source Structural constraint, I suppose that the source is in the first node
    for i in range(1, 9):
        heatProb += lpSum([x_var[i][1]]) == 0

    # Source's heat  generation capacity, i suppose that the source node is in 1
    for i in range(1, 9):
        heatProb += lpSum([Pin_var[1][i]]) <= Qmax[1]

    # Tour elimination, i suppose that source node is in 1
    for i in range(1, 9):
        if i != 1:
            heatProb += lpSum([x_var[i][j]] for j in range(1, 9)) >= 1

    #
