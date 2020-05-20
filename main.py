# Import PuLP modeler functions
from pulp import *
import pandas as pd
import numpy as np
from itertools import combinations
import math as mt


if __name__ == "__main__":
    
    # Read excel file #
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
          if min(values.shape) == 2:  # For single-dimension parameters in Excel
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
 
  ###### Main program ######
  
  ### Data importation ###
   
    card_V = read_excel_data(InputData, "Nodes")
    set_V = [i for i in range(1, card_V + 1)]
    
    # Sommet source
    v_0 = read_excel_data(InputData, "SourceNum")

    # Pertes thermiques fixes (thetaijfix)
    thetafix = read_excel(InputData, "vfix(thetaijfix)")

    # Pertes thermiques variables (thetaijvar)
    thetavar = read_excel(InputData, "vvar(thetaijvar)")

    # Fixed Unit Cost (?)
    fixedUnitCost = read_excel(InputData, "FixedUnitCost")
    
    # Heat generation cost of source
    cheat= read_excel_data(InputData, "cheat(ciheat)")
    
    # Coûts variables dinvestissement    
    cvar = read_excel(InputData, "cvar(cijvar)")
    
    # Couts de maintenance
    cfix = read_excel(InputData, "com(cijom)")
    
    # Revenus
    crev = read_excel(InputData, "crev(cijrev)")
    
    # Temps de fonctionnement de la production
    Tflh = read_excel(InputData, "Tflh(Tiflh)")
    
   #Betta 
   betta = read_excel(InputData, "Betta") 

   # Lambda
   lambda = read_excel(InputData, "Lambda")
   
   # Alpha 
   alpha = read_excel(InputData, "Alpha")
   
   # Edges demand peak
   d = read_excel(InputData, "EdgesDemandPeak(dij)")
   
   #Edges demand annual 
   D = read_excel (InputData, "EdgesDemandAnnual(Dij)")
 
   # Capacité maximale des tuyaux
   Cmax = read_excel(InputData, "Cmax(cijmax)")
 
  # Capacité maximale de la source 
  Qmax = read_excel(InputData, "SourceMaxCap(Qimax)")
 
   # Pénalités de non satisfaction de la demande
   pumd = read_excel(InputData, "pumd(pijumd)")
   
   # longueurs des canaux 
   l= {}
   NodesCord = read_excel(InputData, "NodesCord")
   for i in range(1,int(len(NodesCord)/2)+1):
     for j in range(1,int(len(NodesCord)/2)+1):
         l[i,j] = distance.euclidean([NodesCord[i,1],NodesCord[i,2]],[NodesCord[j,1],NodesCord[j,2]]);
 
 
  ### create the decision variables###
    x_var = LpVariable.dicts('x', (set_V,set_V), lowBound = 0, upBound = 1, cat = 'Integer')
    Pin_var = LpVariable.dicts('Pin', (set_V,set_V), lowBound = 0, cat = 'Continous')
    Pout_var = LpVariable.dicts('Pout', (set_V,set_V), lowBound = 0, cat = 'Continous')
