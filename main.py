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
    
    
    
    
    
    
    
