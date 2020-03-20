import os
import Parser as parser
import XLSX
import plate as plate
import pandas as pd
from pprint import pprint
import json
import numpy as np

alpha = 0.5
lm = 5
lk = alpha * lm
y = [0, lk, lk + lm/2, lk + lm, lm + 2*lk] 
x = 30


FEMIn = {         'dia' : 0.2,
                'alpha' : alpha,
                   'lm' : lm,
                   'lk' : lk,
                    'x' : x,
                    'y' : y,
                'xload' : x/2,
                'yload' : 5,
                    't' : [0.3, 0.3, 0.3, 0.3, 0.3],
             'Ext_list' : ['_My', '_M2', '_Vy', '_Mxy', '_coor'], 
         'materialName' : "C40/50",
                'xsupp' : x,
                'ysupp' : [lk ,  lk + lm],
        'loadIntensity' : 10,
              'poisson' : 0}


ParserIn= {'Direction' : 'x',
                   'y' : [1,2],
                   'x' : [2,1],
           'Parameter' : ['My','Vy'],      
                  'Vy' : int(10),
                  'M2' : int(4),
                  'My' : int(4),
                 'Mxy' : int(5),
          'Section_My' : [1.25, 1.3, 2, 2.5, 3, 4, 5, 6, 7.5, 8], #list(np.arange(0.2,9.8, 0.1)),
         'Section_Mxy' : [1.25, 1.3, 2, 2.5, 3, 4, 5, 6, 7.5, 8], #list(np.arange(0.2,9.8, 0.1)),
          'Section_Vy' : [ 2, 2.15, 2.3, 2.4, 2.45, 2.5, 2.6, 2.7, 3, 4, 5, 7.5],
                'Mesh' : float(0.25),
           'plot_sort' : 3,
                   'W' : [30, 0, 0.5, 1, 1.5, 2, 2.5, 3],
        'aspect_ratio' : float(x/y[len(y)-1])}     

w_col = len(ParserIn['W'])

TableForm= {   'y_head' : 0,
             'Para_avg' : 1,        #First Row in Excel sheet
            'Hand_calc' : w_col + 4,        
                 'q1q2' : w_col*2 + 6, 
            'Eq_header' : w_col*3 + 8,          #Header values for eq width (Hidden)
           '-ve_header' : w_col*4 + 8,
           '+ve_header' : w_col*5 + 8,
          'Main Header' : w_col*6 + 11        #Vertical Table header
          }       

  

if __name__ == '__main__':
    Path = 'Alpha' + str(alpha) + '_Y' + str(y[len(y)-1]) + '_X' + str(x) + '_Load' + str(FEMIn['yload'])
    #plate.Deck(Path, ParserIn, FEMIn, True)
    Para_dict = parser.Analyze(Path, ParserIn, FEMIn)
    #with open('log.txt', 'w') as out:
        #pprint(Para_dict, stream = out)
    #pprint(Para_dict)
    XLSX.xlsExcel(Path, Para_dict, ParserIn, FEMIn, TableForm)
   
    
    