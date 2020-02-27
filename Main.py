import os
import Parser as parser
import XLSX
import plate as plate
import pandas as pd
from pprint import pprint


alpha = 0.5
lm = 5
lk = alpha * lm
y = [0, lk, lk + lm/2, lk + lm, lm + 2*lk] 
x = 30


FEMIn = {         'dia' : 0.2,
                'alpha' : alpha,
                   'lm' : lm,
                    'lk': lk,
                    'x' : x,
                    'y' : y,
                'xload' : x/2,
                'yload' : 0.2,
                    't' : [0.3, 0.3, 0.3, 0.3, 0.3],
             'Ext_list' : ['_My', '_M2', '_Vy', '_coor'], 
         'materialName' : "C40/50",
                'xsupp' : x,
                'ysupp' : [lk ,  lk + lm],
        'loadIntensity' : 10,
              'poisson' : 0}


ParserIn= {'Direction' : 'x',
                   'y' : [1,2],
                   'x' : [2,1],
           'Parameter' : ['M2','Vy'],      
                  'Vy' : int(10),
                  'M2' : int(4),
                  'My' : int(3),
          'Section_M2' : [1.25, 1.3, 2, 2.5, 3, 4, 5],
          'Section_Vy' : [1.25, 2.4, 2.45, 2.6, 2.7, 3, 4, 5, 7.3],
                'Mesh' : float(0.25),
           'plot_sort' : 3,
                   'W' : [30, 0.5, 1, 1.5, 2, 2.5, 3],
        'aspect_ratio' : float(x/y[len(y)-1])}              

  

if __name__ == '__main__':
    Path = 'Alpha' + str(alpha) + '_Y' + str(y[len(y)-1]) + '_X' + str(x)
    #plate.Deck(Path, ParserIn, FEMIn, True)
    Para_dict = parser.Analyze(Path, ParserIn, FEMIn)
    #pprint(Para_dict['M2'][1.25])
    XLSX.xlsExcel(Path, Para_dict, ParserIn, FEMIn)
   
    
    