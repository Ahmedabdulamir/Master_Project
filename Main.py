import os
import Parser as parser
#import plate as plate
import pandas as pd
from pprint import pprint


ParserIn= {'Direction' : 'y',
                   'y' : [1,2],
                   'x' : [2,1],
           'Parameter' : ['M2','Vy'],      
                  'Vy' :  int(9),
                  'M2' : int(4),
                  'My' : int(3),
             'Section' : [1.25, 2.4, 2.5, 2.6, 4, 5],
                'Mesh' : float(0.25),
           'plot_sort' : 3,
                   'W' : [30, 0.5, 1, 1.5, 2]}

                   

alpha = 0.5
lm = 5
lk = alpha * lm
x = [0, lk, lk + lm/2, lk + lm, lm + 2*lk] 
y = 30


FEMIn = {         'dia' : 0.2,
                'alpha' : alpha,
                   'lm' : lm,
                    'lk': lk,
                    'x' : x,
                    'y' : y,
                'xload' : 0.2,
                'yload' : y/2,
                    't' : [0.3, 0.3, 0.3, 0.3, 0.3],
             'Ext_list' : ['_My', '_M2', '_Vy', '_coor'], 
         'materialName' : "C40/50",
                'xsupp' : [lk ,  lk + lm],
                'ysupp' : y,
        'loadIntensity' : 10,
              'poisson' : 0}

  

if __name__ == '__main__':
    Path = 'Alpha' + str(alpha) + '_Y' + str(y) + '_X' + str(x[len(x)-1])
    #plate.Deck(x, y, t, xsupp, y, xload, y/2, dia, loadIntensity, materialName, Mesh, '0', Ext_list, , Path, FEMIn, True)
    Para_dict = parser.Analyze(Path, ParserIn, FEMIn)
    pprint(Para_dict['M2'][1.25])
    parser.xlsExcel(Path, Para_dict, ParserIn)
    #parser._parser(path, Moment, Section, Direction, float(Mesh), plot_sort, y, w, False)   
    
    