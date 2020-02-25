import os
import Parser as parser
import plate as plate
import pandas as pd
from pprint import pprint


ParserIn= {'Direction' : 'y',
                   'y' : [1,2],
                   'x' : [2,1],
           'Parameter' : ['M2','Vy'],      
                  'Vy' : int(9),
                  'M2' : int(4),
                  'My' : int(3),
             'Section' : [1.25], #, 2.4, 2.5, 2.6, 4, 5
                'Mesh' : float(0.25),
       'loadIntensity' : 10,
           'plot_sort' : 3,
                   'W' : [30]} #, 0.5, 1, 1.5, 2

dia = 0.2

alpha = 0.5
lm = 5
lk = alpha * lm
x = [0, lk, lk + lm/2, lk + lm, lm + 2*lk] 
t = [0.3, 0.3, 0.3, 0.3, 0.3]
y = 30
xsupp = [lk ,  lk + lm]
ysupp = y
xload = 0.2
#yload = y/2
Ext_list = ['_My', '_M2', '_Vy', '_coor'], 
materialName = "C40/50",


 

    

if __name__ == '__main__':
    Path = 'Alpha' + str(alpha) + '_Y' + str(y) + '_X' + str(x[len(x)-1])
    #plate.Deck(x, y, t, xsupp, y, xload, y/2, dia, loadIntensity, materialName, Mesh, '0', Ext_list, True, path, Section)
    #os.remove('xlsx/' + path + '/'+ path + Direction + '.xlsm')
    #for j in range (0, len(Para)):
        #Moment = Para[j]
        #for i in range (0, len(Section)):
            #parser._parser(path, Moment, Section[i], Direction, float(Mesh), plot_sort, y, w, False)
    
    
    
    Para_dict = parser.Analyze(Path, ParserIn)
    pprint(Para_dict)
    #parser._parser(path, Moment, Section, Direction, float(Mesh), plot_sort, y, w, False)   
    #parser._xlsExcel(Section, path, Direction, Moment, w, True)
    