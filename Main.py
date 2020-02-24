import os
import Parser_ver3 as parser
import plate_ver3 as plate
import testxlsx as xlsx
import pandas as pd



Direction = "y"
Para = ["M2", "Vy"]               #My , M2, V
Section = [1.25, 2.4, 2.5, 2.6, 4, 5]
Mesh = '0.25'
materialName = "C40/50"
loadIntensity = 10
Ext_list = ['_My', '_M2', '_coor']   #Not needed
plot_sort = 3               # Y :3      Moment: 1


# Geometry
# x = [0, 10, 15, 20, 30, 35, 40, 50] 
# t = [1, 3, 3, 1, 1, 3, 3, 1]
# y = 100
# xsupp = [12.5, 37.5]
# ysupp = 100
# xload = 5
# yload = 50
# dia = 5                 #half of load dia

dia = 0.2
w = [30, 0.5, 1, 1.5, 2]
alpha = 0.5
lm = 5
lk = alpha * lm
x = [0, lk, lk + lm/2, lk + lm, lm + 2*lk] 
t = [0.3, 0.3, 0.3, 0.3, 0.3]
y = 30
xsupp = [lk ,  lk + lm]
#ysupp = y
xload = 0.2
#yload = y/2
 

    

if __name__ == '__main__':
    path = 'Alpha' + str(alpha) + '_Y' + str(y) + '_X' + str(x[len(x)-1])
    #plate.Deck(x, y, t, xsupp, y, xload, y/2, dia, loadIntensity, materialName, Mesh, '0', Ext_list, True, path, Section)
    os.remove('xlsx/' + path + '/'+ path + Direction + '.xlsm')
    for j in range (0, len(Para)):
        Moment = Para[j]
        for i in range (0, len(Section)):
            parser._parser(path, Moment, Section[i], Direction, float(Mesh), plot_sort, y, w, False)
    #parser._parser(path, Moment, Section, Direction, float(Mesh), plot_sort, y, w, False)   
    #parser._xlsExcel(Section, path, Direction, Moment, w, True)
    