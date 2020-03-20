
# assumption : all the arrays have to be sorted in such sequence that y from smallest to biggest!
# retrun q1 and q2 if selected section at suport.
# q1 and q2 base on averaging width "w"

def average (Lcan, Para, x_sec, y, m, w, arg):
    import numpy as np 
    # look up the variables 
    q = 0
   
    # define arrays
    y_a = np.array(y)
    m_a = np.array(m)

    if w !=0 :
        # extract and modify the needed arrays
        m_a_mod = m_a[0:-1] + m_a[1:] # top + bottom
        y_a_mod = (y_a[1:] - y_a[0:-1])/2 # height / 2

        # intergration
        m_tot = sum((m_a_mod * y_a_mod))
        
        # averrayge force by given width
        
        m_avg = m_tot / w
    
    elif w == 0 :
        m_tot = float (m_a)
        m_avg = m_tot

    # return result
    if arg == True and Para == 'My':
        q = (-m_tot*2) / (Lcan**2)
        return (q)
    elif arg == True and Para == 'Vy':
        q = (-m_tot) / (Lcan)
        return (q)
    else:
        return (m_avg)



# testing in put 
# node = [1.75,     2.5, 3.5,  4.5, 5.75]
# x =    2.5
# y =    [2.5,      10,  20,   30,  30.5]
# m =    [-187.5,   60,  250,  250, 137.5]

# Sec = 7.5

# lk = 2.5
# lm = 5
# alpha = 0.5
# w = 20

# FEMIn = {         'dia' : 0.2,
#                 'alpha' : alpha,
#                    'lm' : lm,
#                    'lk' : lk,
#                     'x' : x,
#                     'y' : y,
#                 'xload' : x/2,
#                 'yload' : 5,
#                     't' : [0.3, 0.3, 0.3, 0.3, 0.3],
#              'Ext_list' : ['_My', '_M2', '_Vy', '_Mxy', '_coor'], 
#          'materialName' : "C40/50",
#                 'xsupp' : x,
#                 'ysupp' : [lk ,  lk + lm],
#         'loadIntensity' : 10,
#               'poisson' : 0}


# (m_avg, q1, q2) = average (FEMIn, Sec, y, m, w)
# print(m_avg, q1, q2)
