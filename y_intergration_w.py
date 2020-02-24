
# assumption : all the arrays have to be sorted in such sequence that y from smallest to biggest!
# 基本前提 ：所有数列必须以y从小到大的顺序统一排列!!

def averrage (y, m, w):
    import numpy as np 

    # define arrays
    y_a = np.array(y)
    m_a = np.array(m)

    # extract and modify the needed arrays
    m_a_mod = m_a[0:-1] + m_a[1:] # top + bottom
    y_a_mod = (y_a[1:] - y_a[0:-1])/2 # height / 2

    # intergration
    m_tot = sum((m_a_mod * y_a_mod))

    # fancy way to intergrate 
    m_tot1 = float(m_a_mod @ y_a_mod.reshape((np.size(y_a_mod)),1))
    #print ("force intergration=",[m_tot]*len(m))
    #print ("force intergration in a fancy way :) =", m_tot1)
    
    # averrayge force by given width
    m_avg = m_tot / w
   # print ("averrage force=", m_avg)

    return (m_avg)


# testing in put 
# node = [1.75,     2.5, 3.5,  4.5, 5.75]
# x =    [15,       15,  15,   15,  15]
# y =    [2.5,      10,  20,   30,  30.5]
# m =    [-187.5,   60,  250,  250, 137.5]

# w = 35

# averrage (y, m, w)
