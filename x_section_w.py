import numpy as np
import pandas as pd

def section_interp (x_need, df, w):
    
    # part 1 =====================================================================  
    # define arrays
    m =  df[1].tolist()       #Moment/ Shear values
    y =  df[3].tolist()       #Y coordinates
    node = df[0].tolist()      #Nodes
    x =  df[4].tolist()       #X coordinates
    
    node_a = np.array(node)
    x_a = np.array(x)
    y_a = np.array(y)
    m_a = np.array(m)

    # find left and right section
    x_left = min(x_a)
    x_right = max(x_a)

    # indexing that separate the array in left and right
    ind_lfet = x_a == x_left
    ind_right = x_a == x_right

    # separate the arrays in lfet and right
    node_a_left = node_a[ind_lfet]
    node_a_right = node_a[ind_right]
    x_a_left = x_a[ind_lfet]
    x_a_right = x_a[ind_right]
    x_a_need = np.full(np.shape(x_a_left ),x_need)
    y_a_left = y_a[ind_lfet]
    y_a_right = y_a[ind_right]

    # force at left and right section
    m_a_left = m_a[ind_lfet]
    m_a_right = m_a[ind_right]

    if x_right == x_left and x_left == x_need:
        print("existing section")
        m_a_interp = m_a_left

    elif x_right == x_left and x_left != x_need:
        print ("selected section didn't match the given section, using given x as x_need")
        x_need = x_left
        m_a_interp = m_a_left

    elif x_right > x_left and x_left <= x_need <= x_right:
        print("new interpolated section")
        # m_a_left and m_a_right as matrix  
        m_m =np.transpose(np.vstack((m_a_left, m_a_right)))
            
        # interpolation
        # predefine the interpolation vector
        m_interp = np.zeros((np.size(m_a_left),1))

        # interpolating as a for loop 
        for i in range (0, np.size(m_a_left)):
            m_interp[i] = np.interp(x_need, [x_left, x_right], m_m[i])

        m_a_interp = m_interp.reshape(np.size(m_interp))
        #print (m_a_interp) 
        # redefine node, x, y, m 
        #node = [float(x) + 0.5 for x in node_a[ind_lfet]]
        #x = [float(x) for x in x_a_need]
        #y = [float(y) for y in (y_a_left + y_a_right)/2]
    
        #m = np.transpose (m_interp).reshape(np.size(m_interp))
        #m = [float(x) for x in m]
        
        # tracking point
        # print (node)
        # print(x)
        # print(y)
        # print(m)

    else :
        print ("input error")
    
    # part 2 in y direction ========================================================  
    # slab length (biggest distribution width can be chosen)
    w_lim = max(y) - min(y)

    if w_lim >= w:
        # define the points 
        y_mid = (max(y)-min(y))/2 + min(y) 
        y_up = y_mid + w/2
        y_down = y_mid - w/2
        print ("middle point y=",y_mid)
        print ("upper point y=",y_up)
        print ("downer point y=",y_down)

        # find out if any node at given length in y direction
        case = [y1 == y_up or y1 == y_down for y1 in y_a_left]
        ca = sum (np.ones(np.shape(case))[case])
        #print (case)
        #print (ca)

        # logic return as a narrowed y array 
        if ca == 2:
            print ("existing nodes")
            #print (y_up, y_down)
            ind_na = [y2 <= y_up and y2 >= y_down for y2 in y_a_left] 
            y_a_na = y_a_left [ind_na]
            #print (ind_na)
            #print (y_a_na)
            # out put
            node_a_na = [node1 + 0.5 for node1 in node_a_left [ind_na]]
            x_a_na = x_a_need [ind_na]
            y_a_na = y_a_left [ind_na]
            m_a_na = m_a_interp[ind_na]
            print(np.vstack((node_a_na, x_a_na, y_a_na, m_a_na)))
            print("chosen width=", w)

        elif ca == 0:
            print ("created new nodes")
            print ("chosen width=", w)
            # indexing for nearest nodes
            ind_left_upup =  y_a_left == min(y_a_left[np.array([y3-y_up  for y3 in y_a_left]) > 0])
            ind_left_updown = y_a_left == max(y_a_left[np.array([y3-y_up  for y3 in y_a_left]) < 0])
            ind_left_downup = y_a_left == min(y_a_left[np.array([y3-y_down  for y3 in y_a_left]) > 0])
            ind_left_downdown = y_a_left == max(y_a_left[np.array([y3-y_down  for y3 in y_a_left]) < 0])
            # print (ind_left_upup)
            # print (ind_left_updown)
            # print (ind_left_downup)
            # print (ind_left_downdown)
            # indexing array y (nodes between upper and lower nodes)
            ind_bt = [y2 < y_up and y2 > y_down for y2 in y_a_left] 
            #print(ind_bt)
            # interpolation for upper and lower nodes
            m_inter_up = np.interp(y_up, np.hstack((y_a_left[ind_left_updown], y_a_left[ind_left_upup])), np.hstack((m_a_interp[ind_left_updown],m_a_interp[ind_left_upup])))
            m_inter_down = np.interp(y_down, np.hstack((y_a_left[ind_left_downdown], y_a_left[ind_left_downup])), np.hstack((m_a_interp[ind_left_downdown],m_a_interp[ind_left_downup])))

            # interpolate upper and downer nodes 
            node_a_na = [node1+0.5 for node1 in np.hstack((np.array(node_a_left[ind_left_downdown]+0.25),node_a_left[ind_bt],np.array(node_a_left[ind_left_upup]+0.25)))]
            x_a_na = np.hstack((x_a_need[ind_left_downdown], x_a_need[ind_bt], x_a_need[ind_left_upup]))
            y_a_na = np.hstack(([y_down], y_a_left[ind_bt], [y_up]))
            m_a_na = np.hstack(([m_inter_down], m_a_interp[ind_bt], [m_inter_up]))

            print (np.vstack((node_a_na, x_a_na, y_a_na, m_a_na)))
            

        else:
            print ("missing point in y direction")    

    else: 
        print ("chosen width out of range, using full slab length")
        # redefine w
        w = w_lim
        # define the points 
        y_mid = (max(y)-min(y))/2 + min(y) 
        y_up = y_mid + w/2
        y_down = y_mid - w/2
        print ("middle point y=",y_mid)
        print ("upper point y=",y_up)
        print ("downer point y=",y_down)
        
        # find out if any node at given length in y direction
        case = [y1 == y_up or y1 == y_down for y1 in y_a_left]
        ca = sum (np.ones(np.shape(case))[case])
        #print (case)
        #print (ca)

        # logic return as a narrowed y array 
        if ca == 2:
            print ("existing nodes")
            print ("chosen width", w)
            #print (y_up, y_down)
            ind_na = [y2 <= y_up and y2 >= y_down for y2 in y_a_left] 
            y_a_na = y_a_left [ind_na]
            #print (ind_na)
            #print (y_a_na)
            # out put
            node_a_na = [node1 + 0.5 for node1 in node_a_left [ind_na]]
            x_a_na = x_a_need [ind_na]
            y_a_na = y_a_left [ind_na]
            m_a_na = m_a_interp[ind_na]
            print(np.vstack((node_a_na, x_a_na, y_a_na, m_a_na)))
            
        else:
            print ("missing point at y direction")    

        
    # final output
    node = node_a_na
    x = x_a_na
    y = y_a_na
    m = m_a_na
    w = w
    return (node, x, y,m)


# define the variables
# node = [1, 6, 5, 10, 2, 3, 7, 8, 4, 9]
# x = [10, 20, 10, 20, 10, 10, 20, 20, 10, 20]
# y = [0, 0, 40, 40, 10, 20, 10, 20, 30, 30]
# m = [-100, 60, 50, 150, 0, 200, 120, 300, 100, 400]
# x_need = 15

# w = 21


# section_interp (x_need, node, x, y, m, w)

