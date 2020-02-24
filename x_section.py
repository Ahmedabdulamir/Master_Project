
def section_interp (x_need, node, x, y, m):
    import numpy as np

    w = [0.5, 1, 1.5, 2]

    # define arrays
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

    # separate the arrays
    x_a_left = x_a[ind_lfet]
    x_a_right = x_a[ind_right]
    x_a_need = np.full(np.shape(x_a_left ),x_need)
    y_a_left = y_a[ind_lfet]
    #y_a_right = y_a[ind_right]
    m_a_left = m_a[ind_lfet]
    m_a_right = m_a[ind_right]

    # transpose the arrays to vectors and horizontal stack
    m_v_left = m_a_left.reshape(np.size(m_a_left),1)
    m_v_right =m_a_right.reshape(np.size(m_a_right),1)
    m_m = np.hstack([m_v_left, m_v_right])
        
    # interpolation
    # predefine the interpolation vector
    m_interp = np.zeros((np.size(m_a_left),1))

    # interpolating as a for loop 
    for i in range (0, np.size(m_a_left)):
        m_interp[i] = np.interp(x_need, [x_left, x_right], m_m[i])

    # redefine node, x, y, m 
    node = [float(x) + 0.5 for x in node_a[ind_lfet]]
    x = [float(x) for x in x_a_need]
    y = [float(x) for x in y_a_left]
 
    m = np.transpose (m_interp).reshape(np.size(m_interp))
    m = [float(x) for x in m]
 
    # print (node)
    # print(x)
    # print(y)
    # print(m)
    return (node, x, y,m)



