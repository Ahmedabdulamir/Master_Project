import numpy as np

# moment influence line
def inf_M (x_sec, x_p, alfa, Lsp):
    # span ratio alfa = Lcan/Lsp
    Lcan = alfa*Lsp
    Ltot = Lcan*2+Lsp

    if x_p < Lcan :
        if x_p < x_sec <= Lcan :
            return -(x_sec-x_p)
        elif Lcan < x_sec <= (Lcan + Lsp):
            return -(Lcan-x_p)*(1-(x_sec-Lcan)/Lsp)
        else:
            return 0
    elif Lcan <= x_p <= Lcan + Lsp and Lcan <= x_sec <= Lcan + Lsp :
        if x_p < x_sec:
            return ((x_p - Lcan)*(Lcan + Lsp - x_p) / Lsp)*(1 - (x_sec - x_p) / (Lcan + Lsp - x_p))
        else:
            return ((x_p - Lcan)*(Lcan + Lsp - x_p) / Lsp)*((x_sec - Lcan) / (x_p - Lcan))
    elif  Lcan + Lsp < x_p :
        if Lcan + Lsp < x_sec <= x_p :
            return -(x_p - x_sec)
        elif Lcan < x_sec <= Lcan + Lsp :
            return -(x_p - Lcan - Lsp)*((x_sec-Lcan) / Lsp)
        else:
            return 0
    else:
        return 0
    return 
   


# Moment for distributed load:
def M_q (alfa, Lsp, x_sec, x_p, q, a, q1, q2):
    # span ratio alfa = Lcan/Lsp
    Lcan = alfa*Lsp
    Ltot = Lcan*2+Lsp

    # part1 for acting load 

    # a: distribution length of load
    # n_q = 100: number of finite grains of n_q to be divide    
    ai = np.arange (-a/2, a/2 + a/100, a/100)
    xi = ai + np.full_like(ai, x_p)
    qi = np.full_like(ai, q*a/(100+1))
    inf_a_M = np.array ([inf_M (x_sec, x, alfa, Lsp) for x in xi])

    # part2 for dummy load:
    
    # q1:
    # a=Lsp: distribution length of load
    # n_q = 100: number of finite grains of n_q to be divide    
    ai_1 = np.arange (-Lcan/2, Lcan/2 + Lcan/100, Lcan/100)
    xi_1 = ai_1 + np.full_like(ai_1, Lcan/2)
    qi_1 = np.full_like(ai_1, q1*Lcan/(100+1))
    inf_a_M_1 = np.array ([inf_M (x_sec, x, alfa, Lsp) for x in xi_1])

    # q2:
    # a=Lcan: distribution length of load
    # n_q = 100: number of finite grains of n_q to be divide    
    ai_2 = np.arange (-Lcan/2, Lcan/2 + Lcan/100, Lcan/100)
    xi_2 = ai_2 + np.full_like(ai_2, (Lcan/2 + Lsp + Lcan))
    qi_2 = np.full_like(ai_2, q2*Lcan/(100+1))
    inf_a_M_2 = np.array ([inf_M (x_sec, x, alfa, Lsp) for x in xi_2])


    return (float (sum(qi*inf_a_M)) + float (sum(qi_1*inf_a_M_1)) + float (sum(qi_2*inf_a_M_2)))


# shear force influence line
def inf_V (x_sec, x_p, alfa, Lsp):
    # span ratio alfa = Lcan/Lsp
    Lcan = alfa*Lsp
    Ltot = Lcan*2+Lsp

    if 0 <= x_p < Lcan :
        if x_p >= x_sec:
            return 0
        if x_p < x_sec < Lcan :
            return -1
        if x_sec == Lcan :
            return 0
        if Lcan < x_sec < Lcan + Lsp :
            return (Lcan - x_p) / Lsp 
        if Lcan + Lsp <= x_sec <= Ltot :
            return 0
    
    elif Lcan <= x_p <= Lcan + Lsp :
        if 0 <= x_sec <= Lcan :
            return 0
        if Lcan < x_sec < x_p :
            return 1 - ((x_p - Lcan) / Lsp)
        if x_sec == x_p :
            return 0
        if x_p < x_sec < Lcan + Lsp :
            return - ((x_p - Lcan) / Lsp)
        if Lcan + Lsp <= x_sec <= Ltot :
            return 0
    
    elif Lcan + Lsp < x_p <= Ltot :
        if 0 <= x_sec <= Lcan :
            return 0
        if Lcan < x_sec < Lcan + Lsp :
            return 1 - ((x_p - Lcan) / Lsp)
        if x_sec == Lcan + Lsp :
            return 0
        if Lcan + Lsp < x_sec < x_p :
            return 1
        if x_p <= x_sec <= Ltot :
            return 0
      
    return 
   


# shear force for distributed load:
def V_q (alfa, Lsp, x_sec, x_p, q, a, q1, q2):
    
    # span ratio alfa = Lcan/Lsp
    Lcan = alfa*Lsp

    # part1 acting load
    # a: distribution length of load
    # n_q = 100: number of finite grains of n_q to be divide 

    ai = np.arange (-a/2, a/2+a/100, a/100)
    xi = ai + np.full_like(ai, x_p)
    qi = np.full_like(ai, q*a/(100+1))
    inf_a_V = np.array ([inf_V (x_sec, x, alfa, Lsp) for x in xi])

    # part2 dummy load
    # q1
    # a=Lcan: distribution length of load
    # n_q = 100: number of finite grains of n_q to be divide 

    ai_1 = np.arange (-Lcan/2, Lcan/2 + Lcan/100, Lcan/100)
    xi_1 = ai_1 + np.full_like(ai_1, Lcan/2)
    qi_1 = np.full_like(ai_1, q1*Lcan/(100+1))
    inf_a_V_1 = np.array ([inf_V (x_sec, x, alfa, Lsp) for x in xi_1])

    # q2
    # a=Lcan: distribution length of load
    # n_q = 100: number of finite grains of n_q to be divide 

    ai_2 = np.arange (-Lcan/2, Lcan/2 + Lcan/100, Lcan/100)
    xi_2 = ai_2 + np.full_like(ai_2, (Lcan/2 + Lsp + Lcan))
    qi_2 = np.full_like(ai_2, q2*Lcan/(100+1))
    inf_a_V_2 = np.array ([inf_V (x_sec, x, alfa, Lsp) for x in xi_2])

    # need an "-" before result for main code to match FEM design
    return ((float (-sum(qi*inf_a_V)) + float (-sum(qi_1*inf_a_V_1)) + float (-sum(qi_2*inf_a_V_2))))


def hand_calculation (FEMIn, q1q2, Para, Sec):
    alfa = FEMIn.get('alpha')
    Lsp = FEMIn.get('lm')
    x_sec = Sec
    x_p = FEMIn.get('yload')
    q = FEMIn.get('loadIntensity') * FEMIn.get('dia') * 2
    a = FEMIn.get('dia')*2
    q1 = q1q2.get('q' + str(FEMIn.get('ysupp')[0]))
    q2 = q1q2.get('q' + str(FEMIn.get('ysupp')[1]))
    
       
    if Para == 'My':
        return M_q (alfa, Lsp, x_sec, x_p, q, a, q1, q2)

    elif  Para == 'Vy':
        return V_q (alfa, Lsp, x_sec, x_p, q, a, q1, q2)
    else :
        print ('FEM output error')
    return
    
def _slice(y0, y1, y2, y3, z0, z1, z2, z3 ):   
    x0 = y0 [z0]
    x1 = y1 [z1]
    x2 = y2 [z2]
    x3 = y3 [z3]
    return (x0, x1, x2, x3)

def _find(m_a):
    up = [m1 < 0 for m1 in m_a]
    down = [m2 > 0 for m2 in m_a]  
    return(up, down)       
    
def _ind(up, down):
    ind_11 = [m1 == max(up) for m1 in up]
    ind_12 = [m2 == min(down) for m2 in down]
    return(ind_11, ind_12)              


def q_moment (FEMIn, x, m, w) :
    q1 = 0
    q2 = 0
    alfa = FEMIn.get('alpha')
    Lsp = FEMIn.get('lm')
    x_p = FEMIn.get('yload')
    q = FEMIn.get('loadIntensity') * FEMIn.get('dia') * 2
    a = FEMIn.get('dia')*2
    Lcan = alfa*Lsp
    Ltot = Lcan * 2 + Lsp
    
    # span ratio alfa = Lcan/Lsp
    if w < FEMIn.get('x'):
        # define the arrays
        # node_a = np.array(node)
        x_a = np.array(x)
        m_a = np.array(m)

        # slice the arrays to left and right
        ind_left = [Lcan < x1 < x_p for x1 in x_a]
        ind_right = [x_p < x2 < (Lcan + Lsp) for x2 in x_a]
        (x_a_left, x_a_right, m_a_left, m_a_right) = _slice(x_a, x_a, m_a, m_a, ind_left, ind_right, ind_left, ind_right)
        
        # slice the arrays to up and down
        (ind_leftup, ind_leftdown) = _find(m_a_left)
        (ind_rightup, ind_rightdown) = _find(m_a_right)

        (x_a_leftup, x_a_leftdown, x_a_rightup, x_a_rightdown) = _slice(x_a_left, x_a_left, x_a_right, 
                                                                    x_a_right, ind_leftup, ind_leftdown, ind_rightup, ind_rightdown)
        (m_a_leftup, m_a_leftdown, m_a_rightup, m_a_rightdown) = _slice(m_a_left, m_a_left, m_a_right, 
                                                                    m_a_right, ind_leftup, ind_leftdown, ind_rightup, ind_rightdown)
        # locate 4 points
        (ind_11, ind_12) = _ind(m_a_leftup, m_a_leftdown)
        (ind_21, ind_22) = _ind(m_a_rightup, m_a_rightdown)

        (x_11, x_12, x_21, x_22) = _slice(x_a_leftup, x_a_leftdown, x_a_rightup, 
                                                                    x_a_rightdown, ind_11, ind_12, ind_21, ind_22)
        (m_11, m_12, m_21, m_22) = _slice(m_a_leftup, m_a_leftdown, m_a_rightup, 
                                                                    m_a_rightdown, ind_11, ind_12, ind_21, ind_22)

        # interpolation to find the 0 moment position
        x_1 =float (x_11 + (abs (m_11) / (abs (m_11) + abs (m_12))) * (x_12 - x_11))
        x_2 =float (x_22 + (abs (m_22) / (abs (m_21) + abs (m_22))) * (x_21 - x_22))    

        # find the moment for x_1 and x_2 at hand calculation model
        m_1 = - M_q (alfa, Lsp, x_1, x_p, q, a, 0, 0)
        m_2 = - M_q (alfa, Lsp, x_2, x_p, q, a, 0, 0)
           
        incl = (m_2 - m_1) / (x_2 - x_1)
        # use inclination to extand the line to suport 
        # moment at suport 1 and 2
        m_s1 = m_1 - (x_1 - Lcan)*incl
        m_s2 = m_2 + (Lcan + Lsp - x_2)*incl

        # calculate q1 and q2
        q1 = -(m_s1*2) / (Lcan**2)
        q2 = -(m_s2*2) / (Lcan**2)

            
    return (q1, q2)


# x=[0,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9,1,1.1,1.2,1.3,1.4,1.5,1.6,1.7,1.8,1.9,2,2.1,2.2,2.3,2.4,2.5,2.6,2.7,2.8,2.9,3,3.1,3.2,3.3,3.4,3.5,3.6,3.7,3.8,3.9,4,4.1,4.2,4.3,4.4,4.5,4.6,4.7,4.8,4.9,5,5.1,5.2,5.3,5.4,5.5,5.6,5.7,5.8,5.9,6,6.1,6.2,6.3,6.4,6.5,6.6,6.7,6.8,6.9,7,7.1,7.2,7.3,7.4,7.5,7.6,7.7,7.8,7.9,8,8.1,8.2,8.3,8.4,8.5,8.6,8.7,8.8,8.9,9,9.1,9.2,9.3,9.4,9.5,9.6,9.7,9.8]
# m=[1,-0.0019,-0.0031,-0.0045,-0.0065,-0.0085,-0.0108,-0.0132,-0.0156,-0.018,-0.0204,-0.0234,-0.0266,-0.0299875,-0.0339375,-0.0378875,-0.042475,-0.047275,-0.0524625,-0.0588125,-0.0651625,-0.0727125,-0.0806625,-0.0857125,-0.0820625,-0.0784125,-0.065875,-0.050375,-0.0353375,-0.0216875,-0.0080375,0.005275,0.018475,0.031675,0.044875,0.058075,0.0719125,0.0859625,0.100525,0.116625,0.132725,0.1523875,0.1732375,0.1972125,0.2305625,0.2639125,0.300825,0.338925,0.3675,0.3675,0.3675,0.338925,0.300825,0.2639125,0.2305625,0.1972125,0.1732375,0.1523875,0.132725,0.116625,0.100525,0.0859625,0.0719125,0.058075,0.044875,0.031675,0.018475,0.005275,-0.0080375,-0.0216875,-0.0353375,-0.050375,-0.065875,-0.0784125,-0.0820625,-0.0857125,-0.0806625,-0.0727125,-0.0651625,-0.0588125,-0.0524625,-0.047275,-0.042475,-0.0378875,-0.0339375,-0.0299875,-0.0266,-0.0234,-0.0204,-0.018,-0.0156,-0.0132,-0.0108,-0.0085,-0.0065,-0.0045,-0.0031,-0.0019]

# q_moment (0.5, 5, 4, 0.4, x, 5, m)


# print (M_q (0.5, 5, 2.5, 5, 4, 0.4, 0.01592, 0.01592))




