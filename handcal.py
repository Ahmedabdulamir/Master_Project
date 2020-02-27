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
def M_q (alfa, Lsp, x_sec, x_p, q, a):
    # a: distribution length of load
    # n_q = 1000: number of finite grains of n_q to be divide 

    import numpy as np
    ai = np.arange (-a/2, a/2+a/1000, a/1000)
    xi = ai + np.full_like(ai, x_p)
    qi = np.full_like(ai, q*a/(1000+1))
    inf_a_M = np.array ([inf_M (x_sec, x, alfa, Lsp) for x in xi])
    return float (sum(qi*inf_a_M))


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
        if Lcan + Lsp <= X_sec <= Ltot :
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
def V_q (alfa, Lsp, x_sec, x_p, q, a):
    # a: distribution length of load
    # n_q = 1000: number of finite grains of n_q to be divide 

    import numpy as np
    ai = np.arange (-a/2, a/2+a/1000, a/1000)
    xi = ai + np.full_like(ai, x_p)
    qi = np.full_like(ai, q*a/(1000+1))
    inf_a_V = np.array ([inf_V (x_sec, x, alfa, Lsp) for x in xi])
    return float (-sum(qi*inf_a_V))


def hand_calculation (FEMIn, Para, Sec):
    alfa = FEMIn.get('alpha')
    Lsp = FEMIn.get('lm')
    x_sec = Sec
    x_p = FEMIn.get('yload')
    q = FEMIn.get('loadIntensity') * FEMIn.get('dia') * 2
    a = FEMIn.get('dia')*2
       
    if Para == 'M2':
        return M_q (alfa, Lsp, x_sec, x_p, q, a)

    elif  Para == 'Vy':
        return V_q (alfa, Lsp, x_sec, x_p, q, a)
    else :
        return print ('FEM output error')
    return
    







