# -*- coding: utf-8 -*-

"""Created on Fri Apr  3 01:40:13 2020
Dimensionamento sistema PV 
@author: Fabio
"""
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import  interp1d, PchipInterpolator
import scipy.integrate as integrate
#Consumo diario de energia(kWh) - Cargas residenciais

def Eg_PV(energia_residencial, n_cargas, penFactor, daily):
    if(daily == True):
        Eg = energia_residencial
    else:
        Eg = energia_residencial/365
    # fator penetracao    
    Eg_pv= Eg*penFactor    
    return Eg_pv
    
def parameters_Module():
    Pvst= np.array([[0,  25 , 75 , 100],[1.2, 1.0, 0.8,  0.6]])
    Eff = np.array([[0.1 , 0.2 , 0.4 ,1.0],[0.86 , 0.9 , 0.93 , 0.97]])
    Irrad = ([0, 0, 0, 0, 0, 0, .1, .2, .3, .5, .8, .9, 1, 1, .99, .9, .7, .4, .1, 0, 0, 0, 0, 0]) 
    Tshape= ([25, 25, 25, 25, 25, 25, 25, 25, 35, 40, 45, 50, 60, 60, 50, 40, 35, 30, 25, 25, 25, 25, 25, 25])
    Pmpp=0.975
    nptos= 24
    kV= 0.22    
    return Pvst, Eff, Irrad, Tshape, Pmpp, nptos, kV

def Module_Energy2(Pvst, Eff, Irrad, Tshape,Pmpp, nptos):
    
    P = np.zeros(nptos)
    for i in range(0, nptos):
        t= Tshape[i]    
        Pvst_curve= interp1d(Pvst[0], Pvst[1], 'linear')
        # Eff_curve = interp1d(Eff[0], Eff[1], 'quadratic')            
        Eff_curve = PchipInterpolator(Eff[0], Eff[1])            
        Factor_T= Pvst_curve(t)        
        efficience= Eff_curve(Factor_T)
        P[i]=Pmpp*Irrad[i]*0.98*Factor_T*efficience*0.98       
    
    # x= np.linspace(0,100, num=10)
    # plt.plot(Eff)
    # plt.figure()
    x2= np.linspace(1,24, num=24)
    # plt.plot(x2,P)
    # plt.grid(True)
    Energy=integrate.trapz(P, x2)    
    return Energy


if __name__ == "__main__":
    Pvst= np.array([[0,  25 , 75 , 100],[1.2, 1.0, 0.8,  0.6]])
    Eff = np.array([[0.1 , 0.2 , 0.4 ,1.0],[0.86 , 0.9 , 0.93 , 0.97]])
    Irrad = np.array([0, 0, 0, 0 ,0, 0, 0.1, 0.2, 0.3, 0.5, 0.8, 0.9 ,1.0, 1.0 ,0.99, 0.9 ,0.7, 0.4, 0.1 ,0, 0 ,0, 0, 0]) 
    Tshape= ([25, 25, 25, 25, 25, 25, 25, 25, 35, 40, 45, 50, 60, 60, 50, 40, 35, 30, 25, 25, 25, 25, 25, 25])
    Pmpp=0.975
    nptos= 24
    kV= 0.22
    energy=Module_Energy(Pvst, Eff, Irrad, Tshape,Pmpp, nptos)  
    print(energy)
    
# #Média do numero de horas de sol pleno em sao Paulo 
    # NSP = 4.64        
    # Pwp= Eg_pv/NSP        
    # #Fator de correcao pela temperatura
    # F=0.85        
    # #Eficiencia Inversor
    # n_inv= 0.92
    # # Perdas gerais
    # n_geral = 0.98        
    # Pwp_corrigida= Pwp/(n_inv*n_geral*F)        
    # print("potencia total a ser instalada: " +str(Pwp_corrigida))
    # # potencia media de uma painel PV
    # Pwp_painel = 975
    # # numero de paineis necessarios
    # n_paineis= Pwp_corrigida/Pwp_painel
    # print( "numero total de paineis: "+ str(n_paineis))        
    # # numero de paineis por consumidor
    # n = n_paineis/n_cargas
    # print("numero de paineis por consumidor: " +str(n))
    # # numero de paineis necessarios
    # n_paineis= Pwp_corrigida/Pwp_painel
    