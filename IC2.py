# -*- coding: utf-8 -*-
"""
Created on Sun Feb 21 11:17:41 2021

@author: Fabio Andrade Zagotto
@utilização de algoritmos de Paulo Radatz, de seu do Youtube canal "Rumo ao Cinco Bola". https://github.com/PauloRadatz/tutorials_files 
"""

from numpy.core.fromnumeric import mean
from win32api import TerminateProcess
import win32com.client
from Dimensionamento_PV import parameters_Module, Module_Energy2
from LoadShape_type import train_model, predict_loadshape_Type
import matplotlib.pyplot as plt
import numpy as np
import scipy.integrate as integrate
import random
import gc
import time
from datetime import date
import os

""" --------//------Algoritmos Paulo Radatz-------//-------- """

class DSS(): 
    def __init__(self, end):
        self.fileAdd = end
        # connection between python and OpenDSS
        self.dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")
        # initiate DSS obj
        if self.dssObj.Start(0) == False:
            raise ValueError("Could not communicate with openDss")
        else:
            # state variables
            self.compiled = False
            # create variables for the main interfaces
            self.dssText = self.dssObj.text
            self.dssCircuit = self.dssObj.ActiveCircuit
            self.dssSettings = self.dssCircuit.Settings
            self.dssSolution = self.dssCircuit.Solution
            self.dssCktElement = self.dssCircuit.ActiveCktElement
            self.dssBus = self.dssCircuit.ActiveBus
            self.dssMonitors = self.dssCircuit.Monitors
            self.dssLines = self.dssCircuit.Lines
            self.dssGenerators = self.dssCircuit.Generators
            self.dssLoads = self.dssCircuit.Loads
            self.dssVsources = self.dssCircuit.Vsources
            self.dssTransformers = self.dssCircuit.Transformers
            self.dssLineCodes = self.dssCircuit.LineCodes
            self.dssLoadShapes = self.dssCircuit.LoadShapes
            self.dssPVSystems = self.dssCircuit.PVSystems
            self.dssMeters = self.dssCircuit.Meters
            self.dssCapacitors = self.dssCircuit.Capacitors

    def __exit__(self, exc_type, exc_value, traceback):
        print("closed")

    def versao_DSS(self):
        return self.dssObj.Version

    def get_nome_circuit(self):
        return self.dssCircuit.Name

    def compile_DSS(self):
        # Clear memory from last simulation
        self.dssObj.clearAll()
        # Compile File
        self.dssText.Command = ("compile " + self.fileAdd)
        self.compiled = True
        

    def solve_DSS_snapshot(self):
        # config solution
        self.dssText.Command = ("set Mode = SnapShot")
        self.dssText.Command = ("set controlmode= Static")
        self.compile_DSS()
        # Solve Power Flow
        self.dssSolution.Solve()

    def get_results_power(self):
        self.dssText.Command = ("show powers kva elements")

    def get_results_voltage(self):
        self.dssText.Command = ("Show voltages in nodes")

    def get_circuit_power(self):
        p = self.dssCircuit.TotalPower[0]
        q = self.dssCircuit.TotalPower[1]
        return p, q

    def get_barras_elemento(self):
        barras = self.dssCktElement.BusNames
        barra1 = barras[0]
        # barra2 = barras[1]
        return barra1
        # , barra2
    
    def get_tensoes_elemento(self):
        return self.dssCktElement.VoltagesMagAng

    def get_potencias_elemento(self):
        return self.dssCktElement.Powers

    def activate_bus(self, nome_barra):
        # Ativa a barra pelo seu nome
        self.dssCircuit.SetActiveBus(nome_barra)
        # Retornar o nome da barra ativada
        return self.dssBus.Name

    def activate_element(self, element_name):
        """
        Activated Element by the complete name
        i.e. type.name

        """
        ativou = self.dssCircuit.SetActiveElement(element_name)
        if(ativou != -1):
            return self.dssCktElement.Name

        else:
            raise Exception("%s inexistente" % element_name)

    def show_energyMeters(self):
        self.dssText.Command = ("Show Meters")
    
    def solve_DSS(self):
        # config solution
        self.dt = stepsize
        if daily==True:
            self.dssText.Command = ("set Mode = Daily")
        else:
            self.dssText.Command = ("set Mode = Yearly")
        self.dssText.Command = ("set stepsize= %dh" % stepsize)
        self.dssText.Command = ("set number= %d" % nHoras)
       
        # Solve Power Flow
        self.dssSolution.Solve()

    """ --------//------Algoritmos de autoria prória-------//-------- """

    def reset_arquives(self, folder_path):
        #cria os arquivos PVSystem.dss e StorageFleet.dss, se não existirem na pasta e os reseta
        f = open(folder_path+"\\StorageFleet.dss", "w")
        f.write("//Zerando")
        f.close()
        f = open(folder_path+"\\PVSystem.dss", "w")
        f.write("//Zerando")
        f.close()        
        self.compile_DSS()
    
    def get_loadConnectedElement(self, loadName):  
        load = self.dssLoads        
        load.Name = loadName
        
        self.dssCircuit.SetActiveElement(loadName)
        # self.activate_element("load." + loadName)
        if (self.dssCktElement.Name != "Load." + loadName):
            raise Exception("não foi possivel ativar a carga %s"%loadName)
        bus = self.dssCktElement.BusNames[0] 
        n_phases = self.dssCktElement.NumPhases
        # print("loadName=", loadName)
        # print("n_phases=", n_phases)
        lines = self.dssLines
        transformers_list = self.dssTransformers.AllNames
        # lista dos transformadores, tirando os do PVSystem
        transformers_notPV_list = [
            x for x in transformers_list if not 'pv' in x]            
        trafos=self.dssTransformers
        if n_phases == 1:            
            for line in lines.AllNames:
                lines.Name = line
                if(lines.Bus2.split(".",1)[0] == bus.split(".",1)[0]) : 
                    if ("." in bus and "." in lines.Bus2):
                        load_phases=bus.split(".",1)[1]
                        line_fases = lines.Bus2.split(".",1)[1]
                        if ("." in load_phases):
                            if (load_phases.split(".",1)[1] in line_fases and load_phases.split(".",1)[0] in line_fases):
                                return "line."+lines.Name
                        elif (load_phases in line_fases):
                            return "line."+lines.Name
                    if ("." in bus and "." not in lines.Bus2):
                        return "line."+lines.Name
            # if the load is connected to a transformer
            for trafo in transformers_notPV_list:
                trafos.Name = trafo
                self.activate_element("transformer."+trafos.Name)
                buses_trafo = obj.dssCktElement.BusNames
                if( bus.split(".",1)[0] in buses_trafo) : 
                    return "transformer." + trafos.Name
                   
        elif n_phases == 3:
            for line in lines.AllNames:
                lines.Name = line 
                obj.activate_element("line."+lines.Name)    
                if ("." in lines.Bus2):
                    if ("." in bus):
                        if(lines.Bus2.split(".",1)[0] == bus.split(".",1)[0]):
                            return "line."+lines.Name 
                    elif(lines.Bus2.split(".",1)[0] == bus):
                        return "line."+lines.Name 
                elif(lines.Bus2 == bus):
                    return "line."+lines.Name     
            for trafo in transformers_notPV_list:
                trafos.Name = trafo
                obj.activate_element("transformer."+trafos.Name)
                buses_trafo = obj.dssCktElement.BusNames
                if(bus.split(".",1)[0] in buses_trafo[0] or bus.split(".",1)[0] in buses_trafo[1]): 
                    return "transformer." + trafos.Name    

        raise Exception("Linha ou transformador conectado a carga %s não encontrado"%loadName)

    def convert_LS_AnualtoDaily(self, folder_path):  
        folder_daily= folder_path + '\daily'      
        if not os.path.exists(folder_daily): # se nao existe pasta "daily", cria uma
            os.makedirs(folder_daily)        
        f = open(folder_daily + "\LoadShapes_daily.dss", "w") # cria o arquivo de texto do loadshape diario
        # obtem o primeiro loadShape
        self.dssLoadShapes.First
        LoadShape = self.dssLoadShapes
        nome = LoadShape.Name
        # print("loadShape:" + nome)
        loadShape_array = np.array(LoadShape.Pmult)
        # converte para diario
        if(nome != "default"):
            i = 0
            ls_daily = np.zeros(24, dtype=float)
            while(i < loadShape_array.size):
                j = 0
                while(j < 24):
                    ls_daily[j] = ls_daily[j] + loadShape_array[i]
                    j = j+1
                    i = i+1
            ls_daily = ls_daily/365
            plt.plot(ls_daily)
            plt.title(nome + "_daily")
        while (LoadShape.Next):
            # obtem os outros loadshapes
            nome = LoadShape.Name
            # print("loadShape:" + nome)
            loadShape_array = np.array(LoadShape.Pmult)
            # converte para diario
            i = 0
            ls_daily = np.zeros(24, dtype=float)
            while(i < loadShape_array.size):
                j = 0
                while(j < 24):
                    ls_daily[j] = ls_daily[j] + loadShape_array[i]
                    j = j+1
                    i = i+1
            ls_daily = ls_daily/365
            ls_daily_str = "(" + str(ls_daily[0])
            k = 1
            while(k < ls_daily.size):
                ls_daily_str = ls_daily_str + ", " + str(ls_daily[k])
                k = k+1
            ls_daily_str = ls_daily_str + ")"
            # print(ls_daily_str)
            f.write("New loadshape.%s npts=24, interval= 1, mult=%s\r" % (
                nome, ls_daily_str))
        f.close()

    def create_monitor(self, elementName, mode, terminal):
        elementName = str.lower(elementName)
        mode_str = ""
        if(mode == 9):
            mode_str = "_loss"
        if(mode == 0):
            mode_str = "_V"
        monitor_name = elementName.split(".", 1)[1] + mode_str
        # check if monitor exists
        # if monitor_name not in self.dssMonitors.AllNames:
            # Cria um monitor com o nome do elemento
        self.dssText.Command =("New Monitor.%s element=%s terminal = %d mode=%d Ppolar=No" % (
            monitor_name, elementName, terminal, mode))
        # self.dssText.Command =("New Monitor.%s element=%s terminal = %d mode=%d Ppolar=No" % (
        #         monitor_name, elementName, terminal, mode))

    def activate_monitor(self, elementName, mode):
        abreviacao = ""
        if (mode == 0):
            abreviacao = "_v"
        if(mode == 9):
            abreviacao = "_loss"
        monitors = self.dssMonitors        
        monitors.Name = (elementName.split(".", 1)[1]+abreviacao)
        # print("elementName=",elementName)
        # print("mode=", mode)
        # print("self.dssMonitors.Element =",self.dssMonitors.Element)
        # print("self.dssMonitors.Mode = ", self.dssMonitors.Mode)
        
        if (self.dssMonitors.Element == elementName and self.dssMonitors.Mode == mode):
            return
        raise Exception("No monitors for %s in mode %d was found" %
                        (elementName, mode))


    def sampleAllMonitors(self):
        self.solve_DSS()
        self.dssMonitors.SampleAll

    def get_monitors_name(self):
        monitor = self.dssMonitors
        monitor.First
        print(monitor.Name + " Element:" + monitor.Element)
        while(monitor.Next):
            print(monitor.Name + " Element:" + monitor.Element)

    def get_MonitorProfile(self, elementName, mode):
        elementName = str.lower(elementName)
        self.activate_element(elementName)
        nphases = self.dssCktElement.NumPhases
        self.activate_monitor(elementName, mode)
        monitor = self.dssMonitors    
        # self.sampleAllMonitors()
        if (monitor.Element == elementName and monitor.Mode == mode):
            if(nphases == 1):    
                CH1_aux = np.array(monitor.Channel(1))
                CH1 = np.zeros(len(CH1_aux)+1)
                # correção para t=0, colocando o valor de ch[24] em ch[0]
                CH1[0] = CH1_aux[23]
                for i in range(1, len(CH1_aux)+1):
                    CH1[i] = CH1_aux[i-1]
            elif(nphases == 2):
                CH1_aux = np.array(monitor.Channel(1))
                CH2_aux = np.array(monitor.Channel(3))
                CH1 = np.zeros(len(CH1_aux)+1)
                CH2 = np.zeros(len(CH2_aux)+1)
                # correção para t=0, colocando o valor de ch[24] em ch[0]
                CH1[0] = CH1_aux[23]
                CH2[0] = CH2_aux[23]               
                for i in range(1, len(CH1_aux)+1):
                    CH1[i] = CH1_aux[i-1]
                    CH2[i] = CH2_aux[i-1]                    
            elif(nphases == 3):
                CH1_aux = np.array(monitor.Channel(1))
                CH2_aux = np.array(monitor.Channel(3))
                CH3_aux = np.array(monitor.Channel(5))
                CH1 = np.zeros(len(CH1_aux)+1)
                CH2 = np.zeros(len(CH2_aux)+1)
                CH3 = np.zeros(len(CH3_aux)+1)
                # correção para t=0, colocando o valor de ch[24] em ch[0]
                CH1[0] = CH1_aux[23]
                CH2[0] = CH2_aux[23]
                CH3[0] = CH3_aux[23]
                for i in range(1, len(CH1_aux)+1):
                    CH1[i] = CH1_aux[i-1]
                    CH2[i] = CH2_aux[i-1]
                    CH3[i] = CH3_aux[i-1]
            if (type(CH1) != np.ndarray or len(CH1) < 1):
                raise Exception(
                    "Perfil da elemento %s não obtido" % elementName)
            tamanho = len(CH1)
            time_array = np.arange(0, tamanho, 1, dtype=int)
            if(nphases == 1):
                return time_array, CH1
            elif(nphases == 2):
                return time_array, CH1, CH2
            elif(nphases == 3):
                return time_array, CH1, CH2, CH3
        print("O monitor do elemento %s nao foi encontrado" % elementName)

    def plot_MonitorProfile(self, elementName, mode):
        if (mode == 0):
            abreviacao = "V"
            extenso = "Tensao"
            unidade = "V"
        if (mode == 1):
            abreviacao = "P"
            extenso = "Pontencia"
            unidade = "kW"
        if (mode == 9):
            abreviacao = "P"
            extenso = "Perdas"
            unidade = "W"
        elementName = str.lower(elementName)
        self.activate_element(elementName)
        nphases = self.dssCktElement.NumPhases
        fig, ax = plt.subplots()
        if(nphases == 1):
            time_array, ch1 = self.get_MonitorProfile(elementName,  mode)
            ax.plot(time_array, ch1, label=abreviacao)
            ax.legend(loc='lower right')
        elif(nphases == 3):
            time_array, ch1, ch2, ch3 = self.get_MonitorProfile(
                elementName, mode)
            ax.plot(time_array, ch1, label=abreviacao+"1")
            ax.legend(loc='lower right')
            ax.plot(time_array, ch2, label=abreviacao+"2")
            ax.legend(loc='lower right')
            ax.plot(time_array, ch3, label=abreviacao+"3")
            ax.legend(loc='lower right')
        plt.grid(True)
        plt.title("Perfil de %s do elemento %s" % (extenso, elementName))
        plt.ylabel("%s (%s)" % (extenso, unidade))
        plt.xlabel("Tempo (horas)")
        plt.show()

    def get_elementEnergy(self, elementName, mode):
        elementName = str(elementName)
        self.activate_element(elementName)
        nphases = self.dssCktElement.NumPhases
        energy = 0
        if(nphases == 1):
            t, P = self.get_MonitorProfile(elementName, mode)
            # integral para obtencao da energia
            energy = integrate.trapz(P, t)
        if(nphases == 3):
            time_array, P1, P2, P3 = self.get_MonitorProfile(elementName, mode)
            energy = integrate.trapz(
                P1, time_array) + integrate.trapz(P2, time_array) + integrate.trapz(P3, time_array)
        return energy

    def get_LoadsPower_and_Energy(self, folder_path):
        # zerando PVSysistem e StorageFleet (para a energia das cargas não ser alterada)
        self.reset_arquives(folder_path)
        # setando variaveis
        load = self.dssLoads
        # cria monitores de potencia para todas as cargas
        loads_list = self.dssLoads.AllNames
        for (loadName) in loads_list:
            self.create_monitor("load."+loadName, 1, 1)
        # Amostra os monitores
        self.sampleAllMonitors()
        # tensao de base, n de fases, potencia nominal, energia consumida
        LoadsPower_and_EnergyMatrix = np.zeros((len(loads_list), 4))
        i = 0
        for (loadName) in loads_list:
            # obtem numero de fases, tensao e potencia de base da carga
            load.Name = loadName
            self.activate_element("load."+loadName)
            kWBase = load.kW
            kVBase = load.kV
            nphases = self.dssCktElement.NumPhases
            load_energy = self.get_elementEnergy("load."+loadName, 1)
            LoadsPower_and_EnergyMatrix[i, :] = (
                [kVBase, nphases, kWBase, load_energy])
            i = i+1        
        return loads_list, LoadsPower_and_EnergyMatrix
                    
    def get_LoadsPowerAndEnergybyClass(self, folder_path, loadshape_Namelist, ls_classification):
        # zerando PVSysistem e StorageFleet (para a energia das cargas não ser alterada)
        self.reset_arquives(folder_path)
        # setando variaveis
        load = self.dssLoads
        # cria monitores de potencia para todas as cargas
        loads_list = self.dssLoads.AllNames    
        for (loadName) in loads_list:
            self.create_monitor("load."+loadName, 1, 1)            
        # Amostra os monitores
        self.sampleAllMonitors()
        residential_nameList, commercial_nameList = [], []  # load name
        # tensao de base, n de fases, potencia nominal, energia consumida
        residential_matrix, commercial_matrix = [], []    
        for (loadName) in loads_list:
            # obtem numero de fases, tensao e potencia de base da carga
            load.Name = loadName
            self.activate_element("load."+loadName)
            kWBase = load.kW
            kVBase = load.kV
            nphases = self.dssCktElement.NumPhases
            load_energy = self.get_elementEnergy("load."+loadName, 1)
            loadshape_name = load.daily
            postion = loadshape_Namelist.index(str.lower(loadshape_name))
            load_classification = ls_classification[postion]
            if load_classification == 1:                            
                residential_nameList.append(loadName)
                residential_matrix.append((kVBase, nphases, kWBase, load_energy))
            elif load_classification == 2:
                commercial_nameList.append(loadName)
                commercial_matrix.append((kVBase, nphases, kWBase, load_energy))  
            #else:      
        residential_matrix = np.array(residential_matrix)
        commercial_matrix = np.array(commercial_matrix)        
        return residential_nameList, residential_matrix, commercial_nameList, commercial_matrix

    def get_classes_power_and_energy(self, residential_loads_matrix, commercial_loads_matrix):
        residential_power, residential_energy, commercial_power, commercial_energy = 0.0, 0.0, 0.0, 0.0
        if len(residential_loads_matrix) > 0:
            residential_power = np.sum(residential_loads_matrix[:, 2])
            residential_energy = np.sum(residential_loads_matrix[:, 3])
        if len(commercial_loads_matrix) > 0:
            commercial_power = np.sum(commercial_loads_matrix[:, 2])
            commercial_energy = np.sum(commercial_loads_matrix[:, 3])
        return residential_power, residential_energy, commercial_power, commercial_energy

    def Module_Energy(self, folder_path, Pvst, Eff, Irrad, Tshape, Pmpp, nptos, kV):
        f = open(folder_path + "\\PVSystem.dss", "w")        
        f.write("// P-T curve is per unit of rated Pmpp vs temperature\n")
        f.write("New XYCurve.MyPvsT npts=%d  xarray=%s  yarray=%s \n" %
                (len(Pvst[0]), str(Pvst[0]), str(Pvst[1])))
        f.write("\n// efficiency curve is per unit eff vs per unit power\n")
        f.write("New XYCurve.MyEff npts=%d  xarray=%s  yarray=%s \n" %
                (len(Eff[0]), str(Eff[0]), str(Eff[1])))
        f.write("\n// per unit irradiance curve (per unit if 'irradiance' property)\n")
        f.write("New Loadshape.MyIrrad npts=%d interval=1 mult=%s\n" %
                (nptos, str(Irrad)))
        f.write("\n//24-hr temp shape curve\n")
        f.write("New Tshape.MyTemp npts=%d interval=1 temp=%s\n\n" %
                (nptos, str(Tshape)))
        # utilizando a primeira carga
        load = self.dssLoads
        load.First
        loadName = load.Name
        loadkV = load.kV        
        self.activate_element("load."+str(loadName))
        bus = self.dssCktElement.BusNames[0]
        kVA = np.ceil(Pmpp)
        n_phases = self.dssCktElement.NumPhases          
        f.write("New PVSystem.PV_%s phases=%d bus1=trafo_pv_%s kV=%f  kVA=1.0  irrad=.98  Pmpp=%f temperature=25 PF=1\n" % (
            loadName, n_phases, loadName, kV, Pmpp))
        f.write(
            "~ %cutin=0.1 %cutout=0.1  effcurve=Myeff  P-TCurve=MyPvsT Daily=MyIrrad  Tdaily=MyTemp\n")
        f.write("New Transformer.pv_%s  phases= %d xhl=5.750000\n" % (str.upper(loadName), n_phases))
        f.write("~ wdg=1 bus=trafo_pv_%s kV=%f kVA=%f conn=wye\n" %
                (str.upper(loadName), kV, kVA))
        f.write("~ wdg=2 bus=%s kV=%f kVA=%f conn=wye\n\n" %
                (str.upper(bus), loadkV, kVA))
        f.write("New Energymeter.PV_%s element=transformer.pv_%s terminal = 1" % (
            str.upper(loadName), str.upper(loadName)))
        f.close()
        meterName = "pv_"+str.lower(loadName)
        self.compile_DSS()  # compilando o DSS denovo
        self.solve_DSS()
        self.dssMeters.Name = meterName  # setando o registrador        
        if self.dssMeters.Name == meterName:
            module_energy = self.dssMeters.RegisterValues[0]# Pegando o valor do Registrador do módulo PV
        else:
            raise Exception("energy meter not found.")
        return module_energy

    # def raffle_PV_Loads(self, loads_nameList, loads_matrix, parameter_PV_wanted, module_power, penFactor_storage):
    #     parameter_PV = 0
    #     PVLoads_list, num_modules_list = [], []  # relacao com os nomes das cargas sorteadas e o numero de modulos ligados a elas
    #     loads_nameList_aux = []
    #     loads_nameList_aux[:] = loads_nameList[:]                        
    #     while (parameter_PV < parameter_PV_wanted and len(loads_nameList_aux) > 0):            
    #         raffled_load = random.choice(loads_nameList_aux) # sorteando uma carga            
    #         i = loads_nameList.index(raffled_load)# obtendo a energia da carga sorteada           
    #         load_power = loads_matrix[i, 2]            
    #         n_modules = np.floor(np.minimum(load_power, parameter_PV_wanted - parameter_PV)/module_power)            
    #         # maxima potencia PV em uma carga = potencia base da carga
    #         n_modules_max = np.floor(load_power/module_power)
    #         if n_modules > n_modules_max:  # limitando numero de modulos pela potencia
    #             n_modules = n_modules_max            
    #         loads_nameList_aux.remove(raffled_load)
    #         if(n_modules >= 1):  # se for igual a 0, descartar
    #             PVLoads_list.append(raffled_load)
    #             num_modules_list.append(n_modules) 
    #             parameter_PV = parameter_PV + n_modules*module_power  
    #     return PVLoads_list, num_modules_list
    
    def raffle_PV_Loads(self, loads_nameList, loads_matrix, parameter_PV_wanted, module_power, module_energy):
        parameter_PV = 0
        PVLoads_list, num_modules_list = [], []  # relacao com os nomes das cargas sorteadas e o numero de modulos ligados a elas
        loads_nameList_aux = []
        loads_nameList_aux[:] = loads_nameList[:]                        
        while (parameter_PV < parameter_PV_wanted and len(loads_nameList_aux) > 0):            
            raffled_load = random.choice(loads_nameList_aux) # sorteando uma carga            
            i = loads_nameList.index(raffled_load)# obtendo a energia da carga sorteada           
            load_energy = loads_matrix[i, 3]            
            n_modules = np.floor(np.minimum(load_energy/module_energy, (parameter_PV_wanted - parameter_PV)/module_power))            
            # maxima potencia PV em uma carga = potencia base da carga
            #n_modules_max = np.floor(load_power/module_power)
            # if n_modules > n_modules_max:  # limitando numero de modulos pela potencia
            #     n_modules = n_modules_max            
            loads_nameList_aux.remove(raffled_load)
            if(n_modules >= 1):  # se for igual a 0, descartar
                PVLoads_list.append(raffled_load)
                num_modules_list.append(n_modules) 
                parameter_PV = parameter_PV + n_modules*module_power  
        return PVLoads_list, num_modules_list

    def raffle_Storage_Loads(self, PVLoads_list, penFactor_storage):        
        num_storages_wanted = int(np.ceil(len(PVLoads_list)*penFactor_storage))            
        StorageLoads_list = []  # relacao de cargas com Storage
        PVLoads_list_aux = []
        PVLoads_list_aux[:] = PVLoads_list[:]        
        for k in range (0, num_storages_wanted):                        
            raffled_load = random.choice(PVLoads_list_aux) # sorteando uma carga 
            StorageLoads_list.append(raffled_load)
            PVLoads_list_aux.remove(raffled_load)        
        return StorageLoads_list

    def create_PVSystem(self, folder_path, residential_PVLoads_NameList, residential_PVLoads_numModulesList, commercial_PVLoads_NameList, commercial_PVLoads_numModulesList, Pvst, Eff, Irrad, Tshape, Pmpp, nptos, kV):
        f = open(folder_path+"\\PVSystem.dss", "w")
        f.write("// P-T curve is per unit of rated Pmpp vs temperature\n")
        f.write("New XYCurve.MyPvsT npts=%d  xarray=%s  yarray=%s \n" %
                (len(Pvst[0]), str(Pvst[0]), str(Pvst[1])))
        f.write("\n// efficiency curve is per unit eff vs per unit power\n")
        f.write("New XYCurve.MyEff npts=%d  xarray=%s  yarray=%s \n" %
                (len(Eff[0]), str(Eff[0]), str(Eff[1])))
        f.write("\n// per unit irradiance curve (per unit if 'irradiance' property)\n")
        f.write("New Loadshape.MyIrrad npts=%d interval=1 mult=%s\n" %
                (nptos, str(Irrad)))
        f.write("\n//24-hr temp shape curve\n")
        f.write("New Tshape.MyTemp npts=%d interval=1 temp=%s\n\n" %
                (nptos, str(Tshape)))
        f.write("//Numero total de modulos PV residenciais: %d\n\n" %
                int(np.sum(residential_PVLoads_numModulesList)))
        f.write("//Numero total de modulos PV comerciais: %d\n\n" %
                int(np.sum(commercial_PVLoads_numModulesList)))
        load = self.dssLoads
        PVLoads_NameList = residential_PVLoads_NameList + commercial_PVLoads_NameList
        PVLoads_numModulesList = residential_PVLoads_numModulesList + commercial_PVLoads_numModulesList   
        for k in range (0, len(PVLoads_NameList)):
            loadName = PVLoads_NameList[k]
            n_modules = int(PVLoads_numModulesList[k])
            load.Name = loadName
            loadkV = load.kV
            loadkW = load.kW
            self.activate_element("load."+str(loadName))
            bus = self.dssCktElement.BusNames[0]
            n_phases = self.dssCktElement.numPhases
            kVA = np.ceil(Pmpp)*n_modules
            f.write("\n//load: %s - Pnom = %s kW - numero de modulos PV: %d\n" %
                    (str.upper(loadName), loadkW, n_modules))
            for i in range(1, n_modules+1):
                f.write("New PVSystem.PV_%s_%d phases= %d bus1=trafo_pv_%s kV=%f  kVA=1.0  irrad=.98  Pmpp=%f temperature=25 PF=1\n" % (
                    loadName, i, n_phases, loadName, kV, Pmpp))
                f.write(
                    "~ %cutin=0.1 %cutout=0.1  effcurve=Myeff  P-TCurve=MyPvsT Daily=MyIrrad  Tdaily=MyTemp\n")
            f.write("New Transformer.pv_%s  phases=%d xhl=1.500\n" %
                    (str.upper(loadName), n_phases))
            f.write("~ wdg=1 bus=trafo_pv_%s kV=%f kVA=%f conn=wye\n" %
                    (str.upper(loadName), kV, kVA))
            f.write("~ wdg=2 bus=%s kV=%f kVA=%f conn=wye\n" %
                    (str.upper(bus), loadkV, kVA))
            f.write("~ %loadloss = 0\n\n") 
            k = k + 1
       

    def create_storage(self, folder_path, residential_StorageLoads_List, commercial_StorageLoads_List, storagekW_percentage):                
        f = open(folder_path + "\\StorageFleet.dss", "w")        
        f.write("//Inverter Efficiency Curve\n")
        f.write(
            # "New XYCurve.Eff npts=4 xarray = [.1 .2 .4 1.0 ]  yarray = [.86 .9 .93 .97]\n")
            "New XYCurve.Eff npts=4 xarray = [.1 .2 .4 1.0 ]  yarray = [1 1 1 1]\n")
        load = self.dssLoads
        storageLoads_List = residential_StorageLoads_List + commercial_StorageLoads_List
        # Cria monitores linhas
        lines_list = self.dssLines.AllNames                
        for (lineName) in lines_list:
            self.create_monitor("line."+lineName, 1, 1)
        transformers_list = self.dssTransformers.AllNames
        # lista dos transformadores, tirando os do PVSystem
        transformers_notPV_list = [
            x for x in transformers_list if not 'pv' in x]
        for (trafoName) in transformers_notPV_list:
            self.create_monitor("transformer."+trafoName, 1, 1)
        self.sampleAllMonitors()
        
        if len(storageLoads_List) > 0:
            for i in range(0, len(storageLoads_List)):
                loadName = storageLoads_List[i]
                load.Name = loadName # setando a carga
                loadkV = load.kV
                loadkW = load.kW
                loadConnectedElement = self.get_loadConnectedElement(loadName)
                if loadConnectedElement:
                # if loadConnectedElement and "transformer" not in loadConnectedElement:
                    # self.activate_element(loadConnectedElement)
                    self.activate_element("load."+loadName)
                    # print("loadName=", loadName)
                    bus = self.dssCktElement.BusNames[0]
                    nPhases = self.dssCktElement.numPhases
                    if(nPhases == 1):
                        connectedPhase = self.get_loadConnectedPhase(loadName)
                        monPhase = str(connectedPhase)  # fase monitorada
                    else:
                        monPhase = "1"
                    # obtendo os valores de target, max e min
                    kWtarget_min, kWtarget_max = self.get_kWTargets(nPhases, loadConnectedElement)
                    kWrated = loadkW*storagekW_percentage
                    kWhrated = 5*kWrated
                    f.write("\n//load: %s - Porcentagem da potencia atendida pela bateria: %.2f\n" %
                            (str.upper(loadName), storagekW_percentage*100))
                    f.write("New Storage.Battery_%s phases=%d Bus1=%s kV=%.2f kWrated=%.2f kWhrated=%.2f \n" % (
                        loadName, nPhases, bus, loadkV, kWrated, kWhrated))
                    f.write(
                        "~ %idlingkW=0 EffCurve=Eff %Charge=100 %discharge=100  %stored=0  %reserve=0\n")                    
                    f.write("\nNew StorageController.sc_%s element=%s terminal=2  MonPhase=%s \n" % (
                        loadName, loadConnectedElement, monPhase))
                    f.write("~ modedis = peakshave kwtarget=%.2f modecharge=peakshavelow kwtargetlow=%.2f \n" % (
                        kWtarget_max, kWtarget_min))
                    f.write("~ %%rateCharge=30 %%reserve=0 elementList = [Battery_%s] eventlog=yes\n"%loadName)   
        f.close()

    def get_loadConnectedPhase(self, loadName):
        # has to be a monophasic load
        self.activate_element("load."+str(loadName))
        busName = self.dssCktElement.BusNames[0]
        if "." in busName:
            connectedPhase = busName.split(".", 1)[1]
        else:
            raise Exception(
                "Nao foi possivel achar a qual fase a carga %s esta conectada" % loadName)
        return connectedPhase

    def get_kWTargets(self, nPhases_Load, loadConnectedElemment):  
        # print("loadConnectedElemment=",loadConnectedElemment)      
        # phase=[]
        # self.create_monitor(loadConnectedElemment, 1, 1)
        # self.sampleAllMonitors()
        # self.activate_monitor(loadConnectedElemment, 1)
        self.activate_element(loadConnectedElemment)                    
        nPhases = self.dssCktElement.numPhases
        # energy = self.get_elementEnergy(loadConnectedElemment, 1)
        # Pmed = energy/24
        # print("nPhases=", nPhases)
        if nPhases == 1:
            t, phase = self.get_MonitorProfile(loadConnectedElemment, 1)
        elif nPhases == 2:
            t, Ph_1, Ph_2 = self.get_MonitorProfile(loadConnectedElemment, 1)
            # max value between 2 phases at any time
            phase = np.zeros(len(Ph_1))
            for i in range(0, len(Ph_1)):
                phase[i] = np.mean([Ph_1[i], Ph_2[i]])
        elif nPhases == 3:
            t, Ph_1, Ph_2, Ph_3 = self.get_MonitorProfile(loadConnectedElemment, 1)
            # max value between 3 phases at any time
            phase = np.zeros(len(Ph_1))
            for i in range(0, len(Ph_1)):
                phase[i] = np.mean([Ph_1[i], Ph_2[i], Ph_3[i]])        
        
        if len(phase)<1:
            raise Exception("erro ao obter as fases")
        mean = np.mean(phase)
        std_dev = np.std(phase)
        # kWtarget_min = (mean - std_dev/3)*3  # arbritary criteria
        # kWtarget_max = (mean + std_dev/5)*3
        kWtarget_min = (mean*0.90)*3  # arbritary criteria
        kWtarget_max = mean*3
        # print("mean*3=",mean*3)
        # print("kWtarget_min=",kWtarget_min)
        # print("kWtarget_max",kWtarget_max)
        return kWtarget_min, kWtarget_max

    def voltageException(self, parameter_overV, parameter_subV):
        # Cria monitores de tensao para todas as cargas e linhas
        loads_list = self.dssLoads.AllNames
        for (loadName) in loads_list:
            self.create_monitor("load."+loadName, 0, 1)
        # lines_list = self.dssLines.AllNames
        # for (lineName) in lines_list:
        #     self.create_monitor("line."+lineName, 0, 1)
        # transformers_list = self.dssTransformers.AllNames
        # # lista dos transformadores, tirando os do PVSystem
        # transformers_notPV_list = [
        #     x for x in transformers_list if not 'pv' in x]
        # for (trafoName) in transformers_notPV_list:
        #     self.create_monitor("transformer."+trafoName, 0, 1)
        self.sampleAllMonitors()
        monitors_list = self.dssMonitors.AllNames
        monitors = self.dssMonitors
        num_overVolt, num_underVolt = 0, 0  # numero de ocorrencias de sobre e sub tensao
        underV_report, overV_report = [], []
        for (monitor) in monitors_list:
            monitors.Name = monitor
            # ativando o elemento sendo monitorado
            elementName = monitors.Element
            self.activate_element(elementName)
            self.dssCircuit.SetActiveElement(elementName)
            # print("self.dssCktElement.Name=",str.lower(self.dssCktElement.Name))
            # print("elementName=",elementName)
            if (str.lower(self.dssCktElement.Name) != elementName):
                raise Exception("Nao foi possivel ativar o elemento")
            
            nphases = self.dssCktElement.NumPhases
            # obtendo a tensao de base
            bus_name = self.dssCktElement.BusNames[0]                        
            self.activate_bus(bus_name)
            # print("self.dssBus.Name=",str.lower(self.dssBus.Name))
            # print("bus_name=",bus_name)
            if (str.lower(self.dssBus.Name) != bus_name.split(".", 1)[0]):
                raise Exception("Nao foi possivel ativar a barra")
            kV_Base = self.dssBus.kVBase
            v1,v2,v3 = np.array([]), np.array([]), np.array([])
            horas_overVolt1, horas_overVolt2, horas_overVolt3 = 0, 0, 0
            horas_underVolt1, horas_underVolt2, horas_underVolt3 = 0, 0, 0
            if(nphases == 1):  # se elemento monitorado eh monifasico
                connected_phase = bus_name.split(".", 1)[1]
                if connected_phase == 1:                    
                    v1 = np.array(monitors.Channel(1))/(kV_Base*1000)  # em pu                    
                elif connected_phase == 2:
                    v2 = np.array(monitors.Channel(1))/(kV_Base*1000)  # em pu                    
                elif connected_phase == 3:
                    v3 = np.array(monitors.Channel(1))/(kV_Base*1000)  # em pu  
            elif(nphases == 3):
                v1 = np.array(monitors.Channel(1))/(kV_Base*1000)
                v2 = np.array(monitors.Channel(3))/(kV_Base*1000)
                v3 = np.array(monitors.Channel(5))/(kV_Base*1000)                
            if v1.size > 0:
                horas_overVolt1 = np.sum(np.fromiter( 
                    (i > parameter_overV for i in v1), int))#horas sobretensao fase 1
                horas_underVolt1 = np.sum(np.fromiter( 
                    (i < parameter_subV for i in v1), int))#horas subtensao fase 1
            if v2.size > 0:
                horas_overVolt2 = np.sum(np.fromiter(
                    (i > parameter_overV for i in v2), int))
                horas_underVolt2 = np.sum(np.fromiter(
                    (i < parameter_subV for i in v2), int))
            if v3.size > 0:
                horas_overVolt3 = np.sum(np.fromiter(
                    (i > parameter_overV for i in v3), int))
                horas_underVolt3 = np.sum(np.fromiter(
                    (i < parameter_subV for i in v3), int))
            if(horas_overVolt1 > 0 or horas_overVolt2 > 0 or horas_overVolt3 > 0):# verifica se ocorreu sobretensao
                num_overVolt = num_overVolt+1
                overV_report.append((elementName, horas_overVolt1, horas_overVolt2, horas_overVolt3))
                # print("sobretensao: elemento: "+str(elementName)+" num horas: "+str(horas_overVolt1))  
            if(horas_underVolt1 > 0 or horas_underVolt2 > 0 or horas_underVolt3 > 0): # verifica se ocorreu subtensao
                num_underVolt = num_underVolt+1
                underV_report.append(
                    (elementName, horas_underVolt1, horas_underVolt2, horas_underVolt3))
                # print("subtensao: elemento: "+str(elementName)+" num horas: "+str(horas_underVolt1))
                # print(nphases)
                # print(kV_Base)
        return num_overVolt, num_underVolt, underV_report, overV_report

    def StatisticalAnalysis(self, folder_path, folder_analysis, parameter_overV, parameter_subV, num_simulations, penFactor_list):
        
        report_matrix = np.zeros((len(penFactor_list), 12))  # matriz com perdas e transgressoes de tensao para cada combinacao de penetracoes
        underV_report_tot, overV_report_tot = [], []
        # underV_report_tot, overV_report_tot = np.array([]), np.array([])        
        Pvst, Eff, Irrad, Tshape, Pmpp, nptos, kV= parameters_Module() # parametros modulo PV     
        module_energy=self.Module_Energy(folder_path, Pvst, Eff, Irrad, Tshape, Pmpp, nptos, kV) #energia de um modulo
        print("module_energy",module_energy)
        #listas de cargas, potencia total e energia de cargas comerciais e residenciais         
        residential_nameList, residential_loads_matrix, commercial_nameList, commercial_loads_matrix = self.get_LoadsPowerAndEnergybyClass(folder_path,loadshape_Namelist, ls_classification)
        residential_power, residential_energy, commercial_power, commercial_energy = self.get_classes_power_and_energy(residential_loads_matrix, commercial_loads_matrix)
        self.compile_DSS()  
        f = open(folder_analysis + "\StatisticalAnalysis_report.txt", "w")       
        j=0
        for (penFactor_residential, penFactor_commercial, penFactor_storage_residential, penFactor_storage_commercial, storagekW_percentage) in penFactor_list:                
            substation_energy_list, total_losses_list, line_losses_list, transf_losses_list = [], [], [], [] # listas que contem a energia das cargas e as perdas em cada uma das 'num_simulations' simulacoes
            num_overVolt_list, num_underVolt_list = [], [] # listas que contem o numero de transgressoes de tensao cada uma das 'num_simulations' simulacoes
            residential_PVLoads_NameList, commercial_PVLoads_NameList = [], []
            residential_PVLoads_numModulesList, commercial_PVLoads_numModulesList = [], []
            residential_StorageLoads_list, commercial_StorageLoads_list = [], []
            f.write("#Penetrations: PV residential = %.2f %%, PV commercial = %.2f %%, residential storage = %.2f %% of residential PV, commercial storage = %.2f %% of commercial PV, storagekW = %.2f %% of loads KW \n" % (penFactor_residential*100, penFactor_commercial*100, penFactor_storage_residential*100, penFactor_storage_commercial*100, storagekW_percentage))
            f.write("i      num_modules_residential    num_modules_commercial     OverVolt        UnderVolt       total_losses (kWh)       line_losses (kWh)         transf_losses (kWh) \n")
            print("\npenFactors case : %d \n" % (j+1))
            for i in range (0, num_simulations):
                residential_PV_wanted=penFactor_residential*residential_power # energia PV residencial desejada
                commercial_PV_wanted=penFactor_commercial*commercial_power# energia PV comercial desejada    
                if len(residential_loads_matrix) > 0:                    
                    residential_PVLoads_NameList, residential_PVLoads_numModulesList=self.raffle_PV_Loads(residential_nameList, residential_loads_matrix, residential_PV_wanted, Pmpp, module_energy)
                if len(commercial_loads_matrix) > 0:
                    commercial_PVLoads_NameList, commercial_PVLoads_numModulesList=self.raffle_PV_Loads(commercial_nameList, commercial_loads_matrix, commercial_PV_wanted, Pmpp, module_energy) 
                #Cria PV 
                self.create_PVSystem(folder_path, residential_PVLoads_NameList, residential_PVLoads_numModulesList, commercial_PVLoads_NameList, commercial_PVLoads_numModulesList, Pvst, Eff, Irrad, Tshape, Pmpp, nptos, kV)
                self.compile_DSS() # Compila denovo para o OpenDSS incluir o PVSystem            
                if len(residential_PVLoads_NameList) > 0:            
                    residential_StorageLoads_list = self.raffle_Storage_Loads(residential_PVLoads_NameList, penFactor_storage_residential)
                if len(commercial_PVLoads_NameList) > 0:
                    commercial_StorageLoads_list = self.raffle_Storage_Loads(commercial_PVLoads_NameList, penFactor_storage_commercial)
                #Cria Storage
                self.create_storage(folder_path, residential_StorageLoads_list, commercial_StorageLoads_list, storagekW_percentage) 
                self.compile_DSS() # Compila denovo para o OpenDSS incluir o Storage
                #plota tensões potências e perdas nas linhas para uma simulacao
                # if i==0:
                #     self.save_monitors_profiles(folder_analysis, j+1)
                #Transgressoes de Tensao
                num_overVolt, num_underVolt, underV_report, overV_report = self.voltageException(parameter_overV, parameter_subV)
                underV_report_tot.append((penFactor_residential, penFactor_commercial, penFactor_storage_residential, penFactor_storage_commercial, storagekW_percentage, underV_report))
                overV_report_tot.append((penFactor_residential, penFactor_commercial, penFactor_storage_residential, penFactor_storage_commercial, storagekW_percentage, overV_report))
                num_overVolt_list.append((num_overVolt)) 
                num_underVolt_list.append((num_underVolt)) 
                #perdas                
                self.dssMeters.Name= "m1" #setando o registrador
                substation_energy, total_losses, line_losses, transf_losses= self.dssMeters.RegisterValues[0], self.dssMeters.RegisterValues[12], self.dssMeters.RegisterValues[22], self.dssMeters.RegisterValues[23]
                substation_energy_list.append(substation_energy)
                total_losses_list.append(total_losses)
                line_losses_list.append(line_losses)
                transf_losses_list.append(transf_losses)
                f.write("%d      %d                          %d                        %d             %d               %.2f                  %.2f                   %.2f \n" % (
                    i+1, int(sum(residential_PVLoads_numModulesList)), int(sum(commercial_PVLoads_numModulesList)), num_overVolt, num_underVolt, total_losses, line_losses, transf_losses))
                # print("sum(residential_PVLoads_numModulesList)", sum(residential_PVLoads_numModulesList))
                # print("sum(commercial_PVLoads_numModulesList)", sum(commercial_PVLoads_numModulesList))
                # print("overVolt: ", num_overVolt ,"underVolt: ", num_underVolt)                
                # print("total_losses: "+str(total_losses))
                # print("substation_energy: " +str(substation_energy))
            f.write("\n")
            # media e desvio padrao das perdas e da energia das cargas
            mean_substation_energy = np.mean(np.array(substation_energy_list))            
            std_substation_energy = np.std(np.array(substation_energy_list))
            mean_total_losses = np.mean(np.array(total_losses_list))
            std_total_losses = np.std(np.array(total_losses_list))  # desvio padrao
            mean_transf_losses = np.mean(np.array(transf_losses_list))
            std_transf_losses = np.std(np.array(transf_losses_list))
            mean_line_losses = np.mean(np.array(line_losses_list))
            std_line_losses = np.std(np.array(line_losses_list))
            mean_num_overVolt = np.mean(np.array(num_overVolt_list))
            std_num_overVolt = np.std(np.array(num_overVolt_list))
            mean_num_underVolt = np.mean(np.array(num_underVolt_list))
            std_num_underVolt = np.std(np.array(num_underVolt_list))            
            print("mean_num_overVol: ", mean_num_overVolt ,"mean_num_underVolt: ", mean_num_underVolt)                
            print("mean_total_losses: "+str(mean_total_losses))
            print("mean_substation_energy: " +str(mean_substation_energy))
            
            # matriz com medias e desvios padrao de perdas e transgressoes de tensao p/ cada combinacao de penetracoes
            report_matrix[j][:] = [mean_substation_energy, std_substation_energy, mean_total_losses, std_total_losses,
                                      mean_transf_losses, std_transf_losses, mean_line_losses, std_line_losses, mean_num_overVolt, std_num_overVolt, mean_num_underVolt, std_num_underVolt]
            j = j+1
        f.close()
        return report_matrix, underV_report_tot, overV_report_tot
        
    def plot_Statistical(self, folder_analysis, num_simulations, penFactor_list, report_matrix, underV_report_tot, overV_report_tot):        
        penFactor_residential=[]
        mean_total_losses, std_total_losses, mean_transf_losses = np.zeros(
            len(penFactor_list)), np.zeros(len(penFactor_list)), np.zeros(len(penFactor_list))
        std_transf_losses, mean_line_losses, std_line_losses = np.zeros(
            len(penFactor_list)), np.zeros(len(penFactor_list)), np.zeros(len(penFactor_list))
        mean_substation_energy, std_substation_energy, mean_num_overVolt, std_num_overVolt, mean_num_underVolt, std_num_underVolt = np.zeros(
            len(penFactor_list)), np.zeros(len(penFactor_list)), np.zeros(len(penFactor_list)), np.zeros(len(penFactor_list)), np.zeros(len(penFactor_list)), np.zeros(len(penFactor_list))
        for i in range(0, len(penFactor_list)):
            mean_substation_energy[i] = report_matrix[i][0]
            std_substation_energy[i] = report_matrix[i][1]
            mean_total_losses[i] = report_matrix[i][2]
            std_total_losses[i] = report_matrix[i][3]
            mean_transf_losses[i] = report_matrix[i][4]
            std_transf_losses[i] = report_matrix[i][5]
            mean_line_losses[i] = report_matrix[i][6]
            std_line_losses[i] = report_matrix[i][7]
            mean_num_overVolt[i] = report_matrix[i][8]
            std_num_overVolt[i] = report_matrix[i][9]
            mean_num_underVolt[i] = report_matrix[i][10]
            std_num_underVolt[i] = report_matrix[i][11]
        for (penFactor_resid, penFactor_commercial, penFactor_storage_residential, penFactor_storage_commercial, storagekW_percentage) in penFactor_list: 
            penFactor_residential.append((penFactor_resid))        
        table_values= 100*np.array(penFactor_list)
        table_values=table_values.transpose()
        columns = ('Case 1', 'Case 2', 'Case 3', 'Case 4')
        rows=["Residential PV Pen(%)", "Commercial PV Pen(%)", "Residential BATT Pen(%)", "Commercial BATT Pen(%)", "%BATT Power"]
        #Total Losses
        # Add a table at the bottom of the axes
        plt.figure()
        plt.table(cellText=table_values,
                            rowLabels=rows,                        
                            colLabels=columns,
                            loc='bottom')
        # Adjust layout to make room for the table:
        plt.subplots_adjust(left=0.2, bottom=0.2)
        plt.title("Total losses x Penetration Factors, %d simulations" %
                    num_simulations)   
                    
        plt.plot([0.2,0.4,0.6,0.8], mean_total_losses, 'o-', color='red', label='Perdas Totais')
        plt.ylabel("Mean of Total Losses")
        plt.yticks(mean_total_losses, fontsize=10)
        plt.xticks([])
        text = 'Standard Deviations:' + str(np.around(std_total_losses, 1))
        plt.text(np.amax([0.2,0.4,0.6,0.8])/2 +0.01, np.percentile(mean_total_losses, 80)+0.01,
                text, bbox=dict(facecolor='white', alpha=0.5)) 
        plt.legend(loc='best')  
        plt.savefig(folder_analysis+'/Total_Losses.png', bbox_inches='tight')
        #plt.show()
        # Transformer Losses
        plt.figure()
        plt.table(cellText=table_values,
                            rowLabels=rows,                        
                            colLabels=columns,
                            loc='bottom')
        # Adjust layout to make room for the table:
        plt.subplots_adjust(left=0.2, bottom=0.2)
        plt.title("Transformers losses x Penetration Factors, %d simulations" %
                    num_simulations)   
        plt.plot([0.2,0.4,0.6,0.8], mean_transf_losses, 'o-', color='red', label='Transformer Losses')
        plt.ylabel("Mean of Transformers Losses")
        plt.yticks(mean_transf_losses, fontsize=10)
        plt.xticks([])
        text = 'Standard Deviations:' + str(np.around(std_transf_losses, 1))
        plt.text(np.amax([0.2,0.4,0.6,0.8])/2 +0.01, np.percentile(mean_transf_losses, 80)+0.01,
                text, bbox=dict(facecolor='white', alpha=0.5)) 
        plt.legend(loc='best')  
        plt.savefig(folder_analysis+'/Transf_Losses.png', bbox_inches='tight')
        
        # Line Losses
        plt.figure()
        plt.table(cellText=table_values,
                            rowLabels=rows,                        
                            colLabels=columns,
                            loc='bottom')
        # Adjust layout to make room for the table:
        plt.subplots_adjust(left=0.2, bottom=0.2)
        plt.title("Lines Losses x Penetration Factors, %d simulations" %
                    num_simulations)   
        # plt.plot(penFactor_residential, mean_line_losses, 'o-', color='red', label='Lines Losses')
        plt.plot([0.2,0.4,0.6,0.8], mean_line_losses, 'o-', color='red', label='Lines Losses')
        plt.ylabel("Mean of Lines Losses")
        plt.yticks(mean_line_losses, fontsize=10)
        plt.xticks([])
        text = 'Standard Deviations:' + str(np.around(std_line_losses, 1))
        plt.text(np.amax([0.2,0.4,0.6,0.8])/2 +0.01, np.percentile(mean_line_losses, 80)+0.01,
                text, bbox=dict(facecolor='white', alpha=0.5)) 
        plt.legend(loc='best')  
        plt.savefig(folder_analysis+'/Line_Losses.png', bbox_inches='tight')
        
        # Overvoltage
        plt.figure()
        plt.table(cellText=table_values,
                            rowLabels=rows,                        
                            colLabels=columns,
                            loc='bottom')
        # Adjust layout to make room for the table:
        plt.subplots_adjust(left=0.2, bottom=0.2)
        plt.title("Number of Overvoltages x Penetration Factors, %d simulations" %
                    num_simulations)                  
        plt.plot([0.2,0.4,0.6,0.8], mean_num_overVolt, 'o-', color='red', label='Overvoltage Occurences')        
        plt.ylabel("Mean of Overvoltages")
        plt.yticks(mean_num_overVolt, fontsize=10)
        plt.xticks([])
        text = 'Number of Overvoltages:' + str(np.around(std_num_overVolt, 1))
        plt.text(np.amax([0.2,0.4,0.6,0.8])/2, np.percentile(mean_num_overVolt, 80),text, bbox=dict(facecolor='white', alpha=0.5)) 
        plt.legend(loc='best')  
        plt.savefig(folder_analysis+'/overVolt.png', bbox_inches='tight') 
        # Undervoltage
        plt.figure()
        plt.table(cellText=table_values,
                            rowLabels=rows,                        
                            colLabels=columns,
                            loc='bottom')
        # Adjust layout to make room for the table:
        plt.subplots_adjust(left=0.2, bottom=0.2)
        plt.title("Number of Undervoltages x Penetration Factors, %d simulations" %
                    num_simulations)                
        plt.plot([0.2,0.4,0.6,0.8], mean_num_underVolt, 'o-', color='red', label='Undervoltage Occurences')        
        plt.ylabel("Number of Undervoltages")
        plt.yticks(mean_num_underVolt, fontsize=10)
        plt.xticks([])
        text = 'Standard Deviations:' + str(np.around(std_num_underVolt, 1))
        plt.text(np.amax([0.2,0.4,0.6,0.8])/2, np.percentile(mean_num_underVolt, 80),text, bbox=dict(facecolor='white', alpha=0.5)) 
        plt.legend(loc='best')  
        plt.savefig(folder_analysis+'/underVolt.png', bbox_inches='tight')       
        # Substation Energy
        plt.figure()
        plt.table(cellText=table_values,
                            rowLabels=rows,                        
                            colLabels=columns,
                            loc='bottom')
        # Adjust layout to make room for the table:
        plt.subplots_adjust(left=0.2, bottom=0.2)
        plt.title("Total Feeder Energy x Penetration Factors, %d simulations" %
                    num_simulations)         
        plt.plot([0.2,0.4,0.6,0.8], mean_substation_energy, 'o-', color='red', label='Feeder Energy')
        plt.ylabel("Mean of feeder Energy")
        plt.yticks(mean_substation_energy, fontsize=10)
        plt.xticks([])
        text = 'Standard Deviations:' + str(np.around(std_substation_energy, 1))
        plt.text(np.amax([0.2,0.4,0.6,0.8])/2 +0.01, np.percentile(mean_substation_energy, 80)+0.01,
                text, bbox=dict(facecolor='white', alpha=0.5)) 
        plt.legend(loc='best')  
        plt.savefig(folder_analysis+'/substation_energy.png', bbox_inches='tight')
        plt.show()

    def customized_Analysis(self, folder_path, folder_analysis, parameter_overV, parameter_subV, penFactor_residential, penFactor_commercial, penFactor_storage_residential, penFactor_storage_commercial, storagekW_percentage):
        residential_PVLoads_NameList, commercial_PVLoads_NameList = [], []
        residential_PVLoads_numModulesList, commercial_PVLoads_numModulesList = [], []
        residential_StorageLoads_list, commercial_StorageLoads_list = [], []
        residential_nameList, residential_loads_matrix, commercial_nameList, commercial_loads_matrix = self.get_LoadsPowerAndEnergybyClass( loadshape_Namelist, ls_classification)
        residential_power, residential_energy, commercial_power, commercial_energy = self.get_classes_power_and_energy(residential_loads_matrix, commercial_loads_matrix)
        ##circuito original
        self.compile_DSS()
        #Transgressoes de Tensao
        num_overVolt_original, num_underVolt_original, underV_report_original, overV_report_original = self.voltageException(parameter_overV, parameter_subV)
        #perdas
        total_losses_original, line_losses_original, transf_losses_original = self.dssMeters.RegisterValues[12], self.dssMeters.RegisterValues[22], self.dssMeters.RegisterValues[23]
        ##circuito com PV   
        Pvst, Eff, Irrad, Tshape, Pmpp, nptos, kV= parameters_Module() # parametros modulo PV     
        module_energy=self.Module_Energy(folder_path, Pvst, Eff, Irrad, Tshape, Pmpp, nptos, kV) #energia de um modulo
        print("module_energy",module_energy)             
        penFactor_str="nominal power"
        residential_PV_wanted=penFactor_residential*residential_power # energia PV residencial desejada
        commercial_PV_wanted=penFactor_commercial*commercial_power# energia PV comercial desejada    
        #sorteio de cargas para PV           
        if len(residential_loads_matrix) > 0:            
            residential_PVLoads_NameList, residential_PVLoads_numModulesList=self.raffle_PV_Loads(residential_nameList, residential_loads_matrix, residential_PV_wanted, Pmpp, module_energy)
        if len(commercial_loads_matrix) > 0:
            commercial_PVLoads_NameList, commercial_PVLoads_numModulesList=self.raffle_PV_Loads(commercial_nameList, commercial_loads_matrix, commercial_PV_wanted, Pmpp, module_energy)
            
        self.create_PVSystem(folder_path, residential_PVLoads_NameList, residential_PVLoads_numModulesList, commercial_PVLoads_NameList, commercial_PVLoads_numModulesList, Pvst, Eff, Irrad, Tshape, Pmpp, nptos, kV)
        self.compile_DSS() # Compila denovo para o OpenDSS incluir o PVSystem
        #Transgressoes de Tensao
        num_overVolt_PV, num_underVolt_PV, underV_report_PV, overV_report_PV = self.voltageException(parameter_overV, parameter_subV)
        #perdas
        total_losses_PV, line_losses_PV, transf_losses_PV = self.dssMeters.RegisterValues[12], self.dssMeters.RegisterValues[22], self.dssMeters.RegisterValues[23]
        ##circuito com PV e Storage 
        #sorteio de cargas para Storage        
        if len(residential_PVLoads_NameList) > 0:            
            residential_StorageLoads_list = self.raffle_Storage_Loads(residential_PVLoads_NameList, penFactor_storage_residential)
        if len(commercial_PVLoads_NameList) > 0:
            commercial_StorageLoads_list = self.raffle_Storage_Loads(commercial_PVLoads_NameList, penFactor_storage_commercial)
        self.create_storage(folder_path, residential_StorageLoads_list, commercial_StorageLoads_list, storagekW_percentage)       
        self.compile_DSS() # Compila denovo para o OpenDSS incluir o PVSystem e a bateria
        #Transgressoes de Tensao
        num_overVolt_PVandStorage, num_underVolt_PVandStorage, underV_report_PVandStorage, overV_report_PVandStorage = self.voltageException(parameter_overV, parameter_subV)
        #perdas
        total_losses_PVandStorage, line_losses_PVandStorage, transf_losses_PVandStorage = self.dssMeters.RegisterValues[12], self.dssMeters.RegisterValues[22], self.dssMeters.RegisterValues[23]
        f=open(folder_analysis + "\Customized_Analysis_summary.txt", "w") # report com resumo de perdas e transgressoes de tensao p/ circuito origal, com PV, com PV e Storage       
        #ciruito original (sem PV e sem Storage)
        f.write("Residential loads: Number: %d (%.1f%%) - class nominal Power (kW): %.2f (%.1f%%) - class absorverd energy (kWh): %.2f (%.1f%%) \n"%(len(residential_nameList), residential_number_per, residential_power, residential_power_per, residential_energy, residential_energy_per))
        f.write("commercial loads: Number: %d (%.1f%%) - class nominal Power (kW): %.2f (%.1f%%) - class absorverd energy (kWh): %.2f (%.1f%%) \n"%(len(commercial_nameList), commercial_number_per, commercial_power,commercial_power_per, commercial_energy, commercial_energy_per))
        f.write("\n###Original Circuit\n")
        f.write("Line Losses (kWh) = %.2f\n"%line_losses_original)
        f.write("Trafo Losses (kWh) = %.2f\n"%transf_losses_original)
        f.write("Total Losses (kWh) = %.2f\n"%total_losses_original)       
        f.write("\nTotal number of Undervoltage (V< %.2f pu) ocurrences: %d\n"%(parameter_subV, num_underVolt_original))     
        f.write("Total number of Overvoltage (V> %.2f pu) ocurrences: %d\n"%(parameter_overV, num_overVolt_original))    
        #circuito com PV
        f.write("\n###Circuit with PV\n")
        f.write("PV Penetration Factor residential: %.2f %% of residential %s\n"%(penFactor_residential*100, penFactor_str))
        f.write("PV Penetration Factor commercial: %.2f %% of commercial %s\n"%(penFactor_commercial*100, penFactor_str))        
        f.write("Line Losses (kWh) = %.2f\n"%line_losses_PV)
        f.write("Trafo Losses (kWh) = %.2f\n"%transf_losses_PV)
        f.write("Total Losses (kWh) = %.2f\n"%total_losses_PV)
        f.write("\nTotal number of Undervoltage (V< %.2f pu) ocurrences: %d\n"%(parameter_subV,num_underVolt_PV))
        f.write("Total number of Overvoltage (V> %.2f pu) ocurrences: %d\n"%(parameter_overV,num_overVolt_PV))
        #circuito com PV e Storage
        f.write("\n###Circuit with PV and Storage\n")
        f.write("storage kW = %.2f percent of connected load power\n"%storagekW_percentage)
        f.write("Line Losses (kWh)= %.2f\n"%line_losses_PVandStorage)
        f.write("Trafo Losses (kWh)= %.2f\n"%transf_losses_PVandStorage)
        f.write("Total Losses (kWh)= %.2f\n"%total_losses_PVandStorage) 
        f.write("\nTotal number of Undervoltage (V< %.2f pu) ocurrences: %d\n"%(parameter_subV, num_underVolt_PVandStorage))
        f.write("Total number of Overvoltage (V> %.2f pu) ocurrences: %d\n"%(parameter_overV, num_overVolt_PVandStorage))    
        f.close()
        f=open(folder_analysis + "\Customized_Analysis_VoltagesExceptions.txt", "w") #report detalhado de transgressoes de tensao
        #ciruito original (sem PV e sem Storage)
        f.write("\n\n###Original Circuit\n")
        f.write("\n##Voltage exceptions:\n")
        f.write("\n#Undervoltage (V< %.2f pu):\n"%parameter_subV)
        f.write("\n#Total number of Undervoltage ocurrences: %d\n"%num_underVolt_original)    
        f.write("element name, hours in phase1,  hours in phase2, hours in phase3\n")
        for (load_name, h_phase1,  h_phase2, h_phase3) in underV_report_original:
            f.write("%s, %d, %d, %d\n"%(load_name, h_phase1,  h_phase2, h_phase3))
        f.write("\n#Overvoltage (V> %.2f pu):\n"%parameter_overV)
        f.write("\n#Total number of Overvoltage ocurrences: %d\n"%num_overVolt_original)
        f.write("element name, hours in phase1,  hours in phase2, hours in phase3\n")
        for (load_name, h_phase1,  h_phase2, h_phase3) in overV_report_original:
            f.write("%s, %d, %d, %d\n"%(load_name, h_phase1,  h_phase2, h_phase3))  
        #circuito com PV
        f.write("\n###Circuit with PV\n")
        f.write("\n##Voltage exceptions:\n")
        f.write("\n#Undervoltage (V< %.2f pu):\n"%parameter_subV)
        f.write("\n#Total number of Undervoltage ocurrences: %d\n"%num_underVolt_PV)
        f.write("element name, hours in phase1,  hours in phase2, hours in phase3\n")
        for (load_name, h_phase1,  h_phase2, h_phase3) in underV_report_PV:
            f.write("%s, %d, %d, %d\n"%(load_name, h_phase1,  h_phase2, h_phase3))
        f.write("\n#Overvoltage (V> %.2f pu):\n"%parameter_overV)
        f.write("\n#Total number of Overvoltage ocurrences: %d\n"%num_overVolt_PV)
        f.write("element name, hours in phase1,  hours in phase2, hours in phase3\n")
        for (load_name, h_phase1,  h_phase2, h_phase3) in overV_report_PV:
            f.write("%s, %d, %d, %d\n"%(load_name, h_phase1,  h_phase2, h_phase3))
        #circuito com PV e Storage
        f.write("\n###Circuit with PV and Storage\n")
        f.write("\n##Voltage exceptions:\n")
        f.write("\n#Undervoltage (V< %.2f pu):\n"%parameter_subV)
        f.write("\n#Total number of Undervoltage ocurrences: %d\n"%num_underVolt_PVandStorage)
        f.write("element name, time in phase1,  time in phase2, time in phase3\n")
        for (load_name, h_phase1,  h_phase2, h_phase3) in underV_report_PVandStorage:
            f.write("%s, %d, %d, %d"%(load_name, h_phase1,  h_phase2, h_phase3))
        f.write("\n#Overvoltage (V> %.2f pu):\n"%parameter_overV)
        f.write("\n#Total number of Overvoltage ocurrences: %d\n"%num_overVolt_PVandStorage)
        f.write("element name, time in phase1,  time in phase2, time in phase3\n")
        for (load_name, h_phase1,  h_phase2, h_phase3) in overV_report_PVandStorage:
            f.write("%s, %d, %d, %d\n"%(load_name, h_phase1,  h_phase2, h_phase3))
        f.close()

    def print_informacoesGeraisCircuito(self):
        print(u"Nome do elemento Circuit: " + self.get_nome_circuit())
        print("numero de vsources:" + str(self.dssVsources.Count))
        print("numero de barras:" + str(self.dssCircuit.NumBuses))
        print("numero de geradores:" + str(self.dssGenerators.Count))
        print("numero de cargas:" + str(self.dssLoads.Count))
        print("numero de linhas:" + str(self.dssLines.Count))
        print("numero de transformadores:" + str(self.dssTransformers.Count))
        print("numero de capacitores:" + str(self.dssCapacitors.Count))
        print("numero de loadshapes:" + str(self.dssLoadShapes.Count - 1))
        print("voltage bases: ", self.dssSettings.VoltageBases)

    def print_resultadosCircuito(self):
        self.solve_DSS_snapshot()
        self.get_results_power()
        self.get_results_voltage()
        # Informações do elemento Circuit
        p, q = self.get_circuit_power()
        print(u"Nosso exemplo apresenta o nome do elemento Circuit: " +
              self.get_nome_circuit())
        print(u"Fornece Potência Ativa: " + str(p) + " kW")
        print(u"Fornece Potência Reativa: " + str(q) + " kvar \n")
    
    def plot_loadShapes(self, folder_analysis, ls_classification):       
        if not os.path.exists(folder_analysis+"\\loadshapes"): # se nao existe pasta "loadshapes", cria uma
            os.makedirs(folder_analysis+"\\loadshapes")
        loadshape_Namelist = self.dssLoadShapes.allNames
        loadshape_Namelist = loadshape_Namelist[1:len(loadshape_Namelist)] #exclude 'default' loadshape
        if 'myirrad' in loadshape_Namelist:
            loadshape_Namelist = tuple(
                x for x in loadshape_Namelist if x != 'myirrad')  # exclude pvsystem loadshapes creared by this program         
        loadShapes = self.dssLoadShapes
        for loadshape_name in loadshape_Namelist:
            loadShapes.Name = loadshape_name
            loadShape_array=loadShapes.Pmult
            postion = loadshape_Namelist.index(str.lower(loadshape_name))
            loadshape_classification = ls_classification[postion]
            text = "Classification = "
            if loadshape_classification==1:
                text += "Residential"
            elif loadshape_classification==2:
                text += "Commercial"
            elif loadshape_classification == 3:
                text += "Public Light."
            plt.figure()
            plt.text(0.8, 0.8, text, bbox=dict(facecolor='white', alpha=0.5)) 
            plt.title(loadshape_name)
            plt.plot(loadShape_array)
            plt.savefig(folder_analysis+"\\loadshapes\\" + loadshape_name +".png")
            plt.show()
        if len (loadshape_Namelist) ==0:
            raise Exception("No loadshape was found!") 

    def get_loadshapes_names_and_values(self):
        loadshape_Namelist = self.dssLoadShapes.allNames
        loadshape_Namelist = loadshape_Namelist[1:len(loadshape_Namelist)] #exclude 'default' loadshape
        if 'myirrad' in loadshape_Namelist:
            loadshape_Namelist = tuple( 
                x for x in loadshape_Namelist if x != 'myirrad') # exclude pvsystem loadshapes creared by this program 
        loadShape_array_list = []        
        loadShapes = self.dssLoadShapes
        for loadshape_name in loadshape_Namelist:
            loadShapes.Name = loadshape_name
            loadShape_array_list.append(loadShapes.Pmult)
        loadShape_array_list = np.array(loadShape_array_list)
        if len (loadShape_array_list) ==0:
            raise Exception("No loadshape was found!") 
        else:
            return loadshape_Namelist, loadShape_array_list

    def classify_loadshapes_manually(self, loadshape_Namelist, loadShape_array_list):
        classification_array = []
        for i in range(0, len(loadshape_Namelist)):
            name = loadshape_Namelist[i]
            array = loadShape_array_list[i, :]
            time_array = np.arange(0, len(array), 1, dtype=int)
            plt.figure()
            plt.title(name)
            plt.grid(True)
            plt.plot(time_array, array)
            plt.xticks(time_array)
            plt.show()
            type_chosen = int(input(
                " classify %s \n 1: Residencial \n 2: Comercial/Industrial\n 3: Ilum. publica\n" % name))
            classification_array.append(type_chosen)
        return classification_array

    def save_monitors_profiles(self, folder_analysis, penCase):       
        if not os.path.exists(folder_analysis+"\\monitors"): # se nao existe pasta "monitors", cria uma
            os.makedirs(folder_analysis+"\\monitors")
        
        if not os.path.exists(folder_analysis+"\\monitors\\"+"\\voltage"): 
            os.makedirs(folder_analysis+"\\monitors\\"+"\\voltage")
        if not os.path.exists(folder_analysis+"\\monitors\\"+"\\power"): 
            os.makedirs(folder_analysis+"\\monitors\\"+"\\power")
        if not os.path.exists(folder_analysis+"\\monitors\\"+"\\losses"): 
            os.makedirs(folder_analysis+"\\monitors\\"+"\\losses")
        # modes = [0, 1, 9]   
        modes = [1]  
        for mode in modes:            
            # # Cria monitores linhas
            lines_list = self.dssLines.AllNames                
            for (lineName) in lines_list:
                self.create_monitor("line."+lineName, mode, 1)
            transformers_list = self.dssTransformers.AllNames
            # lista dos transformadores, tirando os do PVSystem
            transformers_notPV_list = [
                x for x in transformers_list if not 'pv' in x]
            for (trafoName) in transformers_notPV_list:
                self.create_monitor("transformer."+trafoName, mode, 1)
            self.sampleAllMonitors()
            monitors_list = self.dssMonitors.AllNames
            monitors = self.dssMonitors 
            for (monitor) in monitors_list:                
                if (mode == 0):
                    abreviacao = "V"
                    extenso = "Tensao"
                    unidade = "V"
                    save_path = folder_analysis+"\\monitors\\"+"\\voltage\\"
                if (mode == 1):
                    abreviacao = "P"
                    extenso = "Pontencia"
                    unidade = "kW"
                    save_path = folder_analysis+"\\monitors\\"+"\\power\\"
                if (mode == 9):
                    abreviacao = "loss"
                    extenso = "Perdas"
                    unidade = "W"
                    save_path = folder_analysis+"\\monitors\\"+"\\losses\\"
                monitors.Name = monitor
                # ativando o elemento sendo monitorado
                elementName = monitors.Element
                self.activate_element(elementName)
                # print("monitor = ", monitor)
                # print("elementName = ", elementName)
                nphases = self.dssCktElement.NumPhases
                fig, ax = plt.subplots()
                if(nphases == 1):
                    time_array, ch1 = self.get_MonitorProfile(elementName,  mode)
                    ax.plot(time_array, ch1, label=abreviacao)
                    ax.legend(loc='lower right')
                elif(nphases == 2):
                    time_array, ch1, ch2 = self.get_MonitorProfile(
                        elementName, mode)
                    ax.plot(time_array, ch1, label=abreviacao+"1")
                    ax.legend(loc='lower right')
                    ax.plot(time_array, ch2, label=abreviacao+"2")
                    ax.legend(loc='lower right')
                elif(nphases == 3):
                    time_array, ch1, ch2, ch3 = self.get_MonitorProfile(
                        elementName, mode)
                    ax.plot(time_array, ch1, label=abreviacao+"1")
                    ax.legend(loc='lower right')
                    ax.plot(time_array, ch2, label=abreviacao+"2")
                    ax.legend(loc='lower right')
                    ax.plot(time_array, ch3, label=abreviacao+"3")
                    ax.legend(loc='lower right')
                plt.grid(True)
                plt.title("Perfil de %s do elemento %s - %d" % (extenso, elementName, penCase))
                plt.ylabel("%s (%s)" % (extenso, unidade))
                plt.xlabel("Tempo (horas)")            
                plt.savefig(save_path + elementName + "_"+ str(penCase)+".png")
                plt.close()


if __name__ == "__main__":
    classification_manual = []
    classification_predicted = []
    folder_program=os.getcwd()

    """ --------//------USER INPUTS-------//-------- """
    #parametros PVSystem
    global Irrad
    Pvst= np.array([[0,  25 , 75 , 100],[1.2, 1.0, 0.8,  0.6]])
    Eff = np.array([[0.1 , 0.2 , 0.4 ,1.0],[0.86 , 0.9 , 0.93 , 0.97]])
    Irrad = ([0, 0, 0, 0, 0, 0, .1, .2, .3, .5, .8, .9, 1, 1, .99, .9, .7, .4, .1, 0, 0, 0, 0, 0]) 
    Tshape = ([25, 25, 25, 25, 25, 25, 25, 25, 35, 40, 45, 50, 60, 60, 50, 40, 35, 30, 25, 25, 25, 25, 25, 25])
    Pmpp = 0.975
    nptos = 24
    kV = 0.22
    
    parameter_overV, parameter_subV= 1.05, 0.93 #sobretensão e subtensão    
    #PEN (res_PV, COm_PV, RES_BAT, COM_BAT, BAT_POWER)
    # penFactor_list = [(0.0,0.0,0.00,0.00,0.0), (0.15,0.15,0.0,0.0,0.0), (0.30,0.3,0.0,0.0,0.0), (0.50,0.50,0.0,0.0,0.0)]
    # penFactor_list = [(0.3,0.3,0.00,0.00,0.0), (0.3,0.3,0.15,0.15,1.0), (0.30,0.3,0.30,0.30,1.0), (0.30,0.30,0.5,0.5,1.0)]

    penFactor_list = [(0.6,0.6,0.00,0.00,0.0), (0.6,0.6,0.15,0.15,1.0), (0.60,0.6,0.30,0.30,1.0), (0.60,0.60,0.99,0.99,1.0)]
    
    
    global daily 
    global nHoras
    global stepsize
    num_simulations = 100  # NUMERO DE SIMULACOES p/ cada pen
    daily = True  # simulacao diaria
    nHoras = 24
    stepsize =1

    path_master= r"C:\Users\Fabio\Documents\IC\ckt7\ckt7_daily\Master_ckt7.dss"
    # path_master= (r"C:\Users\Fabio\Documents\IC\13Bus\IEEE13MASTER.dss")     
    # path_master=(r"C:\Users\Fabio\Documents\ibirapuera\Master.dss")
    # path_master=(r"C:\Users\Fabio\Documents\IC\Storage_test\Master.dss")       
    

    folder_path=path_master.rsplit('\\',1)[0]
    
    folder_analysis= folder_path + "\\Analysis"      
    if not os.path.exists(folder_analysis): # se nao existe pasta "Analysis", cria uma
        os.makedirs(folder_analysis)

    obj = DSS(path_master) # Criar um objeto da classe DSS

    print(""" Autor: Fabio Andrade Zagotto \n Data: %s\n """ % date.today())
    print(u"Versão do OpenDSS: " + obj.versao_DSS() + "\n")
    print(folder_analysis)

    obj.reset_arquives(folder_path)
    

    # Resolver o Fluxo de Potência
    obj.compile_DSS()
    obj.solve_DSS()
    # obj.solve_DSS_yearly() 


    """ --------//------PASSO PRELIMINAR: Conversao loadshapes anuais para diarios-------//-------- """
    #obj.convert_LS_AnualtoDaily(folder_path)
    #obj.adjust_files_to_daily_loadshape(folder_path)
    # criar PVSystem.dss e StorageFleet.dss
    # Incluir  PVSystem.dss e StorageFleet.dss no master.dss
    # criar EnergyMeter para trafo do PVSystem no master.dss

    """ --------//------PASSO 1: classificação dos loadshapes -------//-------- """
    loadshape_Namelist, loadShape_array_list = obj.get_loadshapes_names_and_values()
    # # print(loadShape_array_list)
    # # classification_manual = obj.classify_loadshapes_manually(loadshape_Namelist, loadShape_array_list) # classificacao manual
    # # print(classification_manual)
    # # ls_classification=np.array([2, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 1, 1, 1, 1, 1, 1])
    # # train_model(loadShape_array_list, ls_classification, loadshape_Namelist, folder_analysis, folder_program) # treinar modelo kNearestNeibours
    classification_predicted = predict_loadshape_Type(loadShape_array_list, folder_program) # classificao obtida pelo algoritmo
    # # print("classification_predicted", classification_predicted)
    if len(classification_predicted) > 0:
        ls_classification = classification_predicted
    elif len(classification_manual) > 0:
        ls_classification = classification_manual
    
    # """ -------------// ------ INFORMACOES DO CIRCUITO ------// ------------"""
    # obj.print_informacoesGeraisCircuito()
    # # obj.print_resultadosCircuito()
    # # obj.plot_loadShapes(folder_analysis, ls_classification)



    # """ --------//------PASSO 2: Obtencao de energia total de cargas -------//-------- """
    loads_nameList, loads_matrix = obj.get_LoadsPower_and_Energy(folder_path)    
    residential_nameList, residential_loads_matrix, commercial_nameList, commercial_loads_matrix = obj.get_LoadsPowerAndEnergybyClass(folder_path, loadshape_Namelist, ls_classification)
    residential_power, residential_energy, commercial_power, commercial_energy = obj.get_classes_power_and_energy(residential_loads_matrix, commercial_loads_matrix)
    # tensao de base, n de fases, potencia nominal, energia consumida
    commercial_number_per = len(commercial_nameList)/(len(residential_nameList) + len(commercial_nameList))*100 #porcentagem de cargas comerciais
    residential_number_per= len(residential_nameList)/(len(residential_nameList)+ len(commercial_nameList)) *100
    residential_power_per= residential_power/(residential_power+ commercial_power)*100 #porcentagem de potencia residencial
    commercial_power_per= commercial_power/(residential_power+commercial_power)*100
    residential_energy_per= residential_energy/( residential_energy+ commercial_energy)*100
    commercial_energy_per= commercial_energy/( residential_energy+ commercial_energy)*100
    f = open(folder_analysis + "\LoadsConsumption_report.txt", "w")
    f.write("Residential loads: Number: %d (%.1f%%) - class nominal Power (kW): %.2f (%.1f%%) - class absorverd energy (kWh): %.2f (%.1f%%) \n"%(len(residential_nameList), residential_number_per, residential_power, residential_power_per, residential_energy, residential_energy_per))
    f.write("commercial loads: Number: %d (%.1f%%) - class nominal Power (kW): %.2f (%.1f%%) - class absorverd energy (kWh): %.2f (%.1f%%) \n"%(len(commercial_nameList), commercial_number_per, commercial_power,commercial_power_per, commercial_energy, commercial_energy_per))
    f.write("\nload class, load name, base voltage(kV), number of phases, nominal power(kW), absorverd energy(kWh)\n")
    if len(residential_nameList) > 0:
        i=0
        for (load_name) in residential_nameList:
            f.write("Residential, %s, %.2f, %d, %.2f, %.2f\n"%(load_name,residential_loads_matrix[i,0], residential_loads_matrix[i,1], residential_loads_matrix[i,2], residential_loads_matrix[i,3]))
            i=i+1
    if len(commercial_nameList) > 0:
        i=0
        for (load_name) in commercial_nameList:
            f.write("Commercial, %s, %.2f, %d, %.2f, %.2f\n"%(load_name,commercial_loads_matrix[i,0], commercial_loads_matrix[i,1], commercial_loads_matrix[i,2], commercial_loads_matrix[i,3]))
            i=i+1
    f.close()

    # """ --------//------PASSO 3 Opcao 1: Analise customizada-------//-------- """
    # t1_start = time.perf_counter()
    # parameter_overV, parameter_subV= 1.03, 0.95
    
    # penFactor_residential = .10 # penetracao  residencial em porcentagem
    # penFactor_commercial = .10 #penetracao  comercial em porcentagem    
    # penFactor_storage_residential = .50 # porcentagem das cargas residenciais, com PV, com storage 
    # penFactor_storage_commercial = .50 # porcentagem das cargas comercias, com PV, com storage    
    # storagekW_percentage = 0.7 #potencia do armazenamento como porcentagem da potencia da carga
    # obj.customized_Analysis(folder_analysis, parameter_overV, parameter_subV,  penFactor_residential, penFactor_commercial, penFactor_storage_residential, penFactor_storage_commercial, storagekW_percentage)
    # t1_stop = time.perf_counter()
    # print("Elapsed time B:", t1_stop- t1_start)
    
    # """ --------//------PASSO 3 Opcao 2: Analise estatistica de perdas e niveis de tensao-------//-------- """
     # lista de (penetracao PV residencial,penetracao PV comercial, penetracao armazenamento residencial em rel a PV, penetracao armazenamento comercial em rel a PV, %pot_armazenamento em rel a da carga)  
    #penFactor_list = [(0.05,0.05,0.00,0.00,0.7), (0.10,0.10,0.0,0.0,0.7), (0.20,0.20,0.0,0.0,0.7), (0.30,0.30,0.0,0.0,0.7)]  
    t1_start = time.perf_counter()

    report_matrix, underV_report_tot, overV_report_tot = obj.StatisticalAnalysis(folder_path, folder_analysis, parameter_overV, parameter_subV, num_simulations, penFactor_list)
    t1_stop = time.perf_counter()
    print("Elapsed time B:", t1_stop- t1_start)

    # # "parÂmetro: reducao energia gerada"
    # # """ --------//------PASSO 3 Opcao 2.1: Plotagem dos resultados estatisticos-------//-------- """

    obj.plot_Statistical(folder_analysis, num_simulations, penFactor_list, report_matrix, underV_report_tot, overV_report_tot)

    

    # fig, ax = plt.subplots()
    # ax.plot(t_pv, ch1_sist, label= "1")
    # ax.legend(loc='lower right')
    # ax.plot(t_pv, ch2_sist, label= "2")
    # ax.legend(loc='lower right')
    # ax.plot(t_pv, ch3_sist, label= "3")
    # ax.legend(loc='lower right')

    # plt.grid(True)
    # plt.title("Perfil de Potencia do sistema pv + carga")
    # plt.ylabel("Potencia (kW)")
    # plt.xlabel("Tempo (horas)")
    # plt.show()

    # substation_energy, monofasica, trifasica = obj.dssMeters.RegisterValues[1], obj.dssMeters.RegisterValues[62], obj.dssMeters.RegisterValues[61]
    # print("\nEnergias registrador OpenDSS")
    # print("-total:"+str(substation_energy))
    # print("-monofasica:"+str(monofasica))
    # print("-trifasica:"+str(trifasica))

    # Informações do elemento escolhido
    # print (u"Elemento Ativo: " + obj.activate_element("Load.1007732-DC1"))
    # barra1 = obj.get_barras_elemento()
    # print (u"Esse elemento está conectado na barra: " + barra1 )
    # print( u"As tensões nodais desse elemento (V): " + str(obj.get_tensoes_elemento()))
    # print (u"As potências desse elemento (kW) e (kvar): " + str(obj.get_potencias_elemento()) + "\n")
    # loadName= "1007805-DC2"
