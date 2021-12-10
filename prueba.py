# -*- coding: utf-8 -*-
"""
Created on Fri Dec 10 06:27:10 2021

@author: schica
"""
import pandas as pd
import os
import powerfactory
## Funciones #####################################################################################################################################################
restriccionDeTensionSuperiorPorBarras = lambda barra: 1.05 if str(barra.uknom) == '500' else 1.1
def esRestriccionBajaTension(elemento):  
    try:
        esBarra = elemento.GetClassName() == 'ElmTerm'
        esSub = elemento.GetClassName() == 'ElmSubstat'
        if esBarra:  
            esAisladoElElemento = elemento.GetAttribute('m:u') > 0.1
            if esAisladoElElemento:
                esBajaTension = elemento.GetAttribute('m:u') < 0.9
                if esBajaTension:   
                    return True
        elif esSub: 
            esAisladoElElemento = elemento.GetAttribute('m:umin') > 0.1
            if esAisladoElElemento:
                esBajaTension = elemento.GetAttribute('m:umin') < 0.9
                if esBajaTension:
                    return True
    except:
        pass
    return False 
def esRestriccionSobreTension(elemento):  
    try:
        esBarra = elemento.GetClassName() == 'ElmTerm'
        esSub = elemento.GetClassName() == 'ElmSubstat'
        if esBarra:
            esAisladoElElemento = elemento.GetAttribute('m:u') > 0.1
            if esAisladoElElemento:
                esBajaTension = elemento.GetAttribute('m:u') > restriccionDeTensionSuperiorPorBarras(elemento)   
        elif esSub:
            esAisladoElElemento = elemento.GetAttribute('m:umax') > 0.1
            if esAisladoElElemento:
                 esBajaTension = elemento.GetAttribute('m:umax') > restriccionDeTensionSuperiorPorBarras(elemento)
                 if esBajaTension:
                    return True       
    except:
        pass  
    return False 
def esRestriccionCargabilidad(elemento):  
    try:
        esNoBarra = elemento.GetClassName() != 'ElmTerm'
        if esNoBarra:
                esCargabilidad = elemento.GetAttribute('c:loading') > elemento.maxload
                if esCargabilidad: 
                    return True
    except:
        pass
    return False 

## Codigo######################################################################################################################################################
app = powerfactory.GetApplication()
script = app.GetCurrentScript()   
Ruta = script.Ruta
listForDataFrame = []
generador = script.generador
barra = script.barra
app.EchoOff()

#obtiene los objetos de los sets del powerFactory    
for i in script.GetContents():    
    if i.loc_name == 'Casos_Estudio':
        set_Casos_Estudio = i.All()
    elif i.loc_name == 'Perdidas':   
        set_Perdidas = i.All()
    elif i.loc_name == 'Restricciones_Pot':   
        set_Restricciones_Pot = i.All()
    elif i.loc_name == 'Monitorear':
        set_Monitorear = i.All()
        
## Recorrer casos de Estudio del SET
for StudyC in set_Casos_Estudio:
    # Activar caso de estudio
    StudyC.Activate()
    Ldf = app.GetFromStudyCase('*.ComLdf')
    reloj = app.GetFromStudyCase('*.SetTime')
    # guardar nombre del caso de estudio
    casoDeEstudioName = StudyC.loc_name
     # Escenario operativo 
    escenarioActivo = app.GetActiveScenario() 
    # flujo de cargas HORARIO
    for hour in range (24):
        reloj.SetTime(hour)
        Ldf.Execute() 
    # Suma de las perdidas por las lineas de transmision clasificado por nivel de tension
        perdidasLocalesSDL = 0 
        perdidasLocalesSTR = 0
        perdidasLocalesSTN = 0
    
        for elemento in set_Perdidas:
            try:         
                # Condicionales para calculo de perdidas en funcion del nivel de tension
                esSDL = elemento.GetAttribute('t:uline') < 57.5
                esSTR = elemento.GetAttribute('t:uline') >= 57.5 and elemento.GetAttribute('t:uline') < 220
                esSTN = elemento.GetAttribute('t:uline') >= 220
                if esSDL:
                    perdidasLocalesSDL += elemento.GetAttribute('c:Losses')
                elif esSTR:
                    perdidasLocalesSTR += elemento.GetAttribute('c:Ploss')
                elif esSTN:
                    perdidasLocalesSTN += elemento.GetAttribute('c:Ploss')
            except:
                pass     
        tensiones = []
        for tensiones_importantes in set_Monitorear:
            tensiones.append(tensiones_importantes.loc_name + '-' + str(round(tensiones_importantes.GetAttribute('m:u'),3)) + 'pu, ')
        ## Restricciones
        restricciones = []
        esSobreTension = False
        esBajaTension = False
        esCargabilidad = False
        for elemento_para_monitorear in set_Restricciones_Pot:
            try:
                esBarra = elemento_para_monitorear.GetClassName() == 'ElmTerm'
                esSub = elemento_para_monitorear.GetClassName() == 'ElmSubstat'
                esBajaTension = esRestriccionBajaTension(elemento_para_monitorear)
                esSobreTension = esRestriccionSobreTension(elemento_para_monitorear)
                esCargabilidad = esRestriccionCargabilidad(elemento_para_monitorear)
            except:
                pass
            if esSobreTension:
                if esBarra :
                    restricciones.append(elemento_para_monitorear.loc_name + ' Tension-(' + str(elemento_para_monitorear.GetAttribute('m:u')) + ' pu), ')
                elif esSub :
                    restricciones.append(elemento_para_monitorear.loc_name + ' Tension-(' + str(elemento_para_monitorear.GetAttribute('m:umax')) + ' pu), ')
                    
            elif esBajaTension:
                if esBarra:
                    restricciones.append(elemento_para_monitorear.loc_name + ' Tension-(' + str(elemento_para_monitorear.GetAttribute('m:u')) + ' pu), ')
                elif esSub:
                    restricciones.append(elemento_para_monitorear.loc_name + ' Tension-(' + str(elemento_para_monitorear.GetAttribute('m:umin')) + ' pu), ')
            elif esCargabilidad:
                restricciones.append(elemento_para_monitorear.loc_name + ' Cargabilidad-(' + str(elemento_para_monitorear.GetAttribute('c:loading')) + ' %), ') 
        # Agregar los valores hallados a los vectores
        listForDataFrame.append([casoDeEstudioName, hour, perdidasLocalesSDL, perdidasLocalesSTR, perdidasLocalesSTN, barra.GetAttribute('m:u'), generador.GetAttribute('m:Q:bus1'), tensiones, restricciones])
     
    escenarioActivo.DiscardChanges()
        
dataFrame = pd.DataFrame(listForDataFrame, columns=['Caso de estudio', 'Hora', 'Perdidas del SDL', 'Perdidas del STR', 'Perdidas del STN', 'Tensión barra de ' + barra.loc_name, 'potencia reactiva compensador', 'restricciones', 'tensiones zona'])
dataFrame.set_index(['Caso de estudio', 'Hora'], inplace=True)
dataFrame.to_excel(os.path.join(Ruta, 'La Palma Compensacion estatica 7mvar.xlsx'))
app.EchoOn()

   
        
        