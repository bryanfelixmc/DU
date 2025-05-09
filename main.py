from pyautocad import Autocad, APoint

def dibujar(bahia,coord_xy):
    
    acad = Autocad(create_if_not_exists=True, visible=True)
    doc = acad.ActiveDocument
    lib_objetos = doc.blocks
    listado = [elem.name for elem in lib_objetos]

    defined_width=80
    separacion_derecha_eje=30
    for instance in bahia:  # Use class variable
        block_name = instance.get("block_name")
        if block_name in listado:
            acad.model.InsertBlock(APoint(instance.get("x")+coord_xy[0], instance.get("y"), 0), block_name, 1, 1, 1, 0)
            acad.model.AddMText(APoint(instance.get("x")+separacion_derecha_eje+coord_xy[0], instance.get("y")+3, 0), defined_width, instance.get("block_name"))             
            # Access the "Descripcion" dictionary
            descripcion = instance.get("Descripcion")
            if isinstance(descripcion, dict):  # Verificar si "Descripcion" is un diccionario
                for key, value in descripcion.items():  # Iterar a travez de cada par llave-valor 
                    # agregar cada par llave-valor como texto en AutoCAD
                    acad.model.AddMText(APoint(instance.get("x") + separacion_derecha_eje+coord_xy[0], instance.get("y") - (list(descripcion.keys()).index(key) * 3), 0), defined_width, f"{key}: {value}")

def clear_autocad(): ##ok
    try:
        acad = Autocad(create_if_not_exists=True, visible=True)
        doc = acad.ActiveDocument
        for obj in list(doc.ModelSpace):
            obj.Delete()
        print("Cleaning & Ploting again...") 
    except Exception as e:
        print(f"Error al limpiar AutoCAD: {e}")

def main():
    
    #------------------- Datos Generales --------------------

    dict_responsables = {
            "Jefe del Proyecto          ": "Value 1",
            "Elaborado por              ": "Value 2",
            "Revisado por               ": "Value 3",
            "Aprobado por               ": "Value 4"
        }

    dict_informacion_proyecto={
            "Nombre de la Subestación Eléctrica ": "Podras Jamas",
            "Ubicacion - Pais ": "Peru",
            "Ubicacion - Departamento ": "Arequipa",
            "Ubicacion - Provincia ": "Valle Grande",
            "Ubicacion - Distrito ": "Costa Azul",
        }

    dict_condiciones_ambientales= {
            "Altura sobre el nivel del mar   (m.s.n.m.)": 560,
            "Temperatura promedio  (°C) ": 25,
            "Resistividad del terreno (Ω-m)": 103,
            "Nivel de contaminación ": "Severo",
            "Humedad Relativa (%) ": 60
        }

    dict_sistema_electrico={
        "Tensión (kV)": 220,
            "Frecuencia (Hz)": 60,
            "Corriente nominal (A)": 2500,
            "Corriente de cortocircuito (kA)": 40,
            "Nivel de aislamiento (BIL)": 1050}


    #---------------------- Definir parametros aplicables-------------------------
    bil=dict_sistema_electrico.get("Nivel de aislamiento (BIL)")
    v_n=dict_sistema_electrico.get("Tensión (kV)")
    i_n=dict_sistema_electrico.get("Corriente nominal (A)")
    #----------------------Ingresar Bahias-------------------------   
    #Simple barra down
    lista_sb_down=[
                [0, -230, {"Nombre": "Hacia SE_XXXX", "Codigo de Linea": "L-XXXX", "Tramo linea (km)":"xx km"}, "SALIDA_LINEA_DWN"],
                [0,-200, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Clase": 0}, "Pararrayos c-cd1"],
                [0, -160, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "Clase de Proteccion": 0, "Tipo": 0}, "TTC_AT_2S"],
                [0, -120, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Tipo": 0}, "SL_AT_AV_DOWN1"],
                [0, -85, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Tipo": 0, "Relacion devanado primario": 0, "Relacion devanado secundario": 0, "Burden": 0}, "CT_AT_4S"],
                [0, -50, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n}, "IP_AT"],
                [0, -20, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Tipo": 0}, "SB_AT_AC_DOWN"],
            ]

    # Convert to a list of dictionaries
    dict_sb_down = [{"x": item[0],"y": item[1],"Descripcion": item[2],"block_name": item[3]}for item in lista_sb_down]

    #Simple barra up
    lista_sb_up= [
                [ 0, 230, {"Nombre": "Hacia SE_XXXX", "Codigo de Linea": "L-XXXX", "Tramo linea (km)":"xx km"}, "SALIDA_LINEA"],
                [ 0, 200, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Clase": 0,}, "Pararrayos c-cd1"],
                [ 0, 160, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "Clase de Proteccion": 0, "Tipo": 0}, "TTC_AT_2S"],
                [ 0, 120, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Tipo": 0}, "SL_AT_AV_DOWN2"],
                [ 0, 85, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Tipo": 0, "Relacion devanado primario": 0, "Relacion devanado secundario": 0, "Burden": 0}, "CT_AT_4S_DOWN"],
                [ 0, 50, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n}, "IP_AT_DOWN"],
                [ 0, 20, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Tipo": 0}, "SB_AT_AC"],
            ]

    # Convert to a list of dictionaries
    dict_sb_up = [{"x": item[0],"y": item[1],"Descripcion": item[2],"block_name": item[3]}for item in lista_sb_up]

    ''' (BMC:) ESTA PENDIENTE CREAR LOS BLOQUES PARA DOBLE BARRA!
    #Doble barra down
    lista_db_down=[
                [ 0, 280+280-14, {"Nombre": "Hacia SE_XXXX", "Codigo de Linea": "L-XXXX", "Tramo linea (km)":"xx km"}, "SALIDA_LINEA_DWN"],
                [ 18.5, 280+240-14, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Clase": 0}, "Pararrayos c-cd1"],
                [ 10.5, 280+200-14, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "Clase de Proteccion": 0, "Tipo": 0}, "TTC_AT_2S"],
                [ -5, 280+160-14, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Tipo": 0}, "SL_AT_AV_DOWN1"],
                [ 0.5, 280+120-14, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Tipo": 0, "Relacion devanado primario": 0, "Relacion devanado secundario": 0, "Burden": 0}, "CT_AT_4S"],
                [ 9.5, 280+80-14, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n}, "IP_AT"],
                [ 0, 280+40-14, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Tipo": 0}, "SB_AT_AC_DOWN"],
            ]
    # Convert to a list of dictionaries
    dict_db_down = [{"x": item[0],"y": item[1],"Descripcion": item[2],"block_name": item[3]}for item in lista_db_down]

    #Doble barra up

    lista_db_up=[
                [ 0, 40, {"Nombre": "Hacia SE_XXXX", "Codigo de Linea": "L-XXXX", "Tramo linea (km)":"xx km"}, "SALIDA_LINEA"],
                [ 18.5, 80, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Clase": 0,}, "Pararrayos c-cd1"],
                [ 10.5, 120, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "Clase de Proteccion": 0, "Tipo": 0}, "TTC_AT_2S"],
                [ -5, 160, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Tipo": 0}, "SL_AT_AV_DOWN2"],
                [ 0.5, 200, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Tipo": 0, "Relacion devanado primario": v_n, "Relacion devanado secundario": 0, "Burden": 0}, "CT_AT_4S_DOWN"],
                [ 9.5, 240, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n}, "IP_AT_DOWN"],
                [ 0, 280, {"Tension nominal (kV)": v_n, "BIL (kVp)": bil, "I nominal (A)": i_n, "Tipo": 0}, "SB_AT_AC"],
            ]
    # Convert to a list of dictionaries
    dict_db_up = [{"x": item[0],"y": item[1],"Descripcion": item[2],"block_name": item[3]}for item in lista_db_up]
    '''
    #-------------------- Limpiar AutoCad ----------------------
    clear_autocad()  # Limpiar AutoCAD antes de dibujar  

    #----------------------Dibujar-------------------------------
    f=85
    lista=[
           [0,0,"down"],
           [0,0,"up"],
           [1,0,"up"],
           [1,0,"down"],
           [2,0,"up"],
           ]
    for elem in lista:
        if elem[2]=="up":
            #Dibujar bahia simple barra up
            dibujar(bahia=dict_sb_up,coord_xy=[f*elem[0],elem[1]])
        if elem[2]=="down":
            #Dibujar bahia simple barra down
            dibujar(bahia=dict_sb_down,coord_xy=[f*elem[0],elem[1]])       

if __name__ == "__main__":
    main()