from pyautocad import Autocad, APoint

class bahia():
    def __init__(self,metadatos,p0):
        self.elegir_pr()
        self.elegir_tt()
        self.elegir_secc1()
        self.elegir_tc()   
        self.elegir_ip()
        self.elegir_secc2()
        self.metadatos=metadatos
        self.p0=p0

    def elegir_secc2(self):
        self.secc2="SB_AT_AC"

    def elegir_ip(self):
        self.int="IP_AT"

    def elegir_secc1(self):
        self.secc1="SL_AT_AV"

    def elegir_pr(self):
        self.pr="Pararrayos c-cd1"

    def elegir_tt(self):
        while True:
            print("\nElija el tipo de transformador de tension:\n1)Inductivo\n2)Capacitivo")
            val=input("-->")
            if val=="1":
                print("Selecciono TT Inductivo")
                self.tt="TT1_AT_2S"
                break
            if val=="2":
                print("Selecciono TT Capacitivo")
                self.tt="TTC_AT_2S"
                break
            else:
                print("Eliga la opcion nuevamente")    

    def elegir_tc(self):
        while True:
            print("\nElija el tipo de transformador de corriente:\n1) 3 devanados\n2) 4 devanados")
            val=input("-->")
            if val=="1":
                print("Selecciono 3 devanados")
                self.tc="CT_AT_3S"
                break
            if val=="2":
                print("Selecciono 4 devanados")
                self.tc="CT_AT_4S"
                break
            else:
                print("Eliga la opcion nuevamente")
      
    def dibujar(self):
        
        acad = Autocad(create_if_not_exists=True, visible=True)
        doc = acad.ActiveDocument
        lib_objetos = doc.blocks
        listado = [elem.name for elem in lib_objetos]

        defined_width=80
        separacion_derecha_eje=30
        for instance in self.metadatos:  # Use class variable
            block_name = instance.get("block_name")
            if block_name in listado:
                acad.model.InsertBlock(APoint(instance.get("x")+self.p0[0], instance.get("y"), 0), block_name, 1, 1, 1, 0)
                acad.model.AddMText(APoint(instance.get("x")+separacion_derecha_eje+self.p0[0], instance.get("y")+3, 0), defined_width, instance.get("block_name"))             
                # Access the "Descripcion" dictionary
                descripcion = instance.get("Descripcion")
                if isinstance(descripcion, dict):  # Verificar si "Descripcion" is un diccionario
                    for key, value in descripcion.items():  # Iterar a travez de cada par llave-valor 
                        # agregar cada par llave-valor como texto en AutoCAD
                        acad.model.AddMText(APoint(instance.get("x") + separacion_derecha_eje+self.p0[0], instance.get("y") - (list(descripcion.keys()).index(key) * 3), 0), defined_width, f"{key}: {value}")


    def elegir_TC():
        while True:
            print("\nElija el tipo de transformador de corriente:\n1) 3 devanados\n2) 4 devanados")
            val=input("-->")
            if val=="1":
                print("Selecciono 3 devanados")
                break
            if val=="2":
                print("Selecciono 4 devanados")
                break
            else:
                print("Eliga la opcion nuevamente")
        return val

def clear_autocad(): 
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
    metadatos = [{"x": item[0],"y": item[1],"Descripcion": item[2],"block_name": item[3]}for item in lista_sb_up]

    #-------------------- Limpiar AutoCad ----------------------
    clear_autocad()  # Limpiar AutoCAD antes de dibujar  

    #-------------------- Establecer condiciones ----------------
    p0=[0,0]
    bahia1=bahia(metadatos,p0)
    #----------------------Dibujar-------------------------------

    #Dibujar bahia simple barra up
    bahia1.dibujar()
    

if __name__ == "__main__":
    main()