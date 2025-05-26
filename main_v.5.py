from pyautocad import Autocad, APoint

class bahia():
    def __init__(self,p0,bil,v_Um,i_n):
        self.p0=p0
        self.bil=bil
        self.v_Um=v_Um
        self.i_n=i_n
        self.salida={ "x":0, "y":200, "Descripcion":{"Nombre": "Hacia SE_XXXX", "Codigo de Linea": "L-XXXX", "Tramo linea (km)":"xx km"}, "block_name":""}
        self.pr={ "x":0, "y":170, "Descripcion":{"Tension nominal (kV)": self.v_Um, "BIL (kVp)": self.bil, "I nominal (A)": self.i_n, "Clase": 0,}, "block_name":""}
        self.tt={ "x":0, "y":140, "Descripcion":{"Tension nominal (kV)": self.v_Um, "BIL (kVp)": self.bil, "Clase de Proteccion": 0, "Tipo": 0},"block_name": ""}
        self.secc1={ "x":0, "y": 110, "Descripcion":{"Tension nominal (kV)": self.v_Um, "BIL (kVp)": self.bil, "I nominal (A)": self.i_n, "Tipo": 0}, "block_name":""}
        self.tc={ "x":0, "y": 85, "Descripcion":{"Tension nominal (kV)": self.v_Um, "BIL (kVp)": self.bil, "I nominal (A)": self.i_n, "Tipo": 0, "Relacion devanado primario": 0, "Relacion devanado secundario": 0, "Burden": 0}, "block_name":""}
        self.ip={ "x":0, "y": 50, "Descripcion":{"Tension nominal (kV)": self.v_Um, "BIL (kVp)": self.bil, "I nominal (A)": self.i_n}, "block_name":""}
        self.secc2={ "x":0, "y": 20, "Descripcion":{"Tension nominal (kV)": self.v_Um, "BIL (kVp)": self.bil, "I nominal (A)": self.i_n, "Tipo": 0}, "block_name":""}
        self.elementos=[self.salida,self.pr,self.tt,self.secc1,self.tc,self.ip,self.secc2]
        self.input_usuario()
    
    def input_usuario(self):
        self.elegir_salida()
        self.elegir_pr()
        self.elegir_tt()
        self.elegir_secc1()
        self.elegir_tc()
        self.elegir_ip()
        self.elegir_secc2()
        
    def elegir_salida(self):
        self.salida["block_name"] = "SALIDA_LINEA"    
    
    def elegir_pr(self):
        self.pr["block_name"] ="Pararrayos c-cd1"
    def elegir_tt(self):
        while True:
            print("\nElija el tipo de transformador de tension:\n1) Inductivo 2 dev.\n2) Capacitivo 1 dev.\n3) Capacitivo 2 dev.\n4) Capacitivo 3 dev.")
            val=input("-->")
            if val=="1":
                print("Selecciono TT Inductivo 2 dev")
                self.tt["block_name"] ="TTI_AT_2S"
                break
            if val=="2":
                print("Selecciono TT Capacitivo 1 dev.")
                self.tt["block_name"] ="TTC_AT_1S"
                break
            if val=="3":
                print("Selecciono TT Capacitivo 2 dev.")
                self.tt["block_name"] ="TTC_AT_2S"
                break
            if val=="4":
                print("Selecciono TT Capacitivo 3 dev.")
                self.tt["block_name"] ="TTC_AT_3S"
                break
            else:
                print("Eliga la opcion nuevamente")    
    def elegir_secc1(self):
        self.secc1["block_name"] ="SL_AT_AV"

    def elegir_tc(self):
        while True:
            print("\nElija el tipo de transformador de corriente:\n1) tres devanados\n2) cuatro devanados")
            val=input("-->")
            if val=="1":
                print("Selecciono tres devanados")
                self.tc["block_name"] ="CT_AT_3S"
                break
            if val=="2":
                print("Selecciono cuatro devanados")
                self.tc["block_name"] ="CT_AT_4S"
                break
            else:
                print("Eliga la opcion nuevamente")
        
    def elegir_ip(self):
        self.ip["block_name"] ="IP_AT"

    def elegir_secc2(self):
        self.secc2["block_name"] = "SB_AT_AC"

    def dibujar(self):
        
        acad = Autocad(create_if_not_exists=True, visible=True)
        doc = acad.ActiveDocument
        lib_objetos = doc.blocks
        listado = [elem.name for elem in lib_objetos]

        defined_width=80
        separacion_derecha_eje=30
        for instance in self.elementos:  # Use class variable
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
        "Tensión Um (kV)": 245,
            "Frecuencia (Hz)": 60,
            "Corriente nominal (A)": 2500,
            "Corriente de cortocircuito (kA)": 40,
            "Nivel de aislamiento (BIL)": 1050}


    #---------------------- Definir parametros aplicables-------------------------
    bil=dict_sistema_electrico.get("Nivel de aislamiento (BIL)")
    v_Um=dict_sistema_electrico.get("Tensión Um (kV)")
    i_n=dict_sistema_electrico.get("Corriente nominal (A)")

    #-------------------- Limpiar AutoCad ----------------------
    clear_autocad()  # Limpiar AutoCAD antes de dibujar  

    #-------------------- Establecer condiciones ----------------
    bahia1=bahia(p0=[0,0],
                 bil=bil,
                 v_Um=v_Um,
                 i_n=i_n,
                )
    #----------------------Dibujar-------------------------------

    #Dibujar bahia simple barra up
    bahia1.dibujar()
    

if __name__ == "__main__":
    main()
'''
Generar ejecutable

instalar:  pip install pyinstaller
en cmd , ir a ruta donde se encuentra el archivo python:  cd "C:\Users\BMALPARTIDA\Downloads\DU-main (1)\DU-main"
en cmd escribir: pyinstaller --onefile main_v.5.py
'''