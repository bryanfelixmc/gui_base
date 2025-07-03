from pyautocad import Autocad, APoint
from openpyxl import load_workbook

class LectorExcel:
    def __init__(self, ruta):
        self.hoja = load_workbook(ruta, data_only=True)["bahia"]

    def obtener_datos_salida(self):
        return self.hoja["E2"].value

    def obtener_datos_pararrayo(self):
        h = self.hoja
        nombre_equipo = h["A3"].value or "Pararrayo"
        datos = []

        if h["B3"].value:
            datos.append(f"{h['B3'].value} = {h['C3'].value}{h['D3'].value or ''}".strip())
        if h["B4"].value:
            datos.append(f"{h['B4'].value} = {h['C4'].value}{h['D4'].value or ''}".strip())

        parte1 = f"{h['C6'].value}{h['D6'].value}".strip() if h["C6"].value and h["D6"].value else ""
        parte2 = f"{h['B5'].value} {h['C5'].value}".strip() if h["B5"].value and h["C5"].value else ""
        linea_combinada = ", ".join(filter(None, [parte1, parte2]))
        if linea_combinada:
            datos.append(linea_combinada)

        if h["C9"].value and h["D9"].value:
            datos.append(f"{h['C9'].value} {h['D9'].value}".strip())

        return nombre_equipo, datos

    def obtener_datos_ttc(self):
        h = self.hoja
        nombre_equipo = h["A18"].value or "Transformador_Tension"
        datos = []

        if h["C18"].value:
            datos.append(f"{h['C18'].value} {h['D18'].value}, {h['C21'].value} {h['D21'].value or ''}".strip())
        if h["C12"].value:
            datos.append(f"{h['C22'].value} / {h['C23'].value} / {h['C24'].value}".strip())

            parte1 = f"{h['C25'].value}{h['D25'].value}".strip() if h["C25"].value and h["D25"].value else ""
            parte2 = f"{h['B26'].value} {h['C26'].value}".strip() if h["B26"].value and h["C26"].value else ""
            linea_combinada = " - ".join(filter(None, [parte1, parte2]))
            if linea_combinada:
                datos.append(linea_combinada)

            parte1 = f"{h['C25'].value}{h['D25'].value}".strip() if h["C25"].value and h["D25"].value else ""
            parte2 = f"{h['B26'].value} {h['C27'].value}".strip() if h["B26"].value and h["C26"].value else ""
            linea_combinada = " - ".join(filter(None, [parte1, parte2]))
            if linea_combinada:
                datos.append(linea_combinada)

        return nombre_equipo, datos

    def obtener_datos_seccionador(self):
        h = self.hoja
        nombre_equipo = h["A28"].value or "SeccionadorLinea"
        datos = []

        if h["C31"].value:
            datos.append(f"{h['C31'].value} {h['D31'].value}, {h['C32'].value}{h['D32'].value or ''}".strip())
        if h["C28"].value:
            datos.append(f"{h['C28'].value} {h['D28'].value}, {h['C30'].value}{h['D30'].value or ''}".strip())

        return nombre_equipo, datos

    def obtener_datos_transformadorcorriente(self):
        h = self.hoja
        nombre_equipo = h["A37"].value or "Transformador_Corriente"
        datos = []

        if h["C37"].value:
            datos.append(f"{h['C37'].value} {h['D37'].value}, {h['C30'].value} {h['D30'].value or ''}".strip())
        if h["C38"].value:
            datos.append(f"{h['C38'].value} / {h['C39'].value} / {h['C40'].value}".strip())
        if h["C44"].value:
            datos.append(f"{h['C44'].value}{h['D44'].value} , {h['C45'].value} {h['D45'].value}".strip())
        if h["C41"].value:
            datos.append(f"{h['C41'].value}x({h['C46'].value} {h['D46'].value} , {h['C43'].value})".strip())
        if h["C40"].value:
            datos.append(f"{h['C40'].value}x({h['C46'].value} {h['D46'].value} , {h['B42'].value} {h['C42'].value})".strip())

        return nombre_equipo, datos
    
    def obtener_datos_interruptor(self):
        h = self.hoja
        nombre_equipo = h["A47"].value or "Interruptor"
        datos = []
        
        if h["C50"].value:
            datos.append(f"{h['C50'].value} {h['D50'].value}, {h['C51'].value} {h['D51'].value or ''}".strip())
        if h["C47"].value:
            datos.append(f"{h['C47'].value} {h['D47'].value},  {h['C49'].value} {h['D49'].value or ''}".strip())

        return nombre_equipo, datos
    
    def obtener_datos_seccionadorbarra(self):
        h = self.hoja
        nombre_equipo = h["A59"].value or "seccionador barra"
        datos = []

        if h["C62"].value:
            datos.append(f"{h['C62'].value} {h['D62'].value}, {h['C63'].value}  {h['D63'].value or ''}".strip())
        if h["C59"].value:
            datos.append(f"{h['C59'].value} {h['D59'].value},  {h['C61'].value} {h['D61'].value or ''}".strip())

        return nombre_equipo, datos

class AutoCAD:
    def __init__(self):
        self.acad = Autocad(create_if_not_exists=True, visible=True)

    def borrar_todo(self):
        for obj in list(self.acad.ActiveDocument.ModelSpace):
            try:
                obj.Delete()
            except Exception as e:
                print(f"No se pudo borrar un objeto: {e}")

    def poner_bloque_con_texto(self, nombre_bloque, punto_bloque, punto_texto, nombre_equipo, datos):
        if nombre_bloque not in [b.Name for b in self.acad.ActiveDocument.Blocks]:
            print(f"Bloque '{nombre_bloque}' no encontrado")
            return

        # Insertar el bloque
        self.acad.model.InsertBlock(punto_bloque, nombre_bloque, 1, 1, 1, 0)

        # Insertar texto desde posición libre
        x = punto_texto.x
        y = punto_texto.y
        altura_texto = 2.5
        altura_linea = 4

        self.acad.model.AddText(nombre_equipo, APoint(x, y), altura_texto)

        for i, texto in enumerate(datos):
            y_linea = y - (i + 1) * altura_linea
            self.acad.model.AddText(texto, APoint(x, y_linea), altura_texto)

class Proyecto:
    def __init__(self, ruta_excel):
        self.datos = LectorExcel(ruta_excel)
        self.autocad = AutoCAD()

    def correr(self):
        self.autocad.borrar_todo()

        # SALIDA DE LÍNEA
        nombre_sal = self.datos.obtener_datos_salida()
        self.autocad.poner_bloque_con_texto("SALIDA_LINEA", APoint(0, 240), APoint(-30, 250), nombre_sal, [])

        # PARARRAYO
        nombre_p, datos_p = self.datos.obtener_datos_pararrayo()
        self.autocad.poner_bloque_con_texto("Pararrayos c-cd2", APoint(0, 200), APoint(40, 210), nombre_p, datos_p)

        # TTC
        nombre_ttc, datos_ttc = self.datos.obtener_datos_ttc()
        self.autocad.poner_bloque_con_texto("TTC_AT_1S", APoint(0, 160), APoint(40, 160), nombre_ttc, datos_ttc)

        # SECCIONADOR
        nombre_sl, datos_seccionador = self.datos.obtener_datos_seccionador()
        self.autocad.poner_bloque_con_texto("SL_AT_AC", APoint(0, 130), APoint(40, 130), nombre_sl, datos_seccionador)

        # TC
        nombre_tc, datos_transformadorcorriente = self.datos.obtener_datos_transformadorcorriente()
        self.autocad.poner_bloque_con_texto("CT_AT_4S", APoint(0, 100), APoint(40, 100), nombre_tc, datos_transformadorcorriente)

       # IN
        nombre_in, datos_interruptor = self.datos.obtener_datos_interruptor()
        self.autocad.poner_bloque_con_texto("INTERRUPTOR", APoint(0, 60), APoint(40, 60), nombre_in, datos_interruptor)

       # SB
        nombre_sb, datos_seccionadorbarra = self.datos.obtener_datos_seccionadorbarra()
        self.autocad.poner_bloque_con_texto("SB_AT_AC", APoint(0, 30), APoint(40, 30), nombre_sb, datos_seccionadorbarra)

if __name__ == "__main__":
    ruta = r"C:\Users\JGUERRAV\Desktop\INFORMACION\IDI\Base de datos de partida (2) (1).xlsx"
    Proyecto(ruta).correr()
