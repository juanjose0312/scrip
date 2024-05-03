import os
import shutil
import pandas as pd
import win32com.client as win32
from script_xlsx_sql import Script_xlsx_sql

class Extraer_tablas_dinamicas:
    def __init__(self, url_completa, diccionario_filtros_valor, nombre_hoja, rango_hoja, rango_inicio_tabla):

        # que tenga menos inputs

        self.url = url_completa
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.wb = self.excel.Workbooks.Open(self.url)
        self.nombre_hojas_inicio = []
        self.nombre_hojas_fin = []
        self.nombre_hoja = nombre_hoja
        self.rango_hoja = rango_hoja
        self.rango_inicio_tabla = rango_inicio_tabla
        self.diccionario_filtros_valor = diccionario_filtros_valor
        self.pvtTable = self.wb.Sheets(nombre_hoja).Range(rango_hoja).PivotTable

    def get_sheet_names(self):

        # meto que se encarga de decir el nombre de las hojas
        lista = []
        for sheet in self.wb.Worksheets:
            lista.append(sheet.Name)
        return lista
    
    def process_excel(self):
        # guarda el nombre de las hojas del libro de excel
        self.nombre_hojas_inicio = self.get_sheet_names()

        print("\033[1;30;42m Aviso : inicio \033[0m")

        # limpia los filtros
        for filtro in self.diccionario_filtros_valor:
            self.pvtTable.PivotFields(filtro).ClearAllFilters()

        # selecciona las variable a incluir en el filtro atravez de un diccionario
        for filtro in self.diccionario_filtros_valor:
            for item in self.pvtTable.PivotFields(filtro).PivotItems():
                if item.Name not in self.diccionario_filtros_valor[filtro]:
                    item.Visible = False
        
        print('aviso : ya aplico los filtros')

        # Verificar si la macro existe
        component_exists = False
        for component in self.wb.VBProject.VBComponents:
            if component.Name == "Módulo1":
                component_exists = True
                break
        
        # si el componente no existe, crearlo
        if not component_exists:
            # crear un nuevo modulo de VBA
            excelModule = self.wb.VBProject.VBComponents.Add(1)
            excelModule.CodeModule.AddFromString(f"""
            Sub explotar_tabla()
                Sheets("{self.nombre_hoja }").Select
                Range("{self.rango_inicio_tabla}").Select
                Selection.End(xlToRight).Select
                Selection.End(xlDown).Select
                Selection.ShowDetail = True
            End Sub
            """)

        # ejecutar la macro

        self.excel.Run("Módulo1.explotar_tabla")

        print('aviso : ya ejecuto el macro')

        # Accede al módulo de VBA y lo elimina
        VBAProject = self.wb.VBProject
        VBAComponents = VBAProject.VBComponents
        VBAComponents.Remove(VBAComponents.Item('Módulo1'))

        # guarda el nombre de las hojas del libro de excel
        self.nombre_hojas_fin = self.get_sheet_names()

        # determina la hoja nueva
        hoja_nueva = list(set(self.nombre_hojas_fin) - set(self.nombre_hojas_inicio))

        # guarda y cierra el archivo
        print('aviso : guardando ...')
        print("\033[1;37;41m Alerta : dar a guardar en el archivo de excel y luego aceptar \033[0m")

        self.wb.Close()
        self.excel.Quit()

        print('aviso : guardo y cerro el archivo')

        # lee la hoja nueva
        df = pd.read_excel(self.url, sheet_name=hoja_nueva[0])

        # elimina el contenido de la carpeta
        carpeta = self.url.rsplit('/',1)[0]        
        nombres_archivos = os.listdir(carpeta)

        #try:
        #    for archivo in nombres_archivos:
        #        os.remove(f"{carpeta}/{archivo}")
        #except:
        #    pass

        #print('aviso : elimino los archivos')

        # crea un excel con la informacion
        if 'radicadas' in self.url.lower():
            # guarda la informacion en un archivo xlsx
            df.to_excel(f"correo_3_info/informe_avanza/radicadas_diario.xlsx", index=False)
            print("\033[1;30;42m Aviso : guardo la informacion en un archivo xlsx \033[0m")
        
        if 'avanza' in self.url.lower():
            # guarda la informacion en jun archivo xlsx
            df.to_excel("correo_3_info/informe_avanza/avanza_diario.xlsx", index=False)
            print("\033[1;30;42m Aviso : guardo la informacion en un archivo xlsx \033[0m")
            
        # Script_xlsx_sql('credenciales.json','correo_3_info/informe_avanza/radicadas_diario.xlsx')

        # print('aviso : se subio la informacion a la base de datos')

        return df