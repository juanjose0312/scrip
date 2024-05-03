import os
import json
import shutil
import pandas as pd
import mysql.connector
from openpyxl import load_workbook
from datetime import datetime, timedelta

class Script_xlsx_sql:
    def __init__(self,ruta_credenciales, ruta_carpeta_a_subir):
        self.ruta_credenciales = ruta_credenciales
        self.ruta_carpeta_a_subir = ruta_carpeta_a_subir

    def cargar_credenciales(self):

        # Cargar las credenciales desde el archivo JSON

        with open(self.ruta_credenciales) as archivo:
            credenciales = json.load(archivo)
        return credenciales
    
    def extraer_fecha(self, nombre_documento):

        # funcion encargada leer la fecha del nombre del documento para devolverla en formato de fecha
        # divide el texto para tener el string de fecha en un valor de la lista
        nombre_documento = nombre_documento.rsplit("_", 2)
        fecha_en_texto = nombre_documento[1]
        fecha = datetime.strptime(fecha_en_texto, '%Y%m%d')
        return fecha.strftime('%Y-%m-%d')
        
#    def formatear_fecha(self, fecha_str):
#
#        # Función para formatear la fecha ingresa un texto con la estructura
#        # YYYYMM y sale en tipo fecha asi YYYY-MM-DD
#
#        if pd.isnull(fecha_str):
#            return None # Retorna None si es NaN
#        
#        fecha = datetime.strptime(str(fecha_str), '%Y%m')
#        return fecha.strftime('%Y-%m-%d')
    
    def convertir_fecha_numerica(self, fecha_numerica):
        if pd.isnull(fecha_numerica):
            return None
        fecha_base = datetime(1899, 12, 30)
        fecha = fecha_base + timedelta(int(fecha_numerica))
        return fecha

    def limpiar_porcentajes(self, texto_porcentaje):

        # funcion para elminar el simbolo de % de los textos

        try:
            # Elimina cualquier símbolo de porcentaje y espacios en blanco, luego convierte a float
            porcentaje_limpiado = float(texto_porcentaje.replace('%', '').strip())
            return porcentaje_limpiado / 100
        except ValueError:
            # Si hay errores de conversión, devuelve NaN (puede manejar estos valores más tarde)
            return pd.NA

    def limpiar_caracteres_no_validos(self, valor):
        # Reemplazar los caracteres no válidos con un espacio en blanco
        valor_limpio = valor.replace("→", "")
        valor_limpio = valor_limpio.replace('','?')
        return valor_limpio

    def formatear_documento(self, ruta, documento):
        fallo = False

        # extraer la extension, nombre del documento y genera el nombre de .xls y de .xlsx
        # para futuros usos de las variables

        extension_documento = documento.rsplit(".",1)[1]
        nombre_documento    = documento.rsplit(".",1)[0]
        archivo_excel_xls   = f'{nombre_documento}.xls'
        archivo_excel_xlsx  = f'{nombre_documento}.xlsx'

        if extension_documento == 'xls':
            try:
                # Leer el archivo .xls en un DataFrame de pandas
                df = pd.read_excel(f'{ruta}/{documento}')

                if 'NOMBRE_FACTURA' in df.columns:
                    df['NOMBRE_FACTURA'] = df['NOMBRE_FACTURA'].apply(self.limpiar_caracteres_no_validos)

                if 'DIRECCION_1' in df.columns:
                    df['DIRECCION_1'] = df['DIRECCION_1'].apply(self.limpiar_caracteres_no_validos)

                # Guardar el DataFrame en un nuevo archivo .xlsx
                df.to_excel(f'{ruta}/{archivo_excel_xlsx}', index=False)

                # informa del proceso que se esta haciendo 
                print("Aviso : archivo Excel convertido guardado como:", archivo_excel_xlsx)

                # Eliminar el archivo original .xls
                os.remove(f'{ruta}/{documento}')

                # informa del proceso que se esta haciendo 
                print("Aviso : archivo original eliminado:", archivo_excel_xls)

            except Exception as e :
                print(f'Fallo el intento de cambiar la extension de documento por {e}')
                fallo = True

        if not fallo:
            # Ruta al archivo Excel
            archivo_excel = f'{ruta}/{archivo_excel_xlsx}'

            # Cargar el libro de trabajo
            libro_trabajo = load_workbook(archivo_excel)

            # Acceder a la hoja activa
            hoja_activa = libro_trabajo.active

            # Iterar sobre todas las celdas de la hoja activa y establecer el formato como "General"
            for fila in hoja_activa:
                for celda in fila:
                    celda.number_format = 'General'

            # Guardar los cambios en el archivo Excel
            libro_trabajo.save(archivo_excel)
            print(f'Aviso : el documento {archivo_excel_xlsx} cuenta con todas su celdas en formato "General"')

        return archivo_excel_xlsx, fallo

    def procesar_archivos(self):
        nombres_archivos = os.listdir(self.ruta_carpeta_a_subir)
        cantidad_de_archivos = len(nombres_archivos)
        cantidad_de_archivos_nosubidos = 0
        
        col_fecha_nf =  ['fecha_alta_cuenta']
        col_fecha_nm =  []
        col_fecha_nfp = ['fecharegistropeticion','fecha_alta_cuenta','fechafactura','fechavencimiento','maximocuenta']
        col_fecha_nmp = ['fecha_venta','fechafactura','fechavencimiento','maximocuenta']
        col_fecha_am =  ['fecha_horaventa','fecha_desbloqueo','fecha_venta','fechaventana','fecha_inicioventa','periodo_legaliza','fechatrafico','fecha_renorepo','fechaentregasim','fechareactivacion']
        col_fecha_af =  ['fecharegistropeticion','fechainiestpetatis','fechainiestadosubpetatis','fechaingresoatiempo','fechainiestadoatiempo','fec_ini_ult_actividad','fec_est_ult_actividad','fechaseguimiento','fechaaltafactelect','fec_comercializacion_cto','fecha_registro_movil_mtotal']

        # recopila el nombre y cantidad de los archivos con los que se va a trabajar
        nombres_archivos = os.listdir(self.ruta_carpeta_a_subir)
        cantidad_de_archivos = len(nombres_archivos)
        cantidad_de_archivos_nosubidos = 0

        # repite cuantos documentos sean para subirlos
        for nombre_archivo in nombres_archivos:
        
            if not ("Radicadas_Consolidado" in nombre_archivo or
                    "Avanza_Detalle" in nombre_archivo or
                    "radicadas_diario" in nombre_archivo or
                    "avanza_diario" in nombre_archivo
                ):
            
                # revisa que tipo de archivo es, en base a eso determina el tipo de mecanismo que usa para sacar la fecha
                try:
                    fecha_reporte = self.extraer_fecha(nombre_archivo)
                except Exception as e :
                    cantidad_de_archivos_nosubidos += 1
                    print(f"El documento {nombre_archivo} no se pudo subi por mal formato en nombre de fecha, revistar el readme en los parametros de nombre")
                    # salta al siguiente ciclo  para que no suba este archivo
                    continue
                
            # Con esta funcion se su extension y el formato que traen sus celdas de la direccion de la primera casilla 
            nombre_archivo, fallo = self.formatear_documento( self.ruta_carpeta_a_subir ,nombre_archivo)

            if fallo:
                continue
            
            # determina en que tabla se va a subir
            if "Fija_Nunca" in nombre_archivo:
                tabla_mysql = 'nopago_fija'
                tipo_producto = 'fija'
                columna_fecha = col_fecha_nf
                columna_a_concatenar_1 = 'id_peticion'
                columna_a_concatenar_2 = 'producto_cg'
                columna_a_concatenar_3 = 'fecha_reporte'

            elif "Movil_Nunca" in nombre_archivo:
                tabla_mysql = 'nopago_movil'
                tipo_producto = 'movil'
                columna_fecha = col_fecha_nm
                columna_a_concatenar_1 = 'cod_cliente'
                columna_a_concatenar_2 = 'celular'
                columna_a_concatenar_3 = 'fecha_reporte'

            elif "Pago_Fija" in nombre_archivo:
                tabla_mysql = 'nopago_fija_preventivo'
                tipo_producto = 'fija'
                columna_fecha = col_fecha_nfp
                columna_a_concatenar_1 = 'id_peticion'
                columna_a_concatenar_2 = 'producto_cg'
                columna_a_concatenar_3 = 'fecha_reporte'

            elif "Pago_Movil" in nombre_archivo:
                tabla_mysql = 'nopago_movil_preventivo'
                tipo_producto = 'movil'
                columna_fecha = col_fecha_nmp
                columna_a_concatenar_1 = 'cod_cliente'
                columna_a_concatenar_2 = 'celular'
                columna_a_concatenar_3 = 'fecha_reporte'

            elif "Radicadas_Consolidado" in nombre_archivo :            
                tabla_mysql = 'alta_movil'
                tipo_producto = 'movil'
                columna_fecha = col_fecha_am
                columna_a_concatenar_1 = 'cod_cliente'
                columna_a_concatenar_2 = 'celular'
                columna_a_concatenar_3 = 'fecha_desbloqueo'

            elif "radicadas_diario" in nombre_archivo :
                tabla_mysql = 'alta_movil_diario'
                tipo_producto = 'movil'
                columna_fecha = col_fecha_am
                columna_a_concatenar_1 = 'cod_cliente'
                columna_a_concatenar_2 = 'celular'
                columna_a_concatenar_3 = 'fecha_desbloqueo'

            elif "Avanza_Detalle" in nombre_archivo:                
                tabla_mysql = 'alta_fija'
                tipo_producto = 'fija'
                columna_fecha = col_fecha_af
                columna_a_concatenar_1 = 'id_peticion'
                columna_a_concatenar_2 = 'producto_hom'
                columna_a_concatenar_3 = 'fechaseguimiento'

            elif "avanza_diario" in nombre_archivo:
                tabla_mysql = 'alta_fija_diario'
                tipo_producto = 'fija'
                columna_fecha = col_fecha_af
                columna_a_concatenar_1 = 'id_peticion'
                columna_a_concatenar_2 = 'producto_hom'
                columna_a_concatenar_3 = 'fechaseguimiento'
            else:
                print(f"Aviso : la tabla {nombre_archivo} no se pudo ubicar")
                continue
            
            # Cargar datos desde el archivo Excel
            df = pd.read_excel(f"{self.ruta_carpeta_a_subir}/{nombre_archivo}")

            # Reestructura el nombre de las colulmnas
            columnas = list(df.columns)
            columnas_formateadas = [valor.replace('%', 'porcentaje') for valor in columnas]
            columnas_formateadas = [valor.replace(' ', '_') for valor in columnas_formateadas]
            df.columns = [col.lower() for col in columnas_formateadas]

            # Aplica la función de limpieza a la columna de porcentajes y fomateando fechas
            if tabla_mysql == 'nopago_movil':
                df['porcentaje_pago'] = df['porcentaje_pago'].apply(self.limpiar_porcentajes)
            elif tabla_mysql == 'alta_fija_diario':
                df = df.drop('autenticacion_correo', axis=1)
                df = df.drop('tipo_venta', axis=1)
                df = df.drop('valor_vision_cliente', axis=1)
                

            # se formatea las columnas que contegan fecha para que pase de expresion numerica a formato de fecha
            for col_fecha in columna_fecha:
                    df[col_fecha] = df[col_fecha].apply(self.convertir_fecha_numerica)

            # Genera 3 nuevas columnas y las rellena
                # fecha de reporte
                # id concatendando 3 columnas para hacer un registro unico
                # tipo de producto con respescto a la db que se este tratando

            if ("Radicadas_Consolidado" in nombre_archivo or
                "Avanza_Detalle" in nombre_archivo or
                "radicadas_diario" in nombre_archivo or
                "avanza_diario" in nombre_archivo):
                df.insert(0, 'fecha_reporte', df[columna_a_concatenar_3])
            else:
                df.insert(0, 'fecha_reporte', fecha_reporte)

            df.insert(1, 'id_unico', df[columna_a_concatenar_1].astype(str) + "-" + df[columna_a_concatenar_2].astype(str) + "-" + df[columna_a_concatenar_3].astype(str))
            df.insert(2, 'tipo_producto', tipo_producto)

            # se rellenan todos los espacios vacios para que se pueda ingresar en la base
            try:
                pd.set_option('future.no_silent_downcasting', True)
                df = df.fillna(0)
            except:
                pass    
            
            # Conectar a la base de datos MySQL
            conexion = mysql.connector.connect(
                host='192.168.2.86',
                user= self.cargar_credenciales()['usuario'],
                password= self.cargar_credenciales()['clave'],
                database='movistardb'
            )

            # se crea un cursor para hacer la query
            cursor = conexion.cursor()

            # Insertar los datos en la tabla
            try:
                for i, fila in df.iterrows():
                    insert_query = f"INSERT IGNORE INTO {tabla_mysql} ({', '.join(df.columns)}) VALUES ({', '.join(['%s' for _ in range(len(df.columns))])})"
                    cursor.execute(insert_query, tuple(fila))

                # Commit para guardar los cambios en la base de datos
                conexion.commit()

                # Mover el archivo de tablas por subir a tablas subidas
                nombre_carpeta = tabla_mysql
                try:
                     shutil.move(f'{self.ruta_carpeta_a_subir}/{nombre_archivo}',f'tablas_subidas/{nombre_carpeta}/{nombre_archivo}')
                except:
                    pass
                # imprime el mensaje que se movio exitosamente
                print(f"Successful : la tabla {nombre_archivo} se a subido con exito")

            except Exception as e:
                cantidad_de_archivos_nosubidos += 1
                print(f"Aviso : la tabla {nombre_archivo} no se pudo subir con exito por el siguiente error: {e}")

        try:
            # intenta cerrar la conexion, si no puede significa que nunca se pudo abrir 

            cursor.close()
            conexion.close()
        except:
            print("Aviso : no se pudo subir niniguna base")

        print(cantidad_de_archivos_nosubidos)
        cantidad_de_archivos_subidos = cantidad_de_archivos-cantidad_de_archivos_nosubidos
        print(f"informacion : de {cantidad_de_archivos} se subieron {cantidad_de_archivos_subidos}")

        # informa que la tabla se subio con exito y informe de tiempo de ejecucion

        # informe final y mensaje de exito
    
    def select_query(self, tabla):
        # Conectar a la base de datos MySQL
        conexion = mysql.connector.connect(
                host='192.168.2.86',
                user= self.cargar_credenciales()['usuario'],
                password= self.cargar_credenciales()['clave'],
                database='movistardb'
        )

        # se crea un cursor para hacer la query
        cursor = conexion.cursor()

        # se hace la query
        cursor.execute(f"SELECT * FROM {tabla}")

        # se obtiene los datos de la query
        datos = cursor.fetchall()

        # se cierra la conexion
        cursor.close()
        conexion.close()

        return datos