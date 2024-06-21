import pandas as pd  # Importa la biblioteca Pandas y la asigna al alias 'pd'
import numpy as np  # Importa la biblioteca NumPy y la asigna al alias 'np'
import os  # Importa el módulo os para interactuar con el sistema operativo
import datetime  # Importa el módulo datetime para trabajar con fechas y horas
import warnings  # Importa el módulo warnings para manejar advertencias
from datetime import timedelta  # Importa la clase timedelta del módulo datetime

warnings.filterwarnings("ignore")  # Ignora las advertencias durante la ejecución

# Leemos los feriados desde el archivo (este archivo se debe verificar que este actualizado con los feriados de todo el periodo a evaluar)
directorio_modulo = os.path.dirname(os.path.abspath(__file__))
ruta_relativa_feriados = os.path.join(directorio_modulo, "data", "feriados.xlsx")
feriados = pd.read_excel(ruta_relativa_feriados)['Feriados'].apply(lambda x: x.date()).tolist()

# Parámetros dinámicos
Fecha_Inicio = "Fecha_Inicio"  # Ingrese la fecha de inicio
Fecha_Fin = "Fecha_Fin"     # Ingrese la fecha de fin
Fecha_para_vacios = "17/06/2024" # En caso la celda de fecha fin este vacia, entonces rellenaremos con esta fecha para el calculo de dias habiles

# Creamos nombres para las columnas auxiliares
Fecha_Inicio_AUX = Fecha_Inicio + "_AUX"
Fecha_Fin_AUX = Fecha_Fin + "_AUX"
    
# Ingrese el nombre del archivo de entrada y el nombre de su salida
ruta_relativa_archivo = "prueba.xlsx"
ruta_salida_archivo = "salida.xlsx"

# Función para capturar la hora actual (ayudará a determinar el tiempo de demora para ejecutar el programa)
def obtener_hora_actual():
    return datetime.datetime.now()  # Devuelve la hora actual del sistema

# Definimos la función para calcular los días hábiles entre dos fechas, excluyendo feriados y fines de semana
def calcular_dias_habiles(row):
    return np.busday_count(row[Fecha_Inicio_AUX].date() + timedelta(days=1), row[Fecha_Fin_AUX].date() + timedelta(days=1), holidays=feriados)

def main():
    hora_inicio = obtener_hora_actual()  # Registra la hora de inicio del proceso
    
    df = pd.read_excel(ruta_relativa_archivo)
      
    # Convertimos las columnas originales a objetos datetime y asignamos este valor a las columnas auxiliares, especificando dayfirst=True para evitar la advertencia
    df[Fecha_Inicio_AUX] = pd.to_datetime(df[Fecha_Inicio], dayfirst=True)
    
    # Reemplazamos los valores vacíos en la columna auxiliar de fecha fin por la fecha predeterminada
    df[Fecha_Fin_AUX] = df[Fecha_Fin].fillna(Fecha_para_vacios)
    df[Fecha_Fin_AUX] = pd.to_datetime(df[Fecha_Fin_AUX], dayfirst=True)
    
    # Calculamos la diferencia en días hábiles entre las fechas en las columnas Fecha_inicio y Fecha_fin
    df['Dias_transcurridos(habiles)'] = df.apply(calcular_dias_habiles, axis=1)
    
    # Eliminamos las columnas auxiliares
    df = df.drop (Fecha_Inicio_AUX, axis = 1)
    df = df.drop (Fecha_Fin_AUX, axis = 1)
    
    # Guardamos en nuestro nuevo archivo excel
    df.to_excel(ruta_salida_archivo, index=False)
    
    print("Diferencia en días hábiles calculada y guardada en la columna 'Dias_transcurridos(habiles)'.")
    
    hora_final = obtener_hora_actual()  # Registra la hora de finalización del proceso
    print(f"Creación Exitosa. Proceso terminado. Datos guardados en '{ruta_salida_archivo}'.")
        
    # Calcula la diferencia de tiempo entre la hora de inicio y la hora de finalización
    tiempo_transcurrido = hora_final - hora_inicio
    print("Tiempo transcurrido:", tiempo_transcurrido)

if __name__ == "__main__":
    main()  # Llama a la función principal si el script se ejecuta directamente