import pandas as pd  # Importa la biblioteca Pandas y la asigna al alias 'pd'
import numpy as np  # Importa la biblioteca NumPy y la asigna al alias 'np'
import os  # Importa el módulo os para interactuar con el sistema operativo
import datetime  # Importa el módulo datetime para trabajar con fechas y horas
import warnings  # Importa el módulo warnings para manejar advertencias
from datetime import timedelta  # Importa la clase timedelta del módulo datetime

warnings.filterwarnings("ignore")  # Ignora las advertencias durante la ejecución

# Feriados
directorio_modulo = os.path.dirname(os.path.abspath(__file__))  # Obtiene el directorio del módulo actual
ruta_relativa_archivo = os.path.join(directorio_modulo, "data", "feriados.xlsx")  # Construye la ruta al archivo de feriados
feriados = pd.read_excel(ruta_relativa_archivo)['Feriados'].apply(lambda x: x.date()).tolist()  # Lee los feriados del archivo y los convierte a una lista de objetos de fecha

# Función para capturar la hora actual (ayudará a determinar el tiempo de demora para ejecutar el programa)
def obtener_hora_actual():
    return datetime.datetime.now()  # Devuelve la hora actual del sistema

def calcular_dias(row):
    row['DIAS_CALCULADOS'] = np.busday_count(row['Fecha_Inicio_AUX'].date() + timedelta(days=1), row['Fecha_Fin_AUX'].date() + timedelta(days=1), holidays=feriados)  # Calcula la cantidad de días hábiles entre dos fechas, excluyendo los feriados
    return row

def main():
    hora_inicio = obtener_hora_actual()  # Registra la hora de inicio del proceso

    df_excel_de_entrada_global = "Excel_prueba.xlsx"  # Nombre del archivo Excel de entrada
    df = pd.read_excel(df_excel_de_entrada_global)  # Lee el archivo Excel de entrada y carga los datos en un DataFrame
    excel_de_salida_global = "Excel_prueba_calculado.xlsx"  # Nombre del archivo Excel de salida

    print("Calculando días. Espere un momento por favor (dependiendo del tamaño puede tardar unos minutos)...")

    # Convierte las fechas de inicio y fin del DataFrame a objetos de fecha y hora (usaremos auxiiares para no modificar los originales)
    df['Fecha_Inicio_AUX'] = pd.to_datetime(df['Fecha_Inicio'], format='%d/%m/%Y')
    df['Fecha_Fin_AUX'] = pd.to_datetime(df['Fecha_Fin'], format='%d/%m/%Y')

    # Aplica la función 'calcular_dias' a cada fila del DataFrame para calcular los días hábiles
    df = df.apply(calcular_dias, axis=1)

    # Elimina las columnas auxiliares utilizadas para cálculos
    df.drop(columns=['Fecha_Inicio_AUX', 'Fecha_Fin_AUX'], inplace=True)

    # Guarda el DataFrame modificado en un nuevo archivo Excel
    df.to_excel(excel_de_salida_global, sheet_name='CALCULO', index=False)

    hora_final = obtener_hora_actual()  # Registra la hora de finalización del proceso
    print(f"Creación Exitosa. Proceso terminado. Datos guardados en '{excel_de_salida_global}'.")
    
    # Calcula la diferencia de tiempo entre la hora de inicio y la hora de finalización
    tiempo_transcurrido = hora_final - hora_inicio
    print("Tiempo transcurrido:", tiempo_transcurrido)

if __name__ == "__main__":
    main()  # Llama a la función principal si el script se ejecuta directamente
