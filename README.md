# PY-CalculadorDiasHabiles
<br>
Este script de Python calcula la cantidad de días hábiles entre dos fechas, excluyendo los días sabado, domingo y feriados. Puede ser útil para cálculos relacionados con el tiempo de trabajo o proyectos que requieren estimaciones de duración en días hábiles. <br>

##Funcionalidad<br>
El script realiza las siguientes tareas:<br>

**Carga de Datos de Feriados**: Lee un archivo Excel que contiene la lista de feriados. Este archivo debe estar ubicado en la carpeta "data" junto al script principal.<br>

**Cálculo de Días Hábiles**: Utiliza la biblioteca Pandas para cargar un archivo Excel de entrada que contiene fechas de inicio y fin. Luego calcula la cantidad de días hábiles entre estas fechas, excluyendo los feriados.<br>

**Salida de Resultados**: Guarda los resultados en un nuevo archivo Excel, eliminando las columnas auxiliares utilizadas durante el cálculo.<br>

**Registro del Tiempo de Ejecución**: Registra el tiempo de inicio y finalización del proceso, mostrando la duración total de la ejecución al usuario.<br>

Uso<br>
Asegúrate de tener instaladas las bibliotecas requeridas. Puedes instalarlas utilizando pip:<br>

```bash
pip install pandas numpy

```

Coloca el archivo de feriados en la carpeta "data" con el nombre "feriados.xlsx".<br>

Ejecuta el script proporcionando el nombre del archivo Excel de entrada. Por ejemplo:<br>

python calcular_dias_habiles.py<br>

Espera a que el script calcule los días hábiles. El tiempo de ejecución puede variar dependiendo del tamaño de los datos de entrada.<br>

Una vez que el proceso haya finalizado, encontrarás los resultados en un nuevo archivo Excel llamado "Excel_prueba_calculado.xlsx".<br>
