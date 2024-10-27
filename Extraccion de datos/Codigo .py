import openpyxl
import pandas as pd
from tabulate import tabulate

# Cargar el archivo de Excel
excel_dataframe = openpyxl.load_workbook("Copia de Ejemplo Disponibilidad.xlsx")

# Definir los rangos para filas y columnas
fila_inicio_datos = 13
fila_fin_datos = 27
columna_inicio_disponibilidad = 5
columna_fin_disponibilidad = 10

# Diccionario de colores para verificar la disponibilidad
colores_disponibilidad = {
    "FFFFFFFF": "sin disponibilidad",  # Blanco (sin relleno)
    "FFEA9999": "sin disponibilidad",  # Rojo (#ea9a99)
    "FFCCCCCC": "poca disponibilidad",  # Gris (#cccccc)
    "FFFFE599": "disponible"           # Amarillo (#ffe599)
}

# Diccionario para mapear cada horario a un ID secuencial
horario_ids = {
    "7:00-8:00": 1,
    "8:00-9:00": 2,
    "9:00-10:00": 3,
    "10:00-11:00": 4,
    "11:00-12:00": 5,
    "12:00-13:00": 6,
    "13:00-14:00": 7,
    "14:00-15:00": 8,
    "15:00-16:00": 9,
    "16:00-17:00": 10,
    "17:00-18:00": 11,
    "18:00-19:00": 12,
    "19:00-20:00": 13,
    "20:00-21:00": 14,
    "21:00-22:00": 15
}

#Diccionario para los dias de la semana, convertir a ID
dias_ids = {
    "LUNES": 1,
    "MARTES": 2,
    "MIÉRCOLES" or "MIERCOLES": 3,
    "JUEVES": 4,
    "VIERNES": 5,
    "SÁBADO (virtual)": 6,
    "DOMINGO": 7
}

# Inicializar lista para almacenar todos los datos
todos_datos = []
Datos = []

# Iterar sobre cada hoja de trabajo, comenzando desde la segunda hoja
for nombre_hoja in excel_dataframe.sheetnames[0:]:
    # Usar el nombre de la hoja como id_empleado
    id_empleado = nombre_hoja

    # Seleccionar la hoja actual
    hoja = excel_dataframe[nombre_hoja]

    # Extraer los horarios y días
    horarios = [hoja.cell(row=row, column=3).value for row in range(fila_inicio_datos, fila_fin_datos + 1)]
    dias = [hoja.cell(row=12, column=col).value for col in range(columna_inicio_disponibilidad, columna_fin_disponibilidad + 1)]

    # Crear registros para cada combinación de día y horario en la hoja actual
    for i, row in enumerate(range(fila_inicio_datos, fila_fin_datos + 1)):
        horario = horarios[i]  # Obtener el horario actual
        id_horario = horario_ids.get(horario, "No definido")  # Obtener el ID del horario

        for j, col in enumerate(range(columna_inicio_disponibilidad, columna_fin_disponibilidad + 1)):
            dia = dias[j]  # Obtener el día correspondiente
            id_dias = dias_ids.get(dia, "No definido")  # Obtener el ID del horario
            # Obtener la celda actual
            celda = hoja.cell(row=row, column=col)

            # Verificar el color de la celda, si tiene relleno y color específico
            if celda.fill and celda.fill.start_color:
                color_celda = celda.fill.start_color.rgb
                disponibilidad = colores_disponibilidad.get(color_celda, "sin disponibilidad")
            else:
                # Sin color
                disponibilidad = "sin disponibilidad"
                
            # En caso de que se tomen en cuenta las celdas con texto en color rojo, esto lo marca como "sin disponibilidad"
            #if celda.font and celda.font.color and celda.font.color.rgb == "FFFF0000":
                #disponibilidad = "sin disponibilidad"
            
            #Esta parte la es para obtener solo los registros con disponibilidad y poca disponibilidad, y ignora los que estan en color blanco o rojo, tambien los que tienen letras o numros rojos!!
            #if disponibilidad in ["disponible", "poca disponibilidad"]: 
                #todos_datos.append([id_empleado, dia, horario, disponibilidad])
            
             # Agregar solo los registros con disponibilidad "disponible" o "poca disponibilidad"
            if disponibilidad in ["disponible", "poca disponibilidad"]:
                # Asigno ID basado en la disponibilidad
                 if disponibilidad == "disponible":
                     id_disponibilidad = 1
                 elif disponibilidad == "poca disponibilidad":
                     id_disponibilidad = 2
                 todos_datos.append([id_empleado, id_dias, id_horario, id_disponibilidad])
                #Si necesita verificar los datos, descomentar esta parte y la de las ultimas lineas
                 #Datos.append([id_empleado, dia, horario, disponibilidad]) 

df = pd.DataFrame(todos_datos, columns=["ID_Empleado", "ID_Dias", "ID_Horario", "ID_Disponibilidad"])
df.to_csv("disponibilidades_ids.csv", index=False, encoding="utf-8")
print(tabulate(df, headers="keys", tablefmt="fancy_grid"))
print("Todos los datos con ids han sido guardados en 'disponibilidades_ids.csv'.")


#Descomentar esta parte para la visualizacion completa de datos
#df2 = pd.DataFrame(Datos, columns=["ID_Empleado", "Dia", "Horario", "Disponibilidad"])
#df2.to_csv("disponibilidades.csv", index=False, encoding="utf-8")
#print(tabulate(df2, headers="keys", tablefmt="fancy_grid"))
#print("Todos los datos completos han sido guardados en 'disponibilidades.csv'.")
