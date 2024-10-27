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

# Inicializar lista para almacenar todos los datos
todos_datos = []

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

        for j, col in enumerate(range(columna_inicio_disponibilidad, columna_fin_disponibilidad + 1)):
            dia = dias[j]  # Obtener el día correspondiente
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
            
            todos_datos.append([id_empleado, dia, horario, disponibilidad])

# Guardar todos los datos en un único archivo CSV llamado 'disponibilidades.csv'
df = pd.DataFrame(todos_datos, columns=["ID_Empleado", "Día", "Horario", "Disponibilidad"])
df.to_csv("disponibilidades.csv", index=False, encoding="utf-8")


#por si quiere visualizar la tabla de los datos recabados
#print(tabulate(df, headers="keys", tablefmt="fancy_grid")) 

print("Todos los datos han sido guardados en 'disponibilidades.csv'.")
