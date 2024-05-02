import pandas as pd

# Cargar el archivo Excel
archivo_excel = 'cruce.xlsx'

# Nombres de las hojas
nombres_hojas = ['hos11', 'ssd3', 'hos12', 'hos14', 'hos3', 'hos7', 'server2', 'server2hostingcom', 'ssd4', 'ssd5', 'uh2', 'ssd1']

# Cargar todas las hojas en un diccionario de DataFrames, añadiendo una columna para el nombre de la hoja
hojas = {nombre: pd.read_excel(archivo_excel, sheet_name=nombre, engine='openpyxl') for nombre in nombres_hojas}
for nombre, df in hojas.items():
    df['hoja'] = nombre  # Añadir columna que indica el nombre de la hoja
    df['estado'] = df['Is Suspended'].apply(lambda x: 'Suspendido' if x == 1 else 'Activo')

# Concatenar todos los DataFrames en uno solo para facilitar el análisis
df_total = pd.concat(hojas.values(), ignore_index=True)

# Convertir la columna 'dominio' a texto en el DataFrame total
df_total['dominio'] = df_total['dominio'].astype(str)

# Agrupar por 'dominio' y contar las hojas únicas en las que aparece cada dominio
dominios_hojas = df_total.groupby('dominio')['hoja'].nunique()

# Filtrar para obtener solo los dominios que aparecen en más de una hoja
dominios_repetidos = dominios_hojas[dominios_hojas > 1]

# Para cada dominio repetido, obtener las hojas donde aparece y su estado
repetidos_detalle = df_total[df_total['dominio'].isin(dominios_repetidos.index)].groupby('dominio').apply(lambda x: list(zip(x['hoja'], x['estado'])))

# Agregar una columna para verificar si el dominio está en 'hos11'
df_dominios_repetidos = pd.DataFrame({
    'Dominio': repetidos_detalle.index,
    'Hojas y Estado': repetidos_detalle.values,
    'Número de Repeticiones': repetidos_detalle.map(len),
    'En hos11': repetidos_detalle.apply(lambda x: 'hos11' in [hoja for hoja, _ in x])
})

# Ordenar los resultados para mostrar primero los dominios en 'hos11'
df_dominios_repetidos = df_dominios_repetidos.sort_values(by=['En hos11', 'Número de Repeticiones', 'Dominio'], ascending=[False, False, True])

# Imprimir o guardar los resultados
print("Dominios que se repiten en varias hojas:")
print(df_dominios_repetidos)

# Guardar el DataFrame en una nueva hoja del mismo archivo Excel
with pd.ExcelWriter(archivo_excel, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    df_dominios_repetidos.to_excel(writer, sheet_name='resultado', index=False)
