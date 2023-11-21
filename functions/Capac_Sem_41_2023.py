import pandas as pd
from connection import conn, cursor
import timeit


def process_Capac_Sem_41_2023(file_path):
    xls_path = './files/Capac Sem 41_2023.xlsx'
    sheet = 'Listado acum semana 41'

    table_name = 'Capac_Sem_41_2023'

    # Read each target sheet into a DataFrame
    print(f"Processing sheet: {sheet}")
    df = pd.read_excel(xls_path, dtype=str, engine="openpyxl", sheet_name=sheet, skiprows=1)

    df = df.where(pd.notna(df), None)
    print("Se eliminan nulos")

    # Mapeo de nombres de columna excel con tabla destino
    df.rename(columns={'Area': 'area'}, inplace=True)
    df.rename(columns={'Capacitación': 'capacitacion'}, inplace=True)
    df.rename(columns={'Nombre Trabajador. ': 'nombreTrabajador'}, inplace=True)
    print("Se renombran filas")
    df2 = df.loc[:, ['area', 'capacitacion', 'nombreTrabajador']]
    print("Se seleccionan las 3 columnas")

    insert_sql = f"INSERT INTO {table_name} ([{'], ['.join(df2.columns)}]) VALUES ({', '.join(['%s'] * len(df2.columns))})"
    print("Se crea INSERT")

    # Iterar en las filas del dataframe e insertarlas en la tabla
    batch_size = 100
    print("Se define tamaño del batch = 100")

    # Iterar en las filas del dataframe e insertarlas en la tabla
    for index, row in df.iterrows():
        values = tuple(row)
        cursor.execute(insert_sql, values)

    print("Se ejecuta insert: ", insert_sql, values)

    # Poblar fecha registro
    poblar_fecha_registro = f"""
    UPDATE {table_name}
    SET fechaRegistro = CONVERT(VARCHAR, GETDATE(), 23)
    WHERE fechaRegistro IS NULL;
    """
    cursor.execute(poblar_fecha_registro)
    print("Se crea fecha de registro")


    conn.commit()

    cursor.close()

# 10:33 - 
