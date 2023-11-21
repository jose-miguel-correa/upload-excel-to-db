import pandas as pd
from connection import conn, cursor


def process_Capac_Sem_41_2023(file_path):
    xls_path = './files/Capac Sem 41_2023.xlsx'
    sheet = 'Listado acum semana 41'
    table_name = 'Capac_Sem_41_2023'
 
    # Read each target sheet into a DataFrame
    df = pd.read_excel(xls_path, dtype=str, engine="openpyxl", sheet_name=sheet, skiprows=1)

    df = df.where(pd.notna(df), None)

    # Mapeo de nombres de columna excel con tabla destino
    df.rename(columns={'Area': 'area'}, inplace=True)
    df.rename(columns={'Capacitación': 'capacitacion'}, inplace=True)
    df.rename(columns={'Nombre Trabajador. ': 'nombreTrabajador'}, inplace=True)
    
    df2 = df.loc[:, ['area', 'capacitacion', 'nombreTrabajador']]

    insert_sql = f"INSERT INTO {table_name} ([{'], ['.join(df2.columns)}]) VALUES ({', '.join(['%s'] * len(df2.columns))})"

    # Iterar en las filas del dataframe e insertarlas en la tabla
    batch_size = 1000
    for chunk in range(0, len(df2), batch_size):
        chunk_df = df2.iloc[chunk:chunk+batch_size]
        values = [tuple(row) for _, row in chunk_df.iterrows()]
        cursor.executemany(insert_sql, values)
        conn.commit()

    poblar_fecha_registro = f"""
    UPDATE {table_name}
    SET fechaRegistro = CONVERT(VARCHAR, GETDATE(), 23)
    WHERE fechaRegistro IS NULL;
    """
    cursor.execute(poblar_fecha_registro)

    conn.commit()
    cursor.close()
