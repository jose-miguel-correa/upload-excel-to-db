# Informe_STAFF_SEM_41_Status_C_de_A

import pandas as pd
from connection import conn, cursor


def process_Informe_STAFF_SEM_41_Status_C_de_A(file_path):
    xls_path = './files/Informe STAFF SEM 41.xlsx'
    sheet = "Status C de A"
    table_name = 'Informe_STAFF_SEM_41_Status_C_de_A'
 
    df = pd.read_excel(xls_path, dtype=str, engine="openpyxl", sheet_name=sheet)

    df = df.where(pd.notna(df), None)

    # Mapeo de nombres de columna excel con tabla destino
    df.rename(columns={'Gerencia': 'gerencia'}, inplace=True)
    df.rename(columns={'Empresa': 'empresa'}, inplace=True)
    df.rename(columns={'Subcontrato': 'subcontrato'}, inplace=True)
    df.rename(columns={'Dotación': 'dotacion'}, inplace=True)
    df.rename(columns={'Fecha inicio de revisión': 'fechaIniciodeRevision'}, inplace=True)
    df.rename(columns={'N° de revisión': 'numdeRevision'}, inplace=True)
    df.rename(columns={'% avance  C de A  Semana 41': 'ptjeAvanceCdeA'}, inplace=True)

    insert_sql = f"INSERT INTO {table_name} ([{'], ['.join(df.columns)}]) VALUES ({', '.join(['%s'] * len(df.columns))})"

    # Iterar en las filas del dataframe e insertarlas en la tabla
    for index, row in df.iterrows():
        values = tuple(row)
        cursor.execute(insert_sql, values)

        # Poblar id del área
    poblar_id_area = f"""
    UPDATE {table_name}
    SET idArea = 1
    WHERE idArea IS NULL;
    """
    cursor.execute(poblar_id_area)

    # Poblar fecha registro
    poblar_fecha_registro = f"""
    UPDATE {table_name}
    SET fechaRegistro = CONVERT(VARCHAR, GETDATE(), 23)
    WHERE fechaRegistro IS NULL;
    """
    cursor.execute(poblar_fecha_registro)

    conn.commit()
    cursor.close()