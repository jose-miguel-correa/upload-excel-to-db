# Informe_VPP_Sem_41_Alertas_para_Ejecutivos

import pandas as pd
from connection import conn, cursor


def process_Informe_VPP_Sem_41_Alertas_para_Ejecutivos(file_path):
    xls_path = './files/Informe VPP SEM 41.xlsx'
    sheet = "Alertas para Ejecutivos"
    table_name = 'Informe_VPP_Sem_41_Alertas_para_Ejecutivos'
 
    df = pd.read_excel(xls_path, dtype=str, engine="openpyxl", sheet_name=sheet)

    df = df.where(pd.notna(df), None)

    # Mapeo de nombres de columna excel con tabla destino
    df.rename(columns={'Gerencia': 'gerencia'}, inplace=True)
    df.rename(columns={'Empresa': 'empresa'}, inplace=True)
    df.rename(columns={'Descripción Alerta': 'descripcionAlerta'}, inplace=True)

    insert_sql = f"INSERT INTO {table_name} ([{'], ['.join(df.columns)}]) VALUES ({', '.join(['%s'] * len(df.columns))})"

    # Iterar en las filas del dataframe e insertarlas en la tabla
    for index, row in df.iterrows():
        values = tuple(row)
        cursor.execute(insert_sql, values)

        # Poblar id del área
    poblar_id_area = f"""
    UPDATE {table_name}
    SET idArea = 7 --vppecord
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

    # Eliminar saltos de línea
    eliminar_saltos_de_linea = f"""
    EXEC RemoveLineBreaks {table_name}, 'id';
    """
    cursor.execute(eliminar_saltos_de_linea)


    conn.commit()
    cursor.close()