# Informe_VPPE_Cordillera_Sem_41_Status_C_de_A

import pandas as pd
import numpy as np
from connection import conn, cursor


def process_Informe_VPPE_Cordillera_Sem_41_Status_C_de_A(file_path):
    xls_path = './files/Informe VPPE Cordillera Sem 41_.xlsx'
    sheet = "Staus C de A "
    table_name = 'Informe_VPPE_Cordillera_Sem_41_Status_C_de_A'
 
    df = pd.read_excel(xls_path, dtype=str, engine="openpyxl", sheet_name=sheet)

    df = df.where(pd.notna(df), None)

    # Mapeo de nombres de columna excel con tabla destino
    df.rename(columns={'Gerencia': 'gerencia'}, inplace=True)
    df.rename(columns={'Empresa': 'empresa'}, inplace=True)
    df.rename(columns={'Subcontrato': 'subcontrato'}, inplace=True)
    df.rename(columns={'Dotación': 'dotacion'}, inplace=True)
    df.rename(columns={'Fecha Inicio de Revisión': 'fechaIniciodeRevision'}, inplace=True)
    df.rename(columns={'N° de Revisión': 'numDeRevision'}, inplace=True)
    df.rename(columns={'% avance  C de A Semana 41': 'ptjeAvanceCdeA'}, inplace=True)

    df['ptjeAvanceCdeA'] = pd.to_numeric(df['ptjeAvanceCdeA'], errors='coerce')

    df = df.dropna(subset=['numDeRevision'])


    df['ptjeAvanceCdeA'] = np.where((df['ptjeAvanceCdeA'] > 1) & (df['ptjeAvanceCdeA'] <= 100), df['ptjeAvanceCdeA'] / 100, df['ptjeAvanceCdeA'])

    
    df2 = df.loc[:,['gerencia','empresa','subcontrato','dotacion','fechaIniciodeRevision','numDeRevision','ptjeAvanceCdeA']]
    print(df2)
    insert_sql = f"INSERT INTO {table_name} ([{'], ['.join(df2.columns)}]) VALUES ({', '.join(['%s'] * len(df2.columns))})"

    # Iterar en las filas del dataframe e insertarlas en la tabla
    for index, row in df2.iterrows():
        values = tuple(row)
        cursor.execute(insert_sql, values)

    # Borrar 5to Molino en columna empresa
    sentencia_borrar_quinto_molino = f"""
    DELETE FROM {table_name}
    WHERE gerencia like '5to%'
    """
    cursor.execute(sentencia_borrar_quinto_molino)


    # Poblar id del área
    poblar_id_area = f"""
    UPDATE {table_name}
    SET idArea = 7
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