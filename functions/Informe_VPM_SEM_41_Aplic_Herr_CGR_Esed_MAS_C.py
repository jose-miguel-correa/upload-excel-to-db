# Informe_VPM_SEM_41_Aplic_Herr_CGR_Esed_MAS_C

import pandas as pd
from connection import conn, cursor


def process_Informe_VPM_SEM_41_Aplic_Herr_CGR_Esed_MAS_C(file_path):
    xls_path = './files/Informe VPM SEM 41.xlsx'
    sheet = "Aplic Herr CGR Esed MAS+C"
    table_name = 'Informe_VPM_SEM_41_Aplic_Herr_CGR_Esed_MAS_C'
 
    df = pd.read_excel(xls_path, dtype=str, engine="openpyxl", sheet_name=sheet)

    df = df.where(pd.notna(df), None)

    # Mapeo de nombres de columna excel con tabla destino
    df.rename(columns={'Empresa': 'empresa'}, inplace=True)
    df.rename(columns={'GRT s/desv': 'grtSinDesv'}, inplace=True)
    df.rename(columns={'GRT  c/desv <': 'grtConDesvMenor'}, inplace=True)
    df.rename(columns={'GRT C/desv >': 'grtConDesvMayor'}, inplace=True)
    df.rename(columns={'RITUS': 'ritus'}, inplace=True)
    df.rename(columns={'VATS': 'vats'}, inplace=True)
    
    df2 = df.loc[:,['empresa','grtSinDesv','grtConDesvMenor','grtConDesvMayor','ritus','vats']]
    
    insert_sql = f"INSERT INTO {table_name} ([{'], ['.join(df2.columns)}]) VALUES ({', '.join(['%s'] * len(df2.columns))})"
    
    # Iterar en las filas del dataframe e insertarlas en la tabla
    for index, row in df2.iterrows():
        values = tuple(row)
        cursor.execute(insert_sql, values)

    # Borrar TOTAL_ESED en columna empresa
    sentencia_borrar_total_esed = f"""
    DELETE FROM {table_name}
    WHERE empresa like '%total%'
    """
    cursor.execute(sentencia_borrar_total_esed)

    # Borrar valores numericos sueltos
    sentencia_borrar_valores_sueltos = f"""
    DELETE FROM {table_name} 
    WHERE empresa IS NULL
    """
    cursor.execute(sentencia_borrar_valores_sueltos)

    # Poblar id del área
    poblar_id_area = f"""
    UPDATE {table_name}
    SET idArea = 5
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

    # Poblar idHerrDelCgr
    poblar_id_herr_del_cgr = f"""
    UPDATE {table_name}
    SET idHerrDelCgr = 1
    WHERE idHerrDelCgr IS NULL
    """
    cursor.execute(poblar_id_herr_del_cgr)

    # Eliminar saltos de línea
    eliminar_saltos_de_linea = f"""
    EXEC RemoveLineBreaks {table_name}, 'id';
    """
    cursor.execute(eliminar_saltos_de_linea)


    conn.commit()
    cursor.close()

