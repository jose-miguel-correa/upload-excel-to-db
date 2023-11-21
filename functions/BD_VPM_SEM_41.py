# BD_VPM_SEM_41

import pandas as pd
from connection import conn, cursor

def process_BD_VPM_SEM_41(file_path):
    xls_path = './files/BD VPM SEM 41.xlsx'
    sheet = "VPM Cttos Finales <22"
    table_name = 'BD_VPM_SEM_41'
 
    df = pd.read_excel(xls_path, dtype=str, engine="openpyxl", sheet_name=sheet, skiprows=1)

    df = df.where(pd.notna(df), None)

    # Mapeo de nombres de columna excel con tabla destino
    df.rename(columns={'ESED Subcontrato u OS': 'esedSubcontratouOs'}, inplace=True)
    df.rename(columns={'E (Esporádica) / P (Permanente)': 'esporadicaPermanente'}, inplace=True)
    df.rename(columns={'Ingreso a SIMIN': 'ingresoaSimin'}, inplace=True)
    df.rename(columns={'Estado Final ante SIMIN': 'estadoFinalAnteSimin'}, inplace=True)
    df.rename(columns={'Status Ctto u OS (VIGENCIA)': 'statusCttosuOsVigencia'}, inplace=True)
    df.rename(columns={'Dotación': 'dotacion'}, inplace=True)

    # Seleccionar columnas
    df2 = df.loc[:,['esedSubcontratouOs', 'esporadicaPermanente', 'ingresoaSimin', 'estadoFinalAnteSimin', 'statusCttosuOsVigencia', 'dotacion']]
    
    insert_sql = f"INSERT INTO {table_name} ([{'], ['.join(df2.columns)}]) VALUES ({', '.join(['%s'] * len(df2.columns))})"

    # Iterar en las filas del dataframe e insertarlas en la tabla
    for index, row in df2.iterrows():
        values = tuple(row)
        cursor.execute(insert_sql, values)

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

    # Poblar idCarpetasDeArranque con valor 'Permanente'
    actualizar_id_permanente = f"""
    UPDATE {table_name}
    SET idCarpetasDeArranque = 2
    WHERE esporadicaPermanente like '%erma%' 
    """
    cursor.execute(actualizar_id_permanente)

    # Poblar idCarpetasDeArranque con valor 'Esporádica'
    actualizar_id_esporadica = f"""
    UPDATE {table_name}
    SET idCarpetasDeArranque = 1
    WHERE esporadicaPermanente like '%spor%' 
    """
    cursor.execute(actualizar_id_esporadica)


    conn.commit()
    cursor.close()