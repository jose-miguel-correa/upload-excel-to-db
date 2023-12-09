# Subir archivos xlsx a tablas de un SQL SERVER usando Python.
### Objetivo general:
- Automatizar la carga de datos desde archivos excel a una base de datos.
### Objetivos específicos:
- Usar Python para subir recursivamente archivos xlsx ubicados en carpeta local a distintas tablas en un SQL Server.
- Reconocer nombre de archivo e insertar sus datos en la tabla correspondiente.
- Reconocer nombres de hojas y subir datos a tabla correspondiente.
- Reconocer columnas específicas y subir sus datos a tabla correspondiente.
- Los datos deben quedar en un stage.
- Los datos deben pasar a una Fact Table.
- Diseñar base de datos y relacionar tablas adecuadamente.
- Generar procedimientos almacenados gestionados por Python.

### Modelo Base de inserción de datos desde planillas Excel a MSSQL Server

### Crear base de datos MSSQL Server en entorno Linux:

- Bajar docker de mssql.                               <br>
```
docker pull mcr.microsoft.com/mssql/server
```
- Correr imagen.
```
docker run <nombre_imagen>
```
- SA_PASSWORD=<tu-contraseña>.
- La base de datos para este modelo siempre es "master".<br>
```
docker run -e 'ACCEPT_EULA=Y' -e 'SA_PASSWORD=MSSQL_Server2023' -p 1433:1433 -d mcr.microsoft.com/mssql/server
```

### Con alguna extensión del VSCODE, conectar a la db.
