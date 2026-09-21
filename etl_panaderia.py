import pandas as pd
from datetime import date
import sys 
import os 
from dotenv import load_dotenv
from sqlalchemy import create_engine
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

sys.path.append(r"C:\Users\Admin\OneDrive\Documentos\mis_funciones")
from etl_funciones import(
    cargar_excel,
    validar_calidad,
    convertir_numero_texto,
    limpiar_dataframe,
    cargar_postgresql,
    identificar_inexistentes,
    ejecutar_query,
    exportar_excel,
    enviar_email
)

#CONECTAR A POSTGRES
load_dotenv()

cargar_postgresql(
    database=os.getenv("DB_NAME"),
    user=os.getenv("DB_USER"),
    password=os.getenv("DB_PASSWORD")
)

db_url = (
    f"postgresql://{os.getenv('DB_USER')}:"
    f"{os.getenv('DB_PASSWORD')}@"
    f"{os.getenv('DB_HOST')}:"
    f"{os.getenv('DB_PORT')}/"
    f"{os.getenv('DB_NAME')}"
)

engine = create_engine(db_url)

#CARGAR EXCEL
df_productos = cargar_excel("panaderia_productos.xlsx")
df_clientes = cargar_excel("panaderia_clientes.xlsx")
df_ventas = cargar_excel("panaderia_ventas.xlsx")

print("carga completada")

#VALIDAR CALIDAD
validar_productos = validar_calidad(df_productos, "Precio_unitario")
validar_clientes = validar_calidad(df_clientes)
validar_ventas = validar_calidad(df_ventas)
comprobar_ventas = validar_calidad(df_ventas, "Cantidad")

print("NULOS CLIENTES ")
print(validar_clientes["nulos"])

print("DUPLICADO CLIENTES")
print(validar_clientes["duplicados"])

print("NULOS VENTAS")
print(validar_ventas["nulos"])

print("DUPLICADO VENTAS")
print(validar_ventas["duplicados"])

print("NULOS PRODUCTOS")
print(validar_productos["nulos"])

print("DUPLICADO PRODUCTOS")
print(validar_productos["duplicados"])

print("VALORES INVALIDOS PRODUCTOS (Precio_unitario)")
print(validar_productos["valores_invalidos"])

print("VALORES INVALIDOS VENTAS (Cantidad)")
print(comprobar_ventas["valores_invalidos"])

print(df_clientes.dtypes)

#CONVERTIR NUMEROS A TEXTO
df_clientes = convertir_numero_texto(df_clientes, ["Telefono"])

#LIMPIAR NULOS
df_productos["Categoria"] = df_productos["Categoria"].fillna("Sin categoria")
    
#LIMPIAR DATAFRAME 
df_productos = limpiar_dataframe(
    df_productos,
    columnas_texto=["Nombre_producto", "Categoria",]
)
df_clientes = limpiar_dataframe(
    df_clientes,
    columnas_texto=["Nombre_cliente", "Tipo_cliente","Ciudad","Telefono"]
)
df_ventas = limpiar_dataframe(
    df_ventas,
    columnas_texto=["Estado"],
    columnas_fecha=["Fecha"],
    formato_latino=False
)   

print("limpieza completada")

#CARGAR A POSTGRESQL
df_productos.to_sql("productos", engine, index=False, if_exists="replace")
df_clientes.to_sql("clientes", engine, index=False, if_exists="replace")
df_ventas.to_sql("ventas", engine, index=False, if_exists="replace")

print("tablas cargadas a postgresql")

#IDENTIFICAR INEXISTENTES
huerfanas = identificar_inexistentes(df_ventas, df_productos, "Id_producto")
print(huerfanas)

#MERGUE UNIR TABLAS
df_paso1 = df_ventas.merge(df_productos, on="Id_producto", how="inner")
df_combinado = df_paso1.merge(df_clientes, on="Id_cliente", how="inner")

df_combinado.to_sql("combinado", engine, index=False, if_exists="replace")

print(df_combinado.columns.tolist())

#EJECUTAR QUERY 
#QUERY 1 Ventas por categoria
df_categoria = pd.read_sql("""
    SELECT "Categoria",
        COUNT(*) AS cantidad_ventas,
        SUM("Cantidad") AS unidades_vendidas,
        ROUND(SUM("Precio_unitario" * "Cantidad")::numeric, 2) AS ingresos_totales
    FROM combinado
    WHERE "Estado" = 'Pagada'
    GROUP BY "Categoria"
    ORDER BY ingresos_totales DESC
""", engine)

print(df_categoria)

#QUERY 2 Top 5 productos
df_productos_top_5 = pd.read_sql("""
    SELECT "Nombre_producto", "Categoria",
        SUM("Cantidad") AS unidades_vendidas,
        ROUND(SUM("Precio_unitario" * "Cantidad")::numeric, 2) AS ingresos_totales
    FROM combinado
    WHERE "Estado" = 'Pagada'
    GROUP BY "Nombre_producto", "Categoria"
    ORDER BY unidades_vendidas DESC
    LIMIT 5
""", engine)

print(df_productos_top_5)

#QUERY 3 Cliente que mas compro
df_clientes_top = pd.read_sql("""
    SELECT "Nombre_cliente", "Tipo_cliente", "Ciudad", "Id_cliente",
        COUNT(*) AS cantidad_compras,
        ROUND(SUM("Precio_unitario" * "Cantidad")::numeric, 2) AS total_comprado
    FROM combinado
    WHERE "Estado" = 'Pagada'
    GROUP BY "Nombre_cliente", "Tipo_cliente", "Ciudad", "Id_cliente"
    ORDER BY total_comprado DESC
    LIMIT 10
""", engine)

print(df_clientes_top)

#QUERY 4 Ventas por tipo de clientes 
df_tipo_cliente_top = pd.read_sql("""
    SELECT "Tipo_cliente",
        COUNT(*) AS cantidad_compras,
        ROUND(SUM("Precio_unitario" * "Cantidad")::numeric, 2) AS total_comprado
    FROM combinado
    WHERE "Estado" = 'Pagada'
    GROUP BY "Tipo_cliente"
    ORDER BY total_comprado DESC
""", engine)

print(df_tipo_cliente_top)

#QUERY 5 Ventas por ciudad
df_ciudad_top = pd.read_sql("""
    SELECT "Ciudad",
        COUNT(*) AS cantidad_compras,
        ROUND(SUM("Precio_unitario" * "Cantidad")::numeric, 2) AS total_comprado
    FROM combinado
    WHERE "Estado" = 'Pagada'
    GROUP BY "Ciudad"
    ORDER BY total_comprado DESC
""", engine)

print(df_ciudad_top)

#QUERY 6 Ranking por categoria 
df_ranking = pd.read_sql("""
    SELECT "Categoria", "Nombre_producto",
        ROUND(SUM("Precio_unitario" * "Cantidad")::numeric, 2) AS ingresos_totales,
        RANK() OVER (PARTITION BY "Categoria" ORDER BY ROUND(SUM("Precio_unitario" * "Cantidad")::numeric, 2) DESC) AS ranking 
    FROM combinado
    WHERE "Estado" = 'Pagada'
    GROUP BY "Categoria", "Nombre_producto"
    ORDER BY "Categoria", ranking 
""", engine)

print(df_ranking)

#QUERY 7 Ventas mes a mes lANG
df_crecimiento = pd.read_sql("""
    SELECT
        TO_CHAR("Fecha" ::date, 'MM-YYYY') AS mes,
        ROUND(SUM("Precio_unitario" * "Cantidad")::numeric, 2) AS ingresos_totales,
        LAG(ROUND(SUM("Precio_unitario" * "Cantidad")::numeric, 2)) OVER (
            ORDER BY TO_CHAR("Fecha" ::date, 'MM-YYYY')) AS mes_anterior,
        ROUND(SUM("Precio_unitario" * "Cantidad")::numeric,2) - LAG(
            ROUND(SUM("Precio_unitario" * "Cantidad")::numeric, 2)) OVER (
            ORDER BY TO_CHAR("Fecha" ::date, 'MM-YYYY')) AS diferencia 
    FROM combinado
    WHERE "Estado" = 'Pagada'
    GROUP BY TO_CHAR("Fecha" ::date, 'MM-YYYY')
    ORDER BY TO_CHAR("Fecha" ::date, 'MM-YYYY')
""", engine)
    
print(df_crecimiento)   

#EXPORTAR EXCEL 
df_combinado["Fecha"] = pd.to_datetime(df_combinado["Fecha"]).dt.date

hoy = date.today().strftime("%Y%m%d")

exportar_excel(f"reporte_panaderia_etl_{hoy}.xlsx",{
    "datos_completos": df_combinado,
    "ventas_por_categoria": df_categoria,
    "top_5_productos": df_productos_top_5,
    "clientes_top": df_clientes_top,
    "tipo_cliente_top": df_tipo_cliente_top,
    "ciudad_top": df_ciudad_top,
    "ranking": df_ranking,
    "crecimiento": df_crecimiento,
    "inexistentes": huerfanas
})

#FORMATO PROFESIONAL
wb = load_workbook(f"reporte_panaderia_etl_{hoy}.xlsx")

azul_oscuro = "1F4E79"
azul_medio  = "2E75B6"
azul_claro  = "D6E4F0"
blanco      = "FFFFFF"

borde_fino = Border(
    left=Side(style='thin', color='CCCCCC'),
    right=Side(style='thin', color='CCCCCC'),
    top=Side(style='thin', color='CCCCCC'),
    bottom=Side(style='thin', color='CCCCCC')
)

for nombre_hoja in wb.sheetnames:
    ws = wb[nombre_hoja]

    #Encabezados fila 1
    for cell in ws["1"]:
        cell.font      = Font(name="Arial", size=9, bold=True, color=blanco)
        cell.fill      = PatternFill("solid", fgColor=azul_oscuro)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = borde_fino
    ws.row_dimensions[1].height = 35

    #Filas datos
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        color = azul_claro if i % 2 == 0 else blanco
        for j, cell in enumerate(row, start=1):
            cell.fill   = PatternFill("solid", fgColor=color)
            cell.border = borde_fino
            cell.font   = Font(name="Arial", size=9)
            if j == 1:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.font = Font(name="Arial", size=9, bold=True)
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'
        ws.row_dimensions[i].height = 18

    #Ancho automatico 
    for col in ws.iter_cols():
        max_len = 0
        col_letter = None
        for cell in col:
            if hasattr(cell, 'column_letter'):
                col_letter = cell.column_letter
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
        if col_letter:
            ws.column_dimensions[col_letter].width = min(max_len + 2, 15)

     # Ajustar ancho según encabezado
    for cell in ws[1]:
        if cell.value:
            col_letter = cell.column_letter
            ancho_encabezado = len(str(cell.value)) + 2
            ancho_actual = ws.column_dimensions[col_letter].width
            ws.column_dimensions[col_letter].width = max(ancho_actual, ancho_encabezado)

wb.save(f"reporte_panaderia_etl_{hoy}.xlsx")
print("Formato aplicado correctamente")

#ENVIAR EMAIL
enviar_email(
    destinatario="feletbass@gmail.com",
    asunto=f"Reporte Panaderia ETL {hoy}",
    mensaje="Adjunto el reporte de ventas actualizado",
    archivo=f"reporte_panaderia_etl_{hoy}.xlsx",
    user=os.getenv("EMAIL_USER"),
    password=os.getenv("EMAIL_PASSWORD")
)

















