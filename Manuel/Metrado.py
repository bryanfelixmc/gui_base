import pandas as pd

# Leer archivo de entrada
archivo = "Base de datos de partida.xlsx"
df_bahia = pd.read_excel(archivo, sheet_name="bahia")
df_metrado = pd.read_excel(archivo, sheet_name="metrado")

# Limpiar textos
df_bahia["Elemento"] = df_bahia["Elemento"].str.strip().str.upper()
df_metrado["DESCRIPCION"] = df_metrado["DESCRIPCION"].str.upper()

# Relacionar cada DESCRIPCION con su TAG base
tags = {
    "DESCARGADORES DE SOBRETENSIÓN": "PR",
    "TRAMPA DE ONDA": "TO",
    "TRANSFORMADOR DE TENSIÓN CAPACITIVO": "TTC",
    "SECCIONADOR DE LÍNEA": "SL",
    "TRANSFORMADOR DE CORRIENTE": "TC",
    "INTERRUPTOR DE POTENCIA": "IN",
    "SECCIONADOR DE BARRA": "SB"
}

# Generar metrado por bahía
rows_bahia = []
item_n = 0

for i, row in df_metrado.iterrows():
    descripcion = row["DESCRIPCION"]
    cantidad = row["CANTIDAD"]

    if descripcion.startswith("BAHÍA DE"):
        item_n += 1
        item = f"{item_n}"
        rows_bahia.append({"ITEM": item, "DESCRIPCION": descripcion.title(), "CANTIDAD": cantidad})
        subitem = 1
    else:
        item = f"{item_n}.{subitem}"
        rows_bahia.append({"ITEM": item, "DESCRIPCION": descripcion.title(), "CANTIDAD": cantidad})
        tag_base = tags.get(descripcion.upper(), "")
        if tag_base:
            tag_rows = df_bahia[df_bahia["Elemento"].str.startswith(tag_base)]
            tag_rows = tag_rows.drop_duplicates(subset=["Descripción", "Valor", "Unidad"])
            for _, det in tag_rows.iterrows():
                if str(det["Descripción"]).strip().upper() != "N.A.":
                    detalle = f"    {det['Descripción']} = {det['Valor']} {det['Unidad']}"
                    rows_bahia.append({"ITEM": "", "DESCRIPCION": detalle, "CANTIDAD": ""})
        subitem += 1

df_metrado_bahia = pd.DataFrame(rows_bahia)

# Generar metrado total por equipo con numeración
df_total = df_metrado[~df_metrado["DESCRIPCION"].str.startswith("BAHÍA DE")].copy()
df_total = df_total.groupby("DESCRIPCION", as_index=False)["CANTIDAD"].sum()
df_total = df_total.sort_values(by="DESCRIPCION")  # Orden alfabético opcional

usados = set()
rows_total = []

for i, row in enumerate(df_total.itertuples(), start=1):
    descripcion = row.DESCRIPCION.title()
    cantidad = row.CANTIDAD
    item = f"{i}"
    rows_total.append({"ITEM": item, "EQUIPO": descripcion, "CANTIDAD": cantidad})

    tag_base = tags.get(row.DESCRIPCION.upper(), "")
    if tag_base and tag_base not in usados:
        usados.add(tag_base)
        tag_rows = df_bahia[df_bahia["Elemento"].str.startswith(tag_base)]
        tag_rows = tag_rows.drop_duplicates(subset=["Descripción", "Valor", "Unidad"])
        for _, det in tag_rows.iterrows():
            if str(det["Descripción"]).strip().upper() != "N.A.":
                detalle = f"    {det['Descripción']} = {det['Valor']} {det['Unidad']}"
                rows_total.append({"ITEM": "", "EQUIPO": detalle, "CANTIDAD": ""})

df_metrado_equipo = pd.DataFrame(rows_total)

# Exportar resultado a nuevo Excel
with pd.ExcelWriter("Metrado_Final.xlsx", engine="openpyxl") as writer:
    df_metrado_bahia.to_excel(writer, sheet_name="Metrado por bahía", index=False)
    df_metrado_equipo.to_excel(writer, sheet_name="Metrado total por equipo", index=False)

print("Metrado generado como 'Metrado_Final.xlsx'")