# import pandas as pd
# from typing import List, Optional, Tuple
# import sys
# import unicodedata

# def normalizar(texto):
#     if pd.isna(texto):
#         return ""
#     texto = str(texto).upper().strip()
#     # Quitar tildes
#     texto = ''.join(
#         c for c in unicodedata.normalize('NFD', texto)
#         if unicodedata.category(c) != 'Mn'
#     )
#     return texto

# def encontrar_combinacion_indices(values: List[int], target: int) -> Optional[List[int]]:
#     n = len(values)
#     memo = set()

#     def dfs(pos: int, remaining: int) -> Optional[List[int]]:
#         if remaining == 0:
#             return []
#         if pos >= n or remaining < 0:
#             return None
#         key = (pos, remaining)
#         if key in memo:
#             return None

#         v = values[pos]

#         if v <= remaining:
#             res_take = dfs(pos + 1, remaining - v)
#             if res_take is not None:
#                 return [pos] + res_take

#         res_skip = dfs(pos + 1, remaining)
#         if res_skip is not None:
#             return res_skip

#         memo.add(key)
#         return None

#     return dfs(0, target)

# def procesar_combinaciones(archivo_excel: str, archivo_salida: str):
#     print(f"Leyendo archivo: {archivo_excel}")

#     df_liq = pd.read_excel(archivo_excel, sheet_name="LIQUIDACIONES")
#     df_con = pd.read_excel(archivo_excel, sheet_name="CON")

#     print("Columnas LIQUIDACIONES:", list(df_liq.columns))
#     print("Columnas CON:", list(df_con.columns))

#     df_liq["FORMA_PAGO_N"] = df_liq["FORMA PAGO"].apply(normalizar)

#     mask_pse = (
#         df_liq["FORMA_PAGO_N"].str.contains("PSE", na=False) &
#         df_liq["FORMA_PAGO_N"].str.contains("DEBITO", na=False)
#     )

#     df_liq = df_liq[mask_pse].copy()
#     df_con = df_con[df_con["CONCEPTO"].astype(str).str.upper().str.contains("PAGO VIRTUAL PSE", na=False)].copy()

#     print("Filas LIQUIDACIONES luego de filtrar PSE:", len(df_liq))
#     print("Filas CON luego de filtrar PAGO VIRTUAL PSE:", len(df_con))

#     df_liq["FECHA TRANSACCIÓN"] = pd.to_datetime(df_liq["FECHA TRANSACCIÓN"], errors="coerce")
#     df_con["FECHA"] = pd.to_datetime(df_con["FECHA"], errors="coerce")

#     df_liq = df_liq[~df_liq["FECHA TRANSACCIÓN"].isna()].copy()
#     df_con = df_con[~df_con["FECHA"].isna()].copy()

#     df_liq["FECHA_STR"] = df_liq["FECHA TRANSACCIÓN"].dt.strftime("%Y-%m-%d")
#     df_con["FECHA_STR"] = df_con["FECHA"].dt.strftime("%Y-%m-%d")

#     fechas = sorted(df_liq["FECHA_STR"].unique())
#     print("Fechas a procesar:", fechas)

#     resultado_rows = []

#     for fecha in fechas:
#         print("\nProcesando fecha:", fecha)

#         df_liq_fecha = df_liq[df_liq["FECHA_STR"] == fecha].copy()
#         df_con_fecha = df_con[df_con["FECHA_STR"] == fecha].copy()

#         print(" LIQ rows:", len(df_liq_fecha), " CON rows:", len(df_con_fecha))

#         if len(df_con_fecha) == 0:
#             print(" ⚠ No hay objetivos para esta fecha.")
#             continue

#         if len(df_liq_fecha) == 0:
#             print(" ⚠ No hay valores LIQ para esta fecha.")
#             continue

#         df_liq_fecha["VALOR_NUM"] = pd.to_numeric(df_liq_fecha["VALOR PAGADO"], errors="coerce").fillna(0).astype(int)
#         df_con_fecha["VALOR_NUM"] = pd.to_numeric(df_con_fecha["VALOR"], errors="coerce").fillna(0).astype(int)

#         valores_list = df_liq_fecha["VALOR_NUM"].tolist()
#         original_indices = list(df_liq_fecha.index)

#         paired = sorted(list(enumerate(valores_list)), key=lambda x: x[1], reverse=True)
#         pool_positions = [p[0] for p in paired]
#         pool_values = [p[1] for p in paired]
#         pool = list(zip(pool_positions, pool_values))

#         def pool_values_list():
#             return [v for (_, v) in pool]

#         for _, row_con in df_con_fecha.iterrows():
#             objetivo = int(row_con["VALOR_NUM"])
#             dxc = row_con.get("DXC", "")
#             suma_encontrada = None
#             estado = "No encontrado"
#             valores_usados_str = ""

#             if objetivo == 0 or not pool:
#                 resultado_rows.append({
#                     "FECHA": fecha,
#                     "OBJETIVO": row_con["VALOR"],
#                     "SUMA_ENCONTRADA": None,
#                     "VALORES_USADOS": "",
#                     "DXC": dxc,
#                     "ESTADO": estado
#                 })
#                 continue

#             curr_values = pool_values_list()

#             if sum(curr_values) < objetivo:
#                 resultado_rows.append({
#                     "FECHA": fecha,
#                     "OBJETIVO": row_con["VALOR"],
#                     "SUMA_ENCONTRADA": None,
#                     "VALORES_USADOS": "",
#                     "DXC": dxc,
#                     "ESTADO": "No - pool insuficiente"
#                 })
#                 continue

#             indices_rel = encontrar_combinacion_indices(curr_values, objetivo)

#             if indices_rel is not None:
#                 used_pairs = [pool[i] for i in indices_rel]

#                 used_orig_positions = [p for (p, _) in used_pairs]
#                 pool = [item for item in pool if item[0] not in used_orig_positions]

#                 used_display = []
#                 for orig_pos, val in used_pairs:
#                     df_index = original_indices[orig_pos]
#                     used_display.append((df_index, val))

#                 used_display.sort(key=lambda x: x[0])

#                 valores_usados_str = " + ".join([f"{v:,d}".replace(",", ".") for (_, v) in used_display])

#                 suma_encontrada = objetivo
#                 estado = "✔ Correcto"

#                 print(f"  ✔ Encontrado {objetivo:,d} → DXC {dxc} usando {valores_usados_str}")

#             else:
#                 print(f"  ✖ No se encontró combinación para {objetivo:,d}")

#             resultado_rows.append({
#                 "FECHA": fecha,
#                 "OBJETIVO": row_con["VALOR"],
#                 "SUMA_ENCONTRADA": suma_encontrada,
#                 "VALORES_USADOS": valores_usados_str,
#                 "DXC": dxc,
#                 "ESTADO": estado
#             })

#     df_result = pd.DataFrame(resultado_rows)
#     df_result.to_excel(archivo_salida, index=False)

#     print(f"\nProceso terminado. Archivo generado: {archivo_salida}")
#     print("Filas:", len(df_result))
#     return df_result

# if __name__ == "__main__":
#     archivo_entrada = "DIRECCION GENERAL OCT 2025.xlsx"
#     archivo_salida = "resultado_combinaciones.xlsx"

#     procesar_combinaciones(archivo_entrada, archivo_salida)


#--------------------------------------
import pandas as pd
from typing import List, Optional
import unicodedata

# =====================================================
# UTILIDADES
# =====================================================

def normalizar(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).upper().strip()
    texto = ''.join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )
    return texto


def obtener_columna(df, posibles):
    for col in posibles:
        if col in df.columns:
            return df[col]
    return pd.Series([""] * len(df))


def mapear_descripcion_servicio(concepto):
    c = normalizar(concepto)
    if "CONTRATO DE APRENDIZAJE" in c:
        return "LIQUIDACION-CA"
    if "FIC" in c:
        return "LIQUIDACION-FIC"
    if "APORTES" in c:
        return "LIQUIDACION-APORTES"
    return ""


def encontrar_combinacion_indices(values: List[int], target: int) -> Optional[List[int]]:
    n = len(values)
    memo = set()

    def dfs(pos: int, remaining: int):
        if remaining == 0:
            return []
        if pos >= n or remaining < 0:
            return None
        key = (pos, remaining)
        if key in memo:
            return None
        v = values[pos]
        if v <= remaining:
            res = dfs(pos + 1, remaining - v)
            if res is not None:
                return [pos] + res
        res = dfs(pos + 1, remaining)
        if res is not None:
            return res
        memo.add(key)
        return None

    return dfs(0, target)

# =====================================================
# PROCESO PRINCIPAL
# =====================================================

def procesar_combinaciones(
    archivo_entrada: str,
    archivo_resultados: str,
    archivo_salida_dg: str
):

    print(f"Leyendo archivo: {archivo_entrada}")

    # ---------- LEER HOJAS ----------
    df_liq = pd.read_excel(archivo_entrada, sheet_name="LIQUIDACIONES")
    df_con = pd.read_excel(archivo_entrada, sheet_name="CON")
    df_dg_original = pd.read_excel(archivo_entrada, sheet_name="DIRECCION GENERAL")

    # ---------- FILTRO PSE (PARA CONCILIACIÓN) ----------
    df_liq["FORMA_PAGO_N"] = df_liq["FORMA PAGO"].apply(normalizar)

    mask_pse = (
        df_liq["FORMA_PAGO_N"].str.contains("PSE", na=False) &
        df_liq["FORMA_PAGO_N"].str.contains("DEBITO", na=False)
    )
    df_liq = df_liq[mask_pse].copy()

    df_con = df_con[
        df_con["CONCEPTO"]
        .astype(str)
        .str.upper()
        .str.contains("PAGO VIRTUAL PSE", na=False)
    ].copy()

    # ---------- FECHAS ----------
    df_liq["FECHA TRANSACCIÓN"] = pd.to_datetime(df_liq["FECHA TRANSACCIÓN"], errors="coerce")
    df_con["FECHA"] = pd.to_datetime(df_con["FECHA"], errors="coerce")

    df_liq = df_liq.dropna(subset=["FECHA TRANSACCIÓN"])
    df_con = df_con.dropna(subset=["FECHA"])

    df_liq["FECHA_STR"] = df_liq["FECHA TRANSACCIÓN"].dt.strftime("%Y-%m-%d")
    df_con["FECHA_STR"] = df_con["FECHA"].dt.strftime("%Y-%m-%d")

    fechas = sorted(df_liq["FECHA_STR"].unique())

    # ---------- ACUMULADORES ----------
    resultado_rows = []
    liq_idx_to_dxc = {}
    liq_usadas_indices = set()

    # =====================================================
    # CONCILIACIÓN DXC
    # =====================================================

    for fecha in fechas:

        df_liq_f = df_liq[df_liq["FECHA_STR"] == fecha].copy()
        df_con_f = df_con[df_con["FECHA_STR"] == fecha].copy()

        if df_liq_f.empty or df_con_f.empty:
            continue

        df_liq_f["VALOR_NUM"] = pd.to_numeric(
            df_liq_f["VALOR PAGADO"], errors="coerce"
        ).fillna(0).astype(int)

        df_con_f["VALOR_NUM"] = pd.to_numeric(
            df_con_f["VALOR"], errors="coerce"
        ).fillna(0).astype(int)

        valores = df_liq_f["VALOR_NUM"].tolist()
        original_idx = list(df_liq_f.index)

        pool = sorted(
            list(enumerate(valores)),
            key=lambda x: x[1],
            reverse=True
        )

        def pool_vals():
            return [v for (_, v) in pool]

        for _, row_con in df_con_f.iterrows():

            objetivo = int(row_con["VALOR_NUM"])
            dxc = row_con.get("DXC", "")

            if objetivo <= 0 or not pool:
                continue

            if sum(pool_vals()) < objetivo:
                continue

            idx_rel = encontrar_combinacion_indices(pool_vals(), objetivo)

            if idx_rel is not None:
                usados = [pool[i] for i in idx_rel]
                usados_pos = [p for (p, _) in usados]

                pool = [x for x in pool if x[0] not in usados_pos]

                for pos, _ in usados:
                    idx_real = original_idx[pos]
                    liq_usadas_indices.add(idx_real)
                    liq_idx_to_dxc[idx_real] = dxc

                resultado_rows.append({
                    "FECHA": fecha,
                    "DXC": dxc,
                    "VALOR_CON": row_con["VALOR"]
                })

    # ---------- ARCHIVO RESULTADOS DXC ----------
    pd.DataFrame(resultado_rows).to_excel(archivo_resultados, index=False)

    # =====================================================
    # CONSTRUIR NUEVOS REGISTROS DIRECCION GENERAL
    # =====================================================

    if liq_usadas_indices:

        df_base = df_liq.loc[list(liq_usadas_indices)].copy()
        df_base["DXC"] = df_base.index.map(liq_idx_to_dxc)

        df_base["FORMA_PAGO_N"] = df_base["FORMA PAGO"].apply(normalizar)
        df_base["BANCO_N"] = df_base["BANCO"].apply(normalizar)

        # 🔥 REGLA FINAL DE CÓDIGO / CICLO
        def calcular_codigo_ciclo(row):

            # ✅ PSE → 0 / 0
            if "PSE" in row["FORMA_PAGO_N"]:
                return 0, 0

            # Bancos
            if "PAGO EN EFECTIVO CAJA BANCO" in row["FORMA_PAGO_N"]:
                if "BANCOLOMBIA" in row["BANCO_N"]:
                    return 100, 1
                if row["BANCO_N"] in ["VISA", "MASTERCARD", "AMERICAN EXPRESS"]:
                    return 1, 1

            return "", ""

        df_base[["Código Sistema de Pago", "Ciclo bancario"]] = df_base.apply(
            lambda r: pd.Series(calcular_codigo_ciclo(r)),
            axis=1
        )

        df_dg_nuevo = pd.DataFrame({
            "Nit": df_base["NIT"],
            "Primer Nombre ó Razón Social": df_base["RAZÓN SOCIAL"],
            "Primer Apellido (si es Persona natural)": df_base.get("PRIMER APELLIDO", ""),
            "Descripción de Servicio": df_base["CONCEPTO"].apply(mapear_descripcion_servicio),
            "REGISTRAR EN": obtener_columna(df_base, ["REGISTRAR EN", "REGIONAL PAGO"]),
            "Código SIIF": df_base["CODIGO SIIF"],
            "Fecha pago": df_base["FECHA PAGO"],
            "Valor pago": df_base["VALOR PAGADO"],
            "Regional": obtener_columna(
                df_base,
                ["REGIONAL", "REGIONAL PAGO", "REGIONAL DOMICILIO EMPRESARIAL"]
            ),
            "ticketID": df_base["TICKETID"],
            "Entidad Financiera": df_base["BANCO"],
            "Código Sistema de Pago": df_base["Código Sistema de Pago"],
            "Ciclo bancario": df_base["Ciclo bancario"],
            "DXC": df_base["DXC"]
        })

        df_dg_final = pd.concat(
            [df_dg_original, df_dg_nuevo],
            ignore_index=True
        )

    else:
        df_dg_final = df_dg_original.copy()

    # =====================================================
    # ARCHIVO FINAL NUEVO
    # =====================================================

    with pd.ExcelWriter(archivo_salida_dg) as writer:
        df_dg_final.to_excel(
            writer,
            sheet_name="DIRECCION GENERAL",
            index=False
        )

    print("✔ Archivo generado:", archivo_salida_dg)

if __name__ == "__main__":
    procesar_combinaciones(
        archivo_entrada="DIRECCION GENERAL OCT 2025.xlsx",
        archivo_resultados="resultado_combinaciones.xlsx",
        archivo_salida_dg="DIRECCION_GENERAL_GENERADO.xlsx"
    )
