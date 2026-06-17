# -*- coding: utf-8 -*-
"""
=============================================================================
 PLATAFORMA DE DISOLUCIÓN DE FERTILIZANTES PARA FERTIRRIEGO
=============================================================================
 Tres modos de trabajo:
   1) DIRECTO   : dosis (kg/g) -> concentraciones, CE, compatibilidad, pH...
   2) INVERSO   : objetivo (ppm de cada nutriente) -> receta (gramos de cada
                  fertilizante), resuelta con minimos cuadrados no negativos
                  (NNLS) ponderados por error relativo.
   3) TEMPORADA : cultivo + Kc FAO-56 + ETo -> oferta hidrica, calendario de
                  riego editable y recetas de fertirriego por evento.

 Catálogo: se lee desde una Google Sheet PUBLICA fija (constante SHEET_URL).
           Si no está configurada o falla, usa un catálogo interno de respaldo.
 Ejecutar:  streamlit run fertirriego_app.py
 Requiere:  pip install streamlit pandas openpyxl scipy
=============================================================================
"""

import re
from io import BytesIO
from datetime import date, timedelta

import numpy as np
import pandas as pd
import streamlit as st

# ===========================================================================
#  >>>  URL FIJA DE LA GOOGLE SHEET DEL CATALOGO  <<<
#  Pega aqui el enlace de tu hoja PUBLICA (acceso por enlace, rol Lector).
#  Mientras este vacia, la app usa el catalogo interno por defecto.
#  Ejemplo: "https://docs.google.com/spreadsheets/d/XXXXXXXX/edit#gid=0"
# ===========================================================================
SHEET_URL = "https://docs.google.com/spreadsheets/d/1ICkXbGILdm0SAw-ubaHs6xdfEkBCdjinNP4P8YkiOxQ/edit?usp=sharing"

# Logo UCHILECREA embebido (SVG en base64), para no depender de archivos externos
LOGO_UCHILECREA_B64 = "PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiIHN0YW5kYWxvbmU9Im5vIj8+CjxzdmcKICAgdmVyc2lvbj0iMS4wIgogICB3aWR0aD0iODc2LjExNjMzIgogICBoZWlnaHQ9IjM1MS4xOTg3NiIKICAgdmlld0JveD0iMCAwIDg3Ni4xMTYzMyAzNTEuMTk4NzYiCiAgIHByZXNlcnZlQXNwZWN0UmF0aW89InhNaWRZTWlkIgogICBpZD0ic3ZnMjYiCiAgIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIKICAgeG1sbnM6c3ZnPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+CiAgPGRlZnMKICAgICBpZD0iZGVmczI2IiAvPgogIDxnCiAgICAgZmlsbD0iIzBhOTM1NCIKICAgICBpZD0iZzExIgogICAgIHRyYW5zZm9ybT0ibWF0cml4KDEuNDE0OTMxMiwwLDAsMS40MDg3NTQ2LC0yNzIuMDE5NDQsLTIxNy43NzI1MSkiPgogICAgPHBhdGgKICAgICAgIGQ9Im0gNTM5LjYyLDQwMy40OCBjIC05LjQzLC0xLjYgLTIwLjk2LC04Ljc4IC0yNy4yLC0xNy4wNCAtMTEuMDIsLTE0LjU4IC0xMS43NSwtMzIuMTMgLTIuMSwtNTIgMy43NywtNy45OCA3LjE4LC0xMi44NCAxNS44MSwtMjIuOTIgMy40OCwtMy45OSA3LjE4LC05LjA3IDguMjcsLTExLjI0IGwgMS45NiwtMy45OSB2IC0yMy45MyBjIDAsLTI3LjcxIC0wLjI5LC0yOS4zNyAtNC40MiwtMjkuMzcgLTYuNDYsMCAtMTkuNTEsLTcuNjIgLTI4LjU4LC0xNi42OCAtNi4wOSwtNi4wMiAtNy40LC03LjgzIC0xMC43MywtMTQuNTEgLTIuMzksLTQuODYgLTQuNzEsLTExLjAyIC02LjMxLC0xNy4wNCAtNS4wOCwtMTguNjQgLTcuMDQsLTI0LjggLTEwLjE1LC0zMS45MSAtMS45NiwtNC4zNSAtMi45NywtNy40NyAtMi42MSwtNy44MyAxLjIzLC0xLjIzIDUuNDQsMC4xNSAxMi40OCw0LjEzIDE3Ljc3LDEwLjAxIDQxLjQ5LDI5Ljk1IDQ4LjE2LDQwLjU0IDEuMjMsMS44OSAyLjM5LDMuNDEgMi42MSwzLjQxIDAuMjIsMCAwLjY1LC00LjM1IDAuOTQsLTkuNTcgMC45NCwtMTYuMjUgMi4xOCwtMjcuNjMgMy4xMiwtMjcuOTIgMS4yMywtMC40NCAxLjUyLDUuMDggMi4zOSw0MS40OSAxLjAyLDQxLjc4IDMuMjYsOTMuNDIgNC4zNSw5Ny43NyAxLjMxLDQuOTMgNC43OSw4LjQ5IDE1LjgxLDE1Ljk2IDExLjg5LDcuOTggMTguNDIsMTQgMjIuMTIsMjAuMTYgNi43NSwxMS40NiA4LjYzLDI4LjA3IDQuMjEsMzYuODQgLTIuMDMsMy45MiAtNy42OSw4Ljg1IC0xMy44NSwxMS45NyAtMi45NywxLjYgLTguNjMsNC45MyAtMTIuNDgsNy40NyAtMy44NCwyLjYxIC04LjU2LDUuMTUgLTEwLjQ0LDUuNzMgLTMuMTksMC45NCAtOS4yOCwxLjE2IC0xMy4zNSwwLjUxIHoiCiAgICAgICBpZD0icGF0aDUiIC8+CiAgICA8cGF0aAogICAgICAgZD0ibSA1ODIuMTksMjM4LjMzIGMgLTExLjI0LC00LjA2IC0xNC40MywtMTYuODMgLTcuNjksLTMwLjMyIDIuMDMsLTQuMjEgMTMuMTMsLTIwLjg5IDEzLjc4LC0yMC44OSAwLjg3LDAgOC4yNywxMC45NSAxMS43NSwxNy40OCA1LjI5LDkuODYgNi42NywxNi43NSA0LjcxLDIzLjI4IC0xLjAyLDMuNTUgLTUuODcsOS4xNCAtOS4wNywxMC40NCAtMi45NywxLjMxIC05Ljg2LDEuMjMgLTEzLjQ5LDAgeiIKICAgICAgIGlkPSJwYXRoMTAiIC8+CiAgICA8cGF0aAogICAgICAgZD0ibSA1NjguOTksMTkyLjY0IGMgLTUuMjIsLTIuMzIgLTYuODksLTguMjcgLTQuMDYsLTE0LjY1IDIuMTgsLTQuOTMgNi43NSwtMTEuODkgNy44MywtMTEuODkgMS44MSwwIDcuNjksMTAuMDggOC44NSwxNS4yMyAwLjg3LDMuOTIgMCw2Ljg5IC0yLjk3LDkuOTQgLTIuNjEsMi42OCAtNS43MywzLjEyIC05LjY1LDEuMzggeiIKICAgICAgIGlkPSJwYXRoMTEiIC8+CiAgPC9nPgogIDxnCiAgICAgZmlsbD0iIzA3ODE0OSIKICAgICBpZD0iZzI2IgogICAgIHN0eWxlPSJmaWxsOiNmZmZmZmY7ZmlsbC1vcGFjaXR5OjEiCiAgICAgdHJhbnNmb3JtPSJtYXRyaXgoMS40MTQ5MzEyLDAsMCwxLjQwODc1NDYsLTI3Mi4wMTk0NCwtMjE3Ljc3MjUxKSI+CiAgICA8cGF0aAogICAgICAgZD0ibSA1MzQuMjUsMjM5LjcxIGMgLTAuMjksLTEuNTIgLTEuMTYsLTMuNyAtMS44MSwtNC43MSAtMC43MywtMS4wMiAtMi45NywtNS4yMiAtNS4wOCwtOS40MyAtMy43LC03LjU0IC0xMi44NCwtMjEuMDMgLTIwLjE2LC0yOS44OCAtMy44NCwtNC41IC00LjY0LC02LjM4IC0yLjksLTYuMzggMS41MiwwIDEzLjQ5LDEzLjA2IDE4LjQyLDIwLjA5IDYuMTYsOC45MiAxMC4wMSwxNy4xOSAxMS40NiwyNS4xIDAuNjUsMy40OCAxLjA5LDYuNzUgMC44Nyw3LjI1IC0wLjE1LDAuNTEgLTAuNTEsLTAuNDQgLTAuOCwtMi4wMyB6IgogICAgICAgaWQ9InBhdGgyNiIKICAgICAgIHN0eWxlPSJmaWxsOiNmZmZmZmY7ZmlsbC1vcGFjaXR5OjEiIC8+CiAgPC9nPgogIDx0ZXh0CiAgICAgeG1sOnNwYWNlPSJwcmVzZXJ2ZSIKICAgICBzdHlsZT0iZm9udC1zaXplOjE5OC4yMDRweDtmb250LWZhbWlseTpBcmlhbDstaW5rc2NhcGUtZm9udC1zcGVjaWZpY2F0aW9uOkFyaWFsO3RleHQtYWxpZ246c3RhcnQ7d3JpdGluZy1tb2RlOmxyLXRiO2RpcmVjdGlvbjpsdHI7dGV4dC1hbmNob3I6c3RhcnQ7ZmlsbDojMGE5MzU0O2ZpbGwtb3BhY2l0eToxO3N0cm9rZS13aWR0aDoyLjA2NDYiCiAgICAgeD0iLTE0LjMyMzMzNiIKICAgICB5PSIzMjAuNTQ5MzUiCiAgICAgaWQ9InRleHQyNyIKICAgICB0cmFuc2Zvcm09InNjYWxlKDAuOTE4OTA4NDQsMS4wODgyNDc3KSI+PHRzcGFuCiAgICAgICBpZD0idHNwYW4yNyIKICAgICAgIHg9Ii0xNC4zMjMzMzYiCiAgICAgICB5PSIzMjAuNTQ5MzUiCiAgICAgICBzdHlsZT0iZm9udC1zdHlsZTpub3JtYWw7Zm9udC12YXJpYW50Om5vcm1hbDtmb250LXdlaWdodDpub3JtYWw7Zm9udC1zdHJldGNoOm5vcm1hbDtmb250LXNpemU6MTk4LjIwNHB4O2ZvbnQtZmFtaWx5OmNhbGlicmk7LWlua3NjYXBlLWZvbnQtc3BlY2lmaWNhdGlvbjpjYWxpYnJpO3N0cm9rZS13aWR0aDoyLjA2NDYiPjx0c3BhbgogICBzdHlsZT0iZm9udC1zdHlsZTpub3JtYWw7Zm9udC12YXJpYW50Om5vcm1hbDtmb250LXdlaWdodDpub3JtYWw7Zm9udC1zdHJldGNoOm5vcm1hbDtmb250LXNpemU6MTk4LjIwNHB4O2ZvbnQtZmFtaWx5OmNhbGlicmk7LWlua3NjYXBlLWZvbnQtc3BlY2lmaWNhdGlvbjpjYWxpYnJpO2ZpbGw6IzI0NzFhZjtmaWxsLW9wYWNpdHk6MC45OTYwNzg7c3Ryb2tlLXdpZHRoOjIuMDY0NiIKICAgaWQ9InRzcGFuMjgiPnVjaGlsZTx0c3BhbgogICBzdHlsZT0iZm9udC1zaXplOjE3My4zMzNweCIKICAgaWQ9InRzcGFuMjkiPiAgPC90c3Bhbj48L3RzcGFuPjx0c3BhbgogICBzdHlsZT0iZm9udC1zaXplOjE3My4zMzNweCIKICAgaWQ9InRzcGFuMzAiPiAgPC90c3Bhbj5jcmVhPC90c3Bhbj48L3RleHQ+Cjwvc3ZnPgo="

CFG = {
    "titulo": "Gestor Fertirriego",
    "volumen_default_L": 1000.0,
    "razon_iny_default": 100,
    "k_tds": 640.0,
    "eq": {"Ca": 20.04, "Mg": 12.15, "K": 39.10, "Na": 23.0,
           "N": 14.0, "S": 16.03, "P": 30.97, "Cl": 35.45, "HCO3": 61.0},
    "ox": {"P2O5_a_P": 0.4364, "K2O_a_K": 0.8301,
           "CaO_a_Ca": 0.7147, "MgO_a_Mg": 0.6032},
    "objetivo_default": {"N": 150.0, "P": 40.0, "K": 200.0,
                         "Ca": 150.0, "Mg": 45.0, "S": 60.0},
}

COLUMNAS_CATALOGO = [
    "nombre", "formula", "tipo",
    "N_pct", "P_pct", "K_pct", "Ca_pct", "Mg_pct", "S_pct", "Cl_pct",
    "fracN_NO3", "fracN_NH4", "fracN_ureico",
    "solubilidad_g_L", "ce_factor", "reaccion",
    "sulfato", "fosfato", "cloruro", "es_acido", "es_quelato_Fe",
    "meq_H_g", "densidad", "notas",
]

ELEMENTOS = ["N", "P", "K", "Ca", "Mg", "S", "Cl"]
MACROS_OBJETIVO = ["N", "P", "K", "Ca", "Mg", "S"]


def cargar_catalogo_default() -> pd.DataFrame:
    F = [
        ("Nitrato de calcio", "Ca(NO3)2*NH4NO3*10H2O", "solido",
         15.5, 0, 0, 19.0, 0, 0, 0, 0.94, 0.06, 0, 1200, 1.10, "basica",
         False, False, False, False, False, 0, None,
         "Base del Estanque A. Incompatible concentrado con sulfatos y fosfatos."),
        ("Nitrato de potasio", "KNO3", "solido",
         13.0, 0, 38.2, 0, 0, 0, 0, 1.0, 0, 0, 316, 1.30, "neutra",
         False, False, False, False, False, 0, None,
         "Compatible con ambos estanques."),
        ("Nitrato de amonio", "NH4NO3", "solido",
         34.0, 0, 0, 0, 0, 0, 0, 0.5, 0.5, 0, 1900, 1.50, "acida",
         False, False, False, False, False, 0, None,
         "Muy soluble. Acidificante por nitrificación del NH4."),
        ("Sulfato de amonio", "(NH4)2SO4", "solido",
         21.0, 0, 0, 0, 0, 24.0, 0, 0, 1.0, 0, 750, 1.80, "acida",
         True, False, False, False, False, 0, None,
         "Aporta SO4 -> Estanque B. Fuertemente acidificante."),
        ("Urea", "CO(NH2)2", "solido",
         46.0, 0, 0, 0, 0, 0, 0, 0, 0, 1.0, 1080, 0.05, "acida",
         False, False, False, False, False, 0, None,
         "No ionica: aporte despreciable a la CE."),
        ("Fosfato monoamonico (MAP)", "NH4H2PO4", "solido",
         12.0, 26.6, 0, 0, 0, 0, 0, 0, 1.0, 0, 400, 0.90, "acida",
         False, True, False, False, False, 0, None,
         "Aporta fosfato -> Estanque B. pH de solución ~4.5."),
        ("Fosfato monopotasico (MKP)", "KH2PO4", "solido",
         0, 22.7, 28.2, 0, 0, 0, 0, 0, 0, 0, 230, 0.70, "acida",
         False, True, False, False, False, 0, None,
         "Aporta fosfato -> Estanque B. pH de solución ~4.5."),
        ("Sulfato de potasio", "K2SO4", "solido",
         0, 0, 41.5, 0, 0, 18.0, 0, 0, 0, 0, 110, 1.70, "neutra",
         True, False, False, False, False, 0, None,
         "Solubilidad BAJA (~110 g/L). Aporta SO4 -> Estanque B."),
        ("Sulfato de magnesio (heptahidr.)", "MgSO4*7H2O", "solido",
         0, 0, 0, 0, 9.7, 13.0, 0, 0, 0, 0, 710, 0.80, "neutra",
         True, False, False, False, False, 0, None,
         "Sal de Epsom. Aporta SO4 -> Estanque B."),
        ("Nitrato de magnesio", "Mg(NO3)2*6H2O", "solido",
         11.0, 0, 0, 0, 9.5, 0, 0, 1.0, 0, 0, 1250, 1.00, "neutra",
         False, False, False, False, False, 0, None,
         "Mg sin sulfato -> compatible con Estanque A (Ca)."),
        ("Cloruro de potasio", "KCl", "solido",
         0, 0, 49.8, 0, 0, 0, 47.6, 0, 0, 0, 340, 1.70, "neutra",
         False, False, True, False, False, 0, None,
         "Alto Cl: cuidado en cultivos sensibles a cloruro."),
        ("Ácido fosforico 85%", "H3PO4", "liquido",
         0, 26.8, 0, 0, 0, 0, 0, 0, 0, 0, None, 1.00, "acida",
         False, True, False, True, False, 8.67, 1.685,
         "Liquido. Añadir SIEMPRE el ácido al agua. ~1 proton efectivo."),
        ("Ácido nitrico 65%", "HNO3", "liquido",
         14.4, 0, 0, 0, 0, 0, 0, 1.0, 0, 0, None, 1.20, "acida",
         False, False, False, True, False, 10.30, 1.39,
         "Liquido corrosivo. Añadir SIEMPRE el ácido al agua. EPP obligatorio."),
        ("Quelato de hierro EDTA (Fe 6%)", "Fe-EDTA", "solido",
         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 900, 0.20, "neutra",
         False, False, False, False, True, 0, None,
         "Fe 6%. Estable hasta pH ~6.5; sobre eso precipita el Fe."),
        ("Quelato de hierro EDDHA (Fe 6%)", "Fe-EDDHA", "solido",
         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 900, 0.15, "neutra",
         False, False, False, False, True, 0, None,
         "Fe 6%. Estable hasta pH ~9; ideal para aguas/suelos calcáreos."),
        ("Sulfato de manganeso", "MnSO4*H2O", "solido",
         0, 0, 0, 0, 0, 19.0, 0, 0, 0, 0, 700, 0.10, "neutra",
         True, False, False, False, False, 0, None,
         "Mn 32%. Microelemento (dosis bajas). Aporta SO4."),
        ("Sulfato de zinc (heptahidr.)", "ZnSO4*7H2O", "solido",
         0, 0, 0, 0, 0, 11.0, 0, 0, 0, 0, 580, 0.10, "neutra",
         True, False, False, False, False, 0, None,
         "Zn 22%. Microelemento (dosis bajas). Aporta SO4."),
        ("Ácido borico", "H3BO3", "solido",
         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 50, 0.05, "acida",
         False, False, False, False, False, 0, None,
         "B 17%. Microelemento. Solubilidad baja en frio."),
    ]
    return pd.DataFrame(F, columns=COLUMNAS_CATALOGO)


def url_csv_gsheet(url: str) -> str:
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", url)
    if not m:
        raise ValueError("No se reconoce un ID de hoja en la URL.")
    sid = m.group(1)
    g = re.search(r"[#&?]gid=(\d+)", url)
    gid = g.group(1) if g else "0"
    return f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv&gid={gid}"


@st.cache_data(show_spinner=False)
def cargar_desde_gsheet(url: str) -> pd.DataFrame:
    return pd.read_csv(url_csv_gsheet(url))


def validar_catalogo(df: pd.DataFrame):
    faltan = [c for c in COLUMNAS_CATALOGO if c not in df.columns]
    if faltan:
        return None, faltan
    num = ["N_pct", "P_pct", "K_pct", "Ca_pct", "Mg_pct", "S_pct", "Cl_pct",
           "fracN_NO3", "fracN_NH4", "fracN_ureico",
           "solubilidad_g_L", "ce_factor", "meq_H_g", "densidad"]
    def _to_num(series):
        return pd.to_numeric(
            series.astype(str)
                  .str.replace(",", ".", regex=False)
                  .replace({"nan": np.nan, "None": np.nan, "": np.nan}),
            errors="coerce"
        )

    for c in num:
        df[c] = _to_num(df[c])

    for c in ["N_pct", "P_pct", "K_pct", "Ca_pct", "Mg_pct", "S_pct", "Cl_pct",
              "fracN_NO3", "fracN_NH4", "fracN_ureico", "ce_factor", "meq_H_g"]:
        df[c] = df[c].fillna(0.0)
    for c in ["sulfato", "fosfato", "cloruro", "es_acido", "es_quelato_Fe"]:
        df[c] = df[c].astype(str).str.lower().isin(
            ["true", "1", "si", "sí", "x", "verdadero"])
    return df, []


def calcular_aportes(sel: pd.DataFrame, volumen_L: float) -> pd.DataFrame:
    df = sel.copy()
    df["g_L"] = df["gramos"] / volumen_L
    for el in ELEMENTOS:
        df[f"{el}_ppm"] = df["gramos"] * df[f"{el}_pct"] / 100.0 * 1000.0 / volumen_L
    df["CE_parcial"] = df["g_L"] * df["ce_factor"]
    df["excede_solub"] = df.apply(
        lambda r: bool(pd.notna(r["solubilidad_g_L"]) and r["g_L"] > r["solubilidad_g_L"]),
        axis=1)
    return df


def totales(df_aportes: pd.DataFrame, volumen_L: float) -> dict:
    t = {el: float(df_aportes[f"{el}_ppm"].sum()) for el in ELEMENTOS}
    t["CE"] = float(df_aportes["CE_parcial"].sum())
    t["masa_total_g"] = float(df_aportes["gramos"].sum())
    t["pct_pv"] = t["masa_total_g"] / (volumen_L * 10.0)
    t["TDS_aprox"] = t["CE"] * CFG["k_tds"]
    t["P2O5"] = t["P"] / CFG["ox"]["P2O5_a_P"] if t["P"] else 0.0
    t["K2O"] = t["K"] / CFG["ox"]["K2O_a_K"] if t["K"] else 0.0
    t["CaO"] = t["Ca"] / CFG["ox"]["CaO_a_Ca"] if t["Ca"] else 0.0
    t["MgO"] = t["Mg"] / CFG["ox"]["MgO_a_Mg"] if t["Mg"] else 0.0
    return t


def balance_ionico(sel: pd.DataFrame, volumen_L: float) -> pd.DataFrame:
    df = sel.copy()
    eq = CFG["eq"]
    Ca = (df["gramos"] * df["Ca_pct"] / 100 * 1000 / volumen_L).sum() / eq["Ca"]
    Mg = (df["gramos"] * df["Mg_pct"] / 100 * 1000 / volumen_L).sum() / eq["Mg"]
    K = (df["gramos"] * df["K_pct"] / 100 * 1000 / volumen_L).sum() / eq["K"]
    n_ppm = df["gramos"] * df["N_pct"] / 100 * 1000 / volumen_L
    NH4 = (n_ppm * df["fracN_NH4"]).sum() / eq["N"]
    NO3 = (n_ppm * df["fracN_NO3"]).sum() / eq["N"]
    SO4 = (df["gramos"] * df["S_pct"] / 100 * 1000 / volumen_L).sum() / eq["S"]
    P = (df["gramos"] * df["P_pct"] / 100 * 1000 / volumen_L).sum() / eq["P"]
    Cl = (df["gramos"] * df["Cl_pct"] / 100 * 1000 / volumen_L).sum() / eq["Cl"]
    filas = [("Ca++", Ca, "cation"), ("Mg++", Mg, "cation"),
             ("K+", K, "cation"), ("NH4+", NH4, "cation"),
             ("NO3-", NO3, "anion"), ("SO4--", SO4, "anion"),
             ("H2PO4-", P, "anion"), ("Cl-", Cl, "anion")]
    return pd.DataFrame(filas, columns=["ion", "meq_L", "grupo"])


def evaluar_compatibilidad(sel: pd.DataFrame) -> list:
    alertas = []
    ca = sel[sel["Ca_pct"] > 0]["nombre"].tolist()
    mg = sel[sel["Mg_pct"] > 0]["nombre"].tolist()
    sulf = sel[sel["sulfato"]]["nombre"].tolist()
    fosf = sel[sel["fosfato"]]["nombre"].tolist()
    cl = sel[sel["cloruro"]]["nombre"].tolist()
    acid = sel[sel["es_acido"]]["nombre"].tolist()
    quel = sel[sel["es_quelato_Fe"]]["nombre"].tolist()
    if ca and sulf:
        alertas.append(("ALTA", "Precipitacion de sulfato de calcio (CaSO4 / yeso)",
                        f"Ca: {', '.join(ca)} + sulfatos: {', '.join(sulf)}. "
                        "No mezclar en el mismo estanque concentrado."))
    if ca and fosf:
        alertas.append(("ALTA", "Precipitacion de fosfato de calcio",
                        f"Ca: {', '.join(ca)} + fosfatos: {', '.join(fosf)}. "
                        "Separar en estanques distintos (A / B)."))
    if mg and fosf:
        alertas.append(("MEDIA", "Riesgo de precipitación de fosfato de magnesio",
                        f"{', '.join(mg)} + {', '.join(fosf)}. Mayor riesgo a pH alto."))
    if quel and fosf:
        alertas.append(("MEDIA", "Posible interacción quelato de Fe + fosfato",
                        f"{', '.join(quel)} con {', '.join(fosf)}: vigilar pH."))
    if acid:
        alertas.append(("INFO", "Manejo de ácidos",
                        f"{', '.join(acid)}: añadir SIEMPRE el ácido sobre el agua, con EPP."))
    if cl:
        alertas.append(("INFO", "Aporte de cloruro",
                        f"{', '.join(cl)}: revisar tolerancia del cultivo al Cl-."))
    if not alertas:
        alertas.append(("OK", "Sin incompatibilidades evidentes",
                        "La combinación no dispara reglas de precipitación."))
    return alertas


def asignar_estanques(sel: pd.DataFrame) -> dict:
    """Propuesta de separación en dos estanques concentrados (A y B), con verificación.

    Regla dura: el calcio y los quelatos de Fe NO deben compartir estanque concentrado
    con sulfatos ni fosfatos (precipitan). Por eso:
      A = fuentes de calcio, quelatos de Fe y nitratos compatibles.
      B = sulfatos, fosfatos y ácidos.
    Las sales neutras sin Ca/SO4/PO4 (KNO3, KCl, urea, Mg(NO3)2, microelementos sin
    sulfato...) son flexibles y se colocan donde no generen incompatibilidad.

    Devuelve un dict con: tabla (nombre/estanque/motivo), listas de cada estanque,
    el estado verificado de A y B (errores graves y notas), y si hay conflicto interno.
    """
    filas = []
    for _, r in sel.iterrows():
        ca = bool(r["Ca_pct"] > 0)
        fe = bool(r["es_quelato_Fe"])
        su = bool(r["sulfato"])
        ph = bool(r["fosfato"])
        ac = bool(r["es_acido"])
        if ca and (su or ph):
            tk, motivo = "CONFLICTO", "el producto mezcla Ca con sulfato/fosfato"
        elif ca:
            tk, motivo = "A", "fuente de calcio"
        elif fe:
            tk, motivo = "A", "quelato de Fe (mantener lejos de fosfatos)"
        elif su or ph or ac:
            tk, motivo = "B", ("aporta sulfato" if su else
                               "aporta fosfato" if ph else "ácido")
        else:
            tk, motivo = "flexible", "sal neutra (compatible en cualquiera)"
        filas.append({"nombre": r["nombre"], "tk": tk, "motivo": motivo,
                      "Ca": ca, "Mg": bool(r["Mg_pct"] > 0),
                      "su": su, "ph": ph, "ac": ac, "fe": fe})
    a = pd.DataFrame(filas)

    # Colocar las flexibles: por defecto en A (no aportan Ca/SO4/PO4, seguras en cualquiera),
    # salvo que en A ya exista calcio y la flexible sea... siguen siendo seguras igual.
    a["estanque"] = a["tk"].replace({"flexible": "A"})

    def verificar(tk):
        sub = a[a["estanque"] == tk]
        graves, notas = [], []
        if sub.empty:
            return {"estado": "vacío", "graves": [], "notas": []}
        ca, su, ph = sub["Ca"].any(), sub["su"].any(), sub["ph"].any()
        mg, fe = sub["Mg"].any(), sub["fe"].any()
        if ca and su:
            graves.append("Ca + sulfato (precipita CaSO4)")
        if ca and ph:
            graves.append("Ca + fosfato (precipita fosfato de calcio)")
        if mg and ph:
            notas.append("Mg + fosfato: mantené este estanque ácido (pH < 6) para evitar precipitación")
        if fe and ph:
            notas.append("quelato de Fe + fosfato: vigilar el pH")
        return {"estado": "ERROR" if graves else ("ok con nota" if notas else "ok"),
                "graves": graves, "notas": notas}

    tabla = a[["nombre", "estanque", "motivo"]].copy()
    return {"tabla": tabla,
            "A": a[a["estanque"] == "A"]["nombre"].tolist(),
            "B": a[a["estanque"] == "B"]["nombre"].tolist(),
            "estado_A": verificar("A"), "estado_B": verificar("B"),
            "conflicto": a[a["tk"] == "CONFLICTO"]["nombre"].tolist()}


def estimar_ph_tendencia(sel: pd.DataFrame) -> str:
    if sel.empty:
        return "Sin datos"
    peso = {"acida": -1.0, "neutra": 0.0, "basica": 1.0}
    s = sel.copy()
    s["w"] = s["reaccion"].map(peso).fillna(0.0)
    s.loc[s["es_acido"], "w"] = -3.0
    score = (s["gramos"] * s["w"]).sum()
    if score < -50000:
        return "Fuertemente acidificante"
    if score < -5000:
        return "Acidificante"
    if score <= 5000:
        return "Neutra"
    if score <= 50000:
        return "Basificante"
    return "Fuertemente basificante"


def neutralizacion_bicarbonatos(hco3_mg_L, volumen_L, meq_H_total) -> dict:
    hco3_meq = (hco3_mg_L / CFG["eq"]["HCO3"]) * volumen_L
    residual_meq = hco3_meq - meq_H_total
    residual_mg_L = max(residual_meq, 0) / volumen_L * CFG["eq"]["HCO3"] if volumen_L else 0
    if hco3_meq == 0:
        estado = "El agua no aporta bicarbonatos (o no se ingreso dato)."
    elif residual_meq > 0.05 * hco3_meq:
        estado = "Sub-neutralizado: queda alcalinidad residual (pH alto)."
    elif residual_meq < -0.05 * hco3_meq:
        estado = "Sobre-neutralizado: exceso de ácido (riesgo de pH muy bajo)."
    else:
        estado = "Neutralización ~completa de los bicarbonatos."
    return {"hco3_meq": hco3_meq, "H_meq": meq_H_total,
            "residual_meq": residual_meq, "residual_mg_L": residual_mg_L,
            "estado": estado}


def resolver_receta_objetivo(catalogo, objetivos, volumen_L, disponibles,
                             ce_cap_madre=None):
    """Resuelve la receta (gramos por fertilizante) para acercarse a los objetivos
    de ppm, sin masas negativas y, opcionalmente, sin superar una CE crítica.

    ce_cap_madre: CE maxima permitida en el estanque madre (dS/m). None o <=0 = sin
    tope. Si NNLS ya cumple el tope, se devuelve esa solución (es la optima). Si el
    tope es activo, se reoptimiza con SLSQP minimizando el error relativo de los
    nutrientes sujeto a  CE(x) <= tope  y  x >= 0  (problema cuadratico convexo).
    """
    from scipy.optimize import nnls, minimize
    nut = [el for el in MACROS_OBJETIVO if objetivos.get(el, 0) and objetivos[el] > 0]
    sub = catalogo[catalogo["nombre"].isin(disponibles)].reset_index(drop=True)
    if sub.empty or not nut:
        return None

    # A[i,j] = ppm del nutriente i por gramo del fertilizante j
    A = np.zeros((len(nut), len(sub)))
    for i, el in enumerate(nut):
        A[i, :] = sub[f"{el}_pct"].values.astype(float) * 10.0 / volumen_L
    b = np.array([float(objetivos[el]) for el in nut], dtype=float)

    if not np.isfinite(A).all() or not np.isfinite(b).all():
        raise ValueError("El catálogo u objetivos contienen NaN/inf tras la lectura.")

    w = 1.0 / b
    Aw, bw = A * w[:, None], b * w

    # CE por gramo de cada fertilizante:  ce_factor / V   (dS/m por g en el estanque)
    ce_coef = np.nan_to_num(sub["ce_factor"].values.astype(float)) / volumen_L

    # 1) Solucion sin restriccion de CE (exacta y rapida)
    x, rnorm = nnls(Aw, bw)
    ce_nnls = float(ce_coef @ x)

    ce_limitada = False
    tope = float(ce_cap_madre) if (ce_cap_madre and ce_cap_madre > 0) else None
    if tope is not None and ce_nnls > tope + 1e-9:
        # 2) El tope es ACTIVO -> reoptimizar con la restriccion CE <= tope
        ce_limitada = True

        def f(z):
            r = Aw @ z - bw
            return float(r @ r)

        def jac(z):
            return 2.0 * (Aw.T @ (Aw @ z - bw))

        cons = [{"type": "ineq",
                 "fun": lambda z: tope - float(ce_coef @ z),
                 "jac": lambda z: -ce_coef}]
        bounds = [(0.0, None)] * len(x)
        x0 = x * min(1.0, tope / ce_nnls) if ce_nnls > 0 else np.zeros_like(x)
        res = minimize(f, x0, jac=jac, bounds=bounds, constraints=cons,
                       method="SLSQP", options={"maxiter": 800, "ftol": 1e-10})
        x = np.clip(res.x, 0.0, None)
        rnorm = float(np.linalg.norm(Aw @ x - bw))

    sub["gramos"] = x
    receta = sub[sub["gramos"] > 1e-4][["nombre", "gramos"]].reset_index(drop=True)
    ce_madre = float(ce_coef @ x)

    filas = []
    for el in MACROS_OBJETIVO:
        coef = sub[f"{el}_pct"].values.astype(float) * 10.0 / volumen_L
        logrado = float((coef * x).sum())
        meta = float(objetivos.get(el, 0) or 0)
        err = (logrado - meta) / meta * 100 if meta > 0 else None
        filas.append({"Nutriente": el, "Objetivo (ppm)": meta,
                      "Logrado (ppm)": logrado, "Error (%)": err,
                      "optimizado": el in nut})
    return {"receta": receta, "df_objetivo": pd.DataFrame(filas),
            "nut": nut, "rnorm": float(rnorm),
            "ce_madre": ce_madre, "ce_sin_tope": ce_nnls,
            "ce_limitada": ce_limitada, "ce_tope": tope}


def ferts_macro_por_defecto(catalogo) -> list:
    macro = (catalogo[["N_pct", "P_pct", "K_pct", "Ca_pct", "Mg_pct"]]
             .fillna(0).sum(axis=1) > 0)
    no_acido = ~catalogo["es_acido"].fillna(False)
    return catalogo[macro & no_acido]["nombre"].tolist()


def exportar_excel(df_aportes, t, alertas, estanques, bal, df_objetivo=None) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        resumen = pd.DataFrame({
            "Parametro": ["N (ppm)", "P (ppm)", "K (ppm)", "Ca (ppm)", "Mg (ppm)",
                          "S (ppm)", "Cl (ppm)", "CE (dS/m)", "% p/v",
                          "P2O5 (ppm)", "K2O (ppm)"],
            "Valor": [t["N"], t["P"], t["K"], t["Ca"], t["Mg"], t["S"], t["Cl"],
                      t["CE"], t["pct_pv"], t["P2O5"], t["K2O"]]})
        resumen.to_excel(w, sheet_name="Resumen", index=False)
        cols = ["nombre", "gramos", "g_L"] + [f"{e}_ppm" for e in ELEMENTOS] +\
               ["CE_parcial", "excede_solub"]
        df_aportes[cols].to_excel(w, sheet_name="Aportes", index=False)
        if df_objetivo is not None:
            df_objetivo.to_excel(w, sheet_name="Objetivo_vs_logrado", index=False)
        pd.DataFrame(alertas, columns=["Severidad", "Titulo", "Detalle"]).to_excel(
            w, sheet_name="Compatibilidad", index=False)
        estanques.to_excel(w, sheet_name="Estanques", index=False)
        bal.to_excel(w, sheet_name="Balance_ionico", index=False)
    return buf.getvalue()


def analizar_y_mostrar(sel, catalogo, volumen_L, razon, ce_agua, hco3_agua,
                       df_objetivo=None):
    if "N_pct" not in sel.columns:
        sel = sel.merge(catalogo, on="nombre", how="left")
    df_ap = calcular_aportes(sel, volumen_L)
    t = totales(df_ap, volumen_L)
    bal = balance_ionico(sel, volumen_L)
    alertas = evaluar_compatibilidad(sel)
    estanques = asignar_estanques(sel)
    ph_tend = estimar_ph_tendencia(sel)
    ce_madre = t["CE"]
    factor = 1.0 / razon
    ce_final = ce_agua + ce_madre / razon   # solución diluida en el gotero

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("CE estanque madre", f"{ce_madre:.2f} dS/m")
    c2.metric("CE final (gotero)", f"{ce_final:.2f} dS/m",
              help=f"CE del agua ({ce_agua:.2f}) + CE madre ÷ {razon}")
    c3.metric("Concentración total", f"{t['pct_pv']:.2f} % p/v")
    c4.metric("Tendencia de pH", ph_tend)

    sev_max = max((a[0] for a in alertas),
                  key=lambda s: {"OK": 0, "INFO": 1, "MEDIA": 2, "ALTA": 3}.get(s, 0))
    if sev_max == "ALTA":
        st.error("Hay incompatibilidades de severidad ALTA (ver pestaña Compatibilidad).")
    elif sev_max == "MEDIA":
        st.warning("Hay advertencias de compatibilidad de severidad MEDIA.")

    tabs = st.tabs(["Nutrientes", "Compatibilidad",
                    "Solubilidad", "pH / bicarbonatos", "Balance iónico"])
    with tabs[0]:
        st.markdown("**Concentración de nutrientes (ppm = mg/L):**")
        nut = pd.DataFrame({
            "Nutriente": ELEMENTOS,
            "Estanque madre (ppm)": [t[e] for e in ELEMENTOS],
            f"Final 1:{razon} (ppm)": [t[e] * factor for e in ELEMENTOS]})
        st.dataframe(nut.round(2), use_container_width=True, hide_index=True)
        st.bar_chart(nut.set_index("Nutriente")["Estanque madre (ppm)"])
        with st.expander("Equivalente en óxidos (P2O5, K2O, CaO, MgO)"):
            st.dataframe(pd.DataFrame({
                "Forma": ["P2O5", "K2O", "CaO", "MgO"],
                "ppm (madre)": [t["P2O5"], t["K2O"], t["CaO"], t["MgO"]]}).round(1),
                hide_index=True, use_container_width=True)
        with st.expander("Relaciones de interes"):
            def ratio(a, b):
                return round(t[a] / t[b], 2) if t[b] else None
            st.write(f"- N : K  =  {ratio('N','K')}")
            st.write(f"- K : Ca  =  {ratio('K','Ca')}")
            st.write(f"- Ca : Mg  =  {ratio('Ca','Mg')}")
        st.markdown("**Aporte por fertilizante (g/L y ppm):**")
        cols = ["nombre", "gramos", "g_L"] + [f"{e}_ppm" for e in ELEMENTOS]
        st.dataframe(df_ap[cols].round(2), use_container_width=True, hide_index=True)
    with tabs[1]:
        st.markdown("**Alertas de compatibilidad / precipitación:**")
        colores = {"ALTA": st.error, "MEDIA": st.warning, "INFO": st.info, "OK": st.success}
        for sev, titulo, detalle in alertas:
            colores.get(sev, st.info)(f"**[{sev}] {titulo}** — {detalle}")
        st.divider()
        st.markdown("**Propuesta de estanques** "
                    "(A = calcio · quelatos · nitratos | B = sulfatos · fosfatos · ácidos):")
        if estanques["conflicto"]:
            st.error("No se puede separar en dos estanques: "
                     + ", ".join(estanques["conflicto"]) +
                     " mezcla(n) calcio con sulfato/fosfato en el mismo producto.")
        ca, cb = st.columns(2)
        for col, tk in ((ca, "A"), (cb, "B")):
            contenido = estanques[tk]
            estado = estanques[f"estado_{tk}"]
            with col:
                st.markdown(f"**Estanque {tk}**")
                if not contenido:
                    st.caption("(vacío)")
                else:
                    for n in contenido:
                        st.write(f"• {n}")
                if estado["graves"]:
                    st.error("Incompatible: " + "; ".join(estado["graves"]))
                elif estado["notas"]:
                    st.warning("; ".join(estado["notas"]))
                elif contenido:
                    st.success("Verificado: sin incompatibilidades.")
        if not estanques["conflicto"] and not estanques["estado_A"]["graves"] \
                and not estanques["estado_B"]["graves"]:
            st.caption("Propuesta verificada: el calcio queda separado de sulfatos y "
                       "fosfatos. Prepara cada estanque por separado y dilúyelos juntos "
                       "solo en el agua de riego ya diluida.")
    with tabs[2]:
        st.markdown("**Solubilidad** (dosis g/L vs límite a ~20 C):")
        sol = df_ap[["nombre", "g_L", "solubilidad_g_L", "excede_solub"]].rename(
            columns={"g_L": "dosis (g/L)", "solubilidad_g_L": "solubilidad (g/L)",
                     "excede_solub": "excede?"})
        st.dataframe(sol.round(1), use_container_width=True, hide_index=True)
        if df_ap["excede_solub"].any():
            st.error("Hay dosis que superan la solubilidad: no se disolveran del todo. "
                     "Reduce dosis, sube volumen o usa agua más tibia.")
        else:
            st.success("Todas las dosis están por debajo del límite de solubilidad.")
    with tabs[3]:
        st.markdown(f"**Tendencia de pH de la mezcla:** {ph_tend}")
        st.caption("Estimación cualitativa ponderada por masa; el pH real debe medirse.")
        st.divider()
        st.markdown("**Neutralización de bicarbonatos** (ácidos vs HCO3- del agua):")
        meq_H = float((df_ap["gramos"] * df_ap["meq_H_g"]).sum())
        res = neutralizacion_bicarbonatos(hco3_agua, volumen_L, meq_H)
        cc1, cc2, cc3 = st.columns(3)
        cc1.metric("meq HCO3- en estanque", f"{res['hco3_meq']:.0f}")
        cc2.metric("meq H+ por ácidos", f"{res['H_meq']:.0f}")
        cc3.metric("HCO3- residual", f"{res['residual_mg_L']:.0f} mg/L")
        st.info(res["estado"])
    with tabs[4]:
        st.markdown("**Balance iónico (meq/L)** — cationes vs aniones:")
        st.dataframe(bal.round(2), use_container_width=True, hide_index=True)
        cat = bal[bal["grupo"] == "cation"]["meq_L"].sum()
        ani = bal[bal["grupo"] == "anion"]["meq_L"].sum()
        d1, d2, d3 = st.columns(3)
        d1.metric("Suma cationes", f"{cat:.2f} meq/L")
        d2.metric("Suma aniones", f"{ani:.2f} meq/L")
        d3.metric("Desbalance", f"{((cat-ani)/cat*100 if cat else 0):+.1f} %")
        st.bar_chart(bal.set_index("ion")["meq_L"])
        st.caption("Un desbalance grande suele indicar un ion acompanante no contemplado.")

    st.subheader("Exportar", anchor=False)
    st.download_button(
        "Descargar receta en Excel",
        data=exportar_excel(df_ap, t, alertas, estanques["tabla"], bal, df_objetivo),
        file_name="receta_fertirriego.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def sidebar_config():
    with st.sidebar:
        st.markdown(
            f'<div style="text-align:center;margin-bottom:0.5rem;">'
            f'<img src="data:image/svg+xml;base64,{LOGO_UCHILECREA_B64}" '
            f'style="width:78%;max-width:220px;"/></div>',
            unsafe_allow_html=True)
        st.divider()
        st.header("Configuración", anchor=False)
        volumen_L = st.number_input("Capacidad del estanque (L)", min_value=1.0,
                                    value=CFG["volumen_default_L"], step=50.0)
        razon = st.number_input("Razón de inyección 1:R", min_value=1,
                                value=CFG["razon_iny_default"], step=1,
                                help="Por cada 1 L de solución madre, R litros de agua de "
                                     "riego. Usa 1 si el estanque es la solución final.")
        st.divider()
        st.subheader("Agua de riego", anchor=False)
        ce_agua = st.number_input("CE del agua (dS/m)", min_value=0.0, value=0.0, step=0.1)
        hco3_agua = st.number_input("Bicarbonatos HCO3- (mg/L)", min_value=0.0,
                                    value=0.0, step=10.0)
        st.divider()
        st.subheader("Catálogo", anchor=False)
        catalogo = cargar_catalogo_fija()
    return volumen_L, razon, ce_agua, hco3_agua, catalogo


def cargar_catalogo_fija() -> pd.DataFrame:
    """Carga el catálogo desde la Google Sheet fija (SHEET_URL).
    Si la URL está vacía o falla, usa el catálogo interno como respaldo.
    Muestra el estado en la barra lateral."""
    if not SHEET_URL.strip():
        st.info("Usando catálogo interno (la URL de la hoja aún no está configurada).")
        return cargar_catalogo_default()
    try:
        df_g, faltan = validar_catalogo(cargar_desde_gsheet(SHEET_URL))
        if faltan:
            st.error("La hoja no tiene las columnas requeridas: " + ", ".join(faltan) +
                     ". Se usa el catálogo interno.")
            return cargar_catalogo_default()
        st.success(f"Catálogo cargado desde Google Sheet: {len(df_g)} fertilizantes.")
        return df_g
    except Exception as e:
        st.warning(f"No se pudo leer la hoja fija ({e}). Se usa el catálogo interno.")
        return cargar_catalogo_default()


def modo_directo(catalogo, volumen_L, razon, ce_agua, hco3_agua):
    st.subheader("Fertilizantes y dosis", anchor=False)
    unidad = st.radio("Unidad de dosis", ["kg", "g"], horizontal=True)
    st.caption(f"Elige fertilizante y dosis en **{unidad}** para **{volumen_L:g} L**.")
    base = pd.DataFrame({"nombre": pd.Series(dtype="str"),
                         "dosis": pd.Series(dtype="float")})
    editor = st.data_editor(
        base, num_rows="dynamic", use_container_width=True, key="sel_dir",
        column_config={
            "nombre": st.column_config.SelectboxColumn(
                "Fertilizante", options=sorted(catalogo["nombre"].tolist()), width="large"),
            "dosis": st.column_config.NumberColumn(f"Dosis ({unidad})", min_value=0.0,
                                                   step=0.1, format="%.2f")})
    sel = editor.dropna(subset=["nombre"]).copy()
    sel = sel[sel["nombre"].isin(catalogo["nombre"])]
    if sel.empty:
        st.info("Agrega al menos un fertilizante para ver resultados.")
        return
    sel["gramos"] = sel["dosis"].fillna(0) * (1000.0 if unidad == "kg" else 1.0)
    st.subheader("Resultados", anchor=False)
    analizar_y_mostrar(sel[["nombre", "gramos"]], catalogo, volumen_L, razon,
                       ce_agua, hco3_agua)


def modo_inverso(catalogo, volumen_L, razon, ce_agua, hco3_agua):
    st.subheader("Objetivo (ppm)", anchor=False)
    st.caption("Objetivo de concentración en **el estanque** (la solución que prepararás). "
               "Deja 0 para no exigir ese nutriente (igual se reporta si aparece como "
               "efecto colateral).")
    cols = st.columns(6)
    objetivos = {}
    for i, el in enumerate(MACROS_OBJETIVO):
        objetivos[el] = cols[i].number_input(
            f"{el}", min_value=0.0, value=float(CFG["objetivo_default"][el]),
            step=5.0, key=f"obj_{el}")
    st.subheader("Fertilizantes disponibles", anchor=False)
    opciones = sorted(catalogo["nombre"].tolist())
    default = [n for n in ferts_macro_por_defecto(catalogo) if n in opciones]
    disponibles = st.multiselect(
        "El solver elegira las cantidades solo entre estos:",
        options=opciones, default=default,
        help="Por defecto se sugieren fuentes de macronutrientes (sin ácidos ni micros).")
    if not disponibles:
        st.info("Selecciona al menos un fertilizante disponible.")
        return

    st.subheader("CE crítica", anchor=False)
    cce1, cce2 = st.columns([1, 1.4])
    ce_crit = cce1.number_input(
        "CE crítica (dS/m)", min_value=0.0, value=0.0, step=0.1,
        help="Tope de conductividad que la receta NO debe superar. 0 = sin límite.")
    refiere = cce2.radio(
        "La CE crítica se refiere a:",
        ["Solución final (gotero)", "Estanque madre"], horizontal=True,
        help="Final = lo que recibe la planta (CE agua + CE madre ÷ R). "
             "Madre = la solución concentrada del estanque.")

    # Traducir el tope a un límite sobre la CE del ESTANQUE MADRE (lo que ve el solver)
    ce_cap_madre = None
    if ce_crit > 0:
        if refiere.startswith("Solución final"):
            ce_cap_madre = (ce_crit - ce_agua) * razon
            if ce_cap_madre <= 0:
                st.error(f"La CE del agua ({ce_agua:.2f}) ya iguala o supera la CE crítica "
                         f"final ({ce_crit:.2f}). Sin margen para fertilizar: sube la CE "
                         "crítica o reduce la CE del agua.")
                return
        else:
            ce_cap_madre = ce_crit

    if not st.button("Resolver receta", type="primary"):
        st.stop()
    res = resolver_receta_objetivo(catalogo, objetivos, volumen_L, disponibles,
                                   ce_cap_madre=ce_cap_madre)
    if res is None:
        st.error("No se pudo resolver: revisa objetivos > 0 y fertilizantes.")
        return
    receta, df_obj = res["receta"], res["df_objetivo"]
    if receta.empty:
        st.warning("El solver no asigno masa a ningun fertilizante. "
                   "Revisa que los elegidos aporten los nutrientes pedidos.")
        return
    st.subheader("Receta propuesta", anchor=False)
    vista = receta.copy()
    vista["kg"] = vista["gramos"] / 1000.0
    vista["g/L"] = vista["gramos"] / volumen_L
    st.dataframe(
        vista.rename(columns={"nombre": "Fertilizante", "gramos": "g (total)"})
             .round({"g (total)": 1, "kg": 3, "g/L": 3}),
        use_container_width=True, hide_index=True)
    st.subheader("Ajuste a objetivos", anchor=False)
    show = df_obj[df_obj["optimizado"] | (df_obj["Objetivo (ppm)"] > 0)
                  | (df_obj["Logrado (ppm)"] > 1)].copy()
    st.dataframe(show.drop(columns=["optimizado"]).round(1),
                 use_container_width=True, hide_index=True)
    problemas = []
    for _, r in df_obj.iterrows():
        if r["optimizado"] and r["Error (%)"] is not None and abs(r["Error (%)"]) > 10:
            problemas.append(f"{r['Nutriente']} ({r['Error (%)']:+.0f}%)")
    if problemas:
        st.warning("Objetivos no alcanzados (error >10%): " + ", ".join(problemas) +
                   ". Suele faltar una fuente adecuada de ese nutriente (p. ej. Ca sin "
                   "nitrato de calcio entre los disponibles).")
    else:
        st.success("Todos los nutrientes objetivo se alcanzaron con error <= 10%.")
    # --- Estado del tope de CE ---
    if res.get("ce_tope"):
        ce_madre = res["ce_madre"]
        ce_final = ce_agua + ce_madre / razon
        m1, m2, m3 = st.columns(3)
        m1.metric("CE estanque madre", f"{ce_madre:.2f} dS/m")
        m2.metric("CE final (gotero)", f"{ce_final:.2f} dS/m")
        etiqueta = "final" if refiere.startswith("Solución final") else "madre"
        m3.metric(f"CE crítica ({etiqueta})", f"{ce_crit:.2f} dS/m")
        if res["ce_limitada"]:
            # Cuantificar el compromiso real: cuanto bajaron los nutrientes objetivo
            faltantes = [r for _, r in df_obj.iterrows()
                         if r["optimizado"] and r["Error (%)"] is not None
                         and r["Error (%)"] < -2]
            recorte = (1 - ce_madre / res["ce_sin_tope"]) * 100 if res["ce_sin_tope"] else 0
            msg = (f"La CE crítica está activa: sin ella la receta llegaría a "
                   f"{res['ce_sin_tope']:.2f} dS/m, y se ajustó a {ce_madre:.2f} dS/m para "
                   f"no superar el máximo de {ce_crit:.2f}.")
            if faltantes:
                detalle = ", ".join(f"{r['Nutriente']} {r['Error (%)']:+.0f}%"
                                    for r in faltantes)
                st.warning(
                    msg + f" Esto obliga a un compromiso: para bajar la CE un {recorte:.0f}%, "
                    f"estos nutrientes quedan bajo su objetivo: {detalle}. Si necesitas "
                    "acercarte más, sube la CE crítica o el volumen del estanque, o usa "
                    "fuentes menos salinas.")
            else:
                st.info(msg + " Aun así, todos los nutrientes objetivo se mantienen cerca "
                        "de su meta (ver el ajuste de arriba).")
        else:
            st.success("La receta cumple la CE crítica con holgura (el tope no fue activo).")

    cap_txt = (f" · tope CE madre={res['ce_tope']:.2f} dS/m"
               f"{' (ACTIVO)' if res['ce_limitada'] else ''}") if res.get("ce_tope") else ""
    metodo = "SLSQP con restricción de CE" if res["ce_limitada"] else "NNLS"
    st.caption(f"Solver: {metodo}, ponderado por error relativo · residuo={res['rnorm']:.3f} "
               f"· optimizados: {', '.join(res['nut'])}{cap_txt}.")
    st.divider()
    st.subheader("Análisis de la receta", anchor=False)
    analizar_y_mostrar(receta, catalogo, volumen_L, razon, ce_agua, hco3_agua,
                       df_objetivo=df_obj)


# ===========================================================================
#  MODULO TEMPORADA  ·  FAO-56 (Kc, ETc), oferta hidrica, calendario de riego
#  y recetas de fertirriego por evento.
# ===========================================================================

# Etapas FAO-56 (etiquetas usadas en toda la app de temporada)
ETAPAS_FAO = ["Inicial", "Desarrollo", "Media", "Final"]

# Especies y coeficientes unicos de cultivo Kc segun FAO-56 (Estudio Riego y
# Drenaje 56, Cuadros 11 y 12). Kc_ini / Kc_mid / Kc_end y duracion de etapas
# en dias (L_ini, L_dev, L_mid, L_late). p = fraccion de agotamiento admisible.
# Estos valores quedan FIJOS en el codigo y son editables por el usuario en UI.
CROPS = {
    "Papa / Patata (Solanum tuberosum)": {
        "kc_ini": 0.50, "kc_mid": 1.15, "kc_end": 0.75,
        "L_ini": 30, "L_dev": 35, "L_mid": 50, "L_late": 30,
        "p": 0.35, "altura_m": 0.6,
        "ref": "FAO-56 Cuadro 12 (0.50/1.15/0.75); Cuadro 11 Europa Abr 30/35/50/30."},
    "Vid de mesa / pasas (Vitis vinifera)": {
        "kc_ini": 0.30, "kc_mid": 0.85, "kc_end": 0.45,
        "L_ini": 20, "L_dev": 50, "L_mid": 75, "L_late": 60,
        "p": 0.45, "altura_m": 2.0,
        "ref": "FAO-56 Cuadro 12 uvas mesa (0.30/0.85/0.45); Cuadro 11 Calif 20/50/75/60."},
    "Vid vinifera (Vitis vinifera)": {
        "kc_ini": 0.30, "kc_mid": 0.70, "kc_end": 0.45,
        "L_ini": 30, "L_dev": 60, "L_mid": 40, "L_late": 80,
        "p": 0.45, "altura_m": 1.8,
        "ref": "FAO-56 Cuadro 12 uvas vino (0.30/0.70/0.45); Cuadro 11 lat. medias 30/60/40/80."},
    "Cerezo (Prunus avium, sin cobertura, con heladas)": {
        "kc_ini": 0.45, "kc_mid": 0.95, "kc_end": 0.70,
        "L_ini": 20, "L_dev": 70, "L_mid": 90, "L_late": 30,
        "p": 0.50, "altura_m": 4.0,
        "ref": "FAO-56 Cuadro 12 Manzanas/Cerezas/Peras sin cobertura, fuertes heladas (0.45/0.95/0.70)."},
    "Maiz grano (Zea mays)": {
        "kc_ini": 0.30, "kc_mid": 1.20, "kc_end": 0.60,
        "L_ini": 30, "L_dev": 40, "L_mid": 50, "L_late": 30,
        "p": 0.55, "altura_m": 2.0,
        "ref": "FAO-56 Cuadro 12 maiz grano (0.30/1.20/0.60); Cuadro 11 Esp/Calif 30/40/50/30."},
    "Tomate (Solanum lycopersicum)": {
        "kc_ini": 0.60, "kc_mid": 1.15, "kc_end": 0.80,
        "L_ini": 30, "L_dev": 40, "L_mid": 40, "L_late": 25,
        "p": 0.40, "altura_m": 0.6,
        "ref": "FAO-56 Cuadro 12 tomate (0.60/1.15/0.70-0.90)."},
}

OBJETIVO_NUTRI_DEFAULT = {
    "Inicial":    {"N": 18.0, "P": 4.0,  "K": 25.0,  "Ca": 4.0,  "Mg": 2.5, "S": 3.5},
    "Desarrollo": {"N": 54.0, "P": 11.0, "K": 75.0,  "Ca": 12.0, "Mg": 7.5, "S": 10.5},
    "Media":      {"N": 81.0, "P": 16.0, "K": 112.0, "Ca": 18.0, "Mg": 11.0, "S": 16.0},
    "Final":      {"N": 27.0, "P": 4.0,  "K": 38.0,  "Ca": 6.0,  "Mg": 4.0, "S": 5.0},
}

DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]


def kc_diario(crop: dict, dia: int, L: tuple) -> tuple:
    """Curva Kc de FAO-56 para el dia 'dia' (1-indexado desde la siembra).
    Meseta en Kc_ini durante L_ini, rampa lineal en L_dev hasta Kc_mid, meseta
    en Kc_mid durante L_mid y descenso lineal en L_late hasta Kc_end."""
    ci, cm, cf = crop["kc_ini"], crop["kc_mid"], crop["kc_end"]
    li, ld, lm, ll = L
    if dia <= li:
        return ci, "Inicial"
    if dia <= li + ld:
        return ci + (cm - ci) * (dia - li) / max(ld, 1), "Desarrollo"
    if dia <= li + ld + lm:
        return cm, "Media"
    return cm + (cf - cm) * (dia - (li + ld + lm)) / max(ll, 1), "Final"


def parsear_eto_agromet(file) -> tuple:
    """Lee un xlsx de agrometeorologia.cl (Red INIA). Devuelve (df, estaciones).
    df: columnas 'fecha' + una columna de ETo (mm) por estacion. Soporta varias
    estaciones (cada una con su columna '... % de datos' que se ignora)."""
    raw = pd.read_excel(file, header=None, engine="openpyxl")
    hdr = None
    for i in range(len(raw)):
        v = str(raw.iloc[i, 0]).strip().lower()
        if v.startswith("tiempo"):
            hdr = i
            break
    if hdr is None:
        raise ValueError("No se encontró la fila de encabezado ('Tiempo ...').")
    header = raw.iloc[hdr].tolist()
    data = raw.iloc[hdr + 1:].reset_index(drop=True)
    estaciones = {}
    for j in range(1, len(header)):
        name = str(header[j]).strip()
        low = name.lower()
        if low in ("none", "nan", "") or low.endswith("% de datos"):
            continue
        estaciones[name] = j
    if not estaciones:
        raise ValueError("No se reconocieron columnas de estaciones de ETo.")
    fechas = pd.to_datetime(data.iloc[:, 0].astype(str), dayfirst=True, errors="coerce")
    out = pd.DataFrame({"fecha": fechas})
    for name, j in estaciones.items():
        out[name] = pd.to_numeric(
            data.iloc[:, j].astype(str).str.replace(",", ".", regex=False),
            errors="coerce")
    out = out[out["fecha"].notna()].reset_index(drop=True)
    return out, list(estaciones.keys())


def construir_serie_diaria(crop, fecha_siembra, L, factores, eto_df, estacion,
                           relleno="media") -> tuple:
    """Serie diaria de la temporada: fecha, dia, etapa, Kc, ETo, ETc y oferta de
    riego (factor de irrigacion por etapa x ETc). Rellena dias sin ETo medida.

    Devuelve (serie, info) donde info trae cobertura y valor de relleno usado."""
    li, ld, lm, ll = L
    total = int(li + ld + lm + ll)
    eto_map = {}
    if eto_df is not None and estacion in eto_df.columns:
        for _, r in eto_df.iterrows():
            if pd.notna(r[estacion]):
                eto_map[pd.Timestamp(r["fecha"]).date()] = float(r[estacion])
    fechas = [fecha_siembra + timedelta(days=d) for d in range(total)]
    medidos = [eto_map.get(f) for f in fechas]
    disponibles = [v for v in medidos if v is not None]
    if relleno == "media" and disponibles:
        val_relleno = float(np.mean(disponibles))
    elif isinstance(relleno, (int, float)):
        val_relleno = float(relleno)
    else:
        val_relleno = float(np.mean(disponibles)) if disponibles else 0.0
    filas = []
    n_rellenos = 0
    for d in range(1, total + 1):
        fecha = fecha_siembra + timedelta(days=d - 1)
        kc, etapa = kc_diario(crop, d, L)
        medido = eto_map.get(fecha)
        if medido is None:
            eto = val_relleno
            origen = "relleno"
            n_rellenos += 1
        else:
            eto = medido
            origen = "medido"
        etc = kc * eto
        f = float(factores.get(etapa, 1.0))
        filas.append({
            "dia": d, "fecha": fecha, "dia_semana": DIAS_SEMANA[fecha.weekday()],
            "etapa": etapa, "Kc": round(kc, 3), "ETo_mm": round(eto, 2),
            "origen_ETo": origen, "factor_riego": f,
            "ETc_mm": round(etc, 3), "oferta_riego_mm": round(f * etc, 3)})
    serie = pd.DataFrame(filas)
    info = {"total_dias": total, "n_medidos": len(disponibles),
            "n_rellenos": n_rellenos, "val_relleno": val_relleno,
            "cobertura_pct": 100.0 * len(disponibles) / total if total else 0.0}
    return serie, info


def generar_calendario(serie, freq_etapa, dias_operativos, cfg_riego) -> pd.DataFrame:
    """Calendario de riego propuesto a partir de la serie diaria.

    Para cada dia acumula la lamina neta ajustada = oferta_riego x Kr - Pe (>=0).
    Dispara un evento cuando los dias desde el ultimo riego alcanzan la frecuencia
    de la etapa Y el dia es operativo. Si la frecuencia se cumple en dia no
    operativo, sigue acumulando hasta el proximo dia operativo (conserva el agua).
    El ultimo dia cierra el evento con lo acumulado.

    lamina_bruta = lamina_neta / (Ea x Es) ; tiempo_h = lamina_bruta / PPeq."""
    Ea = max(cfg_riego["Ea"], 1e-6)
    Es = max(cfg_riego["Es"], 1e-6)
    PPeq = cfg_riego["PPeq"]
    Kr = cfg_riego["Kr"]
    Pe = cfg_riego["Pe_mm_dia"]
    eventos = []
    acum_neta = 0.0
    dias_desde = 0
    etc_acum = 0.0
    n = len(serie)
    n_evt = 0
    for idx, row in serie.iterrows():
        neta_dia = max(row["oferta_riego_mm"] * Kr - Pe, 0.0)
        acum_neta += neta_dia
        etc_acum += row["ETc_mm"]
        dias_desde += 1
        etapa = row["etapa"]
        freq = int(freq_etapa.get(etapa, 3))
        fecha = row["fecha"]
        operativo = fecha.weekday() in dias_operativos
        es_ultimo = (idx == n - 1)
        debe = dias_desde >= freq
        if ((debe and operativo) or es_ultimo) and acum_neta > 1e-9:
            n_evt += 1
            lb = acum_neta / (Ea * Es)
            t_h = lb / PPeq if PPeq > 0 else 0.0
            eventos.append({
                "evento": n_evt, "fecha": fecha,
                "dia_semana": DIAS_SEMANA[fecha.weekday()],
                "etapa": etapa, "intervalo_dias": dias_desde,
                "ETc_acum_mm": round(etc_acum, 2),
                "lamina_neta_mm": round(acum_neta, 2),
                "lamina_bruta_mm": round(lb, 2),
                "tiempo_riego_h": round(t_h, 2),
                "incluir": True})
            acum_neta = 0.0
            dias_desde = 0
            etc_acum = 0.0
    return pd.DataFrame(eventos)


def distribuir_nutrientes(cal, objetivos_ha, sup_ha, modo="agua") -> pd.DataFrame:
    """Reparte los kg/ha objetivo de cada etapa entre los eventos de esa etapa.
    modo='agua' -> proporcional a la lamina bruta (concentracion ~constante);
    modo='igual' -> partes iguales por evento."""
    cal = cal.copy()
    for nut in MACROS_OBJETIVO:
        cal[f"kg_{nut}"] = 0.0
    for etapa in cal["etapa"].unique():
        sub = cal[cal["etapa"] == etapa]
        if sub.empty:
            continue
        if modo == "agua":
            peso = sub["lamina_bruta_mm"].astype(float)
        else:
            peso = pd.Series(1.0, index=sub.index)
        tot = float(peso.sum())
        for nut in MACROS_OBJETIVO:
            kg_ha = float(objetivos_ha.loc[etapa, nut]) if etapa in objetivos_ha.index else 0.0
            kg_campo = kg_ha * sup_ha
            if tot > 0:
                cal.loc[sub.index, f"kg_{nut}"] = kg_campo * peso / tot
    return cal


def receta_evento(catalogo, kg_event, lamina_bruta_mm, sup_ha, disponibles,
                  razon, ce_agua, ce_crit_gotero) -> dict:
    """Receta de un evento de fertirriego. Resuelve los gramos totales (al campo)
    para alcanzar los ppm objetivo en el gotero, derivados de los kg objetivo del
    evento y del volumen de agua aplicado.

    V_riego (L) = lamina_bruta_mm x sup_ha x 10000   (1 mm.ha = 10000 L)
    ppm_gotero  = kg_nut x 1e6 / V_riego
    El solver entrega los gramos para esos ppm sobre V_riego (= masa por evento).
    Estanque madre: la solucion inyectada = V_riego / razon."""
    V = float(lamina_bruta_mm) * float(sup_ha) * 10000.0
    if V <= 0:
        return None
    objetivos_ppm = {n: (kg_event.get(n, 0.0) * 1e6 / V) for n in MACROS_OBJETIVO}
    ce_cap = (ce_crit_gotero - ce_agua) if ce_crit_gotero and ce_crit_gotero > 0 else None
    if ce_cap is not None and ce_cap <= 0:
        ce_cap = 1e-6
    try:
        res = resolver_receta_objetivo(catalogo, objetivos_ppm, V, disponibles,
                                       ce_cap_madre=ce_cap)
    except Exception as e:
        return {"error": str(e)}
    if res is None or res["receta"].empty:
        return {"error": "sin solución", "objetivos_ppm": objetivos_ppm}
    sel = res["receta"].merge(catalogo, on="nombre", how="left")
    ap = calcular_aportes(sel, V)
    t = totales(ap, V)
    V_madre = V / razon if razon else V
    rec = res["receta"].copy()
    rec["kg"] = rec["gramos"] / 1000.0
    rec["g_L_madre"] = rec["gramos"] / V_madre if V_madre else np.nan
    sel_m = rec.merge(catalogo[["nombre", "solubilidad_g_L"]], on="nombre", how="left")
    sel_m["excede_solub"] = sel_m.apply(
        lambda r: bool(pd.notna(r["solubilidad_g_L"]) and r["g_L_madre"] > r["solubilidad_g_L"]),
        axis=1)
    ce_gotero = ce_agua + t["CE"]
    return {
        "receta": rec, "solub": sel_m, "objetivos_ppm": objetivos_ppm,
        "ppm_gotero": {e: t[e] for e in ELEMENTOS},
        "ce_gotero": ce_gotero, "ce_madre_aprox": t["CE"] * razon,
        "V_riego_L": V, "V_madre_L": V_madre,
        "df_objetivo": res["df_objetivo"], "ce_limitada": res["ce_limitada"],
        "alguno_excede": bool(sel_m["excede_solub"].any())}


def exportar_excel_temporada(meta, serie, calendario, objetivos_ha,
                             recetas_largo, ppm_largo, ce_eventos) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(list(meta.items()), columns=["Parametro", "Valor"]).to_excel(
            w, sheet_name="Cultivo", index=False)
        serie.to_excel(w, sheet_name="Serie_diaria", index=False)
        calendario.to_excel(w, sheet_name="Calendario", index=False)
        objetivos_ha.reset_index().rename(columns={"index": "Etapa"}).to_excel(
            w, sheet_name="Objetivos_kg_ha", index=False)
        if recetas_largo is not None and not recetas_largo.empty:
            recetas_largo.to_excel(w, sheet_name="Recetas_por_evento", index=False)
        if ppm_largo is not None and not ppm_largo.empty:
            ppm_largo.to_excel(w, sheet_name="PPM_gotero", index=False)
        if ce_eventos is not None and not ce_eventos.empty:
            ce_eventos.to_excel(w, sheet_name="CE_eventos", index=False)
    return buf.getvalue()


def _tab_cultivo():
    st.markdown("**1 · Especie y etapas FAO-56**")
    c0, c1 = st.columns([1.4, 1])
    especie = c0.selectbox("Especie (Kc fijos en el código, editables abajo)",
                           list(CROPS.keys()), key="temp_especie")
    crop_base = CROPS[especie]
    sup_ha = c1.number_input("Superficie (ha)", min_value=0.01, value=1.0, step=0.1,
                             key="temp_sup_ha")
    st.caption(crop_base["ref"])
    cc = st.columns(4)
    kc_ini = cc[0].number_input("Kc inicial", min_value=0.0, value=float(crop_base["kc_ini"]),
                                step=0.05, key="temp_kc_ini")
    kc_mid = cc[1].number_input("Kc medio", min_value=0.0, value=float(crop_base["kc_mid"]),
                                step=0.05, key="temp_kc_mid")
    kc_end = cc[2].number_input("Kc final", min_value=0.0, value=float(crop_base["kc_end"]),
                                step=0.05, key="temp_kc_end")
    p_ag = cc[3].number_input("p (agotamiento)", min_value=0.05, max_value=0.9,
                              value=float(crop_base["p"]), step=0.05, key="temp_p")
    cl = st.columns(4)
    li = cl[0].number_input("L inicial (d)", min_value=1, value=int(crop_base["L_ini"]),
                            step=1, key="temp_li")
    ld = cl[1].number_input("L desarrollo (d)", min_value=1, value=int(crop_base["L_dev"]),
                            step=1, key="temp_ld")
    lm = cl[2].number_input("L media (d)", min_value=1, value=int(crop_base["L_mid"]),
                            step=1, key="temp_lm")
    ll = cl[3].number_input("L final (d)", min_value=1, value=int(crop_base["L_late"]),
                            step=1, key="temp_ll")
    fecha_siembra = st.date_input("Fecha de siembra / brotación", value=date(2026, 6, 1),
                                  key="temp_fsiembra")
    crop = {"kc_ini": kc_ini, "kc_mid": kc_mid, "kc_end": kc_end, "p": p_ag}
    L = (li, ld, lm, ll)
    total = li + ld + lm + ll
    st.info(f"Temporada: {total} días "
            f"({fecha_siembra:%d-%m-%Y} → {fecha_siembra + timedelta(days=total-1):%d-%m-%Y}). "
            f"Etapas: Inicial {li} · Desarrollo {ld} · Media {lm} · Final {ll}.")
    st.session_state["temp_crop"] = crop
    st.session_state["temp_L"] = L
    st.session_state["temp_total"] = total
    st.session_state["temp_fecha"] = fecha_siembra
    st.session_state["temp_sup"] = sup_ha
    st.session_state["temp_especie_nombre"] = especie


def _tab_eto():
    st.markdown("**2 · ETo de referencia (formato agrometeorología.cl / Red INIA)**")
    up = st.file_uploader("Sube el xlsx de ETo", type=["xlsx"], key="temp_eto_up")
    if up is None:
        st.info("Sube el archivo de Evapotranspiración para mapear la ETo a la temporada. "
                "Mientras tanto puedes definir una ETo constante de respaldo abajo.")
    eto_df, estaciones = None, []
    if up is not None:
        try:
            eto_df, estaciones = parsear_eto_agromet(up)
            st.success(f"ETo leída: {len(eto_df)} días · estaciones: {', '.join(estaciones)}.")
            est = st.selectbox("Estación a usar", estaciones, key="temp_estacion")
            vis = eto_df[["fecha", est]].dropna().rename(columns={est: "ETo_mm"})
            c1, c2 = st.columns([1, 1.3])
            with c1:
                st.dataframe(vis.assign(fecha=vis["fecha"].dt.strftime("%d-%m-%Y")),
                             use_container_width=True, hide_index=True, height=260)
            with c2:
                st.line_chart(vis.set_index("fecha")["ETo_mm"])
                st.caption(f"ETo media medida: {vis['ETo_mm'].mean():.2f} mm/d · "
                           f"min {vis['ETo_mm'].min():.2f} · máx {vis['ETo_mm'].max():.2f}.")
            st.session_state["temp_eto_df"] = eto_df
            st.session_state["temp_estacion_sel"] = est
        except Exception as e:
            st.error(f"No se pudo leer el archivo: {e}")
            st.session_state.pop("temp_eto_df", None)
    relleno_modo = st.radio(
        "Relleno de días sin ETo medida (la temporada suele exceder los datos)",
        ["Media de los días medidos", "Valor constante"], horizontal=True,
        key="temp_relleno_modo")
    if relleno_modo.startswith("Valor"):
        st.session_state["temp_relleno"] = st.number_input(
            "ETo de relleno (mm/d)", min_value=0.0, value=3.0, step=0.1,
            key="temp_relleno_val")
    else:
        st.session_state["temp_relleno"] = "media"


def _tab_oferta():
    st.markdown("**3 · Oferta hídrica y oferta de riego (factor de irrigación por etapa)**")
    if "temp_crop" not in st.session_state:
        st.warning("Define primero el cultivo en la pestaña 1.")
        return
    st.caption("Factor de irrigación por etapa = fracción de la ETc a reponer "
               "(p. ej. 0,7·ETc). La oferta de riego diaria = factor · Kc · ETo.")
    cols = st.columns(4)
    defaults = {"Inicial": 0.70, "Desarrollo": 0.85, "Media": 1.00, "Final": 0.80}
    factores = {}
    for i, etapa in enumerate(ETAPAS_FAO):
        factores[etapa] = cols[i].number_input(
            f"Factor {etapa}", min_value=0.0, max_value=2.0,
            value=float(defaults[etapa]), step=0.05, key=f"temp_fac_{etapa}")
    serie, info = construir_serie_diaria(
        st.session_state["temp_crop"], st.session_state["temp_fecha"],
        st.session_state["temp_L"], factores,
        st.session_state.get("temp_eto_df"),
        st.session_state.get("temp_estacion_sel", ""),
        relleno=st.session_state.get("temp_relleno", "media"))
    st.session_state["temp_serie"] = serie
    st.session_state["temp_factores"] = factores

    if info["cobertura_pct"] < 100:
        st.warning(f"Cobertura de ETo medida: {info['cobertura_pct']:.0f}% "
                   f"({info['n_medidos']}/{info['total_dias']} días). "
                   f"Se rellenaron {info['n_rellenos']} días con "
                   f"{info['val_relleno']:.2f} mm/d. Para planificación robusta usa una "
                   "serie de ETo normal o histórica de la temporada completa.")
    else:
        st.success("ETo medida cubre toda la temporada.")

    etc_tot = serie["ETc_mm"].sum()
    oferta_tot = serie["oferta_riego_mm"].sum()
    m = st.columns(4)
    m[0].metric("ETc temporada", f"{etc_tot:.0f} mm")
    m[1].metric("Oferta de riego (neta)", f"{oferta_tot:.0f} mm")
    m[2].metric("ETc media", f"{serie['ETc_mm'].mean():.2f} mm/d")
    m[3].metric("Sup. (ha)", f"{st.session_state['temp_sup']:.2f}")

    st.markdown("**Curva diaria** (ETo, ETc y oferta de riego):")
    st.line_chart(serie.set_index("fecha")[["ETo_mm", "ETc_mm", "oferta_riego_mm"]])
    st.markdown("**Acumulado de la temporada:**")
    acum = serie[["fecha"]].copy()
    acum["ETc_acum_mm"] = serie["ETc_mm"].cumsum()
    acum["Oferta_riego_acum_mm"] = serie["oferta_riego_mm"].cumsum()
    st.line_chart(acum.set_index("fecha"))
    resumen_etapa = (serie.groupby("etapa")
                     .agg(dias=("dia", "count"), ETo_mm=("ETo_mm", "sum"),
                          ETc_mm=("ETc_mm", "sum"), oferta_riego_mm=("oferta_riego_mm", "sum"))
                     .reindex(ETAPAS_FAO).dropna(how="all"))
    st.markdown("**Resumen por etapa (mm):**")
    st.dataframe(resumen_etapa.round(1), use_container_width=True)
    with st.expander("Ver serie diaria completa"):
        st.dataframe(serie.assign(fecha=serie["fecha"].apply(lambda d: d.strftime("%d-%m-%Y"))),
                     use_container_width=True, hide_index=True, height=320)


def _tab_calendario():
    st.markdown("**4 · Calendario de riego (mm y tiempo) — propuesto y editable**")
    if "temp_serie" not in st.session_state:
        st.warning("Genera primero la oferta hídrica en la pestaña 3.")
        return
    st.markdown("Parámetros del equipo y manejo:")
    c = st.columns(4)
    Ea = c[0].number_input("Eficiencia de aplicación", min_value=0.05, max_value=1.0,
                           value=0.90, step=0.01, key="temp_Ea",
                           help="Fracción del agua aplicada que queda disponible (riego "
                                "localizado típico 0,85–0,95).")
    Es = c[1].number_input("Eficiencia de almacenamiento", min_value=0.05, max_value=1.0,
                           value=0.95, step=0.01, key="temp_Es",
                           help="Fracción del agua que efectivamente se almacena en la "
                                "zona radicular.")
    PPeq = c[2].number_input("Precipitación del equipo PPeq (mm/h)", min_value=0.1,
                             value=3.5, step=0.1, key="temp_PPeq",
                             help="Lámina horaria que aplica el equipo (caudal/superficie).")
    Kr = c[3].number_input("Kr (reducción localizada)", min_value=0.1, max_value=1.0,
                           value=1.00, step=0.05, key="temp_Kr",
                           help="Coeficiente de reducción por sombreamiento/mojado parcial "
                                "en riego localizado.")
    c2 = st.columns(4)
    Pe = c2[0].number_input("Precipitación efectiva (mm/d)", min_value=0.0, value=0.0,
                            step=0.5, key="temp_Pe",
                            help="Aporte de lluvia que descuenta la lámina neta diaria.")
    cfg_riego = {"Ea": Ea, "Es": Es, "PPeq": PPeq, "Kr": Kr, "Pe_mm_dia": Pe}

    st.markdown("**Días operativos** (desmarca, p. ej., el domingo):")
    cd = st.columns(7)
    dias_operativos = set()
    default_op = [True, True, True, True, True, True, False]
    for i, dia in enumerate(DIAS_SEMANA):
        if cd[i].checkbox(dia, value=default_op[i], key=f"temp_op_{i}"):
            dias_operativos.add(i)
    if not dias_operativos:
        st.error("Selecciona al menos un día operativo.")
        return

    st.markdown("**Frecuencia de riego estimada por etapa (días entre riegos):**")
    cf = st.columns(4)
    freq_def = {"Inicial": 4, "Desarrollo": 3, "Media": 2, "Final": 3}
    freq_etapa = {}
    for i, etapa in enumerate(ETAPAS_FAO):
        freq_etapa[etapa] = cf[i].number_input(
            f"Frecuencia {etapa}", min_value=1, max_value=30,
            value=int(freq_def[etapa]), step=1, key=f"temp_freq_{etapa}")

    st.session_state["temp_cfg_riego"] = cfg_riego
    if st.button("Generar / regenerar calendario propuesto", type="primary",
                 key="temp_btn_cal"):
        prop = generar_calendario(st.session_state["temp_serie"], freq_etapa,
                                  dias_operativos, cfg_riego)
        st.session_state["temp_cal_prop"] = prop
        st.session_state.pop("temp_cal_edit", None)

    if "temp_cal_prop" not in st.session_state:
        st.info("Pulsa el botón para generar la propuesta de calendario.")
        return

    prop = st.session_state["temp_cal_prop"]
    if prop.empty:
        st.warning("No se generaron eventos: revisa frecuencias, días operativos o la "
                   "oferta de riego (puede ser 0 si el factor o Kr son muy bajos).")
        return

    base = st.session_state.get("temp_cal_edit", prop).copy()
    base["fecha"] = pd.to_datetime(base["fecha"])
    st.caption("Edita fechas, láminas o tiempos; desmarca 'incluir' para saltar un riego. "
               "Las recetas usarán esta tabla editada.")
    edit = st.data_editor(
        base, use_container_width=True, hide_index=True, num_rows="dynamic",
        key="temp_cal_editor",
        column_config={
            "evento": st.column_config.NumberColumn("Evento", disabled=True),
            "fecha": st.column_config.DateColumn("Fecha", format="DD-MM-YYYY"),
            "dia_semana": st.column_config.TextColumn("Día", disabled=True),
            "etapa": st.column_config.SelectboxColumn("Etapa", options=ETAPAS_FAO),
            "intervalo_dias": st.column_config.NumberColumn("Interv. (d)"),
            "ETc_acum_mm": st.column_config.NumberColumn("ETc acum (mm)", format="%.1f"),
            "lamina_neta_mm": st.column_config.NumberColumn("Neta (mm)", format="%.2f"),
            "lamina_bruta_mm": st.column_config.NumberColumn("Bruta (mm)", format="%.2f"),
            "tiempo_riego_h": st.column_config.NumberColumn("Tiempo (h)", format="%.2f"),
            "incluir": st.column_config.CheckboxColumn("Incluir")})
    edit["fecha"] = pd.to_datetime(edit["fecha"]).dt.date
    st.session_state["temp_cal_edit"] = edit

    activos = edit[edit["incluir"] == True]
    sup = st.session_state["temp_sup"]
    vol_tot = (activos["lamina_bruta_mm"].sum() * sup * 10.0)
    mm = st.columns(4)
    mm[0].metric("N° de riegos", f"{len(activos)}")
    mm[1].metric("Lámina bruta total", f"{activos['lamina_bruta_mm'].sum():.0f} mm")
    mm[2].metric("Tiempo total", f"{activos['tiempo_riego_h'].sum():.1f} h")
    mm[3].metric("Volumen temporada", f"{vol_tot:.0f} m³")

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        edit.to_excel(w, sheet_name="Calendario", index=False)
    st.download_button("Descargar calendario en Excel", data=buf.getvalue(),
                       file_name="calendario_riego.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="temp_dl_cal")


def _tab_nutricion():
    st.markdown("**5 · Objetivos nutricionales (kg de nutriente por hectárea y etapa)**")
    st.caption("Define el requerimiento de cada nutriente por etapa, en kg/ha. El total "
               "de la temporada es la suma de las etapas (× superficie en el reparto).")
    especie = st.session_state.get("temp_especie_nombre", "")
    base = pd.DataFrame(OBJETIVO_NUTRI_DEFAULT).T.reindex(ETAPAS_FAO)[MACROS_OBJETIVO]
    base.index.name = "Etapa"
    if "temp_obj_nutri" in st.session_state:
        guardado = st.session_state["temp_obj_nutri"]
        base = guardado.reindex(ETAPAS_FAO)[MACROS_OBJETIVO]
    edit = st.data_editor(
        base, use_container_width=True, key="temp_nutri_editor",
        column_config={n: st.column_config.NumberColumn(f"{n} (kg/ha)", min_value=0.0,
                       step=1.0, format="%.1f") for n in MACROS_OBJETIVO})
    st.session_state["temp_obj_nutri"] = edit
    tot = edit.sum()
    st.markdown("**Totales de temporada (kg/ha):**")
    res = pd.DataFrame({"Nutriente": MACROS_OBJETIVO,
                        "kg/ha": [tot[n] for n in MACROS_OBJETIVO]})
    res["kg/ha (óxido)"] = [
        tot["N"], tot["P"] / CFG["ox"]["P2O5_a_P"], tot["K"] / CFG["ox"]["K2O_a_K"],
        tot["Ca"] / CFG["ox"]["CaO_a_Ca"], tot["Mg"] / CFG["ox"]["MgO_a_Mg"], tot["S"]]
    res["forma óxido"] = ["N", "P2O5", "K2O", "CaO", "MgO", "S"]
    st.dataframe(res.round(1), use_container_width=True, hide_index=True)
    sup = st.session_state.get("temp_sup", 1.0)
    st.caption(f"Para {sup:.2f} ha, el total de N sería "
               f"{tot['N']*sup:.1f} kg, K {tot['K']*sup:.1f} kg, etc.")
    st.bar_chart(edit)


def _tab_recetas(catalogo, razon, ce_agua, hco3_agua):
    st.markdown("**6 · Recetas de fertirriego por evento de riego**")
    if "temp_cal_edit" not in st.session_state:
        st.warning("Genera el calendario en la pestaña 4 antes de calcular recetas.")
        return
    if "temp_obj_nutri" not in st.session_state:
        st.warning("Define los objetivos nutricionales en la pestaña 5.")
        return

    opciones = sorted(catalogo["nombre"].tolist())
    default = [n for n in ferts_macro_por_defecto(catalogo) if n in opciones]
    disponibles = st.multiselect("Fertilizantes disponibles para el solver",
                                 options=opciones, default=default, key="temp_disp")
    c = st.columns(3)
    modo_rep = c[0].radio("Reparto del nutriente dentro de la etapa",
                          ["Proporcional al agua", "Partes iguales"], key="temp_modo_rep")
    ce_crit = c[1].number_input("CE crítica en gotero (dS/m)", min_value=0.0, value=0.0,
                                step=0.1, key="temp_ce_crit",
                                help="Tope de CE de la solución que recibe la planta. "
                                     "0 = sin límite. Incluye la CE del agua.")
    c[2].metric("CE agua / razón", f"{ce_agua:.2f} / 1:{razon}")
    if not disponibles:
        st.info("Selecciona al menos un fertilizante.")
        return

    if not st.button("Generar recetas de la temporada", type="primary",
                     key="temp_btn_rec"):
        st.stop()

    cal = st.session_state["temp_cal_edit"].copy()
    cal = cal[cal["incluir"] == True].reset_index(drop=True)
    if cal.empty:
        st.warning("No hay eventos activos en el calendario.")
        return
    sup = st.session_state["temp_sup"]
    obj_ha = st.session_state["temp_obj_nutri"].reindex(ETAPAS_FAO)[MACROS_OBJETIVO]
    modo = "agua" if modo_rep.startswith("Proporcional") else "igual"
    cal = distribuir_nutrientes(cal, obj_ha, sup, modo=modo)

    recetas_largo, ppm_largo, ce_rows = [], [], []
    resumen_rows = []
    detalles = []
    barra = st.progress(0.0, text="Resolviendo recetas...")
    n = len(cal)
    for i, (_, ev) in enumerate(cal.iterrows()):
        kg_event = {nut: float(ev[f"kg_{nut}"]) for nut in MACROS_OBJETIVO}
        out = receta_evento(catalogo, kg_event, ev["lamina_bruta_mm"], sup,
                            disponibles, razon, ce_agua, ce_crit)
        barra.progress((i + 1) / n, text=f"Evento {int(ev['evento'])}/{n}")
        if out is None or "error" in out:
            resumen_rows.append({
                "evento": int(ev["evento"]),
                "fecha": ev["fecha"].strftime("%d-%m-%Y"),
                "etapa": ev["etapa"], "estado": out.get("error", "sin volumen") if out else "sin volumen"})
            detalles.append((ev, out, kg_event))
            continue
        for _, r in out["receta"].iterrows():
            recetas_largo.append({
                "evento": int(ev["evento"]), "fecha": ev["fecha"].strftime("%d-%m-%Y"),
                "etapa": ev["etapa"], "fertilizante": r["nombre"],
                "gramos_total": round(r["gramos"], 1), "kg_total": round(r["kg"], 3),
                "g_L_madre": round(r["g_L_madre"], 3)})
        fila_ppm = {"evento": int(ev["evento"]), "fecha": ev["fecha"].strftime("%d-%m-%Y"),
                    "etapa": ev["etapa"]}
        for e in ELEMENTOS:
            fila_ppm[f"{e}_ppm_gotero"] = round(out["ppm_gotero"][e], 1)
        ppm_largo.append(fila_ppm)
        ce_rows.append({"evento": int(ev["evento"]),
                        "fecha": ev["fecha"].strftime("%d-%m-%Y"), "etapa": ev["etapa"],
                        "CE_gotero_dS_m": round(out["ce_gotero"], 2),
                        "CE_madre_aprox_dS_m": round(out["ce_madre_aprox"], 2),
                        "V_riego_L": round(out["V_riego_L"], 0),
                        "V_madre_L": round(out["V_madre_L"], 1),
                        "excede_solub_madre": out["alguno_excede"],
                        "CE_limitada": out["ce_limitada"]})
        resumen_rows.append({
            "evento": int(ev["evento"]), "fecha": ev["fecha"].strftime("%d-%m-%Y"),
            "etapa": ev["etapa"], "estado": "ok",
            "CE_gotero": round(out["ce_gotero"], 2),
            "masa_total_kg": round(out["receta"]["kg"].sum(), 2)})
        detalles.append((ev, out, kg_event))
    barra.empty()

    recetas_largo = pd.DataFrame(recetas_largo)
    ppm_largo = pd.DataFrame(ppm_largo)
    ce_eventos = pd.DataFrame(ce_rows)
    st.session_state["temp_recetas_largo"] = recetas_largo

    st.subheader("Resumen por evento", anchor=False)
    st.dataframe(pd.DataFrame(resumen_rows), use_container_width=True, hide_index=True)

    if not recetas_largo.empty:
        st.subheader("Masa total de fertilizante por temporada", anchor=False)
        tot_fert = (recetas_largo.groupby("fertilizante")[["kg_total"]].sum()
                    .sort_values("kg_total", ascending=False).round(2))
        st.dataframe(tot_fert, use_container_width=True)
        st.subheader("ppm a gotero por evento", anchor=False)
        st.dataframe(ppm_largo, use_container_width=True, hide_index=True, height=300)
        if not ce_eventos.empty:
            st.line_chart(ce_eventos.set_index("fecha")["CE_gotero_dS_m"])
            if ce_eventos["excede_solub_madre"].any():
                st.warning("En algunos eventos la concentración en el estanque madre "
                           "supera la solubilidad: sube el volumen del estanque madre, baja "
                           "la razón de inyección o fracciona la inyección.")

    st.subheader("Detalle por evento", anchor=False)
    for ev, out, kg_event in detalles:
        etiqueta = f"Evento {int(ev['evento'])} · {ev['fecha']:%d-%m-%Y} · {ev['etapa']} · " \
                   f"{ev['lamina_bruta_mm']:.1f} mm · {ev['tiempo_riego_h']:.1f} h"
        with st.expander(etiqueta):
            if out is None or "error" in out:
                st.error("No se resolvió: " + (out.get("error", "sin volumen") if out else "sin volumen"))
                continue
            cc = st.columns(4)
            cc[0].metric("CE gotero", f"{out['ce_gotero']:.2f} dS/m")
            cc[1].metric("CE madre ~", f"{out['ce_madre_aprox']:.2f} dS/m")
            cc[2].metric("V riego", f"{out['V_riego_L']:.0f} L")
            cc[3].metric("V madre/evento", f"{out['V_madre_L']:.0f} L")
            vista = out["receta"][["nombre", "gramos", "kg", "g_L_madre"]].rename(
                columns={"nombre": "Fertilizante", "gramos": "g (total)",
                         "kg": "kg (total)", "g_L_madre": "g/L (madre)"})
            st.dataframe(vista.round({"g (total)": 1, "kg (total)": 3, "g/L (madre)": 3}),
                         use_container_width=True, hide_index=True)
            ppm_df = pd.DataFrame({
                "Nutriente": ELEMENTOS,
                "Objetivo gotero (ppm)": [round(out["objetivos_ppm"].get(e, 0), 1)
                                          if e in MACROS_OBJETIVO else 0 for e in ELEMENTOS],
                "Logrado gotero (ppm)": [round(out["ppm_gotero"][e], 1) for e in ELEMENTOS],
                "kg objetivo evento": [round(kg_event.get(e, 0), 3)
                                       if e in MACROS_OBJETIVO else 0 for e in ELEMENTOS]})
            st.dataframe(ppm_df, use_container_width=True, hide_index=True)
            if out["alguno_excede"]:
                st.warning("Concentración en madre supera la solubilidad de algún producto.")

    meta = {
        "Especie": st.session_state.get("temp_especie_nombre", ""),
        "Superficie (ha)": sup,
        "Fecha siembra": st.session_state["temp_fecha"].strftime("%d-%m-%Y"),
        "Razón de inyección": f"1:{razon}", "CE agua (dS/m)": ce_agua,
        "HCO3 agua (mg/L)": hco3_agua, "CE crítica gotero (dS/m)": ce_crit,
        "Reparto": modo_rep, "N eventos": len(cal)}
    xls = exportar_excel_temporada(
        meta, st.session_state["temp_serie"], st.session_state["temp_cal_edit"],
        obj_ha, recetas_largo, ppm_largo, ce_eventos)
    st.download_button("Descargar temporada completa en Excel", data=xls,
                       file_name="fertirriego_temporada.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="temp_dl_full")


def modo_temporada(catalogo, volumen_L, razon, ce_agua, hco3_agua):
    st.caption("Flujo de temporada: cultivo y Kc FAO-56 → ETo → oferta hídrica → "
               "calendario de riego editable → objetivos nutricionales → recetas por evento.")
    tabs = st.tabs(["1 · Cultivo (FAO-56)", "2 · ETo", "3 · Oferta hídrica",
                    "4 · Calendario", "5 · Objetivos nutri.", "6 · Recetas"])
    with tabs[0]:
        _tab_cultivo()
    with tabs[1]:
        _tab_eto()
    with tabs[2]:
        _tab_oferta()
    with tabs[3]:
        _tab_calendario()
    with tabs[4]:
        _tab_nutricion()
    with tabs[5]:
        _tab_recetas(catalogo, razon, ce_agua, hco3_agua)


_OCULTAR_UI = """
<style>
[data-testid="stToolbar"] {visibility: hidden; height: 0; position: fixed;}
[data-testid="stDecoration"] {display: none;}
[data-testid="stStatusWidget"] {display: none;}
[data-testid="stHeaderActionElements"] {display: none !important;}
.stDeployButton {display: none;}
#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}
a.stHeadingAnchor {display: none !important;}
h1 a, h2 a, h3 a {display: none !important;}
</style>
"""


def main():
    st.set_page_config(page_title=CFG["titulo"], page_icon="🧪", layout="wide",
                       menu_items={"Get help": None, "Report a bug": None, "About": None})
    st.markdown(_OCULTAR_UI, unsafe_allow_html=True)
    st.title(CFG["titulo"], anchor=False)
    st.caption("Calcula concentraciones, CE, compatibilidad y pH. "
               "Modo directo (dosis → resultado), inverso (objetivo → receta) "
               "o temporada (FAO-56 → calendario de riego → recetas por evento).")
    volumen_L, razon, ce_agua, hco3_agua, catalogo = sidebar_config()
    modo = st.radio("Modo de trabajo",
                    ["Directo  (dosis → resultado)", "Inverso  (objetivo → receta)",
                     "Temporada  (FAO-56 → calendario → recetas)"],
                    horizontal=True)
    st.divider()
    if modo.startswith("Directo"):
        modo_directo(catalogo, volumen_L, razon, ce_agua, hco3_agua)
    elif modo.startswith("Inverso"):
        modo_inverso(catalogo, volumen_L, razon, ce_agua, hco3_agua)
    else:
        modo_temporada(catalogo, volumen_L, razon, ce_agua, hco3_agua)
    st.divider()
    st.caption("La CE usa factores empíricos por sal (g/L→dS/m); la urea no aporta CE. "
               "El pH es indicativo. El solver minimiza el error relativo sin masas "
               "negativas, pero no impone compatibilidad: revisa la separación A/B.")


if __name__ == "__main__":
    main()
