# -*- coding: utf-8 -*-
"""
=============================================================================
 PLATAFORMA DE DISOLUCIÓN DE FERTILIZANTES PARA FERTIRRIEGO
=============================================================================
 Dos modos de trabajo:
   1) DIRECTO : dosis (kg/g) -> concentraciones, CE, compatibilidad, pH...
   2) INVERSO : objetivo (ppm de cada nutriente) -> receta (gramos de cada
                fertilizante), resuelta con minimos cuadrados no negativos
                (NNLS) ponderados por error relativo.

 Catalogo: se lee desde una Google Sheet PUBLICA fija (constante SHEET_URL).
           Si no esta configurada o falla, usa un catalogo interno de respaldo.
 Ejecutar:  streamlit run fertirriego_app.py
 Requiere:  pip install streamlit pandas openpyxl scipy
=============================================================================
"""

import re
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st

# ===========================================================================
#  >>>  URL FIJA DE LA GOOGLE SHEET DEL CATALOGO  <<<
#  Pega aqui el enlace de tu hoja PUBLICA (acceso por enlace, rol Lector).
#  Mientras este vacia, la app usa el catalogo interno por defecto.
#  Ejemplo: "https://docs.google.com/spreadsheets/d/XXXXXXXX/edit#gid=0"
# ===========================================================================
SHEET_URL = "https://docs.google.com/spreadsheets/d/11LHtPNCzdEKSzqMTrXhKLyr-9aI-CoSO/edit?usp=sharing&ouid=104728458573663018595&rtpof=true&sd=true"

CFG = {
    "titulo": "Plataforma de Disolución de Fertilizantes",
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
         "Muy soluble. Acidificante por nitrificacion del NH4."),
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
         "Aporta fosfato -> Estanque B. pH de solucion ~4.5."),
        ("Fosfato monopotasico (MKP)", "KH2PO4", "solido",
         0, 22.7, 28.2, 0, 0, 0, 0, 0, 0, 0, 230, 0.70, "acida",
         False, True, False, False, False, 0, None,
         "Aporta fosfato -> Estanque B. pH de solucion ~4.5."),
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
        ("Acido fosforico 85%", "H3PO4", "liquido",
         0, 26.8, 0, 0, 0, 0, 0, 0, 0, 0, None, 1.00, "acida",
         False, True, False, True, False, 8.67, 1.685,
         "Liquido. Anadir SIEMPRE el acido al agua. ~1 proton efectivo."),
        ("Acido nitrico 65%", "HNO3", "liquido",
         14.4, 0, 0, 0, 0, 0, 0, 1.0, 0, 0, None, 1.20, "acida",
         False, False, False, True, False, 10.30, 1.39,
         "Liquido corrosivo. Anadir SIEMPRE el acido al agua. EPP obligatorio."),
        ("Quelato de hierro EDTA (Fe 6%)", "Fe-EDTA", "solido",
         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 900, 0.20, "neutra",
         False, False, False, False, True, 0, None,
         "Fe 6%. Estable hasta pH ~6.5; sobre eso precipita el Fe."),
        ("Quelato de hierro EDDHA (Fe 6%)", "Fe-EDDHA", "solido",
         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 900, 0.15, "neutra",
         False, False, False, False, True, 0, None,
         "Fe 6%. Estable hasta pH ~9; ideal para aguas/suelos calcareos."),
        ("Sulfato de manganeso", "MnSO4*H2O", "solido",
         0, 0, 0, 0, 0, 19.0, 0, 0, 0, 0, 700, 0.10, "neutra",
         True, False, False, False, False, 0, None,
         "Mn 32%. Microelemento (dosis bajas). Aporta SO4."),
        ("Sulfato de zinc (heptahidr.)", "ZnSO4*7H2O", "solido",
         0, 0, 0, 0, 0, 11.0, 0, 0, 0, 0, 580, 0.10, "neutra",
         True, False, False, False, False, 0, None,
         "Zn 22%. Microelemento (dosis bajas). Aporta SO4."),
        ("Acido borico", "H3BO3", "solido",
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
    for c in num:
        df[c] = pd.to_numeric(df[c], errors="coerce")
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
        alertas.append(("MEDIA", "Riesgo de precipitacion de fosfato de magnesio",
                        f"{', '.join(mg)} + {', '.join(fosf)}. Mayor riesgo a pH alto."))
    if quel and fosf:
        alertas.append(("MEDIA", "Posible interaccion quelato de Fe + fosfato",
                        f"{', '.join(quel)} con {', '.join(fosf)}: vigilar pH."))
    if acid:
        alertas.append(("INFO", "Manejo de acidos",
                        f"{', '.join(acid)}: anadir SIEMPRE el acido sobre el agua, con EPP."))
    if cl:
        alertas.append(("INFO", "Aporte de cloruro",
                        f"{', '.join(cl)}: revisar tolerancia del cultivo al Cl-."))
    if not alertas:
        alertas.append(("OK", "Sin incompatibilidades evidentes",
                        "La combinacion no dispara reglas de precipitacion."))
    return alertas


def asignar_estanques(sel: pd.DataFrame) -> pd.DataFrame:
    def grupo(r):
        a = (r["Ca_pct"] > 0) or r["es_quelato_Fe"]
        b = r["sulfato"] or r["fosfato"] or r["es_acido"]
        if a and not b:
            return "A"
        if b and not a:
            return "B"
        if a and b:
            return "CONFLICTO"
        return "A o B"
    out = sel[["nombre"]].copy()
    out["estanque"] = sel.apply(grupo, axis=1)
    return out


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
        return "Tendencia neutra"
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
        estado = "Sobre-neutralizado: exceso de acido (riesgo de pH muy bajo)."
    else:
        estado = "Neutralizacion ~completa de los bicarbonatos."
    return {"hco3_meq": hco3_meq, "H_meq": meq_H_total,
            "residual_meq": residual_meq, "residual_mg_L": residual_mg_L,
            "estado": estado}


def resolver_receta_objetivo(catalogo, objetivos, volumen_L, disponibles,
                             ce_cap_madre=None):
    """Resuelve la receta (gramos por fertilizante) para acercarse a los objetivos
    de ppm, sin masas negativas y, opcionalmente, sin superar una CE critica.

    ce_cap_madre: CE maxima permitida en el estanque madre (dS/m). None o <=0 = sin
    tope. Si NNLS ya cumple el tope, se devuelve esa solucion (es la optima). Si el
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
    b = np.array([float(objetivos[el]) for el in nut])
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
        cols = ["nombre", "gramos", "g_L"] + [f"{e}_ppm" for e in ELEMENTOS] + \
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
    ce_final = ce_agua + ce_madre / razon

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("CE estanque madre", f"{ce_madre:.2f} dS/m")
    c2.metric("CE final (gotero)", f"{ce_final:.2f} dS/m",
              help=f"CE agua ({ce_agua:.2f}) + CE madre / {razon}")
    c3.metric("Concentracion total", f"{t['pct_pv']:.2f} % p/v")
    c4.metric("Tendencia de pH", ph_tend)

    sev_max = max((a[0] for a in alertas),
                  key=lambda s: {"OK": 0, "INFO": 1, "MEDIA": 2, "ALTA": 3}.get(s, 0))
    if sev_max == "ALTA":
        st.error("Hay incompatibilidades de severidad ALTA (ver pestana Compatibilidad).")
    elif sev_max == "MEDIA":
        st.warning("Hay advertencias de compatibilidad de severidad MEDIA.")

    tabs = st.tabs(["Nutrientes", "Compatibilidad y estanques",
                    "Solubilidad", "pH / bicarbonatos", "Balance ionico"])
    with tabs[0]:
        st.markdown("**Concentracion de nutrientes (ppm = mg/L):**")
        nut = pd.DataFrame({
            "Nutriente": ELEMENTOS,
            "Estanque madre (ppm)": [t[e] for e in ELEMENTOS],
            f"Final 1:{razon} (ppm)": [t[e] * factor for e in ELEMENTOS]})
        st.dataframe(nut.round(2), use_container_width=True, hide_index=True)
        st.bar_chart(nut.set_index("Nutriente")["Estanque madre (ppm)"])
        with st.expander("Equivalente en oxidos (P2O5, K2O, CaO, MgO)"):
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
        st.markdown("**Alertas de compatibilidad / precipitacion:**")
        colores = {"ALTA": st.error, "MEDIA": st.warning, "INFO": st.info, "OK": st.success}
        for sev, titulo, detalle in alertas:
            colores.get(sev, st.info)(f"**[{sev}] {titulo}** — {detalle}")
        st.divider()
        st.markdown("**Separacion sugerida en estanques** "
                    "(A = calcio/quelatos/nitratos · B = sulfatos/fosfatos/acidos):")
        st.dataframe(estanques, use_container_width=True, hide_index=True)
        if (estanques["estanque"] == "CONFLICTO").any():
            st.warning("Algun fertilizante cae en A y B a la vez: revisa la receta.")
    with tabs[2]:
        st.markdown("**Solubilidad** (dosis g/L vs limite a ~20 C):")
        sol = df_ap[["nombre", "g_L", "solubilidad_g_L", "excede_solub"]].rename(
            columns={"g_L": "dosis (g/L)", "solubilidad_g_L": "solubilidad (g/L)",
                     "excede_solub": "excede?"})
        st.dataframe(sol.round(1), use_container_width=True, hide_index=True)
        if df_ap["excede_solub"].any():
            st.error("Hay dosis que superan la solubilidad: no se disolveran del todo. "
                     "Reduce dosis, sube volumen o usa agua mas tibia.")
        else:
            st.success("Todas las dosis estan por debajo del limite de solubilidad.")
    with tabs[3]:
        st.markdown(f"**Tendencia de pH de la mezcla:** {ph_tend}")
        st.caption("Estimacion cualitativa ponderada por masa; el pH real debe medirse.")
        st.divider()
        st.markdown("**Neutralizacion de bicarbonatos** (acidos vs HCO3- del agua):")
        meq_H = float((df_ap["gramos"] * df_ap["meq_H_g"]).sum())
        res = neutralizacion_bicarbonatos(hco3_agua, volumen_L, meq_H)
        cc1, cc2, cc3 = st.columns(3)
        cc1.metric("meq HCO3- en estanque", f"{res['hco3_meq']:.0f}")
        cc2.metric("meq H+ por acidos", f"{res['H_meq']:.0f}")
        cc3.metric("HCO3- residual", f"{res['residual_mg_L']:.0f} mg/L")
        st.info(res["estado"])
    with tabs[4]:
        st.markdown("**Balance ionico (meq/L)** — cationes vs aniones:")
        st.dataframe(bal.round(2), use_container_width=True, hide_index=True)
        cat = bal[bal["grupo"] == "cation"]["meq_L"].sum()
        ani = bal[bal["grupo"] == "anion"]["meq_L"].sum()
        d1, d2, d3 = st.columns(3)
        d1.metric("Suma cationes", f"{cat:.2f} meq/L")
        d2.metric("Suma aniones", f"{ani:.2f} meq/L")
        d3.metric("Desbalance", f"{((cat-ani)/cat*100 if cat else 0):+.1f} %")
        st.bar_chart(bal.set_index("ion")["meq_L"])
        st.caption("Un desbalance grande suele indicar un ion acompanante no contemplado.")

    st.subheader("Exportar receta")
    st.download_button(
        "Descargar receta en Excel",
        data=exportar_excel(df_ap, t, alertas, estanques, bal, df_objetivo),
        file_name="receta_fertirriego.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def sidebar_config():
    with st.sidebar:
        st.header("Configuracion")
        volumen_L = st.number_input("Capacidad del estanque (L)", min_value=1.0,
                                    value=CFG["volumen_default_L"], step=50.0)
        razon = st.number_input("Razon de inyeccion 1:R", min_value=1,
                                value=CFG["razon_iny_default"], step=1,
                                help="Usa 1 si el estanque es la solucion final aplicada.")
        st.divider()
        st.subheader("Agua de riego (opcional)")
        ce_agua = st.number_input("CE del agua (dS/m)", min_value=0.0, value=0.0, step=0.1)
        hco3_agua = st.number_input("Bicarbonatos HCO3- (mg/L)", min_value=0.0,
                                    value=0.0, step=10.0)
        st.divider()
        st.subheader("Catalogo de fertilizantes")
        catalogo = cargar_catalogo_fija()
    return volumen_L, razon, ce_agua, hco3_agua, catalogo


def cargar_catalogo_fija() -> pd.DataFrame:
    """Carga el catalogo desde la Google Sheet fija (SHEET_URL).
    Si la URL esta vacia o falla, usa el catalogo interno como respaldo.
    Muestra el estado en la barra lateral."""
    if not SHEET_URL.strip():
        st.info("Usando catalogo interno (la URL de la hoja aun no esta configurada).")
        return cargar_catalogo_default()
    try:
        df_g, faltan = validar_catalogo(cargar_desde_gsheet(SHEET_URL))
        if faltan:
            st.error("La hoja no tiene las columnas requeridas: " + ", ".join(faltan) +
                     ". Se usa el catalogo interno.")
            return cargar_catalogo_default()
        st.success(f"Catalogo cargado desde Google Sheet: {len(df_g)} fertilizantes.")
        return df_g
    except Exception as e:
        st.warning(f"No se pudo leer la hoja fija ({e}). Se usa el catalogo interno.")
        return cargar_catalogo_default()


def modo_directo(catalogo, volumen_L, razon, ce_agua, hco3_agua):
    st.subheader("Seleccion de fertilizantes y dosis")
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
    st.subheader("Resultados")
    analizar_y_mostrar(sel[["nombre", "gramos"]], catalogo, volumen_L, razon,
                       ce_agua, hco3_agua)


def modo_inverso(catalogo, volumen_L, razon, ce_agua, hco3_agua):
    st.subheader("Objetivo de la solucion (ppm de cada nutriente)")
    st.caption("Objetivo del **estanque madre**. Deja 0 para no exigir ese nutriente "
               "(igual se reporta si aparece como efecto colateral).")
    cols = st.columns(6)
    objetivos = {}
    for i, el in enumerate(MACROS_OBJETIVO):
        objetivos[el] = cols[i].number_input(
            f"{el}", min_value=0.0, value=float(CFG["objetivo_default"][el]),
            step=5.0, key=f"obj_{el}")
    st.subheader("Fertilizantes disponibles para la formula")
    opciones = sorted(catalogo["nombre"].tolist())
    default = [n for n in ferts_macro_por_defecto(catalogo) if n in opciones]
    disponibles = st.multiselect(
        "El solver elegira las cantidades solo entre estos:",
        options=opciones, default=default,
        help="Por defecto se sugieren fuentes de macronutrientes (sin acidos ni micros).")
    if not disponibles:
        st.info("Selecciona al menos un fertilizante disponible.")
        return

    st.subheader("CE critica (limitante)")
    cce1, cce2 = st.columns([1, 1.4])
    ce_crit = cce1.number_input(
        "CE critica (dS/m)", min_value=0.0, value=0.0, step=0.1,
        help="Tope de conductividad que la receta NO debe superar. 0 = sin limite.")
    refiere = cce2.radio(
        "La CE critica se refiere a:",
        ["Solucion final (gotero)", "Estanque madre"], horizontal=True,
        help="Final = lo que recibe la planta (CE agua + CE madre / razon). "
             "Madre = la solucion concentrada del estanque.")

    # Traducir el tope a un limite sobre la CE del ESTANQUE MADRE (lo que ve el solver)
    ce_cap_madre = None
    if ce_crit > 0:
        if refiere.startswith("Solucion final"):
            ce_cap_madre = (ce_crit - ce_agua) * razon
            if ce_cap_madre <= 0:
                st.error(f"La CE del agua ({ce_agua:.2f}) ya iguala o supera la CE critica "
                         f"final ({ce_crit:.2f}). Sin margen para fertilizar: sube la CE "
                         "critica o reduce la CE del agua.")
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
    st.subheader("Receta propuesta")
    vista = receta.copy()
    vista["kg"] = vista["gramos"] / 1000.0
    vista["g/L"] = vista["gramos"] / volumen_L
    st.dataframe(
        vista.rename(columns={"nombre": "Fertilizante", "gramos": "g (total)"})
             .round({"g (total)": 1, "kg": 3, "g/L": 3}),
        use_container_width=True, hide_index=True)
    st.subheader("Ajuste a los objetivos")
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
        m1.metric("CE madre lograda", f"{ce_madre:.2f} dS/m")
        m2.metric("CE final (gotero)", f"{ce_final:.2f} dS/m")
        etiqueta = "final" if refiere.startswith("Solucion final") else "madre"
        m3.metric(f"CE critica ({etiqueta})", f"{ce_crit:.2f} dS/m")
        if res["ce_limitada"]:
            st.warning(
                f"La CE critica fue **activa**: limito la receta. Sin tope, la CE del "
                f"estanque madre habria sido {res['ce_sin_tope']:.2f} dS/m. Para respetar "
                "el limite, el solver redujo sales y por eso algunos nutrientes quedan por "
                "debajo del objetivo (ver errores arriba). Sube la CE critica, el volumen "
                "del estanque, o usa fuentes mas concentradas/menos salinas si necesitas "
                "acercarte mas a los objetivos.")
        else:
            st.success("La receta cumple la CE critica con holgura (el tope no fue activo).")

    cap_txt = (f" · tope CE madre={res['ce_tope']:.2f} dS/m"
               f"{' (ACTIVO)' if res['ce_limitada'] else ''}") if res.get("ce_tope") else ""
    metodo = "SLSQP con restriccion de CE" if res["ce_limitada"] else "NNLS"
    st.caption(f"Solver: {metodo}, ponderado por error relativo · residuo={res['rnorm']:.3f} "
               f"· optimizados: {', '.join(res['nut'])}{cap_txt}.")
    st.divider()
    st.subheader("Analisis de la receta resultante")
    analizar_y_mostrar(receta, catalogo, volumen_L, razon, ce_agua, hco3_agua,
                       df_objetivo=df_obj)


def main():
    st.set_page_config(page_title=CFG["titulo"], page_icon="🧪", layout="wide")
    st.title(CFG["titulo"])
    st.caption("Concentracion (ppm, %) · CE (dS/m) · Compatibilidad · pH · "
               "y dosificacion inversa (objetivo -> receta).")
    volumen_L, razon, ce_agua, hco3_agua, catalogo = sidebar_config()
    modo = st.radio("Modo de trabajo",
                    ["Directo  (dosis → resultado)", "Inverso  (objetivo → receta)"],
                    horizontal=True)
    st.divider()
    if modo.startswith("Directo"):
        modo_directo(catalogo, volumen_L, razon, ce_agua, hco3_agua)
    else:
        modo_inverso(catalogo, volumen_L, razon, ce_agua, hco3_agua)
    st.divider()
    st.caption("La CE usa factores empiricos por sal (g/L→dS/m); la urea no aporta CE. "
               "El pH es indicativo. El solver minimiza el error relativo sin masas "
               "negativas, pero no impone compatibilidad: revisa la separacion A/B.")


if __name__ == "__main__":
    main()
