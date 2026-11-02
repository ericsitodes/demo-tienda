"""Dashboard de ventas de ejemplo para "Abarrotes El Puerto".

Uso: streamlit run app.py
"""

import calendar
import re
import unicodedata
from datetime import date, timedelta
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

ARCHIVO_DEMO = Path(__file__).parent / "ventas_tienda.xlsx"
PLANTILLA = Path(__file__).parent / "plantilla-tienda-farmacia.xlsx"
NEGOCIO_DEMO = "Abarrotes El Puerto"
NEGOCIO_VACIO = "Tu negocio"
MIME_EXCEL = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ---------------------------------------------------------------------------
# Paleta: la misma de la página. Los colores de datos están validados contra
# daltonismo sobre el fondo SUPERFICIE.
# ---------------------------------------------------------------------------
FONDO = "#0B1628"
SUPERFICIE = "#12213A"
BORDE = "#27405F"
TEXTO = "#EDF2F9"
TEXTO_SUAVE = "#B0BED2"
ACENTO = "#2DE2C5"
COLORES_SERIES = ["#10A68F", "#4C8EF0", "#BC8520", "#A77BE0", "#D9687F"]
COLOR_OTROS = "#6B7F99"

# El color sigue a la categoría, no a su posición, para que no cambie al filtrar
COLOR_CATEGORIA = dict(zip(['Abarrotes', 'Bebidas', 'Botanas', 'Lácteos', 'Limpieza'], COLORES_SERIES))

# Fechas en español escritas a mano: en la nube no hay locale es_MX instalado
MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
DIAS = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]

COLUMNAS_VENTAS = ["Fecha", "Hora", "Ticket", "Producto", "Categoría", "Cantidad",
                   "Precio unitario", "Total", "Método de pago", "Cajero"]
COLUMNAS_GASTOS = ["Mes", "Concepto", "Monto"]
COLUMNAS_INVENTARIO = ["Producto", "Existencia"]

# Resurtido: días de venta que se promedian y umbrales de alerta
DIAS_CONSUMO = 30
DIAS_URGENTE = 7
DIAS_PRONTO = 15


class ErrorFormato(Exception):
    """El Excel subido no tiene el formato esperado."""


# ---------------------------------------------------------------------------
# Formatos
# ---------------------------------------------------------------------------
def moneda(valor):
    """1234.5 -> $1,234.50"""
    signo = "-" if valor < 0 else ""
    return f"{signo}${abs(valor):,.2f}"


def fecha_es(dia):
    """Convierte una fecha en texto como 5 jul 2026."""
    return f"{dia.day} {MESES[dia.month - 1]} {dia.year}"


def fecha_larga(dia):
    """Convierte una fecha en texto como Dom 5 jul 2026."""
    return f"{DIAS[dia.weekday()]} {fecha_es(dia)}"


def mes_es(periodo):
    """Convierte un periodo mensual en texto como jul 2026."""
    return f"{MESES[periodo.month - 1]} {periodo.year}"


def nombre_archivo(negocio):
    """Convierte "Doña Lupe, Centro" en dona_lupe_centro para nombrar archivos."""
    sin_acentos = unicodedata.normalize("NFKD", negocio).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "_", sin_acentos.lower()).strip("_") or "tu_negocio"


# ---------------------------------------------------------------------------
# Carga y validación del Excel
# ---------------------------------------------------------------------------
def leer_hoja(hojas, nombre, columnas):
    """Toma una hoja y comprueba que traiga todas las columnas."""
    if nombre not in hojas:
        raise ErrorFormato(f'Falta la hoja "{nombre}".')
    tabla = hojas[nombre]
    tabla.columns = [str(c).strip() for c in tabla.columns]
    faltantes = [c for c in columnas if c not in tabla.columns]
    if faltantes:
        raise ErrorFormato(f'A la hoja "{nombre}" le faltan columnas: {", ".join(faltantes)}.')
    return tabla[columnas].copy()


def limpiar_ventas(ventas):
    """Convierte los tipos de la hoja Ventas y descarta filas inservibles."""
    ventas["Fecha"] = pd.to_datetime(ventas["Fecha"], errors="coerce", dayfirst=True).dt.normalize()
    ventas["Total"] = pd.to_numeric(ventas["Total"], errors="coerce")
    ventas["Cantidad"] = pd.to_numeric(ventas["Cantidad"], errors="coerce")
    ventas["HoraNum"] = pd.to_datetime(ventas["Hora"].astype(str), errors="coerce", format="mixed").dt.hour
    ventas = ventas.dropna(subset=["Fecha", "Total", "Cantidad", "HoraNum"])
    if ventas.empty:
        raise ErrorFormato('La hoja "Ventas" no tiene filas válidas.')
    ventas = ventas.assign(HoraNum=ventas["HoraNum"].astype(int))
    for columna in ("Producto", "Categoría", "Método de pago", "Cajero"):
        ventas[columna] = ventas[columna].astype(str).str.strip()
    return ventas


def limpiar_gastos(gastos):
    """Convierte los tipos de la hoja Gastos y agrega el total por mes."""
    gastos["Mes"] = pd.to_datetime(gastos["Mes"], errors="coerce", dayfirst=True).dt.to_period("M")
    gastos["Monto"] = pd.to_numeric(gastos["Monto"], errors="coerce")
    return gastos.dropna(subset=["Mes", "Monto"]).groupby("Mes")["Monto"].sum()


def limpiar_inventario(inventario):
    """Convierte los tipos de la hoja Inventario y suma productos repetidos."""
    inventario["Producto"] = inventario["Producto"].astype(str).str.strip()
    inventario["Existencia"] = pd.to_numeric(inventario["Existencia"], errors="coerce")
    return inventario.dropna(subset=["Existencia"]).groupby("Producto", as_index=False)["Existencia"].sum()


def leer_libro(contenido):
    """Recibe los bytes del Excel y devuelve (ventas, gastos por mes, inventario)."""
    try:
        hojas = pd.read_excel(BytesIO(contenido), sheet_name=None, engine="openpyxl")
    except Exception as error:
        raise ErrorFormato("El archivo no se pudo abrir como Excel (.xlsx).") from error
    ventas = limpiar_ventas(leer_hoja(hojas, "Ventas", COLUMNAS_VENTAS))
    gastos = limpiar_gastos(leer_hoja(hojas, "Gastos", COLUMNAS_GASTOS))
    inventario = limpiar_inventario(leer_hoja(hojas, "Inventario", COLUMNAS_INVENTARIO))
    return ventas, gastos, inventario


@st.cache_data(show_spinner="Cargando datos…")
def cargar_demo():
    """Lee el Excel de ejemplo; solo estos datos ficticios van a la caché."""
    return leer_libro(ARCHIVO_DEMO.read_bytes())


def cargar_subido(subido):
    """Lee el Excel subido y lo guarda solo en la sesión de esta visita.

    No usa st.cache_data: esa caché es del servidor y sobreviviría a la visita.
    """
    if st.session_state.get("libro_id") != subido.file_id:
        with st.spinner("Cargando datos…"):
            libro = leer_libro(subido.getvalue())
        st.session_state["libro"] = libro
        st.session_state["libro_id"] = subido.file_id
    return st.session_state["libro"]


# ---------------------------------------------------------------------------
# Cálculos
# ---------------------------------------------------------------------------
def filtrar_fechas(ventas, inicio, fin):
    """Devuelve las ventas entre dos fechas, ambas incluidas."""
    return ventas[ventas["Fecha"].between(pd.Timestamp(inicio), pd.Timestamp(fin))]


def filtrar_ventas(ventas, inicio, fin, categorias, metodos):
    """Aplica todos los filtros. Una lista vacía significa "todas"."""
    filtradas = filtrar_fechas(ventas, inicio, fin)
    if categorias:
        filtradas = filtradas[filtradas["Categoría"].isin(categorias)]
    if metodos:
        filtradas = filtradas[filtradas["Método de pago"].isin(metodos)]
    return filtradas


def periodo_anterior(inicio, fin):
    """Devuelve el periodo equivalente que termina justo antes de inicio.

    Si el rango son meses completos, compara contra los mismos meses previos
    (jul-sep contra abr-jun); si no, contra los mismos días.
    """
    fin_anterior = inicio - timedelta(days=1)
    if inicio.day == 1 and fin.day == calendar.monthrange(fin.year, fin.month)[1]:
        meses = (fin.year - inicio.year) * 12 + fin.month - inicio.month + 1
        return (pd.Period(inicio, freq="M") - meses).start_time.date(), fin_anterior
    return fin_anterior - (fin - inicio), fin_anterior


def gastos_del_rango(gastos, inicio, fin):
    """Reparte el gasto mensual por días cuando el rango cubre solo parte del mes.

    Devuelve una serie por mes, o None si falta algún mes en la hoja Gastos.
    """
    resultado = {}
    for mes in pd.period_range(inicio, fin, freq="M"):
        if mes not in gastos.index:
            return None
        dias_mes = calendar.monthrange(mes.year, mes.month)[1]
        desde = max(inicio, date(mes.year, mes.month, 1))
        hasta = min(fin, date(mes.year, mes.month, dias_mes))
        resultado[mes] = gastos[mes] * ((hasta - desde).days + 1) / dias_mes
    return pd.Series(resultado)


def calcular_utilidad(ventas, gastos, inicio, fin):
    """Ventas de todo el negocio menos gastos del rango; None si faltan gastos.

    No usa los filtros de categoría ni de pago: los gastos no se pueden
    repartir entre categorías, y restarlos completos daría una utilidad falsa.
    """
    gastos_rango = gastos_del_rango(gastos, inicio, fin)
    if gastos_rango is None:
        return None
    return filtrar_fechas(ventas, inicio, fin)["Total"].sum() - gastos_rango.sum()


def calcular_indicadores(ventas):
    """Calcula ventas, tickets y ticket promedio de un conjunto de filas."""
    total = ventas["Total"].sum()
    tickets = ventas["Ticket"].nunique()
    return {"ventas": total, "tickets": tickets, "ticket_promedio": total / tickets if tickets else 0}


def variacion(actual, anterior):
    """Devuelve el texto de variación porcentual, o None si no se puede calcular."""
    if anterior is None or actual is None or anterior == 0:
        return None
    return f"{(actual - anterior) / abs(anterior) * 100:+.1f}%"


def excel_filtrado(ventas):
    """Devuelve los bytes de un Excel con las ventas filtradas."""
    salida = ventas[COLUMNAS_VENTAS].copy()
    salida["Fecha"] = salida["Fecha"].dt.date
    memoria = BytesIO()
    with pd.ExcelWriter(memoria, engine="openpyxl") as escritor:
        salida.to_excel(escritor, sheet_name="Ventas", index=False)
    return memoria.getvalue()


def tabla_resurtir(ventas, inventario):
    """Calcula para cuántos días alcanza cada producto al ritmo de venta reciente.

    Usa siempre las últimas ventas del archivo, sin filtros: lo que se debe
    comprar no cambia por el periodo que se esté mirando en el tablero.
    """
    ultima = ventas["Fecha"].max()
    recientes = ventas[ventas["Fecha"] > ultima - pd.Timedelta(days=DIAS_CONSUMO)]
    venta_diaria = recientes.groupby("Producto")["Cantidad"].sum() / DIAS_CONSUMO

    tabla = inventario.copy()
    tabla["Venta diaria"] = tabla["Producto"].map(venta_diaria).fillna(0)
    con_venta = tabla["Venta diaria"] > 0
    tabla["Días de existencia"] = (tabla["Existencia"] / tabla["Venta diaria"]).where(con_venta)
    tabla["Estado"] = "Sin ventas recientes"
    tabla.loc[con_venta, "Estado"] = "Suficiente"
    tabla.loc[con_venta & (tabla["Días de existencia"] < DIAS_PRONTO), "Estado"] = "Pronto"
    tabla.loc[con_venta & (tabla["Días de existencia"] < DIAS_URGENTE), "Estado"] = "Urgente"
    return tabla.sort_values("Días de existencia", na_position="last").reset_index(drop=True)


# ---------------------------------------------------------------------------
# Gráficas
# ---------------------------------------------------------------------------
def estilo(fig, alto=340, leyenda=False):
    """Aplica la misma apariencia a todas las gráficas."""
    fig.update_layout(
        height=alto,
        paper_bgcolor=SUPERFICIE,
        plot_bgcolor=SUPERFICIE,
        font=dict(family="Inter, Segoe UI, sans-serif", color=TEXTO_SUAVE, size=13),
        margin=dict(l=12, r=16, t=16, b=12),
        hoverlabel=dict(bgcolor=FONDO, bordercolor=BORDE, font=dict(color=TEXTO, size=13)),
        showlegend=leyenda,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(color=TEXTO)),
        dragmode=False,  # en celular, deslizar debe mover la página y no la gráfica
    )
    fig.update_xaxes(gridcolor=BORDE, zerolinecolor=BORDE, linecolor=BORDE, automargin=True)
    fig.update_yaxes(gridcolor=BORDE, zerolinecolor=BORDE, linecolor=BORDE, automargin=True)
    return fig


def marcas_fecha(dias):
    """Elige marcas del eje en español: por día, semana o mes según el rango."""
    if len(dias) <= 10:
        marcas = list(dias)
    elif len(dias) <= 62:
        marcas = list(dias[::7])
    else:
        marcas = [d for d in dias if d.day == 1] or [dias[0]]
    return marcas, [f"{d.day} {MESES[d.month - 1]}" for d in marcas]


def grafica_ventas_diarias(ventas, inicio, fin):
    """Línea de ventas por día; los días sin venta aparecen en cero."""
    dias = pd.date_range(inicio, fin)
    serie = ventas.groupby("Fecha")["Total"].sum().reindex(dias, fill_value=0)
    fig = go.Figure(go.Scatter(
        x=dias, y=serie.values, mode="lines", line=dict(color=ACENTO, width=2),
        customdata=[[fecha_larga(d), moneda(v)] for d, v in serie.items()],
        hovertemplate="%{customdata[0]}<br><b>%{customdata[1]}</b><extra></extra>",
    ))
    marcas, textos = marcas_fecha(dias)
    fig.update_xaxes(tickvals=marcas, ticktext=textos, showgrid=False)
    fig.update_yaxes(tickprefix="$", tickformat=",.2f", rangemode="tozero")
    return estilo(fig)


def grafica_calor(ventas):
    """Mapa de calor de ventas por día de la semana y hora."""
    tabla = ventas.assign(Dia=ventas["Fecha"].dt.day_of_week).pivot_table(
        index="Dia", columns="HoraNum", values="Total", aggfunc="sum", fill_value=0)
    horas = range(tabla.columns.min(), tabla.columns.max() + 1)
    tabla = tabla.reindex(index=range(7), columns=horas, fill_value=0)
    fig = go.Figure(go.Heatmap(
        z=tabla.values, x=[f"{h}h" for h in horas], y=DIAS,
        colorscale=[[0, "#16273F"], [1, ACENTO]], xgap=2, ygap=2,
        customdata=[[moneda(v) for v in fila] for fila in tabla.values],
        hovertemplate="%{y} a las %{x}<br><b>%{customdata}</b><extra></extra>",
        # Horizontal: en celular una barra vertical con centavos le quita ancho al mapa
        colorbar=dict(orientation="h", y=-0.14, yanchor="top", thickness=10, nticks=4,
                      tickprefix="$", tickformat=",.2f", outlinewidth=0),
    ))
    fig.update_xaxes(showgrid=False, type="category")
    fig.update_yaxes(showgrid=False, autorange="reversed")
    return estilo(fig, alto=410)


def grafica_barras_h(etiquetas, valores, hover, texto, alto=None):
    """Barras horizontales de una sola serie, con la mayor arriba."""
    fig = go.Figure(go.Bar(
        x=valores, y=etiquetas, orientation="h", marker=dict(color=ACENTO, cornerradius=4),
        text=texto, textposition="outside", textfont=dict(color=TEXTO), cliponaxis=False,
        customdata=hover, hovertemplate="%{y}<br><b>%{customdata}</b><extra></extra>",
    ))
    fig.update_xaxes(visible=False, range=[0, max(valores) * 1.5])
    fig.update_yaxes(autorange="reversed", showgrid=False, tickfont=dict(color=TEXTO))
    return estilo(fig, alto=alto or 56 * len(etiquetas) + 40)


def grafica_productos(productos):
    """Barras de unidades vendidas por producto."""
    return grafica_barras_h(
        productos["Producto"], productos["Cantidad"],
        [f"{c:,.0f} piezas · {moneda(t)}" for c, t in zip(productos["Cantidad"], productos["Total"])],
        [f"{c:,.0f}" for c in productos["Cantidad"]],
        alto=38 * len(productos) + 40,
    )


def grafica_montos(serie):
    """Barras de ventas por grupo, de mayor a menor (pago y cajero)."""
    serie = serie.sort_values(ascending=False)
    textos = [moneda(v) for v in serie]
    return grafica_barras_h(serie.index, serie.values, textos, textos)


def color_categoria(categorias):
    """Asigna colores fijos; las categorías desconocidas toman los libres."""
    libres = [c for c in COLORES_SERIES if c not in
              [COLOR_CATEGORIA[n] for n in categorias if n in COLOR_CATEGORIA]]
    colores = []
    for categoria in categorias:
        if categoria in COLOR_CATEGORIA:
            colores.append(COLOR_CATEGORIA[categoria])
        else:
            colores.append(libres.pop(0) if libres else COLOR_OTROS)
    return colores


def grafica_categorias(ventas):
    """Dona de ventas por categoría, con leyenda y porcentaje directo."""
    serie = ventas.groupby("Categoría")["Total"].sum().sort_values(ascending=False)
    fig = go.Figure(go.Pie(
        labels=serie.index, values=serie.values, hole=0.62, sort=False, direction="clockwise",
        marker=dict(colors=color_categoria(list(serie.index)), line=dict(color=SUPERFICIE, width=2)),
        textinfo="percent", textposition="outside", textfont=dict(color=TEXTO), automargin=True,
        customdata=[moneda(v) for v in serie.values],
        hovertemplate="%{label}<br><b>%{customdata}</b> · %{percent}<extra></extra>",
    ))
    return estilo(fig, alto=380, leyenda=True)


def grafica_mensual(ventas, gastos, inicio, fin):
    """Ingresos, gastos y utilidad por mes en barras agrupadas."""
    gastos_rango = gastos_del_rango(gastos, inicio, fin)
    rango = filtrar_fechas(ventas, inicio, fin)
    ingresos = rango.groupby(rango["Fecha"].dt.to_period("M"))["Total"].sum().reindex(gastos_rango.index, fill_value=0)
    utilidad = ingresos - gastos_rango
    meses = [mes_es(m) for m in gastos_rango.index]

    series = (("Ingresos", ingresos, COLORES_SERIES[0]), ("Gastos", gastos_rango, COLORES_SERIES[1]),
              ("Utilidad", utilidad, COLORES_SERIES[2]))
    fig = go.Figure()
    for nombre, valores, color in series:
        fig.add_trace(go.Bar(
            name=nombre, x=meses, y=valores.values, marker=dict(color=color, cornerradius=4),
            customdata=[moneda(v) for v in valores.values],
            hovertemplate=f"{nombre}: <b>%{{customdata}}</b><extra></extra>",
        ))
    # Solo la utilidad lleva etiqueta: es el dato que importa de esta gráfica
    fig.data[2].update(text=[moneda(v) for v in utilidad.values],
                       textposition="outside", textangle=-90, textfont=dict(color=TEXTO), cliponaxis=False,
                       constraintext="none")
    fig.update_layout(barmode="group", bargap=0.3, bargroupgap=0.06, hovermode="x unified")
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(tickprefix="$", tickformat=",.2f")
    return estilo(fig, alto=400, leyenda=True)


def mostrar_grafica(fig):
    """Muestra una gráfica a todo el ancho y sin barra de herramientas."""
    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False, "scrollZoom": False})


# ---------------------------------------------------------------------------
# Interfaz
# ---------------------------------------------------------------------------
def aplicar_css():
    """Ajustes que config.toml no cubre: tarjetas y márgenes en celular."""
    st.markdown(f"""
    <style>
      h1 {{ font-size: clamp(1.5rem, 1.1rem + 2vw, 2.25rem) !important; line-height: 1.2 !important; }}
      [data-testid="stMetric"] {{ background: {SUPERFICIE}; border: 1px solid {BORDE}; border-radius: 12px; padding: 14px 16px; }}
      [data-testid="stMetricValue"] {{ font-size: clamp(1.4rem, 1.2rem + .6vw, 1.75rem); font-variant-numeric: tabular-nums; }}
      .pie-demo {{ text-align: center; color: {TEXTO_SUAVE}; font-size: .9rem; margin-top: 2rem; padding-top: 1rem; border-top: 1px solid {BORDE}; }}
      @media (max-width: 640px) {{
        .block-container {{ padding: 3.5rem 1rem 2rem !important; }}
      }}
    </style>
    """, unsafe_allow_html=True)


def elegir_origen():
    """Dibuja la sección de datos y devuelve (archivo subido o None, nombre del negocio)."""
    with st.sidebar.expander("Usar mi propio Excel", expanded=True):
        negocio = st.text_input("Nombre del negocio", max_chars=60, placeholder=NEGOCIO_VACIO)
        st.caption("Debe tener las hojas Ventas, Gastos e Inventario con las mismas columnas que el ejemplo.")
        subido = st.file_uploader("Sube tu archivo de Excel", type="xlsx")
        st.caption("Tu archivo se procesa solo durante tu visita y no se guarda.")
        st.download_button("Descargar hoja para llenar", PLANTILLA.read_bytes(), file_name=PLANTILLA.name,
                           mime=MIME_EXCEL, width="stretch")
        st.download_button("Descargar Excel de ejemplo", ARCHIVO_DEMO.read_bytes(),
                           file_name=ARCHIVO_DEMO.name, mime=MIME_EXCEL, width="stretch")
    return subido, negocio.strip()


def elegir_filtros(ventas, origen):
    """Dibuja los filtros y devuelve (inicio, fin, categorías, métodos)."""
    primera = ventas["Fecha"].min().date()
    ultima = ventas["Fecha"].max().date()
    # Por defecto, los últimos 3 meses: deja un periodo anterior para comparar
    inicial = max(primera, (pd.Period(ultima, freq="M") - 2).start_time.date())

    st.sidebar.header("Ver por fecha o categoría")
    rango = st.sidebar.date_input("Rango de fechas", value=(inicial, ultima), min_value=primera,
                                  max_value=ultima, format="DD/MM/YYYY", key=f"fechas_{origen}")
    # Mientras la persona elige la segunda fecha, llega solo una
    inicio, fin = rango if len(rango) == 2 else (rango[0], rango[0])
    categorias = st.sidebar.multiselect("Categoría", sorted(ventas["Categoría"].unique()),
                                        placeholder="Todas", key=f"categorias_{origen}")
    metodos = st.sidebar.multiselect("Método de pago", sorted(ventas["Método de pago"].unique()),
                                     placeholder="Todos", key=f"metodos_{origen}")
    return inicio, fin, categorias, metodos


def mostrar_indicadores(ventas, gastos, filtradas, inicio, fin, categorias, metodos):
    """Fila de indicadores con su variación contra el periodo anterior."""
    inicio_ant, fin_ant = periodo_anterior(inicio, fin)
    hay_anterior = inicio_ant >= ventas["Fecha"].min().date()

    actual = calcular_indicadores(filtradas)
    anterior = {}
    if hay_anterior:
        anterior = calcular_indicadores(filtrar_ventas(ventas, inicio_ant, fin_ant, categorias, metodos))
    utilidad = calcular_utilidad(ventas, gastos, inicio, fin)
    utilidad_ant = calcular_utilidad(ventas, gastos, inicio_ant, fin_ant) if hay_anterior else None

    columnas = st.columns(4)
    columnas[0].metric("Ventas totales", moneda(actual["ventas"]), variacion(actual["ventas"], anterior.get("ventas")))
    columnas[1].metric("Tickets", f"{actual['tickets']:,}", variacion(actual["tickets"], anterior.get("tickets")))
    columnas[2].metric("Ticket promedio", moneda(actual["ticket_promedio"]),
                       variacion(actual["ticket_promedio"], anterior.get("ticket_promedio")))
    columnas[3].metric("Utilidad", moneda(utilidad) if utilidad is not None else "Sin gastos",
                       variacion(utilidad, utilidad_ant))

    if hay_anterior:
        st.caption(f"Variación contra el periodo anterior: {fecha_es(inicio_ant)} a {fecha_es(fin_ant)}.")
    else:
        st.caption("Sin variación: los datos no alcanzan para un periodo anterior equivalente.")
    if utilidad is None:
        st.warning('La hoja "Gastos" no tiene todos los meses de este rango; la utilidad no se puede calcular.')
    elif categorias or metodos:
        st.caption("La utilidad y la gráfica de ingresos vs. gastos consideran todas las ventas del rango, "
                   "porque los gastos no se pueden repartir por categoría ni por método de pago.")


def mostrar_graficas(ventas, gastos, filtradas, inicio, fin):
    """Dibuja las siete gráficas en el orden del tablero."""
    st.subheader("Ventas por día")
    mostrar_grafica(grafica_ventas_diarias(filtradas, inicio, fin))

    st.subheader("¿Cuándo se vende más?")
    mostrar_grafica(grafica_calor(filtradas))

    productos = (filtradas.groupby("Producto")[["Cantidad", "Total"]].sum()
                 .sort_values("Cantidad", ascending=False).reset_index())
    izquierda, derecha = st.columns([3, 2])
    with izquierda:
        st.subheader("Top 10 productos")
        mostrar_grafica(grafica_productos(productos.head(10)))
    with derecha:
        st.subheader("Los 5 menos vendidos")
        if len(productos) > 5:
            mostrar_grafica(grafica_productos(productos.tail(5).iloc[::-1]))
        else:
            st.caption("Con estos filtros hay muy pocos productos para comparar.")

    izquierda, derecha = st.columns(2)
    with izquierda:
        st.subheader("Ventas por categoría")
        mostrar_grafica(grafica_categorias(filtradas))
    with derecha:
        st.subheader("Ventas por método de pago")
        mostrar_grafica(grafica_montos(filtradas.groupby("Método de pago")["Total"].sum()))
        st.subheader("Ventas por cajero")
        mostrar_grafica(grafica_montos(filtradas.groupby("Cajero")["Total"].sum()))

    st.subheader("Ingresos vs. gastos por mes")
    if gastos_del_rango(gastos, inicio, fin) is None:
        st.caption('Faltan meses en la hoja "Gastos" para mostrar esta gráfica.')
    else:
        mostrar_grafica(grafica_mensual(ventas, gastos, inicio, fin))


def mostrar_resurtir(ventas, inventario):
    """Muestra los productos que se van a acabar, del más urgente al menos."""
    st.subheader("¿Qué conviene resurtir?")
    tabla = tabla_resurtir(ventas, inventario)
    urgentes = int((tabla["Estado"] == "Urgente").sum())
    pronto = int((tabla["Estado"] == "Pronto").sum())
    columnas = st.columns(2)
    columnas[0].metric(f"Urgente: menos de {DIAS_URGENTE} días", urgentes)
    columnas[1].metric(f"Pronto: menos de {DIAS_PRONTO} días", pronto)
    st.dataframe(
        tabla[tabla["Estado"].isin(["Urgente", "Pronto"])], hide_index=True, width="stretch",
        column_config={
            "Existencia": st.column_config.NumberColumn(format="%d"),
            "Venta diaria": st.column_config.NumberColumn(format="%.1f"),
            "Días de existencia": st.column_config.NumberColumn(format="%.0f"),
        },
    )
    st.caption(f"Calculado con las ventas de los últimos {DIAS_CONSUMO} días del archivo; no cambia con los filtros.")


def mostrar_pie(es_demo):
    """Pie de página; "de ejemplo" solo aplica al Excel de ejemplo."""
    texto = "Tablero de ejemplo creado por Eleva Dev" if es_demo else "Tablero creado por Eleva Dev"
    st.markdown(f'<div class="pie-demo">{texto}</div>', unsafe_allow_html=True)


def main():
    st.set_page_config(page_title=f"Tablero de ventas · {NEGOCIO_DEMO}", page_icon="📊", layout="wide")
    aplicar_css()
    zona_titulo = st.empty()  # el título depende del Excel, que se conoce después
    subido, negocio = elegir_origen()

    libro, origen = None, "demo"
    if subido is None:
        st.session_state.pop("libro", None)
        st.session_state.pop("libro_id", None)
    else:
        try:
            libro, origen = cargar_subido(subido), subido.file_id
        except ErrorFormato as error:
            st.error(f"No pude leer tu Excel. {error} Te muestro los datos de ejemplo.")

    es_demo = libro is None
    if es_demo:
        libro = cargar_demo()
        st.info("Este es un tablero de ejemplo. Sube tu Excel en el menú lateral para ver el tuyo.")

    nombre_negocio = NEGOCIO_DEMO if es_demo else negocio or NEGOCIO_VACIO
    titulo = f"Tablero de ventas · {nombre_negocio}"
    st.set_page_config(page_title=titulo, page_icon="📊", layout="wide")
    zona_titulo.title(titulo)

    ventas, gastos, inventario = libro
    inicio, fin, categorias, metodos = elegir_filtros(ventas, origen)
    filtradas = filtrar_ventas(ventas, inicio, fin, categorias, metodos)

    if filtradas.empty:
        st.info("No hay ventas con estos filtros. Prueba con otro rango de fechas.")
    else:
        mostrar_indicadores(ventas, gastos, filtradas, inicio, fin, categorias, metodos)
        mostrar_graficas(ventas, gastos, filtradas, inicio, fin)
        st.download_button("Descargar datos filtrados en Excel", excel_filtrado(filtradas),
                           file_name=f"ventas_{nombre_archivo(nombre_negocio)}.xlsx", mime=MIME_EXCEL)
    mostrar_resurtir(ventas, inventario)
    mostrar_pie(es_demo)


if __name__ == "__main__":
    main()
