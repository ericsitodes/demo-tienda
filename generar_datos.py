"""Genera el Excel de ejemplo con datos ficticios de "Abarrotes El Puerto".

Uso: python generar_datos.py

La semilla es fija, así que cada ejecución produce exactamente los mismos datos.
"""

from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import Font

FECHA_INICIO = "2026-04-01"
FECHA_FIN = "2026-09-30"
DIAS_RECIENTES = 30

SEMILLA = 2027
ARCHIVO = Path(__file__).parent / "ventas_tienda.xlsx"

# Catálogo: (producto, categoría, precio en MXN, popularidad relativa)
CATALOGO = [
    ("Refresco de cola 600 ml", "Bebidas", 20, 60), ("Agua natural 1 L", "Bebidas", 14, 44),
    ("Jugo de naranja 1 L", "Bebidas", 32, 16), ("Cerveza lata 355 ml", "Bebidas", 24, 30),
    ("Papas fritas 45 g", "Botanas", 19, 34), ("Cacahuates 100 g", "Botanas", 22, 14),
    ("Galletas de chocolate", "Botanas", 21, 26), ("Dulces surtidos", "Botanas", 8, 22),
    ("Leche entera 1 L", "Lácteos", 29, 42), ("Yogur natural 1 L", "Lácteos", 45, 12),
    ("Queso fresco 400 g", "Lácteos", 68, 14), ("Crema 200 ml", "Lácteos", 24, 11),
    ("Arroz 1 kg", "Abarrotes", 32, 24), ("Frijol negro 1 kg", "Abarrotes", 42, 22),
    ("Aceite vegetal 1 L", "Abarrotes", 48, 18), ("Azúcar 1 kg", "Abarrotes", 34, 18),
    ("Huevo 12 piezas", "Abarrotes", 44, 30), ("Tortillas 1 kg", "Abarrotes", 26, 36),
    ("Atún en lata", "Abarrotes", 23, 15), ("Café soluble 100 g", "Abarrotes", 72, 6),
    ("Detergente 1 kg", "Limpieza", 46, 10), ("Cloro 1 L", "Limpieza", 18, 10),
    ("Jabón de trastes", "Limpieza", 25, 9), ("Papel higiénico 4 rollos", "Limpieza", 36, 12),
]

METODOS_PAGO = ["Efectivo", "Tarjeta", "Transferencia"]
PESOS_PAGO = [0.68, 0.22, 0.10]
CAJEROS = {"Marta": (0, 1.3), "Beto": (2, 1.1), "Lupita": (4, 1.0), "Iván": (6, 0.8)}
TICKETS_POR_DIA = [108, 100, 102, 106, 128, 150, 132]
FACTOR_MES = {4: 1.0, 5: 1.03, 6: 0.98, 7: 1.08, 8: 1.05, 9: 0.88}
PESOS_HORA = {7: 5, 8: 8, 9: 8, 10: 6, 11: 6, 12: 7, 13: 8, 14: 8, 15: 6, 16: 5,
              17: 7, 18: 10, 19: 12, 20: 11, 21: 7}
PRODUCTOS_POR_TICKET = ([1, 2, 3, 4, 5, 6, 8], [0.28, 0.27, 0.18, 0.11, 0.08, 0.05, 0.03])
COSTO_MERCANCIA = (0.71, 0.74)
GASTOS_FIJOS = {"Renta": 14000, "Nómina": 34000, "Otros": 6000}
LUZ_ALTA, LUZ_BAJA = (7000, 8500), (5000, 6000)


def generar_ventas(rng):
    """Genera una fila por producto de cada ticket."""
    horas = np.array(list(PESOS_HORA))
    pesos_hora = np.array(list(PESOS_HORA.values()), dtype=float)
    pesos_hora /= pesos_hora.sum()
    pesos_catalogo = np.array([p[3] for p in CATALOGO], dtype=float)
    pesos_catalogo /= pesos_catalogo.sum()
    tamanos, pesos_tamano = PRODUCTOS_POR_TICKET

    filas = []
    folio = 1
    for dia in pd.date_range(FECHA_INICIO, FECHA_FIN):
        media = TICKETS_POR_DIA[dia.day_of_week] * FACTOR_MES[dia.month]
        turno = [(n, p) for n, (descanso, p) in CAJEROS.items() if descanso != dia.day_of_week]
        nombres = [n for n, _ in turno]
        pesos = np.array([p for _, p in turno])

        registros = sorted(
            (int(rng.choice(horas, p=pesos_hora)), int(rng.integers(0, 60)))
            for _ in range(rng.poisson(media))
        )
        for hora, minuto in registros:
            ticket = f"T-{folio:05d}"
            folio += 1
            cajero = str(rng.choice(nombres, p=pesos / pesos.sum()))
            metodo = str(rng.choice(METODOS_PAGO, p=PESOS_PAGO))
            indices = rng.choice(len(CATALOGO), size=rng.choice(tamanos, p=pesos_tamano), p=pesos_catalogo)
            for indice, cantidad in zip(*np.unique(indices, return_counts=True)):
                producto, categoria, precio, _ = CATALOGO[indice]
                filas.append({
                    "Fecha": dia.date(),
                    "Hora": f"{hora:02d}:{minuto:02d}",
                    "Ticket": ticket,
                    "Producto": producto,
                    "Categoría": categoria,
                    "Cantidad": int(cantidad),
                    "Precio unitario": precio,
                    "Total": int(cantidad) * precio,
                    "Método de pago": metodo,
                    "Cajero": cajero,
                })
    return pd.DataFrame(filas)


def generar_gastos(rng, ventas):
    """Genera los gastos mensuales; la mercancía depende de lo vendido."""
    meses = pd.to_datetime(ventas["Fecha"]).dt.to_period("M")
    filas = []
    for mes, total in ventas.groupby(meses)["Total"].sum().items():
        luz = LUZ_ALTA if mes.month in (5, 6, 7, 8) else LUZ_BAJA  # el aire acondicionado sube la luz
        conceptos = {"Mercancía": total * rng.uniform(*COSTO_MERCANCIA), **GASTOS_FIJOS, "Luz": rng.integers(*luz)}
        for concepto, monto in conceptos.items():
            filas.append({"Mes": mes.to_timestamp().date(), "Concepto": concepto, "Monto": round(float(monto), 2)})
    return pd.DataFrame(filas)


def generar_inventario(rng, ventas):
    """Da a cada producto entre 2 y 45 días de existencia al ritmo reciente."""
    fechas = pd.to_datetime(ventas["Fecha"])
    recientes = ventas[fechas > fechas.max() - pd.Timedelta(days=DIAS_RECIENTES)]
    venta_diaria = recientes.groupby("Producto")["Cantidad"].sum() / DIAS_RECIENTES
    filas = []
    for producto, _, _, _ in CATALOGO:
        dias = rng.uniform(2, 45)
        filas.append({"Producto": producto, "Existencia": int(round(venta_diaria.get(producto, 0) * dias))})
    return pd.DataFrame(filas)


def guardar_excel(ventas, gastos, inventario):
    """Escribe las tres hojas con formato de fecha, moneda y anchos legibles."""
    with pd.ExcelWriter(ARCHIVO, engine="openpyxl") as escritor:
        ventas.to_excel(escritor, sheet_name="Ventas", index=False)
        gastos.to_excel(escritor, sheet_name="Gastos", index=False)
        inventario.to_excel(escritor, sheet_name="Inventario", index=False)

        formatos = {"Fecha": "dd/mm/yyyy", "Mes": "mm/yyyy", "Precio unitario": '"$"#,##0.00',
                    "Total": '"$"#,##0.00', "Monto": '"$"#,##0.00'}
        for hoja in escritor.sheets.values():
            hoja.freeze_panes = "A2"
            for columna in hoja.columns:
                encabezado = columna[0].value
                columna[0].font = Font(bold=True)
                ancho = 28 if encabezado == "Producto" else max(len(str(encabezado)), 12)
                hoja.column_dimensions[columna[0].column_letter].width = ancho + 2
                if encabezado in formatos:
                    for celda in columna[1:]:
                        celda.number_format = formatos[encabezado]


def main():
    rng = np.random.default_rng(SEMILLA)
    ventas = generar_ventas(rng)
    gastos = generar_gastos(rng, ventas)
    inventario = generar_inventario(rng, ventas)
    guardar_excel(ventas, gastos, inventario)
    print(f"Listo: {ARCHIVO.name}")
    print(f"{len(ventas):,} filas, {ventas['Ticket'].nunique():,} tickets, ${ventas['Total'].sum():,.2f} en ventas")


if __name__ == "__main__":
    main()
