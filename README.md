# Dashboard de Ventas · Abarrotes El Puerto

Demo de Eleva Dev: un tablero de ventas para tienda de abarrotes hecho con Python, pandas, Plotly y Streamlit. **Todos los datos son ficticios.**

## Correrlo en tu computadora

Necesitas Python 3.11 o más nuevo.

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Se abre en http://localhost:8501. Para detenerlo, presiona `Ctrl + C` en la terminal.

## Archivos

| Archivo | Para qué sirve |
| --- | --- |
| `app.py` | El dashboard |
| `generar_datos.py` | Crea `ventas_tienda.xlsx`; la semilla fija da siempre los mismos datos |
| `ventas_tienda.xlsx` | Datos de ejemplo |
| `plantilla-tienda-farmacia.xlsx` | Plantilla en blanco que se descarga desde la barra lateral |
| `.gitignore` | Evita subir Excel de clientes, entornos y secretos |
| `requirements.txt` | Librerías que instala Streamlit Community Cloud |
| `.streamlit/config.toml` | Colores del tema |

## Formato para subir tu propio Excel

Hoja **Ventas**: `Fecha`, `Hora`, `Ticket`, `Producto`, `Categoría`, `Cantidad`, `Precio unitario`, `Total`, `Método de pago`, `Cajero`. Cada fila es un producto de un ticket.

Hoja **Gastos**: `Mes` (cualquier fecha de ese mes), `Concepto`, `Monto`.

Hoja **Inventario**: `Producto`, `Existencia`.

## Cómo se publica

Demo en línea: https://demo-tienda.streamlit.app/

Está en Streamlit Community Cloud y se actualiza solo con cada push a la rama `main`:

```powershell
git add .
git commit -m "Describe tu cambio"
git push origin main
```

Streamlit apaga las apps gratuitas sin visitas: abre el demo antes de cada reunión.

## Dónde se cambian los datos

- **Datos de ejemplo:** edita los catálogos y parámetros de `generar_datos.py` y ejecuta `python generar_datos.py`. Se reescribe `ventas_tienda.xlsx`.
- **Nombre del negocio de ejemplo:** constante `NEGOCIO_DEMO`, al inicio de `app.py`.
- **Plantilla:** reemplaza `plantilla-tienda-farmacia.xlsx` por la versión de la carpeta `plantillas/`.
- **Excel de un cliente:** se sube desde la barra lateral. Nunca se copia a este repositorio; se guarda solo en `C:\Users\eric\clientes\`.

## Privacidad del Excel subido

La app no escribe el archivo en disco ni lo manda a otro servicio. Se procesa en el servidor de Streamlit y queda solo en la sesión de la visita (`st.session_state`), que se borra al cerrarla. No uses `st.cache_data` con datos subidos: esa caché sobrevive a la visita, y dejaría de ser cierto el texto que se muestra bajo el botón de subir.

## Cómo se calculan los indicadores

- **Periodo anterior:** si eliges meses completos, se compara contra los mismos meses previos. Si no, contra la misma cantidad de días justo antes.
- **Utilidad:** ventas menos gastos. Si el rango cubre solo parte de un mes, el gasto de ese mes se reparte por días.
- La utilidad y la gráfica de ingresos vs. gastos ignoran los filtros de categoría y método de pago, porque los gastos no se pueden repartir de esa forma.
- **Resurtido:** existencia entre la venta diaria de los últimos 30 días del archivo. Es **Urgente** con menos de 7 días y **Pronto** con menos de 15. No cambia con los filtros.
