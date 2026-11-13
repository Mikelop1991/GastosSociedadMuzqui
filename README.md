# Sociedad Muzki

Las cuentas de la soci: apuntar un sobre (cena, bebida y leña) o un recibo en dos toques, y sacar la misma Excel de siempre (`Gastos_Muzqui_<año>-<año+1>.xlsx`, el año de la soci va de marzo a febrero).

- **App**: `index.html` (PWA, sin dependencias). Se publica con GitHub Pages.
- **Datos**: una hoja de Google Sheets (`Muzki_datos_app`, sube `Muzki_datos_app.xlsx` a Drive y ábrela con Hojas de cálculo) con la pestaña `Movimientos` (id, fecha, concepto, importe, nota, sobre, creado, modificado). Cada sobre son hasta tres filas (SOBRE CENA, SOBRE BEBIDA, SOBRE LEÑA) con el mismo código en `sobre`.
- **Back**: `Code.gs`, pegado en Extensiones → Apps Script de esa hoja e implementado como aplicación web (ejecutar como: yo · acceso: cualquier usuario).
- **Excel**: `plantilla.xlsx` es la Excel de la soci vacía, sin ningún dato, con la identidad de la app (escudo, verde entra / rojo sale). RESUMEN con tarjetas y gráficos, un mes por pestaña (mismas columnas de siempre, totales en la fila 4, datos desde la 6) y FIESTAS. `{ANIO}` y `{FIESTAS}` se cambian por el año al rellenarla. La app la rellena en el propio aparato: un sobre por fila, cada pago o ingreso en su fila, con «Observaciones» como antes.

Conceptos (= columnas de la Excel). Sale: Iberdrola (luz), Leña, Internet, Urtasun, Bebida, Mancomunidad, Varios, Seguro (en la Excel va en la columna EXTRA con «Seguro» en Observaciones), Extra.

Fiestas: una cuenta por año, aparte (pestaña Fiestas de la app y hoja FIESTAS de la Excel). Cada línea escrita a mano son hasta dos filas con el mismo código en `sobre`: FIESTAS GASTO y FIESTAS INGRESO, o FIESTAS AYTO si lo paga el ayuntamiento. No suman en los meses ni en el resumen anual. Entra: Sobre cena, Sobre bebida, Ingreso caja, Sobre leña, Internet (fútbol), Ingreso extra.

La app se configura sola al abrir el enlace de instalación `…/<repo>/#u=<URL /exec>&k=<clave>`. La clave no está en este repositorio.

## La Excel por correo cada mes

En `Code.gs` rellena `CORREO_EXCEL` (uno o varios correos separados por comas) y `URL_PLANTILLA` (`https://<usuario>.github.io/<repo>/plantilla.xlsx`), guarda y ejecuta una vez `activarEnvioMensual` (Google pedirá permiso para Gmail, Drive y conexiones externas). El día 1 de cada mes, hacia las 8, llega un correo con el resumen del mes cerrado y la Excel del año (marzo a febrero) adjunta, y se guarda una copia en la carpeta «Muzki · Excel» de Drive. `probarEnvio` lo manda en el momento.

## Probar sin Google

Abre `index.html` desde GitHub Pages sin enlace de instalación: pide la conexión en Ajustes. Los datos de la soci están en la hoja de Google, nunca en este repositorio (es público).
