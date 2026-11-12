/**
 * SOCIEDAD MUZKI · unificar las notas escritas de distintas formas (se ejecuta UNA vez).
 * En Apps Script: «+» junto a Archivos → Script → llámalo «Unificar», pega esto, guarda,
 * elige «unificarNotas» en el desplegable y dale a Ejecutar. Hace antes una copia de la hoja en Drive.
 * Mira el registro (abajo): dice cuántas notas ha cambiado. Después puedes borrar este archivo.
 */
const UNIFICAR = {
 "INTERNET": {
  "Ikai internet": "Ikai",
  "Ikai 3.6 internet": "Ikai",
  "Ikai 2.7 internet": "Ikai",
  "Ikai 5.8 internet": "Ikai",
  "Ikay": "Ikai",
  "IKAY": "Ikai",
  "": "Ikai"
 },
 "COBRO INTERNET": {
  "Futbol": "Fútbol",
  "Futbol 12 y": "Fútbol",
  "Futbol e internet": "Fútbol",
  "Furtbol": "Fútbol",
  "": "Fútbol"
 },
 "EXTRA": {
  "Gastos mantto.cuenta": "Comisión mantenimiento cuenta",
  "Mto cta.": "Comisión mantenimiento cuenta",
  "1.10.26 comision mto.cuenta": "Comisión mantenimiento cuenta",
  "Comision mto.cueta": "Comisión mantenimiento cuenta",
  "Comision servicio b.d": "Comisión servicio banco",
  "Dejados cambios de 2  y 05 euros": "Cambios"
 },
 "VARIOS": {
  "Se queda para cambios": "Cambios",
  "Se quedan para cambios": "Cambios",
  "Pago bm": "BM",
  "Pago fra. bm": "BM",
  "Pago fra. mercadona": "Mercadona",
  "Pago fra. inaci": "Factura Inaci",
  "Pago fra, inaci": "Factura Inaci",
  "Pago fra. aceite coop": "Aceite Coop",
  "Pago fra, ccop hostel": "Coop Hostel",
  "SAGASTI chuches": "Sagasti chuches",
  "Pago fra. unide/sagasti": "Unide / Sagasti",
  "Servi hielo": "Servi Hielo",
  "Pago 2 facturas 16,50y 7,5": "2 facturas (16,50 y 7,50)"
 },
 "BEBIDA": {
  "Pago eroski": "Eroski",
  "Pago fra. Cervezas sagrario": "Cervezas Sagrario",
  "Bodegas sagasti(vino)": "Bodegas Sagasti (vino)",
  "Compra bebida fiestas (15.5)": "Bebida para fiestas",
  "Licores aza.": "Licores Aza"
 },
 "MANCOMUNIDAD": {
  "Recibo agua": ""
 },
 "INGRESO CAJA": {
  "(cambiosmonedad 0,20 y 010": "Cambios en monedas"
 },
 "INGRESO EXTRA": {
  "Cuota ana belen": "Cuota Ana Belén",
  "Cuota patxi": "Cuota Patxi"
 },
 "FIESTAS": {
  "Electriicidad Iraitz": "Electricidad Iraitz",
  "Carniceria": "Carnicería",
  "Préstamo y devolucion de dinero a soci": "Préstamo y devolución de dinero a la soci",
  "Bebidas que sobra y compra la soci": "Bebida que sobra y compra la soci",
  "Chuches sagasti": "Chuches Sagasti"
 }
};

function unificarNotas() {
  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName('Movimientos');
  DriveApp.getFileById(ss.getId()).makeCopy(ss.getName() + ' · antes de unificar');
  const n = hoja.getLastRow() - 1;
  if (n < 1) return;
  const rango = hoja.getRange(2, 3, n, 3);            // concepto, importe, nota
  const v = rango.getValues();
  let cambios = 0;
  const notas = v.map(fila => {
    const c = String(fila[0]).toUpperCase().trim();
    const tabla = UNIFICAR[c.indexOf('FIESTAS') === 0 ? 'FIESTAS' : c];
    const nota = String(fila[2]);
    if (tabla && Object.prototype.hasOwnProperty.call(tabla, nota) && tabla[nota] !== nota) { cambios++; return [tabla[nota]]; }
    return [nota];
  });
  hoja.getRange(2, 5, n, 1).setValues(notas);
  Logger.log('Notas unificadas: ' + cambios);
}
