// ============================================================
// Convierte el Excel de la lista de precios en productos.json
//
// Uso:   node herramientas/lista.js "<ruta al Excel>" [dd/mm/aaaa]
//        (si no se pasa la fecha, se usa la de hoy)
//
// Columnas que usa (el resto se ignora, por ej. PROVEEDOR y NOMBRE no se publican):
//   CODIGO | DETALLE | RUBRO | ESTADO | PRECIO (sin IVA) | IVA (%, opcional: si falta, 21)
// ============================================================
const X = require('xlsx');
const fs = require('fs');
const path = require('path');

const [, , src, fechaArg] = process.argv;
if (!src) { console.error('Falta la ruta al Excel de la lista'); process.exit(1); }
const fecha = fechaArg || new Date().toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric' });
const out = path.join(__dirname, '..', 'productos.json');

const wb = X.readFile(src);
const filas = X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' })
  .map(f => Object.fromEntries(Object.entries(f).map(([k, v]) => [String(k).trim().toUpperCase(), v])));
for (const col of ['CODIGO', 'DETALLE', 'PRECIO']) {
  if (!(col in (filas[0] || {}))) { console.error(`Falta la columna ${col}`); process.exit(1); }
}

const txt = v => String(v ?? '').replace(/\s+/g, ' ').trim();
const orden = ['ACTIVO', 'ACTIVO VEP', 'A PEDIDO', 'PAUSADO', 'SUSPENDIDO', 'DISCONTINUO', 'INACTIVO'];
const estadosUsados = new Set(filas.map(f => txt(f.ESTADO)));
const estados = [...new Set([...orden, ...estadosUsados])].filter(e => estadosUsados.has(e));
const rubros = [...new Set(filas.map(f => txt(f.RUBRO)))].sort();

const productos = filas.filter(f => txt(f.DETALLE)).map(f => {
  const fila = [txt(f.CODIGO), txt(f.DETALLE), rubros.indexOf(txt(f.RUBRO)), estados.indexOf(txt(f.ESTADO)),
    Math.round((Number(f.PRECIO) || 0) * 100) / 100];
  const iva = Number(f.IVA);
  if (f.IVA !== undefined && f.IVA !== '' && iva !== 21) fila.push(iva);   // solo se guarda si no es 21%
  return fila;
});

// Resumen de cambios contra la lista publicada
let antes = new Map();
try { antes = new Map(JSON.parse(fs.readFileSync(out, 'utf8')).productos.map(p => [p[0], p[4]])); } catch (e) {}
let cambian = 0, nuevos = 0;
for (const p of productos) {
  if (!antes.has(p[0])) nuevos++;
  else if (Math.abs(antes.get(p[0]) - p[4]) > 0.009) cambian++;
}
const nuevosCodigos = new Set(productos.map(p => p[0]));
const quitados = [...antes.keys()].filter(c => !nuevosCodigos.has(c)).length;

fs.writeFileSync(out, JSON.stringify({ fecha, estados, rubros, productos }));
console.log(`Lista al ${fecha}: ${productos.length} productos`);
console.log(`  precios que cambian: ${cambian} · nuevos: ${nuevos} · quitados: ${quitados}`);
console.log(`  con IVA distinto de 21%: ${productos.filter(p => p.length > 5).length}`);
