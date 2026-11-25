// ============================================================
// Genera las condiciones comerciales cifradas de cada cliente.
//
// Uso:   node herramientas/clientes.js "<ruta al Excel de clientes>"
//
// Excel (primera hoja), una fila por cliente:
//   CUIT | RAZON SOCIAL | CLAVE | DESCUENTO
//   - CLAVE vacía  -> se genera una y se guarda en el mismo Excel.
//   - DESCUENTO    -> % que se le aplica al cliente cuando se identifica.
//
// Salida: carpeta clientes/ con un archivo cifrado por cliente. Solo se puede
// abrir con CUIT + CLAVE, así que se puede publicar aunque el repo sea público.
// El Excel de clientes NO va al repositorio (tiene las claves).
// ============================================================
const X = require('xlsx');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const ITER = 150000;
const ALFABETO = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // sin 0/O ni 1/I para no confundir

const src = process.argv[2];
if (!src) { console.error('Falta la ruta al Excel de clientes'); process.exit(1); }
const outDir = path.join(__dirname, '..', 'clientes');

const soloDigitos = s => String(s ?? '').replace(/\D/g, '');
const nuevaClave = () => Array.from(crypto.randomBytes(8), b => ALFABETO[b % ALFABETO.length]).join('');
const pct = v => { const n = parseFloat(String(v ?? '').replace(',', '.').replace('%', '')); return isFinite(n) ? Math.max(0, Math.min(99, n)) : null; };

// Misma derivación que la página: 16 bytes -> nombre del archivo, 32 bytes -> clave AES
function derivar(cuit, clave) {
  const bits = crypto.pbkdf2Sync(`${cuit}:${clave.toUpperCase()}`, `ambito-lista:${cuit}`, ITER, 48, 'sha256');
  return { id: bits.subarray(0, 16).toString('hex'), key: bits.subarray(16) };
}
function cifrar(obj, key) {
  const iv = crypto.randomBytes(12);
  const c = crypto.createCipheriv('aes-256-gcm', key, iv);
  const data = Buffer.concat([c.update(JSON.stringify(obj), 'utf8'), c.final(), c.getAuthTag()]);
  return { v: 1, iv: iv.toString('base64'), data: data.toString('base64') };
}

const wb = X.readFile(src);
const ws = wb.Sheets[wb.SheetNames[0]];
const filas = X.utils.sheet_to_json(ws, { defval: '' });
const norm = s => String(s).trim().toUpperCase();

fs.rmSync(outDir, { recursive: true, force: true });
fs.mkdirSync(outDir, { recursive: true });

let ok = 0, nuevas = 0;
const errores = [];
for (const [i, f] of filas.entries()) {
  const k = Object.fromEntries(Object.entries(f).map(([a, b]) => [norm(a), b]));
  const cuit = soloDigitos(k['CUIT']);
  const nombre = String(k['RAZON SOCIAL'] || k['RAZÓN SOCIAL'] || k['CLIENTE'] || k['NOMBRE'] || '').trim();
  if (!cuit && !nombre) continue;
  if (cuit.length !== 11) { errores.push(`Fila ${i + 2}: CUIT inválido (${k['CUIT']})`); continue; }
  let clave = String(k['CLAVE'] || '').trim().toUpperCase();
  if (!clave) { clave = nuevaClave(); f[Object.keys(f).find(c => norm(c) === 'CLAVE') || 'CLAVE'] = clave; nuevas++; }
  const cliente = { nombre, cuit, descuento: pct(k['DESCUENTO']) ?? 0 };
  const { id, key } = derivar(cuit, clave);
  fs.writeFileSync(path.join(outDir, id + '.json'), JSON.stringify(cifrar(cliente, key)));
  ok++;
}

if (nuevas) {
  const cols = Object.keys(filas[0] || {});
  if (!cols.some(c => norm(c) === 'CLAVE')) cols.splice(2, 0, 'CLAVE');
  const nueva = X.utils.json_to_sheet(filas, { header: cols });
  wb.Sheets[wb.SheetNames[0]] = nueva;
  X.writeFile(wb, src);
}
console.log(`Clientes generados: ${ok}${nuevas ? ` (claves nuevas: ${nuevas}, guardadas en el Excel)` : ''}`);
errores.forEach(e => console.log('  ⚠ ' + e));
