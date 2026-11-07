// Lógica pura de QORVEX_Label: sin acceso al DOM ni al archivo.
// Se prueba con Node (ver tests/ids.test.js) y se usa desde app.js.
(function (root) {
  'use strict';

  // Formatea un consecutivo como AA001, AA002, ... (mínimo 3 dígitos).
  function formatId(prefix, n) {
    return prefix + String(n).padStart(3, '0');
  }

  // Dado el último consecutivo usado y una cantidad, devuelve el rango nuevo.
  // Nunca reutiliza números: el siguiente siempre es last + 1.
  function nextRange(last, qty) {
    if (!Number.isInteger(last) || last < 0) throw new Error('Consecutivo base inválido');
    if (!Number.isInteger(qty) || qty < 1) throw new Error('La cantidad debe ser un entero mayor a 0');
    const from = last + 1;
    const to = last + qty;
    return { from: from, to: to, qty: qty };
  }

  // Lista de IDs de un rango. Ej: prefix "AA", from 1, to 3 -> AA001, AA002, AA003.
  function idsForRange(prefix, from, to) {
    const out = [];
    for (let n = from; n <= to; n++) out.push(formatId(prefix, n));
    return out;
  }

  // Valida el prefijo configurable: letras mayúsculas, 1 a 4 caracteres.
  function validPrefix(p) {
    return typeof p === 'string' && /^[A-Z]{1,4}$/.test(p);
  }

  // Valida una contraseña de administrador o de usuario (no vacía, sin espacios al inicio o final).
  function validSecret(s) {
    return typeof s === 'string' && s.length > 0 && s === s.trim();
  }

  // Reconciliación al reiniciar: IDs en estado Enviado se marcan como "Pendiente de confirmar",
  // nunca se reutilizan. Devuelve los consecutivos que quedaron pendientes.
  function pendingFromRegistro(rows) {
    return rows.filter(function (r) { return r.Estado === 'Enviado'; });
  }

  // Tiempo de inactividad: true si ya venció el límite (en milisegundos).
  function isIdleExpired(lastActivityMs, nowMs, limitMs) {
    return nowMs - lastActivityMs >= limitMs;
  }


  // Límite de consecutivos por prefijo. Al llegar a 999 se pasa al siguiente prefijo.
  const MAX_PER_PREFIX = 999;

  // Siguiente prefijo: AA -> AB, AZ -> BA, ZZ -> AAA.
  function nextPrefix(p) {
    if (!validPrefix(p)) throw new Error('Prefijo inválido: ' + p);
    const c = p.split('');
    let i = c.length - 1;
    while (i >= 0) {
      if (c[i] !== 'Z') { c[i] = String.fromCharCode(c[i].charCodeAt(0) + 1); return c.join(''); }
      c[i] = 'A';
      i--;
    }
    return 'A' + c.join('');
  }

  // Asigna qty IDs continuando desde 'last' (último número usado con 'prefix').
  // Pura: no modifica nada. Devuelve los IDs y el nuevo estado (prefijo y último número).
  function allocate(prefix, last, qty) {
    if (!validPrefix(prefix)) throw new Error('Prefijo inválido: ' + prefix);
    if (!Number.isInteger(last) || last < 0) throw new Error('Consecutivo base inválido');
    if (!Number.isInteger(qty) || qty < 1) throw new Error('La cantidad debe ser un entero mayor a 0');
    const ids = [];
    let p = prefix, n = last;
    for (let i = 0; i < qty; i++) {
      if (n >= MAX_PER_PREFIX) { p = nextPrefix(p); n = 0; }
      n++;
      ids.push(formatId(p, n));
    }
    return { ids: ids, prefix: p, last: n };
  }


  // Clave de línea del contador: Validación no usa línea; Secuencia sí.
  function lineaKey(op, linea) { return op === 'Secuencia' ? String(linea) : ''; }

  // Busca el contador de una operación, componente y línea.
  function findCounter(rows, op, comp, linea) {
    const L = lineaKey(op, linea);
    return rows.find(function (r) { return r.Operacion === op && r.Componente === comp && r.Linea === L; }) || null;
  }

  const api = {
    formatId: formatId,
    nextRange: nextRange,
    idsForRange: idsForRange,
    validPrefix: validPrefix,
    validSecret: validSecret,
    pendingFromRegistro: pendingFromRegistro,
    isIdleExpired: isIdleExpired,
    nextPrefix: nextPrefix,
    allocate: allocate,
    MAX_PER_PREFIX: MAX_PER_PREFIX,
    lineaKey: lineaKey,
    findCounter: findCounter
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.QorvexIds = api;
})(typeof window !== 'undefined' ? window : globalThis);
