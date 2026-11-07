// Pruebas de la lógica pura. Ejecutar con: node tests/ids.test.js
const assert = require('assert');
const ids = require('../js/ids.js');

// Formato de ID
assert.strictEqual(ids.formatId('AA', 1), 'AA001');
assert.strictEqual(ids.formatId('AA', 100), 'AA100');
assert.strictEqual(ids.formatId('AA', 1000), 'AA1000'); // se amplía sin romper

// Rangos consecutivos sin reinicio ni reutilización
const r1 = ids.nextRange(0, 100);
assert.deepStrictEqual([r1.from, r1.to], [1, 100]);
const r2 = ids.nextRange(r1.to, 100);
assert.deepStrictEqual([r2.from, r2.to], [101, 200]);
assert.deepStrictEqual(ids.idsForRange('AA', 1, 3), ['AA001', 'AA002', 'AA003']);

// Entradas inválidas
assert.throws(() => ids.nextRange(0, 0));
assert.throws(() => ids.nextRange(0, 1.5));
assert.throws(() => ids.nextRange(-1, 5));

// Prefijos
assert.ok(ids.validPrefix('AA'));
assert.ok(!ids.validPrefix('aa'));
assert.ok(!ids.validPrefix('ABCDE'));
assert.ok(!ids.validPrefix(''));

// Contraseñas
assert.ok(ids.validSecret('abc123'));
assert.ok(!ids.validSecret(' abc'));
assert.ok(!ids.validSecret(''));

// Inactividad: 5 minutos
const MIN5 = 5 * 60 * 1000;
assert.strictEqual(ids.isIdleExpired(0, MIN5 - 1, MIN5), false);
assert.strictEqual(ids.isIdleExpired(0, MIN5, MIN5), true);

// Pendientes: solo Enviado
const pend = ids.pendingFromRegistro([
  { ID: 'AA001', Estado: 'Impreso' },
  { ID: 'AA002', Estado: 'Enviado' },
  { ID: 'AA003', Estado: 'Reservado' }
]);
assert.deepStrictEqual(pend.map(r => r.ID), ['AA002']);

// Prefijos siguientes
assert.strictEqual(ids.nextPrefix('AA'),'AB');
assert.strictEqual(ids.nextPrefix('AZ'),'BA');
assert.strictEqual(ids.nextPrefix('ZZ'),'AAA');

// Continuidad: 20 impresas, luego 21 sigue; al 999 pasa a AB001
let a=ids.allocate('AA',0,20);
assert.strictEqual(a.ids[0],'AA001'); assert.strictEqual(a.ids[19],'AA020');
let b=ids.allocate('AA',a.last,5);
assert.deepStrictEqual(b.ids,['AA021','AA022','AA023','AA024','AA025']);
let c=ids.allocate('AA',998,3);
assert.deepStrictEqual(c.ids,['AA999','AB001','AB002']);
assert.strictEqual(c.prefix,'AB'); assert.strictEqual(c.last,2);
let d=ids.allocate('AA',999,1);
assert.deepStrictEqual(d.ids,['AB001']);
// allocate no muta: misma entrada, mismo resultado
assert.deepStrictEqual(ids.allocate('AA',0,2).ids,['AA001','AA002']);

// Contadores: Validación por componente; Secuencia por componente y línea
const cons = [
  { Operacion: 'Validación', Componente: 'Llaves', Linea: '', Prefijo: 'FA', Ultimo: 0 },
  { Operacion: 'Secuencia', Componente: 'Llaves', Linea: '011', Prefijo: 'FA', Ultimo: 20 },
  { Operacion: 'Secuencia', Componente: 'Llaves', Linea: '065', Prefijo: 'FA', Ultimo: 0 }
];
// Validación de Llaves no depende de la línea
assert.strictEqual(ids.findCounter(cons,'Validación','Llaves','011'), cons[0]);
assert.strictEqual(ids.findCounter(cons,'Validación','Llaves','065'), cons[0]);
// Secuencia separa por línea
assert.strictEqual(ids.findCounter(cons,'Secuencia','Llaves','011'), cons[1]);
assert.strictEqual(ids.findCounter(cons,'Secuencia','Llaves','065'), cons[2]);
// Secuencia de una línea sin contador todavía -> no existe (se crea al usar)
assert.strictEqual(ids.findCounter(cons,'Secuencia','Llaves','014'), null);
// Secuencia 065 empieza en 001 aunque 011 vaya en 020
assert.deepStrictEqual(ids.allocate('FA', cons[2].Ultimo, 2).ids, ['FA001','FA002']);
// Secuencia 011 sigue en 021
assert.deepStrictEqual(ids.allocate('FA', cons[1].Ultimo, 1).ids, ['FA021']);

console.log('OK: todas las pruebas de ids.js pasaron');
