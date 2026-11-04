'use strict';

/* =========================================================
   SISFIS_VANESA – Informes de supervisión
   Formulario + vista previa idéntica al Word modelo
   ========================================================= */

const $ = id => document.getElementById(id);

// Todos los campos de texto del formulario (ids del HTML)
const CAMPOS = [
  'expediente', 'empresa', 'placa', 'registro', 'supervisor',
  'supervisorCargo', 'empresaSupervisora', 'numeroInforme', 'ciudad', 'fecha',
  'destinatario', 'copias', 'asunto', 'referencia',
  'objetivo', 'baseLegal',
  'descripcionGeneral', 'imagenTitulo', 'consultaRegistro', 'direccion',
  'analisis', 'fuente', 'cierreVisita',
  'anexoIntro', 'anexos',
  'descargos', 'conclusiones', 'recomendaciones',
  'firma1Nombre', 'firma1Documento', 'firma1Cargo', 'firma1Empresa',
  'firma2Nombre', 'firma2Documento', 'firma2Cargo', 'firma2Empresa'
];

// Datos variables de "2. ANTECEDENTES" (bloques 2.1 a 2.6 de la tarjeta Antecedentes).
// Cada uno es un marcador {{campo}}: si se cambia, cambia en todo el informe.
const CAMPOS_ANT_TEXTO = [
  'capacidad', 'compartimientos', 'direccionLegal', 'distrito', 'tracto',  // 2.1
  'resolucionNumero',                                                      // 2.2
  'lugarFiscalizacion',                                                    // 2.3
  'especialista', 'asignacionNumero', 'cartaLinea',                        // 2.4
  'fiscalizador'                                                           // 2.5
];
// Fechas (input type="date"): se escriben en formato largo (24 de septiembre de 2026).
// {{fecha}} es la "Fecha del informe" de Datos generales (2.6).
const CAMPOS_ANT_FECHA = ['fechaRegistro', 'resolucionFecha', 'cartaFecha', 'asignacionFecha', 'fechaVisita'];
const CAMPOS_ANT = [...CAMPOS_ANT_TEXTO, ...CAMPOS_ANT_FECHA];
CAMPOS.push(...CAMPOS_ANT);

// Campos que se pueden usar como marcador {{campo}} dentro de los textos
const MARCADORES = ['expediente', 'empresa', 'placa', 'registro'];
const MARCADORES_TEXTO = [...MARCADORES, ...CAMPOS_ANT_TEXTO];
const MARCADORES_FECHA = ['fecha', ...CAMPOS_ANT_FECHA];

// Campos donde se buscan/reemplazan los marcadores
const CAMPOS_CON_MARCADOR = CAMPOS.filter(
  k => !MARCADORES.includes(k) && k !== 'fecha' && !CAMPOS_ANT.includes(k)
);

const CAMPOS_HECHO = ['numero', 'incumplimiento', 'baseLegal', 'descripcion', 'medio'];

let editId = null;

/* ---------------------------------------------------------
   Logo de la empresa supervisora (public/logo.png, 199 x 130 px)
   Se usa tal cual. Ancho en el informe: 110 pt (como en el modelo).
   --------------------------------------------------------- */
const LOGO = {
  urlWeb: new URL('logo.png', location.href).href,  // vista previa e impresión
  rutaWord: 'informe_archivos/logo.png',            // ruta interna dentro del .doc
  anchoPt: 110,
  anchoPx: 147,   // 110 pt a 96 ppp, proporción original 199:130
  altoPx: 96,
  base64: null    // contenido del PNG para incrustarlo en el Word
};

// Lee logo.png una sola vez y lo guarda en base64 para incrustarlo en el Word
const logoListo = fetch('logo.png')
  .then(r => (r.ok ? r.arrayBuffer() : Promise.reject(new Error('logo.png no encontrado'))))
  .then(buf => { LOGO.base64 = bytesABase64(new Uint8Array(buf)); })
  .catch(err => console.warn('No se pudo cargar el logo:', err.message));

function bytesABase64(bytes) {
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(bin);
}

// Base64 partido en líneas de 76 caracteres (formato MIME)
function base64MIME(b64) {
  return b64.replace(/.{1,76}/g, '$&\r\n');
}

/* ---------------------------------------------------------
   Imagen de "4. DESCRIPCIÓN GENERAL" (consulta del Registro)
   Se guarda en informes.json como base64 (data URL) junto con el informe.
   --------------------------------------------------------- */
const TITULO_IMAGEN = 'IMAGEN N° 01 CONSULTA EN LÍNEA DEL REGISTRO DE HIDROCARBUROS';
const IMG_MAX_ANCHO = 1000;   // px: se reduce con canvas antes de guardar
const IMG_WORD_ANCHO = 560;   // px: ancho del recuadro en Word (420 pt, como el modelo)

// Ruta interna de la imagen dentro del .doc (extensión según su tipo)
function rutaImagenWord(dataUrl) {
  const p = partesDataURL(dataUrl);
  const ext = p ? (p.tipo.split('/')[1] || 'png').replace('jpeg', 'jpg').replace(/\+.*/, '') : 'png';
  return `informe_archivos/imagen01.${ext}`;
}

const IMAGEN = { dataUrl: null, w: 0, h: 0 };
let imagenEjemplo = null;     // data URL de public/consulta-ejemplo.png

function leerComoDataURL(blob) {
  return new Promise((ok, mal) => {
    const fr = new FileReader();
    fr.onload = () => ok(fr.result);
    fr.onerror = () => mal(fr.error);
    fr.readAsDataURL(blob);
  });
}

function cargarImagen(src) {
  return new Promise((ok, mal) => {
    const img = new Image();
    img.onload = () => ok(img);
    img.onerror = () => mal(new Error('No se pudo leer la imagen'));
    img.src = src;
  });
}

// Reduce la imagen a máximo 1000 px de ancho. PNG para capturas (texto nítido);
// si queda muy pesada, se pasa a JPEG.
async function reducirImagen(archivo) {
  const original = await leerComoDataURL(archivo);
  const img = await cargarImagen(original);
  const escala = Math.min(1, IMG_MAX_ANCHO / img.naturalWidth);
  const w = Math.round(img.naturalWidth * escala);
  const h = Math.round(img.naturalHeight * escala);

  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, w, h);
  ctx.drawImage(img, 0, 0, w, h);

  let dataUrl = canvas.toDataURL('image/png');
  if (dataUrl.length > 1500000) dataUrl = canvas.toDataURL('image/jpeg', 0.85);
  return dataUrl;
}

// Fija la imagen actual del informe (o null para quitarla)
async function ponerImagen(dataUrl) {
  IMAGEN.dataUrl = dataUrl || null;
  IMAGEN.w = 0;
  IMAGEN.h = 0;

  if (IMAGEN.dataUrl) {
    try {
      const img = await cargarImagen(IMAGEN.dataUrl);
      if (IMAGEN.dataUrl !== dataUrl) return;   // se cambió mientras cargaba
      IMAGEN.w = img.naturalWidth;
      IMAGEN.h = img.naturalHeight;
    } catch (err) {
      IMAGEN.dataUrl = null;
    }
  }
  mostrarMiniatura();
  programarVista();
}

function mostrarMiniatura() {
  const hay = !!IMAGEN.dataUrl;
  $('imagenMiniatura').hidden = !hay;
  $('btnQuitarImagen').hidden = !hay;
  if (hay) $('imagenMiniatura').src = IMAGEN.dataUrl;
  else $('imagenMiniatura').removeAttribute('src');
  $('imagenInfo').textContent = hay
    ? `Imagen cargada (${IMAGEN.w || '?'} × ${IMAGEN.h || '?'} px). Puede reemplazarla subiendo otra.`
    : 'Sin imagen: el recuadro no se mostrará en el informe.';
}

// Tamaño de la imagen para Word (atributos width/height en px)
function tamanoWord() {
  if (!IMAGEN.w || !IMAGEN.h) return { w: IMG_WORD_ANCHO, h: Math.round(IMG_WORD_ANCHO * 0.45) };
  const w = Math.min(IMG_WORD_ANCHO, IMAGEN.w);
  return { w, h: Math.round(IMAGEN.h * w / IMAGEN.w) };
}

// Separa un data URL en tipo y base64 (para incrustarlo en el Word)
function partesDataURL(dataUrl) {
  const m = /^data:(image\/[\w.+-]+);base64,(.*)$/.exec(dataUrl || '');
  return m ? { tipo: m[1], base64: m[2] } : null;
}

const imagenEjemploLista = fetch('consulta-ejemplo.png')
  .then(r => (r.ok ? r.blob() : Promise.reject(new Error('consulta-ejemplo.png no encontrado'))))
  .then(leerComoDataURL)
  .then(d => { imagenEjemplo = d; })
  .catch(err => console.warn('No se pudo cargar la imagen de ejemplo:', err.message));

/* ---------------------------------------------------------
   Plantilla fija de "2. ANTECEDENTES" (una línea por punto: 2.1 a 2.6).
   Los marcadores {{ }} solo se usan aquí, por dentro: el formulario muestra
   un campo por dato y la vista previa / Word arman el texto completo.
   --------------------------------------------------------- */
const PLANTILLA_ANTECEDENTES = [
  'Con fecha {{fechaRegistro}}, {{empresa}} (en adelante, el Agente Fiscalizado) obtuvo el Registro de Hidrocarburos N.° {{registro}}, para operar la unidad vehicular de placa principal N.° {{placa}} y tracto {{tracto}}, en la actividad de transporte de combustibles líquidos y OPDH, con capacidad autorizada de {{capacidad}} galones y {{compartimientos}} compartimientos, con dirección legal en {{direccionLegal}}, distrito de {{distrito}}, provincia y departamento de HUÁNUCO, conforme a la consulta del Registro de Hidrocarburos.',
  'Mediante la Resolución de Autoridad de Fiscalización – Oficina Regional Huánuco N.° {{resolucionNumero}}, de fecha {{resolucionFecha}}, se dispuso, en calidad de MANDATO, que {{empresa}} ponga a disposición la unidad de placa principal N.° {{placa}}, con Registro de Hidrocarburos N.° {{registro}}.',
  'Con fecha {{cartaFecha}}, mediante carta dirigida al Especialista Regional en Hidrocarburos de la Oficina Regional Huánuco, {{empresa}} comunicó que la fiscalización de la unidad {{placa}} se realizaría en {{lugarFiscalizacion}}.',
  'Con fecha {{asignacionFecha}}, el Especialista Regional en Hidrocarburos, Ing. {{especialista}}, mediante Asignación N.° {{asignacionNumero}}, asignó la Carta Línea N.° {{cartaLinea}} a la empresa supervisora SUPERVISIÓN TÉCNICA DEMOSTRATIVA E.I.R.L., a fin de efectuar la fiscalización operativa a medios de transporte terrestre de CL, con visita, a la unidad vehicular de placa {{placa}}.',
  'Con fecha {{fechaVisita}}, el Ing. {{fiscalizador}} realizó la visita de fiscalización al medio de transporte de placa principal N.° {{placa}}, operado por {{empresa}}, y suscribió el Acta de Fiscalización correspondiente al Expediente N.° {{expediente}}.',
  'Con fecha {{fecha}}, se elabora el presente informe de fiscalización, correspondiente a la Carta Línea N.° {{cartaLinea}} y al Expediente N.° {{expediente}}, para su evaluación.'
].join('\n');

/* ---------------------------------------------------------
   Texto modelo (datos ficticios, igual que el Word modelo)
   --------------------------------------------------------- */
const MODELO = {
  expediente: '202600000000',
  empresa: 'TRANSPORTES MODELO DEMOSTRATIVO S.A.C.',
  placa: 'ABC-000',
  registro: '00000-060-000000',
  supervisor: 'Carlos Ejemplo Rivera',
  supervisorCargo: 'Supervisor ES',
  empresaSupervisora: 'SUPERVISIÓN TÉCNICA DEMOSTRATIVA E.I.R.L.',
  numeroInforme: '0001-2026',
  ciudad: 'Huánuco',
  fecha: '2026-10-25',
  destinatario: 'Oficina Regional Huánuco - OSINERGMIN',
  copias: '',
  asunto: 'Fiscalización de SEGURIDAD DE MEDIOS DE TRANSPORTE DE COMBUSTIBLE LIQUIDOS Y OPDH, operado por {{empresa}}',
  referencia: 'Registro de Hidrocarburos de OSINERGMIN N° {{registro}}',

  // Datos de los antecedentes (valores ficticios por defecto)
  fechaRegistro: '2024-08-29',
  capacidad: '9000',
  compartimientos: 'seis',
  direccionLegal: 'DIRECCIÓN FICTICIA PARA CAPACITACIÓN, HUÁNUCO',
  distrito: 'CHURUBAMBA',
  tracto: 'XYZ-000',
  resolucionNumero: '000-2026-OS-GSE/DSR-OR HUÁNUCO',
  resolucionFecha: '2026-09-10',
  cartaFecha: '2026-09-11',
  lugarFiscalizacion: 'la Carretera Central Huánuco–Tingo María, Km 23+270, Centro Poblado de Chullqui, distrito de Churubamba, Huánuco',
  asignacionFecha: '2026-09-18',
  especialista: 'Luis Modelo Fernández',
  asignacionNumero: '00000-2026',
  cartaLinea: '0000000-0',
  fechaVisita: '2026-09-24',
  fiscalizador: 'Carlos Ejemplo Rivera',

  objetivo: 'El objetivo de la presente visita de fiscalización es de verificar el cumplimiento de las obligaciones relativas a las condiciones de seguridad del MEDIO DE TRANSPORTE DE COMBUSTIBLE LIQUIDOS Y OPDH de placa de rodaje principal Nº {{placa}}, con Registro de Hidrocarburos N° {{registro}}, operado por {{empresa}}, con Registro Único de Contribuyentes (R.U.C.) N° 20000000001 (en adelante, el Agente Fiscalizado).',

  baseLegal: [
    'Ley Nº 26734, Ley de Creación de Osinergmin.',
    'Ley Nº 27699, Ley Complementaria de Fortalecimiento Institucional de Osinergmin.',
    'Decreto Supremo Nº 054-2001-PCM, que aprueba el Reglamento General del Osinergmin.',
    'Decreto Supremo N° 006-2026-JUS, que aprueba el Texto Único Ordenado de la Ley N.º 27444, Ley del Procedimiento Administrativo General.',
    'Resolución de Consejo Directivo Nº 150-2024-OS/CD, que aprueba el Reglamento del Registro de Hidrocarburos de Osinergmin. Y la Resolución de Consejo Directivo Nº 56-2026-OS/CD, que modifica el Reglamento del Registro de Hidrocarburos de Osinergmin.',
    'Resolución de Consejo Directivo Nº 208-2020-OS/CD, que aprueba el Reglamento de Fiscalización y Sanción de las actividades energéticas y mineras a cargo de Osinergmin.',
    'Resolución de Consejo Directivo N° 271-2012-OS/CD, que aprueba la Tipificación y Escala de Multas y Sanciones de Hidrocarburos de OSINERGMIN y modificatorias.',
    'Decreto Supremo N° 026-94-EM y sus modificatorias, que aprueba el Reglamento.',
    'Decreto Supremo Nº 030-98-EM y sus modificatorias, que aprueba el Reglamento.',
    'Decreto Supremo Nº 019-99-EM y sus modificatorias, que amplía el plazo para que camiones cisterna de transporte de combustibles no equipados para carga inferior, adecuen sus instalaciones.',
    'Norma Técnica Peruana, NTP: 350.043-1, Extintores Portátiles.'
  ].join('\n'),

  descripcionGeneral: 'El medio de transporte terrestre de combustibles líquidos y OPDH, operado por {{empresa}}; cuenta con la siguiente información:',
  imagenTitulo: TITULO_IMAGEN,
  consultaRegistro: 'Consulta del Registro de Hidrocarburos de fecha 22 de octubre de 2026: Registro N.° {{registro}}, titular {{empresa}}, placa principal {{placa}}, tracto {{tracto}}, capacidad de {{capacidad}} galones y {{compartimientos}} compartimientos.',
  direccion: 'La fiscalización realizada al medio de transporte terrestre de combustibles líquidos y OPDH, de placa principal {{placa}}, se llevó a cabo en DIRECCIÓN FICTICIA PARA CAPACITACIÓN, HUÁNUCO, distrito de CHURUBAMBA, provincia y departamento de HUÁNUCO.',

  analisis: 'Con fecha {{fechaVisita}}, siendo las 10:16 horas, se inició la fiscalización a la unidad vehicular de placa principal {{placa}}, operada por {{empresa}}, con Registro de Hidrocarburos N.° {{registro}}. Se constataron las observaciones consignadas en el apartado II «Otros» del acta, conforme al siguiente cuadro:',
  fuente: 'Acta de Fiscalización de Seguridad de Medios de Transporte de Combustibles Líquidos y OPDH N° {{expediente}}.',
  cierreVisita: 'Culminada la fiscalización, se suscribió el ACTA DE FISCALIZACIÓN DE SEGURIDAD DE MEDIOS DE TRANSPORTE DE COMBUSTIBLES LÍQUIDOS Y OPDH correspondiente al Expediente N.° {{expediente}}, recibida por el Sr. PEDRO MODELO GARCIA, identificado con DNI N.° 00000000, en calidad de encargado de la unidad vehicular. La diligencia concluyó a las 15:20 horas.',

  anexoIntro: 'De la visita realizada a la unidad vehicular se recabó la documentación señalada en el Acta de Fiscalización.',
  anexos: [
    'ANEXO 1. REPORTE DE CIERRE SFH - CL {{cartaLinea}}',
    'ANEXO 2. {{expediente}}-060 - PLAN DE CONTINGENCIAS',
    'ANEXO 3. {{expediente}}-060 - GUÍA DE REMISIÓN',
    'ANEXO 4. {{expediente}}-060 - CAPACITACIÓN DEL CONDUCTOR',
    'ANEXO 5. {{expediente}}-060 - CONSULTA RHO',
    'ANEXO 6. {{expediente}}-060 - CONSULTA RUC',
    'ANEXO 7. {{expediente}}-060 - PÓLIZA DE SEGURO',
    'ANEXO 8. {{expediente}}-060 – CONSULTA EN LINEA RHO',
    'ANEXO 9. {{expediente}}-060 - CERTIFICADO DE INSTALACIÓN GPS',
    'ANEXO 10. {{expediente}}-060 - CERTIFICADO DE VERIFICACIÓN',
    'ANEXO 11. {{expediente}}-060 - TARJETA E INSPECCIÓN TÉCNICA {{placa}}',
    'ANEXO 12. {{expediente}}-060 - TARJETA E INSPECCIÓN TÉCNICA {{tracto}}',
    'VIDEO DE OPERATIVIDAD DE LUCES'
  ].join('\n'),

  descargos: '',
  conclusiones: '{{empresa}}, con Ficha de Registro de Hidrocarburos N.° {{registro}}, NO CUMPLIÓ con lo establecido de las condiciones de seguridad de Medio De Transporte De Combustible Líquidos Y OPDH.',
  recomendaciones: [
    'Comunicar a {{empresa}}, con Ficha de Registro de Hidrocarburos N.° {{registro}}, las observaciones consignadas en el apartado II «Otros» del Acta de Fiscalización, a fin de que acredite su subsanación respecto de la unidad vehicular de placa principal {{placa}}.',
    'Asimismo, se otorgue el plazo establecido de 30 días calendario para que la empresa pueda subsanar los incumplimientos de orden técnico y que no constituyen un alto riesgo para la seguridad (TIPO B), bajo apercibimiento de iniciar Procedimiento Administrativo Sancionador – PAS^[El plazo se contabiliza a partir del día siguiente de habérsele trasladado al administrado el oficio de notificación del informe de fiscalización, vía SNE, en el que se deja constancia de los incumplimientos técnicos verificados. Versión 1.0 del Lineamientos de actuación en la fiscalización de seguridad de Medios de Transporte de Combustibles Líquidos y OPDH a través de camiones tanque y camiones cisterna, División De Supervisión Regional Gerencia de Supervisión de Energía.].'
  ].join('\n'),

  firma1Nombre: 'Elena Referencia Torres',
  firma1Documento: 'DNI N° 00000000',
  firma1Cargo: 'Representante Legal',
  firma1Empresa: 'SUPERVISIÓN TÉCNICA DEMOSTRATIVA E.I.R.L.',
  firma2Nombre: 'Carlos Ejemplo Rivera',
  firma2Documento: 'CIP N° 000000',
  firma2Cargo: 'Fiscalizador',
  firma2Empresa: 'SUPERVISIÓN TÉCNICA DEMOSTRATIVA E.I.R.L.',

  tieneDescargo: false,

  hechos: [
    {
      numero: 'OTRAS OBSERVACIONES',
      incumplimiento: 'Cada compartimiento del tanque de cargamento deberá ser provisto con venteos normales de presión y vacío, cada uno teniendo un área de sección transversal libre neta de 0.44 pulg2 (284 mm2). Venteos de presión deberán ser fijados para abrir en una presión manométrica de no más de 1 psi (6.9 kPa). Venteos de vacío deberán ser fijados para abrir en 6 oz (170 g). Donde un compartimiento del tanque de cargamento es diseñado para ser cargado o descargado con la cubierta del domo cerrada, el venteo o venteos descritos en 5.3.12.2 o venteos adicionales deberán limitar el vacío a 1 psi (6.9 kPa) y la presión del tanque a una presión manométrica de 3 psi (20.7 kPa) basada en la tasa máxima de transferencia de producto.',
      baseLegal: '5.3.12.2 y 5.3.12.4 del Estándar NFPA 385 Estándar para vehículos tanque para líquidos inflamables y combustibles. Edición 2007, en concordancia con la cuarta disposición transitoria de Texto Único Ordenado de la Ley Orgánica de Hidrocarburos, aprobada mediante Decreto Supremo N° 042-2005-EM.',
      descripcion: 'Tipo B: Durante la visita de fiscalización se verificó que las válvulas de presión y vacío de los compartimientos N.° 1, 2, 3, 4, 5 y 6 no especifican que limiten el vacío a 1 PSI (6.9 kPa) y la presión manométrica a 3 PSI (20.7 kPa).',
      medio: 'Foto N.º 23, 30, 39, 47, 54 y 61.'
    },
    {
      numero: 'OTRAS OBSERVACIONES',
      incumplimiento: '**Venteo de Emergencia para Exposición a Incendio:**\nCada compartimiento del tanque de cargamento deberá ser equipado con uno o más venteos accionados por presión fijados para abrir en una presión manométrica de no menos de 3 psi (20.7 kPa) y cerrar cuando las caídas de presión debajo de 3 psi (20.7 kPa) del punto de seteo.',
      baseLegal: '5.3.13 del Estándar NFPA 385 Estándar para vehículos tanque para líquidos inflamables y combustibles. Edición 2007, en concordancia con la cuarta disposición transitoria de Texto Único Ordenado de la Ley Orgánica de Hidrocarburos, aprobada mediante Decreto Supremo N° 042-2005-EM.',
      descripcion: 'Tipo B: Durante la visita de fiscalización se verificó que el Agente Fiscalizado deberá acreditar que los compartimientos N.° 1, 2, 3, 4 y 6 cuentan con venteos de emergencia para exposición a incendios, calibrados para abrir a no menos de 3 PSI y cerrar por debajo de 3 PSI del punto de venteo.',
      medio: 'Foto N.º 22, 29, 38, 46, 53 y 60.'
    }
  ]
};

/* ---------------------------------------------------------
   Utilidades de texto
   --------------------------------------------------------- */
function esc(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Reemplaza {{expediente}}, {{empresa}}, {{placa}}, {{registro}} y los datos de los
// antecedentes por su valor actual. Las fechas salen en formato largo.
function reemplazar(texto, x) {
  return String(texto ?? '').replace(/\{\{\s*(\w+)\s*\}\}/g, (m, k) => {
    if (MARCADORES_FECHA.includes(k)) return fechaLarga(x[k]);
    if (!MARCADORES_TEXTO.includes(k)) return m;
    const v = (x[k] || '').trim();
    return v || '________';
  });
}

// Texto listo para el informe: marcadores reemplazados, escapado y con **negrita**
// ^[texto] = nota al pie (se numera sola: ¹, ², ...)
function txt(texto, x) {
  const formato = s => esc(s).replace(/\*\*(.+?)\*\*/g, '<b>$1</b>');
  const crudo = reemplazar(texto, x);
  let salida = '';
  let desde = 0;
  for (const m of crudo.matchAll(/\^\[([^\]]+)\]/g)) {
    salida += formato(crudo.slice(desde, m.index)) + refNota(m[1].trim());
    desde = m.index + m[0].length;
  }
  return salida + formato(crudo.slice(desde));
}

function lineas(texto) {
  return String(texto ?? '')
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(Boolean);
}

const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio',
  'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

function fechaLarga(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
  if (!m) return '________';
  return `${Number(m[3])} de ${MESES[Number(m[2]) - 1]} de ${m[1]}`;
}

function romano(n) {
  const t = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
    [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let r = '';
  for (const [v, s] of t) while (n >= v) { r += s; n -= v; }
  return r;
}

// Convierte en marcador los valores escritos a mano dentro de los textos.
// Ej.: si el expediente es 202600000000 y aparece escrito en un texto,
// se cambia por {{expediente}} para que luego se actualice solo.
function tokenizar(x) {
  const pares = MARCADORES
    .map(k => [k, String(x[k] || '').trim()])
    .filter(([, v]) => v.length >= 6)
    .sort((a, b) => b[1].length - a[1].length);

  if (!pares.length) return x;

  const tk = s => {
    if (typeof s !== 'string') return s;
    for (const [k, v] of pares) s = s.split(v).join('{{' + k + '}}');
    return s;
  };

  const y = { ...x };
  CAMPOS_CON_MARCADOR.forEach(k => { y[k] = tk(y[k]); });
  y.hechos = (x.hechos || []).map(h => {
    const f = {};
    CAMPOS_HECHO.forEach(k => { f[k] = tk(h[k] || ''); });
    return f;
  });
  return y;
}

/* ---------------------------------------------------------
   Estilos del documento (vista previa, Word e impresión)
   Igual al Word modelo: Poppins 9,5 pt, interlineado 1,15,
   A4 con márgenes 3,5 / 2 / 1,5 / 2,5 cm y encabezado a 0,5 cm.
   --------------------------------------------------------- */
const FUENTE = 'Poppins, Arial, sans-serif';

const DOC_CSS = `
.doc{font-family:${FUENTE};font-size:9.5pt;line-height:1.43;color:#000;background:#fff;}
.doc p{margin:0;text-align:justify;font-family:${FUENTE};font-size:9.5pt;}
.doc .num{display:inline-block;text-indent:0;}
.doc table{border-collapse:collapse;}
.doc td,.doc th{font-family:${FUENTE};}
.doc sup{font-size:65%;line-height:0;}
.doc .notas{margin-top:24pt;}
.doc .notas hr{width:33%;margin:0 0 3pt 0;border:none;border-top:0.75pt solid #000;}
.doc .notas p{font-size:7pt;line-height:1.3;}
`;

// Ajustes propios de Word (interlineado "múltiple 1,15" y notas al pie)
const WORD_CSS = `
p,li,div{margin:0;font-family:${FUENTE};font-size:9.5pt;line-height:115%;text-align:justify;}
table{border-collapse:collapse;}
td,th{font-family:${FUENTE};}
span.MsoFootnoteReference{vertical-align:super;}
p.MsoFootnoteText{margin:0;font-size:7pt;line-height:normal;text-align:justify;font-family:${FUENTE};}
`;

const B = 'border:0.75pt solid #000;';
const CELDA = `${B}padding:2pt 4pt;vertical-align:middle;font-size:7.5pt;line-height:1.4;`;
const CAB = `${B}padding:2pt 4pt;vertical-align:middle;font-family:${FUENTE};font-size:7.5pt;line-height:1.3;background:#0070C0;color:#fff;font-weight:bold;text-align:center;`;
const PC = 'margin:0;font-size:7.5pt;line-height:1.4;text-align:justify;';
const ROJO = '#EE0000';

/* ---------------------------------------------------------
   Notas al pie: escriba ^[texto de la nota] dentro de cualquier texto.
   En Word salen como notas al pie reales; en la vista previa, al final.
   --------------------------------------------------------- */
let NOTAS = null;        // lista de notas del informe que se está armando
let MODO_NOTAS = 'web';

function refNota(texto) {
  if (!NOTAS) return '';
  NOTAS.push(texto);
  const k = NOTAS.length;
  return MODO_NOTAS === 'word'
    ? `<a style="mso-footnote-id:ftn${k}" href="#_ftn${k}" name="_ftnref${k}" title=""><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a>`
    : `<sup>${k}</sup>`;
}

function listaNotas() {
  if (!NOTAS || !NOTAS.length) return '';
  if (MODO_NOTAS === 'word') {
    return `<div style="mso-element:footnote-list"><![if !supportFootnotes]><br clear="all"><hr align="left" size="1" width="33%"><![endif]>`
      + NOTAS.map((t, i) => `<div style="mso-element:footnote" id="ftn${i + 1}"><p class="MsoFootnoteText"><a style="mso-footnote-id:ftn${i + 1}" href="#_ftnref${i + 1}" name="_ftn${i + 1}" title=""><span class="MsoFootnoteReference"><span style="mso-special-character:footnote"></span></span></a> ${esc(t)}</p></div>`).join('')
      + `</div>`;
  }
  return `<div class="notas"><hr>`
    + NOTAS.map((t, i) => `<p><sup>${i + 1}</sup> ${esc(t)}</p>`).join('')
    + `</div>`;
}

/* ---------------------------------------------------------
   Encabezado de página (logo a la izquierda, empresa a la derecha
   y una línea debajo). En Word se repite en todas las páginas.
   --------------------------------------------------------- */
function encabezadoHTML(x, modo) {
  const logo = modo === 'word'
    ? (LOGO.base64 ? `<img src="logo.png" width="${LOGO.anchoPx}" height="${LOGO.altoPx}" alt="Logo">` : '&nbsp;')
    : `<img src="${LOGO.urlWeb}" alt="Logo" style="width:${LOGO.anchoPt}pt;height:auto;display:block;">`;

  return `<table style="width:100%;border:none;border-collapse:collapse;"><tr>
      <td valign="middle" style="width:47%;border:none;padding:0;vertical-align:middle;">${logo}</td>
      <td valign="middle" style="width:53%;border:none;padding:0;vertical-align:middle;">
        <p style="margin:0;text-align:right;font-family:${FUENTE};font-size:10pt;font-weight:bold;">${txt(x.empresaSupervisora, x) || '&nbsp;'}</p>
      </td>
    </tr></table>
    <p style="margin:0;font-size:2pt;line-height:normal;border:none;border-bottom:solid windowtext 0.75pt;padding:0 0 1pt 0;">&nbsp;</p>`;
}

// Archivo de encabezado para Word (va dentro del .doc, junto al logo)
function encabezadoWord(x) {
  return `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><link id="Main-File" rel="Main-File" href="../informe.htm"></head>
<body lang="ES"><div style="mso-element:header" id="h1">${encabezadoHTML(x, 'word')}</div></body></html>`;
}

/* ---------------------------------------------------------
   Armado del informe (misma estructura que el Word modelo)
   modo: 'web' (vista previa / imprimir) o 'word' (descarga)
   --------------------------------------------------------- */
function informeHTML(x, modo = 'web') {
  NOTAS = [];
  MODO_NOTAS = modo === 'word' ? 'word' : 'web';
  const T = s => txt(s, x);
  const word = modo === 'word';

  // Párrafo con número en sangría francesa (número en negrita, como el modelo)
  const np = (num, html, ml, h, extra = '') => {
    const estilo = `margin-left:${ml}pt;text-indent:-${h}pt;${extra}`;
    return word
      ? `<p style="${estilo}tab-stops:${ml}pt ${ml + 5}pt ${ml + 10}pt ${ml + 15}pt;">${num}<span style="mso-tab-count:1">&nbsp;&nbsp;&nbsp;</span>${html}</p>`
      : `<p style="${estilo}"><span class="num" style="min-width:${h}pt">${num}</span>${html}</p>`;
  };

  const parrafo = (html, ml, extra = '') =>
    `<p style="margin-left:${ml}pt;${extra}">${html}</p>`;

  const parrafos = (texto, ml, extra = '') =>
    lineas(texto).map(l => parrafo(T(l), ml, extra)).join('');

  let n = 0;
  const titulo = t => np(`<b>${++n}.</b>`, `<b><u>${t}:</u></b>`, 14.2, 14.2, 'margin-top:16pt;margin-bottom:8pt;');
  const items = (texto, ml = 36, h = 21.8, negrita = false, extra = '') =>
    lineas(texto).map((l, i) =>
      np(`<b>${n}.${i + 1}.</b>`, negrita ? `<b>${T(l)}</b>` : T(l), ml, h, `margin-top:6pt;${extra}`)).join('');

  let h = '';

  /* ---------- Encabezado de página (solo en la vista previa;
     en Word va en el encabezado real y al imprimir se repite) ---------- */
  if (modo === 'web') h += `<div style="margin-bottom:22pt;">${encabezadoHTML(x, 'web')}</div>`;

  /* ---------- 1. Título, N°, ciudad/fecha/expediente y "Otros Destinatarios" ---------- */
  // Fecha a la izquierda y expediente a la derecha, en su propia línea
  const fechaLinea = `<table style="width:100%;border:none;border-collapse:collapse;"><tr>
      <td valign="bottom" style="border:none;padding:0;vertical-align:bottom;">
        <p style="margin:0;text-align:left;">${T(x.ciudad) || '________'}, ${fechaLarga(x.fecha)}</p>
      </td>
      <td valign="bottom" style="border:none;padding:0 6pt 0 0;vertical-align:bottom;">
        <p style="margin:0;text-align:right;white-space:nowrap;">Expediente: ${T(x.expediente) || '________'}</p>
      </td>
    </tr></table>`;

  // Título centrado en la página (el recuadro "Otros Destinatarios" queda a la derecha)
  h += `<table style="width:100%;border:none;border-collapse:collapse;border-bottom:solid windowtext 0.75pt;"><tr>
    <td valign="bottom" style="width:330pt;border:none;border-bottom:solid windowtext 0.75pt;padding:0 0 2pt 0;vertical-align:bottom;">
      <p style="margin:10pt 0 0 138pt;text-align:center;font-size:12pt;font-weight:bold;line-height:1.3;">INFORME DE SUPERVISIÓN</p>
      <p style="margin:0 0 8pt 138pt;text-align:center;font-size:12pt;font-weight:bold;line-height:1.3;">N° ${T(x.numeroInforme) || '________'}</p>
      ${fechaLinea}
    </td>
    <td valign="top" style="border:none;border-left:solid windowtext 0.75pt;border-bottom:solid windowtext 0.75pt;padding:2pt 0 0 4pt;vertical-align:top;">
      <p style="margin:0;text-align:left;font-size:8pt;line-height:1.3;">Otros Destinatarios</p>
      <p style="margin:0;text-align:left;font-size:8pt;line-height:1.3;">C/c: ${T(x.copias)}</p>
    </td>
  </tr></table>`;

  /* ---------- 2. A / De / Asunto / Referencia ---------- */
  // Tabla SIN bordes de 3 columnas: etiqueta (2,5 cm) · dos puntos (0,5 cm) · texto.
  // Así los ":" quedan uno debajo del otro y el texto, aunque ocupe varias
  // líneas, sigue alineado en la columna 3. Raya fina bajo Referencia (como el modelo).
  const ANCHO_ETQ = 70.9;   // 2,5 cm
  const ANCHO_DOS = 14.2;   // 0,5 cm
  const pBloque = (html, alin = 'left') => `<p style="margin:0;text-align:${alin};">${html}</p>`;

  const filaBloque = (etiqueta, valor, { ultima = false } = {}) => {
    const pad = ultima ? 'padding:0 0 2pt 0;' : 'padding:0 0 12pt 0;';
    const td = (ancho, html) =>
      `<td valign="top" style="${ancho ? `width:${ancho}pt;` : ''}border:none;${pad}vertical-align:top;">${html}</td>`;
    return `<tr>${td(ANCHO_ETQ, pBloque(`<b>${etiqueta}</b>`))}${td(ANCHO_DOS, pBloque('<b>:</b>'))}${td(0, valor)}</tr>`;
  };

  const cargoSup = [T(x.supervisorCargo), T(x.empresaSupervisora)].filter(Boolean).join(' ');

  // Espacio sobre el bloque (Word ignora el margin de las tablas)
  h += `<p style="margin:0;line-height:normal;font-size:9pt;">&nbsp;</p>`;
  h += `<table style="width:100%;border:none;border-collapse:collapse;">
    <col style="width:${ANCHO_ETQ}pt"><col style="width:${ANCHO_DOS}pt"><col>
    ${filaBloque('A', pBloque(T(x.destinatario) || '&nbsp;'))}
    ${filaBloque('De', pBloque(`<b>${T(x.supervisor) || '&nbsp;'}</b>`) + (cargoSup ? pBloque(cargoSup) : ''))}
    ${filaBloque('Asunto', pBloque(T(x.asunto) || '&nbsp;', 'justify'))}
    ${filaBloque('Referencia', pBloque(T(x.referencia) || '&nbsp;'), { ultima: true })}
  </table>`;
  // Raya fina bajo Referencia (como en el Word modelo)
  h += `<p style="margin:0;font-size:2pt;line-height:normal;border:none;border-bottom:solid windowtext 0.5pt;padding:0 0 1pt 0;">&nbsp;</p>`;

  /* ---------- 3. Secciones numeradas ---------- */

  // OBJETIVO
  h += titulo('OBJETIVO');
  h += parrafos(x.objetivo, 14.2, 'margin-bottom:6pt;');

  // ANTECEDENTES
  h += titulo('ANTECEDENTES');
  h += items(x.antecedentes);

  // BASE LEGAL
  h += titulo('BASE LEGAL');
  h += items(x.baseLegal);

  // DESCRIPCIÓN GENERAL
  h += titulo('DESCRIPCIÓN GENERAL');
  const sec = n;
  let sub = 0;

  if (lineas(x.descripcionGeneral).length) {
    h += np(`<b>${sec}.${++sub}.</b>`, lineas(x.descripcionGeneral).map(T).join('<br>'), 36, 21.8);
  }

  // Imagen N° 01: título centrado en negrita, imagen en recuadro con borde
  // y, debajo, la descripción en otro recuadro. Sin imagen no hay recuadro vacío.
  const hayImagen = /^data:image\/[\w.+-]+;base64,[A-Za-z0-9+/=]+$/.test(x.imagen || '');
  const hayDescImagen = lineas(x.consultaRegistro).length > 0;

  if (hayImagen || hayDescImagen) {
    h += `<p style="margin:14pt 0 0 48pt;text-align:center;width:420pt;"><b>${T(x.imagenTitulo) || esc(TITULO_IMAGEN)}</b></p>`;
    h += `<table style="width:420pt;margin:0 0 0 48pt;border-collapse:collapse;">`;

    if (hayImagen) {
      let img;
      if (word) {
        const t = tamanoWord();
        img = `<img src="${rutaImagenWord(x.imagen)}" width="${t.w}" height="${t.h}" alt="Imagen N° 01">`;
      } else {
        img = `<img src="${x.imagen}" alt="Imagen N° 01" style="width:100%;height:auto;display:block;">`;
      }
      h += `<tr><td align="center" style="${B}padding:0;text-align:center;vertical-align:middle;">${img}</td></tr>`;
    }

    if (hayDescImagen) {
      h += `<tr><td style="${B}padding:1pt 3pt;vertical-align:top;">
        <p style="margin:0;font-size:8pt;line-height:1.4;text-align:justify;">${lineas(x.consultaRegistro).map(T).join('<br>')}</p>
      </td></tr>`;
    }

    h += `</table>`;
  }

  if (lineas(x.direccion).length) {
    h += np(`<b>${sec}.${++sub}.</b>`, '<b><u>DIRECCIÓN</u></b>', 24, 24, 'margin-top:14pt;');
    h += parrafos(x.direccion, 36);
  }

  // DE LA VISITA DE FISCALIZACIÓN (subsección, igual que en el modelo)
  const subVisita = ++sub;
  h += np(`<b>${sec}.${subVisita}.</b>`, '<b><u>DE LA VISITA DE FISCALIZACIÓN</u></b>', 24, 24, 'margin-top:14pt;');

  let sub2 = 0;
  if (lineas(x.analisis).length) {
    h += np(`<b>${sec}.${subVisita}.${++sub2}.</b>`, lineas(x.analisis).map(T).join('<br>'), 50, 36, 'margin-top:6pt;');
  }

  h += `<p style="margin:2pt 0 0 50pt;width:404pt;text-align:center;"><b><u>TABLA Nº 01 Hechos Constatados, Consignados En El Acta De Fiscalización:</u></b></p>`;

  const hechos = (x.hechos || []).filter(f => CAMPOS_HECHO.some(k => (f[k] || '').trim()));
  const filasHechos = hechos.length
    ? hechos.map((f, i) => {
      const inc = lineas(f.incumplimiento).map((l, j, arr) =>
        `<p style="${PC}${j < arr.length - 1 ? 'margin-bottom:7pt;' : ''}">${T(l)}</p>`).join('');
      const bl = (f.baseLegal || '').trim().replace(/^base legal\s*:?\s*/i, '');
      const base = bl ? `<p style="${PC}margin-top:7pt;"><b><i>Base legal:</i></b> ${T(bl)}</p>` : '';
      const desc = lineas(f.descripcion).map(l => `<p style="${PC}color:${ROJO};">${T(l)}</p>`).join('');
      const medio = lineas(f.medio).map(l => `<p style="${PC}text-align:left;">${T(l)}</p>`).join('');
      return `<tr>
        <td valign="middle" style="${CELDA}text-align:center;"><p style="${PC}text-align:center;">${T(f.numero) || i + 1}</p></td>
        <td valign="top" style="${CELDA}vertical-align:top;">${inc}${base}</td>
        <td valign="middle" style="${CELDA}">${desc}</td>
        <td valign="middle" style="${CELDA}">${medio}</td>
      </tr>`;
    }).join('')
    : `<tr><td style="${CELDA}text-align:center;">—</td><td style="${CELDA}">&nbsp;</td><td style="${CELDA}">&nbsp;</td><td style="${CELDA}">&nbsp;</td></tr>`;

  h += `<table style="width:404pt;margin:0 0 0 50pt;border-collapse:collapse;">
    <tr>
      <th style="${CAB}width:66pt;">N.º</th>
      <th style="${CAB}width:171pt;">INCUMPLIMIENTO NORMATIVO – BASE LEGAL</th>
      <th style="${CAB}width:110pt;">DESCRIPCIÓN DEL INCUMPLIMIENTO VERIFICADO</th>
      <th style="${CAB}width:57pt;">MEDIO PROBATORIO</th>
    </tr>
    ${filasHechos}
  </table>`;

  h += parrafo(`Fuente: ${T(x.fuente) || '________'}`, 36, 'margin-right:-22pt;');

  if (lineas(x.cierreVisita).length) {
    h += np(`<b>${sec}.${subVisita}.${++sub2}.</b>`, lineas(x.cierreVisita).map(T).join('<br>'), 50, 36, 'margin-top:6pt;');
  }

  // ANEXO
  h += titulo('ANEXO');
  h += parrafos(x.anexoIntro, 14.2, 'margin-bottom:4pt;');
  h += lineas(x.anexos).map((l, i) =>
    np(romano(i + 1) + '.', T(l), 50, 14, 'font-size:9pt;line-height:1.2;')).join('');

  // ANÁLISIS DE DESCARGOS (solo si el agente presentó descargos)
  if (x.tieneDescargo && lineas(x.descargos).length) {
    h += titulo('ANÁLISIS DE DESCARGOS');
    h += items(x.descargos);
  }

  // CONCLUSIONES
  h += titulo('CONCLUSIONES');
  h += parrafo('De la visita de fiscalización a la unidad vehicular se concluye lo siguiente:', 14.2);
  h += items(x.conclusiones, 50, 36, true, 'font-size:9pt;');

  // RECOMENDACIONES
  h += titulo('RECOMENDACIONES');
  h += parrafo('Por lo expuesto en el presente informe, se recomienda:', 14.2);
  h += items(x.recomendaciones, 50, 36, false, 'font-size:9pt;');

  /* ---------- 5. Cierre y firmas ---------- */
  h += parrafo('Atentamente.', 14.2, 'margin-top:22pt;');
  h += `<p style="margin:0;line-height:normal;">&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>`;

  const firma = (nom, doc, cargo, emp) => `
    <td valign="top" style="width:50%;border:none;padding:0 4pt;vertical-align:top;text-align:center;">
      <p style="text-align:center;margin:0;font-size:9pt;line-height:1.3;"><b>${T(nom) || '&nbsp;'}</b></p>
      <p style="text-align:center;margin:0;font-size:9pt;line-height:1.3;">${T(doc) || '&nbsp;'}</p>
      <p style="text-align:center;margin:0;font-size:9pt;line-height:1.3;">${T(cargo) || '&nbsp;'}</p>
      <p style="text-align:center;margin:0;font-size:9pt;line-height:1.3;"><b>${T(emp) || '&nbsp;'}</b></p>
    </td>`;

  h += `<table style="width:100%;border:none;border-collapse:collapse;"><tr>
    ${firma(x.firma1Nombre, x.firma1Documento, x.firma1Cargo, x.firma1Empresa)}
    ${firma(x.firma2Nombre, x.firma2Documento, x.firma2Cargo, x.firma2Empresa)}
  </tr></table>`;

  // Notas al pie (^[...])
  h += listaNotas();

  NOTAS = null;
  return h;
}

// Documento HTML completo: 'word' (descarga) o 'print' (imprimir / PDF)
function documentoCompleto(x, modo) {
  const titulo = `Informe de Supervisión ${esc(reemplazar('{{expediente}}', x))}`;

  if (modo === 'word') {
    return `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><title>${titulo}</title>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom><w:DoNotOptimizeForBrowser/></w:WordDocument></xml><![endif]-->
<style>
@page WordSection1{size:595.3pt 841.9pt;margin:99.25pt 56.7pt 42.55pt 70.9pt;mso-header-margin:14.15pt;mso-footer-margin:17pt;mso-header:url("informe_archivos/header.htm") h1;}
div.WordSection1{page:WordSection1;}
${WORD_CSS}
</style></head>
<body lang="ES"><div class="WordSection1">${informeHTML(x, 'word')}</div></body></html>`;
  }

  // Imprimir / PDF: el encabezado va en un <thead>, que el navegador repite en cada hoja
  return `<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><title>${titulo}</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Poppins:ital,wght@0,400;0,700;1,400;1,700&display=swap">
<style>
@page{size:A4;margin:0.5cm 2cm 1.5cm 2.5cm;}
html,body{margin:0;background:#fff;}
*{-webkit-print-color-adjust:exact;print-color-adjust:exact;}
.doc table:not(.pagina) tr{page-break-inside:avoid;}
.pagina{width:100%;border-collapse:collapse;}
.pagina > thead > tr > td,.pagina > tbody > tr > td{padding:0;}
.pagina > thead > tr > td{padding-bottom:22pt;}
${DOC_CSS}
</style></head>
<body><div class="doc"><table class="pagina">
<thead><tr><td>${encabezadoHTML(x, 'web')}</td></tr></thead>
<tbody><tr><td>${informeHTML(x, 'print')}</td></tr></tbody>
</table></div></body></html>`;
}

/* ---------------------------------------------------------
   Tabla de hechos constatados (editor del formulario)
   --------------------------------------------------------- */
function agregarHecho(h = {}) {
  const div = document.createElement('div');
  div.className = 'hecho';
  div.innerHTML = `
    <div class="hecho-cab">
      <strong class="hecho-titulo">Fila</strong>
      <button type="button" class="btn-quitar">Quitar fila</button>
    </div>
    <div class="form-grid">
      <div class="form-group">
        <label>N.º / categoría <span class="nota">(vacío = número automático)</span></label>
        <input type="text" data-k="numero" placeholder="Ejemplo: OTRAS OBSERVACIONES">
      </div>
      <div class="form-group">
        <label>Medio probatorio</label>
        <input type="text" data-k="medio" placeholder="Ejemplo: Foto N.º 23, 30">
      </div>
    </div>
    <label>Incumplimiento normativo</label>
    <textarea data-k="incumplimiento" rows="3"></textarea>
    <label>Base legal</label>
    <textarea data-k="baseLegal" rows="2"></textarea>
    <label>Descripción del incumplimiento verificado</label>
    <textarea data-k="descripcion" rows="3"></textarea>`;

  div.querySelectorAll('[data-k]').forEach(el => { el.value = h[el.dataset.k] || ''; });

  div.querySelector('.btn-quitar').onclick = () => {
    div.remove();
    numerarFilas();
    programarVista();
  };

  $('hechos').append(div);
  numerarFilas();
}

function numerarFilas() {
  [...$('hechos').children].forEach((d, i) => {
    d.querySelector('.hecho-titulo').textContent = `Fila ${i + 1}`;
  });
}

function leerHechos() {
  return [...$('hechos').children].map(d => {
    const f = {};
    d.querySelectorAll('[data-k]').forEach(el => { f[el.dataset.k] = el.value; });
    return f;
  });
}

function ponerHechos(lista) {
  $('hechos').replaceChildren();
  (lista && lista.length ? lista : [{}]).forEach(agregarHecho);
}

/* ---------------------------------------------------------
   Formulario
   --------------------------------------------------------- */
function datos() {
  const x = {};
  CAMPOS.forEach(k => { x[k] = $(k).value; });
  x.antecedentes = PLANTILLA_ANTECEDENTES;
  x.tieneDescargo = $('tieneDescargo').checked;
  x.hechos = leerHechos();
  x.imagen = IMAGEN.dataUrl || '';
  return x;
}

function cargar(x, id = null) {
  const y = tokenizar(x);
  CAMPOS.forEach(k => { $(k).value = y[k] ?? ''; });
  if (!$('imagenTitulo').value) $('imagenTitulo').value = TITULO_IMAGEN;
  $('tieneDescargo').checked = !!y.tieneDescargo;
  $('zonaDescargo').hidden = !y.tieneDescargo;
  $('imagenArchivo').value = '';
  ponerHechos(y.hechos);
  editId = id;
  actualizarModo();
  ponerImagen(y.imagen || null);
  vista(datos());
}

// Texto modelo con la imagen de ejemplo (public/consulta-ejemplo.png)
function cargarModelo() {
  cargar({ ...MODELO, imagen: imagenEjemplo }, null);
}

function limpiar() {
  CAMPOS.forEach(k => { $(k).value = ''; });
  $('imagenTitulo').value = TITULO_IMAGEN;
  $('tieneDescargo').checked = false;
  $('zonaDescargo').hidden = true;
  $('imagenArchivo').value = '';
  ponerHechos([]);
  editId = null;
  actualizarModo();
  ponerImagen(null);
  vista(datos());
}

function actualizarModo() {
  $('btnGuardar').textContent = editId ? 'Actualizar informe' : 'Guardar informe';
  $('tituloForm').textContent = editId
    ? `Editando informe – Expediente ${$('expediente').value}`
    : 'Registro de informe de fiscalización';
}

function estado(msg, tipo = 'ok') {
  $('estado').textContent = msg;
  $('estado').className = 'estado ' + tipo;
  clearTimeout(estado.t);
  estado.t = setTimeout(() => { $('estado').textContent = ''; }, 5000);
}

/* ---------------------------------------------------------
   Vista previa en tiempo real
   --------------------------------------------------------- */
function vista(x) {
  $('vista').innerHTML = `<div class="doc">${informeHTML(x, 'web')}</div>`;
}

// La hoja mide 21 cm (como en Word); si la columna es más angosta,
// se reduce con zoom para que el texto corte las líneas igual que en el A4.
function ajustarHoja() {
  const marco = $('vistaMarco');
  const hoja = $('vista');
  if (!marco || !hoja) return;
  hoja.style.zoom = '';
  const disponible = marco.clientWidth - 32;
  const ancho = hoja.offsetWidth;
  if (ancho > 0 && disponible < ancho) hoja.style.zoom = String(disponible / ancho);
}

let pendiente = false;
function programarVista() {
  if (pendiente) return;
  pendiente = true;
  requestAnimationFrame(() => {
    pendiente = false;
    vista(datos());
    if (editId) actualizarModo();
  });
}

/* ---------------------------------------------------------
   Descargar en Word / Imprimir – PDF
   --------------------------------------------------------- */
function nombreArchivo(x, ext) {
  const exp = (x.expediente || 'sin_expediente').trim().replace(/[^\w.-]+/g, '_');
  return `Informe_Supervision_${exp}.${ext}`;
}

// Arma un .doc en formato MHTML ("página web de un solo archivo"):
// el HTML del informe y las imágenes (logo e Imagen N° 01) van DENTRO
// del mismo archivo, así Word las muestra sin depender de ninguna carpeta.
// imagenes: [{ ruta, tipo, base64 }]
function paqueteWord(html, imagenes = []) {
  const limite = '----=_SISFIS_VANESA_' + Date.now();
  const base = 'file:///C:/SISFIS_VANESA/';
  const htmlB64 = bytesABase64(new TextEncoder().encode(html));

  let mhtml = 'MIME-Version: 1.0\r\n'
    + `Content-Type: multipart/related; boundary="${limite}"; type="text/html"\r\n\r\n`
    + `--${limite}\r\n`
    + 'Content-Type: text/html; charset="utf-8"\r\n'
    + 'Content-Transfer-Encoding: base64\r\n'
    + `Content-Location: ${base}informe.htm\r\n\r\n`
    + base64MIME(htmlB64) + '\r\n';

  imagenes.filter(i => i && i.base64).forEach(i => {
    mhtml += `--${limite}\r\n`
      + `Content-Type: ${i.tipo}\r\n`
      + 'Content-Transfer-Encoding: base64\r\n'
      + `Content-Location: ${base}${i.ruta}\r\n\r\n`
      + base64MIME(i.base64) + '\r\n';
  });

  return mhtml + `--${limite}--\r\n`;
}

async function descargarWord() {
  await logoListo;
  if (!LOGO.base64) estado('No se encontró public/logo.png: el Word saldrá sin logo.', 'error');

  const x = datos();
  const html = documentoCompleto(x, 'word');

  // Partes que van DENTRO del .doc: encabezado de página, logo e imagen 01
  const cabecera = bytesABase64(new TextEncoder().encode(encabezadoWord(x)));
  const imagenes = [
    { ruta: 'informe_archivos/header.htm', tipo: 'text/html; charset="utf-8"', base64: cabecera },
    { ruta: LOGO.rutaWord, tipo: 'image/png', base64: LOGO.base64 }
  ];
  const img01 = partesDataURL(x.imagen);
  if (img01) imagenes.push({ ruta: rutaImagenWord(x.imagen), tipo: img01.tipo, base64: img01.base64 });

  const blob = new Blob([paqueteWord(html, imagenes)], { type: 'application/msword' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = nombreArchivo(x, 'doc');
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function imprimir() {
  const html = documentoCompleto(datos(), 'print');
  const marco = document.createElement('iframe');
  marco.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;';
  document.body.append(marco);

  const w = marco.contentWindow;
  w.document.open();
  w.document.write(html);
  w.document.close();

  // Espera a que el logo (y cualquier imagen) termine de cargar antes de imprimir
  const imagenes = [...w.document.images].map(img =>
    img.complete ? Promise.resolve() : new Promise(ok => { img.onload = img.onerror = ok; }));

  Promise.all(imagenes).then(() => {
    setTimeout(() => {
      w.focus();
      w.print();
      setTimeout(() => marco.remove(), 1000);
    }, 200);
  });
}

/* ---------------------------------------------------------
   Lista de informes (CRUD)
   --------------------------------------------------------- */
async function listar() {
  try {
    const r = await fetch('/api/informes');
    const a = await r.json();

    $('lista').replaceChildren();

    if (!a.length) {
      $('lista').textContent = 'Aún no hay informes registrados.';
      return;
    }

    a.forEach(x => {
      const p = document.createElement('p');
      const info = document.createElement('span');
      info.textContent = `${x.expediente} – ${x.empresa}`
        + (x.numeroInforme ? ` – Informe N° ${x.numeroInforme}` : '')
        + (x.fecha ? ` – ${fechaLarga(x.fecha)}` : '');

      const e = document.createElement('button');
      e.type = 'button';
      e.className = 'btn-editar';
      e.textContent = 'Editar';
      e.onclick = async () => {
        const rr = await fetch('/api/informes/' + x.id);
        if (!rr.ok) { estado('No se encontró el informe', 'error'); listar(); return; }
        cargar(await rr.json(), x.id);
        $('seccionRegistro').scrollIntoView({ behavior: 'smooth' });
      };

      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'btn-eliminar';
      b.textContent = 'Eliminar';
      b.onclick = async () => {
        if (!confirm('¿Eliminar el informe del expediente ' + x.expediente + '?')) return;
        await fetch('/api/informes/' + x.id, { method: 'DELETE' });
        if (editId === x.id) { editId = null; actualizarModo(); }
        listar();
      };

      p.append(info, e, b);
      $('lista').append(p);
    });
  } catch (err) {
    $('lista').textContent = 'No se pudo cargar la lista de informes.';
  }
}

/* ---------------------------------------------------------
   Eventos
   --------------------------------------------------------- */
$('tieneDescargo').onchange = () => {
  $('zonaDescargo').hidden = !$('tieneDescargo').checked;
  programarVista();
};

// Cualquier cambio en el formulario actualiza la vista previa al instante
$('form').addEventListener('input', programarVista);
$('form').addEventListener('change', programarVista);

$('agregarHecho').onclick = () => {
  agregarHecho();
  programarVista();
};

$('btnModelo').onclick = () => {
  cargarModelo();
  estado('Se cargó el texto modelo. Modifique los datos y guarde el informe.');
};

// Subir / reemplazar la imagen de la sección 4
$('imagenArchivo').addEventListener('change', async () => {
  const archivo = $('imagenArchivo').files[0];
  if (!archivo) return;
  if (!archivo.type.startsWith('image/')) {
    estado('El archivo seleccionado no es una imagen.', 'error');
    $('imagenArchivo').value = '';
    return;
  }
  try {
    $('imagenInfo').textContent = 'Procesando imagen...';
    await ponerImagen(await reducirImagen(archivo));
    estado('Imagen cargada. Se guardará junto con el informe.');
  } catch (err) {
    estado('No se pudo procesar la imagen.', 'error');
    mostrarMiniatura();
  }
});

$('btnQuitarImagen').onclick = () => {
  $('imagenArchivo').value = '';
  ponerImagen(null);
};

// Caja de marcadores: se puede cerrar (se recuerda en este navegador)
function guardarPreferencia(clave, valor) {
  try { localStorage.setItem(clave, valor); } catch (err) { /* sin almacenamiento */ }
}
function leerPreferencia(clave) {
  try { return localStorage.getItem(clave); } catch (err) { return null; }
}

if (leerPreferencia('sisfis_avisoCerrado') === '1') $('avisoMarcadores').hidden = true;
$('cerrarAviso').onclick = () => {
  $('avisoMarcadores').hidden = true;
  guardarPreferencia('sisfis_avisoCerrado', '1');
};

window.addEventListener('resize', ajustarHoja);

$('form').addEventListener('reset', ev => {
  ev.preventDefault();
  limpiar();
});

$('btnWord').onclick = descargarWord;
$('btnImprimir').onclick = imprimir;

$('form').onsubmit = async ev => {
  ev.preventDefault();

  // Si el usuario escribió el expediente/empresa/placa a mano en algún texto,
  // se guarda como marcador para que se actualice solo en el futuro
  const x = tokenizar(datos());

  try {
    const r = await fetch('/api/informes' + (editId ? '/' + editId : ''), {
      method: editId ? 'PUT' : 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(x)
    });

    if (!r.ok) {
      const e = await r.json().catch(() => ({}));
      estado(e.error || 'Revise expediente y empresa', 'error');
      return;
    }

    const saved = await r.json();
    const eraEdicion = !!editId;
    cargar(saved, saved.id);
    estado(eraEdicion ? 'Informe actualizado correctamente.' : 'Informe guardado correctamente.');
    listar();
  } catch (err) {
    estado('No se pudo conectar con el servidor.', 'error');
  }
};

// Estilos del documento también para la vista previa
const estiloDoc = document.createElement('style');
estiloDoc.textContent = DOC_CSS;
document.head.append(estiloDoc);

// Inicio: se muestra el texto modelo (con la imagen de ejemplo) para ver el formato
(async () => {
  await imagenEjemploLista;
  cargarModelo();
  ajustarHoja();
})();
listar();
