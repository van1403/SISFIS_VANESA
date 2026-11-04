const express = require('express');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

const file = path.join(__dirname, 'informes.json');

// Límite ampliado: los informes traen textos largos, la tabla de hechos
// y la imagen de la sección 4 en base64
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Campos de texto del informe (los textos pueden contener los marcadores
// {{expediente}}, {{empresa}}, {{placa}} y {{registro}}; se guardan tal cual
// para que, al cambiar el expediente, cambie en todo el informe)
const CAMPOS_TEXTO = [
 'expediente', 'empresa', 'placa', 'registro', 'supervisor',
 'supervisorCargo', 'empresaSupervisora', 'numeroInforme', 'ciudad', 'fecha',
 'destinatario', 'copias', 'asunto', 'referencia',
 'objetivo', 'antecedentes', 'baseLegal',
 'descripcionGeneral', 'imagenTitulo', 'consultaRegistro', 'direccion',
 'analisis', 'fuente', 'cierreVisita',
 'anexoIntro', 'anexos',
 'descargos', 'conclusiones', 'recomendaciones',
 'firma1Nombre', 'firma1Documento', 'firma1Cargo', 'firma1Empresa',
 'firma2Nombre', 'firma2Documento', 'firma2Cargo', 'firma2Empresa',
 // Datos de los antecedentes (2.1 a 2.5)
 'fechaRegistro', 'capacidad', 'compartimientos', 'direccionLegal', 'distrito', 'tracto',
 'resolucionNumero', 'resolucionFecha', 'cartaFecha', 'lugarFiscalizacion',
 'asignacionFecha', 'especialista', 'asignacionNumero', 'cartaLinea',
 'fechaVisita', 'fiscalizador'
];

const CAMPOS_HECHO = [
 'numero', 'incumplimiento', 'baseLegal', 'descripcion', 'medio'
];

const IMAGEN_VALIDA = /^data:image\/(png|jpeg|jpg|gif|webp|bmp);base64,[A-Za-z0-9+/=]+$/;

// Leer informes
function read() {
 if (!fs.existsSync(file)) {
  fs.writeFileSync(file, '[]', 'utf8');
  return [];
 }

 const contenido = fs.readFileSync(file, 'utf8').trim();

 if (!contenido) return [];

 const datos = JSON.parse(contenido);

 if (!Array.isArray(datos)) {
  throw new Error('informes.json debe contener un arreglo');
 }

 return datos;
}

// Guardar informes
function save(data) {
 fs.writeFileSync(
  file,
  JSON.stringify(data, null, 2),
  'utf8'
 );
}

// Validar datos
function validar(x) {
 if (!x || typeof x !== 'object') {
  return 'No se recibieron datos válidos';
 }

 if (typeof x.expediente !== 'string' ||
   !x.expediente.trim()) {
  return 'Debe ingresar el número de expediente';
 }

 if (typeof x.empresa !== 'string' ||
   !x.empresa.trim()) {
  return 'Debe ingresar el nombre de la empresa';
 }

 if (x.hechos !== undefined && !Array.isArray(x.hechos)) {
  return 'La tabla de hechos constatados no es válida';
 }

 return null;
}

// Deja solo los campos conocidos, como texto, y la tabla de hechos limpia
function normalizar(x) {
 const limpio = {};

 CAMPOS_TEXTO.forEach(k => {
  limpio[k] = typeof x[k] === 'string' ? x[k] : '';
 });

 limpio.expediente = limpio.expediente.trim();
 limpio.empresa = limpio.empresa.trim();
 limpio.placa = limpio.placa.trim();
 limpio.registro = limpio.registro.trim();
 limpio.tieneDescargo = !!x.tieneDescargo;

 // Imagen N° 01 (base64). Solo se acepta un data URL de imagen válido.
 limpio.imagen = typeof x.imagen === 'string' && IMAGEN_VALIDA.test(x.imagen)
  ? x.imagen
  : '';

 limpio.hechos = (Array.isArray(x.hechos) ? x.hechos : [])
  .filter(h => h && typeof h === 'object')
  .map(h => {
   const fila = {};
   CAMPOS_HECHO.forEach(k => {
    fila[k] = typeof h[k] === 'string' ? h[k] : '';
   });
   return fila;
  })
  .filter(h => CAMPOS_HECHO.some(k => h[k].trim()));

 return limpio;
}

// Listar
app.get('/api/informes', (req, res) => {
 res.json(read());
});

// Buscar por ID
app.get('/api/informes/:id', (req, res) => {
 const informe = read().find(
  x => x.id === req.params.id
 );

 if (!informe) {
  return res.status(404).json({
   error: 'Informe no encontrado'
  });
 }

 res.json(informe);
});

// Crear informe
app.post('/api/informes', (req, res) => {
 const error = validar(req.body);

 if (error) {
  return res.status(400).json({ error });
 }

 const informes = read();
 const ahora = new Date().toISOString();

 const nuevo = {
  ...normalizar(req.body),
  id: String(Date.now()),
  creado: ahora,
  actualizado: ahora
 };

 informes.push(nuevo);
 save(informes);

 res.status(201).json(nuevo);
});

// Actualizar informe
app.put('/api/informes/:id', (req, res) => {
 const error = validar(req.body);

 if (error) {
  return res.status(400).json({ error });
 }

 const informes = read();

 const indice = informes.findIndex(
  x => x.id === req.params.id
 );

 if (indice === -1) {
  return res.status(404).json({
   error: 'Informe no encontrado'
  });
 }

 informes[indice] = {
  ...normalizar(req.body),
  id: req.params.id,
  creado: informes[indice].creado || new Date().toISOString(),
  actualizado: new Date().toISOString()
 };

 save(informes);
 res.json(informes[indice]);
});

// Eliminar informe
app.delete('/api/informes/:id', (req, res) => {
 const informes = read();

 const nuevos = informes.filter(
  x => x.id !== req.params.id
 );

 if (nuevos.length === informes.length) {
  return res.status(404).json({
   error: 'Informe no encontrado'
  });
 }

 save(nuevos);
 res.status(204).end();
});

// Control de errores
app.use((err, req, res, next) => {
 console.error('Error del servidor:', err);

 res.status(err.status || 500).json({
  error: err.status === 400
   ? 'El JSON enviado no es válido'
   : err.status === 413
    ? 'El informe es demasiado grande'
    : 'Error interno del servidor'
 });
});

app.listen(PORT, () => {
 console.log('================================');
 console.log(' SISFIS_VANESA INICIADO');
 console.log(' http://localhost:' + PORT);
 console.log('================================');
});
