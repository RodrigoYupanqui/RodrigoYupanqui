// Plan de marca y marketing de Despeja, en Word (A4).
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, ImageRun, LevelFormat, BorderStyle, PageBreak, Footer, PageNumber } = require('docx');

const GREEN = '23443A', YELLOW = 'F3D34A', SOFT = '56606E';
const FULL = 9026; // ancho útil A4 con márgenes de 1" (DXA)
const p = (text, o = {}) => new Paragraph({ spacing: { after: 120 }, ...o, children: Array.isArray(text) ? text : [new TextRun(text)] });
const b = (t) => new TextRun({ text: t, bold: true });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const bullet = (runs) => new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 60 }, children: typeof runs === 'string' ? [new TextRun(runs)] : runs });
const num = (t) => new Paragraph({ numbering: { reference: 'steps', level: 0 }, spacing: { after: 60 }, children: [new TextRun(t)] });
const border = { style: BorderStyle.SINGLE, size: 4, color: 'D5DDD8' };
const cell = (content, w, opts = {}) => new TableCell({
  width: { size: w, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
  borders: { top: border, bottom: border, left: border, right: border },
  shading: opts.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: opts.fill } : undefined,
  children: (Array.isArray(content) ? content : [content]).map(c => typeof c === 'string'
    ? new Paragraph({ children: [new TextRun({ text: c, bold: !!opts.bold, color: opts.color })] }) : c) });
const table = (widths, rows, head = true) => new Table({
  width: { size: widths.reduce((a, c) => a + c, 0), type: WidthType.DXA }, columnWidths: widths,
  rows: rows.map((r, i) => new TableRow({ tableHeader: head && i === 0,
    children: r.map((c, j) => cell(c, widths[j], head && i === 0 ? { bold: true, fill: GREEN, color: 'FFFFFF' } : {})) })) });
const img = (path, w, h) => new ImageRun({ type: 'png', data: fs.readFileSync(path), transformation: { width: w, height: h } });

const children = [
  // Portada
  new Paragraph({ spacing: { before: 1800, after: 120 }, children: [new TextRun({ text: 'Δ', bold: true, size: 120, color: GREEN, font: 'Arial' })] }),
  new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: 'Despeja', bold: true, size: 80, color: GREEN })] }),
  new Paragraph({ spacing: { after: 600 }, children: [new TextRun({ text: 'Plan de marca y marketing', size: 36, color: SOFT })] }),
  p([new TextRun({ text: 'Clases particulares y grupales de números y ciencias para secundaria, preuniversitario y universidad, con profesores de Ingeniería de la Pontificia Universidad Católica del Perú.', color: SOFT })]),
  p([new TextRun({ text: 'Octubre de 2026', color: SOFT })]),
  new Paragraph({ children: [new PageBreak()] }),

  h1('1. Decisiones tomadas'),
  table([2600, 6426], [
    ['Tema', 'Decisión'],
    ['Nombre', 'Despeja (alternativa: Método 20). Falta confirmar usuario de Instagram y búsqueda en Indecopi.'],
    ['A quién le habla', 'A los padres de familia: son quienes pagan. Todo en "usted" y hablando de "su hijo(a)".'],
    ['Niveles', 'Secundaria, preuniversitario y universidad (en colegio solo secundaria).'],
    ['Cursos', 'Secundaria y pre: Álgebra, Aritmética, Geometría, Trigonometría, Física y Química. Universidad: Cálculo, Física, Química y cursos de Ingeniería.'],
    ['Oferta', 'Plan mensual con 10 % de descuento y facilidades de pago. Sin clases gratuitas.'],
    ['Servicio', 'Diagnóstico, plan semanal y reporte mensual de avance para los padres. Grupos de máximo 4 alumnos.'],
    ['Respaldo', 'Profesores de Ingeniería de la Pontificia Universidad Católica del Perú, sin usar logo ni colores de la universidad.'],
    ['Contacto', 'WhatsApp +51 961 956 660 (QR con mensaje ya escrito) e Instagram @despeja.clases.'],
    ['Diseño elegido', 'Tarjeta E "Bloques": logotipo Δ + Despeja sobre amarillo, mensaje para padres sobre verde.'],
    ['Herramientas', 'PowerPoint y Canva. Adobe descartado.'],
  ]),

  h1('2. Nombre'),
  p([b('Despeja '), new TextRun('viene de "despejar la x", lo primero que se aprende en álgebra, y también significa quitar dudas. Es un verbo: dice lo que hace la academia. Es corto, se dicta fácil por teléfono y no suena a academia genérica.')]),
  table([2000, 2000, 5026], [
    ['Nombre', 'Usuario', 'Por qué sí / por qué no'],
    ['Despeja (elegido)', '@despeja.clases', 'Juego de palabras con álgebra y con "quitar dudas". Fácil de recordar.'],
    ['Método 20', '@metodo20', 'El 20 es la nota máxima en Perú. Nombra el resultado y el cómo.'],
    ['Milimetrado', '@milimetrado.pe', 'Suena a ingeniería; más largo y difícil de dictar.'],
    ['Ingenio', '@ingenio.clases', 'Positivo pero muy común; probablemente ya existe.'],
    ['Delta', '@delta.clases', 'Δ = cambio. Muy usado por empresas.'],
  ]),
  p([new TextRun({ text: 'Antes de imprimir: revisar que el usuario esté libre en Instagram y buscar el nombre en la base de marcas de Indecopi.', italics: true, color: SOFT })], { spacing: { before: 120 } }),

  h1('3. Identidad visual'),
  h2('Logotipo'),
  p('La letra griega Δ (delta, que en ciencias significa "cambio") sobre el nombre Despeja, en verde sobre amarillo. Debajo, el lema "clases de números y ciencias". Usar el mismo logotipo en tarjeta, foto de perfil de Instagram y WhatsApp.'),
  h2('Colores'),
  new Table({ width: { size: FULL, type: WidthType.DXA }, columnWidths: [1400, 2200, 1600, 3826], rows: [
    new TableRow({ tableHeader: true, children: ['Muestra', 'Nombre', 'Código', 'Uso'].map((t, j) => cell(t, [1400, 2200, 1600, 3826][j], { bold: true, fill: GREEN, color: 'FFFFFF' })) }),
    ...[['23443A', 'Verde pizarra', 'Principal: fondos, logotipo, titulares'], ['F3D34A', 'Amarillo tiza', 'Acento: bloque del logo, ofertas'], ['FFFFFF', 'Blanco', 'Fondo del reverso y textos sobre verde'],
        ['1B2430', 'Grafito', 'Texto principal sobre fondo claro'], ['56606E', 'Gris', 'Textos secundarios'], ['CF2338', 'Rojo', 'Solo para avisos puntuales (cupos)']]
      .map(([hex, n, u]) => new TableRow({ children: [cell('', 1400, { fill: hex }), cell(n, 2200), cell('#' + hex, 1600), cell(u, 3826)] })) ] }),
  h2('Tipografías'),
  bullet([b('Barlow Condensed (negrita): '), new TextRun('titulares y logotipo.')]),
  bullet([b('Barlow: '), new TextRun('textos y datos de contacto.')]),
  p('Ambas son gratuitas en Google Fonts y están disponibles en Canva.'),

  new Paragraph({ children: [new PageBreak()] }),
  h1('4. Tarjeta de presentación'),
  p('Tamaño final 9 × 5,5 cm, con 3 mm de sangrado por lado para imprenta (9,6 × 6,1 cm en el archivo).'),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [img('../propuesta-E/frente.png', 454, 288)] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [new TextRun({ text: 'Frente', color: SOFT })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [img('../propuesta-E/reverso.png', 454, 288)] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [new TextRun({ text: 'Reverso', color: SOFT })] }),
  h2('Archivos'),
  bullet([b('Canva (editable): '), new TextRun('"Despeja – Tarjeta de presentación (final)" en su cuenta.')]),
  bullet([b('PowerPoint (editable): '), new TextRun('tarjeta-clases/propuesta-E/tarjeta-despeja.pptx')]),
  bullet([b('PDF para imprenta: '), new TextRun('tarjeta-clases/propuesta-E/tarjeta-despeja.pdf')]),

  h1('5. Mensaje y tono'),
  bullet('Se le habla al padre o madre de familia, de "usted", sobre "su hijo(a)".'),
  bullet('Titular: "La nota de su hijo(a) también puede cambiar." Promete un cambio posible, no una nota garantizada.'),
  bullet('El método en tres pasos: Diagnóstico (ubicamos dónde se traba), Plan semanal (metas hasta su examen) y Reporte a padres (cada mes, su avance).'),
  bullet('Evitar promesas como "aprobación garantizada" o "resultados asegurados".'),

  h1('6. Oferta'),
  table([2600, 6426], [
    ['Elemento', 'Detalle'],
    ['Clase suelta', 'Precio normal. Sirve como referencia para que el plan se vea conveniente.'],
    ['Plan mensual', '10 % de descuento sobre el precio de las clases sueltas.'],
    ['Facilidades de pago', 'Por ejemplo, pago en dos partes o por Yape/Plin. Definir las condiciones exactas.'],
    ['Grupos', 'Máximo 4 alumnos: es una escasez real y un argumento de calidad.'],
    ['Referidos (sugerencia)', 'Entregar 3 tarjetas a cada alumno inscrito. Definir si habrá un beneficio por recomendar.'],
  ]),

  h1('7. Por qué la tarjeta funciona'),
  table([2600, 6426], [
    ['Principio', 'Cómo se aplica'],
    ['Autoridad', 'El nombre completo de la Pontificia Universidad Católica del Perú da confianza inmediata.'],
    ['Aversión a la pérdida', 'La palabra "nota" toca la preocupación del padre por el rendimiento de su hijo.'],
    ['Esperanza concreta', '"También puede cambiar" y la Δ (cambio) muestran que hay salida.'],
    ['Anclaje', 'El precio de la clase suelta hace que el plan mensual se sienta conveniente.'],
    ['Compromiso', 'Quien elige un plan mensual se vuelve constante y cumple más.'],
    ['Menos fricción', 'El QR abre WhatsApp con el mensaje ya escrito.'],
    ['Exposición repetida', 'Mismo logotipo y colores en tarjeta, Instagram y estados de WhatsApp.'],
  ]),

  h1('8. Plan de marketing'),
  h2('Instagram (@despeja.clases)'),
  p('Bio sugerida:'),
  ...['Clases de números y ciencias', 'Secundaria · Pre · Universidad', 'Profes de Ingeniería PUCP', 'Plan con reporte mensual para padres', 'WhatsApp: +51 961 956 660']
    .map(t => new Paragraph({ indent: { left: 360 }, spacing: { after: 20 }, children: [new TextRun({ text: t, color: GREEN })] })),
  p('Tres tipos de contenido que se alternan (3 publicaciones por semana e historias diarias):', { spacing: { before: 160, after: 80 } }),
  bullet([b('Ejercicio en 60 segundos: '), new TextRun('un problema típico de examen resuelto en reel.')]),
  bullet([b('Antes y después: '), new TextRun('avance de un alumno, con permiso escrito y sin nombre.')]),
  bullet([b('Detrás del método: '), new TextRun('quiénes son los profesores, cómo es un diagnóstico, un reporte para padres.')]),
  p('Historias destacadas: Cursos, Planes, Cómo trabajamos, Testimonios.'),
  h2('WhatsApp Business'),
  bullet('Mensaje de bienvenida que pregunta curso, grado y fecha del examen.'),
  bullet('Catálogo con la clase suelta y el plan mensual.'),
  bullet('Respuestas rápidas para precios, horarios, modalidad y facilidades de pago.'),
  bullet('Etiquetas por etapa: nuevo, diagnóstico agendado, plan activo, renovar.'),
  bullet('Un estado semanal con un tip o ejercicio.'),
  h2('Medir qué zona funciona'),
  p('Imprimir las tarjetas en tandas con un QR distinto por zona. Cada QR abre WhatsApp con un mensaje que incluye un código, por ejemplo "(SM1)". Así se sabe qué colegio, librería o barrio trae alumnos.'),
  h2('Dónde repartir'),
  bullet('Librerías y fotocopiadoras cerca de colegios, sobre todo en semanas de exámenes.'),
  bullet('Grupos de padres del colegio y grupos de Facebook del distrito (foto de la tarjeta).'),
  bullet('Cada alumno inscrito recibe 3 tarjetas para recomendar.'),
  bullet('Si hay clases presenciales: perfil de Google Maps para pedir reseñas.'),

  h1('9. Primeros 30 días'),
  ...['Confirmar el nombre y que @despeja.clases esté libre; buscar la marca en Indecopi.',
      'Fijar precios: clase suelta y plan mensual (10 % de descuento) y las facilidades de pago.',
      'Abrir Instagram y WhatsApp Business con bio, catálogo y respuestas rápidas.',
      'Publicar 6 posts antes de repartir tarjetas, para que el perfil no se vea vacío.',
      'Imprimir 500 tarjetas en 2 o 3 tandas con QR por zona y repartir.',
      'Al final del mes, contar mensajes por zona y repetir en las que funcionaron.'].map(num),

  h1('10. Cuidados legales'),
  bullet('No usar logo, escudo ni colores de la PUCP; mencionar la universidad solo como respaldo real de los profesores.'),
  bullet('No prometer resultados garantizados: Indecopi puede considerarlo publicidad engañosa.'),
  bullet('Si el nombre funciona, registrarlo como marca en Indecopi (revisar la tasa vigente).'),
  bullet('Testimonios y notas de alumnos solo con permiso escrito de los padres.'),

  h1('11. Pendientes'),
  bullet('Confirmar el usuario de Instagram.'),
  bullet('Definir precios y condiciones de las facilidades de pago.'),
  bullet('Confirmar "Máximo 4 por grupo".'),
  bullet('Elegir imprenta y pedir prueba de color antes del tiraje completo.'),
];

const doc = new Document({
  creator: 'Despeja', title: 'Plan de marca y marketing – Despeja',
  styles: {
    default: { document: { run: { font: 'Arial', size: 22, color: '1B2430' } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 32, bold: true, color: GREEN }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 26, bold: true, color: GREEN }, paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1 } },
    ] },
  numbering: { config: [
    { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: 'steps', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] } ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'Despeja · Plan de marca  ', color: SOFT, size: 18 }), new TextRun({ children: [PageNumber.CURRENT], color: SOFT, size: 18 })] })] }) },
    children }],
});
Packer.toBuffer(doc).then(buf => fs.writeFileSync('Plan-de-marca-Despeja.docx', buf));
