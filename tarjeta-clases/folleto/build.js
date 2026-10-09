// Folleto tríptico "Despeja": A4 apaisado (297x210 mm) + 3 mm de sangrado por lado = 303x216 mm.
// Todas las coordenadas están en mm medidos desde el corte (trim); el sangrado se suma al dibujar.
// Cara exterior (de izquierda a derecha): solapa 0-95 | contraportada 95-196 | portada 196-297
// Cara interior (de izquierda a derecha): panel 1 0-101 | panel 2 101-202 | panel 3 202-297
const pptxgen = require('pptxgenjs');
const sharp = require('sharp');
const fs = require('fs');

const IN = mm => mm / 25.4, BL = 3;
const COND = 'Barlow Condensed', BODY = 'Barlow';
const GREEN = '23443A', YELLOW = 'F3D34A', WHITE = 'FFFFFF', INK = '1B2430', SOFT = '56606E', MUTED = '7FA796', CHALK = 'E2E8E4', RED = 'CF2338';

(async () => {
  const pres = new pptxgen();
  pres.defineLayout({ name: 'FOLLETO', width: IN(297 + 2 * BL), height: IN(210 + 2 * BL) });
  pres.layout = 'FOLLETO';
  pres.title = 'Despeja: folleto tríptico';

  const pos = (x, y, w, h) => ({ x: IN(x + BL), y: IN(y + BL), w: IN(w), h: IN(h) });
  const rect = (s, x, y, w, h, fill, name) => s.addShape(pres.shapes.RECTANGLE, { ...pos(x, y, w, h), fill: { color: fill }, line: { type: 'none' }, objectName: name });
  const round = (s, x, y, w, h, fill, r, name, line) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { ...pos(x, y, w, h), rectRadius: IN(r), fill: fill ? { color: fill } : { type: 'none' }, line: line || { type: 'none' }, objectName: name });
  const dot = (s, x, y, d, name, fill = YELLOW, edge = GREEN) => s.addShape(pres.shapes.OVAL, { ...pos(x, y, d, d), fill: { color: fill }, line: { color: edge, width: 0.5 }, objectName: name });
  const text = (s, runs, x, y, w, h, o = {}) => s.addText(runs, { ...pos(x, y, w, h), isTextBox: true, margin: 0, fontFace: BODY, valign: 'top', ...o });
  const dotted = (s, x, y, w, color, name) => s.addShape(pres.shapes.LINE, { ...pos(x, y, w, 0), line: { color, width: 0.75, dashType: 'sysDot' }, objectName: name });

  const qrPng = await sharp(fs.readFileSync('../propuesta-A-html/qr-whatsapp.svg'), { density: 1500 }).resize(1000, 1000).png().toBuffer();
  const qrData = 'image/png;base64,' + qrPng.toString('base64');

  // =====================================================  CARA EXTERIOR
  const o = pres.addSlide();
  o.background = { color: WHITE };

  // ---- Portada (x 196-297)
  rect(o, 196, -BL, 101 + BL, 98, YELLOW, 'portada · bloque amarillo');
  rect(o, 196, 95, 101 + BL, 115 + BL, GREEN, 'portada · bloque verde');
  text(o, 'Δ', 196, 2, 101, 56, { fontFace: COND, bold: true, fontSize: 164, color: GREEN, align: 'center', valign: 'middle', objectName: 'logo delta' });
  text(o, 'Despeja', 196, 58, 101, 22, { fontFace: COND, bold: true, fontSize: 68, color: GREEN, align: 'center', valign: 'middle', objectName: 'logo nombre' });
  text(o, 'clases de números y ciencias', 196, 85, 101, 6, { bold: true, fontSize: 12, color: GREEN, align: 'center', valign: 'middle', objectName: 'logo lema' });
  text(o, 'Clases particulares y grupales', 205, 104, 83, 6, { bold: true, fontSize: 11, color: MUTED, objectName: 'portada · servicio' });
  text(o, [{ text: 'La nota de su hijo(a) también puede ', options: { color: WHITE } }, { text: 'cambiar.', options: { color: YELLOW } }],
    205, 112, 83, 42, { fontFace: COND, bold: true, fontSize: 38, lineSpacingMultiple: 0.92, objectName: 'portada · titular' });
  let px = 205;
  ['Secundaria', 'Pre', 'Universidad'].forEach(t => {
    const w = t.length * 2 + 6.5;
    round(o, px, 164, w, 7, null, 3.5, 'nivel ' + t, { color: 'A9BDB4', width: 0.75 });
    text(o, t, px, 164, w, 7, { bold: true, fontSize: 10, color: WHITE, align: 'center', valign: 'middle' });
    px += w + 2.5;
  });
  dotted(o, 205, 183, 83, MUTED, 'portada · repisa');
  text(o, [{ text: 'Profesores de Ingeniería de la', options: { breakLine: true } }, { text: 'Pontificia Universidad Católica del Perú', options: { bold: true } }],
    205, 187, 83, 13, { fontSize: 10, color: CHALK, objectName: 'portada · autoridad' });

  // ---- Contraportada (x 95-196)
  rect(o, 95, -BL, 101, 216, GREEN, 'contraportada · fondo');
  text(o, [{ text: 'Δ ', options: { color: YELLOW } }, { text: 'Despeja', options: { color: YELLOW } }], 104, 12, 83, 14, { fontFace: COND, bold: true, fontSize: 32, valign: 'middle', objectName: 'contraportada · marca' });
  text(o, 'Escríbanos hoy y coordinemos el diagnóstico de su hijo(a).', 104, 38, 83, 44, { fontFace: COND, bold: true, fontSize: 31, color: WHITE, lineSpacingMultiple: 0.95, objectName: 'contraportada · titular' });
  round(o, 119.5, 90, 52, 52, WHITE, 3, 'contraportada · fondo QR');
  o.addImage({ ...pos(124, 94.5, 43, 43), data: qrData, altText: 'Código QR que abre WhatsApp', objectName: 'QR WhatsApp' });
  text(o, 'Escanee y escríbanos por WhatsApp', 104, 146, 83, 6, { fontSize: 10.5, color: CHALK, align: 'center', valign: 'middle', objectName: 'contraportada · QR leyenda' });
  text(o, [{ text: 'WhatsApp', options: { fontSize: 9, color: MUTED, bold: true, breakLine: true } }, { text: '+51 961 956 660', options: { fontSize: 17, color: WHITE, bold: true } }],
    104, 160, 83, 13, { fontFace: BODY, objectName: 'WhatsApp' });
  text(o, [{ text: 'Instagram', options: { fontSize: 9, color: MUTED, bold: true, breakLine: true } }, { text: '@despeja.clases', options: { fontSize: 17, color: WHITE, bold: true } }],
    104, 177, 83, 13, { fontFace: BODY, objectName: 'Instagram' });
  text(o, [{ text: 'Profesores de Ingeniería de la', options: { breakLine: true } }, { text: 'Pontificia Universidad Católica del Perú', options: { bold: true } }], 104, 193, 83, 10, { fontSize: 9, color: CHALK, objectName: 'contraportada · autoridad' });

  // ---- Solapa (x 0-95): preguntas de los padres
  text(o, [{ text: 'Despejamos', options: { breakLine: true } }, { text: 'sus dudas.' }], 9, 12, 77, 24, { fontFace: COND, bold: true, fontSize: 32, color: GREEN, lineSpacingMultiple: 0.92, objectName: 'solapa · titulo' });
  [['¿Cómo sé si mi hijo(a) avanza?', 'Cada mes usted recibe un reporte de avance. Cada semana trabajamos con metas claras.'],
   ['¿Cuántos alumnos hay por grupo?', 'Máximo 4. También ofrecemos clases particulares, solo para su hijo(a).'],
   ['¿Cómo se paga?', 'El plan mensual tiene 10% de descuento y facilidades de pago.'],
   ['¿Qué niveles atienden?', 'Secundaria, preuniversitario y universidad.']].forEach(([q, a], i) => {
    const y = 46 + i * 39;
    dot(o, 9, y + 1.4, 3.4, 'solapa · viñeta ' + (i + 1));
    text(o, [{ text: q, options: { fontFace: COND, bold: true, fontSize: 15.5, color: GREEN, breakLine: true } }, { text: a, options: { fontSize: 10.5, color: SOFT } }],
      15.5, y, 71, 33, { lineSpacingMultiple: 1.05, objectName: 'solapa · pregunta ' + (i + 1) });
  });

  text(o, [{ text: 'Δ ', options: { color: YELLOW } }, { text: 'Despeja', options: { color: GREEN } }], 9, 188, 77, 12, { fontFace: COND, bold: true, fontSize: 26, valign: 'middle', objectName: 'solapa · marca' });

  // =====================================================  CARA INTERIOR
  const n = pres.addSlide();
  n.background = { color: WHITE };

  // ---- Panel 1 (x 0-101): método
  text(n, 'Así acompañamos a su hijo(a)', 9, 12, 83, 24, { fontFace: COND, bold: true, fontSize: 32, color: GREEN, lineSpacingMultiple: 0.92, objectName: 'método · titulo' });
  text(n, 'Un plan estructurado de estudio: usted siempre sabe en qué punto está su hijo(a).', 9, 39, 83, 17, { fontSize: 10.5, color: SOFT, lineSpacingMultiple: 1.1, objectName: 'método · intro' });
  [['1', 'Diagnóstico', 'Ubicamos el tema exacto donde su hijo(a) se traba.'],
   ['2', 'Plan semanal', 'Metas claras, semana a semana, hasta su examen.'],
   ['3', 'Reporte a padres', 'Cada mes, usted ve cómo avanza.']].forEach(([k, t, d], i) => {
    const y = 64 + i * 35;
    round(n, 9, y, 12, 12, GREEN, 6, 'método · círculo ' + k);
    text(n, k, 9, y, 12, 12, { fontFace: COND, bold: true, fontSize: 20, color: YELLOW, align: 'center', valign: 'middle' });
    text(n, [{ text: t, options: { fontFace: COND, bold: true, fontSize: 19, color: GREEN, breakLine: true } }, { text: d, options: { fontSize: 10.5, color: SOFT } }],
      26, y - 0.5, 66, 28, { lineSpacingMultiple: 1.05, objectName: 'método · paso ' + k });
  });
  round(n, 9, 172, 83, 26, GREEN, 3, 'método · recuadro grupos');
  text(n, [{ text: 'Máx. 4 alumnos por grupo', options: { fontFace: COND, bold: true, fontSize: 20, color: YELLOW, breakLine: true } }, { text: 'Atención cercana para cada alumno.', options: { fontSize: 10.5, color: WHITE } }],
    15, 172, 71, 26, { valign: 'middle', objectName: 'método · grupos' });

  // ---- Panel 2 (x 101-202): cursos
  text(n, 'Qué enseñamos', 110, 12, 83, 14, { fontFace: COND, bold: true, fontSize: 36, color: GREEN, objectName: 'cursos · titulo' });
  text(n, [{ text: 'Clases particulares y grupales', options: { breakLine: true } }, { text: 'de números y ciencias.' }], 110, 28, 83, 12, { fontSize: 10.5, color: SOFT, lineSpacingMultiple: 1.1, objectName: 'cursos · intro' });
  round(n, 110, 46, 83, 50, YELLOW, 3, 'cursos · bloque secundaria y pre');
  text(n, 'Secundaria y pre', 117, 51, 70, 10, { fontFace: COND, bold: true, fontSize: 22, color: GREEN, valign: 'middle' });
  [['Álgebra', 'Aritmética', 'Geometría'], ['Trigonometría', 'Física', 'Química']].forEach((col, c) => col.forEach((t, r) => {
    dot(n, 117 + c * 37, 66.2 + r * 8.6, 3, 'viñeta ' + t, GREEN, GREEN);
    text(n, t, 122 + c * 37, 65 + r * 8.6, 32, 6, { fontSize: 12, color: GREEN, bold: true, valign: 'middle' });
  }));
  round(n, 110, 104, 83, 52, GREEN, 3, 'cursos · bloque universidad');
  text(n, 'Universidad', 117, 109, 70, 10, { fontFace: COND, bold: true, fontSize: 22, color: YELLOW, valign: 'middle' });
  ['Cálculo', 'Física', 'Química', 'Cursos de Ingeniería'].forEach((t, r) => {
    dot(n, 117, 124.2 + r * 8.2, 3, 'viñeta ' + t);
    text(n, t, 122, 123 + r * 8.2, 66, 6, { fontSize: 12, color: WHITE, bold: true, valign: 'middle' });
  });
  text(n, [{ text: 'Quiénes enseñan', options: { fontFace: COND, bold: true, fontSize: 19, color: GREEN, breakLine: true } },
           { text: 'Profesores de Ingeniería de la Pontificia Universidad Católica del Perú, con experiencia en docencia.', options: { fontSize: 10.5, color: SOFT } }],
    110, 165, 83, 32, { lineSpacingMultiple: 1.08, objectName: 'cursos · quiénes' });

  // ---- Panel 3 (x 202-297): plan mensual
  rect(n, 202, -BL, 98, 216, GREEN, 'plan · fondo');
  text(n, 'Plan mensual', 211, 13, 77, 12, { fontFace: COND, bold: true, fontSize: 28, color: WHITE, valign: 'middle', objectName: 'plan · titulo' });
  text(n, '10%', 211, 24, 77, 44, { fontFace: COND, bold: true, fontSize: 104, color: YELLOW, valign: 'top', objectName: 'plan · 10%' });
  text(n, 'de descuento', 211, 66, 77, 12, { fontFace: COND, bold: true, fontSize: 30, color: WHITE, valign: 'middle', objectName: 'plan · descuento' });
  text(n, 'y facilidades de pago', 211, 79, 77, 8, { bold: true, fontSize: 14, color: YELLOW, valign: 'middle', objectName: 'plan · facilidades' });
  dotted(n, 211, 98, 77, MUTED, 'plan · línea');
  text(n, 'Cómo empezar', 211, 104, 77, 10, { fontFace: COND, bold: true, fontSize: 24, color: WHITE, valign: 'middle', objectName: 'plan · cómo empezar' });
  [['1', 'Escríbanos por WhatsApp o escanee el QR.'], ['2', 'Coordinamos el diagnóstico de su hijo(a).'], ['3', 'Usted recibe el plan semanal y, cada mes, su reporte.']].forEach(([k, d], i) => {
    const y = 120 + i * 21;
    round(n, 211, y, 10, 10, YELLOW, 5, 'plan · círculo ' + k);
    text(n, k, 211, y, 10, 10, { fontFace: COND, bold: true, fontSize: 17, color: GREEN, align: 'center', valign: 'middle' });
    text(n, d, 226, y - 1, 62, 14, { fontSize: 11, color: WHITE, valign: 'top', lineSpacingMultiple: 1.08, objectName: 'plan · paso ' + k });
  });
  text(n, [{ text: 'WhatsApp  ', options: { color: MUTED, fontSize: 10, bold: true } }, { text: '+51 961 956 660', options: { color: YELLOW, fontSize: 17, bold: true } }],
    211, 189, 77, 9, { valign: 'middle', objectName: 'plan · contacto' });

  o.addNotes('CARA EXTERIOR. Orden de izquierda a derecha: solapa (0-95 mm) | contraportada (95-196 mm) | portada (196-297 mm). Dobleces en 95 mm y 196 mm. Tamaño de página 303 x 216 mm con 3 mm de sangrado por lado.');
  n.addNotes('CARA INTERIOR. Orden de izquierda a derecha: panel 1 (0-101 mm) | panel 2 (101-202 mm) | panel 3 (202-297 mm). Dobleces en 101 mm y 202 mm. Imprimir a doble cara, volteando por el lado corto. La solapa (exterior izquierda) queda detrás del panel 3.');
  await pres.writeFile({ fileName: 'folleto-despeja.pptx' });
})();
