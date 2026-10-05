// Tarjeta E "Bloques": 90x55 mm + 3 mm de sangrado (96x61 mm). Frente y reverso, editable en PowerPoint y Canva.
const pptxgen = require('pptxgenjs');
const sharp = require('sharp');
const fs = require('fs');

const MM = 1 / 25.4, W = 96 * MM, H = 61 * MM;
const COND = 'Barlow Condensed', BODY = 'Barlow';
const GREEN = '23443A', YELLOW = 'F3D34A', WHITE = 'FFFFFF', SOFT = '56606E', INK = '1B2430', RED = 'CF2338';

(async () => {
  const pres = new pptxgen();
  pres.defineLayout({ name: 'TARJETA', width: W, height: H });
  pres.layout = 'TARJETA';
  pres.title = 'Despeja: tarjeta de presentación';
  const T = (s, txt, o) => s.addText(txt, Object.assign({ isTextBox: true, margin: 0, fontFace: BODY }, o));
  const box = (s, x, y, w, h, color, name) => s.addShape(pres.shapes.RECTANGLE, { x: x * MM, y: y * MM, w: w * MM, h: h * MM, fill: { color }, line: { type: 'none' }, objectName: name });

  // ---------- FRENTE ----------
  const f = pres.addSlide();
  f.background = { color: GREEN };
  box(f, 0, 0, 38, 61, YELLOW, 'bloque amarillo');
  // logotipo: Δ + nombre, centrado en el bloque amarillo (x 3–38)
  T(f, 'Δ', { x: 3 * MM, y: 11 * MM, w: 35 * MM, h: 22 * MM, fontFace: COND, bold: true, fontSize: 80, color: GREEN, align: 'center', valign: 'middle', objectName: 'logo delta' });
  T(f, 'Despeja', { x: 3 * MM, y: 35 * MM, w: 35 * MM, h: 8.5 * MM, fontFace: COND, bold: true, fontSize: 26, color: GREEN, align: 'center', valign: 'middle', charSpacing: 0.5, objectName: 'logo nombre' });
  T(f, 'clases de números y ciencias', { x: 3 * MM, y: 45.6 * MM, w: 35 * MM, h: 3 * MM, fontSize: 6, bold: true, color: GREEN, align: 'center', valign: 'middle', objectName: 'logo lema' });
  // mensaje (columna derecha x 44–89)
  T(f, [{ text: 'La nota de su hijo(a) también puede ', options: { color: WHITE } }, { text: 'cambiar.', options: { color: YELLOW } }],
    { x: 44 * MM, y: 10 * MM, w: 45 * MM, h: 22 * MM, fontFace: COND, bold: true, fontSize: 20, lineSpacingMultiple: 0.95, valign: 'top', objectName: 'titular' });
  ['Secundaria', 'Pre', 'Universidad'].reduce((x, t) => {
    const w = t.length * 1.25 + 4;
    f.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x * MM, y: 36 * MM, w: w * MM, h: 4.4 * MM, rectRadius: 0.08, fill: { type: 'none' }, line: { color: 'A9BDB4', width: 0.6 }, objectName: 'nivel ' + t });
    T(f, t, { x: x * MM, y: 36 * MM, w: w * MM, h: 4.4 * MM, fontSize: 6.3, bold: true, color: WHITE, align: 'center', valign: 'middle' });
    return x + w + 1.5;
  }, 44);
  T(f, [{ text: 'Profesores de Ingeniería de la', options: { breakLine: true } }, { text: 'Pontificia Universidad Católica del Perú', options: { bold: true } }],
    { x: 44 * MM, y: 47 * MM, w: 45 * MM, h: 7 * MM, fontSize: 6.3, color: 'E2E8E4', valign: 'bottom', objectName: 'autoridad' });

  // ---------- REVERSO ----------
  const b = pres.addSlide();
  b.background = { color: WHITE };
  box(b, 0, 0, 96, 24.5, YELLOW, 'franja oferta');
  T(b, 'Plan mensual', { x: 7 * MM, y: 4.8 * MM, w: 55 * MM, h: 5.5 * MM, fontFace: COND, bold: true, fontSize: 13, color: GREEN, valign: 'middle' });
  T(b, [{ text: '10%', options: { fontSize: 25 } }, { text: ' de descuento', options: { fontSize: 13 } }],
    { x: 7 * MM, y: 10.8 * MM, w: 58 * MM, h: 8.5 * MM, fontFace: COND, bold: true, color: GREEN, valign: 'bottom', objectName: 'oferta' });
  T(b, 'y facilidades de pago', { x: 7 * MM, y: 19.6 * MM, w: 55 * MM, h: 3 * MM, fontSize: 6.8, bold: true, color: GREEN, valign: 'middle' });
  b.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 73.6 * MM, y: 4 * MM, w: 15.4 * MM, h: 15.4 * MM, rectRadius: 0.05, fill: { color: WHITE }, line: { type: 'none' }, objectName: 'fondo QR' });
  const qr = await sharp(fs.readFileSync('../propuesta-A-html/qr-whatsapp.svg'), { density: 1200 }).resize(600, 600).png().toBuffer();
  b.addImage({ data: 'image/png;base64,' + qr.toString('base64'), x: 74.8 * MM, y: 5.2 * MM, w: 13 * MM, h: 13 * MM, altText: 'QR a WhatsApp', objectName: 'QR WhatsApp' });
  // pasos
  [['Diagnóstico', 'Ubicamos dónde se traba.'], ['Plan semanal', 'Metas hasta su examen.'], ['Reporte a padres', 'Cada mes, su avance.']].forEach(([k, d], i) => {
    T(b, [{ text: k, options: { fontFace: COND, bold: true, fontSize: 8.5, color: GREEN, breakLine: true } }, { text: d, options: { fontSize: 6.2, color: SOFT } }],
      { x: (7 + i * 27.3) * MM, y: 28 * MM, w: 26 * MM, h: 7 * MM, valign: 'top', objectName: 'paso ' + (i + 1) });
  });
  // cursos en dos viñetas (el punto es una forma)
  [[38.5, 'Secundaria y pre: ', 'Álgebra, Aritmética, Geometría, Trigonometría, Física y Química.'],
   [42, 'Universidad: ', 'Cálculo, Física, Química y cursos de Ingeniería.']].forEach(([y, k, d]) => {
    b.addShape(pres.shapes.OVAL, { x: 7.2 * MM, y: (y + 0.85) * MM, w: 1.15 * MM, h: 1.15 * MM, fill: { color: YELLOW }, line: { color: GREEN, width: 0.3 }, objectName: 'viñeta ' + k });
    T(b, [{ text: k, options: { bold: true, color: GREEN } }, { text: d, options: { color: SOFT } }],
      { x: 9.8 * MM, y: y * MM, w: 79 * MM, h: 3 * MM, fontSize: 6, valign: 'top', objectName: 'cursos ' + k });
  });
  b.addShape(pres.shapes.LINE, { x: 7 * MM, y: 47.6 * MM, w: 82 * MM, h: 0, line: { color: 'DDE3DF', width: 0.6 } });
  T(b, [{ text: 'WhatsApp  ', options: { color: SOFT } }, { text: '+51 961 956 660', options: { bold: true, color: INK } }], { x: 7 * MM, y: 48.6 * MM, w: 31 * MM, h: 5.4 * MM, fontSize: 6.5, valign: 'middle', objectName: 'whatsapp' });
  T(b, [{ text: 'Instagram  ', options: { color: SOFT } }, { text: '@despeja.clases', options: { bold: true, color: INK } }], { x: 37.5 * MM, y: 48.6 * MM, w: 30 * MM, h: 5.4 * MM, fontSize: 6.5, valign: 'middle', objectName: 'instagram' });
  T(b, 'Máx. 4 por grupo', { x: 69 * MM, y: 48.6 * MM, w: 20 * MM, h: 5.4 * MM, fontSize: 6.5, bold: true, color: RED, align: 'right', valign: 'middle', objectName: 'grupos' });

  f.addNotes('Frente. Tamaño final 90x55 mm; el borde exterior de 3 mm es sangrado y se corta.');
  b.addNotes('Reverso. Confirmar el usuario de Instagram antes de imprimir.');
  await pres.writeFile({ fileName: 'tarjeta-despeja.pptx' });
})();
