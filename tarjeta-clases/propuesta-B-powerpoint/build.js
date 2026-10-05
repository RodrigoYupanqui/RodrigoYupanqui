// Propuesta B: "Pizarra". Tarjeta 90x55 mm + 3 mm de sangrado (96x61 mm), 2 diapositivas: frente y reverso.
const pptxgen = require('pptxgenjs');
const sharp = require('sharp');
const fs = require('fs');
const { applyTheme } = require(process.env.SKILL + '/scripts/apply_theme.js');

const MM = 1 / 25.4, W = 96 * MM, H = 61 * MM, S = 7 * MM; // S = margen seguro desde el borde con sangrado
const THEME = { name: 'Pizarra', headFontFace: 'Century Schoolbook', bodyFontFace: 'Calibri',
  colors: { dk1: '1D2B25', lt1: 'FFFFFF', dk2: '23443A', lt2: 'EEF2EE', accent1: 'F3D34A', accent2: 'F2F0E6',
    accent3: '7FA796', accent4: 'C9452E', accent5: '3E6B5A', accent6: '9AA7A1', hlink: '23443A', folHlink: '3E6B5A' } };

(async () => {
  const pres = new pptxgen();
  pres.defineLayout({ name: 'TARJETA', width: W, height: H });
  pres.layout = 'TARJETA';
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = 'Tarjeta de presentación: clases';
  const C = pres.SchemeColor;
  const T = (s, txt, o) => s.addText(txt, Object.assign({ isTextBox: true, margin: 0 }, o));

  // Retícula común a ambas caras (mm, medidos desde el borde con sangrado):
  // área segura x 7–89, y 7–54; columna principal x 7–62; columna derecha x 66–89;
  // línea divisoria en y 44 y fila inferior en y 46–52.
  const L = 7, R = 89, COL = 66, COLW = R - COL, LINE_Y = 44, ROW_Y = 46, ROW_H = 6;

  // ---------- FRENTE ----------
  const f = pres.addSlide();
  f.background = { color: '23443A' };
  // fórmulas en tiza tenue, solo en zonas libres
  [['F = m·a', 50, 6.2, 8, -5], ['Δx / Δt', 70, 31, 8, 4], ['∫ f(x) dx', 64, 37.2, 7, -3]]
    .forEach(([t, x, y, pt, r]) => T(f, t, { x: x * MM, y: y * MM, w: 22 * MM, h: 5 * MM, fontFace: 'Century Schoolbook',
      fontSize: pt, italic: true, color: 'F2F0E6', transparency: 80, rotate: r, objectName: 'formula ' + t }));
  T(f, 'Clases particulares y grupales', { x: L * MM, y: L * MM, w: 40 * MM, h: 3.5 * MM, fontSize: 7, color: '7FA796', bold: true, valign: 'top', objectName: 'servicio' });
  T(f, [
    { text: 'Despejamos', options: { breakLine: true } },
    { text: 'tus dudas.' }
  ], { x: L * MM, y: 12 * MM, w: 56 * MM, h: 18 * MM, fontFace: 'Century Schoolbook', fontSize: 23, bold: true, color: 'F2F0E6', lineSpacingMultiple: 0.9, valign: 'top', objectName: 'titular' });
  T(f, [
    { text: 'Matemática, Física y Química', options: { breakLine: true } },
    { text: 'Colegio, preuniversitario y universidad' }
  ], { x: L * MM, y: 32 * MM, w: 58 * MM, h: 7 * MM, fontSize: 7.5, color: 'F2F0E6', transparency: 12, valign: 'top', lineSpacingMultiple: 1.05, objectName: 'cursos' });
  // la x despejada: centrada en la columna derecha y a la altura del titular
  const XC = COL + COLW / 2, XY = 21;
  f.addShape(pres.shapes.OVAL, { x: (XC - 8.5) * MM, y: (XY - 7.5) * MM, w: 17 * MM, h: 15 * MM, rotate: -8, line: { color: 'F3D34A', width: 0.9 }, fill: { type: 'none' }, objectName: 'circulo tiza' });
  T(f, 'x', { x: (XC - 8.5) * MM, y: (XY - 8.5) * MM, w: 17 * MM, h: 15 * MM, fontFace: 'Century Schoolbook', italic: true, fontSize: 40, color: 'F3D34A', align: 'center', valign: 'middle', objectName: 'x' });
  // repisa de la pizarra
  f.addShape(pres.shapes.LINE, { x: L * MM, y: LINE_Y * MM, w: (R - L) * MM, h: 0, line: { color: '7FA796', width: 0.5, dashType: 'sysDot' }, objectName: 'repisa' });
  T(f, 'TU MARCA', { x: L * MM, y: ROW_Y * MM, w: 40 * MM, h: ROW_H * MM, fontFace: 'Century Schoolbook', fontSize: 11, bold: true, color: 'F3D34A', valign: 'middle', objectName: 'marca (pendiente)' });
  T(f, 'Profesores de Ingeniería PUCP', { x: 46 * MM, y: ROW_Y * MM, w: (R - 46) * MM, h: ROW_H * MM, fontSize: 7, color: 'F2F0E6', align: 'right', valign: 'middle', objectName: 'autoridad' });

  // ---------- REVERSO ----------
  const b = pres.addSlide();
  b.background = { color: 'FFFFFF' };
  T(b, 'Así trabajamos', { x: L * MM, y: L * MM, w: 55 * MM, h: 5.5 * MM, fontFace: 'Century Schoolbook', fontSize: 12, bold: true, color: '23443A', valign: 'top', objectName: 'titulo reverso' });
  const steps = [['1', 'Diagnóstico', 'ubicamos el tema exacto donde te trabas.'], ['2', 'Plan semanal', 'metas claras hasta tu examen.'], ['3', 'Seguimiento', 'reporte de avance para ti y tus papás.']];
  const D = 4.4; // diámetro del número
  steps.forEach(([n, k, d], i) => {
    const y = 16 + i * 9;
    b.addShape(pres.shapes.OVAL, { x: L * MM, y: y * MM, w: D * MM, h: D * MM, fill: { color: '23443A' }, line: { type: 'none' }, objectName: 'paso ' + n });
    T(b, n, { x: L * MM, y: y * MM, w: D * MM, h: D * MM, fontSize: 7, bold: true, color: 'F3D34A', align: 'center', valign: 'middle' });
    T(b, [{ text: k, options: { bold: true, color: '23443A', breakLine: true } }, { text: d, options: { color: '3B4A44' } }],
      { x: (L + 6.5) * MM, y: (y - 0.9) * MM, w: (62 - L - 6.5) * MM, h: 6.4 * MM, fontSize: 7, valign: 'middle', lineSpacingMultiple: 1.0 });
  });
  // oferta: plan estructurado (columna derecha, alineada arriba con el título)
  b.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: COL * MM, y: L * MM, w: COLW * MM, h: 15 * MM, rectRadius: 0.06, fill: { color: '23443A' }, line: { type: 'none' }, objectName: 'oferta (pendiente %)' });
  T(b, [{ text: 'Con plan mensual', options: { breakLine: true, fontSize: 6.5, color: 'F2F0E6' } },
        { text: '−20%', options: { breakLine: true, fontSize: 18, bold: true, color: 'F3D34A', fontFace: 'Century Schoolbook' } },
        { text: 'en cada clase', options: { fontSize: 6.5, color: 'F2F0E6' } }],
    { x: COL * MM, y: L * MM, w: COLW * MM, h: 15 * MM, align: 'center', valign: 'middle' });
  // QR centrado en la columna derecha
  const QS = 15;
  const qr = await sharp(fs.readFileSync('../propuesta-A-html/qr-whatsapp.svg'), { density: 1200 }).resize(600, 600).png().toBuffer();
  b.addImage({ data: 'image/png;base64,' + qr.toString('base64'), x: (XC - QS / 2) * MM, y: 24.5 * MM, w: QS * MM, h: QS * MM, altText: 'QR a WhatsApp', objectName: 'QR (pendiente)' });
  T(b, 'Escanea y escríbenos', { x: COL * MM, y: 40.2 * MM, w: COLW * MM, h: 2.6 * MM, fontSize: 6, color: '3E6B5A', align: 'center', valign: 'top' });
  // contacto: misma línea y fila inferior que el frente
  b.addShape(pres.shapes.LINE, { x: L * MM, y: LINE_Y * MM, w: (R - L) * MM, h: 0, line: { color: 'C9D3CE', width: 0.5 } });
  T(b, [{ text: 'WhatsApp  ', options: { color: '7FA796', bold: true } }, { text: '9XX XXX XXX', options: { color: '1D2B25', bold: true } }],
    { x: L * MM, y: ROW_Y * MM, w: 29 * MM, h: ROW_H * MM, fontSize: 7, valign: 'middle', objectName: 'whatsapp (pendiente)' });
  T(b, [{ text: 'Instagram  ', options: { color: '7FA796', bold: true } }, { text: '@tu_usuario', options: { color: '1D2B25', bold: true } }],
    { x: 37 * MM, y: ROW_Y * MM, w: 28 * MM, h: ROW_H * MM, fontSize: 7, valign: 'middle', objectName: 'instagram (pendiente)' });
  T(b, 'Máx. 4 por grupo', { x: COL * MM, y: ROW_Y * MM, w: COLW * MM, h: ROW_H * MM, fontSize: 6.5, color: 'C9452E', bold: true, align: 'right', valign: 'middle', objectName: 'grupos (pendiente)' });

  f.addNotes('Frente. Tamaño final 90x55 mm; el borde exterior de 3 mm es sangrado y se corta.');
  b.addNotes('Reverso. Reemplazar TU MARCA, WhatsApp, Instagram, % de descuento y QR.');
  await pres.writeFile({ fileName: 'tarjeta-pizarra.pptx' });
  await applyTheme('tarjeta-pizarra.pptx', THEME);
})();
