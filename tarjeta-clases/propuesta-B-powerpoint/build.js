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

  // ---------- FRENTE ----------
  const f = pres.addSlide();
  f.background = { color: '23443A' };
  // fórmulas en tiza tenue de fondo
  [['F = m·a', 52, 6, 9, -6], ['Δx/Δt', 74, 36, 8, 4], ['a² + b² = c²', 40, 37, 7, -3], ['∫ f(x) dx', 56, 2.5, 7, 6], ['pH = −log[H⁺]', 4, 2.5, 6, 0]]
    .forEach(([t, x, y, pt, r]) => T(f, t, { x: x * MM, y: y * MM, w: 30 * MM, h: 6 * MM, fontFace: 'Century Schoolbook',
      fontSize: pt, italic: true, color: 'F2F0E6', transparency: 82, rotate: r, objectName: 'formula ' + t }));
  T(f, 'Clases particulares y grupales', { x: S, y: 9 * MM, w: 60 * MM, h: 4 * MM, fontSize: 7, color: '7FA796', bold: true, objectName: 'servicio' });
  T(f, [
    { text: 'Despejamos', options: { breakLine: true } },
    { text: 'tus dudas.' }
  ], { x: S, y: 14 * MM, w: 60 * MM, h: 20 * MM, fontFace: 'Century Schoolbook', fontSize: 23, bold: true, color: 'F2F0E6', lineSpacingMultiple: 0.9, valign: 'top', objectName: 'titular' });
  // la x despejada, en tiza amarilla
  T(f, 'x', { x: 70 * MM, y: 13.5 * MM, w: 18 * MM, h: 17 * MM, fontFace: 'Century Schoolbook', italic: true, fontSize: 40, color: 'F3D34A', align: 'center', valign: 'middle', objectName: 'x' });
  f.addShape(pres.shapes.OVAL, { x: 70.5 * MM, y: 14.5 * MM, w: 17 * MM, h: 15 * MM, rotate: -8, line: { color: 'F3D34A', width: 0.9 }, fill: { type: 'none' }, objectName: 'circulo tiza' });
  T(f, 'Matemática, Física y Química para colegio, pre y universidad.', { x: S, y: 33.5 * MM, w: 56 * MM, h: 7 * MM, fontSize: 7.5, color: 'F2F0E6', transparency: 15, objectName: 'cursos' });
  // borde inferior de la pizarra: la repisa de tiza
  f.addShape(pres.shapes.LINE, { x: S, y: 45 * MM, w: W - 2 * S, h: 0, line: { color: '7FA796', width: 0.5, dashType: 'sysDot' }, objectName: 'repisa' });
  T(f, 'TU MARCA', { x: S, y: 47 * MM, w: 40 * MM, h: 6 * MM, fontFace: 'Century Schoolbook', fontSize: 11, bold: true, color: 'F3D34A', objectName: 'marca (pendiente)' });
  T(f, 'Profesores de Ingeniería PUCP', { x: 46 * MM, y: 48.2 * MM, w: 43 * MM, h: 4 * MM, fontSize: 7, color: 'F2F0E6', align: 'right', objectName: 'autoridad' });

  // ---------- REVERSO ----------
  const b = pres.addSlide();
  b.background = { color: 'FFFFFF' };
  T(b, 'Así trabajamos', { x: S, y: S, w: 50 * MM, h: 6 * MM, fontFace: 'Century Schoolbook', fontSize: 12, bold: true, color: '23443A', objectName: 'titulo reverso' });
  const steps = [['1', 'Diagnóstico', 'ubicamos el tema exacto donde te trabas.'], ['2', 'Plan semanal', 'metas claras hasta tu examen.'], ['3', 'Seguimiento', 'reporte de avance para ti y tus papás.']];
  steps.forEach(([n, k, d], i) => {
    const y = (15.5 + i * 7) * MM;
    b.addShape(pres.shapes.OVAL, { x: S, y, w: 4.4 * MM, h: 4.4 * MM, fill: { color: '23443A' }, line: { type: 'none' }, objectName: 'paso ' + n });
    T(b, n, { x: S, y, w: 4.4 * MM, h: 4.4 * MM, fontSize: 7, bold: true, color: 'F3D34A', align: 'center', valign: 'middle' });
    T(b, [{ text: k + ': ', options: { bold: true, color: '23443A' } }, { text: d, options: { color: '3B4A44' } }],
      { x: S + 6 * MM, y: y - 0.3 * MM, w: 52 * MM, h: 5 * MM, fontSize: 7, valign: 'middle' });
  });
  // oferta: plan estructurado
  b.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 66 * MM, y: S, w: 23 * MM, h: 17 * MM, rectRadius: 0.06, fill: { color: '23443A' }, line: { type: 'none' }, objectName: 'oferta (pendiente %)' });
  T(b, [{ text: 'Con plan mensual', options: { breakLine: true, fontSize: 6.5, color: 'F2F0E6' } },
        { text: '−20%', options: { breakLine: true, fontSize: 18, bold: true, color: 'F3D34A', fontFace: 'Century Schoolbook' } },
        { text: 'en cada clase', options: { fontSize: 6.5, color: 'F2F0E6' } }],
    { x: 66 * MM, y: S + 0.8 * MM, w: 23 * MM, h: 15.4 * MM, align: 'center', valign: 'middle' });
  // QR
  const qr = await sharp(fs.readFileSync('../propuesta-A-html/qr-whatsapp.svg'), { density: 1200 }).resize(600, 600).png().toBuffer();
  b.addImage({ data: 'image/png;base64,' + qr.toString('base64'), x: 69.5 * MM, y: 26 * MM, w: 16 * MM, h: 16 * MM, altText: 'QR a WhatsApp', objectName: 'QR (pendiente)' });
  T(b, 'Escanea y escríbenos', { x: 64 * MM, y: 42.3 * MM, w: 27 * MM, h: 3 * MM, fontSize: 6, color: '3E6B5A', align: 'center' });
  // contacto
  b.addShape(pres.shapes.LINE, { x: S, y: 46.5 * MM, w: W - 2 * S, h: 0, line: { color: 'C9D3CE', width: 0.5 } });
  T(b, [{ text: 'WhatsApp  ', options: { color: '7FA796', bold: true } }, { text: '9XX XXX XXX', options: { color: '1D2B25', bold: true } }],
    { x: S, y: 48 * MM, w: 32 * MM, h: 5 * MM, fontSize: 7, valign: 'middle', objectName: 'whatsapp (pendiente)' });
  T(b, [{ text: 'Instagram  ', options: { color: '7FA796', bold: true } }, { text: '@tu_usuario', options: { color: '1D2B25', bold: true } }],
    { x: 39 * MM, y: 48 * MM, w: 28 * MM, h: 5 * MM, fontSize: 7, valign: 'middle', objectName: 'instagram (pendiente)' });
  T(b, 'Máx. 4 por grupo', { x: 66 * MM, y: 48 * MM, w: 23 * MM, h: 5 * MM, fontSize: 6.5, color: 'C9452E', bold: true, align: 'right', valign: 'middle', objectName: 'grupos (pendiente)' });

  f.addNotes('Frente. Tamaño final 90x55 mm; el borde exterior de 3 mm es sangrado y se corta.');
  b.addNotes('Reverso. Reemplazar TU MARCA, WhatsApp, Instagram, % de descuento y QR.');
  await pres.writeFile({ fileName: 'tarjeta-pizarra.pptx' });
  await applyTheme('tarjeta-pizarra.pptx', THEME);
})();
