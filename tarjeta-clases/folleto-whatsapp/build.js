// Folleto de una sola cara para enviar por WhatsApp: imagen vertical 1080 x 1920 px.
// Todas las medidas están en px de pantalla; 1 px = 1/96 pulg. y las fuentes se convierten px -> pt (x 0,75).
const pptxgen = require('pptxgenjs');

const P = v => v / 96, PT = v => v * 0.75;
const COND = 'Barlow Condensed', BODY = 'Barlow';
const GREEN = '23443A', YELLOW = 'F3D34A', WHITE = 'FFFFFF', SOFT = '56606E', MUTED = '7FA796', CHALK = 'E2E8E4', EDGE = 'A9BDB4';

const pres = new pptxgen();
pres.defineLayout({ name: 'WHATSAPP', width: P(1080), height: P(1920) });
pres.layout = 'WHATSAPP';
pres.title = 'Despeja: folleto para WhatsApp';
const s = pres.addSlide();
s.background = { color: WHITE };

const box = (x, y, w, h) => ({ x: P(x), y: P(y), w: P(w), h: P(h) });
const rect = (x, y, w, h, fill, name) => s.addShape(pres.shapes.RECTANGLE, { ...box(x, y, w, h), fill: { color: fill }, line: { type: 'none' }, objectName: name });
const round = (x, y, w, h, fill, r, name, line) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { ...box(x, y, w, h), rectRadius: P(r), fill: fill ? { color: fill } : { type: 'none' }, line: line || { type: 'none' }, objectName: name });
const dot = (x, y, d, fill, edge, name) => s.addShape(pres.shapes.OVAL, { ...box(x, y, d, d), fill: { color: fill }, line: { color: edge, width: 1 }, objectName: name });
const run = (t, o = {}) => ({ text: t, options: { ...o, ...(o.fontSize ? { fontSize: PT(o.fontSize) } : {}) } });
const text = (content, x, y, w, h, o = {}) => s.addText(content, { ...box(x, y, w, h), isTextBox: true, margin: 0, fontFace: BODY, valign: 'top', ...o, ...(o.fontSize ? { fontSize: PT(o.fontSize) } : {}) });

// ---------- 1. Cabecera amarilla: logotipo
rect(0, 0, 1080, 210, YELLOW, 'cabecera');
text('Δ', 48, 12, 200, 186, { fontFace: COND, bold: true, fontSize: 232, color: GREEN, align: 'center', valign: 'middle', objectName: 'logo delta' });
text('Despeja', 250, 4, 760, 126, { fontFace: COND, bold: true, fontSize: 126, color: GREEN, valign: 'middle', objectName: 'logo nombre' });
text('clases particulares y grupales', 256, 148, 750, 44, { bold: true, fontSize: 36, color: GREEN, objectName: 'logo lema' });

// ---------- 2. Titular
rect(0, 210, 1080, 486, GREEN, 'titular · fondo');
text([run('La nota de su', { color: WHITE, breakLine: true }), run('hijo(a) también', { color: WHITE, breakLine: true }), run('puede ', { color: WHITE }), run('cambiar.', { color: YELLOW })], 72, 244, 936, 340,
  { fontFace: COND, bold: true, fontSize: 108, lineSpacingMultiple: 0.92, objectName: 'titular' });
let px = 72;
['Secundaria', 'Pre', 'Universidad'].forEach(t => {
  const w = t.length * 19 + 64;
  round(px, 622, w, 56, null, 28, 'nivel ' + t, { color: EDGE, width: 1.5 });
  text(t, px, 622, w, 56, { bold: true, fontSize: 34, color: WHITE, align: 'center', valign: 'middle' });
  px += w + 20;
});

// ---------- 3. Método
text('Así acompañamos a su hijo(a)', 72, 726, 936, 72, { fontFace: COND, bold: true, fontSize: 64, color: GREEN, valign: 'middle', objectName: 'método · título' });
[['1', 'Diagnóstico', 'Dónde se traba.'], ['2', 'Plan semanal', 'Metas semanales.'], ['3', 'Reporte a padres', 'Avance mensual.']].forEach(([k, t, d], i) => {
  const x = 72 + i * 326;
  round(x, 818, 72, 72, GREEN, 36, 'método · círculo ' + k);
  text(k, x, 818, 72, 72, { fontFace: COND, bold: true, fontSize: 46, color: YELLOW, align: 'center', valign: 'middle' });
  text(t, x, 904, 300, 48, { fontFace: COND, bold: true, fontSize: 40, color: GREEN, valign: 'middle', objectName: 'método · paso ' + k });
  text(d, x, 954, 300, 40, { fontSize: 30, color: SOFT, objectName: 'método · detalle ' + k });
});

// ---------- 4. Cursos
text('Qué enseñamos', 72, 1018, 936, 72, { fontFace: COND, bold: true, fontSize: 64, color: GREEN, valign: 'middle', objectName: 'cursos · título' });
round(72, 1106, 512, 298, YELLOW, 28, 'cursos · bloque secundaria y pre');
text('Secundaria y pre', 100, 1128, 460, 62, { fontFace: COND, bold: true, fontSize: 52, color: GREEN, valign: 'middle' });
[['Álgebra', 'Aritmética', 'Geometría'], ['Trigonometría', 'Física', 'Química']].forEach((col, c) => col.forEach((t, r) => {
  const x = 100 + c * 245, y = 1202 + r * 46;
  dot(x, y + 17, 14, GREEN, GREEN, 'viñeta ' + t);
  text(t, x + 30, y, 215, 44, { fontFace: COND, bold: true, fontSize: 36, color: GREEN, valign: 'middle' });
}));
round(608, 1106, 400, 298, GREEN, 28, 'cursos · bloque universidad');
text('Universidad', 636, 1128, 340, 62, { fontFace: COND, bold: true, fontSize: 52, color: YELLOW, valign: 'middle' });
['Cálculo', 'Física', 'Química', 'Cursos de Ingeniería'].forEach((t, r) => {
  const y = 1202 + r * 46;
  dot(636, y + 17, 14, YELLOW, YELLOW, 'viñeta ' + t);
  text(t, 666, y, 320, 44, { fontFace: COND, bold: true, fontSize: 36, color: WHITE, valign: 'middle' });
});

// ---------- 5. Oferta
rect(0, 1432, 1080, 128, YELLOW, 'oferta · fondo');
text('10%', 72, 1432, 270, 128, { fontFace: COND, bold: true, fontSize: 120, color: GREEN, valign: 'middle', objectName: 'oferta · 10%' });
text('de descuento con el plan mensual', 350, 1444, 660, 58, { fontFace: COND, bold: true, fontSize: 50, color: GREEN, valign: 'middle', objectName: 'oferta · descuento' });
text('Facilidades de pago · Máx. 4 por grupo', 350, 1504, 660, 40, { bold: true, fontSize: 30, color: GREEN, objectName: 'oferta · detalle' });

// ---------- 6. Llamado a la acción
rect(0, 1560, 1080, 360, GREEN, 'contacto · fondo');
text([run('Responda este mensaje y coordinamos el', { breakLine: true }), run('diagnóstico de su hijo(a).')], 72, 1588, 936, 90, { bold: true, fontSize: 38, color: WHITE, lineSpacingMultiple: 1.05, objectName: 'contacto · llamado' });
text([run('WhatsApp', { fontSize: 28, color: MUTED, bold: true, breakLine: true }), run('+51 961 956 660', { fontFace: COND, fontSize: 56, color: YELLOW, bold: true })], 72, 1704, 440, 96, { objectName: 'contacto · WhatsApp' });
text([run('Instagram', { fontSize: 28, color: MUTED, bold: true, breakLine: true }), run('@despeja.clases', { fontFace: COND, fontSize: 56, color: YELLOW, bold: true })], 560, 1704, 448, 96, { objectName: 'contacto · Instagram' });
s.addShape(pres.shapes.LINE, { ...box(72, 1814, 936, 0), line: { color: MUTED, width: 1, dashType: 'sysDot' }, objectName: 'contacto · línea' });
text([run('Profesores de Ingeniería de la', { breakLine: true }), run('Pontificia Universidad Católica del Perú', { bold: true })], 72, 1830, 936, 68, { fontSize: 28, color: CHALK, objectName: 'contacto · respaldo' });

s.addNotes('Imagen de 1080 x 1920 px para enviar por WhatsApp. Confirmar el usuario de Instagram antes de enviar.');
pres.writeFile({ fileName: 'folleto-whatsapp.pptx' });
