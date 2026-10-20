/**
 * Creador IG · MÖÖnna — backend en Apps Script.
 *
 * Qué hace:
 *   - Guarda ideas y grabaciones en una Google Sheet y los audios en una carpeta de Drive.
 *   - Genera ideas y transcribe audios con Gemini (nivel gratis).
 *   - Expone una API JSON que usa la app web (GitHub Pages) desde el iPhone.
 *
 * Primera vez (una sola vez, desde el editor de Apps Script):
 *   1. Proyecto → Configuración → Propiedades del script → agrega GEMINI_API_KEY.
 *   2. Ejecuta la función `instalar` y acepta los permisos.
 *   3. Copia el CÓDIGO que imprime en el registro: lo pide la app la primera vez que la abres.
 *   4. Implementar → Nueva implementación → Aplicación web → "Cualquier usuario" → copia la URL /exec.
 */

const PROPS = PropertiesService.getScriptProperties();
const TZ = 'America/Mexico_City';
// gemini-flash-latest se satura seguido en el nivel gratis: alternamos con la versión lite.
const MODELOS = ['gemini-flash-latest', 'gemini-flash-lite-latest', 'gemini-flash-latest', 'gemini-flash-lite-latest'];

const COLUMNAS = {
  ideas: ['id', 'creada', 'estado', 'titulo', 'pilar', 'gancho', 'guion', 'duracion_seg', 'estilo', 'visuales', 'por_que_funciona'],
  proyectos: ['id', 'creado', 'estado', 'idea_id', 'titulo', 'estilo', 'guion', 'archivo_id', 'mime', 'nombre_archivo', 'duracion', 'texto', 'error'],
};

const SISTEMA = `Eres el estratega de contenido de Instagram de la marca descrita abajo.
Propones ideas para Reels de 20 a 45 segundos que el dueño grabará SOLO CON SU VOZ
(el video se arma después con subtítulos y visuales automáticos).

Reglas:
- El gancho debe atrapar en los primeros 2 segundos.
- El guion se lee en voz alta de forma natural, en español de México, frases cortas.
- Cada guion cierra con el llamado a la acción de la marca.
- No repitas ideas ya usadas.
- Responde SOLO con un arreglo JSON, sin texto antes ni después.

Cada elemento del arreglo:
{"titulo": str, "pilar": str, "gancho": str, "guion": str,
 "duracion_seg": int, "estilo": str, "visuales": str, "por_que_funciona": str}

"estilo" es uno de: "stock" (clips de stock con subtítulos grandes), "ia" (imágenes
generadas con IA), "texto" (fondo de color con texto animado) o "propio" (fotos/videos
que el dueño sube). Elige el que mejor sirva a cada idea y varía entre ideas.
"visuales" describe en una línea qué se debería ver (sirve para el editor automático).

=== PERFIL DE MARCA ===
`;

const MARCA_INICIAL = `# Perfil de marca: MÖÖnna

Edita esta celda cuando quieras: es lo que la IA lee para proponerte ideas.

## Qué es MÖÖnna
Empresa de tecnología que crea soluciones hiper-automatizadas con herramientas de vanguardia (IA, bots,
automatizaciones, apps a la medida) para cualquier rubro.
Lema: "Si lo sueñas, se puede crear."
La cuenta la lleva su fundador, con voz en primera persona: alguien que de verdad construye estas cosas
y lo cuenta sin rollo técnico.

## A quién le hablamos
Profesionistas, emprendedores y dueños de negocios de cualquier giro (no importa la edad) que sienten que
la tecnología avanza más rápido que ellos. Les preocupa quedarse desactualizados, perder tiempo en tareas
repetitivas o que la competencia los rebase. Quieren saber qué pueden aplicar YA, sin ser programadores.

## Objetivo de la cuenta
Todo a la vez: entretener, construir comunidad, posicionar a MÖÖnna como referente en hiper-automatización
y generar clientes. Primero que entretenga y aporte; la venta llega sola.

## Tono de voz
Relajado, con humor, como cotorreando con un amigo, pero directo: va al punto, sin rodeos ni tecnicismos.
Español de México, coloquial. Se valen groserías moderadas de vez en cuando (no en cada video).
Divertirse con la tecnología es parte del mensaje: nada de tono de catedrático ni de "gurú".

## Pilares de contenido
1. Entretenimiento con IA (pilar principal): situaciones chistosas, memes, "la IA ya es parte de tu vida
   y ni cuenta te diste", comparaciones del antes vs. ahora.
2. "Esto ya se puede automatizar": tareas de un negocio o profesión que hoy se resuelven solas
   (en cualquier rubro: doctores, abogados, tiendas, restaurantes, contadores...).
3. IA en la vida diaria / estilo de vida: cómo la IA ya está en lo cotidiano y cómo aprovecharla.
4. Detrás de cámaras de MÖÖnna: lo que construimos, proyectos, retos y aprendizajes.
5. Noticias y herramientas nuevas: lo nuevo explicado en cristiano y qué significa para tu negocio.

## Lo que NO queremos
- No mostrar la cara (por ahora): todo con voz en off. El material propio son grabaciones de pantalla,
  manos, escritorio, oficina, demos de lo que construimos; nunca el rostro.
- Nada de tono de vendedor agresivo ni promesas de "hazte millonario con IA".
- Sin tecnicismos sin explicar.

## Llamado a la acción preferido
Principalmente: "Sígueme" y "Escríbeme por DM" (p. ej. "mándame DM con la palabra AUTOMATIZA y te cuento
cómo aplicarlo en tu negocio"). De vez en cuando: "guárdalo" o "mándaselo a alguien que lo necesita".`;

// ============================================================
// Instalación (se corre una vez desde el editor)
// ============================================================

function instalar() {
  let sheetId = PROPS.getProperty('SHEET_ID');
  if (!sheetId) {
    const ss = SpreadsheetApp.create('Creador IG · MÖÖnna');
    for (const [nombre, cols] of Object.entries(COLUMNAS)) {
      const h = ss.insertSheet(nombre);
      h.getRange('A:Z').setNumberFormat('@'); // todo como texto: fechas e ids no se deforman
      h.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
      h.setFrozenRows(1);
    }
    const m = ss.insertSheet('marca');
    m.getRange('A1').setValue(MARCA_INICIAL).setWrap(true).setVerticalAlignment('top');
    m.setColumnWidth(1, 720);
    ss.deleteSheet(ss.getSheets()[0]); // la "Hoja 1" vacía
    sheetId = ss.getId();
    PROPS.setProperty('SHEET_ID', sheetId);
  }

  let folderId = PROPS.getProperty('FOLDER_ID');
  if (!folderId) {
    const f = DriveApp.createFolder('Creador IG · MÖÖnna');
    f.createFolder('audios');
    f.createFolder('videos');
    folderId = f.getId();
    PROPS.setProperty('FOLDER_ID', folderId);
  }

  if (!PROPS.getProperty('TOKEN')) {
    PROPS.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, '').slice(0, 20));
  }

  Logger.log('Listo.\n' +
    'Hoja: https://docs.google.com/spreadsheets/d/' + sheetId + '\n' +
    'Carpeta: https://drive.google.com/drive/folders/' + folderId + '\n' +
    'CÓDIGO para la app: ' + PROPS.getProperty('TOKEN') + '\n' +
    (PROPS.getProperty('GEMINI_API_KEY') ? 'Gemini: configurado' : 'FALTA GEMINI_API_KEY en Propiedades del script'));
}

// ============================================================
// API HTTP
// ============================================================

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!autorizado_(p.token)) return json_({ error: 'Código incorrecto' });
  try {
    switch (p.accion) {
      case 'ping': {
        const k = claveGemini_();
        return json_({ ok: true, gemini: !!k, clave: k ? k.slice(0, 4) + '…' + k.slice(-3) + ' (' + k.length + ')' : '' });
      }
      case 'ideas':
        return json_(tabla_('ideas').reverse()); // las más nuevas primero
      case 'proyectos':
        return json_(tabla_('proyectos').reverse());
      case 'audio': {
        const pr = fila_('proyectos', p.id);
        if (!pr || !pr.archivo_id) return json_({ error: 'Sin audio' });
        const f = DriveApp.getFileById(pr.archivo_id);
        return json_({ mime: f.getMimeType(), base64: Utilities.base64Encode(f.getBlob().getBytes()) });
      }
      default:
        return json_({ error: 'Acción desconocida: ' + p.accion });
    }
  } catch (err) {
    return json_({ error: String(err.message || err) });
  }
}

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (!autorizado_(p.token)) return json_({ error: 'Código incorrecto' });
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (_) { return json_({ error: 'Cuerpo inválido' }); }
  try {
    switch (b.accion) {
      case 'generar_ideas':
        return json_(generarIdeas_(b.tema || '', Math.min(10, Math.max(1, Number(b.cantidad) || 5))));
      case 'estado_idea': {
        if (!['pendiente', 'usada', 'descartada'].includes(b.estado)) return json_({ error: 'Estado inválido' });
        if (!actualizar_('ideas', b.id, { estado: b.estado })) return json_({ error: 'Idea no encontrada' });
        return json_({ ok: true });
      }
      case 'subir_audio':
        return json_(subirAudio_(b));
      case 'borrar_proyecto': {
        const pr = fila_('proyectos', b.id);
        if (!pr) return json_({ error: 'Grabación no encontrada' });
        if (pr.estado === 'transcribiendo') return json_({ error: 'Espera a que termine de transcribir' });
        if (pr.archivo_id) { try { DriveApp.getFileById(pr.archivo_id).setTrashed(true); } catch (_) {} }
        borrar_('proyectos', b.id);
        return json_({ ok: true });
      }
      default:
        return json_({ error: 'Acción desconocida: ' + b.accion });
    }
  } catch (err) {
    return json_({ error: String(err.message || err) });
  }
}

function autorizado_(token) {
  const esperado = PROPS.getProperty('TOKEN');
  return !!esperado && token === esperado;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// Ideas
// ============================================================

function generarIdeas_(tema, cantidad) {
  const previas = tabla_('ideas').slice(-40).map(i => i.titulo); // las últimas 40 = las más recientes
  let instruccion = 'Dame ' + cantidad + ' ideas nuevas.';
  if (tema.trim()) instruccion += ' Enfócate en: ' + tema.trim() + '.';
  if (previas.length) instruccion += '\nYa usadas o propuestas (no repetir):\n- ' + previas.join('\n- ');

  const texto = preguntarIA_(SISTEMA + marca_(), instruccion);
  const ini = texto.indexOf('['), fin = texto.lastIndexOf(']');
  if (ini < 0 || fin < 0) throw new Error('La IA no devolvió una lista. Intenta de nuevo.');
  const nuevas = JSON.parse(texto.slice(ini, fin + 1));

  const ahora = ahora_();
  for (const idea of nuevas) {
    idea.id = Utilities.getUuid().slice(0, 8);
    idea.creada = ahora;
    idea.estado = 'pendiente';
    agregar_('ideas', idea);
  }
  return nuevas;
}

function marca_() {
  return String(hoja_('marca').getRange('A1').getValue() || MARCA_INICIAL);
}

// ============================================================
// Grabaciones
// ============================================================

function subirAudio_(b) {
  if (!b.base64 || b.base64.length < 1000) throw new Error('El audio está vacío o es demasiado corto');
  const idea = b.idea_id ? fila_('ideas', b.idea_id) : null;
  const nombre = b.nombre || 'grabacion.wav';
  const mime = b.mime || 'audio/wav';
  const bytes = Utilities.base64Decode(b.base64);

  const id = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 4);
  const archivo = subcarpeta_('audios').createFile(Utilities.newBlob(bytes, mime, id + '-' + nombre));

  const proyecto = {
    id: id,
    creado: ahora_(),
    estado: 'transcribiendo',
    idea_id: idea ? idea.id : '',
    titulo: idea ? idea.titulo : 'Grabación libre',
    estilo: idea ? idea.estilo : '',
    guion: idea ? idea.guion : '',
    archivo_id: archivo.getId(),
    mime: mime,
    nombre_archivo: nombre,
    duracion: mime === 'audio/wav' ? duracionWav_(bytes) : '',
    texto: '',
    error: '',
  };
  agregar_('proyectos', proyecto);
  if (idea) actualizar_('ideas', idea.id, { estado: 'usada' });

  // La transcripción corre aparte para que la app no se quede esperando.
  ScriptApp.newTrigger('procesarPendientes').timeBased().after(1000).create();
  return proyecto;
}

function procesarPendientes() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // ya hay otra corrida en curso
  try {
    // Limpia los disparadores de un solo uso que ya cumplieron.
    for (const t of ScriptApp.getProjectTriggers()) {
      if (t.getHandlerFunction() === 'procesarPendientes') ScriptApp.deleteTrigger(t);
    }
    for (const pr of tabla_('proyectos')) {
      if (pr.estado === 'transcribiendo') transcribir_(pr);
    }
  } finally {
    lock.releaseLock();
  }
}

function transcribir_(pr) {
  try {
    const f = DriveApp.getFileById(pr.archivo_id);
    const mime = mimeGemini_(pr.mime || f.getMimeType());
    const prompt = 'Transcribe literalmente este audio, en español de México, respetando muletillas y ' +
      'groserías tal como se dicen. No resumas ni corrijas. Pistas de contexto (nombres y palabras ' +
      'que pueden aparecer): MÖÖnna, automatizar, IA, bot, DM. ' +
      (pr.guion ? 'El hablante leía más o menos este guion: «' + String(pr.guion).slice(0, 700) + '». ' : '') +
      'Responde SOLO con JSON: {"texto": "...", "duracion_seg": número}';

    const texto = preguntarIA_(null, null, [
      { inlineData: { mimeType: mime, data: Utilities.base64Encode(f.getBlob().getBytes()) } },
      { text: prompt },
    ]);
    const ini = texto.indexOf('{'), fin = texto.lastIndexOf('}');
    const r = JSON.parse(texto.slice(ini, fin + 1));
    actualizar_('proyectos', pr.id, {
      estado: 'transcrito',
      texto: String(r.texto || '').trim(),
      duracion: pr.duracion || (r.duracion_seg ? Math.round(Number(r.duracion_seg)) : ''),
      error: '',
    });
  } catch (err) {
    actualizar_('proyectos', pr.id, { estado: 'error', error: String(err.message || err).slice(0, 400) });
  }
}

function mimeGemini_(mime) {
  const mapa = { 'audio/x-m4a': 'audio/mp4', 'audio/m4a': 'audio/mp4', 'audio/mpeg': 'audio/mp3', 'audio/x-wav': 'audio/wav' };
  return mapa[mime] || mime;
}

function duracionWav_(bytes) {
  // WAV PCM 16-bit mono a 16 kHz (así lo manda la app): (bytes - cabecera) / (16000 * 2)
  return Math.round((bytes.length - 44) / 32000);
}

// ============================================================
// Gemini
// ============================================================

/** preguntarIA_(sistema, mensaje) para texto, o preguntarIA_(null, null, partes) para multimodal. */
/** Lee la clave tolerando que se haya pegado con "GEMINI_API_KEY=", comillas o espacios. */
function claveGemini_() {
  let k = String(PROPS.getProperty('GEMINI_API_KEY') || '').trim();
  k = k.replace(/^GEMINI_API_KEY\s*=\s*/i, '').replace(/^["']|["']$/g, '').trim();
  return k;
}

function preguntarIA_(sistema, mensaje, partes) {
  const clave = claveGemini_();
  if (!clave) throw new Error('Falta GEMINI_API_KEY en Propiedades del script');

  const cuerpo = {
    contents: [{ role: 'user', parts: partes || [{ text: mensaje }] }],
    generationConfig: { responseMimeType: 'application/json' },
  };
  if (sistema) cuerpo.systemInstruction = { parts: [{ text: sistema }] };
  const payload = JSON.stringify(cuerpo);

  let ultimo = null;
  for (let n = 0; n < MODELOS.length; n++) {
    const r = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/' + MODELOS[n] + ':generateContent',
      { method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': clave }, payload: payload, muteHttpExceptions: true }
    );
    const code = r.getResponseCode();
    if (code === 200) {
      const d = JSON.parse(r.getContentText());
      return d.candidates[0].content.parts.map(p => p.text || '').join('');
    }
    ultimo = code + ': ' + r.getContentText().slice(0, 300);
    if (![429, 500, 503].includes(code)) break;
    Utilities.sleep(2000 * (n + 1));
  }
  if (/^(429|503)/.test(ultimo)) throw new Error('Gemini está saturado en este momento. Intenta en un minuto.');
  throw new Error('Gemini respondió ' + ultimo);
}

// ============================================================
// Sheets como base de datos
// ============================================================

function hoja_(nombre) {
  return SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID')).getSheetByName(nombre);
}

function tabla_(nombre) {
  const v = hoja_(nombre).getDataRange().getValues();
  const cab = v[0];
  return v.slice(1).filter(r => r[0] !== '').map(r => {
    const o = {};
    cab.forEach((c, i) => { o[c] = r[i] instanceof Date ? Utilities.formatDate(r[i], TZ, "yyyy-MM-dd'T'HH:mm:ss") : r[i]; });
    return o;
  });
}

function fila_(nombre, id) {
  return tabla_(nombre).find(r => String(r.id) === String(id)) || null;
}

function agregar_(nombre, obj) {
  const h = hoja_(nombre);
  const cab = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0];
  h.appendRow(cab.map(c => (obj[c] === undefined || obj[c] === null) ? '' : obj[c]));
}

function actualizar_(nombre, id, cambios) {
  const h = hoja_(nombre);
  const v = h.getDataRange().getValues();
  const cab = v[0];
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(id)) {
      for (const [k, val] of Object.entries(cambios)) {
        const c = cab.indexOf(k);
        if (c >= 0) h.getRange(i + 1, c + 1).setValue(val);
      }
      return true;
    }
  }
  return false;
}

function borrar_(nombre, id) {
  const h = hoja_(nombre);
  const v = h.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]) === String(id)) { h.deleteRow(i + 1); return true; }
  }
  return false;
}

function subcarpeta_(nombre) {
  const raiz = DriveApp.getFolderById(PROPS.getProperty('FOLDER_ID'));
  const it = raiz.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : raiz.createFolder(nombre);
}

function ahora_() {
  return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
}
