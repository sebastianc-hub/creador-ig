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
  proyectos: ['id', 'creado', 'estado', 'idea_id', 'titulo', 'estilo', 'guion', 'archivo_id', 'mime', 'nombre_archivo', 'duracion', 'texto', 'error',
              'video_id', 'caption', 'visuales', 'gancho',
              'ig_container', 'ig_media_id', 'permalink', 'publicado_en'],
};

// Estados de un proyecto:
//   transcribiendo → transcrito → por_editar → editando → listo → publicando → publicado
//   y "error" / "error_video" / "error_publicar" cuando algo falla.

const GRAPH = 'https://graph.facebook.com/v23.0';

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
  asegurarColumnas_();

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
        return json_({ ok: true, gemini: !!k, clave: k ? k.slice(0, 4) + '…' + k.slice(-3) + ' (' + k.length + ')' : '',
                       instagram: PROPS.getProperty('IG_USERNAME') || '' });
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
      case 'cola_render': // lo consulta la PC
        return json_({ pendientes: tabla_('proyectos').filter(x => x.estado === 'por_editar'), marca: marca_() });
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
      case 'estado_proyecto': { // teléfono: "Armar video" / "Rehacer"; PC: editando / error_video
        const permitidos = ['por_editar', 'editando', 'error_video', 'transcrito', 'listo'];
        if (!permitidos.includes(b.estado)) return json_({ error: 'Estado inválido' });
        const cambios = { estado: b.estado, error: b.error || '' };
        if (!actualizar_('proyectos', b.id, cambios)) return json_({ error: 'Grabación no encontrada' });
        return json_({ ok: true });
      }
      case 'subir_video':
        return json_(subirVideo_(b));
      case 'guardar_caption': {
        if (!actualizar_('proyectos', b.id, { caption: String(b.caption || '').slice(0, 2200) })) return json_({ error: 'Grabación no encontrada' });
        return json_({ ok: true });
      }
      case 'publicar':
        return json_(pedirPublicacion_(b.id));
      case 'borrar_proyecto': {
        const pr = fila_('proyectos', b.id);
        if (!pr) return json_({ error: 'Grabación no encontrada' });
        if (pr.estado === 'transcribiendo') return json_({ error: 'Espera a que termine de transcribir' });
        for (const fid of [pr.archivo_id, pr.video_id]) {
          if (fid) { try { DriveApp.getFileById(fid).setTrashed(true); } catch (_) {} }
        }
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

// Fechas (mes-día) que vale la pena aprovechar con contenido o promociones.
const FESTIVIDADES = [
  ['01-06', 'Día de Reyes'], ['02-14', 'San Valentín'], ['03-08', 'Día de la Mujer'],
  ['04-30', 'Día del Niño'], ['05-10', 'Día de las Madres (México)'], ['05-15', 'Día del Maestro'],
  ['06-15', 'Día del Padre (tercer domingo de junio)'], ['09-15', 'Fiestas Patrias / Grito'],
  ['10-31', 'Halloween'], ['11-02', 'Día de Muertos'], ['11-14', 'El Buen Fin'],
  ['11-28', 'Black Friday / Cyber Monday'], ['12-12', 'Posadas y cierre de año'], ['12-24', 'Navidad'],
  ['12-31', 'Año Nuevo'],
];

function festividadesProximas_(dias) {
  const hoy = new Date();
  const out = [];
  for (const [md, nombre] of FESTIVIDADES) {
    for (const anio of [hoy.getFullYear(), hoy.getFullYear() + 1]) {
      const f = new Date(anio, Number(md.slice(0, 2)) - 1, Number(md.slice(3)));
      const faltan = Math.round((f - hoy) / 86400000);
      if (faltan >= -1 && faltan <= dias) out.push(nombre + ' (en ' + Math.max(0, faltan) + ' días)');
    }
  }
  return out;
}

function generarIdeas_(tema, cantidad) {
  const previas = tabla_('ideas').slice(-40).map(i => i.titulo); // las últimas 40 = las más recientes
  let instruccion = 'Hoy es ' + Utilities.formatDate(new Date(), TZ, 'EEEE d MMMM yyyy') + '. Dame ' + cantidad + ' ideas nuevas.';
  if (tema.trim()) instruccion += ' Enfócate en: ' + tema.trim() + '.';
  const fiestas = festividadesProximas_(21);
  if (fiestas.length && cantidad >= 3) {
    instruccion += '\nSe acercan: ' + fiestas.join(', ') + '. Incluye al menos 1 idea que las aproveche ' +
      '(contenido con ese tema o un promocional de MÖÖnna ligado a la fecha), sin forzar el chiste.';
  }
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
  asegurarColumnas_();
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
    estilo: idea ? idea.estilo : 'texto',
    guion: idea ? idea.guion : '',
    visuales: idea ? idea.visuales : '',
    gancho: idea ? idea.gancho : '',
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

// ============================================================
// Video terminado (lo sube la PC)
// ============================================================

function subirVideo_(b) {
  const pr = fila_('proyectos', b.id);
  if (!pr) throw new Error('Grabación no encontrada');
  if (!b.base64 || b.base64.length < 1000) throw new Error('Video vacío');
  asegurarColumnas_();

  const bytes = Utilities.base64Decode(b.base64);
  const archivo = subcarpeta_('videos').createFile(Utilities.newBlob(bytes, 'video/mp4', pr.id + '.mp4'));
  // Con enlace para que el iPhone lo reproduzca sin iniciar sesión (va a Instagram de todos modos).
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  if (pr.video_id) { try { DriveApp.getFileById(pr.video_id).setTrashed(true); } catch (_) {} }

  let caption = pr.caption || '';
  if (!caption) {
    try { caption = redactarCaption_(pr); } catch (err) { caption = ''; }
  }
  actualizar_('proyectos', pr.id, { estado: 'listo', video_id: archivo.getId(), caption: caption, error: '' });
  return { ok: true, video_id: archivo.getId(), caption: caption };
}

function redactarCaption_(pr) {
  const sistema = 'Eres community manager de la marca descrita abajo. Escribes textos de publicación (caption) ' +
    'para Reels de Instagram en español de México: 2 a 4 líneas con gancho, un llamado a la acción al final y ' +
    'de 8 a 12 hashtags relevantes en una línea aparte. Responde SOLO con JSON: {"caption": "..."}\n\n=== MARCA ===\n' + marca_();
  const texto = preguntarIA_(sistema, 'Título: ' + pr.titulo + '\nGuion/transcripción: ' + (pr.texto || pr.guion));
  const ini = texto.indexOf('{'), fin = texto.lastIndexOf('}');
  return String(JSON.parse(texto.slice(ini, fin + 1)).caption || '').trim();
}

// ============================================================
// Instagram (API oficial de Meta)
// ============================================================

/**
 * Se corre UNA vez desde el editor. Pega el token de corta duración que te da el
 * Explorador de la API Graph (con permisos instagram_basic, instagram_content_publish,
 * pages_show_list, pages_read_engagement, business_management) y guarda lo necesario.
 * Antes, en Propiedades del script: META_APP_ID y META_APP_SECRET.
 */
function conectarInstagram(tokenCorto) {
  if (!tokenCorto) throw new Error('Pasa el token: conectarInstagram("EAAB...")');
  const appId = PROPS.getProperty('META_APP_ID'), secreto = PROPS.getProperty('META_APP_SECRET');
  if (!appId || !secreto) throw new Error('Faltan META_APP_ID y META_APP_SECRET en Propiedades del script');

  const largo = graph_('GET', '/oauth/access_token', {
    grant_type: 'fb_exchange_token', client_id: appId, client_secret: secreto, fb_exchange_token: tokenCorto,
  }).access_token;

  const paginas = graph_('GET', '/me/accounts', {
    fields: 'id,name,access_token,instagram_business_account{id,username}', access_token: largo,
  }).data || [];
  const pag = paginas.find(p => p.instagram_business_account);
  if (!pag) throw new Error('Ninguna de tus páginas de Facebook tiene una cuenta de Instagram profesional vinculada.');

  PROPS.setProperties({
    IG_USER_ID: pag.instagram_business_account.id,
    IG_USERNAME: pag.instagram_business_account.username,
    IG_TOKEN: pag.access_token, // token de página de larga duración: no caduca
  });
  Logger.log('Instagram conectado: @' + pag.instagram_business_account.username + ' (página "' + pag.name + '")');
}

function pedirPublicacion_(id) {
  const pr = fila_('proyectos', id);
  if (!pr) throw new Error('Grabación no encontrada');
  if (!PROPS.getProperty('IG_TOKEN')) throw new Error('Instagram no está conectado todavía (función conectarInstagram).');
  if (!pr.video_id) throw new Error('Este proyecto no tiene video');
  if (!['listo', 'error_publicar'].includes(pr.estado)) throw new Error('El video no está listo para publicar');
  asegurarColumnas_();

  const ig = PROPS.getProperty('IG_USER_ID'), token = PROPS.getProperty('IG_TOKEN');
  const base = { media_type: 'REELS', caption: pr.caption || '', share_to_feed: 'true', access_token: token };
  let contenedor;
  try {
    // 1) Subida directa del archivo (no depende de que Drive sirva el video públicamente)
    const c = graph_('POST', '/' + ig + '/media', Object.assign({ upload_type: 'resumable' }, base));
    const bytes = DriveApp.getFileById(pr.video_id).getBlob().getBytes();
    const r = UrlFetchApp.fetch(c.uri, {
      method: 'post', contentType: 'application/octet-stream', payload: bytes, muteHttpExceptions: true,
      headers: { Authorization: 'OAuth ' + token, offset: '0', file_size: String(bytes.length) },
    });
    const d = JSON.parse(r.getContentText() || '{}');
    if (r.getResponseCode() !== 200 || !d.success) throw new Error('subida: ' + r.getContentText().slice(0, 200));
    contenedor = c.id;
  } catch (e1) {
    // 2) Respaldo: que Meta descargue el video desde Drive
    const url = 'https://drive.google.com/uc?export=download&id=' + pr.video_id;
    contenedor = graph_('POST', '/' + ig + '/media', Object.assign({ video_url: url }, base)).id;
  }

  actualizar_('proyectos', id, { estado: 'publicando', ig_container: contenedor, error: '' });
  ScriptApp.newTrigger('procesarPublicaciones').timeBased().after(45 * 1000).create();
  return { ok: true, contenedor: contenedor };
}

/** Disparador: espera a que Meta procese el video y lo publica. Se reprograma si aún no termina. */
function procesarPublicaciones() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    for (const t of ScriptApp.getProjectTriggers()) {
      if (t.getHandlerFunction() === 'procesarPublicaciones') ScriptApp.deleteTrigger(t);
    }
    const ig = PROPS.getProperty('IG_USER_ID'), token = PROPS.getProperty('IG_TOKEN');
    let quedan = false;
    for (const pr of tabla_('proyectos')) {
      if (pr.estado !== 'publicando' || !pr.ig_container) continue;
      try {
        const st = graph_('GET', '/' + pr.ig_container, { fields: 'status_code,status', access_token: token });
        if (st.status_code === 'FINISHED') {
          const pub = graph_('POST', '/' + ig + '/media_publish', { creation_id: pr.ig_container, access_token: token });
          let permalink = '';
          try { permalink = graph_('GET', '/' + pub.id, { fields: 'permalink', access_token: token }).permalink || ''; } catch (_) {}
          actualizar_('proyectos', pr.id, { estado: 'publicado', ig_media_id: pub.id, permalink: permalink, publicado_en: ahora_(), error: '' });
        } else if (st.status_code === 'ERROR' || st.status_code === 'EXPIRED') {
          actualizar_('proyectos', pr.id, { estado: 'error_publicar', error: 'Meta rechazó el video: ' + (st.status || st.status_code) });
        } else {
          quedan = true; // IN_PROGRESS
        }
      } catch (err) {
        actualizar_('proyectos', pr.id, { estado: 'error_publicar', error: String(err.message || err).slice(0, 300) });
      }
    }
    if (quedan) ScriptApp.newTrigger('procesarPublicaciones').timeBased().after(60 * 1000).create();
  } finally {
    lock.releaseLock();
  }
}

function graph_(metodo, ruta, params) {
  const opciones = { method: metodo.toLowerCase(), muteHttpExceptions: true };
  let url = GRAPH + ruta;
  if (metodo === 'GET') {
    url += '?' + Object.entries(params).map(([k, v]) => k + '=' + encodeURIComponent(v)).join('&');
  } else {
    opciones.payload = params;
  }
  const r = UrlFetchApp.fetch(url, opciones);
  const d = JSON.parse(r.getContentText() || '{}');
  if (r.getResponseCode() >= 400 || d.error) {
    const e = d.error || {};
    throw new Error('Meta: ' + (e.error_user_msg || e.message || r.getContentText().slice(0, 200)));
  }
  return d;
}

/** Agrega a la hoja las columnas que se hayan definido después de instalar. */
function asegurarColumnas_() {
  for (const [nombre, cols] of Object.entries(COLUMNAS)) {
    const h = hoja_(nombre);
    const actuales = h.getRange(1, 1, 1, Math.max(1, h.getLastColumn())).getValues()[0].filter(String);
    const faltan = cols.filter(c => !actuales.includes(c));
    if (faltan.length) {
      h.getRange(1, actuales.length + 1, 1, faltan.length).setValues([faltan]).setFontWeight('bold');
    }
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
  return k.split(/\s+/)[0]; // si se pegaron más líneas (p. ej. MODELO_GEMINI=...), solo la clave

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
