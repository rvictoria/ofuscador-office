// The panel of the add-in of Word: it calls the service of the machine (http://127.0.0.1) as the panel
// of the service does, with a session it gets for a one-time code (/panel/session), and changes the
// document in place: each value found by its label, each label back by its value, so the text keeps
// its format. The labels of a document are those of a conversation of its own, whose identifier the
// document keeps in its settings.

const $ = id => document.getElementById(id);

// The session lives as long as the panel is open: the origin of the page is shared by other pages
// of the same site, which must not find it later.
const SESSION = 'ofuscador.sesion';
const PORT = 'ofuscador.puerto';
const CONVERSATION = 'ofuscador.conversacion';

const read = (storage, key) => { try { return storage.getItem(key); } catch { return null; } };
const write = (storage, key, value) => { try { value === null ? storage.removeItem(key) : storage.setItem(key, value); } catch {} };

const port = () => Number(read(localStorage, PORT)) || 8790;
let session = read(sessionStorage, SESSION);
let busy = false;

function say(text, error = false) {
  $('estado').textContent = text;
  $('estado').className = error ? 'error' : '';
}

function show() {
  $('conectar').hidden = Boolean(session);
  $('acciones').hidden = !session;
  for (const button of document.querySelectorAll('button')) button.disabled = busy;
}

async function call(path, body = null, method = 'POST') {
  let response;
  try {
    response = await fetch(`http://127.0.0.1:${port()}${path}`, {
      method,
      headers: { 'Content-Type': 'application/json', ...(session ? { Authorization: `Bearer ${session}` } : {}) },
      body: body === null ? undefined : JSON.stringify(body),
    });
  } catch (error) {
    // What the proof of concept is about: whether Word lets this page reach the machine.
    console.error('[Ofuscador] sin respuesta del servicio', error);
    throw new Error(`No se llega al Ofuscador en 127.0.0.1:${port()}. ¿Está abierto? (${error.message})`);
  }
  const data = await response.json().catch(() => ({}));
  if (response.status === 403 && /Token/.test(data.error ?? '')) {
    session = null;
    write(sessionStorage, SESSION, null);
    show();
    throw new Error('La conexión con el Ofuscador ha caducado: conéctalo otra vez.');
  }
  if (!response.ok) throw new Error(data.error ?? `El Ofuscador respondió ${response.status}.`);
  return data;
}

// The conversation of the document: made the first time, kept in the document.
async function conversation() {
  const settings = Office.context.document.settings;
  let id = settings.get(CONVERSATION);
  if (!id) {
    id = `word-${crypto.randomUUID()}`;
    settings.set(CONVERSATION, id);
    await new Promise(done => settings.saveAsync(done));
  }
  return id;
}

// A text as Word looks for it: «^» begins its special characters. It finds 255 characters at most.
const searchable = text => text.replaceAll('^', '^^');
const wholeWord = text => /^[\p{L}\p{N}]/u.test(text) && /[\p{L}\p{N}]$/u.test(text);

// Each pair [what is written, what goes instead] changed in `range`, the longest first, every place
// it is: how many places, and the pairs that could not be found as they are (across paragraphs).
async function replaceIn(context, range, pairs) {
  const once = new Map();
  for (const [from, to] of pairs) if (from && !once.has(from)) once.set(from, to);
  const sorted = [...once].sort(([a], [b]) => b.length - a.length);
  let changed = 0;
  const missed = [];
  for (const [from, to] of sorted) {
    if (from.length > 255) { missed.push(to); continue; }
    const found = range.search(searchable(from), { matchCase: true, matchWholeWord: wholeWord(from) });
    found.load('items');
    await context.sync();
    if (!found.items.length) { missed.push(to); continue; }
    for (const item of found.items) item.insertText(to, 'Replace');
    await context.sync();
    changed += found.items.length;
  }
  return { changed, missed };
}

async function protect(whole) {
  await Word.run(async context => {
    const range = whole ? context.document.body : context.document.getSelection();
    range.load('text');
    await context.sync();
    if (!range.text.trim()) {
      say(whole ? 'El documento está vacío.' : 'Selecciona primero el texto que quieres proteger.');
      return;
    }
    const result = await call('/protect', { session: await conversation(), text: range.text });
    const pairs = [...result.replacements.map(item => [item.value, item.token]), ...(result.generalized ?? [])];
    if (!pairs.length) {
      say('No hay nada que ocultar.');
      return;
    }
    const { changed, missed } = await replaceIn(context, range, pairs);
    // What is left of the values, read again from the document.
    range.load('text');
    await context.sync();
    const left = pairs.filter(([value]) => value && range.text.includes(value)).map(([, token]) => token);
    const lost = [...new Set([...missed, ...left])];
    say(lost.length
      ? `Cambiados ${changed} sitios, pero no se ha podido cambiar en su sitio: ${lost.join(', ')}. Cámbialo a mano antes de enviarlo.`
      : `Protegido: ${changed} ${changed === 1 ? 'dato cambiado' : 'datos cambiados'} por su etiqueta.`, lost.length > 0);
  });
}

async function restore() {
  await Word.run(async context => {
    const { labels } = await call(`/map?session=${encodeURIComponent(await conversation())}`, null, 'GET');
    const pairs = Object.entries(labels ?? {});
    if (!pairs.length) {
      say('Este documento aún no tiene etiquetas que restaurar.');
      return;
    }
    const { changed } = await replaceIn(context, context.document.body, pairs);
    say(changed ? `Restaurado: ${changed} ${changed === 1 ? 'etiqueta' : 'etiquetas'} con su dato.` : 'No hay etiquetas de este documento en él.');
  });
}

async function connect() {
  const code = $('codigo').value.trim();
  if (!code) { say('Escribe el código que da el Ofuscador.', true); return; }
  write(localStorage, PORT, String(Number($('puerto').value) || 8790));
  const answer = await call('/panel/session', { code });
  session = answer.session;
  write(sessionStorage, SESSION, session);
  $('codigo').value = '';
  say('Conectado con el Ofuscador.');
}

async function leave() {
  await call('/panel/logout').catch(() => {});
  session = null;
  write(sessionStorage, SESSION, null);
  say('Desconectado.');
}

function action(work) {
  return async () => {
    if (busy) return;
    busy = true;
    show();
    try {
      await work();
    } catch (error) {
      say(error.message, true);
    } finally {
      busy = false;
      show();
    }
  };
}

Office.onReady(() => {
  $('puerto').value = String(port());
  $('entrar').addEventListener('click', action(connect));
  $('codigo').addEventListener('keydown', event => { if (event.key === 'Enter') action(connect)(); });
  $('proteger-seleccion').addEventListener('click', action(() => protect(false)));
  $('proteger-documento').addEventListener('click', action(() => protect(true)));
  $('restaurar').addEventListener('click', action(restore));
  $('salir').addEventListener('click', action(leave));
  show();
});
