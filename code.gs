const SPREADSHEET_ID    = '1r7RcpcjfFqFVPvEsHhZX21mNDDDOoZ1LPhWmW8csvnE';
const USUARIOS_SHEET    = 'Usuarios';
const SHEET_NAME        = 'Salas';
const RESERVAS_SHEET    = 'Reservas';
const CAMBIOS_RESERVAS_SHEET = 'Cambios reservas';
const CIUDADES_SHEET    = 'Ciudades';
const CENTROS_SHEET     = 'Centros';
const CITY_IMAGES_FOLDER_NAME = 'ReservaSalas_Ciudades';
const CENTER_IMAGES_FOLDER_NAME = 'ReservaSalas_Centros';
const SALA_IMAGES_FOLDER_NAME = 'ReservaSalas_Salas';
// Dirección que recibirá copia de cada reserva
const RESPONSABLE_EMAIL = 'francisco.benavente.salgado@intelcia.com';
// URL de la webapp para incluir enlaces en los correos
const WEBAPP_URL       = (ScriptApp.getService && ScriptApp.getService().getUrl) ?
  ScriptApp.getService().getUrl() : '';
  
/**
 * Normaliza un valor (trim + toLowerCase) para comparaciones uniformes.
 */
function normalize(val) {
  // Devuelve el valor normalizado (trim + minúsculas) para comparaciones.
  return String(val || '').trim().toLowerCase();
}

/**
 * Devuelve la hoja por nombre ignorando mayúsculas/minúsculas y espacios.
 */
function getSheetByNameIC(ss, name) {
  // Obtiene una hoja por nombre ignorando mayúsculas/minúsculas y espacios.
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  const target = String(name).trim().toLowerCase();
  return ss.getSheets().find(s => s.getName().trim().toLowerCase() === target) || null;
}

/**
 * Comprueba si el usuario está en la base de datos de usuarios.
 * Devuelve null si no existe, o un objeto con sus datos (email, nombre, rol, etc).
 */
function getUserData() {
  // Busca el usuario activo en la hoja 'Usuarios' y devuelve sus datos o null.
  const email = Session.getActiveUser().getEmail();
  if (!email) return null;

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheetByNameIC(ss, USUARIOS_SHEET);
  const data  = sheet.getDataRange().getValues();
  const headers = data[0].map(h => normalize(h));
  const idxEmail = headers.indexOf('email');
  if (idxEmail === -1) return null;

  for (let i = 1; i < data.length; i++) {
    if (normalize(data[i][idxEmail]) === normalize(email)) {
      const usuario = {};
      headers.forEach((h, j) => usuario[h] = data[i][j]);
      return usuario;
    }
  }
  return null;
}

/**
 * Incluye un fragmento HTML por nombre de archivo.
 */
function include(filename) {
  // Evalúa el parcial como plantilla y le pasa 'usuario' y 'role'
  var t = HtmlService.createTemplateFromFile(filename);
  var u = getUserData();                       // lee tu usuario de la hoja
  t.usuario = u || {};                         // disponible en el parcial
  t.role    = normalize((u && u.rol) || 'usuario');
  return t.evaluate().getContent();            // <-- ahora sí se evalúan <?= ... ?>
}


/**
 * Genera la URL de Gravatar para el usuario activo.
 */
function getUserAvatarUrl() {
  // Calcula el hash MD5 del email y construye la URL de Gravatar (64px).
  const email = Session.getActiveUser().getEmail() || '';
  const hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    email.trim().toLowerCase()
  )
  .map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2))
  .join('');
  return `https://www.gravatar.com/avatar/${hash}?d=mp&s=64`;
}

/**
 * Devuelve el correo del usuario activo.
 */
function getActiveUserEmail() {
  // Devuelve el correo del usuario activo (o cadena vacía).
  return Session.getActiveUser().getEmail() || '';
}

/**
 * Punto de entrada web: comprueba acceso y pasa el objeto usuario a la plantilla.
 */
function doGet() {
  // Punto de entrada de la webapp: valida acceso y renderiza la plantilla con el rol.
  const user = getUserData();
  if (!user) { /* …acceso denegado… */ }

  const role = normalize(user.rol || 'usuario');
  const template = HtmlService.createTemplateFromFile('index');
  template.usuario = user;
  template.role    = role;
  return template
    .evaluate()
    .setTitle('Reserva salas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Lee y normaliza toda la hoja de Salas, ignorando filas vacías y cabeceras “raras”. */
function getAllSalas() {
  // Lee la hoja 'Salas', detecta cabecera válida, omite filas vacías y devuelve objetos.
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const name = (typeof SHEET_NAME === 'string' && SHEET_NAME) ? SHEET_NAME : 'Salas';
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error("Hoja '"+ name +"' no encontrada.");

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  // Cogemos todo lo “realmente usado”
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();

  // 1) Localiza la primera fila que tenga algo => cabecera
  const headerIdx = values.findIndex(row => row.some(c => String(c).trim() !== ''));
  if (headerIdx === -1 || headerIdx >= values.length - 1) return [];

  const headers = values[headerIdx].map(h => String(h).trim());

  // 2) Filtra filas totalmente vacías
  const dataRows = values
    .slice(headerIdx + 1)
    .filter(row => row.some(c => String(c).trim() !== ''));

  // 3) Mapea a objetos {Header: Valor}
  return dataRows.map(row => {
    const o = {};
    headers.forEach((h, i) => { o[h] = row[i]; });
    return o;
  });
}


/**
 * Devuelve todos los usuarios de la hoja de Usuarios.
 */
function getAllUsuarios() {
  // Lee toda la hoja 'Usuarios' y devuelve una lista de objetos por fila.
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(USUARIOS_SHEET);
  if (!sheet) throw new Error(`Hoja '${USUARIOS_SHEET}' no encontrada`);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values.shift().map(h => String(h).trim());
  return values.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

/** Devuelve todas las ciudades de la hoja 'Ciudades' */
function getAllCiudades() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheetByNameIC(ss, CIUDADES_SHEET);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values.shift().map(h => String(h).trim());
  return values.map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);
    return o;
  });
}

/** Devuelve todos los centros de la hoja 'Centros' */
function getAllCentros() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheetByNameIC(ss, CENTROS_SHEET);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values.shift().map(h => String(h).trim());
  return values.map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);
    return o;
  });
}

/** Devuelve la lista de roles (columna A de la hoja 'Roles') */
function getRoles() {
  // Lee la hoja 'Roles' (columna A) y devuelve la lista única de roles.
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheetByNameIC(ss, 'Roles');
  if (!sheet) return [];
  return [...new Set(
    sheet
      .getRange('A:A')
      .getValues()
      .flat()
      .map(r => String(r).trim())
      .filter(v => v)
  )];
}

/** Actualiza un usuario identificado por email */
function actualizarUsuario(usuario) {
  // Actualiza los campos del usuario identificado por email en la hoja 'Usuarios'.
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(USUARIOS_SHEET);
  if (!sheet) throw new Error(`Hoja '${USUARIOS_SHEET}' no encontrada`);

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idxEmail    = headers.indexOf('email');
  const idxNombre   = headers.indexOf('nombre');
  const idxRol      = headers.indexOf('rol');
  const idxCampania = headers.indexOf('Campaña');

  for (let i = 1; i < data.length; i++) {
    if (normalize(data[i][idxEmail]) === normalize(usuario.originalEmail || usuario.email)) {
      const row = i + 1;
      if (idxEmail    !== -1) sheet.getRange(row, idxEmail + 1).setValue(usuario.email);
      if (idxNombre   !== -1) sheet.getRange(row, idxNombre + 1).setValue(usuario.nombre);
      if (idxRol      !== -1) sheet.getRange(row, idxRol + 1).setValue(usuario.rol);
      if (idxCampania !== -1) sheet.getRange(row, idxCampania + 1).setValue(usuario['Campaña']);
      return true;
    }
  }
  throw new Error('Usuario no encontrado');
}

/** Crea un nuevo usuario */
function crearUsuario(usuario) {
  // Crea un nuevo registro en 'Usuarios' si el email no existe aún.
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(USUARIOS_SHEET);
  if (!sheet) throw new Error(`Hoja '${USUARIOS_SHEET}' no encontrada`);

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idxEmail    = headers.indexOf('email');
  const idxNombre   = headers.indexOf('nombre');
  const idxRol      = headers.indexOf('rol');
  const idxCampania = headers.indexOf('Campaña');

  for (let i = 1; i < data.length; i++) {
    if (normalize(data[i][idxEmail]) === normalize(usuario.email)) {
      throw new Error('El usuario ya existe');
    }
  }

  const row = headers.map(() => '');
  if (idxEmail    !== -1) row[idxEmail]    = usuario.email;
  if (idxNombre   !== -1) row[idxNombre]   = usuario.nombre;
  if (idxRol      !== -1) row[idxRol]      = usuario.rol;
  if (idxCampania !== -1) row[idxCampania] = usuario['Campaña'];
  sheet.appendRow(row);
  return true;
}



/** Elimina un usuario por email */
function eliminarUsuario(email) {
  // Elimina el usuario cuyo email coincida; devuelve true/false.
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(USUARIOS_SHEET);
  if (!sheet) throw new Error(`Hoja '${USUARIOS_SHEET}' no encontrada`);

  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idxEmail = headers.indexOf('email');

  for (let i = 1; i < data.length; i++) {
    if (normalize(data[i][idxEmail]) === normalize(email)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}


/** Devuelve lista única de ciudades */
function getCiudades() {
  const ciudadesDesdeTabla = getAllCiudades()
    .map(c => String(c['Nombre Ciudad'] || c.NombreCiudad || c.Nombre || '').trim())
    .filter(v => v)
    .sort();

  if (ciudadesDesdeTabla.length) {
    return ciudadesDesdeTabla;
  }

  // Fallback: si la tabla de ciudades está vacía, obtenemos las ciudades
  // directamente desde la hoja de salas.
  return [...new Set(
    getAllSalas()
      .map(r => String(r.Ciudad || '').trim())
      .filter(v => v)
  )].sort();
}

/** Devuelve lista única de centros filtrados por ciudad desde la hoja 'Salas' */
function getCentros(ciudad) {
  const ciudadNormalizada = normalize(ciudad);
  return [...new Set(
    getAllSalas()
      .filter(r => !ciudadNormalizada || normalize(r.Ciudad) === ciudadNormalizada)
      .map(r => String(r.Centro || '').trim())
      .filter(v => v)
  )].sort();
}

/** Devuelve lista única de salas (filtrada por ciudad y centro) */
function getSalas(ciudad, centro) {
  // Devuelve salas únicas (nombre + capacidad) filtrando por ciudad/centro.
  ciudad = normalize(ciudad);
  centro = normalize(centro);
  const all = getAllSalas();
  
  const seen = new Map();
  all.forEach(r => {
    if (ciudad && normalize(r.Ciudad) !== ciudad) return;
    if (centro && normalize(r.Centro) !== centro) return;

    const nombre = String(r.Nombre || '').trim();
    if (!nombre) return;

    const key = normalize(nombre);
    if (seen.has(key)) return;

    const capacidad = formatSalaCapacityValue(r.Capacidad);
    seen.set(key, {
      nombre: nombre,
      capacidad: capacidad
    });
  });

  return Array.from(seen.values()).sort((a, b) =>
    String(a.nombre || '').localeCompare(
      String(b.nombre || ''),
      'es',
      { sensitivity: 'base' }
    )
  );
}

function formatSalaCapacityValue(cap) {
  if (cap === null || cap === undefined) return '';
  if (typeof cap === 'number') {
    if (!isFinite(cap)) return '';
    return String(Math.round(cap));
  }
  const str = String(cap).trim();
  if (!str) return '';
  const normalized = str.replace(',', '.');
  const num = Number(normalized);
  if (!isNaN(num)) {
    return String(Math.round(num));
  }
  return str;
}

function normalizePersonas(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') {
    if (!isFinite(value)) return '';
    const rounded = Math.round(value);
    return rounded >= 1 ? String(rounded) : '';
  }
  const str = String(value).trim();
  if (!str) return '';
  const normalized = str.replace(',', '.');
  const num = Number(normalized);
  if (!isNaN(num) && isFinite(num)) {
    const rounded = Math.round(num);
    return rounded >= 1 ? String(rounded) : '';
  }
  return '';
}

/**
 * 1) Capa común: lectura de Spreadsheet y mapeo a objetos “crudos”
 */
function fetchReservasData() {
  // Lee la hoja 'Reservas' y devuelve {headers, rows} ya separados.
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESERVAS_SHEET);
  if (!sheet) return { headers: [], rows: [] };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { headers: [], rows: [] };

  const headers = data.shift().map(h => String(h).trim());
  return { headers, rows: data };
}

function mapRowToReserva(headers, row) {
  // Proyecta una fila de 'Reservas' a un objeto con campos tipados.
  const idx = name => headers.indexOf(name);
  const idxPersonas = idx('Personas');
  return {
    id:      row[idx('ID Reserva')],
    nombre:  row[idx('Nombre')],
    ciudad:  row[idx('Ciudad')],
    centro:  row[idx('Centro')],
    usuario: row[idx('Usuario')],
    motivo:  row[idx('Motivo')],
    inicio:  row[idx('Fecha inicio')],
    fin:     row[idx('Fecha fin')],
    personas: idxPersonas > -1 ? row[idxPersonas] : ''
  };
}

/**
 * Función genérica: recibe filtros y formateadores
 */
function getReservasGeneric({ filterFn, formatFn }) {
  // Aplica mapeo, filtro y formateo genérico sobre todas las reservas.
  const { headers, rows } = fetchReservasData();
  return rows
    .map(row => mapRowToReserva(headers, row))
    .filter(filterFn)
    .map(formatFn);
}

/**
 * Para el calendario: filtro por ciudad/centro/sala y formateo FullCalendar
 */
function getReservas(ciudad, centro, nombre) {
  // Devuelve eventos troceados por día (FullCalendar) para una sala concreta.
  ciudad = normalize(ciudad);
  centro = normalize(centro);
  nombre = normalize(nombre);
  if (!ciudad || !centro || !nombre) return [];

  const raw = getReservasGeneric({
    filterFn: r =>
      normalize(r.ciudad) === ciudad &&
      normalize(r.centro) === centro &&
      normalize(r.nombre) === nombre,
    formatFn: r => r
  });

  const events = [];
  raw.forEach(r => {
    const slots = splitRangeByDay(r.inicio, r.fin);
    slots.forEach(slot => {
      events.push({
        id:    r.id,
        title: r.motivo || r.usuario,
        start: toIso(slot.start),
        end:   toIso(slot.end),
        extendedProps: {
          usuario: r.usuario,
          motivo:  r.motivo,
          personas: r.personas
        }
      });
    });
  });

  return events.filter(ev => ev.start && ev.end);
}

/**
 * Devuelve todas las reservas en formato de tabla,
 * exactamente igual que getMisReservas() pero sin filtrar por usuario.
 */
function getTodasReservas() {
  // Devuelve todas las reservas formateadas para tabla y ordenadas por inicio.
  return getReservasGeneric({
    // no filtramos nada, devolvemos todo
    filterFn: () => true,
    // mismo formateo plano que en mis reservas
    formatFn: r => ({
      ID:      r.id,
      Sala:    r.nombre,
      Ciudad:  r.ciudad,
      Centro:  r.centro,
      Usuario: r.usuario,
      Motivo:  r.motivo,
      Inicio:  formatFecha(r.inicio),
      Fin:     formatFecha(r.fin)
    })
  })
  // opcional: ordenamos por fecha de inicio ascendente
  .sort((a, b) => {
    const pa = a.Inicio.split(' ')[0].split('/').reverse().join('-') + 'T' + a.Inicio.split(' ')[1];
    const pb = b.Inicio.split(' ')[0].split('/').reverse().join('-') + 'T' + b.Inicio.split(' ')[1];
    return new Date(pa) - new Date(pb);
  });
}


/**
 * Para la tabla de “mis reservas”: filtro por usuario y formateo plano
 */
function getMisReservas(email) {
  // Devuelve las reservas del usuario indicado (o el actual) formateadas para tabla.
  const me = normalize(email || getActiveUserEmail());
  return getReservasGeneric({
    filterFn: r => normalize(r.usuario) === me,
    formatFn: r => ({
      ID:      r.id,
      Sala:    r.nombre,
      Ciudad:  r.ciudad,
      Centro:  r.centro,
      Motivo:  r.motivo,
      Inicio:  formatFecha(r.inicio),
      Fin:     formatFecha(r.fin)
    })
  });
}


/**
 * Crea una reserva: la guarda, envía mail y crea evento en Calendar
 */
/**
 * Crea una reserva: valida solapamientos, la guarda, envía mail y crea evento en Calendar
 */
function crearReserva(reserva) {
  // Valida solapamientos, inserta la reserva, notifica por email y crea evento.
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(RESERVAS_SHEET);
  const PERSONAS_HEADER = 'Personas';
  if (!sheet) {
    sheet = ss.insertSheet(RESERVAS_SHEET);
    sheet.appendRow([
      'ID Reserva','Nombre','Ciudad','Centro','Usuario','Motivo','Fecha inicio','Fecha fin', PERSONAS_HEADER
    ]);
  }

  // 1) Leer todas las reservas actuales
  let { headers, rows } = fetchReservasData();
  if (headers.length && headers.indexOf(PERSONAS_HEADER) === -1) {
    headers = headers.concat([PERSONAS_HEADER]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  const personasValue = normalizePersonas(reserva.personas);
  reserva.personas = personasValue;
  const idxId = headers.indexOf('ID Reserva');
  let maxId = 0;
  rows.forEach(r => {
    const match = String(r[idxId]).match(/^RE(\d+)$/);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxId) maxId = num;
    }
  });
  const newId = 'RE' + String(maxId + 1).padStart(8, '0');

  const existing = rows
    .map(row => mapRowToReserva(headers, row))
    .filter(r =>
      normalize(r.ciudad) === normalize(reserva.ciudad) &&
      normalize(r.centro) === normalize(reserva.centro) &&
      normalize(r.nombre) === normalize(reserva.nombre)
    );
  
  // 2) Parsear fechas de la nueva reserva
  const newStart = new Date(reserva.fechaInicio);
  const newEnd   = new Date(reserva.fechaFin);
  if (newEnd <= newStart) {
    throw new Error('La fecha fin no puede ser anterior a la fecha inicio.');
  }

  // 3) Comprobar solapamientos
  const newSlots = splitRangeByDay(newStart, newEnd);
  const existingSlots = [];
  existing.forEach(r => {
    splitRangeByDay(new Date(r.inicio), new Date(r.fin)).forEach(s => existingSlots.push(s)); 
  });
  const overlap = newSlots.some(ns =>
    existingSlots.some(es => ns.start < es.end && es.start < ns.end)
  );
  if (overlap) {
    throw new Error('La franja elegida ya está reservada en ese rango de fechas.');
  }

  // 4) Si no hay solapamiento, seguimos con la creación
  const idReserva = reserva.idReserva ? String(reserva.idReserva) : newId;
  sheet.appendRow([
    idReserva,
    reserva.nombre,
    reserva.ciudad,
    reserva.centro,
    reserva.usuario,
    reserva.motivo,
    reserva.fechaInicio,
    reserva.fechaFin,
    reserva.personas
  ]);

  // 5) Notificar por email y Calendar
  if (!reserva.skipMail) {
    enviarMail(reserva, idReserva);
  }
  crearEventoCalendar(reserva, idReserva);

  return idReserva;
}


/**
 * Envía email de confirmación
 */
function enviarMail(reserva, idReserva) {
const tz = Session.getScriptTimeZone();
const hoy = new Date();
const fechaLarga = `Madrid, a ${Utilities.formatDate(hoy, tz, 'd')} de
${getNombreMes(hoy.getMonth())} de ${hoy.getFullYear()}`;
const subject = `Reserva ${idReserva} confirmada: ${reserva.nombre}`;
const body = `
<div style="font-family: Calibri, Arial, sans-serif; max-width:650px; margin:0 auto;
background:#fff;">
<div style="width:100%; text-align:center; margin:24px 0 32px;">
<img src="https://drive.google.com/uc?
export=view&id=1vNS8n_vYYJL9VQKx7jZxjZmMXUz0uECG" alt="Intelcia" style="width:100%;
max-width:500px; border-radius:8px;">
</div>
<div style="color:#222; font-size:16px; margin-bottom:24px;">
En ${fechaLarga}
</div>
<div style="margin-bottom:16px;">
Desde <span style="color:#C9006C; font-weight:bold;">Intelcia Spanish Region</span> te
informamos que tu reserva ha sido confirmada. Estos son los detalles:
</div>
<ul style="color:#C9006C; font-size:16px; margin-left:32px;">
<li><b>ID Reserva:</b> ${idReserva}</li>
<li><b>Sala:</b> ${reserva.nombre}</li>
<li><b>Ciudad:</b> ${reserva.ciudad}</li>
<li><b>Centro:</b> ${reserva.centro}</li>
<li><b>Inicio:</b> ${reserva.fechaInicio}</li>
<li><b>Fin:</b> ${reserva.fechaFin}</li>
<li><b>Motivo:</b> ${reserva.motivo}</li>
<li><b>Personas:</b> ${reserva.personas || 'No indicado'}</li>
</ul>
<div style="font-size:15px; margin:14px 0 16px; color:#222;">
Puedes consultar y gestionar tus reservas desde la aplicación
<span style="color:#C9006C; font-weight:bold;">Reserva de salas</span>
<a href="${WEBAPP_URL}" style="color:#C9006C; text-decoration:underline;">(enlace a la
aplicación)</a>.
</div>
<div style="margin-top:24px; font-size:15px;">Gracias</div>
</div>
`;
const backup = getBackupEmail(reserva.usuario);
MailApp.sendEmail({
to: reserva.usuario,
cc: [RESPONSABLE_EMAIL, backup].filter(Boolean).join(','),
subject: subject,
htmlBody: body
});
}

/**
 * Envía email de notificación por actualización de reserva
 */
function enviarMailActualizacion(reserva) {
const tz = Session.getScriptTimeZone();
const hoy = new Date();
const fechaLarga = `Madrid, a ${Utilities.formatDate(hoy, tz, 'd')} de
${getNombreMes(hoy.getMonth())} de ${hoy.getFullYear()}`;
const subject = `Reserva ${reserva.id} modificada: ${reserva.nombre}`;
const body = `
<div style="font-family: Calibri, Arial, sans-serif; max-width:650px; margin:0 auto;
background:#fff;">
<div style="width:100%; text-align:center; margin:24px 0 32px;">
<img src="https://drive.google.com/uc?
export=view&id=1vNS8n_vYYJL9VQKx7jZxjZmMXUz0uECG"
alt="Intelcia" style="width:100%; max-width:500px; border-radius:8px;">
</div>
<div style="color:#222; font-size:16px; margin-bottom:24px;">
En ${fechaLarga}
</div>
<div style="margin-bottom:16px;">
Desde <span style="color:#C9006C; font-weight:bold;">Intelcia Spanish Region</span>
te informamos que tu reserva ha sido <b>modificada</b>. Estos son los nuevos detalles:
</div>
<ul style="color:#C9006C; font-size:16px; margin-left:32px;">
<li><b>ID Reserva:</b> ${reserva.id}</li>
<li><b>Sala:</b> ${reserva.nombre}</li>
<li><b>Ciudad:</b> ${reserva.ciudad}</li>
<li><b>Centro:</b> ${reserva.centro}</li>
<li><b>Inicio:</b> ${reserva.fechaInicio}</li>
<li><b>Fin:</b> ${reserva.fechaFin}</li>
<li><b>Motivo:</b> ${reserva.motivo}</li>
<li><b>Personas:</b> ${reserva.personas || 'No indicado'}</li>
<li><b>Motivo de cambio:</b> ${reserva.motivoCambio || 'No indicado'}</li>
</ul>
<div style="font-size:15px; margin:14px 0 16px; color:#222;">
Puedes consultar y gestionar tus reservas desde la aplicación
<span style="color:#C9006C; font-weight:bold;">Reserva de salas</span>
<a href="${WEBAPP_URL}" style="color:#C9006C; text-decoration:underline;">(enlace a la
aplicación)</a>.
</div>
<div style="margin-top:24px; font-size:15px;">Gracias</div>
</div>
`;
const backup = getBackupEmail(reserva.usuario);
MailApp.sendEmail({
to: reserva.usuario,
cc: [RESPONSABLE_EMAIL, backup].filter(Boolean).join(','),
subject: subject,
htmlBody: body
});
}

/**
 * Envía un único correo con varias franjas reservadas
 */
function sendBulkReservationEmail(params) {
const emailUsuario = params && params.emailUsuario;
const motivo = params && params.motivo;
const reservas = Array.isArray(params && params.reservas) ? params.reservas : [];
const idReservaGlobal = params && params.idReserva;
if (!emailUsuario || !reservas.length) return;
const tz = Session.getScriptTimeZone();
const hoy = new Date();
const fechaLarga = `Madrid, a ${Utilities.formatDate(hoy, tz, 'd')} de
${getNombreMes(hoy.getMonth())} de ${hoy.getFullYear()}`;
// Construimos la lista de franjas reservadas
const listItems = reservas.map(r => {
const id = idReservaGlobal || r.idReserva || r.id || '';
const rango = `${r.fechaInicio} → ${r.fechaFin}`;
return `<li>${id ? '<b>' + id + '</b>: ' : ''}${rango}</li>`;
}).join('');
const subject = 'Reservas confirmadas';
const body = `
<div style="font-family: Calibri, Arial, sans-serif; max-width:650px; margin:0 auto;
background:#fff;">
<div style="width:100%; text-align:center; margin:24px 0 32px;">
<img src="https://drive.google.com/uc?
export=view&id=1vNS8n_vYYJL9VQKx7jZxjZmMXUz0uECG"
alt="Intelcia" style="width:100%; max-width:500px; border-radius:8px;">
</div>
<div style="color:#222; font-size:16px; margin-bottom:24px;">
En ${fechaLarga}
</div>
<div style="margin-bottom:16px;">
Desde <span style="color:#C9006C; font-weight:bold;">Intelcia Spanish Region</span>
te informamos que tu solicitud de reserva ha sido confirmada. Estos son los detalles:
</div>
<ul style="color:#C9006C; font-size:16px; margin-left:32px;">
${listItems}
</ul>
${motivo ? `<div style="margin:14px 0; font-size:15px; color:#222;"><b>Motivo:</b>
${motivo}</div>` : ''}
<div style="font-size:15px; margin:14px 0 16px; color:#222;">
Puedes consultar y gestionar tus reservas desde la aplicación
<span style="color:#C9006C; font-weight:bold;">Reserva de salas</span>
<a href="${WEBAPP_URL}" style="color:#C9006C; text-decoration:underline;">(enlace a la
aplicación)</a>.
</div>
<div style="margin-top:24px; font-size:15px;">Gracias</div>
</div>
`;
const backup = getBackupEmail(emailUsuario);
MailApp.sendEmail({
to: emailUsuario,
cc: [RESPONSABLE_EMAIL, backup].filter(Boolean).join(','),

subject: subject,
htmlBody: body
});
}

/**
 * Crea el evento en Google Calendar
 */
function crearEventoCalendar(reserva, idReserva) {
  // Crea un evento en el calendario por defecto con el resumen y los invitados.
  const perfil = getUserProfile(reserva.usuario);
  if (perfil && perfil.prefs && perfil.prefs.syncCalendar === false) return;
  const calendar = CalendarApp.getDefaultCalendar();
  const start    = new Date(reserva.fechaInicio);
  const end      = new Date(reserva.fechaFin);
  const backup   = getBackupEmail(reserva.usuario);
  calendar.createEvent(
    `Reserva ${idReserva} – ${reserva.nombre}`,
    start,
    end,
    {
      description: `Usuario: ${reserva.usuario}\nMotivo: ${reserva.motivo}` +
        (reserva.personas ? `\nPersonas: ${reserva.personas}` : ''),
      guests:      `${reserva.usuario},${RESPONSABLE_EMAIL}${backup ? ',' + backup : ''}`,
      sendInvites: true
    }
  );
}

/**
 * Divide un rango de fechas en franjas diarias manteniendo la hora.
 * Devuelve un array de objetos {start:Date, end:Date}.
 */
function splitRangeByDay(start, end) {
  // Divide un rango en franjas diarias manteniendo horas/minutos; valida fechas.
  try {
    start = start instanceof Date ? new Date(start) : new Date(start);
    end   = end instanceof Date ? new Date(end)   : new Date(end);
    if (isNaN(start) || isNaN(end)) return [];
  } catch(e) {
    return [];
  }

  const result = [];
  const hStart = start.getHours();
  const mStart = start.getMinutes();
  const hEnd   = end.getHours();
  const mEnd   = end.getMinutes();

  const current = new Date(start);
  current.setHours(0,0,0,0);
  const last = new Date(end);
  last.setHours(0,0,0,0);

  while (current <= last) {
    const s = new Date(current);
    s.setHours(hStart, mStart, 0, 0);
    const e = new Date(current);
    e.setHours(hEnd, mEnd, 0, 0);
    result.push({ start: s, end: e });
    current.setDate(current.getDate() + 1);
  }
  return result;
}

/**
 * Convierte fecha a ISO string para FullCalendar
 */
function toIso(fecha) {
  // Convierte fechas a ISO local para FullCalendar admitiendo varios formatos.
  if (!fecha) return null;
  if (fecha instanceof Date) {
    return Utilities.formatDate(
      fecha,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd'T'HH:mm:ss"
    );
  }
  const s = String(fecha).trim().replace(/\s+/g,' ');
  let m = s.match(/^(\d{4}-\d{2}-\d{2}) (\d{1,2}):(\d{2})$/);
  if (m) {
    let [_, d, h, mi] = m;
    h = h.padStart(2,'0');
    return `${d}T${h}:${mi}:00`;
  }
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{1,2}):(\d{2})$/);
  if (m) {
    let [_, day, mon, yr, h, mi] = m;
    day = day.padStart(2,'0'); mon = mon.padStart(2,'0');
    h   = h.padStart(2,'0');
    return `${yr}-${mon}-${day}T${h}:${mi}:00`;
  }
  return (/T\d{2}:\d{2}/.test(s) ? s : null);
}

/**
 * Formatea para la tabla: '2025-07-05 14:00' → '05/07/2025 14:00'
 */
function formatFecha(fecha) {
  // Formatea una fecha a 'dd/MM/yyyy HH:mm' para listados.
  if (!fecha) return '';
  try {
    const d = fecha instanceof Date
      ? fecha
      : new Date(fecha.replace(/-/g,'/'));
    const pad = n => ('0'+n).slice(-2);
    return `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()} `
         + `${pad(d.getHours())}:${pad(d.getMinutes())}`;
  } catch (e) {
    return fecha;
  }
}



/** Helper: devuelve el correo del usuario activo */
function getCorreoUsuario() {
  // Alias de utilidad para obtener el email del usuario activo.
  return Session.getActiveUser().getEmail() || '';
}

/**
 * Actualiza una reserva existente a partir de su ID
 */
function actualizarReserva(reserva) {
  // Revalida solapamientos y actualiza motivo/fechas de una reserva por ID.
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESERVAS_SHEET);
  if (!sheet) throw new Error('Hoja de reservas no encontrada');

  const start = new Date(reserva.fechaInicio);
  const end   = new Date(reserva.fechaFin);
  if (end <= start) {
    throw new Error('La fecha fin no puede ser anterior a la fecha inicio.');
  }
  if (!reserva.motivoCambio) {
    throw new Error('Debe indicar motivo de la modificación');
  }
  
  const data = sheet.getDataRange().getValues();
  let headers = data[0].map(h => String(h).trim());
  const personasValue = reserva.personas !== undefined
    ? normalizePersonas(reserva.personas)
    : undefined;
  if (personasValue !== undefined && headers.indexOf('Personas') === -1) {
    headers = headers.concat(['Personas']);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  const idxId = headers.indexOf('ID Reserva');
  const idxMotivo = headers.indexOf('Motivo');
  const idxInicio = headers.indexOf('Fecha inicio');
  const idxFin = headers.indexOf('Fecha fin');
  const idxNombre = headers.indexOf('Nombre');
  const idxCiudad = headers.indexOf('Ciudad');
  const idxCentro = headers.indexOf('Centro');
  const idxUsuario = headers.indexOf('Usuario');
  const idxPersonas = headers.indexOf('Personas');

  const reservas = data.slice(1).map(row => mapRowToReserva(headers, row));
  const actual = reservas.find(r => String(r.id) === String(reserva.id));
  if (!actual) {
    throw new Error('Reserva no encontrada');
  }

  const otros = reservas.filter(r =>
    String(r.id) !== String(reserva.id) &&
    normalize(r.ciudad) === normalize(actual.ciudad) &&
    normalize(r.centro) === normalize(actual.centro) &&
    normalize(r.nombre) === normalize(actual.nombre)
  );

  const newSlots = splitRangeByDay(start, end);
  const existingSlots = [];
  otros.forEach(r => {
    splitRangeByDay(new Date(r.inicio), new Date(r.fin)).forEach(s => existingSlots.push(s));
  });
  const overlap = newSlots.some(ns =>
    existingSlots.some(es => ns.start < es.end && es.start < ns.end)
  );
  if (overlap) {
    throw new Error('La franja elegida ya está reservada en ese rango de fechas.');
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(reserva.id)) {
      const row = i + 1;
      sheet.getRange(row, idxMotivo + 1).setValue(reserva.motivo);
      sheet.getRange(row, idxInicio + 1).setValue(reserva.fechaInicio);
      sheet.getRange(row, idxFin + 1).setValue(reserva.fechaFin);
      if (idxPersonas > -1 && personasValue !== undefined) {
        sheet.getRange(row, idxPersonas + 1).setValue(personasValue);
      }

      const cambios = [];
      const tz = Session.getScriptTimeZone();
      if (String(actual.motivo) !== String(reserva.motivo)) {
        cambios.push(`Motivo: ${actual.motivo} -> ${reserva.motivo}`);
      }
      if (new Date(actual.inicio).getTime() !== start.getTime()) {
        cambios.push(`Inicio: ${Utilities.formatDate(new Date(actual.inicio), tz, 'dd/MM/yyyy HH:mm')} -> ${Utilities.formatDate(start, tz, 'dd/MM/yyyy HH:mm')}`);
      }
      if (new Date(actual.fin).getTime() !== end.getTime()) {
        cambios.push(`Fin: ${Utilities.formatDate(new Date(actual.fin), tz, 'dd/MM/yyyy HH:mm')} -> ${Utilities.formatDate(end, tz, 'dd/MM/yyyy HH:mm')}`);
      }
      if (idxPersonas > -1 && personasValue !== undefined) {
        const actualPersonas = data[i][idxPersonas];
        if (String(actualPersonas || '') !== String(personasValue || '')) {
          cambios.push(`Personas: ${actualPersonas || '—'} -> ${personasValue || '—'}`);
        }
      }

      let logSheet = ss.getSheetByName(CAMBIOS_RESERVAS_SHEET);
      if (!logSheet) {
        logSheet = ss.insertSheet(CAMBIOS_RESERVAS_SHEET);
      }
      if (logSheet.getLastRow() === 0) {
        logSheet.appendRow(['Fecha modificación', 'ID Reserva', 'Cambios', 'Motivo modificación', 'Usuario']);
      }
      logSheet.appendRow([
        new Date(),
        reserva.id,
        cambios.join('; '),
        reserva.motivoCambio,
        Session.getActiveUser().getEmail()
      ]);

      const reservaCompleta = {
        id:        reserva.id,
        nombre:    data[i][idxNombre],
        ciudad:    data[i][idxCiudad],
        centro:    data[i][idxCentro],
        usuario:   data[i][idxUsuario],
        motivo:    reserva.motivo,
        fechaInicio: reserva.fechaInicio,
        fechaFin:    reserva.fechaFin,
        motivoCambio: reserva.motivoCambio,
        personas: idxPersonas > -1
          ? (personasValue !== undefined ? personasValue : data[i][idxPersonas])
          : ''
      };
      enviarMailActualizacion(reservaCompleta);
      return true;
    }
  }
  throw new Error('Reserva no encontrada');
}

/**
 * Elimina una reserva por ID
 */
function eliminarReserva(id) {
  // Elimina una reserva por su ID; devuelve true/false.
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESERVAS_SHEET);
  if (!sheet) throw new Error('Hoja de reservas no encontrada');

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idxId = headers.indexOf('ID Reserva');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) === String(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

/** Devuelve los datos completos de una reserva por ID */
function getReservaPorId(id) {
  // Recupera un objeto-reserva completo a partir de su ID o null.
  const { headers, rows } = fetchReservasData();
  const res = rows.map(r => mapRowToReserva(headers, r))
    .find(r => String(r.id) === String(id));
  return res || null;
}

/**
 * Convierte lista de reservas a CSV
 */
function _reservasToCsv(list) {
  // Serializa una lista de objetos reserva a CSV con cabeceras.
  if (!Array.isArray(list) || !list.length) return '';
  const headers = Object.keys(list[0]);
  const esc = v => '"' + String(v == null ? '' : v).replace(/"/g,'""') + '"';
  const lines = [headers.join(',')];
  list.forEach(r => {
    lines.push(headers.map(h => esc(r[h])).join(','));
  });
  return lines.join('\n');
}

/**
 * Exporta todas las reservas en formato CSV
 */
function exportTodasReservasCsv() {
  // Exporta todas las reservas en CSV (usa _reservasToCsv).
  return _reservasToCsv(getTodasReservas());
}

/**
 * Exporta las reservas del usuario actual en formato CSV
 */
function exportMisReservasCsv() {
  // Exporta solo las reservas del usuario actual en CSV.
  return _reservasToCsv(getMisReservas());
}

function _parseFecha(str) {
  // Parsea 'dd/MM/yyyy HH:mm' o deja que Date interprete la cadena.
  if (!str) return null;
  const m = str.match(/(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{2}):(\d{2})/);
  if (m) {
    const [,d,M,y,h,mi] = m;
    return new Date(`${y}-${(''+M).padStart(2,'0')}-${(''+d).padStart(2,'0')}T${h}:${mi}:00`);
  }
  return new Date(str);
}

function _formatIcsDate(d) {
  // Formatea una Date a formato UTC iCalendar (DTSTAMP/DTSTART/DTEND).
  return Utilities.formatDate(d, 'UTC', "yyyyMMdd'T'HHmmss'Z'");
}

function _reservasToIcs(list) {
  // Convierte una lista de reservas a un texto .ics (VCALENDAR/VEVENT).
  if (!Array.isArray(list) || !list.length) return '';
  const lines = ['BEGIN:VCALENDAR','VERSION:2.0','PRODID:-//ReservaSalas//EN'];
  list.forEach(r => {
    const start = _parseFecha(r.Inicio);
    const end   = _parseFecha(r.Fin);
    lines.push('BEGIN:VEVENT');
    lines.push('UID:' + r.ID);
    lines.push('SUMMARY:' + (r.Sala || ''));
    if (start) lines.push('DTSTART:' + _formatIcsDate(start));
    if (end)   lines.push('DTEND:' + _formatIcsDate(end));
    if (r.Motivo) lines.push('DESCRIPTION:' + r.Motivo);
    lines.push('END:VEVENT');
  });
  lines.push('END:VCALENDAR');
  return lines.join('\n');
}

function exportTodasReservasIcs() {
  // Exporta todas las reservas a iCalendar (.ics).
  return _reservasToIcs(getTodasReservas());
}

function exportMisReservasIcs() {
  // Exporta las reservas del usuario actual a iCalendar (.ics).
  return _reservasToIcs(getMisReservas());
}

/** Devuelve las salas libres para un rango y filtros opcionales */
function buscarSalasDisponibles(filtro) {
  // Devuelve las salas que no presentan solapamiento en el rango solicitado.
  const start = new Date(filtro && filtro.fechaInicio);
  const end   = new Date(filtro && filtro.fechaFin);
  if (isNaN(start) || isNaN(end) || end <= start) {
    throw new Error('Rango de fechas inválido');
  }

  let salas = getAllSalas();
  if (filtro && filtro.ciudad) {
    const c = normalize(filtro.ciudad);
    salas = salas.filter(s => normalize(s.Ciudad) === c);
  }
  if (filtro && filtro.centro) {
    const ce = normalize(filtro.centro);
    salas = salas.filter(s => normalize(s.Centro) === ce);
  }
  if (filtro && filtro.sala) {
    const sa = normalize(filtro.sala);
    salas = salas.filter(s => normalize(s.Nombre) === sa);
  }

  const { headers, rows } = fetchReservasData();
  const reservas = rows.map(r => mapRowToReserva(headers, r));

  function overlap(a1, a2, b1, b2) {
    return a1 < b2 && b1 < a2;
  }

  return salas.filter(room => {
    const resSala = reservas.filter(r =>
      normalize(r.ciudad) === normalize(room.Ciudad) &&
      normalize(r.centro) === normalize(room.Centro) &&
      normalize(r.nombre) === normalize(room.Nombre)
    );
    return !resSala.some(res => overlap(start, end, new Date(res.inicio), new Date(res.fin)));
  });
}

/**
 * Devuelve todas las salas (según filtros) y para cada una si está ocupada o disponible en ese rango.
 * Formato: [{Nombre, Ciudad, Centro, Capacidad, estado}]
 */
function buscarTodasSalasConEstado(filtro) {
  // Lista todas las salas (con filtros) marcando 'ocupada' o 'disponible'.
  const start = new Date(filtro && filtro.fechaInicio);
  const end   = new Date(filtro && filtro.fechaFin);
  if (isNaN(start) || isNaN(end) || end <= start) {
    throw new Error('Rango de fechas inválido');
  }

  let salas = getAllSalas();
  if (filtro && filtro.ciudad) {
    const c = normalize(filtro.ciudad);
    salas = salas.filter(s => normalize(s.Ciudad) === c);
  }
  if (filtro && filtro.centro) {
    const ce = normalize(filtro.centro);
    salas = salas.filter(s => normalize(s.Centro) === ce);
  }
  if (filtro && filtro.sala) {
    const sa = normalize(filtro.sala);
    salas = salas.filter(s => normalize(s.Nombre) === sa);
  }

  const { headers, rows } = fetchReservasData();
  const reservas = rows.map(r => mapRowToReserva(headers, r));

  function overlap(a1, a2, b1, b2) {
    return a1 < b2 && b1 < a2;
  }

  return salas.map(room => {
    const resSala = reservas.filter(r =>
      normalize(r.ciudad) === normalize(room.Ciudad) &&
      normalize(r.centro) === normalize(room.Centro) &&
      normalize(r.nombre) === normalize(room.Nombre)
    );
    const ocupada = resSala.some(res => overlap(start, end, new Date(res.inicio), new Date(res.fin)));
    return {
      ...room, // <-- esto mete TODOS los campos de la sala automáticamente
      estado: ocupada ? "ocupada" : "disponible"
    };
  }); // <-- El cierre del map es aquí, y no hay punto y coma dentro del return
}

/**
 * Envía una solicitud de uso de una sala ocupada al responsable.
 * Recibe: { sala, centro, ciudad, horario, motivo }
 */
function enviarSolicitudUso(datos) {
  // Envía email al responsable para solicitar uso de una sala ocupada y registra log.
  // 1. Obtener email del responsable (ajusta a tu lógica real)
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(USUARIOS_SHEET);
  var data = sheet.getDataRange().getValues();
  var headersUsuarios = data[0].map(h => normalize(h));
  var idxCentro = headersUsuarios.indexOf('centro');
  var idxRol = headersUsuarios.indexOf('rol');
  var idxEmail = headersUsuarios.indexOf('email');

  // Busca responsable de ese centro
  var responsableEmail = RESPONSABLE_EMAIL; // Valor por defecto si no encuentra otro
  for (let i = 1; i < data.length; i++) {
    if (
      normalize(data[i][idxCentro]) === normalize(datos.centro) &&
      normalize(data[i][idxRol]) === 'responsable'
    ) {
      responsableEmail = data[i][idxEmail];
      break;
    }
  }

  // Emails de administradores
  const adminEmails = data
    .slice(1)
    .filter(r => normalize(r[idxRol]) === 'admin')
    .map(r => r[idxEmail])
    .filter(e => e);

  // 2. Buscar la reserva existente que solapa con las fechas indicadas
  const { headers: reservaHeaders, rows } = fetchReservasData();
  const reservas = rows.map(r => mapRowToReserva(reservaHeaders, r));
  const reservaSolapada = reservas.find(r =>
    normalize(r.ciudad) === normalize(datos.ciudad) &&
    normalize(r.centro) === normalize(datos.centro) &&
    normalize(r.nombre) === normalize(datos.sala) &&
    new Date(r.inicio) < new Date(datos.fechaFin) &&
    new Date(datos.fechaInicio) < new Date(r.fin)
  );
  const idSolapada = reservaSolapada ? reservaSolapada.id : '';
  const ownerEmail = reservaSolapada ? reservaSolapada.usuario : responsableEmail;

  // 3. Componer email HTML corporativo
  const tz = Session.getScriptTimeZone();
  const hoy = new Date();
  const fechaLarga = `Madrid, a ${Utilities.formatDate(hoy, tz, 'd')} de ${getNombreMes(hoy.getMonth())} de ${hoy.getFullYear()}`;

  const solicitante = Session.getActiveUser().getEmail() || '-';
  const backupSolicitante = getBackupEmail(solicitante);

  const cuerpoHtml = `
    <div style="font-family: Calibri, Arial, sans-serif; max-width:650px; margin:0 auto; background:#fff;">
      <div style="width:100%; text-align:center; margin:24px 0 32px;">
        <img src="https://drive.google.com/uc?export=view&id=1vNS8n_vYYJL9VQKx7jZxjZmMXUz0uECG" 
             alt="Intelcia" style="width:100%; max-width:500px; border-radius:8px;">
      </div>
      <div style="color:#222; font-size:16px; margin-bottom:24px;">
        En ${fechaLarga}
      </div>
      <div style="margin-bottom:16px;">
        Desde <span style="color:#C9006C; font-weight:bold;">Intelcia Spanish Region</span> 
        te informamos que se ha registrado una <b>solicitud de uso de sala ocupada</b>. Estos son los detalles:
      </div>
      <ul style="color:#C9006C; font-size:16px; margin-left:32px;">
        <li><b>Sala:</b> ${datos.sala}</li>
        <li><b>Centro:</b> ${datos.centro}</li>
        <li><b>Ciudad:</b> ${datos.ciudad}</li>
        <li><b>Horario sala:</b> ${datos.horario || '-'}</li>
        <li><b>Fecha/hora solicitadas:</b> ${datos.fechaInicio || '-'} a ${datos.fechaFin || '-'}</li>
        <li><b>Motivo:</b> ${datos.motivo}</li>
        <li><b>Fecha/Hora solicitud:</b> ${Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm')}</li>
        <li><b>Solicitante:</b> ${solicitante}</li>
      </ul>
      ${idSolapada && WEBAPP_URL ? `
        <div style="margin-top:20px;">
          <a href="${WEBAPP_URL}?edit=${idSolapada}#mybookings"
             style="display:inline-block;padding:10px 16px;background:#bc348b;color:#fff;
                    text-decoration:none;border-radius:4px;">
            Liberar reserva
          </a>
        </div>` : ''}
      <div style="font-size:15px; margin:14px 0 16px; color:#222;">
        Puedes gestionar esta solicitud desde la aplicación 
        <span style="color:#C9006C; font-weight:bold;">Reserva de salas</span> 
        <a href="${WEBAPP_URL}" style="color:#C9006C; text-decoration:underline;">(enlace a la aplicación)</a>.
      </div>
      <div style="margin-top:24px; font-size:15px;">Gracias</div>
    </div>
  `;

  const mensaje =
    "Nueva solicitud de uso de sala ocupada:\n\n" +
    "Sala: " + datos.sala + "\n" +
    "Centro: " + datos.centro + "\n" +
    "Ciudad: " + datos.ciudad + "\n" +
    "Horario sala: " + (datos.horario || '') + "\n" +
    "Fecha/hora solicitadas: " + (datos.fechaInicio || '-') + " a " + (datos.fechaFin || '-') + "\n" +
    "Motivo: " + datos.motivo + "\n" +
    "Fecha/Hora solicitud: " + Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm") + "\n" +
    "Solicitante: " + solicitante;

  // Envío mails
  MailApp.sendEmail({
    to: ownerEmail,
    subject: 'Solicitud de uso de sala ocupada (' + datos.sala + ')',
    body: mensaje,
    htmlBody: cuerpoHtml
  });

  MailApp.sendEmail({
    to: solicitante,
    cc: backupSolicitante || '',
    subject: 'Copia de solicitud de uso (' + datos.sala + ')',
    body: mensaje,
    htmlBody: cuerpoHtml
  });

  if (adminEmails.length) {
    MailApp.sendEmail({
      to: adminEmails.join(','),
      subject: 'Solicitud de uso de sala ocupada (' + datos.sala + ')',
      body: mensaje,
      htmlBody: cuerpoHtml
    });
  }

  // 4. Guardar log en hoja "SolicitudesUso"
  var sheetLog = ss.getSheetByName('SolicitudesUso');
  if (!sheetLog) {
    sheetLog = ss.insertSheet('SolicitudesUso');
    sheetLog.appendRow([
      'Fecha solicitud', 'Sala', 'Centro', 'Ciudad',
      'Horario sala', 'Motivo', 'Responsable',
      'Fecha inicio', 'Fecha fin', 'Solicitante'
    ]);
  }
  sheetLog.appendRow([
    new Date(), datos.sala, datos.centro, datos.ciudad,
    datos.horario, datos.motivo, responsableEmail,
    datos.fechaInicio || '', datos.fechaFin || '', solicitante
  ]);

  // 5. Registrar notificación en hoja "Notificaciones"
  let notifSheet = ss.getSheetByName('Notificaciones');
  if (!notifSheet) {
    notifSheet = ss.insertSheet('Notificaciones');
    notifSheet.appendRow(['id','timestamp','type','actorEmail','reservaId','sala','centro','ciudad','fechaInicio','fechaFin','ownerEmail','message','link','status']);
  }
  const notifId = generarIdNotificacion(notifSheet);
  const fmt = d => d ? Utilities.formatDate(new Date(d), tz, 'dd/MM HH:mm') : '';
  const msgNotif = `Solicitud de uso para ${datos.sala} el ${fmt(datos.fechaInicio)}–${fmt(datos.fechaFin)}`;
  const link = WEBAPP_URL ? `${WEBAPP_URL}#reservas?id=${idSolapada}` : `#reservas?id=${idSolapada}`;
  notifSheet.appendRow([
    notifId,
    new Date().toISOString(),
    'solicitud_uso',
    solicitante,
    idSolapada,
    datos.sala,
    datos.centro,
    datos.ciudad,
    datos.fechaInicio || '',
    datos.fechaFin || '',
    ownerEmail,
    msgNotif,
    link,
    'open'
  ]);
}

/**
 * Generar una notificación de necesidad de uso de sala
 */

function generarIdNotificacion(sheet) {
  const lastRow = sheet.getLastRow();
  let next = 1;
  if (lastRow >= 2) {
    const lastId = String(sheet.getRange(lastRow, 1).getValue() || '');
    const match = /NTF-(\d+)/.exec(lastId);
    if (match) next = Number(match[1]) + 1;
  }
  return 'NTF-' + String(next).padStart(6, '0');
}

/**
 * Lee la hoja "Notificaciones" y devuelve las filas con fechas normalizadas a ISO.
 * - Abre el spreadsheet por ID (evita ActiveSpreadsheet nulo o libros equivocados).
 * - Convierte timestamp, fechaInicio y fechaFin a ISO 8601 seguro.
 * - Ordena por timestamp descendente.
 */
function getAllNotifications() {
  const SPREADSHEET_ID = '1r7RcpcjfFqFVPvEsHhZX21mNDDDOoZ1LPhWmW8csvnE';
  const SHEET_NAME = 'Notificaciones';
  const EXPECTED_HEADERS = [
    'id','timestamp','type','actorEmail','reservaId','sala','centro','ciudad',
    'fechaInicio','fechaFin','ownerEmail','message','link','status'
  ];

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return [];

  const range = sh.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return []; // solo cabeceras o vacío

  // 1) Cabeceras reales en la hoja
  const headers = values[0].map(h => String(h).trim());
  // 2) Crea un índice cabecera->col
  const idx = {};
  headers.forEach((h, i) => (idx[h] = i));

  // 3) Mapeo de filas a objetos con claves esperadas
  const rows = values.slice(1).filter(r => r.some(c => c !== '' && c !== null));

  const data = rows.map(r => {
    // Helper para coger un campo por nombre de cabecera; si no existe, devuelve ''
    const get = (name) => {
      const i = idx[name];
      return (i == null) ? '' : r[i];
    };

    // Normaliza fechas a ISO seguro
    const tsISO   = toISO(get('timestamp'));    // puede venir como Date, ISO, o texto
    const iniISO  = toISO(get('fechaInicio'));
    const finISO  = toISO(get('fechaFin'));

    return {
      id:          String(get('id') || '').trim(),
      timestamp:   tsISO,
      type:        String(get('type') || '').trim(),
      actorEmail:  String(get('actorEmail') || '').trim(),
      reservaId:   String(get('reservaId') || '').trim(),
      sala:        String(get('sala') || '').trim(),
      centro:      String(get('centro') || '').trim(),
      ciudad:      String(get('ciudad') || '').trim(),
      fechaInicio: iniISO,
      fechaFin:    finISO,
      ownerEmail:  String(get('ownerEmail') || '').trim(),
      message:     String(get('message') || '').trim(),
      link:        String(get('link') || '').trim(),
      status:      String(get('status') || '').trim(),
    };
  });

  // 4) Ordenar por timestamp descendente (lo más reciente primero)
  data.sort((a, b) => {
    const ta = Date.parse(a.timestamp || '') || 0;
    const tb = Date.parse(b.timestamp || '') || 0;
    return tb - ta;
  });

  // Log opcional para debug (ver en Ejecuciones)
  // Logger.log(JSON.stringify(data, null, 2));

  return data;

  /**
   * Convierte un valor (Date, string ISO, "YYYY-MM-DD HH:mm", número de Sheets) a ISO 8601.
   * Si no puede parsearse, devuelve cadena vacía.
   */
  function toISO(val) {
    if (val == null || val === '') return '';

    // Si ya es un objeto Date (Google Sheets entrega Date cuando la celda tiene formato fecha)
    if (val instanceof Date) {
      const d = val;
      // Asegura milisegundos para consistencia
      return new Date(d.getTime()).toISOString();
    }

    // Si es número (serial de Sheets convertido a número)
    if (typeof val === 'number') {
      // En Apps Script los valores de fecha ya llegan como Date, pero por si acaso…
      const millis = Math.round((val - 25569) * 86400 * 1000); // Excel epoch → ms
      return new Date(millis).toISOString();
    }

    // Si es string
    let s = String(val).trim();

    // Caso: ya es ISO válido
    // Date.parse reconoce bien 'YYYY-MM-DDTHH:mm:ss.sssZ'
    if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
      const t = Date.parse(s);
      return isNaN(t) ? '' : new Date(t).toISOString();
    }

    // Caso: 'YYYY-MM-DD HH:mm' o 'YYYY-MM-DD HH:mm:ss'
    // → convierto a 'YYYY-MM-DDTHH:mm(:ss)' y dejo que el motor lo trate como local
    if (/^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}(:\d{2})?$/.test(s)) {
      s = s.replace(' ', 'T');
      const t = Date.parse(s);
      return isNaN(t) ? '' : new Date(t).toISOString();
    }

    // Caso: solo fecha 'YYYY-MM-DD'
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      const t = Date.parse(s + 'T00:00:00');
      return isNaN(t) ? '' : new Date(t).toISOString();
    }

    // Último intento: Date.parse genérico
    const t = Date.parse(s);
    return isNaN(t) ? '' : new Date(t).toISOString();
  }
}


/** ========= CRUD SALAS (por 'ID Sala') ========= **/

function _getSalasSheet_() {
  // Devuelve el objeto Sheet de 'Salas' validando su existencia.
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheetByNameIC(ss, SHEET_NAME); // SHEET_NAME = 'Salas'
  if (!sheet) throw new Error(`Hoja '${SHEET_NAME}' no encontrada`);
  return sheet;
}

function _readSalasTable_() {
  // Lee la tabla completa de 'Salas' y separa cabecera/filas.
  const sheet = _getSalasSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { headers: [], rows: [], sheet };
  const headers = values.shift().map(h => String(h).trim());
  _ensureImageColumn_(sheet, headers, values, 'Imagen Sala');
  return { headers, rows: values, sheet };
}

function _indexOfHeader_(headers, name) {
  // Devuelve el índice de una cabecera (nombre exacto).
  return headers.indexOf(String(name).trim());
}

function _ensureImageColumn_(sheet, headers, rows, columnName) {
  if (!Array.isArray(headers)) return -1;
  let idx = _indexOfHeader_(headers, columnName);
  if (idx !== -1) return idx;

  if (!headers.length) {
    sheet.getRange(1, 1).setValue(columnName);
    headers.push(columnName);
    if (Array.isArray(rows)) rows.forEach(r => r.push(''));
    return 0;
  }

  const colPosition = headers.length;
  sheet.insertColumnAfter(colPosition);
  const newCol = colPosition + 1;
  sheet.getRange(1, newCol).setValue(columnName);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, newCol, lastRow - 1).clearContent();
  }
  headers.push(columnName);
  if (Array.isArray(rows)) rows.forEach(r => r.push(''));
  return headers.length - 1;
}

function _extractBytesFromUpload_(data) {
  let bytes = [];
  if (Array.isArray(data.bytes) && data.bytes.length) {
    bytes = data.bytes;
  } else if (data.bytes && typeof data.bytes.length === 'number') {
    bytes = Array.from(data.bytes);
  } else if (Array.isArray(data.content) && data.content.length) {
    bytes = data.content;
  } else if (data.dataUrl) {
    let base64 = String(data.dataUrl);
    const base64Index = base64.indexOf('base64,');
    if (base64Index !== -1) {
      base64 = base64.substring(base64Index + 7);
    }
    base64 = base64.trim();
    if (!base64) throw new Error('Datos de imagen inválidos.');
    bytes = Utilities.base64Decode(base64);
  } else {
    throw new Error('Imagen no proporcionada.');
  }

  if (bytes instanceof Uint8Array) {
    bytes = Array.from(bytes);
  }
  if (!bytes || !bytes.length) throw new Error('Datos de imagen inválidos.');
  return bytes;
}

/** Crea una nueva sala. Si no envías idSala, genero 'S' + timestamp. */
function crearSala(sala) {
  // Inserta una nueva sala (genera ID si falta) respetando el orden de cabeceras.
  const { headers, sheet } = _readSalasTable_();
  const h = (n) => _indexOfHeader_(headers, n);

  const idxId           = h('ID Sala');
  const idxCentro       = h('Centro');
  const idxCiudad       = h('Ciudad');
  const idxNombre       = h('Nombre');
  const idxCapacidad    = h('Capacidad');
  const idxEquipamiento = h('Equipamiento');
  const idxObs          = h('Observaciones');
  const idxHorario      = h('Horario');
  const idxImagen       = h('Imagen Sala');

  // ID (si viene vacío, generamos uno)
  const idSala = sala.idSala && String(sala.idSala).trim() ? String(sala.idSala).trim() : 'S' + Date.now();

  // Evitar duplicado por ID
  const { rows } = _readSalasTable_();
  for (let i = 0; i < rows.length; i++) {
    if (idxId !== -1 && String(rows[i][idxId]).trim() === idSala) {
      throw new Error('Ya existe una sala con ese ID.');
    }
  }

  // Construir fila según el orden de headers
  const row = headers.map(() => '');
  if (idxId           !== -1) row[idxId]           = idSala;
  if (idxCentro       !== -1) row[idxCentro]       = sala.Centro || '';
  if (idxCiudad       !== -1) row[idxCiudad]       = sala.Ciudad || '';
  if (idxNombre       !== -1) row[idxNombre]       = sala.Nombre || '';
  if (idxCapacidad    !== -1) row[idxCapacidad]    = sala.Capacidad || '';
  if (idxEquipamiento !== -1) row[idxEquipamiento] = sala.Equipamiento || '';
  if (idxObs          !== -1) row[idxObs]          = sala.Observaciones || '';
  if (idxHorario      !== -1) row[idxHorario]      = sala.Horario || '';
  if (idxImagen      !== -1) row[idxImagen]       = sala['Imagen Sala'] || sala.ImagenSala || sala.MapaSala || sala['Mapa Sala'] || '';

  sheet.appendRow(row);
  return idSala;
}

/** Actualiza una sala existente por ID Sala */
function actualizarSala(sala) {
  // Actualiza los campos de una sala existente identificada por 'ID Sala'.
  const { headers, rows, sheet } = _readSalasTable_();
  const h = (n) => _indexOfHeader_(headers, n);

  const idxId           = h('ID Sala');
  const idxCentro       = h('Centro');
  const idxCiudad       = h('Ciudad');
  const idxNombre       = h('Nombre');
  const idxCapacidad    = h('Capacidad');
  const idxEquipamiento = h('Equipamiento');
  const idxObs          = h('Observaciones');
  const idxHorario      = h('Horario');

  if (idxId === -1) throw new Error("No se encuentra la columna 'ID Sala'.");

  const targetId = String(sala.idSala || sala['ID Sala'] || '').trim();
  if (!targetId) throw new Error('ID Sala no especificado.');

  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][idxId]).trim() === targetId) {
      const rowNumber = i + 2; // +1 por cabecera y +1 para 1-based
      if (idxCentro       !== -1) sheet.getRange(rowNumber, idxCentro + 1).setValue(sala.Centro || '');
      if (idxCiudad       !== -1) sheet.getRange(rowNumber, idxCiudad + 1).setValue(sala.Ciudad || '');
      if (idxNombre       !== -1) sheet.getRange(rowNumber, idxNombre + 1).setValue(sala.Nombre || '');
      if (idxCapacidad    !== -1) sheet.getRange(rowNumber, idxCapacidad + 1).setValue(sala.Capacidad || '');
      if (idxEquipamiento !== -1) sheet.getRange(rowNumber, idxEquipamiento + 1).setValue(sala.Equipamiento || '');
      if (idxObs          !== -1) sheet.getRange(rowNumber, idxObs + 1).setValue(sala.Observaciones || '');
      if (idxHorario      !== -1) sheet.getRange(rowNumber, idxHorario + 1).setValue(sala.Horario || '');
      if (idxImagen       !== -1 && (sala['Imagen Sala'] || sala.ImagenSala || sala.MapaSala || sala['Mapa Sala'])) {
        sheet.getRange(rowNumber, idxImagen + 1).setValue(sala['Imagen Sala'] || sala.ImagenSala || sala.MapaSala || sala['Mapa Sala'] || '');
      }
      return true;
    }
  }
  throw new Error('Sala no encontrada.');
}

/** Elimina una sala por ID Sala */
function eliminarSala(idSala) {
  // Elimina una sala por su 'ID Sala'; devuelve true/false.
  const { headers, rows, sheet } = _readSalasTable_();
  const idxId = _indexOfHeader_(headers, 'ID Sala');
  const idxImagen = _indexOfHeader_(headers, 'Imagen Sala');
  if (idxId === -1) throw new Error("No se encuentra la columna 'ID Sala'.");

  const targetId = String(idSala || '').trim();
  if (!targetId) throw new Error('ID Sala no especificado.');

  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][idxId]).trim() === targetId) {
      if (idxImagen !== -1) {
        const prev = String(rows[i][idxImagen] || '').trim();
        const fileId = _extractDriveFileId_(prev);
        if (fileId) {
          try {
            DriveApp.getFileById(fileId).setTrashed(true);
          } catch (err) {
            Logger.log('No se pudo eliminar la imagen de la sala: ' + err);
          }
        }
      }
      sheet.deleteRow(i + 2); // +1 cabecera +1 base-1
      return true;
    }
  }
  return false;
}

function uploadSalaImagen(data) {
  if (!data || !data.idSala) throw new Error('ID de sala no proporcionado.');

  const allowedMimes = ['image/png', 'image/x-png'];
  let mime = String(data.mimeType || '').toLowerCase();
  if (mime && allowedMimes.indexOf(mime) === -1) throw new Error('Solo se permiten imágenes PNG.');
  if (!mime) mime = 'image/png';

  const { headers, rows, sheet } = _readSalasTable_();
  const idxId = _indexOfHeader_(headers, 'ID Sala');
  const idxImagen = _indexOfHeader_(headers, 'Imagen Sala');
  if (idxId === -1 || idxImagen === -1) throw new Error('Estructura de la tabla de salas no válida.');

  const id = String(data.idSala).trim();
  const rowIndex = rows.findIndex(r => String(r[idxId]).trim() === id);
  if (rowIndex === -1) throw new Error('Sala no encontrada.');

  const bytes = _extractBytesFromUpload_(data);

  const fileNameRaw = data.fileName && String(data.fileName).trim() ? String(data.fileName).trim() : `${id}.png`;
  const fileName = fileNameRaw.toLowerCase().endsWith('.png') ? fileNameRaw : `${fileNameRaw}.png`;
  const folder = _getSalaImagesFolder_();
  const blob = Utilities.newBlob(bytes, 'image/png', fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const prevUrl = String(rows[rowIndex][idxImagen] || '').trim();
  const prevId = _extractDriveFileId_(prevUrl);
  if (prevId) {
    try {
      DriveApp.getFileById(prevId).setTrashed(true);
    } catch (err) {
      Logger.log('No se pudo eliminar la imagen anterior de la sala: ' + err);
    }
  }

  const viewUrl = `https://drive.google.com/thumbnail?sz=w1000&id=${file.getId()}`;
  sheet.getRange(rowIndex + 2, idxImagen + 1).setValue(viewUrl);
  return viewUrl;
}

/** ========= CRUD CIUDADES ========= **/

function _getCiudadesSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheetByNameIC(ss, CIUDADES_SHEET);
  if (!sheet) throw new Error(`Hoja '${CIUDADES_SHEET}' no encontrada`);
  return sheet;
}

function _readCiudadesTable_() {
  const sheet = _getCiudadesSheet_();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { headers: [], rows: [], sheet };
  }
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  if (!values.length) return { headers: [], rows: [], sheet };
  const headers = values.shift().map(h => String(h).trim());
  const rows = values;
  _ensureCiudadImagenColumn_(sheet, headers, rows);
  return { headers, rows, sheet };
}

function _ensureCiudadImagenColumn_(sheet, headers, rows) {
  if (!Array.isArray(headers)) return -1;
  let idx = _indexOfHeader_(headers, 'Imagen Ciudad');
  if (idx !== -1) return idx;
  if (!headers.length) {
    sheet.getRange(1, 1).setValue('Imagen Ciudad');
    headers.push('Imagen Ciudad');
    if (Array.isArray(rows)) rows.forEach(r => r.push(''));
    return 0;
  }
  const colPosition = headers.length;
  sheet.insertColumnAfter(colPosition);
  const newCol = colPosition + 1;
  sheet.getRange(1, newCol).setValue('Imagen Ciudad');
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, newCol, lastRow - 1).clearContent();
  }
  headers.push('Imagen Ciudad');
  if (Array.isArray(rows)) rows.forEach(r => r.push(''));
  return headers.length - 1;
}

function crearCiudad(ciudad) {
  const { headers, sheet } = _readCiudadesTable_();
  const h = (n) => _indexOfHeader_(headers, n);
  const idxId = h('ID Ciudad');
  const idxNombre = h('Nombre Ciudad');
  const idxImagen = h('Imagen Ciudad');  
  const id = ciudad.idCiudad && String(ciudad.idCiudad).trim() ? String(ciudad.idCiudad).trim() : 'C' + Date.now();
  const { rows } = _readCiudadesTable_();
  if (rows.some(r => idxId !== -1 && String(r[idxId]).trim() === id)) throw new Error('Ya existe una ciudad con ese ID.');
  const row = headers.map(() => '');
  if (idxId !== -1) row[idxId] = id;
  if (idxNombre !== -1) row[idxNombre] = ciudad.NombreCiudad || ciudad.nombre || '';
  if (idxImagen !== -1) row[idxImagen] = ciudad['Imagen Ciudad'] || ciudad.ImagenCiudad || ciudad.imagen || '';
  sheet.appendRow(row);
  return id;
}

function actualizarCiudad(ciudad) {
  const { headers, rows, sheet } = _readCiudadesTable_();
  const h = (n) => _indexOfHeader_(headers, n);
  const idxId = h('ID Ciudad');
  const idxNombre = h('Nombre Ciudad');
  const idxImagen = h('Imagen Ciudad');
  if (idxId === -1) throw new Error("No se encuentra la columna 'ID Ciudad'.");
  const id = String(ciudad.idCiudad || ciudad['ID Ciudad'] || '').trim();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][idxId]).trim() === id) {
      const row = i + 2;
      if (idxNombre !== -1) sheet.getRange(row, idxNombre + 1).setValue(ciudad.NombreCiudad || ciudad.nombre || '');
      if (idxImagen !== -1 && (ciudad['Imagen Ciudad'] || ciudad.ImagenCiudad || ciudad.imagen)) {
        sheet.getRange(row, idxImagen + 1).setValue(ciudad['Imagen Ciudad'] || ciudad.ImagenCiudad || ciudad.imagen || '');
      }
      return true;
    }
  }
  throw new Error('Ciudad no encontrada.');
}

function eliminarCiudad(idCiudad) {
  const { headers, rows, sheet } = _readCiudadesTable_();
  const idxId = _indexOfHeader_(headers, 'ID Ciudad');
  const idxImagen = _indexOfHeader_(headers, 'Imagen Ciudad');
  if (idxId === -1) throw new Error("No se encuentra la columna 'ID Ciudad'.");
  const id = String(idCiudad || '').trim();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][idxId]).trim() === id) {
      if (idxImagen !== -1) {
        const prev = String(rows[i][idxImagen] || '').trim();
        const fileId = _extractDriveFileId_(prev);
        if (fileId) {
          try { DriveApp.getFileById(fileId).setTrashed(true); } catch (err) { Logger.log('No se pudo eliminar la imagen de ciudad: ' + err); }
        }
      }
      sheet.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

function _getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function _getCiudadImagesFolder_() {
  return _getOrCreateFolder_(CITY_IMAGES_FOLDER_NAME);
}

function _getCentroImagesFolder_() {
  return _getOrCreateFolder_(CENTER_IMAGES_FOLDER_NAME);
}

function _getSalaImagesFolder_() {
  return _getOrCreateFolder_(SALA_IMAGES_FOLDER_NAME);
}

function _extractDriveFileId_(url) {
  if (!url) return '';
  const str = String(url);
  const idParam = str.indexOf('id=') !== -1 ? str.split('id=')[1].split('&')[0] : '';
  if (idParam) return decodeURIComponent(idParam);
  const match = str.match(/[-\w]{25,}/);
  return match ? match[0] : '';
}

function uploadCiudadImagen(data) {
  if (!data || !data.idCiudad) throw new Error('ID de ciudad no proporcionado.');

  const allowedMimes = ['image/png', 'image/x-png'];
  let mime = String(data.mimeType || '').toLowerCase();
  if (mime && allowedMimes.indexOf(mime) === -1) throw new Error('Solo se permiten imágenes PNG.');
  if (!mime) mime = 'image/png';

  const { headers, rows, sheet } = _readCiudadesTable_();
  const idxId = _indexOfHeader_(headers, 'ID Ciudad');
  const idxImagen = _indexOfHeader_(headers, 'Imagen Ciudad');
  if (idxId === -1 || idxImagen === -1) throw new Error('Estructura de la tabla de ciudades no válida.');

  const id = String(data.idCiudad).trim();
  const rowIndex = rows.findIndex(r => String(r[idxId]).trim() === id);
  if (rowIndex === -1) throw new Error('Ciudad no encontrada.');

  let bytes = [];
  if (Array.isArray(data.bytes) && data.bytes.length) {
    bytes = data.bytes;
  } else if (data.bytes && typeof data.bytes.length === 'number') {
    bytes = Array.from(data.bytes);
  } else if (Array.isArray(data.content) && data.content.length) {
    bytes = data.content;
  } else if (data.dataUrl) {
    let base64 = String(data.dataUrl);
    const base64Index = base64.indexOf('base64,');
    if (base64Index !== -1) {
      base64 = base64.substring(base64Index + 7);
    }
    base64 = base64.trim();
    if (!base64) throw new Error('Datos de imagen inválidos.');
    bytes = Utilities.base64Decode(base64);
  } else {
    throw new Error('Imagen no proporcionada.');
  }

  if (bytes instanceof Uint8Array) {
    bytes = Array.from(bytes);
  }
  if (!bytes || !bytes.length) throw new Error('Datos de imagen inválidos.');

  const fileNameRaw = data.fileName && String(data.fileName).trim() ? String(data.fileName).trim() : `${id}.png`;
  const fileName = fileNameRaw.toLowerCase().endsWith('.png') ? fileNameRaw : `${fileNameRaw}.png`;
  const folder = _getCiudadImagesFolder_();
  const blob = Utilities.newBlob(bytes, 'image/png', fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const prevUrl = String(rows[rowIndex][idxImagen] || '').trim();
  const prevId = _extractDriveFileId_(prevUrl);
  if (prevId) {
    try { DriveApp.getFileById(prevId).setTrashed(true); } catch (err) { Logger.log('No se pudo eliminar la imagen anterior de la ciudad: ' + err); }
  }

  const viewUrl = `https://drive.google.com/thumbnail?sz=w5000&id=${file.getId()}`;
  sheet.getRange(rowIndex + 2, idxImagen + 1).setValue(viewUrl);
  return viewUrl;
}

/** ========= CRUD CENTROS ========= **/

function _getCentrosSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheetByNameIC(ss, CENTROS_SHEET);
  if (!sheet) throw new Error(`Hoja '${CENTROS_SHEET}' no encontrada`);
  return sheet;
}

function _readCentrosTable_() {
  const sheet = _getCentrosSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { headers: [], rows: [], sheet };
  const headers = values.shift().map(h => String(h).trim());
  _ensureImageColumn_(sheet, headers, values, 'Imagen Centro');
  return { headers, rows: values, sheet };
}

function crearCentro(centro) {
  const { headers, sheet } = _readCentrosTable_();
  const h = (n) => _indexOfHeader_(headers, n);
  const idxId   = h('ID Centro');
  const idxNom  = h('Nombre Centro');
  const idxDir  = h('Direccion Centro');
  const idxTel  = h('Telefono');
  const idxEmail= h('Email');
  const idxResp = h('Responsable');
  const idxImagen = h('Imagen Centro');
  const id = centro.idCentro && String(centro.idCentro).trim() ? String(centro.idCentro).trim() : 'CT' + Date.now();
  const { rows } = _readCentrosTable_();
  if (rows.some(r => idxId !== -1 && String(r[idxId]).trim() === id)) throw new Error('Ya existe un centro con ese ID.');
  const row = headers.map(() => '');
  if (idxId   !== -1) row[idxId]   = id;
  if (idxNom  !== -1) row[idxNom]  = centro.NombreCentro || centro.nombre || '';
  if (idxDir  !== -1) row[idxDir]  = centro.DireccionCentro || centro['Direccion Centro'] || '';
  if (idxTel  !== -1) row[idxTel]  = centro.Telefono || '';
  if (idxEmail!== -1) row[idxEmail]= centro.Email || '';
  if (idxResp !== -1) row[idxResp] = centro.Responsable || '';
  if (idxImagen !== -1) row[idxImagen] = centro['Imagen Centro'] || centro.ImagenCentro || centro.PlanoCentro || centro['Plano Centro'] || '';
  sheet.appendRow(row);
  return id;
}

function actualizarCentro(centro) {
  const { headers, rows, sheet } = _readCentrosTable_();
  const h = (n) => _indexOfHeader_(headers, n);
  const idxId   = h('ID Centro');
  const idxNom  = h('Nombre Centro');
  const idxDir  = h('Direccion Centro');
  const idxTel  = h('Telefono');
  const idxEmail= h('Email');
  const idxResp = h('Responsable');
  const idxImagen = h('Imagen Centro');
  if (idxId === -1) throw new Error("No se encuentra la columna 'ID Centro'.");
  const id = String(centro.idCentro || centro['ID Centro'] || '').trim();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][idxId]).trim() === id) {
      const row = i + 2;
      if (idxNom  !== -1) sheet.getRange(row, idxNom  + 1).setValue(centro.NombreCentro || centro.nombre || '');
      if (idxDir  !== -1) sheet.getRange(row, idxDir  + 1).setValue(centro.DireccionCentro || centro['Direccion Centro'] || '');
      if (idxTel  !== -1) sheet.getRange(row, idxTel  + 1).setValue(centro.Telefono || '');
      if (idxEmail!== -1) sheet.getRange(row, idxEmail+ 1).setValue(centro.Email || '');
      if (idxResp !== -1) sheet.getRange(row, idxResp + 1).setValue(centro.Responsable || '');
      if (idxImagen !== -1 && (centro['Imagen Centro'] || centro.ImagenCentro || centro.PlanoCentro || centro['Plano Centro'])) {
        sheet.getRange(row, idxImagen + 1).setValue(centro['Imagen Centro'] || centro.ImagenCentro || centro.PlanoCentro || centro['Plano Centro'] || '');
      }
      return true;
    }
  }
  throw new Error('Centro no encontrado.');
}

function eliminarCentro(idCentro) {
  const { headers, rows, sheet } = _readCentrosTable_();
  const idxId = _indexOfHeader_(headers, 'ID Centro');
  const idxImagen = _indexOfHeader_(headers, 'Imagen Centro');
  if (idxId === -1) throw new Error("No se encuentra la columna 'ID Centro'.");
  const id = String(idCentro || '').trim();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][idxId]).trim() === id) {
      if (idxImagen !== -1) {
        const prev = String(rows[i][idxImagen] || '').trim();
        const prevId = _extractDriveFileId_(prev);
        if (prevId) {
          try {
            DriveApp.getFileById(prevId).setTrashed(true);
          } catch (err) {
            Logger.log('No se pudo eliminar la imagen del centro: ' + err);
          }
        }
      }
      sheet.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

function uploadCentroImagen(data) {
  if (!data || !data.idCentro) throw new Error('ID de centro no proporcionado.');

  const allowedMimes = ['image/png', 'image/x-png'];
  let mime = String(data.mimeType || '').toLowerCase();
  if (mime && allowedMimes.indexOf(mime) === -1) throw new Error('Solo se permiten imágenes PNG.');
  if (!mime) mime = 'image/png';

  const { headers, rows, sheet } = _readCentrosTable_();
  const idxId = _indexOfHeader_(headers, 'ID Centro');
  const idxImagen = _indexOfHeader_(headers, 'Imagen Centro');
  if (idxId === -1 || idxImagen === -1) throw new Error('Estructura de la tabla de centros no válida.');

  const id = String(data.idCentro).trim();
  const rowIndex = rows.findIndex(r => String(r[idxId]).trim() === id);
  if (rowIndex === -1) throw new Error('Centro no encontrado.');

  const bytes = _extractBytesFromUpload_(data);

  const fileNameRaw = data.fileName && String(data.fileName).trim() ? String(data.fileName).trim() : `${id}.png`;
  const fileName = fileNameRaw.toLowerCase().endsWith('.png') ? fileNameRaw : `${fileNameRaw}.png`;
  const folder = _getCentroImagesFolder_();
  const blob = Utilities.newBlob(bytes, 'image/png', fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const prevUrl = String(rows[rowIndex][idxImagen] || '').trim();
  const prevId = _extractDriveFileId_(prevUrl);
  if (prevId) {
    try {
      DriveApp.getFileById(prevId).setTrashed(true);
    } catch (err) {
      Logger.log('No se pudo eliminar la imagen anterior del centro: ' + err);
    }
  }

  const viewUrl = `https://drive.google.com/thumbnail?sz=w2000&id=${file.getId()}`;
  sheet.getRange(rowIndex + 2, idxImagen + 1).setValue(viewUrl);
  return viewUrl;
}

/**
 * Devuelve el perfil del usuario por email.
 */
function getUserProfile(email) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(USUARIOS_SHEET);
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const idxEmail = headers.indexOf('email');
  const idxNombre = headers.indexOf('nombre');
  const idxRol = headers.indexOf('rol');
  const idxCamp = headers.indexOf('campaña');
  const idxPrefs = headers.indexOf('prefs_json');
  const idxBackup = headers.indexOf('backup_email');
  const idxLast = headers.indexOf('last_login');
  for (let i = 1; i < data.length; i++) {
    if (normalize(data[i][idxEmail]) === normalize(email)) {
      return {
        email: data[i][idxEmail],
        nombre: data[i][idxNombre],
        rol: data[i][idxRol],
        campaña: data[i][idxCamp],
        prefs: idxPrefs > -1 && data[i][idxPrefs] ? JSON.parse(data[i][idxPrefs]) : {},
        backup_email: idxBackup > -1 ? data[i][idxBackup] : '',
        last_login: idxLast > -1 ? data[i][idxLast] : '',
        avatar: getUserAvatarUrl()
      };
    }
  }
  return null;
}

/** Obtiene el correo de backup de un usuario */
function getBackupEmail(email) {
  const prof = getUserProfile(email);
  return prof && prof.backup_email ? prof.backup_email : '';
}

/**
 * Actualiza nombre, campaña y preferencias del usuario actual.
 */
function updateUserProfile(data) {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(USUARIOS_SHEET);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim().toLowerCase());

  const idxEmail   = headers.indexOf('email');
  const idxNombre  = headers.indexOf('nombre');
  const idxCamp    = headers.indexOf('campaña');       // o 'campana' si tu hoja no lleva tilde
  const idxPrefs   = headers.indexOf('prefs_json');
  const idxBackup  = headers.indexOf('backup_email');  // <- ¡declarado!
  // (opcional) const idxLast = headers.indexOf('last_login');

  for (let i = 1; i < values.length; i++) {
    if (normalize(values[i][idxEmail]) === normalize(email)) {
      const row = i + 1;

      if (idxNombre > -1) sh.getRange(row, idxNombre + 1).setValue(data.nombre || '');
      if (idxCamp   > -1) sh.getRange(row, idxCamp   + 1).setValue(data.campaña || data.campana || '');
      if (idxPrefs  > -1) sh.getRange(row, idxPrefs  + 1).setValue(JSON.stringify(data.prefs || {}));
      if (idxBackup > -1) sh.getRange(row, idxBackup + 1).setValue(data.backup_email || '');

      return true;
    }
  }
  return false;
}

/**
 * Devuelve las últimas acciones del usuario.
 */
function getUserActivity(email) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(CAMBIOS_RESERVAS_SHEET);
  if (!sh) return [];

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  // Cabeceras reales en "Cambios reservas"
  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const idxUsuario = headers.indexOf('usuario');               // existe en tu log
  const idxCambios = headers.indexOf('cambios');               // lo mapeamos a 'accion'
  const idxFechaMod = headers.indexOf('fecha modificación');   // lo mapeamos a 'fecha'
  if (idxUsuario === -1 || idxCambios === -1 || idxFechaMod === -1) return [];

  const rows = values
    .slice(1)
    .filter(r => normalize(r[idxUsuario]) === normalize(email))
    .map(r => ({ accion: r[idxCambios] || '', fecha: r[idxFechaMod] || '' }))
    .slice(-10);

  return rows;
}


/**
 * Devuelve métricas personales del usuario.
 */
function getUserMetrics(email) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(RESERVAS_SHEET);
  if (!sh) return {};

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return {};

  // Cabeceras reales en “Reservas”
  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const idxUsuario = headers.indexOf('usuario');          // << antes: 'email' (no existe)
  const idxInicio  = headers.indexOf('fecha inicio');     // << antes: 'inicio' (no existe)
  const idxFin     = headers.indexOf('fecha fin');        // << antes: 'fin' (no existe)
  if (idxUsuario === -1 || idxInicio === -1 || idxFin === -1) return {};

  // Si el activo es admin, no filtramos por email
  const user = getUserData(); // lee de la hoja Usuarios
  const isAdmin = user && normalize(user.rol) === 'admin';

  const now = new Date();
  let res30 = 0, res90 = 0, horas = 0;

  values.slice(1).forEach(r => {
    if (!isAdmin && normalize(r[idxUsuario]) !== normalize(email)) return;

    const inicio = new Date(r[idxInicio]);
    const fin    = new Date(r[idxFin]);
    if (inicio instanceof Date && !isNaN(inicio) && fin instanceof Date && !isNaN(fin)) {
      const diffH = (fin - inicio) / 36e5;
      if (!isNaN(diffH)) horas += diffH;
      if (now - inicio <= 30 * 24 * 60 * 60 * 1000) res30++;
      if (now - inicio <= 90 * 24 * 60 * 60 * 1000) res90++;
    }
  });

  return { res30, res90, horas: Math.round(horas) };
}

function importarUsuariosDesdeExcel(data) {
  // 1. Crear archivo temporal en Drive
  const blob = Utilities.newBlob(data.content, MimeType.MICROSOFT_EXCEL, data.name);
  const file = DriveApp.createFile(blob);

  // 2. Convertir Excel a Google Sheets con la API de Drive (v3)
  const resource = {
    name: "TEMP_Usuarios_" + new Date().toISOString(),
    mimeType: "application/vnd.google-apps.spreadsheet"
  };
  const newFile = Drive.Files.copy(resource, file.getId());
  const tempSS = SpreadsheetApp.openById(newFile.id);

  // 3. Leer la primera pestaña (asumimos que ahí está la tabla de usuarios)
  const hojaOrigen = tempSS.getSheets()[0];
  const datos = hojaOrigen.getDataRange().getValues();
  if (datos.length < 2) throw new Error("El archivo no contiene datos suficientes.");

  // 4. Obtener cabeceras
  const headers = datos[0].map(h => String(h).trim().toLowerCase());
  if (headers.indexOf("email") === -1) throw new Error("Falta la columna 'email' en el archivo.");
  if (headers.indexOf("nombre") === -1) throw new Error("Falta la columna 'nombre' en el archivo.");
  if (headers.indexOf("rol") === -1) throw new Error("Falta la columna 'rol' en el archivo.");

  // 5. Insertar los usuarios en la hoja 'Usuarios'
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hojaDestino = ss.getSheetByName(USUARIOS_SHEET);
  if (!hojaDestino) throw new Error("Hoja 'Usuarios' no encontrada");

  let nuevos = 0, duplicados = 0;

  for (let i = 1; i < datos.length; i++) {
    const row = datos[i];
    if (!row.some(v => String(v).trim() !== "")) continue; // saltar filas vacías

    const usuario = {
      email: row[headers.indexOf("email")] || "",
      nombre: row[headers.indexOf("nombre")] || "",
      rol: row[headers.indexOf("rol")] || "",
      campaña: (headers.indexOf("campaña") !== -1
                  ? row[headers.indexOf("campaña")]
                  : headers.indexOf("campana") !== -1
                    ? row[headers.indexOf("campana")]
                    : "")
    };

    if (usuario.email) {
      try {
        crearUsuario(usuario); // usas tu función ya existente
        nuevos++;
      } catch (e) {
        duplicados++;
        Logger.log("Usuario ya existente: " + usuario.email);
      }
    }
  }

  // 6. Limpiar archivos temporales
  DriveApp.getFileById(newFile.id).setTrashed(true);
  file.setTrashed(true);

  return `Importación completada. Nuevos: ${nuevos}, ya existentes: ${duplicados}`;
}

/**
 * Función para chatbot Gradio
 */

/***** Chatbot local en Apps Script: reglas de ayuda para reservas *****/

/***** LEE FAQ DESDE GOOGLE SHEETS *****/
function getFaqsFromSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID); // 👈 usamos tu constante
  const sheet = ss.getSheetByName("FAQ");
  if (!sheet) throw new Error("No se encontró la hoja FAQ");

  const rows = sheet.getDataRange().getValues();
  let faqs = [];

  for (let i = 1; i < rows.length; i++) { // saltamos cabecera
    const pregunta = rows[i][0];  // Columna A
    const respuesta = rows[i][1]; // Columna B
    if (pregunta && respuesta) {
      faqs.push({ q: pregunta.toString(), a: respuesta });
    }
  }
  return faqs;
}

/***** FUNCIONES DE SIMILITUD (Levenshtein) *****/
function similarity(s1, s2) {
  s1 = s1.toLowerCase();
  s2 = s2.toLowerCase();
  let longer = s1.length > s2.length ? s1 : s2;
  let shorter = s1.length > s2.length ? s2 : s1;
  let longerLength = longer.length;
  if (longerLength === 0) return 1.0;
  return (longerLength - editDistance(longer, shorter)) / longerLength;
}

function editDistance(s1, s2) {
  s1 = s1.toLowerCase();
  s2 = s2.toLowerCase();

  let costs = [];
  for (let i = 0; i <= s1.length; i++) {
    let lastValue = i;
    for (let j = 0; j <= s2.length; j++) {
      if (i === 0) costs[j] = j;
      else if (j > 0) {
        let newValue = costs[j - 1];
        if (s1.charAt(i - 1) !== s2.charAt(j - 1))
          newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
        costs[j - 1] = lastValue;
        lastValue = newValue;
      }
    }
    if (i > 0) costs[s2.length] = lastValue;
  }
  return costs[s2.length];
}

/***** CHATBOT RESPONSE CON SIMILITUD *****/
function getChatbotResponse(pregunta) {
  if (!pregunta) return "¿Puedes repetir la pregunta?";
  const faqs = getFaqsFromSheet();

  let mejorCoincidencia = null;
  let mejorSimilitud = 0;

  for (let f of faqs) {
    let sim = similarity(pregunta, f.q);
    if (sim > mejorSimilitud) {
      mejorSimilitud = sim;
      mejorCoincidencia = f.a;
    }
  }

  // 👇 Ajusta el umbral a tu gusto (0.5 = 50% parecido, 0.7 más exigente)
  if (mejorSimilitud > 0.5) {
    return mejorCoincidencia;
  }
  return "🤔 No encontré esa respuesta en las FAQs. Intenta con otra palabra clave.";
}

/***** TEST RAPIDO DESDE SCRIPT EDITOR *****/
function testChat() {
  const ejemplos = [
    "cómo hago una reserva",
    "quiero anular mi reserva",
    "hay wifi en la sala"
  ];
  ejemplos.forEach(q => Logger.log(q + " → " + getChatbotResponse(q)));
}

/***** NUEVA FUNCIÓN GET NOMBRE MES *****/
function getNombreMes(numeroMes) {
  const meses = [
    'enero','febrero','marzo','abril','mayo','junio',
    'julio','agosto','septiembre','octubre','noviembre','diciembre'
  ];
  return meses[numeroMes] || '';
}

