const SPREADSHEET_ID    = '1r7RcpcjfFqFVPvEsHhZX21mNDDDOoZ1LPhWmW8csvnE';
const USUARIOS_SHEET    = 'Usuarios';
const SHEET_NAME        = 'Salas';
const RESERVAS_SHEET    = 'Reservas';
// Dirección que recibirá copia de cada reserva
const RESPONSABLE_EMAIL = 'francisco.benavente.salgado@intelcia.com';
// URL de la webapp para incluir enlaces en los correos
const WEBAPP_URL       = (ScriptApp.getService && ScriptApp.getService().getUrl) ?
  ScriptApp.getService().getUrl() : '';
  
/**
 * Normaliza un valor (trim + toLowerCase) para comparaciones uniformes.
 */
function normalize(val) {
  return String(val || '').trim().toLowerCase();
}

/**
 * Devuelve la hoja por nombre ignorando mayúsculas/minúsculas y espacios.
 */
function getSheetByNameIC(ss, name) {
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
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Genera la URL de Gravatar para el usuario activo.
 */
function getUserAvatarUrl() {
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
  return Session.getActiveUser().getEmail() || '';
}

/**
 * Punto de entrada web: comprueba acceso y pasa el objeto usuario a la plantilla.
 */
function doGet() {
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

/**
 * Lee y normaliza toda la hoja de Salas.
 */
function getAllSalas() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheetByNameIC(ss, SHEET_NAME);
  if (!sheet) throw new Error(`Hoja '${SHEET_NAME}' no encontrada`);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values.shift().map(h => String(h).trim());
  return values.map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}



/**
 * Devuelve todos los usuarios de la hoja de Usuarios.
 */
function getAllUsuarios() {
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

/** Devuelve la lista de roles (columna A de la hoja 'Roles') */
function getRoles() {
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
  return [...new Set(
    getAllSalas()
      .map(r => String(r.Ciudad || '').trim())
      .filter(v => v)
  )].sort();
}

/** Devuelve lista única de centros (opcionalmente filtrada por ciudad) */
function getCentros(ciudad) {
  ciudad = normalize(ciudad);
  const all = getAllSalas();
  return [...new Set(
    all
      .filter(r => !ciudad || normalize(r.Ciudad) === ciudad)
      .map(r => String(r.Centro || '').trim())
      .filter(v => v)
  )].sort();
}

/** Devuelve lista única de salas (filtrada por ciudad y centro) */
function getSalas(ciudad, centro) {
  ciudad = normalize(ciudad);
  centro = normalize(centro);
  const all = getAllSalas();
  return [...new Set(
    all
      .filter(r => (!ciudad || normalize(r.Ciudad) === ciudad)
                 && (!centro || normalize(r.Centro) === centro))
      .map(r => String(r.Nombre || '').trim())
      .filter(v => v)
  )].sort();
}

/**
 * 1) Capa común: lectura de Spreadsheet y mapeo a objetos “crudos”
 */
function fetchReservasData() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESERVAS_SHEET);
  if (!sheet) return { headers: [], rows: [] };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { headers: [], rows: [] };

  const headers = data.shift().map(h => String(h).trim());
  return { headers, rows: data };
}

function mapRowToReserva(headers, row) {
  const idx = name => headers.indexOf(name);
  return {
    id:      row[idx('ID Reserva')],
    nombre:  row[idx('Nombre')],
    ciudad:  row[idx('Ciudad')],
    centro:  row[idx('Centro')],
    usuario: row[idx('Usuario')],
    motivo:  row[idx('Motivo')],
    inicio:  row[idx('Fecha inicio')],
    fin:     row[idx('Fecha fin')],
  };
}

/**
 * Función genérica: recibe filtros y formateadores
 */
function getReservasGeneric({ filterFn, formatFn }) {
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
          motivo:  r.motivo
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
function getMisReservas() {
  const me = normalize(getActiveUserEmail());
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
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(RESERVAS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RESERVAS_SHEET);
    sheet.appendRow([
      'ID Reserva','Nombre','Ciudad','Centro','Usuario','Motivo','Fecha inicio','Fecha fin'
    ]);
  }

  // 1) Leer todas las reservas actuales
  const { headers, rows } = fetchReservasData();
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
  const idReserva = reserva.idReserva ? String(reserva.idReserva) : 'R' + Date.now();
  sheet.appendRow([
    idReserva,
    reserva.nombre,
    reserva.ciudad,
    reserva.centro,
    reserva.usuario,
    reserva.motivo,
    reserva.fechaInicio,
    reserva.fechaFin
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
  const subject = `Reserva ${idReserva} confirmada: ${reserva.nombre}`;
  const body    = `
Hola,

Tu reserva ha sido registrada con éxito:

• ID Reserva: ${idReserva}
• Sala:       ${reserva.nombre}
• Ciudad:     ${reserva.ciudad}
• Centro:     ${reserva.centro}
• Inicio:     ${reserva.fechaInicio}
• Fin:        ${reserva.fechaFin}
• Motivo:     ${reserva.motivo}

Gracias por usar el sistema de reservas.
`;
  MailApp.sendEmail({
    to:      reserva.usuario,
    cc:      RESPONSABLE_EMAIL,
    subject: subject,
    body:    body
  });
}

/**
 * Envía email de notificación por actualización de reserva
 */
function enviarMailActualizacion(reserva) {
  const subject = `Reserva ${reserva.id} modificada: ${reserva.nombre}`;
  const body    = `
Hola,

Tu reserva ha sido modificada correctamente:

• ID Reserva: ${reserva.id}
• Sala:       ${reserva.nombre}
• Ciudad:     ${reserva.ciudad}
• Centro:     ${reserva.centro}
• Inicio:     ${reserva.fechaInicio}
• Fin:        ${reserva.fechaFin}
• Motivo:     ${reserva.motivo}

Gracias por usar el sistema de reservas.
`;
  MailApp.sendEmail({
    to:      reserva.usuario,
    cc:      RESPONSABLE_EMAIL,
    subject: subject,
    body:    body
  });
}

/**
 * Envía un único correo con varias franjas reservadas
 */
function sendBulkReservationEmail(params) {
  const emailUsuario = params && params.emailUsuario;
  const motivo       = params && params.motivo;
  const reservas     = Array.isArray(params && params.reservas) ? params.reservas : [];
  const idReservaGlobal = params && params.idReserva;
  if (!emailUsuario || !reservas.length) return;

  const list = reservas
    .map(r => {
      const id = idReservaGlobal || r.idReserva || r.id || '';
      const rango = `${r.fechaInicio} \u2192 ${r.fechaFin}`;
      return `<li>${id ? '<b>' + id + '</b>: ' : ''}${rango}</li>`;
    })
    .join('');

  const htmlBody = `
    <p>Hola,</p>
    <p>Tu reserva ha sido registrada con &eacute;xito para las siguientes franjas:</p>
    <ul>${list}</ul>
    <p>Motivo: ${motivo || ''}</p>
    <p>Gracias por usar el sistema de reservas.</p>
  `;

  MailApp.sendEmail({
    to: emailUsuario,
    cc: RESPONSABLE_EMAIL,
    subject: 'Reservas confirmadas',
    htmlBody: htmlBody
  });
}

/**
 * Crea el evento en Google Calendar
 */
function crearEventoCalendar(reserva, idReserva) {
  const calendar = CalendarApp.getDefaultCalendar();
  const start    = new Date(reserva.fechaInicio);
  const end      = new Date(reserva.fechaFin);
  calendar.createEvent(
    `Reserva ${idReserva} – ${reserva.nombre}`,
    start,
    end,
    {
      description: `Usuario: ${reserva.usuario}\nMotivo: ${reserva.motivo}`,
      guests:      `${reserva.usuario},${RESPONSABLE_EMAIL}`,
      sendInvites: true
    }
  );
}

/**
 * Divide un rango de fechas en franjas diarias manteniendo la hora.
 * Devuelve un array de objetos {start:Date, end:Date}.
 */
function splitRangeByDay(start, end) {
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
  return Session.getActiveUser().getEmail() || '';
}

/**
 * Actualiza una reserva existente a partir de su ID
 */
function actualizarReserva(reserva) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RESERVAS_SHEET);
  if (!sheet) throw new Error('Hoja de reservas no encontrada');

  const start = new Date(reserva.fechaInicio);
  const end   = new Date(reserva.fechaFin);
  if (end <= start) {
    throw new Error('La fecha fin no puede ser anterior a la fecha inicio.');
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idxId = headers.indexOf('ID Reserva');
  const idxMotivo = headers.indexOf('Motivo');
  const idxInicio = headers.indexOf('Fecha inicio');
  const idxFin = headers.indexOf('Fecha fin');
  const idxNombre = headers.indexOf('Nombre');
  const idxCiudad = headers.indexOf('Ciudad');
  const idxCentro = headers.indexOf('Centro');
  const idxUsuario = headers.indexOf('Usuario');

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

      const reservaCompleta = {
        id:        reserva.id,
        nombre:    data[i][idxNombre],
        ciudad:    data[i][idxCiudad],
        centro:    data[i][idxCentro],
        usuario:   data[i][idxUsuario],
        motivo:    reserva.motivo,
        fechaInicio: reserva.fechaInicio,
        fechaFin:    reserva.fechaFin
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
  const { headers, rows } = fetchReservasData();
  const res = rows.map(r => mapRowToReserva(headers, r))
    .find(r => String(r.id) === String(id));
  return res || null;
}

/**
 * Convierte lista de reservas a CSV
 */
function _reservasToCsv(list) {
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
  return _reservasToCsv(getTodasReservas());
}

/**
 * Exporta las reservas del usuario actual en formato CSV
 */
function exportMisReservasCsv() {
  return _reservasToCsv(getMisReservas());
}

function _parseFecha(str) {
  if (!str) return null;
  const m = str.match(/(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{2}):(\d{2})/);
  if (m) {
    const [,d,M,y,h,mi] = m;
    return new Date(`${y}-${(''+M).padStart(2,'0')}-${(''+d).padStart(2,'0')}T${h}:${mi}:00`);
  }
  return new Date(str);
}

function _formatIcsDate(d) {
  return Utilities.formatDate(d, 'UTC', "yyyyMMdd'T'HHmmss'Z'");
}

function _reservasToIcs(list) {
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
  return _reservasToIcs(getTodasReservas());
}

function exportMisReservasIcs() {
  return _reservasToIcs(getMisReservas());
}

/** Devuelve las salas libres para un rango y filtros opcionales */
function buscarSalasDisponibles(filtro) {
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

  // 2. Buscar la reserva existente que solapa con las fechas indicadas
  const { headers: reservaHeaders, rows } = fetchReservasData();
  const reservas = rows.map(r => mapRowToReserva(reservaHeaders, r));
  const idSolapada = reservas.find(r =>
    normalize(r.ciudad) === normalize(datos.ciudad) &&
    normalize(r.centro) === normalize(datos.centro) &&
    normalize(r.nombre) === normalize(datos.sala) &&
    new Date(r.inicio) < new Date(datos.fechaFin) &&
    new Date(datos.fechaInicio) < new Date(r.fin)
  )?.id || '';

  // 3. Componer email HTML con botón para liberar la reserva
  const solicitante = Session.getActiveUser().getEmail() || '-';
  const cuerpoHtml = `
    <p>Nueva solicitud de uso de sala ocupada:</p>
    <ul>
      <li><b>Sala:</b> ${datos.sala}</li>
      <li><b>Centro:</b> ${datos.centro}</li>
      <li><b>Ciudad:</b> ${datos.ciudad}</li>
      <li><b>Horario sala:</b> ${datos.horario || ''}</li>
      <li><b>Fecha/hora solicitadas:</b> ${datos.fechaInicio || '-'} a ${datos.fechaFin || '-'}</li>
      <li><b>Motivo:</b> ${datos.motivo}</li>
      <li><b>Fecha/Hora solicitud:</b> ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}</li>
      <li><b>Solicitante:</b> ${solicitante}</li>
    </ul>
    ${idSolapada && WEBAPP_URL ? '<a href="' + WEBAPP_URL + '?edit=' + idSolapada + '#mybookings" style="display:inline-block;margin-top:10px;padding:8px 14px;background:#bc348b;color:#fff;text-decoration:none;border-radius:4px">Liberar reserva</a>' : ''}
  `;

  const mensaje =
    "Nueva solicitud de uso de sala ocupada:\n\n" +
    "Sala: " + datos.sala + "\n" +
    "Centro: " + datos.centro + "\n" +
    "Ciudad: " + datos.ciudad + "\n" +
    "Horario sala: " + (datos.horario || '') + "\n" +
    "Fecha/hora solicitadas: " + (datos.fechaInicio || '-') + " a " + (datos.fechaFin || '-') + "\n" +
    "Motivo: " + datos.motivo + "\n" +
    "Fecha/Hora solicitud: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") + "\n" +
    "Solicitante: " + solicitante;

  MailApp.sendEmail({
    to: responsableEmail,
    subject: 'Solicitud de uso de sala ocupada (' + datos.sala + ')',
    body: mensaje,
    htmlBody: cuerpoHtml
  });

  // 3. (Opcional) Guarda la solicitud en una hoja de registro
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
}

