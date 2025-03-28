// CONFIGURACIÓN GENERAL
const CONFIG = {
  spreadsheetId: 'YOUR_SPREADSHEET_ID', // ID de la planilla de Google Sheets
  sheetName: 'Hoja 2',                                          // Nombre de la hoja dentro de la planilla
  dataRange: 'A2:F9',                                           // Rango que contiene los datos de los exámenes
  calendar: CalendarApp.getDefaultCalendar(),                   // Calendario de Google donde se agregarán los eventos
  eventDurationHours: 2,                                        // Duración de cada evento (en horas)
  rowDelayMs: 500,                                              // Espera entre cada fila procesada para no saturar la API
  apiRateLimitDelayMs: 10000,                                   // Espera (en ms) si se detecta un límite de uso de la API
  avoidDuplicates: true,                                        // Si es true, evita crear eventos que ya existan
  emailRemindersDays: [7, 2],                                   // Días antes para enviar recordatorios por correo
  popupRemindersDays: [7, 2],                                   // Días antes para mostrar recordatorios emergentes
};

// Función principal: procesa los datos del spreadsheet y crea los eventos en el calendario
function addFinalExamsToCalendar() {
  try {
    // Validación del ID de la planilla
    if (!CONFIG.spreadsheetId || CONFIG.spreadsheetId === 'YOUR_SPREADSHEET_ID') {
      throw new Error('El ID de la planilla no está configurado. Reemplazá "YOUR_SPREADSHEET_ID" en la configuración.');
    }

    // Abrimos la planilla y la hoja
    Logger.log('Abriendo planilla: ' + CONFIG.spreadsheetId);
    const sheet = SpreadsheetApp.openById(CONFIG.spreadsheetId).getSheetByName(CONFIG.sheetName);
    if (!sheet) throw new Error('No se encontró la hoja: ' + CONFIG.sheetName);

    // Convertimos todas las celdas a formato texto para evitar errores al leer fechas u horas
    Logger.log('Formateando celdas como texto sin formato...');
    sheet.getDataRange().setNumberFormat('@');

    // Leemos los datos del rango especificado
    Logger.log('Leyendo datos del rango: ' + CONFIG.dataRange);
    const data = sheet.getRange(CONFIG.dataRange).getValues();

    // Calculamos desde qué fila empieza el rango (ej: A2 → empieza en la fila 2)
    const startRow = parseInt(CONFIG.dataRange.match(/\d+/)[0], 10);

    const calendar = CONFIG.calendar;
    Logger.log('Calendario cargado: ' + calendar.getName());

    // Procesamos cada fila
    data.forEach(function (row, index) {
      const rowNum = startRow + index;

      // Si la fila está vacía (todas las celdas vacías), la omitimos
      if (row.every(cell => cell === "")) {
        Logger.log('Fila vacía en la fila ' + rowNum + ', se omite.');
        return;
      }

      Logger.log('---\nProcesando fila ' + rowNum + ': ' + JSON.stringify(row));

      // Leemos los valores relevantes de la fila
      const subject = row[1];     // Materia
      const firstDate = row[2];   // Fecha 1er llamado
      const firstTime = row[3];   // Hora 1er llamado
      const secondDate = row[4];  // Fecha 2do llamado (opcional)
      const secondTime = row[5];  // Hora 2do llamado (opcional)

      // Si hay datos para el primer llamado, lo procesamos
      if (subject && firstDate && firstTime) {
        procesarEvento(subject, firstDate, firstTime, '1° llamado', calendar, rowNum);
      }

      // Si hay datos para el segundo llamado, lo procesamos
      if (subject && secondDate && secondTime) {
        procesarEvento(subject, secondDate, secondTime, '2° llamado', calendar, rowNum);
      }

      // Esperamos antes de procesar la siguiente fila
      Utilities.sleep(CONFIG.rowDelayMs);
    });

    Logger.log('Ejecución del script finalizada.');
  } catch (e) {
    Logger.log('Error general en la ejecución: ' + e.toString());
  }
}

// Crea un evento en el calendario (si no es duplicado) y agrega los recordatorios
function procesarEvento(subject, dateStr, timeStr, llamado, calendar, rowNum) {
  try {
    Logger.log(`Interpretando fecha/hora del ${llamado}: ${dateStr} ${timeStr}`);
    const eventDateTime = parseDateTime(dateStr, timeStr);
    Logger.log(`Fecha/hora interpretada: ${eventDateTime}`);

    const title = `${subject} - ${llamado}`;

    // Si se debe evitar duplicados y ya existe el evento, lo omitimos
    if (CONFIG.avoidDuplicates && eventExists(calendar, title, eventDateTime)) {
      Logger.log(`Evento duplicado: "${title}" @ ${eventDateTime}`);
    } else {
      // Creamos el evento
      const event = calendar.createEvent(
        title,
        eventDateTime,
        new Date(eventDateTime.getTime() + CONFIG.eventDurationHours * 60 * 60 * 1000)
      );
      Logger.log(`Evento creado: ${event.getTitle()} @ ${event.getStartTime()}`);

      // Agregamos recordatorios, y si hay error por límite, esperamos y seguimos
      try {
        addReminders(event);
      } catch (e) {
        if (e.toString().includes("too many calendars or calendar events")) {
          Logger.log(`Límite de la API alcanzado. Esperando ${CONFIG.apiRateLimitDelayMs / 1000} segundos...`);
          Utilities.sleep(CONFIG.apiRateLimitDelayMs);
        } else {
          throw e;
        }
      }
    }
  } catch (e) {
    Logger.log(`Error en la fila ${rowNum} (${llamado}): ${e.toString()}`);
  }
}

// Interpreta una fecha y una hora dadas en texto y las convierte en un objeto Date
function parseDateTime(dateStr, timeStr) {
  try {
    const meses = {
      "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
      "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
      "septiembre": "09", "setiembre": "09", "octubre": "10",
      "noviembre": "11", "diciembre": "12"
    };

    // Limpiamos el texto: eliminamos acentos y removemos el día de la semana
    dateStr = dateStr.toString().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
    timeStr = timeStr.toString().trim();

    const partes = dateStr.replace(/^\w+,\s*/, '').split(' de ');
    if (partes.length !== 3) throw new Error('Formato de fecha inesperado: ' + dateStr);

    const dia = partes[0].padStart(2, '0');
    const mes = meses[partes[1]];
    const anio = partes[2];
    if (!mes) throw new Error('Mes no reconocido: ' + partes[1]);

    const [horas, minutos] = timeStr.split(":").map(v => parseInt(v, 10));
    if (isNaN(horas) || isNaN(minutos)) throw new Error('Formato de hora inválido: "' + timeStr + '"');

    const fechaISO = `${anio}-${mes}-${dia}T${horas.toString().padStart(2, '0')}:${minutos.toString().padStart(2, '0')}:00`;
    const finalDate = new Date(fechaISO);
    if (isNaN(finalDate.getTime())) throw new Error('No se pudo construir la fecha final');

    Logger.log('Fecha y hora combinadas: ' + finalDate);
    return finalDate;
  } catch (e) {
    Logger.log('Error al interpretar fecha y hora: ' + e.toString());
    throw e;
  }
}

// Agrega recordatorios a un evento según la configuración
function addReminders(event) {
  event.removeAllReminders(); // Eliminamos los recordatorios previos (si había)

  CONFIG.emailRemindersDays.forEach(d => {
    event.addEmailReminder(d * 24 * 60); // Convertimos días a minutos
  });

  CONFIG.popupRemindersDays.forEach(d => {
    event.addPopupReminder(d * 24 * 60);
  });

  Logger.log('Recordatorios agregados a: ' + event.getTitle());
}

// Verifica si ya existe un evento con el mismo título y hora aproximada (±1 minuto)
function eventExists(calendar, title, startTime) {
  const start = new Date(startTime.getTime() - 60 * 1000); // un minuto antes
  const end = new Date(startTime.getTime() + 60 * 1000);   // un minuto después
  const events = calendar.getEvents(start, end);
  return events.some(event => event.getTitle() === title);
}