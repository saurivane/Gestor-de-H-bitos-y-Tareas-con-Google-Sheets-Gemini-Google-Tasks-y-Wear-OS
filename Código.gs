/**
 * ============================================================================
 * GESTOR DE TAREAS Y HÁBITOS - BACKEND, SYNC TASKS + WEB APP CON CRUD COMPLETO
 * ============================================================================
 * Pestañas requeridas en Google Sheets:
 * 1. "Plantillas_Habitos" : [ID, Nombre, Categoría, Frecuencia, Hora/Momento, Activo]
 * 2. "Registro_Diario"    : [Fecha, Tipo, Nombre, Categoría, Estado, Timestamp_Completado]
 * 3. "Dashboard"          : Panel de métricas y KPIs.
 * ============================================================================
 */

const CONFIG = {
  SHEETS: {
    PLANTILLAS: 'Plantillas_Habitos',
    REGISTRO: 'Registro_Diario',
    DASHBOARD: 'Dashboard'
  },
  CATEGORIAS: ['Salud', 'Estudio', 'Hogar', 'Mantenimiento', 'Finanzas', 'Trabajo'],
  FRECUENCIAS: ['Diaria', 'L-V', 'Fin de semana', 'Semanal'],
  ESTADOS: ['Pendiente', 'Completado'],
  TASKS_LIST_NAME: 'Hábitos',
  GEMINI_MODELS: [
    'gemini-3.5-flash-lite',
    'gemini-3.6-flash',
    'gemini-3.8-flash',
    'gemini-4-argon',
    'gemini-3.1-flash-lite'
  ],
  TIMEZONE: Session.getScriptTimeZone() || 'Europe/Madrid'
};

/**
 * 1. CONFIGURACIÓN Y UI EN SHEETS (onOpen)
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ Mi Gestor')
    .addItem('🗣️ Entrada rápida por comando (IA)', 'mostrarDialogoComando')
    .addSeparator()
    .addItem('🔄 1. Volcar hábitos de hoy (+ Enviar a Google Tasks)', 'generarHabitosDelDiaManual')
    .addItem('🔃 2. Sincronizar con Google Tasks ("Hábitos") ahora', 'sincronizarGoogleTasksManual')
    .addItem('📊 3. Actualizar tabla KPI por Meses en hoja Dashboard', 'actualizarKpiMensualEnSheetManual')
    .addSeparator()
    .addItem('⏰ Activar automatización total (00:05 + Sync cada 5 min)', 'instalarTriggersAutomaticos')
    .addItem('🎙️ Configurar alias de Asistente de Voz', 'configurarPrefijoAsistente')
    .addItem('🔑 Configurar Gemini API Key', 'configurarApiKey')
    .addToUi();
}

function configurarPrefijoAsistente() {
  const ui = SpreadsheetApp.getUi();
  const actual = PropertiesService.getScriptProperties().getProperty('VOICE_ASSISTANT_PREFIX') || '';
  const res = ui.prompt(
    '🎙️ Configurar alias de Asistente de Voz',
    `Introduce el nombre o alias de tu asistente personalizado en Gemini${actual ? ` (Actual: "${actual}")` : ''}:\n\n` +
    'Si configuras una regla en "Contexto personal" de Gemini para que anteponga [TuAlias] al crear tareas en Google Tasks, el script las detectará automáticamente en cualquier lista de Google Tasks, las procesará y las moverá limpias a tu lista "Hábitos".\n' +
    '(Déjalo en blanco si quieres desactivar el filtro por alias).',
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() === ui.Button.OK) {
    const alias = res.getResponseText().trim().replace(/^\[|\]$/g, '').trim();
    if (alias) {
      PropertiesService.getScriptProperties().setProperty('VOICE_ASSISTANT_PREFIX', alias);
      ui.alert(`✅ Alias de asistente guardado como "${alias}".\n\nCualquier tarea que entre en Google Tasks con el prefijo "[${alias}]" o "${alias}:" será procesada automáticamente.`);
    } else {
      PropertiesService.getScriptProperties().deleteProperty('VOICE_ASSISTANT_PREFIX');
      ui.alert('ℹ️ Se ha eliminado el alias de asistente personalizado.');
    }
  }
}

function configurarApiKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Configurar Gemini API Key',
    'Introduce tu API Key de Google AI Studio (AIzaSy...):',
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() === ui.Button.OK) {
    const key = res.getResponseText().trim();
    if (key) {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
      ui.alert('✅ API Key guardada correctamente en las propiedades del script.');
    } else {
      ui.alert('⚠️ La API Key no puede estar vacía.');
    }
  }
}

function mostrarDialogoComando() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '⚡ Entrada rápida con IA (Gemini)',
    'Escribe tu comando en lenguaje natural (ej. "He terminado de entrenar", "Añadir hábito 20 min diario audio italiano", "Añadir tarea pagar factura"):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const textoUsuario = response.getResponseText().trim();
  if (!textoUsuario) {
    ui.alert('⚠️ Por favor, escribe una orden válida.');
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Analizando con Gemini...', '⏳ Procesando', 5);
    const ordenJson = procesarOrdenGemini(textoUsuario);
    const resultado = ejecutarAccion(ordenJson);

    ui.alert(
      '✅ Resultado de la operación',
      `Interpretación IA:\n• Acción: ${ordenJson.accion}\n• Tipo: ${ordenJson.tipo}\n• Nombre: ${ordenJson.nombre}\n• Categoría: ${ordenJson.categoria}\n\nRespuesta:\n${resultado.mensaje}`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('❌ Error al procesar la orden', error.message, ui.ButtonSet.OK);
  }
}

function generarHabitosDelDiaManual() {
  try {
    const res = generarHabitosDelDia();
    SpreadsheetApp.getUi().alert('✅ Volcado y Envío a Google Tasks completado', res.mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error en el volcado', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function sincronizarGoogleTasksManual() {
  try {
    const res = sincronizarBuzonGoogleTasks();
    SpreadsheetApp.getUi().alert('🔃 Sincronización Bidireccional (Google Tasks)', res.mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      '⚠️ Falta activar el servicio Google Tasks API',
      `Detalle: ${error.message}\n\nPara activarlo en 10 segundos:\n1. En el editor de Apps Script, mira la barra izquierda y haz clic en el símbolo "+" junto a "Servicios".\n2. Busca "Google Tasks API", selecciónalo y pulsa "Añadir".`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function actualizarKpiMensualEnSheetManual() {
  try {
    const estado = obtenerEstadoDashboard();
    SpreadsheetApp.getUi().alert(
      '📊 KPI Mensual Actualizado en "Dashboard"',
      `Se ha actualizado la tabla "📅 CUMPLIMIENTO POR MESES" en tu pestaña Dashboard (${estado.resumenGlobal.porMes.length} mes(es) registrados).`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error al actualizar KPI mensual', error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 2. INTEGRACIÓN NLU CON GEMINI API
 */
function procesarOrdenGemini(textoUsuario) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('No se ha configurado GEMINI_API_KEY. Usa el menú "⚡ Mi Gestor > Configurar Gemini API Key".');
  }

  const textoLimpio = limpiarPrefijoAsistente_(textoUsuario).texto;
  const itemsHoy = obtenerNombresRegistradosHoy_();
  const contextoHoy = itemsHoy.length > 0
    ? `Elementos ya registrados hoy en la hoja: [${itemsHoy.join(', ')}]. Si el usuario se refiere a alguno de estos para completarlo o borrarlo, utiliza EXACTAMENTE ese mismo nombre.`
    : 'Aún no hay elementos registrados hoy.';

  const systemInstructionText = `Eres el motor NLU de un sistema de gestión de hábitos y tareas personales.
Analiza la orden del usuario y devuélvela estrictamente en formato JSON cumpliendo este esquema:
{
  "accion": "CREAR" | "COMPLETAR" | "BORRAR" | "CONSULTAR",
  "tipo": "Hábito" | "Tarea",
  "nombre": "nombre limpio y conciso en infinitivo o sustantivo (con primera letra en mayúscula)",
  "categoria": "Salud" | "Estudio" | "Hogar" | "Mantenimiento" | "Finanzas" | "Trabajo",
  "frecuencia": "Diaria" | "L-V" | "Fin de semana" | "Semanal"
}

Reglas de negocio:
1. "accion":
   - "COMPLETAR": cuando el usuario indica que ya hizo, terminó o cumplió algo (ej. "ya entrené", "hecho leer", "completado meditación", "terminado audio italiano").
   - "CREAR": cuando quiere agendar, recordar o añadir un hábito o tarea nueva (ej. "20 minutos diario audio italiano", "comprar leche", "tengo que estudiar inglés").
   - "BORRAR": cuando quiere eliminar, cancelar o quitar un registro (ej. "borra comprar leche", "quita entrenar de hoy").
   - "CONSULTAR": cuando pregunta por su estado o resumen.
2. "tipo":
   - "Hábito": acciones recurrentes de desarrollo personal, idiomas, salud, lectura, deporte, rutinas o si menciona palabras como "diario", "todos los días", "hábito", "semanal".
   - "Tarea": acciones puntuales, recados, compras, pagos o trámites.
3. "nombre": limpia palabras de comando ("hecho", "completado", "añade") o frecuencias ("diario", "todos los días"). Si coincide semánticamente con uno de los elementos registrados hoy, usa EXACTAMENTE el nombre de la lista de hoy.
4. "categoria": clasifica semánticamente en una de las 6 categorías permitidas ("Salud", "Estudio", "Hogar", "Mantenimiento", "Finanzas", "Trabajo").
5. ${contextoHoy}`;

  const payload = {
    systemInstruction: {
      parts: [{ text: systemInstructionText }]
    },
    contents: [
      {
        role: 'user',
        parts: [{ text: textoLimpio }]
      }
    ],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          accion: {
            type: 'STRING',
            enum: ['CREAR', 'COMPLETAR', 'BORRAR', 'CONSULTAR']
          },
          tipo: {
            type: 'STRING',
            enum: ['Hábito', 'Tarea']
          },
          nombre: {
            type: 'STRING'
          },
          categoria: {
            type: 'STRING',
            enum: ['Salud', 'Estudio', 'Hogar', 'Mantenimiento', 'Finanzas', 'Trabajo']
          },
          frecuencia: {
            type: 'STRING',
            enum: ['Diaria', 'L-V', 'Fin de semana', 'Semanal']
          }
        },
        required: ['accion', 'tipo', 'nombre', 'categoria']
      }
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let ultimoError = '';
  for (let m = 0; m < CONFIG.GEMINI_MODELS.length; m++) {
    const modelName = CONFIG.GEMINI_MODELS[m];
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;

    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (statusCode === 200) {
      const jsonResponse = JSON.parse(responseText);
      const rawContent = jsonResponse?.candidates?.[0]?.content?.parts?.[0]?.text;
      if (rawContent) {
        return JSON.parse(rawContent);
      }
    }

    ultimoError = `HTTP ${statusCode} (${modelName}): ${responseText}`;
    if (statusCode === 400 || statusCode === 403) {
      break;
    }
  }

  throw new Error(`Error Gemini API (${ultimoError})`);
}

/**
 * 3. LÓGICA DE MANIPULACIÓN EN REGISTRO_DIARIO Y PLANTILLAS (CRUD COMPLETO)
 */
function ejecutarAccion(orden, opciones) {
  const opts = opciones || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = obtenerHojaFlexible_(ss, CONFIG.SHEETS.REGISTRO);
    const sheetPlantillas = obtenerHojaFlexible_(ss, CONFIG.SHEETS.PLANTILLAS);
    if (!sheet) throw new Error(`No existe la pestaña "${CONFIG.SHEETS.REGISTRO}".`);

    const hoyStr = obtenerFechaHoyStr_();
    const ahoraTimestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    const accion = (orden.accion || '').toUpperCase();
    const tipo = orden.tipo || 'Tarea';
    const nombre = (orden.nombre || '').trim();
    const categoria = orden.categoria || 'Hogar';
    const frecuencia = orden.frecuencia || 'Diaria';
    const horaMomento = orden.horaMomento || 'Cualquier momento';
    const guardarEnPlantilla = orden.guardarEnPlantilla !== false; // Por defecto true si es Hábito
    const nombreNorm = normalizarTexto_(nombre);

    const data = sheet.getDataRange().getValues();
    let mensaje = '';
    let filaAfectada = null;

    if (accion === 'COMPLETAR') {
      let encontradoHoyIndex = -1;

      if (orden.idFila && orden.idFila > 1 && orden.idFila <= data.length) {
        encontradoHoyIndex = Number(orden.idFila);
      } else {
        for (let i = data.length - 1; i >= 1; i--) {
          const fechaFila = formatearCeldaFecha_(data[i][0]);
          const nombreFilaNorm = normalizarTexto_(String(data[i][2]));

          if (fechaFila === hoyStr && (nombreFilaNorm === nombreNorm || nombreFilaNorm.includes(nombreNorm) || nombreNorm.includes(nombreFilaNorm))) {
            encontradoHoyIndex = i + 1;
            break;
          }
        }
      }

      let nombreReal = nombre;
      if (encontradoHoyIndex !== -1) {
        nombreReal = String(sheet.getRange(encontradoHoyIndex, 3).getValue());
        sheet.getRange(encontradoHoyIndex, 5, 1, 2).setValues([['Completado', ahoraTimestamp]]);
        filaAfectada = encontradoHoyIndex;
        mensaje = `✅ "${nombreReal}" marcado como Completado.`;
      } else {
        filaAfectada = insertarEnPrimeraFilaLibre_(sheet, [hoyStr, tipo, nombre, categoria, 'Completado', ahoraTimestamp], 3);
        mensaje = `✅ Añadido y completado hoy: "${nombre}" (${categoria}).`;
      }

      if (!opts.omitirSyncTasks) {
        marcarCompletadaEnGoogleTasks_(nombreReal);
      }
    }

    else if (accion === 'REABRIR') {
      let encontradoIndex = -1;
      if (orden.idFila && orden.idFila > 1 && orden.idFila <= data.length) {
        encontradoIndex = Number(orden.idFila);
      } else {
        for (let i = data.length - 1; i >= 1; i--) {
          const fechaFila = formatearCeldaFecha_(data[i][0]);
          const nombreFilaNorm = normalizarTexto_(String(data[i][2]));
          if (fechaFila === hoyStr && (nombreFilaNorm === nombreNorm || nombreFilaNorm.includes(nombreNorm))) {
            encontradoIndex = i + 1;
            break;
          }
        }
      }
      if (encontradoIndex !== -1) {
        const nombreReal = String(sheet.getRange(encontradoIndex, 3).getValue());
        sheet.getRange(encontradoIndex, 5, 1, 2).setValues([['Pendiente', '']]);
        filaAfectada = encontradoIndex;
        mensaje = `↩️ "${nombreReal}" reabierto como Pendiente.`;
        if (!opts.omitirSyncTasks) {
          asegurarPendienteEnGoogleTasks_(nombreReal, tipo, categoria);
        }
      } else {
        mensaje = `⚠️ No se encontró el elemento para reabrir.`;
      }
    }

    else if (accion === 'CREAR') {
      let existeHoy = false;
      for (let i = data.length - 1; i >= 1; i--) {
        if (formatearCeldaFecha_(data[i][0]) === hoyStr && normalizarTexto_(String(data[i][2])) === nombreNorm) {
          existeHoy = true;
          filaAfectada = i + 1;
          break;
        }
      }

      // Si es un Hábito y el día de hoy le aplica según su frecuencia (o se pidió explícitamente para hoy), lo insertamos en Registro_Diario
      const aplicaHoy = tipo === 'Tarea' || correspondeFrecuenciaHoy_(frecuencia, new Date().getDay()) || orden.forzarHoy === true;
      if (!existeHoy && aplicaHoy) {
        filaAfectada = insertarEnPrimeraFilaLibre_(sheet, [hoyStr, tipo, nombre, categoria, 'Pendiente', ''], 3);
        if (!opts.omitirSyncTasks) {
          asegurarPendienteEnGoogleTasks_(nombre, tipo, categoria);
        }
      }

      if (tipo === 'Hábito' && guardarEnPlantilla && sheetPlantillas) {
        asegurarEnPlantillasHabitos_(sheetPlantillas, nombre, categoria, frecuencia, horaMomento);
      }

      mensaje = tipo === 'Hábito'
        ? (aplicaHoy
            ? `📌 Hábito "${nombre}" guardado en Plantillas (${frecuencia}) y añadido a hoy.`
            : `📌 Hábito "${nombre}" guardado en Plantillas (${frecuencia}). Se volcará en los días correspondientes.`)
        : `📌 Tarea "${nombre}" (${categoria}) añadida para hoy.`;
    }

    else if (accion === 'EDITAR_HOY') {
      const idFila = Number(orden.idFila);
      if (!idFila || idFila < 2 || idFila > data.length) {
        throw new Error('Fila inválida para editar.');
      }
      const nombreAnterior = String(sheet.getRange(idFila, 3).getValue());
      sheet.getRange(idFila, 2, 1, 3).setValues([[tipo, nombre, categoria]]);
      filaAfectada = idFila;

      if (!opts.omitirSyncTasks && normalizarTexto_(nombreAnterior) !== nombreNorm) {
        eliminarDeGoogleTasks_(nombreAnterior);
        const estadoActual = String(sheet.getRange(idFila, 5).getValue());
        if (estadoActual.toLowerCase() === 'pendiente') {
          asegurarPendienteEnGoogleTasks_(nombre, tipo, categoria);
        }
      }
      mensaje = `✏️ Actualizado: "${nombre}" (${categoria}).`;
    }

    else if (accion === 'BORRAR') {
      let filaABorrar = -1;

      if (orden.idFila && orden.idFila > 1 && orden.idFila <= data.length) {
        filaABorrar = Number(orden.idFila);
      } else {
        for (let i = data.length - 1; i >= 1; i--) {
          const fechaFila = formatearCeldaFecha_(data[i][0]);
          const nombreFilaNorm = normalizarTexto_(String(data[i][2]));

          if (fechaFila === hoyStr && (nombreFilaNorm === nombreNorm || nombreFilaNorm.includes(nombreNorm))) {
            filaABorrar = i + 1;
            break;
          }
        }
      }

      if (filaABorrar !== -1) {
        const nombreBorrado = String(sheet.getRange(filaABorrar, 3).getValue());
        sheet.deleteRow(filaABorrar);
        filaAfectada = filaABorrar;
        if (!opts.omitirSyncTasks) {
          eliminarDeGoogleTasks_(nombreBorrado);
        }
        mensaje = `🗑️ Eliminado de hoy: "${nombreBorrado}".`;
      } else {
        mensaje = `⚠️ No se encontró "${nombre}" para borrar.`;
      }
    }

    else if (accion === 'CONSULTAR') {
      const resumen = obtenerEstadoDashboard();
      mensaje = `📊 Hoy llevas ${resumen.hoy.completados}/${resumen.hoy.total} (${resumen.hoy.porcentajeCumplimiento}%). Pendientes: ${resumen.hoy.pendientes.map(p => p.nombre).join(', ') || '¡Ninguno!'}`;
    }

    SpreadsheetApp.flush();

    return {
      ok: true,
      accionEjecutada: accion,
      ordenInterpretada: orden,
      filaAfectada: filaAfectada,
      mensaje: mensaje,
      dashboard: obtenerEstadoDashboard()
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * CRUD ESPECÍFICO SOBRE LA PESTAÑA "Plantillas_Habitos"
 * Columnas: [0: ID, 1: Nombre, 2: Categoría, 3: Frecuencia, 4: Hora/Momento, 5: Activo]
 */
function gestionarPlantillaHabito(operacion) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPlantillas = obtenerHojaFlexible_(ss, CONFIG.SHEETS.PLANTILLAS);
    const sheetRegistro = obtenerHojaFlexible_(ss, CONFIG.SHEETS.REGISTRO);
    if (!sheetPlantillas) throw new Error('No existe la pestaña Plantillas_Habitos.');

    const modo = (operacion.modo || '').toUpperCase();
    const idFila = Number(operacion.idFila);
    const nombre = (operacion.nombre || '').trim();
    const categoria = (operacion.categoria || 'Hogar').trim();
    const frecuencia = (operacion.frecuencia || 'Diaria').trim();
    const horaMomento = (operacion.horaMomento || 'Cualquier momento').trim();
    const activo = operacion.activo !== undefined ? Boolean(operacion.activo) : true;

    let mensaje = '';

    if (modo === 'CREAR_PLANTILLA') {
      if (!nombre) throw new Error('El nombre del hábito no puede estar vacío.');
      asegurarEnPlantillasHabitos_(sheetPlantillas, nombre, categoria, frecuencia, horaMomento);
      if (activo && correspondeFrecuenciaHoy_(frecuencia, new Date().getDay())) {
        ejecutarAccion({
          accion: 'CREAR',
          tipo: 'Hábito',
          nombre: nombre,
          categoria: categoria,
          frecuencia: frecuencia,
          guardarEnPlantilla: false,
          forzarHoy: true
        });
      }
      mensaje = `✅ Hábito "${nombre}" creado en el catálogo (${frecuencia}).`;
    }

    else if (modo === 'EDITAR_PLANTILLA') {
      if (!idFila || idFila < 2 || idFila > sheetPlantillas.getLastRow()) {
        throw new Error('Fila de plantilla no válida.');
      }
      const nombreAntiguo = String(sheetPlantillas.getRange(idFila, 2).getValue());
      sheetPlantillas.getRange(idFila, 2, 1, 5).setValues([[nombre, categoria, frecuencia, horaMomento, activo]]);

      // Si cambió el nombre o categoría, actualizamos también el registro de hoy si existía
      if (sheetRegistro) {
        const hoyStr = obtenerFechaHoyStr_();
        const regData = sheetRegistro.getDataRange().getValues();
        const normAntiguo = normalizarTexto_(nombreAntiguo);
        for (let i = regData.length - 1; i >= 1; i--) {
          if (formatearCeldaFecha_(regData[i][0]) === hoyStr && normalizarTexto_(String(regData[i][2])) === normAntiguo) {
            sheetRegistro.getRange(i + 1, 3, 1, 2).setValues([[nombre, categoria]]);
            if (normalizarTexto_(nombre) !== normAntiguo) {
              eliminarDeGoogleTasks_(nombreAntiguo);
              if (String(regData[i][4]).toLowerCase() === 'pendiente') {
                asegurarPendienteEnGoogleTasks_(nombre, 'Hábito', categoria);
              }
            }
            break;
          }
        }
      }
      mensaje = `✏️ Plantilla "${nombre}" actualizada correctamente.`;
    }

    else if (modo === 'TOGGLE_ACTIVO') {
      if (!idFila || idFila < 2 || idFila > sheetPlantillas.getLastRow()) {
        throw new Error('Fila de plantilla no válida.');
      }
      const nombreHab = String(sheetPlantillas.getRange(idFila, 2).getValue());
      sheetPlantillas.getRange(idFila, 6).setValue(activo);
      mensaje = activo
        ? `🟢 Hábito "${nombreHab}" activado.`
        : `⏸️ Hábito "${nombreHab}" pausado (no se volcará cada noche).`;
    }

    else if (modo === 'BORRAR_PLANTILLA') {
      if (!idFila || idFila < 2 || idFila > sheetPlantillas.getLastRow()) {
        throw new Error('Fila de plantilla no válida.');
      }
      const nombreBorrado = String(sheetPlantillas.getRange(idFila, 2).getValue());
      sheetPlantillas.deleteRow(idFila);

      // Si también pidió borrarlo de hoy:
      if (operacion.borrarDeHoy && sheetRegistro) {
        const hoyStr = obtenerFechaHoyStr_();
        const regData = sheetRegistro.getDataRange().getValues();
        const normBorr = normalizarTexto_(nombreBorrado);
        for (let i = regData.length - 1; i >= 1; i--) {
          if (formatearCeldaFecha_(regData[i][0]) === hoyStr && normalizarTexto_(String(regData[i][2])) === normBorr) {
            sheetRegistro.deleteRow(i + 1);
            break;
          }
        }
        eliminarDeGoogleTasks_(nombreBorrado);
      }
      mensaje = `🗑️ Hábito "${nombreBorrado}" eliminado del catálogo.`;
    }

    SpreadsheetApp.flush();

    return {
      ok: true,
      mensaje: mensaje,
      dashboard: obtenerEstadoDashboard()
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 4. MOTOR BIDIRECCIONAL CON GOOGLE TASKS (LISTA "Hábitos")
 */
function sincronizarBuzonGoogleTasks() {
  if (typeof Tasks === 'undefined') {
    throw new Error('El servicio avanzado "Google Tasks API" no está habilitado en Apps Script.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRegistro = obtenerHojaFlexible_(ss, CONFIG.SHEETS.REGISTRO);
  if (!sheetRegistro) throw new Error('No se encontró la hoja Registro_Diario.');

  const hoyStr = obtenerFechaHoyStr_();
  const listaHabitos = obtenerOCrearListaHabitosTasks_();
  const targetListId = listaHabitos.id;

  const dataSheet = sheetRegistro.getDataRange().getValues();
  const itemsHoySheet = [];

  for (let i = 1; i < dataSheet.length; i++) {
    if (formatearCeldaFecha_(dataSheet[i][0]) === hoyStr) {
      const nombre = String(dataSheet[i][2] || '').trim();
      if (!nombre) continue;
      itemsHoySheet.push({
        indexFila: i + 1,
        tipo: String(dataSheet[i][1] || 'Hábito'),
        nombre: nombre,
        nombreNorm: normalizarTexto_(nombre),
        categoria: String(dataSheet[i][3] || 'General'),
        estado: String(dataSheet[i][4] || 'Pendiente')
      });
    }
  }

  function buscarCoincidenciaEnSheet(tituloNorm) {
    if (!tituloNorm) return null;
    let exacto = itemsHoySheet.find(it => it.nombreNorm === tituloNorm);
    if (exacto) return exacto;
    let parcial = itemsHoySheet.find(it =>
      it.nombreNorm.includes(tituloNorm) || tituloNorm.includes(it.nombreNorm)
    );
    if (parcial) return parcial;
    const palabrasTask = tituloNorm.split(/\s+/).filter(p => p.length >= 4 && !['para', 'como', 'este', 'esta', 'todo', 'todos', 'diario', 'diaria', 'minutos', 'hora'].includes(p));
    if (palabrasTask.length > 0) {
      return itemsHoySheet.find(it => palabrasTask.some(p => it.nombreNorm.includes(p)));
    }
    return null;
  }

  const listasARevisar = [{ id: targetListId, title: listaHabitos.title, esPrincipalHabitos: true }];
  try {
    const todasLasListasRes = Tasks.Tasklists.list({ maxResults: 100 });
    const todasLasListas = todasLasListasRes.items || [];
    todasLasListas.forEach(lista => {
      if (lista && lista.id && lista.id !== targetListId) {
        listasARevisar.push({ id: lista.id, title: lista.title || 'Secundaria', esPrincipalHabitos: false });
      }
    });
  } catch (e) {}

  const detallesCambios = [];
  const nombresEnTasksPendientesHoy = new Set();
  const ahoraTimestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

  for (let l = 0; l < listasARevisar.length; l++) {
    const infoLista = listasARevisar[l];
    const tasksRes = Tasks.Tasks.list(infoLista.id, {
      showCompleted: true,
      showHidden: true,
      maxResults: 100
    });
    const tasksItems = tasksRes.items || [];

    for (let i = 0; i < tasksItems.length; i++) {
      const t = tasksItems[i];
      const tituloRaw = (t.title || '').trim();
      if (!tituloRaw) continue;

      const infoPrefijo = limpiarPrefijoAsistente_(tituloRaw);
      const titulo = infoPrefijo.texto || tituloRaw;
      const tienePrefijoAsistente = infoPrefijo.teniaPrefijo;

      const tituloNorm = normalizarTexto_(titulo);
      const estaCompletadaEnTasks = (t.status === 'completed' || Boolean(t.completed));

      let fechaCompletadoTaskStr = '';
      if (t.completed) {
        fechaCompletadoTaskStr = Utilities.formatDate(new Date(t.completed), CONFIG.TIMEZONE, 'yyyy-MM-dd');
      }
      const completadaHoyEnTasks = estaCompletadaEnTasks && (fechaCompletadoTaskStr === hoyStr);
      const completadaEnDiaAnterior = estaCompletadaEnTasks && (fechaCompletadoTaskStr !== '' && fechaCompletadoTaskStr !== hoyStr);

      const itemSheet = tienePrefijoAsistente ? null : buscarCoincidenciaEnSheet(tituloNorm);

      if (completadaEnDiaAnterior) {
        if (infoLista.esPrincipalHabitos || itemSheet || tienePrefijoAsistente) {
          try {
            Tasks.Tasks.remove(infoLista.id, t.id);
          } catch (eRemove) {
            Logger.log(`Aviso: No se pudo eliminar de la lista ${infoLista.title} (${t.id}): ${eRemove.message}`);
          }
        }
        continue;
      }

      if (!infoLista.esPrincipalHabitos && !tienePrefijoAsistente) {
        if (itemSheet && completadaHoyEnTasks) {
          if (itemSheet.estado.toLowerCase() !== 'completado') {
            sheetRegistro.getRange(itemSheet.indexFila, 5, 1, 2).setValues([['Completado', ahoraTimestamp]]);
            itemSheet.estado = 'Completado';
            detallesCambios.push(`✅ Completado hoy (detectado en lista secundaria): "${itemSheet.nombre}"`);
          }
          try {
            Tasks.Tasks.remove(infoLista.id, t.id);
          } catch (eRemove) {
            Logger.log(`Aviso: No se pudo eliminar de la lista ${infoLista.title} (${t.id}): ${eRemove.message}`);
          }
        }
        continue;
      }

      if (itemSheet) {
        const mencionaCompletarEnTitulo = tienePrefijoAsistente && /^(completad[oa]|hecho|terminad[oa]|ya\s+h[eé])\b/i.test(titulo);
        if (completadaHoyEnTasks || mencionaCompletarEnTitulo) {
          if (itemSheet.estado.toLowerCase() !== 'completado') {
            sheetRegistro.getRange(itemSheet.indexFila, 5, 1, 2).setValues([['Completado', ahoraTimestamp]]);
            itemSheet.estado = 'Completado';
            detallesCambios.push(`✅ Completado en Google Sheets: "${itemSheet.nombre}"`);
          }
          try {
            Tasks.Tasks.remove(infoLista.id, t.id);
          } catch (eRemove) {
            Logger.log(`Aviso: No se pudo eliminar de la lista ${infoLista.title} (${t.id}): ${eRemove.message}`);
          }
        } else if (!estaCompletadaEnTasks) {
          if (itemSheet.estado.toLowerCase() === 'completado') {
            try {
            Tasks.Tasks.remove(infoLista.id, t.id);
          } catch (eRemove) {
            Logger.log(`Aviso: No se pudo eliminar de la lista ${infoLista.title} (${t.id}): ${eRemove.message}`);
          }
          } else if (!infoLista.esPrincipalHabitos || tienePrefijoAsistente) {
            // Si entró en lista secundaria con prefijo o tenía prefijo en el título, eliminamos la versión con prefijo para que se republique limpia en "Hábitos"
            try {
            Tasks.Tasks.remove(infoLista.id, t.id);
          } catch (eRemove) {
            Logger.log(`Aviso: No se pudo eliminar de la lista ${infoLista.title} (${t.id}): ${eRemove.message}`);
          }
          } else {
            nombresEnTasksPendientesHoy.add(itemSheet.nombreNorm);
          }
        }
      } else {
        try {
          const textoParaGemini = estaCompletadaEnTasks ? `Completado: ${titulo}` : titulo;
          const orden = procesarOrdenGemini(textoParaGemini);
          if (orden.accion === 'CREAR') {
            orden.forzarHoy = true; // Si el usuario lo acaba de crear hoy por voz, incluirlo también en hoy
          }
          const resAccion = ejecutarAccion(orden, { omitirSyncTasks: true });
          detallesCambios.push(`🤖 IA ("${titulo}") → ${resAccion.mensaje}`);

          try {
            Tasks.Tasks.remove(infoLista.id, t.id);
          } catch (eRemove) {
            Logger.log(`Aviso: No se pudo eliminar de la lista ${infoLista.title} (${t.id}): ${eRemove.message}`);
          }

          if (orden.accion === 'CREAR' && resAccion.filaAfectada) {
            itemsHoySheet.push({
              indexFila: resAccion.filaAfectada,
              tipo: orden.tipo,
              nombre: orden.nombre,
              nombreNorm: normalizarTexto_(orden.nombre),
              categoria: orden.categoria,
              estado: 'Pendiente'
            });
          }
        } catch (err) {
          detallesCambios.push(`⚠️ Error procesando "${tituloRaw}": ${err.message}`);
          Logger.log(`Error interpretando tarea "${titulo}": ${err.message}`);
        }
      }
    }
  }

  let publicadosEnTasks = 0;
  itemsHoySheet.forEach(itemSheet => {
    if (itemSheet.estado.toLowerCase() === 'pendiente' && !nombresEnTasksPendientesHoy.has(itemSheet.nombreNorm)) {
      try {
        Tasks.Tasks.insert(
          {
            title: itemSheet.nombre,
            notes: `[${itemSheet.tipo} · ${itemSheet.categoria}] (${hoyStr})`,
            due: construirDueDateRfc3339_(hoyStr)
          },
          targetListId
        );
        publicadosEnTasks++;
      } catch (e) {
        Logger.log(`Error insertando en Tasks "${itemSheet.nombre}": ${e.message}`);
      }
    }
  });

  if (publicadosEnTasks > 0) {
    detallesCambios.push(`📤 ${publicadosEnTasks} hábitos pendientes enviados a la lista "${listaHabitos.title}" de Google Tasks.`);
  }

  SpreadsheetApp.flush();

  const resumenListas = listasARevisar.map(l => l.title).join(', ');
  return {
    ok: true,
    cambios: detallesCambios.length,
    mensaje: detallesCambios.length > 0
      ? detallesCambios.join('\n')
      : `Sincronización verificada en listas [${resumenListas}] (${itemsHoySheet.length} hábitos hoy en hoja). No se detectaron cambios nuevos.`
  };
}

/**
 * 5. AUTOMATIZACIÓN Y RECURRENCIA NOCTURNA (generarHabitosDelDia)
 */
function generarHabitosDelDia() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPlantillas = ss.getSheetByName(CONFIG.SHEETS.PLANTILLAS);
    const sheetRegistro = ss.getSheetByName(CONFIG.SHEETS.REGISTRO);

    if (!sheetPlantillas || !sheetRegistro) {
      throw new Error('Faltan las pestañas requeridas ("Plantillas_Habitos" o "Registro_Diario").');
    }

    const hoy = new Date();
    const hoyStr = obtenerFechaHoyStr_(hoy);
    const diaSemana = hoy.getDay();

    const datosRegistro = sheetRegistro.getDataRange().getValues();
    const existentesHoy = new Set();
    const tareasPendientesPrevias = [];

    for (let i = 1; i < datosRegistro.length; i++) {
      const fechaFila = formatearCeldaFecha_(datosRegistro[i][0]);
      const tipoFila = String(datosRegistro[i][1] || 'Hábito').trim();
      const nombreFila = String(datosRegistro[i][2] || '').trim();
      const catFila = String(datosRegistro[i][3] || 'General').trim();
      const estadoFila = String(datosRegistro[i][4] || 'Pendiente').trim().toLowerCase();
      const nombreFilaNorm = normalizarTexto_(nombreFila);

      if (fechaFila === hoyStr && nombreFilaNorm) {
        existentesHoy.add(nombreFilaNorm);
      }

      // Si es una Tarea puntual de un día anterior que quedó Pendiente, la arrastramos automáticamente a hoy
      if (fechaFila && fechaFila < hoyStr && tipoFila === 'Tarea' && estadoFila === 'pendiente' && nombreFilaNorm) {
        tareasPendientesPrevias.push({
          nombre: nombreFila,
          nombreNorm: nombreFilaNorm,
          categoria: catFila
        });
      }
    }

    const datosPlantillas = sheetPlantillas.getDataRange().getValues();
    const nuevasFilas = [];

    // 1. Arrastrar primero las Tareas puntuales pendientes de días previos (rollover)
    tareasPendientesPrevias.forEach(tp => {
      if (!existentesHoy.has(tp.nombreNorm)) {
        nuevasFilas.push([hoyStr, 'Tarea', tp.nombre, tp.categoria, 'Pendiente', '']);
        existentesHoy.add(tp.nombreNorm);
      }
    });

    // 2. Volcar los Hábitos del catálogo que correspondan a la frecuencia del día de hoy
    for (let i = 1; i < datosPlantillas.length; i++) {
      const fila = datosPlantillas[i];
      const nombreHabito = String(fila[1] || '').trim();
      const categoria = String(fila[2] || 'General').trim();
      const frecuencia = String(fila[3] || 'Diaria').trim();
      const activoRaw = fila[5];

      if (!nombreHabito) continue;

      const estaActivo =
        activoRaw === true ||
         ['si', 'sí', 'true', '1', 'activo', 'x'].includes(String(activoRaw).trim().toLowerCase());

      if (!estaActivo) continue;

      if (correspondeFrecuenciaHoy_(frecuencia, diaSemana)) {
        const nombreNorm = normalizarTexto_(nombreHabito);
        if (!existentesHoy.has(nombreNorm)) {
          nuevasFilas.push([hoyStr, 'Hábito', nombreHabito, categoria, 'Pendiente', '']);
          existentesHoy.add(nombreNorm);
        }
      }
    }

    if (nuevasFilas.length > 0) {
      let ultimaConNombre = 1;
      for (let i = 1; i < datosRegistro.length; i++) {
        if (String(datosRegistro[i][2] || '').trim() !== '') {
          ultimaConNombre = i + 1;
        }
      }
      sheetRegistro
        .getRange(ultimaConNombre + 1, 1, nuevasFilas.length, nuevasFilas[0].length)
        .setValues(nuevasFilas);
      aplicarValidacionesRegistro_(sheetRegistro, ultimaConNombre + 1, nuevasFilas.length);
      SpreadsheetApp.flush();
    }

    let msgTasks = '';
    try {
      const syncRes = sincronizarBuzonGoogleTasks();
      msgTasks = `\n\n📱 Google Tasks: ${syncRes.mensaje}`;
    } catch (e) {
      msgTasks = `\n(Nota: Activa el servicio Google Tasks API para verlos en la app Tareas).`;
    }

    return {
      ok: true,
      fecha: hoyStr,
      insertados: nuevasFilas.length,
      mensaje: `Se han volcado ${nuevasFilas.length} hábitos nuevos para hoy (${hoyStr}).${msgTasks}`
    };
  } finally {
    lock.releaseLock();
  }
}

function correspondeFrecuenciaHoy_(frecuenciaStr, diaSemana) {
  const freq = normalizarTexto_(frecuenciaStr);

  if (freq === 'diaria' || freq === 'diario' || freq === 'todos los dias') {
    return true;
  }
  if (freq === 'l-v' || freq === 'lunes a viernes' || freq === 'laborables') {
    return diaSemana >= 1 && diaSemana <= 5;
  }
  if (freq === 'fin de semana' || freq === 's-d' || freq === 'sabado y domingo') {
    return diaSemana === 0 || diaSemana === 6;
  }
  if (freq.startsWith('semanal')) {
    if (freq.includes('domingo')) return diaSemana === 0;
    if (freq.includes('martes')) return diaSemana === 2;
    if (freq.includes('miercoles')) return diaSemana === 3;
    if (freq.includes('jueves')) return diaSemana === 4;
    if (freq.includes('viernes')) return diaSemana === 5;
    if (freq.includes('sabado')) return diaSemana === 6;
    return diaSemana === 1;
  }

  return false;
}

function instalarTriggersAutomaticos() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    const fn = trigger.getHandlerFunction();
    if (fn === 'generarHabitosDelDia' || fn === 'sincronizarBuzonGoogleTasks') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('generarHabitosDelDia')
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .nearMinute(5)
    .create();

  ScriptApp.newTrigger('sincronizarBuzonGoogleTasks')
    .timeBased()
    .everyMinutes(5)
    .create();

  try {
    SpreadsheetApp.getUi().alert(
      '⏰ Automatización Bidireccional Activada',
      '1. Cada noche (00:05 AM) tus hábitos se vuelcan en la hoja y aparecen solos en tu lista "Hábitos" de Google Tasks.\n2. Cada 5 minutos, cualquier hábito que marques con ✓ en Google Tasks se marcará como Completado en tu Google Sheet.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log('Triggers instalados correctamente.');
  }
}

function instalarTriggerDiario() {
  instalarTriggersAutomaticos();
}

/**
 * FUNCIONES AUXILIARES DE GOOGLE TASKS Y PLANTILLAS
 */
function obtenerOCrearListaHabitosTasks_() {
  const taskLists = Tasks.Tasklists.list({ maxResults: 100 });
  const listas = taskLists.items || [];
  const nombreBuscado = normalizarTexto_(CONFIG.TASKS_LIST_NAME);
  let listaHabitos = listas.find(l => normalizarTexto_(l.title) === nombreBuscado);

  if (!listaHabitos) {
    listaHabitos = Tasks.Tasklists.insert({ title: CONFIG.TASKS_LIST_NAME });
  }
  return listaHabitos;
}

function construirDueDateRfc3339_(fechaStr) {
  const f = fechaStr || obtenerFechaHoyStr_();
  return `${f}T00:00:00.000Z`;
}

function asegurarPendienteEnGoogleTasks_(nombre, tipo, categoria, fechaStr) {
  try {
    if (typeof Tasks === 'undefined') return;
    const fStr = fechaStr || obtenerFechaHoyStr_();
    const lista = obtenerOCrearListaHabitosTasks_();
    const res = Tasks.Tasks.list(lista.id, { showCompleted: false, maxResults: 100 });
    const items = res.items || [];
    const norm = normalizarTexto_(nombre);
    const yaExiste = items.some(t => normalizarTexto_(t.title) === norm);
    if (!yaExiste) {
      Tasks.Tasks.insert(
        {
          title: nombre,
          notes: `[${tipo} · ${categoria}] (${fStr})`,
          due: construirDueDateRfc3339_(fStr)
        },
        lista.id
      );
    }
  } catch (e) {}
}

function marcarCompletadaEnGoogleTasks_(nombre) {
  try {
    if (typeof Tasks === 'undefined') return;
    const lista = obtenerOCrearListaHabitosTasks_();
    const res = Tasks.Tasks.list(lista.id, { showCompleted: true, showHidden: true, maxResults: 100 });
    const items = res.items || [];
    const norm = normalizarTexto_(nombre);
    items.forEach(t => {
      const tNorm = normalizarTexto_(t.title);
      if (tNorm === norm || tNorm.includes(norm) || norm.includes(tNorm)) {
        try { Tasks.Tasks.remove(lista.id, t.id); } catch (e) {}
      }
    });
  } catch (e) {}
}

function eliminarDeGoogleTasks_(nombre) {
  marcarCompletadaEnGoogleTasks_(nombre);
}

function asegurarEnPlantillasHabitos_(sheetPlantillas, nombre, categoria, frecuencia, horaMomento) {
  const datos = sheetPlantillas.getDataRange().getValues();
  const norm = normalizarTexto_(nombre);

  // 1. Recoger todas las filas reales con nombre y compactar huecos en blanco
  const filasValidas = [];
  let maxNumeroId = 0;

  for (let i = 1; i < datos.length; i++) {
    const nombreFila = String(datos[i][1] || '').trim();
    if (nombreFila) {
      const idxPrevia = filasValidas.findIndex(f => normalizarTexto_(String(f[1])) === normalizarTexto_(nombreFila));
      if (idxPrevia === -1) {
        const filaCopia = [...datos[i]];
        const idActual = String(filaCopia[0] || '').trim();

        // Extraer el número si sigue el formato H01, H05, etc.
        const matchH = idActual.match(/^H(\d+)$/i);
        if (matchH) {
          const num = parseInt(matchH[1], 10);
          if (num > maxNumeroId) maxNumeroId = num;
        }
        filasValidas.push(filaCopia);
      }
    }
  }

  // Normalizar cualquier ID que tuviera el formato antiguo "HAB-..." o estuviera vacío al formato "H01", "H06"...
  filasValidas.forEach((f) => {
    const idStr = String(f[0] || '').trim();
    if (!/^H\d+$/i.test(idStr)) {
      maxNumeroId++;
      f[0] = `H${String(maxNumeroId).padStart(2, '0')}`;
    }
  });

  const yaExistia = filasValidas.some(f => normalizarTexto_(String(f[1])) === norm);
  if (!yaExistia) {
    maxNumeroId++;
    const nuevoId = `H${String(maxNumeroId).padStart(2, '0')}`;
    filasValidas.push([
      nuevoId,
      nombre,
      categoria || 'Hogar',
      frecuencia || 'Diaria',
      horaMomento || 'Cualquier momento',
      true
    ]);
  }

  // Limpiar contenido y validaciones huérfanas antiguas desde fila 2 y escribir todas las plantillas seguidas sin huecos
  const maxFilasHoja = Math.max(datos.length, filasValidas.length + 1);
  if (maxFilasHoja > 1) {
    const rangoTotal = sheetPlantillas.getRange(2, 1, maxFilasHoja - 1, 6);
    rangoTotal.clearContent();
    rangoTotal.clearDataValidations();
  }

  if (filasValidas.length > 0) {
    sheetPlantillas.getRange(2, 1, filasValidas.length, 6).setValues(filasValidas);
    aplicarValidacionesPlantillas_(sheetPlantillas, filasValidas.length);
  }
}

/**
 * Aplica automáticamente los desplegables de Categoría, Frecuencia y casilla Activo en Plantillas_Habitos
 */
function aplicarValidacionesPlantillas_(sheetPlantillas, numFilasDatos) {
  if (!sheetPlantillas || numFilasDatos <= 0) return;

  const ruleCategoria = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.CATEGORIAS, true)
    .setAllowInvalid(true)
    .build();

  const ruleFrecuencia = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.FRECUENCIAS, true)
    .setAllowInvalid(true)
    .build();

  const ruleCheckbox = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();

  // Columna C (3): Categoría | Columna D (4): Frecuencia | Columna F (6): Activo
  sheetPlantillas.getRange(2, 3, numFilasDatos, 1).setDataValidation(ruleCategoria);
  sheetPlantillas.getRange(2, 4, numFilasDatos, 1).setDataValidation(ruleFrecuencia);
  sheetPlantillas.getRange(2, 6, numFilasDatos, 1).setDataValidation(ruleCheckbox);
}

/**
 * Aplica automáticamente los desplegables de Tipo, Categoría y Estado en una o varias filas de Registro_Diario
 */
function aplicarValidacionesRegistro_(sheetRegistro, filaInicio, numFilas) {
  if (!sheetRegistro || numFilas <= 0) return;

  const ruleTipo = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Hábito', 'Tarea'], true)
    .setAllowInvalid(true)
    .build();

  const ruleCategoria = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.CATEGORIAS, true)
    .setAllowInvalid(true)
    .build();

  const ruleEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.ESTADOS, true)
    .setAllowInvalid(true)
    .build();

  // Columna B (2): Tipo | Columna D (4): Categoría | Columna E (5): Estado
  sheetRegistro.getRange(filaInicio, 2, numFilas, 1).setDataValidation(ruleTipo);
  sheetRegistro.getRange(filaInicio, 4, numFilas, 1).setDataValidation(ruleCategoria);
  sheetRegistro.getRange(filaInicio, 5, numFilas, 1).setDataValidation(ruleEstado);
}

/**
 * 6. ENDPOINTS WEBHOOK Y APLICACIÓN VISUAL (doGet / doPost)
 */
function doGet(e) {
  try {
    try { sincronizarBuzonGoogleTasks(); } catch (ignore) {}

    if (e && e.parameter && e.parameter.format === 'json') {
      return respuestaJson_({
        ok: true,
        dashboard: obtenerEstadoDashboard()
      });
    }

    return HtmlService.createHtmlOutput(obtenerHtmlApp_())
      .setTitle('⚡ Mi Gestor de Hábitos y Tareas')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  } catch (error) {
    return respuestaJson_({ ok: false, error: error.message });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return respuestaJson_({ ok: false, error: 'Payload vacío. Se esperaba un cuerpo JSON.' });
    }

    const body = JSON.parse(e.postData.contents);

    const expectedToken = PropertiesService.getScriptProperties().getProperty('WEBHOOK_TOKEN');
    if (expectedToken && body.token !== expectedToken) {
      return respuestaJson_({ ok: false, error: 'No autorizado: Token inválido.' });
    }

    if (body.comando === 'DASHBOARD') {
      try { sincronizarBuzonGoogleTasks(); } catch (ignore) {}
      return respuestaJson_({
        ok: true,
        mensaje: 'Métricas obtenidas correctamente',
        dashboard: obtenerEstadoDashboard()
      });
    }

    if (body.comando === 'GENERAR_HABITOS_HOY') {
      const resultadoVolcado = generarHabitosDelDia();
      return respuestaJson_({
        ...resultadoVolcado,
        dashboard: obtenerEstadoDashboard()
      });
    }

    if (body.accionDirecta && body.accionDirecta.accion) {
      const resultadoDirecto = ejecutarAccion(body.accionDirecta);
      return respuestaJson_(resultadoDirecto);
    }

    const textoUsuario = (body.mensaje || '').trim();
    if (!textoUsuario) {
      return respuestaJson_({
        ok: false,
        error: 'El campo "mensaje" está vacío.'
      });
    }

    const ordenInterpretada = procesarOrdenGemini(textoUsuario);
    const resultadoEjecucion = ejecutarAccion(ordenInterpretada);

    return respuestaJson_(resultadoEjecucion);

  } catch (error) {
    return respuestaJson_({
      ok: false,
      error: error.message,
      stack: error.stack
    });
  }
}

/**
 * Funciones puente llamadas desde la interfaz web (google.script.run)
 */
function apiObtenerDatos() {
  try { sincronizarBuzonGoogleTasks(); } catch (ignore) {}
  return obtenerEstadoDashboard();
}

function apiEnviarComandoVoz(texto) {
  const orden = procesarOrdenGemini(texto);
  return ejecutarAccion(orden);
}

function apiAccionDirecta(orden) {
  return ejecutarAccion(orden);
}

function apiGestionarPlantilla(operacion) {
  return gestionarPlantillaHabito(operacion);
}

function apiVolcarHabitosHoy() {
  const res = generarHabitosDelDia();
  return {
    ...res,
    dashboard: obtenerEstadoDashboard()
  };
}

function obtenerEstadoDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRegistro = obtenerHojaFlexible_(ss, CONFIG.SHEETS.REGISTRO);
  const sheetPlantillas = obtenerHojaFlexible_(ss, CONFIG.SHEETS.PLANTILLAS);
  const sheetDashboard = obtenerHojaFlexible_(ss, CONFIG.SHEETS.DASHBOARD);

  const hoyStr = obtenerFechaHoyStr_();
  const data = sheetRegistro ? sheetRegistro.getDataRange().getValues() : [];

  let totalHoy = 0;
  let completadosHoy = 0;
  const pendientesHoy = [];
  const completadosListaHoy = [];
  const porCategoria = {};

  for (let i = 1; i < data.length; i++) {
    const fechaFila = formatearCeldaFecha_(data[i][0]);
    if (fechaFila !== hoyStr) continue;

    const tipo = String(data[i][1] || 'Tarea');
    const nombre = String(data[i][2] || '');
    const categoria = String(data[i][3] || 'General');
    const estado = String(data[i][4] || 'Pendiente');
    const timestamp = data[i][5] ? String(data[i][5]) : null;

    totalHoy++;
    if (!porCategoria[categoria]) {
      porCategoria[categoria] = { total: 0, completados: 0, porcentaje: 0 };
    }
    porCategoria[categoria].total++;

    if (estado.toLowerCase() === 'completado') {
      completadosHoy++;
      porCategoria[categoria].completados++;
      completadosListaHoy.push({
        idFila: i + 1,
        tipo,
        nombre,
        categoria,
        timestamp
      });
    } else {
      pendientesHoy.push({
        idFila: i + 1,
        tipo,
        nombre,
        categoria
      });
    }
  }

  Object.keys(porCategoria).forEach(cat => {
    const c = porCategoria[cat];
    c.porcentaje = c.total > 0 ? Math.round((c.completados / c.total) * 100) : 0;
  });

  const porcentajeHoy = totalHoy > 0 ? Math.round((completadosHoy / totalHoy) * 100) : 0;

  // Leer catálogo completo de Plantillas_Habitos para la pestaña CRUD de la Web App
  const plantillasCatalogo = [];
  if (sheetPlantillas) {
    const pData = sheetPlantillas.getDataRange().getValues();
    for (let i = 1; i < pData.length; i++) {
      const nombreHab = String(pData[i][1] || '').trim();
      if (!nombreHab) continue;
      const activoRaw = pData[i][5];
      const estaActivo =
        activoRaw === true ||
        ['si', 'sí', 'true', '1', 'activo', 'x'].includes(String(activoRaw).trim().toLowerCase());

      plantillasCatalogo.push({
        idFila: i + 1,
        idCodigo: String(pData[i][0] || `H${String(i).padStart(2, '0')}`),
        nombre: nombreHab,
        categoria: String(pData[i][2] || 'General'),
        frecuencia: String(pData[i][3] || 'Diaria'),
        horaMomento: String(pData[i][4] || 'Cualquier momento'),
        activo: estaActivo
      });
    }
  }

  let snapshotDashboardSheet = [];
  if (sheetDashboard && sheetDashboard.getLastRow() > 0 && sheetDashboard.getLastColumn() > 0) {
    const maxRows = Math.min(sheetDashboard.getLastRow(), 40);
    const maxCols = Math.min(sheetDashboard.getLastColumn(), 10);
    snapshotDashboardSheet = sheetDashboard.getRange(1, 1, maxRows, maxCols).getDisplayValues();
  }

  // Calcular resumen histórico de los últimos 7 días, por mes y por categoría desde Registro_Diario
  const hace7Dias = new Date();
  hace7Dias.setDate(hace7Dias.getDate() - 6);
  const hace7DiasStr = obtenerFechaHoyStr_(hace7Dias);
  const mesActualClave = hoyStr.substring(0, 7); // YYYY-MM

  let total7Dias = 0;
  let completados7Dias = 0;
  let totalHistorico = 0;
  let completadosHistorico = 0;
  const historicoPorCategoria = {};
  const mapaMeses = {};

  CONFIG.CATEGORIAS.forEach(cat => {
    historicoPorCategoria[cat] = { total: 0, completados: 0, porcentaje: 0 };
  });

  // Asegurar que el mes actual exista siempre en el desglose mensual
  mapaMeses[mesActualClave] = {
    clave: mesActualClave,
    etiqueta: formatearEtiquetaMes_(mesActualClave),
    total: 0,
    completados: 0,
    pendientes: 0,
    porcentaje: 0
  };

  for (let i = 1; i < data.length; i++) {
    const fechaFila = formatearCeldaFecha_(data[i][0]);
    const nombreFila = String(data[i][2] || '').trim();
    if (!fechaFila || !nombreFila) continue;

    const catFila = String(data[i][3] || 'Hogar').trim();
    const esCompletado = String(data[i][4] || '').trim().toLowerCase() === 'completado';
    const claveMes = fechaFila.substring(0, 7);

    totalHistorico++;
    if (esCompletado) completadosHistorico++;

    if (fechaFila >= hace7DiasStr && fechaFila <= hoyStr) {
      total7Dias++;
      if (esCompletado) completados7Dias++;
    }

    if (!historicoPorCategoria[catFila]) {
      historicoPorCategoria[catFila] = { total: 0, completados: 0, porcentaje: 0 };
    }
    historicoPorCategoria[catFila].total++;
    if (esCompletado) historicoPorCategoria[catFila].completados++;

    if (/^\d{4}-\d{2}$/.test(claveMes)) {
      if (!mapaMeses[claveMes]) {
        mapaMeses[claveMes] = {
          clave: claveMes,
          etiqueta: formatearEtiquetaMes_(claveMes),
          total: 0,
          completados: 0,
          pendientes: 0,
          porcentaje: 0
        };
      }
      mapaMeses[claveMes].total++;
      if (esCompletado) {
        mapaMeses[claveMes].completados++;
      } else {
        mapaMeses[claveMes].pendientes++;
      }
    }
  }

  Object.keys(historicoPorCategoria).forEach(cat => {
    const itemCat = historicoPorCategoria[cat];
    itemCat.porcentaje = itemCat.total > 0 ? Math.round((itemCat.completados / itemCat.total) * 100) : 0;
  });

  const listaPorMes = Object.keys(mapaMeses)
    .sort((a, b) => b.localeCompare(a))
    .map(k => {
      const m = mapaMeses[k];
      m.porcentaje = m.total > 0 ? Math.round((m.completados / m.total) * 100) : 0;
      return m;
    });

  const statsMesActual = mapaMeses[mesActualClave] || {
    clave: mesActualClave,
    etiqueta: formatearEtiquetaMes_(mesActualClave),
    total: 0,
    completados: 0,
    pendientes: 0,
    porcentaje: 0
  };

  // Sincronizar automáticamente el bloque de KPI Mensual en la pestaña "Dashboard" de Google Sheets
  if (sheetDashboard) {
    try {
      sincronizarTablaMensualEnHojaDashboard_(sheetDashboard, listaPorMes);
    } catch (e) {
      Logger.log('Aviso sincronizando tabla mensual en hoja Dashboard: ' + e.message);
    }
  }

  const porcentaje7Dias = total7Dias > 0 ? Math.round((completados7Dias / total7Dias) * 100) : 0;
  const porcentajeHistorico = totalHistorico > 0 ? Math.round((completadosHistorico / totalHistorico) * 100) : 0;

  // URL privada dinámica hacia la pestaña Dashboard de Google Sheets (sin exponer IDs fijos en el código fuente)
  let sheetDashboardUrl = ss.getUrl();
  if (sheetDashboard) {
    sheetDashboardUrl += `#gid=${sheetDashboard.getSheetId()}`;
  }

  return {
    fecha: hoyStr,
    actualizadoEn: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'HH:mm'),
    hoy: {
      total: totalHoy,
      completados: completadosHoy,
      pendientesCount: pendientesHoy.length,
      porcentajeCumplimiento: porcentajeHoy,
      pendientes: pendientesHoy,
      completadosLista: completadosListaHoy,
      porCategoria: porCategoria
    },
    resumenGlobal: {
      semana: {
        total: total7Dias,
        completados: completados7Dias,
        porcentaje: porcentaje7Dias
      },
      mesActual: statsMesActual,
      porMes: listaPorMes,
      historico: {
        total: totalHistorico,
        completados: completadosHistorico,
        porcentaje: porcentajeHistorico
      },
      porCategoria: historicoPorCategoria
    },
    plantillas: plantillasCatalogo,
    vistaDashboard: snapshotDashboardSheet,
    sheetDashboardUrl: sheetDashboardUrl
  };
}

/**
 * Convierte "2026-11" en "Noviembre 2026"
 */
function formatearEtiquetaMes_(claveMes) {
  const nombresMeses = [
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ];
  const partes = String(claveMes || '').split('-');
  if (partes.length !== 2) return claveMes;
  const anio = partes[0];
  const idxMes = parseInt(partes[1], 10) - 1;
  const nombre = nombresMeses[idxMes] || partes[1];
  return `${nombre} ${anio}`;
}

/**
 * Crea o actualiza en la pestaña "Dashboard" de Google Sheets una tabla formateada
 * con el mismo estilo visual (tonos verde azulado #00695C / #E0F2F1) de la hoja y genera
 * también su gráfico de barras agrupadas ("Tareas y Hábitos por Mes") con el mismo diseño.
 */
function sincronizarTablaMensualEnHojaDashboard_(sheetDashboard, listaPorMes) {
  if (!sheetDashboard || !listaPorMes || listaPorMes.length === 0) return;

  const tituloSeccion = '📅 CUMPLIMIENTO POR MESES';
  const tituloGraficoMes = 'Tareas y Hábitos por Mes';
  const maxRows = Math.max(sheetDashboard.getLastRow(), 1);
  const maxCols = Math.max(sheetDashboard.getLastColumn(), 1);
  const valores = sheetDashboard.getRange(1, 1, maxRows, maxCols).getValues();

  // 1. Detectar dónde acaban las celdas previas y dónde acaban los gráficos flotantes (ignorando el nuestro propio mensual)
  let ultimaFilaCeldasNoTabla = 1;
  let filaTablaPrevia = -1;
  let colTablaPrevia = -1;

  for (let r = 0; r < valores.length; r++) {
    for (let c = 0; c < valores[r].length; c++) {
      const txt = String(valores[r][c] || '').trim();
      if (txt === tituloSeccion) {
        filaTablaPrevia = r + 1;
        colTablaPrevia = c + 1;
      } else if (txt !== '' && (filaTablaPrevia === -1 || r + 1 < filaTablaPrevia)) {
        if (r + 1 > ultimaFilaCeldasNoTabla) ultimaFilaCeldasNoTabla = r + 1;
      }
    }
  }

  let maxFilaOcupadaPorGraficos = 0;
  try {
    const charts = sheetDashboard.getCharts();
    charts.forEach(ch => {
      const opts = ch.getOptions();
      const chTitle = opts ? String(opts.get('title') || '') : '';
      if (chTitle === tituloGraficoMes) return; // Ignorar nuestro propio gráfico mensual al calcular el tope superior

      const info = ch.getContainerInfo();
      if (info) {
        const anchorRow = info.getAnchorRow() || 1;
        const heightPx = (opts && opts.get('height')) ? Number(opts.get('height')) : 340;
        const filasAproxGrafico = Math.ceil(heightPx / 21) + 2;
        const finGrafico = anchorRow + filasAproxGrafico;
        if (finGrafico > maxFilaOcupadaPorGraficos) {
          maxFilaOcupadaPorGraficos = finGrafico;
        }
      }
    });
  } catch (e) {}

  const filaMinimaSegura = Math.max(ultimaFilaCeldasNoTabla + 3, maxFilaOcupadaPorGraficos + 2);

  let filaInicio = filaTablaPrevia;
  let colInicio = colTablaPrevia !== -1 ? colTablaPrevia : 2; // Empezar en Columna B (2) alineado con el gráfico de arriba

  if (filaTablaPrevia !== -1 && (filaTablaPrevia < filaMinimaSegura || colTablaPrevia !== 2)) {
    const rangoAntiguo = sheetDashboard.getRange(filaTablaPrevia, colTablaPrevia, 27, 5);
    rangoAntiguo.breakApart();
    rangoAntiguo.clearContent();
    rangoAntiguo.clearFormat();
    filaInicio = filaMinimaSegura;
    colInicio = 2;
  } else if (filaInicio === -1) {
    filaInicio = filaMinimaSegura;
    colInicio = 2;
  }

  // Limpiar área previa de la tabla mensual (hasta 24 meses)
  const rangoLimpieza = sheetDashboard.getRange(filaInicio, colInicio, 27, 5);
  rangoLimpieza.breakApart();
  rangoLimpieza.clearContent();
  rangoLimpieza.clearFormat();

  // Paleta coherente con el resto del Dashboard de la hoja (Teal #00695C y menta suave #E0F2F1)
  const colorHeaderBg = '#00695C';
  const colorSubHeaderBg = '#E0F2F1';
  const colorTealText = '#00695C';
  const colorBorder = '#B2DFDB';

  // Título del bloque
  const rangoTitulo = sheetDashboard.getRange(filaInicio, colInicio, 1, 5);
  rangoTitulo.merge();
  rangoTitulo.setValue(tituloSeccion)
    .setBackground(colorHeaderBg)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('left');

  // Cabeceras de columnas
  const cabeceras = [['Mes', 'Completadas', 'Pendientes', 'Total', '% Cumplimiento']];
  const rangoCabeceras = sheetDashboard.getRange(filaInicio + 1, colInicio, 1, 5);
  rangoCabeceras.setValues(cabeceras)
    .setBackground(colorSubHeaderBg)
    .setFontColor(colorTealText)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Filas por cada mes (escribimos el mes como texto explícito con apóstrofo inicial para que Sheets respete la mayúscula "Noviembre 2026")
  const filasDatos = listaPorMes.slice(0, 24).map(m => [
    m.etiqueta,
    m.completados,
    m.pendientes,
    m.total,
    m.total > 0 ? m.completados / m.total : 0
  ]);

  if (filasDatos.length > 0) {
    const rangoDatos = sheetDashboard.getRange(filaInicio + 2, colInicio, filasDatos.length, 5);
    sheetDashboard.getRange(filaInicio + 2, colInicio, filasDatos.length, 1).setNumberFormat('@');
    sheetDashboard.getRange(filaInicio + 2, colInicio + 1, filasDatos.length, 3).setNumberFormat('0');
    rangoDatos.setValues(filasDatos);

    sheetDashboard.getRange(filaInicio + 2, colInicio, filasDatos.length, 1)
      .setFontWeight('bold')
      .setFontColor(colorTealText)
      .setBackground('#F4FBF9')
      .setHorizontalAlignment('left');

    sheetDashboard.getRange(filaInicio + 2, colInicio + 1, filasDatos.length, 3)
      .setNumberFormat('0')
      .setHorizontalAlignment('center')
      .setBackground('#FFFFFF');

    sheetDashboard.getRange(filaInicio + 2, colInicio + 4, filasDatos.length, 1)
      .setNumberFormat('0.0%')
      .setHorizontalAlignment('center')
      .setFontWeight('bold')
      .setFontColor(colorTealText)
      .setBackground(colorSubHeaderBg);
  }

  const totalFilasTabla = 2 + filasDatos.length;
  sheetDashboard.getRange(filaInicio, colInicio, totalFilasTabla, 5)
    .setBorder(true, true, true, true, true, true, colorBorder, SpreadsheetApp.BorderStyle.SOLID);

  // 3. Crear o actualizar el Gráfico de Columnas Agrupadas ("Tareas y Hábitos por Mes") con el mismo estilo visual (#00695C y #DE7A61)
  try {
    const rangoGrafico = sheetDashboard.getRange(filaInicio + 1, colInicio, filasDatos.length + 1, 3); // Mes | Completadas | Pendientes
    const charts = sheetDashboard.getCharts();
    let chartExistente = null;

    for (let i = 0; i < charts.length; i++) {
      const opts = charts[i].getOptions();
      if (opts && String(opts.get('title') || '') === tituloGraficoMes) {
        chartExistente = charts[i];
        break;
      }
    }

    const colGrafico = colInicio + 6; // A la derecha de la tabla mensual (Columna H)
    if (chartExistente) {
      const modChart = chartExistente.modify()
        .clearRanges()
        .addRange(rangoGrafico)
        .setChartType(Charts.ChartType.COLUMN)
        .setPosition(filaInicio, colGrafico, 0, 0)
        .setOption('title', tituloGraficoMes)
        .setOption('titleTextStyle', { color: '#00695C', fontSize: 12, bold: true })
        .setOption('colors', ['#00695C', '#DE7A61'])
        .setOption('legend', { position: 'bottom' })
        .setOption('hAxis', { title: 'Mes' })
        .setOption('vAxis', { title: 'Cantidad', minValue: 0 })
        .setOption('width', 450)
        .setOption('height', 300)
        .build();
      sheetDashboard.updateChart(modChart);
    } else {
      const nuevoChart = sheetDashboard.newChart()
        .addRange(rangoGrafico)
        .setChartType(Charts.ChartType.COLUMN)
        .setNumHeaders(1)
        .setPosition(filaInicio, colGrafico, 0, 0)
        .setOption('title', tituloGraficoMes)
        .setOption('titleTextStyle', { color: '#00695C', fontSize: 12, bold: true })
        .setOption('colors', ['#00695C', '#DE7A61'])
        .setOption('legend', { position: 'bottom' })
        .setOption('hAxis', { title: 'Mes' })
        .setOption('vAxis', { title: 'Cantidad', minValue: 0 })
        .setOption('width', 450)
        .setOption('height', 300)
        .build();
      sheetDashboard.insertChart(nuevoChart);
    }
  } catch (errChart) {
    Logger.log('Aviso creando gráfico mensual en Dashboard: ' + errChart.message);
  }
}

/**
 * Obtiene una hoja del Spreadsheet ignorando tildes, mayúsculas, espacios o guiones bajos.
 * Así funciona tanto con "Plantillas_Habitos" como con "Plantillas_Hábitos" o "Plantillas Habitos".
 */
function obtenerHojaFlexible_(ss, nombreBuscado) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const exacta = ss.getSheetByName(nombreBuscado);
  if (exacta) return exacta;

  const clave = normalizarTexto_(nombreBuscado).replace(/_/g, '');
  const hojas = ss.getSheets();
  for (let i = 0; i < hojas.length; i++) {
    const nombreHojaNorm = normalizarTexto_(hojas[i].getName()).replace(/_/g, '');
    if (nombreHojaNorm === clave || nombreHojaNorm.includes(clave) || clave.includes(nombreHojaNorm)) {
      return hojas[i];
    }
  }
  return null;
}

/**
 * Inserta una fila en la primera posición donde la columna clave (1-indexed) esté vacía,
 * evitando que appendRow() escriba en la fila 500 o 1000 cuando hay validaciones/checkboxes vacíos.
 */
function insertarEnPrimeraFilaLibre_(sheet, valoresFila, colClaveIndex) {
  const colCheck = colClaveIndex || 2; // Por defecto columna B o C
  const data = sheet.getDataRange().getValues();
  let ultimaOcupada = 1; // Cabecera

  for (let i = 1; i < data.length; i++) {
    const val = String(data[i][colCheck - 1] || '').trim();
    if (val !== '') {
      ultimaOcupada = i + 1;
    }
  }

  const filaDestino = ultimaOcupada + 1;
  sheet.getRange(filaDestino, 1, 1, valoresFila.length).setValues([valoresFila]);
  if (valoresFila.length === 6 && (valoresFila[1] === 'Hábito' || valoresFila[1] === 'Tarea')) {
    aplicarValidacionesRegistro_(sheet, filaDestino, 1);
  }
  return filaDestino;
}

function obtenerFechaHoyStr_(dateObj) {
  const d = dateObj || new Date();
  return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function formatearCeldaFecha_(valor) {
  if (!valor) return '';
  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor)) {
    return Utilities.formatDate(valor, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  }
  const texto = String(valor).trim();
  const matchDMY = texto.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (matchDMY) {
    return `${matchDMY[3]}-${matchDMY[2].padStart(2, '0')}-${matchDMY[1].padStart(2, '0')}`;
  }
  return texto.substring(0, 10);
}

function normalizarTexto_(str) {
  return String(str || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\w\s]/gi, '')
    .trim();
}

/**
 * Detecta y limpia el prefijo/alias de asistente de voz configurado en PropertiesService ('VOICE_ASSISTANT_PREFIX'),
 * o cualquier etiqueta inicial entre corchetes como "[Alias] ...".
 */
function limpiarPrefijoAsistente_(texto) {
  const raw = String(texto || '').trim();
  if (!raw) return { texto: '', teniaPrefijo: false };

  const alias = (PropertiesService.getScriptProperties().getProperty('VOICE_ASSISTANT_PREFIX') || '').trim();

  // 1. Si empieza por etiqueta explícita entre corchetes "[Algo] ...", paréntesis "(Algo) ..." o asteriscos "**[Algo]** ..."
  const matchCorchetes = raw.match(/^[\*\s]*[\[\(]([^\]\)]+)[\]\)][\*\s]*[,:\-]?\s*(.+)$/);
  if (matchCorchetes && matchCorchetes[2]) {
    const tagInterior = matchCorchetes[1].trim().toLowerCase();
    if (tagInterior !== 'habito' && tagInterior !== 'hábito' && tagInterior !== 'tarea') {
      return {
        texto: matchCorchetes[2].trim(),
        teniaPrefijo: true
      };
    }
  }

  // 2. Si además hay un alias configurado en PropertiesService, detectar también "Alias: ...", "Alias, ..." o "Dile a Alias que ..."
  if (alias) {
    const aliasEsc = alias.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regexPrefijo = new RegExp(
      `^(?:\\[${aliasEsc}\\]|(?:dile\\s+a\\s+|oye\\s+|hola\\s+)?${aliasEsc}(?:\\s+que\\s+|\\s*[,:\\-]\\s*|\\s+))(.*)$`,
      'i'
    );

    const match = raw.match(regexPrefijo);
    if (match && match[1]) {
      return {
        texto: match[1].trim(),
        teniaPrefijo: true
      };
    }
  }

  return { texto: raw, teniaPrefijo: false };
}

function obtenerNombresRegistradosHoy_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.REGISTRO);
    if (!sheet) return [];
    const hoyStr = obtenerFechaHoyStr_();
    const data = sheet.getDataRange().getValues();
    const nombres = [];

    for (let i = 1; i < data.length; i++) {
      if (formatearCeldaFecha_(data[i][0]) === hoyStr && data[i][2]) {
        nombres.push(`${data[i][2]} (${data[i][4]})`);
      }
    }
    return nombres;
  } catch (e) {
    return [];
  }
}

function respuestaJson_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ============================================================================
 * 7. INTERFAZ WEB CON CRUD COMPLETO + SELECTOR DE TEMA (DARK / LIGHT / SYSTEM) + MODO RELOJ
 * ============================================================================
 */
function obtenerHtmlApp_() {
  return `<!DOCTYPE html>
<html lang="es" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
  <meta name="theme-color" content="#0B0F19" id="metaThemeColor">
  <title>⚡ Mi Gestor</title>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&display=swap" rel="stylesheet">
  <style>
    :root, [data-theme="dark"] {
      --bg-grad: radial-gradient(circle at top right, #1E1B4B, #0B0F19 60%);
      --card: rgba(22, 30, 49, 0.88);
      --card-border: rgba(255, 255, 255, 0.09);
      --tabs-bg: rgba(15, 23, 42, 0.75);
      --primary: #6366F1;
      --primary-glow: rgba(99, 102, 241, 0.35);
      --accent: #10B981;
      --danger: #EF4444;
      --text: #F8FAFC;
      --muted: #94A3B8;
      --title-grad: linear-gradient(90deg, #FFF, #A5B4FC);
      --ring-track: rgba(255, 255, 255, 0.08);
      --badge-bg: rgba(255, 255, 255, 0.07);
      --act-btn-bg: rgba(255, 255, 255, 0.05);
      --act-btn-hover: rgba(255, 255, 255, 0.14);
      --modal-bg: #131B2E;
      --input-bg: #0B0F19;
      --toast-bg: rgba(15, 23, 42, 0.96);
      --shadow-hero: 0 12px 32px rgba(0, 0, 0, 0.35);
      --cat-pct: #A5B4FC;
      --sync-pill-bg: rgba(99, 102, 241, 0.16);
      --sync-pill-text: #C7D2FE;
    }

    [data-theme="light"] {
      --bg-grad: radial-gradient(circle at top right, #E0E7FF, #F8FAFC 65%);
      --card: rgba(255, 255, 255, 0.92);
      --card-border: rgba(15, 23, 42, 0.09);
      --tabs-bg: rgba(226, 232, 240, 0.85);
      --primary: #4F46E5;
      --primary-glow: rgba(79, 70, 229, 0.25);
      --accent: #059669;
      --danger: #DC2626;
      --text: #0F172A;
      --muted: #64748B;
      --title-grad: linear-gradient(90deg, #1E1B4B, #4F46E5);
      --ring-track: rgba(15, 23, 42, 0.08);
      --badge-bg: rgba(15, 23, 42, 0.06);
      --act-btn-bg: rgba(15, 23, 42, 0.05);
      --act-btn-hover: rgba(15, 23, 42, 0.12);
      --modal-bg: #FFFFFF;
      --input-bg: #F1F5F9;
      --toast-bg: rgba(15, 23, 42, 0.92);
      --shadow-hero: 0 10px 28px rgba(15, 23, 42, 0.08);
      --cat-pct: #4F46E5;
      --sync-pill-bg: rgba(79, 70, 229, 0.12);
      --sync-pill-text: #4338CA;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }
    body {
      font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
      background: var(--bg-grad);
      color: var(--text);
      min-height: 100vh;
      padding: 16px;
      padding-bottom: 100px;
      transition: background 0.25s ease, color 0.25s ease;
    }
    .container { max-width: 500px; margin: 0 auto; }

    /* HEADER */
    .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; }
    .title-group h1 {
      font-size: 1.35rem; font-weight: 800;
      background: var(--title-grad);
      -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    .title-group p { font-size: 0.76rem; color: var(--muted); margin-top: 2px; }
    .top-actions { display: flex; gap: 6px; align-items: center; }
    .icon-btn {
      background: var(--card); border: 1px solid var(--card-border); color: var(--text);
      width: 38px; height: 38px; border-radius: 12px; display: flex; align-items: center;
      justify-content: center; font-size: 1rem; cursor: pointer;
      transition: transform 0.15s ease, background 0.2s;
    }
    .icon-btn:active { transform: scale(0.92); }
    .theme-btn {
      width: auto; padding: 0 10px; gap: 5px; font-size: 0.75rem; font-weight: 700;
    }

    /* TABS NAVIGATION */
    .nav-tabs {
      display: flex; gap: 8px; background: var(--tabs-bg);
      padding: 5px; border-radius: 16px; border: 1px solid var(--card-border);
      margin-bottom: 16px;
    }
    .tab-btn {
      flex: 1; padding: 10px; border: none; border-radius: 12px;
      background: transparent; color: var(--muted); font-family: inherit;
      font-size: 0.84rem; font-weight: 700; cursor: pointer; transition: all 0.2s;
    }
    .tab-btn.active {
      background: var(--primary); color: white;
      box-shadow: 0 4px 14px var(--primary-glow);
    }

    /* HERO CARD */
    .hero-card {
      background: var(--card); backdrop-filter: blur(16px); border: 1px solid var(--card-border);
      border-radius: 24px; padding: 18px; display: flex; align-items: center; gap: 18px;
      margin-bottom: 16px; box-shadow: var(--shadow-hero);
    }
    .ring-wrap { position: relative; width: 92px; height: 92px; flex-shrink: 0; }
    .ring-svg { transform: rotate(-90deg); width: 92px; height: 92px; }
    .ring-bg { fill: none; stroke: var(--ring-track); stroke-width: 9; }
    .ring-progress {
      fill: none; stroke: url(#gradRing); stroke-width: 9; stroke-linecap: round;
      stroke-dasharray: 238.76; stroke-dashoffset: 238.76;
      transition: stroke-dashoffset 0.7s cubic-bezier(0.22, 1, 0.36, 1);
    }
    .ring-label { position: absolute; inset: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; }
    .ring-pct { font-size: 1.3rem; font-weight: 800; }
    .ring-sub { font-size: 0.68rem; color: var(--muted); }
    .hero-info { flex: 1; }
    .hero-info h2 { font-size: 1.02rem; font-weight: 700; margin-bottom: 4px; }
    .hero-info p { font-size: 0.78rem; color: var(--muted); margin-bottom: 10px; }
    .sync-pill {
      display: inline-flex; align-items: center; gap: 6px; font-size: 0.74rem; font-weight: 600;
      padding: 6px 12px; border-radius: 99px; background: var(--sync-pill-bg);
      color: var(--sync-pill-text); border: 1px solid rgba(99, 102, 241, 0.35); cursor: pointer;
    }

    /* CATEGORIES BAR */
    .cats-scroll { display: flex; gap: 8px; overflow-x: auto; padding-bottom: 6px; margin-bottom: 14px; scrollbar-width: none; }
    .cats-scroll::-webkit-scrollbar { display: none; }
    .cat-chip {
      background: var(--card); border: 1px solid var(--card-border); padding: 7px 12px;
      border-radius: 12px; font-size: 0.74rem; white-space: nowrap; display: flex; align-items: center; gap: 6px;
    }
    .cat-pct { font-weight: 800; color: var(--cat-pct); }

    /* LISTS */
    .section-title {
      font-size: 0.8rem; text-transform: uppercase; letter-spacing: 0.06em; color: var(--muted);
      font-weight: 700; margin: 14px 0 10px; display: flex; justify-content: space-between; align-items: center;
    }
    .task-list { display: flex; flex-direction: column; gap: 10px; }
    .task-item {
      background: var(--card); border: 1px solid var(--card-border); border-radius: 18px;
      padding: 13px 15px; display: flex; align-items: center; gap: 12px;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
    }
    .task-item.done { opacity: 0.65; background: rgba(16, 185, 129, 0.08); border-color: rgba(16, 185, 129, 0.25); }
    .task-item.paused { opacity: 0.5; }
    .check-circle {
      width: 28px; height: 28px; border-radius: 50%; border: 2px solid var(--primary);
      display: flex; align-items: center; justify-content: center; cursor: pointer; flex-shrink: 0; font-size: 0.9rem;
    }
    .task-item.done .check-circle { background: var(--accent); border-color: var(--accent); color: #FFFFFF; font-weight: 900; }
    .task-content { flex: 1; min-width: 0; cursor: pointer; }
    .task-name { font-size: 0.94rem; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .task-item.done .task-name { text-decoration: line-through; color: var(--muted); }
    .task-meta { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 4px; font-size: 0.69rem; color: var(--muted); }
    .badge { padding: 2px 8px; border-radius: 6px; background: var(--badge-bg); font-weight: 600; }
    .item-actions { display: flex; gap: 4px; align-items: center; }
    .act-btn {
      background: var(--act-btn-bg); border: none; color: var(--muted);
      width: 32px; height: 32px; border-radius: 10px; font-size: 0.88rem;
      cursor: pointer; display: flex; align-items: center; justify-content: center;
    }
    .act-btn:hover { background: var(--act-btn-hover); color: var(--text); }

    /* TOGGLE SWITCH PARA PLANTILLAS */
    .switch-btn {
      padding: 4px 10px; border-radius: 99px; font-size: 0.7rem; font-weight: 700;
      border: none; cursor: pointer;
    }
    .switch-btn.on { background: rgba(16, 185, 129, 0.18); color: var(--accent); border: 1px solid rgba(16, 185, 129, 0.4); }
    .switch-btn.off { background: rgba(148, 163, 184, 0.15); color: var(--muted); border: 1px solid rgba(148, 163, 184, 0.3); }

    /* FLOATING ACTION BUTTON (+) */
    .fab {
      position: fixed; bottom: 24px; right: 24px;
      background: linear-gradient(135deg, #6366F1, #8B5CF6);
      color: white; border: none; border-radius: 99px;
      padding: 14px 22px; font-family: inherit; font-weight: 800; font-size: 0.92rem;
      box-shadow: 0 10px 25px rgba(99, 102, 241, 0.5);
      display: flex; align-items: center; gap: 8px; cursor: pointer; z-index: 90;
    }
    .fab:active { transform: scale(0.95); }

    /* MODAL CRUD */
    .modal-backdrop {
      position: fixed; inset: 0; background: rgba(5, 8, 17, 0.75);
      backdrop-filter: blur(6px); display: none; align-items: center;
      justify-content: center; padding: 16px; z-index: 200;
    }
    .modal-backdrop.open { display: flex; }
    .modal-card {
      background: var(--modal-bg); border: 1px solid var(--card-border);
      border-radius: 24px; width: 100%; max-width: 420px; padding: 22px;
      box-shadow: 0 20px 50px rgba(0, 0, 0, 0.45);
    }
    .modal-title { font-size: 1.1rem; font-weight: 800; margin-bottom: 16px; color: var(--text); }
    .form-group { margin-bottom: 14px; }
    .form-group label { display: block; font-size: 0.76rem; color: var(--muted); font-weight: 700; margin-bottom: 6px; }
    .form-control {
      width: 100%; padding: 11px 14px; border-radius: 12px;
      background: var(--input-bg); border: 1px solid var(--card-border);
      color: var(--text); font-family: inherit; font-size: 0.9rem; outline: none;
    }
    .form-control:focus { border-color: var(--primary); }
    .modal-actions { display: flex; justify-content: flex-end; gap: 10px; margin-top: 18px; }
    .btn {
      padding: 10px 18px; border-radius: 12px; font-family: inherit;
      font-weight: 700; font-size: 0.86rem; border: none; cursor: pointer;
    }
    .btn-cancel { background: var(--badge-bg); color: var(--muted); }
    .btn-save { background: var(--primary); color: white; }

    /* TOAST */
    .toast {
      position: fixed; bottom: 85px; left: 50%; transform: translateX(-50%) translateY(120px);
      background: var(--toast-bg); border: 1px solid rgba(99, 102, 241, 0.45); color: white;
      padding: 11px 18px; border-radius: 14px; font-size: 0.84rem; font-weight: 600;
      box-shadow: 0 10px 30px rgba(0,0,0,0.4); transition: transform 0.28s ease;
      z-index: 300; max-width: 90vw; text-align: center; pointer-events: none;
    }
    .toast.show { transform: translateX(-50%) translateY(0); }

    /* DASHBOARD TAB STYLES */
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 10px;
      margin-bottom: 16px;
    }
    .kpi-card {
      background: var(--card);
      border: 1px solid var(--card-border);
      border-radius: 18px;
      padding: 14px 10px;
      text-align: center;
      box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
    }
    .kpi-val {
      font-size: 1.35rem;
      font-weight: 800;
      color: var(--cat-pct);
    }
    .kpi-lbl {
      font-size: 0.68rem;
      color: var(--muted);
      font-weight: 700;
      margin-top: 4px;
      text-transform: uppercase;
    }
    .kpi-sub {
      font-size: 0.7rem;
      color: var(--muted);
      margin-top: 2px;
    }
    .cat-bar-card {
      background: var(--card);
      border: 1px solid var(--card-border);
      border-radius: 18px;
      padding: 16px;
      margin-bottom: 16px;
    }
    .cat-bar-row {
      margin-bottom: 12px;
    }
    .cat-bar-row:last-child { margin-bottom: 0; }
    .cat-bar-header {
      display: flex;
      justify-content: space-between;
      font-size: 0.8rem;
      font-weight: 700;
      margin-bottom: 5px;
    }
    .cat-bar-track {
      width: 100%;
      height: 8px;
      background: var(--ring-track);
      border-radius: 99px;
      overflow: hidden;
    }
    .cat-bar-fill {
      height: 100%;
      background: linear-gradient(90deg, var(--primary), var(--accent));
      border-radius: 99px;
      transition: width 0.5s ease;
    }
    /* GRÁFICO DE COLUMNAS AGRUPADAS (ESTILO SHEET TEAL & CORAL) */
    .grouped-chart-card {
      background: var(--card);
      border: 1px solid var(--card-border);
      border-radius: 18px;
      padding: 16px;
      margin-bottom: 16px;
    }
    .grouped-chart-title {
      font-size: 0.86rem;
      font-weight: 800;
      color: var(--accent);
      margin-bottom: 14px;
    }
    .grouped-chart-bars {
      display: flex;
      align-items: flex-end;
      justify-content: space-around;
      gap: 10px;
      height: 150px;
      padding: 18px 8px 6px;
      border-bottom: 1px solid var(--card-border);
      overflow-x: auto;
    }
    .grouped-col-item {
      display: flex;
      flex-direction: column;
      align-items: center;
      min-width: 54px;
      height: 100%;
      justify-content: flex-end;
    }
    .grouped-pair {
      display: flex;
      align-items: flex-end;
      gap: 5px;
      height: 115px;
    }
    .bar-col-wrap {
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: flex-end;
      height: 100%;
    }
    .bar-col-val {
      font-size: 0.66rem;
      font-weight: 700;
      color: var(--muted);
      margin-bottom: 3px;
    }
    .bar-col {
      width: 16px;
      border-radius: 5px 5px 0 0;
      min-height: 4px;
      transition: height 0.5s ease;
    }
    .bar-col.done-col {
      background: #00695C;
    }
    .bar-col.pend-col {
      background: #DE7A61;
    }
    .grouped-col-lbl {
      font-size: 0.68rem;
      font-weight: 700;
      color: var(--text);
      margin-top: 7px;
      text-align: center;
      white-space: nowrap;
    }
    .grouped-col-pct {
      font-size: 0.64rem;
      font-weight: 800;
      color: var(--accent);
    }
    .grouped-legend {
      display: flex;
      justify-content: center;
      gap: 18px;
      margin-top: 12px;
      font-size: 0.74rem;
      font-weight: 700;
      color: var(--muted);
    }
    .legend-dot {
      display: inline-block;
      width: 10px;
      height: 10px;
      border-radius: 3px;
      margin-right: 5px;
      vertical-align: middle;
    }
    .sheet-link-btn {
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      width: 100%;
      padding: 13px 16px;
      margin-bottom: 16px;
      border-radius: 16px;
      background: var(--sync-pill-bg);
      color: var(--sync-pill-text);
      border: 1px solid rgba(99, 102, 241, 0.35);
      font-weight: 700;
      font-size: 0.85rem;
      text-decoration: none;
      transition: transform 0.15s ease;
    }
    .sheet-link-btn:active { transform: scale(0.98); }
    .sheet-table-wrap {
      background: var(--card);
      border: 1px solid var(--card-border);
      border-radius: 18px;
      padding: 12px;
      overflow-x: auto;
    }
    .sheet-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 0.76rem;
    }
    .sheet-table td {
      padding: 7px 8px;
      border-bottom: 1px solid var(--card-border);
      color: var(--text);
      white-space: nowrap;
    }
    .sheet-table tr:last-child td { border-bottom: none; }

    /* MODO RELOJ WEAR OS */
    body.watch-mode { background: #000000 !important; color: #F8FAFC !important; padding: 10px 8px 36px; }
    body.watch-mode .container { max-width: 230px; text-align: center; }
    body.watch-mode .header, body.watch-mode .nav-tabs, body.watch-mode .cats-scroll,
    body.watch-mode .hero-info p, body.watch-mode .item-actions, body.watch-mode .fab { display: none !important; }
    body.watch-mode .hero-card { flex-direction: column; background: transparent; border: none; padding: 4px; gap: 8px; margin-bottom: 8px; box-shadow: none; }
    body.watch-mode .task-item { padding: 10px 12px; border-radius: 99px; text-align: left; background: rgba(22, 30, 49, 0.9); color: #F8FAFC; }
    body.watch-mode .task-meta { display: none; }
  </style>
</head>
<body>
  <div class="container">
    <!-- CABECERA -->
    <div class="header">
      <div class="title-group">
        <h1>⚡ Mi Gestor</h1>
        <p id="dateSubtitle">Sincronizado con Google Tasks</p>
      </div>
      <div class="top-actions">
        <button class="icon-btn theme-btn" id="themeBtn" onclick="ciclarModoTema()" title="Cambiar tema: Oscuro / Claro / Sistema">
          <span id="themeIcon">🖥️</span>
          <span id="themeLabel">Auto</span>
        </button>
        <button class="icon-btn" onclick="abrirModalCrear()" title="Añadir nuevo">➕</button>
        <button class="icon-btn" onclick="document.body.classList.toggle('watch-mode')" title="Modo Reloj">⌚</button>
        <button class="icon-btn" onclick="cargarDatos()" title="Sincronizar ahora">🔄</button>
      </div>
    </div>

    <!-- PESTAÑAS DE NAVEGACIÓN (HOY vs CATÁLOGO vs DASHBOARD) -->
    <div class="nav-tabs">
      <button class="tab-btn active" id="tabHoyBtn" onclick="cambiarTab('hoy')">📅 Hoy (<span id="tabHoyCount">0</span>)</button>
      <button class="tab-btn" id="tabPlantillasBtn" onclick="cambiarTab('plantillas')">⚙️ Hábitos (<span id="tabPlantillasCount">0</span>)</button>
      <button class="tab-btn" id="tabDashboardBtn" onclick="cambiarTab('dashboard')">📊 Dashboard</button>
    </div>

    <!-- ==================== VISTA 1: HOY ==================== -->
    <div id="vistaHoy">
      <div class="hero-card">
        <div class="ring-wrap">
          <svg class="ring-svg" viewBox="0 0 96 96">
            <defs>
              <linearGradient id="gradRing" x1="0%" y1="0%" x2="100%" y2="100%">
                <stop offset="0%" stop-color="#6366F1" />
                <stop offset="100%" stop-color="#10B981" />
              </linearGradient>
            </defs>
            <circle class="ring-bg" cx="48" cy="48" r="38"></circle>
            <circle class="ring-progress" id="ringCircle" cx="48" cy="48" r="38"></circle>
          </svg>
          <div class="ring-label">
            <span class="ring-pct" id="pctText">0%</span>
            <span class="ring-sub" id="countText">0/0</span>
          </div>
        </div>

        <div class="hero-info">
          <h2 id="statusHeadline">¡Vamos a por el día!</h2>
          <p>Gestiona tus tareas aquí o desde Google Tasks.</p>
          <div class="sync-pill" onclick="volcarHabitosHoy()">
            <span>📥 Volcar hábitos de hoy</span>
          </div>
        </div>
      </div>

      <div class="cats-scroll" id="catsContainer"></div>

      <div class="section-title">
        <span>⏳ Pendientes de hoy</span>
        <span id="pendCountBadge">0</span>
      </div>
      <div class="task-list" id="pendientesList"></div>

      <div class="section-title" style="margin-top: 22px;">
        <span>✅ Completados hoy</span>
        <span id="compCountBadge">0</span>
      </div>
      <div class="task-list" id="completadosList"></div>
    </div>

    <!-- ==================== VISTA 2: CATÁLOGO DE HÁBITOS (PLANTILLAS) ==================== -->
    <div id="vistaPlantillas" style="display: none;">
      <div class="section-title">
        <span>📋 Tus Hábitos Recurrentes (Plantillas_Habitos)</span>
      </div>
      <div class="task-list" id="plantillasList"></div>
    </div>

    <!-- ==================== VISTA 3: DASHBOARD Y MÉTRICAS ==================== -->
    <div id="vistaDashboard" style="display: none;">
      <a href="#" target="_blank" rel="noopener noreferrer" class="sheet-link-btn" id="openSheetDashboardBtn">
        <span>↗️ Abrir pestaña Dashboard en Google Sheets</span>
      </a>

      <div class="section-title">
        <span>📈 Resumen de Cumplimiento</span>
      </div>
      <div class="kpi-grid">
        <div class="kpi-card">
          <div class="kpi-val" id="kpiHoyPct">0%</div>
          <div class="kpi-lbl">Hoy</div>
          <div class="kpi-sub" id="kpiHoySub">0/0</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-val" id="kpiSemanaPct">0%</div>
          <div class="kpi-lbl">Últ. 7 Días</div>
          <div class="kpi-sub" id="kpiSemanaSub">0/0</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-val" id="kpiMesPct">0%</div>
          <div class="kpi-lbl" id="kpiMesLbl">Este Mes</div>
          <div class="kpi-sub" id="kpiMesSub">0/0</div>
        </div>
        <div class="kpi-card">
          <div class="kpi-val" id="kpiTotalPct">0%</div>
          <div class="kpi-lbl">Histórico</div>
          <div class="kpi-sub" id="kpiTotalSub">0/0</div>
        </div>
      </div>

      <div class="section-title">
        <span>📅 Completados por Meses</span>
      </div>
      <div class="grouped-chart-card">
        <div class="grouped-chart-title">📊 Tareas y Hábitos por Mes</div>
        <div class="grouped-chart-bars" id="dashboardMonthsChart"></div>
        <div class="grouped-legend">
          <span><i class="legend-dot" style="background:#00695C;"></i>Completadas</span>
          <span><i class="legend-dot" style="background:#DE7A61;"></i>Pendientes</span>
        </div>
      </div>
      <div class="cat-bar-card" id="dashboardMonthsBars"></div>

      <div class="section-title">
        <span>🏷️ Cumplimiento Histórico por Categoría</span>
      </div>
      <div class="grouped-chart-card">
        <div class="grouped-chart-title">📊 Tareas y Hábitos por Categoría</div>
        <div class="grouped-chart-bars" id="dashboardCatsChart"></div>
        <div class="grouped-legend">
          <span><i class="legend-dot" style="background:#00695C;"></i>Completadas</span>
          <span><i class="legend-dot" style="background:#DE7A61;"></i>Pendientes</span>
        </div>
      </div>
      <div class="cat-bar-card" id="dashboardCatsBars"></div>
    </div>
  </div>

  <!-- BOTÓN FLOTANTE (+ NUEVO) -->
  <button class="fab" onclick="abrirModalCrear()">
    <span>＋</span> Nuevo Hábito / Tarea
  </button>

  <!-- MODAL PARA CREAR / EDITAR (CRUD) -->
  <div class="modal-backdrop" id="crudModal">
    <div class="modal-card">
      <div class="modal-title" id="modalTitle">➕ Nuevo Registro</div>

      <input type="hidden" id="formModo" value="CREAR">
      <input type="hidden" id="formIdFila" value="">

      <div class="form-group" id="groupTipo">
        <label>Tipo de registro</label>
        <select class="form-control" id="formTipo" onchange="actualizarVisibilidadFrecuencia()">
          <option value="Tarea">📌 Tarea puntual (solo para hoy)</option>
          <option value="Hábito">🔄 Hábito recurrente (guardar en plantilla)</option>
        </select>
      </div>

      <div class="form-group">
        <label>Nombre</label>
        <input type="text" class="form-control" id="formNombre" placeholder="Ej: 20 min audio italiano, Comprar pilas...">
      </div>

      <div class="form-group">
        <label>Categoría</label>
        <select class="form-control" id="formCategoria">
          <option value="Salud">Salud</option>
          <option value="Estudio">Estudio</option>
          <option value="Hogar">Hogar</option>
          <option value="Mantenimiento">Mantenimiento</option>
          <option value="Finanzas">Finanzas</option>
          <option value="Trabajo">Trabajo</option>
        </select>
      </div>

      <div class="form-group" id="groupFrecuencia" style="display: none;">
        <label>Frecuencia del Hábito</label>
        <select class="form-control" id="formFrecuencia">
          <option value="Diaria">Diaria (Todos los días)</option>
          <option value="L-V">L-V (Lunes a Viernes)</option>
          <option value="Fin de semana">Fin de semana (Sábado y Domingo)</option>
          <option value="Semanal">Semanal (Lunes)</option>
        </select>
      </div>

      <div class="form-group" id="groupMomento" style="display: none;">
        <label>Hora / Momento del día</label>
        <input type="text" class="form-control" id="formMomento" placeholder="Ej: Mañana, Noche, 08:00...">
      </div>

      <div class="modal-actions">
        <button class="btn btn-cancel" onclick="cerrarModal()">Cancelar</button>
        <button class="btn btn-save" onclick="guardarFormularioCrud()">Guardar</button>
      </div>
    </div>
  </div>

  <div class="toast" id="toastMsg"></div>

  <script>
    const CIRCUMFERENCE = 2 * Math.PI * 38;
    let tabActual = 'hoy';
    let ultimoDashboard = null;

    /* ==================== GESTIÓN DE TEMA (DARK / LIGHT / SYSTEM) ==================== */
    const THEME_MODES = ['system', 'dark', 'light'];
    let currentThemeMode = localStorage.getItem('gestor_theme_mode') || 'system';
    const systemDarkQuery = window.matchMedia('(prefers-color-scheme: dark)');

    function aplicarModoTema(modo, mostrarAviso) {
      currentThemeMode = modo;
      localStorage.setItem('gestor_theme_mode', modo);

      const effectiveTheme = modo === 'system'
        ? (systemDarkQuery.matches ? 'dark' : 'light')
        : modo;

      document.documentElement.setAttribute('data-theme', effectiveTheme);
      const metaTheme = document.getElementById('metaThemeColor');
      if (metaTheme) {
        metaTheme.setAttribute('content', effectiveTheme === 'dark' ? '#0B0F19' : '#F8FAFC');
      }

      const iconEl = document.getElementById('themeIcon');
      const labelEl = document.getElementById('themeLabel');
      if (modo === 'system') {
        iconEl.textContent = '🖥️';
        labelEl.textContent = 'Auto';
        if (mostrarAviso) mostrarToast('🖥️ Tema: Automático del sistema');
      } else if (modo === 'dark') {
        iconEl.textContent = '🌙';
        labelEl.textContent = 'Oscuro';
        if (mostrarAviso) mostrarToast('🌙 Tema: Modo Oscuro');
      } else {
        iconEl.textContent = '☀️';
        labelEl.textContent = 'Claro';
        if (mostrarAviso) mostrarToast('☀️ Tema: Modo Claro');
      }
    }

    function ciclarModoTema() {
      const nextIndex = (THEME_MODES.indexOf(currentThemeMode) + 1) % THEME_MODES.length;
      aplicarModoTema(THEME_MODES[nextIndex], true);
    }

    if (systemDarkQuery.addEventListener) {
      systemDarkQuery.addEventListener('change', () => {
        if (currentThemeMode === 'system') aplicarModoTema('system', false);
      });
    }
    aplicarModoTema(currentThemeMode, false);

    if (window.innerWidth <= 280 || window.location.search.includes('watch')) {
      document.body.classList.add('watch-mode');
    }

    function cambiarTab(tab) {
      tabActual = tab;
      document.getElementById('tabHoyBtn').classList.toggle('active', tab === 'hoy');
      document.getElementById('tabPlantillasBtn').classList.toggle('active', tab === 'plantillas');
      document.getElementById('tabDashboardBtn').classList.toggle('active', tab === 'dashboard');
      document.getElementById('vistaHoy').style.display = tab === 'hoy' ? 'block' : 'none';
      document.getElementById('vistaPlantillas').style.display = tab === 'plantillas' ? 'block' : 'none';
      document.getElementById('vistaDashboard').style.display = tab === 'dashboard' ? 'block' : 'none';
    }

    function mostrarToast(texto) {
      const t = document.getElementById('toastMsg');
      t.textContent = texto;
      t.classList.add('show');
      setTimeout(() => t.classList.remove('show'), 3200);
    }

    function actualizarUI(dashboard) {
      if (!dashboard || !dashboard.hoy) return;
      ultimoDashboard = dashboard;
      const h = dashboard.hoy;
      const plantillas = dashboard.plantillas || [];

      document.getElementById('dateSubtitle').textContent = \`📅 \${dashboard.fecha} · Sync \${dashboard.actualizadoEn}\`;
      document.getElementById('pctText').textContent = \`\${h.porcentajeCumplimiento}%\`;
      document.getElementById('countText').textContent = \`\${h.completados}/\${h.total}\`;
      document.getElementById('tabHoyCount').textContent = \`\${h.completados}/\${h.total}\`;
      document.getElementById('tabPlantillasCount').textContent = plantillas.length;

      const offset = CIRCUMFERENCE - (h.porcentajeCumplimiento / 100) * CIRCUMFERENCE;
      document.getElementById('ringCircle').style.strokeDashoffset = offset;

      if (h.total === 0) {
        document.getElementById('statusHeadline').textContent = 'Sin hábitos volcados hoy';
      } else if (h.porcentajeCumplimiento === 100) {
        document.getElementById('statusHeadline').textContent = '🔥 ¡Día 100% completado!';
      } else {
        document.getElementById('statusHeadline').textContent = \`Te quedan \${h.pendientesCount} para hoy\`;
      }

      // Categorías
      const catsCont = document.getElementById('catsContainer');
      catsCont.innerHTML = '';
      Object.keys(h.porCategoria || {}).forEach(cat => {
        const c = h.porCategoria[cat];
        const div = document.createElement('div');
        div.className = 'cat-chip';
        div.innerHTML = \`<span>\${cat}</span> <span class="cat-pct">\${c.completados}/\${c.total}</span>\`;
        catsCont.appendChild(div);
      });

      // Lista de Pendientes Hoy
      document.getElementById('pendCountBadge').textContent = h.pendientes.length;
      const pendList = document.getElementById('pendientesList');
      pendList.innerHTML = '';
      if (h.pendientes.length === 0) {
        pendList.innerHTML = '<div style="color:var(--muted);font-size:0.85rem;text-align:center;padding:14px;">🎉 ¡Todo al día!</div>';
      } else {
        h.pendientes.forEach(item => pendList.appendChild(crearFilaHoy(item, false)));
      }

      // Lista de Completados Hoy
      document.getElementById('compCountBadge').textContent = h.completadosLista.length;
      const compList = document.getElementById('completadosList');
      compList.innerHTML = '';
      h.completadosLista.forEach(item => compList.appendChild(crearFilaHoy(item, true)));

      // Lista de Catálogo de Plantillas
      const plantList = document.getElementById('plantillasList');
      plantList.innerHTML = '';
      if (plantillas.length === 0) {
        plantList.innerHTML = '<div style="color:var(--muted);font-size:0.85rem;text-align:center;padding:14px;">No hay hábitos en Plantillas_Habitos.</div>';
      } else {
        plantillas.forEach(p => plantList.appendChild(crearFilaPlantilla(p)));
      }

      // ==================== ACTUALIZAR PESTAÑA DASHBOARD ====================
      if (dashboard.sheetDashboardUrl) {
        document.getElementById('openSheetDashboardBtn').href = dashboard.sheetDashboardUrl;
      }

      document.getElementById('kpiHoyPct').textContent = \`\${h.porcentajeCumplimiento}%\`;
      document.getElementById('kpiHoySub').textContent = \`\${h.completados}/\${h.total}\`;

      if (dashboard.resumenGlobal) {
        const sem = dashboard.resumenGlobal.semana || { porcentaje: 0, completados: 0, total: 0 };
        const mesAct = dashboard.resumenGlobal.mesActual || { etiqueta: 'Este Mes', porcentaje: 0, completados: 0, total: 0 };
        const hist = dashboard.resumenGlobal.historico || { porcentaje: 0, completados: 0, total: 0 };

        document.getElementById('kpiSemanaPct').textContent = \`\${sem.porcentaje}%\`;
        document.getElementById('kpiSemanaSub').textContent = \`\${sem.completados}/\${sem.total}\`;
        document.getElementById('kpiMesPct').textContent = \`\${mesAct.porcentaje}%\`;
        document.getElementById('kpiMesLbl').textContent = mesAct.etiqueta || 'Este Mes';
        document.getElementById('kpiMesSub').textContent = \`\${mesAct.completados}/\${mesAct.total}\`;
        document.getElementById('kpiTotalPct').textContent = \`\${hist.porcentaje}%\`;
        document.getElementById('kpiTotalSub').textContent = \`\${hist.completados}/\${hist.total}\`;

        // Barras y Gráfico de Completados por Meses
        const monthsCont = document.getElementById('dashboardMonthsBars');
        const monthsChart = document.getElementById('dashboardMonthsChart');
        monthsCont.innerHTML = '';
        monthsChart.innerHTML = '';
        const listaMeses = dashboard.resumenGlobal.porMes || [];
        if (listaMeses.length === 0) {
          monthsCont.innerHTML = '<div style="color:var(--muted);font-size:0.84rem;text-align:center;">Sin registros mensuales todavía.</div>';
          monthsChart.innerHTML = '<div style="color:var(--muted);font-size:0.8rem;">Sin datos</div>';
        } else {
          let maxValMes = 1;
          listaMeses.forEach(m => {
            const pend = Math.max(0, m.total - m.completados);
            if (m.completados > maxValMes) maxValMes = m.completados;
            if (pend > maxValMes) maxValMes = pend;
          });

          listaMeses.slice().reverse().forEach(itemM => {
            const pendM = Math.max(0, itemM.total - itemM.completados);
            const hDone = Math.max(4, Math.round((itemM.completados / maxValMes) * 90));
            const hPend = Math.max(4, Math.round((pendM / maxValMes) * 90));
            const etqCorta = String(itemM.etiqueta || '').replace(' 20', ' ').substring(0, 10);

            const colEl = document.createElement('div');
            colEl.className = 'grouped-col-item';
            colEl.innerHTML = \`
              <div class="grouped-pair">
                <div class="bar-col-wrap">
                  <span class="bar-col-val">\${itemM.completados}</span>
                  <div class="bar-col done-col" style="height:\${hDone}px;" title="Completadas: \${itemM.completados}"></div>
                </div>
                <div class="bar-col-wrap">
                  <span class="bar-col-val">\${pendM}</span>
                  <div class="bar-col pend-col" style="height:\${hPend}px;" title="Pendientes: \${pendM}"></div>
                </div>
              </div>
              <div class="grouped-col-lbl">\${etqCorta}</div>
              <div class="grouped-col-pct">\${itemM.porcentaje}%</div>
            \`;
            monthsChart.appendChild(colEl);
          });

          listaMeses.forEach(itemM => {
            const rowDiv = document.createElement('div');
            rowDiv.className = 'cat-bar-row';
            rowDiv.innerHTML = \`
              <div class="cat-bar-header">
                <span>🗓️ \${itemM.etiqueta}</span>
                <span style="color:var(--muted);">\${itemM.completados}/\${itemM.total} completados (\${itemM.porcentaje}%)</span>
              </div>
              <div class="cat-bar-track">
                <div class="cat-bar-fill" style="width: \${itemM.porcentaje}%"></div>
              </div>
            \`;
            monthsCont.appendChild(rowDiv);
          });
        }

        // Barras y Gráfico de Cumplimiento Histórico por Categoría
        const barsCont = document.getElementById('dashboardCatsBars');
        const catsChart = document.getElementById('dashboardCatsChart');
        barsCont.innerHTML = '';
        catsChart.innerHTML = '';
        const catsHist = dashboard.resumenGlobal.porCategoria || {};
        const catsKeys = Object.keys(catsHist);

        let maxValCat = 1;
        catsKeys.forEach(cat => {
          const c = catsHist[cat];
          const pend = Math.max(0, c.total - c.completados);
          if (c.completados > maxValCat) maxValCat = c.completados;
          if (pend > maxValCat) maxValCat = pend;
        });

        catsKeys.forEach(cat => {
          const itemC = catsHist[cat];
          const pendC = Math.max(0, itemC.total - itemC.completados);
          const hDone = Math.max(4, Math.round((itemC.completados / maxValCat) * 90));
          const hPend = Math.max(4, Math.round((pendC / maxValCat) * 90));
          const catCorta = cat.length > 9 ? cat.substring(0, 8) + '…' : cat;

          const colEl = document.createElement('div');
          colEl.className = 'grouped-col-item';
          colEl.innerHTML = \`
            <div class="grouped-pair">
              <div class="bar-col-wrap">
                <span class="bar-col-val">\${itemC.completados}</span>
                <div class="bar-col done-col" style="height:\${hDone}px;" title="Completadas: \${itemC.completados}"></div>
              </div>
              <div class="bar-col-wrap">
                <span class="bar-col-val">\${pendC}</span>
                <div class="bar-col pend-col" style="height:\${hPend}px;" title="Pendientes: \${pendC}"></div>
              </div>
            </div>
            <div class="grouped-col-lbl">\${catCorta}</div>
            <div class="grouped-col-pct">\${itemC.porcentaje}%</div>
          \`;
          catsChart.appendChild(colEl);

          const rowDiv = document.createElement('div');
          rowDiv.className = 'cat-bar-row';
          rowDiv.innerHTML = \`
            <div class="cat-bar-header">
              <span>\${cat}</span>
              <span style="color:var(--muted);">\${itemC.completados}/\${itemC.total} (\${itemC.porcentaje}%)</span>
            </div>
            <div class="cat-bar-track">
              <div class="cat-bar-fill" style="width: \${itemC.porcentaje}%"></div>
            </div>
          \`;
          barsCont.appendChild(rowDiv);
        });
      }
    }

    function crearFilaHoy(item, estaCompletado) {
      const row = document.createElement('div');
      row.className = 'task-item' + (estaCompletado ? ' done' : '');

      const check = document.createElement('div');
      check.className = 'check-circle';
      check.innerHTML = estaCompletado ? '✓' : '';
      check.onclick = () => cambiarEstadoItem(item, !estaCompletado);

      const content = document.createElement('div');
      content.className = 'task-content';
      content.onclick = () => cambiarEstadoItem(item, !estaCompletado);
      content.innerHTML = \`
        <div class="task-name">\${item.nombre}</div>
        <div class="task-meta">
          <span class="badge">\${item.tipo}</span>
          <span class="badge">\${item.categoria}</span>
        </div>
      \`;

      const actions = document.createElement('div');
      actions.className = 'item-actions';

      const editBtn = document.createElement('button');
      editBtn.className = 'act-btn';
      editBtn.innerHTML = '✏️';
      editBtn.title = 'Editar';
      editBtn.onclick = (e) => { e.stopPropagation(); abrirModalEditarHoy(item); };

      const delBtn = document.createElement('button');
      delBtn.className = 'act-btn';
      delBtn.innerHTML = '🗑️';
      delBtn.title = 'Borrar de hoy';
      delBtn.onclick = (e) => { e.stopPropagation(); borrarItemHoy(item); };

      actions.appendChild(editBtn);
      actions.appendChild(delBtn);

      row.appendChild(check);
      row.appendChild(content);
      row.appendChild(actions);
      return row;
    }

    function crearFilaPlantilla(p) {
      const row = document.createElement('div');
      row.className = 'task-item' + (!p.activo ? ' paused' : '');

      const content = document.createElement('div');
      content.className = 'task-content';
      content.onclick = () => abrirModalEditarPlantilla(p);
      content.innerHTML = \`
        <div class="task-name">\${p.nombre}</div>
        <div class="task-meta">
          <span class="badge">🔄 \${p.frecuencia}</span>
          <span class="badge">\${p.categoria}</span>
          <span class="badge">🕒 \${p.horaMomento}</span>
        </div>
      \`;

      const actions = document.createElement('div');
      actions.className = 'item-actions';

      const toggleBtn = document.createElement('button');
      toggleBtn.className = 'switch-btn ' + (p.activo ? 'on' : 'off');
      toggleBtn.textContent = p.activo ? 'Activo' : 'Pausado';
      toggleBtn.onclick = (e) => {
        e.stopPropagation();
        toggleActivoPlantilla(p);
      };

      const editBtn = document.createElement('button');
      editBtn.className = 'act-btn';
      editBtn.innerHTML = '✏️';
      editBtn.onclick = (e) => { e.stopPropagation(); abrirModalEditarPlantilla(p); };

      const delBtn = document.createElement('button');
      delBtn.className = 'act-btn';
      delBtn.innerHTML = '🗑️';
      delBtn.onclick = (e) => { e.stopPropagation(); borrarPlantilla(p); };

      actions.appendChild(toggleBtn);
      actions.appendChild(editBtn);
      actions.appendChild(delBtn);

      row.appendChild(content);
      row.appendChild(actions);
      return row;
    }

    /* ==================== FUNCIONES DEL MODAL CRUD ==================== */
    function actualizarVisibilidadFrecuencia() {
      const tipo = document.getElementById('formTipo').value;
      const modo = document.getElementById('formModo').value;
      const mostrarHabito = (tipo === 'Hábito' || modo === 'EDITAR_PLANTILLA');
      document.getElementById('groupFrecuencia').style.display = mostrarHabito ? 'block' : 'none';
      document.getElementById('groupMomento').style.display = mostrarHabito ? 'block' : 'none';
    }

    function abrirModalCrear() {
      document.getElementById('formModo').value = 'CREAR';
      document.getElementById('formIdFila').value = '';
      document.getElementById('modalTitle').textContent = tabActual === 'plantillas' ? '➕ Nuevo Hábito en Catálogo' : '➕ Añadir Tarea o Hábito';
      document.getElementById('groupTipo').style.display = 'block';
      document.getElementById('formTipo').value = tabActual === 'plantillas' ? 'Hábito' : 'Tarea';
      document.getElementById('formNombre').value = '';
      document.getElementById('formCategoria').value = 'Hogar';
      document.getElementById('formFrecuencia').value = 'Diaria';
      document.getElementById('formMomento').value = 'Cualquier momento';
      actualizarVisibilidadFrecuencia();
      document.getElementById('crudModal').classList.add('open');
      setTimeout(() => document.getElementById('formNombre').focus(), 100);
    }

    function abrirModalEditarHoy(item) {
      document.getElementById('formModo').value = 'EDITAR_HOY';
      document.getElementById('formIdFila').value = item.idFila;
      document.getElementById('modalTitle').textContent = '✏️ Editar registro de hoy';
      document.getElementById('groupTipo').style.display = 'block';
      document.getElementById('formTipo').value = item.tipo || 'Tarea';
      document.getElementById('formNombre').value = item.nombre || '';
      document.getElementById('formCategoria').value = item.categoria || 'General';
      document.getElementById('groupFrecuencia').style.display = 'none';
      document.getElementById('groupMomento').style.display = 'none';
      document.getElementById('crudModal').classList.add('open');
    }

    function abrirModalEditarPlantilla(p) {
      document.getElementById('formModo').value = 'EDITAR_PLANTILLA';
      document.getElementById('formIdFila').value = p.idFila;
      document.getElementById('modalTitle').textContent = '✏️ Editar Hábito de Plantilla';
      document.getElementById('groupTipo').style.display = 'none';
      document.getElementById('formTipo').value = 'Hábito';
      document.getElementById('formNombre').value = p.nombre || '';
      document.getElementById('formCategoria').value = p.categoria || 'General';
      document.getElementById('formFrecuencia').value = p.frecuencia || 'Diaria';
      document.getElementById('formMomento').value = p.horaMomento || 'Cualquier momento';
      document.getElementById('groupFrecuencia').style.display = 'block';
      document.getElementById('groupMomento').style.display = 'block';
      document.getElementById('crudModal').classList.add('open');
    }

    function cerrarModal() {
      document.getElementById('crudModal').classList.remove('open');
    }

    function guardarFormularioCrud() {
      const modo = document.getElementById('formModo').value;
      const idFila = document.getElementById('formIdFila').value;
      const tipo = document.getElementById('formTipo').value;
      const nombre = document.getElementById('formNombre').value.trim();
      const categoria = document.getElementById('formCategoria').value;
      const frecuencia = document.getElementById('formFrecuencia').value;
      const horaMomento = document.getElementById('formMomento').value.trim() || 'Cualquier momento';

      if (!nombre) {
        mostrarToast('⚠️ Escribe un nombre');
        return;
      }

      cerrarModal();
      mostrarToast('⏳ Guardando cambios...');

      if (modo === 'CREAR') {
        google.script.run
          .withSuccessHandler(res => { actualizarUI(res.dashboard); mostrarToast(res.mensaje); })
          .withFailureHandler(err => mostrarToast('❌ ' + err.message))
          .apiAccionDirecta({
            accion: 'CREAR',
            tipo: tipo,
            nombre: nombre,
            categoria: categoria,
            frecuencia: frecuencia,
            horaMomento: horaMomento,
            guardarEnPlantilla: (tipo === 'Hábito'),
            forzarHoy: true
          });
      } else if (modo === 'EDITAR_HOY') {
        google.script.run
          .withSuccessHandler(res => { actualizarUI(res.dashboard); mostrarToast(res.mensaje); })
          .withFailureHandler(err => mostrarToast('❌ ' + err.message))
          .apiAccionDirecta({
            accion: 'EDITAR_HOY',
            idFila: Number(idFila),
            tipo: tipo,
            nombre: nombre,
            categoria: categoria
          });
      } else if (modo === 'EDITAR_PLANTILLA') {
        google.script.run
          .withSuccessHandler(res => { actualizarUI(res.dashboard); mostrarToast(res.mensaje); })
          .withFailureHandler(err => mostrarToast('❌ ' + err.message))
          .apiGestionarPlantilla({
            modo: 'EDITAR_PLANTILLA',
            idFila: Number(idFila),
            nombre: nombre,
            categoria: categoria,
            frecuencia: frecuencia,
            horaMomento: horaMomento,
            activo: true
          });
      }
    }

    /* ==================== ACCIONES DIRECTAS DE FILAS ==================== */
    function cargarDatos() {
      mostrarToast('⏳ Sincronizando...');
      google.script.run
        .withSuccessHandler(res => { actualizarUI(res); mostrarToast('✅ Sincronizado'); })
        .withFailureHandler(err => mostrarToast('❌ ' + err.message))
        .apiObtenerDatos();
    }

    function cambiarEstadoItem(item, marcarCompletado) {
      mostrarToast(marcarCompletado ? '✅ Completando...' : '↩️ Reabriendo...');
      google.script.run
        .withSuccessHandler(res => { actualizarUI(res.dashboard); mostrarToast(res.mensaje); })
        .withFailureHandler(err => mostrarToast('❌ ' + err.message))
        .apiAccionDirecta({
          accion: marcarCompletado ? 'COMPLETAR' : 'REABRIR',
          idFila: item.idFila,
          nombre: item.nombre,
          tipo: item.tipo,
          categoria: item.categoria
        });
    }

    function borrarItemHoy(item) {
      if (!confirm(\`¿Eliminar "\${item.nombre}" de la lista de hoy?\`)) return;
      mostrarToast('🗑️ Eliminando...');
      google.script.run
        .withSuccessHandler(res => { actualizarUI(res.dashboard); mostrarToast(res.mensaje); })
        .withFailureHandler(err => mostrarToast('❌ ' + err.message))
        .apiAccionDirecta({ accion: 'BORRAR', idFila: item.idFila, nombre: item.nombre });
    }

    function toggleActivoPlantilla(p) {
      mostrarToast('⏳ Actualizando estado...');
      google.script.run
        .withSuccessHandler(res => { actualizarUI(res.dashboard); mostrarToast(res.mensaje); })
        .withFailureHandler(err => mostrarToast('❌ ' + err.message))
        .apiGestionarPlantilla({
          modo: 'TOGGLE_ACTIVO',
          idFila: p.idFila,
          activo: !p.activo
        });
    }

    function borrarPlantilla(p) {
      const borrarTambienHoy = confirm(\`¿Seguro que quieres eliminar el hábito "\${p.nombre}" de tu catálogo Plantillas_Habitos?\\n\\n(También se quitará de la lista de hoy si estaba pendiente).\`);
      if (!borrarTambienHoy) return;
      mostrarToast('🗑️ Eliminando del catálogo...');
      google.script.run
        .withSuccessHandler(res => { actualizarUI(res.dashboard); mostrarToast(res.mensaje); })
        .withFailureHandler(err => mostrarToast('❌ ' + err.message))
        .apiGestionarPlantilla({
          modo: 'BORRAR_PLANTILLA',
          idFila: p.idFila,
          borrarDeHoy: true
        });
    }

    function volcarHabitosHoy() {
      mostrarToast('📥 Volcando a Sheets y Google Tasks...');
      google.script.run
        .withSuccessHandler(res => { actualizarUI(res.dashboard); mostrarToast('✅ Hábitos volcados'); })
        .withFailureHandler(err => mostrarToast('❌ ' + err.message))
        .apiVolcarHabitosHoy();
    }

    window.onload = cargarDatos;
  </script>
</body>
</html>`;
}
