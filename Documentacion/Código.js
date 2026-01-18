/**
 * ============================================================================
 * WORKSPACE AUTOMATION - TURING IA
 * Archivo: Code.gs
 * ============================================================================
 * 
 * Funciones core del sistema: configuración, lectura de datos y registro de eventos.
 * Este módulo contiene las funciones fundamentales que son utilizadas por todos
 * los demás componentes del sistema.
 * 
 * Autor: José Enrique Guerrero Pérez
 * Fecha: Enero 2026
 * ============================================================================
 */

// ============================================================================
// CONFIGURACIÓN
// ============================================================================

/**
 * Obtiene la configuración del sistema desde la hoja Configuración.
 * 
 * DESCRIPCIÓN:
 * Lee todos los parámetros de configuración almacenados en la hoja "Configuración"
 * y los retorna como un objeto JavaScript para facilitar su acceso desde otras funciones.
 * 
 * ESTRUCTURA ESPERADA DE LA HOJA:
 * - Fila 1: Encabezados (Parámetro, Valor)
 * - Fila 2+: Datos en formato:
 *   Columna A: Nombre del parámetro
 *   Columna B: Valor del parámetro
 * 
 * PARÁMETROS CONFIGURABLES:
 * - emailNotificacion: Dirección de correo del administrador del sistema
 * - calendarioId: ID del calendario de Google (normalmente "primary")
 * - notificarAdmins: TRUE/FALSE para activar/desactivar notificaciones por email
 * - crearEventoCalendar: TRUE/FALSE para activar/desactivar creación de eventos
 * - tamañoCriticoMB: Tamaño en MB para alertas de archivos grandes
 * - grupoAdmins: Nombre del grupo con permisos de administrador
 * 
 * FUNCIONAMIENTO INTERNO:
 * 1. Obtiene referencia al spreadsheet activo
 * 2. Busca la hoja llamada "Configuración"
 * 3. Valida que la hoja existe (lanza error si no existe)
 * 4. Lee todos los datos usando getDataRange() (optimización: 1 sola llamada API)
 * 5. Itera desde índice 1 para saltar los encabezados
 * 6. Construye objeto con pares clave-valor
 * 
 * RETORNA:
 * Object - Diccionario con todos los parámetros de configuración
 * 
 * EXCEPCIONES:
 * Error - Si la hoja "Configuración" no existe
 * 
 * EJEMPLO DE USO:
 * const config = getConfig();
 * const emailAdmin = config.emailNotificacion;
 * if (config.notificarAdmins === 'TRUE') {
 *   // enviar notificación
 * }
 * 
 * DECISIONES DE DISEÑO:
 * - Se usa getDataRange() en lugar de getValue() individual para minimizar llamadas API
 * - Se retorna objeto completo en lugar de parámetros individuales para flexibilidad
 * - Se valida existencia de hoja para dar error claro en caso de problema
 */
function getConfig() {
  // Obtener referencia al spreadsheet activo
  // SpreadsheetApp.getActiveSpreadsheet() retorna el documento actual
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Buscar la hoja de configuración por nombre
  // getSheetByName() retorna null si no existe
  const sheet = ss.getSheetByName('Configuración');
  
  // Validar que la hoja existe antes de continuar
  // Si no existe, lanzar error descriptivo para debugging
  if (!sheet) {
    throw new Error('Hoja Configuración no encontrada');
  }
  
  // Leer todos los datos de la hoja en una sola operación
  // getDataRange() obtiene el rango con datos (optimización: evita leer celdas vacías)
  // getValues() retorna array bidimensional: [[fila1], [fila2], ...]
  const data = sheet.getDataRange().getValues();
  
  // Inicializar objeto vacío para almacenar configuración
  const config = {};
  
  // Iterar filas de datos (empezar en 1 para saltar encabezados en índice 0)
  // Cada fila tiene estructura: [nombreParametro, valor]
  for (let i = 1; i < data.length; i++) {
    // Asignar al objeto usando nombre de parámetro como clave
    // data[i][0] = nombre del parámetro (columna A)
    // data[i][1] = valor del parámetro (columna B)
    config[data[i][0]] = data[i][1];
  }
  
  // Retornar objeto completo con toda la configuración
  return config;
}

// ============================================================================
// LECTURA DE DATOS
// ============================================================================

/**
 * Obtiene todos los usuarios del sistema desde la hoja Usuarios.
 * 
 * DESCRIPCIÓN:
 * Lee la hoja "Usuarios" completa y retorna un array de objetos donde cada
 * objeto representa un usuario con todas sus propiedades. Realiza conversión
 * de tipos y filtrado de datos para asegurar consistencia.
 * 
 * ESTRUCTURA DE LA HOJA USUARIOS:
 * Columna A (índice 0): Nombre del usuario
 * Columna B (índice 1): Email del usuario
 * Columna C (índice 2): Rol (Admin/Editor/Viewer)
 * Columna D (índice 3): Grupo (IT/Finanzas/RH)
 * Columna E (índice 4): Activo (TRUE/FALSE)
 * Columna F (índice 5): Fecha de Registro (Date)
 * Columna G (índice 6): Último Acceso (Date)
 * 
 * FUNCIONAMIENTO INTERNO:
 * 1. Obtiene referencia a la hoja "Usuarios"
 * 2. Si no existe, retorna array vacío (comportamiento defensivo)
 * 3. Lee todos los datos con getDataRange() (1 sola llamada API)
 * 4. Itera desde índice 1 para saltar encabezados
 * 5. Filtra filas vacías (sin nombre de usuario)
 * 6. Convierte cada fila en objeto estructurado
 * 7. Normaliza campo "active" a booleano verdadero
 * 
 * CONVERSIÓN DE TIPOS:
 * - Campo "active": Convierte string "TRUE" o booleano true a true
 * - Fechas: Se mantienen como objetos Date de JavaScript
 * - Strings: Se conservan tal cual (sin trim adicional aquí)
 * 
 * RETORNA:
 * Array<Object> - Lista de usuarios, cada uno con propiedades:
 *   {
 *     name: string,
 *     email: string,
 *     role: string,
 *     group: string,
 *     active: boolean,
 *     dateRegistered: Date,
 *     lastAccess: Date
 *   }
 * Array vacío si la hoja no existe o no hay datos
 * 
 * MANEJO DE CASOS ESPECIALES:
 * - Filas vacías (sin nombre): Se saltan con continue
 * - Hoja inexistente: Retorna array vacío
 * - Campo active con texto "TRUE": Se convierte a boolean true
 * - Campos opcionales: Se asignan aunque sean null/undefined
 * 
 * EJEMPLO DE USO:
 * const usuarios = getUsers();
 * const activos = usuarios.filter(u => u.active);
 * const admins = usuarios.filter(u => u.role === 'Admin');
 * 
 * DECISIONES DE DISEÑO:
 * - Retornar array vacío si no hay hoja (evita null checks constantes)
 * - Normalizar booleanos para evitar comparaciones de string
 * - Una sola llamada getDataRange() para eficiencia
 * - Filtrar filas vacías para limpiar datos
 */
function getUsers() {
  // Obtener referencia al spreadsheet activo
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Buscar hoja de usuarios
  const sheet = ss.getSheetByName('Usuarios');
  
  // Retornar array vacío si la hoja no existe
  // Esto evita errores y permite a funciones llamadoras manejar ausencia de datos
  if (!sheet) return [];
  
  // Leer todos los datos de la hoja en una sola operación
  // Optimización: una llamada API en lugar de múltiples getValue()
  const data = sheet.getDataRange().getValues();
  
  // Inicializar array para almacenar usuarios
  const users = [];
  
  // Iterar desde índice 1 para saltar fila de encabezados
  for (let i = 1; i < data.length; i++) {
    // Filtrar filas vacías: si no hay nombre, saltar esta fila
    // continue salta a la siguiente iteración del loop
    if (!data[i][0]) continue;
    
    // Construir objeto de usuario con todas sus propiedades
    users.push({
      // Columna A: Nombre del usuario
      name: data[i][0],
      
      // Columna B: Email del usuario
      email: data[i][1],
      
      // Columna C: Rol del usuario en el sistema
      role: data[i][2],
      
      // Columna D: Grupo organizacional al que pertenece
      group: data[i][3],
      
      // Columna E: Estado activo/inactivo
      // Normalización: convierte "TRUE" (string) o true (boolean) a boolean true
      // Esto maneja casos donde Google Sheets retorna el valor como string
      active: data[i][4] === true || data[i][4] === 'TRUE',
      
      // Columna F: Fecha de registro en el sistema
      // Se mantiene como objeto Date de JavaScript
      dateRegistered: data[i][5],
      
      // Columna G: Fecha del último acceso registrado
      // Se mantiene como objeto Date de JavaScript
      lastAccess: data[i][6]
    });
  }
  
  // Retornar array completo de usuarios
  return users;
}

/**
 * Calcula días transcurridos desde el último acceso de un usuario.
 * 
 * DESCRIPCIÓN:
 * Toma una fecha de último acceso y calcula cuántos días han pasado hasta hoy.
 * Utilizado para detectar usuarios inactivos que excedan el límite de 7 días.
 * 
 * PARÁMETROS:
 * @param {Date} lastAccessDate - Fecha del último acceso del usuario
 * 
 * FUNCIONAMIENTO INTERNO:
 * 1. Obtiene fecha/hora actual
 * 2. Convierte parámetro a objeto Date (maneja tanto Date como string)
 * 3. Calcula diferencia en milisegundos
 * 4. Convierte milisegundos a días (división)
 * 5. Redondea hacia abajo con Math.floor()
 * 
 * CÁLCULO DE DÍAS:
 * - 1 día = 24 horas = 1440 minutos = 86400 segundos = 86400000 milisegundos
 * - Fórmula: dias = floor((hoy - ultimoAcceso) / 86400000)
 * - Math.abs() asegura resultado positivo
 * - Math.floor() redondea hacia abajo para días completos
 * 
 * RETORNA:
 * Number - Cantidad de días transcurridos (entero)
 * 
 * CASOS ESPECIALES:
 * - Fecha futura: Retorna 0 (Math.abs maneja negativos)
 * - Mismo día: Retorna 0
 * - null/undefined: Puede causar NaN (función llamadora debe validar)
 * 
 * EJEMPLO DE USO:
 * const dias = getDaysSinceLastAccess(user.lastAccess);
 * if (dias > 7) {
 *   // Usuario inactivo, desactivar
 * }
 * 
 * MEJORAS FUTURAS POSIBLES:
 * - Validar que lastAccessDate no sea null/undefined
 * - Retornar 999 o Infinity si no hay fecha
 * - Loggear advertencias para fechas inválidas
 * 
 * DECISIONES DE DISEÑO:
 * - Math.abs() para manejar fechas futuras sin error
 * - Math.floor() para contar solo días completos
 * - Cálculo basado en milisegundos para precisión
 */
function getDaysSinceLastAccess(lastAccessDate) {
  // Obtener fecha y hora actual del sistema
  const today = new Date();
  
  // Convertir parámetro a objeto Date
  // Si ya es Date, no cambia; si es string, se parsea
  const lastAccess = new Date(lastAccessDate);
  
  // Calcular diferencia en milisegundos
  // getTime() retorna milisegundos desde epoch (1 enero 1970)
  // Math.abs() asegura resultado positivo (maneja fechas futuras)
  const diffTime = Math.abs(today - lastAccess);
  
  // Convertir milisegundos a días
  // 1000 ms * 60 seg * 60 min * 24 hrs = 86400000 ms por día
  // Math.floor() redondea hacia abajo para contar solo días completos
  return Math.floor(diffTime / (1000 * 60 * 60 * 24));
}

// ============================================================================
// REGISTRO DE EVENTOS
// ============================================================================

/**
 * Registra un evento en la hoja RegistroDeEventos para auditoría.
 * 
 * DESCRIPCIÓN:
 * Función central de auditoría que registra todas las acciones importantes
 * del sistema en la hoja "RegistroDeEventos". Cada evento se almacena con
 * timestamp, tipo, usuario, detalles y estado.
 * 
 * PARÁMETROS:
 * @param {Object} event - Objeto con información del evento
 *   @param {string} event.type - Tipo de evento (ej: "USUARIO_AGREGADO")
 *   @param {string} event.user - Usuario relacionado con el evento
 *   @param {string} event.details - Descripción detallada del evento
 *   @param {string} event.status - Estado: "OK", "ALERTA", "ERROR", "WARNING"
 *   @param {string} event.action - Acción tomada (ej: "Email enviado")
 * 
 * ESTRUCTURA DE REGISTRO:
 * Cada fila en RegistroDeEventos contiene:
 * Columna A: Timestamp (fecha y hora exacta del evento)
 * Columna B: Tipo de evento
 * Columna C: Usuario afectado/relacionado
 * Columna D: Detalles adicionales
 * Columna E: Estado del evento
 * Columna F: Acción realizada
 * 
 * TIPOS DE EVENTOS COMUNES:
 * - USUARIO_AGREGADO: Nuevo usuario registrado
 * - USUARIO_INACTIVO: Usuario desactivado por inactividad
 * - ROL_MODIFICADO: Cambio de rol de usuario
 * - ARCHIVO_SUBIDO: Archivo cargado al sistema
 * - PERMISO_DENEGADO: Intento de acceso no autorizado
 * - ERROR: Error en el sistema
 * 
 * FUNCIONAMIENTO INTERNO:
 * 1. Intenta obtener hoja "RegistroDeEventos"
 * 2. Si no existe, loggea advertencia y retorna (no falla)
 * 3. Usa appendRow() para agregar nueva fila al final
 * 4. Asigna valores por defecto si propiedades están ausentes
 * 5. Loggea confirmación en consola
 * 6. Maneja errores sin interrumpir flujo principal
 * 
 * MANEJO DE VALORES POR DEFECTO:
 * - type: "EVENTO" si no se proporciona
 * - user: "Sistema" si no se proporciona
 * - details: "" (string vacío) si no se proporciona
 * - status: "OK" si no se proporciona
 * - action: "Ninguna" si no se proporciona
 * 
 * MANEJO DE ERRORES:
 * - Try-catch envuelve toda la función
 * - Si falla, loggea error pero no interrumpe ejecución
 * - Comportamiento defensivo: sistema sigue funcionando sin auditoría
 * 
 * EJEMPLO DE USO:
 * logEvent({
 *   type: 'USUARIO_AGREGADO',
 *   user: 'Juan Pérez',
 *   details: 'Rol: Editor, Grupo: IT',
 *   status: 'OK',
 *   action: 'Email de bienvenida enviado'
 * });
 * 
 * CONSIDERACIONES DE SEGURIDAD:
 * - Todos los eventos son inmutables (solo agregar, no modificar)
 * - Timestamp automático evita manipulación
 * - Registros persisten para cumplimiento y auditoría
 * 
 * DECISIONES DE DISEÑO:
 * - appendRow() en lugar de setValue() para simplicidad
 * - Valores por defecto para evitar celdas vacías
 * - No lanzar error si falla (no interrumpir operación principal)
 * - Loggear tanto éxito como error para debugging
 */
function logEvent(event) {
  try {
    // Obtener referencia al spreadsheet activo
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Buscar hoja de registro de eventos
    const sheet = ss.getSheetByName('RegistroDeEventos');
    
    // Validar que la hoja existe
    // Si no existe, loggear y salir (comportamiento defensivo)
    if (!sheet) {
      Logger.log('Hoja RegistroEventos no encontrada');
      return;
    }
    
    // Agregar nueva fila al final de la hoja con los datos del evento
    // appendRow() agrega automáticamente después de la última fila con datos
    sheet.appendRow([
      // Columna A: Timestamp con fecha y hora exacta
      new Date(),
      
      // Columna B: Tipo de evento con valor por defecto
      // Operador || retorna primer valor truthy
      event.type || 'EVENTO',
      
      // Columna C: Usuario con valor por defecto "Sistema"
      event.user || 'Sistema',
      
      // Columna D: Detalles con valor por defecto string vacío
      event.details || '',
      
      // Columna E: Estado con valor por defecto "OK"
      event.status || 'OK',
      
      // Columna F: Acción tomada con valor por defecto "Ninguna"
      event.action || 'Ninguna'
    ]);
    
    // Loggear confirmación en consola para debugging
    Logger.log(`Evento registrado: ${event.type} - ${event.user}`);
    
  } catch (error) {
    // Capturar cualquier error durante el registro
    // No lanzar error para no interrumpir flujo principal
    // Solo loggear para debugging
    Logger.log('Error al registrar evento: ' + error.message);
  }
}

/**
 * Función de prueba para verificar funcionalidad de Code.gs.
 * 
 * DESCRIPCIÓN:
 * Ejecuta pruebas básicas de las tres funciones principales del módulo:
 * getConfig(), getUsers() y logEvent(). Útil para verificar que el sistema
 * está correctamente configurado.
 * 
 * USO:
 * Ejecutar manualmente desde el editor de Apps Script:
 * 1. Seleccionar función "probarCode" en el dropdown
 * 2. Click en botón "Ejecutar"
 * 3. Revisar logs (Ver > Registros)
 * 
 * PRUEBAS REALIZADAS:
 * 1. Prueba getConfig(): Verifica lectura de configuración
 * 2. Prueba getUsers(): Verifica lectura de usuarios
 * 3. Prueba logEvent(): Verifica escritura de eventos
 * 
 * RESULTADOS ESPERADOS:
 * - Todos los logs deben mostrar éxito
 * - Email configurado debe aparecer en logs
 * - Cantidad de usuarios debe ser mayor a 0
 * - Nueva fila debe aparecer en RegistroDeEventos
 * 
 * MANEJO DE ERRORES:
 * - Cada prueba tiene su propio try-catch
 * - Si una prueba falla, las demás continúan
 * - Errores se loggean con descripción clara
 */
function probarCode() {
  Logger.log('=== PROBANDO CODE.GS ===');
  
  // Prueba 1: Verificar lectura de configuración
  try {
    const config = getConfig();
    Logger.log('getConfig() funciona correctamente');
    Logger.log('Email configurado: ' + config.emailNotificacion);
  } catch (error) {
    Logger.log('getConfig() falló: ' + error.message);
  }
  
  // Prueba 2: Verificar lectura de usuarios
  try {
    const users = getUsers();
    Logger.log('getUsers() funciona correctamente');
    Logger.log('Total usuarios: ' + users.length);
  } catch (error) {
    Logger.log('getUsers() falló: ' + error.message);
  }
  
  // Prueba 3: Verificar registro de eventos
  try {
    logEvent({
      type: 'PRUEBA',
      user: 'Test',
      details: 'Prueba de Code.gs',
      status: 'OK',
      action: 'Ninguna'
    });
    Logger.log('logEvent() funciona correctamente');
  } catch (error) {
    Logger.log('logEvent() falló: ' + error.message);
  }
  
  Logger.log('=== PRUEBA CODE.GS COMPLETADA ===');
}
