/**
 * ============================================================================
 * WORKSPACE AUTOMATION - TURING IA
 * Archivo: Triggers.gs
 * ============================================================================
 * 
 * Triggers automáticos y funciones de prueba.
 * Contiene los activadores (triggers) que responden a eventos del sistema:
 * - onEdit: Trigger simple para ediciones de hoja
 * - verificarUsuariosInactivos: Trigger temporal diario
 * También incluye suite completa de funciones de prueba para validación.
 * 
 * Autor: José Enrique Guerrero Pérez
 * Fecha: Enero 2026
 * ============================================================================
 */

// ============================================================================
// TRIGGER: EDICIÓN DE HOJA
// ============================================================================

/**
 * Trigger simple que se ejecuta automáticamente al editar el Spreadsheet.
 * 
 * DESCRIPCIÓN:
 * Función de trigger que Google Apps Script ejecuta automáticamente cada vez
 * que alguien edita una celda en el spreadsheet. Actúa como punto de entrada
 * para manejar diferentes tipos de ediciones según la hoja modificada.
 * 
 * CONFIGURACIÓN COMO TRIGGER:
 * Este es un trigger SIMPLE (no instalable), por lo que:
 * - Se activa automáticamente al nombrar la función "onEdit"
 * - No requiere instalación manual
 * - Tiene limitaciones de permisos (no puede enviar emails/crear eventos)
 * - Para funcionalidad completa, usar trigger instalable "onEditInstalable"
 * 
 * PARÁMETROS:
 * @param {Object} e - Objeto de evento generado por Google Sheets
 *   Propiedades importantes de e:
 *     @param {Range} e.range - Rango de celdas editado
 *     @param {Sheet} e.range.getSheet() - Hoja donde ocurrió la edición
 *     @param {number} e.range.getRow() - Número de fila editada
 *     @param {number} e.range.getColumn() - Número de columna editada
 *     @param {any} e.value - Nuevo valor de la celda (puede ser undefined)
 *     @param {any} e.oldValue - Valor anterior (puede ser undefined)
 * 
 * FUNCIONAMIENTO INTERNO:
 * 1. Captura objeto de evento e
 * 2. Obtiene nombre de la hoja editada
 * 3. Enruta según nombre de hoja:
 *    - "Usuarios" → llama handleUserEdit(e)
 *    - Otras hojas → ignora (no hace nada)
 * 4. Si hay error, lo captura y registra
 * 
 * MANEJO DE ERRORES:
 * - Try-catch envuelve toda la lógica
 * - Errores se loggean en consola
 * - Errores también se registran en RegistroDeEventos
 * - No se lanzan errores para no interrumpir ediciones del usuario
 * 
 * LIMITACIONES DEL TRIGGER SIMPLE:
 * - No puede acceder a servicios que requieren autorización
 * - GmailApp.sendEmail() falla silenciosamente
 * - CalendarApp.createEvent() falla silenciosamente
 * - No puede mostrar UI (Browser.msgBox, etc)
 * - Ejecuta con permisos del usuario que edita, no del owner
 * 
 * SOLUCIÓN A LIMITACIONES:
 * Para funcionalidad completa (emails, calendar), usar trigger INSTALABLE:
 * 1. Crear función "onEditInstalable" con misma lógica
 * 2. Instalar manualmente en: Activadores > Agregar activador
 * 3. Configurar como: Al editar > Desde hoja de cálculo
 * 
 * HOJAS MONITOREADAS:
 * Actualmente solo:
 * - "Usuarios": Detecta cambios en usuarios (nuevos, rol, activo)
 * Futuras expansiones posibles:
 * - "Configuración": Detectar cambios en parámetros
 * - "Grupos": Detectar cambios en estructura organizacional
 * 
 * FLUJO DE EJECUCIÓN:
 * Usuario edita celda → Google activa onEdit(e) → 
 * Obtiene nombre de hoja → Si es "Usuarios" → handleUserEdit(e) →
 * Procesa según columna editada
 * 
 * EJEMPLO DE USO:
 * No se llama manualmente. Google lo ejecuta automáticamente.
 * Para probar manualmente, usar funciones de prueba en este archivo.
 * 
 * DEBUGGING:
 * - Revisar logs: Ver > Registros (Ctrl+Enter)
 * - Revisar RegistroDeEventos para errores
 * - Para testing, usar pruebaCompletaNuevoUsuario()
 * 
 * DECISIONES DE DISEÑO:
 * - Nombre "onEdit" para activación automática
 * - Enrutamiento por nombre de hoja para escalabilidad
 * - Try-catch general para robustez
 * - Registro de errores para debugging
 * - No interrumpir ediciones del usuario
 */
function onEdit(e) {
  try {
    // Obtener nombre de la hoja donde ocurrió la edición
    // e.range = objeto Range con información de la celda editada
    // getSheet() retorna objeto Sheet
    // getName() retorna nombre de la hoja como string
    const sheetName = e.range.getSheet().getName();

    // Enrutar según nombre de hoja
    // Solo procesamos cambios en hoja "Usuarios"
    if (sheetName === 'Usuarios') {
      handleUserEdit(e);
    }
    // Si es otra hoja, no hacer nada
    // Aquí se pueden agregar más casos: else if (sheetName === 'Configuración')
    
  } catch (error) {
    // Capturar cualquier error durante la ejecución
    // Loggear en consola para debugging inmediato
    Logger.log('Error en onEdit: ' + error.message);
    
    // Registrar error en hoja de auditoría
    // Esto permite revisar errores históricos
    logEvent({
      type: 'ERROR',
      user: 'Sistema',
      details: 'Error en trigger: ' + error.message,
      status: 'ERROR'
    });
  }
}

/**
 * Maneja cambios en la hoja Usuarios.
 * 
 * DESCRIPCIÓN:
 * Función central que procesa todas las ediciones en la hoja "Usuarios".
 * Detecta tres tipos de cambios importantes:
 * 1. Nuevo usuario (fila completa sin fecha de registro)
 * 2. Cambio de estado activo (columna E)
 * 3. Cambio de rol (columna C)
 * 
 * PARÁMETROS:
 * @param {Object} e - Objeto de evento del trigger onEdit
 * 
 * ESTRUCTURA DE LA HOJA USUARIOS:
 * Columna A (1): Nombre
 * Columna B (2): Email
 * Columna C (3): Rol
 * Columna D (4): Grupo
 * Columna E (5): Activo
 * Columna F (6): Fecha Registro
 * Columna G (7): Último Acceso
 * 
 * FUNCIONAMIENTO INTERNO - PASO A PASO:
 * 
 * PASO 1: Obtener información de la edición
 * - Número de fila editada (row)
 * - Si es fila 1 (encabezados), salir inmediatamente
 * - Referencia a la hoja
 * - Número de columna editada (col)
 * 
 * PASO 2: Leer datos completos de la fila
 * - getRange(row, 1, 1, 7) obtiene 7 columnas de esa fila
 * - getValues()[0] convierte a array y toma primera fila
 * - Construye objeto user con todas las propiedades
 * 
 * PASO 3: Normalizar datos
 * - Convertir strings a strings limpios con trim()
 * - Convertir "TRUE"/"true"/true a boolean true
 * - Manejar valores null/undefined
 * 
 * PASO 4: Detectar nuevo usuario (columnas A-E)
 * - Solo si la columna editada está entre 1 y 5
 * - Verificar que TODAS las columnas básicas estén llenas
 * - Verificar que NO tenga fecha de registro (es nuevo)
 * - Si ambas condiciones se cumplen → processNewUser()
 * 
 * PASO 5: Detectar cambio de estado activo (columna E)
 * - Solo si columna editada es 5 (Activo)
 * - Solo si usuario YA tiene fecha de registro
 * - Si nuevo valor es false → processInactiveUser()
 * 
 * PASO 6: Detectar cambio de rol (columna C)
 * - Solo si columna editada es 3 (Rol)
 * - Solo si usuario YA tiene fecha de registro
 * - Siempre notificar (alerta especial si es Admin)
 * - notifyRoleChange()
 * 
 * LÓGICA DE DETECCIÓN DE NUEVO USUARIO:
 * Problema original: Al llenar celda por celda, ¿cuándo procesar?
 * Solución: Verificar en cada edición de A-E si la fila está completa
 * 
 * Escenario: Usuario llena A6, B6, C6, D6, E6
 * - Edita A6: filaCompleta=false (falta B,C,D,E) → espera
 * - Edita B6: filaCompleta=false (falta C,D,E) → espera
 * - Edita C6: filaCompleta=false (falta D,E) → espera
 * - Edita D6: filaCompleta=false (falta E) → espera
 * - Edita E6: filaCompleta=true (todo lleno) → PROCESA
 * 
 * CONVERSIÓN DE TIPOS:
 * - userData[0] puede ser string o null
 * - toString().trim() convierte a string y quita espacios
 * - Operador || retorna '' si es null/undefined
 * - Campo active acepta boolean true o string "TRUE"/"true"
 * 
 * VALIDACIONES IMPLEMENTADAS:
 * 1. row === 1: Ignorar encabezados
 * 2. !user.name: Ignorar filas vacías
 * 3. filaCompleta: Asegurar datos mínimos
 * 4. esNuevo: Evitar reprocesar usuarios existentes
 * 5. user.dateRegistered: Solo procesar cambios de usuarios registrados
 * 
 * CASOS ESPECIALES MANEJADOS:
 * - Celda vacía → string vacío
 * - TRUE como texto → boolean true
 * - Fecha null → null (no undefined)
 * - Edición de múltiples celdas → procesa una a la vez
 * 
 * FLUJO PARA NUEVO USUARIO:
 * handleUserEdit detecta fila completa →
 * processNewUser actualiza fechas →
 * envía email HTML →
 * crea evento Calendar →
 * registra en RegistroDeEventos
 * 
 * FLUJO PARA USUARIO INACTIVO:
 * handleUserEdit detecta col=5 y active=false →
 * processInactiveUser registra evento →
 * envía email de alerta
 * 
 * FLUJO PARA CAMBIO DE ROL:
 * handleUserEdit detecta col=3 →
 * notifyRoleChange registra evento →
 * si es Admin, envía email de alerta
 * 
 * OPTIMIZACIONES:
 * - Una sola lectura de fila completa (no 7 lecturas individuales)
 * - Return temprano después de procesar nuevo usuario
 * - Validación dateRegistered para evitar procesar en cascada
 * 
 * LOGGING:
 * - Logs informativos en cada paso
 * - Logs de estado de verificación
 * - Logs de acciones tomadas
 * 
 * LIMITACIONES:
 * - No detecta edición de múltiples filas simultáneas
 * - No detecta copiar/pegar de bloques grandes
 * - No valida formato de email
 * 
 * EJEMPLO DE USO:
 * No se llama directamente. onEdit(e) la llama cuando detecta edición en "Usuarios".
 * 
 * DECISIONES DE DISEÑO:
 * - Leer fila completa para tener contexto total
 * - Verificar completitud antes de procesar
 * - Separar lógica por tipo de cambio (nuevo/inactivo/rol)
 * - Return temprano para evitar procesamiento múltiple
 * - Validar dateRegistered para distinguir nuevos de existentes
 */
function handleUserEdit(e) {
  // PASO 1: Obtener información básica de la edición
  
  // Obtener número de fila editada (base 1)
  const row = e.range.getRow();
  
  // Si es la fila 1 (encabezados), salir inmediatamente
  // No queremos procesar ediciones en los encabezados
  if (row === 1) return;
  
  // Obtener referencia a la hoja editada
  const sheet = e.range.getSheet();
  
  // Obtener número de columna editada (base 1)
  // A=1, B=2, C=3, D=4, E=5, F=6, G=7
  const col = e.range.getColumn();
  
  // PASO 2: Leer datos completos de la fila del usuario
  
  // Obtener todas las columnas de la fila (A-G = 7 columnas)
  // getRange(fila, columnaInicio, numFilas, numColumnas)
  // getValues() retorna array bidimensional: [[col1, col2, ...]]
  // [0] toma el primer (y único) elemento del array externo
  const userData = sheet.getRange(row, 1, 1, 7).getValues()[0];
  
  // PASO 3: Construir objeto de usuario con normalización de datos
  
  const user = {
    // Columna A: Nombre
    // Si existe, convertir a string y limpiar espacios
    // Si es null/undefined, usar string vacío
    name: userData[0] ? userData[0].toString().trim() : '',
    
    // Columna B: Email
    email: userData[1] ? userData[1].toString().trim() : '',
    
    // Columna C: Rol
    role: userData[2] ? userData[2].toString().trim() : '',
    
    // Columna D: Grupo
    group: userData[3] ? userData[3].toString().trim() : '',
    
    // Columna E: Activo
    // Normalizar a boolean: acepta true (boolean) o "TRUE"/"true" (string)
    // Esto maneja casos donde Sheets retorna el valor como string
    active: userData[4] === true || userData[4] === 'TRUE' || userData[4] === 'true',
    
    // Columna F: Fecha de Registro
    // Mantener como Date o null (no convertir)
    dateRegistered: userData[5] || null,
    
    // Columna G: Último Acceso
    // Mantener como Date o null (no convertir)
    lastAccess: userData[6] || null
  };
  
  // ===================================================================
  // DETECCIÓN 1: NUEVO USUARIO
  // Se ejecuta SOLO cuando la fila está COMPLETA y sin fecha de registro
  // ===================================================================
  
  // Verificar si la columna editada es una de las columnas básicas (A-E)
  // Solo verificamos nuevo usuario cuando se edita información básica
  if (col >= 1 && col <= 5) {
    
    // Verificar que TODAS las columnas básicas estén llenas
    // name, email, role y group son obligatorios
    // No verificamos active porque puede ser false
    const filaCompleta = user.name && user.email && user.role && user.group;
    
    // Verificar si NO tiene fecha de registro (indica que es nuevo)
    // dateRegistered será null o '' para usuarios nuevos
    const esNuevo = !user.dateRegistered || user.dateRegistered === '';
    
    // Si ambas condiciones se cumplen, procesar como nuevo usuario
    if (filaCompleta && esNuevo) {
      // Loggear detección de nuevo usuario
      Logger.log('Fila completa detectada - Procesando nuevo usuario');
      
      // Llamar función de procesamiento con datos de usuario y número de fila
      // El número de fila es necesario para actualizar las fechas
      processNewUser(user, row);
      
      // Salir de la función para evitar procesamiento adicional
      // Un usuario nuevo no puede ser también inactivo o cambiar de rol
      return;
    }
  }
  
  // ===================================================================
  // DETECCIÓN 2: CAMBIO DE ESTADO ACTIVO
  // Se ejecuta cuando se edita columna E (Activo)
  // ===================================================================
  
  // Verificar si la columna editada es la 5 (Activo)
  // Y si el nuevo valor es false (usuario desactivado)
  if (col === 5 && user.active === false) {
    // Procesar desactivación de usuario
    // Esto registrará el evento y enviará email de alerta
    processInactiveUser(user);
  }
  
  // ===================================================================
  // DETECCIÓN 3: CAMBIO DE ROL
  // Se ejecuta cuando se edita columna C (Rol)
  // ===================================================================
  
  // Verificar si la columna editada es la 3 (Rol)
  // Y si el usuario tiene fecha de registro (no es nuevo)
  if (col === 3 && user.dateRegistered) {
    // Notificar cambio de rol
    // Si el rol es Admin, enviará email de alerta
    // Si es otro rol, solo registrará el evento
    notifyRoleChange(user);
  }
}

// ============================================================================
// TRIGGER: VERIFICACIÓN DIARIA
// ============================================================================

/**
 * Verifica usuarios inactivos y los desactiva automáticamente.
 * 
 * DESCRIPCIÓN:
 * Función de trigger temporal que se ejecuta automáticamente cada día
 * (típicamente entre 8-9 AM). Revisa todos los usuarios activos y desactiva
 * aquellos que no han accedido en más de 7 días.
 * 
 * CONFIGURACIÓN COMO TRIGGER:
 * Apps Script > Activadores > Agregar activador:
 * - Función: verificarUsuariosInactivos
 * - Tipo de evento: Controlado por tiempo
 * - Tipo de activador: Activador de temporizador diario
 * - Hora del día: 8 a.m. - 9 a.m.
 * - Notificaciones de fallo: Notificarme diariamente
 * 
 * POLÍTICA APLICADA:
 * Si ultimoAcceso > 7 días Y usuario.activo = TRUE
 * Entonces: Desactivar + Registrar + Notificar
 * 
 * FUNCIONAMIENTO INTERNO - PASO A PASO:
 * 
 * PASO 1: Inicialización
 * - Loggear inicio de verificación
 * - Obtener todos los usuarios del sistema
 * - Loggear array completo para debugging
 * - Inicializar array vacío para usuarios a desactivar
 * 
 * PASO 2: Iteración y verificación
 * - Recorrer cada usuario
 * - Saltar usuarios ya inactivos
 * - Calcular días desde último acceso
 * - Comparar contra límite de 7 días
 * 
 * PASO 3: Desactivación (si días > 7)
 * - Buscar usuario en la hoja
 * - Actualizar columna E a FALSE
 * - Registrar evento en auditoría
 * - Agregar a lista de inactivos
 * 
 * PASO 4: Reporte
 * - Si hay usuarios desactivados, enviar email resumen
 * - Loggear cantidad de usuarios procesados
 * - Marcar verificación como completada
 * 
 * ESTRUCTURA DEL ARRAY usuariosInactivos:
 * Cada elemento contiene:
 * - name: Nombre del usuario
 * - group: Grupo al que pertenece
 * - days: Cantidad de días sin acceso
 * 
 * BÚSQUEDA Y ACTUALIZACIÓN EN HOJA:
 * - Lee toda la hoja con getDataRange()
 * - Itera desde índice 1 (saltar encabezados)
 * - Compara data[i][0] (columna A) con user.name
 * - Al encontrar coincidencia:
 *   - Actualiza fila i+1, columna 5 (E) a 'FALSE'
 *   - Break para salir del loop
 * 
 * NOTA SOBRE ÍNDICES:
 * - Array data: índice 0 = encabezados, índice 1+ = datos
 * - Hoja Sheets: fila 1 = encabezados, fila 2+ = datos
 * - Por eso usamos i+1 para getRange (convierte índice array a número fila)
 * 
 * REGISTRO DE EVENTO:
 * Se registra cada desactivación con:
 * - type: 'USUARIO_INACTIVO'
 * - user: Nombre del usuario
 * - details: Cantidad de días sin actividad
 * - status: 'ALERTA'
 * - action: 'Desactivado automáticamente'
 * 
 * REPORTE DE INACTIVOS:
 * Si hay al menos un usuario desactivado:
 * - Llama enviarReporteInactivos()
 * - Envía email HTML con tabla de usuarios
 * - Incluye: nombre, grupo, días de inactividad
 * 
 * LOGGING EXTENSIVO:
 * - Inicio y fin de verificación
 * - Array completo de usuarios (para debugging)
 * - Cantidad de usuarios desactivados
 * - Cada paso importante del proceso
 * 
 * CASOS ESPECIALES:
 * - Usuario sin fecha de último acceso: getDaysSinceLastAccess retorna 999
 * - Usuario ya inactivo: Se salta (no se procesa)
 * - Cero usuarios inactivos: No se envía email
 * 
 * VENTANA DE EJECUCIÓN:
 * Google ejecuta entre 8-9 AM pero no garantiza hora exacta
 * Puede variar ±15 minutos según carga del sistema
 * 
 * EJEMPLO DE EJECUCIÓN:
 * 8:00 AM → Trigger ejecuta verificarUsuariosInactivos()
 * → Lee 10 usuarios
 * → Usuario A: 5 días → OK, no desactivar
 * → Usuario B: 15 días → Desactivar
 * → Usuario C: 8 días → Desactivar
 * → Envía reporte con 2 usuarios desactivados
 * 
 * OPTIMIZACIONES:
 * - Una sola llamada getUsers() al inicio
 * - Filter para obtener solo usuarios activos
 * - Break al encontrar usuario en la hoja
 * 
 * LIMITACIONES:
 * - Solo se ejecuta una vez al día
 * - No desactiva usuarios inactivos que llegaron hoy al límite hasta mañana
 * - Búsqueda por nombre (puede fallar si hay nombres duplicados)
 * 
 * MEJORAS FUTURAS POSIBLES:
 * - Búsqueda por email en lugar de nombre (más único)
 * - Actualización batch en lugar de una por una
 * - Escalar límite de días según rol (admins más tiempo)
 * - Advertencia previa antes de desactivar (email día 6)
 * 
 * MONITOREO:
 * - Revisar logs diarios en: Ver > Registros
 * - Revisar RegistroDeEventos para historial
 * - Recibir email solo si hay inactivos
 * 
 * DECISIONES DE DISEÑO:
 * - Límite fijo de 7 días (podría ser configurable)
 * - Desactivación automática sin confirmación
 * - Reporte solo si hay cambios
 * - Logging extensivo para debugging
 * - No eliminar usuarios, solo desactivar
 */
function verificarUsuariosInactivos() {
  // PASO 1: Inicialización y logging
  
  Logger.log('=== Verificación diaria iniciada ===');
  
  // Obtener todos los usuarios del sistema
  // Esto incluye tanto activos como inactivos
  const users = getUsers();
  
  // Loggear array completo para debugging
  // Útil para ver estado de todos los usuarios
  Logger.log(users);
  
  // Inicializar array para almacenar usuarios que serán desactivados
  const usuariosInactivos = [];
  
  // PASO 2: Iterar sobre cada usuario y verificar inactividad
  
  users.forEach(user => {
    // Saltar usuarios que ya están inactivos
    // No necesitamos procesarlos nuevamente
    if (!user.active) return;
    
    // Calcular días transcurridos desde último acceso
    // Función retorna número de días o 999 si no hay fecha
    const dias = getDaysSinceLastAccess(user.lastAccess);
    
    // PASO 3: Si excede límite de 7 días, desactivar
    
    if (dias > 7) {
      // ACCIÓN 1: Desactivar en la hoja
      
      // Obtener referencia al spreadsheet y hoja
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('Usuarios');
      
      // Leer todos los datos para buscar al usuario
      const data = sheet.getDataRange().getValues();
      
      // Buscar usuario en la hoja por nombre
      // Iterar desde índice 1 para saltar encabezados
      for (let i = 1; i < data.length; i++) {
        // Comparar nombre en columna A con nombre del usuario
        if (data[i][0] === user.name) {
          // Encontrado: actualizar columna E (Activo) a FALSE
          // i+1 porque array empieza en 0 pero fila 1 son encabezados
          // Columna 5 = E (Activo)
          sheet.getRange(i + 1, 5).setValue('FALSE');
          
          // Salir del loop después de actualizar
          break;
        }
      }
      
      // ACCIÓN 2: Registrar evento en auditoría
      
      logEvent({
        type: 'USUARIO_INACTIVO',
        user: user.name,
        details: `${dias} días sin actividad`,
        status: 'ALERTA',
        action: 'Desactivado automáticamente'
      });
      
      // ACCIÓN 3: Agregar a lista de inactivos para reporte
      
      usuariosInactivos.push({
        name: user.name,
        group: user.group,
        days: dias
      });
    }
  });
  
  // PASO 4: Logging y reporte final
  
  // Loggear cantidad de usuarios desactivados
  Logger.log(`Usuarios desactivados: ${usuariosInactivos.length}`);
  
  // Si hay usuarios inactivos, enviar reporte por email
  if (usuariosInactivos.length > 0) {
    enviarReporteInactivos(usuariosInactivos);
  }
  
  // Marcar verificación como completada
  Logger.log('=== Verificación completada ===');
}

// ============================================================================
// FUNCIONES DE PRUEBA
// ============================================================================

/**
 * PRUEBA 1: Simula agregar usuario nuevo.
 * 
 * DESCRIPCIÓN:
 * Función de prueba que ejecuta processNewUser() con datos de prueba.
 * Útil para verificar que el flujo completo de nuevo usuario funciona:
 * registro de evento, email HTML y evento de Calendar.
 * 
 * EJECUCIÓN:
 * Manual desde editor Apps Script:
 * 1. Seleccionar "pruebaAgregarUsuario" en dropdown
 * 2. Click en botón "Ejecutar"
 * 3. Revisar logs
 * 4. Verificar email recibido
 * 5. Verificar evento en Calendar
 * 
 * DATOS DE PRUEBA:
 * - Nombre: Carlos Test
 * - Email: carlos.test@empresa.com
 * - Rol: Editor
 * - Grupo: IT
 * 
 * LIMITACIÓN:
 * Esta versión NO incluye parámetro 'row', lo que causará error
 * al intentar actualizar fechas. Usar pruebaCompletaNuevoUsuario() en su lugar.
 * 
 * RESULTADOS ESPERADOS:
 * - Log de prueba completada
 * - Email HTML en bandeja
 * - Evento de onboarding en Calendar
 * - Nueva fila en RegistroDeEventos
 * 
 * NOTA:
 * Si no recibe email/evento, verificar:
 * - Configuración tiene notificarAdmins = TRUE
 * - Configuración tiene crearEventoCalendar = TRUE
 * - Email en configuración es válido
 * - Permisos de Gmail y Calendar autorizados
 */
function pruebaAgregarUsuario() {
  Logger.log('=== PRUEBA: Agregar Usuario ===');
  
  // Crear objeto de usuario de prueba
  const user = {
    name: 'Carlos Test',
    email: 'carlos.test@empresa.com',
    role: 'Editor',
    group: 'IT'
  };
  
  // Ejecutar función de procesamiento
  // NOTA: Falta parámetro 'row', causará error
  processNewUser(user);
  
  // Loggear finalización
  Logger.log('Prueba completada');
  Logger.log('Revisar: Email recibido + Evento en Calendar');
}

/**
 * PRUEBA 2: Simula usuario inactivo.
 * 
 * DESCRIPCIÓN:
 * Función de prueba que ejecuta processInactiveUser() con datos de prueba.
 * Verifica que el sistema registre correctamente usuarios inactivos y
 * envíe email de alerta.
 * 
 * EJECUCIÓN:
 * Manual desde editor Apps Script:
 * 1. Seleccionar "pruebaUsuarioInactivo" en dropdown
 * 2. Click en botón "Ejecutar"
 * 3. Revisar logs
 * 4. Verificar email de alerta recibido
 * 
 * DATOS DE PRUEBA:
 * - Nombre: Luis Pérez
 * - Email: luis.perez@empresa.com
 * - Rol: Viewer
 * - Grupo: RH
 * 
 * RESULTADOS ESPERADOS:
 * - Log de prueba completada
 * - Email de alerta en bandeja
 * - Nueva fila en RegistroDeEventos con tipo USUARIO_INACTIVO
 * 
 * CONTENIDO DEL EMAIL:
 * - Asunto: Usuario inactivo detectado
 * - Cuerpo: Información del usuario y grupo
 * - Fecha de la acción
 */
function pruebaUsuarioInactivo() {
  Logger.log('=== PRUEBA: Usuario Inactivo ===');
  
  // Crear objeto de usuario de prueba
  const user = {
    name: 'Luis Pérez',
    email: 'luis.perez@empresa.com',
    role: 'Viewer',
    group: 'RH'
  };
  
  // Ejecutar función de procesamiento
  processInactiveUser(user);
  
  // Loggear finalización
  Logger.log('Prueba completada');
  Logger.log('Revisar: Email de alerta recibido');
}

/**
 * PRUEBA 3: Simula cambio de rol a Admin.
 * 
 * DESCRIPCIÓN:
 * Función de prueba que ejecuta notifyRoleChange() con rol Admin.
 * Verifica que el sistema detecte cambios críticos de seguridad y
 * envíe email de alerta apropiado.
 * 
 * EJECUCIÓN:
 * Manual desde editor Apps Script:
 * 1. Seleccionar "pruebaCambioRol" en dropdown
 * 2. Click en botón "Ejecutar"
 * 3. Revisar logs
 * 4. Verificar email de alerta Admin recibido
 * 
 * DATOS DE PRUEBA:
 * - Nombre: Ana López
 * - Email: ana.lopez@empresa.com
 * - Rol: Admin (cambio crítico)
 * - Grupo: Finanzas
 * 
 * RESULTADOS ESPERADOS:
 * - Log de prueba completada
 * - Email de alerta de seguridad en bandeja
 * - Nueva fila en RegistroDeEventos con status WARNING
 * 
 * CONTENIDO DEL EMAIL:
 * - Asunto: Cambio Crítico: Nuevo Administrador
 * - Cuerpo: Advertencia y datos del usuario
 * - Solicitud de verificación de autorización
 * 
 * VARIANTE:
 * Para probar con rol no-Admin (sin email), cambiar role a "Editor"
 */
function pruebaCambioRol() {
  Logger.log('=== PRUEBA: Cambio de Rol ===');
  
  // Crear objeto de usuario con rol Admin
  const user = {
    name: 'Ana López',
    email: 'ana.lopez@empresa.com',
    role: 'Admin',  // Cambio crítico que dispara alerta
    group: 'Finanzas'
  };
  
  // Ejecutar función de notificación
  notifyRoleChange(user);
  
  // Loggear finalización
  Logger.log('Prueba completada');
  Logger.log('Revisar: Email de alerta Admin recibido');
}

/**
 * PRUEBA 4: Ejecuta verificación diaria manualmente.
 * 
 * DESCRIPCIÓN:
 * Función de prueba que ejecuta verificarUsuariosInactivos() manualmente.
 * Permite probar el trigger diario sin esperar a que se ejecute automáticamente.
 * 
 * EJECUCIÓN:
 * Manual desde editor Apps Script:
 * 1. Seleccionar "pruebaVerificacionDiaria" en dropdown
 * 2. Click en botón "Ejecutar"
 * 3. Revisar logs del sistema
 * 4. Verificar cambios en hoja Usuarios
 * 
 * RESULTADOS ESPERADOS:
 * - Logs con lista completa de usuarios
 * - Usuarios con más de 7 días sin acceso desactivados
 * - Email de reporte si hay inactivos
 * - Actualización de columna Activo en hoja
 * 
 * REQUISITOS PREVIOS:
 * Para que la prueba tenga efecto:
 * - Debe haber usuarios activos en la hoja
 * - Al menos uno debe tener lastAccess > 7 días atrás
 * 
 * DEBUGGING:
 * Revisar logs para ver:
 * - Cantidad total de usuarios
 * - Días sin acceso de cada usuario
 * - Usuarios que fueron desactivados
 */
function pruebaVerificacionDiaria() {
  Logger.log('=== PRUEBA: Verificación Diaria ===');
  
  // Ejecutar verificación completa
  verificarUsuariosInactivos();
  
  // Loggear finalización
  Logger.log('Prueba completada');
  Logger.log('Revisar: Logs del sistema');
}

/**
 * PRUEBA COMPLETA: Simula agregar nuevo usuario con todas las validaciones.
 * 
 * DESCRIPCIÓN:
 * Función de prueba más robusta que simula el flujo completo de agregar
 * un nuevo usuario. Escribe datos en la hoja, luego ejecuta processNewUser()
 * con el número de fila correcto.
 * 
 * DIFERENCIAS CON pruebaAgregarUsuario():
 * - Escribe datos reales en la hoja
 * - Busca primera fila vacía automáticamente
 * - Pasa parámetro 'row' a processNewUser()
 * - Actualiza fechas correctamente
 * 
 * FUNCIONAMIENTO INTERNO:
 * 
 * PASO 1: Preparación
 * - Obtener referencia a hoja Usuarios
 * - Definir datos de prueba
 * 
 * PASO 2: Buscar fila vacía
 * - Leer todos los datos de la hoja
 * - Iterar para encontrar primera fila sin nombre
 * - Si no hay vacías, usar siguiente después de última
 * 
 * PASO 3: Escribir datos
 * - Llenar columnas A, B, C, D, E
 * - Dejar F y G vacías (serán llenadas por processNewUser)
 * 
 * PASO 4: Esperar y procesar
 * - sleep(1000) da tiempo para que Sheets procese
 * - Ejecutar processNewUser con usuario y fila
 * 
 * PASO 5: Validación
 * - Loggear pasos completados
 * - Listar verificaciones esperadas
 * 
 * EJECUCIÓN:
 * Manual desde editor Apps Script:
 * 1. Seleccionar "pruebaCompletaNuevoUsuario" en dropdown
 * 2. Click en botón "Ejecutar"
 * 3. Revisar logs
 * 4. Verificar hoja Usuarios tiene nueva fila con fechas
 * 5. Verificar email HTML recibido
 * 6. Verificar evento en Calendar
 * 7. Verificar RegistroDeEventos tiene nuevo evento
 * 
 * DATOS DE PRUEBA:
 * - Nombre: Pedro Ramírez
 * - Email: pedro.ramirez@empresa.com
 * - Rol: Viewer
 * - Grupo: RH
 * - Activo: true (escrito como "TRUE")
 * 
 * BÚSQUEDA DE FILA VACÍA:
 * - Itera desde índice 1 (saltar encabezados)
 * - Busca primera fila donde columna A está vacía
 * - Si no encuentra, usa length+1 (siguiente fila)
 * 
 * ESCRITURA DE DATOS:
 * - setValue() escribe cada celda individualmente
 * - Podría optimizarse con setValues() batch
 * - Se deja así para claridad en prueba
 * 
 * UTILITIES.SLEEP():
 * - Pausa ejecución por 1000 ms (1 segundo)
 * - Da tiempo a Sheets para procesar escrituras
 * - Evita race conditions
 * 
 * MANEJO DE ERRORES:
 * - Try-catch envuelve processNewUser
 * - Si falla, loggea error y stack trace
 * - No interrumpe escritura de datos
 * 
 * VERIFICACIONES ESPERADAS:
 * 1. Fechas auto-llenadas en columnas F y G
 * 2. Email HTML recibido en bandeja
 * 3. Evento de onboarding en Calendar (mañana 10 AM)
 * 4. Nueva fila en RegistroDeEventos
 * 
 * LIMITACIÓN:
 * - Cada ejecución agrega nueva fila
 * - No verifica si usuario ya existe
 * - Limpiar manualmente después de pruebas
 * 
 * VENTAJA:
 * Esta es la prueba más cercana al flujo real de usuario
 * 
 * DECISIONES DE DISEÑO:
 * - Escribir datos reales para simular trigger real
 * - Buscar fila vacía automáticamente
 * - Incluir sleep para estabilidad
 * - Logging extensivo para debugging
 * - Lista de verificaciones para QA
 */
function pruebaCompletaNuevoUsuario() {
  Logger.log('=== PRUEBA COMPLETA: NUEVO USUARIO ===');
  
  // PASO 1: Preparación
  
  // Obtener referencia al spreadsheet y hoja
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Usuarios');
  
  // Definir datos del usuario de prueba
  const testUser = {
    name: 'Pedro Ramírez',
    email: 'pedro.ramirez@empresa.com',
    role: 'Viewer',
    group: 'RH',
    active: true
  };
  
  // PASO 2: Buscar primera fila vacía
  
  // Leer todos los datos de la hoja
  const data = sheet.getDataRange().getValues();
  
  // Variable para almacenar número de fila vacía
  let emptyRow = -1;
  
  // Buscar primera fila sin nombre (columna A vacía)
  for (let i = 1; i < data.length; i++) {
    // Si columna A está vacía o es string vacío
    if (!data[i][0] || data[i][0] === '') {
      // Guardar índice +1 (convertir índice array a número fila)
      emptyRow = i + 1;
      break;
    }
  }
  
  // Si no se encontró fila vacía, usar siguiente después de última
  if (emptyRow === -1) {
    emptyRow = data.length + 1;
  }
  
  // Loggear fila que se usará
  Logger.log(`Fila para prueba: ${emptyRow}`);
  
  // PASO 3: Escribir datos en la hoja
  
  // Columna A: Nombre
  sheet.getRange(emptyRow, 1).setValue(testUser.name);
  
  // Columna B: Email
  sheet.getRange(emptyRow, 2).setValue(testUser.email);
  
  // Columna C: Rol
  sheet.getRange(emptyRow, 3).setValue(testUser.role);
  
  // Columna D: Grupo
  sheet.getRange(emptyRow, 4).setValue(testUser.group);
  
  // Columna E: Activo (como string "TRUE")
  sheet.getRange(emptyRow, 5).setValue('TRUE');
  
  // Loggear datos escritos
  Logger.log('Datos escritos en la hoja');
  
  // PASO 4: Esperar y procesar
  
  // Pausar 1 segundo para que Sheets procese las escrituras
  Utilities.sleep(1000);
  
  // Ejecutar processNewUser con manejo de errores
  try {
    Logger.log('Ejecutando processNewUser...');
    
    // Llamar función de procesamiento con usuario y fila
    processNewUser(testUser, emptyRow);
    
    Logger.log('processNewUser ejecutado');
  } catch (error) {
    // Si hay error, loggear mensaje y stack trace
    Logger.log('Error en processNewUser: ' + error.message);
    Logger.log(error.stack);
  }
  
  // PASO 5: Logging de verificaciones
  
  Logger.log('=== VERIFICAR: ===');
  Logger.log('1. Fechas en columnas F y G');
  Logger.log('2. Email HTML recibido');
  Logger.log('3. Evento en Calendar (mañana 10 AM)');
  Logger.log('4. Fila en RegistroEventos');
  Logger.log('=== FIN DE PRUEBA ===');
}

/**
 * PRUEBA 5: Verifica configuración del sistema.
 * 
 * DESCRIPCIÓN:
 * Función de prueba que valida la configuración completa del sistema.
 * Lee configuración y usuarios, muestra estadísticas básicas.
 * 
 * UTILIDAD:
 * - Verificar que hojas existen
 * - Validar estructura de configuración
 * - Contar usuarios totales y activos
 * - Debugging de problemas de configuración
 * 
 * EJECUCIÓN:
 * Manual desde editor Apps Script:
 * 1. Seleccionar "pruebaConfiguracion" en dropdown
 * 2. Click en botón "Ejecutar"
 * 3. Revisar logs
 * 
 * INFORMACIÓN MOSTRADA:
 * 1. Configuración completa (formato JSON)
 * 2. Total de usuarios en el sistema
 * 3. Cantidad de usuarios activos
 * 
 * RESULTADOS ESPERADOS:
 * - Log con objeto config completo
 * - Cantidad de usuarios mayor a 0
 * - Cantidad de activos menor o igual a total
 * 
 * JSON.stringify():
 * - Primer parámetro: objeto a convertir
 * - Segundo parámetro: null (no reemplazar valores)
 * - Tercer parámetro: 2 (indentación de 2 espacios)
 * - Resultado: JSON legible formateado
 * 
 * EJEMPLO DE OUTPUT EN LOGS:
 * === PRUEBA: Configuración ===
 * Configuración actual:
 * {
 *   "emailNotificacion": "admin@empresa.com",
 *   "calendarioId": "primary",
 *   "notificarAdmins": "TRUE",
 *   "crearEventoCalendar": "TRUE"
 * }
 * Total usuarios: 5
 * Usuarios activos: 4
 * Prueba completada
 * 
 * CASOS DE ERROR:
 * - Si hoja Configuración no existe: getConfig() lanzará error
 * - Si hoja Usuarios no existe: getUsers() retornará array vacío
 * 
 * DECISIONES DE DISEÑO:
 * - Mostrar config completa para auditoría
 * - Filtrar usuarios activos para estadística rápida
 * - Formato JSON para legibilidad
 * - No validar valores, solo mostrar
 */
function pruebaConfiguracion() {
  Logger.log('=== PRUEBA: Configuración ===');
  
  // Obtener configuración del sistema
  const config = getConfig();
  
  // Loggear configuración en formato JSON legible
  Logger.log('Configuración actual:');
  Logger.log(JSON.stringify(config, null, 2));
  
  // Obtener todos los usuarios
  const users = getUsers();
  
  // Loggear estadísticas de usuarios
  Logger.log(`Total usuarios: ${users.length}`);
  Logger.log(`Usuarios activos: ${users.filter(u => u.active).length}`);
  
  // Marcar prueba como completada
  Logger.log('Prueba completada');
}
