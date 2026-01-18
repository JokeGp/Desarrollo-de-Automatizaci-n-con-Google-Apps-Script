/**
 * ============================================================================
 * WORKSPACE AUTOMATION - TURING IA
 * Archivo: Notifications.gs
 * ============================================================================
 * 
 * Funciones de notificaciones: Gmail y Google Calendar.
 * Gestiona el envío de emails (texto plano y HTML) y la creación de eventos
 * en Google Calendar. También incluye funciones de procesamiento para eventos
 * específicos del sistema (nuevos usuarios, usuarios inactivos, cambios de rol).
 * 
 * Autor: José Enrique Guerrero Pérez
 * Fecha: Enero 2026
 * ============================================================================
 */

// ============================================================================
// GMAIL - NOTIFICACIONES POR EMAIL
// ============================================================================

/**
 * Envía notificación por email con formato de texto plano.
 * 
 * DESCRIPCIÓN:
 * Función para enviar emails simples con texto plano. Valida que las
 * notificaciones estén habilitadas en la configuración y que el email
 * destino sea válido antes de enviar.
 * 
 * PARÁMETROS:
 * @param {string} subject - Asunto del email
 * @param {string} body - Cuerpo del mensaje en texto plano
 * 
 * VALIDACIONES REALIZADAS:
 * 1. Verifica que notificarAdmins esté en TRUE en configuración
 * 2. Valida formato de email con expresión regular
 * 3. Maneja errores sin interrumpir flujo del sistema
 * 
 * RETORNA:
 * Boolean - true si el email se envió exitosamente, false en caso contrario
 * 
 * CASOS QUE RETORNAN FALSE:
 * - Notificaciones deshabilitadas en configuración
 * - Email inválido (no cumple formato)
 * - Error al intentar enviar (capturado en catch)
 * 
 * FUNCIONAMIENTO INTERNO:
 * 1. Obtiene configuración del sistema
 * 2. Verifica si las notificaciones están habilitadas
 * 3. Valida formato del email con regex
 * 4. Envía email usando GmailApp
 * 5. Loggea resultado
 * 6. Retorna true/false según éxito
 * 
 * EXPRESIÓN REGULAR PARA VALIDACIÓN:
 * /^[^\s@]+@[^\s@]+\.[^\s@]+$/
 * - ^[^\s@]+ : Inicio, uno o más caracteres que no sean espacio ni @
 * - @ : Arroba obligatoria
 * - [^\s@]+ : Uno o más caracteres que no sean espacio ni @
 * - \. : Punto obligatorio
 * - [^\s@]+$ : Uno o más caracteres hasta el final
 * 
 * EJEMPLO DE USO:
 * const enviado = sendNotification(
 *   'Alerta del Sistema',
 *   'El usuario Juan fue desactivado por inactividad.'
 * );
 * if (enviado) {
 *   // Email enviado correctamente
 * }
 * 
 * LIMITACIONES:
 * - Gmail tiene límite diario de emails (100-500 según tipo de cuenta)
 * - No soporta formato HTML (usar sendHtmlNotification para eso)
 * - Email siempre se envía al admin configurado (no soporta múltiples destinos)
 * 
 * DECISIONES DE DISEÑO:
 * - Retornar boolean en lugar de lanzar error
 * - Validación de email para evitar errores silenciosos
 * - Logging extensivo para debugging
 * - Try-catch para manejo robusto de errores
 */
function sendNotification(subject, body) {
  try {
    // Obtener configuración del sistema
    const config = getConfig();
    
    // Verificar si las notificaciones están habilitadas
    // Soporta tanto boolean true como string "TRUE"
    if (config.notificarAdmins !== true && config.notificarAdmins !== 'TRUE') {
      Logger.log('Notificaciones deshabilitadas');
      return false;
    }
    
    // Validar formato del email con expresión regular
    // Patrón básico: algo@algo.algo
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(config.emailNotificacion)) {
      Logger.log('Email inválido');
      return false;
    }
    
    // Enviar email usando API de Gmail
    // Parámetros: destinatario, asunto, cuerpo
    GmailApp.sendEmail(config.emailNotificacion, subject, body);
    
    // Loggear éxito
    Logger.log('Email enviado: ' + subject);
    return true;
    
  } catch (error) {
    // Capturar cualquier error durante el envío
    // Loggear para debugging pero no interrumpir ejecución
    Logger.log('Error al enviar email: ' + error.message);
    return false;
  }
}

/**
 * Envía notificación con formato HTML.
 * 
 * DESCRIPCIÓN:
 * Función para enviar emails con formato HTML rico. Permite tablas, estilos,
 * colores y mejor presentación visual que el texto plano. Incluye las mismas
 * validaciones que sendNotification().
 * 
 * PARÁMETROS:
 * @param {string} subject - Asunto del email
 * @param {string} htmlBody - Cuerpo del mensaje en formato HTML
 * 
 * VALIDACIONES REALIZADAS:
 * 1. Verifica que notificarAdmins esté en TRUE
 * 2. No valida email aquí (asume que config es correcta)
 * 
 * RETORNA:
 * Boolean - true si el email se envió exitosamente, false en caso contrario
 * 
 * FORMATO HTML SOPORTADO:
 * - Etiquetas básicas: <h1>, <p>, <strong>, <em>
 * - Tablas: <table>, <tr>, <td>
 * - Estilos inline: style="color: red"
 * - Divisores: <hr>
 * - No soporta: JavaScript, CSS externo, imágenes externas sin URL
 * 
 * FUNCIONAMIENTO INTERNO:
 * 1. Obtiene configuración
 * 2. Verifica si notificaciones están habilitadas
 * 3. Envía email con opción htmlBody
 * 4. El tercer parámetro '' es el cuerpo de texto plano (fallback)
 * 5. Loggea resultado
 * 
 * API DE GMAIL UTILIZADA:
 * GmailApp.sendEmail(destinatario, asunto, cuerpoTexto, opciones)
 * Opciones:
 *   - htmlBody: Versión HTML del email
 *   - Gmail automáticamente usa HTML si el cliente lo soporta
 *   - Fallback a texto plano si no soporta HTML
 * 
 * EJEMPLO DE USO:
 * const html = `
 *   <h2>Alerta</h2>
 *   <p>El usuario <strong>Juan</strong> fue desactivado.</p>
 *   <table border="1">
 *     <tr><td>Fecha</td><td>18/01/2026</td></tr>
 *   </table>
 * `;
 * sendHtmlNotification('Alerta del Sistema', html);
 * 
 * MEJORES PRÁCTICAS HTML:
 * - Usar estilos inline (no CSS externo)
 * - Tablas para layout (más compatible que divs)
 * - Colores seguros para web
 * - Probar en múltiples clientes de email
 * 
 * LIMITACIONES:
 * - Mismo límite diario de Gmail que sendNotification
 * - Algunos clientes de email no renderizan HTML complejo
 * - Imágenes deben ser URLs absolutas
 * 
 * DECISIONES DE DISEÑO:
 * - Texto plano vacío como fallback (Gmail lo maneja)
 * - Validación mínima (asume HTML válido)
 * - Try-catch para robustez
 */
function sendHtmlNotification(subject, htmlBody) {
  try {
    // Obtener configuración del sistema
    const config = getConfig();
    
    // Verificar si las notificaciones están habilitadas
    if (config.notificarAdmins !== true && config.notificarAdmins !== 'TRUE') {
      return false;
    }
    
    // Enviar email con formato HTML
    // Parámetros: destinatario, asunto, cuerpoTexto (vacío), opciones
    // Opciones contiene htmlBody para versión HTML
    GmailApp.sendEmail(config.emailNotificacion, subject, '', {
      htmlBody: htmlBody
    });
    
    // Loggear éxito
    Logger.log('Email HTML enviado: ' + subject);
    return true;
    
  } catch (error) {
    // Capturar errores sin interrumpir ejecución
    Logger.log('Error al enviar email HTML: ' + error.message);
    return false;
  }
}

// ============================================================================
// GOOGLE CALENDAR - EVENTOS AUTOMÁTICOS
// ============================================================================

/**
 * Crea un evento en Google Calendar.
 * 
 * DESCRIPCIÓN:
 * Función para crear eventos automáticamente en Google Calendar. Valida que
 * la creación de eventos esté habilitada y que el calendario exista antes
 * de intentar crear el evento.
 * 
 * PARÁMETROS:
 * @param {string} title - Título del evento
 * @param {string} description - Descripción detallada del evento
 * @param {Date} startTime - Fecha y hora de inicio
 * @param {Date} endTime - Fecha y hora de fin
 * 
 * VALIDACIONES REALIZADAS:
 * 1. Verifica que crearEventoCalendar esté en TRUE
 * 2. Valida que el calendario especificado exista
 * 3. Maneja errores sin interrumpir sistema
 * 
 * RETORNA:
 * Boolean - true si el evento se creó exitosamente, false en caso contrario
 * 
 * FUNCIONAMIENTO INTERNO:
 * 1. Obtiene configuración (calendarioId)
 * 2. Verifica si creación de eventos está habilitada
 * 3. Obtiene referencia al calendario por ID
 * 4. Valida que el calendario existe
 * 5. Crea evento con título, horarios y opciones
 * 6. Loggea resultado
 * 
 * OPCIONES DEL EVENTO:
 * - description: Descripción completa del evento
 * - sendInvites: false (no envía invitaciones automáticas)
 *   Nota: Se puede cambiar a true para enviar invites
 * 
 * ID DE CALENDARIO:
 * - "primary": Calendario principal del usuario
 * - Email del calendario: "usuario@gmail.com"
 * - ID específico: Obtenido de configuración del calendario
 * 
 * EJEMPLO DE USO:
 * const inicio = new Date('2026-01-19T10:00:00');
 * const fin = new Date('2026-01-19T11:00:00');
 * createCalendarEvent(
 *   'Reunión con IT',
 *   'Discusión sobre nuevo sistema',
 *   inicio,
 *   fin
 * );
 * 
 * LIMITACIONES:
 * - Requiere permisos de Calendar (autorizados en trigger instalable)
 * - Calendario debe existir y ser accesible
 * - Fechas deben ser objetos Date válidos
 * 
 * MEJORAS FUTURAS POSIBLES:
 * - Agregar recordatorios (addEmailReminder, addPopupReminder)
 * - Soportar invitados (guests)
 * - Especificar ubicación (location)
 * - Eventos recurrentes (recurrence)
 * 
 * DECISIONES DE DISEÑO:
 * - sendInvites: false por defecto (evita spam)
 * - Validación de calendario antes de crear
 * - Retornar boolean para manejo consistente
 * - Try-catch para robustez
 */
function createCalendarEvent(title, description, startTime, endTime) {
  try {
    // Obtener configuración del sistema
    const config = getConfig();
    
    // Verificar si la creación de eventos está habilitada
    if (config.crearEventoCalendar !== true && config.crearEventoCalendar !== 'TRUE') {
      Logger.log('Creación de eventos deshabilitada');
      return false;
    }
    
    // Obtener referencia al calendario por ID
    // calendarioId puede ser "primary" o un email específico
    const calendar = CalendarApp.getCalendarById(config.calendarioId);
    
    // Validar que el calendario existe y es accesible
    if (!calendar) {
      Logger.log('Calendario no encontrado');
      return false;
    }
    
    // Crear evento en el calendario
    // Parámetros: título, inicio, fin, opciones
    calendar.createEvent(title, startTime, endTime, {
      description: description,
      sendInvites: false  // No enviar invitaciones automáticas
    });
    
    // Loggear éxito
    Logger.log('Evento creado: ' + title);
    return true;
    
  } catch (error) {
    // Capturar errores sin interrumpir ejecución
    Logger.log('Error al crear evento: ' + error.message);
    return false;
  }
}

// ============================================================================
// PROCESAMIENTO DE EVENTOS
// ============================================================================

/**
 * Procesa usuario inactivo: registra evento y envía email de alerta.
 * 
 * DESCRIPCIÓN:
 * Función llamada cuando un usuario es marcado como inactivo. Registra el
 * evento en la hoja de auditoría y envía notificación por email al administrador.
 * 
 * PARÁMETROS:
 * @param {Object} user - Objeto con datos del usuario
 *   @param {string} user.name - Nombre del usuario
 *   @param {string} user.email - Email del usuario
 *   @param {string} user.group - Grupo al que pertenece
 * 
 * ACCIONES REALIZADAS:
 * 1. Registra evento en RegistroDeEventos con tipo USUARIO_INACTIVO
 * 2. Construye email de notificación con detalles del usuario
 * 3. Envía email al administrador
 * 4. Loggea resultado en consola
 * 
 * CONTENIDO DEL EMAIL:
 * - Nombre y email del usuario afectado
 * - Grupo al que pertenece
 * - Fecha y hora del evento
 * - Firma del sistema
 * 
 * FUNCIONAMIENTO INTERNO:
 * 1. Llama a logEvent() para auditoría
 * 2. Construye string de email con template literals
 * 3. Envía email usando sendNotification()
 * 4. Loggea procesamiento completo
 * 
 * EJEMPLO DE EMAIL GENERADO:
 * Asunto: "Usuario inactivo detectado"
 * Cuerpo:
 *   El usuario Juan Pérez (juan@empresa.com) del grupo IT
 *   fue marcado como inactivo.
 *   
 *   Fecha: 18/1/2026 15:30:45
 *   
 *   Sistema de Gestión Workspace - Turing IA
 * 
 * CONSIDERACIONES:
 * - Esta función se llama desde handleUserEdit() cuando detecta cambio a inactivo
 * - También puede llamarse desde verificarUsuariosInactivos() en trigger diario
 * - El email se envía independientemente del resultado del registro
 * 
 * DECISIONES DE DISEÑO:
 * - Registro primero, notificación después
 * - Email en texto plano para compatibilidad
 * - Incluir todos los datos relevantes del usuario
 * - Formato de fecha localizado a es-MX
 */
function processInactiveUser(user) {
  // Registrar evento en hoja de auditoría
  logEvent({
    type: 'USUARIO_INACTIVO',
    user: user.name,
    details: 'Usuario desactivado',
    status: 'ALERTA',
    action: 'Email enviado'
  });
  
  // Definir asunto del email
  const subject = 'Usuario inactivo detectado';
  
  // Construir cuerpo del email usando template literals
  // Incluye: nombre, email, grupo, fecha formateada
  const body = `El usuario ${user.name} (${user.email}) del grupo ${user.group} fue marcado como inactivo.

Fecha: ${new Date().toLocaleString('es-MX')}

Sistema de Gestión Workspace - Turing IA`;
  
  // Enviar notificación por email
  sendNotification(subject, body);
  
  // Loggear procesamiento completo
  Logger.log(`Usuario inactivo procesado: ${user.name}`);
}

/**
 * Notifica cambio de rol, con alerta especial si es Admin.
 * 
 * DESCRIPCIÓN:
 * Función llamada cuando se cambia el rol de un usuario. Registra el evento
 * y envía email solo si el nuevo rol es Admin (cambio crítico de seguridad).
 * 
 * PARÁMETROS:
 * @param {Object} user - Objeto con datos del usuario
 *   @param {string} user.name - Nombre del usuario
 *   @param {string} user.email - Email del usuario
 *   @param {string} user.role - Nuevo rol asignado
 *   @param {string} user.group - Grupo al que pertenece
 * 
 * ACCIONES REALIZADAS:
 * 1. Siempre registra evento en RegistroDeEventos
 * 2. Si rol es Admin, envía email de alerta de seguridad
 * 3. Si rol es otro (Editor/Viewer), solo registra sin email
 * 
 * LÓGICA CONDICIONAL:
 * - rol === 'Admin' → status: WARNING, email enviado
 * - rol !== 'Admin' → status: OK, solo registro
 * 
 * CONTENIDO DEL EMAIL (solo para Admin):
 * - Advertencia de cambio crítico
 * - Nombre, email y grupo del usuario
 * - Fecha del cambio
 * - Solicitud de verificación de autorización
 * 
 * FUNCIONAMIENTO INTERNO:
 * 1. Registra evento con status condicional
 * 2. Verifica si el nuevo rol es Admin
 * 3. Si es Admin, construye y envía email de alerta
 * 4. Si no es Admin, termina después del registro
 * 
 * EJEMPLO DE EMAIL PARA ADMIN:
 * Asunto: "Cambio Crítico: Nuevo Administrador"
 * Cuerpo:
 *   ATENCIÓN: Juan Pérez fue promovido a Admin.
 *   
 *   Email: juan@empresa.com
 *   Grupo: IT
 *   Fecha: 18/1/2026 15:30:45
 *   
 *   Verifica que este cambio esté autorizado.
 *   
 *   Sistema de Gestión Workspace - Turing IA
 * 
 * CONSIDERACIONES DE SEGURIDAD:
 * - Rol Admin tiene permisos completos del sistema
 * - Cambio a Admin requiere notificación inmediata
 * - Email pide verificación de autorización
 * - Registro de auditoría con WARNING para revisión
 * 
 * DECISIONES DE DISEÑO:
 * - Solo notificar cambios críticos (Admin)
 * - Siempre registrar en auditoría
 * - Email con tono de advertencia
 * - Solicitar verificación explícita
 */
function notifyRoleChange(user) {
  // Registrar evento en hoja de auditoría
  // Status condicional: WARNING si es Admin, OK si no
  logEvent({
    type: 'ROL_MODIFICADO',
    user: user.name,
    details: `Nuevo rol: ${user.role}`,
    status: user.role === 'Admin' ? 'WARNING' : 'OK',
    action: user.role === 'Admin' ? 'Email enviado' : 'Solo registro'
  });
  
  // Enviar email SOLO si el nuevo rol es Admin (cambio crítico)
  if (user.role === 'Admin') {
    // Definir asunto con advertencia
    const subject = 'Cambio Crítico: Nuevo Administrador';
    
    // Construir cuerpo del email con advertencia de seguridad
    const body = `ATENCIÓN: ${user.name} fue promovido a Admin.

Email: ${user.email}
Grupo: ${user.group}
Fecha: ${new Date().toLocaleString('es-MX')}

Verifica que este cambio esté autorizado.

Sistema de Gestión Workspace - Turing IA`;
    
    // Enviar notificación de alerta
    sendNotification(subject, body);
  }
}

/**
 * Procesa nuevo usuario: actualiza fechas, envía email HTML y crea evento Calendar.
 * 
 * DESCRIPCIÓN:
 * Función completa que se ejecuta cuando se detecta un nuevo usuario en el sistema.
 * Realiza 4 acciones principales: actualizar fechas, registrar evento, enviar email
 * de bienvenida y crear evento de onboarding en el calendario.
 * 
 * PARÁMETROS:
 * @param {Object} user - Objeto con datos del usuario
 *   @param {string} user.name - Nombre del usuario
 *   @param {string} user.email - Email del usuario
 *   @param {string} user.role - Rol asignado
 *   @param {string} user.group - Grupo asignado
 * @param {number} row - Número de fila del usuario en la hoja (base 1)
 * 
 * ACCIONES REALIZADAS EN ORDEN:
 * 1. Actualiza columna F (dateRegistered) con fecha actual
 * 2. Actualiza columna G (lastAccess) con fecha actual
 * 3. Registra evento en RegistroDeEventos
 * 4. Envía email HTML con tabla de datos del usuario
 * 5. Crea evento de onboarding en Calendar para mañana 10 AM
 * 
 * ACTUALIZACIÓN DE FECHAS:
 * - Usa getRange(row, col) para acceder a celdas específicas
 * - setValue() escribe la fecha actual
 * - Columna 6 = dateRegistered (fecha de registro)
 * - Columna 7 = lastAccess (último acceso)
 * 
 * EMAIL HTML GENERADO:
 * - Tabla HTML con borde y padding
 * - Filas con datos: Nombre, Email, Rol, Grupo, Fecha
 * - Footer con identificación del sistema
 * - Estilos inline para compatibilidad
 * 
 * EVENTO DE CALENDAR:
 * - Programado para el día siguiente a las 10:00 AM
 * - Duración: 1 hora (10:00 - 11:00)
 * - Descripción incluye: datos del usuario y agenda de onboarding
 * - Agenda detallada con 4 puntos y tiempos estimados
 * 
 * AGENDA DEL ONBOARDING:
 * 1. Bienvenida al equipo
 * 2. Explicación de roles y permisos
 * 3. Tour por herramientas
 * 4. Asignación de tareas
 * 
 * FUNCIONAMIENTO INTERNO:
 * 1. Obtiene referencia a hoja Usuarios
 * 2. Crea objeto Date con fecha/hora actual
 * 3. Actualiza celdas F y G de la fila
 * 4. Registra evento en auditoría
 * 5. Construye HTML para email
 * 6. Envía email con sendHtmlNotification()
 * 7. Calcula fecha de mañana a las 10 AM
 * 8. Crea evento con createCalendarEvent()
 * 9. Loggea procesamiento completo
 * 
 * CÁLCULO DE FECHA PARA EVENTO:
 * - tomorrow = new Date() : Crea objeto con fecha actual
 * - setDate(getDate() + 1) : Suma 1 día
 * - setHours(10, 0, 0, 0) : Establece a 10:00:00.000 AM
 * - endTime = copia de tomorrow con hora 11:00:00.000 AM
 * 
 * EJEMPLO DE HTML GENERADO:
 * <h2>Nuevo Usuario Registrado</h2>
 * <table border="1" cellpadding="8" style="border-collapse: collapse;">
 *   <tr><td><strong>Nombre:</strong></td><td>Juan Pérez</td></tr>
 *   <tr><td><strong>Email:</strong></td><td>juan@empresa.com</td></tr>
 *   ...
 * </table>
 * <p style="color: #666;">Sistema de Gestión Workspace - Turing IA</p>
 * 
 * CONSIDERACIONES:
 * - Esta función se llama desde handleUserEdit() cuando detecta fila completa
 * - También se usa en funciones de prueba
 * - Requiere que parámetro 'row' sea correcto para actualizar fechas
 * - Email y evento son opcionales (dependen de config)
 * 
 * DEPENDENCIAS:
 * - getConfig(): Para validar si enviar email/crear evento
 * - logEvent(): Para registrar en auditoría
 * - sendHtmlNotification(): Para enviar email
 * - createCalendarEvent(): Para crear evento
 * 
 * DECISIONES DE DISEÑO:
 * - Actualizar fechas primero (datos críticos)
 * - Email HTML para mejor presentación
 * - Evento al día siguiente (dar tiempo de preparación)
 * - Hora laboral estándar (10 AM)
 * - Duración razonable (1 hora)
 */
function processNewUser(user, row) {
  // PASO 1: Actualizar fechas en la hoja de usuarios
  // Obtener referencia al spreadsheet y hoja Usuarios
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Usuarios');
  
  // Crear objeto Date con fecha y hora actual
  const ahora = new Date();
  
  // Actualizar columna F (índice 6): Fecha de Registro
  // getRange(fila, columna) donde columnas empiezan en 1
  sheet.getRange(row, 6).setValue(ahora);
  
  // Actualizar columna G (índice 7): Último Acceso
  sheet.getRange(row, 7).setValue(ahora);
  
  // PASO 2: Registrar evento en hoja de auditoría
  logEvent({
    type: 'USUARIO_AGREGADO',
    user: user.name,
    details: `Rol: ${user.role}, Grupo: ${user.group}`,
    status: 'OK',
    action: 'Email y evento creados'
  });
  
  // PASO 3: Construir y enviar email HTML de bienvenida
  // Template literal con HTML estructurado
  const htmlBody = `
    <h2>Nuevo Usuario Registrado</h2>
    <table border="1" cellpadding="8" style="border-collapse: collapse;">
      <tr><td><strong>Nombre:</strong></td><td>${user.name}</td></tr>
      <tr><td><strong>Email:</strong></td><td>${user.email}</td></tr>
      <tr><td><strong>Rol:</strong></td><td>${user.role}</td></tr>
      <tr><td><strong>Grupo:</strong></td><td>${user.group}</td></tr>
      <tr><td><strong>Fecha:</strong></td><td>${new Date().toLocaleString('es-MX')}</td></tr>
    </table>
    <p style="color: #666; margin-top: 20px;">Sistema de Gestión Workspace - Turing IA</p>
  `;
  
  // Enviar email HTML
  sendHtmlNotification('Nuevo usuario agregado', htmlBody);
  
  // PASO 4: Crear evento de onboarding en Calendar
  // Calcular fecha de mañana a las 10 AM
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);  // Sumar 1 día
  tomorrow.setHours(10, 0, 0, 0);  // Establecer a 10:00:00.000
  
  // Calcular hora de fin (11 AM = 10 AM + 1 hora)
  const endTime = new Date(tomorrow);
  endTime.setHours(11, 0, 0, 0);  // Establecer a 11:00:00.000
  
  // Construir descripción detallada del evento
  const description = `Sesión de Onboarding

Usuario: ${user.name}
Email: ${user.email}
Rol: ${user.role}
Grupo: ${user.group}

Agenda:
1. Bienvenida al equipo
2. Explicación de roles y permisos
3. Tour por herramientas
4. Asignación de tareas

Sistema de Gestión Workspace - Turing IA`;
  
  // Crear evento en Calendar
  createCalendarEvent(
    `Onboarding: ${user.name}`,
    description,
    tomorrow,
    endTime
  );
  
  // Loggear procesamiento completo
  Logger.log(`Nuevo usuario procesado: ${user.name}`);
}

/**
 * Envía reporte diario de usuarios inactivos desactivados.
 * 
 * DESCRIPCIÓN:
 * Función llamada desde verificarUsuariosInactivos() cuando se detectan uno
 * o más usuarios inactivos. Genera y envía un reporte HTML con tabla detallada.
 * 
 * PARÁMETROS:
 * @param {Array<Object>} usuariosInactivos - Array de objetos con usuarios desactivados
 *   Cada objeto contiene:
 *     @param {string} name - Nombre del usuario
 *     @param {string} group - Grupo del usuario
 *     @param {number} days - Días de inactividad
 * 
 * CONTENIDO DEL REPORTE:
 * - Título con cantidad de usuarios afectados
 * - Tabla HTML con columnas: Usuario, Grupo, Días Inactivo
 * - Cada fila representa un usuario desactivado
 * - Footer con fecha y firma del sistema
 * 
 * FUNCIONAMIENTO INTERNO:
 * 1. Inicializa string vacío para la tabla
 * 2. Itera sobre array de usuarios inactivos
 * 3. Construye fila HTML para cada usuario
 * 4. Ensambla HTML completo con encabezado y tabla
 * 5. Envía email con sendHtmlNotification()
 * 
 * ESTRUCTURA DE LA TABLA HTML:
 * - Encabezado azul con texto blanco
 * - Bordes y padding para legibilidad
 * - Ancho 100% para usar espacio disponible
 * - Estilos inline para compatibilidad con clientes de email
 * 
 * EJEMPLO DE HTML GENERADO:
 * <h2>Reporte Diario - Usuarios Inactivos</h2>
 * <p>Se detectaron <strong>2</strong> usuarios sin actividad:</p>
 * <table>
 *   <thead>
 *     <tr><th>Usuario</th><th>Grupo</th><th>Días Inactivo</th></tr>
 *   </thead>
 *   <tbody>
 *     <tr><td>Juan Pérez</td><td>IT</td><td>15 días</td></tr>
 *     <tr><td>Ana López</td><td>RH</td><td>10 días</td></tr>
 *   </tbody>
 * </table>
 * 
 * ASUNTO DEL EMAIL:
 * - Incluye la fecha actual del reporte
 * - Formato: "Reporte Diario - 18/1/2026"
 * 
 * CONSIDERACIONES:
 * - Solo se llama si hay usuarios inactivos (length > 0)
 * - Reporte es informativo (acciones ya fueron tomadas)
 * - Permite al admin revisar desactivaciones automáticas
 * 
 * DECISIONES DE DISEÑO:
 * - Tabla HTML para mejor presentación que texto plano
 * - Incluir cantidad de afectados en resumen
 * - Fecha en formato localizado es-MX
 * - Color azul corporativo Google (#4285f4)
 */
function enviarReporteInactivos(usuariosInactivos) {
  // Inicializar string para construir filas de la tabla
  let tabla = '';
  
  // Iterar sobre cada usuario inactivo
  usuariosInactivos.forEach(u => {
    // Construir fila HTML con datos del usuario
    // Estilos inline: padding y border para cada celda
    tabla += `<tr>
      <td style="padding: 8px; border: 1px solid #ddd;">${u.name}</td>
      <td style="padding: 8px; border: 1px solid #ddd;">${u.group}</td>
      <td style="padding: 8px; border: 1px solid #ddd;">${u.days} días</td>
    </tr>`;
  });
  
  // Construir HTML completo del email
  const htmlBody = `
    <h2>Reporte Diario - Usuarios Inactivos</h2>
    <p>Se detectaron <strong>${usuariosInactivos.length}</strong> usuarios sin actividad:</p>
    <table border="1" cellpadding="8" style="border-collapse: collapse; width: 100%;">
      <thead>
        <tr style="background-color: #4285f4; color: white;">
          <th>Usuario</th><th>Grupo</th><th>Días Inactivo</th>
        </tr>
      </thead>
      <tbody>${tabla}</tbody>
    </table>
    <p style="margin-top: 20px; color: #666;">
      Fecha: ${new Date().toLocaleDateString('es-MX')}<br>
      Sistema de Gestión Workspace - Turing IA
    </p>
  `;
  
  // Enviar email con reporte
  // Asunto incluye fecha del reporte
  sendHtmlNotification(
    `Reporte Diario - ${new Date().toLocaleDateString('es-MX')}`,
    htmlBody
  );
}

/**
 * Función de prueba para verificar funcionalidad de Notifications.gs.
 * 
 * DESCRIPCIÓN:
 * Ejecuta pruebas de las tres funciones principales de notificaciones:
 * sendNotification(), sendHtmlNotification() y createCalendarEvent().
 * Útil para verificar que el sistema de notificaciones está correctamente
 * configurado y que los permisos están otorgados.
 * 
 * PRUEBAS REALIZADAS:
 * 1. Prueba sendNotification() con email de texto plano
 * 2. Prueba sendHtmlNotification() con email HTML
 * 3. Prueba createCalendarEvent() con evento para mañana
 * 
 * USO:
 * Ejecutar manualmente desde el editor de Apps Script:
 * 1. Seleccionar función "probarNotifications" en dropdown
 * 2. Click en botón "Ejecutar"
 * 3. Revisar logs (Ver > Registros)
 * 4. Verificar email recibido
 * 5. Verificar evento en Calendar
 * 
 * RESULTADOS ESPERADOS:
 * - Logs muestran éxito en todas las pruebas
 * - Dos emails recibidos en bandeja
 * - Un evento creado en Calendar para mañana 3 PM
 * 
 * MANEJO DE ERRORES:
 * - Cada prueba tiene su propio try-catch
 * - Si una prueba falla, las demás continúan
 * - Errores se loggean con descripción clara
 * 
 * EVENTO DE PRUEBA:
 * - Programado para mañana a las 15:00 (3 PM)
 * - Duración: 1 hora (15:00 - 16:00)
 * - Título: "Prueba de Evento"
 * - Descripción simple para identificar como prueba
 */
function probarNotifications() {
  Logger.log('=== PROBANDO NOTIFICATIONS.GS ===');
  
  // Prueba 1: Enviar email de texto plano
  try {
    const resultado = sendNotification(
      'Prueba de Email',
      'Este es un email de prueba desde Apps Script.'
    );
    if (resultado) {
      Logger.log('sendNotification() funciona - Revisa tu email');
    } else {
      Logger.log('sendNotification() no envió (revisa config)');
    }
  } catch (error) {
    Logger.log('sendNotification() falló: ' + error.message);
  }
  
  // Prueba 2: Enviar email HTML
  try {
    const resultado = sendHtmlNotification(
      'Prueba de Email HTML',
      '<h2>Prueba HTML</h2><p>Este es un email con <strong>formato</strong>.</p>'
    );
    if (resultado) {
      Logger.log('sendHtmlNotification() funciona - Revisa tu email');
    } else {
      Logger.log('sendHtmlNotification() no envió');
    }
  } catch (error) {
    Logger.log('sendHtmlNotification() falló: ' + error.message);
  }
  
  // Prueba 3: Crear evento en Calendar
  try {
    // Calcular fecha de mañana a las 15:00 (3 PM)
    const manana = new Date();
    manana.setDate(manana.getDate() + 1);
    manana.setHours(15, 0, 0, 0);
    
    // Calcular hora de fin (16:00 = 4 PM)
    const fin = new Date(manana);
    fin.setHours(16, 0, 0, 0);
    
    const resultado = createCalendarEvent(
      'Prueba de Evento',
      'Este es un evento de prueba creado por Apps Script.',
      manana,
      fin
    );
    
    if (resultado) {
      Logger.log('createCalendarEvent() funciona - Revisa tu Calendar');
    } else {
      Logger.log('createCalendarEvent() no creó evento');
    }
  } catch (error) {
    Logger.log('createCalendarEvent() falló: ' + error.message);
  }
  
  Logger.log('=== PRUEBA NOTIFICATIONS.GS COMPLETADA ===');
}
