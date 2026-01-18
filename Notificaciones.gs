/**
 * ============================================================================
 * WORKSPACE AUTOMATION - TURING IA
 * Archivo: Notifications.gs
 * ============================================================================
 * 
 * Funciones de notificaciones: Gmail y Google Calendar.
 * 
 * Autor: José Enrique Guerrero Pérez
 * Fecha: Enero 2026
 * ============================================================================
 */

// ============================================================================
// GMAIL - NOTIFICACIONES POR EMAIL
// ============================================================================

/**
 * Envía notificación por email (texto plano).
 */
function sendNotification(subject, body) {
  try {
    const config = getConfig();
    
    if (config.notificarAdmins !== true && config.notificarAdmins !== 'TRUE') {
      Logger.log('Notificaciones deshabilitadas');
      return false;
    }
    
    // Validar email
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(config.emailNotificacion)) {
      Logger.log('Email inválido');
      return false;
    }
    
    GmailApp.sendEmail(config.emailNotificacion, subject, body);
    Logger.log('Email enviado: ' + subject);
    return true;
    
  } catch (error) {
    Logger.log('Error al enviar email: ' + error.message);
    return false;
  }
}

/**
 * Envía notificación con formato HTML.
 */
function sendHtmlNotification(subject, htmlBody) {
  try {
    const config = getConfig();
    
    if (config.notificarAdmins !== true && config.notificarAdmins !== 'TRUE') {
      return false;
    }
    
    GmailApp.sendEmail(config.emailNotificacion, subject, '', {
      htmlBody: htmlBody
    });
    
    Logger.log('Email HTML enviado: ' + subject);
    return true;
    
  } catch (error) {
    Logger.log('Error al enviar email HTML: ' + error.message);
    return false;
  }
}

// ============================================================================
// GOOGLE CALENDAR - EVENTOS AUTOMÁTICOS
// ============================================================================

/**
 * Crea un evento en Google Calendar.
 */
function createCalendarEvent(title, description, startTime, endTime) {
  try {
    const config = getConfig();
    
    if (config.crearEventoCalendar !== true && config.crearEventoCalendar !== 'TRUE') {
      Logger.log('Creación de eventos deshabilitada');
      return false;
    }
    
    const calendar = CalendarApp.getCalendarById(config.calendarioId);
    
    if (!calendar) {
      Logger.log('Calendario no encontrado');
      return false;
    }
    
    calendar.createEvent(title, startTime, endTime, {
      description: description,
      sendInvites: false
    });
    
    Logger.log('Evento creado: ' + title);
    return true;
    
  } catch (error) {
    Logger.log('Error al crear evento: ' + error.message);
    return false;
  }
}

// ============================================================================
// PROCESAMIENTO DE EVENTOS
// ============================================================================

/**
 * Procesa usuario inactivo: email + log.
 */
function processInactiveUser(user) {
  logEvent({
    type: 'USUARIO_INACTIVO',
    user: user.name,
    details: 'Usuario desactivado',
    status: 'ALERTA',
    action: 'Email enviado'
  });
  
  const subject = '⚠️ Usuario inactivo detectado';
  const body = `El usuario ${user.name} (${user.email}) del grupo ${user.group} fue marcado como inactivo.

Fecha: ${new Date().toLocaleString('es-MX')}

Sistema de Gestión Workspace - Turing IA`;
  
  sendNotification(subject, body);
  Logger.log(`Usuario inactivo procesado: ${user.name}`);
}

/**
 * Notifica cambio de rol (alerta si es Admin).
 */
function notifyRoleChange(user) {
  logEvent({
    type: 'ROL_MODIFICADO',
    user: user.name,
    details: `Nuevo rol: ${user.role}`,
    status: user.role === 'Admin' ? 'WARNING' : 'OK',
    action: user.role === 'Admin' ? 'Email enviado' : 'Solo registro'
  });
  
  if (user.role === 'Admin') {
    const subject = '⚠️ Cambio Crítico: Nuevo Administrador';
    const body = `ATENCIÓN: ${user.name} fue promovido a Admin.

Email: ${user.email}
Grupo: ${user.group}
Fecha: ${new Date().toLocaleString('es-MX')}

Verifica que este cambio esté autorizado.

Sistema de Gestión Workspace - Turing IA`;
    
    sendNotification(subject, body);
  }
}

/**
 * Procesa nuevo usuario: email HTML + evento Calendar.
 */
function processNewUser(user, row) {
  // Actualizar fechas
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Usuarios');
  const ahora = new Date();
  
  sheet.getRange(row, 6).setValue(ahora);
  sheet.getRange(row, 7).setValue(ahora);
  
  // Resto del código original
  logEvent({
    type: 'USUARIO_AGREGADO',
    user: user.name,
    details: `Rol: ${user.role}, Grupo: ${user.group}`,
    status: 'OK',
    action: 'Email y evento creados'
  });
  
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
  
  sendHtmlNotification('✅ Nuevo usuario agregado', htmlBody);
  
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(10, 0, 0, 0);
  
  const endTime = new Date(tomorrow);
  endTime.setHours(11, 0, 0, 0);
  
  const description = `Sesión de Onboarding

👤 Usuario: ${user.name}
📧 Email: ${user.email}
🏷️ Rol: ${user.role}
👥 Grupo: ${user.group}

Agenda:
1. Bienvenida al equipo
2. Explicación de roles y permisos
3. Tour por herramientas
4. Asignación de tareas

Sistema de Gestión Workspace - Turing IA`;
  
  createCalendarEvent(
    `🎯 Onboarding: ${user.name}`,
    description,
    tomorrow,
    endTime
  );
  
  Logger.log(`Nuevo usuario procesado: ${user.name}`);
}

/**
 * Envía reporte de usuarios inactivos.
 */
function enviarReporteInactivos(usuariosInactivos) {
  let tabla = '';
  usuariosInactivos.forEach(u => {
    tabla += `<tr>
      <td style="padding: 8px; border: 1px solid #ddd;">${u.name}</td>
      <td style="padding: 8px; border: 1px solid #ddd;">${u.group}</td>
      <td style="padding: 8px; border: 1px solid #ddd;">${u.days} días</td>
    </tr>`;
  });
  
  const htmlBody = `
    <h2>📊 Reporte Diario - Usuarios Inactivos</h2>
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
  
  sendHtmlNotification(
    `📊 Reporte Diario - ${new Date().toLocaleDateString('es-MX')}`,
    htmlBody
  );
}
/**
 * Prueba Notifications.gs: Email y Calendar
 */
function probarNotifications() {
  Logger.log('=== PROBANDO NOTIFICATIONS.GS ===');
  
  // Probar sendNotification()
  try {
    const resultado = sendNotification(
      'Prueba de Email',
      'Este es un email de prueba desde Apps Script.'
    );
    if (resultado) {
      Logger.log('✓ sendNotification() funciona - Revisa tu email');
    } else {
      Logger.log('✗ sendNotification() no envió (revisa config)');
    }
  } catch (error) {
    Logger.log('✗ sendNotification() falló: ' + error.message);
  }
  
  // Probar sendHtmlNotification()
  try {
    const resultado = sendHtmlNotification(
      'Prueba de Email HTML',
      '<h2>Prueba HTML</h2><p>Este es un email con <strong>formato</strong>.</p>'
    );
    if (resultado) {
      Logger.log('✓ sendHtmlNotification() funciona - Revisa tu email');
    } else {
      Logger.log('✗ sendHtmlNotification() no envió');
    }
  } catch (error) {
    Logger.log('✗ sendHtmlNotification() falló: ' + error.message);
  }
  
  // Probar createCalendarEvent()
  try {
    const manana = new Date();
    manana.setDate(manana.getDate() + 1);
    manana.setHours(15, 0, 0, 0);
    
    const fin = new Date(manana);
    fin.setHours(16, 0, 0, 0);
    
    const resultado = createCalendarEvent(
      'Prueba de Evento',
      'Este es un evento de prueba creado por Apps Script.',
      manana,
      fin
    );
    
    if (resultado) {
      Logger.log('✓ createCalendarEvent() funciona - Revisa tu Calendar');
    } else {
      Logger.log('✗ createCalendarEvent() no creó evento');
    }
  } catch (error) {
    Logger.log('✗ createCalendarEvent() falló: ' + error.message);
  }
  
  Logger.log('=== PRUEBA NOTIFICATIONS.GS COMPLETADA ===');
}
