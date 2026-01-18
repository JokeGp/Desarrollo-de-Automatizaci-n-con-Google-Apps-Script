/**
 * ============================================================================
 * WORKSPACE AUTOMATION - TURING IA
 * Archivo: Code.gs
 * ============================================================================
 * 
 * Funciones core: configuración, lectura de datos y registro de eventos.
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
 */
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Configuración');
  
  if (!sheet) {
    throw new Error('Hoja Configuración no encontrada');
  }
  
  const data = sheet.getDataRange().getValues();
  const config = {};
  
  for (let i = 1; i < data.length; i++) {
    config[data[i][0]] = data[i][1];
  }
  
  return config;
}

// ============================================================================
// LECTURA DE DATOS
// ============================================================================

/**
 * Obtiene todos los usuarios del sistema.
 */
function getUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Usuarios');
  
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const users = [];
  
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    
    users.push({
      name: data[i][0],
      email: data[i][1],
      role: data[i][2],
      group: data[i][3],
      active: data[i][4] === true || data[i][4] === 'TRUE',
      dateRegistered: data[i][5],
      lastAccess: data[i][6]
    });
  }
  
  return users;
}

/**
 * Calcula días desde el último acceso.
 */
function getDaysSinceLastAccess(lastAccessDate) {
  const today = new Date();
  const lastAccess = new Date(lastAccessDate);
  const diffTime = Math.abs(today - lastAccess);
  return Math.floor(diffTime / (1000 * 60 * 60 * 24));
}

// ============================================================================
// REGISTRO DE EVENTOS
// ============================================================================

/**
 * Registra un evento en la hoja RegistroEventos para auditoría.
 */
function logEvent(event) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('RegistroDeEventos');
    
    if (!sheet) {
      Logger.log('Hoja RegistroEventos no encontrada');
      return;
    }
    
    sheet.appendRow([
      new Date(),
      event.type || 'EVENTO',
      event.user || 'Sistema',
      event.details || '',
      event.status || 'OK',
      event.action || 'Ninguna'
    ]);
    
    Logger.log(`Evento registrado: ${event.type} - ${event.user}`);
    
  } catch (error) {
    Logger.log('Error al registrar evento: ' + error.message);
  }
}

/**
 * Prueba Code.gs: Lee configuración y usuarios
 */
function probarCode() {
  Logger.log('=== PROBANDO CODE.GS ===');
  
  // Probar getConfig()
  try {
    const config = getConfig();
    Logger.log('✓ getConfig() funciona');
    Logger.log('Email configurado: ' + config.emailNotificacion);
  } catch (error) {
    Logger.log('✗ getConfig() falló: ' + error.message);
  }
  
  // Probar getUsers()
  try {
    const users = getUsers();
    Logger.log('✓ getUsers() funciona');
    Logger.log('Total usuarios: ' + users.length);
  } catch (error) {
    Logger.log('✗ getUsers() falló: ' + error.message);
  }
  
  // Probar logEvent()
  try {
    logEvent({
      type: 'PRUEBA',
      user: 'Test',
      details: 'Prueba de Code.gs',
      status: 'OK',
      action: 'Ninguna'
    });
    Logger.log('✓ logEvent() funciona');
  } catch (error) {
    Logger.log('✗ logEvent() falló: ' + error.message);
  }
  
  Logger.log('=== PRUEBA CODE.GS COMPLETADA ===');
}
