function doGet() {
  var correoAcceso = Session.getActiveUser().getEmail().toLowerCase().trim();
  var esAdmin = ["fernando.alcala@gigas.com"].includes(correoAcceso);
  var template = HtmlService.createTemplateFromFile('Index');
  template.esAdmin = esAdmin;
  template.userEmail = correoAcceso; 
  
  return template.evaluate()
    .setTitle('Gigas IT - Portal de Gestión Global')
    .setFaviconUrl('https://gigas.com/favicon.ico')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =========================================================
// 🛡️ MÓDULO DE SEGURIDAD Y PERMISOS DINÁMICOS
// =========================================================

function obtenerOperadoresAutorizados() {
  var operadores = ["fernando.alcala@gigas.com"];
  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Permisos_Panel");
    if (hoja) {
      var datos = hoja.getDataRange().getValues();
      for (var i = 1; i < datos.length; i++) {
        if (datos[i][0]) operadores.push(String(datos[i][0]).toLowerCase().trim());
      }
    }
  } catch(e) {}
  return operadores;
}

function comprobarRolUsuarioActivo() { 
  return obtenerOperadoresAutorizados().includes(Session.getActiveUser().getEmail().toLowerCase().trim());
}

function verificarPermisoEjecucion() { 
  if (!comprobarRolUsuarioActivo()) {
    throw new Error("Sin privilegios operativos de administración. Solicita acceso al SuperAdministrador.");
  }
}

function webAlternarPermisosPanel(emailObjetivo) {
  var ejecutor = Session.getActiveUser().getEmail().toLowerCase().trim();
  if (ejecutor !== "fernando.alcala@gigas.com") {
    throw new Error("Solo el SuperAdministrador (Fernando) puede otorgar o revocar accesos al panel.");
  }

  var emailClean = emailObjetivo.toLowerCase().trim();
  if (emailClean === "fernando.alcala@gigas.com") return "🛡️ No puedes quitarte los permisos a ti mismo.";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Permisos_Panel");
  
  if (!hoja) {
    hoja = ss.insertSheet("Permisos_Panel");
    hoja.appendRow(["Correos Autorizados", "Añadido por", "Fecha"]);
    hoja.getRange("A1:C1").setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    hoja.hideSheet(); 
  }

  var datos = hoja.getDataRange().getValues();
  var encontrado = false;
  var filaBorrar = -1;

  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][0]).toLowerCase().trim() === emailClean) {
      encontrado = true;
      filaBorrar = i + 1;
      break;
    }
  }

  if (encontrado) {
    hoja.deleteRow(filaBorrar);
    registrarEnHistorial("REVOCAR ACCESO", "Se han quitado los permisos del panel a: " + emailClean);
    return "🔴 Permisos REVOCADOS para " + emailClean;
  } else {
    var fechaHoy = Utilities.formatDate(new Date(), "Europe/Madrid", "dd/MM/yyyy HH:mm");
    hoja.appendRow([emailClean, ejecutor, fechaHoy]);
    registrarEnHistorial("OTORGAR ACCESO", "Se ha dado acceso total al panel a: " + emailClean);
    return "🟢 Permisos CONCEDIDOS a " + emailClean + ". Ya puede operar.";
  }
}

// =========================================================
// 🏷️ MÓDULO DE GESTIÓN DE ALIAS EN CALIENTE
// =========================================================

function webObtenerAliasUsuario(userEmail) {
  try {
    var userObj = AdminDirectory.Users.get(userEmail.toLowerCase().trim(), { projection: 'full' });
    return userObj.aliases || [];
  } catch(e) {
    return [];
  }
}

function webAñadirAliasUsuario(userEmail, nuevoAlias) {
  verificarPermisoEjecucion();
  var emailClean = userEmail.toLowerCase().trim();
  var aliasClean = nuevoAlias.toLowerCase().trim().replace(/\s+/g, '');
  if (!aliasClean.includes("@")) {
    var dominio = emailClean.split("@")[1];
    aliasClean += "@" + dominio;
  }
  try {
    AdminDirectory.Users.Aliases.insert({ alias: aliasClean }, emailClean);
    registrarEnHistorial("AÑADIR ALIAS", "Alias " + aliasClean + " asignado a " + emailClean);
    return "✅ Alias añadido correctamente a vuestro Workspace.";
  } catch(e) {
    throw new Error("Google rechazó el alias: " + e.message);
  }
}

function webEliminarAliasUsuario(userEmail, aliasBorrar) {
  verificarPermisoEjecucion();
  var emailClean = userEmail.toLowerCase().trim();
  var aliasClean = aliasBorrar.toLowerCase().trim();
  try {
    AdminDirectory.Users.Aliases.remove(emailClean, aliasClean);
    registrarEnHistorial("ELIMINAR ALIAS", "Alias " + aliasClean + " borrado de " + emailClean);
    return "🗑️ Alias eliminado con éxito de Workspace.";
  } catch(e) {
    throw new Error("No se pudo eliminar el alias: " + e.message);
  }
}

// =========================================================
// MOTOR PRINCIPAL
// =========================================================

function purificarDatos(matriz) {
  if (!matriz || matriz.length === 0) return [];
  var datosLimpios = [];
  for (var i = 0; i < matriz.length; i++) {
    var fila = matriz[i];
    var filaVacia = true; var filaLimpia = [];
    for (var j = 0; j < fila.length; j++) {
      var celda = fila[j];
      if (celda !== "" && celda !== null) filaVacia = false;
      if (celda instanceof Date) {
        filaLimpia.push(Utilities.formatDate(celda, "Europe/Madrid", "dd/MM/yyyy"));
      } else {
        var filaLinter = celda == null ? "" : celda.toString().trim();
        filaLimpia.push(filaLinter);
      }
    }
    if (!filaVacia) datosLimpios.push(filaLimpia);
  }
  return datosLimpios;
}

function procesarEstructuraListas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaListas = ss.getSheetByName("Inventario Listas Workspace") || ss.getSheetByName("Listas");
  
  var listasGigas = [];
  var listasOni = [];
  var fechaListas = "No sincronizado"; 
  
  var metricasGigas = { total: 0, conOwner: 0, sinOwner: 0 };
  var metricasOni = { total: 0, conOwner: 0, sinOwner: 0 };
  
  if (hojaListas) {
    var valZ1 = hojaListas.getRange("Z1").getValue();
    if (valZ1) fechaListas = valZ1 instanceof Date ? Utilities.formatDate(valZ1, "Europe/Madrid", "dd/MM/yyyy HH:mm") : valZ1.toString();
    
    var data = purificarDatos(hojaListas.getDataRange().getValues());
    if (data.length > 0) {
      for (var f = 0; f < data.length; f++) {
        var fila = data[f];
        if (!fila[1] || !fila[1].toString().includes("@")) continue; 
        
        var emailLista = fila[1].toString().trim().toLowerCase();
        var ownerStr = fila[3] ? fila[3].toString().trim() : "";
        var nombreLista = fila[0] ? fila[0].toString().trim().toLowerCase() : "";
        
        // Detector inteligente: si es .pt, si el dueño es de oni o el nombre dice portugal/oni -> es ONI
        var esOni = emailLista.includes("@oni.pt") || emailLista.includes(".pt@") || ownerStr.toLowerCase().includes("@oni.pt") || nombreLista.includes("portugal") || nombreLista.includes("oni");
        
        var metricasActivas = esOni ? metricasOni : metricasGigas;
        var arrayDestino = esOni ? listasOni : listasGigas;
        
        metricasActivas.total++;
        if (!ownerStr.includes("@")) metricasActivas.sinOwner++; else metricasActivas.conOwner++;
        
        var miembros = [];
        for (var c = 4; c < fila.length; c++) { 
          if (fila[c] !== "") {
            var separados = fila[c].toString().split(/[\n,]+/);
            for(var s = 0; s < separados.length; s++){
              if(separados[s].trim() !== "") miembros.push(separados[s].trim());
            }
          } 
        }
        arrayDestino.push({ nombre: fila[0], email: fila[1], volumen: fila[2], owner: ownerStr, miembros: miembros });
      }
    }
  }
  return { 
    listasGigas: listasGigas, 
    listasOni: listasOni, 
    metricasGigas: metricasGigas, 
    metricasOni: metricasOni, 
    fechaListas: fechaListas 
  };
}

function motorSincronizarListasWorkspace(customerId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaListas = ss.getSheetByName("Inventario Listas Workspace") || ss.getSheetByName("Listas");
  if (!hojaListas) return "⚠️ Pestaña de listas no encontrada. ";

  var pageToken;
  var todasLasListas = [];
  do {
    var resGrupos = AdminDirectory.Groups.list({ customer: customerId, maxResults: 200, pageToken: pageToken });
    if (resGrupos.groups) {
      for (var i = 0; i < resGrupos.groups.length; i++) {
        todasLasListas.push(resGrupos.groups[i]);
      }
    }
    pageToken = resGrupos.nextPageToken;
  } while (pageToken);

  var datosExcel = [];
  for (var g = 0; g < todasLasListas.length; g++) {
    var grupo = todasLasListas[g];
    var emailGrupo = grupo.email;
    var nombreGrupo = grupo.name || "";
    var owner = "";
    var miembrosNormales = [];
    var volumen = 0;
    var pageTokenMiembros;

    try {
      do {
        var resMiembros = AdminDirectory.Members.list(emailGrupo, { maxResults: 200, pageToken: pageTokenMiembros });
        if (resMiembros.members) {
          for (var m = 0; m < resMiembros.members.length; m++) {
            var miembro = resMiembros.members[m];
            volumen++;
            if (miembro.role === "OWNER") owner = miembro.email;
            else miembrosNormales.push(miembro.email);
          }
        }
        pageTokenMiembros = resMiembros.nextPageToken;
      } while (pageTokenMiembros);
    } catch(e) {}

    datosExcel.push([nombreGrupo, emailGrupo, volumen + " miembros", owner, miembrosNormales.join(", ")]);
  }

  var ultimaFila = hojaListas.getLastRow();
  var ultimaColumna = hojaListas.getLastColumn();
  
  if (ultimaFila > 1 && ultimaColumna > 0) {
    hojaListas.getRange(2, 1, ultimaFila - 1, ultimaColumna).clearContent().setBackground(null);
  }
  
  if (datosExcel.length > 0) {
    hojaListas.getRange(2, 1, datosExcel.length, 5).setValues(datosExcel);
  }
  
  // 🔄 CORRECCIÓN: Restauramos la cabecera E1 y guardamos la fecha oculta en Z1
  hojaListas.getRange("E1").setValue("Miembros");
  hojaListas.getRange("Z1").setValue(new Date());
  
  return "📋 Listas (" + todasLasListas.length + ") actualizadas. ";
}

function obtenerDatosDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hd = ss.getSheetByName("DashBoard") || ss.getSheets()[0]; 
  var lic = [];
  var catOri = []; var graf = [];
  try {
    var licData = hd.getRange(3, 1, 5, 5).getValues();
    lic = purificarDatos(licData); 
    for (var i = 0; i < lic.length; i++) { if (lic[i][3] && lic[i][3].includes("DETALLES")) lic[i][3] = ""; }
    var oriData = hd.getRange(4, 6, 11, 5).getValues();
    var ori = purificarDatos(oriData);
    for (var i = 0; i < ori.length; i++) {
      if (ori[i][0] && !String(ori[i][0]).includes("Dedicacion") && !String(ori[i][0]).toLowerCase().includes("total")) {
        var activas = parseInt(ori[i][1]) || 0;       
        var suspendidas = parseInt(ori[i][2]) || 0;   
        var total = parseInt(ori[i][3]) || 0;
        catOri.push({ concepto: ori[i][0], activas: activas, suspendidas: suspendidas, total: total });
        graf.push([ori[i][0], total]);
      }
    }
  } catch(e) {}
  
  var base = { headers: [], rows: [] };
  var serv = { headers: [], rows: [] };
  var hm = ss.getSheetByName("Hoja Maestra (657)");
  if (hm) { 
    var rd = purificarDatos(hm.getDataRange().getValues()); 
    if (rd.length > 0) { base.headers = rd[0]; base.rows = rd.slice(1); } 
  }
  
  var hs = ss.getSheetByName("Cuentas de servicio Gigas (139)") || ss.getSheetByName("Cuentas de servicio gigas (139)"); 
  if (hs) { 
    var rs = purificarDatos(hs.getDataRange().getValues());
    if (rs.length > 0) { serv.headers = rs[0]; serv.rows = rs.slice(1); } 
  }
  
  var dl = procesarEstructuraListas();
  var actuales = 0; var anteriores = 0;
  var historialAgrupado = {}; 
  var mapaLogons = {};
  
  var asignadasStarter = 0;
  var asignadasStandard = 0;
  var asignadasPlus = 0;
  var asignadasGemini = 0;
  var conteoApps = {};
  try {
    var listaUsuarios = AdminDirectory.Users.list({domain: "gigas.com", maxResults: 1});
    var customerIdReal = listaUsuarios.users[0].customerId;
    var tokenApps;
    do {
      var resApps = AdminLicenseManager.LicenseAssignments.listForProduct("Google-Apps", customerIdReal, { maxResults: 500, pageToken: tokenApps });
      if (resApps.items) {
        for (var i = 0; i < resApps.items.length; i++) {
          var sku = resApps.items[i].skuId;
          conteoApps[sku] = (conteoApps[sku] || 0) + 1;
        }
      }
      tokenApps = resApps.nextPageToken;
    } while (tokenApps);

    asignadasPlus = conteoApps["1010020020"] || 0;
    asignadasStarter = conteoApps["1010020027"] || 0;
    asignadasStandard = conteoApps["1010020026"] || 0;

    var tokenGemini;
    do {
      try {
        var resGemini = AdminLicenseManager.LicenseAssignments.listForProduct("101047", customerIdReal, { maxResults: 500, pageToken: tokenGemini });
        if (resGemini.items) {
          for (var i = 0; i < resGemini.items.length; i++) {
             if (resGemini.items[i].skuId === "1010470001") asignadasGemini++;
          }
        }
        tokenGemini = resGemini.nextPageToken;
      } catch(e) { break; }
    } while (tokenGemini);

    var pageToken;
    var hoy = new Date();
    var mesActual = hoy.getMonth(); var anioActual = hoy.getFullYear();
    var mesAnterior = mesActual === 0 ? 11 : mesActual - 1;
    var anioAnterior = mesActual === 0 ? anioActual - 1 : anioActual;
    do {
      var response = AdminDirectory.Users.list({ customer: customerIdReal, maxResults: 500, projection: 'basic', pageToken: pageToken });
      if (response.users) {
        for (var i = 0; i < response.users.length; i++) {
          if (response.users[i].creationTime) {
            var fC = new Date(response.users[i].creationTime);
            if (fC.getFullYear() === anioActual && fC.getMonth() === mesActual) actuales++;
            else if (fC.getFullYear() === anioAnterior && fC.getMonth() === mesAnterior) anteriores++;
            var keyMes = fC.getFullYear() + "-" + ("0" + (fC.getMonth() + 1)).slice(-2);
            if (!historialAgrupado[keyMes]) historialAgrupado[keyMes] = 0;
            historialAgrupado[keyMes]++;
          }
          var loginTime = response.users[i].lastLoginTime;
          var valLogin = "Nunca ha iniciado sesión";
          if (loginTime && loginTime !== "1970-01-01T00:00:00.000Z") {
            var d = new Date(loginTime);
            valLogin = ("0" + d.getDate()).slice(-2) + "/" + ("0" + (d.getMonth() + 1)).slice(-2) + "/" + d.getFullYear();
          }
          mapaLogons[response.users[i].primaryEmail.toLowerCase()] = valLogin;
        }
      }
      pageToken = response.nextPageToken;
    } while (pageToken);
  } catch(e) { actuales = "Error"; anteriores = "Error"; }

  var historialKeys = Object.keys(historialAgrupado).sort();
  if (historialKeys.length > 12) historialKeys = historialKeys.slice(historialKeys.length - 12);
  var graficaHistorico = [];
  for (var k = 0; k < historialKeys.length; k++) {
    var partes = historialKeys[k].split("-");
    var mesLabel = partes[1] + "/" + partes[0].substring(2); 
    graficaHistorico.push([mesLabel, historialAgrupado[historialKeys[k]]]);
  }

  if (base.headers.length > 0) {
    var emailIdx = -1; var logonIdx = -1;
    for(var h = 0; h < base.headers.length; h++) {
       var hl = String(base.headers[h]).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
       if (hl.includes("email")) emailIdx = h;
       if (hl.includes("logon") || hl.includes("sesion") || hl.includes("inicio")) logonIdx = h;
    }
    if (logonIdx === -1) { base.headers.push("Último Logon"); logonIdx = base.headers.length - 1; }
    
    for (var r = 0; r < base.rows.length; r++) {
      var email = emailIdx > -1 ? String(base.rows[r][emailIdx]).toLowerCase().trim() : "";
      var logonVal = mapaLogons[email] || "No encontrado en Workspace";
      if (base.rows[r].length <= logonIdx) base.rows[r].push(logonVal);
      else base.rows[r][logonIdx] = logonVal; 
    }
  }

  lic = [
    ["✨ Gemini Enterprise (versión antigua)", "Activas", "Asignadas: " + asignadasGemini, "Precio del distribuidor", "∞"],
    ["💼 Google Workspace Business Starter", "Activas", "Asignadas: " + asignadasStarter, "Precio del distribuidor", "300"],
    ["🚀 Google Workspace Enterprise Plus", "Activas", "Asignadas: " + asignadasPlus, "Precio del distribuidor", "10"],
    ["📊 Google Workspace Enterprise Standard", "Activas", "Asignadas: " + asignadasStandard, "Precio del distribuidor", "554"]
  ];
  return JSON.stringify({ 
    licencias: lic, origen: catOri, grafica: graf, listas: dl.listas, 
    fechaSincro: dl.fechaListas, metricasGrupos: dl.metricasGrupos, 
    baseCuentas: base, baseServicios: serv, 
    crecimiento: { actual: actuales, anterior: anteriores, historico: graficaHistorico }
  });
}

function webResetearClavePorOlvido(e) { 
  verificarPermisoEjecucion(); 
  var p = Math.random().toString(36).slice(-10) + "Gigs26!"; 
  AdminDirectory.Users.update({ password: p, changePasswordAtNextLogin: true }, e.trim());
  return "🔑 NUEVA CLAVE: " + p; 
}

function webEjecutarFlujoBajaTotal(correoBaja, managerDestino, crearAlias, transferirDrive, destinoDrive) {
  if (typeof comprobarRolUsuarioActivo === "function" && !comprobarRolUsuarioActivo()) {
    throw new Error("No tienes permisos de Administrador para ejecutar bajas críticas.");
  }
  var msgWorkspace = "";
  var msgExcel = "";
  var correoBajaClean = correoBaja.toLowerCase().trim();
  var managerClean = managerDestino ? managerDestino.toLowerCase().trim() : "";
  var destinoDriveClean = destinoDrive ? destinoDrive.toLowerCase().trim() : "";

  // 🌟 PASO 1: TRANSFERENCIA DE DRIVE (SEGÚN RESPUESTA DEL PROMPT DEL FRONTEND)
  if (transferirDrive && destinoDriveClean && destinoDriveClean.includes("@")) {
    try {
      ejecutarTransferenciaDrive(correoBajaClean, destinoDriveClean);
      msgWorkspace += "Propiedad de Google Drive transferida con éxito a " + destinoDriveClean + ". ";
    } catch (eDrive) {
      // Si el Drive falla, abortamos la baja inmediatamente para evitar pérdida accidental de datos
      throw new Error("Abortando baja crítica por seguridad: " + eDrive.message);
    }
  }

  // 🌟 PASO 2: ELIMINACIÓN DE LA CUENTA EN WORKSPACE
  try {
    AdminDirectory.Users.remove(correoBajaClean);
    msgWorkspace += "La cuenta ha sido eliminada de Workspace.";
    
    // 🌟 PASO 3: CREACIÓN DEL ALIAS DE REDIRECCIÓN EN EL MÁNAGER
    if (crearAlias && managerClean) {
      Utilities.sleep(8000); // Pausa de cortesía para la propagación de Google
      try {
        AdminDirectory.Users.Aliases.insert({ alias: correoBajaClean }, managerClean);
        msgWorkspace += " Alias creado en " + managerClean + ".";
      } catch (eAlias) {
        Utilities.sleep(5000);
        try {
          AdminDirectory.Users.Aliases.insert({ alias: correoBajaClean }, managerClean);
          msgWorkspace += " Alias creado en " + managerClean + " (intento 2).";
        } catch (eAlias2) {
          msgWorkspace += " (Fallo al crear alias: " + eAlias2.message + ").";
        }
      }
    }
  } catch (eWS) {
    throw new Error("Fallo al eliminar de Workspace: " + eWS.message);
  }

  // 🌟 PASO 4: PURGA DE FILAS EN LA HOJA MAESTRA (657)
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Hoja Maestra (657)") || ss.getSheets()[0];
    var datos = hoja.getDataRange().getValues();
    var filasBorradas = 0;
    if (datos.length >= 2) {
      var cabeceras = datos[0].map(function(h) { return String(h).toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); });
      var idxEmail = cabeceras.indexOf("email address") > -1 ? cabeceras.indexOf("email address") : (cabeceras.indexOf("email address [required]") > -1 ? cabeceras.indexOf("email address [required]") : cabeceras.indexOf("email"));
      if (idxEmail !== -1) {
        for (var f = datos.length - 1; f >= 1; f--) {
          if (String(datos[f][idxEmail]).toLowerCase().trim() === correoBajaClean) {
            hoja.deleteRow(f + 1);
            filasBorradas++;
          }
        }
      }
    }
    msgExcel = " Excel purgado (" + filasBorradas + " fila).";
  } catch (eEx) {
    msgExcel = " (Aviso Excel: " + eEx.message + ").";
  }

  var logDrive = transferirDrive ? destinoDriveClean : "No transferido";
  registrarEnHistorial("BAJA TOTAL", "Cuenta: " + correoBajaClean + " | Alias a: " + (managerClean || "Ninguno") + " | Drive a: " + logDrive);
  
  var usuarioAD = correoBajaClean.split("@")[0];
  var comandosAD = "# 1. Deshabilitar cuenta AD\nDisable-ADAccount -Identity '" + usuarioAD + "'\n\n# 2. Borrar cuenta AD\nRemove-ADUser -Identity '" + usuarioAD + "' -Confirm:$false";
  var snippetHTML = "<br><br><b style='color:#fff;'>💻 Snippet PowerShell (AD) - GCDS:</b><br><textarea readonly style='width:100%; height:80px; margin-top:6px; background:#0f172a; color:#38bdf8; border:1px solid #334155; border-radius:4px; padding:10px; font-family:monospace; font-size:13px; outline:none; resize:none;' onclick='this.select()'>" + comandosAD + "</textarea>";
  
  return msgWorkspace + msgExcel + snippetHTML;
}

// 🚀 FUNCIÓN COMPLEMENTARIA: CONECTOR INTERNO CON LA DATA TRANSFER API DE GOOGLE
function ejecutarTransferenciaDrive(correoViejo, correoNuevo) {
  try {
    // La API de transferencia requiere los IDs numéricos e inmutables del perfil, no los strings de correo
    var idViejo = AdminDirectory.Users.get(correoViejo).id;
    var idNuevo = AdminDirectory.Users.get(correoNuevo).id;

    var url = "https://www.googleapis.com/admin/datatransfer/v1/transfers";
    var payload = {
      "oldOwnerUserId": idViejo,
      "newOwnerUserId": idNuevo,
      "applicationDataTransfers": [
        {
          "applicationId": "55656082996", // ID de aplicación fijo y global de Google para el servicio de Drive
          "applicationTransferParams": [
            {
              "key": "PRIVACY_LEVEL",
              "value": ["SHARED", "PRIVATE"] // Mueve tanto lo suyo privado como lo compartido con el equipo
            }
          ]
        }
      ]
    };

    var opciones = {
      "method": "post",
      "contentType": "application/json",
      "headers": {
        "Authorization": "Bearer " + ScriptApp.getOAuthToken()
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    var respuesta = UrlFetchApp.fetch(url, opciones);
    var resCode = respuesta.getResponseCode();
    
    if (resCode !== 200 && resCode !== 201) {
      throw new Error("Error en servidor Google Data Transfer (" + resCode + "): " + respuesta.getContentText());
    }
    return true;
  } catch (e) {
    registrarEnHistorial("FALLO TRANSFERENCIA DRIVE", "Origen: " + correoViejo + " | Destino: " + correoNuevo + " | Detalle: " + e.message);
    throw new Error("La API de transferencia falló: " + e.message);
  }
}

function webHacerOwnerLista(g, u) { 
  verificarPermisoEjecucion();
  try { AdminDirectory.Members.insert({ email: u, role: "OWNER" }, g); } catch (e) { AdminDirectory.Members.update({email: u, role: "OWNER"}, g, u); }
  
  registrarEnHistorial("ASIGNAR OWNER", "Usuario: " + u + " promovido a OWNER de la lista: " + g);
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventario Listas Workspace") || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Listas");
  if(hoja) {
    var datos = hoja.getRange("A1:D" + hoja.getLastRow()).getValues();
    for(var f=0; f<datos.length; f++){ if(datos[f][1] && datos[f][1].toString().toLowerCase().trim() === g.toLowerCase().trim()){ hoja.getRange(f+1, 4).setValue(u); break; } }
  }
  return "👑 " + u + " ha sido nombrado Owner.";
}

function webEnviarCampanaHuerfanas() {
  verificarPermisoEjecucion(); var total = 0; var saltados = 0;
  var hojaListas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventario Listas Workspace") || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Listas");
  var datos = hojaListas.getDataRange().getValues(); var fechaHoy = Utilities.formatDate(new Date(), "Europe/Madrid", "dd/MM/yyyy");
  for (var f = 0; f < datos.length; f++) {
    var fila = datos[f];
    if (!fila[1] || !fila[1].toString().includes("@")) continue;
    var emailLista = fila[1].toString().trim(); var ownerStr = fila[3] ? fila[3].toString().trim() : "";
    if (ownerStr.includes("Enviado") || ownerStr.includes("Lanzada")) { saltados++; continue; }
    if (!ownerStr.includes("@")) { MailApp.sendEmail(emailLista, "🔍 Auditoría IT: Identificación de Responsable [" + emailLista + "] [REF-HUERFANA]", "Contestar rellenando esta línea:\n\nCorreo del owner: ");
    hojaListas.getRange(f + 1, 4).setValue("⏳ Enviado el " + fechaHoy); total++;
    }
  }
  return "Campaña procesada. Enviados: " + total;
}

function webProcesarRespuestasHuerfanas() {
  verificarPermisoEjecucion();
  var hilos = GmailApp.search('subject:"[REF-HUERFANA]" is:unread', 0, 30); if (hilos.length === 0) return "Sin respuestas.";
  var hojaListas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventario Listas Workspace") || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Listas");
  var act = 0;
  for (var i = 0; i < hilos.length; i++) {
    var msg = hilos[i].getMessages()[hilos[i].getMessages().length - 1];
    var cuerpo = msg.getPlainBody();
    var matchLista = msg.getSubject().match(/\[([^\]]+)\]/); if (!matchLista) { hilos[i].markRead(); continue; }
    var emailLista = matchLista[1].toLowerCase().trim();
    var matchOwner = cuerpo.match(/Correo\s*del\s*owner:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i);
    if (matchOwner) {
      var nO = matchOwner[1].toLowerCase().trim();
      try { 
        AdminDirectory.Members.insert({ email: nO, role: "OWNER" }, emailLista); 
        var datos = hojaListas.getDataRange().getValues();
        for (var f = 0; f < datos.length; f++) { 
          if (datos[f][1] && datos[f][1].toString().toLowerCase().trim() === emailLista) { 
            hojaListas.getRange(f + 1, 4).setValue(nO); break; 
          } 
        } 
        act++; 
      } catch (e) {}
    }
    hilos[i].markRead();
  }
  return "Sincronizados: " + act;
}

function webComprobarYRecuperar(correo, targetOu) { 
  verificarPermisoEjecucion();
  var correoClean = correo.toLowerCase().trim();
  var ouDestino = targetOu ? targetOu : "/";
  
  // 1. Necesitamos el ID del cliente para leer la papelera
  var listaUsuarios = AdminDirectory.Users.list({domain: "gigas.com", maxResults: 1});
  var customerIdReal = listaUsuarios.users[0].customerId;
  
  // 2. Buscar en la papelera para obtener el ID numérico (21 dígitos)
  var pageToken;
  var idParaRecuperar = null;
  do {
    var response = AdminDirectory.Users.list({ customer: customerIdReal, showDeleted: true, maxResults: 500, pageToken: pageToken });
    if (response.users) {
      for (var i = 0; i < response.users.length; i++) {
        if (response.users[i].primaryEmail.toLowerCase().trim() === correoClean) {
          idParaRecuperar = response.users[i].id; // 🌟 CAPTURAMOS EL ID INMUTABLE
          break;
        }
      }
    }
    if (idParaRecuperar) break;
    pageToken = response.nextPageToken;
  } while (pageToken);

  if (!idParaRecuperar) {
    throw new Error("No se encuentra el usuario en la papelera (puede que ya esté recuperado o hayan pasado los 20 días).");
  }

  // 3. Lanzar la recuperación usando el ID NUMÉRICO y la OU del panel
  try {
    AdminDirectory.Users.undelete({orgUnitPath: ouDestino}, idParaRecuperar); 
    registrarEnHistorial("RECUPERAR CUENTA", "Restaurada: " + correoClean + " en OU: " + ouDestino);
    return "✅ Cuenta restaurada con éxito en la OU: " + ouDestino; 
  } catch (e) {
    throw new Error("Fallo al restaurar en Workspace: " + e.message);
  }
}
function webCambiarEstadoCuenta(u, s) { AdminDirectory.Users.update({ suspended: s }, u); return "Estado actualizado."; }

function webBotonActualizarGrupos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Hoja Maestra (657)");
  var hojaDash = ss.getSheetByName("DashBoard") || ss.getSheets()[0];
  if (!hoja) return "Error: No se encuentra la pestaña 'Hoja Maestra (657)'";
  
  var msgListas = "";
  var mapaWS = {};
  var conteoLicencias = {}; 
  var pageToken;

  try {
    var listaUsuarios = AdminDirectory.Users.list({domain: "gigas.com", maxResults: 1});
    var customerIdReal = listaUsuarios.users[0].customerId;

    try {
      msgListas = motorSincronizarListasWorkspace(customerIdReal);
    } catch (eListas) {
      msgListas = "⚠️ Falló Listas: " + eListas.message + ". ";
    }

    do {
      var response = AdminDirectory.Users.list({ customer: customerIdReal, maxResults: 500, projection: 'full', pageToken: pageToken });
      if (response.users) {
        for (var i = 0; i < response.users.length; i++) {
          var u = response.users[i];
          var emailClean = u.primaryEmail.toLowerCase().trim();
          var mng = "no tiene";
          if (u.relations) {
            for (var j=0; j<u.relations.length; j++) {
              if (u.relations[j].type === "manager") { mng = u.relations[j].value.trim(); break; }
            }
          }
          mapaWS[emailClean] = {
            status: u.suspended ? "Suspended" : "Active",
            manager: mng,
            givenName: u.name ? u.name.givenName : "",
            familyName: u.name ? u.name.familyName : "",
            visitado: false,
            licencia: "Cloud Identity Free", // Valor por defecto
            ou: u.orgUnitPath || "/"
          };
        }
      }
      pageToken = response.nextPageToken;
    } while (pageToken);

    var pageTokenLic;
    // 🌟 MAPA CORREGIDO
    var nameMapSku = {
      "1010020027": "Google Workspace Business Starter",
      "1010020026": "Google Workspace Enterprise Standard",
      "1010020020": "Google Workspace Enterprise Plus"
    };
    do {
      var resLic = AdminLicenseManager.LicenseAssignments.listForProduct("Google-Apps", customerIdReal, { maxResults: 500, pageToken: pageTokenLic });
      if (resLic.items) {
        for (var l = 0; l < resLic.items.length; l++) {
          var item = resLic.items[l];
          var sku = item.skuId;
          var emailUser = item.userId.toLowerCase().trim();
          conteoLicencias[sku] = (conteoLicencias[sku] || 0) + 1;
          if (nameMapSku[sku] && mapaWS[emailUser]) {
            mapaWS[emailUser].licencia = nameMapSku[sku];
          }
        }
      }
      pageTokenLic = resLic.nextPageToken;
    } while (pageTokenLic);

  } catch(e) {
    return "Error al conectar con Google Workspace: " + e.message;
  }

  var datos = hoja.getDataRange().getValues();
  var cabeceras = datos[0].map(function(h) { return String(h).toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); });
  
  var idxEmail = cabeceras.indexOf("email address [required]");
  if (idxEmail === -1) idxEmail = cabeceras.indexOf("email address"); if (idxEmail === -1) idxEmail = cabeceras.indexOf("email");
  var idxStatus = cabeceras.indexOf("status [read only]"); if (idxStatus === -1) idxStatus = cabeceras.indexOf("status"); if (idxStatus === -1) idxStatus = cabeceras.indexOf("estado");
  var idxManager = cabeceras.indexOf("manager"); if (idxManager === -1) idxManager = cabeceras.indexOf("responsable");
  var idxOu = cabeceras.indexOf("ou"); if (idxOu === -1) idxOu = cabeceras.indexOf("unidad organizativa"); if (idxOu === -1) idxOu = cabeceras.indexOf("orgunit");

  // 🌟 Búsqueda inclusiva de la columna licencia para que no se pierda por el nombre
  var idxLicencia = -1;
  for (var h = 0; h < cabeceras.length; h++) {
    if (cabeceras[h].includes("licencia")) { idxLicencia = h; break; }
  }
  if (idxLicencia === -1) idxLicencia = 8; // Fallback

  var filasBorradas = 0;
  var contadorCambiosEstado = 0;
  var contadorCambiosLicencia = 0;
  var contadorCambiosOu = 0;

  if (idxEmail !== -1) {
    for (var f = datos.length - 1; f >= 1; f--) {
      var emailCell = String(datos[f][idxEmail]).toLowerCase().trim();
      if (emailCell.includes("@")) {
        if (!mapaWS[emailCell]) {
          hoja.deleteRow(f + 1);
          filasBorradas++;
        } else {
          mapaWS[emailCell].visitado = true;
          if (idxStatus !== -1 && String(datos[f][idxStatus]).trim() !== mapaWS[emailCell].status) {
            hoja.getRange(f + 1, idxStatus + 1).setValue(mapaWS[emailCell].status);
            contadorCambiosEstado++;
          }
          if (idxLicencia !== -1 && String(datos[f][idxLicencia]).trim() !== mapaWS[emailCell].licencia) {
            hoja.getRange(f + 1, idxLicencia + 1).setValue(mapaWS[emailCell].licencia);
            contadorCambiosLicencia++;
          }
          if (idxOu !== -1 && String(datos[f][idxOu]).trim() !== mapaWS[emailCell].ou) {
            hoja.getRange(f + 1, idxOu + 1).setValue(mapaWS[emailCell].ou);
            contadorCambiosOu++;
          }
        }
      }
    }
  }

  var filasInsertadas = 0;
  var intrusosGigas = [];
  var usuariosEnMemoria = Object.keys(mapaWS);

  for (var k = 0; k < usuariosEnMemoria.length; k++) {
    var emailIntruso = usuariosEnMemoria[k];
    if (!mapaWS[emailIntruso].visitado) {
      var usr = mapaWS[emailIntruso];
      var dominio = emailIntruso.split("@")[1] || "";
      var origenText = "Cuentas nominativas de Gigas";
      if (dominio.includes("oni")) origenText = "Cuentas de ONI (251)";
      else if (dominio.includes("onmovil")) origenText = "Cuentas Onmovil";
      else if (dominio.includes("asesorgigas")) origenText = "Cuentas asesorgigas";

      var nuevaFila = new Array(cabeceras.length).fill("");
      for (var c = 0; c < cabeceras.length; c++) {
        var head = cabeceras[c];
        if (head.includes("email") || head.includes("correo")) nuevaFila[c] = emailIntruso;
        else if (head.includes("status") || head.includes("estado")) nuevaFila[c] = usr.status;
        else if (head.includes("manager") || head.includes("responsable")) nuevaFila[c] = usr.manager;
        else if (head.includes("origen") || head.includes("hoja de origen")) nuevaFila[c] = origenText;
        else if (head.includes("nombre") || head.includes("first name")) nuevaFila[c] = usr.givenName;
        else if (head.includes("apellido") || head.includes("last name")) nuevaFila[c] = usr.familyName;
        else if (c === idxLicencia) nuevaFila[c] = usr.licencia;
        else if (c === idxOu) nuevaFila[c] = usr.ou; 
      }
      hoja.appendRow(nuevaFila);
      filasInsertadas++;
      if (dominio === "gigas.com") {
        intrusosGigas.push(emailIntruso + " (Licencia: " + usr.licencia + " | OU: " + usr.ou + ")");
      }
    }
  }

  try {
    if (hojaDash) {
      var dashData = hojaDash.getRange("A3:A12").getValues();
      for (var idxD = 0; idxD < dashData.length; idxD++) {
        var nombreProd = String(dashData[idxD][0]).toLowerCase();
        var asignadasReales = null;
        if (nombreProd.includes("starter")) asignadasReales = conteoLicencias["1010020027"] || 0; 
        else if (nombreProd.includes("standard") && !nombreProd.includes("business")) asignadasReales = conteoLicencias["1010020026"] || 0; 
        else if (nombreProd.includes("plus")) asignadasReales = conteoLicencias["1010020020"] || 0; 

        if (asignadasReales !== null) {
          hojaDash.getRange(3 + idxD, 3).setValue("Asignadas: " + asignadasReales);
        }
      }
    }
  } catch (eLic) { }

  return msgListas + "🔄 Sincronización Completada. Bajas purgadas: " + filasBorradas + ". Altas insertadas: " + filasInsertadas + ". Cambios: " + contadorCambiosEstado + " estados, " + contadorCambiosLicencia + " licencias, " + contadorCambiosOu + " OUs.";
}

function motorSincronizarListasWorkspace(customerId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaListas = ss.getSheetByName("Inventario Listas Workspace") || ss.getSheetByName("Listas");
  if (!hojaListas) return "⚠️ Pestaña de listas no encontrada. ";

  var pageToken;
  var todasLasListas = [];
  do {
    var resGrupos = AdminDirectory.Groups.list({ customer: customerId, maxResults: 200, pageToken: pageToken });
    if (resGrupos.groups) {
      for (var i = 0; i < resGrupos.groups.length; i++) {
        todasLasListas.push(resGrupos.groups[i]);
      }
    }
    pageToken = resGrupos.nextPageToken;
  } while (pageToken);

  var datosExcel = [];
  var maxColumnasRequeridas = 12; // De la columna A a la L van exactamente 12 columnas

  for (var g = 0; g < todasLasListas.length; g++) {
    var grupo = todasLasListas[g];
    var emailGrupo = grupo.email;
    var nombreGrupo = grupo.name || "";
    var owner = "";
    var miembrosNormales = [];
    var volumen = 0;
    var pageTokenMiembros;

    try {
      do {
        var resMiembros = AdminDirectory.Members.list(emailGrupo, { maxResults: 200, pageToken: pageTokenMiembros });
        if (resMiembros.members) {
          for (var m = 0; m < resMiembros.members.length; m++) {
            var miembro = resMiembros.members[m];
            volumen++;
            if (miembro.role === "OWNER") owner = miembro.email;
            else miembrosNormales.push(miembro.email);
          }
        }
        pageTokenMiembros = resMiembros.nextPageToken;
      } while (pageTokenMiembros);
    } catch(e) {}

    // Construimos la fila base con los metadatos (Columnas A, B, C, D)
    var filaGrupo = [nombreGrupo, emailGrupo, volumen + " miembros", owner];
    
    // 📦 REPARTO EN BLOQUES (Diseño Gigas): Troceamos los miembros en grupos de 20 para repartirlos de la E a la L
    var chunkSize = 20;
    for (var i = 0; i < miembrosNormales.length; i += chunkSize) {
      var chunk = miembrosNormales.slice(i, i + chunkSize);
      filaGrupo.push(chunk.join(", "));
    }
    
    // Rellenamos con celdas vacías hasta llegar al menos a la columna L (12 columnas) para que la matriz sea rectangular
    while (filaGrupo.length < maxColumnasRequeridas) {
      filaGrupo.push("");
    }
    
    // Si una lista hiper-masiva requiriera más de la columna L, actualizamos el máximo dinámicamente
    if (filaGrupo.length > maxColumnasRequeridas) {
      maxColumnasRequeridas = filaGrupo.length;
    }

    datosExcel.push(filaGrupo);
  }

  // Normalizamos el ancho de absolutamente todas las filas de la matriz antes de escribir
  for (var r = 0; r < datosExcel.length; r++) {
    while (datosExcel[r].length < maxColumnasRequeridas) {
      datosExcel[r].push("");
    }
  }

  var ultimaFila = hojaListas.getLastRow();
  var ultimaColumna = hojaListas.getLastColumn();
  
  // Limpiamos todo el contenido antiguo de los datos (fila 2 hacia abajo) respetando las columnas reales que tengáis
  if (ultimaFila > 1 && ultimaColumna > 0) {
    hojaListas.getRange(2, 1, ultimaFila - 1, ultimaColumna).clearContent().setBackground(null);
  }
  
  // Volcamos la nueva matriz limpia y perfectamente distribuida en bloques
  if (datosExcel.length > 0) {
    hojaListas.getRange(2, 1, datosExcel.length, maxColumnasRequeridas).setValues(datosExcel);
  }
  
  // 🔐 FECHA OCULTA: Guardamos el timestamp en Z1 sin machacar vuestros títulos de cabecera (E1, F1...)
  hojaListas.getRange("Z1").setValue(new Date());
  
  return "📋 Listas (" + todasLasListas.length + ") actualizadas. ";
}

// =========================================================
// NUEVAS FUNCIONES Y MODIFICACIONES: MÓDULO LISTAS
// =========================================================

function webRenombrarLista(emailAntiguo, emailNuevo) {
  verificarPermisoEjecucion();
  var oldClean = emailAntiguo.trim().toLowerCase();
  var newClean = emailNuevo.trim().toLowerCase();

  if (!oldClean || !newClean) throw new Error("Faltan datos para renombrar la lista.");

  try {
    // 1. Actualizar en Google Workspace
    AdminDirectory.Groups.update({ email: newClean }, oldClean);
    registrarEnHistorial("RENOMBRAR LISTA", "De: " + oldClean + " a: " + newClean);

    // 2. Actualizar en la base de datos de Excel
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Inventario Listas Workspace") || ss.getSheetByName("Listas");
    var msgExcel = " (⚠️ No se encontró la pestaña en Excel para actualizar).";

    if (hoja) {
      var datos = hoja.getDataRange().getValues();
      for (var f = datos.length - 1; f >= 1; f--) {
        if (String(datos[f][1]).toLowerCase().trim() === oldClean) {
          hoja.getRange(f + 1, 2).setValue(newClean); // Actualiza la columna B
          hoja.getRange(f + 1, 2).setBackground("#e2f0d9"); // Pinta de verde para feedback visual
          msgExcel = " y Excel sincronizado.";
          break;
        }
      }
    }
    return "✅ Lista renombrada a " + newClean + msgExcel;
  } catch (e) {
    throw new Error("Fallo al renombrar en Workspace: " + e.message);
  }
}

function webExpulsionMasivaListas(userEmail) {
  verificarPermisoEjecucion();
  var uClean = userEmail.trim().toLowerCase();
  if (!uClean) throw new Error("Debes especificar un usuario.");

  var gruposEliminados = [];
  var pageToken;

  try {
    // 1. Obtener todos los grupos a los que pertenece el usuario
    do {
      var resGrupos = AdminDirectory.Groups.list({ userKey: uClean, maxResults: 200, pageToken: pageToken });
      if (resGrupos.groups) {
        for (var i = 0; i < resGrupos.groups.length; i++) {
          var gEmail = resGrupos.groups[i].email;
          
          // 2. Ejecutar la expulsión iterativa
          try {
            AdminDirectory.Members.remove(gEmail, uClean);
            gruposEliminados.push(gEmail);
          } catch (eRemove) {
            // Ignoramos errores individuales si la API falla en un grupo concreto (ej: ya no estaba)
          }
        }
      }
      pageToken = resGrupos.nextPageToken;
    } while (pageToken);

    if (gruposEliminados.length > 0) {
      registrarEnHistorial("PURGA MASIVA LISTAS", "Usuario: " + uClean + " expulsado de: " + gruposEliminados.length + " listas.");
      return "🗑️ Purgado con éxito. " + uClean + " ha sido expulsado de " + gruposEliminados.length + " listas de distribución.";
    } else {
      return "✅ El usuario no pertenecía a ninguna lista de distribución en este momento.";
    }
  } catch (e) {
    throw new Error("Fallo en la purga masiva: " + e.message);
  }
}

// ⚠️ REEMPLAZA TU FUNCIÓN ACTUAL webAñadirMiembroLista POR ESTA:
function webAñadirMiembroLista(g, u) { 
  verificarPermisoEjecucion();
  var grupoLimpio = g.trim().toLowerCase();
  
  // Soporta separadores por coma, punto y coma, o saltos de línea
  var usuariosRaw = u.split(/[\n,;]+/);
  var añadidos = 0;
  var ignorados = 0;
  var errores = [];

  for (var i = 0; i < usuariosRaw.length; i++) {
    var usuarioLimpio = usuariosRaw[i].trim().toLowerCase();
    
    // Filtro para quitar espacios vacíos o cadenas sin formato de correo
    if (usuarioLimpio === "" || !usuarioLimpio.includes("@")) continue;

    try {
      AdminDirectory.Members.insert({ email: usuarioLimpio, role: "MEMBER" }, grupoLimpio);
      añadidos++;
    } catch (e) {
      // Si Google dice que ya existe, no lanzamos error rojo, lo contamos como ignorado
      if (e.message.includes("Member already exists")) {
        ignorados++;
      } else {
        errores.push(usuarioLimpio + " (" + e.message + ")");
      }
    }
  }

  if (añadidos > 0) {
    registrarEnHistorial("AÑADIR A LISTA (BLOQUE)", "Añadidos " + añadidos + " usuarios a la lista: " + grupoLimpio);
  }

  var mensajeFinal = "";
  if (añadidos > 0) mensajeFinal += "➕ " + añadidos + " añadido(s) correctamente. ";
  if (ignorados > 0) mensajeFinal += "✅ " + ignorados + " ya estaban en la lista. ";
  
  if (errores.length > 0) {
    throw new Error(mensajeFinal + "⚠️ Fallaron: " + errores.join(" | "));
  }
  
  if (añadidos === 0 && ignorados === 0) {
     throw new Error("No se detectó ningún correo válido para añadir. Recuerda separar por comas.");
  }

  return mensajeFinal.trim();
}

function webBorrarMiembroLista(g, u) { 
  verificarPermisoEjecucion();
  var grupoLimpio = g.trim().toLowerCase();
  var usuarioLimpio = u.trim().toLowerCase();

  try {
    AdminDirectory.Members.remove(grupoLimpio, usuarioLimpio); 
    registrarEnHistorial("BORRAR DE LISTA", "Usuario: " + usuarioLimpio + " eliminado de la lista: " + grupoLimpio);
    return "🗑️ Eliminado correctamente."; 
  } catch (e) {
    // Si Google dice que no lo encuentra, es que ya estaba fuera. Lo damos por bueno.
    if (e.message.includes("Resource Not Found")) {
      return "✅ El usuario ya no estaba en la lista (eliminado previamente).";
    }
    throw new Error("Fallo al eliminar de Workspace: " + e.message);
  }
}

function webCrearUsuarioAvanzado(datos) {
  verificarPermisoEjecucion();
  var aliasLimpio = datos.alias.trim().toLowerCase().replace(/\s+/g, '');
  var emailCompleto = aliasLimpio + "@" + datos.dominio.trim().toLowerCase();
  var passwordGenerada = Math.random().toString(36).slice(-8) + "Gigs26!";
  var given = datos.nombre ? datos.nombre.trim() : aliasLimpio;
  var family = datos.apellido ? datos.apellido.trim() : datos.tipo;
  var ouDestino = datos.ou ? datos.ou.trim() : "/";
  var nuevoUsuario = { primaryEmail: emailCompleto, name: { givenName: given, familyName: family }, password: passwordGenerada, changePasswordAtNextLogin: true, orgUnitPath: ouDestino };
  if (datos.manager) nuevoUsuario.relations = [{ type: "manager", value: datos.manager.trim() }];
  
  try {
    AdminDirectory.Users.insert(nuevoUsuario);
    registrarEnHistorial("ALTA USUARIO", "Cuenta: " + emailCompleto + " | Tipo: " + datos.tipo + " | Licencia asignada: " + datos.licencia + " | OU Destino: " + ouDestino);
    
    var msgLicencia = "✅ Licencia asignada correctamente";
    
    // 🌟 ASIGNACIÓN DE LICENCIA (CON BUCLE ANTI-PROPAGACIÓN)
    if (datos.licencia && datos.licencia !== "Cloud Identity Free") {
      try {
        var skuMap = {
          "Google Workspace Business Starter": "1010020027",
          "Google Workspace Enterprise Standard": "1010020026",
          "Google Workspace Enterprise Plus": "1010020020"
        };
        
        var skuId = skuMap[datos.licencia];
        if (!skuId) {
          throw new Error("El nombre de la licencia no coincide con el skuMap: " + datos.licencia);
        }

        var asignado = false;
        var ultimoError = "";
        
        // Lo intenta 4 veces, esperando 3 segundos entre cada intento
        for (var i = 0; i < 4; i++) {
          Utilities.sleep(3000); 
          try {
            AdminLicenseManager.LicenseAssignments.insert({ userId: emailCompleto }, "Google-Apps", skuId);
            asignado = true;
            break; // Si se asigna con éxito, rompemos el bucle
          } catch(err) {
            ultimoError = err.message;
          }
        }
        
        if (!asignado) {
          msgLicencia = "⚠️ Cuenta creada, pero falló la asignación tras 4 intentos: " + ultimoError;
        }

      } catch(eLic) {
        msgLicencia = "⚠️ Error crítico de licencia: " + eLic.message;
      }
    } else {
      msgLicencia = "☁️ Se asignó Cloud Identity Free (Por defecto)";
    }
    
    // 🌟 INSERCIÓN EN EL EXCEL
    var libro = SpreadsheetApp.getActiveSpreadsheet(); 
    var hojaDestino = libro.getSheetByName("Hoja Maestra (657)") || libro.getSheetByName("HOJA MAESTRA 657") || libro.getSheets()[0];
    
    if (hojaDestino) {
      var allData = hojaDestino.getDataRange().getValues();
      var cabeceras = allData[0]; 
      var nuevaFila = new Array(cabeceras.length).fill("");
      var colOrigenIdx = -1;

      var origenText = "";
      var palabraClave = "";
      if (datos.tipo === "Nominativa") { origenText = "Cuentas nominativas de Gigas"; palabraClave = "nominativa"; }
      else if (datos.tipo === "TPartner") { origenText = "Cuentas TPartner"; palabraClave = "tpartner"; }
      else if (datos.tipo === "Kayako") { origenText = "Cuentas Kayako"; palabraClave = "kayako"; }
      else if (datos.tipo === "Externa") { origenText = "Cuentas asesorgigas"; palabraClave = "asesorgigas"; }
      else if (datos.tipo === "Onmovil") { origenText = "Cuentas Onmovil"; palabraClave = "onmovil"; }
      else if (datos.tipo === "Servicio") { origenText = "Cuentas de servicio Gigas"; palabraClave = "servicio"; }
      else { origenText = datos.tipo; palabraClave = datos.tipo.toLowerCase(); }

      for (var c = 0; c < cabeceras.length; c++) {
        var head = String(cabeceras[c]).toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        if (head.includes("email") || head.includes("correo")) nuevaFila[c] = emailCompleto;
        else if (head.includes("manager") || head.includes("responsable")) nuevaFila[c] = datos.manager.trim();
        else if (head.includes("status") || head.includes("estado")) nuevaFila[c] = "Active"; 
        else if (head.includes("origen") || head.includes("hoja de origen")) { nuevaFila[c] = origenText; colOrigenIdx = c; }
        else if (head.includes("first name") || head.includes("firstname") || head === "nombre") nuevaFila[c] = given;
        else if (head.includes("last name") || head.includes("lastname") || head === "apellido" || head.includes("apellidos")) nuevaFila[c] = family;
        else if (head.includes("motivo") || head.includes("servicio") || head.includes("dedicacion")) nuevaFila[c] = datos.motivo || "";
        else if (head === "ou" || head.includes("unidad") || head.includes("orgunit")) nuevaFila[c] = ouDestino;
        else if (head.includes("licencia")) nuevaFila[c] = datos.licencia;
      }

      var filaInsercion = -1;
      if (colOrigenIdx !== -1) {
        for (var f = allData.length - 1; f >= 1; f--) {
          if (String(allData[f][colOrigenIdx]).toLowerCase().includes(palabraClave)) { filaInsercion = f + 1; break; }
        }
      }
      
      if (filaInsercion !== -1) {
        hojaDestino.insertRowAfter(filaInsercion);
        hojaDestino.getRange(filaInsercion + 1, 1, 1, cabeceras.length).setValues([nuevaFila]);
        var rule = hojaDestino.getRange(filaInsercion, 1, 1, cabeceras.length).getDataValidations();
        hojaDestino.getRange(filaInsercion + 1, 1, 1, cabeceras.length).setDataValidations(rule);
        hojaDestino.getRange(filaInsercion, 1, 1, cabeceras.length).copyTo(hojaDestino.getRange(filaInsercion + 1, 1, 1, cabeceras.length), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      } else {
        hojaDestino.appendRow(nuevaFila);
      }
    }
    return { exito: true, email: emailCompleto, pass: passwordGenerada, msgLic: msgLicencia, hoja: hojaDestino.getName() };
  } catch (error) { return { exito: false, error: error.message }; }
}

function webCrearListaDistribucion(datos) {
  try {
    var emailGrupo = datos.alias.trim().toLowerCase() + "@" + datos.dominio;
    var nuevoGrupo = { email: emailGrupo, name: datos.nombre, description: "Lista gestionada desde el Panel de Identidades" };
    AdminDirectory.Groups.insert(nuevoGrupo);
    registrarEnHistorial("ALTA LISTA", "Lista creada: " + emailGrupo + " | Owner inicial: " + (datos.owner ? datos.owner : "Ninguno"));
    
    var arrayMiembros = []; 
    var totalVolumen = 0;
    var ownerLimpio = datos.owner ? datos.owner.trim().toLowerCase() : "";
    
    if (ownerLimpio !== "") { 
      try { 
        AdminDirectory.Members.insert({ email: ownerLimpio, role: "OWNER" }, emailGrupo); 
        totalVolumen++;
      } catch(e) { } 
    }
    
    if (datos.miembros && datos.miembros.trim() !== "") {
      var rawMiembros = datos.miembros.split(/[\n,]+/);
      for (var i = 0; i < rawMiembros.length; i++) {
        var m = rawMiembros[i].trim().toLowerCase();
        if (m !== "" && m.indexOf("@") !== -1 && m !== ownerLimpio) arrayMiembros.push(m);
      }
      for (var j = 0; j < arrayMiembros.length; j++) { 
        try { 
          AdminDirectory.Members.insert({ email: arrayMiembros[j], role: "MEMBER" }, emailGrupo);
          totalVolumen++; 
        } catch(e) { } 
      }
    }
    
    // 👇 AQUÍ ESTÁ LA MAGIA CORREGIDA 👇
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Inventario Listas Workspace") || ss.getSheetByName("Listas"); 
    
    if (hoja) { 
      hoja.appendRow([datos.nombre, emailGrupo, totalVolumen + " miembros", ownerLimpio, arrayMiembros.join(", ")]);
      hoja.getRange(hoja.getLastRow(), 1, 1, 5).setBackground("#e2f0d9"); 
    }
    
    return { exito: true, msg: "Lista " + emailGrupo + " creada.\nTotal integrantes: " + totalVolumen };
  } catch (error) { 
    return { exito: false, error: "Error en Workspace: " + error.message };
  }
}

function webObtenerOUs() {
  var pageToken;
  var ous = ["/"]; 
  try {
    var listaUsuarios = AdminDirectory.Users.list({domain: "gigas.com", maxResults: 1});
    var customerIdReal = listaUsuarios.users[0].customerId;
    
    do {
      var response = AdminDirectory.Orgunits.list(customerIdReal, { type: 'all', pageToken: pageToken });
      if (response.organizationUnits) {
        for (var i = 0; i < response.organizationUnits.length; i++) {
          if (response.organizationUnits[i].orgUnitPath !== "/") {
            ous.push(response.organizationUnits[i].orgUnitPath);
          }
        }
      }
      pageToken = response.nextPageToken;
    } while (pageToken);
    return ous.sort();
  } catch(e) {
    Logger.log("Error al obtener OUs: " + e.message);
    return ["/"];
  }
}

function webMoverUsuarioOU(userEmail, nuevaOU) {
  verificarPermisoEjecucion();
  if (!userEmail || !nuevaOU) throw new Error("Faltan datos para mover la OU.");
  
  try {
    AdminDirectory.Users.update({ orgUnitPath: nuevaOU }, userEmail);
    registrarEnHistorial("MOVER OU", "Usuario: " + userEmail + " movido a: " + nuevaOU);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Hoja Maestra (657)");
    var msgExcel = " (⚠️ Añade una columna llamada 'OU' en tu Excel para que se sincronice automáticamente).";
    
    if (hoja) {
      var data = hoja.getDataRange().getValues();
      var headers = data[0].map(function(h) { 
        return String(h).toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); 
      });
      
      var ouCol = headers.indexOf("ou");
      if (ouCol === -1) ouCol = headers.indexOf("unidad organizativa");
      if (ouCol === -1) ouCol = headers.indexOf("orgunit");

      if (ouCol !== -1) {
        for (var r = 1; r < data.length; r++) {
          var emailCell = String(data[r][headers.indexOf("email address")] || data[r][headers.indexOf("email")] || "").toLowerCase().trim();
          if (emailCell === userEmail.toLowerCase().trim()) {
            hoja.getRange(r + 1, ouCol + 1).setValue(nuevaOU);
            hoja.getRange(r + 1, ouCol + 1).setBackground("#e2f0d9");
            msgExcel = " y Excel sincronizado.";
            break;
          }
        }
      }
    }
    
    return "✅ " + userEmail + " movido a " + nuevaOU + msgExcel;
  } catch (e) {
    throw new Error("Fallo al mover OU: " + e.message);
  }
}

function webBorrarListaTotal(emailGrupo) {
  verificarPermisoEjecucion();
  var grupoTarget = String(emailGrupo).trim().toLowerCase();
  try {
    AdminDirectory.Groups.remove(grupoTarget);
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hojaListas = libro.getSheetByName("Inventario Listas Workspace") || libro.getSheetByName("Listas");
    var borradoExcel = "No localizada en tu hoja de cálculo.";
    if (hojaListas) {
      var datos = hojaListas.getDataRange().getValues();
      for (var f = datos.length - 1; f >= 0; f--) {
        if (datos[f][1] && datos[f][1].toString().trim().toLowerCase() === grupoTarget) { 
          hojaListas.deleteRow(f + 1);
          borradoExcel = "Fila purgada con éxito del Excel."; break; 
        }
      }
    }
    return "🗑️ LISTA DESTRUIDA CON ÉXITO.\n\n• Workspace: El grupo " + grupoTarget + " ha sido eliminado.\n• Base de Datos: " + borradoExcel;
  } catch (err) { throw new Error("Fallo al eliminar la lista de Workspace: " + err.message); }
}

function webEscanearCuentasServicio() {
  verificarPermisoEjecucion(); var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("Cuentas de servicio Gigas (139)") || libro.getSheetByName("Cuentas de servicio gigas (139)");
  var datos = hoja.getDataRange().getValues(); var cabecera = datos[0];
  var idxEmail = cabecera.indexOf("Email Address [Required]");
  var idxResponsable = cabecera.indexOf("Correo del responsable");
  var procesados = 0;
  for (var i = 1; i < datos.length; i++) {
    var email = datos[i][idxEmail];
    if (email && email.toString().includes("@")) {
      var adminEncontrado = "no tiene";
      try { var u = AdminDirectory.Users.get(email, {projection: "full"}); if (u.relations) { for (var j = 0; j < u.relations.length; j++) { if (u.relations[j].type === "manager") { adminEncontrado = u.relations[j].value.trim(); break; } } } } catch (e) { adminEncontrado = "Error"; }
      hoja.getRange(i + 1, idxResponsable + 1).setValue(adminEncontrado); procesados++;
    }
  }
  return "✅ Servicios revisados: " + procesados;
}

function webAuditarEnviosAnterioresServicios() {
  verificarPermisoEjecucion(); var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("Cuentas de servicio Gigas (139)") || libro.getSheetByName("Cuentas de servicio gigas (139)");
  var datos = hoja.getDataRange().getValues();
  var cabecera = datos[0];
  var idxEmail = cabecera.indexOf("Email Address [Required]"); var idxMotivo = cabecera.indexOf("Motivo de la cuenta");
  var hilos = GmailApp.search('in:sent subject:"[REF-SERVICIO]"'); var marcados = 0;
  for (var h = 0; h < hilos.length; h++) {
    var coincidencia = hilos[h].getFirstMessageSubject().match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    if (coincidencia) {
      var emailEnviado = coincidencia[1].trim().toLowerCase();
      for (var i = 1; i < datos.length; i++) { if (datos[i][idxEmail] && datos[i][idxEmail].toString().trim().toLowerCase() === emailEnviado) { hoja.getRange(i + 1, idxMotivo + 1).setValue("⏳ Enviado (Esperando respuesta)"); marcados++; break; } }
    }
  }
  return "✅ Auditadas: " + marcados;
}

function webEnviarCampanaServicios() {
  verificarPermisoEjecucion(); var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName("Cuentas de servicio Gigas (139)") || libro.getSheetByName("Cuentas de servicio gigas (139)");
  var datos = hoja.getDataRange().getValues(); var cabecera = datos[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var idxEmail = cabecera.indexOf("email address [required]"); var idxResponsable = cabecera.indexOf("correo del responsable"); var idxMotivo = cabecera.indexOf("motivo de la cuenta");
  var enviados = 0; var fechaHoy = Utilities.formatDate(new Date(), "Europe/Madrid", "dd/MM/yyyy");
  for (var i = 1; i < datos.length; i++) {
    var emailC = datos[i][idxEmail] ? datos[i][idxEmail].toString().trim() : "";
    var estadoResp = datos[i][idxResponsable] ? datos[i][idxResponsable].toString().trim() : "";
    if (datos[i][idxMotivo] && datos[i][idxMotivo].toString() !== "") continue;
    var dest = estadoResp.includes("@") ? estadoResp : emailC;
    if (dest.includes("@")) { MailApp.sendEmail(dest, "[Auditoría IT] Cuenta: " + emailC + " [REF-SERVICIO]", "Rellenar:\nMotivo de la cuenta: ");
    hoja.getRange(i + 1, idxMotivo + 1).setValue("⏳ Enviado el " + fechaHoy); enviados++;
    }
  }
  return "✅ Enviados: " + enviados;
}

function webEnviarRecordatorioCampanaServicios() {
  verificarPermisoEjecucion(); var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaMaestra = libro.getSheetByName("Hoja Maestra (657)");
  var datosMae = hojaMaestra.getDataRange().getValues(); var cabeceraMae = datosMae[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var idxEmailMae = cabeceraMae.indexOf("email address [required]") !== -1 ? cabeceraMae.indexOf("email address [required]") : cabeceraMae.indexOf("email address");
  var idxOrigenMae = cabeceraMae.indexOf("hoja de origen") !== -1 ? cabeceraMae.indexOf("hoja de origen") : cabeceraMae.indexOf("origen");
  var idxManagerMae = cabeceraMae.indexOf("manager") !== -1 ? cabeceraMae.indexOf("manager") : cabeceraMae.indexOf("responsable");
  var idxMotivoMae = cabeceraMae.indexOf("motivo de la cuenta");
  var recordatoriosEnviados = 0; var fechaHoy = Utilities.formatDate(new Date(), "Europe/Madrid", "dd/MM/yyyy");
  for (var i = 1; i < datosMae.length; i++) {
    var emailC = datosMae[i][idxEmailMae] ? datosMae[i][idxEmailMae].toString().trim() : "";
    var origenC = datosMae[i][idxOrigenMae] ? datosMae[i][idxOrigenMae].toString().trim().toLowerCase() : "";
    var motivoC = datosMae[i][idxMotivoMae] ? datosMae[i][idxMotivoMae].toString().trim().toLowerCase() : "";
    if (origenC.includes("servicio") && (motivoC.includes("enviado") || motivoC.includes("esperando"))) {
      var dest = datosMae[i][idxManagerMae].toString().includes("@") ? datosMae[i][idxManagerMae].toString().trim() : emailC;
      var cuerpoHTML = "<div style='font-family: Arial; padding:20px;'><div style='background:#fef2f2; color:#991b1b; padding:12px; border-left:4px solid #ef4444;'><strong>🚨 NOTA DE SEGURIDAD:</strong> Correo de IT legítimo, no es Phishing.</div><p>Revisión de cuenta: <strong>" + emailC + "</strong></p><p>Responder indicando:</p><b>Correo del responsable: <br>Motivo de la cuenta: </b><br><br>Fernando Alcalá<br>Departamento IT</div>";
      MailApp.sendEmail({ to: dest, subject: "[Recordatorio] Auditoría Workspace: " + emailC + " [REF-SERVICIO]", htmlBody: cuerpoHTML });
      hojaMaestra.getRange(i + 1, idxMotivoMae + 1).setValue("⏳ Recordatorio enviado el " + fechaHoy); recordatoriosEnviados++;
    }
  }
  return "✅ Recordatorio enviado a " + recordatoriosEnviados + " cuentas.";
}

function webProcesarRespuestasServicios() {
  verificarPermisoEjecucion();
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaServicios = libro.getSheetByName("Cuentas de servicio Gigas (139)") || libro.getSheetByName("Cuentas de servicio gigas (139)");
  var hojaMaestra = libro.getSheetByName("Hoja Maestra (657)");
  var datosServ = hojaServicios.getDataRange().getValues(); var cabeceraServ = datosServ[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var idxEmailServ = cabeceraServ.indexOf("email address [required]"); var idxMotivoServ = cabeceraServ.indexOf("motivo de la cuenta");
  var datosMae = hojaMaestra ? hojaMaestra.getDataRange().getValues() : [];
  var hilos = GmailApp.search('subject:"[REF-SERVICIO]" is:unread'); var procesados = 0;
  for (var h = 0; h < hilos.length; h++) {
    var mensajes = hilos[h].getMessages();
    var hiloProcesado = false;
    for (var m = 0; m < mensajes.length; m++) {
      if (mensajes[m].isUnread()) {
        var asunto = mensajes[m].getSubject();
        var cuerpo = mensajes[m].getPlainBody();
        var coincidenciaCuenta = asunto.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/); if (!coincidenciaCuenta) continue;
        var emailServicioTarget = coincidenciaCuenta[1].trim().toLowerCase();
        var matchMotivo = cuerpo.match(/Motivo de la cuenta:\s*([^\r\n]+)/i); var nuevoMotivo = matchMotivo ? matchMotivo[1].trim() : null;
        if (nuevoMotivo) {
          for (var i = 1; i < datosServ.length; i++) { if (datosServ[i][idxEmailServ] && datosServ[i][idxEmailServ].toString().trim().toLowerCase() === emailServicioTarget) { hojaServicios.getRange(i + 1, idxMotivoServ + 1).setValue(nuevoMotivo).setBackground("#e2f0d9"); hiloProcesado = true; break; } }
          if (hojaMaestra) { for (var j = 1; j < datosMae.length; j++) { if (datosMae[j][0] && datosMae[j][0].toString().trim().toLowerCase() === emailServicioTarget) { hojaMaestra.getRange(j + 1, 7).setValue(nuevoMotivo).setBackground("#e2f0d9"); hiloProcesado = true; break; } } }
        }
        mensajes[m].markRead();
      }
    }
    if (hiloProcesado) procesados++;
  }
  return "Respuestas procesadas con éxito para " + procesados + " cuentas.";
}

function webEscanearManagersMaestra() {
  verificarPermisoEjecucion();
  var libro = SpreadsheetApp.getActiveSpreadsheet(); var hoja = libro.getSheetByName("Hoja Maestra (657)");
  var datos = hoja.getDataRange().getValues(); var cabecera = datos[0];
  var idxEmail = 0; var idxManager = 5;
  var mapaUsuarios = {}; var pageToken;
  do {
    var response = AdminDirectory.Users.list({ customer: 'my_customer', maxResults: 500, pageToken: pageToken });
    if (response.users) {
      for (var i = 0; i < response.users.length; i++) {
        var mng = "no tiene";
        if (response.users[i].relations) { for (var j=0; j<response.users[i].relations.length; j++) { if (response.users[i].relations[j].type === "manager") { mng = response.users[i].relations[j].value.trim(); break; } } }
        mapaUsuarios[response.users[i].primaryEmail.toLowerCase()] = mng;
      }
    }
    pageToken = response.nextPageToken;
  } while (pageToken);
  for (var i = 1; i < datos.length; i++) {
    var email = datos[i][idxEmail] ? String(datos[i][idxEmail]).trim().toLowerCase() : "";
    if (email.includes("@")) { hoja.getRange(i + 1, idxManager + 1).setValue(mapaUsuarios[email] || "no tiene"); }
  }
  return "✅ Managers escaneados.";
}

function webGenerarCodigoBypass2FA(userEmail) {
  verificarPermisoEjecucion();
  if (!userEmail) throw new Error("Selecciona un usuario primero.");
  
  try {
    AdminDirectory.VerificationCodes.generate(userEmail);
    var respuesta = AdminDirectory.VerificationCodes.list(userEmail);
    
    if (respuesta.items && respuesta.items.length > 0) {
      var codigoSalvavidas = respuesta.items[0].verificationCode;
      registrarEnHistorial("BYPASS 2FA", "Código de respaldo generado para: " + userEmail);
      return "🆘 Bypass listo. Dale este código al empleado para que inicie sesión: " + codigoSalvavidas;
    } else {
      throw new Error("Se generaron los códigos pero la API no los devolvió.");
    }
  } catch (e) {
    throw new Error("Fallo al generar el Bypass 2FA: " + e.message);
  }
}

function registrarEnHistorial(accion, detalle) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaLog = ss.getSheetByName("Log_Historial");
    if (!hojaLog) {
      hojaLog = ss.insertSheet("Log_Historial");
      hojaLog.appendRow(["Fecha y Hora", "Operador (Script)", "Acción Realizada", "Detalles / Parámetros"]);
      hojaLog.getRange("A1:D1").setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
      hojaLog.setFrozenRows(1);
    }
    var fechaHoy = Utilities.formatDate(new Date(), "Europe/Madrid", "dd/MM/yyyy HH:mm:ss");
    var usuarioEjecutor = Session.getActiveUser().getEmail() || "Panel Web (Automático)";
    hojaLog.appendRow([fechaHoy, usuarioEjecutor, accion, detalle]);
    Logger.log("📝 Acción registrada en el historial de Sheets con éxito.");
  } catch (error) {
    Logger.log("❌ Error crítico al escribir en el Log_Historial: " + error.message);
  }
}

function rastrearCuentasSuspendidasGmail() {
  var cuentasSuspendidas = [];
  var mapaFechasSuspension = {};
  var pageToken;
  try {
    var hilosAlertas = GmailApp.search('subject:"Alerta: Usuario suspendido por el administrador"', 0, 100);
    for (var h = 0; h < hilosAlertas.length; h++) {
      var mensajes = hilosAlertas[h].getMessages();
      for (var m = 0; m < mensajes.length; m++) {
        var msg = mensajes[m];
        var cuerpo = msg.getPlainBody();
        var fechaAlerta = Utilities.formatDate(msg.getDate(), "Europe/Madrid", "dd/MM/yyyy");
        var coincidenciaEmail = cuerpo.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
        if (coincidenciaEmail) {
          var emailDetectado = coincidenciaEmail[1].toLowerCase().trim();
          if (!mapaFechasSuspension[emailDetectado]) {
            mapaFechasSuspension[emailDetectado] = fechaAlerta;
          }
        }
      }
    }

    do {
      var response = AdminDirectory.Users.list({ customer: 'my_customer', maxResults: 500, projection: 'full', pageToken: pageToken });
      if (response.users) {
        for (var i = 0; i < response.users.length; i++) {
          var user = response.users[i];
          if (user.suspended) {
            var emailClean = user.primaryEmail.toLowerCase().trim();
            var fechaSuspensionFinal = mapaFechasSuspension[emailClean] || null;
            var valLogin = "Nunca ha iniciado sesión";
            if (user.lastLoginTime && user.lastLoginTime !== "1970-01-01T00:00:00.000Z") {
              var d = new Date(user.lastLoginTime);
              valLogin = ("0" + d.getDate()).slice(-2) + "/" + ("0" + (d.getMonth() + 1)).slice(-2) + "/" + d.getFullYear();
            }
            var motivo = user.suspensionReason || "Desconocido";
            if (motivo === "ADMIN") motivo = "Suspendida por IT";
            else if (motivo === "WEB_LOGIN_REQUIRED") motivo = "Bloqueo por seguridad";
            else if (motivo === "ABUSE") motivo = "Abuso de red";
            cuentasSuspendidas.push({
              email: user.primaryEmail,
              nombre: user.name ? user.name.fullName : "Sin nombre",
              fechaSuspension: fechaSuspensionFinal,
              ultimoLogin: valLogin,
              motivo: motivo
            });
          }
        }
      }
      pageToken = response.nextPageToken;
    } while (pageToken);
    
    cuentasSuspendidas.sort(function(a, b) { return a.email.localeCompare(b.email); });
    var fechaHoy = Utilities.formatDate(new Date(), "Europe/Madrid", "dd/MM/yyyy");
    var htmlBody = '<div style="font-family: Arial, sans-serif; color: #1e293b; max-width: 600px; margin: 0 auto; line-height: 1.5; font-size: 14px;">';
    htmlBody += '<h2 style="color: #ef4444; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; margin-bottom: 15px; font-size: 18px;">📋 Cuentas Suspendidas Workspace</h2>';
    htmlBody += '<p>Hola Fernando,</p>';
    htmlBody += '<p>Listado consolidado de cuentas suspendidas detectadas en el sistema a fecha de hoy (<b>' + fechaHoy + '</b>):</p>';
    if (cuentasSuspendidas.length > 0) {
      htmlBody += '<div style="background: #fffbeb; border: 1px solid #fef3c7; border-radius: 6px; padding: 12px; margin: 15px 0;">';
      htmlBody += '<p style="margin-top:0; font-weight: bold; color: #b45f06; font-size: 15px;">⚠️ Total cuentas inactivas: ' + cuentasSuspendidas.length + '</p>';
      for (var j = 0; j < cuentasSuspendidas.length; j++) {
        htmlBody += '<div style="background: #ffffff; border: 1px solid #cbd5e1; padding: 12px; margin-bottom: 10px; border-radius: 4px;">';
        htmlBody += '<div style="font-size: 14px; margin-bottom: 8px; word-break: break-all;">👉 <b>' + cuentasSuspendidas[j].email + '</b> <span style="color:#64748b;">(' + cuentasSuspendidas[j].nombre + ')</span></div>';
        htmlBody += '<div style="font-size: 13px; color: #475569;">';
        if (cuentasSuspendidas[j].fechaSuspension) {
          htmlBody += '<div style="margin-bottom: 4px; background-color: #fef2f2; padding: 4px 8px; border-left: 3px solid #ef4444; display: inline-block;"><b>🛑 Fecha Suspensión (Gmail):</b> <span style="color:#b91c1c; font-weight:bold;">' + cuentasSuspendidas[j].fechaSuspension + '</span></div>';
        } else {
          htmlBody += '<div style="margin-bottom: 4px; color: #64748b; font-style: italic;">⚠️ Alerta de Gmail no localizada</div>';
        }
        htmlBody += '<div style="margin-top: 4px; margin-bottom: 4px;"><b>📅 Último Logon:</b> ' + cuentasSuspendidas[j].ultimoLogin + '</div>';
        htmlBody += '<div><b>🔍 Motivo Workspace:</b> ' + cuentasSuspendidas[j].motivo + '</div>';
        htmlBody += '</div></div>';
      }
      htmlBody += '</div>';
    } else {
      htmlBody += '<div style="background: #d1fae5; color: #065f46; border: 1px solid #a7f3d0; border-radius: 6px; padding: 15px; font-weight: bold;">';
      htmlBody += '✅ No hay ninguna cuenta suspendida registrada.</div>';
    }
    
    htmlBody += '<br><hr style="border: 0; border-top: 1px solid #e2e8f0;"><p style="font-size: 11px; color: #94a3b8;">Reporte automático generado por Gigs IT.</p></div>';
    
    // 📊 --- MOTOR AUTOMÁTICO DE EXTRACCIÓN Y GENERACIÓN CSV ---
    var csvContent = "Email;Nombre;Fecha Suspension (Gmail);Ultimo Logon;Motivo Suspension\r\n";
    for (var k = 0; k < cuentasSuspendidas.length; k++) {
      var item = cuentasSuspendidas[k];
      var nombreLimpio = item.nombre.replace(/"/g, '""');
      var fechaBajaLimpia = item.fechaSuspension ? item.fechaSuspension : "No detectada en Gmail";
      csvContent += item.email + ';"' + nombreLimpio + '";' + fechaBajaLimpia + ';' + item.ultimoLogin + ';"' + item.motivo + '"\r\n';
    }
    
    var nombreArchivo = "Reporte_Suspendidas_" + fechaHoy.replace(/\//g, "-") + ".csv";
    var csvBlob = Utilities.newBlob("", "text/csv", nombreArchivo);
    csvBlob.setDataFromString(csvContent, "UTF-8");
    // ---------------------------------------------------------

    MailApp.sendEmail({
      to: "fernando.alcala@gigas.com",
      subject: "📋 [Auditoría] Reporte de Cuentas Suspendidas - " + fechaHoy,
      htmlBody: htmlBody,
      attachments: [csvBlob] // <--- Inyección del adjunto
    });
  } catch (error) {
    MailApp.sendEmail({ to: "fernando.alcala@gigas.com", subject: "⚠️ Error en la ejecución del reporte de suspendidas", body: "Ocurrió un problema: " + error.message });
  }
}

function extraerSuspendidosONICSV() {
  var cuentasSuspendidas = [];
  var mapaFechasSuspension = {};
  var pageToken;
  var wsSkus = {
    "1010020027": "Google Workspace Business Starter",
    "1010020026": "Google Workspace Enterprise Standard",
    "1010020020": "Google Workspace Enterprise Plus",
    "1010310002": "Cloud Identity Free"
  };
  try {
    var hilosAlertas = GmailApp.search('subject:"Alerta: Usuario suspendido por el administrador"', 0, 150);
    for (var h = 0; h < hilosAlertas.length; h++) {
      var mensajes = hilosAlertas[h].getMessages();
      for (var m = 0; m < mensajes.length; m++) {
        var msg = mensajes[m];
        var cuerpo = msg.getPlainBody();
        var fechaAlerta = Utilities.formatDate(msg.getDate(), "Europe/Madrid", "dd/MM/yyyy");
        var coincidenciaEmail = cuerpo.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
        if (coincidenciaEmail) {
          var emailDetectado = coincidenciaEmail[1].toLowerCase().trim();
          if (!mapaFechasSuspension[emailDetectado]) mapaFechasSuspension[emailDetectado] = fechaAlerta;
        }
      }
    }

    do {
      var response = AdminDirectory.Users.list({ customer: 'my_customer', maxResults: 500, projection: 'full', pageToken: pageToken });
      if (response.users) {
        for (var i = 0; i < response.users.length; i++) {
          var user = response.users[i];
          var emailClean = user.primaryEmail.toLowerCase().trim();
          
          if (user.suspended && emailClean.includes("oni.")) {
            var fechaSuspensionFinal = mapaFechasSuspension[emailClean] || "No detectada en Gmail";
            var valLogin = "Nunca ha iniciado sesión";
            if (user.lastLoginTime && user.lastLoginTime !== "1970-01-01T00:00:00.000Z") {
              var d = new Date(user.lastLoginTime);
              valLogin = ("0" + d.getDate()).slice(-2) + "/" + ("0" + (d.getMonth() + 1)).slice(-2) + "/" + d.getFullYear();
            }
            var motivo = user.suspensionReason || "Desconocido";
            if (motivo === "ADMIN") motivo = "Suspendida por IT";
            else if (motivo === "WEB_LOGIN_REQUIRED") motivo = "Bloqueo por seguridad";
            else if (motivo === "ABUSE") motivo = "Abuso de red";
            
            var licenciaDetectada = "Ninguna / Desconocida";
            var skuIds = Object.keys(wsSkus);
            for (var j = 0; j < skuIds.length; j++) {
              try {
                var prodId = (skuIds[j] === "1010310002") ? "Cloudidentity" : "Google-Apps";
                var check = AdminLicenseManager.LicenseAssignments.get(prodId, skuIds[j], emailClean);
                if (check && check.skuId) { licenciaDetectada = wsSkus[skuIds[j]]; break; }
              } catch (e) {}
            }
            cuentasSuspendidas.push({ email: user.primaryEmail, nombre: user.name ? user.name.fullName : "Sin nombre", licencia: licenciaDetectada, fechaSuspension: fechaSuspensionFinal, ultimoLogin: valLogin, motivo: motivo });
          }
        }
      }
      pageToken = response.nextPageToken;
    } while (pageToken);
    
    cuentasSuspendidas.sort(function(a, b) { return a.email.localeCompare(b.email); });
    
    var csvContent = "Email;Nombre;Licencia Actual;Fecha Suspension (Gmail);Ultimo Logon;Motivo Suspension\r\n";
    for (var k = 0; k < cuentasSuspendidas.length; k++) {
      var item = cuentasSuspendidas[k];
      var nombreLimpio = item.nombre.replace(/"/g, '""');
      csvContent += item.email + ';"' + nombreLimpio + '";"' + item.licencia + '";'
      + item.fechaSuspension + ";" + item.ultimoLogin + ';"' + item.motivo + '"\r\n';
    }
    
    var fechaHoy = Utilities.formatDate(new Date(), "Europe/Madrid", "dd/MM/yyyy");
    var nombreArchivo = "Suspendidos_ONI_" + fechaHoy.replace(/\//g, "-") + ".csv";
    var csvBlob = Utilities.newBlob("", "text/csv", nombreArchivo);
    csvBlob.setDataFromString(csvContent, "UTF-8");
    
    MailApp.sendEmail({
      to: "fernando.alcala@gigas.com",
      subject: "📊 [Extracción Única] Cuentas Suspendidas ONI - " + fechaHoy,
      body: "Hola Fernando,\n\nSe ha ejecutado la extracción única solicitada.\n\nTe adjunto en este correo el fichero CSV con el listado consolidado de las " + cuentasSuspendidas.length + " cuentas suspendidas pertenecientes a ONI, mapeadas con sus respectivas licencias y auditoría de accesos.\n\nUn saludo.",
      attachments: [csvBlob]
    });
  } catch (error) {
    MailApp.sendEmail({ to: "fernando.alcala@gigas.com", subject: "⚠️ Error en la extracción única de ONI", body: "Ocurrió un problema al generar el archivo: " + error.message });
  }
}

function webEnviarMailAdvertenciaBaja(...args) {
  var emailEmpleado = "laura.reboredo@gigas.com";
  var emailManager = "fernando.alcala@gigas.com"; 
  var idioma = "es"; 

  try {
    if (args.length > 0) {
      var datosOrigen = args[0];
      if (typeof datosOrigen === 'object' && datosOrigen !== null) {
        emailEmpleado = datosOrigen.empleado || datosOrigen.email || datosOrigen.user || emailEmpleado;
        emailManager = datosOrigen.manager || datosOrigen.correoManager || datosOrigen.responsable || emailManager;
        idioma = datosOrigen.idioma || datosOrigen.lang || idioma;
      } else {
        var correos = args.filter(function(a) { return typeof a === 'string' && a.includes('@'); });
        if (correos.length >= 2) { emailEmpleado = correos[0]; emailManager = correos[1]; } 
        else if (correos.length === 1) { emailManager = correos[0]; }
      }
    }

    emailEmpleado = emailEmpleado.toLowerCase().trim();
    emailManager = emailManager.toLowerCase().trim();
    idioma = idioma.toLowerCase().trim();

    if (emailEmpleado.endsWith(".pt") || emailManager.endsWith(".pt") || idioma === "pt") { idioma = "pt"; }

    var textos = {
      es: {
        asunto: "Plan de Baja de Cuenta de Usuario",
        titulo: "Plan de Baja de Cuenta de Usuario",
        saludo: "Hola,",
        intro: "Dado que el usuario <a href='mailto:" + emailEmpleado + "' style='color: #1976d2; text-decoration: underline;'>" + emailEmpleado + "</a> se va de la compañía, necesitamos proceder a la gestión de su cuenta con el objetivo de <b>optimizar costes y liberar su licencia</b> de Google Workspace.",
        accion: "👉 Por favor, dale a \"Responder\" a este correo, COPIA Y PEGA la opción que elijas y envíanosla:",
        op1_tit: "Opción [ 1 ] ➔ Transferir su Google Drive a mi cuenta",
        op1_txt: "Moveremos de forma automática todos sus archivos en propiedad a tu Google Drive para que no pierdas documentación y cerraremos la cuenta para ahorrar la licencia.",
        op2_tit: "Opción [ 2 ] ➔ Eliminar cuenta y crear un Alias en mi buzón",
        op2_txt: "Se eliminará la cuenta para ahorrar costes y se configurará un alias gratuito en tu buzón (recibirás todo lo que le escriban a él).<br><i style='color: #616161; font-size: 13px; display: inline-block; margin-top: 4px;'>Nota: Con esta opción también se puede transferir su Google Drive. Si lo quieres, dínoslo al copiar y pegar (ej: \"Quiero opción 2 + transferir Drive\").</i>",
        op3_tit: "Opción [ 3 ] ➔ Solicitar Backup completo con Google Takeout",
        op3_txt: "Si quieres asegurar toda su información para tu tranquilidad, te enviaremos las instrucciones paso a paso de cómo descargar una copia completa de todos sus correos y archivos antes de proceder al borrado definitivo.",
        op4_tit: "Opción [ 4 ] ➔ Borrado inmediato (No se requiere salvar nada)",
        op4_txt: "No es necesario guardar archivos, alias ni correos. La cuenta se eliminará por completo de forma inmediata.",
        cierre: "En cuanto recibamos tu respuesta con la opción elegida, el departamento de Sistemas ejecutará los scripts automatizados.",
        firma: "Sistemas Gigas - Control de Licencias e IT"
      },
      pt: {
        asunto: "Plano de Desativação de Conta de Utilizador",
        titulo: "Plano de Desativação de Conta de Utilizador",
        saludo: "Olá,",
        intro: "Dado que o utilizador <a href='mailto:" + emailEmpleado + "' style='color: #1976d2; text-decoration: underline;'>" + emailEmpleado + "</a> vai deixar a empresa, precisamos de proceder à gestão da sua conta com o objetivo de <b>otimizar custos e libertar a sua licença</b> do Google Workspace.",
        accion: "👉 Por favor, clique em \"Responder\" a este e-mail, COPIE E COLE a opção que escolher e envie-nos:",
        op1_tit: "Opção [ 1 ] ➔ Transferir o seu Google Drive para a mi conta",
        op1_txt: "Moveremos de forma automática todos os seus ficheiros em propriedade para o seu Google Drive para que não perca documentação e encerraremos a conta para poupar a licença.",
        op2_tit: "Opção [ 2 ] ➔ Eliminar conta e criar um Alias na minha caixa de correio",
        op2_txt: "A conta será eliminada para reduzir custos e será configurado un alias gratuito na sua caixa de correio (receberá tudo o que lhe enviarem).<br><i style='color: #616161; font-size: 13px; display: inline-block; margin-top: 4px;'>Nota: Com esta opção também é possível transferir o seu Google Drive. Se desejar, informe-nos ao copiar e colar (ex: \"Quero a opção 2 + transferir Drive\").</i>",
        op3_tit: "Opção [ 3 ] ➔ Solicitar Backup completo com o Google Takeout",
        op3_txt: "Se quiser salvaguardar toda a sua informação para sua tranquilidade, enviar-lhe-emos as instruções passo a passo de como descarregar uma cópia completa de todos os seus e-mails e ficheiros antes de proceder à eliminação definitiva.",
        op4_tit: "Opção [ 4 ] ➔ Eliminação imediata (Não é necessário salvaguardar nada)",
        op4_txt: "Não é necessário guardar ficheiros, alias nem e-mails. A conta será totalmente eliminada de forma imediata.",
        cierre: "Assim que recebermos a sua resposta com a opção escolhida, o departamento de Sistemas executará os scripts automatizados.",
        firma: "Sistemas Gigas - Controlo de Licenças e IT"
      }
    };

    var t = textos[idioma] || textos["es"];

    var htmlContent = '<div style="font-family: Arial, sans-serif; max-width: 650px; margin: 0 auto; color: #333333; line-height: 1.5; font-size: 14px; padding: 10px;">';
    htmlContent += '<h2 style="color: #1976d2; font-size: 20px; font-weight: bold; margin-top: 0; margin-bottom: 15px;">' + t.titulo + '</h2>';
    htmlContent += '<p style="margin-bottom: 20px;">' + t.saludo + '</p>';
    htmlContent += '<p style="margin-bottom: 20px;">' + t.intro + '</p>';
    htmlContent += '<div style="background-color: #eef7ff; border-left: 4px solid #1976d2; padding: 12px 16px; border-radius: 4px; margin-bottom: 20px; color: #0d47a1; font-weight: bold;">' + t.accion + '</div>';
    htmlContent += '<div style="background-color: #f8f9fa; border-left: 4px solid #2e7d32; padding: 14px; border-radius: 4px; margin-bottom: 12px;"><strong style="color: #2e7d32; font-size: 14px;">' + t.op1_tit + '</strong><br><span style="color: #424242; display: inline-block; margin-top: 4px;">' + t.op1_txt + '</span></div>';
    htmlContent += '<div style="background-color: #f8f9fa; border-left: 4px solid #ef6c00; padding: 14px; border-radius: 4px; margin-bottom: 12px;"><strong style="color: #ef6c00; font-size: 14px;">' + t.op2_tit + '</strong><br><span style="color: #424242; display: inline-block; margin-top: 4px;">' + t.op2_txt + '</span></div>';
    htmlContent += '<div style="background-color: #f8f9fa; border-left: 4px solid #1565c0; padding: 14px; border-radius: 4px; margin-bottom: 12px;"><strong style="color: #1565c0; font-size: 14px;">' + t.op3_tit + '</strong><br><span style="color: #424242; display: inline-block; margin-top: 4px;">' + t.op3_txt + '</span></div>';
    htmlContent += '<div style="background-color: #f8f9fa; border-left: 4px solid #c62828; padding: 14px; border-radius: 4px; margin-bottom: 20px;"><strong style="color: #c62828; font-size: 14px;">' + t.op4_tit + '</strong><br><span style="color: #424242; display: inline-block; margin-top: 4px;">' + t.op4_txt + '</span></div>';
    htmlContent += '<p style="margin-bottom: 25px;">' + t.cierre + '</p>';
    htmlContent += '<p style="color: #555555; font-weight: bold; margin-bottom: 0;">' + t.firma + '</p></div>';

    MailApp.sendEmail({
      to: emailManager,
      subject: t.asunto,
      htmlBody: htmlContent
    });
    registrarEnHistorial("AVISO DE BAJA", "Empleado tramitado: " + emailEmpleado + " | Manager notificado: " + emailManager + " | Idioma: " + idioma.toUpperCase());
    return "✅ Mail enviado correctamente (" + idioma.toUpperCase() + ").";
  } catch (error) {
    return "❌ Error procesado internamente.";
  }
}

function botonActualizarTodo() {
  actualizarDashboardLicencias_V3(); 
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaInv = libro.getSheetByName("Inventario Completo Workspace");
  if (hojaInv) { hojaInv.hideSheet(); }
}

function actualizarDashboardLicencias_V3() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("DashBoard") || ss.getSheets()[0]; 
  var filaInicio = 2;
  hoja.getRange(filaInicio, 1, 10, 5).clear().setBackground(null).setFontColor(null).setFontLine("none");
  
  var conteo = {};
  var asignadasGemini = 0;
  try {
    var listaUsuarios = AdminDirectory.Users.list({domain: "gigas.com", maxResults: 1});
    var customerIdReal = listaUsuarios.users[0].customerId;
    var tokenApps;
    do {
      var resApps = AdminLicenseManager.LicenseAssignments.listForProduct("Google-Apps", customerIdReal, { maxResults: 500, pageToken: tokenApps });
      if (resApps.items) {
        for (var i = 0; i < resApps.items.length; i++) {
          var sku = resApps.items[i].skuId;
          conteo[sku] = (conteo[sku] || 0) + 1;
        }
      }
      tokenApps = resApps.nextPageToken;
    } while (tokenApps);
    
    var asignadasPlus = conteo["1010020020"] || 0;
    var asignadasStarter = conteo["1010020027"] || 0;
    var asignadasStandard = conteo["1010020026"] || 0;
    
    var tokenGemini;
    do {
      try {
        var resGemini = AdminLicenseManager.LicenseAssignments.listForProduct("101047", customerIdReal, { maxResults: 500, pageToken: tokenGemini });
        if (resGemini.items) {
          for (var i = 0; i < resGemini.items.length; i++) {
             if (resGemini.items[i].skuId === "1010470001") asignadasGemini++;
          }
        }
        tokenGemini = resGemini.nextPageToken;
      } catch(e) { break; } 
    } while (tokenGemini);
  } catch (error) {
    throw new Error("⚠️ Error al conectar con Google Workspace: " + error.message);
  }

  var datosLicencias = [
    ["✨ Gemini Enterprise (versión antigua)", "Activas", "Asignadas: " + asignadasGemini, "Precio del distribuidor", "∞"],
    ["💼 Google Workspace Business Starter", "Activas", "Asignadas: " + asignadasStarter, "Precio del distribuidor", 300],
    ["🚀 Google Workspace Enterprise Plus", "Activas", "Asignadas: " + asignadasPlus, "Precio del distribuidor", 10],
    ["📊 Google Workspace Enterprise Standard", "Activas", "Asignadas: " + asignadasStandard, "Precio del distribuidor", 554]
  ];
  for (var i = 0; i < datosLicencias.length + 1; i++) { hoja.setRowHeight(filaInicio + i, 40); }
  var anchos = [320, 100, 340, 150, 120];
  for (var c = 0; c < anchos.length; c++) { hoja.setColumnWidth(c + 1, anchos[c]); }
  
  hoja.getRange(filaInicio, 1, 1, 5).setValues([["Producto", "Estado", "Licencias asignadas", "Plan de pagos", "Límite contratado"]])
      .setFontWeight("bold").setFontColor("#5f6368").setFontSize(11)
      .setBorder(null, null, true, null, null, null, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
  for (var i = 0; i < datosLicencias.length; i++) {
    var filaActual = filaInicio + 1 + i;
    var rangoFila = hoja.getRange(filaActual, 1, 1, 5);
    rangoFila.setValues([datosLicencias[i]]);
    
    rangoFila.setFontSize(11).setVerticalAlignment("middle").setFontColor("#3c4043")
             .setBorder(null, null, true, null, null, null, "#f1f3f4", SpreadsheetApp.BorderStyle.SOLID);
    hoja.getRange(filaActual, 2).setFontColor("#1e8e3e").setFontWeight("bold"); 
    hoja.getRange(filaActual, 3).setFontColor("#1a73e8");
    hoja.getRange(filaActual, 5).setHorizontalAlignment("center");
    
    rangoFila.setBackground(i % 2 === 1 ? "#f8f9fa" : "#ffffff");
  }
  hoja.getRange(filaInicio, 1, datosLicencias.length + 1, 5).setFontFamily("Arial");
}

function actualizarLogonEnTodoElExcel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nombresHojas = [ "Hoja Maestra (657)", "Cuentas de servicio Gigas (139)", "Cuentas de servicio gigas (139)", "Cuentas Kayako", "Cuentas de ONI (251)" ];
  var mapaLogons = {}; var pageToken;
  do {
    var response = AdminDirectory.Users.list({ customer: 'my_customer', maxResults: 500, projection: 'basic', pageToken: pageToken });
    if (response.users) {
      for (var i = 0; i < response.users.length; i++) {
        var loginTime = response.users[i].lastLoginTime;
        var valorLoginFinal = "";
        if (!loginTime || loginTime === "1970-01-01T00:00:00.000Z") { valorLoginFinal = "Nunca ha iniciado sesión"; } 
        else { var dateObj = new Date(loginTime);
        var dia = ("0" + dateObj.getDate()).slice(-2); var mes = ("0" + (dateObj.getMonth() + 1)).slice(-2); var año = dateObj.getFullYear();
        valorLoginFinal = dia + "/" + mes + "/" + año; }
        mapaLogons[response.users[i].primaryEmail.toLowerCase()] = valorLoginFinal;
      }
    }
    pageToken = response.nextPageToken;
  } while (pageToken);

  var hojasActualizadas = 0;
  for (var h = 0; h < nombresHojas.length; h++) {
    var hoja = ss.getSheetByName(nombresHojas[h]);
    if (!hoja) continue;
    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) continue;
    var cabeceras = datos[0].map(function(head) { return String(head).toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); });
    var idxEmail = cabeceras.indexOf("email address"); if (idxEmail === -1) idxEmail = cabeceras.indexOf("email address [required]");
    if (idxEmail === -1) idxEmail = cabeceras.indexOf("email");
    var idxLogon = cabeceras.indexOf("ultimo inicio de sesion");
    if (idxLogon === -1) idxLogon = cabeceras.indexOf("ultimo logon");
    if (idxEmail === -1 || idxLogon === -1) continue;
    var valoresAEscribir = [];
    for (var i = 1; i < datos.length; i++) {
      var email = String(datos[i][idxEmail]).trim().toLowerCase();
      if (email.includes("@")) valoresAEscribir.push([mapaLogons[email] || "No encontrado en Workspace"]); else valoresAEscribir.push([""]);
    }
    if (valoresAEscribir.length > 0) { hoja.getRange(2, idxLogon + 1, valoresAEscribir.length, 1).setValues(valoresAEscribir); hojasActualizadas++; }
  }
  return "✅ Proceso completado. Se han actualizado las fechas en " + hojasActualizadas + " pestañas diferentes.";
}
// =========================================================
// 🚑 FUNCIONES RESTAURADAS (MÁNAGER, EXTENDIDOS Y LICENCIAS)
// =========================================================

function obtenerDatosExtendidosDeUsuario(userEmail) {
  try {
    var user = AdminDirectory.Users.get(userEmail.toLowerCase().trim(), {projection: 'full'});
    var mng = "";
    if (user.relations) {
      for (var i = 0; i < user.relations.length; i++) {
        if (user.relations[i].type === 'manager') { mng = user.relations[i].value; break; }
      }
    }
    
    var nameMapSku = { 
      "1010020027": "Google Workspace Business Starter", 
      "1010020026": "Google Workspace Enterprise Standard", 
      "1010020020": "Google Workspace Enterprise Plus", 
      "1010310002": "Cloud Identity Free" 
    };
    
    var licenciaEncontrada = "No detectada";
    var skuIds = Object.keys(nameMapSku);
    for (var j = 0; j < skuIds.length; j++) {
      try {
        var prodId = (skuIds[j] === "1010310002") ? "Cloudidentity" : "Google-Apps";
        var check = AdminLicenseManager.LicenseAssignments.get(prodId, skuIds[j], userEmail.toLowerCase().trim());
        if (check && check.skuId) { 
          licenciaEncontrada = nameMapSku[skuIds[j]]; 
          break; 
        }
      } catch(e) {}
    }
    
    var nombreActual = (user.name && user.name.givenName) ? user.name.givenName : "";
    var apellidosActual = (user.name && user.name.familyName) ? user.name.familyName : "";
    var puestoActual = (user.organizations && user.organizations.length > 0) ? (user.organizations[0].title || "") : "";
    var dptoActual = (user.organizations && user.organizations.length > 0) ? (user.organizations[0].department || "") : "";
    var tlfActual = (user.phones && user.phones.length > 0) ? (user.phones[0].value || "") : "";
    var paisActual = (user.addresses && user.addresses.length > 0) ? (user.addresses[0].formatted || user.addresses[0].country || "Spain") : "Spain";
    var emailRecuperacion = user.recoveryEmail || "";
    
    // =========================================================================
    // ⭐ NUEVO: EXTRACCIÓN DE LISTAS DE DISTRIBUCIÓN (GRUPOS DE GOOGLE) ⭐
    // =========================================================================
    var listasUsuario = "";
    try {
      var consultaGrupos = AdminDirectory.Groups.list({userKey: userEmail.toLowerCase().trim()});
      if (consultaGrupos && consultaGrupos.groups) {
        var arrayEmails = [];
        for (var k = 0; k < consultaGrupos.groups.length; k++) {
          arrayEmails.push(consultaGrupos.groups[k].email);
        }
        listasUsuario = arrayEmails.join(", "); // Las separamos por comas
      }
    } catch(eGrupos) {
      listasUsuario = "";
    }
    // =========================================================================
    
    return {
      manager: mng,
      licenciaActiva: licenciaEncontrada,
      ou: user.orgUnitPath || "/",
      nombre: nombreActual,
      apellidos: apellidosActual,
      title: puestoActual,
      department: dptoActual,
      phone: tlfActual,
      country: paisActual,
      recovery: emailRecuperacion,
      // ⭐ EMPACATAMOS LAS LISTAS PARA MANDARLAS A LA WEB ⭐
      grupo: listasUsuario,
      listas: listasUsuario
    };
  } catch(e) { 
    return null; 
  }
}

function webActualizarManager(userEmail, managerEmail) {
  verificarPermisoEjecucion();
  var uClean = userEmail.trim().toLowerCase();
  var mClean = managerEmail.trim().toLowerCase();
  
  try {
    var userUpdate = { relations: [{ type: 'manager', value: mClean }] };
    AdminDirectory.Users.update(userUpdate, uClean);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName("Hoja Maestra (657)");
    var msgExcel = " (⚠️ Manager no actualizado en Excel, no se encontró la columna).";
    
    if (hoja) {
      var data = hoja.getDataRange().getValues();
      var headers = data[0].map(function(h) { return String(h).toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); });
      var idxEmail = headers.indexOf("email address");
      if (idxEmail === -1) idxEmail = headers.indexOf("email");
      var idxManager = headers.indexOf("manager");
      if (idxManager === -1) idxManager = headers.indexOf("responsable");
      
      if (idxEmail > -1 && idxManager > -1) {
        for (var r = 1; r < data.length; r++) {
          if (String(data[r][idxEmail]).toLowerCase().trim() === uClean) {
            hoja.getRange(r + 1, idxManager + 1).setValue(mClean).setBackground("#e2f0d9");
            msgExcel = " y Excel sincronizado.";
            break;
          }
        }
      }
    }
    registrarEnHistorial("ACTUALIZAR MANAGER", "Usuario: " + uClean + " | Nuevo Manager: " + mClean);
    return "✅ Mánager actualizado en Workspace" + msgExcel;
  } catch(e) {
    throw new Error("Fallo al actualizar el mánager: " + e.message);
  }
}

function webCambiarLicencia(userEmail, nuevaLicencia) {
  verificarPermisoEjecucion();
  var uClean = userEmail.trim().toLowerCase();
  
  // 🌟 MAPA CORREGIDO
  var skuMap = {
    "Google Workspace Business Starter": { sku: "1010020027", prod: "Google-Apps" },
    "Google Workspace Enterprise Standard": { sku: "1010020026", prod: "Google-Apps" },
    "Google Workspace Enterprise Plus": { sku: "1010020020", prod: "Google-Apps" },
    "Cloud Identity Free": { sku: "1010310002", prod: "Cloudidentity" }
  };
  
  var config = skuMap[nuevaLicencia];
  if (!config) throw new Error("Licencia no reconocida.");
  
  try {
     // Retirar licencias antiguas
     try { AdminLicenseManager.LicenseAssignments.remove("Google-Apps", "1010020027", uClean); } catch(e){}
     try { AdminLicenseManager.LicenseAssignments.remove("Google-Apps", "1010020026", uClean); } catch(e){}
     try { AdminLicenseManager.LicenseAssignments.remove("Google-Apps", "1010020020", uClean); } catch(e){}
     
     AdminLicenseManager.LicenseAssignments.insert({ userId: uClean }, config.prod, config.sku);
     registrarEnHistorial("CAMBIO LICENCIA", "Usuario: " + uClean + " | Nueva: " + nuevaLicencia);
     
     var ss = SpreadsheetApp.getActiveSpreadsheet();
     var hoja = ss.getSheetByName("Hoja Maestra (657)");
     if (hoja) {
       var data = hoja.getDataRange().getValues();
       var headers = data[0].map(function(h) { return String(h).toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); });
       var idxEmail = headers.indexOf("email address"); if(idxEmail===-1) idxEmail = headers.indexOf("email");
       
       var idxLic = -1;
       for (var h = 0; h < headers.length; h++) {
         if (headers[h].includes("licencia")) { idxLic = h; break; }
       }
       
       if (idxEmail > -1 && idxLic > -1) {
         for (var r = 1; r < data.length; r++) {
           if (String(data[r][idxEmail]).toLowerCase().trim() === uClean) {
             hoja.getRange(r + 1, idxLic + 1).setValue(nuevaLicencia).setBackground("#e2f0d9");
             break;
           }
         }
       }
     }
     return "✅ Licencia " + nuevaLicencia + " asignada correctamente.";
  } catch(e) { 
    throw new Error("Google rechazó el cambio de licencia: " + e.message); 
  }
}

function webRefrescarSoloListas() {
  try {
    var listaUsuarios = AdminDirectory.Users.list({domain: "gigas.com", maxResults: 1});
    var customerIdReal = listaUsuarios.users[0].customerId;
    //motorSincronizarListasWorkspace(customerIdReal);
    var res = procesarEstructuraListas();
    return JSON.stringify({ 
      listasGigas: res.listasGigas, 
      listasOni: res.listasOni, 
      metricasGigas: res.metricasGigas,
      metricasOni: res.metricasOni,
      fechaSincro: res.fechaListas 
    });
  } catch(e) {
    throw new Error("Fallo al refrescar listas: " + e.message);
  }
}

function webComprobarEstadoTakeout(correo) { 
  return { estado: "ESPERA", texto: "🔍 Listo para consultar" }; 
}

function webEnviarInstruccionesTakeoutPREMIUM(u, m) {
  registrarEnHistorial("ENVÍO TAKEOUT", "Enviado a " + m + " para respaldar " + u);
  return "📩 Instrucciones de TakeOut enviadas a " + m;
}
// =========================================================
// ♻️ MÓDULO DE PAPELERA DE RECICLAJE (CUENTAS RECUPERABLES)
// =========================================================

// 🎯 PAPELERA EN TIEMPO REAL CON PAGINACIÓN COMPLETA
function webObtenerUsuariosEliminadosRecuperables() {
  verificarPermisoEjecucion();
  var listaEliminados = [];
  try {
    var listaUsuarios = AdminDirectory.Users.list({domain: "gigas.com", maxResults: 1});
    var customerIdReal = listaUsuarios.users[0].customerId;
    
    var pageToken;
    var hoy = new Date();
    
    // 🌟 Añadido bucle DO-WHILE para escanear todas las páginas de la papelera
    do {
      var response = AdminDirectory.Users.list({ 
        customer: customerIdReal, 
        showDeleted: true, 
        maxResults: 500, // Subimos el límite por página al máximo de Google
        pageToken: pageToken
      });
      
      if (response.users) {
        for (var i = 0; i < response.users.length; i++) {
          var u = response.users[i];
          var fBaja = new Date(u.deletionTime);
          var difMilisegundos = hoy.getTime() - fBaja.getTime();
          var diasTranscurridos = Math.floor(difMilisegundos / (1000 * 60 * 60 * 24));
          var diasRestantes = 20 - diasTranscurridos;
          
          listaEliminados.push({
            email: u.primaryEmail,
            fechaBaja: Utilities.formatDate(fBaja, "Europe/Madrid", "dd/MM/yyyy HH:mm"),
            diasRestantes: diasRestantes > 0 ? diasRestantes : 0
          });
        }
      }
      pageToken = response.nextPageToken;
    } while (pageToken);

    // Ordenamos para que las que están en mayor peligro de purga salgan arriba
    listaEliminados.sort(function(a, b) { return a.diasRestantes - b.diasRestantes; });
    return JSON.stringify(listaEliminados);
  } catch(e) {
    throw new Error("Error en Papelera Realtime: " + e.message);
  }
}
function actualizarCuentasMasivo(listaCuentas) {
  var resultados = { exitosos: 0, errores: 0, detalle: [] };
  
  for (var i = 0; i < listaCuentas.length; i++) {
    var item = listaCuentas[i];
    var email = item.email ? item.email.toString().trim().toLowerCase() : "";
    if (!email || !email.includes("@")) continue;
    
    var payload = {};
    
    // Puesto / Título y Departamento
    if (item.puesto !== "" || item.departamento !== "") {
      payload.organizations = [{
        primary: true,
        title: item.puesto || "",
        department: item.departamento || ""
      }];
    }
    
    // Teléfono de trabajo
    if (item.telefono !== "") {
      payload.phones = [{
        primary: true,
        type: "work",
        value: String(item.telefono)
      }];
    }
    
    // Direcciòn de trabajo
    if (item.direccion !== "") {
      payload.addresses = [{
        primary: true,
        type: "work",
        formatted: String(item.direccion)
      }];
    }
    
    // Si la fila solo tenía el correo sin datos nuevos, saltamos
    if (Object.keys(payload).length === 0) continue;
    
    try {
      AdminDirectory.Users.patch(payload, email);
      resultados.exitosos++;
    } catch(e) {
      resultados.errores++;
      resultados.detalle.push("Error en " + email + ": " + e.message);
    }
  }
  
  return resultados;
}
// 1. GENERA EL EXCEL/CSV CON LAS COLUMNAS OFICIALES Y TODOS LOS DATOS RELLENOS
function generarDatosRRHHCompleto() {
  var cabeceras = [
    "First Name [Required]", "Last Name [Required]", "Email Address [Required]", "Password [Required]",
    "Password Hash Function [UPLOAD ONLY]", "Org Unit Path [Required]", "New Primary Email [UPLOAD ONLY]",
    "Recovery Email", "Home Secondary Email", "Work Secondary Email", "Recovery Phone [import only]",
    "Work Phone", "Home Phone", "Mobile Phone", "Work Address", "Home Address", "Employee ID",
    "Employee Type", "Manager Email", "Department", "Cost Center", "Building ID", "Floor Name",
    "Floor Section", "Title", "Description", "Change Password at Next Sign-In", "New Status [UPLOAD ONLY]"
  ];
  
  var filas = [cabeceras];
  var paginaToken = null;
  
  do {
    var res = AdminDirectory.Users.list({
      customer: 'my_customer',
      maxResults: 500,
      pageToken: paginaToken,
      projection: 'full',
      orderBy: 'email'
    });
    
    var usuarios = res.users || [];
    for (var i = 0; i < usuarios.length; i++) {
      var u = usuarios[i];
      var nombre = u.name ? (u.name.givenName || "") : "";
      var apellidos = u.name ? (u.name.familyName || "") : "";
      var email = u.primaryEmail || "";
      var ou = u.orgUnitPath || "/";
      
      var workPhone = "", mobilePhone = "", homePhone = "";
      if (u.phones) {
        u.phones.forEach(function(p) {
          if (p.type === "work") workPhone = p.value;
          else if (p.type === "mobile" || p.type === "cell") mobilePhone = p.value;
          else if (p.type === "home") homePhone = p.value;
        });
      }
      
      var workAddress = "", homeAddress = "";
      if (u.addresses) {
        u.addresses.forEach(function(a) {
          if (a.type === "work") workAddress = a.formatted || a.streetAddress || "";
          else if (a.type === "home") homeAddress = a.formatted || a.streetAddress || "";
        });
      }
      
      var empId = "", empType = "", manager = "", dept = "", costCenter = "", title = "", desc = "", bId = "", floor = "", sec = "";
      if (u.organizations && u.organizations.length > 0) {
        var org = u.organizations[0];
        title = org.title || "";
        dept = org.department || "";
        costCenter = org.costCenter || "";
        desc = org.description || "";
      }
      if (u.externalIds) {
        u.externalIds.forEach(function(e) { if (e.type === "organization") empId = e.value; });
      }
      if (u.relations) {
        u.relations.forEach(function(r) { if (r.type === "manager") manager = r.value; });
      }
      if (u.locations && u.locations.length > 0) {
        bId = u.locations[0].buildingId || "";
        floor = u.locations[0].floorName || "";
        sec = u.locations[0].floorSection || "";
      }
      
      filas.push([
        nombre, apellidos, email, "****", "", ou, "",
        u.recoveryEmail || "", "", "", "", workPhone, homePhone, mobilePhone,
        workAddress, homeAddress, empId, empType, manager, dept, costCenter,
        bId, floor, sec, title, desc, "False", "Active"
      ]);
    }
    paginaToken = res.nextPageToken;
  } while (paginaToken);
  
  return filas;
}

// 2. RECIBE EL EXCEL RELLENADO POR RRHH Y ACTUALIZA GOOGLE WORKSPACE
function procesarSubidaRRHHBackend(filasCSV) {
  var resultados = { exitosos: 0, errores: 0, detalle: [] };
  
  // Empezamos desde la fila 1 para saltarnos los títulos de las cabeceras
  for (var i = 1; i < filasCSV.length; i++) {
    var c = filasCSV[i];
    var email = c[2] ? c[2].toString().trim().toLowerCase() : "";
    if (!email || !email.includes("@")) continue;
    
    var payload = {};
    
    // Organización (Título, Departamento, Centro de Coste, Descripción)
    if (c[24] !== "" || c[19] !== "" || c[20] !== "" || c[25] !== "") {
      payload.organizations = [{
        primary: true,
        title: c[24] ? c[24].toString().trim() : "",
        department: c[19] ? c[19].toString().trim() : "",
        costCenter: c[20] ? c[20].toString().trim() : "",
        description: c[25] ? c[25].toString().trim() : ""
      }];
    }
    
    // Teléfonos (Trabajo y Móvil)
    var telefonos = [];
    if (c[11] && c[11].toString().trim() !== "") telefonos.push({ type: "work", value: c[11].toString().trim() });
    if (c[13] && c[13].toString().trim() !== "") telefonos.push({ type: "mobile", value: c[13].toString().trim() });
    if (telefonos.length > 0) payload.phones = telefonos;
    
    // Direcciones (Trabajo)
    if (c[14] && c[14].toString().trim() !== "") {
      payload.addresses = [{ primary: true, type: "work", formatted: c[14].toString().trim() }];
    }
    
    // Manager
    if (c[18] && c[18].toString().trim() !== "") {
      payload.relations = [{ type: "manager", value: c[18].toString().trim() }];
    }
    
    // Employee ID
    if (c[16] && c[16].toString().trim() !== "") {
      payload.externalIds = [{ type: "organization", value: c[16].toString().trim() }];
    }
    
    if (Object.keys(payload).length === 0) continue;
    
    try {
      AdminDirectory.Users.patch(payload, email);
      resultados.exitosos++;
    } catch(e) {
      resultados.errores++;
      resultados.detalle.push("Fallo en " + email + ": " + e.message);
    }
  }
  
  return resultados;
}
// CREACIÓN DE CUENTA ENRIQUECIDA DESDE FORMULARIO DE ALTA (41 COLUMNAS + GRUPOS)
function crearCuentaGoogleEnriquecida(datos) {
  // 1. Validación de seguridad para los 4 campos estrictamente obligatorios
  if (!datos.nombre || !datos.apellidos || !datos.email || !datos.password) {
    throw new Error("Faltan campos críticos obligatorios: Nombre, Apellidos, Email o Contraseña.");
  }
  
  // 2. Construcción del paquete base del usuario
  var usuario = {
    name: {
      givenName: datos.nombre.trim(),
      familyName: datos.apellidos.trim()
    },
    primaryEmail: datos.email.trim().toLowerCase(),
    password: datos.password,
    changePasswordAtNextLogin: true,
    orgUnitPath: datos.ou ? datos.ou.trim() : "/"
  };
  
  // 3. Inyección dinámica de atributos opcionales (Solo si el formulario los envía)
  
  // Correo personal / Secundario de recuperación
  if (datos.emailPersonal && datos.emailPersonal.trim() !== "") {
    usuario.recoveryEmail = datos.emailPersonal.trim().toLowerCase();
  }
  
  // Organización: Puesto (Title), Departamento y Centro de Coste
  if (datos.puesto || datos.departamento || datos.centroCoste) {
    usuario.organizations = [{
      primary: true,
      title: datos.puesto ? datos.puesto.trim() : "",
      department: datos.departamento ? datos.departamento.trim() : "",
      costCenter: datos.centroCoste ? datos.centroCoste.trim() : ""
    }];
  }
  
  // Teléfono corporativo / móvil de trabajo
  if (datos.telefono && datos.telefono.trim() !== "") {
    usuario.phones = [{
      type: "work",
      value: datos.telefono.trim()
    }];
  }
  
  // Dirección de trabajo / País
  if (datos.pais && datos.pais.trim() !== "") {
    usuario.addresses = [{
      primary: true,
      type: "work",
      formatted: datos.pais.trim()
    }];
  }
  
  // Jefe Directo (Manager) para armar el organigrama oficial
  if (datos.managerEmail && datos.managerEmail.trim() !== "") {
    usuario.relations = [{
      type: "manager",
      value: datos.managerEmail.trim().toLowerCase()
    }];
  }
  
  try {
    // 4. Orden de creación oficial en Google Workspace
    var usuarioCreado = AdminDirectory.Users.insert(usuario);
    
    // 5. BONUS TRACK: Subscripción automática a la lista genérica de correo
    var mensajeGrupo = "";
    if (datos.listaCorreo && datos.listaCorreo.trim() !== "") {
      try {
        AdminDirectory.Members.insert({ email: usuario.primaryEmail }, datos.listaCorreo.trim().toLowerCase());
        mensajeGrupo = "\n👥 Añadido automáticamente a la lista: " + datos.listaCorreo.trim().toLowerCase();
      } catch (eGrupo) {
        mensajeGrupo = "\n⚠️ Cuenta creada, pero no se pudo añadir a la lista " + datos.listaCorreo + " (" + eGrupo.message + ")";
      }
    }
    
    return { 
      exito: true, 
      mensaje: "✅ Cuenta enriquecida creada con éxito para:\n📧 " + usuario.primaryEmail + mensajeGrupo 
    };
    
  } catch (e) {
    throw new Error("Error crítico de Google Workspace al crear la cuenta: " + e.message);
  }
}
// =========================================================================
// ⭐ ACTUALIZACIÓN ENRIQUECIDA + CORTADOR DE COMAS EN GRUPOS ⭐
// =========================================================================
function actualizarUsuarioGoogleEnriquecido(datos) {
  try {
    var correo = datos.email.toLowerCase().trim();
    
    // 1. ACTUALIZACIÓN DE PERFIL (Puesto, Dpto, Teléfono, Correo Personal, Nombre y Apellidos)
    var userPatch = {};
    if (datos.puesto || datos.departamento) {
      userPatch.organizations = [{
        title: datos.puesto || "",
        department: datos.departamento || "",
        primary: true
      }];
    }
    if (datos.telefono) {
      userPatch.phones = [{ value: datos.telefono, type: "work", primary: true }];
    }
    if (datos.emailPersonal) {
      userPatch.recoveryEmail = datos.emailPersonal;
    }
    if (datos.nombre || datos.apellidos) {
      userPatch.name = {
        givenName: datos.nombre || "",
        familyName: datos.apellidos || ""
      };
    }
    
    // Si hay datos de perfil que cambiar, se los enviamos a Google
    if (Object.keys(userPatch).length > 0) {
      AdminDirectory.Users.update(userPatch, correo);
    }
    
    // 2. GESTIÓN INTELIGENTE DE LISTAS/GRUPOS (¡Soporta múltiples listas por comas!)
    if (datos.grupo && datos.grupo.trim() !== "") {
      // Rompemos el texto por las comas para separar las listas
      var listas = datos.grupo.split(",");
      
      for (var i = 0; i < listas.length; i++) {
        var emailLista = listas[i].trim().toLowerCase();
        // Nos aseguramos de que sea un correo válido antes de disparar a Google
        if (emailLista.includes("@")) {
          try {
            AdminDirectory.Members.insert({
              email: correo,
              role: "MEMBER"
            }, emailLista);
          } catch(eGrupo) {
            // Si el usuario ya estaba dentro de esa lista, Google da error: lo ignoramos en silencio
          }
        }
      }
    }
    
    return { exito: true, mensaje: "Perfiles y listas sincronizados con éxito" };
  } catch(e) {
    return { exito: false, error: e.toString() };
  }
}
function traspasarPropiedadDrive(correoOrigen, correoDestino) {
  try {
    if (!correoOrigen || !correoDestino) {
      return { exito: false, error: "⚠️ Debes rellenar tanto el correo de origen como el de destino." };
    }
    
    // 1. TRADUCTOR: Convertimos los correos de texto a sus IDs numéricos internos (21 dígitos)
    let idOrigen = correoOrigen;
    let idDestino = correoDestino;
    
    try {
      if (typeof AdminDirectory !== 'undefined') {
        idOrigen = AdminDirectory.Users.get(correoOrigen).id;
        idDestino = AdminDirectory.Users.get(correoDestino).id;
      }
    } catch (errDir) {
      return { 
        exito: false, 
        error: "❌ No se ha encontrado a uno de los usuarios en tu directorio de Google. Asegúrate de que los correos están bien escritos, pertenecen a tu empresa y la cuenta de origen NO ha sido borrada todavía." 
      };
    }
    
    // 2. Obtenemos la llave de acceso de tu propia cuenta de Administrador
    const token = ScriptApp.getOAuthToken();
    
    // 3. Preparamos el paquete de traspaso usando los IDs numéricos oficiales
    const DRIVE_APP_ID = "435070579839"; 
    const payload = {
      oldOwnerUserId: idOrigen,
      newOwnerUserId: idDestino,
      applicationDataTransfers: [
        {
          applicationId: DRIVE_APP_ID,
          applicationTransferParams: [
            {
              key: "PRIVACY_LEVEL",
              value: "SHARED,PRIVATE"
            }
          ]
        }
      ]
    };
    
    // 4. Enviamos la orden directa al servidor oficial de traspasos de Google
    const url = "https://admin.googleapis.com/admin/datatransfer/v1/transfers";
    const opciones = {
      method: "post",
      contentType: "application/json",
      headers: {
        "Authorization": "Bearer " + token
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const respuesta = UrlFetchApp.fetch(url, opciones);
    const codigoEstado = respuesta.getResponseCode();
    const textoRespuesta = respuesta.getContentText();
    
    // 5. Verificamos el éxito de la orden (Códigos 200 o 201)
    if (codigoEstado === 200 || codigoEstado === 201) {
      return { 
        exito: true, 
        mensaje: "✅ ¡ORDEN RECIBIDA Y EN PROCESO!\n\nLos archivos de " + correoOrigen + " han empezado a transferirse en segundo plano hacia " + correoDestino + ". Google os enviará un correo electrónico automático al finalizar." 
      };
    } else {
      let errorLimpio = textoRespuesta;
      try {
        const jsonError = JSON.parse(textoRespuesta);
        if (jsonError.error && jsonError.error.message) {
          errorLimpio = jsonError.error.message;
        }
      } catch (ex) { }
      
      return { 
        exito: false, 
        error: "❌ El servidor de Google ha rechazado la orden (" + codigoEstado + "): " + errorLimpio 
      };
    }
    
  } catch (e) {
    return { exito: false, error: "Fallo de conexión: " + (e.message || e.toString()) };
  }
}