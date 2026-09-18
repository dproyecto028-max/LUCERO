const SPREADSHEET_ID = "1vjqjZYRdcdmKAM-vIIjb_Nq8jM9Rf9mjPKy4VZAdCgQ";
const DRIVE_FOLDER_ID = "1VtfPKlbgDTKqG8ke93MoFdmIufq5s7m3";
const DRIVE_COLABORADORES_FOLDER_ID = "1VtfPKlbgDTKqG8ke93MoFdmIufq5s7m3";
const EMAIL_ADMINISTRACION = "diegodantedanielolivarestello@gmail.com";
const CODIGOS_COLABORADORES = ["LUCERO2026", "ESTANTERIAS", "SERGIO2026", "ADMIN2026", "COLAB2026", "EQUIPO2026", "SUPERVISOR2026", "MENDOZA2026", "SOPORTE2026", "LOGISTICA2026"];

// -------------------------AL REGISTRAR ME DICE CODIGO
// UTILIDADES
// -------------------------

function obtenerHoja(ss, nombreHoja) {
  let sheet = ss.getSheetByName(nombreHoja);
  if (!sheet) {
    sheet = ss.insertSheet(nombreHoja);
  }
  return sheet;
}

function obtenerCarpetaColaboradores() {
  return DriveApp.getFolderById(DRIVE_COLABORADORES_FOLDER_ID);
}

function guardarArchivoEnDrive(base64Data, nombreArchivo, carpetaDestino) {
  if (!base64Data || !nombreArchivo) return "";
  try {
    const partes = base64Data.split(",");
    const mimeType = partes[0].split(";")[0].replace("data:", "");
    const datosDecodificados = Utilities.base64Decode(partes[1]);
    const blob = Utilities.newBlob(datosDecodificados, mimeType, nombreArchivo);

    const folder = carpetaDestino || DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return "Error al subir archivo: " + e.toString();
  }
}

function generarTarjetaEmail(titulo, nombre, contenidoHtml, colorHeader = "#0284c7") {
  return `
    <div style="background-color: #f0f9ff; padding: 20px; font-family: 'Segoe UI', Helvetica, Arial, sans-serif;">
      <div style="max-width: 520px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 12px rgba(2, 132, 199, 0.12); border: 1px solid #bae6fd;">
        <div style="background-color: ${colorHeader}; padding: 24px; text-align: center; color: #ffffff;">
          <h1 style="margin: 0; font-size: 20px; font-weight: 600; letter-spacing: 0.5px;">TOTAL RACK</h1>
          <p style="margin: 5px 0 0 0; font-size: 13px; opacity: 0.95;">${titulo}</p>
        </div>
        <div style="padding: 28px; color: #334155; font-size: 14px; line-height: 1.6;">
          <p style="font-size: 16px; font-weight: 600; margin-top: 0; color: #0f172a;">👋 ¡Hola, ${nombre}!</p>
          ${contenidoHtml}
        </div>
        <div style="background-color: #f8fafc; padding: 16px 28px; text-align: center; font-size: 11px; color: #64748b; border-top: 1px solid #e2e8f0;">
          <p style="margin: 0;">Este es un correo automático de <b>TOTAL RACK</b>.</p>
          <p style="margin: 4px 0 0 0;">© ${new Date().getFullYear()} TOTAL RACK. Todos los derechos reservados.</p>
        </div>
      </div>
    </div>
  `;
}

function enviarCorreoSeguro(destino, asunto, mensajeHtml, mensajePlano) {
  if (!destino) return;
  try {
    GmailApp.sendEmail(destino, asunto, mensajePlano, { htmlBody: mensajeHtml });
  } catch (err) {
    try {
      MailApp.sendEmail({
        to: destino,
        subject: asunto,
        body: mensajePlano,
        htmlBody: mensajeHtml
      });
    } catch (e) {
      Logger.log("Error al enviar correo: " + e.toString());
    }
  }
}

function validarCamposObligatorios(data, campos) {
  for (let i = 0; i < campos.length; i++) {
    const campo = campos[i];
    if (!data[campo] || String(data[campo]).trim() === "") {
      return `El campo '${campo}' es obligatorio.`;
    }
  }
  return null;
}

function existeDNI(ss, dniBuscado) {
  const dniLimpio = String(dniBuscado || "").trim();
  const hojas = ["REGISTRO CLIENTES", "REGISTRO COLABORADORES"];

  for (let h = 0; h < hojas.length; h++) {
    const sheet = ss.getSheetByName(hojas[h]);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][2] || "").trim() === dniLimpio) {
        return true;
      }
    }
  }
  return false;
}

function buscarColaboradorPorEmail(ss, email) {
  const sheet = ss.getSheetByName("REGISTRO COLABORADORES");
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][11] || "").trim() === String(email || "").trim()) {
      return {
        fila: i + 1,
        nombre: data[i][1],
        dni: data[i][2],
        email: data[i][11]
      };
    }
  }
  return null;
}

function normalizarEncabezado(valor) {
  return String(valor || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function buscarUsuarioEnHoja(sheet, dni, password) {
  if (!sheet) return null;

  const filas = sheet.getDataRange().getValues();
  if (!filas.length) return null;

  for (let i = 0; i < filas.length; i++) {
    const fila = filas[i];
    const dniFila = String(fila[2] || "").trim();
    const passwordFila = String(fila[3] || "").trim();

    if (dniFila.toLowerCase() === "dni" || dniFila.toLowerCase() === "documento") {
      continue;
    }

    if (
      dniFila === String(dni || "").trim() &&
      passwordFila === String(password || "").trim()
    ) {
      return {
        name: fila[1] || "",
        email: fila[11] || ""
      };
    }
  }

  return null;
}

function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalizarTipoJornada(tipo) {
  const valor = String(tipo || "").trim().toLowerCase();
  if (["inicio", "iniciar", "start", "entrada"].includes(valor)) return "inicio";
  if (["fin", "finalizar", "finalizar turno", "finish", "salida"].includes(valor)) return "finalizar";
  return valor;
}

// -------------------------
// DO POST
// -------------------------

function doPost(e) {
  try {
    let data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const action = String(data.action || "").trim();

    // 1) REGISTRO CLIENTE
    if (action === "registro_cliente" || action === "registroCliente") {
      const requeridosCliente = ["nombre", "dni", "password", "provincia", "localidad", "calle", "altura", "cp", "celular", "email"];
      const errorValidacion = validarCamposObligatorios(data, requeridosCliente);
      if (errorValidacion) return responseJSON({ status: "error", message: errorValidacion });

      if (existeDNI(ss, data.dni)) {
        return responseJSON({ status: "error", message: "El DNI ingresado ya se encuentra registrado." });
      }

      const sheet = obtenerHoja(ss, "REGISTRO CLIENTES");
      const lastRow = sheet.getLastRow();
      const newId = lastRow <= 1 ? 1 : (Number(sheet.getRange(lastRow, 1).getValue()) || 0) + 1;

      sheet.appendRow([
        newId,
        data.nombre,
        data.dni,
        data.password,
        data.dni,
        data.provincia,
        data.localidad,
        data.calle,
        data.altura,
        data.cp,
        data.celular,
        data.email
      ]);

      const contenido = `
        <p>Nos alegra darle la bienvenida a nuestra plataforma. Su cuenta de <b>Cliente</b> ha sido creada con éxito. 🎉</p>
        <div style="background-color: #f0f9ff; border-left: 4px solid #0284c7; padding: 12px 16px; margin: 20px 0; border-radius: 0 8px 8px 0;">
          <p style="margin: 0; font-weight: 600; color: #0369a1;">📌 Resumen de tu cuenta:</p>
          <p style="margin: 6px 0 0 0;"><b>Usuario (DNI):</b> ${data.dni}</p>
          <p style="margin: 4px 0 0 0;"><b>Estado:</b> Activa ✅</p>
        </div>
        <p>Ya puede acceder a nuestros servicios y gestionar sus solicitudes desde nuestro portal web.</p>
      `;

      const htmlCorreo = generarTarjetaEmail("Confirmación de Registro", data.nombre, contenido, "#0284c7");
      const planoCorreo = `¡Hola ${data.nombre}! Su registro como Cliente en TOTAL RACK ha sido exitoso. Usuario (DNI): ${data.dni}`;
      enviarCorreoSeguro(data.email, "Bienvenido/a a TOTAL RACK", htmlCorreo, planoCorreo);

      return responseJSON({ status: "success", message: "Cliente registrado con éxito", id: newId });
    }

    // 2) REGISTRO COLABORADOR
    if (action === "registro_colaborador" || action === "registroColaborador") {
      const requeridosColab = ["nombre", "dni", "password", "provincia", "localidad", "calle", "altura", "cp", "celular", "email", "cvBase64", "fotoBase64"];
      const errorValidacion = validarCamposObligatorios(data, requeridosColab);
      if (errorValidacion) return responseJSON({ status: "error", message: errorValidacion });

      const codigoIngresado = String(data.codigoAdmin || "").trim().toUpperCase();
      if (!CODIGOS_COLABORADORES.includes(codigoIngresado)) {
        return responseJSON({ status: "error", message: "Código de acceso inválido. Solo puede registrarse el personal autorizado por administración." });
      }

      if (existeDNI(ss, data.dni)) {
        return responseJSON({ status: "error", message: "El DNI ingresado ya se encuentra registrado." });
      }

      const sheet = obtenerHoja(ss, "REGISTRO COLABORADORES");
      const lastRow = sheet.getLastRow();
      const newId = lastRow <= 1 ? 1 : (Number(sheet.getRange(lastRow, 1).getValue()) || 0) + 1;

      const carpetaColaboradores = obtenerCarpetaColaboradores();
      const urlCv = guardarArchivoEnDrive(
        data.cvBase64,
        `CV_${data.dni}_${data.cvNombre || "documento.pdf"}`,
        carpetaColaboradores
      );
      const urlFoto = guardarArchivoEnDrive(
        data.fotoBase64,
        `FOTO_${data.dni}_${data.fotoNombre || "foto.jpg"}`,
        carpetaColaboradores
      );

      sheet.appendRow([
        newId,
        data.nombre,
        data.dni,
        data.password,
        data.dni,
        data.provincia,
        data.localidad,
        data.calle,
        data.altura,
        data.cp,
        data.celular,
        data.email,
        urlCv,
        urlFoto,
        codigoIngresado
      ]);

      const contenido = `
        <p>¡Es un placer tenerte en nuestro equipo! Tu perfil de <b>Colaborador</b> ha sido registrado correctamente. 🚀</p>
        <div style="background-color: #f0f9ff; border-left: 4px solid #0891b2; padding: 12px 16px; margin: 20px 0; border-radius: 0 8px 8px 0;">
          <p style="margin: 0; font-weight: 600; color: #0e7490;">📑 Estado del Registro:</p>
          <p style="margin: 6px 0 0 0;"><b>Documentación:</b> Recibida y almacenada 📁</p>
          <p style="margin: 4px 0 0 0;"><b>Acceso:</b> Habilitado para fichajes ⏱️</p>
        </div>
        <p>Recuerda marcar tus ingresos y salidas a través de la sección de marcación de jornada.</p>
      `;

      const htmlCorreo = generarTarjetaEmail("Registro de Colaborador", data.nombre, contenido, "#0891b2");
      const planoCorreo = `¡Hola ${data.nombre}! Tu registro como Colaborador en TOTAL RACK ha sido exitoso.`;
      enviarCorreoSeguro(data.email, "Bienvenido/a al equipo de TOTAL RACK", htmlCorreo, planoCorreo);

      return responseJSON({ status: "success", message: "Colaborador registrado con éxito", id: newId });
    }

    // 3) TRABAJA CON NOSOTROS / POSTULACIONES
    if (action === "trabaja_con_nosotros" || action === "trabajaConNosotros" || action === "postular_trabajo" || action === "postularTrabajo") {
      const requeridos = ["nombre", "email", "telefono", "puesto", "mensaje"];
      const errorValidacion = validarCamposObligatorios(data, requeridos);
      if (errorValidacion) return responseJSON({ status: "error", message: errorValidacion });

      const cvUrl = guardarArchivoEnDrive(data.cvBase64, `POSTULACION_${String(data.email || "usuario").replace(/\s+/g, "_")}_${data.cvNombre || "curriculum.pdf"}`);
      const sheet = obtenerHoja(ss, "POSTULACIONES");
      const fecha = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");
      sheet.appendRow([
        fecha,
        data.nombre,
        data.email,
        data.telefono,
        data.puesto,
        data.mensaje,
        cvUrl || "No cargado"
      ]);

      const contenido = `
        <p>Se recibió una nueva postulación desde la sección <b>Trabaja con nosotros</b>.</p>
        <div style="background-color: #f0f9ff; border-left: 4px solid #0284c7; padding: 12px 16px; margin: 20px 0; border-radius: 0 8px 8px 0;">
          <p style="margin: 0; font-weight: 600; color: #0369a1;">📄 Datos del postulante:</p>
          <p style="margin: 6px 0 0 0;"><b>Nombre:</b> ${data.nombre}</p>
          <p style="margin: 4px 0 0 0;"><b>Email:</b> ${data.email}</p>
          <p style="margin: 4px 0 0 0;"><b>Teléfono:</b> ${data.telefono}</p>
          <p style="margin: 4px 0 0 0;"><b>Puesto:</b> ${data.puesto}</p>
          <p style="margin: 4px 0 0 0;"><b>CV:</b> ${cvUrl || "No adjuntado"}</p>
        </div>
        <p><b>Mensaje:</b> ${data.mensaje}</p>
      `;

      const htmlCorreo = generarTarjetaEmail("Nueva postulación recibida", data.nombre, contenido, "#0284c7");
      enviarCorreoSeguro(EMAIL_ADMINISTRACION, "Nueva postulación en TOTAL RACK", htmlCorreo, `Nueva postulación de ${data.nombre} para ${data.puesto}.`);

      return responseJSON({ status: "success", message: "Postulación enviada correctamente. Nos comunicaremos pronto." });
    }

    // 4) LOGIN
    if (action === "login") {
      const errorValidacion = validarCamposObligatorios(data, ["dni", "password"]);
      if (errorValidacion) return responseJSON({ status: "error", message: errorValidacion });

      const sheetClienteObj = ss.getSheetByName("REGISTRO CLIENTES");
      const sheetColabObj = ss.getSheetByName("REGISTRO COLABORADORES");

      const sheetCliente = sheetClienteObj ? sheetClienteObj.getDataRange().getValues() : [];
      const sheetColab = sheetColabObj ? sheetColabObj.getDataRange().getValues() : [];

      const cliente = buscarUsuarioEnHoja(sheetClienteObj, data.dni, data.password);
      if (cliente) {
        return responseJSON({
          status: "success",
          type: "cliente",
          name: cliente.name,
          email: cliente.email
        });
      }

      const colaborador = buscarUsuarioEnHoja(sheetColabObj, data.dni, data.password);
      if (colaborador) {
        return responseJSON({
          status: "success",
          type: "colaborador",
          name: colaborador.name,
          email: colaborador.email
        });
      }

      return responseJSON({ status: "error", message: "DNI o contraseña incorrectos" });
    }

    // 4) SOLICITUD DE PROYECTO
    if (action === "contrato_servicio" || action === "contratoServicio") {
      const requeridos = ["nombre", "email", "horas", "tipoTrabajo", "descripcion", "planoBase64"];
      const errorValidacion = validarCamposObligatorios(data, requeridos);
      if (errorValidacion) return responseJSON({ status: "error", message: errorValidacion });

      const planoUrl = guardarArchivoEnDrive(data.planoBase64, `PLANO_${String(data.email || "usuario").replace(/\s+/g, "_")}_${data.planoNombre || "proyecto.pdf"}`);
      const sheet = obtenerHoja(ss, "CONTRATOS SERVICIO");
      const fecha = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");
      sheet.appendRow([
        fecha,
        data.nombre,
        data.email,
        data.horas,
        data.tipoTrabajo,
        data.descripcion,
        planoUrl || "No cargado"
      ]);

      const contenido = `
        <p>Se recibió una nueva solicitud de proyecto de un cliente interesado en una solución Total Rack.</p>
        <div style="background-color: #f0f9ff; border-left: 4px solid #0284c7; padding: 12px 16px; margin: 20px 0; border-radius: 0 8px 8px 0;">
          <p style="margin: 0; font-weight: 600; color: #0369a1;">📋 Detalle del contrato:</p>
          <p style="margin: 6px 0 0 0;"><b>Cliente:</b> ${data.nombre}</p>
          <p style="margin: 4px 0 0 0;"><b>Email:</b> ${data.email}</p>
          <p style="margin: 4px 0 0 0;"><b>Cantidad de módulos:</b> ${data.horas}</p>
          <p style="margin: 4px 0 0 0;"><b>Tipo de proyecto:</b> ${data.tipoTrabajo}</p>
          <p style="margin: 4px 0 0 0;"><b>Plano:</b> ${planoUrl || "No cargado"}</p>
        </div>
        <p><b>Descripción:</b> ${data.descripcion}</p>
      `;

      const htmlCorreo = generarTarjetaEmail("Nueva solicitud de proyecto", data.nombre, contenido, "#f59e0b");
      const planoCorreo = `Se recibió tu solicitud de proyecto, ${data.nombre}. Un asesor se pondrá en contacto con usted para una atención personalizada.`;
      const planoCorreoAdmin = `Nuevo proyecto solicitado por ${data.nombre}. Interés: ${data.tipoTrabajo}. Módulos: ${data.horas}. Plano: ${planoUrl || "No cargado"}. Descripción: ${data.descripcion}`;
      enviarCorreoSeguro(data.email, "Solicitud recibida - TOTAL RACK", htmlCorreo, planoCorreo);
      enviarCorreoSeguro(EMAIL_ADMINISTRACION, `Nuevo proyecto solicitado - ${data.tipoTrabajo}`, htmlCorreo, planoCorreoAdmin);

      return responseJSON({ status: "success", message: "Solicitud recibida. Un asesor se pondrá en contacto con usted para una atención personalizada." });
    }

    // 5) CONFIRMAR PAGO
    if (action === "confirmar_pago" || action === "confirmarPago") {
      const errorValidacion = validarCamposObligatorios(data, ["email", "nombre", "plan"]);
      if (errorValidacion) return responseJSON({ status: "error", message: errorValidacion });

      const contenido = `
        <p>Hemos procesado tu comprobante y confirmamos que tu pago ha sido acreditado exitosamente. 💳</p>
        <div style="background-color: #f0f9ff; border-left: 4px solid #0284c7; padding: 12px 16px; margin: 20px 0; border-radius: 0 8px 8px 0;">
          <p style="margin: 0; font-weight: 600; color: #0369a1;">🧾 Detalle de la Transacción:</p>
          <p style="margin: 6px 0 0 0;"><b>Plan/Servicio:</b> ${data.plan}</p>
          <p style="margin: 4px 0 0 0;"><b>Estado del Pago:</b> Acreditado 🟢</p>
        </div>
        <p>¡Gracias por confiar en nosotros!</p>
      `;

      const htmlCorreo = generarTarjetaEmail("Comprobante de Pago Acreditado", data.nombre, contenido, "#0284c7");
      const planoCorreo = `¡Pago Acreditado! Estimado/a ${data.nombre}, confirmamos que se ha acreditado correctamente su pago por el plan ${data.plan}.`;

      enviarCorreoSeguro(data.email, "Comprobante recibido - TOTAL RACK", htmlCorreo, planoCorreo);

      return responseJSON({ status: "success", message: "Notificación de pago enviada por correo" });
    }

    // 6) MARCACIÓN DE JORNADA
    if (action === "marcar_jornada" || action === "marcarJornada") {
      const tipoRaw = normalizarTipoJornada(data.tipo);

      if (!data.email && !data.nombre) {
        return responseJSON({ status: "error", message: "Se requiere nombre o email del colaborador para registrar la jornada." });
      }

      let nombreColaborador = data.nombre || "";
      if (!nombreColaborador && data.email) {
        const colaborador = buscarColaboradorPorEmail(ss, data.email);
        if (!colaborador) {
          return responseJSON({ status: "error", message: "No se encontró un colaborador asociado a ese email." });
        }
        nombreColaborador = colaborador.nombre;
      }

      const sheetJornadas = obtenerHoja(ss, "JORNADAS LABORALES");
      const ahora = new Date();
      const fechaHoraActual = Utilities.formatDate(ahora, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");

      if (tipoRaw === "inicio") {
        sheetJornadas.appendRow([fechaHoraActual, nombreColaborador, "", "", "", ""]);
        return responseJSON({ status: "success", message: "Inicio de jornada registrado en Sheets." });
      }

      if (tipoRaw === "finalizar") {
        const errorValidacion = validarCamposObligatorios(data, ["email"]);
        if (errorValidacion) return responseJSON({ status: "error", message: errorValidacion });

        const lugarTrabajo = data.lugarTrabajo || "No especificado";
        const novedades = data.novedades && data.novedades.trim() !== "" ? data.novedades : "Sin novedades registradas";

        const dataJornadas = sheetJornadas.getDataRange().getValues();
        let filaEncontrada = -1;

        for (let i = dataJornadas.length - 1; i >= 1; i--) {
          if (String(dataJornadas[i][1] || "").trim() === nombreColaborador && String(dataJornadas[i][4] || "").trim() === "") {
            filaEncontrada = i + 1;
            break;
          }
        }

        if (filaEncontrada === -1) {
          return responseJSON({ status: "error", message: "No se encontró un inicio de jornada activo para este colaborador." });
        }

        const inicioVal = sheetJornadas.getRange(filaEncontrada, 1).getValue();
        const fechaInicio = (inicioVal instanceof Date) ? inicioVal : new Date(inicioVal);

        const totalSegundos = Math.floor((ahora - fechaInicio) / 1000);
        const horas = Math.floor(totalSegundos / 3600);
        const minutos = Math.floor((totalSegundos % 3600) / 60);
        const segundos = totalSegundos % 60;

        const totalHorasTexto =
          (horas < 10 ? "0" + horas : horas) + ":" +
          (minutos < 10 ? "0" + minutos : minutos) + ":" +
          (segundos < 10 ? "0" + segundos : segundos);

        sheetJornadas.getRange(filaEncontrada, 3).setValue(lugarTrabajo);
        sheetJornadas.getRange(filaEncontrada, 4).setValue(novedades);
        sheetJornadas.getRange(filaEncontrada, 5).setValue(fechaHoraActual);
        sheetJornadas.getRange(filaEncontrada, 6).setValue(totalHorasTexto);

        const contenido = `
          <p>Has finalizado tu jornada laboral exitosamente. ⏱️</p>
          <div style="background-color: #f0f9ff; border-left: 4px solid #0284c7; padding: 12px 16px; margin: 20px 0; border-radius: 0 8px 8px 0;">
            <p style="margin: 0; font-weight: 600; color: #0369a1;">📋 Detalle del Turno:</p>
            <p style="margin: 6px 0 0 0;"><b>Colaborador:</b> ${nombreColaborador}</p>
            <p style="margin: 4px 0 0 0;"><b>Lugar de Trabajo:</b> ${lugarTrabajo}</p>
            <p style="margin: 4px 0 0 0;"><b>Tiempo Trabajado:</b> ${totalHorasTexto}</p>
            <p style="margin: 4px 0 0 0;"><b>Hora de Cierre:</b> ${fechaHoraActual}</p>
            <p style="margin: 4px 0 0 0;"><b>Novedades / Reportes:</b> ${novedades}</p>
          </div>
        `;

        const htmlCorreo = generarTarjetaEmail("Resumen de Finalización de Jornada", nombreColaborador, contenido, "#0284c7");
        const planoCorreo = `Resumen de Jornada: Colaborador: ${nombreColaborador} | Lugar: ${lugarTrabajo} | Tiempo: ${totalHorasTexto} | Novedades: ${novedades}`;
        enviarCorreoSeguro(data.email, "Resumen de jornada - TOTAL RACK", htmlCorreo, planoCorreo);

        return responseJSON({ status: "success", message: "Jornada finalizada, calculada y registrada exitosamente." });
      }
    }

    // 7) REPORTES / INCIDENTES
    if (action === "reportar_incidente" || action === "reportarIncidente" || action === "guardar_reporte" || action === "guardarReporte") {
      const errorValidacion = validarCamposObligatorios(data, ["nombreColaborador", "motivo", "descripcion"]);
      if (errorValidacion) return responseJSON({ status: "error", message: errorValidacion });

      const colaborador = data.nombreColaborador;
      const motivo = data.motivo;
      const descripcion = data.descripcion;
      const lugar = data.lugarTrabajo || "No asignado";

      const contenidoAdmin = `
        <p>Se ha recibido un nuevo reporte de incidentes desde el portal del colaborador.</p>
        <div style="background-color: #f0f9ff; border-left: 4px solid #0284c7; padding: 12px 16px; margin: 20px 0; border-radius: 0 8px 8px 0;">
          <p style="margin: 0; font-weight: 600; color: #0369a1;">🚨 Detalle del Reporte:</p>
          <p style="margin: 6px 0 0 0;"><b>Colaborador:</b> ${colaborador}</p>
          <p style="margin: 4px 0 0 0;"><b>Lugar de Trabajo:</b> ${lugar}</p>
          <p style="margin: 4px 0 0 0;"><b>Motivo:</b> ${motivo}</p>
          <p style="margin: 4px 0 0 0;"><b>Descripción:</b> ${descripcion}</p>
        </div>
      `;

      const htmlCorreo = generarTarjetaEmail("Reporte de Incidente / Novedad", "Administrador", contenidoAdmin, "#0284c7");
      const planoCorreo = `Reporte de Incidente: Colaborador: ${colaborador} | Motivo: ${motivo} | Descripcion: ${descripcion}`;
      enviarCorreoSeguro(EMAIL_ADMINISTRACION, `🚨 Nuevo Reporte de Incidente - ${motivo}`, htmlCorreo, planoCorreo);

      return responseJSON({ status: "success", message: "Reporte enviado a Administración exitosamente" });
    }

    return responseJSON({ status: "error", message: "Acción no válida o no especificada" });
  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  }
}
