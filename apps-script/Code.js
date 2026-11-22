const GOOGLE_CLIENT_ID = "252920345538-gq6c1c99ntsk88d9sj19jvjkre2mt1nd.apps.googleusercontent.com";

function doGet(e) {
  const token = (e && e.parameter && e.parameter.token) ? String(e.parameter.token).trim() : '';

  // Si abren directo el link del script sin token, redirigir limpiamente al login de GitHub Pages
  if (!token) {
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html lang="es">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <meta http-equiv="refresh" content="0; url=https://fdvg99.github.io/gestion-hogar/">
        <title>Redirigiendo a Gestión Hogar...</title>
      </head>
      <body style="margin:0;background:#0f172a;color:#f8fafc;font-family:sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;text-align:center;padding:20px;">
        <div>
          <p style="font-size:14px;color:#94a3b8;margin-bottom:16px;">Redirigiendo al inicio de sesión...</p>
          <a href="https://fdvg99.github.io/gestion-hogar/" target="_top" style="display:inline-block;background:#059669;color:#fff;text-decoration:none;font-weight:bold;font-size:13px;padding:12px 20px;border-radius:12px;">
            Ir a Finanzas del Hogar
          </a>
        </div>
        <script>
          try { window.top.location.href = "https://fdvg99.github.io/gestion-hogar/"; } catch (err) {}
        </script>
      </body>
      </html>
    `)
    .setTitle('Gestión Hogar')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.initialToken = token;

  return template.evaluate()
    .setTitle('Gestión Hogar - Finanzas Familiares')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/**
 * Valida el token JWT enviado desde el cliente y devuelve los datos de la app
 */
function autenticarYCargarApp(idToken) {
  const userEmail = verifyGoogleIdToken_(idToken);
  if (!userEmail) {
    throw new Error('TOKEN_INVALIDO: La sesión de Google expiró o no es válida.');
  }
  if (!checkUserAuthorization_(userEmail)) {
    throw new Error('NO_AUTORIZADO: El correo ' + userEmail + ' no está activo en Config_Familia.');
  }
  return getAppData(userEmail);
}

function verifyGoogleIdToken_(idToken) {
  if (!idToken) {
    throw new Error('TOKEN_VACIO: No llegó ningún token en la URL.');
  }
  try {
    // 1. Intento oficial contra endpoint de Google
    const url = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken);
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      const payload = JSON.parse(response.getContentText());
      if (payload.aud === GOOGLE_CLIENT_ID && String(payload.email_verified) === 'true') {
        return String(payload.email || '').toLowerCase().trim();
      }
    }

    // 2. Respaldo decodificando el JWT con relleno Base64 correcto
    const parts = idToken.split('.');
    if (parts.length === 3) {
      let base64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
      while (base64.length % 4 !== 0) {
        base64 += '=';
      }
      const decoded = Utilities.newBlob(Utilities.base64Decode(base64)).getDataAsString();
      const payload = JSON.parse(decoded);
      const now = Math.floor(Date.now() / 1000);

      if (payload.aud !== GOOGLE_CLIENT_ID) {
        throw new Error('AUD_MISMATCH: Client ID distinto.');
      }
      if (payload.exp <= now) {
        throw new Error('TOKEN_EXPIRADO: El token ya venció, inicia sesión de nuevo.');
      }
      if (payload.email) {
        return String(payload.email).toLowerCase().trim();
      }
    }
    throw new Error('FORMATO_INVALIDO: El token recibido tiene ' + idToken.length + ' caracteres.');
  } catch (err) {
    throw new Error(err.message);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function checkUserAuthorization_(email) {
  if (!email) return false;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config_Familia');
  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowEmail = String(data[i][3] || '').toLowerCase().trim();
    const isActive = data[i][5] === true || String(data[i][5]).toUpperCase() === 'TRUE';
    if (rowEmail === email && isActive) {
      return true;
    }
  }
  return false;
}