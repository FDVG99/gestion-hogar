/**
 * Inicializa la estructura de hojas, encabezados y datos base del sistema.
 * Ejecutar una sola vez desde el editor de Apps Script.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    // 1. Definición de esquema de pestañas y encabezados
    const schema = {
      Config_Familia: ['ID_Integrante', 'Nombre', 'Rol', 'Email', 'Cuota_Base', 'Activo'],
      Periodos: [
        'ID_Periodo',
        'Mes_Anio',
        'Responsable_ID',
        'Saldo_Anterior',
        'Total_Ingresos',
        'Total_Gastos',
        'Saldo_Cierre',
        'Estado',
        'Fecha_Cierre',
      ],
      Cuotas: [
        'ID_Cuota',
        'ID_Periodo',
        'ID_Integrante',
        'Nombre_Integrante',
        'Monto_Asignado',
        'Pagado',
        'Fecha_Pago',
      ],
      Gastos: [
        'ID_Gasto',
        'ID_Periodo',
        'Categoria',
        'Subcategoria',
        'Descripcion',
        'Monto',
        'Pagado',
        'Fecha_Pago',
      ],
    };

    // Crear o limpiar pestañas según el esquema
    Object.keys(schema).forEach((sheetName) => {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      } else {
        sheet.clear();
      }

      const headers = schema[sheetName];
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange
        .setValues([headers])
        .setFontWeight('bold')
        .setBackground('#1e293b')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    });

    // Eliminar la hoja por defecto "Hoja 1" o "Sheet1" si existe y hay más hojas
    const defaultSheet = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
    if (defaultSheet && ss.getSheets().length > 1) {
      ss.deleteSheet(defaultSheet);
    }

    // 2. Poblar Config_Familia (Reemplaza los correos placeholder por los 5 Gmail reales)
    const familiaData = [
      ['INT-01', 'José', 'Papá', 'jose.familia@gmail.com', 160, true],
      ['INT-02', 'Judith', 'Mamá', 'judith.familia@gmail.com', 65, true],
      ['INT-03', 'Everin', 'Hermana', 'everin.familia@gmail.com', 115, true],
      ['INT-04', 'Natalie', 'Hermana', 'natalie.familia@gmail.com', 115, true],
      ['INT-05', 'Fermat', 'Hijo', 'fermat.familia@gmail.com', 115, true],
    ];
    ss.getSheetByName('Config_Familia')
      .getRange(2, 1, familiaData.length, familiaData[0].length)
      .setValues(familiaData);

    // 3. Inicializar Periodo Actual (Noviembre 2026)
    const periodoId = 'PER-2026-11';
    const periodoData = [
      [
        periodoId,
        'Noviembre 2026',
        'INT-05', // Fermat como responsable inicial
        0, // Saldo_Anterior (Pozo heredado)
        0, // Total_Ingresos (se calcula dinámicamente)
        0, // Total_Gastos (se calcula dinámicamente)
        0, // Saldo_Cierre
        'ABIERTO',
        '',
      ],
    ];
    ss.getSheetByName('Periodos').getRange(2, 1, 1, periodoData[0].length).setValues(periodoData);

    // 4. Generar Cuotas Base para el periodo inicial
    const cuotasData = familiaData.map((member, idx) => [
      `CUO-202611-0${idx + 1}`,
      periodoId,
      member[0], // ID_Integrante
      member[1], // Nombre
      member[4], // Cuota_Base
      false, // Pagado
      '', // Fecha_Pago
    ]);
    ss.getSheetByName('Cuotas')
      .getRange(2, 1, cuotasData.length, cuotasData[0].length)
      .setValues(cuotasData);

    // 5. Precargar plantilla de servicios fijos del mes (incluyendo los 4 pisos de Cálidda)
    const gastosFijosData = [
      ['GAS-202611-01', periodoId, 'Agua', 'Sedapal', 'Recibo mensual Sedapal', 0, false, ''],
      ['GAS-202611-02', periodoId, 'Luz', 'Enel', 'Recibo mensual Enel', 0, false, ''],
      ['GAS-202611-03', periodoId, 'Internet', 'Win', 'Fibra óptica Win', 0, false, ''],
      ['GAS-202611-04', periodoId, 'Gas', 'Cálidda', 'Cálidda - Piso 1', 0, false, ''],
      ['GAS-202611-05', periodoId, 'Gas', 'Cálidda', 'Cálidda - Piso 2', 0, false, ''],
      ['GAS-202611-06', periodoId, 'Gas', 'Cálidda', 'Cálidda - Piso 3', 0, false, ''],
      ['GAS-202611-07', periodoId, 'Gas', 'Cálidda', 'Cálidda - Piso 4', 0, false, ''],
    ];
    ss.getSheetByName('Gastos')
      .getRange(2, 1, gastosFijosData.length, gastosFijosData[0].length)
      .setValues(gastosFijosData);

    // 6. Formato numérico de moneda (Soles) y autoajuste de columnas
    ss.getSheetByName('Config_Familia').getRange('E2:E6').setNumberFormat('"S/ "#,##0.00');
    ss.getSheetByName('Periodos').getRange('D2:G2').setNumberFormat('"S/ "#,##0.00');
    ss.getSheetByName('Cuotas').getRange('E2:E6').setNumberFormat('"S/ "#,##0.00');
    ss.getSheetByName('Gastos').getRange('F2:F8').setNumberFormat('"S/ "#,##0.00');

    Object.keys(schema).forEach((sheetName) => {
      const sheet = ss.getSheetByName(sheetName);
      sheet.autoResizeColumns(1, schema[sheetName].length);
    });

    Logger.log('Base de datos inicializada correctamente.');
    return { success: true, message: 'Base de datos creada e inicializada correctamente.' };
  } catch (error) {
    Logger.log('Error en setupDatabase: ' + error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Obtiene todo el estado inicial de la app en una sola lectura en bloque.
 */
/**
 * Obtiene todo el estado inicial de la app en una sola lectura en bloque.
 */
function getAppData(loggedEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userEmail = String(loggedEmail || Session.getActiveUser().getEmail() || '')
    .toLowerCase()
    .trim();

  const familiaRows = ss.getSheetByName('Config_Familia').getDataRange().getValues();
  const periodosRows = ss.getSheetByName('Periodos').getDataRange().getValues();
  const cuotasRows = ss.getSheetByName('Cuotas').getDataRange().getValues();
  const gastosRows = ss.getSheetByName('Gastos').getDataRange().getValues();

  // 1. Integrantes activos
  const familia = [];
  let currentUser = null;
  for (let i = 1; i < familiaRows.length; i++) {
    const r = familiaRows[i];
    if (r[5] === true || String(r[5]).toUpperCase() === 'TRUE') {
      const member = {
        id: String(r[0]),
        nombre: String(r[1]),
        rol: String(r[2]),
        email: String(r[3]).toLowerCase().trim(),
        cuotaBase: Number(r[4]) || 0,
      };
      familia.push(member);
      if (member.email === userEmail) currentUser = member;
    }
  }

  // 2. Periodos (historial y periodo abierto actual)
  const periodos = [];
  let activePeriodo = null;
  for (let i = 1; i < periodosRows.length; i++) {
    const r = periodosRows[i];
    const p = {
      rowIndex: i + 1,
      id: String(r[0]),
      mesAnio: formatMesAnio_(r[1]),
      responsableId: String(r[2]),
      saldoAnterior: Number(r[3]) || 0,
      totalIngresos: Number(r[4]) || 0,
      totalGastos: Number(r[5]) || 0,
      saldoCierre: Number(r[6]) || 0,
      estado: String(r[7]),
      fechaCierre: r[8]
        ? Utilities.formatDate(new Date(r[8]), Session.getScriptTimeZone(), 'dd/MM/yyyy')
        : '',
    };
    periodos.push(p);
    if (p.estado === 'ABIERTO') activePeriodo = p;
  }

  // Fallback si no hay abierto
  if (!activePeriodo && periodos.length > 0) {
    activePeriodo = periodos[periodos.length - 1];
  }

  // 3. Cuotas del periodo activo
  const cuotas = [];
  for (let i = 1; i < cuotasRows.length; i++) {
    const r = cuotasRows[i];
    if (String(r[1]) === activePeriodo.id) {
      cuotas.push({
        rowIndex: i + 1,
        id: String(r[0]),
        periodoId: String(r[1]),
        integranteId: String(r[2]),
        nombre: String(r[3]),
        monto: Number(r[4]) || 0,
        pagado: r[5] === true || String(r[5]).toUpperCase() === 'TRUE',
        fechaPago: formatFechaCorta_(r[6]),
      });
    }
  }

  // 4. Gastos del periodo activo
  const gastos = [];
  for (let i = 1; i < gastosRows.length; i++) {
    const r = gastosRows[i];
    if (String(r[1]) === activePeriodo.id) {
      gastos.push({
        rowIndex: i + 1,
        id: String(r[0]),
        periodoId: String(r[1]),
        categoria: String(r[2]),
        subcategoria: String(r[3]),
        descripcion: String(r[4]),
        monto: Number(r[5]) || 0,
        pagado: r[6] === true || String(r[6]).toUpperCase() === 'TRUE',
        fechaPago: formatFechaCorta_(r[7]),
      });
    }
  }

  // Serialización segura sin objetos Date
  return JSON.parse(
    JSON.stringify({
      currentUser,
      familia,
      activePeriodo,
      historialPeriodos: periodos.filter((p) => p.estado === 'CERRADO').reverse(),
      cuotas,
      gastos,
    })
  );
}

/**
 * Actualiza cuota en una sola escritura en bloque (sin recargar las 4 hojas).
 */
function updateCuota(cuotaId, nuevoMonto, pagado) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Cuotas');
    const data = sheet.getDataRange().getValues();
    const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(cuotaId)) {
        const fechaLimpia = pagado ? (data[i][6] ? formatFechaCorta_(data[i][6]) : hoy) : '';
        sheet
          .getRange(i + 1, 5, 1, 3)
          .setValues([[Number(nuevoMonto), Boolean(pagado), fechaLimpia ? "'" + fechaLimpia : '']]);
        return {
          success: true,
          id: cuotaId,
          monto: Number(nuevoMonto),
          pagado: Boolean(pagado),
          fechaPago: fechaLimpia,
        };
      }
    }
    throw new Error('Cuota no encontrada');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Actualiza gasto en una sola escritura en bloque (sin recargar las 4 hojas).
 */
function updateGasto(gastoId, nuevoMonto, pagado, nuevaDescripcion) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Gastos');
    const data = sheet.getDataRange().getValues();
    const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(gastoId)) {
        const desc =
          nuevaDescripcion !== undefined && nuevaDescripcion !== null
            ? String(nuevaDescripcion)
            : String(data[i][4]);
        const fechaLimpia = pagado ? (data[i][7] ? formatFechaCorta_(data[i][7]) : hoy) : '';
        sheet
          .getRange(i + 1, 5, 1, 4)
          .setValues([
            [desc, Number(nuevoMonto), Boolean(pagado), fechaLimpia ? "'" + fechaLimpia : ''],
          ]);
        return {
          success: true,
          id: gastoId,
          descripcion: desc,
          monto: Number(nuevoMonto),
          pagado: Boolean(pagado),
          fechaPago: fechaLimpia,
        };
      }
    }
    throw new Error('Gasto no encontrado');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Agrega un gasto extraordinario y devuelve solo el registro creado.
 */
function addGastoExtra(periodoId, descripcion, monto, pagado) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Gastos');
    const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const newId = 'GAS-' + new Date().getTime();
    const fechaLimpia = pagado ? hoy : '';

    const nextRow = sheet.getLastRow() + 1;
    const range = sheet.getRange(nextRow, 1, 1, 8);
    range.setValues([
      [
        newId,
        periodoId,
        'Extra',
        'Eventual',
        String(descripcion),
        Number(monto),
        Boolean(pagado),
        fechaLimpia ? "'" + fechaLimpia : '',
      ],
    ]);
    sheet.getRange(nextRow, 6).setNumberFormat('"S/ "#,##0.00');

    return {
      success: true,
      gasto: {
        id: newId,
        periodoId: String(periodoId),
        categoria: 'Extra',
        subcategoria: 'Eventual',
        descripcion: String(descripcion),
        monto: Number(monto),
        pagado: Boolean(pagado),
        fechaPago: fechaLimpia,
      },
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Elimina un gasto extraordinario sin recargar todas las hojas.
 */
function deleteGastoExtra(gastoId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Gastos');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(gastoId) && String(data[i][2]) === 'Extra') {
        sheet.deleteRow(i + 1);
        return { success: true, id: gastoId };
      }
    }
    throw new Error('Gasto extra no encontrado');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cambia el responsable del periodo activo sin cerrarlo.
 */
function cambiarResponsableActivo(periodoId, nuevoResponsableId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Periodos');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(periodoId)) {
        sheet.getRange(i + 1, 3).setValue(String(nuevoResponsableId));
        break;
      }
    }
    return getAppData();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cierra el periodo actual, congela el saldo de cierre y abre el siguiente mes
 * transfiriendo el saldo acumulado al nuevo responsable.
 */
function cerrarMesYTransferir(periodoActualId, siguienteMesAnio, nuevoResponsableId, loggedEmail) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const periodosSheet = ss.getSheetByName('Periodos');
    const familiaSheet = ss.getSheetByName('Config_Familia');
    const cuotasSheet = ss.getSheetByName('Cuotas');
    const gastosSheet = ss.getSheetByName('Gastos');

    recalcularTotalesPeriodo_(periodoActualId);

    const pData = periodosSheet.getDataRange().getValues();
    let saldoArrastre = 0;

    for (let i = 1; i < pData.length; i++) {
      if (String(pData[i][0]) === String(periodoActualId)) {
        saldoArrastre = Number(pData[i][6]) || 0;
        periodosSheet.getRange(i + 1, 8).setValue('CERRADO');
        periodosSheet.getRange(i + 1, 9).setValue(new Date());
        break;
      }
    }

    // Crear nuevo periodo
    const nuevoPeriodoId = 'PER-' + new Date().getTime();
    periodosSheet.appendRow([
      nuevoPeriodoId,
      "'" + String(siguienteMesAnio).trim(),
      String(nuevoResponsableId),
      saldoArrastre,
      0,
      0,
      saldoArrastre,
      'ABIERTO',
      '',
    ]);
    const newPRow = periodosSheet.getLastRow();
    periodosSheet.getRange(newPRow, 4, 1, 4).setNumberFormat('"S/ "#,##0.00');

    // Generar cuotas base para el nuevo periodo
    const fData = familiaSheet.getDataRange().getValues();
    const nuevasCuotas = [];
    for (let i = 1; i < fData.length; i++) {
      if (fData[i][5] === true || String(fData[i][5]).toUpperCase() === 'TRUE') {
        nuevasCuotas.push([
          'CUO-' + new Date().getTime() + '-' + i,
          nuevoPeriodoId,
          String(fData[i][0]),
          String(fData[i][1]),
          Number(fData[i][4]) || 0,
          false,
          '',
        ]);
      }
    }
    if (nuevasCuotas.length > 0) {
      const startCRow = cuotasSheet.getLastRow() + 1;
      cuotasSheet.getRange(startCRow, 1, nuevasCuotas.length, 7).setValues(nuevasCuotas);
      cuotasSheet.getRange(startCRow, 5, nuevasCuotas.length, 1).setNumberFormat('"S/ "#,##0.00');
    }

    // Precargar servicios fijos en S/ 0.00
    const ts = new Date().getTime();
    const nuevosGastos = [
      [
        'GAS-' + ts + '-1',
        nuevoPeriodoId,
        'Agua',
        'Sedapal',
        'Recibo mensual Sedapal',
        0,
        false,
        '',
      ],
      ['GAS-' + ts + '-2', nuevoPeriodoId, 'Luz', 'Enel', 'Recibo mensual Enel', 0, false, ''],
      ['GAS-' + ts + '-3', nuevoPeriodoId, 'Internet', 'Win', 'Fibra óptica Win', 0, false, ''],
      ['GAS-' + ts + '-4', nuevoPeriodoId, 'Gas', 'Cálidda', 'Cálidda - Piso 1', 0, false, ''],
      ['GAS-' + ts + '-5', nuevoPeriodoId, 'Gas', 'Cálidda', 'Cálidda - Piso 2', 0, false, ''],
      ['GAS-' + ts + '-6', nuevoPeriodoId, 'Gas', 'Cálidda', 'Cálidda - Piso 3', 0, false, ''],
      ['GAS-' + ts + '-7', nuevoPeriodoId, 'Gas', 'Cálidda', 'Cálidda - Piso 4', 0, false, ''],
    ];
    const startGRow = gastosSheet.getLastRow() + 1;
    gastosSheet.getRange(startGRow, 1, nuevosGastos.length, 8).setValues(nuevosGastos);
    gastosSheet.getRange(startGRow, 6, nuevosGastos.length, 1).setNumberFormat('"S/ "#,##0.00');

    return getAppData(loggedEmail);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Función privada que recalcula Total_Ingresos, Total_Gastos y Saldo_Cierre en la hoja Periodos.
 */
function recalcularTotalesPeriodo_(periodoId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cuotasData = ss.getSheetByName('Cuotas').getDataRange().getValues();
  const gastosData = ss.getSheetByName('Gastos').getDataRange().getValues();
  const periodosSheet = ss.getSheetByName('Periodos');
  const periodosData = periodosSheet.getDataRange().getValues();

  let totalIngresosRecaudados = 0;
  for (let i = 1; i < cuotasData.length; i++) {
    if (
      String(cuotasData[i][1]) === periodoId &&
      (cuotasData[i][5] === true || String(cuotasData[i][5]).toUpperCase() === 'TRUE')
    ) {
      totalIngresosRecaudados += Number(cuotasData[i][4]) || 0;
    }
  }

  let totalGastosPagados = 0;
  for (let i = 1; i < gastosData.length; i++) {
    if (
      String(gastosData[i][1]) === periodoId &&
      (gastosData[i][6] === true || String(gastosData[i][6]).toUpperCase() === 'TRUE')
    ) {
      totalGastosPagados += Number(gastosData[i][5]) || 0;
    }
  }

  for (let i = 1; i < periodosData.length; i++) {
    if (String(periodosData[i][0]) === periodoId) {
      const saldoAnterior = Number(periodosData[i][3]) || 0;
      const saldoCierre = saldoAnterior + totalIngresosRecaudados - totalGastosPagados;
      periodosSheet
        .getRange(i + 1, 5, 1, 3)
        .setValues([[totalIngresosRecaudados, totalGastosPagados, saldoCierre]]);
      break;
    }
  }
}

/**
 * Convierte fechas auto-formateadas por Google Sheets ("Tue Dec 01 2026...")
 * a texto limpio en español ("Diciembre 2026").
 */
function formatMesAnio_(val) {
  if (val instanceof Date && !isNaN(val)) {
    const meses = [
      'Enero',
      'Febrero',
      'Marzo',
      'Abril',
      'Mayo',
      'Junio',
      'Julio',
      'Agosto',
      'Septiembre',
      'Octubre',
      'Noviembre',
      'Diciembre',
    ];
    return `${meses[val.getMonth()]} ${val.getFullYear()}`;
  }
  return String(val || '').trim();
}

/**
 * Formatea celdas de fecha de pago (dd/MM/yyyy) si Google Sheets las convirtió en Date.
 */
function formatFechaCorta_(val) {
  if (!val) return '';
  if (val instanceof Date && !isNaN(val)) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return String(val).trim();
}

/**
 * Obtiene el detalle completo (cuotas y gastos) de un periodo cerrado específico.
 */
function getDetallePeriodo(periodoId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cuotasRows = ss.getSheetByName('Cuotas').getDataRange().getValues();
  const gastosRows = ss.getSheetByName('Gastos').getDataRange().getValues();

  const cuotas = [];
  for (let i = 1; i < cuotasRows.length; i++) {
    const r = cuotasRows[i];
    if (String(r[1]) === String(periodoId)) {
      cuotas.push({
        id: String(r[0]),
        nombre: String(r[3]),
        monto: Number(r[4]) || 0,
        pagado: r[5] === true || String(r[5]).toUpperCase() === 'TRUE',
        fechaPago: formatFechaCorta_(r[6]),
      });
    }
  }

  const gastos = [];
  for (let i = 1; i < gastosRows.length; i++) {
    const r = gastosRows[i];
    if (String(r[1]) === String(periodoId)) {
      gastos.push({
        id: String(r[0]),
        categoria: String(r[2]),
        subcategoria: String(r[3]),
        descripcion: String(r[4]),
        monto: Number(r[5]) || 0,
        pagado: r[6] === true || String(r[6]).toUpperCase() === 'TRUE',
        fechaPago: formatFechaCorta_(r[7]),
      });
    }
  }

  return JSON.parse(JSON.stringify({ cuotas, gastos }));
}

/**
 * ============================================================================
 * IMPORTADOR REUTILIZABLE DE HISTÓRICO (2026 y años anteriores)
 * ============================================================================
 */
function cargarHistorico2026() {
  const historico2026 = [
    {
      id: 'PER-2026-01',
      mesAnio: 'Enero 2026',
      responsableId: 'INT-05',
      saldoAnterior: 411.6,
      totalIngresos: 475.0,
      totalGastos: 461.75,
      saldoCierre: 424.85,
      fechaCierre: '01/02/2026',
      cuotas: [
        { intId: 'INT-03', nombre: 'Evelin', monto: 120.0, pagado: true, fechaPago: '31/01/2026' },
        { intId: 'INT-04', nombre: 'Nathaly', monto: 0.0, pagado: true, fechaPago: '31/01/2026' },
        { intId: 'INT-05', nombre: 'Fermat', monto: 120.0, pagado: true, fechaPago: '31/01/2026' },
        { intId: 'INT-02', nombre: 'Judith', monto: 65.0, pagado: true, fechaPago: '31/01/2026' },
        { intId: 'INT-01', nombre: 'Jose', monto: 170.0, pagado: true, fechaPago: '31/01/2026' },
      ],
      gastos: [
        {
          cat: 'Agua',
          sub: 'Sedapal',
          desc: 'Recibo Sedapal (5021421)',
          monto: 65.35,
          pagado: true,
          fechaPago: '01/02/2026',
        },
        {
          cat: 'Luz',
          sub: 'Enel',
          desc: 'Pluz Energía (1069216)',
          monto: 214.0,
          pagado: true,
          fechaPago: '01/02/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 1 (268988)',
          monto: 55.2,
          pagado: true,
          fechaPago: '01/02/2026',
        },
        {
          cat: 'Internet',
          sub: 'Win',
          desc: 'Fibra Óptica Win (328111)',
          monto: 118.9,
          pagado: true,
          fechaPago: '01/02/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 4 (270026)',
          monto: 2.5,
          pagado: true,
          fechaPago: '01/02/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 3 (270116)',
          monto: 5.8,
          pagado: true,
          fechaPago: '01/02/2026',
        },
      ],
    },
    {
      id: 'PER-2026-02',
      mesAnio: 'Febrero 2026',
      responsableId: 'INT-05',
      saldoAnterior: 424.85,
      totalIngresos: 455.0,
      totalGastos: 432.3,
      saldoCierre: 447.55,
      fechaCierre: '13/03/2026',
      cuotas: [
        { intId: 'INT-03', nombre: 'Evelin', monto: 110.0, pagado: true, fechaPago: '28/02/2026' },
        { intId: 'INT-04', nombre: 'Nathaly', monto: 0.0, pagado: true, fechaPago: '28/02/2026' },
        { intId: 'INT-05', nombre: 'Fermat', monto: 110.0, pagado: true, fechaPago: '28/02/2026' },
        { intId: 'INT-02', nombre: 'Judith', monto: 65.0, pagado: true, fechaPago: '28/02/2026' },
        { intId: 'INT-01', nombre: 'Jose', monto: 170.0, pagado: true, fechaPago: '28/02/2026' },
      ],
      gastos: [
        {
          cat: 'Agua',
          sub: 'Sedapal',
          desc: 'Recibo Sedapal (5021421)',
          monto: 58.4,
          pagado: true,
          fechaPago: '13/03/2026',
        },
        {
          cat: 'Luz',
          sub: 'Enel',
          desc: 'Pluz Energía (1069216)',
          monto: 178.0,
          pagado: true,
          fechaPago: '13/03/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 1 (268988)',
          monto: 67.0,
          pagado: true,
          fechaPago: '13/03/2026',
        },
        {
          cat: 'Internet',
          sub: 'Win',
          desc: 'Fibra Óptica Win (328111)',
          monto: 118.9,
          pagado: true,
          fechaPago: '11/03/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 4 (270026)',
          monto: 2.5,
          pagado: true,
          fechaPago: '13/03/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 3 (270116)',
          monto: 7.5,
          pagado: true,
          fechaPago: '13/03/2026',
        },
      ],
    },
    {
      id: 'PER-2026-03',
      mesAnio: 'Marzo 2026',
      responsableId: 'INT-05',
      saldoAnterior: 447.55,
      totalIngresos: 450.0,
      totalGastos: 373.78,
      saldoCierre: 523.77,
      fechaCierre: '31/03/2026',
      cuotas: [
        { intId: 'INT-03', nombre: 'Evelin', monto: 110.0, pagado: true, fechaPago: '31/03/2026' },
        { intId: 'INT-04', nombre: 'Nathaly', monto: 0.0, pagado: true, fechaPago: '31/03/2026' },
        { intId: 'INT-05', nombre: 'Fermat', monto: 110.0, pagado: true, fechaPago: '31/03/2026' },
        { intId: 'INT-02', nombre: 'Judith', monto: 60.0, pagado: true, fechaPago: '31/03/2026' },
        { intId: 'INT-01', nombre: 'Jose', monto: 170.0, pagado: true, fechaPago: '31/03/2026' },
      ],
      gastos: [
        {
          cat: 'Agua',
          sub: 'Sedapal',
          desc: 'Recibo Sedapal (5021421)',
          monto: 53.9,
          pagado: true,
          fechaPago: '31/03/2026',
        },
        {
          cat: 'Luz',
          sub: 'Enel',
          desc: 'Pluz Energía (1069216)',
          monto: 162.5,
          pagado: true,
          fechaPago: '31/03/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 1 (268988)',
          monto: 23.9,
          pagado: true,
          fechaPago: '31/03/2026',
        },
        {
          cat: 'Internet',
          sub: 'Win',
          desc: 'Fibra Óptica Win (328111)',
          monto: 123.48,
          pagado: true,
          fechaPago: '31/03/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 4 (270026)',
          monto: 2.5,
          pagado: true,
          fechaPago: '31/03/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 3 (270116)',
          monto: 7.5,
          pagado: true,
          fechaPago: '31/03/2026',
        },
      ],
    },
    {
      id: 'PER-2026-04',
      mesAnio: 'Abril 2026',
      responsableId: 'INT-05',
      saldoAnterior: 523.77,
      totalIngresos: 450.0,
      totalGastos: 398.2,
      saldoCierre: 580.57,
      fechaCierre: '30/04/2026',
      cuotas: [
        { intId: 'INT-03', nombre: 'Evelin', monto: 110.0, pagado: true, fechaPago: '30/04/2026' },
        { intId: 'INT-04', nombre: 'Nathaly', monto: 0.0, pagado: true, fechaPago: '30/04/2026' },
        { intId: 'INT-05', nombre: 'Fermat', monto: 110.0, pagado: true, fechaPago: '30/04/2026' },
        { intId: 'INT-02', nombre: 'Judith', monto: 60.0, pagado: true, fechaPago: '30/04/2026' },
        { intId: 'INT-01', nombre: 'Jose', monto: 170.0, pagado: true, fechaPago: '30/04/2026' },
      ],
      gastos: [
        {
          cat: 'Agua',
          sub: 'Sedapal',
          desc: 'Recibo Sedapal (5021421)',
          monto: 50.1,
          pagado: true,
          fechaPago: '30/04/2026',
        },
        {
          cat: 'Luz',
          sub: 'Enel',
          desc: 'Pluz Energía (1069216)',
          monto: 170.5,
          pagado: true,
          fechaPago: '30/04/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 1 (268988)',
          monto: 49.8,
          pagado: true,
          fechaPago: '30/04/2026',
        },
        {
          cat: 'Internet',
          sub: 'Win',
          desc: 'Fibra Óptica Win (328111)',
          monto: 118.9,
          pagado: true,
          fechaPago: '30/04/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 4 (270026)',
          monto: 2.5,
          pagado: true,
          fechaPago: '30/04/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 3 (270116)',
          monto: 6.4,
          pagado: true,
          fechaPago: '30/04/2026',
        },
      ],
    },
    {
      id: 'PER-2026-05',
      mesAnio: 'Mayo 2026',
      responsableId: 'INT-05',
      saldoAnterior: 580.57,
      totalIngresos: 450.0,
      totalGastos: 419.4,
      saldoCierre: 611.17,
      fechaCierre: '31/05/2026',
      cuotas: [
        { intId: 'INT-03', nombre: 'Evelin', monto: 110.0, pagado: true, fechaPago: '31/05/2026' },
        { intId: 'INT-04', nombre: 'Nathaly', monto: 0.0, pagado: true, fechaPago: '31/05/2026' },
        { intId: 'INT-05', nombre: 'Fermat', monto: 110.0, pagado: true, fechaPago: '31/05/2026' },
        { intId: 'INT-02', nombre: 'Judith', monto: 60.0, pagado: true, fechaPago: '31/05/2026' },
        { intId: 'INT-01', nombre: 'Jose', monto: 170.0, pagado: true, fechaPago: '31/05/2026' },
      ],
      gastos: [
        {
          cat: 'Agua',
          sub: 'Sedapal',
          desc: 'Recibo Sedapal (5021421)',
          monto: 66.8,
          pagado: true,
          fechaPago: '31/05/2026',
        },
        {
          cat: 'Luz',
          sub: 'Enel',
          desc: 'Pluz Energía (1069216)',
          monto: 171.0,
          pagado: true,
          fechaPago: '31/05/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 1 (268988)',
          monto: 50.7,
          pagado: true,
          fechaPago: '31/05/2026',
        },
        {
          cat: 'Internet',
          sub: 'Win',
          desc: 'Fibra Óptica Win (328111)',
          monto: 118.9,
          pagado: true,
          fechaPago: '31/05/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 4 (270026)',
          monto: 4.5,
          pagado: true,
          fechaPago: '31/05/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 3 (270116)',
          monto: 7.5,
          pagado: true,
          fechaPago: '31/05/2026',
        },
      ],
    },
    {
      id: 'PER-2026-06',
      mesAnio: 'Junio 2026',
      responsableId: 'INT-05',
      saldoAnterior: 611.17,
      totalIngresos: 450.0,
      totalGastos: 458.2,
      saldoCierre: 602.97,
      fechaCierre: '30/06/2026',
      cuotas: [
        { intId: 'INT-03', nombre: 'Evelin', monto: 110.0, pagado: true, fechaPago: '30/06/2026' },
        { intId: 'INT-04', nombre: 'Nathaly', monto: 0.0, pagado: true, fechaPago: '30/06/2026' },
        { intId: 'INT-05', nombre: 'Fermat', monto: 110.0, pagado: true, fechaPago: '30/06/2026' },
        { intId: 'INT-02', nombre: 'Judith', monto: 60.0, pagado: true, fechaPago: '30/06/2026' },
        { intId: 'INT-01', nombre: 'Jose', monto: 170.0, pagado: true, fechaPago: '30/06/2026' },
      ],
      gastos: [
        {
          cat: 'Agua',
          sub: 'Sedapal',
          desc: 'Recibo Sedapal (5021421)',
          monto: 121.6,
          pagado: true,
          fechaPago: '30/06/2026',
        },
        {
          cat: 'Luz',
          sub: 'Enel',
          desc: 'Pluz Energía (1069216)',
          monto: 158.5,
          pagado: true,
          fechaPago: '30/06/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 1 (268988)',
          monto: 48.8,
          pagado: true,
          fechaPago: '30/06/2026',
        },
        {
          cat: 'Internet',
          sub: 'Win',
          desc: 'Fibra Óptica Win (328111)',
          monto: 118.9,
          pagado: true,
          fechaPago: '30/06/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 4 (270026)',
          monto: 3.2,
          pagado: true,
          fechaPago: '30/06/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 3 (270116)',
          monto: 7.2,
          pagado: true,
          fechaPago: '30/06/2026',
        },
      ],
    },
    {
      id: 'PER-2026-07',
      mesAnio: 'Julio 2026',
      responsableId: 'INT-05',
      saldoAnterior: 602.97,
      totalIngresos: 455.0,
      totalGastos: 445.1,
      saldoCierre: 612.87,
      fechaCierre: '31/07/2026',
      cuotas: [
        { intId: 'INT-03', nombre: 'Evelin', monto: 110.0, pagado: true, fechaPago: '31/07/2026' },
        { intId: 'INT-04', nombre: 'Nathaly', monto: 0.0, pagado: true, fechaPago: '31/07/2026' },
        { intId: 'INT-05', nombre: 'Fermat', monto: 110.0, pagado: true, fechaPago: '31/07/2026' },
        { intId: 'INT-02', nombre: 'Judith', monto: 65.0, pagado: true, fechaPago: '31/07/2026' },
        { intId: 'INT-01', nombre: 'Jose', monto: 170.0, pagado: true, fechaPago: '31/07/2026' },
      ],
      gastos: [
        {
          cat: 'Agua',
          sub: 'Sedapal',
          desc: 'Recibo Sedapal (5021421)',
          monto: 71.2,
          pagado: true,
          fechaPago: '31/07/2026',
        },
        {
          cat: 'Luz',
          sub: 'Enel',
          desc: 'Pluz Energía (1069216)',
          monto: 180.0,
          pagado: true,
          fechaPago: '31/07/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 1 (268988)',
          monto: 67.9,
          pagado: true,
          fechaPago: '31/07/2026',
        },
        {
          cat: 'Internet',
          sub: 'Win',
          desc: 'Fibra Óptica Win (328111)',
          monto: 118.9,
          pagado: true,
          fechaPago: '31/07/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 4 (270026)',
          monto: 3.3,
          pagado: true,
          fechaPago: '31/07/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 3 (270116)',
          monto: 3.8,
          pagado: true,
          fechaPago: '31/07/2026',
        },
      ],
    },
    {
      id: 'PER-2026-08',
      mesAnio: 'Agosto 2026',
      responsableId: 'INT-05',
      saldoAnterior: 612.87,
      totalIngresos: 455.0,
      totalGastos: 502.1,
      saldoCierre: 565.77,
      fechaCierre: '31/08/2026',
      cuotas: [
        { intId: 'INT-03', nombre: 'Evelin', monto: 110.0, pagado: true, fechaPago: '31/08/2026' },
        { intId: 'INT-04', nombre: 'Nathaly', monto: 0.0, pagado: true, fechaPago: '31/08/2026' },
        { intId: 'INT-05', nombre: 'Fermat', monto: 110.0, pagado: true, fechaPago: '31/08/2026' },
        { intId: 'INT-02', nombre: 'Judith', monto: 65.0, pagado: true, fechaPago: '31/08/2026' },
        { intId: 'INT-01', nombre: 'Jose', monto: 170.0, pagado: true, fechaPago: '31/08/2026' },
      ],
      gastos: [
        {
          cat: 'Agua',
          sub: 'Sedapal',
          desc: 'Recibo Sedapal (5021421)',
          monto: 71.0,
          pagado: true,
          fechaPago: '31/08/2026',
        },
        {
          cat: 'Luz',
          sub: 'Enel',
          desc: 'Pluz Energía (1069216)',
          monto: 242.5,
          pagado: true,
          fechaPago: '31/08/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 1 (268988)',
          monto: 62.7,
          pagado: true,
          fechaPago: '31/08/2026',
        },
        {
          cat: 'Internet',
          sub: 'Win',
          desc: 'Fibra Óptica Win (328111)',
          monto: 118.9,
          pagado: true,
          fechaPago: '31/08/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 4 (270026)',
          monto: 3.2,
          pagado: true,
          fechaPago: '31/08/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 3 (270116)',
          monto: 3.8,
          pagado: true,
          fechaPago: '31/08/2026',
        },
      ],
    },
    {
      id: 'PER-2026-09',
      mesAnio: 'Septiembre 2026',
      responsableId: 'INT-05',
      saldoAnterior: 565.77,
      totalIngresos: 450.0,
      totalGastos: 447.9,
      saldoCierre: 567.87,
      fechaCierre: '30/09/2026',
      cuotas: [
        { intId: 'INT-03', nombre: 'Evelin', monto: 110.0, pagado: true, fechaPago: '30/09/2026' },
        { intId: 'INT-04', nombre: 'Nathaly', monto: 0.0, pagado: true, fechaPago: '30/09/2026' },
        { intId: 'INT-05', nombre: 'Fermat', monto: 110.0, pagado: true, fechaPago: '30/09/2026' },
        { intId: 'INT-02', nombre: 'Judith', monto: 60.0, pagado: true, fechaPago: '30/09/2026' },
        { intId: 'INT-01', nombre: 'Jose', monto: 170.0, pagado: true, fechaPago: '30/09/2026' },
      ],
      gastos: [
        {
          cat: 'Agua',
          sub: 'Sedapal',
          desc: 'Recibo Sedapal (5021421)',
          monto: 58.5,
          pagado: true,
          fechaPago: '30/09/2026',
        },
        {
          cat: 'Luz',
          sub: 'Enel',
          desc: 'Pluz Energía (1069216)',
          monto: 195.5,
          pagado: true,
          fechaPago: '30/09/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 1 (268988)',
          monto: 64.6,
          pagado: true,
          fechaPago: '30/09/2026',
        },
        {
          cat: 'Internet',
          sub: 'Win',
          desc: 'Fibra Óptica Win (328111)',
          monto: 118.9,
          pagado: true,
          fechaPago: '30/09/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 4 (270026)',
          monto: 3.2,
          pagado: true,
          fechaPago: '30/09/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 3 (270116)',
          monto: 7.2,
          pagado: true,
          fechaPago: '30/09/2026',
        },
      ],
    },
    {
      id: 'PER-2026-10',
      mesAnio: 'Octubre 2026',
      responsableId: 'INT-05',
      saldoAnterior: 567.87,
      totalIngresos: 455.0,
      totalGastos: 413.6,
      saldoCierre: 609.27,
      fechaCierre: '31/10/2026',
      cuotas: [
        { intId: 'INT-03', nombre: 'Evelin', monto: 110.0, pagado: true, fechaPago: '31/10/2026' },
        { intId: 'INT-04', nombre: 'Nathaly', monto: 0.0, pagado: true, fechaPago: '31/10/2026' },
        { intId: 'INT-05', nombre: 'Fermat', monto: 110.0, pagado: true, fechaPago: '31/10/2026' },
        { intId: 'INT-02', nombre: 'Judith', monto: 65.0, pagado: true, fechaPago: '31/10/2026' },
        { intId: 'INT-01', nombre: 'Jose', monto: 170.0, pagado: true, fechaPago: '31/10/2026' },
      ],
      gastos: [
        {
          cat: 'Agua',
          sub: 'Sedapal',
          desc: 'Recibo Sedapal (5021421)',
          monto: 62.6,
          pagado: true,
          fechaPago: '31/10/2026',
        },
        {
          cat: 'Luz',
          sub: 'Enel',
          desc: 'Pluz Energía (1069216)',
          monto: 174.0,
          pagado: true,
          fechaPago: '31/10/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 1 (268988)',
          monto: 47.7,
          pagado: true,
          fechaPago: '31/10/2026',
        },
        {
          cat: 'Internet',
          sub: 'Win',
          desc: 'Fibra Óptica Win (328111)',
          monto: 118.9,
          pagado: true,
          fechaPago: '31/10/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 4 (270026)',
          monto: 3.2,
          pagado: true,
          fechaPago: '31/10/2026',
        },
        {
          cat: 'Gas',
          sub: 'Cálidda',
          desc: 'Cálidda - Piso 3 (270116)',
          monto: 7.2,
          pagado: true,
          fechaPago: '31/10/2026',
        },
      ],
    },
  ];

  importarPeriodosHistoricos_(historico2026);
}

/**
 * Motor reutilizable: inserta o reemplaza periodos cerrados sin duplicados,
 * ordena cronológicamente todos los periodos y actualiza el Saldo_Anterior
 * del periodo ABIERTO actual.
 */
function importarPeriodosHistoricos_(listaPeriodos) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const periodosSheet = ss.getSheetByName('Periodos');
    const cuotasSheet = ss.getSheetByName('Cuotas');
    const gastosSheet = ss.getSheetByName('Gastos');

    const idsImportar = new Set(listaPeriodos.map((p) => String(p.id)));

    // 1. Limpiar registros previos con los mismos IDs para evitar duplicados
    const filtrarFilas = (sheet, colIndex) => {
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return [data[0]];
      const conservadas = [data[0]];
      for (let i = 1; i < data.length; i++) {
        if (!idsImportar.has(String(data[i][colIndex]))) {
          conservadas.push(data[i]);
        }
      }
      return conservadas;
    };

    const pConservados = filtrarFilas(periodosSheet, 0);
    const cConservadas = filtrarFilas(cuotasSheet, 1);
    const gConservados = filtrarFilas(gastosSheet, 1);

    // 2. Construir filas nuevas
    const nuevasPeriodos = [];
    const nuevasCuotas = [];
    const nuevosGastos = [];

    listaPeriodos.forEach((p) => {
      const codeMes = p.id.replace('PER-', '').replace('-', '');
      nuevasPeriodos.push([
        p.id,
        "'" + p.mesAnio,
        p.responsableId,
        Number(p.saldoAnterior),
        Number(p.totalIngresos),
        Number(p.totalGastos),
        Number(p.saldoCierre),
        'CERRADO',
        "'" + p.fechaCierre,
      ]);

      p.cuotas.forEach((c, idx) => {
        nuevasCuotas.push([
          `CUO-${codeMes}-0${idx + 1}`,
          p.id,
          c.intId,
          c.nombre,
          Number(c.monto),
          Boolean(c.pagado),
          c.fechaPago ? "'" + c.fechaPago : '',
        ]);
      });

      p.gastos.forEach((g, idx) => {
        nuevosGastos.push([
          `GAS-${codeMes}-0${idx + 1}`,
          p.id,
          g.cat,
          g.sub,
          g.desc,
          Number(g.monto),
          Boolean(g.pagado),
          g.fechaPago ? "'" + g.fechaPago : '',
        ]);
      });
    });

    // 3. Combinar y ordenar Periodos cronológicamente por ID_Periodo (PER-AAAA-MM)
    const pHeader = pConservados[0];
    const pBody = pConservados.slice(1).concat(nuevasPeriodos);
    pBody.sort((a, b) => String(a[0]).localeCompare(String(b[0])));

    // 4. Sincronizar Saldo_Anterior del periodo ABIERTO con el último CERRADO
    let ultimoSaldoCierre = 0;
    for (let i = 0; i < pBody.length; i++) {
      if (String(pBody[i][7]) === 'CERRADO') {
        ultimoSaldoCierre = Number(pBody[i][6]) || 0;
      } else if (String(pBody[i][7]) === 'ABIERTO') {
        pBody[i][3] = ultimoSaldoCierre;
        pBody[i][6] = ultimoSaldoCierre + (Number(pBody[i][4]) || 0) - (Number(pBody[i][5]) || 0);
      }
    }

    // 5. Escribir en bloque en las 3 hojas
    const reescribirHoja = (sheet, filas, monedaCols) => {
      sheet.clearContents();
      sheet.getRange(1, 1, filas.length, filas[0].length).setValues(filas);
      if (filas.length > 1) {
        monedaCols.forEach((colRange) => {
          sheet
            .getRange(`${colRange}2:${colRange}${filas.length}`)
            .setNumberFormat('"S/ "#,##0.00');
        });
      }
    };

    reescribirHoja(periodosSheet, [pHeader].concat(pBody), ['D', 'E', 'F', 'G']);
    reescribirHoja(cuotasSheet, cConservadas.concat(nuevasCuotas), ['E']);
    reescribirHoja(gastosSheet, gConservados.concat(nuevosGastos), ['F']);

    // Recalcular totales del periodo abierto por si ya tenía movimientos registrados
    const abierto = pBody.find((r) => String(r[7]) === 'ABIERTO');
    if (abierto) recalcularTotalesPeriodo_(String(abierto[0]));

    Logger.log(
      `Importación completada: ${listaPeriodos.length} periodos históricos cargados. Pozo actual heredado: S/ ${ultimoSaldoCierre.toFixed(2)}`
    );
  } finally {
    lock.releaseLock();
  }
}

/**
 * Devuelve métricas agregadas y series temporales de todos los periodos
 * (cerrados + abierto) para la vista de Analítica Histórica.
 */
/**
 * Devuelve métricas agregadas, series temporales globales y series mensuales
 * desglosadas por cada servicio básico para la vista de Analítica Histórica.
 */
function getAnaliticaData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const periodosRows = ss.getSheetByName('Periodos').getDataRange().getValues();
  const cuotasRows = ss.getSheetByName('Cuotas').getDataRange().getValues();
  const gastosRows = ss.getSheetByName('Gastos').getDataRange().getValues();

  // 1. Ordenar periodos cronológicamente por ID (PER-AAAA-MM)
  const periodos = [];
  for (let i = 1; i < periodosRows.length; i++) {
    const r = periodosRows[i];
    if (!r[0]) continue;
    periodos.push({
      id: String(r[0]),
      mesAnio: formatMesAnio_(r[1]),
      saldoAnterior: Number(r[3]) || 0,
      totalIngresos: Number(r[4]) || 0,
      totalGastos: Number(r[5]) || 0,
      saldoCierre: Number(r[6]) || 0,
      estado: String(r[7]),
    });
  }
  periodos.sort((a, b) => a.id.localeCompare(b.id));

  const periodosConGasto = periodos.filter((p) => p.totalGastos > 0);
  const idsValidos = new Set(periodosConGasto.map((p) => p.id));

  // Inicializar matriz mes x servicio
  const seriesPorServicio = {
    'Luz (Enel)': {},
    'Agua (Sedapal)': {},
    'Internet (Win)': {},
    'Gas (Total 4 Pisos)': {},
  };

  // 2. Agrupar gastos por categoría, piso y mes
  const porServicio = {};
  for (let i = 1; i < gastosRows.length; i++) {
    const r = gastosRows[i];
    const pId = String(r[1]);
    const pagado = r[6] === true || String(r[6]).toUpperCase() === 'TRUE';
    const monto = Number(r[5]) || 0;
    if (!idsValidos.has(pId) || !pagado || monto <= 0) continue;

    const cat = String(r[2]);
    const desc = String(r[4]);
    let clave;

    if (cat === 'Gas') {
      const matchPiso = desc.match(/Piso\s*\d+/i);
      clave = matchPiso ? `Gas (${matchPiso[0]})` : 'Gas Cálidda';
      seriesPorServicio['Gas (Total 4 Pisos)'][pId] =
        (seriesPorServicio['Gas (Total 4 Pisos)'][pId] || 0) + monto;
    } else if (cat === 'Luz') {
      clave = 'Luz (Enel)';
    } else if (cat === 'Agua') {
      clave = 'Agua (Sedapal)';
    } else if (cat === 'Internet') {
      clave = 'Internet (Win)';
    } else {
      clave = 'Extras';
    }

    if (!seriesPorServicio[clave]) seriesPorServicio[clave] = {};
    seriesPorServicio[clave][pId] = (seriesPorServicio[clave][pId] || 0) + monto;

    if (!porServicio[clave]) {
      porServicio[clave] = { nombre: clave, total: 0, meses: 0 };
    }
    porServicio[clave].total += monto;
    porServicio[clave].meses += 1;
  }

  // Convertir mapas por periodo a arreglos alineados con periodosConGasto
  const seriesAlineadas = {};
  Object.keys(seriesPorServicio).forEach((srv) => {
    const valores = periodosConGasto.map((p) =>
      Number((seriesPorServicio[srv][p.id] || 0).toFixed(2))
    );
    const activos = valores.filter((v) => v > 0);
    const suma = activos.reduce((a, b) => a + b, 0);
    seriesAlineadas[srv] = {
      valores,
      promedio: activos.length > 0 ? Number((suma / activos.length).toFixed(2)) : 0,
      minimo: activos.length > 0 ? Math.min(...activos) : 0,
      maximo: activos.length > 0 ? Math.max(...activos) : 0,
    };
  });

  const serviciosLista = Object.values(porServicio)
    .map((s) => ({
      nombre: s.nombre,
      total: Number(s.total.toFixed(2)),
      promedio: Number((s.total / Math.max(1, periodosConGasto.length)).toFixed(2)),
    }))
    .sort((a, b) => b.total - a.total);

  // 3. Agrupar aportes pagados por integrante
  const porIntegrante = {};
  for (let i = 1; i < cuotasRows.length; i++) {
    const r = cuotasRows[i];
    const pagado = r[5] === true || String(r[5]).toUpperCase() === 'TRUE';
    const monto = Number(r[4]) || 0;
    if (!pagado || monto <= 0) continue;

    const nombre = String(r[3]);
    if (!porIntegrante[nombre]) {
      porIntegrante[nombre] = { nombre, total: 0, aportes: 0 };
    }
    porIntegrante[nombre].total += monto;
    porIntegrante[nombre].aportes += 1;
  }

  const integrantesLista = Object.values(porIntegrante)
    .map((m) => ({
      nombre: m.nombre,
      total: Number(m.total.toFixed(2)),
      aportes: m.aportes,
    }))
    .sort((a, b) => b.total - a.total);

  // 4. KPIs globales
  const totalGastadoHistorico = periodosConGasto.reduce((acc, p) => acc + p.totalGastos, 0);
  const promedioGastoMensual =
    periodosConGasto.length > 0 ? totalGastadoHistorico / periodosConGasto.length : 0;

  return JSON.parse(
    JSON.stringify({
      serieMensual: periodos,
      mesesValidosLabels: periodosConGasto.map((p) => p.mesAnio.split(' ')[0].substring(0, 3)),
      seriesPorServicio: seriesAlineadas,
      servicios: serviciosLista,
      integrantes: integrantesLista,
      kpis: {
        mesesAnalizados: periodosConGasto.length,
        totalGastado: Number(totalGastadoHistorico.toFixed(2)),
        promedioMensual: Number(promedioGastoMensual.toFixed(2)),
      },
    })
  );
}
