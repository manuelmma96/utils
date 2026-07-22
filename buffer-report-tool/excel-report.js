'use strict';

const ExcelJS = require('exceljs');

// --- .xlsx format limits (Excel's constraints, not our logic) ---
const EXCEL_ROW_LIMIT = 1048576;
const EXCEL_ROW_SAFETY_MARGIN = 48576; // keep some distance from Excel's real limit
const MAX_ROWS_PER_SHEET = EXCEL_ROW_LIMIT - EXCEL_ROW_SAFETY_MARGIN;
const EXCEL_SHEET_NAME_MAX_LENGTH = 31; // Excel's hard limit for sheet names
const EXCEL_SHEET_NAME_FORBIDDEN_CHARS = /[:\\/?*[\]]/g;

function sanitizeSheetName(name) {
  const cleaned = String(name).replace(EXCEL_SHEET_NAME_FORBIDDEN_CHARS, '_');
  return cleaned.substring(0, EXCEL_SHEET_NAME_MAX_LENGTH);
}

function buildDetailNotes(thresholdMinutes, ignitionPropertyName) {
  return {
    imei: 'Identificador único del dispositivo GPS (IMEI del equipo que envió el paquete).',
    docId:
      'Equivalente al _id de MongoDB/Mongoose. De aquí se deriva "Fin": los primeros 8 ' +
      'caracteres hex del ObjectId son segundos desde epoch Unix, generados por Mongo al ' +
      'insertar el documento.',
    recordIndex:
      'Solo se llena cuando el documento trae más de un dataRecord agrupado en el mismo ' +
      'paquete (recordsNumber > 1). Indica la posición del registro dentro del paquete.',
    fechaGpsFix:
      'Inicio del buffer: fecha/hora en que el dispositivo tomó el fix GPS. Viene tal cual ' +
      'del campo "timestamp" del dataRecord, sin cálculos.',
    createdAtDerivado:
      'Fin del buffer: fecha/hora en que el paquete se guardó en Mongo. NO es un campo ' +
      'explícito del documento: se calcula tomando los primeros 8 caracteres hex del ' +
      'doc_id (ObjectId) y se interpretan como segundos desde epoch Unix.',
    bufferMinutos:
      'Duración del buffer en minutos = (Fin - Inicio) / 60000. Es cuánto tiempo pasó ' +
      'desde que se tomó el fix GPS hasta que el paquete llegó/se guardó en la base de ' +
      `datos. Solo se listan filas donde esta duración supera el umbral (${thresholdMinutes} min).`,
    latitude: 'Latitud del fix GPS, tal cual viene del dataRecord.',
    longitude: 'Longitud del fix GPS, tal cual viene del dataRecord.',
    speed: 'Velocidad reportada por el dispositivo en el momento del fix (km/h).',
    ignition:
      `Estado del contacto/ignición del vehículo (io property "${ignitionPropertyName}"): ` +
      '1 = encendido, 0 = apagado, vacío si el paquete no trae esa propiedad.',
  };
}

function buildSummaryNotes(thresholdMinutes, dateFilterActive) {
  return {
    imei: 'IMEI del dispositivo (tomado del primer documento leído del archivo); si no viene, se usa el nombre del archivo.',
    totalDocs: 'Cantidad de documentos (paquetes) leídos del archivo JSON.',
    totalRecords:
      'Cantidad de registros GPS (dataRecords) analizados, sumando todos los documentos del ' +
      'archivo' + (dateFilterActive ? ' que caen dentro del rango de fechas filtrado.' : '.'),
    bufferedCount:
      'Cantidad de registros cuya brecha (createdAt_derivado - fecha_gps_fix) superó el ' +
      `umbral configurado (${thresholdMinutes} min). Son los que aparecen en la hoja de detalle.`,
    pctBuffered: 'Porcentaje de records con buffer sobre el total de records del archivo.',
    maxGapMinutes: 'Brecha máxima (en minutos) encontrada, calculada solo sobre records que superaron el umbral.',
    avgGapMinutes: 'Brecha promedio (en minutos), calculada solo sobre records que superaron el umbral.',
    medianGapMinutes: 'Brecha mediana (en minutos), calculada solo sobre records que superaron el umbral.',
    anomalyCount:
      'Registros donde fecha_gps_fix es posterior a createdAt_derivado (brecha negativa). ' +
      'No debería ocurrir en condiciones normales; puede indicar desfase de reloj del ' +
      'dispositivo o del servidor. No se incluyen en la hoja de detalle.',
  };
}

const BUCKET_NOTE = 'Cantidad de registros (dentro de los que superaron el umbral) cuya brecha en minutos cae en este rango.';

function applyHeaderNotes(sheet, notesMap) {
  const headerRow = sheet.getRow(1);
  sheet.columns.forEach((column, idx) => {
    const note = notesMap[column.key];
    if (note) headerRow.getCell(idx + 1).note = note;
  });
}

function addBufferedSheets(workbook, result, { thresholdMinutes, ignitionPropertyName }) {
  const { imei, bufferedRows } = result;
  const detailNotes = buildDetailNotes(thresholdMinutes, ignitionPropertyName);
  const totalChunks = Math.max(1, Math.ceil(bufferedRows.length / MAX_ROWS_PER_SHEET));

  if (totalChunks > 1) {
    console.warn(
      `[aviso] ${result.fileName}: ${bufferedRows.length} filas con buffer exceden ` +
      `${MAX_ROWS_PER_SHEET} filas por hoja, se dividirá en ${totalChunks} hojas.`
    );
  }

  for (let chunk = 0; chunk < totalChunks; chunk += 1) {
    const suffix = totalChunks > 1 ? `_pt${chunk + 1}` : '';
    const sheetName = sanitizeSheetName(`Buffer_${imei}${suffix}`);
    const sheet = workbook.addWorksheet(sheetName);

    sheet.columns = [
      { header: 'IMEI', key: 'imei', width: 18 },
      { header: 'doc_id', key: 'docId', width: 26 },
      { header: 'record_index', key: 'recordIndex', width: 12 },
      { header: 'Inicio', key: 'fechaGpsFix', width: 22 },
      { header: 'Fin', key: 'createdAtDerivado', width: 22 },
      { header: 'Duracion (minutos)', key: 'bufferMinutos', width: 18 },
      { header: 'latitude', key: 'latitude', width: 14 },
      { header: 'longitude', key: 'longitude', width: 14 },
      { header: 'speed', key: 'speed', width: 10 },
      { header: 'ignition', key: 'ignition', width: 10 },
    ];
    sheet.getRow(1).font = { bold: true };
    applyHeaderNotes(sheet, detailNotes);

    const slice = bufferedRows.slice(chunk * MAX_ROWS_PER_SHEET, (chunk + 1) * MAX_ROWS_PER_SHEET);
    slice.forEach((row) => {
      sheet.addRow({
        imei: row.imei,
        docId: row.docId,
        recordIndex: row.recordIndex,
        fechaGpsFix: row.fechaGpsFix.toISOString(),
        createdAtDerivado: row.createdAtDerivado.toISOString(),
        bufferMinutos: row.bufferMinutos,
        latitude: row.latitude,
        longitude: row.longitude,
        speed: row.speed,
        ignition: row.ignition,
      }).commit();
    });

    sheet.commit();
  }
}

function addSummarySheet(workbook, results, { thresholdMinutes, dateFilterActive, buckets }) {
  const sheet = workbook.addWorksheet('Resumen Ejecutivo');
  const summaryNotes = buildSummaryNotes(thresholdMinutes, dateFilterActive);

  const columns = [
    { header: 'IMEI / Archivo', key: 'imei', width: 20 },
    { header: 'Total documentos', key: 'totalDocs', width: 16 },
    { header: 'Total records', key: 'totalRecords', width: 14 },
    { header: 'Records con buffer >5min', key: 'bufferedCount', width: 22 },
    { header: '% con buffer', key: 'pctBuffered', width: 14 },
    { header: 'Brecha máx (min)', key: 'maxGapMinutes', width: 14 },
    { header: 'Brecha promedio (min)', key: 'avgGapMinutes', width: 16 },
    { header: 'Brecha mediana (min)', key: 'medianGapMinutes', width: 16 },
    { header: 'Anomalías (brecha negativa)', key: 'anomalyCount', width: 20 },
  ];
  buckets.forEach((b) => columns.push({ header: b.label, key: b.label, width: 12 }));

  sheet.columns = columns;
  sheet.getRow(1).font = { bold: true };
  buckets.forEach((b) => { summaryNotes[b.label] = BUCKET_NOTE; });
  applyHeaderNotes(sheet, summaryNotes);

  let grandTotalDocs = 0;
  let grandTotalRecords = 0;
  let grandBuffered = 0;
  let grandAnomalies = 0;
  const grandBuckets = {};
  buckets.forEach((b) => { grandBuckets[b.label] = 0; });

  results.forEach((r) => {
    const pct = r.totalRecords > 0 ? ((r.bufferedCount / r.totalRecords) * 100).toFixed(2) : '0.00';
    const row = {
      imei: r.imei,
      totalDocs: r.totalDocs,
      totalRecords: r.totalRecords,
      bufferedCount: r.bufferedCount,
      pctBuffered: `${pct}%`,
      maxGapMinutes: r.maxGapMinutes,
      avgGapMinutes: r.avgGapMinutes,
      medianGapMinutes: r.medianGapMinutes,
      anomalyCount: r.anomalyCount,
    };
    buckets.forEach((b) => { row[b.label] = r.bucketCounts[b.label]; });
    sheet.addRow(row).commit();

    grandTotalDocs += r.totalDocs;
    grandTotalRecords += r.totalRecords;
    grandBuffered += r.bufferedCount;
    grandAnomalies += r.anomalyCount;
    buckets.forEach((b) => { grandBuckets[b.label] += r.bucketCounts[b.label]; });
  });

  const totalPct = grandTotalRecords > 0 ? ((grandBuffered / grandTotalRecords) * 100).toFixed(2) : '0.00';
  const totalRow = {
    imei: 'TOTAL',
    totalDocs: grandTotalDocs,
    totalRecords: grandTotalRecords,
    bufferedCount: grandBuffered,
    pctBuffered: `${totalPct}%`,
    maxGapMinutes: '',
    avgGapMinutes: '',
    medianGapMinutes: '',
    anomalyCount: grandAnomalies,
  };
  buckets.forEach((b) => { totalRow[b.label] = grandBuckets[b.label]; });
  const addedTotalRow = sheet.addRow(totalRow);
  addedTotalRow.font = { bold: true };
  addedTotalRow.commit();

  sheet.commit();
}

/**
 * Builds and writes the consolidated .xlsx report (streaming writer, so memory
 * stays bounded regardless of how many buffered rows were collected).
 */
async function writeReport({ outputFile, results, thresholdMinutes, ignitionPropertyName, dateFilterActive, buckets }) {
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: outputFile, useStyles: true });

  addSummarySheet(workbook, results, { thresholdMinutes, dateFilterActive, buckets });
  results.forEach((result) => addBufferedSheets(workbook, result, { thresholdMinutes, ignitionPropertyName }));

  await workbook.commit();
}

module.exports = { writeReport };
