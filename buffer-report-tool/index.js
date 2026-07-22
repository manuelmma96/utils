'use strict';

const fs = require('fs');
const path = require('path');
const JSONStream = require('JSONStream');
const { writeReport } = require('./excel-report');

// --- GPS dataset conventions (expected field names in the source JSON) ---
const IGNITION_PROPERTY_NAME = 'Ignition';
const TAIL_INSPECT_CHUNK_BYTES = 256; // enough to see the file's last few characters

const MIN_DATE = new Date(-8640000000000000); // JS Date's representable range bounds
const MAX_DATE = new Date(8640000000000000);

function getArg(flag, fallback) {
  const idx = process.argv.indexOf(flag);
  if (idx !== -1 && process.argv[idx + 1]) return process.argv[idx + 1];
  return fallback;
}

function parseDateArg(flag) {
  const raw = getArg(flag, null);
  if (!raw) return null;
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) {
    console.error(`[ERROR] ${flag} "${raw}" no es una fecha ISO válida (ej: 2026-07-01T00:00:00.000Z).`);
    process.exit(1);
  }
  return date;
}

const THRESHOLD_MINUTES = Number(getArg('--threshold', 5));
const INPUT_DIR = path.resolve(getArg('--input', path.join(__dirname, '..', 'gps_packets_json')));
const OUTPUT_FILE = path.resolve(
  getArg('--output', path.join(__dirname, '..', 'output', 'reporte_consolidado_gps.xlsx'))
);

// Optional date filter: with no --from/--to, no record is discarded.
const RANGE_START = parseDateArg('--from') || MIN_DATE;
const RANGE_END = parseDateArg('--to') || MAX_DATE;
const DATE_FILTER_ACTIVE = RANGE_START !== MIN_DATE || RANGE_END !== MAX_DATE;

const BUCKETS = [
  { label: '5-15 min', min: 5, max: 15 },
  { label: '15-30 min', min: 15, max: 30 },
  { label: '30-60 min', min: 30, max: 60 },
  { label: '1-3 h', min: 60, max: 180 },
  { label: '3-24 h', min: 180, max: 1440 },
  { label: '>24 h', min: 1440, max: Infinity },
];

/**
 * Derives a document's "createdAt" from its MongoDB ObjectId.
 * The first 4 bytes (8 hex chars) of an ObjectId are seconds since epoch,
 * assigned when the id was generated — used here as a proxy for "when this
 * packet was saved to Mongo", since the source documents have no explicit
 * createdAt field.
 */
function decodeObjectIdTimestamp(oidHex) {
  if (!oidHex || oidHex.length < 8) return null;
  const seconds = parseInt(oidHex.substring(0, 8), 16);
  if (Number.isNaN(seconds)) return null;
  return new Date(seconds * 1000);
}

/**
 * Reads only the file's last bytes (no full load) to detect a truncated/
 * malformed JSON array.
 * @returns {'ok'|'mild'|'severe'}
 *   'ok'     - ends in "]", root array closed properly.
 *   'mild'   - missing "]" but the last document closes with "}" (a complete
 *              document, only the root array's closing bracket is missing;
 *              JSONStream still parses every document without loss).
 *   'severe' - ends in neither "}" nor "]", the cut happened mid-document.
 */
function inspectFileTail(filePath) {
  const { size } = fs.statSync(filePath);
  const readSize = Math.min(TAIL_INSPECT_CHUNK_BYTES, size);
  const buffer = Buffer.alloc(readSize);
  const fd = fs.openSync(filePath, 'r');
  try {
    fs.readSync(fd, buffer, 0, readSize, size - readSize);
  } finally {
    fs.closeSync(fd);
  }
  const tail = buffer.toString('utf8').trimEnd();
  if (tail.endsWith(']')) return 'ok';
  if (tail.endsWith('}')) return 'mild';
  return 'severe';
}

function bucketFor(minutes) {
  const b = BUCKETS.find((x) => minutes >= x.min && minutes < x.max);
  return b ? b.label : null;
}

function median(sortedValues) {
  if (sortedValues.length === 0) return null;
  const mid = Math.floor(sortedValues.length / 2);
  if (sortedValues.length % 2 === 0) {
    return (sortedValues[mid - 1] + sortedValues[mid]) / 2;
  }
  return sortedValues[mid];
}

/**
 * Streams a single GPS packets JSON file (JSONStream, no full-file JSON.parse)
 * and computes, per dataRecord: gap = createdAt (derived from _id) - timestamp.
 * Only rows above THRESHOLD_MINUTES are kept in memory for the detail sheet;
 * this keeps memory bounded even for files with millions of lines.
 */
function processFile(filePath) {
  return new Promise((resolve, reject) => {
    const fileName = path.basename(filePath);
    let imei = null;
    let totalDocs = 0;
    let totalRecords = 0;
    let outOfRangeCount = 0;
    let bufferedCount = 0;
    let anomalyCount = 0; // negative gap: fecha_gps_fix later than createdAt_derivado
    let sumGapMinutes = 0;
    let maxGapMinutes = -Infinity;
    const gapValues = []; // buffered records only, for the median
    const bufferedRows = [];
    const bucketCounts = {};
    BUCKETS.forEach((b) => { bucketCounts[b.label] = 0; });

    const stream = fs.createReadStream(filePath, { encoding: 'utf8' });
    const parser = JSONStream.parse('*');

    parser.on('data', (doc) => {
      totalDocs += 1;
      if (!imei && doc.imei) imei = doc.imei;

      const oid = doc._id && doc._id.$oid;
      const createdAt = decodeObjectIdTimestamp(oid);
      if (!createdAt) return;

      const records = (doc.data && doc.data.dataRecords) || [];
      const hasMultiple = records.length > 1;

      records.forEach((record, idx) => {
        if (!record.timestamp) return;
        const gpsFix = new Date(record.timestamp);
        if (Number.isNaN(gpsFix.getTime())) return;

        if (gpsFix < RANGE_START || gpsFix >= RANGE_END) {
          outOfRangeCount += 1;
          return;
        }

        totalRecords += 1;

        const gapMinutes = (createdAt.getTime() - gpsFix.getTime()) / 60000;

        if (gapMinutes < 0) {
          anomalyCount += 1;
          return;
        }

        if (gapMinutes > THRESHOLD_MINUTES) {
          bufferedCount += 1;
          sumGapMinutes += gapMinutes;
          if (gapMinutes > maxGapMinutes) maxGapMinutes = gapMinutes;
          gapValues.push(gapMinutes);

          const label = bucketFor(gapMinutes);
          if (label) bucketCounts[label] += 1;

          const ignitionProp = (record.ioProperties || []).find((p) => p.name === IGNITION_PROPERTY_NAME);

          bufferedRows.push({
            imei: doc.imei || imei || '',
            docId: oid || '',
            recordIndex: hasMultiple ? idx + 1 : '',
            fechaGpsFix: gpsFix,
            createdAtDerivado: createdAt,
            bufferMinutos: Number(gapMinutes.toFixed(2)),
            latitude: record.latitude,
            longitude: record.longitude,
            speed: record.speed,
            ignition: ignitionProp ? ignitionProp.value : '',
          });
        }
      });
    });

    parser.on('error', (err) => reject(err));
    stream.on('error', (err) => reject(err));

    parser.on('end', () => {
      gapValues.sort((a, b) => a - b);
      resolve({
        fileName,
        imei: imei || fileName.replace(/\.json$/i, ''),
        totalDocs,
        totalRecords,
        outOfRangeCount,
        bufferedCount,
        anomalyCount,
        maxGapMinutes: bufferedCount > 0 ? Number(maxGapMinutes.toFixed(2)) : 0,
        avgGapMinutes: bufferedCount > 0 ? Number((sumGapMinutes / bufferedCount).toFixed(2)) : 0,
        medianGapMinutes: bufferedCount > 0 ? Number(median(gapValues).toFixed(2)) : 0,
        bucketCounts,
        bufferedRows,
      });
    });

    stream.pipe(parser);
  });
}

async function main() {
  if (!fs.existsSync(INPUT_DIR)) {
    console.error(`No existe el directorio de entrada: ${INPUT_DIR}`);
    process.exit(1);
  }

  fs.mkdirSync(path.dirname(OUTPUT_FILE), { recursive: true });

  const files = fs.readdirSync(INPUT_DIR).filter((f) => f.toLowerCase().endsWith('.json'));
  if (files.length === 0) {
    console.error(`No se encontraron archivos .json en ${INPUT_DIR}`);
    process.exit(1);
  }

  console.log(`Directorio de entrada: ${INPUT_DIR}`);
  console.log(`Archivos a procesar: ${files.join(', ')}`);
  console.log(`Umbral de buffer: ${THRESHOLD_MINUTES} minutos`);
  console.log(
    DATE_FILTER_ACTIVE
      ? `Filtro de fecha activo: [${RANGE_START.toISOString()}, ${RANGE_END.toISOString()})`
      : 'Filtro de fecha: ninguno (se procesan todos los records)'
  );

  const results = [];
  for (const file of files) {
    const filePath = path.join(INPUT_DIR, file);
    const tailStatus = inspectFileTail(filePath);

    if (tailStatus === 'severe') {
      console.error(
        `[ERROR] ${file}: el JSON está severamente mal formado (el corte no ocurrió al ` +
        'final de un documento completo). Procesarlo podría perder u omitir registros sin ' +
        'aviso. Se detiene la generación del reporte — corrige o excluye este archivo antes ' +
        'de reintentar.'
      );
      process.exit(1);
    }

    console.log(`Procesando ${file}...`);
    if (tailStatus === 'mild') {
      console.warn(
        `  [advertencia] ${file}: al archivo le falta el "]" de cierre del arreglo raíz, ` +
        'pero el último documento está completo. Se procesa con normalidad.'
      );
    }

    const start = Date.now();
    const result = await processFile(filePath);
    const elapsedSec = ((Date.now() - start) / 1000).toFixed(1);
    const outOfRangeNote = DATE_FILTER_ACTIVE ? ` (${result.outOfRangeCount} fuera de rango, descartados)` : '';
    console.log(
      `  -> docs=${result.totalDocs} records=${result.totalRecords}${outOfRangeNote} ` +
      `buffer>${THRESHOLD_MINUTES}min=${result.bufferedCount} anomalías=${result.anomalyCount} (${elapsedSec}s)`
    );
    results.push(result);
  }

  await writeReport({
    outputFile: OUTPUT_FILE,
    results,
    thresholdMinutes: THRESHOLD_MINUTES,
    ignitionPropertyName: IGNITION_PROPERTY_NAME,
    dateFilterActive: DATE_FILTER_ACTIVE,
    buckets: BUCKETS,
  });
  console.log(`Reporte generado en: ${OUTPUT_FILE}`);
}

main().catch((err) => {
  console.error('Error generando el reporte:', err);
  process.exit(1);
});
