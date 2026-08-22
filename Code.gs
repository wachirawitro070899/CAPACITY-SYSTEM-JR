const SPREADSHEET_ID = '1sKHUxWULtgUedTBuI_a41FU5WkCASTSuXTis0t12XRI';
const SOURCE_GID = 1349772114;
const MAX_HEADER_SCAN_ROWS = 60;

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || '').trim().toLowerCase();
  if (action === 'capacity') {
    const data = getCapacityData();
    const callback = String((e && e.parameter && e.parameter.callback) || '').trim();
    const json = JSON.stringify(data);
    if (callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
      return ContentService.createTextOutput(callback + '(' + json + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    service: 'Capacity System API',
    spreadsheetId: SPREADSHEET_ID,
    sourceGid: SOURCE_GID,
    message: 'Use ?action=capacity'
  })).setMimeType(ContentService.MimeType.JSON);
}

function getCapacityData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getSheetById_(ss, SOURCE_GID);
    if (!sheet) throw new Error('ไม่พบชีต gid=' + SOURCE_GID + ' ใน Google Sheet ตัวใหม่');

    const values = sheet.getDataRange().getDisplayValues();
    const header = detectHeader_(values);
    if (!header) throw new Error('ไม่พบหัวตาราง Part No. / Process / M/C / Step ในชีต ' + sheet.getName());

    const records = [];
    const diagnostics = [];
    let lastPartNo = '';
    let lastPartName = '';
    let lastProcess = '';
    let lastItem = '';

    for (let r = header.row + 1; r < values.length; r++) {
      const row = values[r] || [];
      if (!row.some(v => clean_(v))) continue;

      let item = clean_(getByAlias_(row, header.map, ALIASES.item));
      let partNo = clean_(getByAlias_(row, header.map, ALIASES.partNo));
      let partName = clean_(getByAlias_(row, header.map, ALIASES.partName));
      let process = clean_(getByAlias_(row, header.map, ALIASES.process));
      let step = clean_(getByAlias_(row, header.map, ALIASES.step));
      const machineRaw = clean_(getByAlias_(row, header.map, ALIASES.machine));

      // รองรับ Google Sheet ที่ merge cell หรือกรอก Part/Process เฉพาะบรรทัดแรก
      if (!partNo && (machineRaw || process || step)) partNo = lastPartNo;
      if (!partName && partNo === lastPartNo) partName = lastPartName;
      if (!process && (machineRaw || step)) process = lastProcess;
      if (!item && partNo === lastPartNo) item = lastItem;

      if (partNo) lastPartNo = partNo;
      if (partName) lastPartName = partName;
      if (process) lastProcess = process;
      if (item) lastItem = item;

      if (!partNo && !machineRaw && !process && !step) continue;

      const speedMinPerPc = number_(getByAlias_(row, header.map, ALIASES.speedMinPerPc));
      let ct = number_(getByAlias_(row, header.map, ALIASES.ct));
      if (!(ct > 0) && speedMinPerPc > 0) ct = speedMinPerPc * 60;

      const outputCycle = positive_(number_(getByAlias_(row, header.map, ALIASES.outputCycle)), 1);
      const efficiency = percent_(getByAlias_(row, header.map, ALIASES.efficiency), 100);
      const hoursPerShift = positive_(number_(getByAlias_(row, header.map, ALIASES.hoursPerShift)), 8);
      const shiftsPerDay = positive_(number_(getByAlias_(row, header.map, ALIASES.shiftsPerDay)), 2);
      const cap100 = number_(getByAlias_(row, header.map, ALIASES.cap100));
      const cap90 = number_(getByAlias_(row, header.map, ALIASES.cap90));
      const cap85 = number_(getByAlias_(row, header.map, ALIASES.cap85));

      const machineList = splitMachines_(machineRaw);
      const machines = machineList.length ? machineList : ['ไม่ระบุเครื่อง'];

      machines.forEach(machine => {
        records.push({
          item,
          machine,
          machineRaw,
          sheetName: sheet.getName(),
          sheetId: sheet.getSheetId(),
          row: r + 1,
          partNo,
          partName,
          process,
          step,
          speedMinPerPc,
          ct,
          outputCycle,
          efficiency,
          eff: efficiency,
          hoursPerShift,
          shiftsPerDay,
          cap100,
          cap90,
          cap85,
          status: clean_(getByAlias_(row, header.map, ALIASES.status)) || 'Active',
          remark: clean_(getByAlias_(row, header.map, ALIASES.remark))
        });
      });
    }

    const machineSheets = [...new Set(records.map(r => r.machine).filter(Boolean))];
    diagnostics.push({
      sheet: sheet.getName(),
      gid: SOURCE_GID,
      status: records.length ? 'ok' : 'empty',
      headerRow: header.row + 1,
      rows: records.length,
      columnsDetected: header.found
    });

    return {
      ok: true,
      spreadsheetId: SPREADSHEET_ID,
      sourceGid: SOURCE_GID,
      spreadsheetName: ss.getName(),
      sourceSheetName: sheet.getName(),
      machineSheets,
      records,
      diagnostics,
      generatedAt: new Date().toISOString()
    };
  } catch (err) {
    return {
      ok: false,
      error: String(err && err.message ? err.message : err),
      spreadsheetId: SPREADSHEET_ID,
      sourceGid: SOURCE_GID,
      records: [],
      machineSheets: [],
      diagnostics: [],
      generatedAt: new Date().toISOString()
    };
  }
}

function getSheetById_(ss, gid) {
  const target = Number(gid);
  return ss.getSheets().find(s => Number(s.getSheetId()) === target) || null;
}

const ALIASES = {
  item: ['Item', 'No.', 'No', 'ลำดับ'],
  partNo: ['Part No.', 'Part No', 'PartNo', 'Part Number', 'Part_Number', 'Material No', 'Item No', 'Product No'],
  partName: ['Part Name.', 'Part Name', 'PartName', 'Part Description', 'Description', 'Product Name', 'Item Name', 'Name'],
  process: ['Process', 'Operation', 'Process Name', 'Operation Name', 'Process/Operation', 'Secondary Process'],
  machine: ['M/C', 'MC', 'M.C.', 'Machine', 'Machine No.', 'Machine No', 'Machine Name', 'Equipment'],
  step: ['Step', 'Process Step', 'Operation Step', 'Step No', 'Step No.', 'Process No', 'Process No.', 'Sequence', 'Seq'],
  speedMinPerPc: ['speed 1 min./pcs', 'Speed 1 min./pcs', '1 min/pcs', 'min/pcs', 'Min/Pcs', 'Production Time (min)', 'Production Time (minutes)', '生产工时（分钟）', '生产工时(分钟)'],
  ct: ['CT (sec/pc)', 'CT (sec)', 'CT (s)', 'CT', 'Cycle Time', 'Cycle Time (sec)', 'Cycle Time (s)', 'Time (sec)'],
  cap100: ['100%', 'Capacity 100%', '100 %'],
  cap90: ['90%', 'Capacity 90%', '90 %'],
  cap85: ['85%', 'Capacity 85%', '85 %'],
  outputCycle: ['Output/Cycle', 'Output per Cycle', 'Output / Cycle', 'Qty/Cycle', 'Output per cycle #1', 'No. of unit #1', 'No of unit #1', 'No. of unit', 'No of unit'],
  efficiency: ['Efficiency %', 'Eff %', 'Efficiency', 'Eff', 'Eff % #1', 'Efficiency #1'],
  hoursPerShift: ['Working Hours/Shift', 'Hours/Shift', 'Hours', 'Daily Working Hrs #1', 'Daily Working Hrs', 'Working Hours'],
  shiftsPerDay: ['Shifts/Day', 'Shift/Day', 'Shifts', 'No. of Shift', 'No of Shift'],
  status: ['Status', 'Active/Inactive'],
  remark: ['Remark', 'Remarks', 'Note', 'Notes', 'Comment']
};

function detectHeader_(values) {
  let best = null;
  const max = Math.min(MAX_HEADER_SCAN_ROWS, values.length);

  for (let r = 0; r < max; r++) {
    const map = {};
    (values[r] || []).forEach((v, c) => {
      const key = norm_(v);
      if (key && map[key] === undefined) map[key] = c;
    });

    const found = {};
    Object.keys(ALIASES).forEach(key => {
      const col = aliasCol_(map, ALIASES[key]);
      if (col !== undefined) found[key] = col;
    });

    let score = 0;
    if (found.partNo !== undefined) score += 12;
    if (found.partName !== undefined) score += 3;
    if (found.process !== undefined) score += 6;
    if (found.machine !== undefined) score += 10;
    if (found.step !== undefined) score += 6;
    if (found.speedMinPerPc !== undefined || found.ct !== undefined) score += 5;
    if (found.cap100 !== undefined) score += 2;

    const valid = found.partNo !== undefined && (found.machine !== undefined || found.process !== undefined || found.step !== undefined);
    if (valid && (!best || score > best.score)) best = { row: r, map, score, found };
  }
  return best;
}

function splitMachines_(value) {
  const s = clean_(value);
  if (!s) return [];
  return [...new Set(s.split(/[\n,;]+/).map(clean_).filter(Boolean))];
}

function aliasCol_(map, aliases) {
  for (const alias of aliases) {
    const key = norm_(alias);
    if (map[key] !== undefined) return map[key];
  }
  const entries = Object.keys(map);
  for (const alias of aliases) {
    const target = loose_(alias);
    for (const key of entries) {
      if (loose_(key) === target) return map[key];
    }
  }
}

function getByAlias_(row, map, aliases) {
  const col = aliasCol_(map, aliases);
  return col === undefined ? '' : (row[col] ?? '');
}

function norm_(v) {
  return String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').trim().toLowerCase().replace(/[_-]+/g, ' ').replace(/\s+/g, ' ');
}

function loose_(v) {
  return norm_(v).replace(/#\s*\d+/g, '').replace(/[()./%]/g, ' ').replace(/\s+/g, ' ').trim();
}

function clean_(v) { return String(v == null ? '' : v).trim(); }

function number_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const s = clean_(v).replace(/,/g, '').replace(/%/g, '');
  const n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

function positive_(n, fallback) { return n > 0 ? n : fallback; }

function percent_(v, fallback) {
  const s = clean_(v);
  if (!s) return fallback;
  let n = number_(s);
  if (n > 0 && n <= 1 && s.indexOf('%') === -1) n *= 100;
  return n > 0 ? n : fallback;
}
