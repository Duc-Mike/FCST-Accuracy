import * as XLSX from 'xlsx';
import { MergedRow } from '../types';

export async function processExcelFiles(
  files: Record<string, File>,
  hasBaseData: boolean
): Promise<Uint8Array> {
  const allRows: MergedRow[] = [];

  // Process SO files
  if (files['SO']) {
    const soRows = await processSOFile(files['SO']);
    allRows.push(...soRows);
  }

  // Process SHIP files
  if (files['SHIP']) {
    const shipRows = await processShipFile(files['SHIP']);
    allRows.push(...shipRows);
  }

  // Process FCST/Base Data files
  const fcstFiles = hasBaseData 
    ? ['FCST', 'Base Data'] 
    : ['M FCST', 'M+1 FCST', 'M+2 FCST', 'M+3 FCST', 'M+4 FCST'];

  for (const key of fcstFiles) {
    if (files[key]) {
      const rows = await processFCSTFile(files[key]);
      allRows.push(...rows);
    }
  }

  // Filter out rows where Net price is 0
  const filteredRows = allRows.filter(row => row['Net price'] !== 0);

  // Create workbook
  const worksheet = XLSX.utils.json_to_sheet(filteredRows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'FCST SO Accuracy');

  // Generate buffer
  return XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
}

function getVal(row: any, headerName: string): any {
  const normalizedHeader = headerName.toLowerCase().trim();
  const key = Object.keys(row).find(k => k.toLowerCase().trim() === normalizedHeader);
  return key ? row[key] : undefined;
}

function trimStr(val: any): string {
  if (val === undefined || val === null) return '';
  return String(val).trim();
}

const CATEGORY_MAP: Record<string, string> = {
  'UV PKG': 'PKG',
  'Module': 'MD',
  'TL PKG': 'PKG',
  'Z-FILM': 'PKG',
  'PL PKG': 'PKG',
  'SL PKG': 'PKG',
  'CL PKG': 'PKG',
  'Parts': 'MD',
  'LP PKG': 'PKG',
  'SOC CHIP': 'CHIP',
  'VCESL': 'PKG',
  'Micro LED': 'PKG'
};

const DIVISION_MAP: Record<string, string> = {
  'SD99': 'AM',
  'SE01': 'AM',
  'SC01': 'AM',
  'SU01': 'AM',
  'SA02': 'AM',
  'VN14': 'AM'
};

const REGION_MAP: Record<string, string> = {
  'D10': 'AM.KR',
  'M05': 'AM.NA',
  'M01': 'AM.CN',
  'M03': 'AM.JP',
  'M04': 'AM.EU',
  'M02': 'AM.CN',
  'E12': 'AM.EU',
  'E10': 'AM.EU',
  'E11': 'AM.EU',
  'C11': 'AM.CN',
  'C26': 'AM.TW',
  'U11': 'AM.NA',
  'A13': 'AM.JP',
  'V71': 'AM.EU',
  'M06': 'AM.IN',
  'V74': 'AM.CN',
  'V73': 'AM.KR'
};

const REG_WEEK_TO_MONTH: Record<string, string> = {
  '2025.38': '2025.09',
  '2025.42': '2025.10',
  '2025.47': '2025.11',
  '2025.51': '2025.12',
  '2026.01': '2026.01',
  '2026.02': '2026.01',
  '2026.03': '2026.01',
  '2026.04': '2026.01',
  '2026.05': '2026.02',
  '2026.06': '2026.02',
  '2026.07': '2026.02',
  '2026.08': '2026.02',
  '2026.09': '2026.03',
  '2026.10': '2026.03',
  '2026.11': '2026.03',
  '2026.12': '2026.03',
  '2026.13': '2026.04',
  '2026.14': '2026.04',
  '2026.15': '2026.04',
  '2026.16': '2026.04',
  '2026.17': '2026.04',
  '2026.18': '2026.05',
  '2026.19': '2026.05',
  '2026.20': '2026.05',
  '2026.21': '2026.05',
  '2026.22': '2026.06',
  '2026.23': '2026.06',
  '2026.24': '2026.06',
  '2026.25': '2026.06',
  '2026.26': '2026.07',
  '2026.27': '2026.07',
  '2026.28': '2026.07',
  '2026.29': '2026.07',
  '2026.30': '2026.07',
  '2026.31': '2026.08',
  '2026.32': '2026.08',
  '2026.33': '2026.08',
  '2026.34': '2026.08',
  '2026.35': '2026.09',
  '2026.36': '2026.09',
  '2026.37': '2026.09',
  '2026.38': '2026.09',
  '2026.39': '2026.10',
  '2026.40': '2026.10',
  '2026.41': '2026.10',
  '2026.42': '2026.10',
  '2026.43': '2026.10',
  '2026.44': '2026.11',
  '2026.45': '2026.11',
  '2026.46': '2026.11',
  '2026.47': '2026.11',
  '2026.48': '2026.12',
  '2026.49': '2026.12',
  '2026.50': '2026.12',
  '2026.51': '2026.12',
  '2026.52': '2027.01'
};

function getCategoryValue(input: any): string {
  const val = trimStr(input);
  return CATEGORY_MAP[val] || '';
}

function getDivisionValue(office: any): string {
  const val = trimStr(office).toUpperCase();
  return DIVISION_MAP[val] || '';
}

function getRegionValue(group: any): string {
  const val = trimStr(group).toUpperCase();
  return REGION_MAP[val] || '';
}

function getRegistrationMonthValue(week: any): string {
  const val = trimStr(week);
  return REG_WEEK_TO_MONTH[val] || val;
}

function getNetPriceValue(val: any): number {
  return Number(val || 0);
}

function formatYearMonth(dateVal: any): string {
  if (dateVal === undefined || dateVal === null || dateVal === '') return '';
  
  // If it's already in YYYY.MM format (string)
  if (typeof dateVal === 'string' && /^\d{4}\.\d{2}$/.test(dateVal)) {
    return dateVal;
  }

  let date: Date;
  if (typeof dateVal === 'number') {
    const parsed = XLSX.SSF.parse_date_code(dateVal);
    return `${parsed.y}.${String(parsed.m).padStart(2, '0')}`;
  } else {
    date = new Date(dateVal);
  }

  if (isNaN(date.getTime())) {
    return String(dateVal);
  }

  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  return `${y}.${m}`;
}

function formatFCSTPlanMonth(monthVal: any, planMonth: any): string {
  const monthStr = trimStr(monthVal);
  if (!monthStr || planMonth === undefined || planMonth === null) return '';
  const year = monthStr.substring(0, 4);
  const m = String(planMonth).padStart(2, '0');
  return `${year}.${m}`;
}

function getMCode(regMonth: string, planMonth: string): string {
  if (!regMonth || !planMonth) return 'out';
  
  const rParts = regMonth.split('.');
  const pParts = planMonth.split('.');
  
  if (rParts.length < 2 || pParts.length < 2) return 'out';

  const ry = parseInt(rParts[0]);
  const rm = parseInt(rParts[1]);
  const py = parseInt(pParts[0]);
  const pm = parseInt(pParts[1]);
  
  if (isNaN(ry) || isNaN(rm) || isNaN(py) || isNaN(pm)) return 'out';

  const regTotal = ry * 12 + rm;
  const planTotal = py * 12 + pm;
  const diff = planTotal - regTotal;

  if (diff >= 1 && diff <= 4) return `M+${diff}`;
  return 'out';
}

async function processSOFile(file: File): Promise<MergedRow[]> {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<any>(sheet);

  return rows.map(row => ({
    Code: 'SO',
    'Registration month': '',
    'Plan month': formatYearMonth(getVal(row, 'Requested deliv. date')),
    Division: getDivisionValue(getVal(row, 'Sales office')),
    Region: getRegionValue(getVal(row, 'Sales group')),
    Material: trimStr(getVal(row, 'Material name')),
    Customer: trimStr(getVal(row, 'Customer name')),
    'End customer': trimStr(getVal(row, 'End customer name')),
    'Sales employee': trimStr(getVal(row, 'Sales employee name')),
    'Quantity (PCS)': Number(getVal(row, 'Order qty.(A)') || 0),
    'Amount (KRW)': Number(getVal(row, 'Net value(KRW)') || 0),
    'Net price': getNetPriceValue(getVal(row, 'Net price')),
    Category: getCategoryValue(getVal(row, 'Division name'))
  }));
}

async function processShipFile(file: File): Promise<MergedRow[]> {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<any>(sheet);

  return rows.map(row => ({
    Code: 'Ship',
    'Registration month': '',
    'Plan month': trimStr(getVal(row, 'Month')),
    Division: getDivisionValue(getVal(row, 'Sales office')),
    Region: getRegionValue(getVal(row, 'Sales group')),
    Material: trimStr(getVal(row, 'Material name')),
    Customer: trimStr(getVal(row, 'Customer name')),
    'End customer': trimStr(getVal(row, 'End customer name')),
    'Sales employee': trimStr(getVal(row, 'Sales employee name')),
    'Quantity (PCS)': Number(getVal(row, 'Delivery qty.') || 0),
    'Amount (KRW)': Number(getVal(row, 'Value of supply(KRW)') || 0),
    'Net price': getNetPriceValue(getVal(row, 'Net price of S/O')),
    Category: getCategoryValue(getVal(row, 'Division name'))
  }));
}

async function processFCSTFile(file: File): Promise<MergedRow[]> {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  
  // Use header: 1 to get raw arrays and handle duplicate "Name" columns
  const jsonData = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1 });
  if (jsonData.length === 0) return [];

  const rawHeaders = jsonData[0] as any[];
  const dataRows = jsonData.slice(1);

  // Rename "Name" columns: first -> Name1, second -> Name2
  let nameCount = 0;
  const processedHeaders = rawHeaders.map(h => {
    const header = trimStr(h);
    if (header.toLowerCase() === 'name') {
      nameCount++;
      return `Name${nameCount}`;
    }
    return header;
  });

  // Map data rows to objects using processed headers
  const rows = dataRows.map(row => {
    const obj: any = {};
    processedHeaders.forEach((h, i) => {
      if (h) obj[h] = row[i];
    });
    return obj;
  });

  return rows.map(row => {
    const regWeek = trimStr(getVal(row, 'Registration week'));
    const regMonth = getRegistrationMonthValue(regWeek);
    const planMonth = formatFCSTPlanMonth(getVal(row, 'Month'), getVal(row, 'Plan month'));
    
    return {
      Code: getMCode(regMonth, planMonth),
      'Registration month': regMonth,
      'Plan month': planMonth,
      Division: getDivisionValue(getVal(row, 'Sales Office')),
      Region: getRegionValue(getVal(row, 'Sales Group')),
      Material: trimStr(getVal(row, 'Material Description')),
      Customer: trimStr(getVal(row, 'Name1')),
      'End customer': trimStr(getVal(row, 'District Name')),
      'Sales employee': trimStr(getVal(row, 'Employee/app.name')),
      'Quantity (PCS)': Number(getVal(row, 'FCST balance') || 0),
      'Amount (KRW)': Number(getVal(row, 'Net Value') || 0),
      'Net price': getNetPriceValue(getVal(row, 'Net Price')),
      Category: getCategoryValue(getVal(row, 'Name2'))
    };
  });
}
