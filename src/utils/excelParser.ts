import * as XLSX from 'xlsx';
import { 
  ExemptionPair, 
  ColumnMapping, 
  TeacherConflict, 
  TeacherInfo, 
  ExtractionSettings,
  CellRange
} from '../types';

/**
 * Normalizes a class name for comparison (removes dots, spaces, converts to lowercase)
 */
export function normalizeClassName(className: string): string {
  if (!className) return '';
  return className.trim().replace(/[\s\.]+/g, '').toLowerCase();
}

/**
 * Checks if two class names match after normalization
 */
export function isClassMatch(classA: string, classB: string): boolean {
  return normalizeClassName(classA) === normalizeClassName(classB);
}

/**
 * Normalizes a teacher name for comparison (trims, removes excess spaces, lowercase)
 */
export function normalizeTeacherName(name: string): string {
  if (!name) return '';
  return name.trim().replace(/\s+/g, ' ').toLowerCase();
}

/**
 * Extracts a teacher's name and subject from a cell's string value based on extraction settings
 */
export function extractTeacher(cellValue: any, settings: ExtractionSettings): TeacherInfo | null {
  if (cellValue === undefined || cellValue === null) return null;
  
  const valStr = String(cellValue).trim();
  if (!valStr || valStr === '-' || valStr === 'x' || valStr === 'X') return null;

  // Check if cell is in ignore list
  if (settings.ignoreList.some(ignore => valStr.toLowerCase() === ignore.toLowerCase())) {
    return null;
  }

  const delimiter = settings.delimiter;
  if (!delimiter || !valStr.includes(delimiter)) {
    // No delimiter found, treat whole cell as teacher name
    return {
      originalValue: valStr,
      teacherName: valStr
    };
  }

  const parts = valStr.split(delimiter).map(p => p.trim());
  if (parts.length < 2) {
    return {
      originalValue: valStr,
      teacherName: valStr
    };
  }

  if (settings.takeFirstPart) {
    // Format: "Teacher - Subject"
    const teacherName = parts[0];
    const subjectName = parts.slice(1).join(delimiter);
    return {
      originalValue: valStr,
      teacherName,
      subjectName
    };
  } else {
    // Format: "Subject - Teacher"
    const subjectName = parts[0];
    const teacherName = parts.slice(1).join(delimiter);
    return {
      originalValue: valStr,
      teacherName,
      subjectName
    };
  }
}

/**
 * Checks if a pair of classes is exempted
 */
export function isPairExempted(class1: string, class2: string, exemptions: ExemptionPair[]): boolean {
  return false; // Completely disabled - check all classes for conflicts as requested
}

/**
 * Analyzes the timetable grid and finds all conflicts
 */
export function detectConflicts(
  rows: string[][],
  mappings: ColumnMapping[],
  exemptions: ExemptionPair[],
  extractionSettings: ExtractionSettings
): TeacherConflict[] {
  const conflicts: TeacherConflict[] = [];

  // Identify column indices for each role
  const dayCol = mappings.find(m => m.role === 'day');
  const periodCol = mappings.find(m => m.role === 'period');
  const classCols = mappings.filter(m => m.role === 'class');

  if (classCols.length === 0) return [];

  // Track the last seen day to handle merged cells (if "Thứ" is only filled on the first row of a block)
  let lastSeenDay = 'Chưa xác định';

  // Process rows starting from the content rows (usually index 1 is header, index 2+ are content)
  // Our row data contains rows parsed from the spreadsheet.
  // We'll skip the first row if it's the header.
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    if (!row || row.length === 0) continue;

    // Determine Day and Period
    let day = 'Chưa xác định';
    if (dayCol !== undefined) {
      const cellVal = row[dayCol.index];
      if (cellVal && String(cellVal).trim()) {
        day = String(cellVal).trim();
        lastSeenDay = day;
      } else {
        day = lastSeenDay;
      }
    } else {
      day = lastSeenDay;
    }

    let period = 'Chưa xác định';
    if (periodCol !== undefined) {
      const cellVal = row[periodCol.index];
      if (cellVal && String(cellVal).trim()) {
        period = String(cellVal).trim();
      }
    }

    // Map: Teacher Name (Normalized) -> Array of { classHeader: string, originalCell: string }
    const teacherAssignments = new Map<string, Array<{ className: string; originalVal: string }>>();

    classCols.forEach(col => {
      const cellValue = row[col.index];
      const extracted = extractTeacher(cellValue, extractionSettings);
      
      if (extracted && extracted.teacherName) {
        const normName = normalizeTeacherName(extracted.teacherName);
        if (!teacherAssignments.has(normName)) {
          teacherAssignments.set(normName, []);
        }
        teacherAssignments.get(normName)!.push({
          className: col.header,
          originalVal: extracted.originalValue
        });
      }
    });

    // Analyze teacher assignments for conflicts
    teacherAssignments.forEach((assignments, normTeacherName) => {
      if (assignments.length < 2) return; // No conflict possible with 1 class

      const displayTeacherName = assignments[0].originalVal.includes(extractionSettings.delimiter)
        ? extractTeacher(assignments[0].originalVal, extractionSettings)?.teacherName || normTeacherName
        : assignments[0].originalVal;

      // Find conflicting pairs of classes
      const conflictPairs: [string, string][] = [];
      const classesList = assignments.map(a => a.className);

      for (let i = 0; i < assignments.length; i++) {
        for (let j = i + 1; j < assignments.length; j++) {
          const class1 = assignments[i].className;
          const class2 = assignments[j].className;

          if (!isPairExempted(class1, class2, exemptions)) {
            conflictPairs.push([class1, class2]);
          }
        }
      }

      if (conflictPairs.length > 0) {
        conflicts.push({
          id: `${day}-${period}-${normTeacherName}-${rowIndex}`,
          day,
          period,
          teacher: displayTeacherName,
          classes: classesList,
          conflictPairs
        });
      }
    });
  }

  return conflicts;
}

/**
 * Generates a mock timetable Excel file for the user to download
 */
export function generateSampleExcel(): void {
  // Create worksheets
  const wb = XLSX.utils.book_new();

  // Create empty grid representing the spreadsheet rows
  const data: any[][] = [];
  for (let i = 0; i < 12; i++) {
    data.push([]);
  }

  // Row index 0 corresponds to Row 1 in Excel (where headers/class names are placed)
  data[0] = [
    '', '', '', // A, B, C empty
    'Thứ', 'Tiết', // D, E
    '10.1 TN', '10.1 XH', '11.1 TN', '11.2 XH', '12.1 TN', '12.2 XH', '9.1', '9.2' // F to M
  ];

  // Row index 4 corresponds to Row 5 in Excel (1-indexed), which is the start of the timetable block
  data[4] = [
    '', '', '',
    'Thứ Hai', 'Tiết 1', 
    'Toán - Thầy Hải', 'Toán - Thầy Hải', 'Lý - Cô Bình', 'Lý - Cô Vy', 'Văn - Thầy Nam', 'Hóa - Cô Lan', 'Sinh - Cô Tâm', 'Anh - Thầy Minh'
  ];
  
  data[5] = [
    '', '', '',
    'Thứ Hai', 'Tiết 2', 
    'Anh - Cô Mai', 'Anh - Cô Mai', 'Sử - Cô Hương', 'Sử - Cô Hương', 'Địa - Thầy Tuấn', 'Công nghệ - Thầy Bắc', 'Văn - Cô Lan', 'Văn - Cô Lan'
  ];

  data[6] = [
    '', '', '',
    'Thứ Hai', 'Tiết 3', 
    'Văn - Cô Lan', 'Địa - Thầy Tuấn', 'Toán - Thầy Hải', 'Toán - Thầy Hải', 'Sinh - Cô Tâm', 'Sử - Cô Hương', 'Mỹ thuật - Cô Nhã', 'Nhạc - Thầy Sơn'
  ];

  data[7] = [
    '', '', '',
    'Thứ Ba', 'Tiết 1', 
    'Sử - Cô Hương', 'Sinh - Cô Vy', 'Hóa - Thầy Sơn', 'Tin - Cô Nhi', 'Sử - Cô Hương', 'Toán - Thầy Hải', 'Địa - Thầy Tuấn', 'Sinh - Cô Vy'
  ];

  data[8] = [
    '', '', '',
    'Thứ Ba', 'Tiết 2', 
    'Hóa - Thầy Sơn', 'Anh - Cô Mai', 'Tin - Cô Nhi', 'Tin - Cô Nhi', 'Địa - Thầy Tuấn', 'Văn - Thầy Nam', 'Nhạc - Thầy Sơn', 'Nhạc - Thầy Sơn'
  ];

  data[9] = [
    '', '', '',
    'Thứ Ba', 'Tiết 3', 
    'Lý - Cô Vy', 'Sử - Cô Hương', 'Anh - Cô Mai', 'Lý - Cô Vy', 'Hóa - Thầy Sơn', 'Hóa - Thầy Sơn', 'Tin - Thầy Hùng', 'Tin - Thầy Hùng'
  ];

  const ws = XLSX.utils.aoa_to_sheet(data);

  // Set column widths
  const wscols = [
    { wch: 5 },  // A
    { wch: 5 },  // B
    { wch: 5 },  // C
    { wch: 12 }, // D
    { wch: 10 }, // E
    { wch: 18 }, // F: 10.1 TN
    { wch: 18 }, // G: 10.1 XH
    { wch: 18 }, // H: 11.1 TN
    { wch: 18 }, // I: 11.2 XH
    { wch: 18 }, // J: 12.1 TN
    { wch: 18 }, // K: 12.2 XH
    { wch: 18 }, // L: 9.1
    { wch: 18 }, // M: 9.2
  ];
  ws['!cols'] = wscols;

  XLSX.utils.book_append_sheet(wb, ws, 'Thời khóa biểu mẫu');

  // Trigger download
  XLSX.writeFile(wb, 'thoi_khoa_bieu_mau_trung_lich.xlsx');
}

/**
 * Converts Excel column name (e.g. "F", "V", "AA") to 0-based column index
 */
export function colNameToIndex(colName: string): number {
  let index = 0;
  const upperCol = colName.toUpperCase().trim();
  for (let i = 0; i < upperCol.length; i++) {
    index = index * 26 + (upperCol.charCodeAt(i) - 64);
  }
  return index - 1; // 0-indexed
}

/**
 * Converts 0-based column index to Excel column name (e.g. 0 -> "A", 5 -> "F")
 */
export function indexToColName(index: number): string {
  let colName = '';
  let temp = index;
  while (temp >= 0) {
    colName = String.fromCharCode((temp % 26) + 65) + colName;
    temp = Math.floor(temp / 26) - 1;
  }
  return colName;
}

/**
 * Parses an Excel range string like "F39:V100" into a CellRange object
 */
export function parseExcelRange(rangeStr: string): CellRange | null {
  if (!rangeStr) return null;
  const regex = /^([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$/i;
  const match = rangeStr.trim().match(regex);
  if (!match) return null;

  const startColName = match[1];
  const startRowStr = match[2];
  const endColName = match[3];
  const endRowStr = match[4];

  const startRow = parseInt(startRowStr, 10) - 1;
  const endRow = parseInt(endRowStr, 10) - 1;
  const startCol = colNameToIndex(startColName);
  const endCol = colNameToIndex(endColName);

  // Validate the coordinates
  if (isNaN(startRow) || isNaN(endRow) || startRow < 0 || endRow < 0 || startCol < 0 || endCol < 0) {
    return null;
  }

  return {
    startRow: Math.min(startRow, endRow),
    endRow: Math.max(startRow, endRow),
    startCol: Math.min(startCol, endCol),
    endCol: Math.max(startCol, endCol)
  };
}
