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
  extractionSettings: ExtractionSettings,
  profile?: 'THPT' | 'TieuHoc' | 'All'
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
      // Auto compute day for active profiles based on the exact row offsets requested
      if (profile === 'THPT' || profile === 'TieuHoc') {
        const rowNum = rowIndex + 6; // index 0 matches Excel Row 6 (since rawContentRows starts from row 6)
        if (rowNum >= 10 && rowNum <= 24) {
          day = 'Thứ 2';
        } else if (rowNum >= 28 && rowNum <= 43) {
          day = 'Thứ 3';
        } else if (rowNum >= 47 && rowNum <= 62) {
          day = 'Thứ 4';
        } else if (rowNum >= 66 && rowNum <= 81) {
          day = 'Thứ 5';
        } else if (rowNum >= 85 && rowNum <= 100) {
          day = 'Thứ 6';
        } else {
          day = 'Nghỉ/Trống';
        }
      } else {
        day = lastSeenDay;
      }
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

  // Create empty grid representing the spreadsheet rows (at least 105 rows to comfortably fit up to row 100)
  const data: any[][] = [];
  for (let i = 0; i < 105; i++) {
    data.push(Array(23).fill(''));
  }

  // Row index 4 corresponds to Row 5 in Excel (1-indexed), which is the HEADER row containing class names
  data[4] = [
    '', '', '', '', // A, B, C, D empty
    'Thời gian Tiểu học', 'Lớp 1A', 'Lớp 2A', 'Lớp 3A', 'Lớp 4A', 'Lớp 5A', // E (4) to J (9) - Tiểu học
    'Thời gian THPT', 'Lớp 10A1', 'Lớp 10A2', 'Lớp 11A1', 'Lớp 11A2', 'Lớp 12A1', 'Lớp 12A2', 'Lớp 10B1', 'Lớp 10B2', 'Lớp 11B1', 'Lớp 11B2', 'Lớp 12B1' // K (10) to V (21) - THPT
  ];

  // Helper to populate day blocks matching the exact rows:
  // - Thứ 2: Dòng 9 đến 24 (index 8 đến 23)
  // - Thứ 3: Dòng 28 đến 43 (index 27 đến 42)
  // - Thứ 4: Dòng 47 đến 62 (index 46 đến 61)
  // - Thứ 5: Dòng 66 đến 81 (index 65 đến 80)
  // - Thứ 6: Dòng 85 đến 100 (index 84 đến 99)
  const fillDayBlock = (dayName: string, startRowIdx: number, endRowIdx: number) => {
    for (let rowIdx = startRowIdx; rowIdx <= endRowIdx; rowIdx++) {
      const slotNum = rowIdx - startRowIdx + 1;
      
      // Introduce dynamic teacher assignments with realistic conflicts to show system's accuracy
      let cls10A1 = `Toán - Thầy Hải`;
      let cls10A2 = slotNum === 1 ? `Văn - Thầy Hải` : `Văn - Cô Vy`; // Conflict on Monday/Tuesday etc. Tiết 1
      
      let cls11A1 = `Lý - Cô Hương`;
      let cls11A2 = slotNum === 2 ? `Sử - Cô Hương` : `Hóa - Thầy Nam`; // Conflict on Tiết 2
      
      let cls1A = `Toán - Cô Vy`;
      let cls2A = slotNum === 3 ? `Văn - Cô Vy` : `Anh - Cô Vy`; // Conflict on Tiết 3

      data[rowIdx] = [
        '', '', '', '',
        `${dayName} - Tiết ${slotNum}`, cls1A, cls2A, 'Toán - Thầy Hải', 'Địa - Cô Tâm', 'Mỹ thuật', // E to J (Tiểu học)
        `${dayName} - Tiết ${slotNum}`, cls10A1, cls10A2, cls11A1, cls11A2, 'Tin học - Cô Nhi', 'Sinh - Thầy Bình', 'GDCD', 'Địa', 'Sử', 'QPAN', 'Sinh hoạt' // K to V (THPT)
      ];
    }
  };

  // Populate Monday (Thứ 2) -> Rows 10-24 (Indices 9-23)
  fillDayBlock('Thứ 2', 9, 23);

  // Populate Tuesday (Thứ 3) -> Rows 28-43 (Indices 27-42)
  fillDayBlock('Thứ 3', 27, 42);

  // Populate Wednesday (Thứ 4) -> Rows 47-62 (Indices 46-61)
  fillDayBlock('Thứ 4', 46, 61);

  // Populate Thursday (Thứ 5) -> Rows 66-81 (Indices 65-80)
  fillDayBlock('Thứ 5', 65, 80);

  // Populate Friday (Thứ 6) -> Rows 85-100 (Indices 84-99)
  fillDayBlock('Thứ 6', 84, 99);

  // Set visual separating tags for the spacing rows (optional but makes it extremely easy to read)
  data[24][4] = '--- NGHỈ TRƯA ---'; data[24][10] = '--- NGHỈ TRƯA ---';
  data[43][4] = '--- NGHỈ TRƯA ---'; data[43][10] = '--- NGHỈ TRƯA ---';
  data[62][4] = '--- NGHỈ TRƯA ---'; data[62][10] = '--- NGHỈ TRƯA ---';
  data[81][4] = '--- NGHỈ TRƯA ---'; data[81][10] = '--- NGHỈ TRƯA ---';

  const ws = XLSX.utils.aoa_to_sheet(data);

  // Set column widths for high legibility
  const wscols = [
    { wch: 5 },  // A
    { wch: 5 },  // B
    { wch: 5 },  // C
    { wch: 5 },  // D
    { wch: 20 }, // E: Thời gian Tiểu học
    { wch: 18 }, // F: Lớp 1A
    { wch: 18 }, // G: Lớp 2A
    { wch: 18 }, // H: Lớp 3A
    { wch: 18 }, // I: Lớp 4A
    { wch: 18 }, // J: Lớp 5A
    { wch: 20 }, // K: Thời gian THPT
    { wch: 18 }, // L: Lớp 10A1
    { wch: 18 }, // M: Lớp 10A2
    { wch: 18 }, // N: Lớp 11A1
    { wch: 18 }, // O: Lớp 11A2
    { wch: 18 }, // P: Lớp 12A1
    { wch: 18 }, // Q: Lớp 12A2
    { wch: 18 }, // R: Lớp 10B1
    { wch: 18 }, // S: Lớp 10B2
    { wch: 18 }, // T: Lớp 11B1
    { wch: 18 }, // U: Lớp 11B2
    { wch: 18 }, // V: Lớp 12B1
  ];
  ws['!cols'] = wscols;

  XLSX.utils.book_append_sheet(wb, ws, 'Thời khóa biểu mẫu');

  // Trigger download
  XLSX.writeFile(wb, 'thoi_khoa_bieu_mau_dong5.xlsx');
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
