import * as XLSX from 'xlsx';
import { 
  ExemptionPair, 
  ColumnMapping, 
  TeacherConflict, 
  TeacherInfo, 
  ExtractionSettings,
  CellRange,
  DaySession,
  SplitPeriodSlot,
  SplitPeriodIssue
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

    // Skip processing dividers, blank rows, or rows outside the school timetable block
    if ((profile === 'THPT' || profile === 'TieuHoc') && day === 'Nghỉ/Trống') {
      continue;
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
      
      // Realistic timetable slots:
      // - 10A1: Toán Thầy Hải (tiết 1, 2) -> Consecutive morning (KHÔNG bị chia)
      // - 10A2: Văn Cô Vy (tiết 1 sáng & tiết 6 chiều) -> Bị chia 2 buổi (Sáng & Chiều)
      // - 11A1: Hóa Thầy Nam (tiết 1 & tiết 3 sáng) -> Loãng xương buổi sáng
      // - 11A2: Tiết 2 trùng giáo viên (xung đột giờ)
      let cls10A1 = 'Sinh hoạt';
      if (slotNum === 1 || slotNum === 2) cls10A1 = 'Toán - Thầy Hải';
      else if (slotNum === 3 || slotNum === 4) cls10A1 = 'Lý - Cô Hương';
      else if (slotNum === 5) cls10A1 = 'Sinh - Thầy Bình';
      else if (slotNum === 6 || slotNum === 7) cls10A1 = 'Anh - Cô Mai';

      let cls10A2 = 'Tự chọn';
      if (slotNum === 1) cls10A2 = 'Văn - Cô Vy'; // Sáng tiết 1
      else if (slotNum === 2) cls10A2 = 'Sử - Cô Hương';
      else if (slotNum === 3 || slotNum === 4) cls10A2 = 'Toán - Thầy Nam';
      else if (slotNum === 6) cls10A2 = 'Văn - Cô Vy'; // Chiều tiết 6 -> Bị chia sáng & chiều!

      let cls11A1 = 'Tin học';
      if (slotNum === 1) cls11A1 = 'Hóa - Thầy Nam'; // Tiết 1
      else if (slotNum === 2) cls11A1 = 'Toán - Thầy Hải'; // Tiết 2
      else if (slotNum === 3) cls11A1 = 'Hóa - Thầy Nam'; // Tiết 3 -> Loãng xương sáng!
      else if (slotNum === 4 || slotNum === 5) cls11A1 = 'Văn - Cô Vy';

      let cls11A2 = 'GDCD';
      if (slotNum === 1) cls11A2 = 'Địa - Cô Tâm';
      else if (slotNum === 2) cls11A2 = 'Sử - Cô Hương'; // Xung đột với 10A2 nếu Cô Hương dạy cả hai
      else if (slotNum === 6 || slotNum === 7) cls11A2 = 'GDCD - Thầy Bình';

      let cls1A = slotNum <= 2 ? 'Toán - Cô Vy' : (slotNum <= 4 ? 'Tiếng Việt' : 'Mỹ thuật');
      let cls2A = slotNum === 1 ? 'Toán - Cô Vy' : (slotNum <= 3 ? 'Anh văn' : 'Âm nhạc');

      const periodLabel = slotNum <= 5 ? `Tiết ${slotNum} (Sáng)` : `Tiết ${slotNum} (Chiều)`;

      data[rowIdx] = [
        '', '', '', '',
        `${dayName} - ${periodLabel}`, cls1A, cls2A, 'Toán - Thầy Hải', 'Địa - Cô Tâm', 'Mỹ thuật', // E to J (Tiểu học)
        `${dayName} - ${periodLabel}`, cls10A1, cls10A2, cls11A1, cls11A2, 'Tin học - Cô Nhi', 'Sinh - Thầy Bình', 'GDCD', 'Địa', 'Sử', 'QPAN', 'Sinh hoạt' // K to V (THPT)
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

/**
 * Parses the period text, dayRowIndex, or timetable row to determine:
 * - periodNum: 1..5 for Morning ('Sáng'), 6..8+ for Afternoon ('Chiều')
 * - session: 'Sáng' | 'Chiều'
 * - displayName: standardized display label (e.g. "Tiết 1 (Sáng)", "Tiết 6 (Chiều)")
 */
export function parsePeriodInfo(
  periodText: string,
  dayRowIndex?: number,
  rowIndex?: number,
  profile?: 'THPT' | 'TieuHoc' | 'All'
): { periodNum: number; session: DaySession; displayName: string } {
  const p = (periodText || '').toLowerCase().trim();

  // 1. Text mentioning relative afternoon period, e.g. "Tiết 1 chiều", "T1 chiều", "1 chiều"
  const relChieuMatch = p.match(/(?:tiết|t)?\s*([1-5])\s*(?:chiều|chieu|pm)/i);
  if (relChieuMatch) {
    const relNum = parseInt(relChieuMatch[1], 10);
    const periodNum = relNum + 5; // 1 chiều -> 6, 2 chiều -> 7, 3 chiều -> 8
    return {
      periodNum,
      session: 'Chiều',
      displayName: `Tiết ${periodNum} (Chiều)`
    };
  }

  // 2. Explicit "Tiết X" or "TX" where X is a number
  const tietMatch = p.match(/(?:tiết|t)\s*(\d+)/i);
  if (tietMatch) {
    const num = parseInt(tietMatch[1], 10);
    if (num <= 5) {
      return { periodNum: num, session: 'Sáng', displayName: `Tiết ${num} (Sáng)` };
    } else {
      return { periodNum: num, session: 'Chiều', displayName: `Tiết ${num} (Chiều)` };
    }
  }

  // 3. Standalone digits (e.g. "1", "2", "3", "4", "5", "6", "7", "8")
  const digitMatch = p.match(/^(\d+)$/);
  if (digitMatch) {
    const num = parseInt(digitMatch[1], 10);
    if (num <= 5) {
      return { periodNum: num, session: 'Sáng', displayName: `Tiết ${num} (Sáng)` };
    } else {
      return { periodNum: num, session: 'Chiều', displayName: `Tiết ${num} (Chiều)` };
    }
  }

  // 4. Time matching:
  // 7h.. -> Tiết 1; 8h.. -> Tiết 2; 9h.. -> Tiết 3; 10h.. -> Tiết 4; 11h.. -> Tiết 5
  // 13h.. -> Tiết 6; 14h.. -> Tiết 7; 15h.. -> Tiết 8; 16h.. -> Tiết 9
  const hourMatch = p.match(/(\d{1,2})[h:]/);
  if (hourMatch) {
    const hour = parseInt(hourMatch[1], 10);
    if (hour >= 6 && hour < 12) {
      let pNum = 1;
      if (hour <= 7) pNum = 1;
      else if (hour === 8) pNum = 2;
      else if (hour === 9) pNum = 3;
      else if (hour === 10) pNum = 4;
      else pNum = 5;
      return { periodNum: pNum, session: 'Sáng', displayName: `Tiết ${pNum} (Sáng)` };
    } else if (hour >= 12 && hour <= 18) {
      let pNum = 6;
      if (hour <= 13) pNum = 6;
      else if (hour === 14) pNum = 7;
      else if (hour === 15) pNum = 8;
      else pNum = 9;
      return { periodNum: pNum, session: 'Chiều', displayName: `Tiết ${pNum} (Chiều)` };
    }
  }

  // 5. If profile THPT or TieuHoc, use row offset relative to day block:
  // In each day block (e.g. Thứ 2 rows 10-24):
  // First 5 periods (offset 0..4) are Tiết 1..5 (Sáng)
  // Next 3+ periods (offset 5..7) are Tiết 6..8 (Chiều)
  if (rowIndex !== undefined && (profile === 'THPT' || profile === 'TieuHoc')) {
    const rowNum = rowIndex + 6;
    let dayBase = 0;
    if (rowNum >= 10 && rowNum <= 24) dayBase = 10;
    else if (rowNum >= 28 && rowNum <= 43) dayBase = 28;
    else if (rowNum >= 47 && rowNum <= 62) dayBase = 47;
    else if (rowNum >= 66 && rowNum <= 81) dayBase = 66;
    else if (rowNum >= 85 && rowNum <= 100) dayBase = 85;

    if (dayBase > 0) {
      const offset = rowNum - dayBase;
      const pNum = offset + 1;
      if (pNum <= 5) {
        return { periodNum: pNum, session: 'Sáng', displayName: `Tiết ${pNum} (Sáng)` };
      } else {
        return { periodNum: pNum, session: 'Chiều', displayName: `Tiết ${pNum} (Chiều)` };
      }
    }
  }

  // 6. DayRowIndex fallback
  if (dayRowIndex !== undefined && dayRowIndex >= 0) {
    const pNum = dayRowIndex + 1;
    if (pNum <= 5) {
      return { periodNum: pNum, session: 'Sáng', displayName: `Tiết ${pNum} (Sáng)` };
    } else {
      return { periodNum: pNum, session: 'Chiều', displayName: `Tiết ${pNum} (Chiều)` };
    }
  }

  return { periodNum: 1, session: 'Sáng', displayName: periodText || 'Tiết 1' };
}

/**
 * Backward compatibility helper for session determination
 */
export function determineSession(
  period: string,
  rowIndex: number,
  profile?: 'THPT' | 'TieuHoc' | 'All'
): DaySession {
  return parsePeriodInfo(period, undefined, rowIndex, profile).session;
}

/**
 * Checks for "Tiết bị chia / Loãng xương" in a single day for each class:
 * - Morning: Periods 1, 2, 3, 4, 5. Consecutive teaching is NOT split.
 * - Afternoon: Periods 6, 7, 8. Consecutive teaching is NOT split.
 * - ONLY flagged as split / loãng xương if:
 *   + A teacher teaches a period, then has NO period for that class, and then later teaches that class again!
 *     (e.g., Morning period 1, then no period, then morning period 3 -> Loãng xương buổi sáng).
 *     (e.g., Afternoon period 6, then no period, then afternoon period 8 -> Loãng xương buổi chiều).
 *     (e.g., Morning period 1, then no periods, then afternoon period 6 -> Bị chia 2 buổi Sáng & Chiều).
 */
export function detectSplitPeriods(
  rows: string[][],
  mappings: ColumnMapping[],
  extractionSettings: ExtractionSettings,
  profile?: 'THPT' | 'TieuHoc' | 'All'
): SplitPeriodIssue[] {
  const dayCol = mappings.find(m => m.role === 'day');
  const periodCol = mappings.find(m => m.role === 'period');
  const classCols = mappings.filter(m => m.role === 'class');

  if (classCols.length === 0) return [];

  let lastSeenDay = 'Chưa xác định';
  const dayRowCounters = new Map<string, number>();

  // Map: Key (day__class__teacher[__subject]) -> { day, className, teacher, subject, slots }
  const teacherClassDayMap = new Map<string, {
    day: string;
    className: string;
    teacher: string;
    subject?: string;
    slots: SplitPeriodSlot[];
  }>();

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    if (!row || row.length === 0) continue;

    // Determine Day
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
      if (profile === 'THPT' || profile === 'TieuHoc') {
        const rowNum = rowIndex + 6;
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

    // Skip empty or rest rows
    if ((profile === 'THPT' || profile === 'TieuHoc') && day === 'Nghỉ/Trống') {
      continue;
    }
    if (day === 'Chưa xác định') continue;

    const currentDayRow = dayRowCounters.get(day) || 0;
    dayRowCounters.set(day, currentDayRow + 1);

    // Determine Period text
    let periodText = '';
    if (periodCol !== undefined) {
      const cellVal = row[periodCol.index];
      if (cellVal && String(cellVal).trim()) {
        periodText = String(cellVal).trim();
      }
    }

    const { periodNum, session, displayName } = parsePeriodInfo(
      periodText,
      currentDayRow,
      rowIndex,
      profile
    );

    // Scan each class column
    classCols.forEach(col => {
      const cellValue = row[col.index];
      const extracted = extractTeacher(cellValue, extractionSettings);
      if (!extracted || !extracted.teacherName) return;

      const normTeacher = normalizeTeacherName(extracted.teacherName);
      const normClass = normalizeClassName(col.header);
      const normSubject = extracted.subjectName ? normalizeTeacherName(extracted.subjectName) : '';

      const key = `${day}__${normClass}__${normTeacher}${normSubject ? `__${normSubject}` : ''}`;

      if (!teacherClassDayMap.has(key)) {
        teacherClassDayMap.set(key, {
          day,
          className: col.header,
          teacher: extracted.teacherName,
          subject: extracted.subjectName,
          slots: []
        });
      }

      teacherClassDayMap.get(key)!.slots.push({
        period: periodText || displayName,
        session,
        periodNum,
        rowIndex,
        originalValue: extracted.originalValue
      });
    });
  }

  const issues: SplitPeriodIssue[] = [];

  teacherClassDayMap.forEach((entry) => {
    const { day, className, teacher, subject, slots } = entry;
    if (slots.length <= 1) return; // Single period taught, cannot be split

    // Sort slots by periodNum ascending
    slots.sort((a, b) => a.periodNum - b.periodNum);

    // Deduplicate any slots having the exact same periodNum
    const uniqueSlots: SplitPeriodSlot[] = [];
    slots.forEach(slot => {
      if (uniqueSlots.length === 0 || uniqueSlots[uniqueSlots.length - 1].periodNum !== slot.periodNum) {
        uniqueSlots.push(slot);
      }
    });

    if (uniqueSlots.length <= 1) return;

    // Morning periods (1 to 5) and Afternoon periods (6 to 8+)
    const morningSlots = uniqueSlots.filter(s => s.periodNum <= 5);
    const afternoonSlots = uniqueSlots.filter(s => s.periodNum >= 6);

    // 1. Check for gap in Morning: teaching a period, then no period, then teaching again
    const morningGaps: string[] = [];
    for (let i = 0; i < morningSlots.length - 1; i++) {
      const current = morningSlots[i].periodNum;
      const next = morningSlots[i + 1].periodNum;
      if (next - current > 1) {
        const missing: number[] = [];
        for (let m = current + 1; m < next; m++) {
          missing.push(m);
        }
        morningGaps.push(`Dạy Tiết ${current}, trống Tiết ${missing.join(', ')}, đến Tiết ${next} mới dạy`);
      }
    }
    const hasMorningGap = morningGaps.length > 0;

    // 2. Check for gap in Afternoon: teaching a period, then no period, then teaching again
    const afternoonGaps: string[] = [];
    for (let i = 0; i < afternoonSlots.length - 1; i++) {
      const current = afternoonSlots[i].periodNum;
      const next = afternoonSlots[i + 1].periodNum;
      if (next - current > 1) {
        const missing: number[] = [];
        for (let m = current + 1; m < next; m++) {
          missing.push(m);
        }
        afternoonGaps.push(`Dạy Tiết ${current}, trống Tiết ${missing.join(', ')}, đến Tiết ${next} mới dạy`);
      }
    }
    const hasAfternoonGap = afternoonGaps.length > 0;

    // 3. Check for morning & afternoon split
    const isMorningAfternoonSplit = morningSlots.length > 0 && afternoonSlots.length > 0;

    // STRICT USER RULE:
    // "Nếu các tiết 1,2,3,4,5 là buổi sáng thì không xem là bị chia. Chiều là tiết 6,7,8 là buổi chiều và không xem là bị chia.
    // trừ trường hợp dạy 1 tiết rồi k có tiết lớp đó sau đó lại có tiết lớp đó mới gọi bị chia hay loãng xương"
    //
    // => If all periods are in morning and consecutive without gaps: NOT split!
    // => If all periods are in afternoon and consecutive without gaps: NOT split!
    // => Only if there is a gap (morning gap, afternoon gap, or split between morning and afternoon): FLAGGED!
    if (!isMorningAfternoonSplit && !hasMorningGap && !hasAfternoonGap) {
      return;
    }

    let splitType: 'morning_afternoon' | 'isolated_periods' | 'both';
    if (isMorningAfternoonSplit && (hasMorningGap || hasAfternoonGap)) {
      splitType = 'both';
    } else if (isMorningAfternoonSplit) {
      splitType = 'morning_afternoon';
    } else {
      splitType = 'isolated_periods';
    }

    const morningPeriods = morningSlots.map(s => s.period);
    const afternoonPeriods = afternoonSlots.map(s => s.period);

    // Build clear descriptive explanation
    const descParts: string[] = [];
    if (isMorningAfternoonSplit) {
      descParts.push(`Bị chia 2 buổi (Sáng ${morningSlots.length} tiết: ${morningPeriods.join(', ')} & Chiều ${afternoonSlots.length} tiết: ${afternoonPeriods.join(', ')})`);
    }
    if (hasMorningGap) {
      descParts.push(`Loãng xương sáng (${morningGaps.join('; ')})`);
    }
    if (hasAfternoonGap) {
      descParts.push(`Loãng xương chiều (${afternoonGaps.join('; ')})`);
    }

    const description = descParts.join(' • ');

    issues.push({
      id: `${day}-${className}-${teacher}-${splitType}`,
      day,
      className,
      teacher,
      subject,
      splitType,
      morningPeriods,
      afternoonPeriods,
      totalPeriods: uniqueSlots.length,
      description,
      slots: uniqueSlots
    });
  });

  return issues;
}
