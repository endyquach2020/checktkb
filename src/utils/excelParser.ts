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
  SplitPeriodIssue,
  SimultaneousMatch,
  SimultaneousCheckConfig
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
    'Thời gian THPT', 'Lớp 10A1', 'Lớp 10.2', 'Lớp 11.1', 'Lớp 11.2', 'Lớp 12.1', 'Lớp 12.2', 'Lớp 10B1', 'Lớp 10B2', 'Lớp 11B1', 'Lớp 11B2', 'Lớp 12B1' // K (10) to V (21) - THPT
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
      // - 10.2: Thầy Mịnh & Thầy Đ. Minh trùng giờ với Lớp 12.2, kèm Văn Cô Vy (tiết 1 sáng & tiết 6 chiều)
      // - 11.1: Hóa Thầy Nam (tiết 1 & tiết 3 sáng) -> Loãng xương buổi sáng
      // - 10.2 & 12.2: Kiểm tra trùng giờ Thầy Mịnh & Thầy Đ. Minh
      let cls10A1 = 'Sinh hoạt';
      if (slotNum === 1 || slotNum === 2) cls10A1 = 'Toán - Thầy Hải';
      else if (slotNum === 3 || slotNum === 4) cls10A1 = 'Lý - Cô Hương';
      else if (slotNum === 5) cls10A1 = 'Sinh - Thầy Bình';
      else if (slotNum === 6 || slotNum === 7) cls10A1 = 'Anh - Cô Mai';

      let cls10_2 = 'Tự chọn';
      if (slotNum === 1) cls10_2 = 'Văn - Cô Vy'; // Sáng tiết 1
      else if (slotNum === 2) cls10_2 = 'Sử - Cô Hương';
      else if (slotNum === 3) cls10_2 = 'Toán - Thầy Nam';
      else if (slotNum === 6) cls10_2 = 'Văn - Cô Vy'; // Chiều tiết 6 -> Bị chia sáng & chiều!

      let cls11_1 = 'Tin học';
      if (slotNum === 1) cls11_1 = 'Hóa - Thầy Nam'; // Tiết 1
      else if (slotNum === 2) cls11_1 = 'Toán - Thầy Hải'; // Tiết 2
      else if (slotNum === 3) cls11_1 = 'Hóa - Thầy Nam'; // Tiết 3 -> Loãng xương sáng!
      else if (slotNum === 4 || slotNum === 5) cls11_1 = 'Văn - Cô Vy';

      // 10.2 and 12.2 assignments:
      let cls11_2 = 'GDCD';
      let cls12_1 = 'Tin học - Cô Nhi';
      let cls12_2 = 'Sinh hoạt';

      if (dayName === 'Thứ 2') {
        if (slotNum === 4) {
          cls10_2 = 'Toán - Thầy Mịnh';
          cls12_2 = 'Văn - Thầy Đ. Minh'; // Simultaneous collision for 10.2 & 12.2 on Monday Period 4!
        } else if (slotNum === 1) {
          cls11_2 = 'Địa - Cô Tâm';
          cls12_2 = 'Anh - Cô Mai';
        } else if (slotNum === 2) {
          cls11_2 = 'Sử - Cô Hương';
          cls12_2 = 'Hóa - Thầy Nam';
        }
      } else if (dayName === 'Thứ 4') {
        if (slotNum === 2) {
          cls10_2 = 'Lý - Thầy Đ. Minh';
          cls12_2 = 'Sử - Thầy Mịnh'; // Simultaneous collision for 10.2 & 12.2 on Wednesday Period 2!
        } else if (slotNum === 3) {
          cls11_2 = 'Hóa - Thầy Nam';
          cls12_2 = 'Toán - Thầy Hải';
        }
      } else if (dayName === 'Thứ 3') {
        if (slotNum === 1) {
          cls10_2 = 'Toán - Thầy Mịnh'; // Only 10.2 has Thầy Mịnh, 12.2 has Sinh -> No collision
          cls12_2 = 'Sinh - Thầy Bình';
        }
      }

      let cls1A = slotNum <= 2 ? 'Toán - Cô Vy' : (slotNum <= 4 ? 'Tiếng Việt' : 'Mỹ thuật');
      let cls2A = slotNum === 1 ? 'Toán - Cô Vy' : (slotNum <= 3 ? 'Anh văn' : 'Âm nhạc');

      const periodLabel = slotNum <= 5 ? `Tiết ${slotNum} (Sáng)` : `Tiết ${slotNum} (Chiều)`;

      data[rowIdx] = [
        '', '', '', '',
        `${dayName} - ${periodLabel}`, cls1A, cls2A, 'Toán - Thầy Hải', 'Địa - Cô Tâm', 'Mỹ thuật', // E to J (Tiểu học)
        `${dayName} - ${periodLabel}`, cls10A1, cls10_2, cls11_1, cls11_2, cls12_1, cls12_2, 'Lớp 10B1', 'Lớp 10B2', 'Lớp 11B1', 'Lớp 11B2', 'Lớp 12B1' // K to V (THPT)
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

/**
 * Flexible class name matching:
 * E.g. query "11.2" matches "11.2", "Lớp 11.2", "11.2 TN", "11.2 XH", "11/2", "11A2"
 */
export function matchClassFlexible(query: string, header: string): boolean {
  if (!query || !header) return false;
  const q = query.trim().toLowerCase();
  const h = header.trim().toLowerCase();

  // Direct include
  if (h.includes(q)) return true;

  // Normalized (strip spaces, dots, slashes, dashes)
  const qNorm = q.replace(/[\s\._\-\/]+/g, '');
  const hNorm = h.replace(/[\s\._\-\/]+/g, '');
  if (!qNorm || !hNorm) return false;

  if (hNorm === qNorm) return true;
  if (hNorm.startsWith(qNorm) || hNorm.startsWith('lop' + qNorm)) return true;

  // Grade & Section matching: e.g. "10.2" matches "10.2", "10A2", "10B2", "10/2", "10-2", "Lớp 10A2"
  const parts = q.split(/[\.\/\-_\s]+/);
  if (parts.length === 2 && /^\d+$/.test(parts[0]) && /^\d+$/.test(parts[1])) {
    const grade = parts[0];
    const section = parts[1];
    const classNumPattern = new RegExp(`(^|[^0-9])${grade}[a-z\\.\\/\\-_\\s]*${section}([^0-9]|$)`, 'i');
    if (classNumPattern.test(h)) return true;
  }

  // Pattern matching: e.g. "10.2" matching "10.2", "10/2", "10-2"
  const dotPattern = q.replace(/\./g, '[\\.\\/\\-_\\s]?');
  const regex = new RegExp(`(^|[^0-9])${dotPattern}([^0-9]|$)`, 'i');
  return regex.test(h);
}

/**
 * Flexible teacher name matching:
 * Matches "Thầy Mịnh" vs "Thầy Đ. Minh", handling abbreviation, title, and diacritics
 */
export function matchTeacherFlexible(query: string, candidate: string): { matches: boolean; label: string } {
  if (!query || !candidate) return { matches: false, label: '' };

  const q = normalizeTeacherName(query).toLowerCase().replace(/thầy|cô/gi, '').trim();
  const c = normalizeTeacherName(candidate).toLowerCase().replace(/thầy|cô/gi, '').trim();

  if (!q || !c) return { matches: false, label: '' };

  // Check if query targets "Đ. Minh" / "Đoàn Minh"
  const isQueryDMinh = q.includes('đ.') || q.includes('đ ') || q.includes('đoàn') || q.includes('đặng') || q.includes('đinh') || q.includes('d.');
  const isCandidateDMinh = c.includes('đ.') || c.includes('đ ') || c.includes('đoàn') || c.includes('đặng') || c.includes('đinh') || c.includes('d.');

  if (isQueryDMinh) {
    if (isCandidateDMinh && (c.includes('minh') || c.includes('mịnh'))) {
      return { matches: true, label: 'Thầy Đ. Minh' };
    }
    return { matches: false, label: '' };
  }

  // Check if query targets "Thầy Mịnh" (or Minh without Đ)
  const isQueryMinhOrMinh = q.includes('mịnh') || q.includes('minh');
  if (isQueryMinhOrMinh && !isQueryDMinh) {
    // If candidate has "Đ." then it is Thầy Đ. Minh, not Thầy Mịnh!
    if (isCandidateDMinh) return { matches: false, label: '' };
    if (c.includes('mịnh') || c.includes('minh')) {
      return { matches: true, label: 'Thầy Mịnh' };
    }
  }

  // Fallback: substring matching
  if (c.includes(q) || q.includes(c)) {
    return { matches: true, label: query };
  }

  return { matches: false, label: '' };
}

/**
 * Detects whether Class 1 and Class 2 have simultaneous teaching sessions with Teacher 1 and Teacher 2:
 * "Cùng 1 thời điểm 2 lớp có học 2 thầy cùng lúc không"
 */
export function detectSimultaneousTeaching(
  rows: string[][],
  colMappings: ColumnMapping[],
  settings: ExtractionSettings,
  config: SimultaneousCheckConfig = {
    classQuery1: '10.2',
    classQuery2: '12.2',
    teacherQuery1: 'Thầy Mịnh',
    teacherQuery2: 'Thầy Đ. Minh'
  },
  profile?: 'THPT' | 'TieuHoc' | 'All'
): SimultaneousMatch[] {
  const matches: SimultaneousMatch[] = [];

  const dayCols = colMappings.filter(c => c.role === 'day');
  const periodCols = colMappings.filter(c => c.role === 'period');
  const classCols = colMappings.filter(c => c.role === 'class');

  // Filter columns matching classQuery1 and classQuery2
  const classCols1 = classCols.filter(c => matchClassFlexible(config.classQuery1, c.header));
  const classCols2 = classCols.filter(c => matchClassFlexible(config.classQuery2, c.header));

  if (classCols1.length === 0 || classCols2.length === 0) {
    return [];
  }

  let lastSeenDay = 'Chưa xác định';
  const dayRowCounters = new Map<string, number>();

  rows.forEach((row, rowIndex) => {
    // Determine Day
    let day = '';
    if (dayCols.length > 0) {
      for (const dCol of dayCols) {
        const val = row[dCol.index];
        if (val && String(val).trim()) {
          const dClean = String(val).trim();
          if (dClean.toLowerCase().includes('thứ') || dClean.toLowerCase().includes('chủ nhật')) {
            day = dClean;
            lastSeenDay = day;
            break;
          }
        }
      }
      if (!day) day = lastSeenDay;
    } else {
      if (profile === 'THPT' || profile === 'TieuHoc') {
        const rowNum = rowIndex + 6;
        if (rowNum >= 10 && rowNum <= 24) day = 'Thứ 2';
        else if (rowNum >= 28 && rowNum <= 43) day = 'Thứ 3';
        else if (rowNum >= 47 && rowNum <= 62) day = 'Thứ 4';
        else if (rowNum >= 66 && rowNum <= 81) day = 'Thứ 5';
        else if (rowNum >= 85 && rowNum <= 100) day = 'Thứ 6';
        else day = 'Nghỉ/Trống';
      } else {
        day = lastSeenDay;
      }
    }

    if ((profile === 'THPT' || profile === 'TieuHoc') && day === 'Nghỉ/Trống') {
      return;
    }

    const currentDayRow = dayRowCounters.get(day) || 0;
    dayRowCounters.set(day, currentDayRow + 1);

    // Determine Period & Session
    let periodText = '';
    for (const pCol of periodCols) {
      const cellVal = row[pCol.index];
      if (cellVal && String(cellVal).trim()) {
        periodText = String(cellVal).trim();
        break;
      }
    }

    const { session, displayName } = parsePeriodInfo(
      periodText,
      currentDayRow,
      rowIndex,
      profile
    );

    const periodDisplay = periodText || displayName;

    // Check Class 1 teachers in this row
    const class1Found: {
      className: string;
      teacher: string;
      subject?: string;
      originalValue: string;
      matchedQuery: string;
    }[] = [];

    for (const col of classCols1) {
      const cellVal = row[col.index];
      const extracted = extractTeacher(cellVal, settings);
      if (!extracted || !extracted.teacherName) continue;

      const m1 = matchTeacherFlexible(config.teacherQuery1, extracted.teacherName);
      const m2 = matchTeacherFlexible(config.teacherQuery2, extracted.teacherName);

      if (m1.matches) {
        class1Found.push({
          className: col.header,
          teacher: extracted.teacherName,
          subject: extracted.subjectName,
          originalValue: extracted.originalValue,
          matchedQuery: m1.label || config.teacherQuery1
        });
      } else if (m2.matches) {
        class1Found.push({
          className: col.header,
          teacher: extracted.teacherName,
          subject: extracted.subjectName,
          originalValue: extracted.originalValue,
          matchedQuery: m2.label || config.teacherQuery2
        });
      }
    }

    // Check Class 2 teachers in this row
    const class2Found: {
      className: string;
      teacher: string;
      subject?: string;
      originalValue: string;
      matchedQuery: string;
    }[] = [];

    for (const col of classCols2) {
      const cellVal = row[col.index];
      const extracted = extractTeacher(cellVal, settings);
      if (!extracted || !extracted.teacherName) continue;

      const m1 = matchTeacherFlexible(config.teacherQuery1, extracted.teacherName);
      const m2 = matchTeacherFlexible(config.teacherQuery2, extracted.teacherName);

      if (m1.matches) {
        class2Found.push({
          className: col.header,
          teacher: extracted.teacherName,
          subject: extracted.subjectName,
          originalValue: extracted.originalValue,
          matchedQuery: m1.label || config.teacherQuery1
        });
      } else if (m2.matches) {
        class2Found.push({
          className: col.header,
          teacher: extracted.teacherName,
          subject: extracted.subjectName,
          originalValue: extracted.originalValue,
          matchedQuery: m2.label || config.teacherQuery2
        });
      }
    }

    // If BOTH class 1 and class 2 have one of the target teachers at this same slot:
    if (class1Found.length > 0 && class2Found.length > 0) {
      class1Found.forEach(c1 => {
        class2Found.forEach(c2 => {
          const isSameTeacher = normalizeTeacherName(c1.teacher) === normalizeTeacherName(c2.teacher);
          const type: 'diff_teachers' | 'same_teacher' = isSameTeacher ? 'same_teacher' : 'diff_teachers';

          let description = '';
          if (isSameTeacher) {
            description = `Cùng thời điểm: Cả 2 lớp (${c1.className} và ${c2.className}) cùng học ${c1.teacher}!`;
          } else {
            description = `Cùng thời điểm: Lớp ${c1.className} học ${c1.teacher} và lớp ${c2.className} học ${c2.teacher}!`;
          }

          matches.push({
            id: `sim-${day}-${rowIndex}-${c1.className}-${c2.className}`,
            day,
            period: periodDisplay,
            session,
            rowIndex,
            class1: {
              className: c1.className,
              teacher: c1.teacher,
              subject: c1.subject,
              originalValue: c1.originalValue
            },
            class2: {
              className: c2.className,
              teacher: c2.teacher,
              subject: c2.subject,
              originalValue: c2.originalValue
            },
            type,
            description
          });
        });
      });
    }
  });

  return matches;
}
