export interface ExemptionPair {
  id: string;
  classA: string;
  classB: string;
  isActive: boolean;
}

export interface CellRange {
  startRow: number; // 0-indexed, inclusive
  endRow: number;   // 0-indexed, inclusive
  startCol: number; // 0-indexed, inclusive
  endCol: number;   // 0-indexed, inclusive
}

export type ColumnRole = 'day' | 'period' | 'class' | 'skip';

export interface ColumnMapping {
  index: number;
  header: string;
  role: ColumnRole;
}

export interface ParseResult {
  headers: string[];
  rows: string[][]; // 2D array representation of parsed sheet
  sheetNames: string[];
  activeSheetName: string;
}

export interface TeacherConflict {
  id: string;
  day: string;
  period: string;
  teacher: string;
  classes: string[]; // List of classes the teacher is teaching in this slot
  conflictPairs: [string, string][]; // List of actual conflicting class pairs (excluding exempted ones)
}

export interface TeacherInfo {
  originalValue: string;
  teacherName: string;
  subjectName?: string;
}

export interface ExtractionSettings {
  delimiter: string;
  takeFirstPart: boolean; // if false, take the second part as teacher name
  ignoreList: string[]; // words to ignore
}
