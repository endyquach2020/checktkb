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

export type DaySession = 'Sáng' | 'Chiều' | 'Khác';

export interface SplitPeriodSlot {
  period: string;
  session: DaySession;
  periodNum: number;
  rowIndex: number;
  originalValue: string;
}

export interface SplitPeriodIssue {
  id: string;
  day: string;
  className: string;
  teacher: string;
  subject?: string;
  splitType: 'morning_afternoon' | 'isolated_periods' | 'both';
  morningPeriods: string[];
  afternoonPeriods: string[];
  totalPeriods: number;
  description: string;
  slots: SplitPeriodSlot[];
}

export interface SimultaneousMatch {
  id: string;
  day: string;
  period: string;
  session: DaySession;
  rowIndex: number;
  class1: {
    className: string;
    teacher: string;
    subject?: string;
    originalValue: string;
  };
  class2: {
    className: string;
    teacher: string;
    subject?: string;
    originalValue: string;
  };
  type: 'diff_teachers' | 'same_teacher';
  description: string;
}

export interface SimultaneousCheckConfig {
  classQuery1: string;
  classQuery2: string;
  teacherQuery1: string;
  teacherQuery2: string;
}

export type ConstraintType = 
  | 'simultaneous_classes_teachers'  // 2 lớp & 2 giáo viên (không được cùng giờ hoặc phải cùng giờ)
  | 'teacher_unavailable'           // Giáo viên không dạy vào Thứ / Buổi / Tiết cụ thể
  | 'teacher_max_periods_day'       // Số tiết tối đa trong 1 ngày của GV
  | 'no_split_day'                  // Không chia môn của GV ra cả 2 buổi sáng & chiều
  | 'subject_period_restriction'    // Môn học không được xếp vào Tiết/Buổi/Thứ cụ thể
  | 'class_pair_no_overlap'         // Hai lớp không được học cùng giờ
  | 'teacher_no_gap';               // GV không bị tiết trống (loãng xương) trong buổi

export interface ConstraintViolation {
  id: string;
  day: string;
  period: string;
  session?: DaySession;
  className?: string;
  teacher?: string;
  subject?: string;
  rowIndex?: number;
  message: string;
}

export interface TimetableConstraint {
  id: string;
  name: string;
  description?: string;
  type: ConstraintType;
  isActive: boolean;
  params: {
    // For simultaneous_classes_teachers:
    class1?: string;
    class2?: string;
    teacher1?: string;
    teacher2?: string;
    mode?: 'must_not_overlap' | 'must_overlap';
    
    // For teacher_unavailable:
    teacher?: string;
    days?: string[];      // ['Thứ 2', 'Thứ 7']
    sessions?: DaySession[]; // ['Sáng', 'Chiều']
    periods?: string[];   // ['Tiết 1', 'Tiết 5']
    
    // For teacher_max_periods_day:
    maxPeriodsPerDay?: number;
    
    // For no_split_day:
    targetClass?: string;
    targetTeacher?: string;

    // For subject_period_restriction:
    subject?: string;
    restrictedDays?: string[];
    restrictedPeriods?: string[];
    restrictedSessions?: DaySession[];

    // For class_pair_no_overlap:
    overlapClass1?: string;
    overlapClass2?: string;
  };
}

export interface ConstraintCheckResult {
  constraintId: string;
  constraint: TimetableConstraint;
  isSatisfied: boolean;
  satisfied?: boolean;
  violationsCount: number;
  violations: ConstraintViolation[];
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
