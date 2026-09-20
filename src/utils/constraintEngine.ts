import { 
  TimetableConstraint, 
  ConstraintCheckResult, 
  ConstraintViolation, 
  ColumnMapping, 
  ExtractionSettings, 
  DaySession 
} from '../types';
import { 
  extractTeacher, 
  parsePeriodInfo, 
  matchClassFlexible, 
  matchTeacherFlexible,
  detectSimultaneousTeaching
} from './excelParser';

export function normalizeDay(day: string): string {
  if (!day) return 't2';
  const d = day.trim().toLowerCase();
  if (d.includes('2') || d.includes('hai')) return 't2';
  if (d.includes('3') || d.includes('ba')) return 't3';
  if (d.includes('4') || d.includes('tư') || d.includes('tu')) return 't4';
  if (d.includes('5') || d.includes('năm') || d.includes('nam')) return 't5';
  if (d.includes('6') || d.includes('sáu') || d.includes('sau')) return 't6';
  if (d.includes('7') || d.includes('bảy') || d.includes('bay')) return 't7';
  if (d.includes('nhật') || d.includes('cn')) return 'cn';
  return d;
}

export function formatDay(day: string): string {
  const norm = normalizeDay(day);
  switch (norm) {
    case 't2': return 'Thứ 2';
    case 't3': return 'Thứ 3';
    case 't4': return 'Thứ 4';
    case 't5': return 'Thứ 5';
    case 't6': return 'Thứ 6';
    case 't7': return 'Thứ 7';
    case 'cn': return 'Chủ Nhật';
    default: return day || 'Thứ 2';
  }
}

export function formatClassName(name: string): string {
  if (!name) return '';
  const trimmed = name.trim();
  if (/^(lớp|lop)\s*/i.test(trimmed)) return trimmed;
  return `Lớp ${trimmed}`;
}

export const CONSTRAINTS_STORAGE_KEY = 'checksched_custom_constraints_v1';

export const DEFAULT_CONSTRAINTS: TimetableConstraint[] = [
  {
    id: 'c-simul-10-12',
    name: 'Kiểm tra trùng giờ Lớp 10.2 & 12.2 (Thầy Mịnh & Thầy Đ. Minh)',
    description: 'Không được xếp trùng cùng một tiết giữa Lớp 10.2 và 12.2 khi một lớp học Thầy Mịnh và lớp còn lại học Thầy Đ. Minh.',
    type: 'simultaneous_classes_teachers',
    isActive: true,
    params: {
      class1: '10.2',
      class2: '12.2',
      teacher1: 'Thầy Mịnh',
      teacher2: 'Thầy Đ. Minh',
      mode: 'must_not_overlap'
    }
  },
  {
    id: 'c-no-split-session',
    name: 'Không chia môn của GV trong ngày (cả Sáng và Chiều)',
    description: 'Tránh việc cùng 1 môn học của cùng 1 giáo viên trong cùng 1 lớp bị xé lẻ dạy ở cả buổi sáng lẫn buổi chiều.',
    type: 'no_split_day',
    isActive: true,
    params: {}
  },
  {
    id: 'c-max-periods-teacher',
    name: 'Giáo viên không dạy quá 5 tiết trong 1 ngày',
    description: 'Khống chế số tiết tối đa trong một ngày của mỗi giáo viên để đảm bảo chất lượng giảng dạy và định mức lao động.',
    type: 'teacher_max_periods_day',
    isActive: true,
    params: {
      maxPeriodsPerDay: 5
    }
  },
  {
    id: 'c-no-gap-morning',
    name: 'Hạn chế giáo viên có tiết trống cách quãng (loãng xương)',
    description: 'Tránh trường hợp giáo viên dạy tiết 1 và tiết 4 hoặc 5 mà bỏ trống các tiết ở giữa trong cùng một buổi.',
    type: 'teacher_no_gap',
    isActive: false,
    params: {}
  }
];

/**
 * Load constraints from localStorage, falling back to default constraints
 */
export function getStoredConstraints(): TimetableConstraint[] {
  try {
    const raw = localStorage.getItem(CONSTRAINTS_STORAGE_KEY);
    if (!raw) return DEFAULT_CONSTRAINTS;
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed) && parsed.length > 0) {
      return parsed;
    }
  } catch (err) {
    console.error('Failed to load constraints from localStorage:', err);
  }
  return DEFAULT_CONSTRAINTS;
}

/**
 * Save constraints to localStorage
 */
export function saveStoredConstraints(constraints: TimetableConstraint[]): void {
  try {
    localStorage.setItem(CONSTRAINTS_STORAGE_KEY, JSON.stringify(constraints));
  } catch (err) {
    console.error('Failed to save constraints to localStorage:', err);
  }
}

interface ParsedSlot {
  day: string;
  dayNorm: string;
  period: string;
  periodNum: number;
  session: DaySession;
  className: string;
  teacher: string;
  subject?: string;
  originalValue: string;
  rowIndex: number;
}

/**
 * Evaluate all user-defined constraints against the parsed Excel timetable
 */
export function evaluateConstraints(
  rows: string[][],
  colMappings: ColumnMapping[],
  settings: ExtractionSettings,
  constraints: TimetableConstraint[],
  profile?: 'THPT' | 'TieuHoc' | 'All'
): ConstraintCheckResult[] {
  if (!rows || rows.length === 0 || !colMappings || colMappings.length === 0) {
    return constraints.map(c => ({
      constraintId: c.id,
      constraint: c,
      isSatisfied: true,
      violationsCount: 0,
      violations: []
    }));
  }

  // 1. Identify Day column, Period column, and Class columns
  const dayCol = colMappings.find(c => c.role === 'day');
  const periodCol = colMappings.find(c => c.role === 'period');
  const classCols = colMappings.filter(c => c.role === 'class');

  // 2. Parse all valid slots across the entire timetable
  const slots: ParsedSlot[] = [];
  let lastSeenDay = 'Thứ 2';

  for (let rIdx = 0; rIdx < rows.length; rIdx++) {
    const row = rows[rIdx];
    if (!row) continue;

    // Day handling with carryover & profile awareness
    let day = '';
    if (dayCol && row[dayCol.index]) {
      const rawDay = String(row[dayCol.index]).trim();
      if (rawDay && (rawDay.toLowerCase().includes('thứ') || rawDay.toLowerCase().includes('chủ nhật'))) {
        day = rawDay;
        lastSeenDay = day;
      } else {
        day = lastSeenDay;
      }
    } else {
      if (profile === 'THPT' || profile === 'TieuHoc') {
        const rowNum = rIdx + 6; // rIdx = 0 matches Excel Row 6
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
      continue;
    }

    const dayNorm = normalizeDay(day);

    // Period handling
    let rawPeriod = periodCol && row[periodCol.index] ? String(row[periodCol.index]).trim() : '';
    const periodInfo = parsePeriodInfo(rawPeriod, undefined, rIdx, profile);
    const periodDisplay = rawPeriod || periodInfo.displayName || `Tiết ${periodInfo.periodNum}`;

    // Check each class column
    for (const cCol of classCols) {
      const cellVal = row[cCol.index];
      if (!cellVal || typeof cellVal !== 'string' || !cellVal.trim()) continue;

      const teacherInfo = extractTeacher(cellVal, settings);
      if (!teacherInfo || !teacherInfo.teacherName) continue;

      slots.push({
        day,
        dayNorm,
        period: periodDisplay,
        periodNum: periodInfo.periodNum,
        session: periodInfo.session,
        className: cCol.header,
        teacher: teacherInfo.teacherName,
        subject: teacherInfo.subjectName,
        originalValue: cellVal,
        rowIndex: rIdx + 6 // Excel row number (1-indexed matching sheet)
      });
    }
  }

  // 3. Evaluate each constraint
  const results: ConstraintCheckResult[] = [];

  for (const constraint of constraints) {
    // If not active, mark as satisfied with 0 violations
    if (!constraint.isActive) {
      results.push({
        constraintId: constraint.id,
        constraint,
        isSatisfied: true,
        satisfied: true,
        violationsCount: 0,
        violations: []
      });
      continue;
    }

    const violations: ConstraintViolation[] = [];

    switch (constraint.type) {
      case 'simultaneous_classes_teachers': {
        const qClass1 = constraint.params.class1?.trim() || '10.2';
        const qClass2 = constraint.params.class2?.trim() || '12.2';
        const qTeacher1 = constraint.params.teacher1?.trim() || 'Thầy Mịnh';
        const qTeacher2 = constraint.params.teacher2?.trim() || 'Thầy Đ. Minh';

        // Direct call to detectSimultaneousTeaching ensures 100% synchronization between Card 3 and Tab 4!
        const simMatches = detectSimultaneousTeaching(
          rows,
          colMappings,
          settings,
          {
            classQuery1: qClass1,
            classQuery2: qClass2,
            teacherQuery1: qTeacher1,
            teacherQuery2: qTeacher2
          },
          profile
        );

        for (const m of simMatches) {
          violations.push({
            id: `viol-${constraint.id}-${m.id}`,
            day: formatDay(m.day),
            period: m.period,
            session: m.session,
            className: `${formatClassName(m.class1.className)} & ${formatClassName(m.class2.className)}`,
            teacher: `${m.class1.teacher} / ${m.class2.teacher}`,
            subject: `${m.class1.subject || 'Môn 1'} & ${m.class2.subject || 'Môn 2'}`,
            rowIndex: m.rowIndex + 6,
            message: m.description
          });
        }
        break;
      }

      case 'no_split_day': {
        const targetClass = constraint.params.targetClass?.trim().toLowerCase();
        const targetTeacher = constraint.params.targetTeacher?.trim().toLowerCase();

        // Group by (className, teacher, subject, dayNorm)
        const groupMap = new Map<string, ParsedSlot[]>();
        for (const slot of slots) {
          if (targetClass && !matchClassFlexible(targetClass, slot.className)) continue;
          if (targetTeacher && !slot.teacher.toLowerCase().includes(targetTeacher)) continue;

          const key = `${slot.className}__${slot.teacher}__${slot.subject || 'all'}__${slot.dayNorm}`;
          if (!groupMap.has(key)) groupMap.set(key, []);
          groupMap.get(key)!.push(slot);
        }

        groupMap.forEach((gSlots) => {
          const morning = gSlots.filter(s => s.session === 'Sáng');
          const afternoon = gSlots.filter(s => s.session === 'Chiều');

          if (morning.length > 0 && afternoon.length > 0) {
            const firstSlot = gSlots[0];
            const morningPeriodsStr = morning.map(s => s.period).join(', ');
            const afternoonPeriodsStr = afternoon.map(s => s.period).join(', ');

            violations.push({
              id: `viol-${constraint.id}-${firstSlot.className}-${firstSlot.teacher}-${firstSlot.dayNorm}`,
              day: formatDay(firstSlot.day),
              period: `Sáng: ${morningPeriodsStr} | Chiều: ${afternoonPeriodsStr}`,
              session: 'Khác',
              className: formatClassName(firstSlot.className),
              teacher: firstSlot.teacher,
              subject: firstSlot.subject,
              rowIndex: firstSlot.rowIndex,
              message: `Môn ${firstSlot.subject || 'học'} lớp ${formatClassName(firstSlot.className)} của ${firstSlot.teacher} bị chia cả Sáng (${morningPeriodsStr}) và Chiều (${afternoonPeriodsStr}) vào ${formatDay(firstSlot.day)}.`
            });
          }
        });
        break;
      }

      case 'teacher_max_periods_day': {
        const max = constraint.params.maxPeriodsPerDay || 5;
        const targetTeacher = constraint.params.teacher?.trim().toLowerCase();

        // Group by (teacher, dayNorm)
        const teacherDayMap = new Map<string, ParsedSlot[]>();
        for (const slot of slots) {
          if (targetTeacher && !slot.teacher.toLowerCase().includes(targetTeacher)) continue;
          const key = `${slot.teacher}__${slot.dayNorm}`;
          if (!teacherDayMap.has(key)) teacherDayMap.set(key, []);
          teacherDayMap.get(key)!.push(slot);
        }

        teacherDayMap.forEach((tSlots) => {
          // Count unique periods taught in this day (in case teacher teaches multiple classes simultaneously)
          const periodKeys = new Set(tSlots.map(s => `${s.session}_${s.periodNum}`));
          const periodCount = periodKeys.size;

          if (periodCount > max) {
            const firstSlot = tSlots[0];
            violations.push({
              id: `viol-${constraint.id}-${firstSlot.teacher}-${firstSlot.dayNorm}`,
              day: formatDay(firstSlot.day),
              period: `${periodCount} tiết`,
              session: 'Khác',
              teacher: firstSlot.teacher,
              className: Array.from(new Set(tSlots.map(s => formatClassName(s.className)))).join(', '),
              rowIndex: firstSlot.rowIndex,
              message: `Giáo viên ${firstSlot.teacher} dạy ${periodCount} tiết vào ${formatDay(firstSlot.day)} (vượt mức quy định tối đa ${max} tiết/ngày).`
            });
          }
        });
        break;
      }

      case 'teacher_unavailable': {
        const targetTeacher = constraint.params.teacher?.trim().toLowerCase() || '';
        const restrictedDays = (constraint.params.days || []).map(d => normalizeDay(d));
        const restrictedSessions = constraint.params.sessions || [];
        const restrictedPeriods = (constraint.params.periods || []).map(p => p.trim().toLowerCase());

        for (const slot of slots) {
          if (targetTeacher && !slot.teacher.toLowerCase().includes(targetTeacher)) continue;

          let isViolated = true;
          if (restrictedDays.length > 0 && !restrictedDays.includes(slot.dayNorm)) {
            isViolated = false;
          }
          if (restrictedSessions.length > 0 && !restrictedSessions.includes(slot.session)) {
            isViolated = false;
          }
          if (restrictedPeriods.length > 0) {
            const pStr = `tiết ${slot.periodNum}`;
            const matchesPeriod = restrictedPeriods.some(rp => 
              slot.period.toLowerCase().includes(rp) || 
              rp.includes(`${slot.periodNum}`) ||
              pStr === rp
            );
            if (!matchesPeriod) isViolated = false;
          }

          if (isViolated) {
            violations.push({
              id: `viol-${constraint.id}-${slot.rowIndex}-${slot.className}`,
              day: formatDay(slot.day),
              period: slot.period,
              session: slot.session,
              teacher: slot.teacher,
              className: formatClassName(slot.className),
              subject: slot.subject,
              rowIndex: slot.rowIndex,
              message: `Giáo viên ${slot.teacher} có tiết dạy tại lớp ${formatClassName(slot.className)} vào ${slot.period} ${slot.session} ${formatDay(slot.day)} (thời điểm đã đăng ký bận/nghỉ).`
            });
          }
        }
        break;
      }

      case 'subject_period_restriction': {
        const targetSubject = constraint.params.subject?.trim().toLowerCase() || '';
        const restrictedDays = (constraint.params.restrictedDays || []).map(d => normalizeDay(d));
        const restrictedSessions = constraint.params.restrictedSessions || [];
        const restrictedPeriods = (constraint.params.restrictedPeriods || []).map(p => p.trim().toLowerCase());

        for (const slot of slots) {
          const subjectNorm = (slot.subject || '').toLowerCase();
          if (targetSubject && !subjectNorm.includes(targetSubject)) continue;

          let isViolated = true;
          if (restrictedDays.length > 0 && !restrictedDays.includes(slot.dayNorm)) {
            isViolated = false;
          }
          if (restrictedSessions.length > 0 && !restrictedSessions.includes(slot.session)) {
            isViolated = false;
          }
          if (restrictedPeriods.length > 0) {
            const matchesPeriod = restrictedPeriods.some(rp => 
              slot.period.toLowerCase().includes(rp) || rp.includes(`${slot.periodNum}`)
            );
            if (!matchesPeriod) isViolated = false;
          }

          if (isViolated) {
            violations.push({
              id: `viol-${constraint.id}-${slot.rowIndex}-${slot.className}`,
              day: formatDay(slot.day),
              period: slot.period,
              session: slot.session,
              className: formatClassName(slot.className),
              teacher: slot.teacher,
              subject: slot.subject,
              rowIndex: slot.rowIndex,
              message: `Môn ${slot.subject || targetSubject} của lớp ${formatClassName(slot.className)} được xếp vào ${slot.period} ${slot.session} ${formatDay(slot.day)} (trái với ràng buộc quy định môn học).`
            });
          }
        }
        break;
      }

      case 'class_pair_no_overlap': {
        const c1 = constraint.params.overlapClass1?.trim() || '';
        const c2 = constraint.params.overlapClass2?.trim() || '';

        // Group by (dayNorm, periodNum, session)
        const slotMap = new Map<string, ParsedSlot[]>();
        for (const slot of slots) {
          const key = `${slot.dayNorm}__${slot.periodNum}__${slot.session}`;
          if (!slotMap.has(key)) slotMap.set(key, []);
          slotMap.get(key)!.push(slot);
        }

        slotMap.forEach((timeSlots, key) => {
          const matchesC1 = timeSlots.filter(s => matchClassFlexible(c1, s.className));
          const matchesC2 = timeSlots.filter(s => matchClassFlexible(c2, s.className));

          if (matchesC1.length > 0 && matchesC2.length > 0) {
            const s1 = matchesC1[0];
            const s2 = matchesC2[0];
            violations.push({
              id: `viol-${constraint.id}-${key}`,
              day: formatDay(s1.day),
              period: s1.period,
              session: s1.session,
              className: `${formatClassName(s1.className)} & ${formatClassName(s2.className)}`,
              teacher: `${s1.teacher} / ${s2.teacher}`,
              subject: `${s1.subject || 'Môn 1'} / ${s2.subject || 'Môn 2'}`,
              rowIndex: s1.rowIndex,
              message: `Lớp ${formatClassName(s1.className)} và Lớp ${formatClassName(s2.className)} cùng học vào ${s1.period} ${formatDay(s1.day)} (vi phạm không được trùng phòng/sân bãi).`
            });
          }
        });
        break;
      }

      case 'teacher_no_gap': {
        const targetTeacher = constraint.params.teacher?.trim().toLowerCase();

        // Group by (teacher, dayNorm, session)
        const tSessionMap = new Map<string, ParsedSlot[]>();
        for (const slot of slots) {
          if (targetTeacher && !slot.teacher.toLowerCase().includes(targetTeacher)) continue;
          const key = `${slot.teacher}__${slot.dayNorm}__${slot.session}`;
          if (!tSessionMap.has(key)) tSessionMap.set(key, []);
          tSessionMap.get(key)!.push(slot);
        }

        tSessionMap.forEach((tsSlots) => {
          const periodNums = Array.from(new Set(tsSlots.map(s => s.periodNum))).sort((a, b) => a - b);
          if (periodNums.length < 2) return;

          // Check if there are gaps between min and max period
          const minP = periodNums[0];
          const maxP = periodNums[periodNums.length - 1];
          const gaps: number[] = [];

          for (let p = minP + 1; p < maxP; p++) {
            if (!periodNums.includes(p)) {
              gaps.push(p);
            }
          }

          if (gaps.length > 0) {
            const firstSlot = tsSlots[0];
            violations.push({
              id: `viol-${constraint.id}-${firstSlot.teacher}-${firstSlot.dayNorm}-${firstSlot.session}`,
              day: formatDay(firstSlot.day),
              period: `Tiết dạy: ${periodNums.join(', ')} (Trống tiết ${gaps.join(', ')})`,
              session: firstSlot.session,
              teacher: firstSlot.teacher,
              className: Array.from(new Set(tsSlots.map(s => formatClassName(s.className)))).join(', '),
              rowIndex: firstSlot.rowIndex,
              message: `Giáo viên ${firstSlot.teacher} bị trống tiết ${gaps.join(', ')} vào buổi ${firstSlot.session} ${formatDay(firstSlot.day)}.`
            });
          }
        });
        break;
      }
    }

    results.push({
      constraintId: constraint.id,
      constraint,
      isSatisfied: violations.length === 0,
      satisfied: violations.length === 0,
      violationsCount: violations.length,
      violations
    });
  }

  return results;
}
