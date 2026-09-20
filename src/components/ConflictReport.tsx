import { useState, useMemo } from 'react';
import { TeacherConflict, ExemptionPair, SplitPeriodIssue } from '../types';
import { 
  CheckCircle2, 
  AlertTriangle, 
  Search, 
  Clock, 
  User, 
  Users, 
  Filter, 
  Download,
  BookOpen,
  ArrowRight,
  Sun,
  Moon,
  Layers,
  Sparkles
} from 'lucide-react';
import * as XLSX from 'xlsx';

// Helper to format days to match standard format (e.g., "Thứ Hai" -> "Thứ 2")
const formatDay = (day: string): string => {
  if (!day) return 'Chưa xác định';
  const d = day.trim().toLowerCase();
  if (d === 'thứ hai' || d === 'thu hai' || d === 'thứ 2' || d === 't2' || d === 'hai') return 'Thứ 2';
  if (d === 'thứ ba' || d === 'thu ba' || d === 'thứ 3' || d === 't3' || d === 'ba') return 'Thứ 3';
  if (d === 'thứ tư' || d === 'thu tu' || d === 'thứ 4' || d === 't4' || d === 'tư' || d === 'tu') return 'Thứ 4';
  if (d === 'thứ năm' || d === 'thu nam' || d === 'thứ 5' || d === 't5' || d === 'năm' || d === 'nam') return 'Thứ 5';
  if (d === 'thứ sáu' || d === 'thu sau' || d === 'thứ 6' || d === 't6' || d === 'sáu' || d === 'sau') return 'Thứ 6';
  if (d === 'thứ bảy' || d === 'thu bay' || d === 'thứ 7' || d === 't7' || d === 'bảy' || d === 'bay') return 'Thứ 7';
  if (d === 'chủ nhật' || d === 'chu nhat' || d === 'cn') return 'Chủ Nhật';
  return day;
};

// Helper to format class name to include "Lớp " prefix if not already present
const formatClassName = (className: string): string => {
  if (!className) return '';
  const trimmed = className.trim();
  if (trimmed.toLowerCase().startsWith('lớp') || trimmed.toLowerCase().startsWith('lop')) {
    const content = trimmed.replace(/^(lớp|lop)\s*/i, '').trim();
    return `Lớp ${content}`;
  }
  return `Lớp ${trimmed}`;
};

// Helper to strip trailing "TN" and "XH" (case-insensitive) for name base comparison
const cleanClassBase = (clsName: string): string => {
  if (!clsName) return '';
  let cleaned = clsName.trim().replace(/^(lớp|lop)\s*/i, '').trim();
  // Strip trailing "TN" or "XH" (case-insensitive)
  cleaned = cleaned.replace(/\s*(tn|xh)$/i, '').trim();
  return cleaned;
};

// Check if two classes are a valid partner pair (same base, or 10.1 & 10.2, 11.1 & 11.2, 12.1 & 12.2)
const isWithinValidPair = (cls1: string, cls2: string): boolean => {
  const c1 = cleanClassBase(cls1);
  const c2 = cleanClassBase(cls2);
  if (c1 === c2) return true;
  if ((c1 === '10.1' && c2 === '10.2') || (c1 === '10.2' && c2 === '10.1')) return true;
  if ((c1 === '11.1' && c2 === '11.2') || (c1 === '11.2' && c2 === '11.1')) return true;
  if ((c1 === '12.1' && c2 === '12.2') || (c1 === '12.2' && c2 === '12.1')) return true;
  return false;
};

// Determine class badge styling based on whether it conflicts with an unexpected class
const getClassBadgeStyles = (clsName: string, allClassNames: string[]): string => {
  const hasInvalidConflict = allClassNames.some(other => other !== clsName && !isWithinValidPair(clsName, other));
  if (hasInvalidConflict) {
    return 'bg-purple-100 text-purple-800 border-purple-300 shadow-purple-600/5 ring-1 ring-purple-300/10 font-black';
  }
  return 'bg-rose-50 text-rose-700 border-rose-100 shadow-rose-600/5';
};

// Determine conflict pair styling based on whether the specific pair is a valid partner pair
const getPairBadgeStyles = (pair: [string, string]): string => {
  const isPairValid = isWithinValidPair(pair[0], pair[1]);
  if (!isPairValid) {
    return 'bg-purple-50 text-purple-700 border-purple-200 shadow-sm shadow-purple-600/5 ring-1 ring-purple-300/10 font-bold';
  }
  return 'bg-rose-50/70 text-rose-700 border-rose-100 shadow-sm shadow-rose-600/5';
};

interface ConflictReportProps {
  conflicts: TeacherConflict[];
  splitIssues?: SplitPeriodIssue[];
  exemptions: ExemptionPair[];
  totalCellsChecked: number;
  totalTeachersFound: number;
}

export default function ConflictReport({
  conflicts,
  splitIssues = [],
  exemptions,
  totalCellsChecked,
  totalTeachersFound
}: ConflictReportProps) {
  // Tab state: 'conflicts' or 'splitPeriods'
  const [activeTab, setActiveTab] = useState<'conflicts' | 'splitPeriods'>(() => {
    if (conflicts.length === 0 && splitIssues.length > 0) {
      return 'splitPeriods';
    }
    return 'conflicts';
  });

  // Conflicts filter state
  const [searchTerm, setSearchTerm] = useState('');
  const [dayFilter, setDayFilter] = useState('all');

  // Split periods filter state
  const [splitSearchTerm, setSplitSearchTerm] = useState('');
  const [splitDayFilter, setSplitDayFilter] = useState('all');
  const [splitTypeFilter, setSplitTypeFilter] = useState<'all' | 'morning_afternoon' | 'isolated_periods' | 'both'>('all');

  // Extract unique days for the conflict filter dropdown
  const uniqueConflictDays = useMemo(() => {
    const days = new Set<string>();
    conflicts.forEach(c => {
      if (c.day) days.add(c.day);
    });
    return Array.from(days);
  }, [conflicts]);

  // Extract unique days for the split periods filter dropdown
  const uniqueSplitDays = useMemo(() => {
    const days = new Set<string>();
    splitIssues.forEach(s => {
      if (s.day) days.add(s.day);
    });
    return Array.from(days);
  }, [splitIssues]);

  // Filter conflicts
  const filteredConflicts = useMemo(() => {
    return conflicts.filter(c => {
      const matchSearch = 
        c.teacher.toLowerCase().includes(searchTerm.toLowerCase()) ||
        c.classes.some(cls => cls.toLowerCase().includes(searchTerm.toLowerCase())) ||
        c.period.toLowerCase().includes(searchTerm.toLowerCase());

      const matchDay = dayFilter === 'all' || c.day === dayFilter;

      return matchSearch && matchDay;
    });
  }, [conflicts, searchTerm, dayFilter]);

  // Filter split issues
  const filteredSplitIssues = useMemo(() => {
    return splitIssues.filter(item => {
      const term = splitSearchTerm.toLowerCase();
      const matchSearch = 
        !term ||
        item.teacher.toLowerCase().includes(term) ||
        item.className.toLowerCase().includes(term) ||
        (item.subject && item.subject.toLowerCase().includes(term)) ||
        item.morningPeriods.some(p => p.toLowerCase().includes(term)) ||
        item.afternoonPeriods.some(p => p.toLowerCase().includes(term));

      const matchDay = splitDayFilter === 'all' || item.day === splitDayFilter;

      const matchType = 
        splitTypeFilter === 'all' || 
        (splitTypeFilter === 'morning_afternoon' && (item.splitType === 'morning_afternoon' || item.splitType === 'both')) ||
        (splitTypeFilter === 'isolated_periods' && (item.splitType === 'isolated_periods' || item.splitType === 'both'));

      return matchSearch && matchDay && matchType;
    });
  }, [splitIssues, splitSearchTerm, splitDayFilter, splitTypeFilter]);

  // Export conflict report to Excel
  const handleExportConflicts = () => {
    if (conflicts.length === 0) return;

    const wb = XLSX.utils.book_new();

    const exportData = filteredConflicts.map((c, index) => ({
      'STT': index + 1,
      'Thứ': formatDay(c.day),
      'Tiết / Khung giờ': c.period || 'Chưa xác định',
      'Tên Giáo Viên bị trùng': c.teacher,
      'Các Lớp bị trùng': c.classes.map(formatClassName).join(', '),
      'Chi tiết các cặp trùng thực tế': c.conflictPairs.map(p => `${formatClassName(p[0])} vs ${formatClassName(p[1])}`).join('; ')
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    
    ws['!cols'] = [
      { wch: 6 },  // STT
      { wch: 15 }, // Thứ
      { wch: 20 }, // Tiết / Khung giờ
      { wch: 25 }, // Tên Giáo Viên bị trùng
      { wch: 35 }, // Các Lớp bị trùng
      { wch: 60 }, // Chi tiết các cặp trùng thực tế
    ];

    XLSX.utils.book_append_sheet(wb, ws, 'Báo cáo trùng lịch');
    XLSX.writeFile(wb, 'bao_cao_trung_lich_day.xlsx');
  };

  // Export split periods report to Excel
  const handleExportSplitPeriods = () => {
    if (splitIssues.length === 0) return;

    const wb = XLSX.utils.book_new();

    const exportData = filteredSplitIssues.map((item, index) => {
      let typeText = 'Chia 2 buổi (Sáng & Chiều)';
      if (item.splitType === 'isolated_periods') typeText = 'Cách quãng trong buổi';
      if (item.splitType === 'both') typeText = 'Cả 2 buổi & Cách quãng';

      return {
        'STT': index + 1,
        'Thứ': formatDay(item.day),
        'Lớp học': formatClassName(item.className),
        'Tên Giáo viên': item.teacher,
        'Môn học': item.subject || 'Chung',
        'Phân loại bị chia': typeText,
        'Tiết Sáng': item.morningPeriods.length > 0 ? item.morningPeriods.join(', ') : 'Không có',
        'Tiết Chiều': item.afternoonPeriods.length > 0 ? item.afternoonPeriods.join(', ') : 'Không có',
        'Tổng số tiết': item.totalPeriods,
        'Mô tả chi tiết': item.description
      };
    });

    const ws = XLSX.utils.json_to_sheet(exportData);

    ws['!cols'] = [
      { wch: 6 },  // STT
      { wch: 15 }, // Thứ
      { wch: 18 }, // Lớp học
      { wch: 25 }, // Tên Giáo viên
      { wch: 20 }, // Môn học
      { wch: 28 }, // Phân loại bị chia
      { wch: 30 }, // Tiết Sáng
      { wch: 30 }, // Tiết Chiều
      { wch: 15 }, // Tổng số tiết
      { wch: 45 }, // Mô tả chi tiết
    ];

    XLSX.utils.book_append_sheet(wb, ws, 'Tiet_Day_Bi_Chia');
    XLSX.writeFile(wb, 'bao_cao_tiet_day_bi_chia.xlsx');
  };

  const activeExemptionsCount = exemptions.filter(e => e.isActive).length;

  return (
    <div id="conflict-report-container" className="space-y-6">
      
      {/* Summary Stat Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Total Cells */}
        <div className="bg-white rounded-2xl border border-slate-200/60 p-4 shadow-sm flex items-center gap-3.5 transition-all hover:shadow-md/5" id="stats-total-checked">
          <div className="p-2.5 bg-teal-50 text-teal-600 rounded-xl shrink-0">
            <BookOpen className="w-5 h-5" />
          </div>
          <div>
            <div className="text-xl font-extrabold text-slate-800 font-mono leading-none font-display">{totalCellsChecked}</div>
            <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider mt-1.5">Ô đã rà soát</div>
          </div>
        </div>

        {/* Teachers */}
        <div className="bg-white rounded-2xl border border-slate-200/60 p-4 shadow-sm flex items-center gap-3.5 transition-all hover:shadow-md/5" id="stats-teachers-found">
          <div className="p-2.5 bg-emerald-50 text-emerald-600 rounded-xl shrink-0">
            <User className="w-5 h-5" />
          </div>
          <div>
            <div className="text-xl font-extrabold text-slate-800 font-mono leading-none font-display">{totalTeachersFound}</div>
            <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider mt-1.5">Số Giáo viên</div>
          </div>
        </div>

        {/* Conflict Count (Clickable to switch tab) */}
        <div 
          onClick={() => setActiveTab('conflicts')}
          role="button"
          tabIndex={0}
          className={`rounded-2xl border p-4 shadow-sm flex items-center gap-3.5 transition-all cursor-pointer hover:shadow-md/10 ${
            activeTab === 'conflicts' ? 'ring-2 ring-rose-400/30' : ''
          } ${
            conflicts.length > 0 
              ? 'bg-rose-50/60 border-rose-200/80 text-rose-800' 
              : 'bg-emerald-50/60 border-emerald-200/80 text-emerald-800'
          }`} 
          id="stats-conflict-count"
        >
          <div className={`p-2.5 rounded-xl shrink-0 ${
            conflicts.length > 0 ? 'bg-rose-100 text-rose-600' : 'bg-emerald-100 text-emerald-600'
          }`}>
            {conflicts.length > 0 ? <AlertTriangle className="w-5 h-5" /> : <CheckCircle2 className="w-5 h-5" />}
          </div>
          <div className="flex-1">
            <div className="flex items-center justify-between">
              <div className="text-xl font-extrabold font-mono leading-none font-display">
                {conflicts.length}
              </div>
              {activeTab === 'conflicts' && (
                <span className="text-[9px] px-1.5 py-0.5 bg-rose-200 text-rose-800 rounded font-bold uppercase">Đang xem</span>
              )}
            </div>
            <div className="text-[10px] opacity-75 font-bold uppercase tracking-wider mt-1.5">Trùng tiết cùng giờ</div>
          </div>
        </div>

        {/* Split Periods Count (Clickable to switch tab) */}
        <div 
          onClick={() => setActiveTab('splitPeriods')}
          role="button"
          tabIndex={0}
          className={`rounded-2xl border p-4 shadow-sm flex items-center gap-3.5 transition-all cursor-pointer hover:shadow-md/10 ${
            activeTab === 'splitPeriods' ? 'ring-2 ring-amber-400/40' : ''
          } ${
            splitIssues.length > 0 
              ? 'bg-amber-50/70 border-amber-200/80 text-amber-900' 
              : 'bg-emerald-50/60 border-emerald-200/80 text-emerald-800'
          }`} 
          id="stats-split-periods-count"
        >
          <div className={`p-2.5 rounded-xl shrink-0 ${
            splitIssues.length > 0 ? 'bg-amber-100 text-amber-700' : 'bg-emerald-100 text-emerald-600'
          }`}>
            {splitIssues.length > 0 ? <Clock className="w-5 h-5" /> : <CheckCircle2 className="w-5 h-5" />}
          </div>
          <div className="flex-1">
            <div className="flex items-center justify-between">
              <div className="text-xl font-extrabold font-mono leading-none font-display">
                {splitIssues.length}
              </div>
              {activeTab === 'splitPeriods' && (
                <span className="text-[9px] px-1.5 py-0.5 bg-amber-200 text-amber-800 rounded font-bold uppercase">Đang xem</span>
              )}
            </div>
            <div className="text-[10px] opacity-75 font-bold uppercase tracking-wider mt-1.5">Tiết bị chia (Sáng/Chiều)</div>
          </div>
        </div>
      </div>

      {/* Mode Switcher Tabs */}
      <div className="flex items-center gap-2 p-1.5 bg-slate-200/60 rounded-2xl w-fit" id="report-view-tabs">
        <button
          onClick={() => setActiveTab('conflicts')}
          className={`flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
            activeTab === 'conflicts'
              ? 'bg-white text-slate-900 shadow-sm shadow-slate-200'
              : 'text-slate-600 hover:text-slate-900'
          }`}
          id="tab-btn-conflicts"
        >
          <AlertTriangle className={`w-4 h-4 ${conflicts.length > 0 ? 'text-rose-500' : 'text-slate-400'}`} />
          <span>Trùng lịch dạy cùng giờ</span>
          <span className={`px-2 py-0.5 rounded-full text-[11px] font-extrabold ${
            conflicts.length > 0 ? 'bg-rose-100 text-rose-700' : 'bg-slate-100 text-slate-500'
          }`}>
            {conflicts.length}
          </span>
        </button>

        <button
          onClick={() => setActiveTab('splitPeriods')}
          className={`flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
            activeTab === 'splitPeriods'
              ? 'bg-white text-slate-900 shadow-sm shadow-slate-200'
              : 'text-slate-600 hover:text-slate-900'
          }`}
          id="tab-btn-split-periods"
        >
          <Clock className={`w-4 h-4 ${splitIssues.length > 0 ? 'text-amber-500' : 'text-slate-400'}`} />
          <span>Tiết dạy bị chia trong ngày (Sáng / Chiều)</span>
          <span className={`px-2 py-0.5 rounded-full text-[11px] font-extrabold ${
            splitIssues.length > 0 ? 'bg-amber-100 text-amber-800' : 'bg-slate-100 text-slate-500'
          }`}>
            {splitIssues.length}
          </span>
        </button>
      </div>

      {/* TAB 1: SAME-SLOT CONFLICTS */}
      {activeTab === 'conflicts' && (
        <>
          {conflicts.length === 0 ? (
            <div className="p-8 bg-emerald-50/30 border border-emerald-100/60 rounded-2xl flex flex-col items-center text-center max-w-xl mx-auto my-6 shadow-sm" id="success-banner">
              <div className="p-4 bg-emerald-100 text-emerald-600 rounded-full mb-4 shadow-inner shadow-emerald-200/10">
                <CheckCircle2 className="w-10 h-10" />
              </div>
              <h3 className="text-base font-bold text-slate-900 font-display">
                Thời khóa biểu hoàn toàn không bị trùng lịch!
              </h3>
              <p className="text-xs text-slate-500 mt-2 max-w-sm leading-relaxed font-medium">
                Không phát hiện giáo viên nào phụ trách nhiều lớp cùng 1 thời điểm.
              </p>
            </div>
          ) : (
            <div className="bg-white rounded-2xl border border-slate-200/60 p-6 shadow-sm space-y-5" id="conflicts-section">
              {/* Controls Bar */}
              <div className="flex flex-col md:flex-row gap-3.5 items-stretch md:items-center justify-between">
                <div>
                  <h3 className="text-sm font-bold text-slate-800 flex items-center gap-2 font-display">
                    <AlertTriangle className="w-5 h-5 text-rose-500 shrink-0" />
                    Danh sách chi tiết cảnh báo trùng lịch ({filteredConflicts.length} / {conflicts.length})
                  </h3>
                  <p className="text-xs text-slate-400 font-medium mt-0.5">
                    Các tiết học có cùng 1 giáo viên phụ trách ở nhiều lớp khác nhau tại cùng một thời điểm.
                  </p>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <button
                    onClick={handleExportConflicts}
                    className="flex items-center gap-1.5 px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-xl transition-all cursor-pointer shadow-sm shadow-emerald-600/15"
                    id="export-conflicts-btn"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Xuất file Excel báo cáo
                  </button>
                </div>
              </div>

              {/* Filter Area */}
              <div className="flex flex-col md:flex-row gap-3 p-3 bg-slate-50/60 rounded-xl border border-slate-200/40">
                {/* Search */}
                <div className="relative flex-1">
                  <span className="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none text-slate-400">
                    <Search className="w-4 h-4" />
                  </span>
                  <input
                    type="text"
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    placeholder="Tìm theo Tên Giáo Viên, Lớp hoặc Tiết..."
                    className="w-full pl-9 pr-4 py-1.5 bg-white border border-slate-200/60 rounded-lg text-xs font-medium focus:outline-none focus:ring-2 focus:ring-teal-500/10 focus:border-teal-400 transition-all text-slate-700"
                    id="search-conflicts-input"
                  />
                </div>

                {/* Filter Day */}
                <div className="flex items-center gap-2 shrink-0">
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1">
                    <Filter className="w-3.5 h-3.5" />
                    Thứ:
                  </span>
                  <select
                    value={dayFilter}
                    onChange={(e) => setDayFilter(e.target.value)}
                    className="px-3 py-1.5 bg-white border border-slate-200/60 rounded-lg text-xs text-slate-600 font-bold focus:outline-none focus:ring-2 focus:ring-teal-500/10 focus:border-teal-400 cursor-pointer transition-all"
                    id="filter-day-select"
                  >
                    <option value="all">Tất cả các ngày</option>
                    {uniqueConflictDays.map(day => (
                      <option key={day} value={day}>{formatDay(day)}</option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Result Table */}
              <div className="overflow-hidden border border-slate-200/50 rounded-xl shadow-sm">
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm" id="conflicts-table">
                    <thead className="bg-slate-50/50 border-b border-slate-200/50 text-slate-400 font-extrabold uppercase text-[10px] tracking-wider">
                      <tr>
                        <th className="px-4 py-3 w-16 text-center">STT</th>
                        <th className="px-4 py-3 w-28">Thứ</th>
                        <th className="px-4 py-3 w-36">Tiết / Khung giờ</th>
                        <th className="px-4 py-3 w-48">Tên Giáo Viên bị trùng</th>
                        <th className="px-4 py-3">Các Lớp bị trùng</th>
                        <th className="px-4 py-3">Chi tiết các cặp trùng thực tế</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100 bg-white">
                      {filteredConflicts.length === 0 ? (
                        <tr id="filtered-conflicts-empty">
                          <td colSpan={6} className="px-5 py-8 text-center text-slate-400 italic bg-slate-50/20 font-medium">
                            Không tìm thấy cảnh báo nào khớp với bộ lọc hiện tại.
                          </td>
                        </tr>
                      ) : (
                        filteredConflicts.map((c, index) => (
                          <tr key={c.id} className="hover:bg-rose-50/10 transition-colors" id={`conflict-row-${c.id}`}>
                            {/* STT */}
                            <td className="px-4 py-3 text-center font-mono font-bold text-slate-400">
                              {index + 1}
                            </td>

                            {/* Thứ */}
                            <td className="px-4 py-3 font-extrabold text-slate-800">
                              {formatDay(c.day)}
                            </td>

                            {/* Tiết / Khung giờ */}
                            <td className="px-4 py-3">
                              <span className="px-2.5 py-1 bg-slate-100 text-slate-700 rounded-lg font-bold text-xs border border-slate-200/30">
                                {c.period || <span className="text-slate-300 italic">Chưa xác định</span>}
                              </span>
                            </td>

                            {/* Tên Giáo Viên bị trùng */}
                            <td className="px-4 py-3 font-extrabold text-rose-700">
                              {c.teacher}
                            </td>

                            {/* Các Lớp bị trùng */}
                            <td className="px-4 py-3">
                              <div className="flex flex-wrap items-center gap-1.5">
                                {c.classes.map((className, idx) => (
                                  <span 
                                    key={idx} 
                                    className={`px-2.5 py-1 rounded-lg text-[11px] font-bold shadow-sm transition-all border ${getClassBadgeStyles(className, c.classes)}`}
                                  >
                                    {formatClassName(className)}
                                  </span>
                                ))}
                              </div>
                            </td>

                            {/* Chi tiết các cặp trùng thực tế */}
                            <td className="px-4 py-3 font-semibold text-slate-500 text-xs">
                              <div className="flex flex-wrap items-center gap-1.5">
                                {c.conflictPairs.map((pair, idx) => {
                                  const isPairValid = isWithinValidPair(pair[0], pair[1]);
                                  return (
                                    <span 
                                      key={idx}
                                      className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg font-bold text-[11px] border transition-all ${getPairBadgeStyles(pair)}`}
                                    >
                                      {formatClassName(pair[0])}
                                      <ArrowRight className={`w-3.5 h-3.5 shrink-0 ${isPairValid ? 'text-rose-400' : 'text-purple-400'}`} />
                                      {formatClassName(pair[1])}
                                    </span>
                                  );
                                })}
                              </div>
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}
        </>
      )}

      {/* TAB 2: SPLIT PERIODS IN DAY (SÁNG / CHIỀU) */}
      {activeTab === 'splitPeriods' && (
        <div className="bg-white rounded-2xl border border-slate-200/60 p-6 shadow-sm space-y-5" id="split-periods-section">
          {/* Controls Bar */}
          <div className="flex flex-col md:flex-row gap-3.5 items-stretch md:items-center justify-between">
            <div>
              <h3 className="text-sm font-bold text-slate-800 flex items-center gap-2 font-display">
                <Clock className="w-5 h-5 text-amber-500 shrink-0" />
                Kiểm tra tiết bị chia / loãng xương ({filteredSplitIssues.length} trường hợp)
              </h3>
              <p className="text-xs text-slate-500 font-medium mt-0.5">
                Sáng (Tiết 1-5) và Chiều (Tiết 6-8) dạy liền nhau không tính bị chia. Chỉ cảnh báo khi dạy 1 tiết rồi không có tiết lớp đó sau đó lại có tiết lớp đó.
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <button
                onClick={handleExportSplitPeriods}
                disabled={splitIssues.length === 0}
                className={`flex items-center gap-1.5 px-4 py-2 font-bold text-xs rounded-xl transition-all cursor-pointer shadow-sm ${
                  splitIssues.length > 0 
                    ? 'bg-amber-600 hover:bg-amber-700 text-white shadow-amber-600/15'
                    : 'bg-slate-100 text-slate-400 cursor-not-allowed'
                }`}
                id="export-split-periods-btn"
              >
                <Download className="w-3.5 h-3.5" />
                Xuất file Excel tiết bị chia
              </button>
            </div>
          </div>

          {/* Quy tắc tính Tiết Bị Chia / Loãng Xương */}
          <div className="flex items-start gap-2.5 p-3 bg-amber-50/60 border border-amber-200/60 rounded-xl text-xs text-amber-900 leading-relaxed">
            <Sparkles className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
            <div>
              <span className="font-bold">Quy tắc chuẩn: </span>
              Buổi sáng gồm các Tiết 1, 2, 3, 4, 5 và buổi chiều gồm các Tiết 6, 7, 8. Nếu giáo viên dạy liên tiếp trong buổi sáng hoặc chiều thì <strong>không tính là bị chia</strong>. Chỉ ghi nhận vi phạm khi <strong>dạy 1 tiết rồi không có tiết ở lớp đó sau đó lại có tiết ở lớp đó</strong> (ví dụ sáng 1 tiết, chiều 1 tiết; hoặc sáng dạy Tiết 1 rồi Tiết 3 mới dạy).
            </div>
          </div>

          {/* Filter Area */}
          <div className="flex flex-col md:flex-row gap-3 p-3 bg-slate-50/60 rounded-xl border border-slate-200/40">
            {/* Search */}
            <div className="relative flex-1">
              <span className="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none text-slate-400">
                <Search className="w-4 h-4" />
              </span>
              <input
                type="text"
                value={splitSearchTerm}
                onChange={(e) => setSplitSearchTerm(e.target.value)}
                placeholder="Tìm theo Lớp, Giáo Viên, Môn học hoặc Tiết..."
                className="w-full pl-9 pr-4 py-1.5 bg-white border border-slate-200/60 rounded-lg text-xs font-medium focus:outline-none focus:ring-2 focus:ring-amber-500/10 focus:border-amber-400 transition-all text-slate-700"
                id="search-split-input"
              />
            </div>

            {/* Filter Split Type */}
            <div className="flex items-center gap-2 shrink-0">
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1">
                <Layers className="w-3.5 h-3.5" />
                Kiểu:
              </span>
              <select
                value={splitTypeFilter}
                onChange={(e) => setSplitTypeFilter(e.target.value as any)}
                className="px-3 py-1.5 bg-white border border-slate-200/60 rounded-lg text-xs text-slate-600 font-bold focus:outline-none focus:ring-2 focus:ring-amber-500/10 focus:border-amber-400 cursor-pointer transition-all"
                id="filter-split-type-select"
              >
                <option value="all">Tất cả kiểu chia / loãng xương</option>
                <option value="morning_afternoon">Chia 2 buổi (Sáng & Chiều)</option>
                <option value="isolated_periods">Loãng xương (cách quãng trong buổi)</option>
                <option value="both">Cả 2 buổi & Loãng xương</option>
              </select>
            </div>

            {/* Filter Day */}
            <div className="flex items-center gap-2 shrink-0">
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1">
                <Filter className="w-3.5 h-3.5" />
                Thứ:
              </span>
              <select
                value={splitDayFilter}
                onChange={(e) => setSplitDayFilter(e.target.value)}
                className="px-3 py-1.5 bg-white border border-slate-200/60 rounded-lg text-xs text-slate-600 font-bold focus:outline-none focus:ring-2 focus:ring-amber-500/10 focus:border-amber-400 cursor-pointer transition-all"
                id="filter-split-day-select"
              >
                <option value="all">Tất cả các ngày</option>
                {uniqueSplitDays.map(day => (
                  <option key={day} value={day}>{formatDay(day)}</option>
                ))}
              </select>
            </div>
          </div>

          {/* Split Periods Data Table */}
          <div className="overflow-hidden border border-slate-200/50 rounded-xl shadow-sm">
            <div className="overflow-x-auto">
              <table className="w-full text-left text-sm" id="split-periods-table">
                <thead className="bg-slate-50/50 border-b border-slate-200/50 text-slate-400 font-extrabold uppercase text-[10px] tracking-wider">
                  <tr>
                    <th className="px-4 py-3 w-16 text-center">STT</th>
                    <th className="px-4 py-3 w-28">Thứ</th>
                    <th className="px-4 py-3 w-36">Lớp học</th>
                    <th className="px-4 py-3 w-48">Giáo viên & Môn</th>
                    <th className="px-4 py-3 w-52">Phân loại vi phạm</th>
                    <th className="px-4 py-3">Chi tiết các tiết & Diễn giải khoảng trống</th>
                    <th className="px-4 py-3 w-28 text-center">Tổng số tiết</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 bg-white">
                  {filteredSplitIssues.length === 0 ? (
                    <tr id="filtered-split-empty">
                      <td colSpan={7} className="px-5 py-10 text-center text-slate-400 italic bg-slate-50/20 font-medium">
                        {splitIssues.length === 0 ? (
                          <div className="flex flex-col items-center justify-center gap-2 py-4">
                            <CheckCircle2 className="w-8 h-8 text-emerald-500" />
                            <span className="text-slate-700 font-bold text-sm not-italic">
                              Không có lớp nào bị chia tiết hoặc loãng xương trong ngày!
                            </span>
                            <span className="text-xs text-slate-400 max-w-md not-italic font-normal">
                              Tất cả giáo viên đều được bố trí dạy tiết liền mạch (không bị chia sáng 1 tiết, chiều 1 tiết và không có tiết trống ngắt quãng giữa giờ).
                            </span>
                          </div>
                        ) : (
                          'Không tìm thấy trường hợp nào khớp với bộ lọc hiện tại.'
                        )}
                      </td>
                    </tr>
                  ) : (
                    filteredSplitIssues.map((issue, index) => {
                      const isBoth = issue.splitType === 'both';
                      const isMorningAfternoon = issue.splitType === 'morning_afternoon';
                      const isIsolated = issue.splitType === 'isolated_periods';

                      return (
                        <tr key={issue.id} className="hover:bg-amber-50/15 transition-colors" id={`split-row-${issue.id}`}>
                          {/* STT */}
                          <td className="px-4 py-3 text-center font-mono font-bold text-slate-400">
                            {index + 1}
                          </td>

                          {/* Thứ */}
                          <td className="px-4 py-3 font-extrabold text-slate-800">
                            {formatDay(issue.day)}
                          </td>

                          {/* Lớp học */}
                          <td className="px-4 py-3">
                            <span className="px-2.5 py-1 bg-teal-50 text-teal-800 border border-teal-200/80 rounded-lg text-xs font-extrabold shadow-sm">
                              {formatClassName(issue.className)}
                            </span>
                          </td>

                          {/* Giáo viên & Môn */}
                          <td className="px-4 py-3">
                            <div className="font-extrabold text-slate-900 leading-snug">
                              {issue.teacher}
                            </div>
                            {issue.subject && (
                              <div className="text-[11px] font-bold text-slate-400 mt-0.5">
                                Môn: <span className="text-slate-600">{issue.subject}</span>
                              </div>
                            )}
                          </td>

                          {/* Phân loại bị chia */}
                          <td className="px-4 py-3">
                            {isBoth ? (
                              <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[11px] font-bold bg-rose-50 text-rose-800 border border-rose-200 shadow-sm">
                                <AlertTriangle className="w-3.5 h-3.5 text-rose-500 shrink-0" />
                                <span>Cả 2 buổi & Loãng xương</span>
                              </span>
                            ) : isMorningAfternoon ? (
                              <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[11px] font-bold bg-amber-50 text-amber-800 border border-amber-200/80 shadow-sm">
                                <Sun className="w-3.5 h-3.5 text-amber-500 shrink-0" />
                                <span>Chia 2 buổi (Sáng & Chiều)</span>
                                <Moon className="w-3.5 h-3.5 text-indigo-500 shrink-0" />
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[11px] font-bold bg-purple-50 text-purple-800 border border-purple-200 shadow-sm">
                                <Clock className="w-3.5 h-3.5 text-purple-500 shrink-0" />
                                <span>Loãng xương trong buổi</span>
                              </span>
                            )}
                          </td>

                          {/* Chi tiết các tiết trong ngày */}
                          <td className="px-4 py-3">
                            <div className="space-y-1.5 text-xs">
                              {/* Morning slots */}
                              {issue.morningPeriods.length > 0 && (
                                <div className="flex items-center gap-2">
                                  <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-amber-100/70 text-amber-800 font-bold text-[10px] shrink-0 border border-amber-200/60">
                                    <Sun className="w-3 h-3 text-amber-600" />
                                    Sáng ({issue.morningPeriods.length})
                                  </span>
                                  <div className="flex flex-wrap gap-1">
                                    {issue.morningPeriods.map((p, idx) => (
                                      <span key={idx} className="px-2 py-0.5 bg-slate-100 text-slate-700 rounded text-[11px] font-semibold border border-slate-200/60">
                                        {p}
                                      </span>
                                    ))}
                                  </div>
                                </div>
                              )}

                              {/* Afternoon slots */}
                              {issue.afternoonPeriods.length > 0 && (
                                <div className="flex items-center gap-2">
                                  <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-indigo-100/70 text-indigo-800 font-bold text-[10px] shrink-0 border border-indigo-200/60">
                                    <Moon className="w-3 h-3 text-indigo-600" />
                                    Chiều ({issue.afternoonPeriods.length})
                                  </span>
                                  <div className="flex flex-wrap gap-1">
                                    {issue.afternoonPeriods.map((p, idx) => (
                                      <span key={idx} className="px-2 py-0.5 bg-slate-100 text-slate-700 rounded text-[11px] font-semibold border border-slate-200/60">
                                        {p}
                                      </span>
                                    ))}
                                  </div>
                                </div>
                              )}

                              {/* Description of gap / split */}
                              {issue.description && (
                                <div className="text-[11px] font-medium text-slate-500 bg-slate-50/80 px-2 py-1 rounded border border-slate-200/40">
                                  {issue.description}
                                </div>
                              )}
                            </div>
                          </td>

                          {/* Tổng số tiết */}
                          <td className="px-4 py-3 text-center">
                            <span className="px-2.5 py-1 bg-slate-100 text-slate-800 rounded-lg text-xs font-extrabold border border-slate-200/80 font-mono">
                              {issue.totalPeriods} tiết
                            </span>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
