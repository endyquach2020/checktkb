import { useState, useMemo } from 'react';
import { TeacherConflict, ExemptionPair } from '../types';
import { 
  CheckCircle2, 
  AlertTriangle, 
  Search, 
  Calendar, 
  Clock, 
  User, 
  Users, 
  Filter, 
  Download,
  BookOpen,
  ArrowRight
} from 'lucide-react';
import * as XLSX from 'xlsx';

// Helper to format days to match the user's specific format (e.g., "Thứ Hai" -> "Thứ 2")
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
  exemptions: ExemptionPair[];
  totalCellsChecked: number;
  totalTeachersFound: number;
}

export default function ConflictReport({
  conflicts,
  exemptions,
  totalCellsChecked,
  totalTeachersFound
}: ConflictReportProps) {
  const [searchTerm, setSearchTerm] = useState('');
  const [dayFilter, setDayFilter] = useState('all');

  // Extract unique days for the filter dropdown
  const uniqueDays = useMemo(() => {
    const days = new Set<string>();
    conflicts.forEach(c => {
      if (c.day) days.add(c.day);
    });
    return Array.from(days);
  }, [conflicts]);

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

  // Export conflict report to Excel
  const handleExport = () => {
    if (conflicts.length === 0) return;

    const wb = XLSX.utils.book_new();

    // Prepare data EXACTLY as shown in the user's reference image
    const exportData = filteredConflicts.map((c, index) => ({
      'STT': index + 1,
      'Thứ': formatDay(c.day),
      'Tiết / Khung giờ': c.period || 'Chưa xác định',
      'Tên Giáo Viên bị trùng': c.teacher,
      'Các Lớp bị trùng': c.classes.map(formatClassName).join(', '),
      'Chi tiết các cặp trùng thực tế': c.conflictPairs.map(p => `${formatClassName(p[0])} vs ${formatClassName(p[1])}`).join('; ')
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    
    // Header style helpers & column width setup to match the image beautifully
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

  const activeExemptionsCount = exemptions.filter(e => e.isActive).length;

  return (
    <div id="conflict-report-container" className="space-y-6">
      
      {/* Summary Cards */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <div className="bg-white rounded-2xl border border-slate-200/50 p-4 shadow-sm flex items-center gap-3.5 transition-all hover:shadow-md/5" id="stats-total-checked">
          <div className="p-2.5 bg-teal-50 text-teal-600 rounded-xl shrink-0">
            <BookOpen className="w-5 h-5" />
          </div>
          <div>
            <div className="text-xl font-extrabold text-slate-800 font-mono leading-none font-display">{totalCellsChecked}</div>
            <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider mt-1.5">Ô đã rà soát</div>
          </div>
        </div>

        <div className="bg-white rounded-2xl border border-slate-200/50 p-4 shadow-sm flex items-center gap-3.5 transition-all hover:shadow-md/5" id="stats-teachers-found">
          <div className="p-2.5 bg-emerald-50 text-emerald-600 rounded-xl shrink-0">
            <User className="w-5 h-5" />
          </div>
          <div>
            <div className="text-xl font-extrabold text-slate-800 font-mono leading-none font-display">{totalTeachersFound}</div>
            <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider mt-1.5">Số Giáo viên</div>
          </div>
        </div>

        <div className="bg-white rounded-2xl border border-slate-200/50 p-4 shadow-sm flex items-center gap-3.5 transition-all hover:shadow-md/5" id="stats-exemptions">
          <div className="p-2.5 bg-teal-50 text-teal-600 rounded-xl shrink-0">
            <Users className="w-5 h-5" />
          </div>
          <div>
            <div className="text-xl font-extrabold text-slate-800 font-mono leading-none font-display">{activeExemptionsCount}</div>
            <div className="text-[10px] text-slate-400 font-bold uppercase tracking-wider mt-1.5">Miễn trùng active</div>
          </div>
        </div>

        <div className={`rounded-2xl border p-4 shadow-sm flex items-center gap-3.5 transition-all hover:shadow-md/5 ${
          conflicts.length > 0 
            ? 'bg-rose-50/50 border-rose-100 text-rose-800' 
            : 'bg-emerald-50/50 border-emerald-100 text-emerald-800'
        }`} id="stats-conflict-count">
          <div className={`p-2.5 rounded-xl shrink-0 ${
            conflicts.length > 0 ? 'bg-rose-100 text-rose-600' : 'bg-emerald-100 text-emerald-600'
          }`}>
            {conflicts.length > 0 ? <AlertTriangle className="w-5 h-5" /> : <CheckCircle2 className="w-5 h-5" />}
          </div>
          <div>
            <div className="text-xl font-extrabold font-mono leading-none font-display">
              {conflicts.length}
            </div>
            <div className="text-[10px] opacity-75 font-bold uppercase tracking-wider mt-1.5">Số Lỗi trùng tiết</div>
          </div>
        </div>
      </div>

      {/* Main Validation Result Banner */}
      {conflicts.length === 0 ? (
        <div className="p-8 bg-emerald-50/30 border border-emerald-100/60 rounded-2xl flex flex-col items-center text-center max-w-xl mx-auto my-6 shadow-sm" id="success-banner">
          <div className="p-4 bg-emerald-100 text-emerald-600 rounded-full mb-4 shadow-inner shadow-emerald-200/10">
            <CheckCircle2 className="w-10 h-10" />
          </div>
          <h3 className="text-base font-bold text-slate-900 font-display">
            Thời khóa biểu hoàn toàn hợp lệ!
          </h3>
          <p className="text-xs text-slate-500 mt-2 max-w-sm leading-relaxed font-medium">
            Không phát hiện bất kỳ trùng lặp lịch dạy nào của giáo viên trên hệ thống. Tất cả các giáo viên đều được sắp xếp tiết dạy hợp lý.
          </p>
        </div>
      ) : (
        <div className="bg-white rounded-2xl border border-slate-200/50 p-6 shadow-sm space-y-5">
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
                onClick={handleExport}
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
                {uniqueDays.map(day => (
                  <option key={day} value={day}>{formatDay(day)}</option>
                ))}
              </select>
            </div>
          </div>

          {/* Result Table in exactly the image structure */}
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
    </div>
  );
}
