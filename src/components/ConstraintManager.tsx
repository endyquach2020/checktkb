import React, { useState, useMemo } from 'react';
import { 
  TimetableConstraint, 
  ConstraintCheckResult, 
  ConstraintType, 
  ConstraintViolation,
  DaySession 
} from '../types';
import { 
  CheckCircle2, 
  AlertTriangle, 
  Plus, 
  Trash2, 
  Edit3, 
  Copy, 
  RotateCcw, 
  Download, 
  ChevronDown, 
  ChevronUp, 
  ShieldCheck, 
  Sparkles, 
  Sliders, 
  Check, 
  X,
  Search,
  Filter,
  FileSpreadsheet,
  Upload,
  ArrowRight,
  Clock
} from 'lucide-react';
import * as XLSX from 'xlsx';

interface ConstraintManagerProps {
  constraints: TimetableConstraint[];
  checkResults?: ConstraintCheckResult[];
  availableTeachers?: string[];
  availableClasses?: string[];
  onSaveConstraint: (constraint: TimetableConstraint) => void;
  onDeleteConstraint: (id: string) => void;
  onToggleConstraint: (id: string, isActive: boolean) => void;
  onResetDefaults: () => void;
  isPreUpload?: boolean;
  onProceedToUpload?: () => void;
}

const CONSTRAINT_TYPE_OPTIONS: { type: ConstraintType; label: string; desc: string }[] = [
  {
    type: 'simultaneous_classes_teachers',
    label: 'Trùng giờ 2 Lớp & 2 Giáo viên',
    desc: 'Không xếp trùng tiết giữa 2 lớp cụ thể khi dạy 2 giáo viên cụ thể (hoặc bắt buộc cùng giờ)'
  },
  {
    type: 'teacher_unavailable',
    label: 'Giáo viên bận / Nghỉ dạy',
    desc: 'Giáo viên không được xếp dạy vào các Thứ / Buổi / Tiết đăng ký bận hoặc nghỉ'
  },
  {
    type: 'teacher_max_periods_day',
    label: 'Số tiết tối đa / ngày của GV',
    desc: 'Khống chế giáo viên không dạy quá một số tiết nhất định trong một ngày'
  },
  {
    type: 'no_split_day',
    label: 'Không chia môn trong ngày (Sáng & Chiều)',
    desc: 'Tránh việc cùng môn học của giáo viên bị xé lẻ dạy ở cả sáng và chiều'
  },
  {
    type: 'subject_period_restriction',
    label: 'Ràng buộc tiết cho Môn học',
    desc: 'Môn học không được xếp vào Thứ / Buổi / Tiết nhất định (VD: Thể dục không xếp tiết 5)'
  },
  {
    type: 'class_pair_no_overlap',
    label: 'Hai lớp không được trùng giờ',
    desc: 'Hai lớp dùng chung phòng chức năng, phòng máy tính hoặc sân bãi không được trùng tiết'
  },
  {
    type: 'teacher_no_gap',
    label: 'Không có tiết trống cách quãng (loãng xương)',
    desc: 'Hạn chế việc giáo viên dạy tiết đầu và tiết cuối của buổi mà trống tiết ở giữa'
  }
];

export const ConstraintManager: React.FC<ConstraintManagerProps> = ({
  constraints,
  checkResults = [],
  availableTeachers = [],
  availableClasses = [],
  onSaveConstraint,
  onDeleteConstraint,
  onToggleConstraint,
  onResetDefaults,
  isPreUpload = false,
  onProceedToUpload
}) => {
  const [filterTab, setFilterTab] = useState<'all' | 'violated' | 'satisfied' | 'active'>('all');
  const [searchQuery, setSearchQuery] = useState('');
  const [expandedConstraintIds, setExpandedConstraintIds] = useState<Set<string>>(new Set());
  const [editingConstraint, setEditingConstraint] = useState<TimetableConstraint | null>(null);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [copySuccessId, setCopySuccessId] = useState<string | null>(null);

  // Form State for Add / Edit Modal
  const [formName, setFormName] = useState('');
  const [formDesc, setFormDesc] = useState('');
  const [formType, setFormType] = useState<ConstraintType>('simultaneous_classes_teachers');
  const [formIsActive, setFormIsActive] = useState(true);

  // Params state
  const [formClass1, setFormClass1] = useState('');
  const [formClass2, setFormClass2] = useState('');
  const [formTeacher1, setFormTeacher1] = useState('');
  const [formTeacher2, setFormTeacher2] = useState('');
  const [formMode, setFormMode] = useState<'must_not_overlap' | 'must_overlap'>('must_not_overlap');
  
  const [formTeacher, setFormTeacher] = useState('');
  const [formDays, setFormDays] = useState<string[]>([]);
  const [formSessions, setFormSessions] = useState<DaySession[]>([]);
  const [formPeriods, setFormPeriods] = useState<string[]>([]);
  const [formMaxPeriods, setFormMaxPeriods] = useState<number>(5);

  const [formSubject, setFormSubject] = useState('');
  const [formOverlapClass1, setFormOverlapClass1] = useState('');
  const [formOverlapClass2, setFormOverlapClass2] = useState('');

  // Results lookup by constraint ID
  const resultMap = useMemo(() => {
    const map = new Map<string, ConstraintCheckResult>();
    for (const r of checkResults) {
      map.set(r.constraintId, r);
    }
    return map;
  }, [checkResults]);

  const hasResults = checkResults && checkResults.length > 0;

  // Overall Stats
  const stats = useMemo(() => {
    const activeConstraints = constraints.filter(c => c.isActive);
    let satisfied = 0;
    let violated = 0;
    let totalViolations = 0;

    for (const c of activeConstraints) {
      const res = resultMap.get(c.id);
      if (res) {
        if (res.isSatisfied) {
          satisfied++;
        } else {
          violated++;
          totalViolations += res.violationsCount;
        }
      }
    }

    return {
      total: constraints.length,
      active: activeConstraints.length,
      satisfied,
      violated,
      totalViolations
    };
  }, [constraints, resultMap]);

  // Filtered constraints
  const filteredConstraints = useMemo(() => {
    return constraints.filter(c => {
      const res = resultMap.get(c.id);
      const isSatisfied = res?.isSatisfied ?? true;

      // Filter by tab
      if (filterTab === 'active' && !c.isActive) return false;
      if (filterTab === 'violated' && (!c.isActive || !hasResults || isSatisfied)) return false;
      if (filterTab === 'satisfied' && (!c.isActive || !hasResults || !isSatisfied)) return false;

      // Filter by search query
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const matchName = c.name.toLowerCase().includes(q);
        const matchDesc = (c.description || '').toLowerCase().includes(q);
        const matchParams = JSON.stringify(c.params).toLowerCase().includes(q);
        if (!matchName && !matchDesc && !matchParams) return false;
      }

      return true;
    });
  }, [constraints, resultMap, filterTab, searchQuery, hasResults]);

  const toggleExpand = (id: string) => {
    setExpandedConstraintIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const openAddModal = () => {
    setEditingConstraint(null);
    setFormName('');
    setFormDesc('');
    setFormType('simultaneous_classes_teachers');
    setFormIsActive(true);
    setFormClass1('10.2');
    setFormClass2('12.2');
    setFormTeacher1('Thầy Mịnh');
    setFormTeacher2('Thầy Đ. Minh');
    setFormMode('must_not_overlap');
    setFormTeacher('');
    setFormDays([]);
    setFormSessions([]);
    setFormPeriods([]);
    setFormMaxPeriods(5);
    setFormSubject('');
    setFormOverlapClass1('');
    setFormOverlapClass2('');
    setIsModalOpen(true);
  };

  const openEditModal = (constraint: TimetableConstraint) => {
    setEditingConstraint(constraint);
    setFormName(constraint.name);
    setFormDesc(constraint.description || '');
    setFormType(constraint.type);
    setFormIsActive(constraint.isActive);
    setFormClass1(constraint.params.class1 || '');
    setFormClass2(constraint.params.class2 || '');
    setFormTeacher1(constraint.params.teacher1 || '');
    setFormTeacher2(constraint.params.teacher2 || '');
    setFormMode(constraint.params.mode || 'must_not_overlap');
    setFormTeacher(constraint.params.teacher || '');
    setFormDays(constraint.params.days || []);
    setFormSessions(constraint.params.sessions || []);
    setFormPeriods(constraint.params.periods || []);
    setFormMaxPeriods(constraint.params.maxPeriodsPerDay || 5);
    setFormSubject(constraint.params.subject || '');
    setFormOverlapClass1(constraint.params.overlapClass1 || '');
    setFormOverlapClass2(constraint.params.overlapClass2 || '');
    setIsModalOpen(true);
  };

  const handleDuplicate = (constraint: TimetableConstraint) => {
    const newConstraint: TimetableConstraint = {
      ...constraint,
      id: 'c-' + Date.now(),
      name: `${constraint.name} (Bản sao)`,
      isActive: true
    };
    onSaveConstraint(newConstraint);
  };

  const handleSaveForm = (e: React.FormEvent) => {
    e.preventDefault();
    const finalName = formName.trim() || getDefaultNameForType(formType);

    const newConstraint: TimetableConstraint = {
      id: editingConstraint ? editingConstraint.id : 'c-' + Date.now(),
      name: finalName,
      description: formDesc.trim() || undefined,
      type: formType,
      isActive: formIsActive,
      params: {}
    };

    if (formType === 'simultaneous_classes_teachers') {
      newConstraint.params = {
        class1: formClass1.trim() || '10.2',
        class2: formClass2.trim() || '12.2',
        teacher1: formTeacher1.trim() || 'Thầy Mịnh',
        teacher2: formTeacher2.trim() || 'Thầy Đ. Minh',
        mode: formMode
      };
    } else if (formType === 'teacher_unavailable') {
      newConstraint.params = {
        teacher: formTeacher.trim(),
        days: formDays,
        sessions: formSessions,
        periods: formPeriods
      };
    } else if (formType === 'teacher_max_periods_day') {
      newConstraint.params = {
        teacher: formTeacher.trim() || undefined,
        maxPeriodsPerDay: Number(formMaxPeriods) || 5
      };
    } else if (formType === 'no_split_day') {
      newConstraint.params = {
        targetClass: formClass1.trim() || undefined,
        targetTeacher: formTeacher.trim() || undefined
      };
    } else if (formType === 'subject_period_restriction') {
      newConstraint.params = {
        subject: formSubject.trim(),
        restrictedDays: formDays,
        restrictedSessions: formSessions,
        restrictedPeriods: formPeriods
      };
    } else if (formType === 'class_pair_no_overlap') {
      newConstraint.params = {
        overlapClass1: formOverlapClass1.trim(),
        overlapClass2: formOverlapClass2.trim()
      };
    } else if (formType === 'teacher_no_gap') {
      newConstraint.params = {
        teacher: formTeacher.trim() || undefined
      };
    }

    onSaveConstraint(newConstraint);
    setIsModalOpen(false);
  };

  const getDefaultNameForType = (type: ConstraintType): string => {
    switch (type) {
      case 'simultaneous_classes_teachers':
        return `Ràng buộc trùng giờ 2 lớp (${formClass1 || '10.2'} & ${formClass2 || '12.2'})`;
      case 'teacher_unavailable':
        return `GV ${formTeacher || 'bận'} nghỉ dạy`;
      case 'teacher_max_periods_day':
        return `Tối đa ${formMaxPeriods || 5} tiết/ngày`;
      case 'no_split_day':
        return 'Không chia môn trong ngày (Sáng & Chiều)';
      case 'subject_period_restriction':
        return `Hạn chế giờ dạy môn ${formSubject || 'học'}`;
      case 'class_pair_no_overlap':
        return `Không trùng giờ ${formOverlapClass1 || 'Lớp 1'} & ${formOverlapClass2 || 'Lớp 2'}`;
      case 'teacher_no_gap':
        return 'Hạn chế tiết trống cách quãng của GV';
      default:
        return 'Ràng buộc thời khóa biểu';
    }
  };

  // Export all violations of active constraints to Excel
  const handleExportAllViolations = () => {
    const wb = XLSX.utils.book_new();
    const rows: any[] = [];

    constraints.forEach(c => {
      const res = resultMap.get(c.id);
      if (res && res.violations.length > 0) {
        res.violations.forEach((v, vIdx) => {
          rows.push({
            'Tên Ràng Buộc': c.name,
            'Loại Ràng Buộc': c.type,
            'STT': vIdx + 1,
            'Thứ': v.day,
            'Tiết / Khung giờ': v.period,
            'Buổi': v.session || 'Khác',
            'Lớp học': v.className || '',
            'Giáo viên': v.teacher || '',
            'Môn học': v.subject || '',
            'Dòng trong Excel': v.rowIndex ? `Dòng ${v.rowIndex}` : '',
            'Nội dung vi phạm': v.message
          });
        });
      }
    });

    if (rows.length === 0) {
      alert('Hiện không có vi phạm nào để xuất Excel. Tất cả các ràng buộc đang kiểm tra đều Thỏa mãn (Đạt)!');
      return;
    }

    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [
      { wch: 35 },
      { wch: 25 },
      { wch: 6 },
      { wch: 15 },
      { wch: 25 },
      { wch: 12 },
      { wch: 20 },
      { wch: 25 },
      { wch: 18 },
      { wch: 16 },
      { wch: 65 }
    ];

    XLSX.utils.book_append_sheet(wb, ws, 'Vi_Pham_Rang_Buoc');
    XLSX.writeFile(wb, 'danh_sach_vi_pham_rang_buoc_tkb.xlsx');
  };

  const copyViolationText = (constraint: TimetableConstraint, violations: ConstraintViolation[]) => {
    const text = violations.map((v, i) => `${i + 1}. [${v.day} - ${v.period}] ${v.message}`).join('\n');
    navigator.clipboard.writeText(text);
    setCopySuccessId(constraint.id);
    setTimeout(() => setCopySuccessId(null), 2000);
  };

  return (
    <div className="space-y-6" id="constraint-manager-container">
      {/* Top Banner & Dashboard Summary */}
      <div className="bg-gradient-to-br from-indigo-900 via-slate-900 to-indigo-950 rounded-2xl p-5 sm:p-6 text-white shadow-lg border border-indigo-500/20">
        <div className="flex flex-col lg:flex-row items-start lg:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2.5">
              <div className="p-2 bg-indigo-500/20 rounded-xl border border-indigo-400/30 text-indigo-300">
                <ShieldCheck className="w-6 h-6" />
              </div>
              <div>
                <h2 className="text-lg sm:text-xl font-extrabold tracking-tight font-display flex items-center gap-2">
                  Trung Tâm Quản Lý & Rà Soát Ràng Buộc Tự Động
                  <span className="text-[11px] font-mono px-2 py-0.5 bg-indigo-400/20 text-indigo-200 rounded-full border border-indigo-400/30">
                    Tự động lưu & kiểm tra
                  </span>
                </h2>
                <p className="text-xs sm:text-sm text-slate-300 mt-1">
                  Nhập và lưu trữ các quy tắc phân công. Hệ thống tự động đối chiếu với thời khóa biểu để chỉ rõ các vị trí vi phạm giúp bạn xếp lại lịch dễ dàng.
                </p>
              </div>
            </div>
          </div>

          {/* Quick Actions */}
          <div className="flex flex-wrap items-center gap-2.5 shrink-0 w-full lg:w-auto">
            <button
              onClick={openAddModal}
              id="btn-add-new-constraint"
              className="px-4 py-2.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-bold transition-all shadow-md hover:shadow-indigo-500/30 flex items-center gap-2 cursor-pointer"
            >
              <Plus className="w-4 h-4" />
              Thêm Ràng Buộc Mới
            </button>

            {stats.totalViolations > 0 && (
              <button
                onClick={handleExportAllViolations}
                id="btn-export-violations-excel"
                className="px-3.5 py-2.5 bg-rose-500/20 hover:bg-rose-500/30 text-rose-200 border border-rose-400/40 rounded-xl text-xs font-bold transition-all flex items-center gap-2 cursor-pointer"
              >
                <Download className="w-4 h-4 text-rose-300" />
                Xuất Excel Vi Phạm ({stats.totalViolations})
              </button>
            )}

            <button
              onClick={onResetDefaults}
              title="Khôi phục danh sách ràng buộc chuẩn"
              className="px-3 py-2.5 bg-white/10 hover:bg-white/15 text-slate-200 rounded-xl text-xs font-semibold transition-all flex items-center gap-1.5 cursor-pointer"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              Khôi phục mẫu
            </button>
          </div>
        </div>

        {/* Status Metrics Bar */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-5 pt-5 border-t border-white/10">
          <div className="bg-white/5 backdrop-blur-sm rounded-xl p-3 border border-white/5">
            <div className="text-xl sm:text-2xl font-black font-mono text-white">{stats.active} / {stats.total}</div>
            <div className="text-[11px] text-slate-400 font-bold uppercase tracking-wider mt-0.5">Ràng buộc đang bật</div>
          </div>

          {!hasResults ? (
            <>
              <div className="bg-indigo-500/10 backdrop-blur-sm rounded-xl p-3 border border-indigo-500/20 col-span-2">
                <div className="text-sm font-bold text-indigo-200 flex items-center gap-2">
                  <Clock className="w-4 h-4 text-indigo-300 shrink-0" />
                  Sẵn sàng áp dụng khi tải file Excel
                </div>
                <div className="text-[11px] text-slate-400 mt-0.5">
                  Nhập xong các yêu cầu, bấm "Tải File TKB Lên" để hệ thống tự động quét kiểm tra ngay.
                </div>
              </div>

              {onProceedToUpload && (
                <button
                  onClick={onProceedToUpload}
                  className="bg-teal-500 hover:bg-teal-400 text-slate-950 font-black text-xs rounded-xl p-3 shadow-md flex items-center justify-center gap-1.5 transition-all cursor-pointer"
                >
                  <span>Tải File TKB Lên</span>
                  <ArrowRight className="w-4 h-4" />
                </button>
              )}
            </>
          ) : (
            <>
              <div className="bg-emerald-500/10 backdrop-blur-sm rounded-xl p-3 border border-emerald-500/20">
                <div className="text-xl sm:text-2xl font-black font-mono text-emerald-400 flex items-center gap-1.5">
                  <CheckCircle2 className="w-5 h-5" />
                  {stats.satisfied}
                </div>
                <div className="text-[11px] text-emerald-300 font-bold uppercase tracking-wider mt-0.5">Đạt (Thỏa mãn)</div>
              </div>

              <div className="bg-rose-500/10 backdrop-blur-sm rounded-xl p-3 border border-rose-500/20">
                <div className="text-xl sm:text-2xl font-black font-mono text-rose-400 flex items-center gap-1.5">
                  <AlertTriangle className="w-5 h-5" />
                  {stats.violated}
                </div>
                <div className="text-[11px] text-rose-300 font-bold uppercase tracking-wider mt-0.5">Vi phạm (Cần xếp lại)</div>
              </div>

              <div className="bg-amber-500/10 backdrop-blur-sm rounded-xl p-3 border border-amber-500/20">
                <div className="text-xl sm:text-2xl font-black font-mono text-amber-300">{stats.totalViolations}</div>
                <div className="text-[11px] text-amber-300/80 font-bold uppercase tracking-wider mt-0.5">Tổng số tiết vi phạm</div>
              </div>
            </>
          )}
        </div>
      </div>

      {/* Filter and Search Toolbar */}
      <div className="bg-white rounded-2xl p-4 border border-slate-200/80 shadow-sm flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
        {/* Filter Tabs */}
        <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-xl text-xs font-bold overflow-x-auto">
          <button
            onClick={() => setFilterTab('all')}
            className={`px-3 py-1.5 rounded-lg transition-all whitespace-nowrap cursor-pointer ${
              filterTab === 'all' ? 'bg-white text-slate-800 shadow-xs' : 'text-slate-500 hover:text-slate-700'
            }`}
          >
            Tất cả ({constraints.length})
          </button>
          {hasResults && (
            <>
              <button
                onClick={() => setFilterTab('violated')}
                className={`px-3 py-1.5 rounded-lg transition-all whitespace-nowrap cursor-pointer flex items-center gap-1.5 ${
                  filterTab === 'violated' ? 'bg-rose-600 text-white shadow-xs' : 'text-slate-500 hover:text-rose-600'
                }`}
              >
                <AlertTriangle className="w-3.5 h-3.5" />
                Vi phạm ({stats.violated})
              </button>
              <button
                onClick={() => setFilterTab('satisfied')}
                className={`px-3 py-1.5 rounded-lg transition-all whitespace-nowrap cursor-pointer flex items-center gap-1.5 ${
                  filterTab === 'satisfied' ? 'bg-emerald-600 text-white shadow-xs' : 'text-slate-500 hover:text-emerald-600'
                }`}
              >
                <CheckCircle2 className="w-3.5 h-3.5" />
                Đạt ({stats.satisfied})
              </button>
            </>
          )}
          <button
            onClick={() => setFilterTab('active')}
            className={`px-3 py-1.5 rounded-lg transition-all whitespace-nowrap cursor-pointer ${
              filterTab === 'active' ? 'bg-white text-slate-800 shadow-xs' : 'text-slate-500 hover:text-slate-700'
            }`}
          >
            Đang bật ({stats.active})
          </button>
        </div>

        {/* Search */}
        <div className="relative w-full md:w-72">
          <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            placeholder="Tìm theo tên ràng buộc, lớp, GV..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium text-slate-800 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 focus:bg-white"
          />
        </div>
      </div>

      {/* Constraints List */}
      <div className="space-y-4">
        {filteredConstraints.length === 0 ? (
          <div className="bg-white rounded-2xl border border-dashed border-slate-300 p-8 text-center">
            <ShieldCheck className="w-10 h-10 text-slate-400 mx-auto mb-3" />
            <div className="text-sm font-bold text-slate-700">Không tìm thấy ràng buộc nào phù hợp</div>
            <p className="text-xs text-slate-400 mt-1">Hãy thử thay đổi bộ lọc hoặc thêm ràng buộc mới.</p>
            <button
              onClick={openAddModal}
              className="mt-4 px-4 py-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded-xl text-xs font-bold transition-all inline-flex items-center gap-2 cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" />
              Thêm Ràng Buộc Mới
            </button>
          </div>
        ) : (
          filteredConstraints.map((constraint) => {
            const res = resultMap.get(constraint.id);
            const isSatisfied = res?.isSatisfied ?? true;
            const violations = res?.violations || [];
            const isExpanded = expandedConstraintIds.has(constraint.id);

            return (
              <div
                key={constraint.id}
                id={`constraint-card-${constraint.id}`}
                className={`bg-white rounded-2xl border transition-all shadow-xs overflow-hidden ${
                  !constraint.isActive 
                    ? 'opacity-60 border-slate-200 bg-slate-50/50'
                    : !isSatisfied
                    ? 'border-rose-300 ring-2 ring-rose-400/10'
                    : 'border-slate-200/80 hover:border-emerald-200'
                }`}
              >
                {/* Constraint Header */}
                <div className="p-4 sm:p-5 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
                  <div className="flex items-start gap-3.5 flex-1">
                    {/* Toggle Switch */}
                    <button
                      type="button"
                      role="switch"
                      aria-checked={constraint.isActive}
                      onClick={() => onToggleConstraint(constraint.id, !constraint.isActive)}
                      className={`relative inline-flex h-6 w-11 shrink-0 cursor-pointer rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out focus:outline-none mt-0.5 ${
                        constraint.isActive ? 'bg-indigo-600' : 'bg-slate-200'
                      }`}
                      title={constraint.isActive ? 'Đang bật (Nhấn để tắt)' : 'Đang tắt (Nhấn để bật)'}
                    >
                      <span
                        className={`pointer-events-none inline-block h-5 w-5 transform rounded-full bg-white shadow-sm ring-0 transition duration-200 ease-in-out ${
                          constraint.isActive ? 'translate-x-5' : 'translate-x-0'
                        }`}
                      />
                    </button>

                    <div className="flex-1">
                      <div className="flex flex-wrap items-center gap-2">
                        <h3 className="text-sm sm:text-base font-extrabold text-slate-800 font-display">
                          {constraint.name}
                        </h3>

                        {/* Status Badge */}
                        {!constraint.isActive ? (
                          <span className="px-2.5 py-0.5 bg-slate-100 text-slate-500 rounded-full text-[11px] font-bold">
                            Tạm tắt
                          </span>
                        ) : !hasResults ? (
                          <span className="px-2.5 py-0.5 bg-indigo-50 text-indigo-700 border border-indigo-200 rounded-full text-[11px] font-bold flex items-center gap-1">
                            <Clock className="w-3.5 h-3.5 text-indigo-600" />
                            Sẵn sàng kiểm tra
                          </span>
                        ) : isSatisfied ? (
                          <span className="px-2.5 py-0.5 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded-full text-[11px] font-bold flex items-center gap-1">
                            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                            Đạt (Thỏa mãn)
                          </span>
                        ) : (
                          <span className="px-2.5 py-0.5 bg-rose-100 text-rose-700 border border-rose-300 rounded-full text-[11px] font-extrabold flex items-center gap-1 animate-pulse">
                            <AlertTriangle className="w-3.5 h-3.5 text-rose-600" />
                            Vi phạm ({violations.length} vị trí)
                          </span>
                        )}

                        {/* Type Chip */}
                        <span className="px-2 py-0.5 bg-slate-100 text-slate-600 rounded-md text-[10px] font-semibold">
                          {getTypeLabel(constraint.type)}
                        </span>
                      </div>

                      {constraint.description && (
                        <p className="text-xs text-slate-500 mt-1 leading-relaxed">
                          {constraint.description}
                        </p>
                      )}

                      {/* Params summary preview */}
                      <div className="flex flex-wrap items-center gap-1.5 mt-2.5 text-[11px]">
                        {renderParamsPills(constraint)}
                      </div>
                    </div>
                  </div>

                  {/* Actions & Controls */}
                  <div className="flex items-center gap-1.5 self-end md:self-center shrink-0">
                    {!isSatisfied && violations.length > 0 && (
                      <button
                        onClick={() => toggleExpand(constraint.id)}
                        className="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-xl text-xs font-bold transition-all flex items-center gap-1 cursor-pointer"
                      >
                        {isExpanded ? 'Thu gọn' : `Xem chi tiết (${violations.length})`}
                        {isExpanded ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
                      </button>
                    )}

                    <button
                      onClick={() => openEditModal(constraint)}
                      title="Chỉnh sửa ràng buộc"
                      className="p-2 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-xl transition-all cursor-pointer"
                    >
                      <Edit3 className="w-4 h-4" />
                    </button>

                    <button
                      onClick={() => handleDuplicate(constraint)}
                      title="Nhân bản ràng buộc này"
                      className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition-all cursor-pointer"
                    >
                      <Copy className="w-4 h-4" />
                    </button>

                    <button
                      onClick={() => onDeleteConstraint(constraint.id)}
                      title="Xóa ràng buộc này"
                      className="p-2 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-xl transition-all cursor-pointer"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                </div>

                {/* Expanded Violation List */}
                {isExpanded && violations.length > 0 && (
                  <div className="border-t border-rose-100 bg-rose-50/40 p-4 sm:p-5">
                    <div className="flex items-center justify-between gap-2 mb-3">
                      <div className="text-xs font-bold text-rose-900 flex items-center gap-1.5">
                        <AlertTriangle className="w-4 h-4 text-rose-600" />
                        Danh sách {violations.length} tiết vi phạm ràng buộc (Cần điều chỉnh xếp lại):
                      </div>

                      <button
                        onClick={() => copyViolationText(constraint, violations)}
                        className="px-2.5 py-1 bg-white hover:bg-rose-50 text-rose-700 border border-rose-200 rounded-lg text-[11px] font-bold flex items-center gap-1 shadow-2xs transition-all cursor-pointer"
                      >
                        {copySuccessId === constraint.id ? (
                          <>
                            <Check className="w-3 h-3 text-emerald-600" />
                            Đã sao chép!
                          </>
                        ) : (
                          <>
                            <Copy className="w-3 h-3" />
                            Sao chép danh sách
                          </>
                        )}
                      </button>
                    </div>

                    <div className="overflow-x-auto rounded-xl border border-rose-200/80 bg-white shadow-2xs">
                      <table className="w-full text-left text-xs">
                        <thead className="bg-rose-50/70 text-[11px] font-bold text-rose-900 uppercase tracking-wider border-b border-rose-200">
                          <tr>
                            <th className="py-2.5 px-3">STT</th>
                            <th className="py-2.5 px-3">Thời Gian</th>
                            <th className="py-2.5 px-3">Lớp / Đối tượng</th>
                            <th className="py-2.5 px-3">Giáo Viên</th>
                            <th className="py-2.5 px-3">Vị trí Excel</th>
                            <th className="py-2.5 px-3">Chi Tiết Vi Phạm</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-rose-100">
                          {violations.map((viol, vIdx) => (
                            <tr key={viol.id} className="hover:bg-rose-50/30 transition-colors">
                              <td className="py-2.5 px-3 font-mono font-bold text-slate-500">{vIdx + 1}</td>
                              <td className="py-2.5 px-3 whitespace-nowrap">
                                <span className="font-bold text-slate-800">{viol.day}</span>
                                <span className="text-slate-400 mx-1">•</span>
                                <span className="text-indigo-600 font-semibold">{viol.period}</span>
                              </td>
                              <td className="py-2.5 px-3 font-bold text-slate-800 whitespace-nowrap">
                                {viol.className || '-'}
                              </td>
                              <td className="py-2.5 px-3 font-semibold text-slate-700 whitespace-nowrap">
                                {viol.teacher || '-'}
                              </td>
                              <td className="py-2.5 px-3 whitespace-nowrap">
                                {viol.rowIndex ? (
                                  <span className="px-2 py-0.5 bg-slate-100 text-slate-600 rounded font-mono font-bold text-[11px]">
                                    Dòng {viol.rowIndex}
                                  </span>
                                ) : '-'}
                              </td>
                              <td className="py-2.5 px-3 text-rose-700 font-medium leading-relaxed">
                                {viol.message}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>

      {/* Bottom CTA for Pre-upload mode */}
      {isPreUpload && onProceedToUpload && (
        <div className="bg-gradient-to-r from-teal-600 via-indigo-600 to-indigo-700 rounded-2xl p-5 sm:p-6 text-white flex flex-col sm:flex-row items-center justify-between gap-4 shadow-lg shadow-indigo-500/10 border border-white/10">
          <div className="flex items-center gap-3.5">
            <div className="p-3 bg-white/15 rounded-xl backdrop-blur-xs text-teal-200">
              <Upload className="w-6 h-6" />
            </div>
            <div>
              <div className="font-extrabold text-sm sm:text-base font-display">
                Đã sẵn sàng với {stats.active} yêu cầu ràng buộc
              </div>
              <div className="text-xs text-indigo-100 mt-0.5 max-w-lg">
                Tất cả các ràng buộc trên đã được lưu. Chuyển sang bước tải file để hệ thống nạp bảng Excel và đối soát tự động toàn bộ.
              </div>
            </div>
          </div>
          <button
            onClick={onProceedToUpload}
            id="btn-proceed-to-upload"
            className="w-full sm:w-auto px-6 py-3.5 bg-white text-slate-950 hover:bg-teal-50 font-black text-xs sm:text-sm rounded-xl shadow-lg transition-all flex items-center justify-center gap-2 cursor-pointer shrink-0"
          >
            <span>Tiếp Tục: Tải File TKB Lên</span>
            <ArrowRight className="w-4 h-4 text-indigo-600" />
          </button>
        </div>
      )}

      {/* Add / Edit Constraint Modal */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-white rounded-3xl max-w-2xl w-full max-h-[90vh] flex flex-col shadow-2xl border border-slate-200 overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            {/* Modal Header */}
            <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/80">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-indigo-50 text-indigo-600 rounded-xl">
                  <Sliders className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-extrabold text-slate-800 font-display">
                    {editingConstraint ? 'Chỉnh Sửa Ràng Buộc' : 'Thêm Ràng Buộc Mới'}
                  </h3>
                  <p className="text-xs text-slate-400">
                    Cấu hình điều kiện xếp lịch cần tự động rà soát
                  </p>
                </div>
              </div>

              <button
                onClick={() => setIsModalOpen(false)}
                className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-200/60 rounded-xl transition-all cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Body Form */}
            <form onSubmit={handleSaveForm} className="p-6 space-y-5 overflow-y-auto flex-1 text-xs">
              {/* Type Selection */}
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-2">
                  Loại Ràng Buộc <span className="text-rose-500">*</span>
                </label>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  {CONSTRAINT_TYPE_OPTIONS.map((opt) => {
                    const isSelected = formType === opt.type;
                    return (
                      <div
                        key={opt.type}
                        onClick={() => setFormType(opt.type)}
                        role="button"
                        tabIndex={0}
                        className={`p-3 rounded-xl border text-left cursor-pointer transition-all ${
                          isSelected 
                            ? 'border-indigo-600 bg-indigo-50/60 ring-2 ring-indigo-500/10' 
                            : 'border-slate-200 hover:border-slate-300 bg-white'
                        }`}
                      >
                        <div className="flex items-center justify-between">
                          <span className={`font-bold ${isSelected ? 'text-indigo-900' : 'text-slate-800'}`}>
                            {opt.label}
                          </span>
                          {isSelected && <Check className="w-4 h-4 text-indigo-600" />}
                        </div>
                        <p className="text-[11px] text-slate-400 mt-1 leading-snug">
                          {opt.desc}
                        </p>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Name & Description */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div className="sm:col-span-2">
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    Tên Ràng Buộc
                  </label>
                  <input
                    type="text"
                    value={formName}
                    onChange={(e) => setFormName(e.target.value)}
                    placeholder="VD: Kiểm tra trùng giờ Lớp 10.2 & 12.2"
                    className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 focus:bg-white"
                  />
                </div>

                <div className="sm:col-span-2">
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    Ghi Chú / Mô Tả
                  </label>
                  <input
                    type="text"
                    value={formDesc}
                    onChange={(e) => setFormDesc(e.target.value)}
                    placeholder="Mô tả lý do hoặc nguyên tắc phân công..."
                    className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 focus:bg-white"
                  />
                </div>
              </div>

              {/* Dynamic Parameter Fields based on Selected Type */}
              <div className="p-4 bg-slate-50 rounded-2xl border border-slate-200/80 space-y-4">
                <div className="text-[11px] font-extrabold uppercase tracking-wider text-indigo-900 flex items-center gap-1.5">
                  <Sliders className="w-3.5 h-3.5 text-indigo-600" />
                  Cấu Hình Chi Tiết Cho Loại Ràng Buộc
                </div>

                {/* 1. Simultaneous Classes & Teachers */}
                {formType === 'simultaneous_classes_teachers' && (
                  <div className="space-y-3">
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">Lớp thứ nhất</label>
                        <input
                          type="text"
                          value={formClass1}
                          onChange={(e) => setFormClass1(e.target.value)}
                          placeholder="VD: 10.2"
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-bold text-slate-800"
                        />
                      </div>
                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">Lớp thứ hai</label>
                        <input
                          type="text"
                          value={formClass2}
                          onChange={(e) => setFormClass2(e.target.value)}
                          placeholder="VD: 12.2"
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-bold text-slate-800"
                        />
                      </div>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">Giáo viên 1</label>
                        <input
                          type="text"
                          value={formTeacher1}
                          onChange={(e) => setFormTeacher1(e.target.value)}
                          placeholder="VD: Thầy Mịnh"
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-bold text-slate-800"
                        />
                      </div>
                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">Giáo viên 2</label>
                        <input
                          type="text"
                          value={formTeacher2}
                          onChange={(e) => setFormTeacher2(e.target.value)}
                          placeholder="VD: Thầy Đ. Minh"
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-bold text-slate-800"
                        />
                      </div>
                    </div>

                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">Chế độ kiểm tra</label>
                      <select
                        value={formMode}
                        onChange={(e) => setFormMode(e.target.value as any)}
                        className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-semibold text-slate-800"
                      >
                        <option value="must_not_overlap">Không được xếp trùng giờ (Cấm trùng tiết)</option>
                        <option value="must_overlap">Bắt buộc phải học cùng giờ (Ghép lớp/Song song)</option>
                      </select>
                    </div>
                  </div>
                )}

                {/* 2. Teacher Unavailable */}
                {formType === 'teacher_unavailable' && (
                  <div className="space-y-3">
                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">Tên Giáo Viên</label>
                      <input
                        type="text"
                        value={formTeacher}
                        onChange={(e) => setFormTeacher(e.target.value)}
                        placeholder="VD: Thầy Hưng hoặc Thầy Nam..."
                        className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-bold text-slate-800"
                      />
                      {availableTeachers.length > 0 && (
                        <div className="mt-1.5 flex flex-wrap gap-1">
                          <span className="text-[10px] text-slate-400">Gợi ý từ file:</span>
                          {availableTeachers.slice(0, 6).map(t => (
                            <button
                              key={t}
                              type="button"
                              onClick={() => setFormTeacher(t)}
                              className="text-[10px] px-1.5 py-0.5 bg-slate-200/70 hover:bg-indigo-100 hover:text-indigo-700 rounded text-slate-700"
                            >
                              {t}
                            </button>
                          ))}
                        </div>
                      )}
                    </div>

                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">Các Thứ bận / nghỉ</label>
                      <div className="flex flex-wrap gap-1.5">
                        {['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7'].map(d => {
                          const isChecked = formDays.includes(d);
                          return (
                            <button
                              key={d}
                              type="button"
                              onClick={() => {
                                setFormDays(prev => 
                                  isChecked ? prev.filter(x => x !== d) : [...prev, d]
                                );
                              }}
                              className={`px-2.5 py-1 rounded-lg text-xs font-bold border transition-all cursor-pointer ${
                                isChecked 
                                  ? 'bg-indigo-600 text-white border-indigo-600 shadow-2xs' 
                                  : 'bg-white text-slate-700 border-slate-200 hover:border-slate-300'
                              }`}
                            >
                              {d}
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">Buổi bận</label>
                      <div className="flex gap-2">
                        {(['Sáng', 'Chiều'] as DaySession[]).map(s => {
                          const isChecked = formSessions.includes(s);
                          return (
                            <button
                              key={s}
                              type="button"
                              onClick={() => {
                                setFormSessions(prev => 
                                  isChecked ? prev.filter(x => x !== s) : [...prev, s]
                                );
                              }}
                              className={`px-3 py-1 rounded-lg text-xs font-bold border transition-all cursor-pointer ${
                                isChecked 
                                  ? 'bg-indigo-600 text-white border-indigo-600 shadow-2xs' 
                                  : 'bg-white text-slate-700 border-slate-200 hover:border-slate-300'
                              }`}
                            >
                              {s}
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  </div>
                )}

                {/* 3. Teacher Max Periods */}
                {formType === 'teacher_max_periods_day' && (
                  <div className="space-y-3">
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">Số tiết tối đa / ngày</label>
                        <input
                          type="number"
                          min={1}
                          max={10}
                          value={formMaxPeriods}
                          onChange={(e) => setFormMaxPeriods(Number(e.target.value))}
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-bold text-slate-800"
                        />
                      </div>
                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">Áp dụng cho Giáo viên cụ thể</label>
                        <input
                          type="text"
                          value={formTeacher}
                          onChange={(e) => setFormTeacher(e.target.value)}
                          placeholder="Để trống nếu áp dụng cho TẤT CẢ giáo viên"
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-bold text-slate-800"
                        />
                      </div>
                    </div>
                  </div>
                )}

                {/* 4. No Split Day */}
                {formType === 'no_split_day' && (
                  <div className="space-y-2 text-xs text-slate-600">
                    <p>
                      Quy tắc này tự động kiểm tra xem trong cùng 1 ngày, giáo viên dạy môn học cho lớp có bị chia thành 2 buổi (sáng và chiều) hay không.
                    </p>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mt-2">
                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">Chỉ áp dụng cho lớp (tùy chọn)</label>
                        <input
                          type="text"
                          value={formClass1}
                          onChange={(e) => setFormClass1(e.target.value)}
                          placeholder="Để trống để kiểm tra tất cả lớp"
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-slate-800 font-semibold"
                        />
                      </div>
                      <div>
                        <label className="block text-[11px] font-bold text-slate-700 mb-1">Chỉ áp dụng cho GV (tùy chọn)</label>
                        <input
                          type="text"
                          value={formTeacher}
                          onChange={(e) => setFormTeacher(e.target.value)}
                          placeholder="Để trống để kiểm tra tất cả GV"
                          className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-slate-800 font-semibold"
                        />
                      </div>
                    </div>
                  </div>
                )}

                {/* 5. Subject Period Restriction */}
                {formType === 'subject_period_restriction' && (
                  <div className="space-y-3">
                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">Tên Môn Học</label>
                      <input
                        type="text"
                        value={formSubject}
                        onChange={(e) => setFormSubject(e.target.value)}
                        placeholder="VD: Thể dục, Chào cờ, Tin học..."
                        className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-bold text-slate-800"
                      />
                    </div>
                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">Không xếp vào Buổi</label>
                      <div className="flex gap-2">
                        {(['Sáng', 'Chiều'] as DaySession[]).map(s => {
                          const isChecked = formSessions.includes(s);
                          return (
                            <button
                              key={s}
                              type="button"
                              onClick={() => {
                                setFormSessions(prev => 
                                  isChecked ? prev.filter(x => x !== s) : [...prev, s]
                                );
                              }}
                              className={`px-3 py-1 rounded-lg text-xs font-bold border transition-all cursor-pointer ${
                                isChecked 
                                  ? 'bg-indigo-600 text-white border-indigo-600' 
                                  : 'bg-white text-slate-700 border-slate-200'
                              }`}
                            >
                              {s}
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  </div>
                )}

                {/* 6. Class Pair No Overlap */}
                {formType === 'class_pair_no_overlap' && (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">Lớp 1</label>
                      <input
                        type="text"
                        value={formOverlapClass1}
                        onChange={(e) => setFormOverlapClass1(e.target.value)}
                        placeholder="VD: 10A1"
                        className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-bold text-slate-800"
                      />
                    </div>
                    <div>
                      <label className="block text-[11px] font-bold text-slate-700 mb-1">Lớp 2</label>
                      <input
                        type="text"
                        value={formOverlapClass2}
                        onChange={(e) => setFormOverlapClass2(e.target.value)}
                        placeholder="VD: 10A2"
                        className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-bold text-slate-800"
                      />
                    </div>
                  </div>
                )}

                {/* 7. Teacher No Gap */}
                {formType === 'teacher_no_gap' && (
                  <div>
                    <label className="block text-[11px] font-bold text-slate-700 mb-1">Giáo viên áp dụng</label>
                    <input
                      type="text"
                      value={formTeacher}
                      onChange={(e) => setFormTeacher(e.target.value)}
                      placeholder="Để trống để kiểm tra tất cả giáo viên"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-bold text-slate-800"
                    />
                  </div>
                )}
              </div>

              {/* Status active */}
              <div className="flex items-center justify-between pt-2">
                <span className="text-xs font-bold text-slate-700">Kích hoạt kiểm tra ngay</span>
                <button
                  type="button"
                  onClick={() => setFormIsActive(!formIsActive)}
                  className={`relative inline-flex h-6 w-11 shrink-0 cursor-pointer rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out focus:outline-none ${
                    formIsActive ? 'bg-indigo-600' : 'bg-slate-200'
                  }`}
                >
                  <span
                    className={`pointer-events-none inline-block h-5 w-5 transform rounded-full bg-white shadow-sm ring-0 transition duration-200 ease-in-out ${
                      formIsActive ? 'translate-x-5' : 'translate-x-0'
                    }`}
                  />
                </button>
              </div>

              {/* Submit Buttons */}
              <div className="flex items-center justify-end gap-3 pt-4 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setIsModalOpen(false)}
                  className="px-4 py-2 text-slate-600 hover:bg-slate-100 rounded-xl text-xs font-bold transition-all cursor-pointer"
                >
                  Hủy bỏ
                </button>
                <button
                  type="submit"
                  className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-xs font-bold transition-all shadow-md hover:shadow-indigo-500/30 flex items-center gap-2 cursor-pointer"
                >
                  <Check className="w-4 h-4" />
                  {editingConstraint ? 'Lưu Thay Đổi' : 'Tạo Ràng Buộc'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

function getTypeLabel(type: ConstraintType): string {
  switch (type) {
    case 'simultaneous_classes_teachers': return 'Trùng giờ 2 lớp';
    case 'teacher_unavailable': return 'GV bận / nghỉ';
    case 'teacher_max_periods_day': return 'Số tiết tối đa';
    case 'no_split_day': return 'Không chia buổi';
    case 'subject_period_restriction': return 'Ràng buộc môn học';
    case 'class_pair_no_overlap': return 'Chung phòng/sân';
    case 'teacher_no_gap': return 'Tránh loãng xương';
    default: return 'Ràng buộc';
  }
}

function renderParamsPills(c: TimetableConstraint) {
  const pills: React.ReactNode[] = [];

  if (c.type === 'simultaneous_classes_teachers') {
    pills.push(
      <span key="classes" className="px-2 py-0.5 bg-indigo-50 text-indigo-700 font-bold rounded">
        Lớp: {c.params.class1 || '10.2'} & {c.params.class2 || '12.2'}
      </span>
    );
    pills.push(
      <span key="teachers" className="px-2 py-0.5 bg-purple-50 text-purple-700 font-bold rounded">
        GV: {c.params.teacher1 || 'Thầy Mịnh'} & {c.params.teacher2 || 'Thầy Đ. Minh'}
      </span>
    );
  } else if (c.type === 'teacher_unavailable') {
    if (c.params.teacher) {
      pills.push(<span key="t" className="px-2 py-0.5 bg-blue-50 text-blue-700 font-bold rounded">GV: {c.params.teacher}</span>);
    }
    if (c.params.days && c.params.days.length > 0) {
      pills.push(<span key="d" className="px-2 py-0.5 bg-slate-100 text-slate-700 rounded">Thứ: {c.params.days.join(', ')}</span>);
    }
    if (c.params.sessions && c.params.sessions.length > 0) {
      pills.push(<span key="s" className="px-2 py-0.5 bg-slate-100 text-slate-700 rounded">Buổi: {c.params.sessions.join(', ')}</span>);
    }
  } else if (c.type === 'teacher_max_periods_day') {
    pills.push(
      <span key="max" className="px-2 py-0.5 bg-amber-50 text-amber-800 font-bold rounded">
        Tối đa: {c.params.maxPeriodsPerDay || 5} tiết / ngày
      </span>
    );
    if (c.params.teacher) {
      pills.push(<span key="t" className="px-2 py-0.5 bg-blue-50 text-blue-700 rounded">GV: {c.params.teacher}</span>);
    }
  } else if (c.type === 'subject_period_restriction') {
    if (c.params.subject) {
      pills.push(<span key="sub" className="px-2 py-0.5 bg-teal-50 text-teal-700 font-bold rounded">Môn: {c.params.subject}</span>);
    }
  } else if (c.type === 'class_pair_no_overlap') {
    pills.push(
      <span key="pair" className="px-2 py-0.5 bg-rose-50 text-rose-700 font-bold rounded">
        {c.params.overlapClass1 || 'Lớp 1'} & {c.params.overlapClass2 || 'Lớp 2'}
      </span>
    );
  }

  return pills;
}
