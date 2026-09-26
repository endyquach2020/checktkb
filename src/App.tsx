import React, { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  ExemptionPair, 
  ColumnMapping, 
  ExtractionSettings, 
  ParseResult,
  CellRange,
  SimultaneousMatch,
  SimultaneousCheckConfig
} from './types';
import { 
  detectConflicts, 
  detectSplitPeriods,
  detectSimultaneousTeaching,
  generateSampleExcel, 
  extractTeacher, 
  normalizeTeacherName,
  parseExcelRange,
  indexToColName
} from './utils/excelParser';
import ExemptionSettingsComponent from './components/ExemptionSettings';
import MappingSetup from './components/MappingSetup';
import ConflictReport from './components/ConflictReport';
import { ConstraintManager } from './components/ConstraintManager';
import { 
  TimetableConstraint, 
  ConstraintCheckResult 
} from './types';
import { 
  getStoredConstraints, 
  saveStoredConstraints, 
  DEFAULT_CONSTRAINTS, 
  evaluateConstraints 
} from './utils/constraintEngine';
import { 
  Upload, 
  FileSpreadsheet, 
  Settings2, 
  AlertTriangle, 
  RefreshCw, 
  HelpCircle, 
  FileDown, 
  Check, 
  ArrowRight,
  Sparkles,
  Clock,
  Users,
  ShieldCheck
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

export default function App() {
  // Step State
  const [step, setStep] = useState<'upload' | 'mapping' | 'results'>('upload');

  // File & Excel Data State
  const [file, setFile] = useState<File | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [parseResult, setParseResult] = useState<ParseResult | null>(null);
  const [mappings, setMappings] = useState<ColumnMapping[]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Exemption Pairs State (default 3 pairs as requested)
  const [exemptions, setExemptions] = useState<ExemptionPair[]>([
    { id: 'def-10', classA: '10.1 TN', classB: '10.1 XH', isActive: true },
    { id: 'def-11', classA: '11.1 TN', classB: '11.2 XH', isActive: true },
    { id: 'def-12', classA: '12.1 TN', classB: '12.2 XH', isActive: true },
  ]);

  // Teacher Name Extraction Settings (persisted in localStorage)
  const [extractionSettings, setExtractionSettings] = useState<ExtractionSettings>(() => {
    try {
      const saved = localStorage.getItem('check_sched_extraction_settings');
      if (saved) {
        return JSON.parse(saved);
      }
    } catch (e) {
      console.error('Error loading extractionSettings from localStorage', e);
    }
    return {
      delimiter: '-',
      takeFirstPart: false, // Default Subject - Teacher (take second part)
      ignoreList: ['shcn', 'chào cờ', 'chao co', 'sinh hoạt', 'nghỉ', 'nghi', 'chủ nhiệm', 'x', '-', 'chủ nhật', 'cn'],
    };
  });

  // State for raw ignore list text in settings panel (initialized from the persisted ignoreList)
  const [ignoreListText, setIgnoreListText] = useState(() => extractionSettings.ignoreList.join(', '));

  // Sync extraction settings to localStorage on change
  React.useEffect(() => {
    try {
      localStorage.setItem('check_sched_extraction_settings', JSON.stringify(extractionSettings));
    } catch (e) {
      console.error('Error saving extractionSettings to localStorage', e);
    }
  }, [extractionSettings]);

  // Profile/Cấp học rà soát: 'THPT' (K5:V100) | 'TieuHoc' (E5:J100) | 'All' (Toàn bộ)
  const [checkProfile, setCheckProfile] = useState<'THPT' | 'TieuHoc' | 'All'>(() => {
    try {
      const saved = localStorage.getItem('check_sched_profile');
      if (saved && (saved === 'THPT' || saved === 'TieuHoc' || saved === 'All')) {
        return saved as 'THPT' | 'TieuHoc' | 'All';
      }
    } catch (e) {
      console.error('Error loading checkProfile from localStorage', e);
    }
    return 'All';
  });

  // Sync checkProfile to localStorage on change
  React.useEffect(() => {
    try {
      localStorage.setItem('check_sched_profile', checkProfile);
    } catch (e) {
      console.error('Error saving checkProfile to localStorage', e);
    }
  }, [checkProfile]);

  // Compute active cell range based on the selected profile
  const activeRange = useMemo(() => {
    if (checkProfile === 'THPT') {
      return parseExcelRange('K5:V100');
    }
    if (checkProfile === 'TieuHoc') {
      return parseExcelRange('E5:J100');
    }
    return null;
  }, [checkProfile]);

  // Configuration for checking simultaneous class & teacher schedule (Default: 10.2 & 12.2 with Thầy Mịnh & Thầy Đ. Minh)
  const [simultaneousConfig, setSimultaneousConfig] = useState<SimultaneousCheckConfig>({
    classQuery1: '10.2',
    classQuery2: '12.2',
    teacherQuery1: 'Thầy Mịnh',
    teacherQuery2: 'Thầy Đ. Minh'
  });

  // User-defined constraints for automated timetable verification (persisted in localStorage)
  const [constraints, setConstraints] = useState<TimetableConstraint[]>(() => {
    return getStoredConstraints();
  });

  // Persist constraints to localStorage whenever changed
  useEffect(() => {
    saveStoredConstraints(constraints);
  }, [constraints]);

  const handleSaveConstraint = (item: TimetableConstraint) => {
    setConstraints(prev => {
      const idx = prev.findIndex(c => c.id === item.id);
      if (idx >= 0) {
        const next = [...prev];
        next[idx] = item;
        return next;
      }
      return [item, ...prev];
    });

    if (item.type === 'simultaneous_classes_teachers' && item.params.class1 && item.params.class2) {
      setSimultaneousConfig({
        classQuery1: item.params.class1,
        classQuery2: item.params.class2,
        teacherQuery1: item.params.teacher1 || '',
        teacherQuery2: item.params.teacher2 || ''
      });
    }
  };

  const handleUpdateSimultaneousConfig = (newConfig: SimultaneousCheckConfig) => {
    setSimultaneousConfig(newConfig);
    setConstraints(prev => prev.map(c => {
      if (c.type === 'simultaneous_classes_teachers' && c.id === 'c-simul-10-12') {
        return {
          ...c,
          name: `Kiểm tra trùng giờ Lớp ${newConfig.classQuery1} & ${newConfig.classQuery2} (${newConfig.teacherQuery1} & ${newConfig.teacherQuery2})`,
          params: {
            ...c.params,
            class1: newConfig.classQuery1,
            class2: newConfig.classQuery2,
            teacher1: newConfig.teacherQuery1,
            teacher2: newConfig.teacherQuery2
          }
        };
      }
      return c;
    }));
  };

  const handleDeleteConstraint = (id: string) => {
    setConstraints(prev => prev.filter(c => c.id !== id));
  };

  const handleToggleConstraint = (id: string, isActive: boolean) => {
    setConstraints(prev => prev.map(c => c.id === id ? { ...c, isActive } : c));
  };

  const handleResetConstraints = () => {
    setConstraints(DEFAULT_CONSTRAINTS);
  };

  // Initial workflow tab: allow entering all requirements/constraints first, then uploading timetable
  const [uploadTab, setUploadTab] = useState<'setup_constraints' | 'upload_file'>('setup_constraints');

  // Handle excel parsing
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      processFile(files[0]);
    }
  };

  const processFile = (uploadedFile: File, activeSheetName?: string) => {
    setFile(uploadedFile);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const bstr = e.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const sheetNames = wb.SheetNames;
        if (sheetNames.length === 0) {
          alert('File Excel không có sheet nào hợp lệ!');
          return;
        }

        // Filter out sheets that look like individual teacher/subject sheets (containing hyphen '-')
        // We always keep the very first sheet, and any sheet containing master keywords like "TKB", "tổng", "master", "chung"
        const filteredSheetNames = sheetNames.filter((name, idx) => {
          const nLower = name.toLowerCase().trim();
          if (idx === 0) return true;
          if (nLower.includes('tkb') || nLower.includes('tổng') || nLower.includes('tong') || nLower.includes('master') || nLower.includes('chung')) {
            return true;
          }
          if (name.includes('-')) {
            return false;
          }
          return true;
        });

        const activeSheet = activeSheetName || filteredSheetNames[0] || sheetNames[0];
        const ws = wb.Sheets[activeSheet];
        
        // Parse sheet to 2D array of strings
        const rawData = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1, defval: '' });
        if (rawData.length === 0) {
          alert('Sheet Excel trống!');
          return;
        }

        // Clean raw rows and filter trailing/leading blank rows
        const cleanedData = rawData.map(row => 
          Array.isArray(row) 
            ? row.map(cell => cell !== null && cell !== undefined ? String(cell).trim() : '') 
            : []
        );

        // Class names (headers) are on Row 5 (index 4) for profiles, or Row 1 (index 0) for All mode
        const isProfileActive = checkProfile === 'THPT' || checkProfile === 'TieuHoc';
        const headerRowIdx = isProfileActive ? 4 : 0;
        const dataRowStartIdx = isProfileActive ? 5 : 4;

        if (cleanedData.length < dataRowStartIdx + 1) {
          alert(`Tệp Excel phải có ít nhất ${dataRowStartIdx + 1} dòng để thực hiện kiểm tra!`);
          return;
        }

        const rawHeaders = cleanedData[headerRowIdx] || [];
        const rawContentRows = cleanedData.slice(dataRowStartIdx);

        // Fill empty headers with Cột index
        const headers = rawHeaders.map((h, idx) => h || `Cột ${idx + 1}`);

        // Automatically detect column mappings
        let detectedDay = false;
        let detectedPeriod = false;

        const autoMappings = headers.map((header, index) => {
          const hLower = header.toLowerCase().trim();
          let role: 'day' | 'period' | 'class' | 'skip' = 'class';

          if (checkProfile === 'THPT') {
            if (index === 10) {
              role = 'period'; // Cột K (10) là Tiết / Khung giờ
              detectedPeriod = true;
            } else if (index >= 11 && index <= 21) {
              role = 'class'; // Cột L đến V là các lớp học
            } else {
              role = 'skip'; // Các cột khác, bao gồm cột thứ, để "chưa xác định" / skip
            }
          } else if (checkProfile === 'TieuHoc') {
            if (index === 4) {
              role = 'period'; // Cột E (4) là Tiết / Khung giờ
              detectedPeriod = true;
            } else if (index >= 5 && index <= 9) {
              role = 'class'; // Cột F đến J là các lớp học
            } else {
              role = 'skip'; // Các cột khác, bao gồm cột thứ, để "chưa xác định" / skip
            }
          } else {
            // Standard / All mode
            // Detect Day Column
            if (!detectedDay && (hLower.includes('thứ') || hLower.includes('thu') || hLower.includes('ngày') || hLower.includes('day'))) {
              role = 'day';
              detectedDay = true;
            } 
            // Detect Period Column
            else if (!detectedPeriod && (hLower.includes('tiết') || hLower.includes('tiet') || hLower.includes('giờ') || hLower.includes('gio') || hLower.includes('thời gian') || hLower.includes('period') || hLower.includes('time'))) {
              role = 'period';
              detectedPeriod = true;
            }
            // Skip index or blank columns
            else if (!header || hLower === 'stt' || hLower === 'no' || hLower === '') {
              role = 'skip';
            }
          }

          return {
            index,
            header: header || `Cột ${index + 1}`,
            role: role as any,
          };
        });

        // Fallbacks if not auto-detected (Only in "All" mode)
        if (checkProfile === 'All') {
          if (!detectedDay && headers.length > 0) {
            autoMappings[0].role = 'day' as any;
          }
          if (!detectedPeriod && headers.length > 1) {
            autoMappings[1].role = 'period' as any;
          }
        }

        setParseResult({
          headers,
          rows: rawContentRows,
          sheetNames: filteredSheetNames,
          activeSheetName: activeSheet
        });
        
        setMappings(autoMappings);
        setStep('results');
      } catch (err) {
        console.error(err);
        alert('Đã xảy ra lỗi khi đọc file Excel. Vui lòng kiểm tra định dạng file.');
      }
    };
    reader.readAsBinaryString(uploadedFile);
  };

  // Drag and drop handlers
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      const droppedFile = files[0];
      const isExcel = droppedFile.name.endsWith('.xlsx') || droppedFile.name.endsWith('.xls');
      if (isExcel) {
        processFile(droppedFile);
      } else {
        alert('Vui lòng chỉ tải lên file Excel (.xlsx hoặc .xls)!');
      }
    }
  };

  // Switch sheets if multiple sheets exist
  const handleSheetChange = (sheetName: string) => {
    if (!file) return;
    processFile(file, sheetName);
  };

  // Apply ignore list updates
  const handleIgnoreListSave = () => {
    const list = ignoreListText
      .split(',')
      .map(item => item.trim())
      .filter(item => item !== '');
    setExtractionSettings(prev => ({ ...prev, ignoreList: list }));
  };

  // Re-calculate conflicts on demand based on mappings, exemptions, and extraction settings
  const conflictReportData = useMemo(() => {
    if (!parseResult || mappings.length === 0) {
      return { conflicts: [], totalCellsChecked: 0, totalTeachers: 0 };
    }

    const { rows } = parseResult;
    
    // Slice rows based on checkProfile: Row 5 to 100 means index 0 to 95 of rawContentRows (rows)
    let rowsToCheck = rows;
    if (checkProfile === 'THPT' || checkProfile === 'TieuHoc') {
      rowsToCheck = rows.slice(0, 96);
    }

    // Filter class columns based on checkProfile by generating a reliable array covering all columns
    let activeMappings = mappings;
    if (checkProfile === 'THPT') {
      const maxCols = Math.max(22, ...rowsToCheck.map(r => r.length));
      activeMappings = Array.from({ length: maxCols }, (_, i) => {
        let role: 'day' | 'period' | 'class' | 'skip' = 'skip';
        if (i === 10) {
          role = 'period'; // Khung giờ ở cột K (10)
        } else if (i >= 11 && i <= 21) {
          role = 'class'; // Lớp học ở cột L-V (11-21)
        }
        return {
          index: i,
          header: parseResult.headers[i] || `Cột ${XLSX.utils.encode_col(i)}`,
          role,
        };
      });
    } else if (checkProfile === 'TieuHoc') {
      const maxCols = Math.max(10, ...rowsToCheck.map(r => r.length));
      activeMappings = Array.from({ length: maxCols }, (_, i) => {
        let role: 'day' | 'period' | 'class' | 'skip' = 'skip';
        if (i === 4) {
          role = 'period'; // Khung giờ ở cột E (4)
        } else if (i >= 5 && i <= 9) {
          role = 'class'; // Lớp học ở cột F-J (5-9)
        }
        return {
          index: i,
          header: parseResult.headers[i] || `Cột ${XLSX.utils.encode_col(i)}`,
          role,
        };
      });
    }

    const conflicts = detectConflicts(rowsToCheck, activeMappings, exemptions, extractionSettings, checkProfile);
    const splitIssues = detectSplitPeriods(rowsToCheck, activeMappings, extractionSettings, checkProfile);
    const simultaneousMatches = detectSimultaneousTeaching(
      rowsToCheck,
      activeMappings,
      extractionSettings,
      simultaneousConfig,
      checkProfile
    );

    // Evaluate user-defined constraints
    const constraintResults = evaluateConstraints(
      rowsToCheck,
      activeMappings,
      extractionSettings,
      constraints,
      checkProfile
    );

    // Calculate statistics
    let cellCheckCount = 0;
    const teacherSet = new Set<string>();
    const classCols = activeMappings.filter(m => m.role === 'class');

    rowsToCheck.forEach(row => {
      classCols.forEach(col => {
        const cellVal = row[col.index];
        if (cellVal !== undefined && cellVal !== null && String(cellVal).trim() !== '') {
          cellCheckCount++;
          const extracted = extractTeacher(cellVal, extractionSettings);
          if (extracted && extracted.teacherName) {
            teacherSet.add(normalizeTeacherName(extracted.teacherName));
          }
        }
      });
    });

    const availableTeachers = Array.from(teacherSet).sort((a, b) => a.localeCompare(b, 'vi'));
    const availableClasses = classCols.map(c => c.header).filter(Boolean);

    return {
      conflicts,
      splitIssues,
      simultaneousMatches,
      constraintResults,
      availableTeachers,
      availableClasses,
      totalCellsChecked: cellCheckCount,
      totalTeachers: teacherSet.size
    };
  }, [parseResult, mappings, exemptions, extractionSettings, checkProfile, simultaneousConfig, constraints]);

  // Restart / Reset
  const handleReset = () => {
    setFile(null);
    setParseResult(null);
    setMappings([]);
    setStep('upload');
  };

  return (
    <div className="min-h-screen bg-slate-50/50 text-slate-800 font-sans flex flex-col lg:flex-row antialiased overflow-hidden h-screen" id="app-root">
      
      {/* LEFT SIDEBAR: App Branding, Active File Metadata, and Sidebar Settings */}
      <aside className="w-full lg:w-80 bg-[#0a665e] text-white border-b lg:border-b-0 lg:border-r border-teal-900/60 flex flex-col shrink-0 h-auto lg:h-full overflow-y-auto" id="sidebar-container">
        {/* Branding Header */}
        <div className="p-6 border-b border-teal-800/40 flex items-center justify-between shrink-0 bg-teal-950/20" id="sidebar-header">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-amber-400 text-teal-950 rounded-xl shadow-md font-black text-center flex items-center justify-center w-10 h-10 shrink-0 select-none text-sm">
              TKB
            </div>
            <div>
              <h1 className="text-sm font-extrabold text-white tracking-tight flex items-center gap-1.5 font-display">
                CheckSched Pro
                <span className="text-[9px] font-bold bg-teal-850 text-teal-100 px-1.5 py-0.5 rounded-md border border-teal-700/50">
                  v2.5
                </span>
              </h1>
              <p className="text-[10px] text-teal-200/80 font-bold uppercase tracking-wider">Hệ thống rà soát trùng lịch</p>
            </div>
          </div>
        </div>

        {/* Sidebar Body Content */}
        <div className="p-6 flex-1 space-y-6" id="sidebar-body">
          {/* Active File State Indicator */}
          <div className="space-y-2.5" id="sidebar-file-status">
            <div className="text-[10px] font-bold text-teal-200/80 uppercase tracking-wider">Tệp tin đang mở</div>
            {file ? (
              <div className="p-4 bg-teal-950/25 border border-teal-800/40 rounded-xl space-y-3">
                <div className="flex items-start gap-2.5">
                  <div className="p-1.5 bg-emerald-500/20 text-emerald-300 rounded-lg shrink-0 mt-0.5 border border-emerald-500/20">
                    <Check className="w-3.5 h-3.5" />
                  </div>
                  <div className="min-w-0 flex-1">
                    <p className="text-xs font-bold text-white truncate" title={file.name}>{file.name}</p>
                    <p className="text-[10px] text-teal-300/80 font-semibold mt-0.5">Dung lượng: {(file.size / 1024).toFixed(1)} KB</p>
                  </div>
                </div>



                <button
                  onClick={handleReset}
                  className="w-full flex items-center justify-center gap-1.5 py-2 border border-teal-800/60 bg-teal-950/20 hover:bg-teal-950/35 text-[11px] font-bold text-white rounded-xl transition-all cursor-pointer shadow-sm shadow-teal-950/5"
                  id="sidebar-change-file-btn"
                >
                  <RefreshCw className="w-3 h-3 text-teal-300" />
                  Chọn tệp tin khác
                </button>
              </div>
            ) : (
              <div className="p-4 bg-teal-950/20 border border-dashed border-teal-800/50 rounded-xl text-center">
                <p className="text-xs text-teal-200/60 font-semibold italic">Chưa tải tệp tin lên</p>
              </div>
            )}
          </div>

          {/* Cấu hình Vùng Rà Soát (Excel Cell Range) */}
          <div className="space-y-2.5 pt-4 border-t border-teal-800/30" id="sidebar-range-config">
            <div className="text-[10px] font-bold text-teal-200/80 uppercase tracking-wider flex items-center gap-1.5">
              <Settings2 className="w-3.5 h-3.5 text-emerald-400" />
              Cấu hình cấp học rà soát
            </div>

            <div className="flex flex-col gap-2">
              <select
                value={checkProfile}
                onChange={(e) => setCheckProfile(e.target.value as any)}
                className="w-full text-xs px-2.5 py-1.5 bg-teal-950/40 border border-teal-800/80 rounded-xl text-teal-100 font-bold focus:outline-none focus:ring-1 focus:ring-emerald-400 cursor-pointer transition-all"
                id="sidebar-profile-select"
              >
                <option value="All" className="bg-[#0a665e] text-white">🎒 Toàn bộ bảng (Mặc định)</option>
                <option value="THPT" className="bg-[#0a665e] text-white">🏫 Trung học Phổ thông (K5:V100)</option>
                <option value="TieuHoc" className="bg-[#0a665e] text-white">🏫 Tiểu học (E5:J100)</option>
              </select>

              <div className="bg-teal-950/20 p-4 rounded-xl border border-teal-800/40 space-y-2">
                <div className="text-[11px] text-teal-100 space-y-2 font-medium leading-relaxed">
                  <p className="flex items-start gap-1.5">
                    <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full mt-1.5 shrink-0"></span>
                    <span>
                      Thời gian: <strong className="text-white font-bold font-display">
                        {checkProfile === 'THPT' && 'Cột K'}
                        {checkProfile === 'TieuHoc' && 'Cột E'}
                        {checkProfile === 'All' && 'Cột nhận diện tự động'}
                      </strong>
                    </span>
                  </p>
                  <p className="flex items-start gap-1.5">
                    <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full mt-1.5 shrink-0"></span>
                    <span>
                      Cột rà soát: <strong className="text-white font-bold font-display">
                        {checkProfile === 'THPT' && 'Cột lớp L đến V'}
                        {checkProfile === 'TieuHoc' && 'Cột lớp F đến J'}
                        {checkProfile === 'All' && 'Tự động (Tất cả cột)'}
                      </strong>
                    </span>
                  </p>
                  <p className="flex items-start gap-1.5">
                    <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full mt-1.5 shrink-0"></span>
                    <span>
                      Tên lớp tại: <strong className="text-white font-bold font-display">
                        {checkProfile === 'All' ? 'Dòng 1' : 'Dòng 5'}
                      </strong>
                    </span>
                  </p>
                   <p className="flex items-start gap-1.5">
                    <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full mt-1.5 shrink-0"></span>
                    <span>
                      Dữ liệu lịch dạy: <strong className="text-white font-bold font-display">
                        {checkProfile === 'All' ? 'Dòng 5 trở đi' : 'Dòng 10 đến dòng 100'}
                      </strong>
                    </span>
                  </p>
                  <p className="text-[10px] text-teal-300/60 italic pt-2 border-t border-teal-800/30 leading-relaxed">
                    {checkProfile === 'THPT' && 'Cột K là thời gian. Cột L-V là lớp. Thứ 2 (dòng 10-24), Thứ 3 (dòng 28-43), Thứ 4 (dòng 47-62), Thứ 5 (dòng 66-81), Thứ 6 (dòng 85-100).'}
                    {checkProfile === 'TieuHoc' && 'Cột E là thời gian. Cột F-J là lớp. Thứ 2 (dòng 10-24), Thứ 3 (dòng 28-43), Thứ 4 (dòng 47-62), Thứ 5 (dòng 66-81), Thứ 6 (dòng 85-100).'}
                    {checkProfile === 'All' && 'Rà soát toàn bộ tệp Excel (Dòng 1 chứa tên lớp, Dòng 5 trở xuống là lịch học).'}
                  </p>
                </div>
              </div>
            </div>
          </div>

          {/* FULL Trích xuất & loại trừ Settings Moved to Sidebar as requested */}
          {(step === 'results' || step === 'mapping') && (
            <div className="space-y-3 pt-4 border-t border-teal-800/40" id="sidebar-extraction-config">
              <div className="text-[10px] font-bold text-teal-200/80 uppercase tracking-wider flex items-center gap-1.5">
                <Settings2 className="w-3.5 h-3.5 text-emerald-400" />
                Trích xuất & loại trừ
              </div>

              <div className="bg-teal-950/20 p-4 rounded-xl border border-teal-800/40 space-y-4">
                {/* Delimiter */}
                <div className="flex flex-col gap-1.5">
                  <label className="text-[10px] font-bold text-teal-200/80 uppercase tracking-wider">Ký tự phân tách (Môn - GV)</label>
                  <input
                    type="text"
                    value={extractionSettings.delimiter}
                    onChange={(e) => setExtractionSettings(prev => ({ ...prev, delimiter: e.target.value }))}
                    placeholder="Ví dụ: -"
                    className="w-full px-3 py-1.5 bg-teal-950/40 border border-teal-800/80 rounded-lg text-xs focus:outline-none focus:ring-1 focus:ring-emerald-400 font-mono font-bold text-white transition-all placeholder:text-teal-700"
                    maxLength={5}
                  />
                </div>

                {/* Info order */}
                <div className="space-y-2">
                  <label className="text-[10px] font-bold text-teal-200/80 uppercase tracking-wider block">Thứ tự thông tin trong ô</label>
                  <div className="space-y-1.5">
                    <label className="flex items-center gap-2 cursor-pointer text-xs text-teal-100 hover:text-white group">
                      <input
                        type="radio"
                        name="format-order-sidebar"
                        checked={!extractionSettings.takeFirstPart}
                        onChange={() => setExtractionSettings(prev => ({ ...prev, takeFirstPart: false }))}
                        className="text-emerald-500 focus:ring-emerald-400/20 w-3.5 h-3.5 cursor-pointer accent-emerald-500"
                      />
                      <span className="group-hover:text-white transition-colors text-[11px]">Môn học {extractionSettings.delimiter || '-'} <strong className="text-emerald-400 font-bold">Giáo viên</strong></span>
                    </label>
                    <label className="flex items-center gap-2 cursor-pointer text-xs text-teal-100 hover:text-white group">
                      <input
                        type="radio"
                        name="format-order-sidebar"
                        checked={extractionSettings.takeFirstPart}
                        onChange={() => setExtractionSettings(prev => ({ ...prev, takeFirstPart: true }))}
                        className="text-emerald-500 focus:ring-emerald-400/20 w-3.5 h-3.5 cursor-pointer accent-emerald-500"
                      />
                      <span className="group-hover:text-white transition-colors text-[11px]"><strong className="text-emerald-400 font-bold">Giáo viên</strong> {extractionSettings.delimiter || '-'} Môn học</span>
                    </label>
                  </div>
                </div>

                {/* Ignore List */}
                <div className="flex flex-col gap-1.5 pt-3 border-t border-teal-800/30">
                  <div className="flex items-center justify-between mb-0.5">
                    <label className="text-[10px] font-bold text-teal-200/80 uppercase tracking-wider">Từ khóa loại trừ (Ignore List)</label>
                    <button
                      type="button"
                      onClick={handleIgnoreListSave}
                      className="text-[9px] font-extrabold text-emerald-400 hover:text-emerald-300 cursor-pointer bg-teal-950/60 hover:bg-teal-950/90 px-1.5 py-0.5 rounded border border-teal-800 transition-all"
                    >
                      Cập nhật
                    </button>
                  </div>
                  <textarea
                    value={ignoreListText}
                    onChange={(e) => setIgnoreListText(e.target.value)}
                    placeholder="shcn, chào cờ, nghỉ, x..."
                    rows={2}
                    className="w-full px-2.5 py-1.5 bg-teal-950/40 border border-teal-800/80 rounded-lg text-xs focus:outline-none focus:ring-1 focus:ring-emerald-400 text-teal-100 font-medium leading-relaxed transition-all placeholder:text-teal-700 resize-none"
                  />
                </div>
              </div>
            </div>
          )}

          {/* Gợi ý / Hướng dẫn Alert at the bottom */}
          <div className="p-4 bg-teal-950/30 border border-teal-800/40 rounded-xl space-y-2" id="sidebar-tips-box">
            <h4 className="text-xs font-bold text-teal-100 flex items-center gap-1.5">
              <Sparkles className="w-3.5 h-3.5 text-emerald-400" />
              Gợi ý rà soát:
            </h4>
            <p className="text-[10px] text-teal-200/90 leading-relaxed font-medium">
              Bạn có thể tùy chỉnh danh sách từ khóa bỏ qua (Ignore list) như <strong>"Chào cờ"</strong>, <strong>"SHCN"</strong> để tránh phần mềm báo trùng nhầm.
            </p>
          </div>
        </div>

        {/* Sidebar Footer */}
        <div className="p-6 border-t border-teal-800/40 shrink-0 text-[10px] text-teal-300/80 font-bold text-center bg-teal-950/20 tracking-wider uppercase" id="sidebar-footer">
          Bảo mật tuyệt đối — Chạy 100% Offline
        </div>
      </aside>

      {/* RIGHT VIEWPORT: Interactive Action Areas and Table List */}
      <main className="flex-1 flex flex-col overflow-hidden h-full bg-slate-50/40" id="main-content-viewport">
        
        {/* Dynamic Top Header with Page Title and Status Indicators */}
        <header className="h-20 bg-white border-b border-slate-200/50 px-6 lg:px-8 flex items-center justify-between shrink-0" id="viewport-header">
          <div>
            <h2 className="text-base lg:text-lg font-extrabold text-slate-900 flex items-center gap-2 font-display" id="header-page-title">
              {step === 'upload' && 'Bắt đầu rà soát lịch dạy'}
              {step === 'mapping' && 'Cấu hình ánh xạ cột dữ liệu'}
              {step === 'results' && 'Kết quả rà soát thời khóa biểu'}
            </h2>
            <p className="text-xs text-slate-400 font-semibold uppercase tracking-wider hidden sm:block">
              {step === 'upload' && 'Tải lên tệp thời khóa biểu định dạng Excel để tiến hành kiểm tra'}
              {step === 'mapping' && 'Chỉ định cột nào chứa thông tin Ngày, Tiết và tên các Lớp học'}
              {step === 'results' && `Xem chi tiết và xuất báo cáo trùng tiết của giáo viên`}
            </p>
          </div>

          {/* Dashboard Stats / Badges & Reload Action */}
          <div className="flex items-center gap-2" id="header-badge-container">
            <button
              onClick={() => window.location.reload()}
              className="inline-flex items-center gap-2 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 hover:text-slate-900 text-xs font-extrabold rounded-xl border border-slate-200 transition-all cursor-pointer shadow-sm shrink-0"
              title="Tải lại toàn bộ trang web"
              id="reload-page-btn"
            >
              <RefreshCw className="w-3.5 h-3.5 text-slate-500" />
              <span>Tải lại trang</span>
            </button>

            {step === 'results' && (
              <>
                {conflictReportData.conflicts.length > 0 ? (
                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-rose-50 text-rose-700 text-xs font-extrabold rounded-full border border-rose-150">
                    <span className="w-1.5 h-1.5 bg-rose-600 rounded-full animate-ping"></span>
                    {conflictReportData.conflicts.length} lỗi trùng lịch
                  </span>
                ) : (
                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-50 text-emerald-700 text-xs font-extrabold rounded-full border border-emerald-150">
                    <Check className="w-3.5 h-3.5 text-emerald-600" />
                    Không trùng lịch
                  </span>
                )}

                {conflictReportData.splitIssues.length > 0 && (
                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-amber-50 text-amber-700 text-xs font-extrabold rounded-full border border-amber-200 shadow-sm">
                    <Clock className="w-3.5 h-3.5 text-amber-600" />
                    {conflictReportData.splitIssues.length} lớp bị chia tiết
                  </span>
                )}

                {conflictReportData.simultaneousMatches.length > 0 ? (
                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-indigo-50 text-indigo-700 text-xs font-extrabold rounded-full border border-indigo-200 shadow-sm">
                    <Users className="w-3.5 h-3.5 text-indigo-600" />
                    {conflictReportData.simultaneousMatches.length} ca trùng {simultaneousConfig.classQuery1} & {simultaneousConfig.classQuery2}
                  </span>
                ) : (
                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-indigo-50/70 text-indigo-700 text-xs font-extrabold rounded-full border border-indigo-200/60 shadow-sm hidden sm:inline-flex">
                    {simultaneousConfig.classQuery1} & {simultaneousConfig.classQuery2} không trùng
                  </span>
                )}

                {conflictReportData.constraintResults?.some(r => !r.satisfied) ? (
                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-rose-50 text-rose-700 text-xs font-extrabold rounded-full border border-rose-200 shadow-sm">
                    <ShieldCheck className="w-3.5 h-3.5 text-rose-600" />
                    {conflictReportData.constraintResults.filter(r => !r.satisfied).length} vi phạm ràng buộc
                  </span>
                ) : constraints.length > 0 ? (
                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-teal-50 text-teal-700 text-xs font-extrabold rounded-full border border-teal-200 shadow-sm hidden md:inline-flex">
                    <ShieldCheck className="w-3.5 h-3.5 text-teal-600" />
                    {constraints.length} ràng buộc đạt
                  </span>
                ) : null}
              </>
            )}

            {step === 'mapping' && (
              <span className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-amber-50 text-amber-700 text-xs font-extrabold rounded-full border border-amber-150">
                Chờ cấu hình
              </span>
            )}
          </div>
        </header>

        {/* Scrollable Main Area Viewports */}
        <div className="flex-1 overflow-y-auto p-6 lg:p-8 space-y-6" id="viewport-scrollable-content">
          
          <AnimatePresence mode="wait">
            
            {/* STEP 1: PRE-UPLOAD CONSTRAINTS & UPLOAD SCREEN */}
            {step === 'upload' && (
              <motion.div
                key="viewport-upload"
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -15 }}
                transition={{ duration: 0.2 }}
                className={uploadTab === 'setup_constraints' ? 'max-w-6xl mx-auto space-y-6' : 'max-w-3xl mx-auto space-y-6'}
              >
                {/* Stepper Navigation: Bước 1: Nhập yêu cầu & ràng buộc -> Bước 2: Tải file TKB lên */}
                <div className="bg-white/90 backdrop-blur-md rounded-2xl border border-slate-200/80 p-2 shadow-sm flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-2" id="workflow-stepper-bar">
                  <div className="flex items-center gap-2 p-1 bg-slate-100 rounded-xl flex-1">
                    <button
                      type="button"
                      id="tab-setup-constraints"
                      onClick={() => setUploadTab('setup_constraints')}
                      className={`flex-1 px-4 py-3 rounded-xl font-bold text-xs sm:text-sm transition-all flex items-center justify-center gap-2.5 cursor-pointer ${
                        uploadTab === 'setup_constraints'
                          ? 'bg-indigo-600 text-white shadow-md shadow-indigo-600/20'
                          : 'text-slate-600 hover:text-slate-900 hover:bg-white/60'
                      }`}
                    >
                      <ShieldCheck className="w-4 h-4 shrink-0" />
                      <span>Bước 1: Nhập Yêu Cầu Ràng Buộc</span>
                      <span className={`px-2 py-0.5 rounded-full text-[11px] font-mono font-black ${
                        uploadTab === 'setup_constraints' ? 'bg-indigo-500 text-white' : 'bg-slate-200 text-slate-700'
                      }`}>
                        {constraints.filter(c => c.isActive).length}
                      </span>
                    </button>

                    <button
                      type="button"
                      id="tab-upload-file"
                      onClick={() => setUploadTab('upload_file')}
                      className={`flex-1 px-4 py-3 rounded-xl font-bold text-xs sm:text-sm transition-all flex items-center justify-center gap-2.5 cursor-pointer ${
                        uploadTab === 'upload_file'
                          ? 'bg-teal-600 text-white shadow-md shadow-teal-600/20'
                          : 'text-slate-600 hover:text-slate-900 hover:bg-white/60'
                      }`}
                    >
                      <Upload className="w-4 h-4 shrink-0" />
                      <span>Bước 2: Tải File TKB & Kiểm Tra</span>
                    </button>
                  </div>
                </div>

                {/* BƯỚC 1: NHẬP VÀ LƯU RÀNG BUỘC TRƯỚC */}
                {uploadTab === 'setup_constraints' && (
                  <div className="space-y-6">
                    <ConstraintManager
                      constraints={constraints}
                      onSaveConstraint={handleSaveConstraint}
                      onDeleteConstraint={handleDeleteConstraint}
                      onToggleConstraint={handleToggleConstraint}
                      onResetDefaults={handleResetConstraints}
                      isPreUpload={true}
                      onProceedToUpload={() => setUploadTab('upload_file')}
                    />
                  </div>
                )}

                {/* BƯỚC 2: TẢI FILE EXCEL LÊN ĐỂ KIỂM TRA */}
                {uploadTab === 'upload_file' && (
                  <div className="space-y-6">
                    {/* Cấu hình cấp học rà soát */}
                    <div className="bg-white rounded-2xl border border-slate-200/50 p-5 shadow-sm space-y-4">
                      <div className="flex items-center gap-2 border-b border-slate-100 pb-3">
                        <div className="p-1.5 bg-teal-50 text-teal-600 rounded-lg">
                          <Settings2 className="w-4 h-4" />
                        </div>
                        <div>
                          <h3 className="text-xs font-bold text-slate-700 uppercase tracking-wider">Cấu hình cấp học & Phạm vi rà soát</h3>
                          <p className="text-[11px] text-slate-400 font-medium">Hệ thống sẽ giới hạn vùng quét trong tệp Excel theo đúng lựa chọn này</p>
                        </div>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                        {/* THPT Card */}
                        <button
                          type="button"
                          onClick={() => setCheckProfile('THPT')}
                          className={`p-4 rounded-xl text-left border transition-all cursor-pointer relative flex flex-col justify-between ${
                            checkProfile === 'THPT'
                              ? 'border-teal-500 bg-teal-50/10 ring-2 ring-teal-500/5'
                              : 'border-slate-200 hover:border-slate-300 hover:bg-slate-50/40'
                          }`}
                        >
                          <div className="flex items-center justify-between mb-1.5">
                            <span className="text-xs font-bold text-slate-800">Check THPT</span>
                            <span className={`w-4 h-4 rounded-full border flex items-center justify-center text-[10px] ${
                              checkProfile === 'THPT' ? 'border-teal-500 bg-teal-500 text-white' : 'border-slate-300'
                            }`}>
                              {checkProfile === 'THPT' && '✓'}
                            </span>
                          </div>
                          <span className="text-[14px] font-extrabold text-teal-700 font-display">Ô K5 : V100</span>
                          <span className="text-[10px] text-slate-500 font-medium mt-1 leading-normal">Dòng 5 là tên lớp. Cột K là thời gian. Kiểm tra trùng các cột lớp L đến V (dòng 10 - 100).</span>
                        </button>
     
                        {/* Tiểu học Card */}
                        <button
                          type="button"
                          onClick={() => setCheckProfile('TieuHoc')}
                          className={`p-4 rounded-xl text-left border transition-all cursor-pointer relative flex flex-col justify-between ${
                            checkProfile === 'TieuHoc'
                              ? 'border-teal-500 bg-teal-50/10 ring-2 ring-teal-500/5'
                              : 'border-slate-200 hover:border-slate-300 hover:bg-slate-50/40'
                          }`}
                        >
                          <div className="flex items-center justify-between mb-1.5">
                            <span className="text-xs font-bold text-slate-800">Check Tiểu Học</span>
                            <span className={`w-4 h-4 rounded-full border flex items-center justify-center text-[10px] ${
                              checkProfile === 'TieuHoc' ? 'border-teal-500 bg-teal-500 text-white' : 'border-slate-300'
                            }`}>
                              {checkProfile === 'TieuHoc' && '✓'}
                            </span>
                          </div>
                          <span className="text-[14px] font-extrabold text-teal-700 font-display">Ô E5 : J100</span>
                          <span className="text-[10px] text-slate-500 font-medium mt-1 leading-normal">Dòng 5 là tên lớp. Cột E là thời gian. Kiểm tra trùng các cột lớp F đến J (dòng 10 - 100).</span>
                        </button>

                        {/* Toàn bộ bảng Card */}
                        <button
                          type="button"
                          onClick={() => setCheckProfile('All')}
                          className={`p-4 rounded-xl text-left border transition-all cursor-pointer relative flex flex-col justify-between ${
                            checkProfile === 'All'
                              ? 'border-teal-500 bg-teal-50/10 ring-2 ring-teal-500/5'
                              : 'border-slate-200 hover:border-slate-300 hover:bg-slate-50/40'
                          }`}
                        >
                          <div className="flex items-center justify-between mb-1.5">
                            <span className="text-xs font-bold text-slate-800">Check Toàn Bộ</span>
                            <span className={`w-4 h-4 rounded-full border flex items-center justify-center text-[10px] ${
                              checkProfile === 'All' ? 'border-teal-500 bg-teal-500 text-white' : 'border-slate-300'
                            }`}>
                              {checkProfile === 'All' && '✓'}
                            </span>
                          </div>
                          <span className="text-[14px] font-extrabold text-teal-700 font-display">Tự động</span>
                          <span className="text-[10px] text-slate-500 font-medium mt-1 leading-normal">Quét tất cả các cột chứa thông tin lớp học</span>
                        </button>
                      </div>
                    </div>

                    {/* Drag and Drop area - Đơn giản và nhỏ gọn */}
                    <div 
                      onDragOver={handleDragOver}
                      onDragLeave={handleDragLeave}
                      onDrop={handleDrop}
                      className={`bg-white border-2 border-dashed rounded-xl py-6 px-4 text-center transition-all flex flex-col items-center justify-center cursor-pointer group shadow-xs ${
                        isDragging 
                          ? 'border-teal-500 bg-teal-50/30 ring-4 ring-teal-500/10' 
                          : 'border-slate-300 hover:border-teal-500 hover:bg-slate-50/60'
                      }`}
                      onClick={() => fileInputRef.current?.click()}
                      id="drag-drop-zone"
                    >
                      <input
                        type="file"
                        ref={fileInputRef}
                        onChange={handleFileChange}
                        accept=".xlsx, .xls"
                        className="hidden"
                      />
                      <div className={`p-2.5 bg-teal-50 text-teal-600 rounded-xl mb-2 transition-transform duration-200 ${
                        isDragging ? 'bg-teal-100 scale-110' : 'group-hover:scale-105'
                      }`}>
                        <Upload className="w-5 h-5" />
                      </div>
                      <div className="text-sm font-bold text-slate-800">
                        Kéo thả tệp Excel hoặc <span className="text-teal-600 underline underline-offset-2">bấm vào đây để chọn tệp</span>
                      </div>
                      <p className="text-[11px] text-slate-400 mt-1">
                        Hỗ trợ định dạng .xlsx, .xls
                      </p>
                    </div>

                    {/* Instructions card */}
                    <div className="bg-white rounded-2xl border border-slate-200/50 p-6 shadow-sm space-y-4">
                      <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1.5">
                        <HelpCircle className="w-4 h-4 text-teal-500 shrink-0" />
                        Hướng dẫn chuẩn bị bảng biểu Excel:
                      </h4>
                      <div className="space-y-3.5">
                        <div className="flex items-start gap-3">
                          <div className="w-5 h-5 bg-teal-50 text-teal-600 rounded-md flex items-center justify-center text-xs font-bold shrink-0 mt-0.5">1</div>
                          <p className="text-xs text-slate-600 leading-relaxed font-medium">
                            <strong className="text-slate-900">Dòng 1 (Tiêu đề)</strong>: Chứa thông tin <span className="font-semibold text-slate-700">Tên lớp</span> (ví dụ: <code className="px-1.5 py-0.5 bg-slate-100 text-slate-700 font-mono rounded">10.1 TN</code>, <code className="px-1.5 py-0.5 bg-slate-100 text-slate-700 font-mono rounded">11.2 XH</code>...) xếp ngang từ cột F trở đi.
                          </p>
                        </div>
                        <div className="flex items-start gap-3">
                          <div className="w-5 h-5 bg-teal-50 text-teal-600 rounded-md flex items-center justify-center text-xs font-bold shrink-0 mt-0.5">2</div>
                          <p className="text-xs text-slate-600 leading-relaxed font-medium">
                            <strong className="text-slate-900">Các cột đầu</strong>: Nên bao gồm thông tin <span className="font-semibold text-slate-700">Thứ (Thứ Hai, Thứ Ba...)</span> và <span className="font-semibold text-slate-700">Tiết học (Tiết 1, Tiết 2...)</span>.
                          </p>
                        </div>
                        <div className="flex items-start gap-3">
                          <div className="w-5 h-5 bg-teal-50 text-teal-600 rounded-md flex items-center justify-center text-xs font-bold shrink-0 mt-0.5">3</div>
                          <p className="text-xs text-slate-600 leading-relaxed font-medium">
                            <strong className="text-slate-900">Dòng 5 trở xuống</strong>: Là dữ liệu thời khóa biểu thực tế. Mỗi ô điền dạng <code className="text-teal-600 font-bold">Môn học - Giáo viên</code> hoặc chỉ <code className="text-teal-600 font-bold">Tên giáo viên</code>.
                          </p>
                        </div>
                      </div>

                      <div className="mt-2 pt-4 border-t border-slate-100 flex justify-center">
                        <button
                          onClick={generateSampleExcel}
                          className="inline-flex items-center gap-1.5 text-xs font-bold text-teal-700 hover:text-teal-850 bg-teal-50 hover:bg-teal-100/80 px-4 py-2.5 rounded-xl transition-all border border-teal-100/50 cursor-pointer shadow-sm shadow-teal-100/20"
                          id="guide-download-sample-btn"
                        >
                          <FileDown className="w-4 h-4" />
                          Tải mẫu Excel chuẩn từ Dòng 5
                        </button>
                      </div>
                    </div>
                  </div>
                )}
              </motion.div>
            )}

            {/* STEP 2: COLUMN MAPPING SETUP */}
            {step === 'mapping' && parseResult && (
              <motion.div
                key="viewport-mapping"
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -15 }}
                transition={{ duration: 0.2 }}
                className="max-w-5xl mx-auto space-y-6"
              >
                <MappingSetup
                  headers={parseResult.headers}
                  previewRows={parseResult.rows.slice(0, 5)}
                  mappings={mappings}
                  onMappingChange={setMappings}
                  onProceed={() => setStep('results')}
                  range={activeRange}
                />
              </motion.div>
            )}

            {/* STEP 3: RESULTS & CONFLICT LIST */}
            {step === 'results' && parseResult && (
              <motion.div
                key="viewport-results"
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -15 }}
                transition={{ duration: 0.2 }}
                className="space-y-6"
              >
                <div className="w-full">
                  <ConflictReport
                    conflicts={conflictReportData.conflicts}
                    splitIssues={conflictReportData.splitIssues}
                    simultaneousMatches={conflictReportData.simultaneousMatches}
                    simultaneousConfig={simultaneousConfig}
                    onUpdateSimultaneousConfig={handleUpdateSimultaneousConfig}
                    exemptions={exemptions}
                    totalCellsChecked={conflictReportData.totalCellsChecked}
                    totalTeachersFound={conflictReportData.totalTeachers}
                    constraints={constraints}
                    constraintResults={conflictReportData.constraintResults}
                    availableTeachers={conflictReportData.availableTeachers}
                    availableClasses={conflictReportData.availableClasses}
                    onSaveConstraint={handleSaveConstraint}
                    onDeleteConstraint={handleDeleteConstraint}
                    onToggleConstraint={handleToggleConstraint}
                    onResetConstraints={handleResetConstraints}
                  />
                </div>
              </motion.div>
            )}

          </AnimatePresence>

        </div>
      </main>

    </div>
  );
}
