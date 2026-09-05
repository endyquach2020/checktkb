import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  ExemptionPair, 
  ColumnMapping, 
  ExtractionSettings, 
  ParseResult,
  CellRange
} from './types';
import { 
  detectConflicts, 
  generateSampleExcel, 
  extractTeacher, 
  normalizeTeacherName,
  parseExcelRange,
  indexToColName
} from './utils/excelParser';
import ExemptionSettingsComponent from './components/ExemptionSettings';
import MappingSetup from './components/MappingSetup';
import ConflictReport from './components/ConflictReport';
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
  Sparkles
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

  // Teacher Name Extraction Settings
  const [extractionSettings, setExtractionSettings] = useState<ExtractionSettings>({
    delimiter: '-',
    takeFirstPart: false, // Default Subject - Teacher (take second part)
    ignoreList: ['shcn', 'chào cờ', 'chao co', 'sinh hoạt', 'nghỉ', 'nghi', 'chủ nhiệm', 'x', '-', 'chủ nhật', 'cn'],
  });

  // State for raw ignore list text in settings panel
  const [ignoreListText, setIgnoreListText] = useState(extractionSettings.ignoreList.join(', '));

  // Excel Cell Range Selection
  const [useCustomRange, setUseCustomRange] = useState<boolean>(true); // Enabled by default as requested
  const [customRange, setCustomRange] = useState<string>('F39:V100'); // F39:V100 default
  const [rangeInputVal, setRangeInputVal] = useState<string>('F39:V100');

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

        const activeSheet = activeSheetName || sheetNames[0];
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

        // Class names (headers) are on Row 1 (0-indexed 0)
        const rawHeaders = cleanedData[0] || [];
        
        // Timetable data starts from Row 5 (0-indexed 4) down to the end of the data
        const rawContentRows = cleanedData.slice(4);

        if (cleanedData.length < 5) {
          alert('Tệp Excel phải có ít nhất 5 dòng (Dòng 1 chứa tên lớp, Dòng 5 trở đi chứa dữ liệu lịch học)!');
          return;
        }

        // Fill empty headers with Cột index
        const headers = rawHeaders.map((h, idx) => h || `Cột ${idx + 1}`);

        // Automatically detect column mappings
        let detectedDay = false;
        let detectedPeriod = false;

        const autoMappings = headers.map((header, index) => {
          const hLower = header.toLowerCase().trim();
          let role: 'day' | 'period' | 'class' | 'skip' = 'class';

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

          return {
            index,
            header: header || `Cột ${index + 1}`,
            role: role as any,
          };
        });

        // Fallbacks if not auto-detected
        if (!detectedDay && headers.length > 0) {
          autoMappings[0].role = 'day' as any;
        }
        if (!detectedPeriod && headers.length > 1) {
          autoMappings[1].role = 'period' as any;
        }

        setParseResult({
          headers,
          rows: rawContentRows,
          sheetNames,
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
    const conflicts = detectConflicts(rows, mappings, exemptions, extractionSettings);

    // Calculate statistics
    let cellCheckCount = 0;
    const teacherSet = new Set<string>();
    const classCols = mappings.filter(m => m.role === 'class');

    rows.forEach(row => {
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

    return {
      conflicts,
      totalCellsChecked: cellCheckCount,
      totalTeachers: teacherSet.size
    };
  }, [parseResult, mappings, exemptions, extractionSettings]);

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
      <aside className="w-full lg:w-80 bg-white border-b lg:border-b-0 lg:border-r border-slate-200/60 flex flex-col shrink-0 h-auto lg:h-full overflow-y-auto" id="sidebar-container">
        {/* Branding Header */}
        <div className="p-6 border-b border-slate-100 flex items-center justify-between shrink-0 bg-slate-50/30" id="sidebar-header">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-indigo-600 text-white rounded-xl shadow-md shadow-indigo-600/15 flex items-center justify-center">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <h1 className="text-sm font-extrabold text-slate-900 tracking-tight flex items-center gap-1.5 font-display">
                CheckSched Pro
                <span className="text-[9px] font-bold bg-indigo-50 text-indigo-700 px-1.5 py-0.5 rounded-md border border-indigo-100">
                  v2.5
                </span>
              </h1>
              <p className="text-[10px] text-slate-400 font-semibold uppercase tracking-wider">Hệ thống rà soát trùng lịch</p>
            </div>
          </div>
        </div>

        {/* Sidebar Body Content */}
        <div className="p-6 flex-1 space-y-6" id="sidebar-body">
          {/* Active File State Indicator */}
          <div className="space-y-2.5" id="sidebar-file-status">
            <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Tệp tin đang mở</div>
            {file ? (
              <div className="p-4 bg-slate-50/70 border border-slate-200/50 rounded-xl space-y-3">
                <div className="flex items-start gap-2.5">
                  <div className="p-1.5 bg-emerald-100 text-emerald-700 rounded-lg shrink-0 mt-0.5">
                    <Check className="w-3.5 h-3.5" />
                  </div>
                  <div className="min-w-0 flex-1">
                    <p className="text-xs font-bold text-slate-800 truncate" title={file.name}>{file.name}</p>
                    <p className="text-[10px] text-slate-400 font-semibold mt-0.5">Dung lượng: {(file.size / 1024).toFixed(1)} KB</p>
                  </div>
                </div>

                {/* Sheet Switcher */}
                {parseResult && parseResult.sheetNames.length > 1 && (
                  <div className="space-y-1.5 pt-2 border-t border-slate-200/50">
                    <div className="text-[9px] font-bold text-slate-400 uppercase tracking-wider">Chọn Sheet Excel</div>
                    <div className="flex flex-col gap-1">
                      {parseResult.sheetNames.map((name) => (
                        <button
                          key={name}
                          onClick={() => handleSheetChange(name)}
                          className={`w-full px-2.5 py-1.5 rounded-lg text-left text-xs font-semibold transition-all flex items-center justify-between ${
                            parseResult.activeSheetName === name
                              ? 'bg-indigo-50 text-indigo-700 font-bold border-l-2 border-indigo-600 pl-2'
                              : 'text-slate-500 hover:bg-slate-100/70 hover:text-slate-800'
                          }`}
                        >
                          <span className="truncate">{name}</span>
                          {parseResult.activeSheetName === name && <span className="w-1.5 h-1.5 bg-indigo-600 rounded-full shrink-0"></span>}
                        </button>
                      ))}
                    </div>
                  </div>
                )}

                <button
                  onClick={handleReset}
                  className="w-full flex items-center justify-center gap-1.5 py-1.5 border border-slate-200 bg-white hover:bg-slate-50 text-[11px] font-bold text-slate-600 rounded-lg transition-colors cursor-pointer"
                  id="sidebar-change-file-btn"
                >
                  <RefreshCw className="w-3 h-3 text-slate-400" />
                  Chọn tệp tin khác
                </button>
              </div>
            ) : (
              <div className="p-4 bg-slate-50/50 border border-dashed border-slate-200 rounded-xl text-center">
                <p className="text-xs text-slate-400 font-medium italic">Chưa tải tệp tin lên</p>
              </div>
            )}
          </div>

          {/* Cấu hình Vùng Rà Soát (Excel Cell Range) */}
          <div className="space-y-2.5 pt-4 border-t border-slate-150" id="sidebar-range-config">
            <div className="text-[10px] font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1.5">
              <Settings2 className="w-3.5 h-3.5 text-indigo-500" />
              Phạm vi rà soát tự động
            </div>
            
            <div className="bg-indigo-50/30 p-4 rounded-xl border border-indigo-100/40 space-y-2">
              <div className="text-[11px] text-slate-600 space-y-2 font-medium leading-relaxed">
                <p className="flex items-center gap-1.5">
                  <span className="w-1.5 h-1.5 bg-indigo-600 rounded-full"></span>
                  Tên các lớp: <strong className="text-indigo-950 font-bold font-display">Dòng 1</strong>
                </p>
                <p className="flex items-center gap-1.5">
                  <span className="w-1.5 h-1.5 bg-indigo-600 rounded-full"></span>
                  Dữ liệu rà soát: <strong className="text-indigo-950 font-bold font-display">Dòng 5 đến hết file</strong>
                </p>
                <p className="text-[10px] text-slate-400 italic pt-1 border-t border-indigo-100/30 leading-relaxed">
                  Hệ thống tự động phân tích và rà soát trùng lặp tức thì đối với mọi lớp học được tìm thấy.
                </p>
              </div>
            </div>
          </div>

          {/* Quick Exemption Rules directly inside sidebar for premium interaction */}
          {step === 'results' && (
            <div className="space-y-3 pt-4 border-t border-slate-150" id="sidebar-quick-rules">
              <div className="flex items-center justify-between">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Cấu hình rà soát</span>
              </div>

              {/* Minimized Delimiter Form */}
              <div className="bg-slate-50 p-3 rounded-xl border border-slate-200/50 space-y-2.5">
                <div className="flex items-center justify-between gap-2">
                  <span className="text-[11px] font-semibold text-slate-600">Ký tự phân tách:</span>
                  <input
                    type="text"
                    value={extractionSettings.delimiter}
                    onChange={(e) => setExtractionSettings(prev => ({ ...prev, delimiter: e.target.value }))}
                    className="w-10 text-center py-0.5 bg-white border border-slate-200 rounded text-xs font-mono font-bold text-slate-800 focus:outline-none focus:ring-1 focus:ring-indigo-500"
                    maxLength={3}
                  />
                </div>
                <div className="text-[9px] text-slate-400 leading-normal">
                  Ví dụ: <code className="bg-white px-1 py-0.5 border border-slate-200/50 rounded">Toán - Thầy Hải</code>
                </div>
              </div>

              {/* Fast Sample Download Button */}
              <button
                onClick={generateSampleExcel}
                className="w-full flex items-center justify-center gap-1.5 py-2 px-3 bg-slate-50 hover:bg-slate-100 text-slate-700 text-xs font-bold rounded-lg border border-slate-200 transition-colors cursor-pointer"
                id="sidebar-download-sample-btn"
              >
                <FileDown className="w-3.5 h-3.5 text-slate-500" />
                Tải file Excel mẫu
              </button>
            </div>
          )}

          {/* Gợi ý / Hướng dẫn Alert at the bottom */}
          <div className="p-4 bg-indigo-50/50 border border-indigo-100/60 rounded-xl space-y-2" id="sidebar-tips-box">
            <h4 className="text-xs font-bold text-indigo-950 flex items-center gap-1.5">
              <Sparkles className="w-3.5 h-3.5 text-indigo-600" />
              Gợi ý rà soát:
            </h4>
            <p className="text-[10px] text-indigo-800/90 leading-relaxed font-medium">
              Bạn có thể tùy chỉnh danh sách từ khóa bỏ qua (Ignore list) như <strong>"Chào cờ"</strong>, <strong>"SHCN"</strong> để tránh phần mềm báo trùng nhầm.
            </p>
          </div>
        </div>

        {/* Sidebar Footer */}
        <div className="p-6 border-t border-slate-100 shrink-0 text-[10px] text-slate-400 font-semibold text-center bg-slate-50/40 tracking-wide uppercase" id="sidebar-footer">
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

          {/* Dashboard Stats / Badges */}
          <div className="flex items-center gap-2" id="header-badge-container">
            {step === 'results' && (
              <>
                {conflictReportData.conflicts.length > 0 ? (
                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-rose-50 text-rose-700 text-xs font-extrabold rounded-full border border-rose-150">
                    <span className="w-1.5 h-1.5 bg-rose-600 rounded-full animate-ping"></span>
                    Phát hiện {conflictReportData.conflicts.length} lỗi trùng lịch
                  </span>
                ) : (
                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-50 text-emerald-700 text-xs font-extrabold rounded-full border border-emerald-150">
                    <Check className="w-3.5 h-3.5 text-emerald-600" />
                    Thời khóa biểu hợp lệ
                  </span>
                )}
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
            
            {/* STEP 1: UPLOAD SCREEN */}
            {step === 'upload' && (
              <motion.div
                key="viewport-upload"
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -15 }}
                transition={{ duration: 0.2 }}
                className="max-w-2xl mx-auto space-y-6"
              >
                {/* Drag and Drop area */}
                <div 
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onDrop={handleDrop}
                  className={`bg-white border-2 border-dashed rounded-2xl p-10 text-center transition-all flex flex-col items-center justify-center min-h-[340px] cursor-pointer group shadow-sm ${
                    isDragging 
                      ? 'border-indigo-500 bg-indigo-50/20 ring-4 ring-indigo-500/5 shadow-indigo-100' 
                      : 'border-slate-200 hover:border-indigo-400 hover:bg-slate-50/30'
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
                  <div className={`p-4 bg-slate-50 text-slate-400 rounded-xl mb-4 transition-all duration-300 ${
                    isDragging ? 'bg-indigo-100 text-indigo-600 scale-105 shadow-sm' : 'group-hover:bg-indigo-50 group-hover:text-indigo-600 group-hover:scale-105'
                  }`}>
                    <Upload className="w-8 h-8" />
                  </div>
                  <h3 className="text-base font-bold text-slate-800 font-display">Tải lên tệp Excel thời khóa biểu</h3>
                  <p className="text-xs text-slate-500 mt-1 max-w-sm leading-relaxed">
                    Kéo thả tệp tin <strong className="text-slate-700">.xlsx</strong> hoặc <strong className="text-slate-700">.xls</strong> vào đây, hoặc click để chọn từ thiết bị của bạn.
                  </p>

                  <div className="flex flex-wrap items-center justify-center gap-4 mt-6 text-[11px] text-slate-400 font-bold uppercase tracking-wider">
                    <span className="flex items-center gap-1 text-slate-500">
                      <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full"></span> Nhận diện tự động
                    </span>
                    <span className="flex items-center gap-1 text-slate-500">
                      <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full"></span> Cấu hình lớp ghép
                    </span>
                    <span className="flex items-center gap-1 text-slate-500">
                      <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full"></span> Lọc theo Thứ
                    </span>
                  </div>
                </div>

                {/* Instructions card */}
                <div className="bg-white rounded-2xl border border-slate-200/50 p-6 shadow-sm space-y-4">
                  <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1.5">
                    <HelpCircle className="w-4 h-4 text-indigo-500 shrink-0" />
                    Hướng dẫn chuẩn bị bảng biểu Excel:
                  </h4>
                  <div className="space-y-3.5">
                    <div className="flex items-start gap-3">
                      <div className="w-5 h-5 bg-indigo-50 text-indigo-600 rounded-md flex items-center justify-center text-xs font-bold shrink-0 mt-0.5">1</div>
                      <p className="text-xs text-slate-600 leading-relaxed font-medium">
                        <strong className="text-slate-900">Dòng 1 (Tiêu đề)</strong>: Chứa thông tin <span className="font-semibold text-slate-700">Tên lớp</span> (ví dụ: <code className="px-1.5 py-0.5 bg-slate-100 text-slate-700 font-mono rounded">10.1 TN</code>, <code className="px-1.5 py-0.5 bg-slate-100 text-slate-700 font-mono rounded">11.2 XH</code>...) xếp ngang từ cột F trở đi.
                      </p>
                    </div>
                    <div className="flex items-start gap-3">
                      <div className="w-5 h-5 bg-indigo-50 text-indigo-600 rounded-md flex items-center justify-center text-xs font-bold shrink-0 mt-0.5">2</div>
                      <p className="text-xs text-slate-600 leading-relaxed font-medium">
                        <strong className="text-slate-900">Các cột đầu</strong>: Nên bao gồm thông tin <span className="font-semibold text-slate-700">Thứ (Thứ Hai, Thứ Ba...)</span> và <span className="font-semibold text-slate-700">Tiết học (Tiết 1, Tiết 2...)</span>.
                      </p>
                    </div>
                    <div className="flex items-start gap-3">
                      <div className="w-5 h-5 bg-indigo-50 text-indigo-600 rounded-md flex items-center justify-center text-xs font-bold shrink-0 mt-0.5">3</div>
                      <p className="text-xs text-slate-600 leading-relaxed font-medium">
                        <strong className="text-slate-900">Dòng 5 trở xuống</strong>: Là dữ liệu thời khóa biểu thực tế. Mỗi ô điền dạng <code className="text-indigo-600 font-bold">Môn học - Giáo viên</code> hoặc chỉ <code className="text-indigo-600 font-bold">Tên giáo viên</code>.
                      </p>
                    </div>
                  </div>

                  <div className="mt-2 pt-4 border-t border-slate-100 flex justify-center">
                    <button
                      onClick={generateSampleExcel}
                      className="inline-flex items-center gap-1.5 text-xs font-bold text-indigo-600 hover:text-indigo-700 bg-indigo-50 hover:bg-indigo-100/80 px-4 py-2.5 rounded-xl transition-all border border-indigo-100/50 cursor-pointer shadow-sm shadow-indigo-100/20"
                      id="guide-download-sample-btn"
                    >
                      <FileDown className="w-4 h-4" />
                      Tải mẫu Excel chuẩn từ Dòng 5
                    </button>
                  </div>
                </div>
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
                  range={useCustomRange ? parseExcelRange(customRange) : null}
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
                {/* Responsive configuration panels inside viewport */}
                <div className="grid grid-cols-1 xl:grid-cols-3 gap-6 items-start">
                  
                  {/* Side settings cards in Results */}
                  <div className="xl:col-span-1 space-y-6">
                    
                    {/* Active Controls Details */}
                    <div className="bg-white rounded-2xl border border-slate-200/50 p-5 shadow-sm space-y-4">
                      <div className="border-b border-slate-100 pb-3">
                        <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-1.5">
                          <Settings2 className="w-4 h-4 text-indigo-500" />
                          Trích xuất & loại trừ
                        </h3>
                      </div>

                      {/* Delimiter */}
                      <div className="flex flex-col gap-1.5">
                        <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Ký tự phân tách (Môn - Giáo viên)</label>
                        <input
                          type="text"
                          value={extractionSettings.delimiter}
                          onChange={(e) => setExtractionSettings(prev => ({ ...prev, delimiter: e.target.value }))}
                          placeholder="Ví dụ: -"
                          className="w-full px-3.5 py-2 bg-slate-50 border border-slate-200/60 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500/10 font-mono font-bold text-slate-700 transition-all"
                          maxLength={5}
                        />
                      </div>

                      {/* Info order */}
                      <div className="space-y-2">
                        <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider block">Thứ tự thông tin trong ô</label>
                        <div className="space-y-2">
                          <label className="flex items-center gap-2.5 cursor-pointer text-xs font-medium text-slate-600 hover:text-slate-800 group">
                            <input
                              type="radio"
                              name="format-order-right"
                              checked={!extractionSettings.takeFirstPart}
                              onChange={() => setExtractionSettings(prev => ({ ...prev, takeFirstPart: false }))}
                              className="text-indigo-600 focus:ring-indigo-500/20 w-4 h-4 cursor-pointer"
                            />
                            <span className="group-hover:text-slate-900 transition-colors">Môn học {extractionSettings.delimiter || '-'} <strong className="text-indigo-600 font-bold">Giáo viên</strong></span>
                          </label>
                          <label className="flex items-center gap-2.5 cursor-pointer text-xs font-medium text-slate-600 hover:text-slate-800 group">
                            <input
                              type="radio"
                              name="format-order-right"
                              checked={extractionSettings.takeFirstPart}
                              onChange={() => setExtractionSettings(prev => ({ ...prev, takeFirstPart: true }))}
                              className="text-indigo-600 focus:ring-indigo-500/20 w-4 h-4 cursor-pointer"
                            />
                            <span className="group-hover:text-slate-900 transition-colors"><strong className="text-indigo-600 font-bold">Giáo viên</strong> {extractionSettings.delimiter || '-'} Môn học</span>
                          </label>
                        </div>
                      </div>

                      {/* Ignore List */}
                      <div className="flex flex-col gap-1.5 pt-4 border-t border-slate-100">
                        <div className="flex items-center justify-between mb-1">
                          <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Từ khóa loại trừ (Ignore List)</label>
                          <button
                            type="button"
                            onClick={handleIgnoreListSave}
                            className="text-[10px] font-extrabold text-indigo-600 hover:text-indigo-700 cursor-pointer bg-indigo-50 hover:bg-indigo-100/70 px-2 py-1 rounded-md border border-indigo-100/50 transition-all"
                          >
                            Cập nhật
                          </button>
                        </div>
                        <textarea
                          value={ignoreListText}
                          onChange={(e) => setIgnoreListText(e.target.value)}
                          placeholder="shcn, chào cờ, nghỉ, x..."
                          rows={2}
                          className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200/60 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-indigo-500/10 text-slate-600 font-medium leading-relaxed transition-all"
                        />
                        <span className="text-[10px] text-slate-400 font-medium leading-normal pt-1">
                          Các từ khóa giúp hệ thống tự động bỏ qua tiết sinh hoạt dưới cờ, SHCN để tránh cảnh báo nhầm.
                        </span>
                      </div>
                    </div>

                  </div>

                  {/* Main report table card on the right */}
                  <div className="xl:col-span-2">
                    <ConflictReport
                      conflicts={conflictReportData.conflicts}
                      exemptions={exemptions}
                      totalCellsChecked={conflictReportData.totalCellsChecked}
                      totalTeachersFound={conflictReportData.totalTeachers}
                    />
                  </div>

                </div>
              </motion.div>
            )}

          </AnimatePresence>

        </div>
      </main>

    </div>
  );
}
