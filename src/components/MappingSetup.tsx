import { useEffect } from 'react';
import { ColumnMapping, ColumnRole, CellRange } from '../types';
import { Table, ArrowRight, CheckCircle2, AlertCircle, Settings2 } from 'lucide-react';

interface MappingSetupProps {
  headers: string[];
  previewRows: string[][];
  mappings: ColumnMapping[];
  onMappingChange: (newMappings: ColumnMapping[]) => void;
  onProceed: () => void;
  range?: CellRange | null;
}

export default function MappingSetup({
  headers,
  previewRows,
  mappings,
  onMappingChange,
  onProceed,
  range,
}: MappingSetupProps) {
  
  // Run auto-detection whenever headers change
  useEffect(() => {
    if (headers.length > 0 && mappings.length === 0) {
      let detectedDay = false;
      let detectedPeriod = false;

      const autoMappings = headers.map((header, index) => {
        const hLower = header.toLowerCase().trim();
        let role: ColumnRole = 'class';

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

        // If range limits are active, skip class columns outside the range
        if (range) {
          if (index < range.startCol || index > range.endCol) {
            // Keep as day/period, but skip if it was marked as a class column
            if (role === 'class') {
              role = 'skip';
            }
          }
        }

        return {
          index,
          header: header || `Cột ${index + 1}`,
          role,
        };
      });

      // If we didn't detect any day/period but have at least 2 columns, assign them defaults
      if (!detectedDay && headers.length > 0) {
        autoMappings[0].role = 'day';
      }
      if (!detectedPeriod && headers.length > 1) {
        autoMappings[1].role = 'period';
      }

      onMappingChange(autoMappings);
    }
  }, [headers, mappings, onMappingChange, range]);

  const updateColumnRole = (index: number, role: ColumnRole) => {
    // If setting to 'day' or 'period', make sure others are unset to avoid duplicates
    const updated = mappings.map(m => {
      if (m.index === index) {
        return { ...m, role };
      }
      if ((role === 'day' && m.role === 'day') || (role === 'period' && m.role === 'period')) {
        return { ...m, role: 'class' as ColumnRole }; // Fallback other to class
      }
      return m;
    });
    onMappingChange(updated);
  };

  const getRoleBadgeStyle = (role: ColumnRole) => {
    switch (role) {
      case 'day':
        return 'bg-blue-50 text-blue-700 border-blue-200';
      case 'period':
        return 'bg-purple-50 text-purple-700 border-purple-200';
      case 'class':
        return 'bg-emerald-50 text-emerald-700 border-emerald-200';
      case 'skip':
        return 'bg-slate-50 text-slate-400 border-slate-200';
    }
  };

  const getRoleLabel = (role: ColumnRole) => {
    switch (role) {
      case 'day': return 'Cột Thứ';
      case 'period': return 'Cột Tiết học / Giờ';
      case 'class': return 'Lớp học (Kiểm tra)';
      case 'skip': return 'Bỏ qua không kiểm';
    }
  };

  // Validation
  const hasDay = mappings.some(m => m.role === 'day');
  const hasPeriod = mappings.some(m => m.role === 'period');
  const classCount = mappings.filter(m => m.role === 'class').length;
  const isValid = classCount > 0;

  return (
    <div id="mapping-setup-container" className="bg-white rounded-xl border border-slate-100 p-6 shadow-sm">
      <div className="flex items-center gap-2.5 mb-5 border-b border-slate-100 pb-4">
        <div className="p-2 bg-teal-50 text-teal-600 rounded-lg">
          <Settings2 className="w-5 h-5" />
        </div>
        <div>
          <h3 className="text-lg font-bold text-slate-800">
            Cấu hình ánh xạ cột dữ liệu Excel
          </h3>
          <p className="text-sm text-slate-500 mt-0.5">
            Xác nhận hoặc điều chỉnh các cột thông tin để phần mềm trích xuất chính xác.
          </p>
        </div>
      </div>

      {/* Validation Banner */}
      <div className="mb-6 grid grid-cols-1 md:grid-cols-3 gap-3">
        <div className={`p-3.5 rounded-lg border flex items-start gap-2.5 ${hasDay ? 'bg-emerald-50/50 border-emerald-100' : 'bg-amber-50/50 border-amber-100'}`}>
          {hasDay ? (
            <CheckCircle2 className="w-5 h-5 text-emerald-500 shrink-0" />
          ) : (
            <AlertCircle className="w-5 h-5 text-amber-500 shrink-0" />
          )}
          <div>
            <h4 className="text-sm font-semibold text-slate-700">Cột Thứ (Ngày)</h4>
            <p className="text-xs text-slate-500 mt-0.5">
              {hasDay 
                ? `Đã nhận diện: Cột "${mappings.find(m => m.role === 'day')?.header}"`
                : 'Khuyên dùng cột này để phân loại lịch biểu theo ngày.'}
            </p>
          </div>
        </div>

        <div className={`p-3.5 rounded-lg border flex items-start gap-2.5 ${hasPeriod ? 'bg-emerald-50/50 border-emerald-100' : 'bg-amber-50/50 border-amber-100'}`}>
          {hasPeriod ? (
            <CheckCircle2 className="w-5 h-5 text-emerald-500 shrink-0" />
          ) : (
            <AlertCircle className="w-5 h-5 text-amber-500 shrink-0" />
          )}
          <div>
            <h4 className="text-sm font-semibold text-slate-700">Cột Tiết học / Giờ</h4>
            <p className="text-xs text-slate-500 mt-0.5">
              {hasPeriod 
                ? `Đã nhận diện: Cột "${mappings.find(m => m.role === 'period')?.header}"`
                : 'Khuyên dùng cột này để phân chia tiết dạy.'}
            </p>
          </div>
        </div>

        <div className={`p-3.5 rounded-lg border flex items-start gap-2.5 ${isValid ? 'bg-emerald-50/50 border-emerald-100' : 'bg-rose-50/50 border-rose-100'}`}>
          {isValid ? (
            <CheckCircle2 className="w-5 h-5 text-emerald-500 shrink-0" />
          ) : (
            <AlertCircle className="w-5 h-5 text-rose-500 shrink-0" />
          )}
          <div>
            <h4 className="text-sm font-semibold text-slate-700">Số lượng Lớp học</h4>
            <p className="text-xs text-slate-500 mt-0.5">
              {isValid 
                ? `Đã chọn ${classCount} cột lớp để kiểm tra`
                : 'Cần chọn ít nhất 1 cột chứa thông tin Lớp học.'}
            </p>
          </div>
        </div>
      </div>

      {/* Grid mappings editor */}
      <h4 className="text-sm font-semibold text-slate-700 mb-3 flex items-center gap-1.5">
        <Table className="w-4 h-4 text-teal-500" />
        Danh sách các cột trong file Excel của bạn:
      </h4>
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3 mb-6">
        {mappings.map((mapping) => (
          <div 
            key={mapping.index}
            className={`p-3 rounded-lg border transition-all flex flex-col justify-between ${
              mapping.role !== 'skip' ? 'bg-white border-slate-200 shadow-sm' : 'bg-slate-50/50 border-slate-100 opacity-70'
            }`}
            id={`mapping-card-${mapping.index}`}
          >
            <div>
              <div className="flex items-center justify-between gap-2 mb-2">
                <span className="text-xs text-slate-400 font-mono">Cột #{mapping.index + 1}</span>
                <span className={`text-[11px] font-medium px-2 py-0.5 rounded-full border ${getRoleBadgeStyle(mapping.role)}`}>
                  {getRoleLabel(mapping.role)}
                </span>
              </div>
              <h5 className="text-sm font-bold text-slate-800 line-clamp-1 mb-3" title={mapping.header}>
                {mapping.header || <span className="text-slate-400 italic">Trống</span>}
              </h5>
            </div>

            <div className="mt-auto">
              <label className="block text-[11px] font-semibold text-slate-400 uppercase tracking-wider mb-1">Vai trò cột</label>
              <select
                value={mapping.role}
                onChange={(e) => updateColumnRole(mapping.index, e.target.value as ColumnRole)}
                className="w-full text-xs px-2.5 py-1.5 bg-slate-50 border border-slate-200 rounded-lg text-slate-700 font-medium focus:outline-none focus:ring-2 focus:ring-teal-500/20 focus:border-teal-500 transition-colors cursor-pointer"
                id={`mapping-select-${mapping.index}`}
              >
                <option value="class">📚 Lớp học (Kiểm tra trùng)</option>
                <option value="day">📅 Thứ / Ngày</option>
                <option value="period">⏱️ Tiết dạy / Giờ</option>
                <option value="skip">🚫 Bỏ qua (Không kiểm tra)</option>
              </select>
            </div>
          </div>
        ))}
      </div>

      {/* Raw Data Preview */}
      {previewRows.length > 0 && (
        <div className="mb-6">
          <h4 className="text-sm font-semibold text-slate-700 mb-2">Xem trước dữ liệu file đã tải (Tối đa 5 dòng đầu):</h4>
          <div className="overflow-x-auto border border-slate-100 rounded-lg shadow-inner max-h-[220px]">
            <table className="w-full text-left text-xs text-slate-600 font-sans border-collapse" id="preview-raw-table">
              <thead className="bg-slate-50 border-b border-slate-200 sticky top-0 font-medium text-slate-500">
                <tr>
                  {headers.map((h, i) => (
                    <th key={i} className="px-3 py-2 border-r border-slate-200 last:border-0 min-w-[120px] bg-slate-50">
                      <div className="truncate font-bold text-slate-700">{h || `Cột ${i+1}`}</div>
                      <div className="text-[10px] text-slate-400 mt-0.5 font-normal">
                        ({mappings[i]?.role ? getRoleLabel(mappings[i].role) : 'Chưa gán'})
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 bg-white">
                {previewRows.map((row, rIndex) => (
                  <tr key={rIndex} className="hover:bg-slate-50/50">
                    {headers.map((_, cIndex) => (
                      <td key={cIndex} className="px-3 py-2 border-r border-slate-100 last:border-0 truncate max-w-[200px]">
                        {row[cIndex] !== undefined ? String(row[cIndex]) : <span className="text-slate-300">-</span>}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Actions */}
      <div className="flex justify-end pt-3 border-t border-slate-100">
        <button
          onClick={onProceed}
          disabled={!isValid}
          className={`flex items-center gap-1.5 px-5 py-2.5 rounded-lg text-sm font-medium transition-colors cursor-pointer ${
            isValid
              ? 'bg-teal-600 hover:bg-teal-700 text-white shadow-sm shadow-teal-600/10'
              : 'bg-slate-100 text-slate-400 cursor-not-allowed'
          }`}
          id="proceed-to-check-btn"
        >
          Bắt đầu kiểm tra trùng lịch
          <ArrowRight className="w-4 h-4" />
        </button>
      </div>
    </div>
  );
}
