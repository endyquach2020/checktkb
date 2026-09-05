import React, { useState } from 'react';
import { ExemptionPair } from '../types';
import { Plus, Trash2, ShieldCheck, HelpCircle } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

interface ExemptionSettingsProps {
  exemptions: ExemptionPair[];
  onChange: (updatedExemptions: ExemptionPair[]) => void;
}

export default function ExemptionSettings({ exemptions, onChange }: ExemptionSettingsProps) {
  const [classA, setClassA] = useState('');
  const [classB, setClassB] = useState('');

  const handleToggle = (id: string) => {
    const updated = exemptions.map(ex => 
      ex.id === id ? { ...ex, isActive: !ex.isActive } : ex
    );
    onChange(updated);
  };

  const handleDelete = (id: string) => {
    const updated = exemptions.filter(ex => ex.id !== id);
    onChange(updated);
  };

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    if (!classA.trim() || !classB.trim()) return;

    // Avoid duplicate pairs
    const exists = exemptions.some(ex => {
      const aNorm = ex.classA.trim().toLowerCase();
      const bNorm = ex.classB.trim().toLowerCase();
      const newANorm = classA.trim().toLowerCase();
      const newBNorm = classB.trim().toLowerCase();
      return (
        (aNorm === newANorm && bNorm === newBNorm) ||
        (aNorm === newBNorm && bNorm === newANorm)
      );
    });

    if (exists) {
      alert('Cặp lớp này đã tồn tại trong danh sách miễn trùng!');
      return;
    }

    const newPair: ExemptionPair = {
      id: crypto.randomUUID(),
      classA: classA.trim(),
      classB: classB.trim(),
      isActive: true
    };

    onChange([...exemptions, newPair]);
    setClassA('');
    setClassB('');
  };

  const restoreDefaults = () => {
    const defaults: ExemptionPair[] = [
      { id: 'def-10', classA: '10.1 TN', classB: '10.1 XH', isActive: true },
      { id: 'def-11', classA: '11.1 TN', classB: '11.2 XH', isActive: true },
      { id: 'def-12', classA: '12.1 TN', classB: '12.2 XH', isActive: true },
    ];
    onChange(defaults);
  };

  return (
    <div id="exemption-settings-container" className="bg-white rounded-xl border border-slate-100 p-6 shadow-sm">
      <div className="flex items-start justify-between mb-4">
        <div>
          <h3 className="text-lg font-semibold text-slate-800 flex items-center gap-2">
            <ShieldCheck className="w-5 h-5 text-emerald-500" />
            Cặp lớp ghép song song (Miễn trùng lịch)
          </h3>
          <p className="text-sm text-slate-500 mt-1">
            Các cặp lớp học dưới đây được ghép song song, cho phép cùng 1 giáo viên phụ trách tại cùng 1 tiết mà không tính là trùng lịch.
          </p>
        </div>
        <button
          type="button"
          onClick={restoreDefaults}
          className="text-xs font-medium text-emerald-600 hover:text-emerald-700 bg-emerald-50 hover:bg-emerald-100 px-3 py-1.5 rounded-lg transition-colors"
          id="restore-defaults-btn"
        >
          Khôi phục mặc định
        </button>
      </div>

      {/* Form thêm mới */}
      <form onSubmit={handleAdd} className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-6 p-4 bg-slate-50 rounded-lg border border-slate-100">
        <div className="flex flex-col gap-1">
          <label className="text-xs font-medium text-slate-500 uppercase tracking-wider">Lớp thứ nhất</label>
          <input
            type="text"
            required
            value={classA}
            onChange={(e) => setClassA(e.target.value)}
            placeholder="Ví dụ: 10.1 TN"
            className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all placeholder:text-slate-400"
            id="class-a-input"
          />
        </div>
        <div className="flex flex-col gap-1">
          <label className="text-xs font-medium text-slate-500 uppercase tracking-wider">Lớp thứ hai</label>
          <input
            type="text"
            required
            value={classB}
            onChange={(e) => setClassB(e.target.value)}
            placeholder="Ví dụ: 10.1 XH"
            className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all placeholder:text-slate-400"
            id="class-b-input"
          />
        </div>
        <div className="flex items-end">
          <button
            type="submit"
            className="w-full h-10 flex items-center justify-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white font-medium text-sm rounded-lg transition-colors cursor-pointer"
            id="add-pair-btn"
          >
            <Plus className="w-4 h-4" />
            Thêm cặp miễn trùng
          </button>
        </div>
      </form>

      {/* Danh sách hiện tại */}
      <div className="overflow-hidden border border-slate-100 rounded-lg">
        <table className="w-full text-left text-sm" id="exemptions-table">
          <thead className="bg-slate-50 border-b border-slate-100 text-slate-500 font-medium">
            <tr>
              <th className="px-4 py-2.5">Trạng thái</th>
              <th className="px-4 py-2.5">Lớp ghép A</th>
              <th className="px-4 py-2.5">Lớp ghép B</th>
              <th className="px-4 py-2.5 text-right">Thao tác</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            <AnimatePresence initial={false}>
              {exemptions.length === 0 ? (
                <tr id="no-exemptions-row">
                  <td colSpan={4} className="px-4 py-6 text-center text-slate-400 italic">
                    Chưa có cặp lớp ghép nào được miễn trùng. Tất cả các tiết có cùng giáo viên dạy nhiều lớp sẽ được báo trùng lịch.
                  </td>
                </tr>
              ) : (
                exemptions.map((ex) => (
                  <motion.tr
                    key={ex.id}
                    initial={{ opacity: 0, height: 0 }}
                    animate={{ opacity: 1, height: 'auto' }}
                    exit={{ opacity: 0, height: 0 }}
                    transition={{ duration: 0.2 }}
                    className={`hover:bg-slate-50/50 transition-colors ${!ex.isActive ? 'opacity-60 bg-slate-50/30' : ''}`}
                    id={`ex-row-${ex.id}`}
                  >
                    <td className="px-4 py-3 align-middle">
                      <label className="relative inline-flex items-center cursor-pointer">
                        <input
                          type="checkbox"
                          checked={ex.isActive}
                          onChange={() => handleToggle(ex.id)}
                          className="sr-only peer"
                        />
                        <div className="w-9 h-5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-emerald-500"></div>
                        <span className="ml-2 text-xs font-medium text-slate-500">
                          {ex.isActive ? 'Đang kích hoạt' : 'Đang tắt'}
                        </span>
                      </label>
                    </td>
                    <td className="px-4 py-3 font-semibold text-slate-700 align-middle">
                      <span className="px-2 py-1 bg-emerald-50 text-emerald-700 rounded text-xs border border-emerald-100">
                        {ex.classA}
                      </span>
                    </td>
                    <td className="px-4 py-3 font-semibold text-slate-700 align-middle">
                      <span className="px-2 py-1 bg-teal-50 text-teal-700 rounded text-xs border border-teal-100">
                        {ex.classB}
                      </span>
                    </td>
                    <td className="px-4 py-3 text-right align-middle">
                      <button
                        type="button"
                        onClick={() => handleDelete(ex.id)}
                        className="text-slate-400 hover:text-rose-500 p-1.5 hover:bg-rose-50 rounded-lg transition-colors cursor-pointer"
                        title="Xóa cặp lớp này"
                        id={`delete-ex-${ex.id}`}
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </td>
                  </motion.tr>
                ))
              )}
            </AnimatePresence>
          </tbody>
        </table>
      </div>
      <div className="flex items-center gap-1.5 mt-3 text-xs text-slate-400">
        <HelpCircle className="w-3.5 h-3.5 text-slate-400 flex-shrink-0" />
        <span>Tự động chuẩn hóa so sánh (ví dụ: "10.1 TN" và "10.1TN" đều được nhận diện khớp nhau).</span>
      </div>
    </div>
  );
}
