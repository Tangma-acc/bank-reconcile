import React, { useState, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Plus, ArrowRightLeft, Trash2, Download, Search, Calendar, FileUp, Database, Landmark } from 'lucide-react';
import ExcelJS from 'exceljs';

const BankReconcileApp = () => {
  const [activeTab, setActiveTab] = useState('reconcile');
  const [internalRecords, setInternalRecords] = useState([]);
  const [bankStatement, setBankStatement] = useState([]);
  const [selectedInternal, setSelectedInternal] = useState([]);
  const [selectedBank, setSelectedBank] = useState([]);
  const [confirmedMatches, setConfirmedMatches] = useState([]);

  // --- Drag & Drop States ---
  const [isDraggingInternal, setIsDraggingInternal] = useState(false);
  const [isDraggingBank, setIsDraggingBank] = useState(false);

  // --- Search & Filter States ---
  const [searchInternal, setSearchInternal] = useState('');
  const [internalStartDate, setInternalStartDate] = useState('');
  const [internalEndDate, setInternalEndDate] = useState('');
  const [searchBank, setSearchBank] = useState('');
  const [bankStartDate, setBankStartDate] = useState('');
  const [bankEndDate, setBankEndDate] = useState('');

  // --- Helpers ---
  const formatExcelDate = (val) => {
    if (!val) return '-';
    let date;
    if (typeof val === 'number') {
      date = new Date(Math.round((val - 25569) * 86400 * 1000));
    } else if (typeof val === 'string') {
      const parts = val.split('/');
      if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1;
        const year = parseInt(parts[2], 10);
        date = new Date(year, month, day);
      } else {
        date = new Date(val);
      }
    } else {
      date = new Date(val);
    }
    if (isNaN(date.getTime())) return String(val);
    return `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
  };

  const parseDisplayDate = (dateStr) => {
    if (!dateStr || dateStr === '-') return null;
    const parts = dateStr.split('/');
    if (parts.length !== 3) return null;
    return new Date(parts[2], parts[1] - 1, parts[0], 0, 0, 0);
  };

  const sortByDate = (a, b) => {
    const dateA = parseDisplayDate(a.date);
    const dateB = parseDisplayDate(b.date);
    if (!dateA) return 1;
    if (!dateB) return -1;
    return dateA.getTime() - dateB.getTime();
  };

  const formatAccounting = (num) => {
    if (num === 0 || num === null || num === undefined) return "0.00";
    const formatted = Math.abs(num).toLocaleString(undefined, {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
    return num < 0 ? `(${formatted})` : formatted;
  };

  // --- ฟังก์ชันหลักในการประมวลผลไฟล์ (แชร์ใช้ทั้งปุ่มกดและลากวาง) ---
  const processFile = useCallback((file, type) => {
    if (!file) return;
    const isInternal = type === 'internal';
    const reader = new FileReader();
    
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: false });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      const searchKeys = isInternal ? ['เลขที่เอกสาร', 'ต้องชำระ'] : ['วันที่', 'ถอนเงิน/ฝากเงิน'];
      let headerIdx = rows.findIndex(row => 
        Array.isArray(row) && searchKeys.every(k => row.some(c => String(c).includes(k)))
      );
      if (headerIdx === -1) headerIdx = 0;

      const headers = rows[headerIdx];
      const dataRows = rows.slice(headerIdx + 1);

      const formattedData = dataRows.map((row, index) => {
        const item = {};
        headers.forEach((h, i) => { if (h) item[String(h).trim()] = row[i]; });

        if (isInternal) {
          const docNo = String(item['เลขที่เอกสาร'] || '');
          if (!docNo || docNo === 'รวม' || docNo.trim() === '') return null;
          let amount = parseFloat(String(item['ต้องชำระ'] || 0).replace(/,/g, ''));
          
          // ตรวจสอบ EXP ในเลขที่เอกสาร
          if (docNo.toLowerCase().includes('exp') && amount > 0) amount = -amount;
          
          return { id: `peak-${Date.now()}-${index}`, docNo, date: formatExcelDate(item['วันที่'] || item['วันที่ออก']), description: item['คำอธิบาย'] || '', status: item['สถานะ'] || '', amount };
        } else {
          const dateVal = item['วันที่'];
          const amountVal = item['ถอนเงิน/ฝากเงิน'];
          if (!dateVal || amountVal === undefined || amountVal === null || amountVal === "") return null;
          let amount = parseFloat(String(amountVal).replace(/[( )]/g, '').replace(/,/g, ''));
          if (String(amountVal).includes('(')) amount = -Math.abs(amount);
          return { id: `bank-${Date.now()}-${index}`, docNo: `${item['รายละเอียด'] || item['รายการ'] || 'STM'}${item['เวลา'] ? ` [${item['เวลา']}]` : ''}`, date: formatExcelDate(dateVal), amount };
        }
      }).filter(i => i !== null && !isNaN(i.amount) && i.amount !== 0);

      if (isInternal) setInternalRecords(prev => [...prev, ...formattedData].sort(sortByDate));
      else setBankStatement(prev => [...prev, ...formattedData].sort(sortByDate));
    };
    reader.readAsArrayBuffer(file);
  }, []);

  // --- Handlers สำหรับ Drag & Drop ---
  const handleDragOver = (e, type) => {
    e.preventDefault(); e.stopPropagation();
    if (type === 'internal') setIsDraggingInternal(true);
    else setIsDraggingBank(true);
  };

  const handleDragLeave = (e, type) => {
    e.preventDefault(); e.stopPropagation();
    if (type === 'internal') setIsDraggingInternal(false);
    else setIsDraggingBank(false);
  };

  const handleDrop = (e, type) => {
    e.preventDefault(); e.stopPropagation();
    setIsDraggingInternal(false); setIsDraggingBank(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processFile(e.dataTransfer.files[0], type);
    }
  };

  // --- Filtering & Totals ---
  const filteredInternal = useMemo(() => {
    return internalRecords.filter(item => {
      const matchesSearch = searchInternal === '' || Math.abs(item.amount).toString().includes(searchInternal) || item.docNo.toLowerCase().includes(searchInternal.toLowerCase());
      const itemDate = parseDisplayDate(item.date);
      let matchesDate = true;
      if (itemDate) {
        if (internalStartDate) {
          const start = new Date(internalStartDate); start.setHours(0,0,0,0);
          if (itemDate < start) matchesDate = false;
        }
        if (internalEndDate) {
          const end = new Date(internalEndDate); end.setHours(0,0,0,0);
          if (itemDate > end) matchesDate = false;
        }
      }
      return matchesSearch && matchesDate;
    });
  }, [internalRecords, searchInternal, internalStartDate, internalEndDate]);

  const filteredBank = useMemo(() => {
    return bankStatement.filter(item => {
      const matchesSearch = searchBank === '' || Math.abs(item.amount).toString().includes(searchBank) || item.docNo.toLowerCase().includes(searchBank.toLowerCase());
      const itemDate = parseDisplayDate(item.date);
      let matchesDate = true;
      if (itemDate) {
        if (bankStartDate) {
          const start = new Date(bankStartDate); start.setHours(0,0,0,0);
          if (itemDate < start) matchesDate = false;
        }
        if (bankEndDate) {
          const end = new Date(bankEndDate); end.setHours(0,0,0,0);
          if (itemDate > end) matchesDate = false;
        }
      }
      return matchesSearch && matchesDate;
    });
  }, [bankStatement, searchBank, bankStartDate, bankEndDate]);

  const internalSum = useMemo(() => selectedInternal.reduce((acc, curr) => acc + curr.amount, 0), [selectedInternal]);
  const bankSum = useMemo(() => selectedBank.reduce((acc, curr) => acc + curr.amount, 0), [selectedBank]);
  const diff = Math.abs(internalSum - bankSum);

  const toggleSelection = (item, type) => {
    if (type === 'internal') setSelectedInternal(prev => prev.some(i => i.id === item.id) ? prev.filter(i => i.id !== item.id) : [...prev, item]);
    else setSelectedBank(prev => prev.some(i => i.id === item.id) ? prev.filter(i => i.id !== item.id) : [...prev, item]);
  };

  const confirmMatch = () => {
    if (diff < 0.01 && (selectedInternal.length > 0 || selectedBank.length > 0)) {
      const newMatch = { id: Date.now(), internals: [...selectedInternal], banks: [...selectedBank], totalAmount: internalSum };
      setConfirmedMatches(prev => [newMatch, ...prev]);
      setInternalRecords(prev => prev.filter(item => !selectedInternal.some(s => s.id === item.id)));
      setBankStatement(prev => prev.filter(item => !selectedBank.some(s => s.id === item.id)));
      setSelectedInternal([]); setSelectedBank([]);
    }
  };

  const exportToExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Report');
    const accountingFormat = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
    const headers = ["#", "วันที่", "เลขที่เอกสาร/การจับคู่", "รายละเอียด", "ยอดเงิน", "สถานะ"];
    const headerRow = worksheet.addRow(headers);
    headerRow.eachCell(cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
      cell.font = { bold: true };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    });
    const combinedData = [];
    confirmedMatches.forEach(match => {
      match.banks.forEach(bankItem => combinedData.push({ ...bankItem, matchedDocNo: match.internals.map(i => i.docNo).join(', '), status: "กระทบยอดแล้ว" }));
    });
    bankStatement.forEach(bankItem => combinedData.push({ ...bankItem, matchedDocNo: "", status: "ยังไม่กระทบยอด" }));
    combinedData.sort(sortByDate);
    combinedData.forEach((entry, idx) => {
      const row = worksheet.addRow([idx + 1, entry.date, entry.matchedDocNo, entry.docNo, entry.amount, entry.status]);
      row.eachCell((cell, col) => {
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        if (col === 5) { cell.numFmt = accountingFormat; cell.alignment = { horizontal: 'right' }; }
        if (col === 6) { cell.font = { color: { argb: entry.status === "ยังไม่กระทบยอด" ? 'FFFF0000' : 'FF008000' }, bold: true }; }
      });
    });
    worksheet.columns = [{ width: 6 }, { width: 14 }, { width: 35 }, { width: 45 }, { width: 20 }, { width: 18 }];
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = `Bank_Reconcile_${new Date().toISOString().split('T')[0]}.xlsx`; a.click();
  };

  return (
    <div className="min-h-screen bg-[#f1f5f9] p-4 md:p-6 font-sans text-slate-700">
      <div className="max-w-[1500px] mx-auto flex flex-col h-full">
        {/* Header */}
        <div className="flex justify-between items-center mb-6 bg-white p-5 rounded-3xl shadow-sm border border-slate-100">
          <h1 className="text-2xl font-black text-blue-900 italic uppercase">Bank Reconciliation</h1>
          <div className="flex items-center gap-3">
            <button onClick={exportToExcel} className="flex items-center gap-2 bg-emerald-50 text-emerald-700 border border-emerald-100 px-4 py-2 rounded-xl font-black text-xs hover:bg-emerald-100 transition-all uppercase"><Download size={16} /> Export Excel</button>
            <div className="w-px h-6 bg-slate-200 mx-1"></div>
            <button onClick={() => window.location.reload()} className="bg-white text-slate-400 border border-slate-200 px-4 py-2 rounded-xl font-bold text-xs hover:text-red-500 hover:bg-red-50 transition-all uppercase">ล้างข้อมูล</button>
          </div>
        </div>

        {/* Tab Switcher */}
        <div className="flex gap-4 mb-6 ml-2">
          <button onClick={() => setActiveTab('reconcile')} className={`px-8 py-2.5 rounded-full font-black text-xs transition-all ${activeTab === 'reconcile' ? 'bg-blue-600 text-white shadow-xl' : 'text-slate-400'}`}>รอกระทบยอด</button>
          <button onClick={() => setActiveTab('confirmed')} className={`px-8 py-2.5 rounded-full font-black text-xs transition-all flex items-center gap-2 ${activeTab === 'confirmed' ? 'bg-blue-600 text-white shadow-xl' : 'text-slate-400'}`}>กระทบยอดแล้ว {confirmedMatches.length > 0 && <span className="bg-orange-500 text-white px-1.5 py-0.5 rounded-full text-[8px]">{confirmedMatches.length}</span>}</button>
        </div>

        {/* Main Section */}
        <div className="flex-1">
          {activeTab === 'reconcile' ? (
            <div className="space-y-6">
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 h-[580px]">
                
                {/* 1. ฝั่งบันทึกบัญชี (Drop Zone) */}
                <div 
                  onDragOver={(e) => handleDragOver(e, 'internal')}
                  onDragLeave={(e) => handleDragLeave(e, 'internal')}
                  onDrop={(e) => handleDrop(e, 'internal')}
                  className={`bg-white rounded-[2.5rem] shadow-sm border-2 transition-all flex flex-col overflow-hidden relative ${isDraggingInternal ? 'border-blue-500 border-dashed bg-blue-50' : 'border-transparent'}`}
                >
                  <div className="p-5 bg-blue-600 text-white space-y-4">
                    <div className="flex justify-between items-center">
                      <span className="font-black text-[15px] uppercase tracking-widest flex items-center gap-2"><Database size={18}/> รายการบันทึกบัญชี ({internalRecords.length})</span>
                      <label className="bg-white/20 px-4 py-1.5 rounded-xl cursor-pointer text-[10px] font-black border border-white/30 hover:bg-white/40 transition-all uppercase">
                        <Plus size={12} className="inline mr-1"/> นำเข้า
                        <input type="file" onChange={(e) => processFile(e.target.files[0], 'internal')} className="hidden" accept=".xlsx, .xls" />
                      </label>
                    </div>
                    <div className="flex gap-2">
                      <div className="relative flex-1"><Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-white/40" /><input type="text" placeholder="ค้นหา..." value={searchInternal} onChange={e => setSearchInternal(e.target.value)} className="w-full bg-white/10 border border-white/20 rounded-xl pl-8 pr-8 py-2 text-[10px] outline-none placeholder:text-white/30" /></div>
                      <div className="flex bg-white/10 rounded-xl p-1 items-center border border-white/20"><Calendar size={12} className="ml-2 text-white/50" /><input type="date" value={internalStartDate} onChange={e => setInternalStartDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none" /><span className="text-white/50">-</span><input type="date" value={internalEndDate} onChange={e => setInternalEndDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none" /></div>
                    </div>
                  </div>
                  
                  <div className="flex-1 overflow-y-auto p-4 space-y-2 bg-slate-50/30">
                    {isDraggingInternal && (
                      <div className="absolute inset-0 z-50 bg-blue-600/20 flex flex-col items-center justify-center pointer-events-none animate-in fade-in duration-200">
                        <div className="bg-blue-600 text-white p-6 rounded-3xl shadow-2xl flex flex-col items-center gap-2 animate-bounce">
                           <FileUp size={40} />
                           <span className="font-black text-sm uppercase">วางไฟล์เพื่อนำเข้าบัญชี</span>
                        </div>
                      </div>
                    )}
                    {filteredInternal.map(item => (
                      <div key={item.id} onClick={() => toggleSelection(item, 'internal')} className={`p-4 rounded-2xl border-2 transition-all cursor-pointer min-h-[85px] flex items-center ${selectedInternal.some(i => i.id === item.id) ? 'border-blue-500 bg-blue-50' : 'border-white bg-white shadow-sm hover:border-blue-200'}`}>
                        <div className="flex justify-between items-center w-full">
                          <div className="flex flex-col gap-1 flex-1 min-w-0 pr-4">
                            <div className="flex items-center gap-2"><span className="font-bold text-slate-800 text-xs truncate">{item.docNo}</span>{item.status && <span className="text-[8px] px-1.5 py-0.5 rounded-full bg-blue-100 text-blue-600 font-black">{item.status}</span>}</div>
                            <span className="text-[10px] text-slate-400 italic truncate">{item.description || '-'}</span>
                            <span className="text-[9px] text-slate-400 font-bold uppercase">{item.date}</span>
                          </div>
                          <span className={`text-xl font-black tabular-nums ${item.amount < 0 ? 'text-red-500' : 'text-blue-600'}`}>{formatAccounting(item.amount)}</span>
                        </div>
                      </div>
                    ))}
                    {filteredInternal.length === 0 && !isDraggingInternal && (
                      <div className="h-full flex flex-col items-center justify-center text-slate-300 py-10 opacity-50">
                        <FileUp size={48} strokeWidth={1} />
                        <p className="text-[10px] font-bold mt-2 uppercase">ลากไฟล์บัญชีมาวางที่นี่</p>
                      </div>
                    )}
                  </div>
                </div>

                {/* 2. ฝั่งธนาคาร (Drop Zone) */}
                <div 
                  onDragOver={(e) => handleDragOver(e, 'bank')}
                  onDragLeave={(e) => handleDragLeave(e, 'bank')}
                  onDrop={(e) => handleDrop(e, 'bank')}
                  className={`bg-white rounded-[2.5rem] shadow-sm border-2 transition-all flex flex-col overflow-hidden relative ${isDraggingBank ? 'border-slate-800 border-dashed bg-slate-100' : 'border-transparent'}`}
                >
                  <div className="p-5 bg-slate-800 text-white space-y-4">
                    <div className="flex justify-between items-center">
                      <span className="font-black text-[15px] uppercase tracking-widest text-slate-300 flex items-center gap-2"><Landmark size={18}/> รายการธนาคาร ({bankStatement.length})</span>
                      <label className="bg-white/10 px-4 py-1.5 rounded-xl cursor-pointer text-[10px] font-black border border-white/10 hover:bg-white/20 transition-all uppercase">
                        <Plus size={12} className="inline mr-1"/> นำเข้า
                        <input type="file" onChange={(e) => processFile(e.target.files[0], 'bank')} className="hidden" accept=".xlsx, .xls" />
                      </label>
                    </div>
                    <div className="flex gap-2">
                      <div className="relative flex-1"><Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-white/40" /><input type="text" placeholder="ค้นหา..." value={searchBank} onChange={e => setSearchBank(e.target.value)} className="w-full bg-white/5 border border-white/10 rounded-xl pl-8 pr-8 py-2 text-[10px] outline-none placeholder:text-white/20" /></div>
                      <div className="flex bg-white/5 rounded-xl p-1 items-center border border-white/10"><Calendar size={12} className="text-white/30" /><input type="date" value={bankStartDate} onChange={e => setBankStartDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none opacity-60" /><span className="text-white/10">-</span><input type="date" value={bankEndDate} onChange={e => setBankEndDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none opacity-60" /></div>
                    </div>
                  </div>

                  <div className="flex-1 overflow-y-auto p-4 space-y-2 bg-slate-50/30">
                    {isDraggingBank && (
                      <div className="absolute inset-0 z-50 bg-slate-900/20 flex flex-col items-center justify-center pointer-events-none animate-in fade-in duration-200">
                        <div className="bg-slate-800 text-white p-6 rounded-3xl shadow-2xl flex flex-col items-center gap-2 animate-bounce">
                           <FileUp size={40} />
                           <span className="font-black text-sm uppercase">วางไฟล์เพื่อนำเข้าธนาคาร</span>
                        </div>
                      </div>
                    )}
                    {filteredBank.map(item => (
                      <div key={item.id} onClick={() => toggleSelection(item, 'bank')} className={`p-4 rounded-2xl border-2 transition-all cursor-pointer min-h-[85px] flex items-center ${selectedBank.some(i => i.id === item.id) ? 'border-slate-800 bg-slate-100 shadow-md' : 'border-white bg-white shadow-sm hover:border-slate-300'}`}>
                        <div className="flex justify-between items-center w-full">
                          <div className="flex flex-col gap-1 flex-1 min-w-0 pr-4">
                            <span className="font-bold text-slate-700 text-xs line-clamp-1">{item.docNo}</span>
                            <span className="text-[10px] opacity-0">-</span>
                            <span className="text-[9px] text-slate-400 font-bold uppercase">{item.date}</span>
                          </div>
                          <span className={`text-xl font-black tabular-nums ${item.amount < 0 ? 'text-red-500' : 'text-slate-900'}`}>{formatAccounting(item.amount)}</span>
                        </div>
                      </div>
                    ))}
                    {filteredBank.length === 0 && !isDraggingBank && (
                      <div className="h-full flex flex-col items-center justify-center text-slate-300 py-10 opacity-50">
                        <FileUp size={48} strokeWidth={1} />
                        <p className="text-[10px] font-bold mt-2 uppercase">ลากไฟล์สเตทเม้นท์มาวางที่นี่</p>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              {/* Summary Bottom Bar */}
              <div className="bg-white p-8 rounded-[3.5rem] shadow-xl flex flex-col md:flex-row justify-around items-center border border-slate-100 gap-6">
                <div className="text-center"><div className="text-slate-400 text-[10px] font-black uppercase tracking-widest">บัญชีที่เลือก</div><div className="text-5xl font-black text-blue-600 tracking-tighter tabular-nums">{formatAccounting(internalSum)}</div></div>
                <div className="flex flex-col items-center bg-slate-50 px-16 py-6 rounded-[2.5rem] border shadow-inner min-w-[380px]">
                  <div className="text-slate-400 text-[10px] font-black uppercase mb-1 tracking-widest">ผลต่างรวม</div>
                  <div className={`text-6xl font-black tabular-nums tracking-tighter ${diff < 0.01 ? 'text-emerald-500' : 'text-red-500'}`}>{formatAccounting(diff)}</div>
                  {diff < 0.01 && (selectedInternal.length > 0 || selectedBank.length > 0) && (
                    <button onClick={confirmMatch} className="mt-5 bg-blue-600 text-white px-12 py-3.5 rounded-full font-black text-xs hover:bg-blue-700 transition-all flex items-center gap-2 shadow-2xl animate-bounce tracking-widest uppercase">ยืนยันจับคู่ <ArrowRightLeft size={16}/></button>
                  )}
                </div>
                <div className="text-center"><div className="text-slate-400 text-[10px] font-black uppercase tracking-widest">ธนาคารที่เลือก</div><div className="text-5xl font-black text-slate-900 tracking-tighter tabular-nums">{formatAccounting(bankSum)}</div></div>
              </div>
            </div>
          ) : (
            /* Confirmed Matches List */
            <div className="flex flex-col h-full gap-6">
              <div className="bg-white rounded-[3rem] shadow-sm border border-slate-200 overflow-hidden flex-1 flex flex-col min-h-[550px]">
                <div className="p-6 bg-slate-50 border-b grid grid-cols-4 font-black text-[11px] text-slate-400 uppercase tracking-widest"><span>รายการบัญชี</span><span className="text-center">ยอดเงิน</span><span className="pl-8">รายการธนาคาร</span><span className="text-right">ยอดเงิน</span></div>
                <div className="p-8 space-y-6 overflow-y-auto flex-1 bg-slate-50/20">
                  {confirmedMatches.map(m => (
                    <div key={m.id} className="group bg-white border border-slate-100 rounded-[2rem] p-8 grid grid-cols-4 items-center shadow-sm relative hover:border-blue-300 transition-all">
                      <div className="space-y-3">{m.internals.map(i => <div key={i.id} className="flex flex-col"><span className="text-[10px] text-slate-400 font-black">{i.date}</span><span className="text-sm font-bold text-blue-700">{i.docNo}</span></div>)}</div>
                      <div className={`text-center font-black text-2xl border-r border-slate-50 ${m.totalAmount < 0 ? 'text-red-500' : 'text-slate-800'}`}>{formatAccounting(m.totalAmount)}</div>
                      <div className="pl-8 space-y-3">{m.banks.map(b => <div key={b.id} className="flex flex-col"><span className="text-[10px] text-slate-400 font-black">{b.date}</span><span className="text-sm font-bold text-slate-800 line-clamp-1">{b.docNo}</span></div>)}</div>
                      <div className={`text-right font-black text-2xl ${m.totalAmount < 0 ? 'text-red-500' : 'text-slate-800'}`}>{formatAccounting(m.totalAmount)}</div>
                      <button onClick={() => { setConfirmedMatches(prev => prev.filter(x => x.id !== m.id)); setInternalRecords(p => [...p, ...m.internals].sort(sortByDate)); setBankStatement(p => [...p, ...m.banks].sort(sortByDate)); }} className="absolute -right-3 -top-3 bg-white text-rose-500 border-2 border-rose-50 rounded-full p-2.5 opacity-0 group-hover:opacity-100 transition-all shadow-xl hover:bg-rose-500 hover:text-white"><Trash2 size={18}/></button>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default BankReconcileApp;