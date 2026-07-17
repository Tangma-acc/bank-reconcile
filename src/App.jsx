import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { Upload, CheckCircle2, RefreshCw, Plus, ArrowRightLeft, Trash2, Save, Download, Search, Calendar, X } from 'lucide-react';
import ExcelJS from 'exceljs';

const BankReconcileApp = () => {
  const [activeTab, setActiveTab] = useState('reconcile');
  const [internalRecords, setInternalRecords] = useState([]);
  const [bankStatement, setBankStatement] = useState([]);
  const [selectedInternal, setSelectedInternal] = useState([]);
  const [selectedBank, setSelectedBank] = useState([]);
  const [confirmedMatches, setConfirmedMatches] = useState([]);

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
    if (typeof val === 'number') date = new Date(Math.round((val - 25569) * 86400 * 1000));
    else {
      date = new Date(val);
      if (isNaN(date.getTime())) return val;
    }
    return `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
  };

  const parseDisplayDate = (dateStr) => {
    if (!dateStr || dateStr === '-') return null;
    const parts = dateStr.split('/');
    if (parts.length !== 3) return null;
    return new Date(parts[2], parts[1] - 1, parts[0]);
  };

  // --- Filtering Logic ---
  const filteredInternal = useMemo(() => {
    return internalRecords.filter(item => {
      const matchesSearch = searchInternal === '' || Math.abs(item.amount).toString().includes(searchInternal);
      const itemDate = parseDisplayDate(item.date);
      let matchesDate = true;
      if (itemDate) {
        if (internalStartDate && itemDate < new Date(internalStartDate)) matchesDate = false;
        if (internalEndDate && itemDate > new Date(internalEndDate)) matchesDate = false;
      }
      return matchesSearch && matchesDate;
    });
  }, [internalRecords, searchInternal, internalStartDate, internalEndDate]);

  const filteredBank = useMemo(() => {
    return bankStatement.filter(item => {
      const matchesSearch = searchBank === '' || Math.abs(item.amount).toString().includes(searchBank);
      const itemDate = parseDisplayDate(item.date);
      let matchesDate = true;
      if (itemDate) {
        if (bankStartDate && itemDate < new Date(bankStartDate)) matchesDate = false;
        if (bankEndDate && itemDate > new Date(bankEndDate)) matchesDate = false;
      }
      return matchesSearch && matchesDate;
    });
  }, [bankStatement, searchBank, bankStartDate, bankEndDate]);

  // --- File Upload Logic ---
  const handleFileUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;
    const isInternal = type === 'internal';
    const isExpense = file.name.toLowerCase().includes('expense');
    const reader = new FileReader();
    
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
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
          const docNo = item['เลขที่เอกสาร'];
          if (!docNo || docNo === 'รวม') return null;
          let amount = parseFloat(String(item['ต้องชำระ'] || 0).replace(/,/g, ''));
          if (isExpense && amount > 0) amount = -amount;
          return { id: `peak-${Date.now()}-${index}-${Math.random()}`, docNo, date: formatExcelDate(item['วันที่'] || item['วันที่ออก']), description: item['คำอธิบาย'] || '', status: item['สถานะ'] || '', amount };
        } else {
          const dateVal = item['วันที่'];
          const amountVal = item['ถอนเงิน/ฝากเงิน'];
          if (!dateVal || amountVal === undefined || amountVal === null || amountVal === "") return null;
          let amount = parseFloat(String(amountVal).replace(/[( )]/g, '').replace(/,/g, ''));
          if (String(amountVal).includes('(')) amount = -Math.abs(amount);
          return { 
            id: `bank-${Date.now()}-${index}-${Math.random()}`, 
            index: index, // เก็บตำแหน่งบรรทัดเดิม
            docNo: `${item['รายละเอียด'] || item['รายการ'] || 'STM'}${item['เวลา'] ? ` [${item['เวลา']}]` : ''}`, 
            date: formatExcelDate(dateVal), 
            amount 
          };
        }
      }).filter(i => i !== null && !isNaN(i.amount) && i.amount !== 0);

      if (isInternal) setInternalRecords(prev => [...prev, ...formattedData]);
      else setBankStatement(prev => [...prev, ...formattedData]);
    };
    reader.readAsArrayBuffer(file);
    e.target.value = null;
  };

  const internalSum = useMemo(() => selectedInternal.reduce((acc, curr) => acc + curr.amount, 0), [selectedInternal]);
  const bankSum = useMemo(() => selectedBank.reduce((acc, curr) => acc + curr.amount, 0), [selectedBank]);
  const diff = Math.abs(internalSum - bankSum);

  const toggleSelection = (item, type) => {
    if (type === 'internal') setSelectedInternal(prev => prev.some(i => i.id === item.id) ? prev.filter(i => i.id !== item.id) : [...prev, item]);
    else setSelectedBank(prev => prev.some(i => i.id === item.id) ? prev.filter(i => i.id !== item.id) : [...prev, item]);
  };

  const confirmMatch = () => {
    if (diff < 0.01 && internalSum !== 0) {
      const newMatch = { id: Date.now(), internals: [...selectedInternal], banks: [...selectedBank], totalAmount: internalSum };
      setConfirmedMatches(prev => [newMatch, ...prev]);
      setInternalRecords(prev => prev.filter(item => !selectedInternal.some(s => s.id === item.id)));
      setBankStatement(prev => prev.filter(item => !selectedBank.some(s => s.id === item.id)));
      setSelectedInternal([]); setSelectedBank([]);
    }
  };

  // --- Export Logic ---
  const exportToExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Report');

    // 1. หัวตาราง
    const headers = ["#", "วันที่ออก", "กระทบยอด", "หมายเหตุ", "เงินเข้า", "เงินออก", "สถานะ"];
    const headerRow = worksheet.addRow(headers);
    
    for (let i = 1; i <= 7; i++) {
      const cell = headerRow.getCell(i);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
      cell.font = { bold: true, name: 'Sarabun' };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
      };
    }

    // 2. รวบรวมข้อมูลตาม Statement
    const combinedData = [];
    confirmedMatches.forEach(match => {
      const peakDocs = match.internals.map(i => i.docNo).join(', ');
      match.banks.forEach(bankItem => {
        combinedData.push({ ...bankItem, matchedDocNo: peakDocs, status: "กระทบยอดแล้ว" });
      });
    });
    bankStatement.forEach(bankItem => {
      combinedData.push({ ...bankItem, matchedDocNo: "", status: "ยังไม่กระทบยอด" });
    });

    // เรียงตาม Index จริง
    combinedData.sort((a, b) => a.index - b.index);

    // 3. เขียนข้อมูล
    const numFormat = '#,##0.00;[Red](#,##0.00)';
    combinedData.forEach((entry, idx) => {
      const row = worksheet.addRow([
        idx + 1,
        entry.date,
        entry.matchedDocNo,
        entry.docNo,
        entry.amount > 0 ? entry.amount : null,
        entry.amount < 0 ? Math.abs(entry.amount) : null,
        entry.status
      ]);

      // ใส่เส้นขอบและรูปแบบให้ครบทุกช่อง (1-7)
      for (let i = 1; i <= 7; i++) {
        const cell = row.getCell(i);
        cell.border = {
          top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' }
        };
        cell.alignment = { vertical: 'top', horizontal: 'left', wrapText: false };

        if (i === 1) cell.alignment.horizontal = 'center';
        if (i === 5 || i === 6) {
          cell.numFmt = numFormat;
          cell.alignment.horizontal = 'right';
        }
        if (i === 7) {
          if (entry.status === "ยังไม่กระทบยอด") {
            cell.font = { color: { argb: 'FFFF0000' }, bold: true };
          } else {
            cell.font = { color: { argb: 'FF008000' }, bold: true };
          }
        }
      }
    });

    worksheet.columns = [
      { width: 6 }, { width: 14 }, { width: 35 }, { width: 45 }, { width: 15 }, { width: 15 }, { width: 18 }
    ];

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Bank_Reconcile_Report_${new Date().toISOString().split('T')[0]}.xlsx`; // = ชื่อไฟล์ Export
    a.click();
  };

return (
  <div className="min-h-screen bg-[#f1f5f9] p-4 md:p-6 font-sans text-slate-700">
    <div className="max-w-[1500px] mx-auto flex flex-col h-full">
      
      {/* Header */}
      <div className="flex justify-between items-center mb-6 bg-white p-5 rounded-3xl shadow-sm border border-slate-100">
        <h1 className="text-2xl font-black text-blue-900 italic">BANK RECONCILE</h1>
        
        <div className="flex items-center gap-3 flex-shrink-0">
            <button 
              onClick={downloadTemplate} 
              className="flex-shrink-0 flex items-center gap-2 bg-blue-50 text-blue-700 border border-blue-100 px-4 py-2 rounded-xl font-black text-xs hover:bg-blue-100 transition-all shadow-sm uppercase tracking-wider"
            >
              <Save size={16} /> Template
            </button>

            <button 
              onClick={exportToExcel} 
              className="flex-shrink-0 flex items-center gap-2 bg-emerald-50 text-emerald-700 border border-emerald-100 px-4 py-2 rounded-xl font-black text-xs hover:bg-emerald-100 transition-all shadow-sm uppercase tracking-wider"
            >
              <Download size={16} /> Export Excel
            </button>
            
            <div className="w-px h-6 bg-slate-200 mx-1 flex-shrink-0"></div>
            
            <button 
              onClick={() => window.location.reload()} 
              className="flex-shrink-0 bg-white text-slate-400 border border-slate-200 px-4 py-2 rounded-xl font-bold text-xs hover:text-red-500 hover:border-red-200 hover:bg-red-50 transition-all uppercase tracking-wider"
            >
              ล้างข้อมูล
            </button>
        </div>
      </div>

      {/* Tabs */}
      <div className="flex gap-4 mb-6 ml-2">
        <button onClick={() => setActiveTab('reconcile')} className={`px-8 py-2.5 rounded-full font-black text-xs transition-all ${activeTab === 'reconcile' ? 'bg-blue-600 text-white shadow-xl' : 'text-slate-400'}`}>รอกระทบยอด</button>
        <button onClick={() => setActiveTab('confirmed')} className={`px-8 py-2.5 rounded-full font-black text-xs transition-all flex items-center gap-2 ${activeTab === 'confirmed' ? 'bg-blue-600 text-white shadow-xl' : 'text-slate-400'}`}>รอยืนยัน {confirmedMatches.length > 0 && <span className="bg-orange-500 text-white px-1.5 py-0.5 rounded-full text-[8px]">{confirmedMatches.length}</span>}</button>
      </div>

      <div className="flex-1">
        {activeTab === 'reconcile' ? (
          <div className="space-y-6">
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 h-[580px]">
              
              {/* Left Side: PEAK/Internal */}
              <div className="bg-white rounded-[2.5rem] shadow-sm border flex flex-col overflow-hidden">
                <div className="p-5 bg-blue-600 text-white space-y-4">
                  <div className="flex justify-between items-center">
                      <span className="font-black text-[15px] uppercase tracking-widest">รายการบันทึกบัญชี ({internalRecords.length})</span>
                      <label className="bg-white/20 px-4 py-1.5 rounded-xl cursor-pointer text-[10px] font-black border border-white/30 hover:bg-white/40 transition-all uppercase">
                          <Plus size={12} className="inline mr-1"/>นำเข้า
                          <input type="file" onChange={(e) => handleFileUpload(e, 'internal')} className="hidden" accept=".xlsx, .xls" />
                      </label>
                  </div>
                  <div className="flex gap-2">
                    <div className="relative flex-1"><Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-white/40" /><input type="text" placeholder="ยอดเงิน..." value={searchInternal} onChange={e => setSearchInternal(e.target.value)} className="w-full bg-white/10 border border-white/20 rounded-xl pl-8 pr-8 py-2 text-[10px] outline-none" />
                    {searchInternal && (<button onClick={() => setSearchInternal('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-white/40 hover:text-white transition-all"><X size={14} /></button>)}</div>
                    <div className="flex bg-white/10 rounded-xl p-1 items-center border border-white/20"><Calendar size={12} className="ml-2 text-white/50" /><input type="date" value={internalStartDate} onChange={e => setInternalStartDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none" /><span className="text-white/50">-</span><input type="date" value={internalEndDate} onChange={e => setInternalEndDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none" />{(internalStartDate || internalEndDate) && <button onClick={()=>{setInternalStartDate('');setInternalEndDate('');}} className="p-1 text-white"><X size={12}/></button>}</div>
                  </div>
                </div>
                <div className="flex-1 overflow-y-auto p-4 space-y-2 bg-slate-50/30">
                  {filteredInternal.map((item) => (
                    <div key={item.id} onClick={() => toggleSelection(item, 'internal')} className={`p-4 rounded-2xl border-2 transition-all cursor-pointer min-h-[85px] flex items-center ${selectedInternal.some(i => i.id === item.id) ? 'border-blue-500 bg-blue-50' : 'border-white bg-white shadow-sm'}`}>
                      <div className="flex justify-between items-center w-full">
                        <div className="flex flex-col gap-1 flex-1 min-w-0 pr-4">
                            <div className="flex items-center gap-2">
                                <span className="font-bold text-slate-800 text-xs truncate">{item.docNo}</span>
                                {item.status && <span className="text-[8px] px-1.5 py-0.5 rounded-full bg-blue-100 text-blue-600 font-black flex-shrink-0">{item.status}</span>}
                            </div>
                            <span className="text-[10px] text-slate-400 italic truncate">{item.description || '-'}</span>
                            <span className="text-[9px] text-slate-400 font-bold uppercase tracking-tighter">{item.date}</span>
                        </div>
                        <span className={`text-xl font-black tabular-nums flex-shrink-0 ${item.amount < 0 ? 'text-red-500' : 'text-blue-600'}`}>
                            {item.amount.toLocaleString(undefined, {minimumFractionDigits: 2})}
                        </span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Right Side: Bank Statement */}
              <div className="bg-white rounded-[2.5rem] shadow-sm border flex flex-col overflow-hidden">
                <div className="p-5 bg-slate-800 text-white space-y-4">
                  <div className="flex justify-between items-center">
                      <span className="font-black text-[15px] uppercase tracking-widest text-slate-300">รายการธนาคาร ({bankStatement.length})</span>
                      <label className="bg-white/10 px-4 py-1.5 rounded-xl cursor-pointer text-[10px] font-black border border-white/10 hover:bg-white/20 transition-all uppercase">
                          <Plus size={12} className="inline mr-1"/>นำเข้า
                          <input type="file" onChange={(e) => handleFileUpload(e, 'bank')} className="hidden" accept=".xlsx, .xls" />
                      </label>
                  </div>
                  <div className="flex gap-2">
                   <div className="relative flex-1"><Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-white/40" /><input type="text" placeholder="ยอดเงิน..." value={searchBank} onChange={e => setSearchBank(e.target.value)} className="w-full bg-white/5 border border-white/10 rounded-xl pl-8 pr-8 py-2 text-[10px] outline-none" />
                  {searchBank && (<button onClick={() => setSearchBank('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-white/40 hover:text-white transition-all"><X size={14} /></button>)}</div>
                    <div className="flex bg-white/5 rounded-xl p-1 items-center border border-white/10"><Calendar size={12} className="text-white/30" /><input type="date" value={bankStartDate} onChange={e => setBankStartDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none opacity-60" /><span className="text-white/10">-</span><input type="date" value={bankEndDate} onChange={e => setBankEndDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none opacity-60" />{(bankStartDate || bankEndDate) && <button onClick={()=>{setBankStartDate('');setBankEndDate('');}} className="p-1 text-white"><X size={12}/></button>}</div>
                  </div>
                </div>
                <div className="flex-1 overflow-y-auto p-4 space-y-2 bg-slate-50/30">
                  {filteredBank.map((item) => (
                    <div key={item.id} onClick={() => toggleSelection(item, 'bank')} className={`p-4 rounded-2xl border-2 transition-all cursor-pointer min-h-[85px] flex items-center ${selectedBank.some(i => i.id === item.id) ? 'border-slate-800 bg-slate-100 shadow-md' : 'border-white bg-white shadow-sm'}`}>
                      <div className="flex justify-between items-center w-full">
                        <div className="flex flex-col gap-1 flex-1 min-w-0 pr-4">
                            <span className="font-bold text-slate-700 text-xs line-clamp-1 leading-snug">{item.docNo}</span>
                            {/* เพิ่มพื้นที่ว่าง (Placeholder) เพื่อให้ความสูงและการจัดวางตรงกับฝั่งบัญชี */}
                            <span className="text-[10px] text-slate-300 italic truncate opacity-0">-</span>
                            <span className="text-[9px] text-slate-400 font-bold uppercase tracking-tighter">{item.date}</span>
                        </div>
                        <span className={`text-xl font-black tabular-nums flex-shrink-0 ${item.amount < 0 ? 'text-red-500' : 'text-slate-900'}`}>
                            {item.amount.toLocaleString(undefined, {minimumFractionDigits: 2})}
                        </span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

            </div>

            {/* Summary Bottom */}
            <div className="bg-white p-8 rounded-[3.5rem] shadow-xl flex flex-col md:flex-row justify-around items-center border border-slate-100 gap-6">
              <div className="text-center"><div className="text-slate-400 text-[10px] font-black mb-1 uppercase tracking-widest">รวมบัญชี</div><div className="text-5xl font-black text-blue-600 tracking-tighter tabular-nums">{internalSum.toLocaleString(undefined, {minimumFractionDigits: 2})}</div></div>
              <div className="flex flex-col items-center bg-slate-50 px-16 py-6 rounded-[2.5rem] border shadow-inner min-w-[380px]">
                <div className="text-slate-400 text-[10px] font-black uppercase mb-1 tracking-widest">ผลต่างรวม</div>
                <div className={`text-6xl font-black tabular-nums tracking-tighter ${diff < 0.01 ? 'text-emerald-500' : 'text-red-500'}`}>{diff.toLocaleString(undefined, {minimumFractionDigits: 2})}</div>
                {diff < 0.01 && internalSum !== 0 && <button onClick={confirmMatch} className="mt-5 bg-blue-600 text-white px-12 py-3.5 rounded-full font-black text-xs hover:bg-blue-700 transition-all flex items-center gap-2 shadow-2xl animate-bounce tracking-widest uppercase">ยืนยันจับคู่ <ArrowRightLeft size={16}/></button>}
              </div>
              <div className="text-center"><div className="text-slate-400 text-[10px] font-black mb-1 uppercase tracking-widest">รวมธนาคาร</div><div className="text-5xl font-black text-slate-900 tracking-tighter tabular-nums">{bankSum.toLocaleString(undefined, {minimumFractionDigits: 2})}</div></div>
            </div>
          </div>
        ) : (
          /* ส่วนของ Confirmed Tab คงเดิมตาม Code ของคุณ */
          <div className="flex flex-col h-full gap-6">
            <div className="bg-white rounded-[3rem] shadow-sm border border-slate-200 overflow-hidden flex-1 flex flex-col min-h-[550px]">
              <div className="p-6 bg-slate-50 border-b grid grid-cols-4 font-black text-[11px] text-slate-400 uppercase tracking-widest"><span>รายการบัญชี</span><span className="text-center">ยอดเงิน</span><span className="pl-8">รายการธนาคาร</span><span className="text-right">ยอดเงิน</span></div>
              <div className="p-8 space-y-6 overflow-y-auto flex-1 bg-slate-50/20">
                {confirmedMatches.length === 0 ? (
                  <div className="flex flex-col items-center justify-center h-full text-slate-300 gap-4">
                    <div className="bg-slate-100 p-8 rounded-full"><Download size={48} /></div>
                    <p className="font-bold">ยังไม่มีข้อมูลที่จับคู่สำเร็จ</p>
                  </div>
                ) : (
                  confirmedMatches.map((m) => (
                    <div key={m.id} className="group bg-white border border-slate-100 rounded-[2rem] p-8 grid grid-cols-4 items-center shadow-sm relative hover:border-blue-300 transition-all">
                      <div className="space-y-3">{m.internals.map(i => <div key={i.id} className="flex flex-col"><span className="text-[10px] text-slate-400 font-black">{i.date}</span><span className="text-sm font-bold text-blue-700">{i.docNo}</span></div>)}</div>
                      <div className="text-center font-black text-slate-800 text-2xl border-r border-slate-50">{m.totalAmount.toLocaleString(undefined, {minimumFractionDigits: 2})}</div>
                      <div className="pl-8 space-y-3">{m.banks.map(b => <div key={b.id} className="flex flex-col"><span className="text-[10px] text-slate-400 font-black">{b.date}</span><span className="text-sm font-bold text-slate-800 line-clamp-1">{b.docNo}</span></div>)}</div>
                      <div className="text-right font-black text-slate-800 text-2xl">{m.totalAmount.toLocaleString(undefined, {minimumFractionDigits: 2})}</div>
                      <button onClick={() => { setConfirmedMatches(prev => prev.filter(x => x.id !== m.id)); setInternalRecords(prev => [...prev, ...m.internals]); setBankStatement(prev => [...prev, ...m.banks]); }} className="absolute -right-3 -top-3 bg-white text-rose-500 border-2 border-rose-50 rounded-full p-2.5 opacity-0 group-hover:opacity-100 transition-all shadow-xl hover:bg-rose-500 hover:text-white"><Trash2 size={18}/></button>
                    </div>
                  ))
                )}
              </div>
            </div>
            <div className="h-10"></div> 
          </div>
        )}
      </div>
    </div>
  </div>
);
};

export default BankReconcileApp;