import React, { useState, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { 
  Plus, 
  ArrowRightLeft, 
  Trash2, 
  Download, 
  Search, 
  Calendar, 
  FileUp, 
  Database, 
  Landmark, 
  FilterX, 
  CheckCircle2,
  History 
} from 'lucide-react';
import ExcelJS from 'exceljs';

const BankReconciliationApp = () => {
  const [activeTab, setActiveTab] = useState('reconcile');
  const [internalRecords, setInternalRecords] = useState([]);
  const [bankStatement, setBankStatement] = useState([]);
  const [selectedInternal, setSelectedInternal] = useState([]);
  const [selectedBank, setSelectedBank] = useState([]);
  const [confirmedMatches, setConfirmedMatches] = useState([]);

  // Drag & Drop States
  const [isDraggingInternal, setIsDraggingInternal] = useState(false);
  const [isDraggingBank, setIsDraggingBank] = useState(false);

  // Search & Filter States
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
        date = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
      } else { date = new Date(val); }
    } else { date = new Date(val); }
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
    return (dateA?.getTime() || 0) - (dateB?.getTime() || 0);
  };

  const formatAccounting = (num) => {
    if (num === 0 || num === null || num === undefined) return "0.00";
    const formatted = Math.abs(num).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    return num < 0 ? `(${formatted})` : formatted;
  };

  // --- Core Processing (Account & Bank Files) ---
  const processFile = useCallback((file, type) => {
    if (!file) return;
    const isInternal = type === 'internal';
    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      const searchKeys = isInternal ? ['เลขที่เอกสาร', 'จำนวนเงิน'] : ['วันที่', 'ถอนเงิน/ฝากเงิน'];
      let headerIdx = rows.findIndex(row => Array.isArray(row) && searchKeys.every(k => row.some(c => String(c).includes(k))));
      if (headerIdx === -1) headerIdx = 0;

      const headers = rows[headerIdx];
      const dataRows = rows.slice(headerIdx + 1);

      const formattedData = dataRows.map((row, index) => {
        const item = {};
        headers.forEach((h, i) => { if (h) item[String(h).trim()] = row[i]; });

        if (isInternal) {
          const docNo = String(item['เลขที่เอกสาร'] || '');
          if (!docNo || docNo === 'รวม' || docNo.trim() === '') return null;
          let amount = parseFloat(String(item['จำนวนเงิน'] || 0).replace(/,/g, ''));
          const expenseKeywords = ['exp', 'dp', 'pa', 'pv']; 
          if (expenseKeywords.some(k => docNo.toLowerCase().includes(k)) && amount > 0) amount = -amount;
          return { id: `peak-${Date.now()}-${index}`, docNo, date: formatExcelDate(item['วันที่'] || item['วันที่ออก']), description: item['คำอธิบาย'] || item['รายละเอียด'] || '', amount };
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

  // --- Import Previous Report Logic ---
  const processReportFile = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet);

      const newBankRecords = [];
      const newConfirmedMatches = [];

      json.forEach((row, index) => {
        const status = row['สถานะ'] || '';
        const date = row['วันที่'] || '';
        let amount = parseFloat(String(row['ยอดเงิน'] || 0).replace(/,/g, '').replace(/[()]/g, (m) => m === '(' ? '-' : ''));
        const description = row['รายละเอียด'] || '';
        const internalDocsStr = row['เลขที่เอกสาร/การจับคู่'] || row['เลขที่เอกสาร'] || '';

        if (status === 'ยังไม่กระทบยอด') {
          newBankRecords.push({ id: `bank-rep-${Date.now()}-${index}`, date, docNo: description, amount });
        } else if (status === 'กระทบยอดแล้ว') {
          const docNos = internalDocsStr.split(',').map(d => d.trim()).filter(d => d !== '');
          newConfirmedMatches.push({
            id: `match-rep-${Date.now()}-${index}`,
            totalAmount: amount,
            internals: docNos.map((doc, i) => ({ id: `int-rep-${index}-${i}`, docNo: doc, date, amount })),
            banks: [{ id: `bank-rep-${index}`, docNo: description, date, amount }]
          });
        }
      });
      setBankStatement(prev => [...prev, ...newBankRecords].sort(sortByDate));
      setConfirmedMatches(prev => [...prev, ...newConfirmedMatches]);
      alert(`นำเข้าสำเร็จ: รายการธนาคาร ${newBankRecords.length} รายการ และกระทบยอดแล้ว ${newConfirmedMatches.length} กลุ่ม`);
    };
    reader.readAsArrayBuffer(file);
  };

  // --- Calculations ---
  const internalSum = useMemo(() => selectedInternal.reduce((acc, curr) => acc + curr.amount, 0), [selectedInternal]);
  const bankSum = useMemo(() => selectedBank.reduce((acc, curr) => acc + curr.amount, 0), [selectedBank]);
  const diff = Math.abs(internalSum - bankSum);

  const filteredInternal = useMemo(() => internalRecords.filter(item => (searchInternal === '' || Math.abs(item.amount).toString().includes(searchInternal) || item.docNo.toLowerCase().includes(searchInternal.toLowerCase())) && (!internalStartDate || parseDisplayDate(item.date) >= new Date(internalStartDate)) && (!internalEndDate || parseDisplayDate(item.date) <= new Date(internalEndDate))), [internalRecords, searchInternal, internalStartDate, internalEndDate]);
  const filteredBank = useMemo(() => bankStatement.filter(item => (searchBank === '' || Math.abs(item.amount).toString().includes(searchBank) || item.docNo.toLowerCase().includes(searchBank.toLowerCase())) && (!bankStartDate || parseDisplayDate(item.date) >= new Date(bankStartDate)) && (!bankEndDate || parseDisplayDate(item.date) <= new Date(bankEndDate))), [bankStatement, searchBank, bankStartDate, bankEndDate]);

  const toggleSelection = (item, type) => {
    if (type === 'internal') setSelectedInternal(prev => prev.some(i => i.id === item.id) ? prev.filter(i => i.id !== item.id) : [...prev, item]);
    else setSelectedBank(prev => prev.some(i => i.id === item.id) ? prev.filter(i => i.id !== item.id) : [...prev, item]);
  };

  const confirmMatch = () => {
    if (diff < 0.01 && (selectedInternal.length > 0 && selectedBank.length > 0)) {
      setConfirmedMatches(prev => [{ id: Date.now(), internals: [...selectedInternal], banks: [...selectedBank], totalAmount: internalSum }, ...prev]);
      setInternalRecords(prev => prev.filter(item => !selectedInternal.some(s => s.id === item.id)));
      setBankStatement(prev => prev.filter(item => !selectedBank.some(s => s.id === item.id)));
      setSelectedInternal([]); setSelectedBank([]);
    }
  };

  const exportToExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Report');
    const accountingFormat = '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
    worksheet.addRow(["#", "วันที่", "เลขที่เอกสาร/การจับคู่", "รายละเอียด", "ยอดเงิน", "สถานะ"]).font = { bold: true };
    const combinedData = [];
    confirmedMatches.forEach(m => m.banks.forEach(b => combinedData.push({ ...b, matched: m.internals.map(i => i.docNo).join(', '), status: "กระทบยอดแล้ว" })));
    bankStatement.forEach(b => combinedData.push({ ...b, matched: "", status: "ยังไม่กระทบยอด" }));
    combinedData.sort(sortByDate).forEach((e, i) => {
      const row = worksheet.addRow([i + 1, e.date, e.matched, e.docNo, e.amount, e.status]);
      row.getCell(5).numFmt = accountingFormat;
      row.getCell(6).font = { color: { argb: e.status === "ยังไม่กระทบยอด" ? 'FFFF0000' : 'FF008000' }, bold: true };
    });
    const buffer = await workbook.xlsx.writeBuffer();
    const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([buffer])); a.download = `Report_${new Date().toISOString().split('T')[0]}.xlsx`; a.click();
  };

  return (
    <div className="min-h-screen bg-[#f1f5f9] p-4 md:p-6 font-sans text-slate-700">
      <div className="max-w-[1500px] mx-auto flex flex-col h-full">
        
        {/* Header */}
        <div className="flex justify-between items-center mb-6 bg-white p-5 rounded-3xl shadow-sm border border-slate-100">
          <h1 className="text-2xl font-black text-blue-900 italic uppercase">Bank Reconciliation</h1>
          <div className="flex items-center gap-3">
            <label className="flex items-center gap-2 bg-blue-50 text-blue-700 border border-blue-100 px-4 py-2 rounded-xl font-black text-xs hover:bg-blue-100 transition-all cursor-pointer shadow-sm">
              <History size={16} /> นำเข้ารายงานเดิม
              <input type="file" onChange={(e) => processReportFile(e.target.files[0])} className="hidden" accept=".xlsx, .xls" />
            </label>
            <button onClick={exportToExcel} className="flex items-center gap-2 bg-emerald-50 text-emerald-700 border border-emerald-100 px-4 py-2 rounded-xl font-black text-xs hover:bg-emerald-100 uppercase shadow-sm"><Download size={16} /> Export Excel</button>
            <button onClick={() => window.location.reload()} className="bg-white text-slate-400 border border-slate-200 px-4 py-2 rounded-xl font-bold text-xs hover:text-red-500 hover:bg-red-50 uppercase shadow-sm">ล้างทั้งหมด</button>
          </div>
        </div>

        {/* Tab Switcher */}
        <div className="flex gap-4 mb-6 ml-2 items-center">
          <button onClick={() => setActiveTab('reconcile')} className={`px-8 py-2.5 rounded-full font-black text-xs transition-all ${activeTab === 'reconcile' ? 'bg-blue-600 text-white shadow-xl' : 'text-slate-400'}`}>รอกระทบยอด</button>
          <button onClick={() => setActiveTab('confirmed')} className={`px-8 py-2.5 rounded-full font-black text-xs transition-all flex items-center gap-2 ${activeTab === 'confirmed' ? 'bg-blue-600 text-white shadow-xl' : 'text-slate-400'}`}>
            กระทบยอดแล้ว {confirmedMatches.length > 0 && <span className="bg-orange-500 text-white px-1.5 py-0.5 rounded-full text-[8px]">{confirmedMatches.length}</span>}
          </button>
          {(selectedInternal.length > 0 || selectedBank.length > 0) && (
            <button onClick={() => {setSelectedInternal([]); setSelectedBank([])}} className="ml-auto text-rose-500 font-black text-[10px] uppercase flex items-center gap-1 hover:bg-rose-50 px-3 py-1.5 rounded-xl transition-all"><FilterX size={14} /> ล้างการเลือก</button>
          )}
        </div>

        {activeTab === 'reconcile' ? (
          <div className="space-y-6">
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 h-[600px]">
              
              {/* Internal Section */}
              <div 
                onDragOver={(e) => {e.preventDefault(); setIsDraggingInternal(true)}} 
                onDragLeave={() => setIsDraggingInternal(false)} 
                onDrop={(e) => {e.preventDefault(); setIsDraggingInternal(false); processFile(e.dataTransfer.files[0], 'internal')}}
                className={`bg-white rounded-[2.5rem] shadow-sm border-2 transition-all flex flex-col overflow-hidden relative ${isDraggingInternal ? 'border-blue-500 bg-blue-50 scale-[1.01]' : 'border-transparent'}`}
              >
                <div className="p-5 bg-blue-600 text-white space-y-4">
                  <div className="flex justify-between items-center">
                    <span className="font-black text-[15px] uppercase flex items-center gap-2"><Database size={18}/> บันทึกบัญชี ({internalRecords.length})</span>
                    <div className="flex gap-2">
                      {internalRecords.length > 0 && <button onClick={() => {if(confirm('ล้างข้อมูลฝั่งบัญชี?')) setInternalRecords([]);}} className="bg-rose-500/20 px-3 py-1.5 rounded-xl border border-rose-500/30 hover:bg-rose-500"><Trash2 size={12}/></button>}
                      <label className="bg-white/20 px-4 py-1.5 rounded-xl cursor-pointer text-[10px] font-black border border-white/30 uppercase"><Plus size={12} className="inline mr-1"/> นำเข้า<input type="file" onChange={(e) => processFile(e.target.files[0], 'internal')} className="hidden" /></label>
                    </div>
                  </div>
                  <div className="flex gap-2">
                    <div className="relative flex-1"><Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-white/40" /><input type="text" placeholder="ค้นหา..." value={searchInternal} onChange={e => setSearchInternal(e.target.value)} className="w-full bg-white/10 rounded-xl pl-8 pr-2 py-2 text-[10px] outline-none placeholder:text-white/30" /></div>
                    <div className="flex bg-white/10 rounded-xl p-1 items-center border border-white/20"><Calendar size={12} className="ml-1 text-white/50"/><input type="date" value={internalStartDate} onChange={e => setInternalStartDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none" /><span className="opacity-30">-</span><input type="date" value={internalEndDate} onChange={e => setInternalEndDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none" /></div>
                  </div>
                </div>

                <div className="flex-1 overflow-y-auto p-4 space-y-2 bg-slate-50/30">
                  {/* Overlay ขณะลากไฟล์ */}
                  {isDraggingInternal && (
                    <div className="absolute inset-0 z-50 bg-blue-600/10 flex flex-col items-center justify-center pointer-events-none">
                      <div className="bg-blue-600 text-white p-8 rounded-[3rem] shadow-2xl flex flex-col items-center gap-3 animate-bounce">
                         <FileUp size={48} />
                         <span className="font-black text-lg uppercase tracking-widest">วางไฟล์ที่นี่</span>
                      </div>
                    </div>
                  )}

                  {/* รายการข้อมูล */}
                  {filteredInternal.map(item => (
                    <div key={item.id} onClick={() => toggleSelection(item, 'internal')} className={`p-4 rounded-2xl border-2 cursor-pointer flex justify-between items-center transition-all ${selectedInternal.some(i => i.id === item.id) ? 'border-blue-500 bg-blue-50 shadow-md' : 'border-white bg-white hover:border-blue-200'}`}>
                      <div className="flex flex-col min-w-0 pr-4"><span className="font-bold text-slate-800 text-xs truncate">{item.docNo}</span><span className="text-[10px] text-slate-400 italic truncate">{item.description}</span><span className="text-[9px] text-slate-400 font-bold uppercase">{item.date}</span></div>
                      <span className={`text-xl font-black tabular-nums ${item.amount < 0 ? 'text-red-500' : 'text-blue-600'}`}>{formatAccounting(item.amount)}</span>
                    </div>
                  ))}

                  {/* หน้าว่างเมื่อไม่มีข้อมูล */}
                  {internalRecords.length === 0 && !isDraggingInternal && (
                    <label className="h-full flex flex-col items-center justify-center text-slate-300 py-20 border-2 border-dashed border-slate-200 rounded-[3rem] bg-slate-50/50 cursor-pointer hover:bg-blue-50 transition-all group">
                       <FileUp size={64} className="opacity-40 group-hover:scale-110 transition-transform" />
                       <h3 className="font-black text-slate-400 mt-4 uppercase group-hover:text-blue-500">Import Account File</h3>
                       <input type="file" onChange={(e) => processFile(e.target.files[0], 'internal')} className="hidden" />
                    </label>
                  )}
                </div>
              </div>

              {/* Bank Section */}
              <div 
                onDragOver={(e) => {e.preventDefault(); setIsDraggingBank(true)}} 
                onDragLeave={() => setIsDraggingBank(false)} 
                onDrop={(e) => {e.preventDefault(); setIsDraggingBank(false); processFile(e.dataTransfer.files[0], 'bank')}}
                className={`bg-white rounded-[2.5rem] shadow-sm border-2 transition-all flex flex-col overflow-hidden relative ${isDraggingBank ? 'border-slate-800 bg-slate-100 scale-[1.01]' : 'border-transparent'}`}
              >
                <div className="p-5 bg-slate-800 text-white space-y-4">
                  <div className="flex justify-between items-center">
                    <span className="font-black text-[15px] uppercase flex items-center gap-2"><Landmark size={18}/> ธนาคาร ({bankStatement.length})</span>
                    <div className="flex gap-2">
                      {bankStatement.length > 0 && <button onClick={() => {if(confirm('ล้างข้อมูลฝั่งธนาคาร?')) setBankStatement([]);}} className="bg-rose-500/10 px-3 py-1.5 rounded-xl border border-rose-500/20 hover:bg-rose-500"><Trash2 size={12}/></button>}
                      <label className="bg-white/10 px-4 py-1.5 rounded-xl cursor-pointer text-[10px] font-black border border-white/10 uppercase"><Plus size={12} className="inline mr-1"/> นำเข้า<input type="file" onChange={(e) => processFile(e.target.files[0], 'bank')} className="hidden" /></label>
                    </div>
                  </div>
                  <div className="flex gap-2">
                    <div className="relative flex-1"><Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-white/40" /><input type="text" placeholder="ค้นหา..." value={searchBank} onChange={e => setSearchBank(e.target.value)} className="w-full bg-white/5 border border-white/10 rounded-xl pl-8 pr-2 py-2 text-[10px] outline-none placeholder:text-white/20" /></div>
                    <div className="flex bg-white/5 rounded-xl p-1 items-center border border-white/10 text-white/50"><Calendar size={12}/><input type="date" value={bankStartDate} onChange={e => setBankStartDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none" />- <input type="date" value={bankEndDate} onChange={e => setBankEndDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none" /></div>
                  </div>
                </div>

                <div className="flex-1 overflow-y-auto p-4 space-y-2 bg-slate-50/30">
                  {/* Overlay ขณะลากไฟล์ */}
                  {isDraggingBank && (
                    <div className="absolute inset-0 z-50 bg-slate-800/10 flex flex-col items-center justify-center pointer-events-none">
                      <div className="bg-slate-800 text-white p-8 rounded-[3rem] shadow-2xl flex flex-col items-center gap-3 animate-bounce">
                         <FileUp size={48} />
                         <span className="font-black text-lg uppercase tracking-widest">วางไฟล์ที่นี่</span>
                      </div>
                    </div>
                  )}

                  {/* รายการข้อมูล */}
                  {filteredBank.map(item => (
                    <div key={item.id} onClick={() => toggleSelection(item, 'bank')} className={`p-4 rounded-2xl border-2 cursor-pointer flex justify-between items-center transition-all ${selectedBank.some(i => i.id === item.id) ? 'border-slate-800 bg-slate-100 shadow-lg' : 'border-white bg-white hover:border-slate-300'}`}>
                      <div className="min-w-0 pr-4"><span className="font-bold text-slate-700 text-xs line-clamp-1">{item.docNo}</span><span className="text-[9px] text-slate-400 font-bold uppercase">{item.date}</span></div>
                      <span className={`text-xl font-black tabular-nums ${item.amount < 0 ? 'text-red-500' : 'text-slate-900'}`}>{formatAccounting(item.amount)}</span>
                    </div>
                  ))}

                  {/* หน้าว่างเมื่อไม่มีข้อมูล */}
                  {bankStatement.length === 0 && !isDraggingBank && (
                    <label className="h-full flex flex-col items-center justify-center text-slate-300 py-20 border-2 border-dashed border-slate-200 rounded-[3rem] bg-slate-50/50 cursor-pointer hover:bg-slate-100 transition-all group">
                       <FileUp size={64} className="opacity-40 group-hover:scale-110 transition-transform" />
                       <h3 className="font-black text-slate-400 mt-4 uppercase group-hover:text-slate-700">Import Bank File</h3>
                       <input type="file" onChange={(e) => processFile(e.target.files[0], 'bank')} className="hidden" />
                    </label>
                  )}
                </div>
              </div>
            </div>

            {/* Summary Bar */}
            <div className="bg-white p-8 rounded-[3.5rem] shadow-xl flex flex-col md:flex-row justify-around items-center border border-slate-100 gap-6">
              <div className="text-center"><div className="text-slate-400 text-[10px] font-black uppercase">บัญชีที่เลือก</div><div className="text-5xl font-black text-blue-600">{formatAccounting(internalSum)}</div></div>
              <div className="flex flex-col items-center bg-slate-50 px-16 py-6 rounded-[2.5rem] border min-w-[380px]">
                <div className="text-slate-400 text-[10px] font-black uppercase mb-1">ผลต่าง</div>
                <div className={`text-6xl font-black ${diff < 0.01 ? 'text-emerald-500' : 'text-red-500'}`}>{formatAccounting(diff)}</div>
                {diff < 0.01 && (selectedInternal.length > 0 && selectedBank.length > 0) && (
                  <button onClick={confirmMatch} className="mt-5 bg-blue-600 text-white px-12 py-3.5 rounded-full font-black text-xs hover:bg-blue-700 shadow-2xl animate-bounce flex items-center gap-2">ยืนยันจับคู่ <CheckCircle2 size={16}/></button>
                )}
              </div>
              <div className="text-center"><div className="text-slate-400 text-[10px] font-black uppercase">ธนาคารที่เลือก</div><div className="text-5xl font-black text-slate-900">{formatAccounting(bankSum)}</div></div>
            </div>
          </div>
        ) : (
          /* History View (Confirmed) */
          <div className="bg-white rounded-[3rem] shadow-sm border border-slate-200 overflow-hidden min-h-[550px] flex flex-col">
            <div className="p-6 bg-slate-50 border-b grid grid-cols-4 font-black text-[11px] text-slate-400 uppercase"><span>รายการบัญชี</span><span className="text-center">ยอดเงิน</span><span className="pl-8">รายการธนาคาร</span><span className="text-right">ยอดเงิน</span></div>
            <div className="p-8 space-y-6 overflow-y-auto flex-1 bg-slate-50/20">
              {confirmedMatches.map(m => (
                <div key={m.id} className="group bg-white border border-slate-100 rounded-[2rem] p-8 grid grid-cols-4 items-center shadow-sm relative hover:border-blue-300 transition-all">
                  <div className="space-y-3">{m.internals.map(i => <div key={i.id} className="flex flex-col"><span className="text-[10px] text-slate-400 font-black">{i.date}</span><span className="text-sm font-bold text-blue-700">{i.docNo}</span></div>)}</div>
                  <div className="text-center font-black text-2xl border-r">{formatAccounting(m.totalAmount)}</div>
                  <div className="pl-8 space-y-3">{m.banks.map(b => <div key={b.id} className="flex flex-col"><span className="text-[10px] text-slate-400 font-black">{b.date}</span><span className="text-sm font-bold text-slate-800">{b.docNo}</span></div>)}</div>
                  <div className="text-right font-black text-2xl">{formatAccounting(m.totalAmount)}</div>
                  <button onClick={() => { setConfirmedMatches(prev => prev.filter(x => x.id !== m.id)); setInternalRecords(p => [...p, ...m.internals].sort(sortByDate)); setBankStatement(p => [...p, ...m.banks].sort(sortByDate)); }} className="absolute -right-3 -top-3 bg-white text-rose-500 border-2 rounded-full p-2.5 opacity-0 group-hover:opacity-100 shadow-xl transition-all"><Trash2 size={18}/></button>
                </div>
              ))}
              {confirmedMatches.length === 0 && <div className="h-full flex flex-col items-center justify-center text-slate-300 py-20 uppercase font-black opacity-40">ไม่มีรายการที่กระทบยอดแล้ว</div>}
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default BankReconciliationApp;