import React, { useState, useMemo } from 'react';
import {  Download,  Search,  Plus,  Database, Landmark, FileUp, Trash2, ArrowRightLeft, FilterX, Calendar, CheckCircle2 } from 'lucide-react';
import * as XLSX from 'xlsx'; // ต้องติดตั้ง npm install xlsx

const BankReconciliation = () => {
  // --- States ---
  const [internalRecords, setInternalRecords] = useState([]);
  const [bankStatement, setBankStatement] = useState([]);
  const [selectedInternal, setSelectedInternal] = useState([]);
  const [selectedBank, setSelectedBank] = useState([]);
  const [confirmedMatches, setConfirmedMatches] = useState([]);
  const [activeTab, setActiveTab] = useState('reconcile'); // 'reconcile' | 'confirmed'

  // Search & Filter States
  const [searchInternal, setSearchInternal] = useState('');
  const [searchBank, setSearchBank] = useState('');
  const [internalStartDate, setInternalStartDate] = useState('');
  const [internalEndDate, setInternalEndDate] = useState('');
  const [bankStartDate, setBankStartDate] = useState('');
  const [bankEndDate, setBankEndDate] = useState('');

  // Drag and Drop States
  const [isDraggingInternal, setIsDraggingInternal] = useState(false);
  const [isDraggingBank, setIsDraggingBank] = useState(false);

  // --- Helpers ---
  const formatAccounting = (num) => {
    return new Intl.NumberFormat('th-TH', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(num);
  };

  const sortByDate = (a, b) => new Date(a.date) - new Date(b.date);

  // --- Logic Functions ---
  const processFile = (file, type) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(worksheet);

      // Mapping ข้อมูล (ปรับตามโครงสร้าง Excel ของคุณ)
      const mappedData = json.map((row, index) => ({
        id: `${type}-${Date.now()}-${index}`,
        date: row.Date || row['วันที่'] || '',
        docNo: row.DocNo || row['เลขที่เอกสาร'] || row['Description'] || 'N/A',
        description: row.Description || row['รายละเอียด'] || '',
        amount: parseFloat(row.Amount || row['ยอดเงิน'] || 0),
        status: row.Status || ''
      }));

      if (type === 'internal') {
        setInternalRecords(prev => [...prev, ...mappedData].sort(sortByDate));
      } else {
        setBankStatement(prev => [...prev, ...mappedData].sort(sortByDate));
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const toggleSelection = (item, type) => {
    if (type === 'internal') {
      setSelectedInternal(prev => 
        prev.find(i => i.id === item.id) 
          ? prev.filter(i => i.id !== item.id) 
          : [...prev, item]
      );
    } else {
      setSelectedBank(prev => 
        prev.find(i => i.id === item.id) 
          ? prev.filter(i => i.id !== item.id) 
          : [...prev, item]
      );
    }
  };

  const confirmMatch = () => {
    const newMatch = {
      id: Date.now(),
      internals: [...selectedInternal],
      banks: [...selectedBank],
      totalAmount: selectedInternal.reduce((sum, i) => sum + i.amount, 0),
      matchDate: new Date().toLocaleDateString()
    };

    setConfirmedMatches(prev => [newMatch, ...prev]);
    setInternalRecords(prev => prev.filter(i => !selectedInternal.find(s => s.id === i.id)));
    setBankStatement(prev => prev.filter(b => !selectedBank.find(s => s.id === b.id)));
    setSelectedInternal([]);
    setSelectedBank([]);
  };

  const exportToExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(confirmedMatches.map(m => ({
      'Match ID': m.id,
      'Internal Docs': m.internals.map(i => i.docNo).join(', '),
      'Bank Docs': m.banks.map(b => b.docNo).join(', '),
      'Amount': m.totalAmount,
      'Date': m.matchDate
    })));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Reconciled");
    XLSX.writeFile(workbook, "reconciliation_report.xlsx");
  };

  // --- Drag & Drop Handlers ---
  const onDragOver = (e, type) => {
    e.preventDefault();
    if (type === 'internal') setIsDraggingInternal(true);
    else setIsDraggingBank(true);
  };

  const onDragLeave = (e, type) => {
    if (type === 'internal') setIsDraggingInternal(false);
    else setIsDraggingBank(false);
  };

  const onDrop = (e, type) => {
    e.preventDefault();
    setIsDraggingInternal(false);
    setIsDraggingBank(false);
    const file = e.dataTransfer.files[0];
    processFile(file, type);
  };

  // --- Filtering ---
  const filteredInternal = internalRecords.filter(item => 
    item.docNo.toLowerCase().includes(searchInternal.toLowerCase()) &&
    (!internalStartDate || item.date >= internalStartDate) &&
    (!internalEndDate || item.date <= internalEndDate)
  );

  const filteredBank = bankStatement.filter(item => 
    item.docNo.toLowerCase().includes(searchBank.toLowerCase()) &&
    (!bankStartDate || item.date >= bankStartDate) &&
    (!bankEndDate || item.date <= bankEndDate)
  );

  // --- Calculations ---
  const internalSum = selectedInternal.reduce((sum, i) => sum + i.amount, 0);
  const bankSum = selectedBank.reduce((sum, i) => sum + i.amount, 0);
  const diff = Math.abs(internalSum - bankSum);

  return (
    <div className="min-h-screen bg-[#f1f5f9] p-4 md:p-6 font-sans text-slate-700">
      <div className="max-w-[1500px] mx-auto flex flex-col h-full">
        
        {/* Header */}
        <div className="flex justify-between items-center mb-6 bg-white p-5 rounded-3xl shadow-sm border border-slate-100">
          <h1 className="text-2xl font-black text-blue-900 italic uppercase">Bank Reconciliation</h1>
          <div className="flex items-center gap-3">
            <button onClick={exportToExcel} className="flex items-center gap-2 bg-emerald-50 text-emerald-700 border border-emerald-100 px-4 py-2 rounded-xl font-black text-xs hover:bg-emerald-100 transition-all uppercase">
              <Download size={16} /> Export Excel
            </button>
            <div className="w-px h-6 bg-slate-200 mx-1"></div>
            <button onClick={() => window.location.reload()} className="bg-white text-slate-400 border border-slate-200 px-4 py-2 rounded-xl font-bold text-xs hover:text-red-500 hover:bg-red-50 transition-all uppercase">
              ล้างทั้งหมด
            </button>
          </div>
        </div>

        {/* Tab Switcher */}
        <div className="flex gap-4 mb-6 ml-2 items-center">
          <button onClick={() => setActiveTab('reconcile')} className={`px-8 py-2.5 rounded-full font-black text-xs transition-all ${activeTab === 'reconcile' ? 'bg-blue-600 text-white shadow-xl' : 'text-slate-400'}`}>รอกระทบยอด</button>
          <button onClick={() => setActiveTab('confirmed')} className={`px-8 py-2.5 rounded-full font-black text-xs transition-all flex items-center gap-2 ${activeTab === 'confirmed' ? 'bg-blue-600 text-white shadow-xl' : 'text-slate-400'}`}>
            กระทบยอดแล้ว {confirmedMatches.length > 0 && <span className="bg-orange-500 text-white px-1.5 py-0.5 rounded-full text-[8px]">{confirmedMatches.length}</span>}
          </button>
          
          {(selectedInternal.length > 0 || selectedBank.length > 0) && (
            <button onClick={() => {setSelectedInternal([]); setSelectedBank([])}} className="ml-auto text-rose-500 font-black text-[10px] uppercase flex items-center gap-1 hover:bg-rose-50 px-3 py-1.5 rounded-xl transition-all">
               <FilterX size={14} /> ล้างการเลือก
            </button>
          )}
        </div>

        {/* Main Content Area */}
        <div className="flex-1">
          {activeTab === 'reconcile' ? (
            <div className="space-y-6">
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 h-[600px]">
                
                {/* 1. ฝั่งบันทึกบัญชี (Internal) */}
                <div 
                  onDragOver={(e) => onDragOver(e, 'internal')}
                  onDragLeave={(e) => onDragLeave(e, 'internal')}
                  onDrop={(e) => onDrop(e, 'internal')}
                  className={`bg-white rounded-[2.5rem] shadow-sm border-2 transition-all flex flex-col overflow-hidden relative ${isDraggingInternal ? 'border-blue-500 border-dashed bg-blue-50 scale-[1.01]' : 'border-transparent'}`}
                >
                  <div className="p-5 bg-blue-600 text-white space-y-4">
                    <div className="flex justify-between items-center">
                      <span className="font-black text-[15px] uppercase tracking-widest flex items-center gap-2">
                        <Database size={18}/> บันทึกบัญชี ({internalRecords.length})
                      </span>
                      <div className="flex gap-2">
                        {/* ปุ่มล้างฝั่งบัญชี */}
                        {internalRecords.length > 0 && (
                          <button 
                            onClick={() => { if(confirm('ล้างข้อมูลฝั่งบัญชีทั้งหมด?')) { setInternalRecords([]); setSelectedInternal([]); }}}
                            className="bg-rose-500/20 px-3 py-1.5 rounded-xl text-[10px] font-black border border-rose-500/30 hover:bg-rose-500 transition-all uppercase flex items-center gap-1"
                          >
                            <Trash2 size={12} /> ล้าง
                          </button>
                        )}
                        <label htmlFor="internal-upload-btn" className="bg-white/20 px-4 py-1.5 rounded-xl cursor-pointer text-[10px] font-black border border-white/30 hover:bg-white/40 transition-all uppercase">
                          <Plus size={12} className="inline mr-1"/> นำเข้า
                          <input id="internal-upload-btn" type="file" onChange={(e) => processFile(e.target.files[0], 'internal')} className="hidden" accept=".xlsx, .xls" />
                        </label>
                      </div>
                    </div>
                    <div className="flex gap-2">
                      <div className="relative flex-1"><Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-white/40" /><input type="text" placeholder="ค้นหา..." value={searchInternal} onChange={e => setSearchInternal(e.target.value)} className="w-full bg-white/10 border border-white/20 rounded-xl pl-8 pr-8 py-2 text-[10px] outline-none placeholder:text-white/30" /></div>
                      <div className="flex bg-white/10 rounded-xl p-1 items-center border border-white/20"><Calendar size={12} className="ml-2 text-white/50" /><input type="date" value={internalStartDate} onChange={e => setInternalStartDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none" /><span className="text-white/50">-</span><input type="date" value={internalEndDate} onChange={e => setInternalEndDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none" /></div>
                    </div>
                  </div>
                  
                  <div className="flex-1 overflow-y-auto p-4 space-y-2 bg-slate-50/30">
                    {isDraggingInternal && (
                      <div className="absolute inset-0 z-50 bg-blue-600/10 flex flex-col items-center justify-center pointer-events-none animate-in fade-in">
                        <div className="bg-blue-600 text-white p-8 rounded-[3rem] shadow-2xl flex flex-col items-center gap-3 animate-bounce">
                           <FileUp size={48} />
                           <span className="font-black text-lg uppercase tracking-widest">วางไฟล์ที่นี่</span>
                        </div>
                      </div>
                    )}

                    {filteredInternal.map(item => (
                      <div key={item.id} onClick={() => toggleSelection(item, 'internal')} className={`p-4 rounded-2xl border-2 transition-all cursor-pointer min-h-[85px] flex items-center ${selectedInternal.some(i => i.id === item.id) ? 'border-blue-500 bg-blue-50 shadow-md' : 'border-white bg-white shadow-sm hover:border-blue-200'}`}>
                        <div className="flex justify-between items-center w-full">
                          <div className="flex flex-col gap-1 flex-1 min-w-0 pr-4">
                            <div className="flex items-center gap-2">
                              <span className="font-bold text-slate-800 text-xs truncate">{item.docNo}</span>
                              {item.status && <span className="text-[8px] px-1.5 py-0.5 rounded-full bg-blue-100 text-blue-600 font-black">{item.status}</span>}
                            </div>
                            <span className="text-[10px] text-slate-400 italic truncate">{item.description || '-'}</span>
                            <span className="text-[9px] text-slate-400 font-bold uppercase">{item.date}</span>
                          </div>
                          <span className={`text-xl font-black tabular-nums ${item.amount < 0 ? 'text-red-500' : 'text-blue-600'}`}>{formatAccounting(item.amount)}</span>
                        </div>
                      </div>
                    ))}

                    {internalRecords.length === 0 && !isDraggingInternal && (
                      <label htmlFor="internal-main-upload" className="h-full flex flex-col items-center justify-center text-slate-300 py-20 border-2 border-dashed border-slate-200 rounded-[3rem] bg-slate-50/50 cursor-pointer hover:bg-blue-50/50 hover:border-blue-300 transition-all group">
                        <input id="internal-main-upload" type="file" onChange={(e) => processFile(e.target.files[0], 'internal')} className="hidden" accept=".xlsx, .xls" />
                        <FileUp size={64} strokeWidth={1} className="opacity-40 group-hover:scale-110 transition-transform" />
                        <h3 className="font-black text-slate-400 mt-4 uppercase tracking-widest group-hover:text-blue-500">Import Account File</h3>
                      </label>
                    )}
                  </div>
                </div>

                {/* 2. ฝั่งธนาคาร (Bank) */}
                <div 
                  onDragOver={(e) => onDragOver(e, 'bank')}
                  onDragLeave={(e) => onDragLeave(e, 'bank')}
                  onDrop={(e) => onDrop(e, 'bank')}
                  className={`bg-white rounded-[2.5rem] shadow-sm border-2 transition-all flex flex-col overflow-hidden relative ${isDraggingBank ? 'border-slate-800 border-dashed bg-slate-100 scale-[1.01]' : 'border-transparent'}`}
                >
                  <div className="p-5 bg-slate-800 text-white space-y-4">
                    <div className="flex justify-between items-center">
                      <span className="font-black text-[15px] uppercase tracking-widest text-slate-300 flex items-center gap-2">
                        <Landmark size={18}/> รายการธนาคาร ({bankStatement.length})
                      </span>
                      <div className="flex gap-2">
                        {/* ปุ่มล้างฝั่งธนาคาร */}
                        {bankStatement.length > 0 && (
                          <button 
                            onClick={() => { if(confirm('ล้างข้อมูลฝั่งธนาคารทั้งหมด?')) { setBankStatement([]); setSelectedBank([]); }}}
                            className="bg-rose-500/10 px-3 py-1.5 rounded-xl text-[10px] font-black border border-rose-500/20 hover:bg-rose-500 transition-all uppercase flex items-center gap-1"
                          >
                            <Trash2 size={12} /> ล้าง
                          </button>
                        )}
                        <label htmlFor="bank-upload-btn" className="bg-white/10 px-4 py-1.5 rounded-xl cursor-pointer text-[10px] font-black border border-white/10 hover:bg-white/20 transition-all uppercase">
                          <Plus size={12} className="inline mr-1"/> นำเข้า
                          <input id="bank-upload-btn" type="file" onChange={(e) => processFile(e.target.files[0], 'bank')} className="hidden" accept=".xlsx, .xls" />
                        </label>
                      </div>
                    </div>
                    <div className="flex gap-2">
                      <div className="relative flex-1"><Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-white/40" /><input type="text" placeholder="ค้นหา..." value={searchBank} onChange={e => setSearchBank(e.target.value)} className="w-full bg-white/5 border border-white/10 rounded-xl pl-8 pr-8 py-2 text-[10px] outline-none placeholder:text-white/20" /></div>
                      <div className="flex bg-white/5 rounded-xl p-1 items-center border border-white/10"><Calendar size={12} className="text-white/30" /><input type="date" value={bankStartDate} onChange={e => setBankStartDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none opacity-60" /><span className="text-white/10">-</span><input type="date" value={bankEndDate} onChange={e => setBankEndDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none opacity-60" /></div>
                    </div>
                  </div>

                  <div className="flex-1 overflow-y-auto p-4 space-y-2 bg-slate-50/30">
                    {isDraggingBank && (
                      <div className="absolute inset-0 z-50 bg-slate-900/10 flex flex-col items-center justify-center pointer-events-none animate-in fade-in">
                        <div className="bg-slate-800 text-white p-8 rounded-[3rem] shadow-2xl flex flex-col items-center gap-3 animate-bounce">
                           <FileUp size={48} />
                           <span className="font-black text-lg uppercase tracking-widest">วางไฟล์ที่นี่</span>
                        </div>
                      </div>
                    )}

                    {filteredBank.map(item => (
                      <div key={item.id} onClick={() => toggleSelection(item, 'bank')} className={`p-4 rounded-2xl border-2 transition-all cursor-pointer min-h-[85px] flex items-center ${selectedBank.some(i => i.id === item.id) ? 'border-slate-800 bg-slate-100 shadow-lg' : 'border-white bg-white shadow-sm hover:border-slate-300'}`}>
                        <div className="flex justify-between items-center w-full">
                          <div className="flex flex-col gap-1 flex-1 min-w-0 pr-4">
                            <span className="font-bold text-slate-700 text-xs line-clamp-1">{item.docNo}</span>
                            <span className="text-[9px] text-slate-400 font-bold uppercase">{item.date}</span>
                          </div>
                          <span className={`text-xl font-black tabular-nums ${item.amount < 0 ? 'text-red-500' : 'text-slate-900'}`}>{formatAccounting(item.amount)}</span>
                        </div>
                      </div>
                    ))}

                    {bankStatement.length === 0 && !isDraggingBank && (
                      <label htmlFor="bank-main-upload" className="h-full flex flex-col items-center justify-center text-slate-300 py-20 border-2 border-dashed border-slate-200 rounded-[3rem] bg-slate-50/50 cursor-pointer hover:bg-slate-100 hover:border-slate-400 transition-all group">
                        <input id="bank-main-upload" type="file" onChange={(e) => processFile(e.target.files[0], 'bank')} className="hidden" accept=".xlsx, .xls" />
                        <FileUp size={64} strokeWidth={1} className="opacity-40 group-hover:scale-110 transition-transform" />
                        <h3 className="font-black text-slate-400 mt-4 uppercase tracking-widest group-hover:text-slate-700">Import Bank File</h3>
                      </label>
                    )}
                  </div>
                </div>
              </div>

              {/* Summary Bar */}
              <div className="bg-white p-8 rounded-[3.5rem] shadow-xl flex flex-col md:flex-row justify-around items-center border border-slate-100 gap-6">
                <div className="text-center"><div className="text-slate-400 text-[10px] font-black uppercase tracking-widest">บัญชีที่เลือก</div><div className="text-5xl font-black text-blue-600 tracking-tighter tabular-nums">{formatAccounting(internalSum)}</div></div>
                <div className="flex flex-col items-center bg-slate-50 px-16 py-6 rounded-[2.5rem] border shadow-inner min-w-[380px]">
                  <div className="text-slate-400 text-[10px] font-black uppercase mb-1 tracking-widest">ผลต่าง</div>
                  <div className={`text-6xl font-black tabular-nums tracking-tighter ${diff < 0.01 ? 'text-emerald-500' : 'text-red-500'}`}>{formatAccounting(diff)}</div>
                  {diff < 0.01 && (selectedInternal.length > 0 && selectedBank.length > 0) && (
                    <button onClick={confirmMatch} className="mt-5 bg-blue-600 text-white px-12 py-3.5 rounded-full font-black text-xs hover:bg-blue-700 transition-all flex items-center gap-2 shadow-2xl animate-bounce tracking-widest uppercase">
                      ยืนยันจับคู่ <CheckCircle2 size={16}/>
                    </button>
                  )}
                </div>
                <div className="text-center"><div className="text-slate-400 text-[10px] font-black uppercase tracking-widest">ธนาคารที่เลือก</div><div className="text-5xl font-black text-slate-900 tracking-tighter tabular-nums">{formatAccounting(bankSum)}</div></div>
              </div>
            </div>
          ) : (
            /* Confirmed Tab Content */
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
                      <button 
                        onClick={() => { 
                          setConfirmedMatches(prev => prev.filter(x => x.id !== m.id)); 
                          setInternalRecords(p => [...p, ...m.internals].sort(sortByDate)); 
                          setBankStatement(p => [...p, ...m.banks].sort(sortByDate)); 
                        }} 
                        className="absolute -right-3 -top-3 bg-white text-rose-500 border-2 border-rose-50 rounded-full p-2.5 opacity-0 group-hover:opacity-100 transition-all shadow-xl hover:bg-rose-500 hover:text-white"
                      >
                        <Trash2 size={18}/>
                      </button>
                    </div>
                  ))}
                  {confirmedMatches.length === 0 && <div className="h-full flex flex-col items-center justify-center text-slate-300 py-20 uppercase font-black tracking-widest opacity-40">ไม่มีรายการที่กระทบยอดแล้ว</div>}
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default BankReconciliation;