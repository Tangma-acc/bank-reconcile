import React, { useState, useMemo } from 'react';
import { 
  Search, Plus, Calendar, X, Download, 
  Trash2, ArrowRightLeft, FileSpreadsheet, CheckCircle2 
} from 'lucide-react';
import * as XLSX from 'xlsx';

const BankReconcileApp = () => {
  // States สำหรับข้อมูล
  const [internalRecords, setInternalRecords] = useState([]);
  const [bankStatement, setBankStatement] = useState([]);
  const [selectedInternal, setSelectedInternal] = useState([]);
  const [selectedBank, setSelectedBank] = useState([]);
  const [confirmedMatches, setConfirmedMatches] = useState([]);
  const [activeTab, setActiveTab] = useState('reconcile');

  // States สำหรับ Filter และ Search
  const [searchInternal, setSearchInternal] = useState('');
  const [searchBank, setSearchBank] = useState('');
  const [internalStartDate, setInternalStartDate] = useState('');
  const [internalEndDate, setInternalEndDate] = useState('');
  const [bankStartDate, setBankStartDate] = useState('');
  const [bankEndDate, setBankEndDate] = useState('');

  // --- ฟังก์ชันดาวน์โหลด Template สำหรับนำเข้า ---
  const downloadTemplate = () => {
    const worksheetData = [
      ["ลำดับที่*", "วันที่", "เลขที่เอกสาร", "คำอธิบาย", "ต้องชำระ"], // Header ตามรูปภาพ
      [1, "13/07/2026", "RT-20260700010", "รับชำระค่าบริการ", 14850.00], // ตัวอย่างข้อมูล
      [2, "13/07/2026", "RT-20260700009", "รับชำระค่าสินค้า", 14850.00]
    ];
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(worksheetData);
    ws['!cols'] = [{ wch: 10 }, { wch: 12 }, { wch: 20 }, { wch: 35 }, { wch: 15 }];
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, "Template_บันทึกบัญชี.xlsx");
  };

  // --- ฟังก์ชันนำเข้าไฟล์ Excel ---
  const handleFileUpload = (e, type) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);

      const formatted = data.map((row, idx) => ({
        id: `${type}-${idx}-${Date.now()}`,
        date: row['วันที่'] || '',
        docNo: row['เลขที่เอกสาร'] || row['รายการ'] || 'N/A',
        description: row['คำอธิบาย'] || '',
        amount: parseFloat(row['ต้องชำระ'] || row['ยอดเงิน'] || 0),
        status: row['สถานะ'] || ''
      }));

      if (type === 'internal') setInternalRecords(formatted);
      else setBankStatement(formatted);
    };
    reader.readAsBinaryString(file);
  };

  // --- ฟังก์ชันการเลือกรายการ ---
  const toggleSelection = (item, type) => {
    if (type === 'internal') {
      setSelectedInternal(prev => 
        prev.find(i => i.id === item.id) ? prev.filter(i => i.id !== item.id) : [...prev, item]
      );
    } else {
      setSelectedBank(prev => 
        prev.find(i => i.id === item.id) ? prev.filter(i => i.id !== item.id) : [...prev, item]
      );
    }
  };

  // --- ฟังก์ชันยืนยันการจับคู่ ---
  const confirmMatch = () => {
    const newMatch = {
      id: Date.now(),
      internals: [...selectedInternal],
      banks: [...selectedBank],
      totalAmount: selectedInternal.reduce((sum, i) => sum + i.amount, 0)
    };
    setConfirmedMatches([...confirmedMatches, newMatch]);
    setInternalRecords(prev => prev.filter(i => !selectedInternal.find(si => si.id === i.id)));
    setBankStatement(prev => prev.filter(b => !selectedBank.find(sb => sb.id === b.id)));
    setSelectedInternal([]);
    setSelectedBank([]);
  };

  // --- ฟังก์ชัน Export รายงานที่ Reconcile แล้ว ---
  const exportToExcel = () => {
    if (confirmedMatches.length === 0) return alert("ไม่มีข้อมูลที่ยืนยันแล้ว");
    const data = confirmedMatches.flatMap(m => m.internals.map(i => ({
      'วันที่': i.date,
      'เลขที่เอกสาร': i.docNo,
      'ยอดเงิน': i.amount,
      'สถานะ': 'จับคู่สำเร็จ'
    })));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Reconciled_Report");
    XLSX.writeFile(wb, "Bank_Reconcile_Report.xlsx");
  };

  // --- Logic การคำนวณและ Filter ---
  const filteredInternal = internalRecords.filter(item => 
    item.amount.toString().includes(searchInternal) &&
    (!internalStartDate || item.date >= internalStartDate) &&
    (!internalEndDate || item.date <= internalEndDate)
  );

  const filteredBank = bankStatement.filter(item => 
    item.amount.toString().includes(searchBank) &&
    (!bankStartDate || item.date >= bankStartDate) &&
    (!bankEndDate || item.date <= bankEndDate)
  );

  const internalSum = selectedInternal.reduce((sum, i) => sum + i.amount, 0);
  const bankSum = selectedBank.reduce((sum, i) => sum + i.amount, 0);
  const diff = Math.abs(internalSum - bankSum);

  return (
    <div className="min-h-screen bg-[#f1f5f9] p-4 md:p-6 font-sans text-slate-700">
      <div className="max-w-[1500px] mx-auto flex flex-col h-full">
        
        {/* Header Section */}
        <div className="flex justify-between items-center mb-6 bg-white p-5 rounded-3xl shadow-sm border">
           <h1 className="text-2xl font-black text-blue-900 italic uppercase tracking-tighter">Bank Reconcile</h1>
           <div className="flex items-center gap-3">
              <button onClick={downloadTemplate} className="flex items-center gap-2 bg-blue-50 text-blue-700 border border-blue-100 px-4 py-2 rounded-xl font-black text-xs hover:bg-blue-100 transition-all shadow-sm uppercase tracking-wider">
                <FileSpreadsheet size={16} /> Download Template
              </button>
              <div className="w-px h-6 bg-slate-200 mx-1"></div>
              <button onClick={exportToExcel} className="flex items-center gap-2 bg-emerald-50 text-emerald-700 border border-emerald-100 px-4 py-2 rounded-xl font-black text-xs hover:bg-emerald-100 transition-all shadow-sm uppercase tracking-wider">
                <Download size={16} /> Export Excel
              </button>
              <div className="w-px h-6 bg-slate-200 mx-1"></div>
              <button onClick={() => window.location.reload()} className="text-slate-400 font-bold text-xs px-4 py-2 hover:text-red-500 rounded-xl transition-all">ล้างข้อมูล</button>
           </div>
        </div>

        {/* Tabs Control */}
        <div className="flex gap-4 mb-6 ml-2">
          <button onClick={() => setActiveTab('reconcile')} className={`px-8 py-2.5 rounded-full font-black text-xs transition-all ${activeTab === 'reconcile' ? 'bg-blue-600 text-white shadow-xl' : 'text-slate-400'}`}>รอกระทบยอด</button>
          <button onClick={() => setActiveTab('confirmed')} className={`px-8 py-2.5 rounded-full font-black text-xs transition-all flex items-center gap-2 ${activeTab === 'confirmed' ? 'bg-blue-600 text-white shadow-xl' : 'text-slate-400'}`}>รอยืนยัน {confirmedMatches.length > 0 && <span className="bg-orange-500 text-white px-1.5 py-0.5 rounded-full text-[8px]">{confirmedMatches.length}</span>}</button>
        </div>

        <div className="flex-1">
          {activeTab === 'reconcile' ? (
            <div className="space-y-6">
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 h-[580px]">
                
                {/* Left: PEAK Records */}
                <div className="bg-white rounded-[2.5rem] shadow-sm border flex flex-col overflow-hidden">
                  <div className="p-5 bg-blue-600 text-white space-y-4">
                    <div className="flex justify-between items-center">
                        <span className="font-black text-[15px] uppercase tracking-widest">รายการบันทึกบัญชี ({internalRecords.length})</span>
                        <label className="bg-white/20 px-4 py-1.5 rounded-xl cursor-pointer text-[10px] font-black border border-white/30 hover:bg-white/40 transition-all uppercase">
                            <Plus size={12} className="inline mr-1"/> นำเข้า
                            <input type="file" onChange={(e) => handleFileUpload(e, 'internal')} className="hidden" accept=".xlsx, .xls" />
                        </label>
                    </div>
                    <div className="flex gap-2">
                      <div className="relative flex-1">
                        <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-white/40" />
                        <input type="text" placeholder="ยอดเงิน..." value={searchInternal} onChange={e => setSearchInternal(e.target.value)} className="w-full bg-white/10 border border-white/20 rounded-xl pl-8 pr-8 py-2 text-[10px] outline-none" />
                        {searchInternal && <button onClick={() => setSearchInternal('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-white/40 hover:text-white"><X size={14}/></button>}
                      </div>
                      <div className="flex bg-white/10 rounded-xl p-1 items-center border border-white/20">
                        <Calendar size={12} className="ml-2 text-white/50" />
                        <input type="date" value={internalStartDate} onChange={e => setInternalStartDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none" />
                        <span className="text-white/50">-</span>
                        <input type="date" value={internalEndDate} onChange={e => setInternalEndDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none" />
                        {(internalStartDate || internalEndDate) && <button onClick={()=>{setInternalStartDate('');setInternalEndDate('');}} className="p-1 text-white"><X size={12}/></button>}
                      </div>
                    </div>
                  </div>
                  <div className="flex-1 overflow-y-auto p-4 space-y-2 bg-slate-50/30">
                    {filteredInternal.map((item) => (
                      <div key={item.id} onClick={() => toggleSelection(item, 'internal')} className={`p-3 rounded-2xl border-2 transition-all cursor-pointer ${selectedInternal.some(i => i.id === item.id) ? 'border-blue-500 bg-blue-50' : 'border-white bg-white shadow-sm'}`}>
                        <div className="flex justify-between items-start">
                          <div className="flex flex-col gap-0.5">
                            <span className="font-bold text-slate-800 text-xs">{item.docNo} {item.status && <span className="ml-1 text-[9px] px-1.5 py-0.5 rounded-full bg-blue-100 text-blue-600 font-black">{item.status}</span>}</span>
                            <span className="text-[10px] text-slate-400 italic truncate max-w-[200px]">{item.description}</span>
                            <span className="text-[9px] text-slate-400 font-bold uppercase">{item.date}</span>
                          </div>
                          <span className={`text-xl font-black tabular-nums ${item.amount < 0 ? 'text-red-500' : 'text-blue-600'}`}>{item.amount.toLocaleString(undefined, {minimumFractionDigits: 2})}</span>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>

                {/* Right: Bank Statement */}
                <div className="bg-white rounded-[2.5rem] shadow-sm border flex flex-col overflow-hidden">
                  <div className="p-5 bg-slate-800 text-white space-y-4">
                    <div className="flex justify-between items-center">
                        <span className="font-black text-[15px] uppercase tracking-widest text-slate-300">รายการธนาคาร ({bankStatement.length})</span>
                        <label className="bg-white/10 px-4 py-1.5 rounded-xl cursor-pointer text-[10px] font-black border border-white/10 hover:bg-white/20 transition-all uppercase">
                            <Plus size={12} className="inline mr-1"/> นำเข้า
                            <input type="file" onChange={(e) => handleFileUpload(e, 'bank')} className="hidden" accept=".xlsx, .xls" />
                        </label>
                    </div>
                    <div className="flex gap-2">
                      <div className="relative flex-1">
                        <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-white/40" />
                        <input type="text" placeholder="ยอดเงิน..." value={searchBank} onChange={e => setSearchBank(e.target.value)} className="w-full bg-white/5 border border-white/10 rounded-xl pl-8 pr-8 py-2 text-[10px] outline-none" />
                        {searchBank && <button onClick={() => setSearchBank('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-white/40 hover:text-white"><X size={14}/></button>}
                      </div>
                      <div className="flex bg-white/5 rounded-xl p-1 items-center border border-white/10">
                        <Calendar size={12} className="text-white/30" />
                        <input type="date" value={bankStartDate} onChange={e => setBankStartDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none opacity-60" />
                        <span className="text-white/10">-</span>
                        <input type="date" value={bankEndDate} onChange={e => setBankEndDate(e.target.value)} className="bg-transparent text-[9px] font-bold p-1 outline-none opacity-60" />
                        {(bankStartDate || bankEndDate) && <button onClick={()=>{setBankStartDate('');setBankEndDate('');}} className="p-1 text-white"><X size={12}/></button>}
                      </div>
                    </div>
                  </div>
                  <div className="flex-1 overflow-y-auto p-4 space-y-2 bg-slate-50/30">
                    {filteredBank.map((item) => (
                      <div key={item.id} onClick={() => toggleSelection(item, 'bank')} className={`p-3 rounded-2xl border-2 transition-all cursor-pointer ${selectedBank.some(i => i.id === item.id) ? 'border-slate-800 bg-slate-100 shadow-md' : 'border-white bg-white shadow-sm'}`}>
                        <div className="flex justify-between items-center">
                          <div className="flex flex-col"><span className="font-bold text-slate-700 text-xs line-clamp-1 leading-snug">{item.docNo}</span><span className="text-[9px] text-slate-400 font-bold uppercase mt-1 tracking-widest">{item.date}</span></div>
                          <span className={`text-xl font-black tabular-nums ${item.amount < 0 ? 'text-red-500' : 'text-slate-900'}`}>{item.amount.toLocaleString(undefined, {minimumFractionDigits: 2})}</span>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* Bottom Summary & Matching */}
              <div className="bg-white p-8 rounded-[3.5rem] shadow-xl flex flex-col md:flex-row justify-around items-center border border-slate-100 gap-6">
                <div className="text-center"><div className="text-slate-400 text-[10px] font-black mb-1 uppercase tracking-widest">รวมบัญชี</div><div className="text-5xl font-black text-blue-600 tracking-tighter tabular-nums">{internalSum.toLocaleString(undefined, {minimumFractionDigits: 2})}</div></div>
                <div className="flex flex-col items-center bg-slate-50 px-16 py-6 rounded-[2.5rem] border shadow-inner min-w-[380px]">
                  <div className="text-slate-400 text-[10px] font-black uppercase mb-1 tracking-widest">ผลต่างรวม</div>
                  <div className={`text-6xl font-black tabular-nums tracking-tighter ${diff < 0.01 ? 'text-emerald-500' : 'text-red-500'}`}>{diff.toLocaleString(undefined, {minimumFractionDigits: 2})}</div>
                  {diff < 0.01 && internalSum !== 0 && (
                    <button onClick={confirmMatch} className="mt-5 bg-blue-600 text-white px-12 py-3.5 rounded-full font-black text-xs hover:bg-blue-700 transition-all flex items-center gap-2 shadow-2xl animate-bounce tracking-widest uppercase">
                      ยืนยันจับคู่ <CheckCircle2 size={16}/>
                    </button>
                  )}
                </div>
                <div className="text-center"><div className="text-slate-400 text-[10px] font-black mb-1 uppercase tracking-widest">รวมธนาคาร</div><div className="text-5xl font-black text-slate-900 tracking-tighter tabular-nums">{bankSum.toLocaleString(undefined, {minimumFractionDigits: 2})}</div></div>
              </div>
            </div>
          ) : (
            /* Tab: Confirmed List */
            <div className="flex flex-col h-full gap-6">
              <div className="bg-white rounded-[3rem] shadow-sm border border-slate-200 overflow-hidden flex-1 flex flex-col min-h-[550px]">
                <div className="p-6 bg-slate-50 border-b grid grid-cols-4 font-black text-[15px] text-slate-400 uppercase tracking-widest">
                  <span>รายการบัญชี</span><span className="text-center">ยอดเงิน</span><span className="pl-8">รายการธนาคาร</span><span className="text-right">ยอดเงิน</span>
                </div>
                <div className="p-8 space-y-6 overflow-y-auto flex-1 bg-slate-50/20">
                  {confirmedMatches.length === 0 ? (
                    <div className="flex flex-col items-center justify-center h-full text-slate-300 gap-4 opacity-40">
                      <ArrowRightLeft size={64} />
                      <p className="font-bold uppercase tracking-widest">No confirmed matches yet</p>
                    </div>
                  ) : (
                    confirmedMatches.map((m) => (
                      <div key={m.id} className="group bg-white border border-slate-100 rounded-[2rem] p-8 grid grid-cols-4 items-center shadow-sm relative hover:border-blue-300 transition-all">
                        <div className="space-y-3">{m.internals.map(i => <div key={i.id} className="flex flex-col"><span className="text-[10px] text-slate-400 font-black">{i.date}</span><span className="text-sm font-bold text-blue-700">{i.docNo}</span></div>)}</div>
                        <div className="text-center font-black text-slate-800 text-2xl border-r border-slate-50">{m.totalAmount.toLocaleString(undefined, {minimumFractionDigits: 2})}</div>
                        <div className="pl-8 space-y-3">{m.banks.map(b => <div key={b.id} className="flex flex-col"><span className="text-[10px] text-slate-400 font-black">{b.date}</span><span className="text-sm font-bold text-slate-800 line-clamp-1">{b.docNo}</span></div>)}</div>
                        <div className="text-right font-black text-slate-800 text-2xl">{m.totalAmount.toLocaleString(undefined, {minimumFractionDigits: 2})}</div>
                        <button onClick={() => { 
                          setConfirmedMatches(prev => prev.filter(x => x.id !== m.id)); 
                          setInternalRecords(prev => [...prev, ...m.internals]); 
                          setBankStatement(prev => [...prev, ...m.banks]); 
                        }} className="absolute -right-3 -top-3 bg-white text-rose-500 border-2 border-rose-50 rounded-full p-2.5 opacity-0 group-hover:opacity-100 transition-all shadow-xl hover:bg-rose-500 hover:text-white">
                          <Trash2 size={18}/>
                        </button>
                      </div>
                    ))
                  )}
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