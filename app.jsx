import React, { useState, useCallback, useEffect } from 'react';

/**
 * 租金催收執行方式與成效 - 專業數據處理版 (V3.14)
 * 修正重點：
 * 1. Excel 零值處理：若違約金或管理費為 0，輸出時顯示為 "-" (字串型態)，SUM 公式會自動忽略文字，不影響計算。
 * 2. 靈活調整性：若使用者在 Excel 中於 "-" 儲存格手動輸入數字，SUM 公式會自動將其納入加總。
 * 3. Excel 全局置中：所有資料欄位設定為水平置中、垂直置中。
 * 4. 動態欄寬調整：自動計算資料內容長度，確保地址與長姓名不被遮蓋。
 * 5. 維持 V3.13 所有特性：月份補零 (113/01)、貨幣格式 ($#,##0)、深橘/淺橘背景色、所有 UI 正名。
 */

const App = () => {
  const [xlsxLoaded, setXlsxLoaded] = useState(false);
  const [fileAData, setFileAData] = useState(null);
  const [fileBData, setFileBData] = useState(null);
  const [cutoffDate, setCutoffDate] = useState(new Date().toISOString().split('T')[0]);
  const [processedResults, setProcessedResults] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);

  // 載入支援樣式的 Excel 函式庫 (xlsx-js-style)
  useEffect(() => {
    const script = document.createElement('script');
    script.src = "https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js";
    script.async = true;
    script.onload = () => setXlsxLoaded(true);
    document.head.appendChild(script);
    
    // 插入全域樣式：解決日期選擇器圖示在深色主題看不見的問題
    const style = document.createElement('style');
    style.innerHTML = `
      input[type="date"]::-webkit-calendar-picker-indicator {
        filter: invert(1);
        cursor: pointer;
        opacity: 0.8;
      }
      input[type="date"]::-webkit-calendar-picker-indicator:hover {
        opacity: 1;
      }
    `;
    document.head.appendChild(style);

    return () => {
      if (document.head.contains(script)) document.head.removeChild(script);
      if (document.head.contains(style)) document.head.removeChild(style);
    };
  }, []);

  // 西元轉民國日期 (YYYY-MM-DD -> 114/12/31)
  const toMinguoDate = (dateStr) => {
    if (!dateStr) return "";
    const parts = dateStr.split('-');
    if (parts.length !== 3) return dateStr;
    const year = parseInt(parts[0]) - 1911;
    return `${year}/${parts[1]}/${parts[2]}`;
  };

  /**
   * 清洗月份標籤並自動補零 (113/01)
   */
  const formatMonthLabel = (label) => {
    if (!label) return null;
    const str = String(label).trim();
    if (!/\d+/.test(str)) return null;
    let cleaned = str.replace(/年/g, '/').replace(/月/g, '').replace(/\s/g, '');
    const parts = cleaned.split('/');
    if (parts.length === 2) {
      const year = parts[0];
      const month = parts[1].padStart(2, '0');
      return `${year}/${month}`;
    }
    return cleaned;
  };

  const handleFileA = (e) => {
    if (!xlsxLoaded) return;
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = window.XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const data = window.XLSX.utils.sheet_to_json(wb.Sheets[wsname], { header: 1 });
      setFileAData(data); 
    };
    reader.readAsBinaryString(file);
  };

  const handleFileB = (e) => {
    if (!xlsxLoaded) return;
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = window.XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const data = window.XLSX.utils.sheet_to_json(wb.Sheets[wsname], { header: 1 });
      setFileBData(data);
    };
    reader.readAsBinaryString(file);
  };

  const processData = useCallback(() => {
    if (!fileAData || !fileBData || !xlsxLoaded) return;
    setIsProcessing(true);

    try {
      const cases = {};
      let headerRowIndex = -1;
      let headerRowA = [];

      // 智能偵測標題列
      for (let i = 0; i < Math.min(fileAData.length, 20); i++) {
        const row = fileAData[i] || [];
        const rowStr = row.map(cell => String(cell || "").trim()).join("|");
        if (rowStr.includes("逾期天數") || rowStr.includes("合計")) {
          headerRowIndex = i;
          headerRowA = row;
          break;
        }
      }

      if (headerRowIndex === -1) throw new Error("在 A 表中找不到有效的標題列。");

      const dataRowsA = fileAData.slice(headerRowIndex + 1);
      const overdueIdx = headerRowA.findIndex(h => h && String(h).includes('逾期天數'));
      const totalIdx = headerRowA.findIndex(h => h && String(h).includes('合計'));

      if (overdueIdx === -1 || totalIdx === -1) throw new Error("無法定位資料區間邊界。");

      const startMonthIdx = overdueIdx + 1;
      const endMonthIdx = totalIdx;

      dataRowsA.forEach(row => {
        const caseId = row[1]; // B 欄案號
        if (!caseId) return;

        if (!cases[caseId]) {
          cases[caseId] = {
            id: caseId,
            name: row[2] || "未知",
            address: row[3] || "無地址資訊",
            monthlyRent: row[5] || 0,
            arrears: [],
            penaltyTotal: 0,
            startDate: '',
            endDate: ''
          };
        }

        for (let i = startMonthIdx; i < endMonthIdx; i++) {
          const amt = parseFloat(row[i]);
          if (!isNaN(amt) && amt > 0) {
            const cleanedLabel = formatMonthLabel(headerRowA[i]);
            if (cleanedLabel) {
              cases[caseId].arrears.push({ month: cleanedLabel, amount: amt });
            }
          }
        }
      });

      const headerRowB = fileBData[0] || [];
      let penaltyColIdx = headerRowB.findIndex(h => h && String(h).includes('違約金額'));
      if (penaltyColIdx === -1) penaltyColIdx = 13;

      const startColIdxB = headerRowB.findIndex(h => h && String(h).includes('開始'));
      const endColIdxB = headerRowB.findIndex(h => h && String(h).includes('結束'));

      fileBData.slice(1).forEach(row => {
        const caseId = row[2]; // B 表案號
        if (cases[caseId]) {
          const pAmt = parseFloat(row[penaltyColIdx]);
          if (!isNaN(pAmt)) cases[caseId].penaltyTotal += pAmt;
          if (startColIdxB !== -1) cases[caseId].startDate = row[startColIdxB] || cases[caseId].startDate;
          if (endColIdxB !== -1) cases[caseId].endDate = row[endColIdxB] || cases[caseId].endDate;
        }
      });

      // 過濾：僅保留有租金欠費紀錄的案號
      const filteredResults = Object.values(cases).filter(c => c.arrears.length > 0);
      setProcessedResults(filteredResults);
    } catch (err) {
      console.error(err);
      alert(err.message || "處理失敗。");
    } finally {
      setIsProcessing(false);
    }
  }, [fileAData, fileBData, xlsxLoaded]);

  // 下載 Styled Excel (支援 0 值轉 "-" 與全域置中)
  const downloadExcel = () => {
    if (!window.XLSX) return;
    const wb = window.XLSX.utils.book_new();
    const mCutoff = toMinguoDate(cutoffDate);

    // 樣式定義
    const baseAlign = { horizontal: 'center', vertical: 'center' };
    const borderAll = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
    
    const styleDeepOrange = { 
      fill: { fgColor: { rgb: "E67E22" } }, 
      font: { color: { rgb: "FFFFFF" }, bold: true }, 
      border: borderAll,
      alignment: baseAlign
    };
    
    const styleDeepOrangeCurrency = { 
      ...styleDeepOrange,
      numFmt: "$#,##0" 
    };

    const styleLightOrangeHeader = { 
      fill: { fgColor: { rgb: "F39C12" } }, 
      font: { bold: true }, 
      border: borderAll,
      alignment: { ...baseAlign, wrapText: true }
    };

    const styleWhiteValue = { 
      fill: { fgColor: { rgb: "FFFFFF" } }, 
      border: borderAll, 
      alignment: { ...baseAlign, wrapText: true } 
    };
    
    const styleWhiteCurrency = { 
      ...styleWhiteValue,
      numFmt: "$#,##0" 
    };

    const sheetData = [];

    processedResults.forEach((c) => {
      // 承租人
      sheetData.push([
        { v: "承租人", s: styleDeepOrange },
        { v: c.name, s: styleDeepOrange }
      ]);
      // 物件地址
      sheetData.push([
        { v: "物件地址", s: styleDeepOrange },
        { v: c.address, s: styleDeepOrange }
      ]);
      // 每月租金
      sheetData.push([
        { v: "每月租金", s: styleDeepOrangeCurrency },
        { v: Number(c.monthlyRent), t: 'n', s: styleDeepOrangeCurrency }
      ]);
      // 原租約起訖日
      sheetData.push([
        { v: "原租約起訖日", s: styleLightOrangeHeader },
        { v: `${c.startDate} ~ ${c.endDate}`, s: styleWhiteValue }
      ]);

      const sumIndices = [];
      c.arrears.forEach(a => {
        sheetData.push([
          { v: `未繳租金(${a.month})`, s: styleLightOrangeHeader },
          { v: Number(a.amount), t: 'n', s: styleWhiteCurrency }
        ]);
        sumIndices.push(sheetData.length);
      });

      // 違約金備註 (V3.14 修正：0值顯示為 "-")
      const penaltyVal = Number(c.penaltyTotal);
      sheetData.push([
        { v: "已知未繳違約利息、罰款\n(尚有其他違約金待統計)", s: styleLightOrangeHeader },
        penaltyVal === 0 
          ? { v: "-", s: styleWhiteValue } 
          : { v: penaltyVal, t: 'n', s: styleWhiteCurrency }
      ]);
      sumIndices.push(sheetData.length);

      // 管理費 (V3.14 修正：預設0值顯示為 "-")
      sheetData.push([
        { v: `未繳管理費(截至${mCutoff})`, s: styleLightOrangeHeader },
        { v: "-", s: styleWhiteValue }
      ]);
      sumIndices.push(sheetData.length);

      // 目前暫列欠款金額
      const startIdx = sumIndices[0];
      const endIdx = sumIndices[sumIndices.length - 1];
      sheetData.push([
        { v: "目前暫列欠款金額", s: styleLightOrangeHeader },
        { f: `SUM(B${startIdx}:B${endIdx})`, s: { ...styleWhiteCurrency, font: { bold: true, color: { rgb: "FF0000" } } } }
      ]);

      sheetData.push([], []); 
    });

    const ws = window.XLSX.utils.aoa_to_sheet(sheetData);
    
    // 動態計算欄寬
    const colWidths = [15, 20];
    sheetData.forEach(row => {
      row.forEach((cell, i) => {
        if (!cell || cell.v === undefined) return;
        const str = String(cell.v);
        const lines = str.split('\n');
        let maxLineLen = 0;
        lines.forEach(line => {
          let currentLen = 0;
          for (let char of line) {
            currentLen += (char.charCodeAt(0) > 255) ? 2.2 : 1.1;
          }
          if (currentLen > maxLineLen) maxLineLen = currentLen;
        });
        if (maxLineLen > colWidths[i]) colWidths[i] = maxLineLen;
      });
    });

    ws['!cols'] = colWidths.map(w => ({ wch: w + 2 }));

    window.XLSX.utils.book_append_sheet(wb, ws, "租務統整表");
    window.XLSX.writeFile(wb, `租金催收統整_${mCutoff.replace(/\//g, '')}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-[#050505] text-zinc-100 p-6 md:p-12 font-sans selection:bg-yellow-500/30">
      <div className="max-w-4xl mx-auto">
        <header className="mb-12 text-center">
          <h1 className="text-4xl font-black mb-4 tracking-tight">租金催收執行方式與成效</h1>
        </header>

        <div className="bg-zinc-900/40 rounded-3xl border border-white/5 p-8 mb-12 shadow-2xl backdrop-blur-sm">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-8 mb-10">
            <div className={`relative border-2 border-dashed rounded-2xl p-6 transition-all duration-500 ${fileAData ? 'border-emerald-500/50 bg-emerald-500/5' : 'border-zinc-800 hover:border-yellow-600/50'}`}>
              <label className="cursor-pointer block">
                <span className="text-[10px] font-black text-zinc-600 block mb-3 uppercase tracking-widest">A 表：捷運宅催收統計表</span>
                <input type="file" onChange={handleFileA} className="hidden" accept=".xlsx,.xls" />
                <div className="text-xl font-bold">{fileAData ? "🟢 已載入捷運宅數據" : "📂 選取催收統計檔案"}</div>
              </label>
            </div>
            <div className={`relative border-2 border-dashed rounded-2xl p-6 transition-all duration-500 ${fileBData ? 'border-emerald-500/50 bg-emerald-500/5' : 'border-zinc-800 hover:border-yellow-600/50'}`}>
              <label className="cursor-pointer block">
                <span className="text-[10px] font-black text-zinc-600 block mb-3 uppercase tracking-widest">B 表：承租戶延遲繳納罰款未收明細</span>
                <input type="file" onChange={handleFileB} className="hidden" accept=".xlsx,.xls" />
                <div className="text-xl font-bold">{fileBData ? "🟢 已載入罰款明細數據" : "📂 選取罰款未收檔案"}</div>
              </label>
            </div>
          </div>

          <div className="flex flex-col md:flex-row items-center gap-6">
            <div className="flex-1 w-full bg-black/40 p-5 rounded-2xl border border-white/5 flex items-center gap-4">
              <span className="text-2xl">📅</span>
              <div className="flex-1">
                <label className="text-zinc-500 text-[10px] font-black uppercase tracking-widest block mb-1">管理費欠費統計截止日</label>
                <input 
                  type="date" 
                  value={cutoffDate} 
                  onChange={(e) => setCutoffDate(e.target.value)}
                  className="bg-transparent text-white font-bold w-full outline-none focus:text-yellow-500 transition-colors cursor-pointer"
                />
              </div>
            </div>
            <button
              onClick={processData}
              disabled={!fileAData || !fileBData || isProcessing}
              className={`px-12 py-5 rounded-2xl font-black text-lg transition-all active:scale-95 whitespace-nowrap ${
                !fileAData || !fileBData ? "bg-zinc-800 text-zinc-600 cursor-not-allowed" : "bg-yellow-600 hover:bg-yellow-500 text-black shadow-xl shadow-yellow-900/20"
              }`}
            >
              {isProcessing ? "處理中..." : "開始統整資料"}
            </button>
          </div>
        </div>

        {processedResults.length > 0 && (
          <div className="mb-12 flex justify-center">
            <button onClick={downloadExcel} className="px-16 py-4 rounded-full font-black text-xl bg-emerald-600 hover:bg-emerald-500 text-white shadow-2xl shadow-emerald-900/30 transition-all hover:-translate-y-1">
              📥 下載統整 EXCEL 報表
            </button>
          </div>
        )}

        <div className="space-y-16 pb-40">
          {processedResults.map((item, idx) => (
            <div key={idx} className="bg-[#1C1C1C] rounded-xl overflow-hidden border border-white/10 max-w-2xl mx-auto transition-all hover:border-yellow-600/30">
              <div className="bg-zinc-800/40 px-6 py-2 flex justify-between items-center text-[10px] font-black tracking-widest text-zinc-500 uppercase">
                <span>Case Preview {idx + 1}</span>
                <span>Case ID: {item.id}</span>
              </div>
              <table className="w-full text-center border-collapse">
                <tbody className="text-xl">
                  <tr>
                    <td className="w-1/3 bg-[#84592D] text-white py-4 font-black border border-white/10">承租人</td>
                    <td className="w-2/3 bg-[#1C1C1C] text-white py-4 font-black border border-white/10">{item.name}</td>
                  </tr>
                  <tr>
                    <td className="bg-[#84592D] text-white py-3 border border-white/10 font-bold">物件地址</td>
                    <td className="bg-[#1C1C1C] text-zinc-400 py-3 border border-white/10 italic px-4 leading-snug">{item.address}</td>
                  </tr>
                  <tr>
                    <td className="bg-[#84592D] text-white py-3 border border-white/10 font-bold">每月租金</td>
                    <td className="bg-[#1C1C1C] text-white py-3 border border-white/10 font-mono">${Number(item.monthlyRent || 0).toLocaleString()}</td>
                  </tr>
                  <tr>
                    <td className="bg-[#84592D] text-white py-3 border border-white/10 font-bold">原租約起訖日</td>
                    <td className="bg-[#1C1C1C] text-zinc-500 py-3 border border-white/10 font-mono tracking-tighter">{item.startDate} ~ {item.endDate}</td>
                  </tr>
                  {item.arrears.map((arr, aidx) => (
                    <tr key={aidx}>
                      <td className="bg-[#84592D] text-yellow-500 py-3 border border-white/10 font-black italic">未繳租金({arr.month})</td>
                      <td className="bg-[#1C1C1C] text-[#FF5555] py-3 border border-white/10 font-black font-mono">${Number(arr.amount).toLocaleString()}</td>
                    </tr>
                  ))}
                  <tr>
                    <td className="bg-[#84592D] text-yellow-500 py-4 border border-white/10 font-black leading-tight px-4 text-center">
                      <div>已知未繳違約利息、罰款</div>
                      <div className="text-sm font-normal opacity-50 text-white italic mt-1">(尚有其他違約金待統計)</div>
                    </td>
                    <td className="bg-[#1C1C1C] text-[#FF5555] py-4 border border-white/10 font-black font-mono">
                      {item.penaltyTotal === 0 ? "-" : `$${item.penaltyTotal.toLocaleString()}`}
                    </td>
                  </tr>
                  <tr>
                    <td className="bg-[#84592D] text-yellow-500 py-3 border border-white/10 font-black px-4 text-center">
                      未繳管理費(截至{toMinguoDate(cutoffDate)})
                    </td>
                    <td className="bg-[#1C1C1C] text-zinc-700 py-3 border border-white/10 italic text-sm font-black uppercase tracking-wider">— 下載報表後手動填入此欄 —</td>
                  </tr>
                  <tr>
                    <td className="bg-[#84592D] text-white py-6 border border-white/10 font-black uppercase tracking-tighter">目前暫列欠款金額</td>
                    <td className="bg-[#1C1C1C] text-[#FF5555] py-6 border border-white/10 font-black font-mono shadow-[inset_0_0_80px_rgba(255,85,85,0.08)]">
                      ${(item.arrears.reduce((acc, cur) => acc + cur.amount, 0) + item.penaltyTotal).toLocaleString()}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          ))}
          {processedResults.length === 0 && !isProcessing && (
            <div className="py-40 text-center select-none">
              <div className="text-[120px] mb-8 opacity-[0.02] font-black leading-none">NO DATA</div>
              <p className="font-black tracking-[1em] text-zinc-800 text-[10px] uppercase italic">System idle • Standardized Widths Active</p>
            </div>
          )}
        </div>
      </div>
      <footer className="mt-20 border-t border-white/5 pt-12 pb-24 text-center">
        <div className="text-zinc-800 text-[10px] tracking-[1.5em] font-black uppercase opacity-40">Consolidation Engine v3.14 • Zero as Dash</div>
      </footer>
    </div>
  );
};

export default App;