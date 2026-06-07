'use client';

import { useEffect, useState, useRef } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// Khai báo biến toàn cục
declare const luckysheet: any;
declare const LuckyExcel: any;

// ==========================================
// CÁC HÀM HELPER HỖ TRỢ ĐỌC CHỮ EXCEL
// ==========================================
const getColLetter = (colIndex: number) => {
  let letter = '';
  while (colIndex >= 0) {
    letter = String.fromCharCode((colIndex % 26) + 65) + letter;
    colIndex = Math.floor(colIndex / 26) - 1;
  }
  return letter;
};

const extractTextWithHyperlinkFilter = (cell: any, mode: 'NORMAL' | 'RPAC' = 'NORMAL'): string => {
  if (!cell) return "";
  let text = "";
  if (cell.ct && cell.ct.t === 'inlineStr' && cell.ct.s) {
    text = cell.ct.s.map((s: any) => s.v).join("");
  } else if (cell.m !== undefined && cell.m !== null) {
    text = cell.m.toString();
  } else if (cell.v !== undefined && cell.v !== null) {
    text = cell.v.toString();
  }
  
  if (mode === 'NORMAL') {
    return text;
  }
  
  // R-PAC Mode: Xử lý nếu là công thức HYPERLINK ẩn và lọc URL
  let formula = cell.f || "";
  if (formula.toUpperCase().includes("HYPERLINK")) {
    const match = formula.match(/HYPERLINK\s*\(\s*["'].*?["']\s*[,;]\s*["'](.*?)["']\s*\)/i);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  if (text.toUpperCase().includes("HYPERLINK")) {
    const match = text.match(/HYPERLINK\s*\(\s*["'].*?["']\s*[,;]\s*["'](.*?)["']\s*\)/i);
    if (match && match[1]) {
      text = match[1];
    }
  }

  // Loại bỏ URL/Domain để so khớp thuần túy nội dung hiển thị cho công ty R-pac
  text = text.replace(/https?:\/\/[^\s]+/gi, '');
  text = text.replace(/www\.[^\s]+/gi, '');
  return text;
};

const getRangeText = (row: any, c1: number, c2: number, mode: 'NORMAL' | 'RPAC' = 'NORMAL') => {
  if (!row) return "";
  let textArray = [];
  for (let i = c1; i <= c2; i++) {
    let cell = row[i];
    if (cell) {
      let text = extractTextWithHyperlinkFilter(cell, mode);
      textArray.push(text.trim().replace(/\r?\n/g, ' '));
    } else {
      textArray.push("");
    }
  }
  return textArray.join(' | ');
};

const getCleanedRangeText = (row: any, c1: number, c2: number, mode: 'NORMAL' | 'RPAC' = 'NORMAL') => {
  if (!row) return "";
  let textArray = [];
  for (let i = c1; i <= c2; i++) {
    let cell = row[i];
    let text = extractTextWithHyperlinkFilter(cell, mode);
    // Viết thường, loại bỏ toàn bộ khoảng trắng và ký tự đặc biệt để so khớp chính xác nhất
    let cleaned = text.toLowerCase().replace(/\s+/g, '').replace(/[0-9.,:\/\\#\-]/g, '');
    textArray.push(cleaned);
  }
  return textArray.join('|');
};

const shiftWorksheetMerges = (ws: any, rowStart: number, numRows: number) => {
  try {
    const merges = ws.model.merges || [];
    const newMerges: string[] = [];

    const decodeCell = (addr: string) => {
      const match = addr.match(/^([A-Z]+)([0-9]+)$/);
      if (!match) return { row: 0, col: 0 };
      const colStr = match[1];
      const row = parseInt(match[2], 10);
      let col = 0;
      for (let i = 0; i < colStr.length; i++) {
        col = col * 26 + (colStr.charCodeAt(i) - 64);
      }
      return { row, col };
    };

    const encodeCell = (row: number, col: number) => {
      let colStr = "";
      let temp = col;
      while (temp > 0) {
        let r = (temp - 1) % 26;
        colStr = String.fromCharCode(65 + r) + colStr;
        temp = Math.floor((temp - 1) / 26);
      }
      return `${colStr}${row}`;
    };

    for (let mergeStr of merges) {
      const parts = mergeStr.split(':');
      if (parts.length !== 2) continue;
      
      const startCell = decodeCell(parts[0]);
      const endCell = decodeCell(parts[1]);
      
      let sRow = startCell.row;
      let eRow = endCell.row;
      
      // 1. Nằm hoàn toàn trong phạm vi bị xóa -> Bỏ qua
      if (sRow >= rowStart && eRow < rowStart + numRows) {
        continue;
      }
      
      // 2. Nằm hoàn toàn phía dưới phạm vi bị xóa -> Dịch chuyển lên
      if (sRow >= rowStart + numRows) {
        sRow -= numRows;
        eRow -= numRows;
        newMerges.push(`${encodeCell(sRow, startCell.col)}:${encodeCell(eRow, endCell.col)}`);
      } 
      // 3. Nằm hoàn toàn phía trên phạm vi bị xóa -> Giữ nguyên
      else if (eRow < rowStart) {
        newMerges.push(mergeStr);
      }
      // 4. Bị cắt giao -> Thu nhỏ lại
      else {
        if (sRow < rowStart && eRow >= rowStart + numRows) {
          eRow -= numRows;
          newMerges.push(`${encodeCell(sRow, startCell.col)}:${encodeCell(eRow, endCell.col)}`);
        } else if (sRow < rowStart) {
          eRow = rowStart - 1;
          if (eRow >= sRow) {
            newMerges.push(`${encodeCell(sRow, startCell.col)}:${encodeCell(eRow, endCell.col)}`);
          }
        }
      }
    }

    // Xóa tất cả và merge lại bằng tọa độ mới đã căn chỉnh
    ws.model.merges = [];
    ws._merges = {};
    for (let newMerge of newMerges) {
      try {
        ws.mergeCells(newMerge);
      } catch (e) {
        console.error("Lỗi khi merge lại ô trong ExcelJS:", newMerge, e);
      }
    }
  } catch (err) {
    console.error("Lỗi khi đồng bộ ô gộp (merges) trong ExcelJS:", err);
  }
};

const matchSingleRow = (row: any, patternRowText: string, c1: number, c2: number, mode: 'NORMAL' | 'RPAC' = 'NORMAL') => {
  if (!row) return false;
  let pCells = patternRowText.split('|').map(s => s.trim()).filter(s => s !== "");
  if (pCells.length === 0) return true; // Dòng mẫu trống thì luôn khớp
  
  let rowTextCombined = "";
  for (let c = c1; c <= c2; c++) {
    let cell = row[c];
    let text = extractTextWithHyperlinkFilter(cell, mode);
    let cleaned = text.toLowerCase().replace(/\s+/g, '').replace(/[0-9.,:\/\\#\-]/g, '');
    rowTextCombined += cleaned;
  }
  
  let matchCount = 0;
  for (let pCell of pCells) {
    if (rowTextCombined.includes(pCell)) {
      matchCount++;
    }
  }
  
  let requiredMatches = Math.max(1, Math.ceil(pCells.length * 0.8));
  return matchCount >= requiredMatches;
};

const matchHeader = (sheetData: any[], r: number, pattern: string[], c1: number, c2: number, mode: 'NORMAL' | 'RPAC' = 'NORMAL') => {
  if (r >= sheetData.length) return false;
  
  // Chỉ cần kiểm tra dòng đầu tiên r có khớp với pattern[0] không.
  // Nhờ cơ chế khớp 80% số ô có dữ liệu của mẫu, việc này đã đủ chính xác để tránh nhận diện nhầm các dòng chữ ký.
  return matchSingleRow(sheetData[r], pattern[0], c1, c2, mode);
};

const matchFooter = (sheetData: any[], r: number, fuzzyPattern: string[], c1: number, c2: number, mode: 'NORMAL' | 'RPAC' = 'NORMAL') => {
  if (r + fuzzyPattern.length > sheetData.length) return false;
  for (let i = 0; i < fuzzyPattern.length; i++) {
    let text = getCleanedRangeText(sheetData[r + i], c1, c2, mode);
    if (text !== fuzzyPattern[i]) {
      return false;
    }
  }
  return true;
};

export default function Home() {
  const [loadingMsg, setLoadingMsg] = useState<string>('');
  const [showPanel, setShowPanel] = useState<boolean>(false);
  const [statusText, setStatusText] = useState<string>('(Hãy chọn file Excel để xem)');
  const [hasFile, setHasFile] = useState<boolean>(false);
  const [workbookObj, setWorkbookObj] = useState<any>(null);

  // State quản lý logic cắt gộp
  const [currentMode, setCurrentMode] = useState<'HEADER' | 'FOOTER'>('HEADER');
  const [companyMode, setCompanyMode] = useState<'NORMAL' | 'RPAC'>('NORMAL');
  const [headerData, setHeaderData] = useState<any>(null);
  const [footerData, setFooterData] = useState<any>(null);
  const [previewHTML, setPreviewHTML] = useState<string>('');
  
  // State quản lý Kéo thả bảng điều khiển
  const [panelPos, setPanelPos] = useState({ x: 0, y: 0 });
  const [selectionStats, setSelectionStats] = useState<{ average: number, count: number, sum: number } | null>(null);

  // State quản lý Pivot Table cho Polytex
  const [showPivotModal, setShowPivotModal] = useState<boolean>(false);
  const [pivotModalPos, setPivotModalPos] = useState({ x: 0, y: 0 });
  const [pivotColumns, setPivotColumns] = useState<{ index: number; name: string; isNumeric: boolean }[]>([]);
  const [selectedPivotGroupCols, setSelectedPivotGroupCols] = useState<number[]>([]);
  const [selectedPivotAggCols, setSelectedPivotAggCols] = useState<number[]>([]);
  const [splitPolytexCode, setSplitPolytexCode] = useState<boolean>(true);
  const [pivotResult, setPivotResult] = useState<{ headers: string[]; rows: any[] } | null>(null);
  const [pivotActiveRange, setPivotActiveRange] = useState<any>(null);

  // Refs để Luckysheet (nằm ngoài chu kỳ render của React) luôn đọc được state mới nhất
  const currentModeRef = useRef(currentMode);
  const companyModeRef = useRef(companyMode);
  const showPanelRef = useRef(showPanel);

  useEffect(() => { currentModeRef.current = currentMode; }, [currentMode]);
  useEffect(() => { companyModeRef.current = companyMode; }, [companyMode]);
  useEffect(() => { showPanelRef.current = showPanel; }, [showPanel]);

  // ==========================================
  // HÀM XỬ LÝ KHI QUÉT CHỌN Ô (RANGE SELECT)
  // ==========================================
  const handleRangeSelect = (sheet: any, range: any) => {
    const isPanelOpen = showPanelRef.current;

    // Tự động lưu vùng bôi chọn làm Header mẫu bất kỳ lúc nào người dùng quét chọn trên bảng tính
    if (isPanelOpen) {
      if (range && range.length > 0) {
        let r1 = range[0].row[0]; let r2 = range[0].row[1];
        let c1 = range[0].column[0]; let c2 = range[0].column[1];
        
        // @ts-ignore
        let sheetData = luckysheet.getSheetData();
        let rangeName = `${getColLetter(c1)}${r1 + 1}:${getColLetter(c2)}${r2 + 1}`;
        if (r1 === r2 && c1 === c2) rangeName = `${getColLetter(c1)}${r1 + 1}`;

        let pattern = [];
        let html = "";
        for (let r = r1; r <= r2; r++) {
          let text = getCleanedRangeText(sheetData[r], c1, c2, companyModeRef.current);
          pattern.push(text);
          let display = getRangeText(sheetData[r], c1, c2, companyModeRef.current);
          let bgColor = (r % 2 === 0) ? '#f4f7f6' : '#eaf1ee';
          html += `<div style="background:${bgColor}; padding: 3px 0; border-bottom: 1px solid #ddd; color: black;"><b>D. ${r + 1}:</b> ${display}</div>`;
        }

        if (pattern.every(t => t.replace(/\|/g, '').trim() === '')) {
          setPreviewHTML(`<span style="color:red;">Vùng chọn trống, vui lòng quét chọn lại!</span>`);
          setHeaderData(null);
        } else {
          setHeaderData({ pattern, c1, c2, r1, r2, rangeName });
          const isHeader = currentModeRef.current === 'HEADER';
          const label = isHeader ? 'Header mẫu' : 'Footer mẫu';
          const labelColor = isHeader ? '#e53935' : '#2b579a';
          setPreviewHTML(`<span style="color:${labelColor}; font-weight:bold;">Đã nhận diện ${label} (${rangeName}):</span><br><div style="margin-top:5px; max-height:80px; overflow-y:auto; border:1px solid #ddd; padding:5px; background:#fff; border-radius:4px;">${html}</div>`);
        }
      }
    }

    // Tự động tính toán Sum, Count, Average cho các ô được chọn
    if (range && range.length > 0) {
      try {
        // @ts-ignore
        let sheetData = luckysheet.getSheetData();
        let sum = 0;
        let count = 0;
        let hasNumbers = false;

        for (let rng of range) {
          let r1 = rng.row[0];
          let r2 = rng.row[1];
          let c1 = rng.column[0];
          let c2 = rng.column[1];

          for (let r = r1; r <= r2; r++) {
            if (!sheetData[r]) continue;
            for (let c = c1; c <= c2; c++) {
              let cell = sheetData[r][c];
              if (cell && cell.v !== undefined && cell.v !== null) {
                // Loại bỏ định dạng dấu phẩy phân cách phần nghìn để parse số chuẩn xác
                let valStr = cell.v.toString().replace(/,/g, '').trim();
                let parsed = parseFloat(valStr);
                if (!isNaN(parsed)) {
                  sum += parsed;
                  count++;
                  hasNumbers = true;
                }
              }
            }
          }
        }

        if (hasNumbers && count > 0) {
          setSelectionStats({
            sum: parseFloat(sum.toFixed(2)),
            count: count,
            average: parseFloat((sum / count).toFixed(2))
          });
        } else {
          setSelectionStats(null);
        }
      } catch (err) {
        console.error("Lỗi khi tính toán Sum/Count/Average vùng chọn:", err);
        setSelectionStats(null);
      }
    } else {
      setSelectionStats(null);
    }
  };

  // ==========================================
  // HÀM XỬ LÝ PIVOT TABLE CHO POLYTEX
  // ==========================================
  const handlePivotPanelMouseDown = (e: React.MouseEvent<HTMLDivElement>) => {
    const startX = e.clientX - pivotModalPos.x;
    const startY = e.clientY - pivotModalPos.y;

    const handleMouseMove = (mouseMoveEvent: MouseEvent) => {
      setPivotModalPos({
        x: mouseMoveEvent.clientX - startX,
        y: mouseMoveEvent.clientY - startY
      });
    };

    const handleMouseUp = () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };

    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
  };

  const openPivotModal = () => {
    if (typeof luckysheet === 'undefined') return;
    
    // @ts-ignore
    let ranges = luckysheet.getRange();
    if (!ranges || ranges.length === 0) {
      alert("Vui lòng quét chọn (bôi đen) toàn bộ bảng dữ liệu cần tạo Pivot trước!");
      return;
    }
    
    let r1 = ranges[0].row[0];
    let r2 = ranges[0].row[1];
    let c1 = ranges[0].column[0];
    let c2 = ranges[0].column[1];
    
    if (r2 <= r1) {
      alert("Vui lòng quét chọn bảng dữ liệu bao gồm cả dòng tiêu đề ở trên và các dòng dữ liệu ở dưới!");
      return;
    }
    
    // @ts-ignore
    let sheetData = luckysheet.getSheetData();
    
    let cols: { index: number; name: string; isNumeric: boolean }[] = [];
    let groupDefault: number[] = [];
    let aggDefault: number[] = [];
    
    let hasHashCodes = false;
    
    for (let c = c1; c <= c2; c++) {
      let cell = sheetData[r1][c];
      let name = "";
      if (cell) {
        name = extractTextWithHyperlinkFilter(cell, 'NORMAL').trim();
      }
      if (!name) {
        name = `Cột ${getColLetter(c)}`;
      }
      
      // Kiểm tra xem cột có chứa dữ liệu số không
      let numericCount = 0;
      let nonNumericCount = 0;
      for (let r = r1 + 1; r <= r2; r++) {
        if (!sheetData[r]) continue;
        let cellData = sheetData[r][c];
        if (cellData) {
          let textVal = extractTextWithHyperlinkFilter(cellData, 'NORMAL').trim();
          let valStr = textVal.replace(/,/g, '');
          if (valStr !== "") {
            let parsed = parseFloat(valStr);
            if (!isNaN(parsed)) {
              numericCount++;
            } else {
              nonNumericCount++;
            }
          }
          
          if (textVal.match(/^#\d+/) || textVal.includes("\n#") || textVal.includes("\r#")) {
            hasHashCodes = true;
          }
        }
      }
      
      let isNumeric = numericCount > 0 && numericCount >= nonNumericCount;
      cols.push({ index: c, name, isNumeric });
      
      if (isNumeric) {
        aggDefault.push(c);
      } else {
        groupDefault.push(c);
      }
    }
    
    // Thêm cột ảo cho mã "#" nếu tìm thấy
    if (hasHashCodes) {
      cols.push({ index: -100, name: "Mã số (#)", isNumeric: false });
      groupDefault.push(-100);
    }
    
    setPivotColumns(cols);
    setSelectedPivotGroupCols(groupDefault);
    setSelectedPivotAggCols(aggDefault);
    setPivotActiveRange({ r1, r2, c1, c2 });
    setPivotResult(null);
    setPivotModalPos({ x: 0, y: 0 });
    setShowPivotModal(true);
  };

  const handleGeneratePivot = () => {
    if (!pivotActiveRange) return;
    
    // @ts-ignore
    let sheetData = luckysheet.getSheetData();
    const { r1, r2, c1, c2 } = pivotActiveRange;
    
    let groups: { [key: string]: { groupValues: any[], aggValues: { [colIndex: number]: number } } } = {};
    
    for (let r = r1 + 1; r <= r2; r++) {
      if (!sheetData[r]) continue;
      
      // Trích xuất mã hash (ví dụ: #1012081)
      let hashCode = "";
      for (let c = c1; c <= c2; c++) {
        let cell = sheetData[r][c];
        if (cell) {
          let text = extractTextWithHyperlinkFilter(cell, 'NORMAL');
          let match = text.match(/#\d+/);
          if (match) {
            hashCode = match[0];
            break;
          }
        }
      }
      
      let rowGroupValues: any[] = [];
      let isRowEmpty = true;
      
      selectedPivotGroupCols.forEach(colIndex => {
        if (colIndex === -100) {
          rowGroupValues.push(hashCode);
        } else {
          let cell = sheetData[r][colIndex];
          let val = "";
          if (cell) {
            val = extractTextWithHyperlinkFilter(cell, 'NORMAL');
            if (splitPolytexCode && hashCode) {
              val = val.replace(hashCode, "").trim();
              val = val.replace(/^\s*[\r\n]/gm, "").trim();
            }
            if (val !== "") {
              isRowEmpty = false;
            }
          }
          rowGroupValues.push(val);
        }
      });
      
      if (isRowEmpty && !hashCode) continue;
      
      let groupKey = JSON.stringify(rowGroupValues);
      
      if (!groups[groupKey]) {
        groups[groupKey] = {
          groupValues: rowGroupValues,
          aggValues: {}
        };
        selectedPivotAggCols.forEach(colIndex => {
          groups[groupKey].aggValues[colIndex] = 0;
        });
      }
      
      selectedPivotAggCols.forEach(colIndex => {
        let cell = sheetData[r][colIndex];
        if (cell && cell.v !== undefined && cell.v !== null) {
          let valStr = cell.v.toString().replace(/,/g, '').trim();
          let parsed = parseFloat(valStr);
          if (!isNaN(parsed)) {
            groups[groupKey].aggValues[colIndex] += parsed;
          }
        }
      });
    }
    
    let headers: string[] = [];
    selectedPivotGroupCols.forEach(colIndex => {
      if (colIndex === -100) {
        headers.push("Mã số (#)");
      } else {
        let headerCol = pivotColumns.find(col => col.index === colIndex);
        headers.push(headerCol ? headerCol.name : `Cột ${colIndex}`);
      }
    });
    selectedPivotAggCols.forEach(colIndex => {
      let headerCol = pivotColumns.find(col => col.index === colIndex);
      headers.push(headerCol ? `${headerCol.name} (Tổng)` : `Cột ${colIndex} (Tổng)`);
    });
    
    let rows = Object.values(groups).map(g => {
      let rowData = [...g.groupValues];
      selectedPivotAggCols.forEach(colIndex => {
        let val = g.aggValues[colIndex];
        rowData.push(parseFloat(val.toFixed(2)));
      });
      return rowData;
    });
    
    setPivotResult({ headers, rows });
  };

  const handleExportPivotToNewSheet = () => {
    if (!pivotResult) return;
    
    try {
      let luckyData: any[][] = [];
      
      let headerRow = pivotResult.headers.map(h => ({
        v: h,
        m: h,
        fc: "#ffffff",
        bg: "#107c41",
        bl: 1,
        ht: 1,
        vt: 1
      }));
      luckyData.push(headerRow);
      
      pivotResult.rows.forEach((row: any) => {
        let luckyRow = row.map((val: any) => ({
          v: val,
          m: val !== null && val !== undefined ? val.toString() : "",
          ht: 1,
          vt: 1
        }));
        luckyData.push(luckyRow);
      });
      
      let celldata: any[] = [];
      luckyData.forEach((row, r) => {
        row.forEach((cellVal, c) => {
          celldata.push({
            r: r,
            c: c,
            v: cellVal
          });
        });
      });

      // @ts-ignore
      luckysheet.setSheetAdd({
        sheetObject: {
          name: "Pivot Polytex",
          celldata: celldata
        }
      });
      
      alert("Đã tạo và chèn thêm Sheet 'Pivot Polytex' thành công!");
      setShowPivotModal(false);
    } catch (err: any) {
      alert("Lỗi khi chèn sheet mới vào Luckysheet: " + err.message);
    }
  };

  const handleExportPivotToExcel = async () => {
    if (!pivotResult) return;
    
    setLoadingMsg("Đang đóng gói file Pivot Excel...");
    
    setTimeout(async () => {
      try {
        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet("Pivot Polytex");
        
        const headerRow = ws.addRow(pivotResult.headers);
        headerRow.eachCell(cell => {
          cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF107C41" }
          };
          cell.alignment = { horizontal: "center", vertical: "middle" };
        });
        
        pivotResult.rows.forEach(row => {
          ws.addRow(row);
        });
        
        ws.columns.forEach(column => {
          if (!column) return;
          let maxLen = 10;
          column.eachCell?.({ includeEmpty: true }, cell => {
            let val = cell.value ? cell.value.toString() : "";
            if (val.length > maxLen) maxLen = val.length;
          });
          column.width = Math.min(maxLen + 4, 50);
        });
        
        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), "Pivot_Polytex.xlsx");
        alert("Đã tải xuống file Pivot_Polytex.xlsx thành công!");
      } catch (err: any) {
        alert("Lỗi khi xuất file Excel: " + err.message);
      } finally {
        setLoadingMsg('');
      }
    }, 50);
  };

  // ==========================================
  // KHỞI TẠO LUCKYSHEET
  // ==========================================
  useEffect(() => {
    const sheetOptions = {
      container: 'luckysheet-container',
      showinfobar: false,
      lang: 'en',
      hook: {
        rangeSelect: handleRangeSelect
      }
    };
    
    if (typeof luckysheet !== 'undefined') {
      luckysheet.create(sheetOptions);
    }
  }, []); 

  // ==========================================
  // XỬ LÝ KÉO THẢ BẢNG ĐIỀU KHIỂN
  // ==========================================
  const handlePanelMouseDown = (e: React.MouseEvent<HTMLDivElement>) => {
    const startX = e.clientX - panelPos.x;
    const startY = e.clientY - panelPos.y;

    const handleMouseMove = (mouseMoveEvent: MouseEvent) => {
      setPanelPos({
        x: mouseMoveEvent.clientX - startX,
        y: mouseMoveEvent.clientY - startY
      });
    };

    const handleMouseUp = () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };

    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
  };

  // ==========================================
  // XỬ LÝ UPLOAD FILE
  // ==========================================
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoadingMsg("Đang vẽ bảng tính... Vui lòng chờ...");
    
    setTimeout(() => {
      // 1. Đọc và nạp file vào ExcelJS Workbook để giữ nguyên định dạng, cột, nét vẽ gốc
      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          const arrayBuffer = event.target?.result as ArrayBuffer;
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.load(arrayBuffer);
          setWorkbookObj(workbook);
        } catch (err) {
          console.error("Lỗi khi load ExcelJS Workbook gốc:", err);
        }
      };
      reader.readAsArrayBuffer(file);

      // 2. Chuyển đổi và hiển thị lên Luckysheet
      LuckyExcel.transformExcelToLucky(file, function(exportJson: any) {
        if(exportJson.sheets == null || exportJson.sheets.length == 0) { 
          setLoadingMsg(''); 
          return; 
        }
        
        setStatusText("");
        setHasFile(true);
        
        const options = {
          container: 'luckysheet-container',
          showinfobar: false,
          lang: 'en',
          data: exportJson.sheets,
          title: exportJson.info.name,
          hook: {
            rangeSelect: handleRangeSelect
          }
        };
        
        luckysheet.destroy(); 
        luckysheet.create(options);
        setLoadingMsg('');
      });
    }, 100);
  };

  // ==========================================
  // XỬ LÝ XUẤT FILE EXCEL
  // ==========================================
  const handleExport = async () => {
    if (!workbookObj) {
      alert("Không tìm thấy dữ liệu workbook gốc. Vui lòng tải lại tệp và thực hiện thao tác.");
      return;
    }

    setLoadingMsg("Đang đóng gói dữ liệu thành file Excel...");
    
    setTimeout(async () => {
      try {
        const buffer = await workbookObj.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), "VAT.xlsx");
        setStatusText("✅ Tải xuống thành công!");
      } catch(err) {
        alert("Lỗi khi xuất file Excel: " + err);
      } finally {
        setLoadingMsg('');
      }
    }, 50);
  };

  // ==========================================
  // XỬ LÝ CẮT GỌT DỮ LIỆU
  // ==========================================
  const handleProcessData = () => {
    if (!headerData) {
      alert("Vui lòng bôi chọn vùng tiêu đề làm mẫu trực tiếp trên bảng tính trước khi bấm nút!");
      return;
    }

    setLoadingMsg("Đang xử lý dữ liệu...");
    
    setTimeout(() => {
      try {
        let sheetData = luckysheet.getSheetData();
        if (!sheetData || sheetData.length === 0) {
          alert("Không tìm thấy dữ liệu trong sheet hiện tại.");
          setLoadingMsg('');
          return;
        }
        const { pattern, c1, c2, r1, r2 } = headerData;

        // Hàm kiểm tra xem dòng r có chứa bất kỳ ô nào có viền (border) trong khoảng cột từ c1 đến c2 hay không
        const checkRowHasBorder = (r: number) => {
          try {
            const file = luckysheet.getluckysheetfile().find((s: any) => s.status === 1) || luckysheet.getluckysheetfile()[0];
            const borderInfo = file?.config?.borderInfo || [];
            
            for (let item of borderInfo) {
              if (!item) continue;
              
              if (item.rangeType === 'cell') {
                const val = item.value;
                if (val && val.row_index === r && val.col_index >= c1 && val.col_index <= c2) {
                  if (val.l || val.r || val.t || val.b) {
                    return true;
                  }
                }
              }
              
              if (item.rangeType === 'range') {
                const ranges = item.range || (item.value && item.value.range);
                if (!ranges) continue;
                
                for (let rg of ranges) {
                  const rowStart = rg.row[0];
                  const rowEnd = rg.row[1];
                  const colStart = rg.column[0];
                  const colEnd = rg.column[1];
                  
                  if (r >= rowStart && r <= rowEnd) {
                    const intersectColStart = Math.max(c1, colStart);
                    const intersectColEnd = Math.min(c2, colEnd);
                    if (intersectColStart <= intersectColEnd) {
                      if (item.borderType !== 'border-none') {
                        return true;
                      }
                    }
                  }
                }
              }
            }
          } catch (err) {
            console.error("Lỗi khi kiểm tra viền dòng:", err);
          }
          return false;
        };

        let rowsToDelete = new Set<number>();
        const isHeaderMode = currentModeRef.current === 'HEADER';
        const isRpacMode = companyModeRef.current === 'RPAC';

        if (isRpacMode) {
          // ==========================================
          // THUẬT TOÁN ĐẶC THÙ CHO CÔNG TY R-PAC
          // ==========================================
          const getRowCleanedText = (rowCells: any) => {
            if (!rowCells) return "";
            let combinedText = "";
            for (let c = 0; c < rowCells.length; c++) {
              let cell = rowCells[c];
              if (cell) {
                let text = extractTextWithHyperlinkFilter(cell, 'RPAC');
                combinedText += text.toLowerCase();
              }
            }
            // Loại bỏ khoảng trắng, dấu gạch ngang và ký tự đặc biệt để so khớp chính xác
            return combinedText.replace(/\s+/g, '').replace(/[0-9.,:\/\\#\-]/g, '');
          };

          const isRpacStartRow = (rowCells: any) => {
            const cleaned = getRowCleanedText(rowCells);
            return cleaned.includes("côngtytnhhrpacviệtnam") || 
                   cleaned.includes("rpacvietnam") || 
                   cleaned.includes("rpacvietnamlimited");
          };

          const isRpacEndRow = (rowCells: any) => {
            const cleaned = getRowCleanedText(rowCells);
            // Chứa STT và Hàng hóa/Dịch vụ/Mô tả/Đơn vị
            return (cleaned.includes("stt") || cleaned.includes("no")) && 
                   (cleaned.includes("hànghóa") || cleaned.includes("description") || cleaned.includes("diễngiải") || cleaned.includes("tênhàng") || cleaned.includes("đvt") || cleaned.includes("unit"));
          };

          for (let r = 0; r < sheetData.length; r++) {
            if (isRpacStartRow(sheetData[r])) {
              // Tìm dòng kết thúc (dòng STT) phía dưới trong phạm vi gần (tối đa 30 dòng)
              let endHeaderRow = -1;
              for (let i = r; i < Math.min(sheetData.length, r + 30); i++) {
                if (isRpacEndRow(sheetData[i])) {
                  endHeaderRow = i;
                  break;
                }
              }

              if (endHeaderRow !== -1) {
                // Xác định đây là một khối Header của R-pac
                const isMasterBlock = (r >= r1 && r <= r2) || r === r1;
                
                if (isHeaderMode && isMasterBlock) {
                  // Giữ nguyên master header, không thêm vào rowsToDelete
                } else {
                  // Thêm toàn bộ các dòng từ r đến endHeaderRow vào danh sách xóa
                  for (let targetRow = r; targetRow <= endHeaderRow; targetRow++) {
                    rowsToDelete.add(targetRow);
                  }
                }

                // 2. Dọn dẹp các dòng trống phân trang/hàng không viền xung quanh khối Header này
                for (let offset = 1; offset <= 2; offset++) {
                  // Phía trước khối (r - offset)
                  let rowAbove = r - offset;
                  if (rowAbove >= 0 && (isHeaderMode ? !(rowAbove >= r1 && rowAbove <= r2) : true)) {
                    if (!checkRowHasBorder(rowAbove)) {
                      rowsToDelete.add(rowAbove);
                    }
                  }
                  // Phía sau khối (endHeaderRow + offset)
                  let rowBelow = endHeaderRow + offset;
                  if (rowBelow < sheetData.length && (isHeaderMode ? !(rowBelow >= r1 && rowBelow <= r2) : true)) {
                    if (!checkRowHasBorder(rowBelow)) {
                      rowsToDelete.add(rowBelow);
                    }
                  }
                }

                // Nhảy con trỏ quét qua hết khối tiêu đề này
                r = endHeaderRow;
              }
            }
          }
        } else {
          // ==========================================
          // THUẬT TOÁN TIÊU CHUẨN (NORMAL COMPANY)
          // ==========================================
          for (let r = 0; r < sheetData.length; r++) {
            // A. Ở chế độ HEADER: Nếu dòng này nằm trong vùng bôi chọn mẫu ban đầu -> TUYỆT ĐỐI GIỮ NGUYÊN (Lá chắn bảo vệ 1)
            // Ở chế độ FOOTER: Không giữ lại vùng bôi chọn mẫu, xóa sạch tất cả.
            if (isHeaderMode) {
              if (r >= r1 && r <= r2) {
                continue;
              }
            }

            if (matchHeader(sheetData, r, pattern, c1, c2, companyModeRef.current)) {
              // Ở chế độ HEADER: Chỉ xóa các dòng lặp lại nằm dưới dòng mẫu gốc (r > r2)
              // Ở chế độ FOOTER: Xóa ở mọi vị trí bất kỳ
              if (isHeaderMode && r <= r2) {
                continue;
              }

              let lastRowTemplate = pattern[pattern.length - 1]; // Dòng cuối cùng của mẫu (A | B | C...)
              let endHeaderRow = -1;

              // Quét tìm dòng khớp với lastRowTemplate (dòng chứa STT, Hàng hóa...) trong phạm vi gần
              const maxScanRows = Math.min(sheetData.length - r, pattern.length + 3);
              for (let i = 0; i < maxScanRows; i++) {
                let targetRow = r + i;
                if (matchSingleRow(sheetData[targetRow], lastRowTemplate, c1, c2, companyModeRef.current)) {
                  endHeaderRow = targetRow;
                  break;
                }
              }

              // Nếu tìm thấy dòng STT (dòng cuối của khối tiêu đề)
              if (endHeaderRow !== -1) {
                for (let targetRow = r; targetRow <= endHeaderRow; targetRow++) {
                  // Áp dụng bộ lọc bảo vệ mẫu gốc nếu ở chế độ HEADER
                  if (isHeaderMode) {
                    if (!(targetRow >= r1 && targetRow <= r2)) {
                      rowsToDelete.add(targetRow);
                    }
                  } else {
                    rowsToDelete.add(targetRow);
                  }
                }
              } else {
                // Dự phòng: Nếu không tìm thấy dòng STT cuối cùng, dùng cơ chế so khớp từng dòng như cũ
                for (let i = 0; i < pattern.length; i++) {
                  let targetRow = r + i;
                  if (targetRow >= sheetData.length) break;

                  let matchesAny = false;
                  let targetRowText = getCleanedRangeText(sheetData[targetRow], c1, c2, companyModeRef.current);
                  
                  if (targetRowText === "") {
                    matchesAny = true;
                  } else {
                    for (let j = 0; j < pattern.length; j++) {
                      if (matchSingleRow(sheetData[targetRow], pattern[j], c1, c2, companyModeRef.current)) {
                        matchesAny = true;
                        break;
                      }
                    }
                  }

                  if (!matchesAny) {
                    break;
                  }

                  if (isHeaderMode) {
                    if (!(targetRow >= r1 && targetRow <= r2)) {
                      rowsToDelete.add(targetRow);
                    }
                  } else {
                    rowsToDelete.add(targetRow);
                  }

                  if (matchSingleRow(sheetData[targetRow], lastRowTemplate, c1, c2, companyModeRef.current)) {
                    break;
                  }
                }
              }

              // 2. Xóa các dòng trống/không viền xung quanh để dọn dẹp phân trang
              // Sử dụng chỉ số dòng bắt đầu r để dọn dẹp các dòng trống phía trước (r - 1, r - 2)
              for (let offset = 1; offset <= 2; offset++) {
                // Dọn dẹp phía trước (r - offset)
                let rowAbove = r - offset;
                if (rowAbove >= 0 && (isHeaderMode ? !(rowAbove >= r1 && rowAbove <= r2) : true)) {
                  if (!checkRowHasBorder(rowAbove)) {
                    rowsToDelete.add(rowAbove);
                  }
                }
                // Dọn dẹp phía sau (dựa vào dòng cuối thực tế của khối tiêu đề)
                let actualEndRow = endHeaderRow !== -1 ? endHeaderRow : (r + pattern.length - 1);
                let rowBelow = actualEndRow + offset;
                if (rowBelow < sheetData.length && (isHeaderMode ? !(rowBelow >= r1 && rowBelow <= r2) : true)) {
                  if (!checkRowHasBorder(rowBelow)) {
                    rowsToDelete.add(rowBelow);
                  }
                }
              }

              // 3. Cuối cùng mới cập nhật chỉ số r nhảy cóc qua khối tiêu đề đã xử lý
              if (endHeaderRow !== -1) {
                r = endHeaderRow;
              } else {
                r += pattern.length - 1;
              }
            }
          }
        }

        if (rowsToDelete.size === 0) {
          alert(`Không tìm thấy hàng nào khớp với mẫu ${isHeaderMode ? 'Header' : 'Footer'} bôi chọn để cắt gọt!`);
          setLoadingMsg('');
          return;
        }

        // Sắp xếp các dòng cần xóa theo thứ tự giảm dần
        let sortedRows = Array.from(rowsToDelete).sort((a, b) => b - a);

        // Gom nhóm các dòng liên tiếp để xóa hàng loạt một cách tối ưu bằng API Luckysheet
        let deleteBlocks: { start: number; len: number }[] = [];
        if (sortedRows.length > 0) {
          let currentStart = sortedRows[0];
          let currentLen = 1;
          for (let i = 1; i < sortedRows.length; i++) {
            if (sortedRows[i] === currentStart - 1) {
              currentStart = sortedRows[i];
              currentLen++;
            } else {
              deleteBlocks.push({ start: currentStart, len: currentLen });
              currentStart = sortedRows[i];
              currentLen = 1;
            }
          }
          deleteBlocks.push({ start: currentStart, len: currentLen });
        }

        // Sử dụng API native của Luckysheet để xóa dòng.
        for (let block of deleteBlocks) {
          let rowStart = block.start;
          let rowEnd = block.start + block.len - 1;
          luckysheet.deleteRow(rowStart, rowEnd);
        }

        // Xóa đồng thời các dòng này trong ExcelJS Workbook gốc để đồng bộ dữ liệu xuất ra
        if (workbookObj) {
          try {
            const activeSheetIndex = luckysheet.getluckysheetfile().findIndex((s: any) => s.status === 1);
            const ws = workbookObj.worksheets[activeSheetIndex >= 0 ? activeSheetIndex : 0];
            for (let block of deleteBlocks) {
              const startRowExcel = block.start + 1; // ExcelJS là 1-indexed nên cần cộng 1
              ws.spliceRows(startRowExcel, block.len);
              // Đồng bộ điều chỉnh dịch chuyển tọa độ các ô đã gộp (merged cells) phía dưới
              shiftWorksheetMerges(ws, startRowExcel, block.len);
            }
          } catch (err) {
            console.error("Lỗi khi xóa dòng đồng bộ trong ExcelJS:", err);
          }
        }

        let rowNumbers = sortedRows.map(r => r + 1).reverse().join(', '); // Hiển thị tăng dần
        const labelMode = isHeaderMode ? 'Header' : 'Footer';
        setStatusText(`✅ Đã cắt gọt thành công! Đã xóa và dồn ${rowsToDelete.size} hàng.`);
        setPreviewHTML(`<span style="color:green; font-weight:bold;">Đã cắt gọt thành công!</span><br>Đã loại bỏ các ${labelMode} lặp lại và dọn dẹp các hàng không có đường viền xung quanh.<br/><br/><b>Danh sách các hàng đã bị cắt gọt:</b><br/><div style="color:#d35400; font-family:monospace; word-break:break-all; max-height:80px; overflow-y:auto; margin-top:5px; padding:5px; background:#fff; border:1px solid #ddd; border-radius:4px; font-weight:bold;">Hàng: ${rowNumbers}</div>`);

      } catch (err: any) {
        alert("Có lỗi xảy ra khi xử lý dữ liệu: " + err.message);
      } finally {
        setLoadingMsg('');
      }
    }, 200);
  };

  // ==========================================
  // XỬ LÝ DỌN DẸP DÒNG TRỐNG
  // ==========================================
  const handleRemoveBlankRows = () => {
    setLoadingMsg("Đang dọn dẹp các dòng trống...");
    
    setTimeout(() => {
      try {
        let sheetData = luckysheet.getSheetData();
        if (!sheetData || sheetData.length === 0) {
          alert("Không tìm thấy dữ liệu trong sheet hiện tại.");
          setLoadingMsg('');
          return;
        }

        let rowsToDelete = new Set<number>();

        // Duyệt qua tất cả các dòng của sheet
        for (let r = 0; r < sheetData.length; r++) {
          let rowData = sheetData[r];
          if (!rowData) {
            rowsToDelete.add(r);
            continue;
          }

          let isBlank = true;
          for (let c = 0; c < rowData.length; c++) {
            let cell = rowData[c];
            if (cell) {
              let text = "";
              if (cell.ct && cell.ct.t === 'inlineStr' && cell.ct.s) {
                text = cell.ct.s.map((s: any) => s.v).join("");
              } else if (cell.m !== undefined && cell.m !== null) {
                text = cell.m.toString();
              } else if (cell.v !== undefined && cell.v !== null) {
                text = cell.v.toString();
              }
              
              if (text.trim() !== "") {
                isBlank = false;
                break;
              }
            }
          }

          if (isBlank) {
            rowsToDelete.add(r);
          }
        }

        if (rowsToDelete.size === 0) {
          alert("Bảng tính hiện tại đã hoàn toàn sạch sẽ, không có dòng trống nào để dọn dẹp!");
          setLoadingMsg('');
          return;
        }

        // Sắp xếp giảm dần để xóa từ dưới lên
        let sortedRows = Array.from(rowsToDelete).sort((a, b) => b - a);

        // Gom nhóm các dòng liên tục
        let deleteBlocks: { start: number; len: number }[] = [];
        if (sortedRows.length > 0) {
          let currentStart = sortedRows[0];
          let currentLen = 1;
          for (let i = 1; i < sortedRows.length; i++) {
            if (sortedRows[i] === currentStart - 1) {
              currentStart = sortedRows[i];
              currentLen++;
            } else {
              deleteBlocks.push({ start: currentStart, len: currentLen });
              currentStart = sortedRows[i];
              currentLen = 1;
            }
          }
          deleteBlocks.push({ start: currentStart, len: currentLen });
        }

        // Thực hiện xóa dòng thông qua API Native của Luckysheet
        for (let block of deleteBlocks) {
          let rowStart = block.start;
          let rowEnd = block.start + block.len - 1;
          luckysheet.deleteRow(rowStart, rowEnd);
        }

        // Xóa đồng thời các dòng này trong ExcelJS Workbook gốc để đồng bộ dữ liệu xuất ra
        if (workbookObj) {
          try {
            const activeSheetIndex = luckysheet.getluckysheetfile().findIndex((s: any) => s.status === 1);
            const ws = workbookObj.worksheets[activeSheetIndex >= 0 ? activeSheetIndex : 0];
            for (let block of deleteBlocks) {
              const startRowExcel = block.start + 1; // ExcelJS là 1-indexed nên cần cộng 1
              ws.spliceRows(startRowExcel, block.len);
              // Đồng bộ điều chỉnh dịch chuyển tọa độ các ô đã gộp (merged cells) phía dưới
              shiftWorksheetMerges(ws, startRowExcel, block.len);
            }
          } catch (err) {
            console.error("Lỗi khi xóa dòng đồng bộ trong ExcelJS:", err);
          }
        }

        setStatusText(`🧹 Đã dọn dẹp thành công! Đã xóa và dồn ${rowsToDelete.size} hàng trống.`);
        alert(`Đã dọn dẹp thành công! Đã loại bỏ và dồn ${rowsToDelete.size} hàng trống.`);

      } catch (err: any) {
        alert("Có lỗi xảy ra khi dọn dẹp dòng trống: " + err.message);
      } finally {
        setLoadingMsg('');
      }
    }, 200);
  };

  return (
    <div className="h-screen w-screen overflow-hidden flex flex-col relative">
      
      {/* 1. Header Bar */}
      <div className="bg-[#107c41] px-5 py-3 text-white flex items-center gap-4 z-10">
        <strong className="text-lg">Excel Tool</strong>
        
        {/* Input ẩn để nhận file Excel */}
        <input 
          type="file" 
          id="excel-upload-input"
          accept=".xlsx" 
          onChange={handleFileUpload}
          className="hidden" 
        />

        <button 
          onClick={() => document.getElementById('excel-upload-input')?.click()}
          className="bg-[#217346] hover:bg-[#1a5c38] text-white border border-[#217346] hover:border-white px-4 py-1.5 rounded font-bold transition-all text-sm flex items-center gap-1.5 active:scale-95 shadow"
        >
          📥 Nhập Excel
        </button>
        
        <button 
          onClick={() => {
            setShowPanel(true);
            setPanelPos({ x: 0, y: 0 }); // Reset vị trí bảng về giữa khi mở lại
          }}
          className="bg-[#2b579a] hover:bg-[#1e3f6f] border border-white px-4 py-2 rounded font-bold transition-colors"
        >
          ⚙️ Bảng Điều Khiển Cắt Gộp
        </button>
        <button 
          onClick={openPivotModal}
          className="bg-[#0078d4] hover:bg-[#106ebe] border border-white px-4 py-2 rounded font-bold transition-colors"
        >
          📊 Pivot Polytex
        </button>
        <button 
          onClick={handleRemoveBlankRows}
          className="bg-[#512da8] hover:bg-[#311b92] border border-white px-4 py-2 rounded font-bold transition-colors"
        >
          🧹 Xóa Hàng Trống
        </button>
        <button 
          onClick={handleExport}
          className="bg-[#d83b01] hover:bg-[#a82e00] border border-white px-4 py-2 rounded font-bold transition-colors"
        >
          💾 Tải Xuống Excel
        </button>
        <span className="text-sm italic">{statusText}</span>
      </div>

      {/* 2. Container cho Luckysheet */}
      <div id="luckysheet-container" className="flex-1 w-full relative"></div>

      {/* Excel-style Status Bar at the bottom */}
      <div className="bg-[#f3f2f1] border-t border-[#d2d0ce] h-7 px-5 text-xs text-[#323130] flex items-center justify-between select-none z-10">
        <div className="flex items-center gap-2">
          <span className="font-bold text-[#107c41] text-[10px]">●</span> Ready
        </div>
        {selectionStats && (
          <div className="flex items-center gap-4 font-sans text-[#323130] select-text">
            <span>Average: <strong className="text-gray-900 font-bold">{selectionStats.average}</strong></span>
            <span className="text-gray-300">|</span>
            <span>Count: <strong className="text-gray-900 font-bold">{selectionStats.count}</strong></span>
            <span className="text-gray-300">|</span>
            <span>Sum: <strong className="text-gray-900 font-bold">{selectionStats.sum}</strong></span>
          </div>
        )}
      </div>

      {/* 3. Bảng điều khiển nổi (Floating Panel) */}
      {showPanel && (
        <div 
          className="absolute w-[550px] bg-white rounded-lg shadow-2xl z-50 overflow-hidden border border-gray-200"
          style={{
            top: '50%',
            left: '50%',
            transform: `translate(calc(-50% + ${panelPos.x}px), calc(-50% + ${panelPos.y}px))`
          }}
        >
          {/* Header để Kéo thả */}
          <div 
            onMouseDown={handlePanelMouseDown}
            className="bg-gray-800 text-white px-5 py-3 flex justify-between items-center cursor-move select-none"
          >
            <span className="font-bold">✂️ Thiết lập Cắt Gộp Đồng Thời</span>
            <button onClick={() => setShowPanel(false)} className="text-red-400 font-bold hover:text-red-300">X</button>
          </div>
          
          <div className="p-5">
            <div className="bg-slate-50 p-4 border-l-4 border-[#107c41] text-sm text-gray-700 mb-4 rounded shadow-sm">
              <strong className="block text-[#107c41] mb-1">💡 Hướng dẫn sử dụng cực nhanh:</strong>
              1. Bạn hãy quét bôi chọn (bôi đen) vùng tiêu đề hoặc chân trang cần làm mẫu trực tiếp trên bảng tính Excel. <br />
              2. Chọn chức năng tương ứng bên dưới và nhấp nút <b>"Xử lý Toàn bộ"</b> để hệ thống tự động loại bỏ.
            </div>

            {/* Chọn Chức Năng Cắt Gộp */}
            <div className="mb-4">
              <span className="block font-bold text-gray-700 mb-2 text-sm">Chức năng xử lý:</span>
              <div className="flex gap-4">
                <label className={`flex items-center gap-2.5 cursor-pointer border px-4 py-2.5 rounded-lg flex-1 transition-all select-none ${currentMode === 'HEADER' ? 'border-[#107c41] bg-green-50/30' : 'border-gray-200 bg-slate-50'}`}>
                  <input 
                    type="radio" 
                    name="trimMode" 
                    value="HEADER" 
                    checked={currentMode === 'HEADER'}
                    onChange={() => {
                      setCurrentMode('HEADER');
                      if (headerData) {
                        const label = 'Header mẫu';
                        const labelColor = '#e53935';
                        let html = "";
                        // @ts-ignore
                        let sheetData = luckysheet.getSheetData();
                        for (let r = headerData.r1; r <= headerData.r2; r++) {
                          let display = getRangeText(sheetData[r], headerData.c1, headerData.c2, companyMode);
                          let bgColor = (r % 2 === 0) ? '#f4f7f6' : '#eaf1ee';
                          html += `<div style="background:${bgColor}; padding: 3px 0; border-bottom: 1px solid #ddd; color: black;"><b>D. ${r + 1}:</b> ${display}</div>`;
                        }
                        setPreviewHTML(`<span style="color:${labelColor}; font-weight:bold;">Đã nhận diện ${label} (${headerData.rangeName}):</span><br><div style="margin-top:5px; max-height:80px; overflow-y:auto; border:1px solid #ddd; padding:5px; background:#fff; border-radius:4px;">${html}</div>`);
                      }
                    }}
                    className="accent-[#107c41] w-4 h-4 cursor-pointer"
                  />
                  <div>
                    <span className="font-bold text-sm text-gray-800 block">Xóa Header</span>
                    <span className="text-[11px] text-gray-500">Giữ tiêu đề gốc</span>
                  </div>
                </label>

                <label className={`flex items-center gap-2.5 cursor-pointer border px-4 py-2.5 rounded-lg flex-1 transition-all select-none ${currentMode === 'FOOTER' ? 'border-[#2b579a] bg-blue-50/30' : 'border-gray-200 bg-slate-50'}`}>
                  <input 
                    type="radio" 
                    name="trimMode" 
                    value="FOOTER" 
                    checked={currentMode === 'FOOTER'}
                    onChange={() => {
                      setCurrentMode('FOOTER');
                      if (headerData) {
                        const label = 'Footer mẫu';
                        const labelColor = '#2b579a';
                        let html = "";
                        // @ts-ignore
                        let sheetData = luckysheet.getSheetData();
                        for (let r = headerData.r1; r <= headerData.r2; r++) {
                          let display = getRangeText(sheetData[r], headerData.c1, headerData.c2, companyMode);
                          let bgColor = (r % 2 === 0) ? '#f4f7f6' : '#eaf1ee';
                          html += `<div style="background:${bgColor}; padding: 3px 0; border-bottom: 1px solid #ddd; color: black;"><b>D. ${r + 1}:</b> ${display}</div>`;
                        }
                        setPreviewHTML(`<span style="color:${labelColor}; font-weight:bold;">Đã nhận diện ${label} (${headerData.rangeName}):</span><br><div style="margin-top:5px; max-height:80px; overflow-y:auto; border:1px solid #ddd; padding:5px; background:#fff; border-radius:4px;">${html}</div>`);
                      }
                    }}
                    className="accent-[#2b579a] w-4 h-4 cursor-pointer"
                  />
                  <div>
                    <span className="font-bold text-sm text-gray-800 block">Xóa Footer</span>
                    <span className="text-[11px] text-gray-500">Xóa sạch cả vùng mẫu</span>
                  </div>
                </label>
              </div>
            </div>

            {/* Chọn Cấu Hình Theo Công Ty */}
            <div className="mb-4">
              <span className="block font-bold text-gray-700 mb-2 text-sm">Cấu hình theo công ty:</span>
              <div className="flex gap-4">
                <label className={`flex items-center gap-3 cursor-pointer border px-4 py-2.5 rounded-lg flex-1 transition-all select-none ${companyMode === 'NORMAL' ? 'border-[#107c41] bg-green-50/20' : 'border-gray-200 bg-slate-50'}`}>
                  <input 
                    type="radio" 
                    name="companyMode" 
                    value="NORMAL" 
                    checked={companyMode === 'NORMAL'}
                    onChange={() => {
                      setCompanyMode('NORMAL');
                      if (headerData) {
                        const isHeader = currentMode === 'HEADER';
                        const label = isHeader ? 'Header mẫu' : 'Footer mẫu';
                        const labelColor = isHeader ? '#e53935' : '#2b579a';
                        let html = "";
                        // @ts-ignore
                        let sheetData = luckysheet.getSheetData();
                        for (let r = headerData.r1; r <= headerData.r2; r++) {
                          let display = getRangeText(sheetData[r], headerData.c1, headerData.c2, 'NORMAL');
                          let bgColor = (r % 2 === 0) ? '#f4f7f6' : '#eaf1ee';
                          html += `<div style="background:${bgColor}; padding: 3px 0; border-bottom: 1px solid #ddd; color: black;"><b>D. ${r + 1}:</b> ${display}</div>`;
                        }
                        setPreviewHTML(`<span style="color:${labelColor}; font-weight:bold;">Đã nhận diện ${label} (${headerData.rangeName}):</span><br><div style="margin-top:5px; max-height:80px; overflow-y:auto; border:1px solid #ddd; padding:5px; background:#fff; border-radius:4px;">${html}</div>`);
                      }
                    }}
                    className="accent-[#107c41] w-4 h-4 cursor-pointer"
                  />
                  <div className="flex-1">
                    <span className="font-bold text-sm text-gray-800 block">Normal Company</span>
                    <span className="text-[11px] text-gray-500">Thuật toán chuẩn cũ</span>
                  </div>
                </label>

                <label className={`flex items-center gap-3 cursor-pointer border px-4 py-2.5 rounded-lg flex-1 transition-all select-none ${companyMode === 'RPAC' ? 'border-red-500 bg-red-50/20' : 'border-gray-200 bg-slate-50'}`}>
                  <input 
                    type="radio" 
                    name="companyMode" 
                    value="RPAC" 
                    checked={companyMode === 'RPAC'}
                    onChange={() => {
                      setCompanyMode('RPAC');
                      if (headerData) {
                        const isHeader = currentMode === 'HEADER';
                        const label = isHeader ? 'Header mẫu' : 'Footer mẫu';
                        const labelColor = isHeader ? '#e53935' : '#2b579a';
                        let html = "";
                        // @ts-ignore
                        let sheetData = luckysheet.getSheetData();
                        for (let r = headerData.r1; r <= headerData.r2; r++) {
                          let display = getRangeText(sheetData[r], headerData.c1, headerData.c2, 'RPAC');
                          let bgColor = (r % 2 === 0) ? '#f4f7f6' : '#eaf1ee';
                          html += `<div style="background:${bgColor}; padding: 3px 0; border-bottom: 1px solid #ddd; color: black;"><b>D. ${r + 1}:</b> ${display}</div>`;
                        }
                        setPreviewHTML(`<span style="color:${labelColor}; font-weight:bold;">Đã nhận diện ${label} (${headerData.rangeName}):</span><br><div style="margin-top:5px; max-height:80px; overflow-y:auto; border:1px solid #ddd; padding:5px; background:#fff; border-radius:4px;">${html}</div>`);
                      }
                    }}
                    className="accent-red-500 w-4 h-4 cursor-pointer"
                  />
                  <div className="flex items-center justify-between flex-1 gap-2">
                    <div>
                      <span className="font-bold text-sm text-gray-800 block">R-pac</span>
                      <span className="text-[11px] text-gray-500">Lọc link & domain</span>
                    </div>
                    <img src="/rpac_logo.png" alt="R-pac Logo" className="h-6 w-auto object-contain" />
                  </div>
                </label>
              </div>
            </div>

            {/* Vùng Preview */}
            {previewHTML && (
              <div 
                className="bg-green-50 p-3 border-l-4 border-green-500 font-mono text-sm max-h-32 overflow-y-auto mb-4" 
                dangerouslySetInnerHTML={{ __html: previewHTML }}
              />
            )}

            <div className="text-right pt-4 border-t border-gray-100">
              <button onClick={() => setShowPanel(false)} className="bg-gray-300 hover:bg-gray-400 text-black px-4 py-2 rounded mr-2 font-medium transition-colors">Ẩn Bảng</button>
              <button 
                className="bg-[#107c41] hover:bg-[#0c5e31] text-white px-5 py-2 rounded font-bold transition-all shadow-md active:scale-95"
                onClick={handleProcessData}
              >
                ✂️ Xử lý Toàn bộ
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 3.1 Bảng điều khiển Pivot Polytex (Floating Pivot Modal) */}
      {showPivotModal && (
        <div 
          className="absolute w-[650px] bg-white rounded-lg shadow-2xl z-50 overflow-hidden border border-gray-200 flex flex-col max-h-[85vh]"
          style={{
            top: '45%',
            left: '50%',
            transform: `translate(calc(-50% + ${pivotModalPos.x}px), calc(-50% + ${pivotModalPos.y}px))`
          }}
        >
          {/* Header để Kéo thả */}
          <div 
            onMouseDown={handlePivotPanelMouseDown}
            className="bg-sky-800 text-white px-5 py-3 flex justify-between items-center cursor-move select-none shrink-0"
          >
            <div className="flex items-center gap-2">
              <span>📊</span>
              <span className="font-bold">Pivot Table Builder (Polytex)</span>
            </div>
            <button onClick={() => setShowPivotModal(false)} className="text-red-300 font-bold hover:text-red-100 text-lg">×</button>
          </div>
          
          <div className="p-5 overflow-y-auto flex-1 space-y-4">
            <div className="bg-sky-50/50 p-3.5 border-l-4 border-[#0078d4] text-xs text-slate-700 rounded shadow-sm">
              <strong className="block text-[#0078d4] mb-1 font-semibold">💡 Hướng dẫn Pivot:</strong>
              Hệ thống sẽ gom nhóm các dòng dữ liệu trùng khớp dựa trên các cột bạn chọn ở mục <b>Gom Nhóm</b>, và tự động tính tổng (Sum) các cột số lượng/thành tiền được chọn ở mục <b>Tính Tổng</b>.
            </div>

            {/* Config options */}
            <div className="grid grid-cols-2 gap-4">
              {/* Group columns selection */}
              <div className="border border-slate-200 rounded-lg p-3 bg-slate-50/20">
                <span className="block font-bold text-slate-800 mb-2 text-xs border-b pb-1">🗂️ Cột Hiển Thị / Gom Nhóm:</span>
                <div className="space-y-1.5 max-h-40 overflow-y-auto pr-1">
                  {pivotColumns.map(col => (
                    <label key={col.index} className="flex items-center gap-2 cursor-pointer text-xs select-none hover:text-sky-700 transition-colors">
                      <input 
                        type="checkbox" 
                        checked={selectedPivotGroupCols.includes(col.index)}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setSelectedPivotGroupCols([...selectedPivotGroupCols, col.index]);
                          } else {
                            setSelectedPivotGroupCols(selectedPivotGroupCols.filter(i => i !== col.index));
                          }
                        }}
                        className="rounded border-gray-300 text-sky-600 focus:ring-sky-500"
                      />
                      <span>{col.name}</span>
                    </label>
                  ))}
                </div>
              </div>

              {/* Aggregation columns selection */}
              <div className="border border-slate-200 rounded-lg p-3 bg-slate-50/20">
                <span className="block font-bold text-slate-800 mb-2 text-xs border-b pb-1">📈 Cột Tính Tổng (Sum):</span>
                <div className="space-y-1.5 max-h-40 overflow-y-auto pr-1">
                  {pivotColumns.filter(col => col.index !== -100).map(col => (
                    <label key={col.index} className="flex items-center gap-2 cursor-pointer text-xs select-none hover:text-sky-700 transition-colors">
                      <input 
                        type="checkbox" 
                        checked={selectedPivotAggCols.includes(col.index)}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setSelectedPivotAggCols([...selectedPivotAggCols, col.index]);
                          } else {
                            setSelectedPivotAggCols(selectedPivotAggCols.filter(i => i !== col.index));
                          }
                        }}
                        className="rounded border-gray-300 text-sky-600 focus:ring-sky-500"
                      />
                      <span className={col.isNumeric ? "font-medium" : "text-slate-500"}>
                        {col.name} {col.isNumeric ? "🔢" : ""}
                      </span>
                    </label>
                  ))}
                </div>
              </div>
            </div>

            {/* Special Polytex settings */}
            {pivotColumns.some(col => col.index === -100) && (
              <div className="bg-amber-50/40 border border-amber-200 rounded-lg p-3 flex items-center justify-between">
                <div className="text-xs">
                  <span className="font-bold text-amber-800 block">⚙️ Cấu hình đặc biệt Polytex</span>
                  <span className="text-slate-500 text-[11px]">Tự động tách mã số và mô tả hàng hóa để so khớp gom nhóm sạch hơn</span>
                </div>
                <label className="relative inline-flex items-center cursor-pointer select-none">
                  <input 
                    type="checkbox" 
                    checked={splitPolytexCode}
                    onChange={(e) => setSplitPolytexCode(e.target.checked)}
                    className="sr-only peer"
                  />
                  <div className="w-9 h-5 bg-gray-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-amber-500"></div>
                </label>
              </div>
            )}

            {/* Action buttons */}
            <div className="text-center pt-2">
              <button 
                onClick={handleGeneratePivot}
                className="bg-sky-700 hover:bg-sky-800 text-white font-bold py-2 px-6 rounded-md shadow transition-all active:scale-95 text-xs"
              >
                📊 Phân Tích & Tạo Pivot
              </button>
            </div>

            {/* Pivot Results Preview */}
            {pivotResult && (
              <div className="border border-slate-200 rounded-lg overflow-hidden bg-white shadow-inner">
                <div className="bg-slate-100 px-3 py-1.5 border-b text-xs font-bold text-slate-700 flex justify-between items-center">
                  <span>📋 Xem trước kết quả ({pivotResult.rows.length} dòng):</span>
                </div>
                <div className="max-h-56 overflow-auto">
                  <table className="w-full text-left border-collapse text-[11px] font-sans">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-200 sticky top-0">
                        {pivotResult.headers.map((h, i) => (
                          <th key={i} className="p-2 font-bold text-slate-700 border-r border-slate-200">{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {pivotResult.rows.map((row, rowIndex) => (
                        <tr key={rowIndex} className="border-b border-slate-100 hover:bg-slate-50 transition-colors">
                          {row.map((val: any, colIndex: number) => (
                            <td key={colIndex} className="p-2 border-r border-slate-100 text-slate-600 max-w-[200px] truncate" title={val}>
                              {val !== null && val !== undefined ? val.toString() : ""}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>

          {/* Footer controls */}
          <div className="bg-slate-50 p-4 border-t border-slate-100 flex justify-end gap-2 shrink-0">
            <button 
              onClick={() => setShowPivotModal(false)} 
              className="bg-gray-300 hover:bg-gray-400 text-black px-4 py-2 rounded font-medium transition-colors text-xs"
            >
              Đóng
            </button>
            {pivotResult && (
              <>
                <button 
                  onClick={handleExportPivotToNewSheet}
                  className="bg-[#107c41] hover:bg-[#0c5e31] text-white px-4 py-2 rounded font-bold transition-all shadow active:scale-95 text-xs"
                >
                  📝 Xuất ra Sheet mới
                </button>
                <button 
                  onClick={handleExportPivotToExcel}
                  className="bg-[#d83b01] hover:bg-[#a82e00] text-white px-4 py-2 rounded font-bold transition-all shadow active:scale-95 text-xs"
                >
                  💾 Tải Excel Pivot
                </button>
              </>
            )}
          </div>
        </div>
      )}

      {/* 4. Màn hình Loading (Loading Overlay) */}
      {loadingMsg && (
        <div className="fixed inset-0 bg-white/90 z-[9999] flex flex-col justify-center items-center">
          <div className="animate-spin rounded-full h-16 w-16 border-t-4 border-b-4 border-[#107c41]"></div>
          <div className="mt-4 font-bold text-[#107c41] text-lg">{loadingMsg}</div>
        </div>
      )}
      
    </div>
  );
}