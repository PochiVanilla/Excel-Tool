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

const getRangeText = (row: any, c1: number, c2: number) => {
  if (!row) return "";
  let textArray = [];
  for (let i = c1; i <= c2; i++) {
    let cell = row[i];
    if (cell) {
      let text = "";
      if (cell.ct && cell.ct.t === 'inlineStr' && cell.ct.s) {
        text = cell.ct.s.map((s: any) => s.v).join("");
      } else if (cell.m !== undefined && cell.m !== null) {
        text = cell.m.toString();
      } else if (cell.v !== undefined && cell.v !== null) {
        text = cell.v.toString();
      }
      textArray.push(text.trim().replace(/\r?\n/g, ' '));
    } else {
      textArray.push("");
    }
  }
  return textArray.join(' | ');
};

const getCleanedRangeText = (row: any, c1: number, c2: number) => {
  if (!row) return "";
  let textArray = [];
  for (let i = c1; i <= c2; i++) {
    let cell = row[i];
    let text = "";
    if (cell) {
      if (cell.ct && cell.ct.t === 'inlineStr' && cell.ct.s) {
        text = cell.ct.s.map((s: any) => s.v).join("");
      } else if (cell.m !== undefined && cell.m !== null) {
        text = cell.m.toString();
      } else if (cell.v !== undefined && cell.v !== null) {
        text = cell.v.toString();
      }
    }
    // Viết thường, loại bỏ toàn bộ khoảng trắng và ký tự đặc biệt để so khớp chính xác nhất
    let cleaned = text.toLowerCase().replace(/\s+/g, '').replace(/[0-9.,:\/\\#\-]/g, '');
    textArray.push(cleaned);
  }
  return textArray.join('|');
};

const matchSingleRow = (row: any, patternRowText: string, c1: number, c2: number) => {
  if (!row) return false;
  let pCells = patternRowText.split('|').map(s => s.trim()).filter(s => s !== "");
  if (pCells.length === 0) return true; // Dòng mẫu trống thì luôn khớp
  
  let rowTextCombined = "";
  for (let c = c1; c <= c2; c++) {
    let cell = row[c];
    let text = "";
    if (cell) {
      if (cell.ct && cell.ct.t === 'inlineStr' && cell.ct.s) {
        text = cell.ct.s.map((s: any) => s.v).join("");
      } else if (cell.m !== undefined && cell.m !== null) {
        text = cell.m.toString();
      } else if (cell.v !== undefined && cell.v !== null) {
        text = cell.v.toString();
      }
    }
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

const matchHeader = (sheetData: any[], r: number, pattern: string[], c1: number, c2: number) => {
  if (r >= sheetData.length) return false;
  
  // Chỉ cần kiểm tra dòng đầu tiên r có khớp với pattern[0] không.
  // Nhờ cơ chế khớp 80% số ô có dữ liệu của mẫu, việc này đã đủ chính xác để tránh nhận diện nhầm các dòng chữ ký.
  return matchSingleRow(sheetData[r], pattern[0], c1, c2);
};

const matchFooter = (sheetData: any[], r: number, fuzzyPattern: string[], c1: number, c2: number) => {
  if (r + fuzzyPattern.length > sheetData.length) return false;
  for (let i = 0; i < fuzzyPattern.length; i++) {
    let text = getCleanedRangeText(sheetData[r + i], c1, c2);
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

  // State quản lý logic cắt gộp
  const [currentMode, setCurrentMode] = useState<'HEADER' | 'FOOTER' | null>(null);
  const [headerData, setHeaderData] = useState<any>(null);
  const [footerData, setFooterData] = useState<any>(null);
  const [previewHTML, setPreviewHTML] = useState<string>('');
  
  // State quản lý Kéo thả bảng điều khiển
  const [panelPos, setPanelPos] = useState({ x: 0, y: 0 });

  // Refs để Luckysheet (nằm ngoài chu kỳ render của React) luôn đọc được state mới nhất
  const currentModeRef = useRef(currentMode);
  const showPanelRef = useRef(showPanel);

  useEffect(() => { currentModeRef.current = currentMode; }, [currentMode]);
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
          let text = getCleanedRangeText(sheetData[r], c1, c2);
          pattern.push(text);
          let display = getRangeText(sheetData[r], c1, c2);
          let bgColor = (r % 2 === 0) ? '#f4f7f6' : '#eaf1ee';
          html += `<div style="background:${bgColor}; padding: 3px 0; border-bottom: 1px solid #ddd; color: black;"><b>D. ${r + 1}:</b> ${display}</div>`;
        }

        if (pattern.every(t => t.replace(/\|/g, '').trim() === '')) {
          setPreviewHTML(`<span style="color:red;">Vùng chọn trống, vui lòng quét chọn lại!</span>`);
          setHeaderData(null);
        } else {
          setHeaderData({ pattern, c1, c2, r1, r2, rangeName });
          setPreviewHTML(`<span style="color:#e53935; font-weight:bold;">Đã nhận diện Header mẫu (${rangeName}):</span><br><div style="margin-top:5px; max-height:80px; overflow-y:auto; border:1px solid #ddd; padding:5px; background:#fff; border-radius:4px;">${html}</div>`);
        }
      }
    }
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
    setLoadingMsg("Đang đóng gói dữ liệu thành file Excel...");
    
    setTimeout(async () => {
      try {
        let sheets = luckysheet.getluckysheetfile();
        let workbook = new ExcelJS.Workbook();
        
        sheets.forEach((sheet: any) => {
           let worksheet = workbook.addWorksheet(sheet.name || 'Sheet1');
           if(sheet.data) {
             for(let r = 0; r < sheet.data.length; r++) {
               let rowData = sheet.data[r];
               if(!rowData) continue;
               for(let c = 0; c < rowData.length; c++) {
                 if(rowData[c] && rowData[c].v !== undefined) {
                   worksheet.getCell(r+1, c+1).value = (rowData[c].m !== undefined) ? rowData[c].m : rowData[c].v;
                 }
               }
             }
           }
        });

        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), "Du_Lieu_Da_Xu_Ly.xlsx");
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
        let rowsToDelete = new Set<number>();

        for (let r = 0; r < sheetData.length; r++) {
          // A. Nếu dòng này nằm trong vùng bôi chọn mẫu ban đầu -> TUYỆT ĐỐI GIỮ NGUYÊN (Lá chắn bảo vệ 1)
          if (r >= r1 && r <= r2) {
            continue;
          }

          if (matchHeader(sheetData, r, pattern, c1, c2)) {
            // CHỈ xóa nếu vị trí r nằm DƯỚI dòng Header mẫu người dùng đã bôi đen (r > r2).
            // Giữ nguyên toàn bộ các Header ở vị trí r <= r2 (bao gồm cả Header ở đầu trang r = 0).
            if (r > r2) {
              let lastRowTemplate = pattern[pattern.length - 1]; // Dòng cuối cùng của mẫu (A | B | C...)
              
              // 1. Xóa các dòng của Header trùng lặp bằng cách đối chiếu thông minh
              for (let i = 0; i < pattern.length; i++) {
                let targetRow = r + i;
                if (targetRow >= sheetData.length) break;

                // Kiểm tra xem dòng targetRow có khớp với BẤT KỲ dòng nào trong mẫu không
                // Nếu không khớp (tức là đã chạm vào dữ liệu sản phẩm hoặc vùng chữ ký), ta dừng xóa ngay lập tức
                let matchesAny = false;
                let targetRowText = getCleanedRangeText(sheetData[targetRow], c1, c2);
                
                if (targetRowText === "") {
                  matchesAny = true; // Luôn cho phép dòng trống làm spacer trong header
                } else {
                  for (let j = 0; j < pattern.length; j++) {
                    if (matchSingleRow(sheetData[targetRow], pattern[j], c1, c2)) {
                      matchesAny = true;
                      break;
                    }
                  }
                }

                if (!matchesAny) {
                  break; // Chạm vào dữ liệu thực tế hoặc chữ ký -> Dừng khẩn cấp!
                }

                // Lá chắn bảo vệ 2: Chống xóa nhầm dòng mẫu
                if (!(targetRow >= r1 && targetRow <= r2)) {
                  rowsToDelete.add(targetRow);
                }

                // Nếu dòng targetRow khớp với dòng cuối cùng của mẫu (A | B | C | 1 | 2 | 3=1x2), 
                // ta dừng lại ngay lập tức (không xóa các dòng sản phẩm ở dưới nữa)
                if (matchSingleRow(sheetData[targetRow], lastRowTemplate, c1, c2)) {
                  break;
                }
              }
              
              // 2. +1 Xóa thêm 1 dòng ngay phía trên dòng Header trùng hợp (chính là dòng Footer của trang trước)
              let rowAbove = r - 1;
              // Lá chắn bảo vệ 3: Chống xóa nhầm dòng mẫu
              if (rowAbove >= 0 && !(rowAbove >= r1 && rowAbove <= r2)) {
                rowsToDelete.add(rowAbove);
              }
            }
            r += pattern.length - 1;
            continue;
          }
        }

        if (rowsToDelete.size === 0) {
          alert("Không tìm thấy hàng nào khớp với mẫu Header bôi chọn để cắt gọt!");
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

        let rowNumbers = sortedRows.map(r => r + 1).reverse().join(', '); // Hiển thị tăng dần
        setStatusText(`✅ Đã cắt gọt thành công! Đã xóa và dồn ${rowsToDelete.size} hàng.`);
        setPreviewHTML(`<span style="color:green; font-weight:bold;">Đã cắt gọt thành công!</span><br>Đã loại bỏ các Header lặp lại và các dòng Footer nằm liền kề phía trên.<br/><br/><b>Danh sách các hàng đã bị cắt gọt:</b><br/><div style="color:#d35400; font-family:monospace; word-break:break-all; max-height:80px; overflow-y:auto; margin-top:5px; padding:5px; background:#fff; border:1px solid #ddd; border-radius:4px; font-weight:bold;">Hàng: ${rowNumbers}</div>`);

      } catch (err: any) {
        alert("Có lỗi xảy ra khi xử lý dữ liệu: " + err.message);
      } finally {
        setLoadingMsg('');
      }
    }, 200);
  };

  return (
    <div className="h-screen w-screen overflow-hidden flex flex-col relative">
      
      {/* 1. Header Bar */}
      <div className="bg-[#107c41] px-5 py-3 text-white flex items-center gap-4 z-10">
        <strong className="text-lg">Trình xem Excel Online</strong>
        
        <input 
          type="file" 
          accept=".xlsx" 
          onChange={handleFileUpload}
          className="bg-white text-black p-1 rounded cursor-pointer text-sm" 
        />
        
        {hasFile && (
          <>
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
              onClick={handleExport}
              className="bg-[#d83b01] hover:bg-[#a82e00] border border-white px-4 py-2 rounded font-bold transition-colors"
            >
              💾 Tải Xuống Excel
            </button>
          </>
        )}
        <span className="text-sm italic">{statusText}</span>
      </div>

      {/* 2. Container cho Luckysheet */}
      <div id="luckysheet-container" className="flex-1 w-full relative"></div>

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
              1. Bạn hãy quét bôi chọn (bôi đen) vùng tiêu đề cần làm mẫu trực tiếp trên bảng tính Excel. <br />
              2. Nhấp nút <b>"Xử lý Toàn bộ"</b> ở bên dưới để hệ thống tự động tìm xóa tất cả các tiêu đề lặp lại và dồn dòng chân trang lên.
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