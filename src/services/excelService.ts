import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { Supplier, MasterPOItem, PO, NPLItem } from '../types';

export const parseSuppliers = async (file: File): Promise<Supplier[]> => {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
  
  const suppliers: Supplier[] = [];
  // Skip header row (index 0)
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[1]) { // Column B is Mã NCC
      suppliers.push({
        code: String(row[1] || '').trim(), // B
        name: String(row[2] || '').trim(), // C
        contactPerson: String(row[3] || '').trim(), // D
        taxCode: String(row[4] || '').trim(), // E
        phone: String(row[5] || '').trim(), // F
        address: String(row[6] || '').trim(), // G
        email: String(row[7] || '').trim(), // H
        deposit: String(row[8] || '').trim(), // I
        bankAccountName: String(row[9] || '').trim(), // J
        bankAccountNumber: String(row[10] || '').trim(), // K
        bankName: String(row[11] || '').trim(), // L
      });
    }
  }
  return suppliers;
};

export const parseMasterPO = async (file: File): Promise<MasterPOItem[]> => {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
  
  const parseNumber = (val: any): number => {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    // Remove commas or spaces that might cause NaN during string-to-number conversion
    const cleaned = String(val).replace(/,/g, '').replace(/\s/g, '').trim();
    const num = Number(cleaned);
    return isNaN(num) ? 0 : num;
  };

  const items: MasterPOItem[] = [];
  // Skip header row (index 0)
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[7]) { // Column H is Mã NCC
      items.push({
        supplierCode: String(row[7] || '').trim(),
        productName: String(row[1] || '').trim(), // Cột B (Mã Hàng)
        unit: String(row[2] || '').trim(),        // Cột C (Tên Hàng)
        quantity: parseNumber(row[3]),            // Cột D (Số lượng)
        unitPrice: parseNumber(row[4]),           // Cột E (Đơn giá)
      });
    }
  }
  return items;
};

export const parseNPL = async (file: File): Promise<NPLItem[]> => {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];

  const parseNorm = (val: any): number => {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    const cleaned = String(val).replace(/,/g, '.').trim();
    const num = Number(cleaned);
    return isNaN(num) ? 0 : num;
  };

  const items: NPLItem[] = [];
  let currentProductCode = '';
  let currentProductName = '';

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[1]) {
      currentProductCode = String(row[1]).trim();
      currentProductName = String(row[2] || '').trim();
    }
    if (!currentProductCode || !row[3]) continue;
    items.push({
      productCode: currentProductCode,
      productName: currentProductName,
      materialCode: String(row[3] || '').trim(),
      materialName: String(row[4] || '').trim(),
      norm: parseNorm(row[5]),
      unit: String(row[6] || '').trim(),
    });
  }
  return items;
};

export interface NPLDetailRow {
  productCode: string;
  productName: string;
  poQuantity: number;
  materialCode: string;
  materialName: string;
  norm: number;
  quantity: number;
  unit: string;
}

export interface NPLAggRow {
  materialCode: string;
  materialName: string;
  totalQuantity: number;
  unit: string;
}

export const exportNPLExcel = async (
  poNumber: string,
  supplierName: string,
  details: NPLDetailRow[],
  aggregated: NPLAggRow[]
) => {
  const wb = new ExcelJS.Workbook();

  const headerFill: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F81BD' } };
  const headerFont: Partial<ExcelJS.Font> = { bold: true, name: 'Times New Roman', size: 11, color: { argb: 'FFFFFFFF' } };
  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin' }, left: { style: 'thin' },
    bottom: { style: 'thin' }, right: { style: 'thin' },
  };
  const applyHeader = (row: ExcelJS.Row, values: string[]) => {
    values.forEach((v, i) => {
      const cell = row.getCell(i + 1);
      cell.value = v;
      cell.font = headerFont;
      cell.fill = headerFill;
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = thinBorder;
    });
    row.height = 30;
  };

  wb.creator = supplierName;

  // Sheet 1: Chi tiết
  const ws1 = wb.addWorksheet('Chi tiết NPL');
  ws1.columns = [
    { width: 6 }, { width: 14 }, { width: 30 }, { width: 8 },
    { width: 16 }, { width: 35 }, { width: 12 }, { width: 14 }, { width: 10 },
  ];
  applyHeader(ws1.getRow(1), ['STT', 'Mã SP', 'Tên SP', 'SL PO', 'Mã NPL', 'Tên NPL', 'Định Mức', 'SL Cần', 'Đơn Vị']);
  details.forEach((d, idx) => {
    const row = ws1.getRow(idx + 2);
    row.values = [idx + 1, d.productCode, d.productName, d.poQuantity, d.materialCode, d.materialName, d.norm, d.quantity || '', d.unit];
    for (let c = 1; c <= 9; c++) {
      const cell = row.getCell(c);
      cell.font = { name: 'Times New Roman', size: 11 };
      cell.border = thinBorder;
      if ([1, 4, 7, 8].includes(c)) cell.alignment = { horizontal: 'center' };
    }
    row.height = 22;
  });

  // Sheet 2: Tổng hợp
  const ws2 = wb.addWorksheet('Tổng hợp NPL');
  ws2.columns = [{ width: 6 }, { width: 18 }, { width: 40 }, { width: 16 }, { width: 12 }];
  applyHeader(ws2.getRow(1), ['STT', 'Mã NPL', 'Tên NPL', 'Tổng SL', 'Đơn Vị']);
  aggregated.forEach((a, idx) => {
    const row = ws2.getRow(idx + 2);
    row.values = [idx + 1, a.materialCode, a.materialName, a.totalQuantity, a.unit];
    for (let c = 1; c <= 5; c++) {
      const cell = row.getCell(c);
      cell.font = { name: 'Times New Roman', size: 11 };
      cell.border = thinBorder;
      if ([1, 4].includes(c)) cell.alignment = { horizontal: 'center' };
    }
    row.height = 22;
  });

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `NPL_${poNumber}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.URL.revokeObjectURL(url);
};

const base64ToArrayBuffer = (base64: string) => {
  const binaryString = window.atob(base64.split(',')[1] || base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes.buffer;
};

export const generatePOExcel = async (po: PO, templateBase64: string | null) => {
  if (!templateBase64) {
    throw new Error("Không có file mẫu PO");
  }

  const buffer = base64ToArrayBuffer(templateBase64);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  
  const worksheet = workbook.worksheets[0];

  // Map Header Info (Keep existing or adjust if needed)
  worksheet.getCell('D6').value = po.poNumber;
  worksheet.getCell('D7').value = po.orderDate;
  worksheet.getCell('D8').value = po.contractExpiry;

  // Map Supplier Info
  worksheet.getCell('C9').value = po.supplier.name;
  worksheet.getCell('G9').value = po.supplier.contactPerson;
  worksheet.getCell('C10').value = po.supplier.taxCode || po.supplier.code;
  worksheet.getCell('E9').value = po.supplier.phone;
  worksheet.getCell('C11').value = po.supplier.address;

  // 1. Xác định vị trí và dọn dẹp không gian cho bảng
  let startRow = 15; // Default fallback
  for (let i = 1; i <= 50; i++) {
    const row = worksheet.getRow(i);
    const cell1 = String(row.getCell(1).value || '').toUpperCase().trim();
    const cell2 = String(row.getCell(2).value || '').toUpperCase().trim();
    const cell3 = String(row.getCell(3).value || '').toUpperCase().trim();
    
    if (cell1 === 'STT' || cell2 === 'MÃ SP' || cell3 === 'TÊN SẢN PHẨM') {
      startRow = i;
      break;
    }
  }
  
  // Tìm dòng bắt đầu của phần thông tin bên dưới bảng (ví dụ: Thông tin thanh toán, Tên TK Nhận...)
  let belowTableStartRow = startRow + 2; // Default fallback
  for (let i = startRow + 1; i <= startRow + 50; i++) {
    const row = worksheet.getRow(i);
    let found = false;
    row.eachCell((cell) => {
      const val = String(cell.value || '').toUpperCase();
      if (val.includes('ĐIỀU KHOẢN THANH TOÁN') ||
          val.includes('THÔNG TIN THANH TOÁN') || 
          val.includes('TÊN TK NHẬN') || 
          val.includes('TÊN TÀI KHOẢN') ||
          val.includes('NGƯỜI ĐỀ NGHỊ') ||
          val.includes('GIÁM ĐỐC')) {
        found = true;
      }
    });
    if (found) {
      belowTableStartRow = i;
      break;
    }
  }

  const itemsCount = po.items.length;
  const rowsNeeded = 1 + itemsCount + 3; // 1 Header + N Items + 3 Footer
  const rowsAvailable = belowTableStartRow - startRow;

  // Chèn hoặc xóa dòng để đẩy phần bên dưới lên/xuống cho vừa khít
  if (rowsNeeded > rowsAvailable) {
    worksheet.spliceRows(belowTableStartRow, 0, ...Array(rowsNeeded - rowsAvailable).fill([]));
  } else if (rowsNeeded < rowsAvailable) {
    worksheet.spliceRows(belowTableStartRow - (rowsAvailable - rowsNeeded), rowsAvailable - rowsNeeded);
  }

  // Xóa sạch dữ liệu và style cũ trong vùng chuẩn bị render table
  for (let i = startRow; i < startRow + rowsNeeded; i++) {
    const row = worksheet.getRow(i);
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.value = null;
      cell.style = {};
    });
  }

  // 2. Render Header
  const headerRow = worksheet.getRow(startRow);
  const headers = ['STT', 'Mã SP', 'Tên Sản Phẩm', 'SL Đặt', 'Đơn Giá (Chưa VAT)', 'VAT', 'Thành Tiền', 'Ghi Chú'];
  headers.forEach((text, index) => {
    const cell = headerRow.getCell(index + 1);
    cell.value = text;
    cell.font = { bold: true, name: 'Times New Roman', size: 11, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F81BD' } }; // Xanh dương đậm
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' }
    };
  });
  headerRow.height = 30;

  // 3. Render Danh sách sản phẩm
  let totalBeforeTax = 0;
  let totalQuantity = 0;
  let totalVatAmount = 0;

  po.items.forEach((item, index) => {
    const rowNumber = startRow + 1 + index;
    const row = worksheet.getRow(rowNumber);
    
    const amount = item.quantity * item.unitPrice;
    const vatAmount = amount * (po.vatRate || 0) / 100;
    const totalItemAmount = amount + vatAmount;
    
    totalBeforeTax += amount;
    totalQuantity += item.quantity;
    totalVatAmount += vatAmount;
    
    row.getCell(1).value = index + 1;
    row.getCell(2).value = item.productName;
    row.getCell(3).value = item.unit;
    row.getCell(4).value = item.quantity;
    row.getCell(5).value = item.unitPrice;
    row.getCell(6).value = vatAmount;
    row.getCell(7).value = totalItemAmount;
    row.getCell(8).value = ''; // Ghi chú
    
    for (let col = 1; col <= 8; col++) {
      const cell = row.getCell(col);
      cell.font = { name: 'Times New Roman', size: 11 };
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
      
      if (col === 1 || col === 4) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      } else if (col === 2 || col === 3 || col === 8) {
        cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      } else if (col === 5 || col === 6 || col === 7) {
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
        cell.numFmt = '#,##0';
      }
    }
    row.height = 25;
    row.commit();
  });

  // 4. Tính toán và Render Footer
  const totalAfterTax = totalBeforeTax + totalVatAmount;
  
  let depositPercent = 0;
  if (po.supplier.deposit) {
    const depositStr = po.supplier.deposit;
    if (depositStr.includes('%')) {
      const match = depositStr.match(/(\d+(\.\d+)?)/);
      if (match) depositPercent = parseFloat(match[1]);
    } else {
      const val = parseFloat(depositStr);
      if (!isNaN(val)) depositPercent = (val <= 1 && val > 0) ? val * 100 : val;
    }
  }
  const depositAmount = (totalAfterTax * depositPercent) / 100;
  const remainingAmount = totalAfterTax - depositAmount;

  const footerStartRow = startRow + 1 + itemsCount;
  
  const depositText = po.supplier.deposit || '=DATA NCC';
  const remainingText = po.supplier.deposit ? `${100 - depositPercent}%-CỌC` : '=100%-CỌC';

  const footers = [
    { label: 'TỔNG CỘNG', type: 'total' },
    { label: 'CỌC', type: 'deposit' },
    { label: 'CHỐT PO', type: 'remaining' }
  ];

  footers.forEach((f, index) => {
    const rowNumber = footerStartRow + index;
    const row = worksheet.getRow(rowNumber);
    
    // Merge 3 ô đầu tiên cho Label
    worksheet.mergeCells(`A${rowNumber}:C${rowNumber}`);
    const labelCell = row.getCell(1);
    labelCell.value = f.label;
    
    if (f.type === 'total') {
      row.getCell(4).value = totalQuantity;
      worksheet.mergeCells(`E${rowNumber}:F${rowNumber}`);
      row.getCell(7).value = totalAfterTax;
    } else if (f.type === 'deposit') {
      worksheet.mergeCells(`D${rowNumber}:F${rowNumber}`);
      row.getCell(4).value = depositText;
      row.getCell(7).value = po.supplier.deposit ? depositAmount : '';
    } else if (f.type === 'remaining') {
      worksheet.mergeCells(`D${rowNumber}:F${rowNumber}`);
      row.getCell(4).value = remainingText;
      row.getCell(7).value = po.supplier.deposit ? remainingAmount : '';
    }
    
    for (let col = 1; col <= 8; col++) {
      const cell = row.getCell(col);
      cell.font = { bold: true, name: 'Times New Roman', size: 11 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } }; // Xám nhạt
      
      // Chỉ kẻ viền cho các ô không bị merge (hoặc ô đầu tiên của merge)
      // ExcelJS xử lý viền của merged cells hơi đặc biệt, ta cứ set border cho tất cả
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
      
      if (col === 1) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      } else if (col === 4) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        if (f.type === 'total') cell.numFmt = '#,##0';
      } else if (col === 7) {
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
        if (cell.value !== '') cell.numFmt = '#,##0';
      }
    }
    row.height = 25;
    row.commit();
  });

  // 5. Điền thông tin thanh toán (Bank Info) vào cột E tương ứng với các label
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > startRow) {
      let hasTenTK = false;
      let hasSTK = false;
      let hasNganHang = false;

      row.eachCell((cell) => {
        const val = String(cell.value || '').toUpperCase();
        if (val.includes('TÊN TK NHẬN') || val.includes('TÊN TÀI KHOẢN')) hasTenTK = true;
        if (val.includes('STK') || val.includes('SỐ TÀI KHOẢN')) hasSTK = true;
        if (val.includes('NGÂN HÀNG')) hasNganHang = true;
      });

      if (hasTenTK) {
        row.getCell(5).value = po.supplier.bankAccountName || '';
        row.getCell(5).font = { name: 'Times New Roman', size: 11, bold: true };
      }
      if (hasSTK) {
        row.getCell(5).value = po.supplier.bankAccountNumber || '';
        row.getCell(5).font = { name: 'Times New Roman', size: 11, bold: true };
      }
      if (hasNganHang) {
        row.getCell(5).value = po.supplier.bankName || '';
        row.getCell(5).font = { name: 'Times New Roman', size: 11, bold: true };
      }
    }
  });

  // Export using the requested method to avoid corruption
  const outBuffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `PO_${po.poNumber}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.URL.revokeObjectURL(url);
};