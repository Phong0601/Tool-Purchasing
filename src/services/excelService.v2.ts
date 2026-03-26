import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { Supplier, MasterPOItem, PO } from '../types';

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
    const cleaned = String(val).replace(/,/g, '').replace(/\\s/g, '').trim();
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

const base64ToArrayBuffer = (base64: string) => {
  const binaryString = window.atob(base64.split(',')[1] || base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes.buffer;
};

const getCellText = (cell: ExcelJS.Cell): string => {
  if (cell.value && typeof cell.value === 'object' && 'richText' in cell.value) {
    return (cell.value as any).richText.map((rt: any) => rt.text).join('');
  }
  return String(cell.value || '');
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
  const startRow = 15;
  
  // Tìm dòng bắt đầu của phần thông tin bên dưới bảng (ví dụ: 1 ĐIỀU KHOẢN THANH TOÁN:)
  let belowTableStartRow = 24; // Default fallback
  for (let i = startRow + 1; i <= startRow + 50; i++) {
    const row = worksheet.getRow(i);
    const cellB = row.getCell(2); // Cột B
    const val = getCellText(cellB).toUpperCase();
    if (val.includes('ĐIỀU KHOẢN THANH TOÁN:')) {
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
    
    totalBeforeTax += amount;
    totalQuantity += item.quantity;
    totalVatAmount += vatAmount;
    
    row.getCell(1).value = index + 1;
    row.getCell(2).value = item.productName;
    row.getCell(3).value = item.unit;
    row.getCell(4).value = item.quantity;
    row.getCell(5).value = item.unitPrice;
    row.getCell(6).value = `${po.vatRate || 0}%`;
    row.getCell(7).value = amount;
    row.getCell(8).value = ''; // Ghi chú
    
    for (let col = 1; col <= 8; col++) {
      const cell = row.getCell(col);
      cell.font = { name: 'Times New Roman', size: 11 };
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' },
        bottom: { style: 'thin' }, right: { style: 'thin' }
      };
      
      if (col === 1 || col === 4 || col === 6) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      } else if (col === 2 || col === 3 || col === 8) {
        cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      } else if (col === 5 || col === 7) {
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
      const match = depositStr.match(/(\\d+(\\.\\d+)?)/);
      if (match) depositPercent = parseFloat(match[1]);
    } else {
      const val = parseFloat(depositStr);
      if (!isNaN(val)) depositPercent = (val <= 1 && val > 0) ? val * 100 : val;
    }
  }
  const depositAmount = (totalAfterTax * depositPercent) / 100;
  const remainingAmount = totalAfterTax - depositAmount;

  const footerStartRow = startRow + 1 + itemsCount;
  const footers = [
    { label: 'Tổng cộng', qty: totalQuantity, amount: totalAfterTax },
    { label: 'Cọc', qty: '', amount: depositAmount },
    { label: 'Chốt PO', qty: '', amount: remainingAmount }
  ];

  footers.forEach((f, index) => {
    const rowNumber = footerStartRow + index;
    const row = worksheet.getRow(rowNumber);
    
    // Merge 3 ô đầu tiên cho Label
    worksheet.mergeCells(`A${rowNumber}:C${rowNumber}`);
    const labelCell = row.getCell(1);
    labelCell.value = f.label;
    
    row.getCell(4).value = f.qty;
    row.getCell(7).value = f.amount;
    
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
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
      } else if (col === 4) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        if (f.qty !== '') cell.numFmt = '#,##0';
      } else if (col === 7) {
        cell.alignment = { vertical: 'middle', horizontal: 'right' };
        cell.numFmt = '#,##0';
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
        const val = getCellText(cell).toUpperCase();
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
