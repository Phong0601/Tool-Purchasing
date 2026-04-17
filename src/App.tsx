import React, { useState, useEffect } from 'react';
import {
  Upload, FileSpreadsheet, Printer, Download, Settings, FileText,
  CheckCircle2, AlertCircle, Trash2, X, Eye, Package, ClipboardList,
} from 'lucide-react';
import { Supplier, MasterPOItem, PO, NPLItem } from './types';
import { parseSuppliers, parseMasterPO, generatePOExcel, parseNPL, exportNPLExcel, NPLDetailRow, NPLAggRow } from './services/excelService';
import { saveSupplierData, loadSupplierData, saveTemplateFile, loadTemplateFile, saveNPLData, loadNPLData } from './services/storageService';

const formatDate = (date: Date) => {
  const d = String(date.getDate()).padStart(2, '0');
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const y = date.getFullYear();
  return `${d}/${m}/${y}`;
};

const formatShortDate = (date: Date) => {
  const d = String(date.getDate()).padStart(2, '0');
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const y = String(date.getFullYear()).slice(-2);
  return `${d}${m}${y}`;
};

// ── NPL computed types ───────────────────────────────────────────────────────

interface NPLMaterial {
  materialCode: string;
  materialName: string;
  norm: number;
  quantity: number;
  unit: string;
}

interface NPLProductDetail {
  productCode: string;
  productName: string;
  poQuantity: number;
  materials: NPLMaterial[];
}

function computeNPLForPO(
  po: PO,
  nplData: NPLItem[]
): { products: NPLProductDetail[]; aggregated: NPLAggRow[] } {
  const products: NPLProductDetail[] = po.items.map(item => {
    const matching = nplData.filter(
      n => n.productCode.toLowerCase() === item.productName.toLowerCase()
    );
    return {
      productCode: item.productName, // MasterPOItem.productName stores Mã Hàng
      productName: item.unit,        // MasterPOItem.unit stores Tên Hàng
      poQuantity: item.quantity,
      materials: matching.map(n => ({
        materialCode: n.materialCode,
        materialName: n.materialName,
        norm: n.norm,
        quantity: item.quantity * n.norm,
        unit: n.unit,
      })),
    };
  });

  const aggMap = new Map<string, NPLAggRow>();
  for (const product of products) {
    for (const mat of product.materials) {
      const existing = aggMap.get(mat.materialCode);
      if (existing) {
        existing.totalQuantity += mat.quantity;
      } else {
        aggMap.set(mat.materialCode, {
          materialCode: mat.materialCode,
          materialName: mat.materialName,
          totalQuantity: mat.quantity,
          unit: mat.unit,
        });
      }
    }
  }

  return {
    products,
    aggregated: Array.from(aggMap.values()).sort((a, b) =>
      a.materialCode.localeCompare(b.materialCode)
    ),
  };
}

// ── App ──────────────────────────────────────────────────────────────────────

export default function App() {
  const [suppliers, setSuppliers] = useState<Supplier[]>([]);
  const [hasTemplate, setHasTemplate] = useState(false);
  const [pos, setPos] = useState<PO[]>([]);
  const [activeTab, setActiveTab] = useState<'process' | 'settings'>('process');
  const [toast, setToast] = useState<{ message: string; type: 'success' | 'error' | 'warning' | 'info' } | null>(null);
  const [showConfirmDelete, setShowConfirmDelete] = useState(false);
  const [previewPO, setPreviewPO] = useState<PO | null>(null);

  // NPL state
  const [nplData, setNplData] = useState<NPLItem[]>([]);
  const [nplModal, setNplModal] = useState<PO | null>(null);
  const [nplModalTab, setNplModalTab] = useState<'detail' | 'aggregate'>('detail');

  const showToast = (message: string, type: 'success' | 'error' | 'warning' | 'info' = 'info') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 5000);
  };

  useEffect(() => {
    const loadedSuppliers = loadSupplierData();
    if (loadedSuppliers.length > 0) setSuppliers(loadedSuppliers);
    const template = loadTemplateFile();
    if (template) setHasTemplate(true);
    const loadedNPL = loadNPLData();
    if (loadedNPL.length > 0) setNplData(loadedNPL);
  }, []);

  // ── Handlers ────────────────────────────────────────────────────────────────

  const handleSupplierUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const parsed = await parseSuppliers(file);
      setSuppliers(parsed);
      saveSupplierData(parsed);
      showToast('Tải lên dữ liệu Nhà Cung Cấp thành công!', 'success');
    } catch (error) {
      console.error(error);
      showToast('Lỗi khi đọc file DATA NCC.', 'error');
    }
  };

  const handleTemplateUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const base64 = event.target?.result as string;
      saveTemplateFile(base64);
      setHasTemplate(true);
      showToast('Tải lên File mẫu PO thành công!', 'success');
    };
    reader.readAsDataURL(file);
  };

  const handleNPLUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const parsed = await parseNPL(file);
      setNplData(parsed);
      saveNPLData(parsed);
      showToast(`Tải lên ${parsed.length} dòng NPL thành công!`, 'success');
    } catch (error) {
      console.error(error);
      showToast('Lỗi khi đọc file DATA NPL.', 'error');
    }
    e.target.value = '';
  };

  const clearSettings = () => setShowConfirmDelete(true);

  const confirmClearSettings = () => {
    localStorage.removeItem('procurement_suppliers');
    localStorage.removeItem('procurement_template');
    localStorage.removeItem('procurement_npl');
    setSuppliers([]);
    setHasTemplate(false);
    setNplData([]);
    setShowConfirmDelete(false);
    showToast('Đã xóa dữ liệu cài đặt.', 'success');
  };

  const handleMasterPOUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (suppliers.length === 0) {
      showToast('Vui lòng tải lên DATA NCC trước khi xử lý PO.', 'warning');
      e.target.value = '';
      return;
    }

    try {
      const items = await parseMasterPO(file);

      if (items.length === 0) {
        showToast('Không tìm thấy dữ liệu hợp lệ trong file 1.PO Tổng.', 'error');
        e.target.value = '';
        return;
      }

      const grouped = items.reduce((acc, item) => {
        if (!acc[item.supplierCode]) acc[item.supplierCode] = [];
        acc[item.supplierCode].push(item);
        return acc;
      }, {} as Record<string, MasterPOItem[]>);

      const currentDate = new Date();
      const expiryDate = new Date();
      expiryDate.setDate(currentDate.getDate() + 60);

      const orderDateStr = formatDate(currentDate);
      const expiryDateStr = formatDate(expiryDate);
      const shortCurrentDate = formatShortDate(currentDate);
      const shortExpiryDate = formatShortDate(expiryDate);

      const generatedPOs: PO[] = [];
      const missingSuppliers: string[] = [];

      for (const [supplierCode, supplierItems] of Object.entries(grouped)) {
        const supplier = suppliers.find(s => s.code.toLowerCase() === supplierCode.toLowerCase());
        if (!supplier) {
          missingSuppliers.push(supplierCode);
          continue;
        }
        const defaultSuffix = '01';
        const poNumber = `${supplier.code}${shortCurrentDate}${shortExpiryDate}${defaultSuffix}`;
        generatedPOs.push({
          supplier,
          items: supplierItems,
          poNumber,
          orderDate: orderDateStr,
          contractExpiry: expiryDateStr,
          vatRate: 8,
          poSuffix: defaultSuffix,
          shortCurrentDate,
          shortExpiryDate,
        });
      }

      setPos(generatedPOs);

      if (missingSuppliers.length > 0) {
        showToast(`Cảnh báo: Không tìm thấy NCC: ${missingSuppliers.join(', ')}`, 'warning');
      } else if (generatedPOs.length === 0) {
        showToast('Không tạo được PO nào. Kiểm tra lại mã NCC.', 'error');
      } else {
        showToast(`Đã tạo thành công ${generatedPOs.length} PO.`, 'success');
      }
    } catch (error) {
      console.error(error);
      showToast('Lỗi khi xử lý file 1.PO Tổng.', 'error');
    } finally {
      e.target.value = '';
    }
  };

  const handleExportExcel = async (po: PO) => {
    const templateBase64 = loadTemplateFile();
    if (!templateBase64) {
      showToast('Chưa có File mẫu PO. Vui lòng tải lên trong phần Cài đặt.', 'warning');
      return;
    }
    try {
      await generatePOExcel(po, templateBase64);
    } catch (error) {
      console.error(error);
      showToast('Lỗi khi xuất file Excel.', 'error');
    }
  };

  const handlePrint = () => window.print();

  const updatePreviewPO = (updates: Partial<PO>) => {
    if (!previewPO) return;
    const updatedPO = { ...previewPO, ...updates };
    if (updates.poSuffix !== undefined) {
      updatedPO.poNumber = `${updatedPO.supplier.code}${updatedPO.shortCurrentDate}${updatedPO.shortExpiryDate}${updatedPO.poSuffix}`;
    }
    setPreviewPO(updatedPO);
    setPos(current => current.map(p => p.supplier.code === updatedPO.supplier.code ? updatedPO : p));
  };

  // ── NPL handlers ─────────────────────────────────────────────────────────

  const handleNormEdit = (productCode: string, materialCode: string, newNorm: number) => {
    const value = isNaN(newNorm) ? 0 : newNorm;
    const updated = nplData.map(item =>
      item.productCode.toLowerCase() === productCode.toLowerCase() && item.materialCode === materialCode
        ? { ...item, norm: value }
        : item
    );
    setNplData(updated);
    saveNPLData(updated);
  };

  const handleExportNPL = async () => {
    if (!nplModal) return;
    const { products, aggregated } = computeNPLForPO(nplModal, nplData);
    const details: NPLDetailRow[] = products.flatMap(p =>
      p.materials.length > 0
        ? p.materials.map(m => ({ productCode: p.productCode, productName: p.productName, poQuantity: p.poQuantity, ...m }))
        : [{ productCode: p.productCode, productName: p.productName, poQuantity: p.poQuantity, materialCode: '', materialName: '(Chưa có NPL)', norm: 0, quantity: 0, unit: '' }]
    );
    try {
      await exportNPLExcel(nplModal.poNumber, nplModal.supplier.name, details, aggregated);
    } catch (error) {
      console.error(error);
      showToast('Lỗi khi xuất file NPL Excel.', 'error');
    }
  };

  const openNPLModal = (po: PO) => {
    setNplModalTab('detail');
    setNplModal(po);
  };

  // ── Render ────────────────────────────────────────────────────────────────

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans print:bg-white">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10 print:hidden">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2 text-indigo-600">
            <FileSpreadsheet className="w-6 h-6" />
            <h1 className="font-bold text-xl tracking-tight">Tách PO</h1>
          </div>
          <nav className="flex gap-1">
            <button
              onClick={() => setActiveTab('process')}
              className={`px-4 py-2 rounded-md text-sm font-medium transition-colors ${activeTab === 'process' ? 'bg-indigo-50 text-indigo-700' : 'text-slate-600 hover:bg-slate-100'}`}
            >
              Xử lý PO
            </button>
            <button
              onClick={() => setActiveTab('settings')}
              className={`px-4 py-2 rounded-md text-sm font-medium transition-colors flex items-center gap-2 ${activeTab === 'settings' ? 'bg-indigo-50 text-indigo-700' : 'text-slate-600 hover:bg-slate-100'}`}
            >
              <Settings className="w-4 h-4" />
              Cài đặt
            </button>
          </nav>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 print:hidden">

        {/* ── Settings Tab ────────────────────────────────────────────────── */}
        {activeTab === 'settings' && (
          <div className="space-y-6">
            <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
              <h2 className="text-lg font-semibold mb-4">Quản lý Dữ liệu Nền (Local Storage)</h2>

              <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-6">
                {/* Supplier Data */}
                <div className="border-2 border-dashed border-slate-300 rounded-lg p-6 flex flex-col items-center justify-center text-center hover:bg-slate-50 transition-colors relative">
                  <input type="file" accept=".xlsx,.xls" onChange={handleSupplierUpload} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" />
                  <div className={`w-12 h-12 rounded-full flex items-center justify-center mb-3 ${suppliers.length > 0 ? 'bg-green-100 text-green-600' : 'bg-indigo-100 text-indigo-600'}`}>
                    {suppliers.length > 0 ? <CheckCircle2 className="w-6 h-6" /> : <Upload className="w-6 h-6" />}
                  </div>
                  <h3 className="font-medium text-slate-900">1. Tải lên DATA NCC</h3>
                  <p className="text-sm text-slate-500 mt-1">File Excel chứa danh sách Nhà cung cấp</p>
                  {suppliers.length > 0 && (
                    <span className="mt-3 inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">Đã lưu {suppliers.length} NCC</span>
                  )}
                </div>

                {/* Template */}
                <div className="border-2 border-dashed border-slate-300 rounded-lg p-6 flex flex-col items-center justify-center text-center hover:bg-slate-50 transition-colors relative">
                  <input type="file" accept=".xlsx,.xls" onChange={handleTemplateUpload} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" />
                  <div className={`w-12 h-12 rounded-full flex items-center justify-center mb-3 ${hasTemplate ? 'bg-green-100 text-green-600' : 'bg-indigo-100 text-indigo-600'}`}>
                    {hasTemplate ? <CheckCircle2 className="w-6 h-6" /> : <FileText className="w-6 h-6" />}
                  </div>
                  <h3 className="font-medium text-slate-900">2. Tải lên File mẫu PO</h3>
                  <p className="text-sm text-slate-500 mt-1">Template Excel để xuất dữ liệu</p>
                  {hasTemplate && (
                    <span className="mt-3 inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">Đã lưu Template</span>
                  )}
                </div>

                {/* NPL Data */}
                <div className="border-2 border-dashed border-slate-300 rounded-lg p-6 flex flex-col items-center justify-center text-center hover:bg-slate-50 transition-colors relative">
                  <input type="file" accept=".xlsx,.xls" onChange={handleNPLUpload} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" />
                  <div className={`w-12 h-12 rounded-full flex items-center justify-center mb-3 ${nplData.length > 0 ? 'bg-green-100 text-green-600' : 'bg-emerald-100 text-emerald-600'}`}>
                    {nplData.length > 0 ? <CheckCircle2 className="w-6 h-6" /> : <Package className="w-6 h-6" />}
                  </div>
                  <h3 className="font-medium text-slate-900">3. Tải lên DATA NPL</h3>
                  <p className="text-sm text-slate-500 mt-1">File Excel định mức nguyên phụ liệu theo sản phẩm</p>
                  <p className="text-xs text-slate-400 mt-1">Cột: Mã SP · Tên SP · Mã NPL · Tên NPL · Định Mức · ĐVT</p>
                  {nplData.length > 0 && (
                    <span className="mt-3 inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-green-100 text-green-800">Đã lưu {nplData.length} dòng NPL</span>
                  )}
                </div>
              </div>

              <div className="mt-6 pt-6 border-t border-slate-200 flex justify-end">
                <button onClick={clearSettings} className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-red-600 hover:bg-red-50 rounded-md transition-colors">
                  <Trash2 className="w-4 h-4" />
                  Xóa dữ liệu cài đặt
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ── Process Tab ──────────────────────────────────────────────────── */}
        {activeTab === 'process' && (
          <div className="space-y-6">
            <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
              <div className="flex flex-col md:flex-row gap-6 items-start md:items-end">
                <div className="flex-1 w-full">
                  <div className="relative">
                    <input type="file" accept=".xlsx,.xls" onChange={handleMasterPOUpload} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer" />
                    <button className="w-full flex items-center justify-center gap-2 bg-indigo-600 text-white px-6 py-2.5 rounded-md font-medium hover:bg-indigo-700 transition-colors shadow-sm">
                      <Upload className="w-4 h-4" />
                      Tải lên 1.PO Tổng
                    </button>
                  </div>
                  <p className="mt-2 text-xs text-center text-slate-500">Hệ thống sẽ tự động tách và tạo PO.</p>
                </div>
              </div>

              {(!hasTemplate || suppliers.length === 0) && (
                <div className="mt-4 p-4 bg-amber-50 rounded-md flex items-start gap-3">
                  <AlertCircle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
                  <div className="text-sm text-amber-800">
                    <p className="font-medium">Chưa đủ dữ liệu nền!</p>
                    <p className="mt-1">Vui lòng qua tab <strong>Cài đặt</strong> để tải lên DATA NCC và File mẫu PO trước khi xử lý.</p>
                  </div>
                </div>
              )}
            </div>

            {pos.length > 0 && (
              <div className="space-y-4">
                <h2 className="text-lg font-semibold flex items-center gap-2">
                  Danh sách PO dự kiến
                  <span className="bg-indigo-100 text-indigo-700 py-0.5 px-2.5 rounded-full text-sm font-medium">{pos.length}</span>
                </h2>

                <div className="grid gap-4">
                  {pos.map((po, index) => {
                    const totalBeforeTax = po.items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0);
                    const totalAfterTax = totalBeforeTax * (1 + po.vatRate / 100);

                    return (
                      <div key={index} className="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col md:flex-row gap-6 items-start md:items-center justify-between transition-shadow hover:shadow-md">
                        <div className="flex-1">
                          <div className="flex items-center gap-3 mb-1">
                            <h3 className="font-bold text-lg text-slate-900">{po.poNumber}</h3>
                            <span className="text-sm text-slate-500 font-medium">{po.supplier.name}</span>
                          </div>
                          <div className="text-sm text-slate-600 flex flex-wrap gap-x-6 gap-y-1">
                            <span><strong>Ngày đặt:</strong> {po.orderDate}</span>
                            <span><strong>Hạn HĐ:</strong> {po.contractExpiry}</span>
                            <span><strong>Số mặt hàng:</strong> {po.items.length}</span>
                          </div>
                        </div>

                        <div className="flex items-center gap-4 bg-slate-50 p-3 rounded-lg border border-slate-100">
                          <div className="text-right">
                            <div className="text-xs font-medium text-slate-500 mb-1">Tổng thanh toán (VAT {po.vatRate}%)</div>
                            <div className="font-bold text-indigo-600">
                              {new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(totalAfterTax)}
                            </div>
                          </div>
                        </div>

                        <div className="flex items-center gap-2 w-full md:w-auto">
                          <button
                            onClick={() => openNPLModal(po)}
                            className="flex-1 md:flex-none flex items-center justify-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-md hover:bg-emerald-700 transition-colors text-sm font-medium shadow-sm"
                          >
                            <ClipboardList className="w-4 h-4" />
                            Xem NPL
                          </button>
                          <button
                            onClick={() => setPreviewPO(po)}
                            className="flex-1 md:flex-none flex items-center justify-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-md hover:bg-indigo-700 transition-colors text-sm font-medium shadow-sm"
                          >
                            <Eye className="w-4 h-4" />
                            Xem chi tiết
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}
      </main>

      {/* ── Toast ──────────────────────────────────────────────────────────── */}
      {toast && (
        <div className={`fixed bottom-4 right-4 z-50 p-4 rounded-md shadow-lg max-w-md flex items-start gap-3 ${
          toast.type === 'success' ? 'bg-green-50 text-green-800 border border-green-200' :
          toast.type === 'error' ? 'bg-red-50 text-red-800 border border-red-200' :
          toast.type === 'warning' ? 'bg-amber-50 text-amber-800 border border-amber-200' :
          'bg-blue-50 text-blue-800 border border-blue-200'
        }`}>
          <div className="flex-1 text-sm font-medium">{toast.message}</div>
          <button onClick={() => setToast(null)} className="opacity-50 hover:opacity-100 shrink-0"><X className="w-4 h-4" /></button>
        </div>
      )}

      {/* ── Confirm Delete Modal ────────────────────────────────────────────── */}
      {showConfirmDelete && (
        <div className="fixed inset-0 bg-slate-900/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-xl max-w-md w-full p-6">
            <h3 className="text-lg font-bold text-slate-900 mb-2">Xác nhận xóa dữ liệu</h3>
            <p className="text-slate-600 mb-6 text-sm">Bạn có chắc chắn muốn xóa toàn bộ dữ liệu cài đặt (DATA NCC, File mẫu PO & DATA NPL)? Hành động này không thể hoàn tác.</p>
            <div className="flex justify-end gap-3">
              <button onClick={() => setShowConfirmDelete(false)} className="px-4 py-2 text-sm font-medium text-slate-700 hover:bg-slate-100 rounded-md transition-colors">Hủy</button>
              <button onClick={confirmClearSettings} className="px-4 py-2 text-sm font-medium text-white bg-red-600 hover:bg-red-700 rounded-md transition-colors">Xóa dữ liệu</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Preview PO Modal ────────────────────────────────────────────────── */}
      {previewPO && (
        <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-0 sm:p-6 overflow-hidden print:static print:bg-white print:p-0">
          <div className="bg-white sm:rounded-xl shadow-2xl max-w-5xl w-full h-full max-h-[100vh] sm:max-h-[90vh] flex flex-col animate-in zoom-in-95 print:shadow-none print:max-h-none print:h-auto">
            <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between px-6 py-4 border-b border-slate-200 shrink-0 bg-slate-50 sm:rounded-t-xl gap-4 print:hidden">
              <div className="flex items-center gap-3">
                <FileText className="w-5 h-5 text-indigo-600" />
                <h3 className="text-lg font-bold text-slate-900">Chi tiết PO: {previewPO.poNumber}</h3>
              </div>

              <div className="flex items-center gap-4 bg-white px-4 py-2 rounded-lg border border-slate-200 shadow-sm w-full sm:w-auto">
                <div className="flex items-center gap-2">
                  <label className="text-sm font-medium text-slate-700 whitespace-nowrap">Hậu tố PO:</label>
                  <input
                    type="text" maxLength={2} value={previewPO.poSuffix}
                    onChange={(e) => updatePreviewPO({ poSuffix: e.target.value.toUpperCase() })}
                    className="w-16 px-2 py-1 border border-slate-300 rounded text-sm focus:ring-indigo-500 focus:border-indigo-500 uppercase text-center"
                    placeholder="01"
                  />
                </div>
                <div className="w-px h-6 bg-slate-200" />
                <div className="flex items-center gap-2">
                  <label className="text-sm font-medium text-slate-700 whitespace-nowrap">VAT (%):</label>
                  <input
                    type="number" value={previewPO.vatRate}
                    onChange={(e) => updatePreviewPO({ vatRate: Number(e.target.value) })}
                    className="w-16 px-2 py-1 border border-slate-300 rounded text-sm focus:ring-indigo-500 focus:border-indigo-500 text-center"
                  />
                </div>
              </div>

              <div className="flex items-center gap-2 w-full sm:w-auto justify-end">
                <button onClick={handlePrint} className="flex items-center gap-2 px-3 py-1.5 bg-white border border-slate-300 text-slate-700 rounded-md hover:bg-slate-50 transition-colors text-sm font-medium shadow-sm">
                  <Printer className="w-4 h-4" />In phiếu
                </button>
                <button onClick={() => handleExportExcel(previewPO)} className="flex items-center gap-2 px-3 py-1.5 bg-green-600 text-white rounded-md hover:bg-green-700 transition-colors text-sm font-medium shadow-sm">
                  <Download className="w-4 h-4" />Xuất Excel
                </button>
                <button onClick={() => setPreviewPO(null)} className="p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-200 rounded-full transition-colors ml-2">
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>

            {/* A4 Document Container */}
            <div className="flex-1 overflow-y-auto bg-slate-100 p-0 sm:p-8 print:bg-white print:p-0 print:overflow-visible">
              <div
                className="mx-auto bg-white w-full max-w-[210mm] min-h-[297mm] h-max p-8 sm:p-12 shadow-md border border-slate-300 text-[13px] text-black leading-relaxed print:shadow-none print:border-none print:p-0"
                style={{ fontFamily: '"Times New Roman", Times, serif' }}
              >

                {/* Company Header (Simulated) */}
                <div className="flex justify-between items-start mb-4 border-b-2 border-slate-800 pb-4">
                  <div className="flex gap-4">
                    <img src="/logo.png" alt="Leonardo Logo" className="w-16 h-16 object-contain" referrerPolicy="no-referrer" />
                    <div>
                      <h1 className="font-bold text-base">CÔNG TY TNHH LEONARDO</h1>
                      <p>Số 284 Pasteur, Phường Xuân Hòa, Thành phố Hồ Chí Minh</p>
                      <p>Mã số thuế: 0314465951</p>
                    </div>
                  </div>
                  <div className="text-right">
                    <p>{previewPO.poNumber}</p>
                    <p>{previewPO.orderDate}</p>
                    <p>{previewPO.contractExpiry}</p>
                  </div>
                </div>

                <div className="text-center mb-6">
                  <h2 className="text-xl font-bold uppercase">{previewPO.supplier.name}</h2>
                </div>

                <div className="grid grid-cols-12 gap-2 mb-6 text-sm">
                  <div className="col-span-3 font-bold">Số PO:</div><div className="col-span-9">{previewPO.poNumber}</div>
                  <div className="col-span-3 font-bold">Ngày đặt hàng:</div><div className="col-span-9">{previewPO.orderDate}</div>
                  <div className="col-span-3 font-bold">Hạn hợp đồng:</div><div className="col-span-9">{previewPO.contractExpiry}</div>
                  <div className="col-span-2 font-bold mt-2">Nhà Cung Cấp:</div>
                  <div className="col-span-4 mt-2">{previewPO.supplier.name}</div>
                  <div className="col-span-2 font-bold mt-2">Số điện thoại:</div>
                  <div className="col-span-2 mt-2">{previewPO.supplier.phone}</div>
                  <div className="col-span-2 mt-2">{previewPO.supplier.contactPerson}</div>
                  <div className="col-span-2 font-bold">MST:</div><div className="col-span-10">{previewPO.supplier.taxCode || previewPO.supplier.code}</div>
                  <div className="col-span-2 font-bold">Địa chỉ:</div><div className="col-span-10">{previewPO.supplier.address}</div>
                </div>

                <table className="w-full border-collapse border border-slate-800 mb-6 text-sm">
                  <thead>
                    <tr className="bg-slate-200 font-bold text-center">
                      <th className="border border-slate-800 p-2 w-12">STT</th>
                      <th className="border border-slate-800 p-2 w-24">Mã SP</th>
                      <th className="border border-slate-800 p-2">Tên Sản Phẩm</th>
                      <th className="border border-slate-800 p-2 w-20">SL Đặt</th>
                      <th className="border border-slate-800 p-2 w-28">Đơn Giá</th>
                      <th className="border border-slate-800 p-2 w-24">VAT</th>
                      <th className="border border-slate-800 p-2 w-32">Thành Tiền</th>
                      <th className="border border-slate-800 p-2 w-32">Ghi Chú</th>
                    </tr>
                  </thead>
                  <tbody>
                    {previewPO.items.map((item, idx) => {
                      const amount = item.quantity * item.unitPrice;
                      const vatAmount = amount * (previewPO.vatRate || 0) / 100;
                      const totalItemAmount = amount + vatAmount;
                      return (
                        <tr key={idx} className="even:bg-slate-50 hover:bg-slate-100 transition-colors">
                          <td className="border border-slate-800 p-2 text-center">{idx + 1}</td>
                          <td className="border border-slate-800 p-2 text-center">{item.productName}</td>
                          <td className="border border-slate-800 p-2">{item.unit}</td>
                          <td className="border border-slate-800 p-2 text-center">{item.quantity}</td>
                          <td className="border border-slate-800 p-2 text-right">
                            {new Intl.NumberFormat('vi-VN').format(item.unitPrice)}
                          </td>
                          <td className="border border-slate-800 p-2 text-right">
                            {new Intl.NumberFormat('vi-VN').format(vatAmount)}
                          </td>
                          <td className="border border-slate-800 p-2 text-right">
                            {new Intl.NumberFormat('vi-VN').format(totalItemAmount)}
                          </td>
                          <td className="border border-slate-800 p-2"></td>
                        </tr>
                      );
                    })}
                    {(() => {
                      const totalQuantity = previewPO.items.reduce((sum, item) => sum + item.quantity, 0);
                      const totalBeforeTax = previewPO.items.reduce((sum, item) => sum + (item.quantity * item.unitPrice), 0);
                      const totalVat = totalBeforeTax * (previewPO.vatRate || 0) / 100;
                      const totalAmount = totalBeforeTax + totalVat;
                      let depositPercent = 0;
                      if (previewPO.supplier.deposit) {
                        const depositStr = previewPO.supplier.deposit;
                        if (depositStr.includes('%')) {
                          const match = depositStr.match(/(\d+(\.\d+)?)/);
                          if (match) depositPercent = parseFloat(match[1]);
                        } else {
                          const val = parseFloat(depositStr);
                          if (!isNaN(val)) depositPercent = val <= 1 && val > 0 ? val * 100 : val;
                        }
                      }
                      const depositAmount = (totalAmount * depositPercent) / 100;
                      const remainingAmount = totalAmount - depositAmount;
                      return (
                        <>
                          <tr className="font-bold bg-slate-200">
                            <td colSpan={3} className="border border-slate-800 p-2 text-center">TỔNG CỘNG</td>
                            <td className="border border-slate-800 p-2 text-center">{new Intl.NumberFormat('vi-VN').format(totalQuantity)}</td>
                            <td colSpan={2} className="border border-slate-800 p-2" />
                            <td className="border border-slate-800 p-2 text-right">{new Intl.NumberFormat('vi-VN').format(totalAmount)}</td>
                            <td className="border border-slate-800 p-2" />
                          </tr>
                          <tr className="font-bold bg-slate-100">
                            <td colSpan={3} className="border border-slate-800 p-2 text-center">CỌC</td>
                            <td colSpan={3} className={`border border-slate-800 p-2 text-center ${previewPO.supplier.deposit ? '' : 'bg-yellow-200'}`}>{previewPO.supplier.deposit || '=DATA NCC'}</td>
                            <td className="border border-slate-800 p-2 text-right">{previewPO.supplier.deposit ? new Intl.NumberFormat('vi-VN').format(depositAmount) : ''}</td>
                            <td className="border border-slate-800 p-2" />
                          </tr>
                          <tr className="font-bold bg-slate-100">
                            <td colSpan={3} className="border border-slate-800 p-2 text-center">CHỐT PO</td>
                            <td colSpan={3} className={`border border-slate-800 p-2 text-center ${previewPO.supplier.deposit ? '' : 'bg-yellow-200'}`}>{previewPO.supplier.deposit ? `${100 - depositPercent}%-CỌC` : '=100%-CỌC'}</td>
                            <td className="border border-slate-800 p-2 text-right">{previewPO.supplier.deposit ? new Intl.NumberFormat('vi-VN').format(remainingAmount) : ''}</td>
                            <td className="border border-slate-800 p-2" />
                          </tr>
                        </>
                      );
                    })()}
                  </tbody>
                </table>

                <div className="text-sm">
                  <div className="font-bold mb-1">1 ĐIỀU KHOẢN THANH TOÁN:</div>
                  <div className="grid grid-cols-12 gap-2 mb-2">
                    <div className="col-span-3 italic">Phương thức thanh toán:</div><div className="col-span-9">Chuyển Khoản</div>
                    <div className="col-span-3 italic">Thông tin chuyển khoản:</div>
                    <div className="col-span-9">
                      <div>Tên TK Nhận: <span className={previewPO.supplier.bankAccountName ? '' : 'bg-yellow-200 px-8'}>{previewPO.supplier.bankAccountName}</span></div>
                      <div>STK: <span className={previewPO.supplier.bankAccountNumber ? '' : 'bg-yellow-200 px-8'}>{previewPO.supplier.bankAccountNumber}</span></div>
                      <div>Ngân hàng: <span className={previewPO.supplier.bankName ? '' : 'bg-yellow-200 px-8'}>{previewPO.supplier.bankName}</span></div>
                    </div>
                  </div>
                  <div className="font-bold mb-1">2 ĐIỀU KHOẢN GIAO HÀNG:</div>
                  <div className="grid grid-cols-12 gap-2 mb-2">
                    <div className="col-span-3">Địa điểm giao hàng:</div><div className="col-span-9">Hồ Chí Minh</div>
                    <div className="col-span-3">Người nhận hàng:</div>
                    <div className="col-span-9"><span className="bg-yellow-200 px-2">0703800855</span> Kho Leonardo</div>
                  </div>
                  <div className="mb-2">Hàng hoá được giao bắt buộc phải kèm chứng từ (Phiếu giao hàng).</div>
                  <div className="font-bold mb-1">3 ĐIỀU KHOẢN CHUNG:</div>
                  <div>Hàng hoá được mua bởi PO này được tuân thủ theo thoả thuận hợp đồng thoả thuận</div>
                  <div>PO này có giá trị pháp lý như một hợp đồng kinh tế sau khi NCC ký xác nhận và gửi lại bản Scan/Ảnh chụp qua Email chính thức.</div>
                </div>

                <div className="mt-32 text-center text-xs text-slate-500 italic border-t border-slate-200 pt-4 print:hidden">
                  Đây là bản xem trước mô phỏng cấu trúc File mẫu PO. Khi xuất Excel, hệ thống sẽ giữ nguyên định dạng của file gốc.
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ── NPL Modal ──────────────────────────────────────────────────────── */}
      {nplModal && (() => {
        const { products, aggregated } = computeNPLForPO(nplModal, nplData);
        const missingCount = products.filter(p => p.materials.length === 0).length;

        return (
          <div className="fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-0 sm:p-6 overflow-hidden">
            <div className="bg-white sm:rounded-xl shadow-2xl max-w-5xl w-full h-full max-h-[100vh] sm:max-h-[90vh] flex flex-col">

              {/* Modal header */}
              <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between px-6 py-4 border-b border-slate-200 shrink-0 bg-slate-50 sm:rounded-t-xl gap-3">
                <div className="flex items-center gap-3">
                  <ClipboardList className="w-5 h-5 text-emerald-600" />
                  <div>
                    <h3 className="text-lg font-bold text-slate-900">Nguyên Phụ Liệu</h3>
                    <p className="text-sm text-slate-500">{nplModal.supplier.name} — {nplModal.poNumber}</p>
                  </div>
                </div>

                {/* Tabs */}
                <div className="flex gap-1 bg-slate-200 rounded-lg p-1 shrink-0">
                  <button
                    onClick={() => setNplModalTab('detail')}
                    className={`px-3 py-1.5 rounded text-sm font-medium transition-colors ${nplModalTab === 'detail' ? 'bg-white shadow-sm text-slate-900' : 'text-slate-500 hover:text-slate-700'}`}
                  >
                    Chi tiết theo SP
                  </button>
                  <button
                    onClick={() => setNplModalTab('aggregate')}
                    className={`px-3 py-1.5 rounded text-sm font-medium transition-colors ${nplModalTab === 'aggregate' ? 'bg-white shadow-sm text-slate-900' : 'text-slate-500 hover:text-slate-700'}`}
                  >
                    Tổng hợp NPL
                    {aggregated.length > 0 && (
                      <span className="ml-1.5 bg-emerald-100 text-emerald-700 text-xs px-1.5 py-0.5 rounded-full">{aggregated.length}</span>
                    )}
                  </button>
                </div>

                <div className="flex items-center gap-2 w-full sm:w-auto justify-end">
                  <button
                    onClick={handleExportNPL}
                    disabled={aggregated.length === 0}
                    className="flex items-center gap-2 px-3 py-1.5 bg-emerald-600 text-white rounded-md hover:bg-emerald-700 disabled:opacity-40 disabled:cursor-not-allowed transition-colors text-sm font-medium shadow-sm"
                  >
                    <Download className="w-4 h-4" />
                    Xuất Excel
                  </button>
                  <button onClick={() => setNplModal(null)} className="p-2 text-slate-400 hover:text-slate-600 hover:bg-slate-200 rounded-full transition-colors">
                    <X className="w-5 h-5" />
                  </button>
                </div>
              </div>

              {/* Warnings */}
              {nplData.length === 0 && (
                <div className="px-6 py-3 bg-amber-50 border-b border-amber-200 flex items-center gap-2 text-amber-700 text-sm shrink-0">
                  <AlertCircle className="w-4 h-4 shrink-0" />
                  Chưa có dữ liệu NPL. Vui lòng tải lên file DATA NPL trong phần <strong className="mx-1">Cài đặt</strong>.
                </div>
              )}
              {nplData.length > 0 && missingCount > 0 && (
                <div className="px-6 py-2 bg-amber-50 border-b border-amber-200 flex items-center gap-2 text-amber-700 text-sm shrink-0">
                  <AlertCircle className="w-4 h-4 shrink-0" />
                  {missingCount} sản phẩm chưa có dữ liệu NPL (ô Định mức hiển thị màu vàng).
                </div>
              )}

              {/* Body */}
              <div className="flex-1 overflow-y-auto p-6">

                {/* ── Tab: Chi tiết ── */}
                {nplModalTab === 'detail' && (
                  <div className="space-y-4">
                    {products.map((product, pIdx) => (
                      <div key={pIdx} className="border border-slate-200 rounded-lg overflow-hidden">
                        <div className="bg-slate-50 px-4 py-3 flex items-center justify-between border-b border-slate-200">
                          <div className="flex items-center gap-2">
                            <span className="font-bold text-slate-900 text-sm font-mono">{product.productCode}</span>
                            <span className="text-slate-500 text-sm">{product.productName}</span>
                          </div>
                          <div className="text-sm flex items-center gap-1">
                            <span className="text-slate-400">SL PO:</span>
                            <span className="font-bold text-indigo-600">{new Intl.NumberFormat('vi-VN').format(product.poQuantity)}</span>
                          </div>
                        </div>

                        {product.materials.length === 0 ? (
                          <div className="px-4 py-3 text-sm text-slate-400 italic flex items-center gap-2 bg-amber-50">
                            <AlertCircle className="w-4 h-4 text-amber-400 shrink-0" />
                            Chưa có dữ liệu NPL cho sản phẩm này — kiểm tra Mã SP trong file DATA NPL.
                          </div>
                        ) : (
                          <table className="w-full text-sm">
                            <thead>
                              <tr className="text-xs text-slate-500 bg-slate-25 border-b border-slate-100">
                                <th className="px-4 py-2 text-left font-medium">Mã NPL</th>
                                <th className="px-4 py-2 text-left font-medium">Tên NPL</th>
                                <th className="px-4 py-2 text-center font-medium w-36">Định Mức</th>
                                <th className="px-4 py-2 text-right font-medium">SL Cần</th>
                                <th className="px-4 py-2 text-center font-medium w-20">ĐVT</th>
                              </tr>
                            </thead>
                            <tbody>
                              {product.materials.map((mat, mIdx) => (
                                <tr key={mIdx} className="border-t border-slate-100 hover:bg-slate-50">
                                  <td className="px-4 py-2 font-mono text-xs text-slate-600">{mat.materialCode}</td>
                                  <td className="px-4 py-2 text-slate-800">{mat.materialName}</td>
                                  <td className="px-4 py-2 text-center">
                                    <input
                                      type="number"
                                      value={mat.norm}
                                      min={0}
                                      step="any"
                                      onChange={(e) => handleNormEdit(product.productCode, mat.materialCode, parseFloat(e.target.value))}
                                      className={`w-28 px-2 py-1 border rounded text-sm text-center focus:outline-none focus:ring-1 focus:ring-emerald-500 focus:border-emerald-500 ${mat.norm === 0 ? 'border-amber-400 bg-amber-50' : 'border-slate-300 bg-white'}`}
                                    />
                                  </td>
                                  <td className="px-4 py-2 text-right font-medium text-slate-800">
                                    {mat.norm === 0
                                      ? <span className="text-amber-500 font-normal italic">—</span>
                                      : new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 4 }).format(mat.quantity)
                                    }
                                  </td>
                                  <td className="px-4 py-2 text-center text-slate-500 text-xs">{mat.unit}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        )}
                      </div>
                    ))}
                  </div>
                )}

                {/* ── Tab: Tổng hợp ── */}
                {nplModalTab === 'aggregate' && (
                  <div>
                    {aggregated.length === 0 ? (
                      <div className="text-center py-16 text-slate-400">
                        <Package className="w-12 h-12 mx-auto mb-3 opacity-30" />
                        <p>Không có dữ liệu NPL để tổng hợp.</p>
                        <p className="text-sm mt-1">Kiểm tra lại dữ liệu NPL và Định mức trong tab Chi tiết.</p>
                      </div>
                    ) : (
                      <table className="w-full text-sm border-collapse">
                        <thead>
                          <tr className="bg-slate-100 text-slate-700">
                            <th className="border border-slate-300 px-4 py-2.5 text-center font-semibold w-14">STT</th>
                            <th className="border border-slate-300 px-4 py-2.5 text-left font-semibold">Mã NPL</th>
                            <th className="border border-slate-300 px-4 py-2.5 text-left font-semibold">Tên Nguyên Phụ Liệu</th>
                            <th className="border border-slate-300 px-4 py-2.5 text-right font-semibold">Tổng SL Cần</th>
                            <th className="border border-slate-300 px-4 py-2.5 text-center font-semibold w-24">Đơn Vị</th>
                          </tr>
                        </thead>
                        <tbody>
                          {aggregated.map((agg, idx) => (
                            <tr key={idx} className="hover:bg-slate-50 border-b border-slate-200">
                              <td className="border border-slate-200 px-4 py-2 text-center text-slate-500">{idx + 1}</td>
                              <td className="border border-slate-200 px-4 py-2 font-mono text-xs text-slate-600">{agg.materialCode}</td>
                              <td className="border border-slate-200 px-4 py-2 text-slate-800">{agg.materialName}</td>
                              <td className="border border-slate-200 px-4 py-2 text-right font-bold text-emerald-700">
                                {new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 4 }).format(agg.totalQuantity)}
                              </td>
                              <td className="border border-slate-200 px-4 py-2 text-center text-slate-500 text-xs">{agg.unit}</td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr className="bg-slate-50">
                            <td colSpan={3} className="border border-slate-300 px-4 py-2 text-right font-semibold text-slate-600">Tổng số loại NPL:</td>
                            <td colSpan={2} className="border border-slate-300 px-4 py-2 text-center font-bold text-emerald-700">{aggregated.length} loại</td>
                          </tr>
                        </tfoot>
                      </table>
                    )}
                  </div>
                )}
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}
