export interface Supplier {
  code: string;
  name: string;
  address: string;
  phone: string;
  contactPerson: string;
  email: string;
  taxCode?: string;
  deposit?: string;
  bankAccountName?: string;
  bankAccountNumber?: string;
  bankName?: string;
}

export interface MasterPOItem {
  supplierCode: string;
  productName: string;
  quantity: number;
  unitPrice: number;
  unit: string;
}

export interface PO {
  supplier: Supplier;
  items: MasterPOItem[];
  poNumber: string;
  orderDate: string;
  contractExpiry: string;
  vatRate: number;
  poSuffix: string;
  shortCurrentDate: string;
  shortExpiryDate: string;
}

export interface NPLItem {
  productCode: string;
  productName: string;
  materialCode: string;
  materialName: string;
  norm: number;
  unit: string;
}
