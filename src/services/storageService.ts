import { Supplier } from '../types';

export const saveSupplierData = (suppliers: Supplier[]) => {
  localStorage.setItem('procurement_suppliers', JSON.stringify(suppliers));
};

export const loadSupplierData = (): Supplier[] => {
  const data = localStorage.getItem('procurement_suppliers');
  return data ? JSON.parse(data) : [];
};

export const saveTemplateFile = (base64: string) => {
  localStorage.setItem('procurement_template', base64);
};

export const loadTemplateFile = (): string | null => {
  return localStorage.getItem('procurement_template');
};
