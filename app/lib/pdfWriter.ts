// ==========================================
// CẮT TRANG PDF -> FILE PDF MỚI (pdf-lib) · GÓI NHIỀU FILE THÀNH .zip (jszip)
// Trang được sao chép nguyên vẹn (chữ, ảnh, font), không vẽ lại -> chất lượng như file gốc
// ==========================================
import { PDFDocument } from 'pdf-lib';
import JSZip from 'jszip';

export const loadPdfSource = (bytes: Uint8Array) => PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });

export const extractPages = async (src: PDFDocument, pages: number[], title?: string): Promise<Uint8Array> => {
  const out = await PDFDocument.create();
  const copied = await out.copyPages(src, pages);
  copied.forEach((p) => out.addPage(p));
  if (title) out.setTitle(title);
  out.setProducer('Excel Tool');
  return out.save();
};

export const toBlob = (bytes: Uint8Array, type = 'application/pdf') => new Blob([bytes as BlobPart], { type });

export const buildZip = async (files: { name: string; bytes: Uint8Array }[]): Promise<Blob> => {
  const zip = new JSZip();
  for (const f of files) zip.file(f.name, f.bytes);
  return zip.generateAsync({ type: 'blob', compression: 'STORE' });
};
