import * as pdfjsLib from 'pdfjs-dist';
// @ts-ignore - pdfjs-dist worker module exports WorkerMessageHandler
import * as pdfjsWorker from 'pdfjs-dist/build/pdf.worker.min.mjs';

if (typeof globalThis !== 'undefined' && pdfjsWorker?.WorkerMessageHandler) {
  (globalThis as Record<string, unknown>).pdfjsWorker = pdfjsWorker;
}
pdfjsLib.GlobalWorkerOptions.workerSrc = '#pdfjs-inline';

export interface PdfPosItem {
  s: string;
  x: number;
  y: number;
  w: number;
}

export interface PdfViewportItem {
  str: string;
  x: number;
  y: number;
  w: number;
}

export async function extractPdfText(file: File): Promise<string> {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  let text = '';
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    text += content.items.map((it: any) => it.str).join(' ') + '\n';
  }
  return text;
}

export async function extractPdfData(file: File): Promise<{ text: string; pages: PdfPosItem[][] }> {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  let text = '';
  const pages: PdfPosItem[][] = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    text += content.items.map((it: any) => it.str).join(' ') + '\n';
    pages.push(
      (content.items as any[])
        .filter((it) => it.str && it.str.trim())
        .map((it) => ({
          s: it.str.trim(),
          x: it.transform[4],
          y: it.transform[5],
          w: it.width || 0,
        }))
    );
  }
  return { text, pages };
}

export async function extractPdfLayout(file: File): Promise<{ pages: PdfViewportItem[][]; text: string }> {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const pages: PdfViewportItem[][] = [];
  let text = '';
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const vp = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();
    pages.push(
      (content.items as any[]).map((it) => {
        const t = pdfjsLib.Util.transform(vp.transform, it.transform);
        return { str: it.str, x: t[4], y: t[5], w: it.width * vp.scale };
      })
    );
    text += content.items.map((it: any) => it.str).join(' ') + '\n';
  }
  return { pages, text };
}
