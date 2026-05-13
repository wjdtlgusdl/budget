export async function extractPdfText(file) {
  const pdfjsLib = await import('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.7.76/pdf.min.mjs');
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.7.76/pdf.worker.min.mjs';
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const text = content.items.map((item) => item.str).join(' ');
    pages.push({ page: i, text, normalized: normalizeText(text) });
  }
  return { raw: pages.map((p) => p.text).join('\n'), pages, normalized: normalizeText(pages.map((p) => p.text).join('\n')) };
}

export function normalizeText(text) {
  return String(text || '').replace(/\s+/g, ' ').replace(/,/g, '').trim();
}
