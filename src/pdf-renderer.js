import * as pdfjsLib from 'pdfjs-dist';
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('pdfjs-dist/build/pdf.worker.mjs', import.meta.url).href;

function getStandardFontFamily(fontId, page) {
  let fontName = (fontId || '').toLowerCase();
  let fallbackName = '';
  
  if (page && page.commonObjs && page.commonObjs.has(fontId)) {
    try {
      const fontFace = page.commonObjs.get(fontId);
      if (fontFace) {
        if (fontFace.name) {
          fontName = fontFace.name.toLowerCase();
        }
        if (fontFace.fallbackName) {
          fallbackName = fontFace.fallbackName.toLowerCase();
        }
      }
    } catch (e) {
      console.warn('Error reading font from commonObjs:', e);
    }
  }

  // Detect standard serif fonts
  if (
    fontName.includes('times') || 
    fontName.includes('roman') || 
    fontName.includes('serif') || 
    fontName.includes('georgia') || 
    fontName.includes('mincho') || 
    fontName.includes('song') || 
    fallbackName === 'serif'
  ) {
    return 'Times New Roman, Georgia, serif';
  }
  
  // Detect standard monospace fonts
  if (
    fontName.includes('courier') || 
    fontName.includes('mono') || 
    fontName.includes('code') || 
    fontName.includes('console') || 
    fallbackName === 'monospace'
  ) {
    return 'Courier New, Courier, monospace';
  }
  
  // Default to sans-serif
  return 'Arial, Helvetica, sans-serif';
}

export class PdfRenderer {
  constructor() {
    this.pdfDoc = null;
    this.scale = 1.5;
    this.pages = [];
  }

  async load(file) {
    const buf = await file.arrayBuffer();
    this.pdfDoc = await pdfjsLib.getDocument({ data: buf }).promise;
    return this.pdfDoc.numPages;
  }

  async renderPage(pageNum, canvas) {
    const page = await this.pdfDoc.getPage(pageNum);
    const vp = page.getViewport({ scale: this.scale });
    canvas.width = vp.width;
    canvas.height = vp.height;
    const ctx = canvas.getContext('2d');
    await page.render({ canvasContext: ctx, viewport: vp }).promise;
    return { width: vp.width, height: vp.height };
  }

  async renderThumb(pageNum, canvas, maxW = 140) {
    const page = await this.pdfDoc.getPage(pageNum);
    const vp = page.getViewport({ scale: 1 });
    const s = maxW / vp.width;
    const svp = page.getViewport({ scale: s });
    canvas.width = svp.width;
    canvas.height = svp.height;
    await page.render({ canvasContext: canvas.getContext('2d'), viewport: svp }).promise;
  }

  async getTextContent(pageNum) {
    const page = await this.pdfDoc.getPage(pageNum);
    const vp = page.getViewport({ scale: this.scale });
    const textContent = await page.getTextContent();
    const items = [];
    for (const item of textContent.items) {
      if (!item.str) continue;
      
      // Clean text by replacing non-standard space/control characters with standard spaces
      let cleanedStr = item.str
        .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\xA0\u2000-\u200A\u202F\u205F\u3000\uF020]/g, ' ')
        .replace(/[\xAD\u200B-\u200F\uFEFF]/g, '');

      // Convert VNI-Windows to Unicode if detected or if it is a VNI font
      cleanedStr = detectAndConvertVniToUnicode(cleanedStr, item.fontName);

      if (!cleanedStr.trim()) continue; // Ignore empty and whitespace-only items
      
      const tx = pdfjsLib.Util.transform(vp.transform, item.transform);
      const x = tx[4];
      const y = tx[5];
      const fSize = Math.sqrt(tx[2] * tx[2] + tx[3] * tx[3]);
      const w = item.width * this.scale;
      const h = fSize;
      const fn = (item.fontName || '').toLowerCase();
      const isBold = fn.includes('bold') || fn.includes('gothic') || fn.includes('heavy') || fn.includes('black');
      const isItalic = fn.includes('italic') || fn.includes('oblique');
      items.push({
        str: cleanedStr,
        x: x,
        y: y - h,
        width: w,
        height: h * 1.2,
        fontSize: fSize,
        fontName: item.fontName || '',
        fontFamily: getStandardFontFamily(item.fontName || '', page),
        isBold,
        isItalic
      });
    }

    if (items.length === 0) return [];

    // Group items into lines based on Y coordinate closeness (strict 4px threshold)
    const lines = [];
    items.sort((a, b) => a.y - b.y);

    for (const item of items) {
      let added = false;
      for (const line of lines) {
        const lineY = line[0].y;
        // Strict threshold to prevent completely different lines from being grouped together
        if (Math.abs(lineY - item.y) < 4) {
          line.push(item);
          added = true;
          break;
        }
      }
      if (!added) {
        lines.push([item]);
      }
    }

    const mergedItems = [];
    for (const line of lines) {
      // Sort left-to-right
      line.sort((a, b) => a.x - b.x);

      let current = null;
      for (const item of line) {
        if (!current) {
          current = { ...item };
          continue;
        }

        const gap = item.x - (current.x + current.width);
        const maxGap = current.fontSize * 0.8;

        // Do not merge horizontally if font styling (bold or italic) is different
        if (gap < maxGap && current.isBold === item.isBold && current.isItalic === item.isItalic) {
          const needsSpace = gap > current.fontSize * 0.18 && 
                             !current.str.endsWith(' ') && 
                             !item.str.startsWith(' ');
          current.str += (needsSpace ? ' ' : '') + item.str;
          current.width = (item.x + item.width) - current.x;
          current.height = Math.max(current.height, item.height);
          current.fontSize = Math.max(current.fontSize, item.fontSize);
        } else {
          mergedItems.push(current);
          current = { ...item };
        }
      }
      if (current) {
        mergedItems.push(current);
      }
    }

    // Keep single-line/styled-segment blocks separate to preserve precise fonts and styling
    // Sort blocks top-to-bottom, left-to-right
    mergedItems.sort((a, b) => {
      const yDiff = a.y - b.y;
      if (Math.abs(yDiff) < Math.max(a.fontSize, b.fontSize) * 0.4) {
        return a.x - b.x;
      }
      return yDiff;
    });

    return mergedItems;
  }

  getNumPages() { return this.pdfDoc ? this.pdfDoc.numPages : 0; }
}

export function vniToUnicode(str) {
  if (typeof str !== 'string') return str;
  
  const VNI_CHARS = ["AØ", "AÙ", "AÂ", "AÕ", "EØ", "EÙ", "EÂ", "Ì" , "Í" , "OØ",
    "OÙ", "OÂ", "OÕ", "UØ", "UÙ", "YÙ", "aø", "aù", "aâ", "aõ",
    "eø", "eù", "eâ", "ì" , "í" , "oø", "où", "oâ", "oõ", "uø",
    "uù", "yù", "AÊ", "aê", "Ñ" , "ñ" , "Ó" , "ó" , "UÕ", "uõ",
    "Ô" , "ô" , "Ö" , "ö" , "AÏ", "aï", "AÛ", "aû", "AÁ", "aá",
    "AÀ", "aà", "AÅ", "aå", "AÃ", "aã", "AÄ", "aä", "AÉ", "aé",
    "AÈ", "aè", "AÚ", "aú", "AÜ", "aü", "AË", "aë", "EÏ", "eï",
    "EÛ", "eû", "EÕ", "eõ", "EÁ", "eá", "EÀ", "eà", "EÅ", "eå",
    "EÃ", "eã", "EÄ", "eä", "Æ" , "æ" , "Ò" , "ò" , "OÏ", "oï",
    "OÛ", "oû", "OÁ", "oá", "OÀ", "oà", "OÅ", "oå", "OÃ", "oã",
    "OÄ", "oä", "ÔÙ", "ôù", "ÔØ", "ôø", "ÔÛ", "ôû", "ÔÕ", "ôõ",
    "ÔÏ", "ôï", "UÏ", "uï", "UÛ", "uû", "ÖÙ", "öù", "ÖØ", "öø",
    "ÖÛ", "öû", "ÖÕ", "öõ", "ÖÏ", "öï", "YØ", "yø", "Î" , "î" ,
    "YÛ", "yû", "YÕ", "yõ"];

  const UNI_CHARS = ["À", "Á", "Â", "Ã", "È", "É", "Ê", "Ì", "Í", "Ò",
    "Ó", "Ô", "Õ", "Ù", "Ú", "Ý", "à", "á", "â", "ã",
    "è", "é", "ê", "ì", "í", "ò", "ó", "ô", "õ", "ù",
    "ú", "ý", "Ă", "ă", "Đ", "đ", "Ĩ", "ĩ", "Ũ", "ũ",
    "Ơ", "ơ", "Ư", "ư", "Ạ", "ạ", "Ả", "ả", "Ấ", "ấ",
    "Ầ", "ầ", "Ẩ", "ẩ", "Ẫ", "ẫ", "Ậ", "ậ", "Ắ", "ắ",
    "Ằ", "ằ", "Ẳ", "ẳ", "Ẵ", "ẵ", "Ặ", "ặ", "Ẹ", "ẹ",
    "Ẻ", "ẻ", "Ẽ", "ẽ", "Ế", "ế", "Ề", "ề", "Ể", "ể",
    "Ễ", "ễ", "Ệ", "ệ", "Ỉ", "ỉ", "Ị", "ị", "Ọ", "ọ",
    "Ỏ", "ỏ", "Ố", "ố", "Ồ", "ồ", "Ổ", "ổ", "Ỗ", "ỗ",
    "Ộ", "ộ", "Ớ", "ớ", "Ờ", "ờ", "Ở", "ở", "Ỡ", "ỡ",
    "Ợ", "ợ", "Ụ", "ụ", "Ủ", "ủ", "Ứ", "ứ", "Ừ", "ừ",
    "Ử", "ử", "Ữ", "ữ", "Ự", "ự", "Ỳ", "ỳ", "Ỵ", "ỵ",
    "Ỷ", "ỷ", "Ỹ", "ỹ"];

  const mappings = VNI_CHARS.map((vni, idx) => ({ vni, uni: UNI_CHARS[idx] }))
    .sort((a, b) => b.vni.length - a.vni.length);

  let result = str;
  for (const m of mappings) {
    result = result.split(m.vni).join(m.uni);
  }
  return result;
}

export function isVniEncoding(str) {
  if (typeof str !== 'string') return false;
  return /aø|aù|aâ|aõ|eø|eù|eâ|oø|où|oâ|oõ|uø|uù|yù|aê|uõ|aï|aû|aá|aà|aå|aã|aä|aé|aè|aú|aü|aë|eï|eû|eõ|eá|eà|eå|eã|eä|oï|oû|oá|oà|oå|oã|oä|ôù|ôø|ôû|ôõ|ôï|uï|uû|öù|öø|öû|öõ|öï|yø|yû|yõ|ö|ö/i.test(str);
}

export function detectAndConvertVniToUnicode(str, fontName) {
  if (typeof str !== 'string') return str;
  const isVniFont = (fontName || '').toLowerCase().includes('vni');
  if (isVniFont || isVniEncoding(str)) {
    return vniToUnicode(str);
  }
  return str;
}
